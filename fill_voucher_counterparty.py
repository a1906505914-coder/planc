# -*- coding: utf-8 -*-
"""
fill_voucher_counterparty.py —— 凭证/抽凭表「对方科目/对应科目」补全后处理器（与 13 个 builder 解耦）。

背景：GL 源「对方科目(opp)」列 100% 为空，故抽凭表的对方科目须从「同号凭证的其他分录账户」推导。
本脚本扫描所有 *_生成.xlsx 的抽凭/凭证类 sheet，对「对方科目/对应科目」为空的行，按
  (核算主体, 年度, 凭证号, 科目(本方)) 定位 GL 同号凭证，取除本方外的其余账户名拼接填入。

特点：
- 仅填充【原本为空】的单元格，保留 builder 已填值（如费用程序已 enriched），幂等。
- 锁感知：文件被 Excel 占用时跳过并报错，不破坏原文件。
- 逐文件夹运行：python fill_voucher_counterparty.py <文件夹绝对路径或相对Desktop的字母>
"""
import paths as P
import os, sys, glob, re
import openpyxl
from collections import defaultdict

BASE = P.DATA_ROOT
FOLDERS = ['ACB', 'ADZ', 'AZ', 'ATT', 'AYL', 'AS', 'AJJ', 'AQ', 'AFJ', 'ARF']

from audit_common import (read_gl_rows, discover_entities,
                          build_gl_voucher_map, derive_counterparty,
                          filter_inc_cp)

# 文件名片段 -> P&L 本方科目（用于 pl 凭证抽查表，其 sheet 无「科目」列）
PL_SUBJECT_BY_FRAG = {
    "其他收益": "其他收益",
    "所得税费用": "所得税费用",
    "投资收益": "投资收益",
    "营业外收入": "营业外收入",
    "营业外支出": "营业外支出",
    "信用减值损失": "信用减值损失",
    "资产减值损失": "资产减值损失",
    "其他收益": "其他收益",
}

# 命中这些片段的"凭证/抽凭类"工作表，统一由本后处理器按金额精确匹配重写对方科目。
# 2026-07-29 扩展：补入 expense『费用抽查凭证』、longterm『增加检查表/减少检查表』，
# 解决"对方科目空白未被补/loan 乱拼接串未被覆盖"两大盲区；所有片区统一口径。
VOUCH_SHEET_HINTS = ("凭证抽查", "抽凭", "差异凭证汇总", "差异凭证清单",
                     "费用抽查凭证", "增加检查表", "减少检查表", "借方核对")


def norm(s):
    return re.sub(r"\s+", "", str(s)).strip()


def norm_e(s):
    s = str(s).replace("【", "").replace("】", "").strip()
    return s


def norm_y(v):
    if v is None:
        return None
    m = re.search(r"(20\d{2})", str(v))
    return str(m.group(1)) if m else None  # 字符串以匹配 GL map 键值


def find_cols(header_cells):
    """返回 (e_i, y_i, vno_i, vtype_i, subj_i, cp_i, dr_i, cr_i) 的列索引（找不到为 -1）。"""
    e_i = y_i = vno_i = vtype_i = subj_i = cp_i = dr_i = cr_i = -1
    cp_candidates = []
    for i, c in enumerate(header_cells):
        h = str(c)
        if ("对方科目" in h or "对应科目" in h):
            cp_candidates.append(i)
        if e_i < 0 and ("主体" in h) and ("科目" not in h) and ("核算项目" not in h):
            e_i = i
        if y_i < 0 and ("年度" in h or "年份" in h):
            y_i = i
        if vno_i < 0 and ("凭证号" in h or "凭证编号" in h or "凭证号码" in h) \
                and ("凭证字" not in h) and ("凭证类" not in h):
            vno_i = i   # 2026-07-31 修复：增加检查表/减少检查表列名为"凭证编号"，原只认"凭证号"致整表跳过
        if vtype_i < 0 and ("凭证字" in h or "凭证类型" in h or "凭证种类" in h):
            vtype_i = i
        if subj_i < 0 and ("科目(明细)" in h or "本方科目" in h or
                            ("科目" in h and "对方" not in h and "核算项目" not in h and "测试" not in h)):
            subj_i = i
        if dr_i < 0 and ("借方" in h and "借" in h and "贷方" not in h and "对方科目" not in h):
            dr_i = i
        if cr_i < 0 and ("贷方" in h and "贷" in h and "借方" not in h and "对方科目" not in h):
            cr_i = i
    # 优先选「纯对方科目/对应科目」列，避开 对方科目借方/贷方/汇总 等
    if cp_candidates:
        exact = [i for i in cp_candidates if str(header_cells[i]).strip() in ("对方科目", "对应科目")]
        if exact:
            cp_i = exact[0]
        else:
            # 退而求其次：含 对方科目/对应科目 但不含 借/贷/汇总的列
            pref = [i for i in cp_candidates
                    if not any(k in str(header_cells[i]) for k in ("借方", "贷方", "汇总", "发生"))]
            cp_i = pref[0] if pref else cp_candidates[0]
    return e_i, y_i, vno_i, vtype_i, subj_i, cp_i, dr_i, cr_i


def build_gl_map(data_dir):
    """向后兼容别名：实际实现已提至 audit_common.build_gl_voucher_map（单一来源）。"""
    return build_gl_voucher_map(data_dir)


def derive_cp(lines, subject, subj_dr, subj_cr):
    """向后兼容别名：实际实现已提至 audit_common.derive_counterparty（单一来源）。"""
    return derive_counterparty(lines, subject, subj_dr, subj_cr)


def year_from_filename(fn):
    m = re.search(r"(20\d{2})", fn)
    return str(m.group(1)) if m else None  # 字符串以匹配 GL map 键值


def process_file(path, gl_map, pl_subject):
    fn = os.path.basename(path)
    y_file = year_from_filename(fn)
    try:
        wb = openpyxl.load_workbook(path)  # 可写
    except Exception as ex:
        return f"SKIP(打开失败 {ex})"
    changed = 0
    for ws in wb.worksheets:
        sn = ws.title
        if not any(h in sn for h in VOUCH_SHEET_HINTS):
            continue
        # 找表头行
        hdr = None
        hi = -1
        # ⚡⚡ 2026-08-24 修复（抽查表对方科目 100% 空缺根因）：抽查表前几行是抽样
        #   配置说明（审计抽样规模/参数/大额异常样本/抽样配置/标题），表头可能在
        #   row 7+。原只搜前 6 行 → 找不到表头 → 整 sheet 跳过 → 对方科目列全空。
        for ri in range(1, min(ws.max_row, 12) + 1):
            cells = [ws.cell(ri, c).value for c in range(1, min(ws.max_column, 16) + 1)]
            if any("对方科目" in str(c) or "对应科目" in str(c) for c in cells if c is not None):
                hdr = cells
                hi = ri
                break
        if hdr is None:
            continue
        e_i, y_i, vno_i, vtype_i, subj_i, cp_i, dr_i, cr_i = find_cols(hdr)
        if cp_i < 0 or vno_i < 0:
            continue
        for r in range(hi + 1, ws.max_row + 1):
            vno = ws.cell(r, vno_i + 1).value
            if vno is None or not str(vno).strip():
                continue
            vno = norm(vno)
            e = norm_e(ws.cell(r, e_i + 1).value) if e_i >= 0 else None
            y = norm_y(ws.cell(r, y_i + 1).value) if y_i >= 0 else y_file
            if y is None:
                y = y_file
            subject = ws.cell(r, subj_i + 1).value if subj_i >= 0 else pl_subject
            # 本方金额与方向（取借/贷，绝对值）
            subj_dr = abs(float(ws.cell(r, dr_i + 1).value)) if (dr_i >= 0 and ws.cell(r, dr_i + 1).value not in (None, "")) else 0.0
            subj_cr = abs(float(ws.cell(r, cr_i + 1).value)) if (cr_i >= 0 and ws.cell(r, cr_i + 1).value not in (None, "")) else 0.0
            key = (e, y, vno)
            names = gl_map.get(key)
            if not names and "-" in vno:
                # 去掉表头「凭证字-」前缀（如 6-YS51010001 -> YS51010001）再试
                names = gl_map.get((e, y, vno.split("-", 1)[1]))
            # 放宽回退：仅放宽为 (e,*,vno) / (*,y,vno) / 去前缀变体，不放宽到 (*,*,vno) 以免跨主体错配
            if not names:
                base = vno.split("-", 1)[1] if "-" in vno else vno
                for cand in [(e, None, vno), (None, y, vno), (e, None, base), (None, y, base)]:
                    names = gl_map.get(cand)
                    if names:
                        break
            if not names:
                continue
            cp = derive_cp(names, subject, subj_dr, subj_cr)
            if cp:
                # ⚡⚡ 2026-08-23 仅填空（恢复文档『仅填充原本为空』的语义）：非空单元格
                #   保留 builder/源列已填值。旧代码无条件重写——U8 源对方科目列有值（AFJ 等）
                #   被凭证内全科目覆盖，增加检查表剔除失效；费用程序已 enriched 值也被冲掉。
                #   长期资产「增加/减少检查表」填空时仍走真实来源筛选（与 builder 同口径）。
                if '增加检查表' in sn or '减少检查表' in sn:
                    cp = filter_inc_cp(cp)
                cur = ws.cell(r, cp_i + 1).value
                if cur is None or not str(cur).strip():
                    ws.cell(r, cp_i + 1, cp)
                    changed += 1
    if changed:
        try:
            wb.save(path)
            return f"OK(+{changed})"
        except PermissionError:
            wb.close()
            return f"SKIP(文件被Excel占用, +{changed}未保存)"
        except Exception as ex:
            return f"ERR(保存失败 {ex})"
    wb.close()
    return "nochange"


def process_folder(data_dir):
    """对单个文件夹内所有 *_生成.xlsx 的抽凭/凭证类 sheet 补全对方科目（按金额精确匹配）。
    供 audit_finalize.finalize_folder 在 builder 生成后统一调用，使『各程序』凭证抽查表
    均带正确对方科目，无需手动跑。返回填充了对方科目的文件数。"""
    if not data_dir or not os.path.isdir(data_dir):
        return 0
    gl_map = build_gl_map(data_dir)
    total = 0
    files = sorted(glob.glob(os.path.join(data_dir, "*生成*.xlsx")))
    for f in files:
        if os.path.basename(f).startswith("~$"):
            continue
        fn = os.path.basename(f)
        pl_subj = None
        for frag, subj in PL_SUBJECT_BY_FRAG.items():
            if frag in fn:
                pl_subj = subj
                break
        res = process_file(f, gl_map, pl_subj)
        if res.startswith("OK"):
            total += 1
            print(f"  [对方科目] {fn}: {res}")
    return total


def main():
    targets = sys.argv[1:]
    if not targets:
        targets = FOLDERS
    else:
        # 允许传 c/J/... 或完整路径
        mapped = []
        for t in targets:
            if t in FOLDERS:
                mapped.append(t)
            else:
                mapped.append(t)
        targets = mapped
    total = 0
    for t in targets:
        data_dir = t if os.path.isabs(t) else os.path.join(BASE, t)
        if not os.path.isdir(data_dir):
            print(f"[跳过] {t} 不存在")
            continue
        print(f"\n===== 文件夹 {t} ({data_dir}) =====")
        n = process_folder(data_dir)
        total += n
    print(f"\n全部完成。填充了对方科目的文件数: {total}")


if __name__ == "__main__":
    main()
