# -*- coding: utf-8 -*-
"""底稿「合计/小计/总计」统一重算工具（2026-07-31 用户要求）。

背景：部分底稿（如所得税费用勾稽核对表）的合计行是【构建时按全部年份累加】的静态值；
emit_per_year 按年拆分并按年过滤明细行后，合计行（年度列常为空、不被过滤）未跟随重算，
导致 2026 文件合计 = 两年合计、与目标年明细严重不符（如 FY 所得税费用勾稽 2026 合计
2,676,074.59 vs 明细 -31,660.41）。

规则（保守设计，避免误伤）：
  1) 仅处理 A 列值以『合计/小计/总计』开头的行（全集团合计/集团合计/年度合计/主体小计等）；
  2) 重算范围 = 最近一个【表头行】之后 ～ 该合计行之前；
     表头行 = 行内 ≥3 个单元格含常见表头词（核算主体/科目/项目/月份/期间/年度/序号/账户/对方科目/摘要/日期/年份）；
  3) 若范围内还含其他『合计/小计/总计』行（嵌套小计）→ 跳过不重算（避免双计）；
  4) 只覆盖【静态数值或空】单元格，【公式单元格（=开头）一律跳过】——
     审定数 = 期末+审计调整 等联动公式保留（人工填调整数后自动更新）；
  5) 该列明细区无任何数值 → 不动。

用法：python recalc_totals.py <文件夹> [<文件夹>...]
"""
import paths as P
import os
import re
import sys
import glob

import openpyxl

TOTAL_KW = ('合计', '小计', '总计')
HEADER_KW = ('核算主体', '科目', '项目', '月份', '期间', '年度', '年份', '序号', '账户',
             '对方科目', '摘要', '日期', '凭证', '金额', '余额', '数量', '名称',
             # 2026-08-06 二度启用：固定资产等「增加/减少检查表」下半段『主体/账面增加/抽凭增加合计/差额』
             # 核对表头——缺则合计行明细区扩大到上半区 → 39 行标题行(B=序号561)被当数据行 → 42 行合计 +561。
             # 首启用时"16 个新 ERROR"系检查器旧逻辑（说明行含『合计』误判合计行）假报，
             # 检查器已修（勾稽/集团汇总数排除）后此关键词无害。
             '主体', '账面', '差额')


def _is_total_row(row_vals):
    """合计/小计/总计行识别：
    A 列 == 合计/总计/小计 或以 合计/小计/总计 结尾（覆盖 集团合计/全集团合计/年度合计/主体合计）；
    B/C 列 == 合计/小计 或含『小计』（覆盖「（主体小计）」B 列标记表与 gp_other 明细表
    A=主体名/B=空/C=『小计』的三列小计行——2026-08-01 修复：原漏识别导致重算时把小计行
    当数据行求和、年度合计双计）。
    注意 A 列『（全集团全年合计数）』以『计)』结尾 → 不误判为合计行。"""
    if not row_vals:
        return False
    a = str(row_vals[0]).strip() if row_vals[0] is not None else ''
    # 2026-08-06 同步检查器修复：说明/勾稽行（A 列『勾稽：』正文含『合计』字样）排除——
    # 否则 recalc 会把勾稽说明行当合计行、按明细区写入错误数值（v4 污染源头之一）。
    # 2026-08-06 补：职工薪酬『集团汇总数（内容同上面各主体…）』区块标题行同样排除。
    if a.startswith(('勾稽', '注:', '注：', '说明', '备注', '口径', '核对说明', '差异', '集团汇总数')):
        return False
    if a in ('合计', '总计', '小计') or a.endswith(('合计', '小计', '总计')):
        return True
    # 2026-08-06 17:50 修复：A 列【包含】『合计/小计/总计』也识别——loan 明细表
    # 『【01FY本级公司 2025 小计】』『【01FY本级公司 合计】』以『】』结尾 → endswith
    # 不命中 → 小计/合计行被当数据行 → 全集团合计重算 = 数据行+小计+合计 = 3 倍
    # （FY 长期借款明细 459M = 3×153M 回归根因）。
    if '合计' in a or '小计' in a or '总计' in a:
        return True
    for idx in range(1, min(3, len(row_vals))):
        v = row_vals[idx]
        if v is None:
            continue
        x = str(v).strip()
        if x in ('合计', '小计', '总计') or '小计' in x:
            return True
    return False


def _is_header_row(row_vals):
    """表头行：≥3 个单元格含表头关键词（排除合计/小计行本身）。"""
    if _is_total_row(row_vals):
        return False
    hits = 0
    for v in row_vals[:12]:
        if v is None:
            continue
        s = str(v).strip()
        if any(k in s for k in HEADER_KW):
            hits += 1
    return hits >= 3


def recalc_sheet(ws):
    """重算单个 sheet 的合计/小计行。返回修正单元格数。"""
    max_r = ws.max_row
    max_c = min(ws.max_column, 40)
    fixed = 0
    # 预扫：找出所有合计行（A/B 列关键词）
    total_rows = [r for r in range(1, max_r + 1)
                  if _is_total_row([ws.cell(r, c).value for c in range(1, 4)])]
    # 找表头行
    header_rows = []
    for r in range(1, max_r + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, min(max_c, 14) + 1)]
        if _is_header_row(row_vals):
            header_rows.append(r)

    def _nearest_header_before(r):
        hs = [h for h in header_rows if h < r]
        return hs[-1] if hs else None

    for tr in total_rows:
        # 2026-08-07 修复（S 存货 1.75B 双计）：『全集团合计/集团合计』行由 builder
        # 按各主体小计【精确汇总】填入（父级行已剔除不计入小计）——recalc 若按
        # SUM(明细区) 重算会把【父级+子目】全部数据行累加（S 存货 877M 数据行 →
        # 1.75B≈2×877M）。此类表级合计行跳过不重算（信任 builder 填值）。
        _a_tr0 = str(ws.cell(tr, 1).value or '').strip()
        _b_tr0 = str(ws.cell(tr, 2).value or '').strip()
        if '全集团' in _a_tr0 or '集团合计' in _a_tr0 or '全集团' in _b_tr0:
            continue
        hd = _nearest_header_before(tr)
        if hd is None:
            continue  # 无表头 → 无法确定明细区，跳过
        # 2026-08-07 修复：主体块级合计/小计行（A 列『【XX 小计】』『【XX 合计】』，如
        # 借款明细表）范围起点 = 上一个合计/小计行之后（仅该主体块内明细）——原从表头后
        # 起算且只剔除嵌套合计行 → 每个小计=之前所有主体明细+当前（滚动累计：J 借款
        # 仁旭小计=东轴+仁旭 83.35M、金泰小计=全部 222.54M vs 自身 89.59M）。
        # 2026-08-07 再修：块级小计不止方括号格式——Z 应交税费明细表『上分 小计/北分
        # 小计』（A 列=主体名+小计）同样被从表头起算 → 北分小计=上分+北分 滚动累计
        # （1,522.26=-104.04+1,626.30，核对器块级可信合计直接采用后暴露 7 主体假差异）。
        # 块级判定：A 列不以『全集团/集团/年度/全表/本期合计』开头且含『小计』或以
        # 『合计』结尾（含方括号）→ 范围=上一个合计行后；表级合计仍从表头后全范围。
        _a_tr = str(ws.cell(tr, 1).value or '').strip()
        _b_tr = str(ws.cell(tr, 2).value or '').strip()
        _c_tr = str(ws.cell(tr, 3).value or '').strip()
        _prev_tot = [t for t in total_rows if t < tr]
        _prev_tot = _prev_tot[-1] if _prev_tot else None
        _is_grp = any(k in _a_tr for k in ('全集团', '集团合计', '年度合计', '全表', '本期合计')) \
            or _a_tr.startswith(('总计', '小计'))
        # 2026-08-08 修复：块级判定补查 B/C 列——职工薪酬『薪酬分配核对表/分月计提明细表』
        # A 列=年份（2026）、B 列=主体、C 列=『小计』（小计在 C 列），原只看 A 列 → 被当
        # 表级合计从表头起算 → 每主体小计=滚动累计（02FY 小计=01FY+02FY），builder 写的
        # 本主体值被 finalize 重算污染（checker 修复块级规则后暴露 10 ERROR）。
        # 再修：A 列恰为『合计』、B 列=年份/主体（银行存款期间对比按年块）→ 也是块级，
        # 原 _a_tr.startswith('合计') 进 _is_grp → 2026 合计从表头起算=累计（期初 1.39亿+3.70亿=5.09亿）。
        _is_blk = (not _is_grp) and ('小计' in (_a_tr + _b_tr + _c_tr)
                                     or _a_tr.endswith('合计') or _a_tr == '合计')
        if (_a_tr.startswith('【') or _is_blk) and _prev_tot is not None:
            start = _prev_tot + 1
        else:
            start = hd + 1
        # 范围内嵌套合计/小计行：重算时【剔除】（不参与求和，防双计），而非整行跳过——
        # 2026-08-01 修复：旧版整行跳过会保留被旧重算器污染的错误合计（如长投明细表
        # 年度合计双计 3,472,787,675.74，正确应为小计和 1,736,393,837.87）。
        inner = set(t for t in total_rows if start <= t < tr)
        # 逐列重算
        for c in range(1, max_c + 1):
            vals = []
            any_num = False
            for r in range(start, tr):
                if r in inner:
                    continue  # 嵌套小计/合计行剔除出求和
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)):
                    vals.append(v)
                    any_num = True
                elif v is not None and not isinstance(v, str):
                    try:
                        vals.append(float(v))
                        any_num = True
                    except (TypeError, ValueError):
                        pass
            if not any_num:
                continue
            cell = ws.cell(tr, c)
            if isinstance(cell, openpyxl.cell.cell.MergedCell):
                continue  # 2026-08-06：合并区域非左上角只读，跳过该单元格（协议化底稿无合并，旧文件兼容）
            cv = cell.value
            if isinstance(cv, str) and cv.strip().startswith('='):
                continue  # 公式单元格不动
            total = round(sum(vals), 2)
            if cv != total:
                cell.value = total
                fixed += 1
    return fixed


def process_folder(folder):
    files = sorted(glob.glob(os.path.join(folder, '*_生成.xlsx')))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not files:
        print(f'  ⚠️ {folder}：无 *_生成.xlsx 底稿')
        return 0
    n_fix = 0
    n_file = 0
    for fp in files:
        try:
            wb = openpyxl.load_workbook(fp)
        except Exception as ex:
            print(f'  ⚠️ 跳过 {os.path.basename(fp)}：{ex}')
            continue
        changed = False
        # 样式统一（2026-07-31 用户要求：全底稿 Times New Roman 10 号/会计格式/列宽自适应；
        # 覆盖 emit 按年拆分后重建 sheet 与模板残留样式，终审强制兜底）
        try:
            from audit_shell import unify_workbook_style
            unify_workbook_style(wb)
            changed = True
        except Exception as ex:
            print(f'  ⚠️ [{os.path.basename(fp)}] 样式统一失败：{ex}')
        for sn in list(wb.sheetnames):
            # 2026-08-02：职工薪酬 审定表 为「每主体一块、块内含 SUM 汇总行+小计+合计」特殊结构，
            # 通用合计重算会把 SUM 汇总行当明细而改坏值 → 跳过（正确性由构建勾稽保证）
            if sn == '职工薪酬 审定表':
                continue
            ws = wb[sn]
            try:
                n = recalc_sheet(ws)
            except Exception as ex:
                print(f'  ⚠️ [{os.path.basename(fp)}][{sn}] 重算失败：{ex}')
                continue
            if n:
                n_fix += n
                changed = True
        if changed:
            try:
                wb.save(fp)
                n_file += 1
            except Exception as ex:
                print(f'  ⚠️ 保存失败（可能被 Excel 占用，请关闭后重跑）：{os.path.basename(fp)}：{ex}')
        wb.close()
    print(f'  ✅ {folder}：{len(files)} 份底稿，{n_file} 份有修正，共 {n_fix} 个合计单元格')
    return n_fix


def main():
    dirs = sys.argv[1:] or [P.G]
    total = 0
    for d in dirs:
        total += process_folder(d)
    print(f'\n合计重算完成，共修正 {total} 个单元格。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
