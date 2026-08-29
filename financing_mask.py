# -*- coding: utf-8 -*-
"""financing_mask.py —— 融资明细.xls 本地先行脱敏（2026-08-22）。

背景：融资明细.xls 是逐笔借款台账（融资单位/银行简称/担保信息均含真实主体名），
此前 loan_detail 仅在【程序内】脱敏，原始文件本身未脱敏 → 任何直接读取原始文件
的分析（如查结构/看完整度）都会把真实名带入对话。本脚本生成【脱敏版】，
此后分析/对话一律只读 `融资明细_脱敏.xls`。

规则（复用 loan_detail 已确认映射，语义化而非 A 编号）：
  col3  融资单位（主体）     → 借X（_fin_ent_code；未命中→【外部主体】）
  col4  银行简称             → 贷X（_mask_bank；未命中→【外部银行】）
  col5  业务编号             → 保留（核对键，不涉主体名）
  col22 担保信息             → 借X/贷X/外部担保方/个人 泛化（_mask_guarantor）
  col0  所属体系（集团名）   → 借X 泛化（含"集团/公司"特征→借?? 或集团）
  其余（日期/金额/利率/类别）→ 保留（核对必需）

用法：
  python financing_mask.py            # 生成 融资明细_脱敏.xls + 残留校验
  python financing_mask.py --dry      # 只打印列头与脱敏样本（不写文件）
"""
import os
import re
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

APP_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = r'd:/底稿测试/ADF/数据/2026/融资明细.xls'
OUT = r'd:/底稿测试/ADF/数据/2026/融资明细_脱敏.xlsx'

try:
    from loan_detail import _fin_ent_code, _mask_bank, _mask_guarantor, _BANK_CODE_MAP
except Exception:
    sys.path.insert(0, APP_DIR)
    from loan_detail import _fin_ent_code, _mask_bank, _mask_guarantor, _BANK_CODE_MAP

# 所属体系列（col0）泛化：含集团特征 → 集团代码，否则保留
_GROUP_PAT = re.compile(r'(集团|公司|股份|控股|有限|体系)')
_FIN_INST_PAT = re.compile(r'(银行|金租|租赁|信托|证券|保险|财务公司)')


def _mask_group(v):
    s = str(v or '').strip()
    if not s:
        return s
    return _mask_guarantor(s)          # 复用担保泛化映射（含盛虹集团/国望等）


def _mask_bank_sub(s):
    """文本内银行名【子串】替换（区别于整格 _mask_bank：整串替换会毁掉长文本）。"""
    s = str(s)
    for kw, code in _BANK_CODE_MAP:
        if kw in s:
            s = s.replace(kw, code)
    return s


def _mask_fin(v):
    """金融机构名泛化（用于业务编号/项目名称等长文本列）：
    银行名子串→贷X，金租/信托/证券/保险残留→【外部金融机构】，担保/主体名→借X/外部担保方/个人。"""
    s = str(v or '').strip()
    if not s:
        return s
    s = _mask_bank_sub(s)              # 文本内银行名 → 贷X
    s = _mask_guarantor(s)             # 借X / 外部担保方 / 个人
    if _FIN_INST_PAT.search(s):
        return '【外部金融机构】'
    return s


# 列规则：索引 → 处理函数（None=保留）
COL_RULES = {
    0: _mask_group,          # 所属体系
    3: lambda v: _fin_ent_code(v) if _fin_ent_code(v) != '借??' else '【外部主体】',  # 融资单位
    4: _mask_bank,           # 银行简称（整格=银行名）
    5: _mask_fin,            # 业务编号（可能含银行名，如 平安银行商票…）
    9: _mask_fin,            # 项目名称（防含金租/银行/主体名）
    22: _mask_guarantor,     # 担保信息（只做主体/担保方泛化，保留 质押/抵押/担保/信用 关键词供类别判定）
}


def _load_real_names():
    """收集原始文件中的候选真实名（含公司/银行/集团特征，供残留校验）。"""
    import xlrd
    wb = xlrd.open_workbook(SRC)
    ws = wb.sheet_by_index(0)
    names = set()
    pat = re.compile(r'[\u4e00-\u9fff]{2,}(公司|集团|银行|股份|控股|有限|分行|支行|金租|合作社)')
    for i in range(ws.nrows):
        for j in range(ws.ncols):
            v = ws.cell_value(i, j)
            if isinstance(v, str) and pat.search(v):
                # 取匹配片段（前后扩展）作为校验名
                for m in pat.finditer(v):
                    start = max(0, m.start() - 2)
                    end = min(len(v), m.end() + 2)
                    names.add(v[start:end])
    return names


def main():
    dry = '--dry' in sys.argv
    import xlrd
    if not os.path.exists(SRC):
        print('[ERROR] 原始文件不存在：%s' % SRC)
        return 1
    wb = xlrd.open_workbook(SRC)
    ws = wb.sheet_by_index(0)
    ncols = ws.ncols

    # ---- 表头：打印前先过泛化（防列头本身含真实名）----
    hdr = [str(ws.cell_value(0, j)) for j in range(ncols)]
    print('总列数 %d / 总行数 %d' % (ncols, ws.nrows))
    print('表头（已泛化）:')
    for j, h in enumerate(hdr):
        print('  col%-3d %s' % (j, _mask_guarantor(h)[:20]))

    if dry:
        print('\n[DRY] 数据样本（脱敏后）前 3 行：')
        for i in range(1, min(4, ws.nrows)):
            vals = []
            for j in range(ncols):
                v = ws.cell_value(i, j)
                fn = COL_RULES.get(j)
                if fn:
                    v = fn(v)
                vals.append(str(v)[:18])
            print('  r%d: %s' % (i, ' | '.join(vals)))
        return 0

    # ---- 生成脱敏版 ----
    from openpyxl import Workbook
    wout = Workbook()
    wso = wout.active
    wso.title = '融资明细'
    replaced = {0: 0, 3: 0, 4: 0, 5: 0, 9: 0, 22: 0}
    # 表头原样（列头是泛词）
    wso.append(hdr)
    for i in range(1, ws.nrows):
        row = [ws.cell_value(i, j) for j in range(ncols)]
        for j, fn in COL_RULES.items():
            if j < len(row) and row[j] not in (None, ''):
                old = str(row[j])
                new = fn(old)
                if new != old:
                    replaced[j] += 1
                row[j] = new
        wso.append(row)
    wout.save(OUT)
    print('已生成脱敏版：%s' % OUT)
    print('替换统计（按列）: %s' % {k: v for k, v in replaced.items()})

    # ---- 残留校验：脱敏版不得再含原始真实名 ----
    real_names = _load_real_names()
    from openpyxl import load_workbook
    wb2 = load_workbook(OUT, read_only=True, data_only=True)
    ws2 = wb2.worksheets[0]
    residue = []
    for ri, row in enumerate(ws2.iter_rows(values_only=True), 1):
        for j, v in enumerate(row):
            if isinstance(v, str):
                for nm in real_names:
                    if nm and nm in v:
                        residue.append((ri, j, nm))
                        break
    wb2.close()
    if residue:
        print('[FAIL] 残留 %d 处：' % len(residue))
        for r in residue[:20]:
            print('  r%d c%d : %s' % (r[0], r[1], r[2]))
        return 2
    print('[OK] 残留 0（候选真实名 %d 个全部覆盖）' % len(real_names))
    return 0


if __name__ == '__main__':
    sys.exit(main())
