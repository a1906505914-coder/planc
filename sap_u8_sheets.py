# -*- coding: utf-8 -*-
"""sap_u8_sheets.py —— U8 通用 sheet 构建器（2026-08-09，11 类对齐 U8 样式的公共件）。

提供三个 U8 标准 sheet 的通用构建函数，供各生成器复用：
  1. u8_footnote_sheet(ws)  —— 附注汇总（二级科目 × 主体 × 集团合计）
  2. u8_voucher_check_sheet(ws, rows) —— 凭证抽查表（15 列 U8 标准）
  3. u8_detail_sheet(ws, rows) —— 明细表（核算主体/二级科目/上年数/本期数/增减额/变动率）

对齐基准：<DATA_ROOT>/c 账套 78 份 U8 底稿（audit_templates/U8_底稿样式基准清单.txt）。
"""
import paths as P
import sap_common as C


def u8_footnote_sheet(wb, title, rows, comp, group_total=True):
    """附注汇总：行 = [(二级科目, 金额)]。列：二级科目 | 主体 | 集团合计。
    rows 可为 dict {科目: 金额} 或 [(科目, 金额)]。"""
    ws = wb.create_sheet('附注汇总')
    if isinstance(rows, dict):
        items = sorted(rows.items(), key=lambda x: -abs(x[1]))
    else:
        items = sorted(rows, key=lambda x: -abs(x[1]))
    C.sheet_hdr(ws, ['二级科目', comp, '集团合计'])
    r = 2
    for name, amt in items:
        if abs(amt) < 0.005:
            continue
        C.txt(ws, r, 1, name)
        C.money(ws, r, 2, amt)
        if group_total:
            C.money(ws, r, 3, amt)
        r += 1
    if r > 2:
        C.txt(ws, r, 1, '合计', bold=True, fill=C.TOT_FILL)
        C.money(ws, r, 2, sum(v for _, v in items if abs(v) >= 0.005), C.TOT_FILL)
        if group_total:
            C.money(ws, r, 3, sum(v for _, v in items if abs(v) >= 0.005), C.TOT_FILL)
    else:
        C.txt(ws, 2, 1, f'（{title} 在本期无发生）')
    C.fmt_cols(ws, [40, 16, 16])
    return ws


def u8_voucher_check_sheet(wb, title, rows, max_rows=200):
    """凭证抽查表（U8 15 列标准）：
    核算主体 | 测试序号 | 日期 | 凭证字 | 凭证号 | 二级科目 | 摘要 | 借方金额 | 贷方金额 |
    对方科目 | 与原始凭证相符 | 原始凭证内容 | 原始凭证日期 | 会计处理正确 | 所属时间无误
    rows: [(主体, 日期, 凭证字, 凭证号, 二级科目, 摘要, 借方, 贷方, 对方科目)]"""
    ws = wb.create_sheet('凭证抽查表')
    hdrs = ['核算主体', '测试序号', '日期', '凭证字', '凭证号', '二级科目（或客户/供应商）', '摘要',
            '借方金额', '贷方金额', '对方科目', '与原始凭证相符', '原始凭证内容',
            '原始凭证日期', '会计处理正确', '所属时间无误']
    C.sheet_hdr(ws, hdrs)
    r = 2
    for i, row in enumerate(rows[:max_rows], 1):
        comp, date, ztype, vno, subj, txt, dr, cr, opp = (list(row) + [''] * 9)[:9]
        C.txt(ws, r, 1, comp)
        C.txt(ws, r, 2, i)
        C.txt(ws, r, 3, date)
        C.txt(ws, r, 4, ztype)
        C.txt(ws, r, 5, vno)
        C.txt(ws, r, 6, subj)
        C.txt(ws, r, 7, txt)
        if dr:
            C.money(ws, r, 8, dr)
        if cr:
            C.money(ws, r, 9, cr)
        C.txt(ws, r, 10, opp)
        C.txt(ws, r, 11, '')
        C.txt(ws, r, 12, '')
        C.txt(ws, r, 13, '')
        C.txt(ws, r, 14, '')
        C.txt(ws, r, 15, '')
        r += 1
    if r == 2:
        C.txt(ws, 2, 1, f'（{title} 本期无大额凭证可抽查）')
    widths = [10, 8, 12, 8, 14, 24, 30, 14, 14, 22, 12, 16, 14, 12, 12]
    C.fmt_cols(ws, widths)
    return ws


def u8_detail_sheet(wb, title, rows, col2='二级科目名称', extra=('增减额', '变动率%')):
    """明细表：核算主体 | 二级科目名称 | 上年数 | 本期数 | [增减额 | 变动率%]
    rows: [(主体, 二级科目, 上年数, 本期数)]"""
    ws = wb.create_sheet(f'{title}明细表')
    hdrs = ['核算主体', col2, '上年数', '本期数'] + list(extra)
    C.sheet_hdr(ws, hdrs)
    r = 2
    tot_prev = tot_cur = 0.0
    for comp, subj, prev, cur in rows:
        C.txt(ws, r, 1, comp)
        C.txt(ws, r, 2, subj)
        C.money(ws, r, 3, prev)
        C.money(ws, r, 4, cur)
        if '增减额' in hdrs:
            C.money(ws, r, 5, (cur or 0) - (prev or 0))
        if '变动率%' in hdrs:
            pct = round(100 * ((cur - prev) / prev), 2) if prev else None
            C.put if hasattr(C, 'put') else None
            ws.cell(r, 6, pct).number_format = '0.00'
        tot_prev += prev or 0
        tot_cur += cur or 0
        r += 1
    if r > 2:
        C.txt(ws, r, 1, '合计', bold=True, fill=C.TOT_FILL)
        C.money(ws, r, 3, tot_prev, C.TOT_FILL)
        C.money(ws, r, 4, tot_cur, C.TOT_FILL)
        if '增减额' in hdrs:
            C.money(ws, r, 5, tot_cur - tot_prev, C.TOT_FILL)
    C.fmt_cols(ws, [10, 34, 14, 14, 14, 10])
    return ws
