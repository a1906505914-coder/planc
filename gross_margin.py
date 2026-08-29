# -*- coding: utf-8 -*-
"""毛利率分析表（收入-成本按主体/二级）— 审计底稿第3层(科目自身分析)。
数据来源：各主体《科目余额表》(TB, 权威控制数)。
口径：收入=主营业务收入(6001)本期贷方 gross；成本=主营业务成本(6401)本期借方 gross；
      其他业务收支(6051/6402)单列。结平年度(net=0) gross 口径仍有效（铁律3）。
输出：毛利率分析表_生成.xlsx
  · 毛利率分析表(按主体)：收入/成本/毛利/毛利率 2025 vs 2026(YTD)，毛利率变动+异常标记
  · 毛利率分析表(按二级)：各主体 6001/6401 二级明细 收入/成本/毛利/毛利率
  · 勾稽说明：收入/成本与 TB 一级控制数核对；2026 为 1-5 月 YTD 提示
注意：2026 TB=全年、但 GL 仅 1-5 月；毛利率用 TB gross 不受 GL 限制。
"""
import paths as P
import os, sys, re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import audit_common as A

DATA_DIR = P.G
OUT = os.path.join(P.G, '毛利率分析表_生成.xlsx')

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D9E1F2')
SUBFILL = PatternFill('solid', fgColor='F2F6FC')
WARNFILL = PatternFill('solid', fgColor='FFF2CC')
HFONT = Font(bold=True, name='Times New Roman', size=10)
BFONT = Font(bold=True, name='Times New Roman', size=10)
MONEY = '#,##0.00'
PCT = '0.00%'


def _dg(code):
    return re.sub(r'\D', '', str(code))


def tb_rev_cost_l2(tb):
    """返回 (tot, l2)：
    tot[(e,y)] = {rev,cost,rev2,cost2}
    l2[(e,y,'rev'/'cost', l2name)] = amount
    """
    MARGIN_TOL = 1.0
    tot = {}
    l2 = {}
    for (e, code, name, y), v in tb.items():
        dg = _dg(code)
        if len(dg) < 4:
            continue
        nm = name or ''
        is_rev = ('主营业务收入' in nm) or dg.startswith('6001')
        is_cost = ('主营业务成本' in nm) or dg.startswith('6401')
        is_rev2 = ('其他业务收入' in nm) or dg.startswith('6051')
        is_cost2 = ('其他业务成本' in nm) or dg.startswith('6402')
        if not (is_rev or is_cost or is_rev2 or is_cost2):
            continue
        key = (e, y)
        d = tot.setdefault(key, {'rev': 0.0, 'cost': 0.0, 'rev2': 0.0, 'cost2': 0.0})
        if is_rev:
            d['rev'] += v['df']            # 收入 gross = 贷方
            if len(dg) == 6:
                l2[(e, y, 'rev', nm)] = l2.get((e, y, 'rev', nm), 0.0) + v['df']
        if is_cost:
            d['cost'] += v['jf']           # 成本 gross = 借方
            if len(dg) == 6:
                l2[(e, y, 'cost', nm)] = l2.get((e, y, 'cost', nm), 0.0) + v['jf']
        if is_rev2:
            d['rev2'] += v['df']
            if len(dg) == 6:
                l2[(e, y, 'rev', nm)] = l2.get((e, y, 'rev', nm), 0.0) + v['df']
        if is_cost2:
            d['cost2'] += v['jf']
            if len(dg) == 6:
                l2[(e, y, 'cost', nm)] = l2.get((e, y, 'cost', nm), 0.0) + v['jf']
    return tot, l2


def _margin(rev, cost):
    if abs(rev) < 1.0:
        return None
    return (rev - cost) / rev


def _fmt_money(ws, r, c, v):
    cell = ws.cell(r, c, v if v is not None else None)
    cell.number_format = MONEY
    return cell


def build(data_dir, out_path):
    entities = A.discover_entities(data_dir)
    if not entities:
        print('❌ 未发现任何账套主体'); return None
    tb = A.read_tb_full(data_dir, entities)
    years = sorted({y for e, yd in entities.items() for y in yd})
    tot, l2 = tb_rev_cost_l2(tb)
    print(f'· 主体 {len(entities)} 个，年份 {years}')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ---------- Sheet1 按主体 ----------
    ws = wb.create_sheet('毛利率分析表(按主体)')
    ws.cell(1, 1, '毛利率分析表（按主体）').font = Font(bold=True, name='Times New Roman', size=10)
    ws.cell(2, 1, '数据来源：《科目余额表》主营业务收入(6001)/成本(6401) gross 发生额；'
                  '2026 为全年 TB，毛利率可与 2025 直接比较（TB gross 不受 GL 1-5月限制）。').font = \
        Font(italic=True, name='Times New Roman', size=10, color='808080')
    hdr = ['核算主体', '年度', '主营业务收入', '主营业务成本', '毛利', '毛利率',
           '其他业务收入', '其他业务成本',
           '营业收入合计', '营业成本合计', '综合毛利率', '毛利率变动(pp)', '异常标记']
    ws.append([])  # row3 placeholder
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(3, c, h); cell.font = HFONT; cell.fill = HFILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')
    r = 4
    prev = {}  # (e) -> margin 2025
    entity_order = sorted(entities.keys())
    for e in entity_order:
        for y in sorted(years):
            d = tot.get((e, y), {'rev': 0.0, 'cost': 0.0, 'rev2': 0.0, 'cost2': 0.0})
            rev, cost = d['rev'], d['cost']
            rev2, cost2 = d['rev2'], d['cost2']
            m = _margin(rev, cost)
            tot_rev = rev + rev2
            tot_cost = cost + cost2
            m_all = _margin(tot_rev, tot_cost)
            # 毛利率变动
            d_m = None
            if y != years[0] and (e, years[0]) in tot:
                pm = _margin(tot[(e, years[0])]['rev'], tot[(e, years[0])]['cost'])
                if m is not None and pm is not None:
                    d_m = (m - pm) * 100.0
            # 异常标记
            flags = []
            if m is not None:
                if m < 0:
                    flags.append('毛利为负')
                if abs(m) > 0.8:
                    flags.append('毛利率异常高(>80%)')
            if d_m is not None and abs(d_m) > 10:
                flags.append(f'毛利率变动>{abs(d_m):.0f}pp')
            row = [e, y, rev, cost, rev - cost, m, rev2, cost2,
                   tot_rev, tot_cost, m_all, d_m, '；'.join(flags)]
            for c, val in enumerate(row, 1):
                cell = ws.cell(r, c, val); cell.border = BORDER
                if c in (3, 4, 5, 7, 8, 9, 10):
                    cell.number_format = MONEY
                elif c in (6, 11):
                    cell.number_format = PCT
                elif c == 12 and val is not None:
                    cell.number_format = '0.00'
                if c == 13 and flags:
                    cell.fill = WARNFILL
            r += 1
    # 合计行
    tr = r
    ws.cell(tr, 1, '全集团合计').font = BFONT
    g_rev = sum(tot.get((e, y), {}).get('rev', 0.0) for e in entity_order for y in years)
    g_cost = sum(tot.get((e, y), {}).get('cost', 0.0) for e in entity_order for y in years)
    g_rev2 = sum(tot.get((e, y), {}).get('rev2', 0.0) for e in entity_order for y in years)
    g_cost2 = sum(tot.get((e, y), {}).get('cost2', 0.0) for e in entity_order for y in years)
    gm = _margin(g_rev, g_cost)
    gma = _margin(g_rev + g_rev2, g_cost + g_cost2)
    for c, val in [(3, g_rev), (4, g_cost), (5, g_rev - g_cost), (6, gm),
                   (7, g_rev2), (8, g_cost2), (9, g_rev + g_rev2), (10, g_cost + g_cost2), (11, gma)]:
        cell = ws.cell(tr, c, val); cell.font = BFONT; cell.border = BORDER
        cell.number_format = MONEY if c in (3, 4, 5, 7, 8, 9, 10) else PCT
    ws.cell(tr, 2, '').border = BORDER
    for i, w in enumerate([16, 7, 16, 16, 15, 10, 13, 13, 15, 15, 11, 13, 22], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'

    # ---------- Sheet2 按二级 ----------
    ws2 = wb.create_sheet('毛利率分析表(按二级)')
    ws2.cell(1, 1, '毛利率分析表（按二级明细）').font = Font(bold=True, name='Times New Roman', size=10)
    hdr2 = ['核算主体', '年度', '类别', '二级科目', '发生额(收入=贷/成本=借)', '配对成本/收入', '毛利', '毛利率']
    for c, h in enumerate(hdr2, 1):
        cell = ws2.cell(3, c, h); cell.font = HFONT; cell.fill = HFILL; cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')
    r = 4
    keys = sorted(l2.keys())
    for (e, y, kind, nm) in keys:
        amt = l2[(e, y, kind, nm)]
        # 配对：收入找同名成本，成本找同名收入
        pair = None
        if kind == 'rev':
            pair = l2.get((e, y, 'cost', nm))
        else:
            pair = l2.get((e, y, 'rev', nm))
        if kind == 'rev':
            rev_a, cost_a = amt, (pair or 0.0)
        else:
            rev_a, cost_a = (pair or 0.0), amt
        m = _margin(rev_a, cost_a) if kind == 'rev' else _margin(cost_a, rev_a)
        row = [e, y, '收入' if kind == 'rev' else '成本', nm, amt, pair if pair is not None else '',
               (rev_a - cost_a) if kind == 'rev' else (cost_a - rev_a), m]
        for c, val in enumerate(row, 1):
            cell = ws2.cell(r, c, val); cell.border = BORDER
            if c in (5, 6, 7):
                cell.number_format = MONEY
            if c == 8 and m is not None:
                cell.number_format = PCT
        r += 1
    for i, w in enumerate([16, 7, 8, 30, 22, 18, 16, 10], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2.freeze_panes = 'A4'

    # ---------- Sheet3 勾稽说明 ----------
    ws3 = wb.create_sheet('勾稽说明')
    notes = [
        '毛利率分析表 编制与勾稽说明',
        '',
        '一、数据来源',
        '  收入/成本均取自各主体《科目余额表》(TB) 一级/二级科目 gross 发生额：',
        '    主营业务收入(6001) 取「本期贷方」；主营业务成本(6401) 取「本期借方」；',
        '    其他业务收入(6051)/其他业务成本(6402) 单列。结平年度 net=0，gross 口径仍有效。',
        '',
        '二、与 TB 控制数核对',
        '  本表收入/成本即直接摘自 TB，故与 TB 一级控制数天然一致（无需额外差异调整）。',
        '  若某主体无 6001/6401（境外子公司采用本地科目表，如 4100/4200），将未纳入，需人工补正。',
        '',
        '三、口径提示',
        '  · 2026 科目余额表为全年，毛利率可与 2025 直接比较（不受 GL 仅 1-5 月限制）。',
        '  · 毛利率变动>10pp 或毛利为负/异常高(>80%) 已在本表(按主体)「异常标记」列标黄。',
        '  · 本表为科目自身分析（第3层），旨在发现虽为正常分录形成、但毛利率异常波动可能隐含的会计差错。',
    ]
    for i, t in enumerate(notes, 1):
        ws3.cell(i, 1, t)
        if i == 1:
            ws3.cell(i, 1).font = Font(name='Times New Roman', bold=True, size=10)
    ws3.column_dimensions['A'].width = 100

    path, warn = A._safe_save(wb, out_path)
    if warn:
        print('  ⚠️', warn)
    print(f'  ✓ 已保存：{path}')
    return path


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else DATA_DIR
    out = sys.argv[2] if len(sys.argv) > 2 else OUT
    build(data_dir, out)


if __name__ == '__main__':
    main()
