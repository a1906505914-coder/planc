# -*- coding: utf-8 -*-
"""存货监盘支持表（审计底稿第5层·监盘）。
目的：为年结日存货实地盘点提供「账面年末控制数」与分类汇总，供监盘后倒扎核对。
数据来源：各主体《科目余额表》(TB, 14xx 存货科目期末余额；无数量/单价列，品名级实盘由审计师现场完成)。
方法：
  · 提取 TB 中 14 开头存货科目(原材料/在途/库存商品/发出商品/周转材料/半成品/委托加工/存货跌价准备等)；
  · 监盘基准=2025-12-31(TB 2025 期末)；2026期初作衔接校验(应=2025期末)；
  · 存货净值 = 存货资产类(借)合计 − 存货跌价准备(贷)；
  · 父级(等级1)金额已含子目之和，合计仅取父级避免双计(铁律：父+子累加=双计)。
输出：存货监盘支持表_生成.xlsx
  · 存货监盘支持表(分科目)
  · 存货类别汇总(监盘控制数)
  · 监盘说明与倒扎勾稽
口径：FY2025，年结日=2025-12-31；具体存货项目(仓库/品名/规格/批号)实盘为现场程序，本表仅提供账面控制数。
"""
import paths as P
import os, re, sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import audit_common as A

DATA_DIR = P.G
OUT = os.path.join(P.G, '存货监盘支持表_生成.xlsx')

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D9E1F2')
WARNFILL = PatternFill('solid', fgColor='FFF2CC')
BADFILL = PatternFill('solid', fgColor='FCE4D6')
OKFILL = PatternFill('solid', fgColor='E2EFDA')
HFONT = Font(bold=True, name='Times New Roman', size=10)
BFONT = Font(bold=True, name='Times New Roman', size=10)
MONEY = '#,##0.00'

INV_CAT = {
    '1401': '材料采购', '1402': '在途物资', '1403': '原材料', '1404': '材料成本差异',
    '1405': '库存商品', '1406': '发出商品', '1407': '商品进销差价', '1408': '委托加工物资',
    '1411': '周转材料', '1412': '半成品', '1461': '存货跌价准备', '1471': '存货跌价准备',
}
PROVISION_CATS = ('存货跌价准备',)


def inv_category(code):
    g = re.sub(r'\D', '', str(code))[:4]
    return INV_CAT.get(g, '其他存货')


def build(data_dir, out_path):
    entities = A.discover_entities(data_dir)
    tb = A.read_tb_full(data_dir, entities)
    print(f'· TB 控制数 {len(tb)} 项；主体 {len(entities)} 个')

    # 收集 14xx 存货科目：(e, code, name, y) -> v
    inv = {}
    for key, v in tb.items():
        e, code, name, y = key
        if re.sub(r'\D', '', str(code)).startswith('14'):
            inv[key] = v
    print(f'· 存货科目(TB) {len(inv)} 项')

    # 逐主体·年 构建记录（仅父级 level==1 用于合计；子目一并展示）
    recs = []  # (e, y, code, name, cat, level, dirn, qm, qc)
    for (e, code, name, y), v in inv.items():
        cat = inv_category(code)
        qm = v.get('qm', 0.0)
        qc = v.get('qc', 0.0)
        dirn = '借' if qm > 0 else ('贷' if qm < 0 else '平')
        recs.append((e, y, code, name, cat, v.get('level', 1), dirn, qm, qc))

    # 排序：主体, 年, 类别序, 代码
    cat_order = list(INV_CAT.values()) + ['其他存货']
    def cat_rank(c):
        return cat_order.index(c) if c in cat_order else 99
    recs.sort(key=lambda r: (r[0], r[1], cat_rank(r[4]), r[2]))

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def style_header(ws, hdr, row=3):
        for c, h in enumerate(hdr, 1):
            cell = ws.cell(row, c, h)
            cell.font = HFONT; cell.fill = HFILL; cell.border = BORDER
            cell.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')

    # ---------- Sheet1 存货监盘支持表(分科目) ----------
    ws1 = wb.create_sheet('存货监盘支持表(分科目)')
    ws1.cell(1, 1, '存货监盘支持表（账面年末控制数，基准日 2025-12-31）').font = Font(bold=True, name='Times New Roman', size=10)
    ws1.cell(2, 1, '来源=TB(14xx 存货科目)。账面金额=2025期末(监盘基准)；2026期初作衔接校验(应=2025期末)。'
                    '「监盘备注」由审计人员填：盘点日/实盘数量金额/倒扎至年结日数/盘盈盘亏差异及追查。').font = \
        Font(italic=True, name='Times New Roman', size=10, color='808080')
    hdr1 = ['核算主体', '年度', '科目代码', '科目名称', '存货大类', '等级', '期末方向',
            '账面金额(2025-12-31)', '2026期初(衔接)', '监盘备注']
    style_header(ws1, hdr1, row=3)
    r = 4
    for e, y, code, name, cat, level, dirn, qm, qc in recs:
        if y != '2025':
            continue  # 基准年主表仅列 2025
        row = [e, y, code, name, cat, level, dirn, abs(qm), abs(qc), '']
        for c, val in enumerate(row, 1):
            cell = ws1.cell(r, c, val); cell.border = BORDER
            if c in (8, 9):
                cell.number_format = MONEY
            if cat in PROVISION_CATS:
                cell.fill = WARNFILL
        r += 1
    # 合计（父级 level==1，避免双计）
    g_qm = sum(abs(v.get('qm', 0.0)) for (e, code, name, y), v in inv.items() if y == '2025' and v.get('level') == 1)
    ws1.cell(r, 1, '合计(父级)').font = BFONT
    ws1.cell(r, 4, '存货科目账面合计').font = BFONT
    c = ws1.cell(r, 8, g_qm); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
    for cc in (1, 2, 3, 5, 6, 7, 9, 10):
        ws1.cell(r, cc).border = BORDER
    for i, w in enumerate([14, 7, 11, 24, 14, 6, 9, 20, 18, 30], 1):
        ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws1.freeze_panes = 'A4'

    # ---------- Sheet2 存货类别汇总(监盘控制数) ----------
    ws2 = wb.create_sheet('存货类别汇总(监盘控制数)')
    ws2.cell(1, 1, '存货类别汇总（监盘控制数，基准日 2025-12-31）').font = Font(bold=True, name='Times New Roman', size=10)
    ws2.cell(2, 1, '按主体×存货大类汇总父级(等级1)账面金额；存货净值=存货资产类−存货跌价准备。').font = \
        Font(italic=True, name='Times New Roman', size=10, color='808080')
    hdr2 = ['核算主体', '存货大类', '账面金额(2025)', '其中:存货跌价准备', '存货净值', '2026期初', '备注']
    style_header(ws2, hdr2, row=3)
    r = 4
    ents_order = sorted({x[0] for x in recs})
    for e in ents_order:
        # 该主体 2025 父级各科目
        pars = [(code, name, inv_category(code), v) for (ee, code, name, yy), v in inv.items()
                if ee == e and yy == '2025' and v.get('level') == 1 and re.sub(r'\D', '', str(code)).startswith('14')]
        # 按类别聚合
        by_cat = {}
        for code, name, cat, v in pars:
            by_cat.setdefault(cat, {'amt': 0.0, 'prov': 0.0, 'qc': 0.0})
            by_cat[cat]['amt'] += v.get('qm', 0.0)
            by_cat[cat]['qc'] += v.get('qc', 0.0)
            if cat in PROVISION_CATS:
                by_cat[cat]['prov'] += v.get('qm', 0.0)
        cats_sorted = sorted(by_cat.keys(), key=cat_rank)
        ent_gross = 0.0
        ent_prov = 0.0
        ent_qc = 0.0
        for cat in cats_sorted:
            d = by_cat[cat]
            amt = abs(d['amt'])
            if cat in PROVISION_CATS:
                ent_prov += amt
            else:
                ent_gross += amt
                ent_qc += d['qc']
            row = [e, cat, amt, abs(d['prov']), (0.0 if cat in PROVISION_CATS else amt), abs(d['qc']), '']
            for c, val in enumerate(row, 1):
                cell = ws2.cell(r, c, val); cell.border = BORDER
                if c in (3, 4, 5, 6):
                    cell.number_format = MONEY
                if cat in PROVISION_CATS:
                    cell.fill = WARNFILL
            r += 1
        # 主体小计（存货净值 = 存货资产毛额 − 存货跌价准备）
        c = ws2.cell(r, 1, e); c.font = BFONT
        ws2.cell(r, 2, '小计·存货净值').font = BFONT
        c = ws2.cell(r, 3, abs(ent_gross)); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
        c = ws2.cell(r, 4, abs(ent_prov)); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
        c = ws2.cell(r, 5, abs(ent_gross - ent_prov)); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
        c = ws2.cell(r, 6, abs(ent_qc)); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
        ws2.cell(r, 7).border = BORDER
        r += 1
    # 全集团合计（存货净值 = 存货资产毛额 − 存货跌价准备）
    total_gross = sum(abs(v.get('qm', 0.0)) for (e, code, name, y), v in inv.items()
                      if y == '2025' and v.get('level') == 1 and inv_category(code) not in PROVISION_CATS
                      and re.sub(r'\D', '', str(code)).startswith('14'))
    total_prov = sum(abs(v.get('qm', 0.0)) for (e, code, name, y), v in inv.items()
                     if y == '2025' and v.get('level') == 1 and inv_category(code) in PROVISION_CATS
                     and re.sub(r'\D', '', str(code)).startswith('14'))
    ws2.cell(r, 1, '全集团合计').font = BFONT
    ws2.cell(r, 3, abs(total_gross)).font = BFONT
    ws2.cell(r, 4, abs(total_prov)).font = BFONT
    ws2.cell(r, 5, abs(total_gross - total_prov)).font = BFONT
    for cc in (3, 4, 5):
        ws2.cell(r, cc).number_format = MONEY; ws2.cell(r, cc).border = BORDER
    for cc in (1, 2, 6, 7):
        ws2.cell(r, cc).border = BORDER
    for i, w in enumerate([14, 16, 20, 20, 18, 18, 22], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2.freeze_panes = 'A4'

    # ---------- Sheet3 监盘说明与倒扎勾稽 ----------
    ws3 = wb.create_sheet('监盘说明与倒扎勾稽')
    notes = [
        '存货监盘支持表 编制与倒扎勾稽说明',
        '',
        '一、监盘目标与基准（第5层·监盘）',
        '  对年结日(2025-12-31)存货实施实地盘点，验证账面存货存在性、完整性、计价(成本与可变现净值)。',
        '  本表提供「账面年末控制数」(TB 2025 期末)，供监盘后倒扎核对；品名级实盘(仓库/品名/规格/批号/数量)由审计师现场完成。',
        '',
        '二、存货净值计算',
        '  存货净值 = Σ存货资产类科目(借余：原材料/在途/库存商品/发出商品/周转材料/半成品/委托加工) − 存货跌价准备(贷余)。',
        f'  全集团 2025 存货资产账面(父级) = {total_gross:,.2f}；存货跌价准备 = {total_prov:,.2f}；存货净值 = {total_gross - total_prov:,.2f}。',
        '',
        '三、2025期末 ↔ 2026期初 衔接校验（应相等，验证期后结转连续）',
    ]
    for e in ents_order:
        qm25 = sum(v.get('qm', 0.0) for (ee, code, name, yy), v in inv.items()
                   if ee == e and yy == '2025' and v.get('level') == 1 and re.sub(r'\D', '', str(code)).startswith('14'))
        qc26 = sum(v.get('qc', 0.0) for (ee, code, name, yy), v in inv.items()
                   if ee == e and yy == '2026' and v.get('level') == 1 and re.sub(r'\D', '', str(code)).startswith('14'))
        diff = qm25 - qc26
        flag = '' if abs(diff) <= max(1.0, abs(qm25) * 1e-6) else '  <== 差异!'
        notes.append(f"  {e:<12} 2025期末={qm25:,.2f}  2026期初={qc26:,.2f}  差={diff:,.2f}{flag}")
    notes += [
        '',
        '四、数据口径限制',
        '  · TB 仅含金额，无数量/单价/品名明细；具体项目实盘数须由审计师在仓库现场获取，本表不能替代实盘。',
        '  · 2026 GL 仅含 1-5 月(YTD)；2026期初(本表)作衔接校验，监盘基准仍为 2025 期末。',
        '  · 父级(等级1)金额已含子目之和，合计仅取父级，避免父+子双计(铁律)。',
        '',
        '五、监盘与倒扎程序建议',
        '  1. 监盘日前获取末级存货清单(品名/规格/库位/账面数量)，作为监盘表；',
        '  2. 监盘日现场点数(含抽盘+倒推：监盘日+监盘日至年结日入库−出库=年结日应存)，与本表核对；',
        '  3. 关注残冷背次存货与跌价准备计提充分性；对寄存/代销/在途存货执行函证/检查权属；',
        '  4. 盘盈盘亏差异追查(计量误差/出入库截止/盗损)，必要时调整账面。',
    ]
    for i, t in enumerate(notes, 1):
        ws3.cell(i, 1, t)
        if i == 1:
            ws3.cell(i, 1).font = Font(name='Times New Roman', bold=True, size=10)
    ws3.column_dimensions['A'].width = 120

    path, warn = A._safe_save(wb, out_path)
    if warn:
        print('  ⚠️', warn)
    print(f'  ✓ 已保存：{path}  (存货科目 {len(recs)} 项，主体 {len(ents_order)} 个)')
    return path


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else DATA_DIR
    out = sys.argv[2] if len(sys.argv) > 2 else OUT
    build(data_dir, out)


if __name__ == '__main__':
    main()
