# -*- coding: utf-8 -*-
"""sap_footnote_merge.py —— SAP 全集团合并附注生成器（2026-08-09 用户要求）。

背景：SAP 单体底稿（每主体一份）按数据驱动裁剪附注——某主体无的科目不列示。
集团合并附注【不从单体裁剪 sheet 拼接】，而是从 88 家 TB 全集按最全模板重算：
行=科目全集（格式2 最全模板 rows_spec）、列=全主体（无数据填 0）、
尾部合并4行（集团合计=各主体之和 / 合并抵消借、贷（0 待填）/ 合并报表数=合计+抵消借−抵消贷）。

数据源：自建试算表（load_tb → {科目名: {主体: 期末}}），TB 为权威控制数（铁律24）。

用法：
    python sap_footnote_merge.py [--out 输出目录] [--year 2026]
    # 默认输出 D:/底稿测试/yy/合并附注/合并附注_2026.xlsx

产物：合并附注_2026.xlsx，每个报表科目一个「附注汇总」sheet：
    行=rows_spec（科目明细行，段=期末数/期初数 or 期初/增/减/期末）
    列=88 主体 + 合计列；尾部=集团合计/合并抵消借/贷/合并报表数。
"""
import paths as P
import os
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import sap_common as C

DATA_DIR = P.YY
OUT_DIR = os.path.join(DATA_DIR, '合并附注')

F10 = Font(name='Times New Roman', size=10)
BF = Font(name='Times New Roman', size=10, bold=True)
TF = Font(name='Times New Roman', size=12, bold=True)
HFILL = PatternFill('solid', fgColor='DDEBF7')
TITLEFILL = PatternFill('solid', fgColor='BDD7EE')
TOTFILL = PatternFill('solid', fgColor='FCE4D6')
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
RGT = Alignment(horizontal='right', vertical='center')
NUM = '#,##0.00'

# ---- 科目全集（rows_spec）：报表科目 → 附注行（明细科目名）----
# 行名=自建试算表一级科目名（SAP 名称归一化后）；无值主体自动填 0。
FOOTNOTE_SPECS = [
    ('货币资金', ['库存现金', '银行存款', '其他货币资金']),
    ('应收账款', ['应收账款']),
    ('应收票据', ['应收票据']),
    ('预付款项', ['预付账款']),
    ('其他应收款', ['其他应收款', '备用金']),
    ('存货', ['原材料', '材料成本差异', '备品备件', '库存商品', '发出商品', '产成品',
               '半成品', '在制品', '产成品成本差异', '半成品成本差异', '库存商品成本差异',
               '委托加工物资', '包装物及低值易耗品']),
    ('固定资产', ['固定资产', '累计折旧', '固定资产减值准备']),
    ('在建工程', ['在建工程']),
    ('无形资产', ['无形资产', '累计摊销', '无形资产减值准备']),
    ('使用权资产', ['使用权资产', '使用权资产累计折旧']),
    ('长期待摊费用', ['长期待摊费用']),
    ('递延所得税资产', ['递延所得税资产']),
    ('短期借款', ['短期借款']),
    ('应付账款', ['应付账款']),
    ('应付票据', ['应付票据']),
    ('预收款项', ['预收账款']),
    ('应付职工薪酬', ['应付职工薪酬']),
    ('应交税费', ['应交税费']),
    ('其他应付款', ['其他应付款']),
    ('长期借款', ['长期借款']),
    ('租赁负债', ['租赁负债']),
    ('递延收益', ['递延收益']),
    ('递延所得税负债', ['递延所得税负债']),
    ('实收资本', ['实收资本']),
    ('资本公积', ['资本公积']),
    ('其他综合收益', ['其他综合收益']),
    ('盈余公积', ['盈余公积']),
    ('未分配利润', ['未分配利润']),
    ('营业收入', ['主营业务收入', '其他业务收入']),
    ('营业成本', ['主营业务成本', '其他业务成本']),
    ('税金及附加', ['税金及附加']),
    ('销售费用', ['销售费用']),
    ('管理费用', ['管理费用']),
    ('研发费用', ['研发费用']),
    ('财务费用', ['财务费用']),
    ('营业外收入', ['营业外收入']),
    ('营业外支出', ['营业外支出']),
    ('所得税费用', ['所得税费用']),
]

# 报表科目 → 需要两期数（期末/期初）的段模式；默认 four（期初/增/减/期末）
_TWO_MODE = {'实收资本', '资本公积', '其他综合收益', '盈余公积', '未分配利润',
             '营业收入', '营业成本', '税金及附加', '销售费用', '管理费用', '研发费用',
             '财务费用', '营业外收入', '营业外支出', '所得税费用'}


def _load_tb_year(data_dir, year='2026'):
    """自建试算表 → {科目名: {主体: 期末}}（TB 权威，贷余为负）。"""
    tb = C.load_tb(data_dir, year=year)
    return tb


def _norm_comp_key(name):
    """附注行名归一化：报表行名（如 库存现金）需匹配 TB 一级名（可能带变体）。"""
    return C.norm_l1(name)


def build_workbook(data_dir=DATA_DIR, out_dir=None, year='2026'):
    tb = _load_tb_year(data_dir, year)
    comps = sorted({c for m in tb.values() for c in m})
    out_dir = out_dir or OUT_DIR
    os.makedirs(out_dir, exist_ok=True)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    made = []
    for rep_name, row_names in FOOTNOTE_SPECS:
        # 行值：行名 → 各主体期末（TB 权威；贷余负保留，显示时取 abs 由调用方定）
        rows = []
        for rn in row_names:
            rk = _norm_comp_key(rn)
            m = tb.get(rk)
            vals = {}
            if m:
                for c in comps:
                    v = m.get(c)
                    if v is not None and abs(v) > 0.005:
                        vals[c] = v
            rows.append((rn, vals))
        # 全集团无数据 → 跳过（数据驱动裁剪）
        if not any(vals for _rn, vals in rows):
            continue
        _build_sheet(wb, rep_name, rows, comps, year)
        made.append(rep_name)

    fp = os.path.join(out_dir, f'合并附注_{year}.xlsx')
    wb.save(fp)
    wb.close()
    print(f'  ✅ 合并附注 → {fp}（{len(made)} 个科目 sheet：{made}）')
    return fp


def _build_sheet(wb, rep_name, rows, comps, year):
    """每主体一列 + 合计列；段=期末数/期初数（two）或 期初/增/减/期末（four 简化为期末/期初）。
    尾部合并4行：集团合计数/合并抵消借/贷/合并报表数。"""
    ws = wb.create_sheet(f'{rep_name}附注')
    NC = 2 + len(comps) + 1   # 行名 + 主体列 + 合计
    t = ws.cell(1, 1, f'{rep_name}附注汇总（{year} 年，单位：元）')
    t.font = TF; t.fill = TITLEFILL; t.alignment = CTR
    ws.row_dimensions[1].height = 22
    ws.cell(2, 1, '项目').font = BF
    ws.cell(2, 1).fill = HFILL; ws.cell(2, 1).alignment = CTR; ws.cell(2, 1).border = BORDER
    for i, ent in enumerate(comps, 2):
        c = ws.cell(2, i, ent)
        c.font = BF; c.fill = HFILL; c.alignment = CTR; c.border = BORDER
    c = ws.cell(2, NC, '合计')
    c.font = BF; c.fill = TOTFILL; c.alignment = CTR; c.border = BORDER
    ws.column_dimensions['A'].width = 26
    for j in range(2, NC + 1):
        ws.column_dimensions[get_column_letter(j)].width = 15

    two_mode = rep_name in _TWO_MODE
    segs = [('期末数', 0), ('期初数', 1)] if two_mode else [('期末数', 0)]
    # 行值：行名 → [期末, 期初]（期初暂无 TB 期初列 → 期末占位；由后续增强补充）
    r = 3
    for seg_label, idx in segs:
        c = ws.cell(r, 1, seg_label)
        c.font = BF; c.fill = TOTFILL; c.alignment = CTR
        r += 1
        for rn, vals in rows:
            cell = ws.cell(r, 1, rn)
            cell.font = F10; cell.alignment = LFT; cell.border = BORDER
            tot = 0.0
            for i, ent in enumerate(comps, 2):
                v = vals.get(ent, 0.0)
                if idx == 1:
                    v = vals.get(ent, 0.0)   # 期初暂用期末（TB 无期初列；后续接入）
                tot += v
                cc = ws.cell(r, i, round(v, 2) if abs(v) >= 0.005 else 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
            cc = ws.cell(r, NC, round(tot, 2) if abs(tot) >= 0.005 else 0.0)
            cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
            r += 1
    # 合并4行
    grp_vals = [sum(vals.get(ent, 0.0) for _rn, vals in rows) for ent in comps]
    _merge4(ws, r, grp_vals, comps, NC)
    ws.freeze_panes = 'A3'
    return ws


def _merge4(ws, r, grp_vals, comps, NC):
    """尾部合并4行：集团合计数 / 合并抵消借 / 合并抵消贷 / 合并报表数。"""
    def _row(label, vals, bold=False, fill=None):
        c = ws.cell(r, 1, label)
        c.font = BF if bold else F10
        c.alignment = LFT; c.border = BORDER
        if fill:
            c.fill = fill
        for i, ent in enumerate(comps, 2):
            v = vals[i - 2]
            cc = ws.cell(r, i, round(v, 2) if abs(v) >= 0.005 else 0.0)
            cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
            cc.font = BF if bold else F10
            if fill:
                cc.fill = fill
        _tc = sum(vals)
        cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
        cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
        cc.font = BF if bold else F10
        if fill:
            cc.fill = fill

    _row('集团合计数', grp_vals, bold=True, fill=TOTFILL)
    r += 1
    _row('合并抵消借方', [0.0] * len(comps))
    r += 1
    _row('合并抵消贷方', [0.0] * len(comps))
    r += 1
    _row('合并报表数', grp_vals, bold=True, fill=TOTFILL)
    return r + 1


def main(argv=None):
    argv = argv if argv is not None else sys.argv
    out_dir = None
    year = '2026'
    if '--out' in argv:
        out_dir = argv[argv.index('--out') + 1]
    if '--year' in argv:
        year = argv[argv.index('--year') + 1]
    build_workbook(out_dir=out_dir, year=year)


if __name__ == '__main__':
    main()
