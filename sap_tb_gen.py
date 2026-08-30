# -*- coding: utf-8 -*-
"""SAP 自建试算表生成器（sap_tb_gen.py，2026-08-07 新建）。

用 sap_reader 读取 SAP 科目余额表 → 生成自建试算表 xlsx：
  Sheet1 试算表：一级科目（名称首段）× 公司代码，期末余额（贷余负/借余正，符号口径）
  Sheet2 资产负债表：按报表口径聚合（资产/负债/权益），负债权益贷余转正
  Sheet3 利润表：损益科目发生额（收入贷/费用借）
核对基准：企业报表（利润表/资产负债表 .xls，会企格式）——用户 Q5 决定：企业报表仅作核对基准。
"""
import paths as P
import os
import re
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import sap_reader as SR
import sap_common as C

# ---- 样式（与现有底稿一致）----
TITLE_FONT = Font(name='Times New Roman', size=12, bold=True)
HEAD_FONT = Font(name='Times New Roman', size=10, bold=True)
FONT = Font(name='Times New Roman', size=10)
BOLD = Font(name='Times New Roman', size=10, bold=True)
HEAD_FILL = PatternFill('solid', fgColor='DDEBF7')
TOT_FILL = PatternFill('solid', fgColor='FCE4D6')
NUMFMT = '#,##0.00'
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
RGT = Alignment(horizontal='right', vertical='center')
LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# 报表口径行映射（利润表/资产负债表科目 → TB 名称关键词）
BS_ASSET_ROWS = [
    ('货币资金', ['库存现金', '银行存款', '其他货币资金']),
    ('应收账款', ['应收账款']),
    ('应收票据', ['应收票据']),
    ('预付款项', ['预付账款']),
    ('其他应收款', ['其他应收款', '备用金', '应收股利', '应收利息']),
    ('存货', ['存货', '原材料', '库存商品', '半成品', '产成品', '发出商品', '在产品', '委托加工',
              '包装物', '低值易耗品', '材料成本差异', '库存产品成本差异', '生产成本']),
    ('合同资产', ['合同资产']),
    ('长期股权投资', ['长期股权投资']),
    ('投资性房地产', ['投资性房地产']),
    ('固定资产', ['固定资产', '固定资产清理']),
    ('在建工程', ['在建工程']),
    ('无形资产', ['无形资产']),
    ('使用权资产', ['使用权资产']),
    ('长期待摊费用', ['长期待摊费用']),
    ('递延所得税资产', ['递延所得税资产']),
    ('其他非流动资产', ['其他非流动资产', '委托贷款', '持有至到期投资', '其他权益工具投资']),
]
BS_LIAB_ROWS = [
    ('短期借款', ['短期借款']),
    ('应付账款', ['应付账款']),
    ('应付票据', ['应付票据']),
    ('预收款项', ['预收账款']),
    ('合同负债', ['合同负债']),
    ('应付职工薪酬', ['应付职工薪酬']),
    ('应交税费', ['应交税费']),
    ('其他应付款', ['其他应付款', '应付股利', '应付利息']),
    ('一年内到期的非流动负债', ['租赁负债', '长期借款']),
    ('长期借款', ['长期借款']),
    ('应付债券', ['应付债券']),
    ('租赁负债', ['租赁负债']),
    ('长期应付款', ['长期应付款', '专项应付款']),
    ('递延收益', ['递延收益']),
    ('递延所得税负债', ['递延所得税负债']),
    ('预计负债', ['预计负债']),
]
BS_EQ_ROWS = [
    ('实收资本', ['实收资本']),
    ('资本公积', ['资本公积']),
    ('其他权益工具', ['其他权益工具']),
    ('库存股', ['库存股']),
    ('专项储备', ['专项储备']),
    ('盈余公积', ['盈余公积']),
    ('未分配利润', ['利润分配', '本年利润']),
]
PL_ROWS = [
    ('营业收入', ['主营业务收入', '其他业务收入']),
    ('营业成本', ['主营业务成本', '其他业务成本']),
    ('税金及附加', ['税金及附加']),
    ('销售费用', ['销售费用']),
    ('管理费用', ['管理费用']),
    ('研发费用', ['研发费用', '研发支出', '研究开发费']),
    ('财务费用', ['财务费用']),
    ('其他收益', ['其他收益']),
    ('投资收益', ['投资收益']),
    ('信用减值损失', ['信用减值损失']),
    ('资产减值损失', ['资产减值损失']),
    ('营业外收入', ['营业外收入']),
    ('营业外支出', ['营业外支出']),
    ('所得税费用', ['所得税费用']),
]


def _money(ws, r, c, v, fill=None, bold=False, fmt=NUMFMT):
    cell = ws.cell(r, c)
    if v is not None and abs(v) > 0.005:
        cell.value = round(v, 2)
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if bold:
        cell.font = BOLD
    cell.border = BORDER
    cell.alignment = RGT


def _txt(ws, r, c, v, fill=None, bold=False):
    cell = ws.cell(r, c, v)
    if fill:
        cell.fill = fill
    if bold:
        cell.font = BOLD
    cell.border = BORDER
    cell.alignment = LFT


def _l1(name):
    """科目一级名 = 名称首段（SAP 名称以 - 分层，含变体归一化）。"""
    return C.norm_l1(name)


def build_sap_tb(data_dir, out_path=None, year='2026', comps=None):
    """生成 SAP 自建试算表。comps=None 时用全部公司。返回文件路径。"""
    tb = SR.read_sap_tb(data_dir, None, year=year)
    # 实体列表：TB 中实际出现的公司
    all_comps = sorted({c for (c, _cd, _n, _y) in tb})
    if comps:
        comps = [c for c in all_comps if c in comps]
    else:
        comps = all_comps

    # 科目一级聚合：{公司: {一级名: {qc,jf,df,qm}}}
    ent_l1 = {}
    for c in comps:
        ent_l1[c] = {}
    for (comp, code, name, yy), v in tb.items():
        if comp not in ent_l1:
            continue
        l1 = _l1(name)
        a = ent_l1[comp].setdefault(l1, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0})
        a['qc'] += v['qc']; a['jf'] += v['jf']; a['df'] += v['df']; a['qm'] += v['qm']

    out_path = out_path or os.path.join(data_dir, f'自建试算表_{year}.xlsx')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ===== Sheet1 试算表：一级科目 × 公司（期末余额，符号口径） =====
    ws = wb.create_sheet('试算表')
    ws.append(['科目名称'] + comps + ['集团合计'])
    for c in ws[1]:
        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
    # 科目排序：资产类在前（一级名自然序）
    l1_names = set()
    for c in comps:
        l1_names.update(ent_l1[c].keys())
    for i, l1 in enumerate(sorted(l1_names), start=2):
        _txt(ws, i, 1, l1)
        tot = 0.0
        for j, c in enumerate(comps, start=2):
            v = ent_l1[c].get(l1, {}).get('qm', 0.0)
            _money(ws, i, j, v)
            tot += v
        _money(ws, i, len(comps) + 2, tot, bold=True, fill=TOT_FILL)
    # 合计行
    r = len(l1_names) + 2
    _txt(ws, r, 1, '期末合计', bold=True, fill=TOT_FILL)
    for j, c in enumerate(comps, start=2):
        tot = sum(v['qm'] for v in ent_l1[c].values())
        _money(ws, r, j, tot, bold=True, fill=TOT_FILL)
    ws.freeze_panes = 'B2'
    ws.column_dimensions['A'].width = 24

    # ===== Sheet2 资产负债表：资产/负债/权益 =====
    ws2 = wb.create_sheet('资产负债表')
    _txt(ws2, 1, 1, f'自建资产负债表（{year}，核对基准=企业报表）', bold=True)
    _txt(ws2, 1, 2, '编制单位：SAP 集团', bold=True)
    rows = [('资 产', BS_ASSET_ROWS, 1), ('负 债', BS_LIAB_ROWS, -1), ('所有者权益', BS_EQ_ROWS, -1)]
    r = 3
    for sec, cfgs, sgn in rows:
        _txt(ws2, r, 1, sec, bold=True, fill=HEAD_FILL); r += 1
        for rn, kws in cfgs:
            _txt(ws2, r, 1, rn)
            for j, c in enumerate(comps):
                v = sum(a['qm'] for l1, a in ent_l1[c].items()
                        if any(k in l1 for k in kws))
                # ⚡ 2026-08-09 修复：负债/权益贷余科目转正显示（铁律54；原 _sgn 未生效）
                _money(ws2, r, 2 + j, v * sgn)
            r += 1
        r += 1
    # 勾稽
    _txt(ws2, r, 1, '勾稽说明：以企业资产负债表为核对基准', bold=True)

    # ===== Sheet3 利润表：损益发生额 =====
    ws3 = wb.create_sheet('利润表')
    _txt(ws3, 1, 1, f'自建利润表（{year}，发生额口径）', bold=True)
    hdr = ['项目'] + comps + ['集团合计']
    for j, h in enumerate(hdr, 1):
        _txt(ws3, 2, j, h, fill=HEAD_FILL, bold=True)
    r = 3
    for rn, kws in PL_ROWS:
        _txt(ws3, r, 1, rn)
        tot = 0.0
        # ⚡⚡ 2026-08-30 修复：收益类科目取贷方 df（原仅营业收入取 df，投资收益等
        #   错取借方 jf → 334M 显示 1.96M，利润表投资收益严重失真）。
        is_income = rn in ('营业收入', '其他收益', '投资收益', '营业外收入')
        for j, c in enumerate(comps, 2):
            v = sum((a['df'] if is_income else a['jf']) for l1, a in ent_l1[c].items()
                    if any(k in l1 for k in kws))
            _money(ws3, r, j, v)
            tot += v
        _money(ws3, r, len(comps) + 2, tot, bold=True, fill=TOT_FILL)
        r += 1

    wb.save(out_path)
    print(f'✅ 自建试算表已生成: {out_path}')
    print(f'   公司数: {len(comps)}  一级科目数: {len(l1_names)}')
    return out_path


def main(argv=None):
    data_dir = sys.argv[1] if len(sys.argv) > 1 else P.YY
    comps = sys.argv[2].split(',') if len(sys.argv) > 2 and sys.argv[2] else None
    build_sap_tb(data_dir, comps=comps)


if __name__ == '__main__':
    main()
