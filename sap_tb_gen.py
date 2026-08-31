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
import sap_adapter as A   # ⚡⚡ 2026-08-30 支持 U8（XBJ）回退 + 年份推断
import sap_report as RPT  # ⚡⚡ 2026-08-31 企业报表（利润表/资产负债表 .xls）读取


def _infer_year(data):
    """从数据目录推断年份（同 run_u8_on_sap）。消除 year='2026' 硬编码。"""
    m = re.search(r'[\\/]数据[\\/]?(\d{4})', data)
    if m:
        return m.group(1)
    m = re.findall(r'(\d{4})', data)
    return m[-1] if m else '2026'

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
    ('研发费用', ['研发费用', '研究开发费']),
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


def _rev_cost_agg(comp, tb, year, prefix, is_income, mirror_mode=False):
    """收入/成本 TB 聚合（⚡⚡ 2026-08-31 修复 AH 收入/成本双计与冲减根因）：
    SAP 损益科目借贷双方均有发生（父级行 df、GL 流水行 df+jf、其他业务收入 jf 退回等）
    → 真实账套取【净额】：收入 df−jf、成本 jf−df（=企业报表口径）。
    ⚡⚡ 2026-09-01 镜像账套（XBJ/AZ 损益科目借贷同额复制）→ 取发生额（收入 df、成本 jf，
    =底稿口径），否则 jf=df 同额行净额=0 失真。镜像父级（父=子和）排除防双计。"""
    tot = 0.0
    rows = [(str(cd_), nm_, v_) for (cc_, cd_, nm_, yy_), v_ in tb.items()
            if cc_ == comp and str(yy_) == str(year) and str(cd_).startswith(prefix)]
    # ⚡⚡ 2026-09-01 镜像父级排除（AZ 6001 父级=子和同额 → 父+子双计 2×）：有子级
    #   且金额=子级之和的父级行跳过，只计末级叶子（与 revenue._tb_total 一致）。
    _codes = [r[0] for r in rows]
    _excl = set()
    for _c in _codes:
        _kids = [_c2 for _c2 in _codes if _c2 != _c and _c2.startswith(_c)]
        if not _kids:
            continue
        _s = sum(float(r[2].get('df') or 0.0) for r in rows if r[0] in _kids)
        _pv = next((float(r[2].get('df') or 0.0) for r in rows if r[0] == _c), 0.0)
        if abs(abs(_pv) - abs(_s)) < 1.0:
            _excl.add(_c)
    for cd_, nm_, v_ in rows:
        if cd_ in _excl:
            continue
        _d = float(v_.get('df') or 0.0)
        _j = float(v_.get('jf') or 0.0)
        if is_income:
            tot += _d if mirror_mode else (_d - _j)
        else:
            tot += _j if mirror_mode else (_j - _d)
    return tot


def build_sap_tb(data_dir, out_path=None, year=None, comps=None):
    """生成 SAP 自建试算表。comps=None 时用全部公司。返回文件路径。
    ⚡⚡ 2026-08-30：改走 sap_adapter.read_tb_full（SAP/U8 双支持 + 年份自动推断）。"""
    if year is None:
        year = _infer_year(data_dir)
    A._DATA_ROOT = data_dir
    _tb_all = A.read_tb_full(data_dir, None)
    # 过滤目标年份（键 y 可能是 int/str）
    tb = {(c, cd, n, y): v for (c, cd, n, y), v in _tb_all.items()
          if str(y) == str(year)}
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
    # ⚡⚡ 2026-08-31 修复（AZ 泰国账套损益双计 0.5× 根因）：U8 科目余额表损益科目含
    #   【父级行】（6801 所得税费用）+【末级行】（6801.01 当期所得税费用）同额 →
    #   ent_l1 按一级名聚合时父+子双计（上分所得税 60,058=2×30,029）。
    #   精准剔除：仅当【父级金额 = 子级之和】（qc/jf/df/qm 全相等=镜像复制行）
    #   且【有子级名称含父级一级名】（同科目父子，如 当期所得税费用⊇所得税费用；
    #   泰国 6602 管理费用/职工薪酬 子名不含"管理费用"→ 保留父级，防 480K 丢失）。
    #   仅对损益科目（匹配 PL_ROWS kws）；资产负债表科目不动（与底稿同源口径一致）。
    _pl_kws = {kw for _r in PL_ROWS for kw in _r[1]}
    # ⚡⚡ 2026-08-31 AZ 跨期调整兜底：损益科目名称首段不含标准名但码同族
    #   （AZ 6001.02『跨期收入调整』jf=df=-18.5M，一级名≠主营业务收入）→ 名称
    #   不匹配 PL 关键词时按码前缀归入对应损益一级名（与底稿 read_km 父级口径一致）。
    _pl_pfx = {'6001': '主营业务收入', '6051': '其他业务收入',
               '6401': '主营业务成本', '6402': '其他业务成本'}
    _by_comp = {}
    for (comp, code, _n, _y), _v in tb.items():
        # ⚡⚡ read_tb_full 的 name 在 key（非 value）→ 需随行保存，_l1 判定依赖
        _by_comp.setdefault(comp, {})[str(code)] = (_v, str(_n))
    _skip = set()
    for comp, _cds in _by_comp.items():
        for c1 in _cds:
            _v1, _n1 = _cds[c1]
            if _l1(_n1) not in _pl_kws:
                continue
            _kids = [cc for cc in _cds if cc != c1 and cc.startswith(c1)]
            if not _kids:
                continue
            _ksum = {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0}
            for cc in _kids:
                for _k in _ksum:
                    _ksum[_k] += float(_cds[cc][0].get(_k) or 0.0)
            _pv = _v1
            _mirror = all(abs(float(_pv.get(k) or 0.0) - _ksum[k]) < 0.005 for k in _ksum)
            if not _mirror:
                continue
            _pl1 = _l1(_n1)
            if any(_pl1 in _cds[cc][1] for cc in _kids):
                _skip.add((comp, c1))
    # 族一级行保留集：某主体某 pfx 族有【未剔除的 level==1 行】→ 父级已在 ent_l1，
    #   子级（名称不同，如 6051.01 光亮铜）不再 _pl_pfx 兜底归入（防父+子双计）。
    _pl_l1_kept = set()
    for (comp, code, name, yy), v in tb.items():
        if (str(code)[:4] in _pl_pfx and v.get('level') == 1
                and (comp, str(code)) not in _skip):
            _pl_l1_kept.add((comp, str(code)[:4]))
    # ⚡⚡ 2026-09-01 修复（AZ 试算表应收=净额 vs 底稿=余额 口径不一致根因）：备抵科目
    #   1231 坏账准备的子级名撞被备抵科目名（AZ 1231.01 子级名『应收账款』→ 名称聚合
    #   误归入应收账款 → 试算表应收=1122-坏账(净额 376.5M) vs 底稿应收=1122(余额
    #   397.2M) 差 20.7M）。处理：①坏账准备族（1231 前缀）子级强制归位『坏账准备』；
    #   ②资产负债表镜像父级（父级=子级之和）只保留父级、剔除子级——否则父级+子级
    #   同归一个一级名会双计（AZ 1231 父级 -20.8M + 子级 -20.8M = -41.6M）。
    _mirror_kids = set()   # 镜像父级的子级（父级已代表全族）→ 聚合时跳过
    for comp, _cds in _by_comp.items():
        for c1 in _cds:
            _v1, _n1 = _cds[c1]
            if _l1(_n1) in _pl_kws:
                continue   # 损益镜像父级已走 _skip（剔父留叶）
            _kids = [cc for cc in _cds if cc != c1 and cc.startswith(c1)]
            if not _kids:
                continue
            _ksum = {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0}
            for cc in _kids:
                for _k in _ksum:
                    _ksum[_k] += float(_cds[cc][0].get(_k) or 0.0)
            if all(abs(float(_v1.get(k) or 0.0) - _ksum[k]) < 0.005 for k in _ksum):
                for cc in _kids:
                    _mirror_kids.add((comp, cc))
    # ⚡⚡ 2026-09-01 AZ 外币主体（泰国/新加坡）一级名映射：折算试算表行只有明细
    #   （1002.01 中行/外币/开泰银行，无 level==1 父级），名称非标准一级名 → 聚合后
    #   散落成『中行』『外币』等独立行（泰国银行存款/管理费用 TB=0）。修复：从
    #   level==1 行学 code4→一级名（1002→银行存款、6602→管理费用…），明细行名称
    #   无法识别为已知一级名时按 code 前4位映射。
    _code4_l1 = {}
    _l1_known = set()
    _has_l1 = set()
    for (_cc, _cd, _nn, _yy), _vv in tb.items():
        if _vv.get('level') == 1 and _nn:
            _has_l1.add(_cc)
            _nm1 = _l1(str(_nn))
            _l1_known.add(_nm1)
            _c4 = str(_cd)[:4]
            if _c4 and _c4 not in _code4_l1:
                _code4_l1[_c4] = _nm1
    # 外币主体（泰国/新加坡）：折算试算表全明细行无 level==1 → 该主体无一级父行
    _foreign_comps = {c for c in ent_l1 if c not in _has_l1}
    for (comp, code, name, yy), v in tb.items():
        if comp not in ent_l1:
            continue
        if (comp, str(code)) in _skip:
            continue   # 镜像父级（父=子和 且 子名含父级一级名）→ 剔除防双计
        if (comp, str(code)) in _mirror_kids:
            continue   # 资产负债表镜像父级的子级 → 父级已代表全族
        l1 = _l1(name)
        _c4 = str(code)[:4]
        if (comp in _foreign_comps and l1 not in _l1_known
                and _c4 in _code4_l1 and l1 != _code4_l1[_c4]):
            l1 = _code4_l1[_c4]   # 外币主体明细名（中行/外币/应付薪金…）→ code4 一级名
        if str(code)[:4] == '1231':
            l1 = '坏账准备'   # 备抵科目子级名撞被备抵科目名 → 归位坏账准备
        if l1 not in _pl_kws and str(code)[:4] in _pl_pfx:
            _pfx4 = str(code)[:4]
            if (comp, _pfx4) not in _pl_l1_kept:
                l1 = _pl_pfx[_pfx4]   # 跨期调整等码同族科目归入损益一级名
        a = ent_l1[comp].setdefault(l1, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0})
        a['qc'] += v['qc']; a['jf'] += v['jf']; a['df'] += v['df']; a['qm'] += v['qm']

    # ⚡⚡ 2026-09-01 账套损益口径模式：镜像账套（损益科目借贷同额复制，XBJ/AZ）→
    #   损益取【发生额】（收入 df、成本 jf，=底稿口径）；真实账套（损益借贷不对称，
    #   AH 信用减值 jf94.9M/df7.7M）→ 取【净额】（=企业报表口径）。AH 个别同额行
    #   （630101 处置利得 jf=df=1.16M 结转镜像）净额=0 正确（企业报表营业外收入 10,374）。
    _pl_probe = []
    for (comp, code, name, yy), v in tb.items():
        if str(code)[:1] in ('6', '7', '8'):
            _d = float(v.get('df') or 0.0)
            _j = float(v.get('jf') or 0.0)
            if _d or _j:
                _m = max(abs(_d), abs(_j))
                _pl_probe.append(abs(_d - _j) < 0.005 or (_m > 0 and abs(_d - _j) / _m < 0.02))
    _mirror_mode = bool(_pl_probe) and sum(_pl_probe) / len(_pl_probe) > 0.8
    if _mirror_mode:
        print(f'  ⚡ 损益口径=镜像账套（取发生额）')

    out_path = out_path or os.path.join(data_dir, f'自建试算表_{year}.xlsx')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ===== Sheet1 试算表：一级科目 × 公司（期末余额，符号口径） =====
    ws = wb.create_sheet('试算表')
    ws.append(['科目名称'] + comps + ['集团合计'])
    for c in ws[1]:
        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
    # 科目排序：资产→负债→权益→损益 报表顺序（用户方法论，audit_common.sort_report_names）
    l1_names = set()
    for c in comps:
        l1_names.update(ent_l1[c].keys())
    try:
        from audit_common import sort_report_names as _srn
        _ordered = _srn(list(l1_names))
    except Exception:
        _ordered = sorted(l1_names)
    for i, l1 in enumerate(_ordered, start=2):
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
    # ⚡⚡ 2026-08-31 铁律：试算表必须与 TB 同源（底稿与 TB 核对一致）→ 不用企业报表
    #   覆盖。Sheet2 资产负债表 = ent_l1 一级名聚合（TB 口径）。
    r = 3
    for sec, cfgs, sgn in rows:
        _txt(ws2, r, 1, sec, bold=True, fill=HEAD_FILL); r += 1
        for rn, kws in cfgs:
            _txt(ws2, r, 1, rn)
            for j, c in enumerate(comps):
                v = sum(a['qm'] for l1, a in ent_l1[c].items()
                        if any(k in l1 for k in kws))
                # ⚡ 2026-08-09 修复：负债/权益贷余科目转正显示（铁律54；原 _sgn 未生效）
                v = v * sgn
                _money(ws2, r, 2 + j, v)
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
    # ⚡⚡ 2026-08-31 铁律：试算表与 TB 同源（底稿与 TB 核对一致）→ 不直接读企业报表。
    #   管理费用/销售费用/研发费用/制造费用 在 SAP 6600 总池无独立科目 → 读账套
    #   费用目录（6600 池功能范围拆分，=TB 拆分形态）→ 与底稿（expense_detail 同源）一致。
    _fee_scope = {}
    try:
        _fs_cats = ('管理费用', '销售费用', '研发费用', '制造费用')
        _has_rep = RPT.has_enterprise_reports(data_dir)
        for c in comps:
            _fs = RPT.read_fee_scope(data_dir, c)
            if _has_rep:
                # ⚡⚡ 2026-08-31 与底稿(expense_detail)同源：费用目录缺某类别/全缺的
                #   主体回退企业利润表（=TB 6600 池功能范围拆分，费用目录未导出部分）
                _pl = RPT.read_ent_profit(data_dir, c)
                for _cat in ('管理费用', '销售费用', '研发费用'):
                    if (_cat not in _fs
                            and abs(float(_pl.get(_cat, 0.0))) > 0.005):
                        _fs.setdefault(_cat, [(f'{_cat}（企业报表）', float(_pl.get(_cat, 0.0)))])
            if _fs:
                _fee_scope[c] = {k: sum(v for _, v in items) for k, items in _fs.items()}
    except Exception:
        pass
    r = 3
    for rn, kws in PL_ROWS:
        _txt(ws3, r, 1, rn)
        tot = 0.0
        # ⚡⚡ 2026-08-30 修复：收益类科目取贷方 df（原仅营业收入取 df，投资收益等
        #   错取借方 jf → 334M 显示 1.96M，利润表投资收益严重失真）。
        is_income = rn in ('营业收入', '其他收益', '投资收益', '营业外收入')
        for j, c in enumerate(comps, 2):
            if rn == '营业收入':
                # ⚡⚡ 2026-08-31 GL 流水行取净额（AH 父级+GL 双计修复，与底稿/TB 一致）
                v = (_rev_cost_agg(c, tb, year, '6001', True, _mirror_mode)
                     + _rev_cost_agg(c, tb, year, '6051', True, _mirror_mode))
            elif rn == '营业成本':
                v = (_rev_cost_agg(c, tb, year, '6401', False, _mirror_mode)
                     + _rev_cost_agg(c, tb, year, '6402', False, _mirror_mode))
            elif rn in _fs_cats and c in _fee_scope and rn in _fee_scope[c]:
                v = _fee_scope[c][rn]   # 费用目录（TB 6600 池功能范围拆分）
            elif rn == '财务费用':
                # ⚡⚡ 2026-08-31 与底稿一致（expense_detail tb_l2_control 净额）：
                #   财务费用取 TB 净额（利息收入等贷方冲减），非借发。
                #   ⚡⚡ 2026-09-01 镜像账套取借发（XBJ/AZ），真实账套取净额（AH）
                _t = 0.0
                for l1, a in ent_l1[c].items():
                    if any(k in l1 for k in kws):
                        _d = a['df']; _j = a['jf']
                        _t += _j if _mirror_mode else (_j - _d)
                v = _t
            else:
                # ⚡⚡ 2026-08-31 损益科目取净额（AH 信用减值/资产减值/其他收益/
                #   投资收益等借贷双方冲减 → 净额=企业报表口径）；镜像账套（XBJ/AZ）
                #   取发生额（is_income→df、损失→jf，=底稿口径）
                if is_income:
                    v = sum((a['df'] if _mirror_mode else (a['df'] - a['jf']))
                            for l1, a in ent_l1[c].items() if any(k in l1 for k in kws))
                else:
                    v = sum((a['jf'] if _mirror_mode else (a['jf'] - a['df']))
                            for l1, a in ent_l1[c].items() if any(k in l1 for k in kws))
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
