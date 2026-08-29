# -*- coding: utf-8 -*-
"""subjects_registry.py —— 科目注册表（2026-08-02 架构铺垫 P0-①）。

集中登记全部报表科目（即 13 个 builder 产出的每张底稿）的统一元数据，
作为未来大动作的共同前置：
  1. 调整分录台账：按科目生成台账、回读过入审定表调整列；
  2. 期初/上年数两种过入方式：audit_py（期初未审 vs 上年审定核对→过入上年调整）
     / audited（直接采用上年审定数，利润表类）；
  3. 全科目附注汇总：按科目枚举、抽取审定数；
  4. gp_other 收敛分发、业务类型参数化（存货）等。

设计原则：
  · 本表只登记【元数据】（科目是谁、怎么取数、用什么模板），不含取数逻辑；
  · 取值与 13 个 builder 内现有定义保持一致（铁律：以实际账套科目名为准，
    代码只作结构定位）；
  · 新科目只需加一行注册，附注/审定表/台账自动接入，不改 builder 代码。
"""
from __future__ import annotations

# ---------- 附注模板类型 ----------
# 与 audit_common.build_footnote_generic 的 mode 对齐：
#   'money' = 期末数(未审)/审计调整/期末审定数/期初数（货币资金专用）
#   'four'  = 期初数/本期增加/本期减少/期末数（余额增减变动类）
#   'two'   = 期末数/期初数（两期数对比类）
#   'asset3d' = 固定资产三维模板（未审/调整/审定 块 × 原值/折旧/减值/净值 组件 × 段）
#   'none'  = 暂未生成附注汇总（往来/存货/收入 暂缓类）
FOOTNOTE_NONE = 'none'
FOOTNOTE_MONEY = 'money'
FOOTNOTE_FOUR = 'four'
FOOTNOTE_TWO = 'two'
FOOTNOTE_3D = 'asset3d'

# ---------- 期初/上年数过入方式（用户方法论 2026-08-02） ----------
#   'audit_py' = 方式1：账面期初未审数与上年审定数核对，有差异→过入上年调整数
#                （BS 资产/负债类，如长期资产，期初必须带调整）
#   'audited'  = 方式2：直接采用上年审定数，不再追溯上年未审与审定差异
#                （PL 利润表类，上年数为"比较数"性质）
OPENING_AUDIT_PY = 'audit_py'
OPENING_AUDITED = 'audited'

# ---------- 调整列结构（用户方法论 2026-08-02：调整数必须结构化拆分） ----------
#   'simple'  = 单列"审计调整数"（现状，调整数留空待填）
#   'split'   = 拆 借方调整/贷方调整/其中：重分类调整（资产类披露增减变动时启用）
ADJ_SIMPLE = 'simple'
ADJ_SPLIT = 'split'


def _s(name, codes, *, is_credit=False, builder=None, footnote=FOOTNOTE_NONE,
       opening=OPENING_AUDIT_PY, adj=ADJ_SIMPLE, file_key=None, note=''):
    """构造一条科目注册记录。
    file_key：底稿文件名定位键（报表科目名 ≠ 实际文件名时指定，如 货币资金→银行存款、
              应付职工薪酬→职工薪酬；默认=name）。未来大动作（台账/汇总/Word）靠它定位文件。
    """
    return dict(
        name=name,                # 报表科目名（附注/审定表标题用）
        codes=[str(c) for c in codes],   # 取数科目代码（4 位主码；名称匹配仍以实际账套为准）
        is_credit=is_credit,      # 负债/权益类=贷方科目 True，资产/费用类=False
        builder=builder,          # 所属 builder 模块名（不含 .py）
        footnote=footnote,        # 附注模板类型
        opening=opening,          # 期初/上年数过入方式
        adj=adj,                  # 调整列结构
        file_key=file_key or name,   # 底稿文件名定位键
        note=note,                # 备注（暂缓原因、特殊披露等）
    )


# ==================== 科目注册表（登记顺序≈报表顺序） ====================
# 注：取值与各 builder 内现有定义一致；"暂缓附注"科目（往来/存货/收入）footnote='none'。
REGISTRY = {
    # ---- 货币资金（bank_deposit）----
    'cash': _s('货币资金', ['1001', '1002', '1012'], builder='bank_deposit_detail',
               footnote=FOOTNOTE_MONEY, opening=OPENING_AUDIT_PY, file_key='银行存款',
               note='三科目：库存现金/银行存款/其他货币资金；money 模板（未审/调整/审定/期初）；'
                    '底稿文件名=银行存款审计底稿'),

    # ---- 往来类（current_account_detail）—— 附注暂缓 ----
    'ar': _s('应收账款', ['1122'], builder='current_account_detail',
             footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
             note='往来类，附注暂缓（2026-08-02 用户指示先放下）；期初需带上年调整'),
    'ap': _s('应付账款', ['2202'], is_credit=True, builder='current_account_detail',
             footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
             note='往来类，附注暂缓；贷方科目'),
    'ar_other': _s('其他应收款', ['1221'], builder='current_account_detail',
                   footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
                   note='往来类，附注暂缓'),
    'ap_other': _s('其他应付款', ['2241'], is_credit=True, builder='current_account_detail',
                   footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
                   note='往来类，附注暂缓；贷方科目'),
    'note_recv': _s('应收票据', ['1121', '1021'], builder='current_account_detail',
                    footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
                    note='往来类，附注暂缓；2026-08-07 补 Q=1021'),
    'prepay': _s('预付账款', ['1123'], builder='current_account_detail',
                 footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
                 note='往来类，附注暂缓；2026-08-06 补登记'),
    'advance_recv': _s('预收账款', ['2203'], is_credit=True, builder='current_account_detail',
                       footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
                       note='往来类，附注暂缓；2026-08-06 补登记'),
    'contract_liab': _s('合同负债', ['2205', '2207', '2204'], is_credit=True, builder='current_account_detail',
                        footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
                        note='往来类，附注暂缓；FY=2205、JTt=2203（名称=合同负债\预收工程款，按名匹配）、T=2204、JTt/SSS=2207；2026-08-06 补登记，08-07 补 2204（2203 撤——H 2203=预收账款，靠名称区分）'),
    'contract_asset': _s('合同资产', ['1480', '1462', '1125'], builder='current_account_detail',
                         footnote=FOOTNOTE_NONE, opening=OPENING_AUDIT_PY,
                         note='往来类，附注暂缓；SSS=1480 建筑施工企业大科目、JTt=1125、Q=1462；2026-08-06 前瞻登记，08-07 补 1462/1125'),

    # ---- 存货（inventory_detail）—— 附注暂缓；业务类型参数化待办 ----
    'inventory': _s('存货', ['1401', '1402', '1403', '1404', '1405', '1406', '1407', '1408',
                              '1411', '1412', '1413', '1414', '1415'],
                    builder='inventory_detail', footnote=FOOTNOTE_NONE,
                    opening=OPENING_AUDIT_PY,
                    note='附注暂缓；按业务类型（制造/贸易/施工）参数化待办；'
                         '2026-08-06 codes 补 1404/1412/1413/1414/1415（材料成本差异/包装物/低值易耗品/在产品/数据资源）'),

    # ---- 长期资产（longterm_assets_detail）—— 固定资产三维模板 ----
    'fa': _s('固定资产', ['1601', '1602', '1603'], builder='longterm_assets_detail',
             footnote=FOOTNOTE_3D, opening=OPENING_AUDIT_PY, adj=ADJ_SPLIT,
             note='固定资产三维模板（含累计折旧/减值/账面价值组件）；调整列拆分（借/贷/重分类）'),
    'cip': _s('在建工程', ['1604'], builder='longterm_assets_detail',
              footnote=FOOTNOTE_3D, opening=OPENING_AUDIT_PY, adj=ADJ_SPLIT,
              note='三维模板；账面余额+转入固定资产字段扩展'),
    'ia': _s('无形资产', ['1701', '1702', '1703'], builder='longterm_assets_detail',
             footnote=FOOTNOTE_3D, opening=OPENING_AUDIT_PY, adj=ADJ_SPLIT,
             note='三维模板（含累计摊销/减值/净值）'),
    'lta': _s('长期待摊费用', ['1801'], builder='longterm_assets_detail',
              footnote=FOOTNOTE_3D, opening=OPENING_AUDIT_PY, adj=ADJ_SPLIT,
              note='三维模板（长期待摊摊销）'),
    'rua': _s('使用权资产', ['1704', '1705', '1681', '1509', '1510', '1810', '1812'], builder='longterm_assets_detail',
              footnote=FOOTNOTE_3D, opening=OPENING_AUDIT_PY, adj=ADJ_SPLIT,
              note='三维模板；2026-08-06 codes 修正：FY=1704/1705、JTt单体=1509/1510、JTt合并/SSS=1681、旧码 1810/1812；原含 1811 与递延所得税资产(dta)串科目，已剔除'),
    'ipr': _s('投资性房地产', ['1521', '1522', '1523'], builder='longterm_assets_detail',
              footnote=FOOTNOTE_3D, opening=OPENING_AUDIT_PY, adj=ADJ_SPLIT,
              note='三维模板；2026-08-06 补登记（JTt ga 有数据）'),

    # ---- 借款（loan_detail）—— 2026-08-06 拆两个独立 file_key（实际输出两个文件） ----
    'st_loan': _s('短期借款', ['2001'], is_credit=True, builder='loan_detail',
                  footnote=FOOTNOTE_TWO, opening=OPENING_AUDIT_PY,
                  note='短期借款审计底稿（独立文件）；两期数模板'),
    'lt_loan': _s('长期借款', ['2501'], is_credit=True, builder='loan_detail',
                  footnote=FOOTNOTE_TWO, opening=OPENING_AUDIT_PY,
                  note='长期借款审计底稿（独立文件）；两期数模板'),

    # ---- 权益（equity_detail）----
    'paid_in': _s('实收资本(或股本)', ['4001'], is_credit=True, builder='equity_detail',
                  footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
                  note='four 模板（数据驱动裁剪：无增资自动隐藏本期增加）'),
    'capital_res': _s('资本公积', ['4002'], is_credit=True, builder='equity_detail',
                      footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY),
    'surplus_res': _s('盈余公积', ['4101'], is_credit=True, builder='equity_detail',
                      footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY),
    'retained': _s('未分配利润', ['4104'], is_credit=True, builder='equity_detail',
                   footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
                   note='2026-08-06 codes 修正：FY=4104.12/JTt=410409（4104 前缀）；原 4103 是"本年利润"（下挂损益科目）会被误取，已剔除'),
    'reserve_fund': _s('专项储备', ['4004', '4102', '3100', '4301'], is_credit=True, builder='equity_detail',
                       footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
                       note='2026-08-06 补登记：JTt=4004/4301；部分账套 4102 或挂在盈余公积/资本公积下（按名匹配）；2026-08-07 H=3100'),

    # ---- 收入（revenue_detail）—— 附注暂缓 ----
    'revenue': _s('营业收入', ['6001', '6051', '6401', '6402'], builder='revenue_detail',
                  footnote=FOOTNOTE_NONE, opening=OPENING_AUDITED,
                  note='收入类，附注暂缓；分部门明细表=附注模板；期初用方式2（直接上年审定数）'),

    # ---- 费用（expense_detail）----
    'sga': _s('销售费用', ['6601'], builder='expense_detail',
              footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
              note='费用类；期初用方式2'),
    'ga': _s('管理费用', ['6602'], builder='expense_detail',
             footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
             note='费用类；期初用方式2'),
    'finance_exp': _s('财务费用', ['6603'], builder='expense_detail',
                      footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                      note='费用类；期初用方式2'),

    # ---- 利润表（pl_detail）—— 2026-08-06 补登记全部损益科目 ----
    'income_tax': _s('所得税费用', ['6801'], builder='pl_detail',
                     footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                     note='利润表科目；期初用方式2'),
    'tax_surcharge': _s('税金及附加', ['6071', '6403'], builder='pl_detail',
                        footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                        note='FY=6403(营业务税金及附加)、JTt/SSS=6071；2026-08-06 补登记'),
    'invest_income': _s('投资收益', ['6111'], builder='pl_detail',
                        footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                        note='利润表科目；期初用方式2；2026-08-06 补登记'),
    'other_income': _s('其他收益', ['6113'], builder='pl_detail',
                       footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                       note='FY=6113；JTt=6115 特例已入配置（6115 标准属资产处置损益，防撞车）'),
    'nonop_income': _s('营业外收入', ['6301'], builder='pl_detail',
                       footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                       note='利润表科目；2026-08-06 补登记'),
    'nonop_expense': _s('营业外支出', ['6711'], builder='pl_detail',
                        footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                        note='利润表科目；2026-08-06 补登记'),
    'asset_impair': _s('资产减值损失', ['6701'], builder='pl_detail',
                       footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                       note='利润表科目；2026-08-06 补登记'),
    'credit_impair': _s('信用减值损失', ['6702'], builder='pl_detail',
                        footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                        note='利润表科目；2026-08-06 补登记'),
    'asset_disposal': _s('资产处置损益', ['6115'], builder='pl_detail',
                         footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                         note='利润表科目；2026-08-06 补登记'),
    'fair_value': _s('公允价值变动损益', ['6101'], builder='pl_detail',
                     footnote=FOOTNOTE_FOUR, opening=OPENING_AUDITED,
                     note='利润表科目；FY 无数据跳过；2026-08-06 补登记'),

    # ---- 税金（tax_detail）----
    'tax_payable': _s('应交税费', ['2221'], is_credit=True, builder='tax_detail',
                      footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
                      note='按税种展开；负债类期初带调整'),

    # ---- 薪酬（payroll_detail）----
    'payroll': _s('应付职工薪酬', ['2211'], is_credit=True, builder='payroll_detail',
                  footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY, file_key='职工薪酬',
                  note='计提/列支核对+多借多贷核对表；底稿文件名=职工薪酬审计底稿'),

    # ---- 研发支出（rd_expense_detail）----
    'rd': _s('研发支出', ['5301'], builder='rd_expense_detail',
             footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
             note='资本化/费用化；无研发支出科目时如实披露'),

    # ---- 其他类（gp_other）----
    'tfa': _s('交易性金融资产', ['1101', '1003'], builder='gp_other',
              footnote=FOOTNOTE_TWO, opening=OPENING_AUDIT_PY,
              note='2026-08-07 补 H=1003'),
    'lti': _s('长期股权投资', ['1511'], builder='gp_other',
              footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY, adj=ADJ_SPLIT,
              note='按被投资单位展开；调整列拆分'),
    'gw': _s('商誉', ['1711'], builder='gp_other',
             footnote=FOOTNOTE_TWO, opening=OPENING_AUDIT_PY),
    'apn': _s('应付票据', ['2201'], is_credit=True, builder='gp_other',
              footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY),
    'll': _s('租赁负债', ['2601', '2631', '2632', '2271', '2802'], is_credit=True, builder='gp_other',
             footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
             note='2026-08-07 v2：2703 移出（S 一年内到期=2703 撞车，已入配置）；Q=2271、T=2802 在配置'),
    'ncl': _s('一年内到期的非流动负债', ['2242', '2485', '2281'], is_credit=True, builder='gp_other',
              footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY, file_key='一年内到期的非流动负债',
              note='2026-08-07 v2：2401（Q 特例）/2703（S 特例）移出并入配置'),
    'di': _s('递延收益', ['2401'], is_credit=True, builder='gp_other',
             footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
             note='2026-08-07 v2：2601 移出（dq 特例=递延收益 2601，入配置；FY 2601=租赁负债）'),
    'sp': _s('专项应付款', ['2711'], is_credit=True, builder='gp_other',
             footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY),
    'ep': _s('预计负债', ['2801', '2301'], is_credit=True, builder='gp_other',
             footnote=FOOTNOTE_FOUR, opening=OPENING_AUDIT_PY,
             note='2026-08-07 补 H=2301'),
    'dtl': _s('递延所得税负债', ['2901'], is_credit=True, builder='gp_other',
              footnote=FOOTNOTE_TWO, opening=OPENING_AUDIT_PY),
    'dta': _s('递延所得税资产', ['1811', '1901'], builder='gp_other',
              footnote=FOOTNOTE_TWO, opening=OPENING_AUDIT_PY,
              note='2026-08-07 补 dq=1901'),
}


# ==================== 查询与自检 ====================

def all_subjects():
    """全部注册科目（按登记顺序）。"""
    return list(REGISTRY.values())


def by_name(name):
    """按报表科目名查找（精确匹配；返回 None 未注册）。"""
    for v in REGISTRY.values():
        if v['name'] == name:
            return v
    return None


def by_builder(builder):
    """按所属 builder 模块名分组。"""
    return [v for v in REGISTRY.values() if v['builder'] == builder]


def footnote_pending():
    """附注暂缓科目（footnote='none'）。"""
    return [v for v in REGISTRY.values() if v['footnote'] == FOOTNOTE_NONE]


def opening_audited():
    """期初用方式2（直接上年审定数）的科目——利润表类。"""
    return [v for v in REGISTRY.values() if v['opening'] == OPENING_AUDITED]


def file_key(name_or_key):
    """报表科目名或 file_key → 底稿文件名定位键（无 '-' 分隔；未来定位 *_生成.xlsx 用）。
    例：file_key('货币资金') = '银行存款'；file_key('银行存款') = '银行存款'。
    """
    for v in REGISTRY.values():
        if v['name'] == name_or_key or v['file_key'] == name_or_key:
            return v['file_key']
    return None


def find_outputs(data_dir, year, name_or_key):
    """按科目名/file_key 定位某年份底稿文件（未生成返回 None）。
    返回绝对路径：<data_dir>/<file_key>审计底稿_<year>_生成.xlsx
    """
    import os
    fk = file_key(name_or_key)
    if not fk:
        return None
    p = os.path.join(data_dir, f'{fk}审计底稿_{year}_生成.xlsx')
    return p if os.path.exists(p) else None


def self_check():
    """注册表自检：键唯一、科目名唯一、字段合法。返回 (ok, 问题列表)。"""
    problems = []
    names = {}
    fkeys = {}
    for k, v in REGISTRY.items():
        if v['name'] in names:
            problems.append(f'科目名重复：{v["name"]}（{names[v["name"]]} / {k}）')
        names[v['name']] = k
        if v['file_key'] in fkeys:
            problems.append(f'file_key 重复：{v["file_key"]}（{fkeys[v["file_key"]]} / {k}）')
        fkeys[v['file_key']] = k
        if not v['codes']:
            problems.append(f'{k}：无取数代码')
        if v['footnote'] not in (FOOTNOTE_NONE, FOOTNOTE_MONEY, FOOTNOTE_FOUR,
                                 FOOTNOTE_TWO, FOOTNOTE_3D):
            problems.append(f'{k}：未知附注模板 {v["footnote"]}')
        if v['opening'] not in (OPENING_AUDIT_PY, OPENING_AUDITED):
            problems.append(f'{k}：未知期初过入方式 {v["opening"]}')
        if v['adj'] not in (ADJ_SIMPLE, ADJ_SPLIT):
            problems.append(f'{k}：未知调整列结构 {v["adj"]}')
    return (not problems, problems)


if __name__ == '__main__':
    ok, probs = self_check()
    print(f'科目注册表自检：{"✓ 通过" if ok else "✗ 有问题"}')
    for p in probs:
        print('  ⚠️', p)
    print()
    print(f'已注册科目数：{len(REGISTRY)}')
    from collections import Counter
    byb = Counter(v['builder'] for v in REGISTRY.values())
    for b, n in sorted(byb.items(), key=lambda x: -x[1]):
        print(f'  {b}: {n} 个科目')
    print()
    print('附注暂缓科目：', [v['name'] for v in footnote_pending()])
    print('期初方式2（直接上年审定数）：', [v['name'] for v in opening_audited()])
