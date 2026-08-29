# -*- coding: utf-8 -*-
"""sap_subject_map.py —— SAP 科目→归属生成器集中映射（2026-08-09，建议1 轻量版）。

目的：把 13 个生成器的科目边界集中到一张表（防漂移/防双计/防遗漏的统一依据）。
取数逻辑仍在各生成器内（不重构），本表用于：
  ① 科目覆盖完整性自检（_sap_coverage_scan.py 引用）；
  ② 新账套接入时快速核对科目归属；
  ③ 未来重构（subjects_registry 接入）的过渡层。

⚠️ 内部码为 AH 账套 SAP 方言（历史名 yy2026）实测（铁律61），改前必须实测 TB 实际码。
"""
from collections import defaultdict

# ============================================================
# SAP 期间费用特判码集中表（2026-08-10 实测 1010；改前必须实测 TB 实际码）
#   分散在 expense_detail/payroll_detail 里的 660006/6603/6600/660303/660304
#   字面量判断统一收口到本表 —— 防漂移/防遗漏的唯一依据。
# ============================================================
SAP_FEE_CODES = {
    'rd': '660006',          # 研究开发费（SAP 特有码，标准码体系无）
    'finance': '6603',       # 财务费用（期间费用下独立科目）
    'pool': '6600',          # 期间费用总池（管理/销售/制造/研发混合）
    'fx_loss': '660303',     # 汇兑损失（财务费用下子目）
    'fx_gain': '660304',     # 汇兑收益（财务费用下子目）
}

# SAP 功能范围列(col19) → 费用类别（6600 总池无 TB 拆分，只能靠 GL 功能范围列）
SAP_FR_CAT = {
    '6601': '销售费用',
    '6602': '管理费用',
    '5101': '制造费用',
    '5301': '研发费用',
}


def fee_code(key):
    """SAP 费用特判码查询（默认表，供 expense/payroll 等生成器统一引用）。"""
    return SAP_FEE_CODES.get(key)


def fr_category(fr):
    """功能范围列 → 标准费用类别名（无则 None）。"""
    return SAP_FR_CAT.get(str(fr or '').strip())


# 科目 → 归属生成器（4 位内部码前缀）
SUBJECT_MAP = {
    # 货币资金
    '1001': 'bank', '1002': 'bank', '1012': 'bank',
    # 往来（应收/应付/预收预付/合同）
    '1121': 'current_account',   # 应收票据
    '1122': 'current_account',   # 应收账款
    '1123': 'current_account',   # 预付账款
    '1124': 'current_account',   # 合同资产
    '1221': 'current_account',   # 其他应收款
    '1222': 'current_account',   # 备用金
    '2201': 'current_account',   # 应付票据
    '2202': 'current_account',   # 应付账款
    '2203': 'current_account',   # 预收账款
    '2205': 'current_account',   # 合同负债（SAP 无，预留）
    '2241': 'current_account',   # 其他应付款
    # 存货
    '1401': 'inventory', '1403': 'inventory', '1404': 'inventory',
    '1405': 'inventory', '1406': 'inventory', '1407': 'inventory',
    '1411': 'inventory', '1412': 'inventory',   # 1412 包装物及低值易耗品
    '5001': 'inventory',          # 生产成本（四段闭环链）
    # 长期资产
    '1521': 'longterm', '1522': 'longterm',   # 投资性房地产（原值/累计折旧）
    '1601': 'longterm', '1602': 'longterm',   # 固定资产/累计折旧
    '1604': 'longterm',           # 在建工程
    '1605': 'longterm',           # 工程物资
    '1608': 'longterm', '1609': 'longterm',   # 使用权资产/累计折旧
    '1701': 'longterm', '1702': 'longterm',   # 无形资产/累计摊销
    '1705': 'longterm',           # 无形资产-待转（余额 0，防未来遗漏）
    # 权益
    '4001': 'equity', '4002': 'equity', '4003': 'equity',   # 实收/资本公积/库存股
    '4004': 'equity',             # 其他权益工具（权益类）
    '4101': 'equity', '4102': 'equity',   # 盈余公积/专项储备
    '4103': 'equity', '4104': 'equity',   # 本年利润/利润分配
    # 借款
    '2001': 'loan', '2501': 'loan',       # 短期/长期借款
    # 税费/薪酬
    '2221': 'tax',                # 应交税费
    '2211': 'payroll',            # 应付职工薪酬
    # 研发
    '1704': 'rd_expense',         # 研发支出（开发支出类）
    # 损益（revenue/expense/pl）
    '6001': 'revenue', '6051': 'revenue', '6301': 'revenue',   # 主营/其他/营业外收入
    '6401': 'revenue', '6402': 'revenue',                       # 主营/其他成本
    '6403': 'pl',                # 税金及附加
    '6600': 'expense', '6601': 'expense', '6602': 'expense',
    '6603': 'expense', '6604': 'expense',
    '6111': 'pl', '6115': 'pl', '6711': 'pl', '6720': 'pl',
    '6730': 'pl', '6741': 'pl', '6701': 'pl', '6801': 'pl', '6901': 'pl',
    # 其他类（gp_other）
    '1101': 'gp_other',           # 交易性金融资产
    '1131': 'gp_other', '1132': 'gp_other',   # 应收股利/应收利息
    '1231': 'gp_other',           # 坏账准备
    '1301': 'gp_other',           # 委托贷款
    '1471': 'gp_other',           # 存货跌价准备
    '1489': 'gp_other',           # 碳排放权交易
    '1501': 'gp_other', '1503': 'gp_other',   # 持有至到期投资/其他权益工具投资
    '1511': 'gp_other', '1512': 'gp_other',   # 长期股权投资/减值准备
    '1504': 'gp_other',           # 待摊费用
    '1532': 'gp_other',           # 未实现融资收益
    '1606': 'gp_other',           # 固定资产清理
    '1607': 'gp_other',           # 在建工程减值准备
    '1603': 'gp_other',           # 固定资产减值准备（与坏账/跌价准备同 gp_other）
    '1611': 'gp_other', '1621': 'gp_other',   # 使用权/资产减值准备
    '1703': 'gp_other',           # 无形资产减值准备
    '1711': 'gp_other',           # 商誉
    '1801': 'gp_other',           # 长期待摊费用
    '1811': 'gp_other',           # 递延所得税资产
    '1831': 'gp_other',           # 其他非流动资产
    '1901': 'gp_other',           # 待处理财产损溢
    '2231': 'gp_other',           # 应付利息
    '2232': 'gp_other',           # 应付股利
    '2260': 'gp_other',           # 租赁负债
    '2401': 'gp_other',           # 递延收益
    '2402': 'gp_other',           # 预提费用
    '2502': 'gp_other',           # 应付债券
    '2701': 'gp_other',           # 长期应付款
    '2711': 'gp_other',           # 专项应付款
    '2801': 'gp_other',           # 预计负债
    '2901': 'gp_other',           # 递延所得税负债
}

# 反查：生成器 → 科目码集合
GENERATOR_SUBJECTS = defaultdict(set)
for _c, _g in SUBJECT_MAP.items():
    GENERATOR_SUBJECTS[_g].add(_c)


def covered_by(code4):
    """4 位内部码 → 归属生成器（None=未覆盖，应立即补表/补生成器）。"""
    return SUBJECT_MAP.get(code4)


if __name__ == '__main__':
    print('SAP 科目映射表：%d 个主码，%d 个生成器' % (len(SUBJECT_MAP), len(GENERATOR_SUBJECTS)))
    for g in sorted(GENERATOR_SUBJECTS):
        print('  %-16s %s' % (g, sorted(GENERATOR_SUBJECTS[g])))
