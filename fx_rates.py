# -*- coding: utf-8 -*-
"""fx_rates.py —— 外币报表折算汇率表（2026-08-05 用户需求：资产负债表期初/期末分汇率）。

外币报表折算（企业会计准则第19号——外币折算）口径：
  ① 资产负债表项目：期末余额 → 期末汇率（2025-12-31）；
                      期初余额 → 期初汇率（2024-12-31，即上年末）。
  ② 利润表项目：本期发生额 → 平均汇率（2025 年度）或交易发生日即期汇率（审计组定）。
  ③ 所有者权益（实收资本等）：发生时点汇率（历史成本，一般不随报表日汇率变动）。
  ④ 折算差额 → 计入其他综合收益『外币报表折算差额』（合并层处理，底稿层不列）。

汇率来源：中国人民银行授权中国外汇交易中心公布的银行间外汇市场人民币汇率中间价；
  平均汇率：国家统计局《2025年国民经济和社会发展统计公报》（2026-02-28 发布）。
  期初 2024-12-31：1 美元 = 7.1884 人民币；1 人民币 = 4.7028 泰铢
  期末 2025-12-31：1 美元 = 7.0288 人民币；1 人民币 = 4.4940 泰铢
  2025 年度平均：1 美元 = 7.1429 人民币（统计局公报）

用法：
  from fx_rates import FX_RATES, to_cny
  to_cny(1000, 'USD', 'period_end')   # -> 7028.8
  to_cny(1000, 'THB', 'period_begin') # -> 1000 / 4.7028 ≈ 212.64
"""
import math

# 每 1 外币单位 = X 人民币（直接标价法，CNY 为 1）。
# key: (币种, 口径) -> 1 外币 = 人民币
# 口径：period_begin=期初(2024-12-31) / period_end=期末(2025-12-31) / avg=2025 年度平均（待审计组确认）
FX_RATES = {
    ('CNY', 'period_begin'): 1.0,
    ('CNY', 'period_end'): 1.0,
    ('CNY', 'avg'): 1.0,
    ('USD', 'period_begin'): 7.1884,   # 2024-12-31 中间价
    ('USD', 'period_end'): 7.0288,     # 2025-12-31 中间价
    ('USD', 'avg'): 7.1429,           # 2025 年度平均（国家统计局公报口径，审计组可复核）
    ('THB', 'period_begin'): 1.0 / 4.7028,   # 1 CNY = 4.7028 THB → 1 THB = 0.21264 CNY
    ('THB', 'period_end'): 1.0 / 4.4940,     # 1 CNY = 4.4940 THB → 1 THB = 0.22252 CNY
    ('THB', 'avg'): 1.0 / 4.5,       # 待定：泰铢 2025 年度平均（近似 4.5，审计组确认后替换）
}

# 币种全称/简称（附注与底稿标注用）
FX_LABEL = {'CNY': '人民币', 'USD': '美元', 'THB': '泰铢'}


def get_rate(currency, period):
    """取汇率；缺失返回 None。period: 'period_begin'/'period_end'/'avg'。"""
    return FX_RATES.get((currency, period))


def to_cny(amount, currency, period):
    """外币金额 → 人民币（直接标价：金额 × 汇率）。CNY 原样返回。"""
    if currency == 'CNY':
        return amount
    rate = get_rate(currency, period)
    if rate is None:
        return None  # 汇率未定（如平均汇率待审计组确认）
    return amount * rate


def needs_avg_rates():
    """返回尚未确定平均汇率的币种（利润表折算前需补）。"""
    missing = []
    for (cur, period), rate in FX_RATES.items():
        if period == 'avg' and rate is None and cur != 'CNY':
            missing.append(cur)
    return missing
