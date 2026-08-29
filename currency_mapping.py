# -*- coding: utf-8 -*-
"""currency_mapping.py —— 主体→记账本位币识别（2026-08-05 用户需求：Z 集团外币主体分币种处理）。

问题背景：Z 文件夹（浙江兆龙集团）中 龙腾互连（泰国）以泰铢(THB)记账、龙腾控股（新加坡）
以美元(USD)记账，其余主体（母公司/高分子/数链/物联及分公司）以人民币(CNY)记账。
此前所有金额被当作人民币混同相加 → 底稿/附注/合并口径全部失真。

原则（铁律）：【按主体名识别币种】，币种是主体的记账本位币属性，与科目代码无关；
不同币种主体的金额【绝不直接相加】，只在合并层按汇率折算成人民币后统一口径。

用法：
  from currency_mapping import entity_currency, CURRENCY_LABEL, is_cny_entity, currency_of_tb_row
  cur = entity_currency('龙腾互连（泰国）有限公司')   # -> 'THB'
  cur = entity_currency('新加坡')                     # -> 'USD'
  cur = entity_currency('母公司')                     # -> 'CNY'
"""
import re

# 主体名关键词 → 记账本位币（CNY/THB/USD）。
# 注意：泰语主体名可能含『泰国/THAI/Thailand』；新加坡主体名含『新加坡/Singapore』。
# 匹配优先级：先外币（THB/USD），后 CNY（CNY 是兜底默认，无需关键词）。
CURRENCY_KEYWORDS = [
    (('泰国', 'THAI', 'THAILAND', '泰國'), 'THB'),
    (('新加坡', 'SINGAPORE', '星加坡'), 'USD'),
]

CURRENCY_LABEL = {
    'CNY': '人民币',
    'THB': '泰铢',
    'USD': '美元',
}

CURRENCY_CODE = {
    'CNY': 'CNY',
    'THB': 'THB',
    'USD': 'USD',
}


def entity_currency(entity_name):
    """按主体名识别记账本位币：泰国系→THB，新加坡系→USD，其余→CNY（兜底）。"""
    if not entity_name:
        return 'CNY'
    s = str(entity_name).upper()
    for kws, cur in CURRENCY_KEYWORDS:
        if any(k.upper() in s for k in kws):
            return cur
    return 'CNY'


def is_cny_entity(entity_name):
    """是否人民币主体（用于分币种生成时判断是否需单独列示）。"""
    return entity_currency(entity_name) == 'CNY'


def currency_of_tb_row(tb_key_or_entity):
    """从 TB dict key（e, code, name, y）或实体名提取币种。"""
    if isinstance(tb_key_or_entity, (tuple, list)):
        return entity_currency(tb_key_or_entity[0])
    return entity_currency(tb_key_or_entity)


def group_entities_by_currency(entities):
    """把实体名集合按币种分组：{currency: [entities]}（CNY 组放最后）。"""
    groups = {}
    for e in entities:
        groups.setdefault(entity_currency(e), []).append(e)
    order = sorted(groups, key=lambda c: (0 if c == 'CNY' else 1, c))
    return {c: groups[c] for c in order}


def label_with_currency(entity_name):
    """主体显示名 + 币种后缀（如『母公司（人民币）』『泰国（泰铢）』），用于底稿表头标注。"""
    cur = entity_currency(entity_name)
    return '%s（%s）' % (entity_name, CURRENCY_LABEL.get(cur, cur))
