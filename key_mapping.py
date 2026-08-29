# -*- coding: utf-8 -*-
"""key_mapping.py —— 台账↔账套键映射配置框架（2026-08-22）。

勾稽核对的核心：台账（业务事实）与账套（账面记录）之间用什么键对齐。
本模块统一登记【键映射】（natural key），供各专项勾稽复用：
  台账行 ↔ 账套 TB/GL 行 按登记的键字段取值，做匹配/归集。

用法：
  from key_mapping import key_map, get_key, is_match
  cfg = key_map('ADF', 'financing')      # → 该账套该台账的键配置
  k1 = get_key(cfg, ledger_row, '台账')   # 从台账行取键
  k2 = get_key(cfg, tb_row, '账套')       # 从账套行取键
  is_match(k1, k2)                        # 键是否对齐

新增台账/账套 = 在 KEY_MAP 加一段配置，不写代码。
"""
import re

# ============================================================
# 键映射登记表：{账套: {台账类型: {key_fields: {侧: 字段/列名}, match: 匹配方式}}}
#   key_fields  键字段：台账侧 / 账套侧（列名或索引）
#   match       匹配方式：'exact'(默认) / 'norm'(去空格/大小写) / 'date'(日期归一)
#   subject     关联科目（说明性）
# ============================================================
KEY_MAP = {
    'ADF': {
        'financing': {
            'subject': '短期借款/长期借款',
            'key_fields': {'台账': '业务编号', '账套': '合同号'},
            'match': 'norm',
        },
        'fa': {
            'subject': '固定资产',
            'key_fields': {'台账': '资产号码', '账套': '科目子目'},
            'match': 'exact',
        },
    },
    'AZ': {
        'confirm': {
            'subject': '银行存款/应收账款',
            'key_fields': {'台账': '账户号', '账套': '账号'},
            'match': 'norm',
        },
    },
}

_NORM_PAT = re.compile(r'[\s\-/_.]+')


def key_map(acct, ledger_type):
    """取账套×台账的键配置；未登记返回 {}。"""
    return (KEY_MAP.get(acct) or {}).get(ledger_type) or {}


def get_key(cfg, row, side='台账'):
    """按配置从行取键。row 支持 dict（按列名）或 list（按索引）。"""
    fld = (cfg.get('key_fields') or {}).get(side)
    if not fld:
        return None
    if isinstance(row, dict):
        v = row.get(fld)
    else:
        try:
            v = row[int(fld)]
        except (ValueError, TypeError, IndexError):
            v = None
    return '' if v is None else str(v).strip()


def is_match(cfg, k1, k2, fuzzy=False):
    """键是否对齐。fuzzy=True 时忽略空格/分隔符/大小写（norm 匹配）。"""
    if not k1 or not k2:
        return False
    if (cfg.get('match') or 'exact') == 'norm' or fuzzy:
        return _NORM_PAT.sub('', k1).lower() == _NORM_PAT.sub('', k2).lower()
    return k1 == k2


def list_maps():
    print('=== 键映射登记 ===')
    for acct, ledgers in KEY_MAP.items():
        for lt, cfg in ledgers.items():
            kf = cfg.get('key_fields') or {}
            print('  %-5s %-10s 台账键=%s ↔ 账套键=%s（%s）%s'
                  % (acct, lt, kf.get('台账'), kf.get('账套'),
                     cfg.get('match'), cfg.get('subject') or ''))
