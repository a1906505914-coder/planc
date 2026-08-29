# -*- coding: utf-8 -*-
"""confirm_mask.py —— 回函脱敏映射读取（v2 纯数字编码，2026-08-22）。

信息安全规则（用户定）：
  发函公司（被审计单位）→ A 编号代码（含分公司归并母公司）
  回函银行 → B 编号代码
  回函企业 → C 编号代码

⚠️ 本模块只读 confirm_mask.json（由 confirm_mask_local.py 本机生成），
真实名↔代码对应表仅本地保存，不进入任何脱敏输出/对话。
"""
import json
import os
import re

APP_DIR = os.path.dirname(os.path.abspath(__file__))
MASK_FP = os.path.join(APP_DIR, 'confirm_mask.json')

# 企业/银行名通用后缀（用于对文本中残余碎片的兜底替换）
_COMPANY_SUFFIX = [
    '股份有限公司', '有限责任公司', '有限公司', '责任公司', '股份公司',
    '公司', '分公司', '营业部', '经营部', '工厂', '厂',
    '（中国）', '(中国)', '（新加坡）', '(新加坡)', '(THAILAND)', '(S.E.A)', '(泰国)',
    '公众股份公司', '上海分行', '集团', '有限公司德清支行', '支行',
]


def _load_mask():
    if os.path.exists(MASK_FP):
        try:
            with open(MASK_FP, encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def mask_auditee(name):
    """发函公司名 → A 编号代码。未知返回 None。"""
    if not name:
        return None
    n = name.strip()
    data = _load_mask()
    codes = data.get('auditee_codes', {})
    if n in codes:
        return codes[n]
    # 分公司归并：母公司前缀匹配
    for base, code in codes.items():
        if n.startswith(base):
            return code
    return None


def mask_reply(name):
    """回函方名 → B/C 编号代码。未知返回 None。"""
    if not name:
        return None
    n = name.strip()
    data = _load_mask()
    return data.get('reply_map', {}).get(n)


def mask_text(text):
    """文本中的真实主体名 → 脱敏代码（长名优先，容忍跨行空白）。
    ⚡ 2026-08-24 委托统一 desensitizer（confirm 语境 A/B/C，同一映射源
    confirm_mask.json）：实体映射替换 + 中文公司名整体泛化，不再逐后缀 re.sub
    （原实现会把『厂房』误伤成『公司房』）。"""
    if not text:
        return text
    from desensitizer import get_confirm
    return get_confirm().mask(text, 'text')


if __name__ == '__main__':
    d = _load_mask()
    print('发函方代码:')
    for n, c in d.get('auditee_codes', {}).items():
        print(f'  {c} <- {n}')
    print('回函方代码:')
    for n, c in d.get('reply_map', {}).items():
        print(f'  {c} <- {n}')
