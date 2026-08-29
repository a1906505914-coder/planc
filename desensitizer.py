# -*- coding: utf-8 -*-
"""desensitizer.py —— 统一脱敏层（2026-08-24 架构·输出侧）。

现状问题：confirm_mask/contract_mask/fa_ledger_mask/financing_ledger 各自实现
mask，逻辑分散、一致性无保证（同一客户在不同 sheet/底稿打码可能不同 →
脱敏版无法核对）。

统一设计：
- Desensitizer 实例持有【一致性映射】（实体名 → 脱敏码），同一实例跨 sheet/底稿
  调用 mask 返回相同结果
- mask(value, field_type) 按字段类型脱敏：
    'entity'  往来单位/公司名   → 映射优先，未收录则中文公司名泛化
    'bank'    银行名            → 贷X 映射（contract bank_dir），未收录泛称
    'account' 银行账号          → 保留前4后4，中间 *（可核对尾号）
    'person'  人员名            → 保留姓氏
    'code'    编码/号码（凭证号/资产号）→ 保留（配对 key）
    'text'    通用文本          → 实体名 + 银行名部分替换（跨行容忍）+ 公司名泛化
- 保留配对 key：资产号码/凭证号/账号尾号等用于跨期/跨表核对，不脱敏

映射语境（sources 控制，防止 A/B/C 与 借X/贷X 混用）：
    'confirm'  → confirm_mask.json：发函方 A 编号 / 回函方 B/C 编号（函证）
    'contract' → contract_mask.json：借款主体 借X / 银行 贷X（融资/资产台账）
    默认全量加载；各业务程序按语境取实例：
        get_confirm()  → 只加载 confirm 映射（函证回函）
        get_contract() → 只加载 contract 映射（借款/融资/资产台账）

用法：
    from desensitizer import Desensitizer, get_contract
    dz = get_contract()                    # 融资/资产语境默认实例
    dz.register('entity', {'上海万坤实业发展有限公司': 'A001'})
    dz.mask('上海万坤实业发展有限公司', 'entity')   # → 'A001'（后续调用结果相同）
    dz.mask('中国银行', 'bank')                     # → '贷01'
"""
import json
import os
import re

# 公司名后缀 → 泛称（fa_ledger_mask 同款，用于文本兜底泛化）
_COMPANY_SUFFIX = ('股份有限公司', '有限责任公司', '集团有限公司', '有限公司',
                   '集团', '股份公司', '分公司', '支行', '分行', '银行',
                   '控股', '有限', '公司', '工厂', '厂', '合作社', '事务所')

# 中文公司名模式（2+ 中文字 + 后缀）——文本兜底泛化
_COMPANY_PAT = re.compile(
    r'[\u4e00-\u9fff]{2,}(' + '|'.join(_COMPANY_SUFFIX) + r')')

# 银行映射未收录兜底
_DEFAULT_BANK_FALLBACK = '外部银行'

# 映射语境 → 对应 JSON 文件（小程序根目录）
_SOURCE_FILES = {
    'confirm': 'confirm_mask.json',
    'contract': 'contract_mask.json',
}


class Desensitizer:
    """统一脱敏器：一致性映射 + 字段类型泛化。"""

    def __init__(self, load_default=False, sources=('confirm', 'contract')):
        self._maps = {}       # field_type -> {原始名: 脱敏码}
        self._cache = {}      # (field_type, 原始名) -> 脱敏码（一致性保证）
        self._code_seq = 0    # 未收录实体的自增码
        self._sources = tuple(sources)
        if load_default:
            self._load_default()

    # ---------------- 注册 ----------------
    def register(self, field_type, mapping):
        """注册映射 {原始名: 脱敏码}（长名优先）。"""
        self._maps.setdefault(field_type, {}).update(mapping or {})

    def _load_default(self):
        """按语境载入映射 JSON（sources 控制加载哪些）。"""
        self.register_maps_from_json(sources=self._sources)

    def register_maps_from_json(self, app_dir=None, sources=None):
        """统一从映射 JSON 注册（app_dir 默认小程序根目录）。
        sources：'confirm'（A/B/C）/ 'contract'（借X/贷X），默认按实例 sources。"""
        app_dir = app_dir or os.path.dirname(os.path.abspath(__file__))
        for src in (sources or self._sources):
            fn = _SOURCE_FILES.get(src)
            if not fn:
                continue
            try:
                fp = os.path.join(app_dir, fn)
                if not os.path.exists(fp):
                    continue
                with open(fp, encoding='utf-8') as f:
                    d = json.load(f)
            except Exception:
                continue
            if src == 'confirm':
                # 发函方 A 编号 / 回函方 B/C 编号（entity 映射）
                self.register('entity', d.get('auditee_codes', {}))
                self.register('entity', d.get('reply_map', {}))
            else:  # 'contract'
                # borrower_dir key 形如 '3100国望' → 目录代码（3100）为映射键，
                #   名称部分（国望）是简称、注册为 entity 会部分替换产生残渣
                #   （江苏国望高科纤维有限公司 → 江苏借01【客商】）——只注册代码映射，
                #   完整公司名由各程序全名映射（如 financing_ledger._GUARANTOR_MAP）补充。
                bd = d.get('borrower_dir', {}) or {}
                ent = {}
                for k, v in bd.items():
                    m = re.match(r'(\d+)(.*)', str(k))
                    if m and m.group(1):
                        ent[m.group(1)] = v
                self.register('entity', ent)
                self.register('bank', d.get('bank_dir', {}) or {})

    # ---------------- 脱敏 ----------------
    def mask(self, value, field_type='text'):
        """按字段类型脱敏单个值。None/数字/编码 安全。"""
        if value is None:
            return value
        s = str(value)
        if not s.strip():
            return s
        key = (field_type, s)
        if key in self._cache:
            return self._cache[key]
        out = self._mask_by_type(s, field_type)
        self._cache[key] = out
        return out

    def _mask_by_type(self, s, field_type):
        if field_type == 'code':
            return s                       # 编码保留（配对 key）
        if field_type == 'account':
            return self._mask_account(s)
        if field_type == 'person':
            return self._mask_person(s)
        if field_type == 'bank':
            return self._mask_bank(s)
        if field_type == 'entity':
            return self._mask_entity(s)
        return self._mask_text(s)          # 'text' 兜底

    # ---------------- 字段类型实现 ----------------
    def _mask_entity(self, s):
        """实体名：映射优先 → 未收录时中文公司名泛化（长名先试，短名兜底）。"""
        mp = self._maps.get('entity', {})
        # 精确映射优先
        if s in mp:
            return mp[s]
        # 长名优先的包含替换（容忍字段带前后缀，如『客户：上海万坤实业发展有限公司』）
        for name in sorted(mp, key=len, reverse=True):
            if name and name in s:
                return s.replace(name, mp[name])
        # 未收录：中文公司名泛化（长名→短码）
        m = _COMPANY_PAT.search(s)
        if m:
            # 生成稳定码：公司名 hash 前缀（保证同一名同码）
            import hashlib
            code = 'C' + hashlib.md5(s.encode('utf-8')).hexdigest()[:6].upper()
            return s[:m.start()] + code + s[m.end():]
        return s

    def _mask_bank(self, s):
        """银行名：映射优先（贷X）→ 未收录泛称。"""
        mp = self._maps.get('bank', {})
        if s in mp:
            return mp[s]
        for name in sorted(mp, key=len, reverse=True):
            if name and name in s:
                return mp[name]
        return _DEFAULT_BANK_FALLBACK

    def _mask_text(self, s):
        """通用文本：实体名 + 银行名【部分替换】+ 中文公司名泛化。
        部分替换 = 只替换命中的映射子串，不整体兜底（文本中未收录银行保留原样，
        区别于单值 _mask_bank 的整体兜底『外部银行』）。跨行/空格容忍。
        公司名兜底泛化不误伤业务词（『厂房』不会变『公司房』）。"""
        out = s
        for ft in ('entity', 'bank'):
            mp = self._maps.get(ft, {})
            for name in sorted(mp, key=len, reverse=True):
                if not name:
                    continue
                if name in out.replace('\n', '').replace(' ', ''):
                    pat = r'\s*'.join(re.escape(c) for c in name)
                    out = re.sub(pat, mp[name], out)
        out = _COMPANY_PAT.sub('【客商】', out)
        return out

    def _mask_account(self, s):
        """银行账号：保留前4后4，中间 *（可核对尾号）。"""
        digits = re.sub(r'\D', '', s)
        if len(digits) > 8:
            return digits[:4] + '*' * (len(digits) - 8) + digits[-4:]
        if len(digits) > 4:
            return digits[:2] + '*' * (len(digits) - 4) + digits[-2:]
        return s

    def _mask_person(self, s):
        """人员名：保留姓氏（首字）+ *。"""
        if len(s) <= 1:
            return s
        return s[0] + '*' * (len(s) - 1)

    # ---------------- 批量 ----------------
    def mask_df_col(self, values, field_type):
        """批量脱敏（list/可迭代）→ list。"""
        return [self.mask(v, field_type) for v in values]


# ---------------- 语境默认实例（一致性跨模块共享） ----------------
_instances = {}


def _get_instance(sources, load_default=True):
    """按语境 sources 取进程内单例（同一实例保证跨 sheet/底稿一致性）。"""
    key = tuple(sources)
    if key not in _instances:
        _instances[key] = Desensitizer(load_default=load_default, sources=key)
    return _instances[key]


def get_default(load_default=False):
    """全量默认实例（confirm+contract）——通用场景，避免语境串码时用 get_confirm/get_contract。"""
    return _get_instance(('confirm', 'contract'), load_default=load_default)


def get_confirm():
    """函证回函语境默认实例（只加载 confirm_mask.json A/B/C 编号）。"""
    return _get_instance(('confirm',))


def get_contract():
    """融资/资产台账语境默认实例（只加载 contract_mask.json 借X/贷X）。"""
    return _get_instance(('contract',))
