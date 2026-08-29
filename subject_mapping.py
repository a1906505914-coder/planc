# -*- coding: utf-8 -*-
"""subject_mapping.py —— 2026-08-07 新增（任务①科目映射配置化）。

把散落在各 builder 里的账套特异科目识别（if 分支/别名表/结构 flag）抽成
【数据文件】account_profiles.json，程序统一从这里读取。新增账套 = 改配置，
不改代码（用户方法论：几百账套时配置驱动，避免"适配 A 破坏 B"行为漂移）。

用法：
    from subject_mapping import account_profile, feature, norm_name
    p = account_profile('H')                 # -> {'name_map': {...}, ...}
    feature('AJ', 'no_gl')                  # -> True
    norm_name('H', '递延资产')                # -> '长期待摊费用'
    norm_name('Q', '投资损益')                # -> '投资收益'
"""
import json
import os

_CFG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'account_profiles.json')
_CACHE = {}


def _load():
    if not _CACHE:
        try:
            with open(_CFG_PATH, encoding='utf-8') as f:
                _CACHE['data'] = json.load(f)
        except Exception as ex:
            _CACHE['data'] = {'common_name_map': {}, 'accounts': {}}
            print(f'⚠️ account_profiles.json 读取失败：{ex}')
    return _CACHE['data']


def common_name_map():
    return _load().get('common_name_map') or {}


def account_profile(account):
    """返回账套配置 dict（无则空）。account=数据文件夹名（FY/JTt/Z...）。"""
    return (_load().get('accounts') or {}).get(account) or {}


def feature(account, key, default=None):
    """读取账套结构特征 flag：feature('AJ', 'no_gl') -> True。"""
    return account_profile(account).get(key, default)


def norm_name(account, raw_name):
    """账套科目名 → 标准名（一级段）。优先级：账套 name_map > 通用 common_name_map。
    路径分层名（JTt『640303\\税金及附加\\城市维护建设税』）取一级段（第二段）。
    account 为空或未登记 → 仅做通用映射。"""
    nm = str(raw_name or '').strip()
    if not nm:
        return nm
    # 反斜杠路径名：取一级段（JTt『代码\一级名\二级名』→ 一级名）
    if '\\' in nm:
        _parts = nm.split('\\')
        nm = _parts[1].strip() if len(_parts) >= 2 else _parts[-1].strip()
    p = account_profile(account)
    if nm in (p.get('name_map') or {}):
        return p['name_map'][nm]
    cm = common_name_map()
    if nm in cm:
        return cm[nm]
    return nm


def name_variants(account, standard_name):
    """给定标准名 → 该账套可能的实际科目名列表（含映射反查 + 通用别名反查）。
    供取数匹配（铁律5 名称优先）做候选扩展。"""
    out = [standard_name]
    cm = common_name_map()
    for _raw, _std in cm.items():
        if _std == standard_name and _raw not in out:
            out.append(_raw)
    p = account_profile(account)
    for _raw, _std in (p.get('name_map') or {}).items():
        if _std == standard_name and _raw not in out:
            out.append(_raw)
    return out


def resolve_subject_codes(account, subject_key, default_codes=None):
    """账套级科目代码解析（2026-08-07 v2：会计语言优先，代码方言隔离）。

    返回该账套某科目应使用的代码前缀列表：
      · account_profiles.json -> accounts.<账套>.subject_codes.<科目key> 有值 → 用之；
      · 否则返回 default_codes（subjects_registry 的标准码 / builder 内部默认码）。

    设计意图：账套特例码（dq 递延收益=2601、H 实收资本=3003、Q 租赁负债=2271 等）
    全部登记在配置里，【不进全局注册表】→ 消除跨账套 codes 撞车（2601 在 dq=递延
    收益、FY=租赁负债的历史 bug 从此不可能发生）。新增账套 = 改配置，不改代码。
    """
    p = account_profile(account)
    sc = p.get('subject_codes') or {}
    if subject_key in sc:
        return [str(c) for c in sc[subject_key]]
    return [str(c) for c in (default_codes or [])]


if __name__ == '__main__':
    for a, raw in (('AJJ', '递延资产'), ('AQ', '投资损益'), ('AQ', '资产处理损益'),
                   ('AJ', '640303\\税金及附加\\城市维护建设税'), ('GFY', '股本'), ('X', '递延资产')):
        print(f'{a}: {raw!r} -> {norm_name(a, raw)!r}')
    print('JTt no_gl =', feature('AJ', 'no_gl'), '| J alpha =', feature('ADZ', 'alpha_suffix_codes'))
