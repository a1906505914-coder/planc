# -*- coding: utf-8 -*-
"""financing_ledger.py —— 融资台账 adapter（2026-08-22 从 loan_detail 抽取）。

职责：融资明细台账（融资明细.xls / 融资明细_脱敏.xlsx）的统一读取与脱敏出口，
供短期借款/长期借款（loan_detail）、应付票据、表外事项等专项复用。

安全纪律：**脱敏版优先**——优先读 `融资明细_脱敏.xlsx`（主体/银行/担保已泛化，
直接取值不再二次映射）；无脱敏版才回退原始 .xls（本机审计用，不进对话）。

对外接口（loan_detail 同名下划线名，调用点零改动）：
  ENTITY_MAP / _load_entity_map / _ent_disp       公司代码 → 借X
  _fin_ent_code / _xls_date / _guarantee_to_cat   主体/日期/担保类别
  _GUARANTOR_MAP / _BANK_CODE_MAP                 映射表
  _mask_bank / _mask_guarantor                    银行/担保 脱敏
  _FIN_LOANS_CACHE / _load_finance_loans          贷款类逐笔（ST/LT）
"""
import paths as P
import json
import os
import re

APP_DIR = os.path.dirname(os.path.abspath(__file__))


# ----------------------------------------------------------------------------
# 主体目录代码 → 脱敏代号（借X）映射（来自 contract_mask.json borrower_dir）
# ----------------------------------------------------------------------------
def _load_entity_map():
    try:
        fp = os.path.join(APP_DIR, 'contract_mask.json')
        with open(fp, encoding='utf-8') as f:
            d = json.load(f)
        # key 形如 '3100国望'，取前导数字段为目录代码
        return {re.match(r'\d+', k).group(): v for k, v in d.get('borrower_dir', {}).items()
                if re.match(r'\d+', k)}
    except Exception:
        return {}


ENTITY_MAP = _load_entity_map()


def _ent_disp(ent):
    """主体代码 → 脱敏代号（3100 → 借01）；未收录原样。"""
    return ENTITY_MAP.get(str(ent), str(ent))


# ----------------------------------------------------------------------------
# 融资明细.xls（逐笔借款台账：银行/金额/利率/日期/担保→借款类别）——披露口径数据源
# ----------------------------------------------------------------------------
def _fin_ent_code(n):
    n = str(n)
    if '国望' in n: return '借01'
    if '盛虹纤维' in n: return '借02'
    if '中鲈' in n: return '借03'
    if '港虹' in n: return '借04'
    if '苏震' in n: return '借05'
    if '新视界' in n: return '借06'
    return '借??'


def _xls_date(v):
    try:
        if isinstance(v, float):
            import datetime
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=v)
            return dt.strftime('%Y-%m-%d')
        return str(v)[:10]
    except Exception:
        return ''


def _guarantee_to_cat(g):
    """担保信息 → 借款披露类别（质押/抵押/保证/信用/其他）。"""
    g = str(g)
    if '质押' in g:
        return '质押借款'
    if '抵押' in g:
        return '抵押借款'
    if '担保' in g or '保' in g or '保证' in g:
        return '保证借款'
    if '信用' in g:
        return '信用借款'
    return '其他'


# 融资明细主体名 → 脱敏代码（担保信息/银行列替换，防真实名残留）
_GUARANTOR_MAP = [
    ('江苏东方盛虹股份有限公司', '外部担保方'), ('东方盛虹', '外部担保方'),
    ('江苏盛虹新材料集团有限公司', '外部担保方'), ('盛虹新材料', '外部担保方'),
    ('盛虹控股集团有限公司', '外部担保方'), ('盛虹控股', '外部担保方'),
    ('盛虹集团', '外部担保方'), ('盛虹', '外部担保方'),
    ('江苏国望高科纤维有限公司', '借01'), ('国望高科', '借01'), ('国望', '借01'),
    ('苏州盛虹纤维有限公司', '借02'), ('盛虹纤维', '借02'), ('纤维', '借02'),
    ('江苏中鲈科技发展股份有限公司', '借03'), ('中鲈', '借03'),
    ('江苏港虹纤维有限公司', '借04'), ('港虹', '借04'),
    ('江苏新视界先进功能纤维创新中心有限公司', '借06'), ('新视界', '借06'),
    ('缪汉根夫妇', '个人'), ('缪汉根', '个人'),
]
_BANK_CODE_MAP = [
    ('中国银行', '贷01'), ('中行', '贷01'),
    ('农业银行', '贷02'), ('农行', '贷02'),
    ('工商银行', '贷03'), ('工行', '贷03'),
    ('建设银行', '贷04'), ('建行', '贷04'),
    ('江苏银行', '贷05'),
    ('招银金租', '贷06'),
    ('招商银行', '贷08'), ('邮储', '贷09'), ('南京银行', '贷10'), ('中信银行', '贷11'),
    ('交通银行', '贷12'), ('进出口', '贷13'), ('民生银行', '贷14'), ('宁波银行', '贷15'),
    ('浙商银行', '贷16'), ('浦发银行', '贷17'), ('光大银行', '贷18'), ('华夏银行', '贷19'),
    ('兴业银行', '贷20'), ('平安银行', '贷21'), ('杭州银行', '贷22'), ('恒丰', '贷23'), ('渤海', '贷24'),
]


def _mask_bank(b):
    """银行名脱敏：统一走 Desensitizer（contract 语境 bank_dir 贷X）。"""
    from desensitizer import get_contract
    dz = get_contract()
    # 注册本模块银行映射（与 contract_mask bank_dir 同源，补丁保险）
    dz.register('bank', {kw: code for kw, code in _BANK_CODE_MAP})
    return dz.mask(b, 'bank')


def _mask_guarantor(g):
    """担保信息文本脱敏：真实主体名 → 借X/贷X/盛虹集团/个人（统一 Desensitizer）。"""
    from desensitizer import get_contract
    dz = get_contract()
    # 注册本模块担保映射（主体名→借X/贷X/个人），与 contract_mask borrower_dir 合并
    dz.register('entity', {real: code for real, code in _GUARANTOR_MAP})
    return dz.mask(g, 'entity')


_FIN_LOANS_CACHE = None
_FIN_DATA_ROOT = None   # ⚡ 2026-08-28 当前项目数据根（set_fin_data_root 设置）；None=默认 ADF


def set_fin_data_root(root):
    """⚡ 2026-08-28 修复：原硬编码读 `d:/底稿测试/ADF/数据/2026/融资明细`，AH 等
    其他项目底稿会误用 ADF 借款台账（1010 短期借款融资台账曾整表填入 ADF 数据）。
    调用方（loan_detail.process_folder）按当前项目 data_dir 设置；未设置时保持 ADF 默认。"""
    global _FIN_DATA_ROOT, _FIN_LOANS_CACHE
    if root != _FIN_DATA_ROOT:
        _FIN_DATA_ROOT = root
        _FIN_LOANS_CACHE = None


def _load_finance_loans(sk='ST'):
    """读取融资明细，筛选贷款类，返回逐笔贷款 dict（脱敏代码）。
    ST=流动资金贷款+委托贷款+法人透支；LT=项目贷款+中长期流动资金贷款。
    ⚡ 安全纪律：优先读【脱敏版】融资明细_脱敏.xlsx（主体/银行/担保已泛化，
    直接取值不再二次映射）；无脱敏版才回退原始 .xls（本机审计用，不进对话）。
    ⚡ 2026-08-28 数据根取 set_fin_data_root 设置的当前项目（无则默认 ADF）。"""
    global _FIN_LOANS_CACHE
    if _FIN_LOANS_CACHE is not None:
        return [x for x in _FIN_LOANS_CACHE if x['sk'] == sk]
    base = os.path.join(_FIN_DATA_ROOT or os.path.join(P.DATA_DIRS['ADF'], '数据', '2026'), '融资明细')
    src = None
    try:
        fp_mask = base + '_脱敏.xlsx'
        if os.path.exists(fp_mask):
            from openpyxl import load_workbook
            wb = load_workbook(fp_mask, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            it = ws.iter_rows(values_only=True)
            next(it, None)          # 跳过表头
            src = 'masked'
        else:
            import xlrd
            fp_raw = base + '.xls'
            if not os.path.exists(fp_raw):
                _FIN_LOANS_CACHE = []
                return []
            wb = xlrd.open_workbook(fp_raw)
            ws = wb.sheet_by_index(0)
            src = 'raw'
    except Exception:
        _FIN_LOANS_CACHE = []
        return []
    out = []
    if src == 'masked':
        for i, row in enumerate(it, 1):
            big = str(row[6])
            small = str(row[7])
            if big == '流动资金贷款':
                sk_of = 'ST' if small in ('短期流动资金贷款', '法人透支') else 'LT'
            elif big == '委托贷款':
                sk_of = 'ST'
            elif big == '项目贷款':
                sk_of = 'LT'
            else:
                sk_of = None
            if sk_of is None:
                continue
            code = str(row[3]).strip()          # 已是 借X（脱敏版）
            if not code.startswith('借'):
                continue
            try:
                amt = float(row[13] or 0)
            except (TypeError, ValueError):
                amt = 0.0
            try:
                rate = float(row[18]) if row[18] else None
            except (TypeError, ValueError):
                rate = None
            out.append(dict(
                sk=sk_of, ent=code,
                bank=str(row[4]).strip(),
                no=str(row[5]).strip(),
                small=str(row[7]).strip(),
                amt=amt, rate=rate,
                start=_xls_date(row[10]),
                end=_xls_date(row[11]),
                guar=str(row[22]).strip(),
                cat=_guarantee_to_cat(row[22]),
            ))
    else:
        for i in range(1, ws.nrows):
            big = str(ws.cell_value(i, 6))
            small = str(ws.cell_value(i, 7))
            if big == '流动资金贷款':
                sk_of = 'ST' if small in ('短期流动资金贷款', '法人透支') else 'LT'
            elif big == '委托贷款':
                sk_of = 'ST'
            elif big == '项目贷款':
                sk_of = 'LT'
            else:
                sk_of = None
            if sk_of is None:
                continue
            code = _fin_ent_code(ws.cell_value(i, 3))
            if code == '借??':
                continue
            try:
                amt = float(ws.cell_value(i, 13) or 0)
            except (TypeError, ValueError):
                amt = 0.0
            try:
                rate = float(ws.cell_value(i, 18)) if ws.cell_value(i, 18) else None
            except (TypeError, ValueError):
                rate = None
            out.append(dict(
                sk=sk_of, ent=code,
                bank=str(ws.cell_value(i, 4)).strip(),
                no=str(ws.cell_value(i, 5)).strip(),
                small=str(ws.cell_value(i, 7)).strip(),
                amt=amt, rate=rate,
                start=_xls_date(ws.cell_value(i, 10)),
                end=_xls_date(ws.cell_value(i, 11)),
                guar=str(ws.cell_value(i, 22)).strip(),
                cat=_guarantee_to_cat(ws.cell_value(i, 22)),
            ))
    _FIN_LOANS_CACHE = out
    return [x for x in out if x['sk'] == sk]
