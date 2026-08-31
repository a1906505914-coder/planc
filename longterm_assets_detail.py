# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=各长期资产科目底稿(固定资产/投资性房地产/无形资产/使用权资产/长期待摊费用/在建工程)审计底稿_{year}_生成.xlsx | 关键列=原值增减/折旧摊销/减少去向/在建工程转入 | 职责=长期资产增减变动与折旧摊销勾稽(升级版,替代fixed_assets_detail.py)

from audit_common import _safe_save, read_km as _read_km  # 共享库：统一锁感知保存（单一来源）+ 通用读取解析器
from mask_dict import mask_names   # ⚡ console 主体列表打印脱敏（2026-08-22 复扫）
"""长期资产科目审计底稿生成程序（固定资产小程序升级版）
================================================================================
将「固定资产小程序」升级为「长期资产科目小程序」，覆盖以下长期资产科目组
（按科目【名称】自动识别，兼容不同会计科目表；如本集团使用权资产用 1704/1705、
无形资产用 1701/1702、在建工程用 1604）：

  · 固定资产(FA)        1601 原值 / 1602 累计折旧 / 1603 减值准备 / 1606 清理
  · 投资性房地产(IPR)    1521 原值 / 1522 累计折旧(摊销) / 1523 减值准备   [本集团账套无,引擎仍支持]
  · 无形资产(INT)        1701 原值 / 1702 累计摊销 / 1703 减值准备
  · 使用权资产(ROU)      1704 原值 / 1705 使用权资产累计折旧
  · 长期待摊费用(LTDEF)  1801 原值(摊销=1801贷方,无独立备抵科目)
  · 在建工程(CIP)        1604 原值(特殊:仅滚调表+减少去向统计)

数据来源（与固定资产底稿一致）：
  · 科目余额表 = 权威控制数（期初/借发/贷发/期末，带借贷符号）。
  · 综合查询明细表(GL) = 凭证级，用于：增加/减少按类型拆分、折旧/摊销分月分科目
    计提核对、减少逐笔核对、在建工程减少去向。

输出（每个【存在该科目的主体】的年份，分别生成一份底稿）：
  · <被拖文件夹>/固定资产审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/投资性房地产审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/无形资产审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/使用权资产审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/长期待摊费用审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/在建工程审计底稿_{year}_生成.xlsx
规则/格式一律参考《固定资产审计底稿》；在建工程相对特殊，仅生成
期初/本期增加/本期减少/期末/审计调整/审定数 格式，并统计本期减少的去向
（转入固定资产/无形资产/长期待摊费用/其他）。固定资产、无形资产等底稿中
亦统计「从在建工程转入数」（见各底稿《增加检查表》的「在建工程转入」分类及汇总）。
================================================================================
"""

# ---- 启动诊断（先于 import openpyxl，便于排查"拖入闪退无痕迹"）----
import os as _bs_os
import sys as _bs_sys
import subprocess as _bs_sub


def _write_boot_log(tag):
    import os as _bl_os
    if _bl_os.environ.get('AUDIT_LOG') != '1':
        return
    try:
        import datetime
        _p = _bs_os.path.join(_bs_os.path.expanduser('~'), 'Desktop', 'longterm_assets_boot.log')
        with open(_p, 'a', encoding='utf-8') as _f:
            _f.write('[%s] %s sys=%s argv=%s\n' % (
                datetime.datetime.now().isoformat(timespec='seconds'),
                tag, _bs_sys.executable, _bs_sys.argv))
    except Exception:
        pass


_write_boot_log('LOAD')


def _find_managed_python():
    base = _bs_os.path.join(_bs_os.environ.get('USERPROFILE', _bs_os.path.expanduser('~')),
                            '.workbuddy', 'binaries', 'python', 'versions')
    if _bs_os.path.isdir(base):
        for d in sorted(_bs_os.listdir(base), reverse=True):
            p = _bs_os.path.join(base, d, 'python.exe')
            if _bs_os.path.isfile(p):
                return p
    return None


def _bootstrap_relaunch():
    return  # ⚡ 2026-08-22 禁用托管Python重启：venv依赖已齐全，托管3.13 openpyxl损坏


if _bs_os.environ.get('AUDIT_NO_RELAUNCH') != '1':
    _bootstrap_relaunch()

import argparse
import os
import re
import itertools
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TOT_FILL, SHELL_TITLE_FONT, SHELL_SUB_FONT,
                         SHELL_BOLD, SHELL_NUM, SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT,
                         finalize_workbook)


# ----------------------------------------------------------------------------
# 各长期资产科目组定义（按科目【名称】识别，兼容不同科目表）
# ----------------------------------------------------------------------------
# cost_names : 一级原值科目名（精确匹配 level-1）
# accum_names: 累计折旧/累计摊销科目名（精确匹配）
# impair_names: 减值准备科目名（前缀匹配）
# clear_names : 清理科目名（前缀匹配，仅 FA）
# accum_kw    : 该组折旧/摊销费用在 GL 中的关键词（用于费用列支识别）
# amort       : 是否有折旧/摊销计提表
# amort_src   : 计提数来源账户角色（'accum'=累计折旧/摊销贷方；'cost'=原值贷方,如长期待摊费用）
# special     : 在建工程特殊处理（仅滚调表+减少去向）
GROUPS = {
    'FA': {
        'title': '固定资产', 'paper': '固定资产审计底稿',
        'main_sheet': '固定资产分类汇总表',
        'amort_sheet': '折旧分摊表',
        'inc_sheet': '固定资产增加检查表', 'dec_sheet': '固定资产减少检查表',
        'cost_names': ['固定资产'], 'accum_names': ['累计折旧'],
        'impair_names': ['固定资产减值准备'], 'clear_names': ['固定资产清理'],
        'accum_kw': '折旧', 'amort': True, 'amort_src': 'accum', 'special': False,
    },
    'IPR': {
        'title': '投资性房地产', 'paper': '投资性房地产审计底稿',
        'main_sheet': '投资性房地产分类汇总表',
        'amort_sheet': '投资性房地产摊销分摊表',
        'inc_sheet': '投资性房地产增加检查表', 'dec_sheet': '投资性房地产减少检查表',
        'cost_names': ['投资性房地产'], 'accum_names': ['投资性房地产累计折旧', '投资性房地产累计摊销'],
        'impair_names': ['投资性房地产减值准备'], 'clear_names': [],
        'accum_kw': '折旧', 'amort': True, 'amort_src': 'accum', 'special': False,
    },
    'INT': {
        'title': '无形资产', 'paper': '无形资产审计底稿',
        'main_sheet': '无形资产分类汇总表',
        'amort_sheet': '无形资产摊销核对表',
        'inc_sheet': '无形资产增加检查表', 'dec_sheet': '无形资产减少检查表',
        'cost_names': ['无形资产'], 'accum_names': ['累计摊销'],
        'impair_names': ['无形资产减值准备'], 'clear_names': [],
        'accum_kw': '摊销', 'amort': True, 'amort_src': 'accum', 'special': False,
    },
    'ROU': {
        'title': '使用权资产', 'paper': '使用权资产审计底稿',
        'main_sheet': '使用权资产分类汇总表',
        'amort_sheet': '使用权资产折旧分摊表',
        'inc_sheet': '使用权资产增加检查表', 'dec_sheet': '使用权资产减少检查表',
        'cost_names': ['使用权资产'], 'accum_names': ['使用权资产折旧', '使用权资产累计折旧'],
        'impair_names': [], 'clear_names': [],
        'accum_kw': '折旧', 'amort': True, 'amort_src': 'accum', 'special': False,
    },
    'LTDEF': {
        'title': '长期待摊费用', 'paper': '长期待摊费用审计底稿',
        'main_sheet': '长期待摊费用分类汇总表',
        'amort_sheet': '长期待摊费用摊销核对表',
        'inc_sheet': '长期待摊费用增加检查表', 'dec_sheet': '长期待摊费用减少检查表',
        'cost_names': ['长期待摊费用', '递延资产'], 'accum_names': [],
        'impair_names': [], 'clear_names': [],
        'accum_kw': '摊销', 'amort': True, 'amort_src': 'cost', 'special': False,
    },
    'CIP': {
        'title': '在建工程', 'paper': '在建工程审计底稿',
        'main_sheet': '在建工程分类汇总表',
        'inc_sheet': '在建工程增加检查表', 'dec_sheet': '在建工程减少检查表',
        'cost_names': ['在建工程'], 'accum_names': [],
        'impair_names': [], 'clear_names': [],
        'accum_kw': None, 'amort': False, 'amort_src': None, 'special': True,
    },
    'ONCA': {
        # ⚡⚡ 2026-08-30 新增：其他非流动资产（1831 资产购置预付款 等），
        #   之前无生成器 → 科目无处安放被预付 kw='预付' 误吸收。并入 longterm 程序生成。
        # ⚡⚡ 2026-09-01 加『委托贷款』：AH 1010 委托贷款 4.4B（一年以上 3.58B）归入
        #   其他非流动资产（横展表/企业报表口径）；一年内到期委托贷款 0.84B 由匹配逻辑排除
        #   （归『一年内到期的非流动资产』，不混入其他非流动）。
        'title': '其他非流动资产', 'paper': '其他非流动资产审计底稿',
        'main_sheet': '其他非流动资产分类汇总表',
        'inc_sheet': '其他非流动资产增加检查表', 'dec_sheet': '其他非流动资产减少检查表',
        'cost_names': ['其他非流动资产', '委托贷款'], 'accum_names': [],
        'impair_names': [], 'clear_names': [],
        'accum_kw': None, 'amort': False, 'amort_src': None, 'special': False,
    },
}

# 二级分类展示优先顺序（按常见资产类别；其余按名称排序补在后）
CAT_ORDER_PREF = ['房屋及建筑物', '机械设备', '专用设备', '动力设备', '运输设备',
                  '办公设备', '其他设备', '其他']

# 增加/减少 子类型标签（按类别展示不同名称，内部 key 统一）
INC_LABELS = {
    'FA': ['购入', '在建工程转入', '内部转移', '其他'],
    'IPR': ['购入', '在建工程转入', '内部转移', '其他'],
    'INT': ['购入', '在建工程转入', '内部转移', '其他'],
    'ROU': ['购入', '在建工程转入', '内部转移', '其他'],
    'LTDEF': ['发生', '在建工程转入', '内部转移', '其他'],
    'CIP': ['发生', '在建工程转入', '内部转移', '其他'],
    'ONCA': ['发生', '在建工程转入', '内部转移', '其他'],
}
DEC_LABELS = {
    'FA': ['出售', '内部转移', '转入在建工程', '其他'],
    'IPR': ['出售', '内部转移', '转入在建工程', '其他'],
    'INT': ['出售', '内部转移', '转入在建工程', '其他'],
    'ROU': ['出售', '内部转移', '转入在建工程', '其他'],
    'LTDEF': ['摊销完毕/转出', '内部转移', '转入在建工程', '其他'],
    'CIP': ['转入固定资产', '转入无形资产', '转入长期待摊费用', '转入投资性房地产',
            '内部转移', '出售', '其他'],
    'ONCA': ['转出', '内部转移', '转入在建工程', '其他'],
}
INC_KEYS = ['purchase', 'inproject', 'internal', 'other']
DEC_KEYS_DEPR = ['sale', 'internal', 'to_inproject', 'other']
DEC_KEYS_CIP = ['to_fa', 'to_int', 'to_ltdef', 'to_ipr', 'internal', 'sale', 'other']

# 分类关键词
PURCHASE_KW = ['银行存款', '应付账款', '应付票据', '预付账款', '库存现金',
               '其他应付款', '在途物资', '材料采购', '委托加工物资', '工程物资', '应付职工薪酬']
INTERNAL_KW = ['内部转移', '内部往来', '关联方']
GAIN_ACCOUNTS = ['资产处置收益']
NONBIZ_INCOME = ['营业外收入']
NONBIZ_EXPENSE = ['营业外支出']


# ----------------------------------------------------------------------------
# 文件名解析 / 实体发现
# ----------------------------------------------------------------------------
def _split_ent_year(fn, marker):
    base = fn
    for ext in ('.xlsx', '.xls', '.xlsm', '.csv'):
        if base.lower().endswith(ext):
            base = base[:-len(ext)]
            break
    base = re.sub(r'_\d{12,}$', '', base)
    base = re.sub(r'_\d{8}$', '', base)
    m = re.search(r'(19|20)\d{2}', base)
    yr = m.group(0) if m else '2025'
    prefix, _, _suffix = base.partition(marker)
    ent = re.sub(r'(19|20)\d{2}', '', prefix)
    ent = re.sub(r'(?<!\d)(2\d)(?!\d)', '', ent)
    ent = re.sub(r'年度|年\d*[-~]?\d*月', '', ent)  # 去"年度"、"年1-3月"
    ent = ent.strip(' _-')
    # 个位数前缀补0（1��FY本级公司 → 01、FY本级公司）
    _m0 = re.match(r'^(\d)([、,，.．\s])', ent)
    if _m0 and len(ent) > 2:
        ent = '0' + ent
    return ent, yr


def _discover_entities(data_dir):
    # ⚡ 2026-08-09 治本统一：audit_common.discover_entities（U8 全量/SAP adapter 单主体过滤）
    # → 转本程序嵌套结构 {E: {'km': {y: path}, 'gl': {y: path}}}
    # ⚡ 2026-08-12 治理：改用 audit_common.norm_entities（幂等兼容 U8/SAP 双结构）。
    # ⚡⚡ 2026-08-27 P0 修复（1010 固定资产底稿全空根因）：原 `from audit_common import
    #   discover_entities as _de` 是 from-import 绑定，patch_audit_common() 只替换
    #   audit_common 模块属性，不影响已绑定引用 → SAP 场景 _de 仍为 U8 版（只认 .xls 旧格式，
    #   AH .xlsx 数据发现为空 → SKIP/空底稿）。改运行时属性查找，patch 后即 SAP 版。
    import audit_common as _AU
    return _AU.norm_entities(_AU.discover_entities(data_dir))


# ----------------------------------------------------------------------------
# 数据加载
# ----------------------------------------------------------------------------


def _read_gl_fa(path):
    if not path:
        return []
    # ⚡ SAP：gl 是多文件 list，走 adapter（忽略传入，用 current_comp）
    # 2026-08-10 修复：adapter 行 {name,date,vtype,vno,cp,sm,debit,credit} 无 'voucher'
    # → 转 U8 同构（voucher='字-号'），否则 _vouch_key 恒空 → 摊销核对列支全 0。
    if not isinstance(path, str):
        if _adapter is not None:
            # ⚡ 2026-08-12 修复：集团模式（current_comp=None）下无参 _adapter.read_gl()
            #   返回空 → 折旧/摊销分摊表集团版全 0。但【不能】在集团分支全量读 88 家
            #   GL 合成一表——_build_year_workbook 对每个主体调 _read_gl_fa(gl_path)，
            #   全量累积=88 倍重复 + MemoryError（10.8GB 实测爆）。改为从 gl_path
            #   （含 {comp} 路径）解析当前主体，只读该主体（与单主体同源，防 OOM）。
            _cc = _adapter.current_comp()
            if not _cc:
                # 从 gl_path（list[0] 路径）解析主体
                # ⚡⚡ 2026-08-28 修复（1357 多主体摊销表全 0 根因）：原【先匹配目录段
                #    [\\/](\d{4})[\\/]】——路径 d:/底稿测试/AH/数据/2026\序时账\1020-1月.xlsx
                #    中的 `/2026\`（数据根年份目录）先被匹配 → _comp_c='2026'（非主体）
                #    → read_gl('2026') 返回全量 GL（60万行），实体过滤后 0 行 →
                #    摊销/折旧分摊表多主体集团版全 0。修复：优先【basename 前缀】匹配
                #    （1020-1月.xlsx → 1020），仅 basename 匹配不到才回退目录段正则。
                _comp_c = None
                _p0 = path[0] if isinstance(path, list) and path else path
                _m2_c = re.match(r'^(\d{4})', os.path.basename(str(_p0)))
                if _m2_c:
                    _comp_c = _m2_c.group(1)
                else:
                    _m_c = re.search(r'[\\/](\d{4})[\\/]', str(_p0))
                    if _m_c:
                        _comp_c = _m_c.group(1)
                if _comp_c:
                    _r_rows = _adapter.read_gl(_comp_c)
                else:
                    _r_rows = []
            else:
                _r_rows = _adapter.read_gl()
            out = []
            for _r in _r_rows:
                _vt = str(_r.get('vtype') or '')
                _no = str(_r.get('vno') or '')
                _vk = ('%s-%s' % (_vt, _no)).strip('-') or _no
                # ⚡ 2026-08-12 修复（铁律108 结构敏感）：adapter read_gl_rows 返回
                # dr/cr（非 debit/credit）→ 原取 debit/credit 恒 0 → 折旧/摊销分摊表、
                # 增加/减少检查表金额全空（质检 ERROR：固定资产/无形资产/投房/使用权）。
                _db = float(_r.get('dr') if _r.get('dr') is not None else
                            (_r.get('debit') or 0.0))
                _cr = float(_r.get('cr') if _r.get('cr') is not None else
                            (_r.get('credit') or 0.0))
                out.append({'name': str(_r.get('name') or ''),
                            'voucher': _vk,
                            'date': str(_r.get('date') or ''),
                            'cp': str(_r.get('cp') or ''),
                            'sm': str(_r.get('sm') or ''),
                            'fr': str(_r.get('fr') or ''),   # ⚡ 2026-08-28 保留功能范围（折旧费拆分）
                            'debit': _db,
                            'credit': _cr})
            return out
        return []
    if not os.path.exists(path):
        return []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdr = None
    for r in rows[:6]:
        if r and any('科目名称' in str(c) for c in r if c):
            hdr = r
            break
    if hdr is None:
        return []
    idx = {}
    specs = [('name', ['科目名称']), ('date', ['日期']), ('vt', ['字']), ('no', ['号']),
             ('cp', ['对方科目']), ('sm', ['摘要']), ('db', ['借方金额', '借方']), ('cr', ['贷方金额', '贷方'])]
    for key, alts in specs:
        for i, c in enumerate(hdr):
            if c and any(a in str(c) for a in alts):
                idx[key] = i
                break
    out = []
    for r in rows:
        ni = idx.get('name', 0)
        if not r or ni >= len(r) or r[ni] is None:
            continue
        nm = str(r[ni]).strip()
        if nm == '科目名称' or not nm:
            continue
        vt = str(r[idx.get('vt', 2)] or '').strip()
        no = str(r[idx.get('no', 3)] or '').strip()
        voucher = '%s-%s' % (vt, no) if (vt or no) else ''
        dt = str(r[idx.get('date', 1)] or '').strip()
        cp = str(r[idx.get('cp', 4)] or '').strip()
        sm = str(r[idx.get('sm', 5)] or '').strip()
        db_i = idx.get('db', 6)
        cr_i = idx.get('cr', 7)
        db = float(r[db_i]) if db_i < len(r) and r[db_i] not in (None, '') else 0.0
        cr = float(r[cr_i]) if cr_i < len(r) and r[cr_i] not in (None, '') else 0.0
        out.append({'name': nm, 'voucher': voucher, 'date': dt, 'cp': cp, 'sm': sm,
                    'debit': db, 'credit': cr})
    return out


def _ym(date_str):
    if not date_str:
        return None
    for sep in ('-', '/', '.'):
        if sep in date_str:
            parts = date_str.split(sep)
            if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                return (int(parts[0]), int(parts[1]))
    return None


# ----------------------------------------------------------------------------
# 按科目【名称】识别科目组与角色
# ----------------------------------------------------------------------------
def _grp_cat(name):
    """返回 (gkey, role)；role ∈ {cost, accum, impair, clear}。
    铁律：科目名以【实际账套科目名】为准——先按【组名】归属（名称含组名即属该组），
    再按【角色关键词】判断（折旧/摊销→备抵、减值准备→减值、清理→清理、其余→原值）；
    不含组名时按通用关键词兜底（累计折旧→FA、累计摊销→INT、减值准备/清理→FA）。
    故"使用权资产折旧"/"使用权资产累计折旧"/"固定资产累计折旧"/"累计折旧-房屋"等
    名称变体均可正确归类，不依赖写死的精确科目名。"""
    n = (name or '').strip()
    # ① 含组名优先归属；角色判定受【组角色能力】约束：
    #    组未配置备抵/减值/清理科目时，"折旧/摊销/减值准备/清理"不构成该角色
    #    （如"在建工程-土地使用权摊销"属 CIP cost，非 CIP 备抵，避免被误当计提来源）。
    #    摊销备抵须含"累计"语义（无形资产累计摊销/累计摊销），"管理费用-无形资产摊销"
    #    等费用/去向科目虽含组名+摊销但不属备抵，返回 None 由费用列支逻辑处理。
    for gkey, sp in GROUPS.items():
        gname = sp.get('title') or ''
        if gname and gname in n:
            # 名称含组名但属费用/去向科目（管理费用-无形资产摊销 等费用科目），
            # 不应归入资产组；注：『在建工程』不得列入去向关键词——
            # 否则"在建工程-检测费"等在建工程自身科目行会被误判为去向科目而全部排除
            # （2026-07-31 修复：FY 在建工程增加/减少检查表取数全空的根因）。
            _go_kw = ('管理费用', '制造费用', '销售费用', '研发费用', '主营业务成本',
                      '其他业务成本', '其他业务支出', '营业费用', '生产成本', '劳务成本',
                      '应交税费', '应付账款', '其他应付款')
            _is_go = any(k in n for k in _go_kw) or ('待摊费用' in n and not n.startswith('长期待摊'))
            if _is_go:
                return (None, None)
            # ⚡ 2026-08-10 SAP 独立费用科目：名称含"摊销/折旧"动词但【无"累计"】者
            # 即费用/去向科目（SAP 1010『研究开发费-无形资产摊销』『折旧费』等无
            # "管理费用-"前缀，_go_kw 匹配不到）→ 一律不归资产组，交由费用列支逻辑。
            # 资产原值科目名（无形资产-土地/软件）不含摊销/折旧动词，不受影响；
            # 备抵（无形资产累计摊销/累计折旧）含"累计"→ 落到下方 accum 分支。
            if (('摊销' in n and '累计' not in n) or ('折旧' in n and '累计' not in n)):
                return (None, None)
            if '清理' in n and sp.get('clear_names'):
                return (gkey, 'clear')
            if '减值准备' in n and sp.get('impair_names'):
                return (gkey, 'impair')
            if ('折旧' in n or ('摊销' in n and '累计' in n)) and sp.get('accum_names'):
                return (gkey, 'accum')
            return (gkey, 'cost')
    # ② 不含组名的通用关键词兜底：仅认【累计*】形态的备抵科目
    #    （"管理费用-无形资产摊销"等费用科目含"摊销"但非备抵，不归 accum）
    if '清理' in n:
        return ('FA', 'clear')
    if n.startswith('累计摊销'):
        return ('INT', 'accum')
    if n.startswith('累计折旧'):
        return ('FA', 'accum')
    if '减值准备' in n:
        return ('FA', 'impair')
    return (None, None)


def _find_group_codes(km, gkey):
    """从《科目余额表》【动态发现】该组 原值/备抵/减值 实际科目代码。
    铁律：名称以实际账套科目名为准——按 组名+角色关键词 语义识别
    （"使用权资产折旧""固定资产累计折旧""累计摊销"等变体均可命中），
    不依赖写死的精确名称列表。"""
    sp = GROUPS[gkey]
    gname = sp.get('title') or ''
    # 2026-08-06：名称匹配用 cost_names（别名集）而非仅 title——
    # H 账套 1801『递延资产』=长期待摊费用（旧准则名），title 匹配不到 → 底稿静默缺失。
    _gnames = [n for n in (sp.get('cost_names') or []) if n] or [gname]
    cost_cands = []
    accum_cands = []
    impair_cands = []
    for c, d in km.items():
        nm = (d.get('name') or '').strip()
        if not nm:
            continue
        if any(_gn in nm for _gn in _gnames) and '清理' not in nm \
                and not (gkey == 'ONCA' and '一年内到期' in nm):
            # ⚡⚡ 2026-09-01 ONCA 排除『一年内到期的委托贷款』（归一年内到期的非流动资产）
            # 费用/去向科目（管理费用-无形资产摊销 等）虽含组名，但属 P&L 费用列支，绝不构成资产原值。
            # 名称含「摊销」但无「累计」语义 → 费用科目（无形资产摊销/固定资产折旧等）；父级路径在 TB 可能缺失
            # （如 6605.01.04 名称仅"无形资产摊销"，不带"管理费用"前缀），故需按摊销/折旧动词直接排除。
            _is_go = any(k in nm for k in ('管理费用', '制造费用', '销售费用', '研发费用',
                                            '其他业务成本', '其他业务支出', '营业费用', '生产成本', '劳务成本',
                                            '应交税费', '应付账款', '其他应付款',
                                            # ⚡⚡ 2026-08-15 收入类科目（含组名但非资产原值）：
                                            #   XBJ 46 主体『其他业务收入\固定资产出租』(605101) 曾混入
                                            #   FA cost_cands → _p4={1601,6051} → 虚拟父级不触发 → 单选
                                            #   160101 土地资产漏 1.14 亿。收入科目绝不构成资产原值。
                                            '主营业务收入', '其他业务收入', '营业外收入'))
            _is_amort_exp = (('摊销' in nm and '累计' not in nm) or ('折旧' in nm and '累计' not in nm))
            # 2026-08-07 修复：其他资产组的备抵名不得误入本组——FA 组若只含『使用权资产
            # 累计折旧』（1642，上分/北分/数链 无 1602 时）会被『折旧 in nm』误收 →
            # FA 折旧行名变『使用权资产累计折旧』。按组专属前缀排除：
            # FA 不收 使用权/投资性/无形 前缀；INT 不收 使用权/投资性/固定 前缀。
            _other_pref = {'FA': ('使用权资产', '投资性房地产', '无形资产'),
                           'INT': ('使用权资产', '投资性房地产', '固定资产'),
                           'ROU': ('投资性房地产', '无形资产', '固定资产'),
                           'IPR': ('使用权资产', '无形资产', '固定资产')}.get(gkey, ())
            _is_other = any(p in nm for p in _other_pref)
            # ⚡ 2026-08-09 修复：impair_names 配置了具体名称时必须【名称前缀精确匹配】——
            # 原 '减值准备' in nm 会把任意减值科目（如长期股权投资减值准备 1512）误挂固定资产组
            # （SAP 1010 实测：8100 万长期股权投资减值被误作固定资产减值）。CIP 组无 impair_names 配置，
            # 仍按名称含"减值准备"兜底。
            _impair_hit = (gkey == 'CIP' or
                           any(nm.startswith(pn) for pn in (sp.get('impair_names') or [])))
            if '减值准备' in nm and _impair_hit:
                impair_cands.append(c)
            elif ('折旧' in nm or ('摊销' in nm and '累计' in nm)) and sp.get('accum_names') and not _is_other:
                accum_cands.append(c)
            elif _is_go or _is_amort_exp:
                continue   # 费用科目（无形资产摊销/固定资产折旧等）不构成原值候选
            else:
                cost_cands.append(c)
        else:
            # 不含组名的通用备抵/减值（如 一级 累计折旧/累计摊销/固定资产减值准备）
            # 2026-08-07 修复②：通用分支同样排除其他组前缀——上分/北分/数链 无 1602
            # 时『使用权资产累计折旧』(1642) 含"累计折旧"被 FA 误收（else 分支 395 行）。
            _other_pref2 = {'FA': ('使用权资产', '投资性房地产', '无形资产'),
                            'INT': ('使用权资产', '投资性房地产', '固定资产'),
                            'ROU': ('投资性房地产', '无形资产', '固定资产'),
                            'IPR': ('使用权资产', '无形资产', '固定资产')}.get(gkey, ())
            _is_other2 = any(p in nm for p in _other_pref2)
            if gkey == 'FA' and '累计折旧' in nm and '减值' not in nm and not _is_other2:
                accum_cands.append(c)
            elif gkey == 'INT' and '累计摊销' in nm and not _is_other2:
                accum_cands.append(c)
            elif gkey == 'FA' and '减值准备' in nm and not any(
                    k in nm for k in ('长期股权投资', '坏账', '存货', '应收', '在建工程',
                                      '无形资产', '投资性房地产', '使用权资产')):
                impair_cands.append(c)
    # 同名多候选时（如 制造费用下的“固定资产” 与 真实 1601 固定资产）：
    # 优选 二级明细最多者（真实资产科目挂二级），并列取代码最短者（一级科目通常 4 位）。
    def _pick(cands):
        if not cands:
            return None
        # ⚡⚡ 2026-08-15 无父级平行子目：XBJ 46 主体 160101-160108（土地/房屋/机器…）
        #   无 1601 父级行、各子目均无下挂明细 → 原单选取代码最短 '160101'（土地资产）
        #   → 底稿原值只有土地 7.27 亿，漏房屋/机器等 1.14 亿（TB 合计 8.41 亿）。
        #   返回共同 4 位前缀作『虚拟父级』，由 _prep_group 按前缀聚合（原值=Σ 全部分行）。
        if len(cands) > 1:
            _p4 = {str(c)[:4] for c in cands}
            _all_leaf = not any(k != c and str(k).startswith(str(c)) for c in cands for k in km)
            if len(_p4) == 1 and _all_leaf:
                return _p4.pop()
        return max(cands, key=lambda c: (sum(1 for k in km if k.startswith(c) and k != c), -len(c)))
    cost = _pick(cost_cands)
    accum = _pick(accum_cands)
    impair = _pick(impair_cands)
    return cost, accum, impair


def _discover_accum_name(recs, gkey):
    """从各主体《科目余额表》【动态发现】该组实际累计折旧/摊销科目名（一级）。
    铁律：科目名以实际账套为准（如"使用权资产折旧"）；未发现时回退配置通用名。"""
    sp = GROUPS[gkey]
    verb = '摊销' if gkey in ('INT', 'LTDEF') else '折旧'
    for _ent, _rec in (recs or []):
        _km = _rec.get('km') or {}
        _acc = _find_group_codes(_km, gkey)[1]
        if _acc:
            _nm = ((_km.get(_acc) or {}).get('name') or '').strip()
            _base = (_nm.split('-')[0] if '-' in _nm else _nm).strip()
            if _base:
                return _base
    if sp.get('accum_names'):
        return sp['accum_names'][0]
    return '累计%s' % verb


def _top_account(name):
    return (name or '').split('-')[0].strip()


# 长期待摊费用摊销：借费用科目 / 贷长期待摊费用，贷方即摊销数。这些常规摊销
# 不应作为「减少/处置」查证，故在减少明细表中剔除对方科目为费用类的贷方分录。
_AMORT_EXPENSE_PREFIXES = ('管理费用', '制造费用', '销售费用', '研发费用', '研发支出',
                           '生产成本', '劳务成本', '主营业务成本', '其他业务成本',
                           '业务及管理费')


def _is_amort_expense_cp(cp):
    """长期待摊费用的常规摊销（借:费用科目 贷:长期待摊费用）不计入「减少/处置」查证。
    对方科目识别：①费用类一级前缀；②对方科目含"摊销"（如 FY 嵊州用独立『摊销费用』科目，
    2026-07-31 修复：此前未识别导致其摊销被误计入处置抽凭）。"""
    cp = (cp or '').strip()
    if not cp:
        return False
    top = _top_account(cp)
    return top.startswith(_AMORT_EXPENSE_PREFIXES) or '摊销' in top


def _split_cp(cp):
    """对方科目字符串分割（兼容 逗号/分号/全角——AFJ 等 U8 导出混用分隔符）。
    ⚡ 不含顿号（、）——『在建工程-建筑、安装工程』等科目名内部含顿号，误分割会拆散科目。"""
    import re as _re
    return [p.strip() for p in _re.split(r'[,;；，]', cp or '') if p.strip()]


# ⚡⚡ 2026-08-23 增加检查表对方科目筛选（用户方法论）：凭证级多科目只保留真实来源。
#   定义已上移至 audit_common（INC_CP_KEEP/INC_CP_DROP/filter_inc_cp），builder 与终审补全共用。
from audit_common import filter_inc_cp as _filter_inc_cp


def _is_same_group(gkey, cp):
    """对方科目中任一科目属于同一长期资产组（任意角色）-> 组内重分类。"""
    cp = (cp or '').strip()
    if not cp:
        return False
    for part in cp.split(','):
        part = part.strip()
        if not part:
            continue
        gg, _ = _grp_cat(part)
        if gg == gkey:
            return True
    return False


def _is_intragroup_cost_transfer(gkey, cp):
    """对方科目中任一科目属于同一组且为成本(原值)类 -> 组内原值互转(借A贷B)，非真实增减。"""
    cp = (cp or '').strip()
    if not cp:
        return False
    for part in cp.split(','):
        part = part.strip()
        if not part:
            continue
        gg, rl = _grp_cat(part)
        if gg == gkey and rl == 'cost':
            return True
    return False


# ----------------------------------------------------------------------------
# 增加/减少 分类
# ----------------------------------------------------------------------------
def _classify_inc(gkey, cp, sm, role='cost'):
    parts = [p.strip() for p in (cp or '').split(',') if p.strip()]
    if any('在建工程' in p for p in parts):
        return 'inproject'
    if ('内部转移' in (sm or '')) or any(k in p for p in parts for k in INTERNAL_KW):
        return 'internal'
    if role in ('accum', 'impair'):
        return 'purchase'  # 计提数 / 摊销数
    if any(k in p for p in parts for k in PURCHASE_KW):
        return 'purchase'
    return 'other'


def _classify_dec(gkey, cp, sm):
    parts = [p.strip() for p in (cp or '').split(',') if p.strip()]
    if gkey == 'CIP':
        if any('固定资产' in p for p in parts):
            return 'to_fa'
        if any('无形资产' in p for p in parts):
            return 'to_int'
        if any('长期待摊费用' in p for p in parts):
            return 'to_ltdef'
        if any('投资性房地产' in p for p in parts):
            return 'to_ipr'
        if ('内部转移' in (sm or '')) or any('内部转移' in p for p in parts):
            return 'internal'
        if any('固定资产清理' in p for p in parts):
            return 'sale'
        return 'other'
    if any('固定资产清理' in p for p in parts):
        return 'sale'
    if ('内部转移' in (sm or '')) or any('内部转移' in p for p in parts):
        return 'internal'
    if any('在建工程' in p for p in parts):
        return 'to_inproject'
    return 'other'


def _grp_breakdown(gl, gkey, role):
    inc = {k: 0.0 for k in INC_KEYS}
    dec_keys = DEC_KEYS_CIP if gkey == 'CIP' else DEC_KEYS_DEPR
    dec = {k: 0.0 for k in dec_keys}
    for r in gl:
        g, rl = _grp_cat(r['name'])
        if g != gkey or rl != role:
            continue
        if role == 'cost':
            if r['debit'] > 0:
                inc[_classify_inc(gkey, r['cp'], r['sm'], 'cost')] += r['debit']
            if r['credit'] > 0:
                dec[_classify_dec(gkey, r['cp'], r['sm'])] += r['credit']
        else:  # accum / impair：增加=贷, 减少=借
            if r['credit'] > 0:
                inc[_classify_inc(gkey, r['cp'], r['sm'], role)] += r['credit']
            if r['debit'] > 0:
                dec[_classify_dec(gkey, r['cp'], r['sm'])] += r['debit']
    return inc, dec


# ----------------------------------------------------------------------------
# Excel 辅助
# ----------------------------------------------------------------------------
def _hdr(ws, row, headers, start_col=1):
    for j, h in enumerate(headers, start_col):
        c = ws.cell(row, j, h)
        c.fill = SHELL_HFILL
        c.font = SHELL_HFONT
        c.border = SHELL_BORDER
        c.alignment = SHELL_CEN
    return row + 1


def _money(ws, r, c, v):
    if v is not None:
        try:
            v = round(float(v), 2)
        except (TypeError, ValueError):
            pass
    cell = ws.cell(r, c, v)
    cell.number_format = SHELL_NUM
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    return cell


def _txt(ws, r, c, v, bold=False, align=None, fill=None):
    cell = ws.cell(r, c, v)
    cell.border = SHELL_BORDER
    cell.alignment = align or SHELL_LEFT
    if bold:
        cell.font = SHELL_BOLD
    if fill:
        cell.fill = fill
    return cell


def _title(ws, text, ncols, row=1):
    ws.cell(row, 1, text).font = SHELL_TITLE_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(ncols, 1))
    return row + 1


def _sub(ws, text, ncols, row=2):
    c = ws.cell(row, 1, text)
    c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(ncols, 1))
    return row + 1


def _col_letter(idx):
    return get_column_letter(idx)


def _set_formula(ws, r, col, formula):
    cell = ws.cell(r, col, formula)
    cell.number_format = SHELL_NUM
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    return cell


# ----------------------------------------------------------------------------
# Sheet: 附注披露（2026-08-02 用户方法论：所有科目最后加 1 张附注汇总表）
# 每主体一列；A=类别、B=段名；块=未审数 / 审计调整 / 期末审定数；
# 组件=一、原值 / 二、累计折旧(摊销) / 三、减值准备 / 四、账面价值；
# 段=期初数 / 本期增加 / 本期减少 / 期末数（账面价值仅 期初数 / 期末数）。
# 要体现增减变动的科目一律用固定资产模板（含长期待摊费用等）；
# 三平行表（未审/调整/审定）仅用于利润表类简单披露（只需上年数/本年数）。
# ----------------------------------------------------------------------------
def build_footnote_sheet(wb, year, recs, gkey, sp, all_ents=None):
    """长期资产附注披露表（固定资产模板，2026-08-02 用户方法论）。
    每主体一列；A=类别、B=段名；块=未审数/审计调整/期末审定数；
    组件=一、原值/二、累计折旧(摊销)/三、减值准备/四、账面价值；
    段=期初数/本期增加/本期减少/期末数（账面价值仅期初数/期末数）。
    要体现增减变动的科目一律用本模板（含长期待摊费用等）；
    三平行表（未审/调整/审定）仅用于利润表类简单披露（只需上年数/本年数）。
    all_ents（2026-08-03）：传入全部主体名（含无数据主体，列示 0），默认只用 recs 里有数据的主体。"""
    ws = wb.create_sheet('附注汇总')
    is_cip = sp.get('special')
    accum_label = _discover_accum_name(recs, gkey) if not sp.get('special') else None
    accum_label = accum_label or '累计折旧'
    ent_list = all_ents if all_ents else [e for (e, _) in recs]
    n_ent = len(ent_list)
    # —— 收集每主体的二级类别数据：{ent: {cat: {'cost':[o,i,d,c], 'accum':[...], 'impair':[...]}}} ——
    ent_cat = {}
    for ent, rec in recs:
        cd = {}

        def _put(role, code, name, o, i, d, c):
            # 2026-08-03：类别归并到二级——按代码长度（6位/1个点=二级）取类名，三级及以上归并到其二级父类。
            # 名称可能带前缀（固定资产-交通运输设备）或不带（汽车），统一按代码层级归并：
            #   二级码（1601.03/160103）→ 类别=去一级前缀后的首段名；
            #   三级码（1601.04.01）→ 找其二级父码（1601.04）的类名，归并进去。
            cs = str(code).replace('.', '')
            if len(cs) <= 6:
                cat = (name.split('-', 1)[1] if '-' in name else name).split('-')[0].strip()
            else:
                # 三级及以上：由调用方预先把三级归并到二级（见下方 l2_code 映射），此处兜底取首段
                cat = (name.split('-', 1)[1] if '-' in name else name).split('-')[0].strip()
            b = cd.setdefault(cat, {'cost': [0.0] * 4, 'accum': [0.0] * 4, 'impair': [0.0] * 4})
            b[role][0] += o
            b[role][1] += i
            b[role][2] += d
            b[role][3] += c

        # 2026-08-03：三级及以上码 → 归并到其二级父类（码取前 6 位/首个点前段）
        def _put_code(role, code, dd):
            """按码归并：只取二级行（6 位码），三级及以上跳过——U8 父级发生额=子目之和，
            三级值已含在二级父行内，再加即双计（FY本级 土地房屋 673M×2 曾双计 1.35B）。
            2026-08-06 修复：accum/impair（备抵）opening/closing 取 abs——泰国 1420 累计折旧
            二级行 closing 为负（贷方），账面价值=原值−(−折旧)=原值+折旧 → 附注双计
            （458.95M vs 净值 371.3M）。与下方 rec['accum_d'] 非二级路径 abs(_disp(...)) 口径一致。"""
            cs = str(code).replace('.', '')
            # ⚡ 2026-08-10 SAP 码长适配：U8 二级=6 位（170101），SAP 二级=10 位
            # （1701010000）→ 统一只收 6/10 位二级，其余（U8 三级 8 位/SAP 三级 12 位）
            # 跳过（父级发生额=子目之和，三级值已含在二级父行内，再加即双计）。
            if len(cs) not in (6, 10):
                return
            cat = (dd['name'].split('-', 1)[1] if '-' in dd['name'] else dd['name']).split('-')[0].strip()
            if role in ('accum', 'impair'):
                # 2026-08-07 修复：备抵二级类别名≠原值二级名（dq 160201『房屋折旧』vs
                # 160101『房屋及建筑物』；代码一一对应但名称不同）→ 原逻辑 12 个类别全
                # 列出 → 原值块混入『房屋折旧』行、折旧块混入『房屋及建筑物』行。
                # 归并：备抵类别名去『折旧/摊销/减值准备』词根后，与已收集的原值类别匹配
                # （base 含于原值类 或 原值类含于 base → 取原值类；无匹配保留原名兜底）。
                _base = cat.replace('折旧', '').replace('摊销', '').replace('减值', '').replace('准备', '').strip()
                if _base:
                    for _c0 in cd:
                        if _c0 == _base or (_base and (_base in _c0 or _c0 in _base)):
                            cat = _c0
                            break
            b = cd.setdefault(cat, {'cost': [0.0] * 4, 'accum': [0.0] * 4, 'impair': [0.0] * 4})
            if role in ('accum', 'impair'):
                b[role][0] += abs(dd['opening'])
                # ⚡ 2026-08-10 备抵方向修正：备抵科目【贷方=计提(增加)、借方=转回/转销
                # (减少)】——原 debit→增加/credit→减少 与分类汇总表（credit=增加）相反，
                # SAP 1010 附注曾把累计摊销计提 4,828,530.26 显示为"本期减少"。
                b[role][1] += dd.get('credit', 0.0)
                b[role][2] += dd.get('debit', 0.0)
                b[role][3] += abs(dd['closing'])
            else:
                b[role][0] += dd['opening']
                b[role][1] += dd.get('debit', 0.0)
                b[role][2] += dd.get('credit', 0.0)
                b[role][3] += dd['closing']

        if rec['cost_l2']:
            for c, dd in rec['cost_l2']:
                _put_code('cost', c, dd)
        else:
            cd.setdefault(sp['title'], {'cost': [0.0] * 4, 'accum': [0.0] * 4, 'impair': [0.0] * 4})
            if rec['cost_d']:
                cd[sp['title']]['cost'] = [_disp(rec['cost_d'], 'cost', 'opening'),
                                           _acc_inc(rec['cost_d'], 'cost'),
                                           _acc_dec(rec['cost_d'], 'cost'),
                                           _disp(rec['cost_d'], 'cost', 'closing')]
        if not is_cip and rec.get('accum_l2'):
            for c, dd in rec['accum_l2']:
                _put_code('accum', c, dd)
        elif not is_cip and rec['accum_d']:
            cd.setdefault(sp['title'], {'cost': [0.0] * 4, 'accum': [0.0] * 4, 'impair': [0.0] * 4})
            cd[sp['title']]['accum'] = [abs(_disp(rec['accum_d'], 'accum', 'opening')),
                                        _acc_inc(rec['accum_d'], 'accum'),
                                        _acc_dec(rec['accum_d'], 'accum'),
                                        abs(_disp(rec['accum_d'], 'accum', 'closing'))]
        if rec['impair_d']:
            cd.setdefault(sp['title'], {'cost': [0.0] * 4, 'accum': [0.0] * 4, 'impair': [0.0] * 4})
            cd[sp['title']]['impair'] = [abs(_disp(rec['impair_d'], 'impair', 'opening')),
                                         _acc_inc(rec['impair_d'], 'impair'),
                                         _acc_dec(rec['impair_d'], 'impair'),
                                         abs(_disp(rec['impair_d'], 'impair', 'closing'))]
        # 2026-08-02：在建工程也可有独立减值准备（_find_group_codes 对 CIP 已识别）——二级明细一并收集
        if not is_cip and rec.get('impair_l2'):
            for c, dd in rec['impair_l2']:
                _put_code('impair', c, dd)
        ent_cat[ent] = cd
    # —— 类别全集与顺序（2026-08-03 修复：PREF 是排序优先级而非白名单——
    #    旧逻辑把"含 PREF 关键词但不精确等于"的类别（如 房屋装修）两段都不命中→类别丢失→附注未审数全 0。
    #    新逻辑：全部类别保留，按 PREF 首匹配项排序，未匹配的按原序排后。）
    all_cats = []
    for cd in ent_cat.values():
        for cat in cd:
            if cat not in all_cats:
                all_cats.append(cat)

    PREF = ('房屋', '建筑物', '专用设备', '通用设备', '运输工具', '电子', '办公设备', '机器设备',
            '器具', '工具', '家具', '装修', '管道', '设计', '软件', '土地', '在建')

    def _cat_rank(c):
        for i, p in enumerate(PREF):
            if p in c:
                return i
        return len(PREF)

    cats = sorted(all_cats, key=_cat_rank)
    if sp['title'] in cats:
        cats.remove(sp['title'])
        cats.append(sp['title'])
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出『全集团合计』列
    _single_fn = len(ent_list) <= 1
    ncols = 3 + n_ent - (1 if _single_fn else 0)   # A类别 + B段 + 主体… (+合计, 单体省略)
    # —— 标题 + 表头（2026-08-02：附注汇总不设合并单元格，标题仅 A1 加粗）——
    _txt(ws, 1, 1, '%s附注披露（%s年度）' % (sp['title'], year), bold=True, align=SHELL_LEFT)
    _txt(ws, 2, 1, '项目', bold=True, fill=SHELL_HFILL, align=SHELL_CEN)
    _txt(ws, 2, 2, '项目', bold=True, fill=SHELL_HFILL, align=SHELL_CEN)
    for i, ent in enumerate(ent_list, 3):
        _txt(ws, 2, i, ent, bold=True, fill=SHELL_HFILL, align=SHELL_CEN)
    if not _single_fn:
        _txt(ws, 2, ncols, '全集团合计', bold=True, fill=SHELL_HFILL, align=SHELL_CEN)
    ws.freeze_panes = 'C3'
    r = 3
    # —— 组件规格（在建工程：账面余额 + 减值准备(有数据时) + 账面价值；其余 原值/折旧/减值/账面价值）——
    if is_cip:
        _has_cip_impair = any(rec.get('impair') for _e, rec in recs)
        comps_spec = [('cost', '一、账面余额', (('期初数', 0), ('本期增加', 1), ('本期减少', 2), ('期末数', 3)))]
        if _has_cip_impair:
            comps_spec += [('impair', '二、减值准备', (('期初数', 0), ('本期增加', 1), ('本期减少', 2), ('期末数', 3))),
                           ('net', '三、账面价值', (('期初数', 0), ('期末数', 3)))]
    else:
        comps_spec = [
            ('cost', '一、原值', (('期初数', 0), ('本期增加', 1), ('本期减少', 2), ('期末数', 3))),
            ('accum', '二、' + accum_label, (('期初数', 0), ('本期增加', 1), ('本期减少', 2), ('期末数', 3))),
            ('impair', '三、减值准备', (('期初数', 0), ('本期增加', 1), ('本期减少', 2), ('期末数', 3))),
            ('net', '四、账面价值', (('期初数', 0), ('期末数', 3))),
        ]

    def _cip_transfer(ent):
        """在建工程「转入固定资产」：GL 贷方在建工程、对方科目(cp)含『固定资产』且不含『清理』
        （2026-08-02 用户方法论；对方科目为空账套无法识别，返回 0 并在表尾披露限制）。"""
        tot = 0.0
        for _r in (_recs_by_ent.get(ent, {}).get('gl') or []):
            nm = str(_r.get('name') or '')
            if not nm.startswith('在建工程'):
                continue
            if float(_r.get('credit') or 0) <= 0:
                continue
            cp = str(_r.get('cp') or '')
            if '固定资产' in cp and '清理' not in cp:
                tot += float(_r['credit'])
        return tot

    _recs_by_ent = {e: rec for e, rec in recs}

    def _ent_val(ent, cat, role, idx):
        b = ent_cat.get(ent, {}).get(cat)
        if not b:
            return 0.0
        return b[role][idx]

    def _row_val(ent, comp, cat, idx):
        """该主体 组件/类别/段 的值；账面价值 = 原值 − 折旧 − 减值。"""
        if comp == 'net':
            return (_ent_val(ent, cat, 'cost', idx)
                    - _ent_val(ent, cat, 'accum', idx)
                    - _ent_val(ent, cat, 'impair', idx))
        return _ent_val(ent, cat, comp, idx)

    def _block_header(txt):
        """块头（未审数/审计调整/期末审定数、一、原值…）：A 列加粗提示，不合并跨行（2026-08-02）。"""
        nonlocal r
        _txt(ws, r, 1, txt, bold=True, fill=SHELL_TOT_FILL)
        r += 1

    def _write_block(label, mode, ref):
        """写一块：块头 + 各组件（组件头 + 各段 × 类别行 + 小计行）。
        mode='unaudit' 填数值；'adj' 全 0（留待填）；'audit' 公式=未审+调整。
        ref = {'unaudit': comp_rows, 'adj': comp_rows}（audit 用）。
        返回 {comp: {seg: {cat: row}}}（含 '小计'）。"""
        nonlocal r
        _block_header(label)
        comp_rows = {}
        for comp, comp_label, segs in comps_spec:
            _block_header(comp_label)
            seg_rows = {}
            for seg, idx in segs:
                cat_rows = {}
                cat_row_nums = []   # 类别数据行号（小计公式引用，2026-08-03 公式化）
                for cat in cats:
                    row = r
                    cat_rows[cat] = row
                    cat_row_nums.append(row)
                    _txt(ws, row, 1, cat)
                    _txt(ws, row, 2, seg)
                    if mode == 'audit':
                        u = ref['unaudit'][comp][seg][cat]
                        # 2026-08-06 协议：审定数行写数值（=未审+调整0；公式 data_only 读 None 收回读不回）
                        for i in range(3, ncols + 1):
                            _v = ws.cell(u, i).value if isinstance(ws.cell(u, i).value, (int, float)) else 0.0
                            _money(ws, row, i, _v)
                    else:
                        for i, ent in enumerate(ent_list, 3):
                            v = _row_val(ent, comp, cat, idx) if mode == 'unaudit' else 0.0
                            _money(ws, row, i, v)
                        if not _single_fn:
                            tot = sum(_row_val(ent, comp, cat, idx) for ent in ent_list) if mode == 'unaudit' else 0.0
                            _money(ws, row, ncols, tot)
                    r += 1
                # 小计行（2026-08-06 协议：写数值=类别行之和；公式 data_only 读 None 收回读不回）
                row = r
                cat_rows['小计'] = row
                _txt(ws, row, 1, '小  计', bold=True)
                _txt(ws, row, 2, seg, bold=True)
                if mode == 'audit':
                    u = ref['unaudit'][comp][seg]['小计']
                    for i in range(3, ncols + 1):
                        _v = ws.cell(u, i).value if isinstance(ws.cell(u, i).value, (int, float)) else 0.0
                        _money(ws, row, i, _v)
                else:
                    for i in range(3, ncols + 1):
                        if cat_row_nums:
                            _sv = sum(ws.cell(x, i).value for x in cat_row_nums
                                      if isinstance(ws.cell(x, i).value, (int, float)))
                            _money(ws, row, i, _sv)
                        else:
                            _money(ws, row, i, 0.0)
                r += 1
                # 在建工程「本期减少」段：其中：转入固定资产（2026-08-02 用户方法论，参考上市公司在建工程附注变动表）
                if is_cip and comp == 'cost' and seg == '本期减少':
                    row = r
                    cat_rows['其中：转入固定资产'] = row
                    _txt(ws, row, 1, '其中：转入固定资产')
                    _txt(ws, row, 2, seg)
                    if mode == 'audit':
                        u = ref['unaudit'][comp][seg]['其中：转入固定资产']
                        for i in range(3, ncols + 1):
                            _v = ws.cell(u, i).value if isinstance(ws.cell(u, i).value, (int, float)) else 0.0
                            _money(ws, row, i, _v)
                    else:
                        for i, ent in enumerate(ent_list, 3):
                            v = _cip_transfer(ent) if mode == 'unaudit' else 0.0
                            _money(ws, row, i, v)
                        if not _single_fn:
                            _money(ws, row, ncols,
                                   sum(_cip_transfer(ent) for ent in ent_list) if mode == 'unaudit' else 0.0)
                    r += 1
                seg_rows[seg] = cat_rows
            comp_rows[comp] = seg_rows
        return comp_rows

    rows_unaudit = _write_block('未审数', 'unaudit', None)
    rows_adj = _write_block('审计调整', 'adj', None)
    rows_audit = _write_block('期末审定数', 'audit', {'unaudit': rows_unaudit, 'adj': rows_adj})
    # 审定数块末尾：集团合计数 / 合并抵消借方 / 合并抵消贷方 / 合并报表数（2026-08-02 用户方法论）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出集团合计 4 行
    _net_comp = 'cost' if is_cip else 'net'
    _sm_row = rows_audit.get(_net_comp, {}).get('期末数', {}).get('小计')
    if _sm_row and len(ent_list) > 1:
        rg0 = r
        _txt(ws, r, 1, '集团合计数', bold=True, fill=SHELL_TOT_FILL)
        _txt(ws, r, 2, '期末数', bold=True, fill=SHELL_TOT_FILL)
        for i in range(3, ncols + 1):
            # 2026-08-06 协议：集团合计数写数值（=小计行值；公式 data_only 读 None）
            _v = ws.cell(_sm_row, i).value if isinstance(ws.cell(_sm_row, i).value, (int, float)) else 0.0
            _money(ws, r, i, _v)
        r += 1
        _txt(ws, r, 1, '合并抵消借方')
        _txt(ws, r, 2, '期末数')
        for i in range(3, ncols + 1):
            _money(ws, r, i, 0.0)
        r += 1
        _txt(ws, r, 1, '合并抵消贷方')
        _txt(ws, r, 2, '期末数')
        for i in range(3, ncols + 1):
            _money(ws, r, i, 0.0)
        r += 1
        _txt(ws, r, 1, '合并报表数', bold=True, fill=SHELL_TOT_FILL)
        _txt(ws, r, 2, '期末数', bold=True, fill=SHELL_TOT_FILL)
        for i in range(3, ncols + 1):
            # 2026-08-06 协议：合并报表数写数值（=集团合计，抵消 0；公式 data_only 读 None）
            _v = ws.cell(rg0, i).value if isinstance(ws.cell(rg0, i).value, (int, float)) else 0.0
            _money(ws, r, i, _v)
        r += 1
    # 滚动勾稽行（2026-08-03 用户方法论：期初+本期增加-本期减少-期末=0；差异>0.01 标红）
    # 对每个滚动组件（原值/累计折旧(摊销)/减值准备）各加一行；账面价值(net)仅期初/期末无增减不勾稽。
    for comp in ('cost', 'accum', 'impair'):
        seg_rows = rows_audit.get(comp)
        if not seg_rows:
            continue
        need = ('期初数', '本期增加', '本期减少', '期末数')
        if not all(s in seg_rows for s in need):
            continue
        r0 = seg_rows['期初数'].get('小计'); r1 = seg_rows['本期增加'].get('小计')
        r2 = seg_rows['本期减少'].get('小计'); r3 = seg_rows['期末数'].get('小计')
        if not (r0 and r1 and r2 and r3):
            continue
        exp = {}
        for e in ent_list:
            s = 0.0
            for cat in cats:
                s += (_row_val(e, comp, cat, 0) + _row_val(e, comp, cat, 1)
                      - _row_val(e, comp, cat, 2) - _row_val(e, comp, cat, 3))
            exp[e] = s
        _txt(ws, r, 1, '滚动勾稽（期初+本期增加-本期减少-期末）', bold=True, fill=SHELL_TOT_FILL)
        _txt(ws, r, 2, comp, bold=True, fill=SHELL_TOT_FILL)
        for i, ent in enumerate(ent_list, 3):
            # 2026-08-06 协议：滚动勾稽写数值（exp 已算；公式 data_only 读 None）
            _money(ws, r, i, exp[ent])
            cc = ws.cell(r, i)
            cc.font = Font(name='Times New Roman', size=10,
                           color='FF0000' if abs(exp[ent]) > 0.01 else '000000', bold=True)
        # 全集团合计列：数值
        _money(ws, r, ncols, sum(exp.values()))
        r += 1
    # 口径说明（A 列，不合并跨行）
    r += 1
    c = ws.cell(r, 1, '口径：未审数=期初/本期增减/期末账面数（科目余额表，备抵取绝对值）；审计调整留空待填；'
                      '期末审定数=未审+调整（公式）。类别=各主体二级科目名（无二级时并入组名）；'
                      '账面价值=原值−累计折旧(摊销)−减值准备。每主体一列，末列全集团合计。')
    c.font = SHELL_SUB_FONT
    # 列宽
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 22
    for i in range(3, ncols + 1):
        ws.column_dimensions[_col_letter(i)].width = 15
    return ws
def _prep_group(ent, km, gl, gkey):
    cost, accum, impair = _find_group_codes(km, gkey)
    # ⚡⚡ 2026-08-15 虚拟父级聚合：cost 为『共同前缀』（XBJ 1601 无父级、子目平行）时，
    #   km 无该 code 行 → cost_d 按前缀 Σ 全部子目（原值 opening/debit/credit/closing 带符号求和；
    #   name 取首个子目组名，odir/fdir 借=原值贷余负。下游仅按 dict 键访问，接口不变）。
    def _agg_cost(code):
        # ⚡⚡ 2026-08-24 修复（AZ 点分级双算根因）：_rows 若含【父级自身】+【子目】，
        #   sum 全部 = 父级+子目 = 双计（1601 父级金额 = 1601.01~06 子目之和）。
        #   规则：父级行存在 → 只取子目（排除父级自身，防双算）；无父级行（XBJ 平行
        #   子目 160101-160108）→ 全取（子目即全部明细）。
        #   ⚡⚡ 2026-08-24 补充：父级存在但【无子目】（1701 无形资产仅一级无明细）→
        #   排除父级后变空 → 回退用父级自身（不能返回 None 丢原值）。
        _keys = [c for c in km if str(c).startswith(str(code))]
        _has_parent = str(code) in km
        # ⚡⚡ 2026-08-24 只取【末级叶子】：排除父级自身 + 排除中间级（160101 也是
        #   16010101 的父级，AS energy 1601 把中间级 160101 与叶子 16010101 双算 →
        #   原值虚高 146.6M → 核对差 8.35亿）。父级行存在且无任何子目（1701 单级）→ 回退父级。
        _leaf_keys = [c for c in _keys
                      if not any(o.startswith(c) and len(o) > len(c) for o in _keys)]
        # ⚡⚡ 2026-09-01 ONCA 委托贷款排除『一年内到期』：1301 虚拟父级聚合会把
        #   一年内到期委托贷款（0.84B）并入其他非流动资产（企业报表口径：应归一年内
        #   到期的非流动资产）。仅 ONCA 且委托贷款前缀时排除该子目。
        if gkey == 'ONCA' and str(code) == '1301':
            _leaf_keys = [c for c in _leaf_keys
                          if '一年内到期' not in str(km[c].get('name') or '')]
        _rows = [km[c] for c in _leaf_keys if not (_has_parent and str(c) == str(code))]
        if not _rows:
            _rows = [km[str(code)]] if _has_parent and str(code) in km else []
        if not _rows:
            return None
        if len(_rows) == 1 and not _has_parent and str(code) in km:
            return _rows[0]
        return {'name': (str(_rows[0].get('name') or '').split('\\')[0]),
                'odir': '借', 'fdir': '借', 'level': 1,
                'opening': sum(float(d.get('opening') or 0.0) for d in _rows),
                'debit': sum(float(d.get('debit') or 0.0) for d in _rows),
                'credit': sum(float(d.get('credit') or 0.0) for d in _rows),
                'closing': sum(float(d.get('closing') or 0.0) for d in _rows)}
    rec = {'ent': ent, 'km': km, 'gl': gl, 'gkey': gkey,
           'cost': cost, 'accum': accum, 'impair': impair,
           'cost_d': _agg_cost(cost) if cost else None,
           'accum_d': km.get(accum) if accum else None,
           'impair_d': km.get(impair) if impair else None}
    rec['cost_inc'], rec['cost_dec'] = _grp_breakdown(gl, gkey, 'cost')
    rec['accum_inc'], rec['accum_dec'] = _grp_breakdown(gl, gkey, 'accum')
    rec['impair_inc'], rec['impair_dec'] = _grp_breakdown(gl, gkey, 'impair')
    rec['cost_l2'] = [(c, km[c]) for c in km if cost and c.startswith(cost) and c != cost
                      # ⚡⚡ 2026-09-01 ONCA 委托贷款排除『一年内到期』（附注/审定表同源）
                      and not (gkey == 'ONCA' and str(cost) == '1301'
                               and '一年内到期' in str(km[c].get('name') or ''))
                      and not (impair and c.startswith(impair))]
    # 2026-08-02 附注披露所需：累计折旧/摊销 与 减值准备 的二级明细
    rec['accum_l2'] = [(c, km[c]) for c in km if accum and c.startswith(accum) and c != accum]
    rec['impair_l2'] = [(c, km[c]) for c in km if impair and c.startswith(impair) and c != impair]
    return rec


def _acc_inc(acc, role):
    if not acc:
        return 0.0
    return acc.get('credit', 0.0) if role in ('accum', 'impair') else acc.get('debit', 0.0)


def _acc_dec(acc, role):
    if not acc:
        return 0.0
    return acc.get('debit', 0.0) if role in ('accum', 'impair') else acc.get('credit', 0.0)


def _disp(acc, role, field):
    if not acc:
        return 0.0
    v = acc.get(field, 0.0)
    if role in ('accum', 'impair'):
        return abs(v)
    return v


# ----------------------------------------------------------------------------
# Sheet: 增减变动表（主表；在建工程特殊处理）
# ----------------------------------------------------------------------------
def build_rollforward_sheet(wb, year, recs, gkey):
    sp = GROUPS[gkey]
    ws = wb.create_sheet(sp['main_sheet'])
    # 铁律：科目名以实际账套为准——动态发现该组实际备抵科目名（如"使用权资产折旧"）
    accum_label = _discover_accum_name(recs, gkey) if not sp.get('special') else None
    headers = ['主体', '类别', '科目代码', '科目名称', '期初数', '增加', '减少', '期末数',
               '审计调整数', '审定数', '备注']
    ncols = len(headers)
    is_cip = sp.get('special')
    title = '%s（%s 年度）' % (sp['main_sheet'], year)
    sub = ('控制数取自各主体《科目余额表》(权威,带借贷符号):"增加/减少"=科目余额表净发生额(借发/贷发),'
           '为审计控制数,不以GL填列。二级明细期初/增加/减少/期末均取自科目余额表对应二级科目。'
           '审计调整数留空待填,审定数=期末未审数+审计调整数。全集团按二级分类汇总见表末"全集团合计"行。'
           '在建工程相对特殊：仅列期初/本期增加/本期减少/期末/审计调整/审定数（不单列净值与折旧摊销）。'
           ) if is_cip else (
           '控制数取自各主体《科目余额表》(权威,带借贷符号):"增加/减少"=科目余额表净发生额(借发/贷发),'
           '为审计控制数,不以GL填列。各二级明细期初/增加/减少/期末均取自科目余额表对应二级科目。'
           '审计调整数留空待填,审定数=期末未审数+审计调整数;备抵科目(累计折旧/摊销/减值准备)以正数展示。'
           '全集团按二级分类汇总见表末"全集团合计"行(原值/累计折旧(摊销)/减值准备/净值 合计)。'
           '各底稿《增加检查表》列示"在建工程转入"金额(从在建工程转入数)。')
    r = _title(ws, title, ncols)
    r = _sub(ws, sub, ncols, r)
    hr = _hdr(ws, r, headers)
    r = hr
    COL = {h: i + 1 for i, h in enumerate(headers)}

    g_acc = {'o': 0.0, 'i': 0.0, 'd': 0.0, 'c': 0.0}
    ga_acc = {'o': 0.0, 'i': 0.0, 'd': 0.0, 'c': 0.0}
    gi_acc = {'o': 0.0, 'i': 0.0, 'd': 0.0, 'c': 0.0}
    # 按二级分类汇总（全组件滚动）：原值/累计折旧(摊销)/减值准备 各 [期初,本期增加,本期减少,期末]
    g_cat = defaultdict(lambda: {'cost': [0.0, 0.0, 0.0, 0.0],
                                 'accum': [0.0, 0.0, 0.0, 0.0],
                                 'impair': [0.0, 0.0, 0.0, 0.0]})

    def _merge_cat(cat):
        """备抵类别归并到原值类别（2026-08-07，铁律55 同 _put_code）：
        『土地使用费摊销』『土地使用权累计摊销』→ 归并到『土地使用费』原值类，
        使 g_cat 按原值类别聚合（原值/摊销/减值/净值同键），分类净值才能计算。
        无匹配（如费用/其他）保留原名兜底。"""
        base = cat.replace('折旧', '').replace('摊销', '').replace('减值', '').replace('准备', '').strip()
        if base:
            for _c0 in list(g_cat.keys()):
                if _c0 == base or (base and (base in _c0 or _c0 in base)):
                    return _c0
        return cat

    def _write_row(rr, ent, cat, code, name, opening, inc, dec, closing, bold=False, grp=False):
        _txt(ws, rr, COL['主体'], ent, bold=bold or grp)
        _txt(ws, rr, COL['类别'], cat, bold=bold or grp)
        _txt(ws, rr, COL['科目代码'], code, bold=bold or grp)
        _txt(ws, rr, COL['科目名称'], name, bold=bold or grp)
        _money(ws, rr, COL['期初数'], opening)
        _money(ws, rr, COL['增加'], inc)
        _money(ws, rr, COL['减少'], dec)
        _money(ws, rr, COL['期末数'], closing)
        _money(ws, rr, COL['审计调整数'], None)
        # 2026-08-06 协议：审定数写数值（=期末数+调整0；公式 data_only 读 None 收回读不回）
        _money(ws, rr, COL['审定数'], closing)
        if grp:
            for i in range(1, ncols + 1):
                ws.cell(rr, i).fill = SHELL_TOT_FILL

    for ent, rec in recs:
        # 原值（一级 + 二级）
        cost = rec['cost_d']
        code = rec['cost'] or ''
        a_inc = _acc_inc(cost, 'cost'); a_dec = _acc_dec(cost, 'cost')
        g_acc['o'] += _disp(cost, 'cost', 'opening')
        g_acc['i'] += a_inc
        g_acc['d'] += a_dec
        g_acc['c'] += _disp(cost, 'cost', 'closing')
        _write_row(r, ent, sp['title'], code, sp['title'] + '（合计）',
                   _disp(cost, 'cost', 'opening'), a_inc, a_dec, _disp(cost, 'cost', 'closing'),
                   bold=True)
        r += 1
        if rec['cost_l2']:
            for c, d in rec['cost_l2']:
                # 2026-08-03：分类汇总原只列二级（6 位码/1 个点）。
                # 2026-08-07 修复：在建工程展开三级（8 位码）——用户要求"分到二级、三级"
                #（dq 160402 安装工程 136M = 三级 16040201 在安装设备 + 16040203 电工电器
                # 等之和，原仅列二级无法看项目明细）。四级及以上（10 位）仍不列（粒度到三级）。
                # 注意：三级行【不加入 g_cat】（g_cat 只归并二级）——三级加进会与二级父行
                # 双计（父=子和），『全集团合计』行独立按二级计算。
                _cnum = len(str(c).replace('.', ''))
                # ⚡ 2026-08-10 SAP 码长适配：二级=6(U8)/10(SAP) 位，三级=8(U8)/12(SAP)
                # 位；CIP 展开到末级，其余组展开到二级。
                _is_l2 = _cnum in (6, 10)
                if not str(c).replace('.', '').isdigit():
                    continue
                inc2 = d['debit']
                dec2 = d['credit']
                if is_cip:
                    # ⚡ 2026-08-11 DQ1：在建工程分类汇总【层级树】——同时保留一级、二级、末级明细
                    #   （仿管理费用明细表：一级合计行 + 二级加粗行 + 三级/末级缩进行）。
                    #   原 2026-08-10 只列末级叶子（_has_child 过滤父级）→ 用户要求层级化。
                    #   防双计：父级行显示其真实 TB 值（=子级之和），全集团合计行仍按一级（cost）独立计算。
                    # ⚡ 2026-08-11 修复：原内层 `if is_cip: ... else: <写行>` 嵌套错误——FA/IPR/INT/ROU
                    #   等非 CIP 组外层 is_cip=False 跳过整块 → 原值二级明细行永不写入（DQ 固定资产
                    #   汇总表 R5-R16 空、只有合计行）。还原为外层 if/else：CIP 层级树 / 其余组二级行。
                    _ln = len(str(c).replace('.', ''))
                    _lv = 3 if _ln not in (4, 6, 8) else ({4: 0, 6: 1, 8: 2}[_ln])
                    nm = d['name']
                    if _lv >= 2:
                        nm = '\u3000' * (_lv - 1) + str(nm)
                    _write_row(r, '', '', c, nm, d['opening'], d['debit'], d['credit'], d['closing'],
                               bold=(_lv == 1))
                else:
                    if not _is_l2:
                        # 其余资产组（FA/IPR/INT/ROU）：展开到二级（三级及以上归并二级）
                        continue
                    _write_row(r, '', '', c, d['name'], d['opening'], d['debit'], d['credit'], d['closing'])
                if _is_l2:
                    # g_cat 归并到二级名称（与附注汇总 _put 一致；三级如 固定资产-交通运输设备-汽车 归入 交通运输设备）
                    cat = (d['name'].split('-', 1)[1] if '-' in d['name'] else d['name']).split('-')[0].strip()
                    g_cat[cat]['cost'][0] += d['opening']
                    g_cat[cat]['cost'][1] += inc2
                    g_cat[cat]['cost'][2] += dec2
                    g_cat[cat]['cost'][3] += d['closing']
                r += 1
        else:
            g_cat[sp['title']]['cost'][0] += _disp(cost, 'cost', 'opening')
            g_cat[sp['title']]['cost'][1] += a_inc
            g_cat[sp['title']]['cost'][2] += a_dec
            g_cat[sp['title']]['cost'][3] += _disp(cost, 'cost', 'closing')

        if is_cip:
            continue  # 在建工程无备抵/净值

        # 累计折旧/摊销（一级 + 二级）
        if rec['accum_d'] is not None:
            accum = rec['accum_d']
            ai = _acc_inc(accum, 'accum'); ad = _acc_dec(accum, 'accum')
            ga_acc['o'] += abs(_disp(accum, 'accum', 'opening'))
            ga_acc['i'] += ai
            ga_acc['d'] += ad
            ga_acc['c'] += abs(_disp(accum, 'accum', 'closing'))
            _write_row(r, ent, (rec['accum_d']['name'] or '累计折旧').split('-')[0].strip(), rec['accum'] or '',
                       (rec['accum_d']['name'] or '累计折旧') + '（合计）',
                       abs(_disp(accum, 'accum', 'opening')), ai, ad,
                       abs(_disp(accum, 'accum', 'closing')), bold=True)
            r += 1
            a_l2 = [(c, km_c) for c, km_c in [(cc, rec['km'][cc]) for cc in rec['km']
                      if rec['accum'] and cc.startswith(rec['accum']) and cc != rec['accum']]]
            if a_l2:
                for c, d in a_l2:
                    # 2026-08-03：分类汇总只列二级（U8 6 位码）；三级及以上不列
                    # ⚡ 2026-08-10 SAP 码长适配：SAP 二级=10 位（1702010000）同样列为明细
                    if str(c).replace('.', '').isdigit() and len(str(c).replace('.', '')) not in (6, 10):
                        continue
                    _write_row(r, '', '', c, d['name'], abs(d['opening']), d['credit'], d['debit'],
                               abs(d['closing']))
                    cat = (d['name'].split('-', 1)[1] if '-' in d['name'] else d['name']).split('-')[0].strip()
                    cat = _merge_cat(cat)   # 2026-08-07：备抵类别归并到原值类（分类净值前提）
                    g_cat[cat]['accum'][0] += abs(d['opening'])
                    g_cat[cat]['accum'][1] += d['credit']
                    g_cat[cat]['accum'][2] += d['debit']
                    g_cat[cat]['accum'][3] += abs(d['closing'])
                    r += 1
            elif rec['accum_d'] is not None:
                g_cat[sp['title']]['accum'][0] += abs(_disp(rec['accum_d'], 'accum', 'opening'))
                g_cat[sp['title']]['accum'][1] += ai
                g_cat[sp['title']]['accum'][2] += ad
                g_cat[sp['title']]['accum'][3] += abs(_disp(rec['accum_d'], 'accum', 'closing'))

        # 减值准备（一级）
        if rec['impair_d'] is not None:
            imp = rec['impair_d']
            ii = _acc_inc(imp, 'impair'); idc = _acc_dec(imp, 'impair')
            gi_acc['o'] += abs(_disp(imp, 'impair', 'opening'))
            gi_acc['i'] += ii
            gi_acc['d'] += idc
            gi_acc['c'] += abs(_disp(imp, 'impair', 'closing'))
            _write_row(r, ent, '减值准备', rec['impair'] or '',
                       (imp['name'] or '减值准备') + '（合计）',
                       abs(_disp(imp, 'impair', 'opening')), ii, idc,
                       abs(_disp(imp, 'impair', 'closing')), bold=True)
            g_cat[sp['title']]['impair'][0] += abs(_disp(imp, 'impair', 'opening'))
            g_cat[sp['title']]['impair'][1] += ii
            g_cat[sp['title']]['impair'][2] += idc
            g_cat[sp['title']]['impair'][3] += abs(_disp(imp, 'impair', 'closing'))
            r += 1

        # 净值行（2026-08-07：先按二级分类列示分类净值，再列总净值——用户要求
        # 『无形资产净值先是分类净值』。分类净值=该分类原值−累计摊销/折旧−减值，
        # 取自 g_cat（已按 _merge_cat 归并到原值类别）；无备抵/无数据的分类跳过）
        _net_cats = []
        for _cat in g_cat:
            _gc = g_cat[_cat]
            _no = _gc['cost'][0] - _gc['accum'][0] - _gc['impair'][0]
            _nc = _gc['cost'][3] - _gc['accum'][3] - _gc['impair'][3]
            if abs(_no) > 0.005 or abs(_nc) > 0.005:
                _net_cats.append((_cat, _no, _nc))
        for _cat, _no, _nc in sorted(_net_cats, key=lambda x: -(abs(x[2]))):
            _txt(ws, r, COL['主体'], ent)
            _txt(ws, r, COL['类别'], sp['title'] + ' 净值')
            _txt(ws, r, COL['科目代码'], '—')
            _txt(ws, r, COL['科目名称'], '★ %s净值' % _cat)
            _money(ws, r, COL['期初数'], round(_no, 2))
            _money(ws, r, COL['增加'], None)
            _money(ws, r, COL['减少'], None)
            _money(ws, r, COL['期末数'], round(_nc, 2))
            _money(ws, r, COL['审计调整数'], None)
            _money(ws, r, COL['审定数'], round(_nc, 2))
            r += 1
        co = _disp(cost, 'cost', 'opening')
        cc = _disp(cost, 'cost', 'closing')
        ao = abs(_disp(rec['accum_d'], 'accum', 'opening')) if rec['accum_d'] else 0.0
        ac = abs(_disp(rec['accum_d'], 'accum', 'closing')) if rec['accum_d'] else 0.0
        io = abs(_disp(rec['impair_d'], 'impair', 'opening')) if rec['impair_d'] else 0.0
        ic = abs(_disp(rec['impair_d'], 'impair', 'closing')) if rec['impair_d'] else 0.0
        net_op = co - ao - io
        net_cl = cc - ac - ic
        _txt(ws, r, COL['主体'], ent)
        _txt(ws, r, COL['类别'], sp['title'])
        _txt(ws, r, COL['科目代码'], '—')
        _txt(ws, r, COL['科目名称'], '★ %s净值（=%s−累计折旧/摊销−减值准备）' % (sp['title'], sp['title']), bold=True)
        _money(ws, r, COL['期初数'], net_op)
        _money(ws, r, COL['增加'], None)
        _money(ws, r, COL['减少'], None)
        _money(ws, r, COL['期末数'], net_cl)
        _money(ws, r, COL['审计调整数'], None)
        # 2026-08-06 协议：审定数写数值（=期末数+调整0）
        _money(ws, r, COL['审定数'], net_cl)
        r += 1

    # ===== 全集团合计（2026-07-29 起：不再独立 sheet，直接接在分类汇总表末尾） =====
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出全集团合计块
    if not is_cip and len(recs) > 1:
        r += 1
        # 2026-08-06 修复：原『全集团合计』为无数值的区块标题行（真正带数值的合计
        # 在其下方『全集团 XX（原值合计）』行），被检查器当合计行校验 → 明细区扩大到
        # 主体明细+净值行 → 双计误报 5 个单元格。删除冗余标题行。
        # 全集团按二级分类展开（2026-07-29 用户要求：分类汇总展开显示各二级类别）
        _txt(ws, r, 1, '全集团按二级分类汇总', bold=True)
        for i in range(1, ncols + 1):
            ws.cell(r, i).fill = SHELL_TOT_FILL
            ws.cell(r, i).font = SHELL_BOLD
        r += 1
        cats = [c for c in CAT_ORDER_PREF if c in g_cat] + \
               [c for c in sorted(g_cat) if c not in CAT_ORDER_PREF]
        for cat in cats:
            g = g_cat[cat]
            # 原值行
            _txt(ws, r, COL['主体'], '全集团')
            _txt(ws, r, COL['类别'], cat, bold=True)
            _txt(ws, r, COL['科目代码'], '—')
            _txt(ws, r, COL['科目名称'], cat, bold=True)
            _money(ws, r, COL['期初数'], g['cost'][0])
            _money(ws, r, COL['增加'], g['cost'][1])
            _money(ws, r, COL['减少'], g['cost'][2])
            _money(ws, r, COL['期末数'], g['cost'][3])
            _money(ws, r, COL['审计调整数'], None)
            # 2026-08-06 协议：审定数写数值（=期末数+调整0）
            _money(ws, r, COL['审定数'], g['cost'][3])
            r += 1
        # 全集团原值合计
        _txt(ws, r, COL['主体'], '全集团')
        _txt(ws, r, COL['类别'], sp['title'] + '（原值合计）', bold=True)
        _txt(ws, r, COL['科目代码'], '—')
        _txt(ws, r, COL['科目名称'], '全集团 %s（合计）' % sp['title'], bold=True)
        _money(ws, r, COL['期初数'], g_acc['o'])
        _money(ws, r, COL['增加'], g_acc['i'])
        _money(ws, r, COL['减少'], g_acc['d'])
        _money(ws, r, COL['期末数'], g_acc['c'])
        _money(ws, r, COL['审计调整数'], None)
        # 2026-08-06 协议：审定数写数值
        _money(ws, r, COL['审定数'], g_acc['c'])
        for i in range(1, ncols + 1):
            ws.cell(r, i).fill = SHELL_TOT_FILL
            ws.cell(r, i).font = SHELL_BOLD
        r += 1
        # 累计折旧/摊销 按二级分类展开（2026-07-29 用户要求：累计摊销也要按二级展开）
        if g_cat and any(g['accum'][0] or g['accum'][1] or g['accum'][2] or g['accum'][3] for g in g_cat.values()):
            _txt(ws, r, 1, '全集团按二级分类汇总（累计折旧/摊销）', bold=True)
            for i in range(1, ncols + 1):
                ws.cell(r, i).fill = SHELL_TOT_FILL
                ws.cell(r, i).font = SHELL_BOLD
            r += 1
            for cat in cats:
                g = g_cat[cat]
                if not (g['accum'][0] or g['accum'][1] or g['accum'][2] or g['accum'][3]):
                    continue
                _txt(ws, r, COL['主体'], '全集团')
                _txt(ws, r, COL['类别'], cat, bold=True)
                _txt(ws, r, COL['科目代码'], '—')
                _txt(ws, r, COL['科目名称'], cat + '（累计' + ('摊销' if gkey in ('INT','LTDEF') else '折旧') + '）', bold=True)
                _money(ws, r, COL['期初数'], g['accum'][0])
                _money(ws, r, COL['增加'], g['accum'][1])
                _money(ws, r, COL['减少'], g['accum'][2])
                _money(ws, r, COL['期末数'], g['accum'][3])
                _money(ws, r, COL['审计调整数'], None)
                # 2026-08-06 协议：审定数写数值
                _money(ws, r, COL['审定数'], g['accum'][3])
                r += 1
        # 累计折旧全集团合计
        if g_cat and any(g['accum'][0] or g['accum'][1] or g['accum'][2] or g['accum'][3] for g in g_cat.values()):
            _txt(ws, r, COL['主体'], '全集团')
            _txt(ws, r, COL['类别'], accum_label + '（合计）', bold=True)
            _txt(ws, r, COL['科目代码'], '—')
            _txt(ws, r, COL['科目名称'], '全集团 %s（合计）' % accum_label, bold=True)
            _money(ws, r, COL['期初数'], ga_acc['o'])
            _money(ws, r, COL['增加'], ga_acc['i'])
            _money(ws, r, COL['减少'], ga_acc['d'])
            _money(ws, r, COL['期末数'], ga_acc['c'])
            _money(ws, r, COL['审计调整数'], None)
            # 2026-08-06 协议：审定数写数值
            _money(ws, r, COL['审定数'], ga_acc['c'])
            for i in range(1, ncols + 1):
                ws.cell(r, i).fill = SHELL_TOT_FILL
                ws.cell(r, i).font = SHELL_BOLD
            r += 1
        # 全集团净值合计（2026-08-07：原缺净值合计行——分类净值已列，此处补全集团净值
        # = 原值合计 − 累计折旧/摊销合计 − 减值合计，与主体区分类净值之和勾稽）
        _net_o = g_acc['o'] - ga_acc['o'] - gi_acc['o']
        _net_c = g_acc['c'] - ga_acc['c'] - gi_acc['c']
        _txt(ws, r, COL['主体'], '全集团')
        _txt(ws, r, COL['类别'], sp['title'] + ' 净值（合计）', bold=True)
        _txt(ws, r, COL['科目代码'], '—')
        _txt(ws, r, COL['科目名称'], '★ 全集团 %s净值（合计）' % sp['title'], bold=True)
        _money(ws, r, COL['期初数'], _net_o)
        _money(ws, r, COL['增加'], None)
        _money(ws, r, COL['减少'], None)
        _money(ws, r, COL['期末数'], _net_c)
        _money(ws, r, COL['审计调整数'], None)
        _money(ws, r, COL['审定数'], _net_c)
        for i in range(1, ncols + 1):
            ws.cell(r, i).fill = SHELL_TOT_FILL
            ws.cell(r, i).font = SHELL_BOLD
        r += 1

    # 全集团按二级分类汇总 已直接接在分类汇总表末尾，不再独立 sheet。
    widths = [14, 14, 14, 30, 14, 14, 14, 14, 11, 12, 16] + [13] * (ncols + 5)
    for i, w in enumerate(widths[:ncols + 5], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'E4'
    return g_cat


# ----------------------------------------------------------------------------
# Sheet: 全集团汇总明细（独立 sheet，原 分类汇总表 底部“全集团按二级分类汇总”区块移出）
# ----------------------------------------------------------------------------
def build_group_summary_sheet(wb, year, recs, gkey, g_cat):
    """全集团按二级分类汇总（独立 sheet：全集团汇总明细）。
    复用 build_rollforward_sheet 已累计的 g_cat（按二级分类聚合的原值/累计折旧/减值准备 各 期初/增加/减少/期末）。
    审定数 = 期末数 + 审计调整数（修正此前“审定数引用自身”导致的循环引用）。"""
    sp = GROUPS[gkey]
    is_cip = sp.get('special')
    # 铁律：科目名以实际账套为准——动态发现该组实际备抵科目名
    accum_name = _discover_accum_name(recs, gkey) if not is_cip else '累计折旧'
    comps = ['原值'] if is_cip else ['原值', '净值']
    metrics = ['期初数', '本期增加', '本期减少', '期末数', '审计调整数', '审定数']
    # 净值仅需期初期末（用户要求 2026-07-30）
    net_metrics = ['期初数', '期末数']
    ncols = 1 + len(comps) * len(metrics)
    title = '全集团汇总明细（%s，%s 年度）' % (sp['title'], year)
    idx = wb.sheetnames.index(sp['main_sheet']) + 1 if sp['main_sheet'] in wb.sheetnames else len(wb.sheetnames)
    ws = wb.create_sheet('全集团汇总明细', idx)
    r = _title(ws, title, ncols)
    r = _sub(ws, '按二级分类（固定资产/累计折旧/减值准备的二级科目后缀，如“专用设备”“运输设备”）汇总全集团各组件'
              '（原值/累计折旧(摊销)/减值准备/净值）的期初、本期增加、本期减少、期末、审计调整、审定数'
              '（审定数=期末+调整，调整留空待填）。净值=原值−累计折旧(摊销)−减值准备；金额取自各主体《科目余额表》对应二级科目。', ncols, r)
    r += 1
    # 两级表头：一级组件合并，二级指标
    hdr1 = ['二级分类']
    for comp in comps:
        hdr1.append(comp)
        ws.merge_cells(start_row=r, start_column=len(hdr1), end_row=r, end_column=len(hdr1) + len(metrics) - 1)
        hdr1 += [''] * (len(metrics) - 1)
    for j, h in enumerate(hdr1, 1):
        if h:
            cc = ws.cell(r, j, h)
            cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    r += 1
    for j, m in enumerate(metrics, 1):
        cc = ws.cell(r, j + 1, m)
        cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    r += 1
    cats = [c for c in CAT_ORDER_PREF if c in g_cat] + \
          [c for c in sorted(g_cat) if c not in CAT_ORDER_PREF]
    tot = {comp: [0.0, 0.0, 0.0, 0.0] for comp in comps}
    cat_comp_row = {}
    for cat in cats:
        g = g_cat[cat]
        net = [g['cost'][k] - g['accum'][k] - g['impair'][k] for k in range(4)]
        _txt(ws, r, 1, cat, bold=True)
        col = 2
        for comp in comps:
            vals = net if comp == '净值' else g['cost']
            if comp == '净值':
                # 净值仅期初期末
                for k in [0, 3]:  # 期初=0, 期末=3
                    _money(ws, r, col, vals[k])
                    col += 1
            else:
                for k, m in enumerate(metrics):
                    if m in ('期初数', '本期增加', '本期减少', '期末数'):
                        _money(ws, r, col, vals[k])
                    elif m == '审计调整数':
                        _money(ws, r, col, None)
                    else:
                        # 2026-08-06 协议：审定数写数值（=期末数+调整0）
                        _money(ws, r, col, vals[3])
                    col += 1
            cat_comp_row[(cat, comp)] = r
        for i in range(2, col):
            ws.cell(r, i).fill = SHELL_TOT_FILL
        for comp in comps:
            vals = net if comp == '净值' else g['cost']
            for k in range(4):
                tot[comp][k] += vals[k]
        r += 1
    # 合计行
    _txt(ws, r, 1, '合计', bold=True)
    col = 2
    for comp in comps:
        if comp == '净值':
            for k in [0, 3]:
                _money(ws, r, col, tot[comp][k])
                col += 1
        else:
            for k, m in enumerate(metrics):
                if m in ('期初数', '本期增加', '本期减少', '期末数'):
                    _money(ws, r, col, tot[comp][k])
                elif m == '审计调整数':
                    # 2026-08-06 协议：合计行调整数写 0（类别行调整留空）
                    _money(ws, r, col, 0.0)
                else:  # 审定数 = 各类别审定数之和 = 期末合计（2026-08-06 协议：写数值）
                    _money(ws, r, col, tot[comp][3])
                col += 1
        for i in range(1, col):
            ws.cell(r, i).fill = SHELL_TOT_FILL
            ws.cell(r, i).font = SHELL_BOLD
    r += 1
    r += 1
    note = ('说明：上表按二级分类汇总全集团各组件（原值/累计折旧(摊销)/减值准备/净值）的期初、本期增加、本期减少、期末、审计调整、审定数'
            '（审定数=期末+调整，调整留空待填）。净值=原值−累计折旧(摊销)−减值准备；各分类金额取自各主体《科目余额表》对应二级科目；无二级分类的科目以一级汇总。')
    nc = ws.cell(r, 1, note)
    nc.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 1
    return r


# ----------------------------------------------------------------------------
# Sheet: 折旧/摊销计提分月分科目核对
# ----------------------------------------------------------------------------
def _prior_year_acc_group(ents, ent, prior_year, gkey):
    try:
        if not ents:
            return None
        gl_path = ents.get(ent, {}).get('gl', {}).get(prior_year)
        if not gl_path:
            return None
        gl = _read_gl_fa(gl_path)
        py = int(prior_year)
        sp = GROUPS[gkey]
        src_role = sp['amort_src']
        tot = 0.0
        for r in gl:
            g, rl = _grp_cat(r['name'])
            if g == gkey and rl == src_role and r['credit'] > 0:
                ym = _ym(r['date'])
                if ym and ym[0] == py:
                    tot += r['credit']
        return tot
    except Exception:
        return None


def _expense_marker(gkey, nm, sm):
    nm = nm or ''
    sm = sm or ''
    # 排除明显不属于折旧/摊销费用标的的科目（如收到了概括匹配关键词的非费用类账户）。
    # 递延收益之"摊销"语义≠无形资产摊销；应交税费/所得税费用/租赁负债/递延税项 等
    # 余额类科目的摘要含"折旧"时不应被匹配，否则摊销核对表会误列这些科目。
    top = (nm.split('-')[0] if '-' in nm else nm).strip()
    _EXPENSE_EXCLUDE = {'递延收益', '应交税费', '所得税费用', '租赁负债',
                        '递延所得税负债', '递延所得税资产'}
    if top in _EXPENSE_EXCLUDE:
        return False
    if gkey == 'FA':
        # ⚡⚡ 2026-08-23 修复：原『科目名或摘要含"折旧"即认』过宽——AFJ SF 12月 转-2128
        #   混合凭证（折旧计提+福利费计提+租金）中『应付职工薪酬-福利费』『管理费用-
        #   职工薪酬-福利费』（摘要"计提本月折旧费"含"折旧"）被误判为折旧列支 →
        #   折旧分摊表列支虚增 16.3 万、计提数≠列支合计（差 166,901.86）。
        #   收窄：①科目名含"折旧"→ 折旧费科目直接认；②仅摘要含"折旧"→ 必须 top 为
        #   费用科目 且（无子目 或 子目含"折旧"）——排除 应付职工薪酬/其他应付款 等
        #   非费用账户，以及 管理费用-职工薪酬-福利费 这类非折旧子目。
        # ⚡⚡ 2026-08-28 尝试修复（已回退）：加"主营业务成本"识别会造成 FA 折旧分摊表
        #   全线误收（1100/2468 差数亿）——成本结转/销售成本凭证与折旧计提凭证可能同号
        #   或同凭证含累计折旧 → 非折旧主营业务成本行被误判列支。正确方案见
        #   _is_amort_partner 的比例封顶（跨组剥离），此处保持只认"折旧"字样 + 摘要
        #   判定，1240 的"主营业务成本-光伏成本"折旧走披露（见摊销核对表注）。
        if nm.startswith('累计折旧') or '投资性房地产' in nm or '使用权资产' in nm \
                or '使用权资产' in sm or '投资性房地产' in sm:
            return False
        if '折旧' in nm:
            return True
        if '折旧' in sm:
            if top not in ('管理费用', '制造费用', '销售费用', '研发费用',
                           '研发支出', '营业费用', '其他业务成本'):
                return False
            if '-' in nm:
                sub = nm.split('-', 1)[1]
                return '折旧' in sub
            return True
        return False
    if gkey == 'IPR':
        # ⚡ 2026-08-10 SAP 适配：成本模式投房折旧计入【其他业务成本】（SAP 1010
        # 『其他业务成本-租金』与 投资性房地产-累计折旧 同凭证 1:1 配平）；原逻辑
        # 仅认名称含"投资性房地产"→ SAP 独立费用科目（折旧费/其他业务成本）全被
        # 挡在门外 → IPR 列支恒 0。放宽：费用行名称/摘要含 投资性房地产 或
        # 其他业务成本（同凭证判定已确保是本组计提凭证的对方，不会误收 FA 的折旧费）。
        return (('投资性房地产' in nm or '投资性房地产' in sm
                 or '其他业务成本' in nm or '其他业务成本' in sm)
                and not nm.startswith('投资性房地产累计折旧')
                and not nm.startswith('投资性房地产-累计折旧')
                and not nm.startswith('投资性房地产减值准备'))
    if gkey == 'INT':
        # ⚡⚡ 2026-08-27 复盘：SAP 1010 无形资产摊销 622万 中 120万 计入『无形资产摊销』
        #   科目、502万 计入【折旧费】通用科目（与固定资产折旧合并核算，无法拆分）。
        #   因此列支只认『无形资产摊销』等明确摊销科目；『折旧费』不列入支（否则 FA 折旧
        #   461万/月 全被误收 → 核对差 -461万 过度计入）。差异在核对表披露归因。
        # ⚡⚡ 2026-08-28 修复（对应科目未找全/误收）：『研究开发费-折旧费用与长期费用摊销』
        #   (66000605) 是【研发折旧】（FA 归属，08-08 记忆确认），科目名同时含"折旧"+"摊销"
        #   → INT marker 因"摊销"误认 → 反向结转凭证（借 折旧费 / 贷 研发折旧）被 INT 侧
        #   扣减 417,445 → 研发费用列负值。INT 只认无形资产摊销相关，含"折旧"的科目一律排除
        #   （折旧属 FA/ROU/IPR）。『研究开发费-无形资产摊销』(66000607) 不含"折旧"→ 保留。
        return (('摊销' in nm or '摊销' in sm or '无形资产' in nm or '无形资产' in sm)
                and '折旧' not in nm
                and not nm.startswith('累计摊销')
                and not nm.startswith('无形资产-'))
    if gkey == 'ROU':
        # 仅"使用权资产"相关行计列支（科目名/摘要含 使用权资产/租赁/房租，如"管理费用-房租费"
        # 摘要"摊X月房租"——绍兴/金属 ROU 摊销摘要无"使用权资产"字样，需'房租'关键词兜底）
        # ⚡⚡ 2026-08-28 修复（对应科目未找全）：SAP 使用权资产折旧计入【租赁费-房屋租赁费】
        #   (6600171，IFRS16 惯例，08-08 记忆确认"识别折旧族须含 6600171")——科目名不含
        #   "折旧/摊"关键词，原 marker 挡掉 → 1357/2468 使用权资产折旧列支全漏。
        #   补"租赁费"前缀识别（凭证级 _is_amort_partner 已保证只认"使用权资产累计折旧
        #   同凭证"的租赁费行，纯经营租赁租金凭证无 ROU 备抵贷方 → 不会误收）。
        return (('使用权资产' in nm or '使用权资产' in sm or '租赁' in sm or '租赁' in nm
                 or '房租' in nm or '房租' in sm
                 or nm.startswith('租赁费'))
                and ('折旧' in nm or '折旧' in sm or '摊' in sm
                     or nm.startswith('租赁费'))
                and not nm.startswith('使用权资产累计折旧') and not nm.startswith('使用权资产折旧')
                and '投资性房地产' not in nm)
    if gkey == 'LTDEF':
        return ('摊销' in nm or '摊销' in sm) and '长期待摊费用' not in nm
    return False


def _vouch_key(r):
    """凭证分组键：年月|凭证号（U8 凭证号每月重置，各月都有"记-85"，
    跨月同号必须隔离，否则混合/错配）。"""
    vk = (r.get('voucher') or '').strip()
    if not vk:
        return ''
    ym = (r.get('date') or '')[:7]
    return (ym + '|' + vk) if ym else vk


_EXPENSE_TOPS = ('管理费用', '制造费用', '销售费用', '研发费用', '研发支出', '主营业务成本',
                 '其他业务成本', '其他业务支出', '营业费用', '生产成本', '劳务成本', '业务及管理费')

# ⚡⚡ 2026-08-28 对应科目罗列（用户方法论：不追求计提=列支对平，把实际计入的费用/成本
#   科目罗列出来）。折旧/摊销计提凭证（贷方为本组备抵）内，借方行若为下列【STD 外
#   费用/成本类 top】（主营业务成本-光伏成本、研发支出-直接投入、其他业务成本-租金等），
#   亦作为"对应科目"计入核对表列（动态追加列），而非塞进差异归因区。仅当凭证级
#   _is_amort_partner 判定相关（贷方有本组备抵 或 本组费用科目借贷双向结转）时才生效，
#   纯销售成本结转凭证（贷方为存货/在制品，无备抵）不会误收。
_COST_TOPS = ('主营业务成本', '其他业务成本', '其他业务支出', '营业费用', '生产成本',
              '劳务成本', '研发支出', '业务及管理费')
# ⚠️ 勿把『研究开发费』(660006) 加进 _COST_TOPS：其下子目如 直接投入/其他费用 等
#    无折旧/摊销语义，仅靠 top 白名单会在结转凭证（借 研究开发费-直接投入 / 贷 折旧费）
#    中被误收进折旧核对表（1090 研发费用列多计 24.4万 根因）。研发折旧 66000605/研发
#    摊销 66000607 已由本组 marker（名称含折旧/摊销）识别，无需 top 白名单。"""


def _is_cost_subject(nm):
    """费用/成本类一级科目判定（STD 外对应科目罗列用）。"""
    return _top_account(nm) in _COST_TOPS


def _is_cost_subject_own(gkey, nm, sm):
    """STD 外费用/成本类 top 是否【归本组】收列：
    本组 marker 命中，或（top 属费用/成本类 且 不被其他资产组 marker 认领）。
    ⚡⚡ 2026-08-28 修复（混合凭证误收）：1010 计提凭证 2200000000 为三组合并凭证
    （贷 累计折旧+累计摊销+投房累计折旧），『其他业务成本-租金』是 IPR(投房)折旧去向
    （IPR marker 认领），若仅按 _is_cost_subject 无条件收 → FA/INT 表也收 295.7万 →
    摊销核对差 -295.7万。故 STD 外 top 必须【不被其他组 marker 认领】才归本组；
    『主营业务成本-光伏成本』(1240,FA折旧去向) 无任何组 marker 命中 → 归 FA ✓。"""
    if _expense_marker(gkey, nm, sm):
        return True
    if not _is_cost_subject(nm):
        return False
    for _g2, _sp2 in GROUPS.items():
        if _g2 == gkey or not _sp2.get('amort'):
            continue
        if _expense_marker(_g2, nm, sm):
            return False
    return True


def _is_expense_subject(nm):
    """一级费用科目白名单判断（用于"非本期费用列支计提"归因）。"""
    return _top_account(nm) in _EXPENSE_TOPS


def build_amort_sheet(wb, year, recs, ents, gkey):
    sp = GROUPS[gkey]
    ws = wb.create_sheet(sp['amort_sheet'])
    yint = int(year)
    src_role = sp['amort_src']

    ent_src_cr = {}
    ent_src_non = {}
    ent_src_non_detail = []   # 非列支计提凭证级明细（差异归因展开，2026-07-31）
    ent_exp = {}
    STD = ['管理费用', '制造费用', '销售费用', '研发费用']

    for ent, rec in recs:
        gl = rec['gl']
        # 凭证级映射：年月|凭证号 -> 该凭证内全部行。铁律：费用列支行必须与
        # 【本组计提来源科目】同一凭证（即该摊销分录的对方），排除"待摊费用/长期待摊
        # 费用/利息"等其他"摘要含摊销"但与本组无关的分录（FY GL 对方科目列为空）。
        # 键含【年月】：U8 凭证号每月重置（各月都有"记-85"），跨月同号必须隔离。
        # 混合凭证（多笔摊销同号）再按【金额配对】判定。
        vouch_rows = defaultdict(list)
        for r in gl:
            vk = _vouch_key(r)
            if vk:
                vouch_rows[vk].append(r)

        def _is_amort_partner(r):
            """该借方费用行是否为【本组计提来源科目】的摊销对方。
            判定规则：
              · 凭证内存在本组计提来源科目（备抵贷方）为前提；
              · 若凭证内【仅本组一个计提来源组】（纯本组摊销凭证，如折旧凭证借多行贷多行，
                各行金额不对应）→ 存在性即可（全部费用行计入）；
              · 若凭证内【混有多组计提来源】（如 FY 皮革记-91 同时含软件摊销+装修摊销）
                → 需【金额配对】（1:1 / 1:多 / 多:多 子集），排除其他组的摊销行。
            ⚡⚡ 2026-08-28 修复（对应科目未找全）：SAP 存在【费用科目间结转】凭证
            （借 研究开发费-无形资产摊销 / 贷 无形资产摊销；借 研究开发费-折旧费用与
            长期费用摊销 / 贷 折旧费）——贷方是【费用科目】而非备抵，原判定因无备抵贷方
            而排除 → 研发部分列支全丢、转出科目未扣减。扩展：凭证内存在本组【费用科目贷方】
            （即结转凭证）亦视为相关。"""
            vk = _vouch_key(r)
            if not vk:
                return False
            amt = abs(float(r.get('debit') or 0)) or abs(float(r.get('credit') or 0))
            if amt <= 0:
                return False
            src_grps = defaultdict(list)  # gkey -> [备抵贷方金额]
            fee_dr_tops = set()           # 本组费用科目【借方】top（纯结转凭证信号，2026-08-28）
            fee_cr_tops = set()           # 本组费用科目【贷方】top
            for x in vouch_rows.get(vk, []):
                g2, rl2 = _grp_cat((x.get('name') or '').strip())
                xnm = (x.get('name') or '').strip()
                if rl2 in ('accum',) or (rl2 == 'cost' and g2 == 'LTDEF'):
                    c = abs(float(x.get('credit') or 0))
                    if c > 0:
                        src_grps.setdefault(g2, []).append(c)
                elif g2 != gkey and rl2 is None:
                    # 本组费用科目行：记录借贷两侧 top。⚠️ 结转判定必须【借贷双向都是本组
                    # 费用科目】（借 研发摊销 / 贷 无形资产摊销 = 纯 INT 结转）。
                    # ⚡⚡ 2026-08-28 修复（跨组结转误判）：SAP 跨组重分类凭证
                    #   （借 研发折旧+研发摊销 / 贷 折旧费，及 5月冲回 借 折旧费 / 贷
                    #   研发折旧+研发摊销）贷方/借方涉及【折旧费】（非本组费用科目）——
                    #   若按单侧费用科目即认定结转，会把研发摊销 10,183 计入 INT 列支
                    #   造成 4/5 月 ±10,183 月度错配。跨组重分类整体不进入本组列支净额。
                    # ⚡⚡ 2026-08-28 修复（误收）：结转判定须【科目名】含摊销/折旧动词，
                    #   不能靠摘要——SAP『递延收益摊销→其他收益』凭证（借 递延收益-搬迁
                    #   收益 427,067 / 贷 其他收益-政府补助）摘要"搬迁递延收益摊销"含
                    #   "摊销"但【贷方科目是其他收益、非摊销费用科目】→ 用带摘要的
                    #   _expense_marker 会把"其他收益"误判为 INT 费用行 → 结转信号
                    #   误触发 → 其他收益被当摊销扣减（核对差 -427,067 根因）。
                    #   用空摘要调用 marker（仅科目名判定）：其他收益→False、无形资产
                    #   摊销→True、折旧费(FA)→True。
                    if _expense_marker(gkey, xnm, ''):
                        if float(x.get('debit') or 0) != 0:
                            fee_dr_tops.add(_top_account(xnm))
                        if float(x.get('credit') or 0) != 0:
                            fee_cr_tops.add(_top_account(xnm))
            # 结转凭证信号：借贷双向均存在本组费用科目（纯结转）才认定，无备抵贷方也可；
            # 单向（仅借方或仅贷方）属跨组重分类，不认定（避免 4/5 月 ±10,183 月度错配）
            is_carryover = bool(fee_dr_tops and fee_cr_tops)
            if not src_grps and not is_carryover:
                return False
            srcs = src_grps.get(gkey, [])
            # ⚡ 2026-08-28 结转凭证（贷方为费用科目、无本组备抵贷方）：该凭证借方
            #   费用行即本组研发列支，直接相关（存在性即可）。
            if is_carryover and not srcs:
                return True
            if not srcs:
                return False
            # ① 纯本组凭证：存在性即可（折旧凭证借多行贷多行，各行金额不对应）
            # ② FA/ROU/IPR/INT：_expense_marker 已按 组名+关键词 精确互斥
            #    （'折旧'排除使用权资产/投资性房地产；'使用权资产'/'租赁'排除FA），
            #    故混合凭证（如 FA+ROU 同号）中本组费用行也全部计入，无需金额配对
            #    ⚡⚡ 2026-08-27 修复（1010 无形资产摊销核对差 502万 根因）：SAP 1010
            #    每月计提为【混合凭证】2200000000 同含 FA/IPR/INT 三组备抵贷方（借 折旧费
            #    小额多行 / 贷 累计折旧+累计摊销），INT 原不在豁免列表 → 走③金额配对 →
            #    '折旧费'小额行(278 等)与 INT 备抵金额对不上 → 摊销列支全漏(计提622万 vs
            #    列支120万)。SAP 摊销借方是通用"折旧费"（不含"摊销"字样，不会被 LTDEF
            #    误收），INT 加入存在性豁免是安全的。
            if len(src_grps) == 1 or gkey in ('FA', 'ROU', 'IPR', 'INT'):
                return True
            # ③ INT/LTDEF：摊销关键词组间不互斥（"摊销装修费"既似长摊又似无形），
            #    多组混合凭证需金额配对（1:1 / 1:多 / 多:多 子集），排除其他组摊销行
            if any(abs(c - amt) < 0.01 for c in srcs):       # 1:1
                return True
            if abs(amt - sum(srcs)) < 0.01:                   # 1:多
                return True
            if len(srcs) <= 5:                                # 多:多 子集
                for k in range(2, len(srcs) + 1):
                    for comb in itertools.combinations(srcs, k):
                        if abs(amt - sum(comb)) < 0.01:
                            return True
            return False

        def _has_src_debit_in_vouch(r):
            """同凭证内是否存在与本行【金额相等】的本组备抵【借方】行（借A=X 贷B=X 对冲）。
            用途：计提行排除组内备抵互转（借累计折旧 贷累计折旧）。不依赖"对方科目"列文本
            （FY 2026 GL 该列误装摘要）。
            注意：计提折旧与资产处置（借累计折旧 贷固定资产）可同凭证——借方金额≠贷方计提额
            时不算互转，计提必须保留（绍兴记-85 场景：计提 103,261.59 + 报废转出 235,002.04）。"""
            vk = _vouch_key(r)
            if not vk:
                return False
            amt = abs(float(r.get('credit') or 0))
            if amt <= 0:
                return False
            for x in vouch_rows.get(vk, []):
                if x is r:
                    continue
                g2, rl2 = _grp_cat((x.get('name') or '').strip())
                # ⚡ 2026-08-10 跨组备抵互转排除：SAP 1010 FA→IPR 重分类凭证
                # （6500000091/0092：借 累计折旧-房屋及建筑物 1,320 万 / 贷 投资性房地产
                # -累计折旧 1,320 万）——借方是 FA 组备抵、贷方是 IPR 组备抵，金额相等，
                # 属【重分类转入的累计折旧】而非本期计提，原 `g2 == gkey` 只认同组导致
                # 1,320 万被误计入 IPR 计提。改为：借方只要【备抵类】(rl2==src_role)
                # 且金额相等即视为互转/重分类（借备抵贷备抵同向对冲必非计提）。
                if rl2 == src_role:
                    d = abs(float(x.get('debit') or 0))
                    if d > 0.01 and abs(d - amt) < 0.01:
                        return True
            return False

        # 凭证级"非列支计提"归因：vouch -> {src:本组备抵贷方合计, exp:费用科目借方合计, m:月份}
        vouch_meta = defaultdict(lambda: {'src': 0.0, 'exp': 0.0, 'm': None})

        for r in gl:
            g, rl = _grp_cat(r['name'])
            nm = (r['name'] or '').strip(); sm = (r['sm'] or '').strip()
            # ⚡⚡ 2026-08-28 对应科目罗列：费用行识别由 marker 白名单【扩展】为
            #   "marker 或 归本组的 STD 外费用/成本类 top"（主营业务成本-光伏成本/研发支出
            #   -直接投入等）。仅当凭证级 _is_amort_partner 判定相关时才会被收集（下 else
            #   分支再判），故纯销售成本结转凭证（贷方存货、无备抵）不会误收；且被其他
            #   资产组 marker 认领的费用行（如投房折旧→其他业务成本-租金）不归本组。
            is_exp = _expense_marker(gkey, nm, sm) or _is_cost_subject_own(gkey, nm, sm)
            # 仅计入两类凭证行：①本组自身科目(累计折旧/摊销贷方=计提来源；原值/清理等)
            # ②本组折旧/摊销的"费用列支"行(管理费用/制造费用等费用科目借方,科目不属于本组资产账户)；
            # 费用列支取【净】发生额(含红字冲回/错账调整等负借方)，与《科目余额表》借发(净额)勾稽一致
            # ⚡⚡ 2026-08-28 修复（对应科目未找全）：费用列支行【借贷两向】均须进入——
            #    SAP 费用科目间结转凭证（借 研发摊销 / 贷 无形资产摊销）的【贷方】行
            #    用于冲减原科目列支，原 `r['debit'] != 0` 把贷方结转行提前 continue 掉
            #    → 转出部分未从原 top 扣减、研发部分多计。改为：费用行只要有任一方向
            #    发生额且与本组相关（_is_amort_partner 内再判）即进入 else 分支取净额。
            if not (g == gkey or (is_exp and r['debit'] != 0 or (is_exp and r['credit'] != 0))):
                continue
            ym = _ym(r['date'])
            if ym is None or ym[0] != yint:
                continue
            m = ym[1]
            cp = r['cp'] or ''
            if rl == src_role:
                # 计提数取【净】发生额（含红字冲回/错账调整等负贷方），与《科目余额表》贷发(净额)勾稽一致
                if r['credit'] != 0:
                    # 组内备抵互转排除（借A=X 贷B=X 对冲，非真实计提）：凭证级+金额配对判定。
                    # 不依赖"对方科目"列文本（FY 2026 GL 该列误装摘要）；计提折旧与资产处置
                    # （借累计折旧 贷固定资产）同凭证时金额不等 → 不算互转，计提保留。
                    if _has_src_debit_in_vouch(r):
                        continue
                    ent_src_cr[(ent, m)] = ent_src_cr.get((ent, m), 0.0) + r['credit']
                    vk_meta = _vouch_key(r)
                    if vk_meta:
                        vouch_meta[vk_meta]['src'] += r['credit']
                        vouch_meta[vk_meta]['m'] = m
            else:
                # 仅"费用列支"行(科目不属于本组资产账户)才计入各计提对应科目金额；
                # 取【净】发生额（含红字冲回/错账调整等负借方、费用科目间结转贷方），
                # 与《科目余额表》借发(净额)勾稽一致
                # 凭证级判定：必须与本组计提来源科目同凭证（摊销分录的对方），
                # 排除"待摊费用/长期待摊费用/利息"等摘要含"摊销"但与本组无关的分录
                # ⚡⚡ 2026-08-28 修复（对应科目未找全）：SAP 费用科目间结转凭证
                #   （借 研究开发费-无形资产摊销 / 贷 无形资产摊销；借 研究开发费-折旧
                #   费用与长期费用摊销 / 贷 折旧费）——贷方是【费用科目】而非备抵，原判定
                #   因无备抵贷方而排除 → 研发部分列支全丢、转出科目未扣减。现对【借贷两向】
                #   统一归集净额：借方计入 top、贷方从 top 扣减（结转），研究开发费→研发费用列。
                if is_exp and g != gkey and _is_amort_partner(r):
                    top = _top_account(nm)
                    # ROU 折旧分摊表：排除对方科目为 FA 组（固定资产累计折旧）的费用行，
                    # 避免摘要含"折旧"的 FA 折旧行被误计入 ROU 列支（核对差-38M 根因）
                    if gkey == 'ROU' and any(_grp_cat(p.strip())[0] == 'FA' for p in cp.split(',') if p.strip()):
                        continue
                    # 研发费用/研发支出之间的结转(借研发费用贷研发支出 或反向)属重分类,
                    # 不计入折旧/摊销费用列支,避免重复统计
                    if top in ('研发费用', '研发支出') and ('研发支出' in cp or '研发费用' in cp):
                        continue
                    # 折旧/摊销核对表的对应费用科目不含资产减值损失/信用减值损失
                    if top in ('资产减值损失', '信用减值损失'):
                        continue
                    # ⚡⚡ 2026-08-28 对应科目归集：研究开发费（SAP 660006）→ 研发费用列
                    #    （与 expense_detail/payroll_detail 一致，08-08 记忆确认 66000605 研发折旧
                    #    66000607 研发摊销 归属研发费用）
                    top = '研发费用' if top == '研究开发费' else top
                    # ⚡⚡ 2026-08-28 折旧费(660007)按功能范围拆分到 STD 一级科目——SAP 6600
                    #   期间费用池混合科目，1010 借方实际分布 管理 1936万/制造 1104万/研发 392万/
                    #   销售 12万（FR 列），原笼统归『折旧费』列、与 STD 科目对不上（用户质询）。
                    #   FR 无值或未命中保持原 top（折旧费列兜底）。
                    if top == '折旧费' and r.get('fr'):
                        _fr4 = str(r.get('fr')).strip()[:4]
                        _FR2STD = {'6602': '管理费用', '5101': '制造费用',
                                   '6601': '销售费用', '5301': '研发费用'}
                        if _fr4 in _FR2STD:
                            top = _FR2STD[_fr4]
                    net = (r['debit'] or 0) - (r['credit'] or 0)
                    # ⚡⚡ 2026-08-28 修复（跨组结转月度错配 04/05 月 ±10,183）：SAP
                    #   跨组重分类凭证（借 研发折旧43,644+研发摊销10,183 / 贷 折旧费53,827，
                    #   及 5月冲回）——贷方【折旧费】含其他组（摊销）成分，FA 侧若全额扣减
                    #   折旧费会多扣 10,183（摊销成分），且 5月冲回又全额计入 → 04/05 月
                    #   核对差 ±10,183。修复：纯结转凭证（无备抵贷方、借贷双向均含本组费用
                    #   科目）内，本组费用行的计入按【凭证内本组借贷金额比例】封顶——
                    #   贷方行扣减 ≤ 凭证内本组费用借方合计、借方行计入 ≤ 凭证内本组费用
                    #   贷方合计，跨组成分留给他组处理，本组月度即平。
                    # ⚡⚡ 2026-08-28 修复（计提凭证误扣）：比例封顶仅对【纯结转凭证】生效——
                    #   计提凭证（含本组备抵贷方）若也封顶，会把"累计折旧-XX"备抵贷方行
                    #   （经 _is_amort_partner 判定为相关）按 _dr 扣减 → 折旧费列支少计
                    #   （3010 计提 77,221 折旧费列支 52,132，差 25,089 根因）。故先判
                    #   src_grps 是否为空（纯结转 = 无备抵贷方）才封顶。
                    if vk_meta := _vouch_key(r):
                        _has_accum = any(rl2 == 'accum' or (rl2 == 'cost' and g2 == 'LTDEF')
                                         for x in vouch_rows.get(vk_meta, [])
                                         for g2, rl2 in [_grp_cat((x.get('name') or '').strip())])
                        if not _has_accum:
                            _dr = sum(abs(float(x.get('debit') or 0))
                                      for x in vouch_rows.get(vk_meta, [])
                                      if (x.get('name') or '').strip() != nm
                                      and _grp_cat((x.get('name') or '').strip())[0] is None
                                      and _expense_marker(gkey, (x.get('name') or '').strip(), ''))
                            _cr = sum(abs(float(x.get('credit') or 0))
                                      for x in vouch_rows.get(vk_meta, [])
                                      if (x.get('name') or '').strip() != nm
                                      and _grp_cat((x.get('name') or '').strip())[0] is None
                                      and _expense_marker(gkey, (x.get('name') or '').strip(), ''))
                            if abs(r['debit'] or 0) > 0 and _cr > 0:
                                # 借方行：计入不超过凭证内本组费用贷方合计
                                net = min(r['debit'], _cr)
                            elif abs(r['credit'] or 0) > 0 and _dr > 0:
                                # 贷方行：扣减不超过凭证内本组费用借方合计
                                net = -min(r['credit'], _dr)
                    if net != 0:
                        ent_exp[(ent, m, top)] = ent_exp.get((ent, m, top), 0.0) + net
                        vk_meta = _vouch_key(r)
                        if vk_meta:
                            vouch_meta[vk_meta]['exp'] += net

        # 非本期费用列支计提（期初确认/审计调整/补提/重分类）：凭证内本组备抵贷方超过
        # 费用科目借方合计的部分（借方为使用权资产原值/租赁负债/固定资产原值等非费用科目）
        for vk_meta, meta in vouch_meta.items():
            if meta['m'] is None:
                continue
            non = max(0.0, meta['src'] - meta['exp'])
            if non > 0.01:
                ent_src_non[(ent, meta['m'])] = ent_src_non.get((ent, meta['m']), 0.0) + non
                # 凭证级明细：日期/凭证号/金额/对方科目/费用借方行摘要，供披露行展开差异原因
                _detail = {
                    'ent': ent, 'm': meta['m'], 'amt': non,
                    'date': '', 'vtype': '', 'vno': '', 'opp': '', 'summ': '',
                }
                for x in vouch_rows.get(vk_meta, []):
                    gx, rlx = _grp_cat((x.get('name') or '').strip())
                    if gx == gkey and rlx == src_role:
                        if not _detail['date']:
                            _detail['date'] = str(x.get('date') or '')
                            _detail['vtype'] = str(x.get('voucher') or '')
                        _detail['summ'] = str(x.get('sm') or '') or _detail['summ']
                        _detail['opp'] = str(x.get('cp') or '') or _detail['opp']
                    elif abs(float(x.get('debit') or 0)) > 0.01 and not _detail['opp']:
                        _detail['opp'] = str(x.get('cp') or '') or _detail['opp']
                ent_src_non_detail.append(_detail)

    all_tops = sorted({t for (_, _, t) in ent_exp})
    exp_accounts = STD + [t for t in all_tops if t not in STD]
    # ⚡⚡ 2026-08-28 对应科目罗列：列名区分费用化/资本化研发——『研发支出』(5301,资本化)
    #   单列 "研发支出(资本化)"，与『研发费用』(660006 研究开发费,费用化)分列，避免混淆。
    disp_top = {t: ('研发支出(资本化)' if t == '研发支出' else t) for t in exp_accounts}
    all_ents = sorted({e for (e, _) in ent_src_cr} | {e for (e, _, _) in ent_exp})

    verb = '摊销' if gkey in ('INT', 'LTDEF') else '折旧'
    # 计提数列名：用【实际账套备抵科目名】（如"使用权资产折旧"），动态发现，而非写死的通用名
    if src_role == 'accum':
        acc_label = _discover_accum_name(recs, gkey)
    else:
        acc_label = sp['title']
    headers = ['主体', '月份', '计提数(%s贷方)' % acc_label] + [disp_top[t] for t in exp_accounts] + \
              ['列支合计', '核对(计提-列支)']
    ncols = len(headers)
    r = _title(ws, '%s（%s 年度）' % (sp['amort_sheet'], year), ncols)
    r = _sub(ws, '口径:每月"计提数"=%s(来源)%s当月贷方合计(净计提);"对应科目"按一级科目动态罗列'
             '(标准:管理费用/制造费用/销售费用/研发费用;本组归集:折旧费/无形资产摊销/租赁费等;'
             'STD 外:主营业务成本/其他业务成本/研发支出(资本化)等,凡计提凭证借方实际计入科目均罗列),'
             '为该月%s实际计入科目净额。费用列支识别采用【凭证级判定】:仅统计与本组%s同一凭证的借方行,'
             '可排除"待摊费用/长期待摊费用/利息"等摘要含%s但与本组无关的分录。'
             '核对=计提数-列支合计,差额为%s计入其他科目或非本期费用列支(已披露)。'
             '2026 年 GL 仅 1–5 月(YTD)。' % (
                 sp['title'], '累计%s' % verb if src_role == 'accum' else '原值(减少)',
                 verb, '累计%s' % verb if src_role == 'accum' else '原值', verb, verb), ncols, r)
    # ⚡⚡ 2026-08-27 披露：SAP 无形资产摊销部分计入『折旧费』通用科目（与固定资产折旧
    #   合并核算），核对表列支仅认『无形资产摊销』科目 → 计提-列支差异为该部分，
    #   属科目归集差异（非账务差错），在此显式披露供审计判断。
    if gkey == 'INT':
        r = _sub(ws,
                 '注：SAP 摊销部分计入『折旧费』通用科目（与固定资产折旧合并核算，无法按'
                 '科目拆分）；本表对应科目仅统计『无形资产摊销』等明确摊销科目，核对差异即为'
                 '计入『折旧费』的摊销额（科目归集差异，非账务差错）。',
                 ncols, r)
    r = _sub(ws, '注：『研发支出』(5301,资本化归集)与『研发费用』(660006 研究开发费,费用化)为两套独立'
                '科目，本表单列区分；两者之间结转（借 研发费用 贷 研发支出 或反向）属重分类，'
                '不计入列支；折旧/摊销核对表的对应科目不含资产减值损失。', ncols, r)
    hr = _hdr(ws, r, headers)
    r = hr

    col_acc = 3
    col_top = {top: 4 + i for i, top in enumerate(exp_accounts)}
    col_lie = 4 + len(exp_accounts)
    col_chk = col_lie + 1

    gt_acc = gt_lie = 0.0
    for ent in all_ents:
        ent_months = sorted({m for (e, m) in ent_src_cr if e == ent} |
                            {m for (e, m, _) in ent_exp if e == ent})
        e_acc = e_lie = 0.0
        for m in ent_months:
            acc = ent_src_cr.get((ent, m), 0.0)
            lie = sum(ent_exp.get((ent, m, s), 0.0) for s in exp_accounts)
            chk = round(acc - lie, 2)
            _txt(ws, r, 1, ent)
            _txt(ws, r, 2, '%02d月' % m)
            _money(ws, r, col_acc, acc)
            for top in exp_accounts:
                _money(ws, r, col_top[top], ent_exp.get((ent, m, top), 0.0))
            _money(ws, r, col_lie, lie)
            _money(ws, r, col_chk, chk)
            e_acc += acc; e_lie += lie
            r += 1
        _txt(ws, r, 1, ent, bold=True)
        _txt(ws, r, 2, '小计', bold=True)
        _money(ws, r, col_acc, e_acc)
        for top in exp_accounts:
            tv = sum(ent_exp.get((ent, mm, top), 0.0) for mm in ent_months)
            _money(ws, r, col_top[top], tv)
        _money(ws, r, col_lie, e_lie)
        _money(ws, r, col_chk, round(e_acc - e_lie, 2))
        for i in range(1, ncols + 1):
            ws.cell(r, i).font = SHELL_BOLD
            ws.cell(r, i).fill = SHELL_TOT_FILL
        gt_acc += e_acc; gt_lie += e_lie
        r += 1
        prior = _prior_year_acc_group(ents, ent, str(int(year) - 1), gkey)
        _txt(ws, r, 1, ent)
        _txt(ws, r, 2, '上年计提小计')
        _money(ws, r, col_acc, prior)
        r += 1
        diff = round(e_acc - prior, 2) if prior is not None else None
        _txt(ws, r, 1, ent)
        _txt(ws, r, 2, '本期小计-上期小计')
        _money(ws, r, col_acc, diff)
        r += 1

    _txt(ws, r, 1, '合计', bold=True)
    _txt(ws, r, 2, '%s年' % year, bold=True)
    _money(ws, r, col_acc, gt_acc)
    for top in exp_accounts:
        tv = sum(v for (e, mm, t), v in ent_exp.items() if t == top)
        _money(ws, r, col_top[top], tv)
    _money(ws, r, col_lie, gt_lie)
    _money(ws, r, col_chk, round(gt_acc - gt_lie, 2))
    for i in range(1, ncols + 1):
        ws.cell(r, i).font = SHELL_BOLD
    r += 1

    # 非本期费用列支计提披露（期初确认/审计调整/补提/重分类：借方为非费用科目）
    gt_non = sum(v for v in ent_src_non.values())
    if abs(gt_non) > 0.01:
        _txt(ws, r, 1, '其中：非本期费用列支计提', bold=True)
        _txt(ws, r, 2, '（期初确认/审计调整/补提/重分类）', bold=True)
        _money(ws, r, col_acc, round(gt_non, 2))
        _txt(ws, r, col_chk, '= 计提数 - 列支合计(应等于上方核对差异)')
        for i in range(1, ncols + 1):
            ws.cell(r, i).font = SHELL_BOLD
        r += 1
        # 凭证级差异归因展开：逐凭证列示 主体/月份/日期/凭证号/金额/对方科目（2026-07-31 用户要求讲清差异原因）
        if ent_src_non_detail:
            _txt(ws, r, 1, '    └ 差异归因（逐凭证）：', bold=True)
            for i in range(1, ncols + 1):
                ws.cell(r, i).font = Font(name='Times New Roman', size=10, color='606060')
            r += 1
            _hdr(ws, r, ['主体', '月份', '凭证日期', '凭证字-号', '非列支金额',
                         '费用侧借方行摘要/对方科目'] + [''] * (ncols - 6))
            r += 1
            for d in sorted(ent_src_non_detail, key=lambda x: (str(x['ent']), x['m'], str(x['date']))):
                _txt(ws, r, 1, d['ent'])
                _txt(ws, r, 2, '%02d月' % d['m'])
                _txt(ws, r, 3, str(d['date'])[:10])
                _txt(ws, r, 4, d['vtype'])
                _money(ws, r, 5, round(d['amt'], 2))
                _txt(ws, r, 6, ('%s；对方:%s' % (d['summ'], d['opp']))[:80])
                r += 1
            r += 1

    residual = round(gt_acc - gt_lie, 2)
    res_str = '%.2f' % residual
    non_str = '其中"非本期费用列支计提"%.2f元（期初确认/审计调整/补提/重分类，借方为原值/租赁负债等非费用科目，不计入费用列支）' % gt_non \
        if abs(gt_non) > 0.01 else ''
    direction = ('为%s计提计入上述费用之外的科目(如其他业务成本、资本化研发等),已披露,不认定为账务差错。'
                 % verb) if residual >= -0.005 else (
        '为负数,表明费用列支合计大于计提数:系部分标准费用科目行摘要误标"%s"被计入,已披露,不认定为账务差错。'
        % verb)
    note = ('勾稽说明:每月"计提数"(%s贷方)应等于按一级费用科目展开的"费用列支数"之和(列支合计),'
            '核对=计提数-列支合计。本client%s费用依"科目名/摘要含%s"识别(对方科目多为空)。'
            '若核对存在差额(如 %s 元),请查明原因。%s'
            % (verb, verb, verb, res_str, non_str))
    ws.cell(r, 1, note).font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 2

    widths = [14, 8, 16] + [14] * len(exp_accounts) + [14, 16, 18, 18]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'C4'
    return ws


# ----------------------------------------------------------------------------
# Sheet: 增加检查表（含「在建工程转入」分类与汇总）
# ----------------------------------------------------------------------------
_VOUCH_STD_HDR = ['核算主体', '测试序号', '日期', '凭证字', '凭证号', '二级科目', '摘要',
                  '借方金额', '贷方金额', '对方科目',
                  '与原始凭证相符', '原始凭证内容', '原始凭证日期', '会计处理正确', '所属时间无误']


def _write_std_vouch_sheet(ws, sp, year, rows, side):
    """标准 15 列凭证抽查模板（2026-07-31 用户要求：在建工程增加/减少检查表采用普通凭证抽查模板）。
    rows: [dict(ent, date, vt, no, name, sm, cp, amt)]；side='inc' 金额列=借方(8)、'dec'=贷方(9)。
    返回下一空行。"""
    ncols = 15
    sheet_label = sp['inc_sheet'] if side == 'inc' else sp['dec_sheet']
    r = _title(ws, '%s（%s 年度）' % (sheet_label, year), ncols)
    # 抽凭表说明已按用户要求删除（2026-07-31）
    hr = _hdr(ws, r, _VOUCH_STD_HDR)
    r = hr
    amt_col = 8 if side == 'inc' else 9
    sd = 0.0
    for i, d in enumerate(rows, 1):
        _txt(ws, r, 1, d['ent'])
        _txt(ws, r, 2, i)
        _txt(ws, r, 3, str(d['date'])[:10] if d['date'] else '')
        _txt(ws, r, 4, d.get('vt', ''))
        _txt(ws, r, 5, d.get('no', ''))
        _txt(ws, r, 6, d.get('name', ''))
        _txt(ws, r, 7, (d.get('sm', '') or '')[:60])
        if side == 'inc':
            _money(ws, r, 8, d['amt'])
            _txt(ws, r, 9, None)
        else:
            _txt(ws, r, 8, None)
            _money(ws, r, 9, d['amt'])
        _txt(ws, r, 10, d.get('cp', ''))
        for c in range(11, 16):
            _txt(ws, r, c, None)
        sd += d['amt']
        r += 1
    if not rows:
        _txt(ws, r, 1, '（本年度无%s业务或 GL 无相关数据）' % sp['title'])
        r += 1
    else:
        _txt(ws, r, 1, '合计', bold=True, fill=SHELL_TOT_FILL)
        _txt(ws, r, 2, '%d 笔' % len(rows), bold=True, fill=SHELL_TOT_FILL)
        _mc = _money(ws, r, amt_col, sd)
        _mc.fill = SHELL_TOT_FILL
        for c in range(1, ncols + 1):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    for c, w in zip('ABCDEFGHIJKLMNO', [16, 10, 12, 10, 10, 22, 30, 14, 14, 24, 16, 16, 16, 16, 16]):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    return r


_FEE_DEP_KW = ('折旧', '摊销', '租赁费')
# 费用侧汇总排除的资产负债类 top（名称含"摊销"如 待摊费用，或含"租赁费"如 应付租赁费，
# 但均非损益费用科目，罗列会污染费用侧口径；累计*/资产原值已由 _grp_cat 排除）。
_FEE_DEP_EXCLUDE_TOPS = ('待摊费用', '长期待摊费用', '预提费用', '其他应付款', '应付账款',
                         '应付票据', '预付账款', '预收账款', '应收账款', '应收票据',
                         '其他应收款', '应交税费', '专项应付款', '递延收益', '递延所得税',
                         '递延所得税资产', '递延所得税负债', '银行存款', '库存现金',
                         '长期应付款', '租赁负债', '使用权资产')


def build_fee_dep_summary_sheet(wb, year, recs):
    """费用侧【费用中摊销折旧计提汇总】（用户 2026-08-28 方法论）：
    把同时记入费用的各类折旧/摊销/租赁费，从 GL 按【主体×一级费用科目】净额罗列汇总。
    口径：全 GL 中名称含 折旧/摊销/租赁费 的【费用/成本科目】（排除累计折旧/累计摊销等
    备抵、排除资产原值/在建工程/长期待摊费用等资产侧科目），net = 借方-贷方。
    与资产侧《折旧/摊销分摊表》同源同口径呼应，不追求对平，罗列即交付。"""
    from collections import OrderedDict
    ws = wb.create_sheet('费用中摊销折旧计提汇总')
    ent_top = OrderedDict()          # (ent, top) -> [dr, cr]
    ent_set = set()
    dep_ents = set()
    for ent, rec in recs:
        gl = rec.get('gl') or []
        if not gl:
            continue
        for r in gl:
            nm = (r.get('name') or '').strip()
            if not nm or not any(k in nm for k in _FEE_DEP_KW):
                continue
            gx, rlx = _grp_cat(nm)
            # 排除备抵/清理/减值/资产原值（cost）——仅收费用/成本类科目
            if rlx in ('accum', 'clear', 'impair') or rlx == 'cost':
                continue
            if gx is not None:
                continue
            # ⚡⚡ 2026-08-28 排除资产负债类 top（待摊费用/其他应付款-应付租赁费等）：
            #   名称虽含 摊销/租赁费 关键词但属资产/负债，非损益费用，不得入费用侧罗列。
            if _top_account(nm) in _FEE_DEP_EXCLUDE_TOPS:
                continue
            db = float(r.get('debit') or 0); cr = float(r.get('credit') or 0)
            if db <= 0 and cr <= 0:
                continue
            top = _top_account(nm)
            key = (ent, top)
            ent_top.setdefault(key, [0.0, 0.0])
            ent_top[key][0] += db; ent_top[key][1] += cr
            ent_set.add(ent)
    if not ent_top:
        ws.cell(1, 1, '（本年度 GL 无折旧/摊销/租赁费相关费用科目）')
        return 0
    tops = sorted({t for (_, t) in ent_top},
                  key=lambda t: -sum(v[0] - v[1] for (e, tt), v in ent_top.items() if tt == t))
    std = ('折旧费', '无形资产摊销', '研究开发费', '租赁费', '管理费用', '制造费用',
           '销售费用', '研发费用')
    tops = [t for t in std if t in tops] + [t for t in tops if t not in std]
    headers = ['主体'] + tops + ['净额合计']
    ncols = len(headers)
    r = _title(ws, '费用中摊销折旧计提汇总（%s 年度）' % year, ncols)
    r = _sub(ws, '口径：GL 中名称含 折旧/摊销/租赁费 的费用/成本科目（备抵/资产原值已剔除），'
             '按主体×一级科目净额(借-贷)罗列。与各资产底稿《折旧/摊销分摊表》同源呼应；'
             '研发支出(5301)=资本化归集，研发费用(660006研究开发费)=费用化，两者单列。', ncols, r)
    hr = _hdr(ws, r, headers)
    r = hr
    col_top = {t: 2 + i for i, t in enumerate(tops)}
    col_tot = 2 + len(tops)
    gt = [0.0] * len(tops)
    for ent in sorted(ent_set):
        vals = [ent_top.get((ent, t), [0.0, 0.0])[0] - ent_top.get((ent, t), [0.0, 0.0])[1] for t in tops]
        _txt(ws, r, 1, ent)
        for i, t in enumerate(tops):
            if vals[i] != 0:
                _money(ws, r, col_top[t], round(vals[i], 2))
                gt[i] += vals[i]
        _money(ws, r, col_tot, round(sum(vals), 2))
        r += 1
    _txt(ws, r, 1, '合计', bold=True, fill=SHELL_TOT_FILL)
    for i, t in enumerate(tops):
        if gt[i] != 0:
            _mc = _money(ws, r, col_top[t], round(gt[i], 2))
            _mc.fill = SHELL_TOT_FILL
    _mc = _money(ws, r, col_tot, round(sum(gt), 2))
    _mc.fill = SHELL_TOT_FILL
    for c in range(1, ncols + 1):
        ws.cell(r, c).font = SHELL_BOLD
    r += 1
    # 类别标注（哪些科目属折旧/摊销/租赁），供审计师分辨计提类别
    _txt(ws, r, 1, '类别标注：', bold=True)
    _txt(ws, r, 2, ('折旧类→ ' + '、'.join(t for t in tops if '折旧' in t or '租赁' in t)) or '无', ncols)
    r += 1
    _txt(ws, r, 2, ('摊销类→ ' + '、'.join(t for t in tops if '摊销' in t)) or '无', ncols)
    r += 1
    _txt(ws, r, 2, ('其他类→ ' + '、'.join(t for t in tops if '折旧' not in t and '摊销' not in t and '租赁' not in t)) or '无', ncols)
    r += 1
    for i, c in enumerate('ABCDEFGHIJKLMNOPQRSTUVWXYZ'[:ncols]):
        ws.column_dimensions[c].width = 14
    ws.cell(1, 1, '费用中摊销折旧计提汇总（%s 年度）' % year).font = Font(bold=True, size=12)
    ws.freeze_panes = 'A%d' % (hr + 1)
    return r


def build_inc_sheet(wb, year, recs, gkey):
    sp = GROUPS[gkey]
    # 在建工程：2026-07-31 用户要求改用普通凭证抽查模板（标准 15 列）
    if gkey == 'CIP':
        ws = wb.create_sheet(sp['inc_sheet'])
        rows = []
        for ent, rec in recs:
            for line in rec['gl']:
                ym = _ym(line['date'])
                if ym is None or ym[0] != int(year):
                    continue
                g, rl = _grp_cat(line['name'])
                if g != gkey or rl != 'cost' or line['debit'] <= 0:
                    continue
                if _is_same_group(gkey, line['cp']):
                    continue  # 对方科目为同组科目，属组内重分类/调整，非真实增加
                vt = no = ''
                if line['voucher'] and '-' in line['voucher']:
                    vt, no = line['voucher'].split('-', 1)
                rows.append({'ent': ent, 'date': line['date'], 'vt': vt, 'no': no,
                             'name': line['name'], 'sm': line['sm'], 'cp': line['cp'],
                             'amt': line['debit']})
        _write_std_vouch_sheet(ws, sp, year, rows, 'inc')
        return
    ws = wb.create_sheet(sp['inc_sheet'])
    ncols = 22
    r = _title(ws, '%s增加检查表（%s 年度）' % (sp['title'], year), ncols)
    r = _sub(ws, '「取得方式」含「在建工程转入」（从在建工程转入数），表底附其汇总。', ncols, r)
    main = ['主体', '日期', '取得方式', '类别', '增加情况', '', '凭证种类', '凭证编号',
            '摘要（业务内容）', '对方科目', '金额', '', '核对内容', '', '', '', '', '', '', '', '备注', '索引']
    sub = ['', '', '', '', '数量', '原值', '', '', '', '', '借方', '贷方', '1', '2', '3', '4', '5', '6', '7', '8', '', '']
    hr = r
    for j, h in enumerate(main, 1):
        if h == '':
            continue
        c = ws.cell(hr, j, h)
        c.fill = SHELL_HFILL; c.font = SHELL_HFONT; c.border = SHELL_BORDER; c.alignment = SHELL_CEN
    sr = hr + 1
    for j, h in enumerate(sub, 1):
        if h == '':
            continue
        c = ws.cell(sr, j, h)
        c.fill = SHELL_HFILL; c.font = SHELL_HFONT; c.border = SHELL_BORDER; c.alignment = SHELL_CEN
    ws.merge_cells(start_row=hr, start_column=5, end_row=hr, end_column=6)
    ws.merge_cells(start_row=hr, start_column=11, end_row=hr, end_column=12)
    ws.merge_cells(start_row=hr, start_column=13, end_row=hr, end_column=20)
    r = sr + 1

    INC_LBL = dict(zip(INC_KEYS, INC_LABELS[gkey]))
    rows = []
    inproject_total = 0.0
    for ent, rec in recs:
        for line in rec['gl']:
            g, rl = _grp_cat(line['name'])
            if g != gkey or rl != 'cost' or line['debit'] <= 0:
                continue
            if _is_same_group(gkey, line['cp']):
                continue  # 对方科目为同组科目(含累计折旧/减值准备/清理),属组内重分类/调整,非真实增加
            nm = (line['name'] or '').strip()
            cat2 = re.sub(r'^(%s[-－]?)+' % '|'.join(re.escape(n) for n in sp['cost_names']), '', nm).strip() or sp['title']
            method = _classify_inc(gkey, line['cp'], line['sm'], 'cost')
            amt = line['debit']
            if method == 'inproject':
                inproject_total += amt
            vt = no = ''
            if line['voucher'] and '-' in line['voucher']:
                vt, no = line['voucher'].split('-', 1)
            rows.append({'ent': ent, 'date': line['date'], 'method': INC_LBL.get(method, '其他'),
                         'cat2': cat2, 'amt': amt, 'vt': vt, 'no': no,
                         'sm': line['sm'], 'cp': line['cp']})
    if not rows:
        _txt(ws, r, 1, '（本年度无%s增加业务或 GL 无相关数据）' % sp['title'])
        r += 1
    for d in rows:
        _txt(ws, r, 1, d['ent'])
        _txt(ws, r, 2, d['date'])
        _txt(ws, r, 3, d['method'])
        _txt(ws, r, 4, d['cat2'])
        _money(ws, r, 6, d['amt'])
        _txt(ws, r, 7, d['vt'])
        _txt(ws, r, 8, d['no'])
        _txt(ws, r, 9, (d['sm'] or '')[:60])
        # ⚡⚡ 2026-08-23 对方科目只保留真实来源（用户方法论）：AFJ 等 U8 对方科目列为
        #   凭证级多科目（含 应交税费/收入/费用/组内/库存现金），增加检查表应剔除
        #   不相关科目，只留 银行存款/应付账款/在建工程 等真实支付或转入来源。
        _txt(ws, r, 10, _filter_inc_cp(d['cp']))
        _money(ws, r, 11, d['amt'])
        r += 1

    r += 1
    # 仅非在建工程（CIP）的底稿才有「在建工程转入」统计（在建工程本身是转出来源）
    if gkey != 'CIP':
        _txt(ws, r, 2, '填表说明：', bold=True)
        _txt(ws, r, 10, '其中：在建工程转入合计')
        _money(ws, r, 11, inproject_total)
        r += 1
        note = ('核对内容:1．与发票是否一致；2．与付款单据是否一致；3．与购买/建造合同是否一致；4．与验收报告或评估报告等是否一致；'
                '5.审批手续是否齐全；6．与在建工程转出数是否一致(见在建工程底稿《在建工程减少检查表》)；'
                '7．会计处理是否正确(入账日期和入账金额)；8．……')
        c = ws.cell(r, 2, note)
        c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncols)

    # ===== 抽凭合计与账面增加核对 =====
    r += 1
    r += 1
    _txt(ws, r, 1, '抽凭合计与账面增加核对', bold=True)
    r += 1
    rh_cols = ['主体', '账面增加（TB借发）', '抽凭增加合计（GL）', '差额']
    for j, h in enumerate(rh_cols, 1):
        cc = ws.cell(r, j, h)
        cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    r += 1
    inc_vouch_total = defaultdict(float)
    for d in rows:
        inc_vouch_total[d['ent']] += d['amt']
    tc_tb = tc_vc = 0.0
    for ent, rec in recs:
        tb = rec['cost_d']['debit'] if rec['cost_d'] else 0.0
        vc = inc_vouch_total.get(ent, 0.0)
        diff = round(vc - tb, 2)
        _txt(ws, r, 1, ent)
        _money(ws, r, 2, tb)
        _money(ws, r, 3, vc)
        _txt(ws, r, 4, '相符' if abs(diff) <= 0.005 else '%.2f（账面负数=红字冲回〔借红字〕；GL 抽凭仅列正数借方行，差额属分类/科目识别差异，非账务差错）' % diff)
        tc_tb += tb; tc_vc += vc
        r += 1
    _txt(ws, r, 1, '合计', bold=True)
    _money(ws, r, 2, tc_tb); _money(ws, r, 3, tc_vc)
    _txt(ws, r, 4, '%.2f' % round(tc_vc - tc_tb, 2))
    for i in range(1, 5):
        ws.cell(r, i).fill = SHELL_TOT_FILL; ws.cell(r, i).font = SHELL_BOLD
    r += 1
    r += 1
    nc = ws.cell(r, 1, '差额≠0 系分类/科目识别差异（GL 为全量抽取）；账面负数=红字冲回（借红字冲减），抽凭仅列正数借方行，非账务差错。')
    nc.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 1

    widths = [14, 12, 12, 16, 8, 14, 10, 10, 28, 20, 14, 14, 4, 4, 4, 4, 4, 4, 4, 4, 14, 14]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (sr + 1)
    return ws


# ----------------------------------------------------------------------------
# Sheet: 减少明细及核对（折旧类科目组）
# ----------------------------------------------------------------------------
def build_disposal_sheet(wb, year, recs, gkey):
    sp = GROUPS[gkey]
    ws = wb.create_sheet(sp['dec_sheet'])
    ncols = 27
    r = _title(ws, '%s（%s 年度）' % (sp['dec_sheet'], year), ncols)
    r = _sub(ws, '每笔%s减少(按凭证)一行;减少净值=原值−累计折旧/摊销−减值准备,与处置损益'
             '(资产处置收益/营业外收支)核对。内部转移单独标识。' % sp['title'], ncols, r)
    main = ['主体', '日 期', '取得日期', '处置方式', '处置日期', '原值', '累计折旧/摊销', '减值准备',
            '账面价值', '处置收入', '净损益', '凭证种类', '凭证编号', '摘要（业务内容）', '资产处置收益',
            '营业外收入', '营业外支出', '其他', '核对内容', '', '', '', '', '', '', '', '备注']
    sub = ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
           '1', '2', '3', '4', '5', '6', '7', '8', '']
    hr = r
    for j, h in enumerate(main, 1):
        if h == '':
            continue
        c = ws.cell(hr, j, h)
        c.fill = SHELL_HFILL; c.font = SHELL_HFONT; c.border = SHELL_BORDER; c.alignment = SHELL_CEN
    sr = hr + 1
    for j, h in enumerate(sub, 1):
        if h == '':
            continue
        c = ws.cell(sr, j, h)
        c.fill = SHELL_HFILL; c.font = SHELL_HFONT; c.border = SHELL_BORDER; c.alignment = SHELL_CEN
    ws.merge_cells(start_row=hr, start_column=19, end_row=hr, end_column=26)
    r = sr + 1

    DEC_LBL = dict(zip(DEC_KEYS_DEPR, DEC_LABELS[gkey]))
    any_row = False
    any_collided = False
    tot_netpl = 0.0
    tot_gain = tot_nbi = tot_nbe = 0.0
    dec_vouch_total = defaultdict(float)
    for ent, rec in recs:
        gl = rec['gl']
        events = {}
        cost_rows_by_vc = {}
        for row in gl:
            ym = _ym(row['date'])
            if ym is None or ym[0] != int(year):
                continue
            g, rl = _grp_cat(row['name'])
            if g != gkey:
                # 处置损益科目（资产处置收益/营业外收支）仍须累计，即便不属于本组资产账户
                if any(ga in row['name'] for ga in GAIN_ACCOUNTS):
                    pass
                else:
                    if not (any(ga in row['name'] for ga in GAIN_ACCOUNTS)
                            or any(nb in row['name'] for nb in NONBIZ_INCOME)
                            or any(nb in row['name'] for nb in NONBIZ_EXPENSE)):
                        continue
            vc = row['voucher'] or 'NOVOUCHER'
            ev = events.setdefault(vc, {'ent': ent, 'date': row['date'], 'sm': [],
                                        'cost_cr': 0.0, 'accum_db': 0.0, 'impair_db': 0.0,
                                        'gain': 0.0, 'nbi': 0.0, 'nbe': 0.0,
                                        'internal': False, 'vt': '', 'no': '', 'method': None})
            nm = row['name']; cp = row['cp'] or ''; sm = row['sm'] or ''
            if g == gkey:
                if rl == 'cost':
                    if row['credit'] > 0:
                        # 长期待摊费用：摊销数记在自身贷方(借:费用科目)，属常规摊销，
                        # 不计入「减少/处置」查证——对方科目为费用类、或摘要【以"摊销"开头】
                        # （FY 01 本级 GL 对方科目列为空，须用摘要兜底；2026-07-31 修复。
                        # 用 startswith 而非包含，避免"…冲销多计摊销"等转固/冲销摘要被误排除）。
                        if gkey == 'LTDEF' and (_is_amort_expense_cp(cp) or (sm or '').strip().startswith('摊销')):
                            pass
                        # 组内原值科目互转(借A贷B)，属重分类，非真实减少
                        elif _is_intragroup_cost_transfer(gkey, cp):
                            pass
                        else:
                            ev['cost_cr'] += row['credit']
                            cost_rows_by_vc.setdefault(vc, []).append((row['date'], row['credit'], sm))
                            if ('内部转移' in sm) or any(k in cp for k in INTERNAL_KW):
                                ev['internal'] = True
                            if sm and sm not in ev['sm']:
                                ev['sm'].append(sm)
                            m = _classify_dec(gkey, cp, sm)
                            if ev['method'] is None:
                                ev['method'] = m
                    if row['debit'] > 0:
                        if ('内部转移' in sm) or any(k in cp for k in INTERNAL_KW):
                            ev['internal'] = True
                elif rl == 'accum':
                    if row['debit'] > 0:
                        ev['accum_db'] += row['debit']
                elif rl == 'impair':
                    if row['debit'] > 0:
                        ev['impair_db'] += row['debit']
            if any(ga in nm for ga in GAIN_ACCOUNTS):
                ev['gain'] += row['credit']
            if any(nb in nm for nb in NONBIZ_INCOME):
                ev['nbi'] += row['credit']
            if any(nb in nm for nb in NONBIZ_EXPENSE):
                ev['nbe'] += row['debit']
            if row['voucher'] and '-' in row['voucher'] and not ev['vt']:
                ev['vt'], ev['no'] = row['voucher'].split('-', 1)

        collided = {vc for vc, lst in cost_rows_by_vc.items() if len(lst) > 1}

        for vc, ev in sorted(events.items()):
            if ev['cost_cr'] <= 0:
                continue
            dec_vouch_total[ent] += ev['cost_cr']
            if vc in collided:
                any_collided = True
                # 字/号碰撞：同一字/号被多笔不同凭证复用，无法可靠匹配累计折旧/减值准备到各原值行。
                # 改为按同字/号合并为一行，原值/累计折旧/减值准备/账面价值均为同号合计，并在备注标注不可靠。
                any_row = True
                net = ev['cost_cr'] - ev['accum_db'] - ev['impair_db']
                pl_acct = ev['gain'] + ev['nbi'] - ev['nbe']
                netpl = round(pl_acct, 2) if abs(pl_acct) > 0.005 else None
                if netpl is not None:
                    tot_netpl += netpl
                tot_gain += ev['gain']; tot_nbi += ev['nbi']; tot_nbe += ev['nbe']
                mlbl = DEC_LBL.get(ev['method'], '其他') if ev['method'] else '其他'
                _txt(ws, r, 1, ent); _txt(ws, r, 2, ev['date']); _txt(ws, r, 3, '')
                _txt(ws, r, 4, mlbl); _txt(ws, r, 5, ev['date'])
                _money(ws, r, 6, ev['cost_cr']); _money(ws, r, 7, ev['accum_db'])
                _money(ws, r, 8, ev['impair_db']); _money(ws, r, 9, net)
                _money(ws, r, 10, None); _money(ws, r, 11, netpl)
                _txt(ws, r, 12, ev['vt']); _txt(ws, r, 13, ev['no'])
                _txt(ws, r, 14, '；'.join(ev['sm'])[:60])
                _money(ws, r, 15, ev['gain']); _money(ws, r, 16, ev['nbi']); _money(ws, r, 17, ev['nbe'])
                _money(ws, r, 18, None)
                _txt(ws, r, 27, '字/号碰撞·合并行' + ('·内部转移' if ev['internal'] else ''))
                r += 1
                continue
            any_row = True
            net = ev['cost_cr'] - ev['accum_db'] - ev['impair_db']
            pl_acct = ev['gain'] + ev['nbi'] - ev['nbe']
            if abs(pl_acct) > 0.005:
                netpl = round(pl_acct, 2)
            else:
                netpl = None
            if netpl is not None:
                tot_netpl += netpl
            tot_gain += ev['gain']; tot_nbi += ev['nbi']; tot_nbe += ev['nbe']
            mlbl = DEC_LBL.get(ev['method'], '其他') if ev['method'] else '其他'
            _txt(ws, r, 1, ent); _txt(ws, r, 2, ev['date']); _txt(ws, r, 3, '')
            _txt(ws, r, 4, mlbl); _txt(ws, r, 5, ev['date'])
            _money(ws, r, 6, ev['cost_cr']); _money(ws, r, 7, ev['accum_db'])
            _money(ws, r, 8, ev['impair_db']); _money(ws, r, 9, net)
            _money(ws, r, 10, None); _money(ws, r, 11, netpl)
            _txt(ws, r, 12, ev['vt']); _txt(ws, r, 13, ev['no'])
            _txt(ws, r, 14, '；'.join(ev['sm'])[:60])
            _money(ws, r, 15, ev['gain']); _money(ws, r, 16, ev['nbi']); _money(ws, r, 17, ev['nbe'])
            _money(ws, r, 18, None)
            _txt(ws, r, 27, '内部转移' if ev['internal'] else '')
            r += 1

    if not any_row:
        _txt(ws, r, 1, '（本年度无%s减少业务）' % sp['title'])
        r += 1
    if any_row:
        r += 1
        c = ws.cell(r, 2, '处置损益提取汇总（同一减少凭证内提取，供核对参考）：净损益合计 %.2f；'
                        '资产处置收益合计 %.2f；营业外收入合计 %.2f；营业外支出合计 %.2f；'
                        '处置损益净额(提取) %.2f。若处置损益在另一凭证确认或处置未通过清理核算，则本表未含，'
                        '与净损益的差异属正常，需结合凭证逐笔追查。'
                        % (round(tot_netpl, 2), round(tot_gain, 2), round(tot_nbi, 2),
                           round(tot_nbe, 2), round(tot_gain + tot_nbi - tot_nbe, 2)))
        c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncols)
        r += 1
    if any_collided:
        r += 1
        c = ws.cell(r, 2, '⚠ 警告：本Client《综合查询明细表》「字/号」非唯一凭证号（同一字/号被多笔不同凭证复用）。'
                        '减少明细按字/号归集时，同一字/号下的累计折旧/减值准备/处置损益无法可靠匹配到各原值行，'
                        '已按同字/号合并为一行并填列原值/累计折旧/减值准备/账面价值（均为同号合计），配对科目标注「不可靠」。'
                        '如需完整逐笔处置明细，须取得唯一凭证号或逐笔追查凭证。')
        c.font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncols)
        r += 1
    r += 1
    c = ws.cell(r, 2, '填表说明：')
    c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncols)
    r += 1
    c = ws.cell(r, 2, '核对内容：1.与收款单据是否一致；2.与合同是否一致；3.审批手续是否完整；4.会计处理是否正确；5.……')
    c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncols)
    r += 1

    # ===== 抽凭合计与账面减少核对 =====
    # 长期待摊费用：若减少数全部为摊销（无真实处置/转出），依据审计要求不生成凭证抽查表
    if gkey != 'LTDEF' or any_row:
        r += 1
        r += 1
        _txt(ws, r, 1, '抽凭合计与账面减少核对', bold=True)
        r += 1
        rh_cols = ['主体', '账面减少（TB贷发）', '抽凭减少合计（GL）', '差额']
        for j, h in enumerate(rh_cols, 1):
            cc = ws.cell(r, j, h)
            cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
        r += 1
        tc_tb = tc_vc = 0.0
        for ent, rec in recs:
            tb = rec['cost_d']['credit'] if rec['cost_d'] else 0.0
            vc = dec_vouch_total.get(ent, 0.0)
            diff = round(vc - tb, 2)
            _txt(ws, r, 1, ent)
            _money(ws, r, 2, tb)
            _money(ws, r, 3, vc)
            _txt(ws, r, 4, '相符' if abs(diff) <= 0.005 else '%.2f（账面负数=红字冲回〔借红字〕；GL 抽凭仅列正数借方行，差额属分类/科目识别差异，非账务差错）' % diff)
            tc_tb += tb; tc_vc += vc
            r += 1
        _txt(ws, r, 1, '合计', bold=True)
        _money(ws, r, 2, tc_tb); _money(ws, r, 3, tc_vc)
        _txt(ws, r, 4, '%.2f' % round(tc_vc - tc_tb, 2))
        for i in range(1, 5):
            ws.cell(r, i).fill = SHELL_TOT_FILL; ws.cell(r, i).font = SHELL_BOLD
        r += 1
        r += 1
        note_ltdef = '（长期待摊费用摊销不计入减少检查表，差额含摊销部分属正常）' if gkey == 'LTDEF' else ''
        nc = ws.cell(r, 1, '差额≠0 系内部转移/摊销等分类差异（GL 为全量抽取），非账务差错。%s' % note_ltdef)
        nc.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        r += 1
    else:
        # 长期待摊费用且减少数全部为摊销：跳过凭证抽查表，仅作说明
        c = ws.cell(r, 2, '说明：本会计期间长期待摊费用减少数全部为摊销（常规摊销完毕/转出），'
                          '无真实处置/转出业务，依据审计要求未生成凭证抽查表。')
        c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncols)
        r += 1

    widths = [14, 12, 12, 12, 12, 14, 12, 12, 12, 12, 12, 10, 10, 28, 12, 10, 10, 10,
              4, 4, 4, 4, 4, 4, 4, 4, 14]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (sr + 1)
    return ws


# ----------------------------------------------------------------------------
# Sheet: 在建工程减少去向统计（特殊）
# ----------------------------------------------------------------------------
def build_audit_confirm_sheet(wb, year, recs, gkey):
    """{资产} 审定表：分核算主体、分 原值/累计折旧(摊销)/减值准备/净值 列示
    期初/期末未审/审计调整/期末审定；末尾全集团合计。数据取自《科目余额表》(权威控制数)。"""
    sp = GROUPS[gkey]
    title = '%s 审定表' % sp['title']
    ws = wb.create_sheet(title, wb.sheetnames.index(sp['main_sheet']))
    is_cip = sp.get('special')
    # 铁律：科目名以实际账套为准——动态发现该组实际备抵科目名（如"使用权资产折旧"）
    accum_label = _discover_accum_name(recs, gkey) if not is_cip else None
    # ⚡⚡ 2026-08-30 ONCA 无累计折旧/摊销（accum_names 空）→ accum_label None →
    #   只列『原值』列（同 is_cip 但走通用 disposal 逻辑，非 CIP 转固）。
    comps = ['原值'] if (is_cip or not accum_label) else ['原值', accum_label, '减值准备', '净值']
    ncols = 6
    r = _title(ws, '%s（%s 年度）' % (title, year), ncols)
    r = _sub(ws, '期末审定数=期末未审数+审计调整数（审计调整数留空待填）；表末全集团合计。', ncols, r)
    hr = _hdr(ws, r, ['主体', '项目', '期初数', '期末未审数', '审计调整数', '期末审定数'])
    r = hr
    COL = {'ent': 1, 'comp': 2, 'o': 3, 'c': 4, 'adj': 5, 'aud': 6}

    def _comp_vals(rec):
        cost = rec['cost_d']
        co = _disp(cost, 'cost', 'opening') if cost else 0.0
        cc = _disp(cost, 'cost', 'closing') if cost else 0.0
        if is_cip:
            return {'原值': (co, cc)}
        accum = rec['accum_d']
        ao = abs(_disp(accum, 'accum', 'opening')) if accum else 0.0
        ac = abs(_disp(accum, 'accum', 'closing')) if accum else 0.0
        imp = rec['impair_d']
        io = abs(_disp(imp, 'impair', 'opening')) if imp else 0.0
        ic = abs(_disp(imp, 'impair', 'closing')) if imp else 0.0
        return {'原值': (co, cc), accum_label: (ao, ac),
                '减值准备': (io, ic), '净值': (co - ao - io, cc - ac - ic)}

    ent_rows = {}
    for ent, rec in recs:
        vals = _comp_vals(rec)
        ent_rows[ent] = {}
        for comp in comps:
            o, c = vals.get(comp, (0.0, 0.0))
            rr = r + 1
            _txt(ws, rr, COL['ent'], ent)
            _txt(ws, rr, COL['comp'], comp, bold=(comp == '净值'))
            _money(ws, rr, COL['o'], o)
            _money(ws, rr, COL['c'], c)
            _money(ws, rr, COL['adj'], None)  # 审计调整数留空待填
            if comp == '净值':
                # 2026-08-06 协议：净值审定=净值期末(c)+调整(留空=0)，写数值（公式 data_only 读 None）
                _money(ws, rr, COL['aud'], round(c, 2))
            else:
                _money(ws, rr, COL['aud'], round(c, 2))
            ent_rows[ent][comp] = rr
            if comp == '净值':
                for i in range(1, ncols + 1):
                    ws.cell(rr, i).fill = SHELL_TOT_FILL
            r = rr
        r += 1

    # 全集团合计（⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出）
    if len(recs) > 1:
        r += 1
        c0 = ws.cell(r, 1, '全集团合计')
        c0.font = SHELL_BOLD; c0.fill = SHELL_TOT_FILL
        for comp in comps:
            erows = [ent_rows[e][comp] for e in ent_rows if comp in ent_rows[e]]
            rr = r + 1
            _txt(ws, rr, COL['ent'], '全集团合计', bold=True)
            _txt(ws, rr, COL['comp'], comp, bold=True)
            # 2026-08-01 修复：原 SUM(first:last) 范围公式——因 原值/累计折旧(摊销)/减值/净值 行交错，
            # 范围会混入其他类行致合计双计（如长期待摊原值合计=2×原值）。改为离散行求和。
            if erows:
                # 2026-08-06 协议：全集团合计写数值（读数据行已写值求和；公式 data_only 读 None）
                for _cc in ('o', 'c', 'adj', 'aud'):
                    _sv = 0.0
                    for _e in erows:
                        _cv = ws.cell(_e, COL[_cc]).value
                        if isinstance(_cv, (int, float)):
                            _sv += _cv
                    _money(ws, rr, COL[_cc], round(_sv, 2))
            else:
                for _cc in ('o', 'c', 'adj', 'aud'):
                    _money(ws, rr, COL[_cc], 0.0)
            for i in range(1, ncols + 1):
                ws.cell(rr, i).fill = SHELL_TOT_FILL
            if comp == '净值':
                for i in range(1, ncols + 1):
                    ws.cell(rr, i).font = SHELL_BOLD
            r = rr
        r += 1

    widths = [14, 16, 16, 16, 14, 16]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    return ws


def build_cip_disposal_sheet(wb, year, recs):
    sp = GROUPS['CIP']
    ws = wb.create_sheet(sp['dec_sheet'])
    # 2026-07-31 用户要求：在建工程减少检查表改用普通凭证抽查模板（标准 15 列）；
    # 表尾保留「抽凭合计与账面减少核对」（账面=TB 贷发）。
    rows = []
    cip_vouch_total = defaultdict(float)
    for ent, rec in recs:
        gl = rec['gl']
        for row in gl:
            ym = _ym(row['date'])
            if ym is None or ym[0] != int(year):
                continue
            g, rl = _grp_cat(row['name'])
            if g != 'CIP' or rl != 'cost' or row['credit'] <= 0:
                continue
            if _is_intragroup_cost_transfer('CIP', row['cp']):
                continue  # 在建工程组内互转(借A贷B)，属重分类，非真实减少
            vt = no = ''
            if row['voucher'] and '-' in row['voucher']:
                vt, no = row['voucher'].split('-', 1)
            rows.append({'ent': ent, 'date': row['date'], 'vt': vt, 'no': no,
                         'name': row['name'], 'sm': row['sm'], 'cp': row['cp'],
                         'amt': row['credit']})
            cip_vouch_total[ent] += row['credit']
    r = _write_std_vouch_sheet(ws, sp, year, rows, 'dec')

    # ===== 抽凭合计与账面减少核对 =====
    r += 1
    r += 1
    _txt(ws, r, 1, '抽凭合计与账面减少核对', bold=True)
    r += 1
    rh_cols = ['主体', '账面减少（TB贷发）', '抽凭减少合计（GL）', '差额']
    for j, h in enumerate(rh_cols, 1):
        cc = ws.cell(r, j, h)
        cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    r += 1
    tc_tb = tc_vc = 0.0
    for ent, rec in recs:
        tb = rec['cost_d']['credit'] if rec['cost_d'] else 0.0
        vc = cip_vouch_total.get(ent, 0.0)
        diff = round(vc - tb, 2)
        _txt(ws, r, 1, ent)
        _money(ws, r, 2, tb)
        _money(ws, r, 3, vc)
        _txt(ws, r, 4, '相符' if abs(diff) <= 0.005 else '%.2f（账面负数=红字冲回〔借红字〕；GL 抽凭仅列正数借方行，差额属分类/科目识别差异，非账务差错）' % diff)
        tc_tb += tb; tc_vc += vc
        r += 1
    _txt(ws, r, 1, '合计', bold=True)
    _money(ws, r, 2, tc_tb); _money(ws, r, 3, tc_vc)
    _txt(ws, r, 4, '%.2f' % round(tc_vc - tc_tb, 2))
    for i in range(1, 5):
        ws.cell(r, i).fill = SHELL_TOT_FILL; ws.cell(r, i).font = SHELL_BOLD
    r += 1
    r += 1
    nc = ws.cell(r, 1, '差额≠0 系分类/科目识别差异（GL 为全量抽取）；账面负数=红字冲回（借红字冲减），抽凭仅列正数借方行，非账务差错。')
    nc.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=15)
    r += 1
    return ws


# ----------------------------------------------------------------------------
# Sheet: 目录说明
# ----------------------------------------------------------------------------
def build_cover(wb, data_dir, year, ents, recs, gkey, notes):
    sp = GROUPS[gkey]
    ws = wb.create_sheet('目录说明', 0)
    ws.cell(1, 1, '%s审计底稿（%s 年度）' % (sp['title'], year)).font = SHELL_TITLE_FONT
    lines = [
        ('数据目录', data_dir),
        ('年度', year),
        ('主体数', '%d 个：%s' % (len(ents), '、'.join(ents))),
        ('数据来源', '《科目余额表》（控制数）+《综合查询明细表(GL)》（凭证级）'),
        ('科目范围', '%s：%s（按科目名自动识别，兼容不同科目表）' % (
            sp['title'],
            '、'.join([sp['title']] + sp.get('accum_names', []) + sp.get('impair_names', [])
                       + sp.get('clear_names', [])) or sp['title'])),
    ]
    r = 3
    for k, v in lines:
        _txt(ws, r, 1, k, bold=True)
        ws.cell(r, 2, v).alignment = SHELL_LEFT
        r += 1
    r += 1
    _txt(ws, r, 1, '工作表索引', bold=True)
    r += 1
    idx = [
        (sp['main_sheet'], '主表：一级(合计)与二级明细的期初/增加/减少/期末均取自科目余额表；'
         + ('表末"全集团按二级分类汇总"按二级分类列示原值/累计折旧(摊销)/减值准备/净值 各自的期初/本期增加/本期减少/期末/审计调整/审定数。' if not sp.get('special')
            else '仅滚调表格式(期初/本期增加/本期减少/期末/审计调整/审定数),无净值与折旧摊销；表末"全集团按二级分类汇总"按二级分类列示原值(期末)。')),
    ]
    if sp.get('amort'):
        idx.append((sp['amort_sheet'], '%s：计提数=当期%s贷方发生数,按一级费用科目展开的费用列支数与之勾稽相等；'
                    '研发支出与研发费用结转互转不重复统计；每个主体小计后附"上年计提小计/本期小计-上期小计"两行' % (
                        sp['amort_sheet'], '折旧' if sp['accum_kw'] == '折旧' else '摊销')))
    idx.append((sp['inc_sheet'], '全集团各主体%s当年增加(借方>0)的凭证级抽样检查,含「在建工程转入」分类及汇总'
                % sp['title']))
    if sp.get('special'):
        idx.append((sp['dec_sheet'], '在建工程减少检查表：转入固定资产/无形资产/长期待摊费用/投资性房地产/内部转移/出售/其他,附汇总及抽凭合计与账面减少核对'))
    else:
        idx.append((sp['dec_sheet'], '每笔减少(按凭证)原值/折旧/减值/账面价值/处置收入/净损益,与处置损益核对；表末附抽凭减少合计与账面减少(TB贷发)核对'))
    for name, desc in idx:
        _txt(ws, r, 1, name, bold=True)
        ws.cell(r, 2, desc).alignment = SHELL_LEFT
        r += 1
    r += 1
    _txt(ws, r, 1, '重要提示', bold=True)
    r += 1
    for n in notes:
        if not n:
            continue
        c = ws.cell(r, 1, '· ' + n)
        c.alignment = SHELL_LEFT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        r += 1
    for i, w in enumerate([26, 80, 10, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
def process_folder(data_dir, out_path=None):
    if not os.path.isdir(data_dir):
        print(f'[ERROR] 数据目录不存在：{data_dir}')
        return 1
    ents = _discover_entities(data_dir)
    if not ents:
        print(f'[ERROR] 未在 {data_dir} 发现科目余额表/综合查询明细表')
        return 1
    print(f'[INFO] 发现主体 {len(ents)} 个：{mask_names(sorted(ents))}')

    years = set()
    for ent in ents:
        years |= set(ents[ent]['km'].keys())
        years |= set(ents[ent]['gl'].keys())
    years = sorted(years)
    if not years:
        print('[ERROR] 未发现任何年份的账套数据')
        return 1
    print(f'[INFO] 检测到年份：{years}')

    rc = 0
    for year in years:
        for gkey in GROUPS:
            rc_year = _build_year_workbook(data_dir, ents, year, gkey)
            if rc_year != 0:
                rc = rc_year
    return rc


def _build_year_workbook(data_dir, ents, year, gkey):
    sp = GROUPS[gkey]
    recs = []
    # ⚡⚡ 2026-08-30 修复：集团模式 current_comp=None → read_km 把区间文件(2310~2400.xlsx)
    #   解析为区间起点主体 2310 → 2330-2400 固定资产原值全部串号（710.9M）。
    #   逐主体显式 set_comp(ent)（同 revenue #740 修复），函数结束恢复。
    try:
        import sap_adapter as _A
        _orig_comp = _A.current_comp()
    except Exception:
        _orig_comp = None
    for ent in sorted(ents):
        km_path = ents[ent]['km'].get(year)
        if not km_path:
            continue
        try:
            import sap_adapter as _A2
            if hasattr(_A2, 'set_comp'):
                _A2.set_comp(ent)
        except Exception:
            pass
        # ⚡ 2026-08-13 原则（用户定）：生成器只负责生成底稿，数据问题全部在前端解决。
        #   _read_km 在前端（SAP 下被 patch 为 adapter.read_km）内部按 current_comp
        #   解析主体，生成器不感知数据形态、不分支（SAP 区间文件主体解析在 adapter 层）。
        km = _read_km(km_path)
        cost, accum, impair = _find_group_codes(km, gkey)
        if not cost and not accum:
            continue
        gl_path = ents[ent]['gl'].get(year)
        gl = _read_gl_fa(gl_path) if gl_path else []
        recs.append((ent, _prep_group(ent, km, gl, gkey)))
    if not recs:
        # 2026-08-06：空产出告警——区分『账套真无该组科目』与『识别失败』。
        # jtt 曾因 read_km 表头不兼容整表 0 科目导致 6 组资产底稿静默缺失（无任何提示）。
        _probe = False
        _gname = sp.get('title') or ''
        for _ent in sorted(ents):
            _km_path = ents[_ent]['km'].get(year)
            if not _km_path:
                continue
            try:
                _km0 = _read_km(_km_path)
            except Exception:
                _km0 = {}
            if _km0 and any(_gname and _gname in (d.get('name') or '') for d in _km0.values()):
                _probe = True
                break
        if _probe:
            print(f'  ⚠️ [WARN] {_gname} 组：TB 存在含「{_gname}」的科目但全部主体识别失败'
                  f'（cost/accum 为空）—— 请检查 read_km 表头兼容与科目名判定')
        else:
            print(f'  [SKIP] {_gname} 组：本账套 TB 无该组科目（无数据，不产生空底稿）')
        # 该组本年无数据，跳过（不产生空底稿）
        try:
            import sap_adapter as _A3
            if hasattr(_A3, 'set_comp'):
                _A3.set_comp(_orig_comp)
        except Exception:
            pass
        return 0

    out_path = os.path.join(data_dir, '%s_%s_生成.xlsx' % (sp['paper'], year))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    notes = ['GL 为全量抽取（在抽取期间内无缺失行）；差异系分类/科目识别差异，已在各表归因',
             '2026 年 GL 仅 1–5 月（YTD）' if '2026' in str(year) else '',
             '各底稿《增加检查表》列示"在建工程转入"金额；在建工程底稿《在建工程减少检查表》列示其去向']
    # 目录说明已删除（2026-07-29 清理）
    g_cat = build_rollforward_sheet(wb, year, recs, gkey)
    print(f'[OK] {sp["main_sheet"]} ({year})')
    # 附注披露（2026-08-02 用户方法论：所有科目最后加 1 张附注汇总表；行次多/少自动判格式）
    try:
        build_footnote_sheet(wb, year, recs, gkey, sp, all_ents=sorted(ents))
        print(f'[OK] 附注披露 ({year})')
    except Exception as _ex:
        print(f"  ⚠️ 附注披露生成失败：{_ex}")
    # 全集团按二级分类汇总 已直接接在分类汇总表末尾（2026-07-29），不再独立 sheet
    if sp.get('amort'):
        build_amort_sheet(wb, year, recs, ents, gkey)
        print(f'[OK] {sp["amort_sheet"]} ({year})')
        # ⚡⚡ 2026-08-28 费用侧对应科目罗列（用户方法论）：在 FA 底稿附一张全资产组合计的
        #   《费用中摊销折旧计提汇总》，把同时记入费用的各类折旧/摊销/租赁费罗列汇总。
        if gkey == 'FA':
            build_fee_dep_summary_sheet(wb, year, recs)
            print(f'[OK] 费用中摊销折旧计提汇总 ({year})')
    build_inc_sheet(wb, year, recs, gkey)
    print(f'[OK] {sp["inc_sheet"]} ({year})')
    if sp.get('special'):
        build_cip_disposal_sheet(wb, year, recs)
        print(f'[OK] {sp["dec_sheet"]} ({year})')
    else:
        build_disposal_sheet(wb, year, recs, gkey)
        print(f'[OK] {sp["dec_sheet"]} ({year})')

    # ---- 前置插入『{资产} 审定表』（分账套×组件×期初/期末/调整/审定，全集团合计） ----
    try:
        build_audit_confirm_sheet(wb, year, recs, gkey)
        print(f'[OK] {sp["title"]} 审定表 ({year})')
    except Exception as _ex:
        print(f"  ⚠️ {sp['title']}审定表生成失败：{_ex}")
    # ⚡ 2026-08-11 用户需求：{资产} 横展审定表（原值/折旧/减值/净值 × 各主体）
    try:
        build_longterm_wide_sheet(wb, year, recs, gkey)
    except Exception as _ex:
        print(f"  ⚠️ {sp['title']}横展审定表生成失败：{_ex}")

    # 2026-08-03 备抵机制通用化：该资产组若存在减值准备科目（有余额）→ 追加『XX减值准备明细表』
    # （期初/计提/转回及转销/期末；FY 无固定资产/无形/在建减值数据时自动跳过）
    try:
        from baddebt_common import build_baddebt_detail_sheet as _bd_sheet
        _imp_name = (sp.get('impair_names') or [''])
        _tbv = {}
        for _e, _r in recs:
            _imp = _r.get('impair_d')
            if not _imp:
                continue
            _tbv[(_e, _r.get('impair') or '', _imp_name[0], str(year))] = {
                'qc': _imp.get('opening', 0.0), 'jf': _imp.get('debit', 0.0),
                'df': _imp.get('credit', 0.0), 'qm': _imp.get('closing', 0.0)}
        if _tbv:
            _bd_sheet(wb, _tbv, None, year, _imp_name[0], sp['title'])
            print(f'[OK] {sp["title"]}减值准备明细表 ({year})')
    except Exception as _ex:
        print(f"  ⚠️ {sp['title']}减值准备明细表生成失败：{_ex}")

    finalize_workbook(wb)
    # ⚡ 2026-08-28 #875：对方科目核对（模块内集成）
    try:
        from counterparty_recon import inject_into_wb_auto
        inject_into_wb_auto(wb, data_dir, year, sp['title'], ents_set=set(ents))
    except Exception as _ex:
        print(f"  ⚠️ {sp['title']}对方科目核对注入失败：{_ex}")
    out_path, _save_warn = _safe_save(wb, out_path)
    if _save_warn:
        print(f'  [WARN] {_save_warn}')
    if out_path is None:
        raise PermissionError(_save_warn or '保存失败：目标与副本均无法写入')
    wb.close()
    print(f'[DONE] 已生成：{out_path}')
    try:
        import sap_adapter as _A4
        if hasattr(_A4, 'set_comp'):
            _A4.set_comp(_orig_comp)
    except Exception:
        pass
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description='长期资产科目审计底稿生成程序（固定资产小程序升级版）')
    ap.add_argument('folders', nargs='*', default=[],
                    help='一个或多个数据目录（拖入文件夹时自动作为位置参数传入）')
    ap.add_argument('--input', action='append', dest='inputs', default=[],
                    help='数据目录（可多次指定）')
    ap.add_argument('--output', default=None, help='仅单目录时指定输出路径')
    args = ap.parse_args(argv)

    dirs = list(args.folders) + list(args.inputs)
    if not dirs:
        dirs = [P.T]

    rc = 0
    for d in dirs:
        out = args.output if len(dirs) == 1 else None
        r = process_folder(d, out)
        if r != 0:
            rc = r
        from audit_common import finalize_after_build
        finalize_after_build(d)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）
    return rc



# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None

if __name__ == '__main__':
    _bs_sys.exit(main())

def build_longterm_wide_sheet(wb, year, recs, gkey):
    """⚡ 2026-08-11 用户需求：{资产} 横展审定表——行=组件(原值/累计折旧(摊销)/减值准备/净值)、
    列=各核算主体期末+合计（横向对比各单位资产规模/折旧水平）。集团模式（多主体）生成；
    在建工程(special)仅原值一行。数据与审定表同源（科目余额表权威控制数）。"""
    from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_BOLD,
                             SHELL_NUM, SHELL_LEFT, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                             SHELL_TOT_FILL)
    from openpyxl.utils import get_column_letter
    sp = GROUPS[gkey]
    if len(recs) <= 1:
        return
    is_cip = sp.get('special')
    accum_label = _discover_accum_name(recs, gkey) if not is_cip else None
    comps = ['原值'] if is_cip else ['原值', accum_label, '减值准备', '净值']

    def _comp_vals(rec):
        cost = rec['cost_d']
        cc = _disp(cost, 'cost', 'closing') if cost else 0.0
        if is_cip:
            return {'原值': cc}
        accum = rec['accum_d']
        ac = abs(_disp(accum, 'accum', 'closing')) if accum else 0.0
        imp = rec['impair_d']
        ic = abs(_disp(imp, 'impair', 'closing')) if imp else 0.0
        return {'原值': cc, accum_label: ac, '减值准备': ic, '净值': cc - ac - ic}

    ents = sorted(e for e, _r in recs)
    ent_vals = {e: _comp_vals(r) for e, r in recs if r}
    if not ent_vals:
        return
    ws = wb.create_sheet(f'{sp["title"]}横展审定表_{year}')
    nc = 1 + len(ents) + 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, f'{sp["title"]}横展审定表（{year} 年度）· 行=原值/累计折旧/减值/净值、列=各核算主体期末+合计（横向对比各单位资产规模）')
    c.font = SHELL_TITLE_FONT
    hdr = ['项目'] + ents + ['合计']
    for j, h in enumerate(hdr, 1):
        cc = ws.cell(2, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    r = 3
    for comp in comps:
        is_net = (comp == '净值')
        vals_by_e = {e: (ent_vals[e].get(comp, 0.0) if e in ent_vals else None) for e in ents}
        tot = sum(v for v in vals_by_e.values() if isinstance(v, (int, float)))
        ws.cell(r, 1, comp)
        for j, e in enumerate(ents, 2):
            v = vals_by_e[e]
            cc = ws.cell(r, j, round(v, 2) if isinstance(v, (int, float)) and abs(v) >= 0.005 else None)
            cc.border = SHELL_BORDER
            cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
        cc = ws.cell(r, nc, round(tot, 2) if abs(tot) >= 0.005 else None)
        cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
        cc.font = SHELL_BOLD
        for j in range(1, nc + 1):
            ws.cell(r, j).border = SHELL_BORDER
            if is_net:
                ws.cell(r, j).fill = SHELL_TOT_FILL
                ws.cell(r, j).font = SHELL_BOLD
        r += 1
    ws.column_dimensions['A'].width = 22
    for j in range(2, nc + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.freeze_panes = 'B3'
    return ws
