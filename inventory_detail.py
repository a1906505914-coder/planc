# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=存货审计底稿_生成.xlsx | 关键列=存货二级期初/增减/期末/跌价准备 | 职责=存货变动与成本勾稽(独立生成器)

from audit_common import _safe_save, is_zero_amount, read_km as _read_km, read_gl as _read_gl  # 共享库：统一锁感知保存（单一来源）+ 0值判断 + 通用读取解析器
"""
存货审计底稿生成程序（第 7 个小程序）
================================================================================
功能概述
--------------------------------------------------------------------------------
读取 G 各核算主体《科目余额表》+《综合查询明细表(GL)》，自动生成：
  1. 存货审定表（2 级科目：期初/本期增加/本期减少/期末未审/审计调整[留空]/审定数/变动比例，含 1471 跌价准备）；
  2. 原材料 分月变动表 + 变动分析表（最末级明细 期末 vs 期初）；
  3. 库存商品 分月变动表 + 变动分析表；
  4. 生产成本 分月变动表 + 成本分析表（每月各项增加占比，识别异常）；
  5. 其他存货（发出商品/半成品/周转材料/跌价准备）分月变动表 + 变动分析表；
  6. 编制说明（数据来源、口径、非全量抽取披露）。

约定
--------------------------------------------------------------------------------
  · 科目余额表 = 权威控制数（期初/借发/贷发/期末）。
  · 综合查询明细表(GL) = 方向性归集（按对方科目区分增加/减少来源）。
  · 2025 GL 全量(1–12月)；2026 GL 仅 1–5 月(YTD)，差额=凭证抽取不完整，非账务错误。
  · 外币科目余额表不纳入，以标准本位币科目余额表为准。
  · 审计调整数留空，审定数 = 期末未审数 + 审计调整数（公式，审计人员填列后自动更新）。

使用
  python inventory_detail.py [--input 数据目录] [--output 输出xlsx]
  默认 input = 桌面 G；output = <input>/存货审计底稿_生成.xlsx
================================================================================
"""

import os as _bs_os
import sys as _bs_sys
import subprocess as _bs_sub


# ---- 启动诊断日志（绝对路径，纯内置库，任何崩溃前必写，用于定位"闪退无痕迹"）----
# 与费用(费用明细)小程序同款机制：模块一加载就写 'LOAD'，先于 import openpyxl / relaunch，
# 因此即便后续因系统 Python 无 openpyxl 或 relaunch 失败而闪退，也会在
# ~/Desktop/inventory_boot.log 留下启动痕迹，便于事后排查。
def _write_boot_log(tag):
    import os as _bl_os
    if _bl_os.environ.get('AUDIT_LOG') != '1':
        return
    try:
        import datetime
        _p = _bs_os.path.join(_bs_os.path.expanduser('~'), 'Desktop', 'inventory_boot.log')
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
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import re

from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_SUB_FONT,
                         SHELL_BOLD, SHELL_NUM, SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT,
                         SHELL_TOT_FILL, finalize_workbook)

# ----------------------------------------------------------------------------
# 存货科目范围（2 级父 + 末级）
# ----------------------------------------------------------------------------
# 2 级父科目
INV_LEVEL2 = {
    '1403': '原材料',
    '1405': '库存商品',
    '1406': '发出商品',
    '1408': '委托加工物资',   # ⚡⚡ 2026-08-24 修复：AFJ U8 1408 委托加工物资 130万
    #   漏配 → 存货审定表缺该科目 → 阶段③核对差 1,301,354.12（SAP 1411 委托加工走名称匹配）
    '1411': '周转材料',
    '1412': '半成品',
    '1471': '存货跌价准备',
}
# ⚡ 2026-08-10 SAP 存货按【实际科目名称】分类（用户要求；铁律58 名称主键）：
# SAP 1406=库存产品父级（子目=库存商品/产成品/半成品/在制品）≠U8 发出商品；
# SAP 1411=委托加工物资 ≠U8 周转材料；另有 1404 材料成本差异/1407 库存产品成本差异。
# 名称匹配顺序：跌价准备最先（防『存货跌价准备-库存产品』误入库存商品），
# 成本差异其次（防『库存产品成本差异-产成品』误入产成品），其余按关键词。
SAP_INV_CATS = [
    ('存货跌价准备', ['跌价准备']),
    ('材料成本差异', ['材料成本差异', '成本差异']),
    ('原材料', ['原材料']),
    ('库存商品', ['库存商品']),
    ('产成品', ['产成品']),
    ('半成品', ['半成品']),
    ('在制品', ['在制品', '在产品']),
    ('发出商品', ['发出商品']),
    ('委托加工物资', ['委托加工物资']),
    ('周转材料', ['周转材料', '包装物', '低值易耗品']),
    ('在途物资', ['在途']),
]


def _sap_inv_cat(nm):
    """SAP 存货科目名 → 类别名（无匹配返回 None）。"""
    n = str(nm or '')
    for cat, kws in SAP_INV_CATS:
        for kw in kws:
            if kw in n:
                return cat
    return None
# 2026-08-06 泰国存货代码映射：泰国 1140 系列（1140.01 原材料/1140.02 成品/1140.04 发出商品）
# → 标准类别码（审定表/附注汇总取数用；泰国代码未归一化，标准码精确匹配会落空 → 存货空壳）
THAI_INV_MAP = {
    '1403': '1140.01',   # 原材料
    '1405': '1140.02',   # 库存商品（泰国 成品）
    '1406': '1140.04',   # 发出商品
}
# 1411 周转材料 的末级子目
INV_CHILDREN = {
    '1411': [('141101', '周转材料-包装材料'), ('141102', '周转材料-低值易耗品')],
}
# 末级科目（用于变动分析表）：2级父 + 其末级子目集合
LEAF_ACCOUNTS = {}  # code -> name
for c, n in INV_LEVEL2.items():
    LEAF_ACCOUNTS[c] = n
    for cc, cn in INV_CHILDREN.get(c, []):
        LEAF_ACCOUNTS[cc] = cn

# 生产成本
PROD_CODE = '5001'
PROD_NAME = '生产成本'

# ----------------------------------------------------------------------------
# 双轨取数键：交易对手科目(对方科目)优先；为空时回退 GL 末级科目名称
# ----------------------------------------------------------------------------
def _has_kw(parts, kws):
    return any(kw in p for p in parts for kw in (kws if isinstance(kws, (list, tuple)) else [kws]))


def _std_inv_cat(cp, side):
    """标准存货(原材料/库存商品/半成品/发出商品/周转材料) 增加(借)/减少(贷) 按『对方科目』语义归类。

    本项目《综合查询明细表》的『对方科目』列实测 29 个账套 35520 行仅 47 行空（<0.2%），
    故直接按对方科目性质归入业务类别，不再回退到科目自身（原回退会导致『原材料』既作借又
    作贷、毫无审计意义）。用户要求：借=购入对应科目（应付账款购入/银行存款购入/生产成本转入等）、
    贷=转出对应科目（转入生产成本/制造费用/营业费用/管理费用/研发费用、结转主营业务成本等），
    无发生额则不列示。对方科目空值归入『其他增加/其他减少』并单列披露。
    """
    parts = [p.strip() for p in (cp or '').split(',') if p and p.strip()]
    if not parts:
        return '其他增加' if side == 'inc' else '其他减少'
    if side == 'inc':
        if _has_kw(parts, ['应付账款', '应付票据', '暂估', '材料采购']):
            return '应付账款购入'
        if _has_kw(parts, ['预付账款']):
            return '预付账款购入'
        if _has_kw(parts, ['银行存款', '库存现金', '其他货币资金']):
            return '银行存款购入'
        if _has_kw(parts, ['生产成本']):
            return '生产成本转入'
        # 2026-08-03 加强识别：存货间互转/委托加工/损益结转 不再落「其他增加」
        if _has_kw(parts, ['库存商品']):
            return '库存商品转入'
        if _has_kw(parts, ['发出商品']):
            return '发出商品转入'
        if _has_kw(parts, ['委托加工物资', '委托调拨物资']):
            return '委托加工'
        if _has_kw(parts, ['本年利润', '未分配利润', '利润分配', '以前年度损益调整']):
            return '损益结转'
        # 一借多贷采购：借库存商品、贷应付账款+应交税费(进项税) —— 进项税随采购计入存货，归购入
        if _has_kw(parts, ['应交税费']):
            return '购入(含进项税)'
        # 2026-07-29 用户要求：周转材料挂其他应付款（零星报销/福利领用）→ 归为"其他应收款/费用领用"
        if _has_kw(parts, ['其他应付款', '其他应收款', '报销', '福利费']):
            return '其他应收款/费用领用'
        return '其他增加'
    # side == 'dec'：存货减少/转出对应科目
    if _has_kw(parts, ['生产成本']):
        return '转入生产成本'
    if _has_kw(parts, ['制造费用']):
        return '转入制造费用'
    if _has_kw(parts, ['销售费用', '营业费用']):
        return '转入营业费用'
    if _has_kw(parts, ['管理费用']):
        return '转入管理费用'
    if _has_kw(parts, ['研发费用']):
        return '转入研发费用'
    if _has_kw(parts, ['主营业务成本']):
        return '结转主营业务成本'
    if _has_kw(parts, ['其他业务成本']):
        return '结转其他业务成本'
    if _has_kw(parts, ['库存商品']):
        return '转入库存商品'
    if _has_kw(parts, ['半成品', '自制半成品', '中间品']):
        return '转入自制半成品'
    if _has_kw(parts, ['发出商品']):
        return '转入发出商品'
    if _has_kw(parts, ['原材料']):
        return '转入原材料'
    # 2026-08-03 加强识别：委托加工/损益结转 不再落「其他减少」
    if _has_kw(parts, ['委托加工物资', '委托调拨物资']):
        return '委托加工'
    if _has_kw(parts, ['本年利润', '未分配利润', '利润分配', '以前年度损益调整']):
        return '损益结转'
    # 存货领用/转为工程资产或固定资产（自建、存货转固）
    if _has_kw(parts, ['工程施工']):
        return '转入工程施工'
    if _has_kw(parts, ['在建工程']):
        return '转入在建工程'
    if _has_kw(parts, ['固定资产']):
        return '转入固定资产'
    # 退供/冲暂估：对方仍挂应付账款（红字或正常），业务上属采购退回
    if _has_kw(parts, ['应付账款', '应付票据', '暂估']):
        return '采购退回'
    if _has_kw(parts, ['待处理财产损益', '待处理财产损溢']):
        return '待处理财产损益'
    # 2026-07-29 用户要求：库存商品盘点/报废 对方为累计折旧 → 归待处理财产损益
    if _has_kw(parts, ['累计折旧', '累计摊销']):
        return '待处理财产损益'
    if _has_kw(parts, ['研发支出']):
        return '转入研发支出'
    return '其他减少'


def classify_contra(cp, side):
    """1471 存货跌价准备（备抵）：借=转回(减)，贷=计提(增)。"""
    if side == '贷':
        return '计提(增)'
    return '转回(减)'


def classify_prod_inc(name, cp=None):
    """生产成本 增加(借)按【科目名称】明细分类（对方科目为多值串，不可靠）。

    标准成本三要素 + 委外 + 自制半成品：直接材料 / 直接人工(科目名多见"直接工资"/"工资")
    / 制造费用(含水电气动力等) / 委外加工费 / 自制半成品(分步法下自制半成品转入下步骤)。
    GL 中 生产成本 末级科目名为「生产成本-直接原料」「生产成本-工资」「生产成本-制造费用」
    「生产成本-委外加工费」「生产成本-半成品」等——已对齐关键词。
    2026-07-29：对方科目已知为『库存商品』时（存货回转进生产），优先归『库存商品转入』，
    否则仍按科目末级名分类。
    """
    if cp and _has_kw([p.strip() for p in str(cp).split(',') if p and p.strip()], ['库存商品']):
        return '库存商品转入'
    if '直接材料' in name or '直接原料' in name:
        return '直接材料'
    if '直接工资' in name or '直接人工' in name or '工资' in name:
        return '直接人工'
    if ('制造费用' in name or '水电' in name or '动力' in name or '燃料' in name
            or '能源' in name or '动能' in name or '蒸气' in name or '蒸汽' in name):
        return '制造费用转入'
    if '委外' in name or '外协' in name:
        return '委外加工费'
    if '半成品' in name or '自制半成品' in name or '中间品' in name:
        return '自制半成品'
    return '其他增加'


def classify_prod_dec(cp):
    """生产成本 减少(贷)按对方科目分类；转出须区分『自制半成品』与『库存商品』(用户要求)。

    对方科目为空(个别行)时无法区分，归入『其他减少』并在表头披露。
    """
    parts = [p.strip() for p in (cp or '').split(',') if p and p.strip()]
    if _has_kw(parts, ['库存商品']):
        return '转入库存商品'
    if _has_kw(parts, ['半成品', '自制半成品', '中间品']):
        return '转入自制半成品'
    if _has_kw(parts, ['发出商品']):
        return '转入发出商品'
    if _has_kw(parts, ['原材料', '周转材料']):
        return '转入原材料'
    # 退供/冲暂估：对方挂应付账款
    if _has_kw(parts, ['应付账款', '应付票据', '暂估']):
        return '采购退回'
    return '其他减少'


def _prod_cp_empty(ents):
    """扫描所有实体 GL，判断生产成本贷方(减少)的『对方科目』是否全部不可得（含反推）。

    若全部为空且同号凭证反推也无法恢复，则减少侧无法区分完工入库/其他减少，
    需在分月变动表披露：所有贷方已并入『其他减少』（实为完工入库性质），不得解读为异常。
    2026-07-29 起：对方科目为空时先用 _recon_cp 同号凭证反推，反推可得即视为有对方科目。
    """
    for ent in ents:
        for year in sorted({y for e in ents for y in ents[e]}):
            gl_path = ents[ent]['gl'].get(year)
            if not gl_path:
                continue
            gl = _read_gl_cached(gl_path)
            vidx = _build_vidx(gl)
            for row in gl:
                nm = row.get('name', '')
                if not (nm == PROD_NAME or nm.startswith(PROD_NAME + '-')):
                    continue
                if not row.get('credit'):
                    continue
                cp_eff = (row.get('cp') or '').strip() or _recon_cp(row, vidx)
                if cp_eff:
                    return False
    return True


# 标准存货 增加/减少 语义类别（固定顺序；仅列示有发生额者）
# 2026-08-03 修复：必须与 _std_inv_cat 全部返回值一致——此前缺『购入(含进项税)』『其他应收款/费用领用』
# 等类别，Z 母公司"借原材料/贷应付账款+应交税费(进项税)"一借多贷采购被丢进黑洞，13.2 亿借方缺失
# → 原材料分月变动表增加数异常小、期末滚成负数。贷方同样补齐 _std_inv_cat 全部 dec 返回类别。
STD_INC_CATS = ['应付账款购入', '预付账款购入', '银行存款购入', '生产成本转入', '库存商品转入',
                '发出商品转入', '委托加工', '损益结转',
                '购入(含进项税)', '其他应收款/费用领用', '其他增加']
STD_DEC_CATS = ['转入生产成本', '转入制造费用', '转入营业费用', '转入管理费用', '转入研发费用',
                '结转主营业务成本', '结转其他业务成本', '转入库存商品', '转入自制半成品',
                '转入发出商品', '转入原材料', '委托加工', '损益结转',
                '转入工程施工', '转入在建工程', '转入固定资产',
                '采购退回', '待处理财产损益', '转入研发支出', '其他减少']
# 历史兼容常量（已不再用于动态列）
STD_INC = ['购入', '生产成本转入', '其他增加']
STD_DEC = ['生产领用', '制造费用领用', '管理费用领用', '销售费用领用', '研发费用领用',
           '结转主营业务成本', '结转其他业务成本', '其他减少']
CONTRA_INC = ['计提(增)']
CONTRA_DEC = ['转回(减)']
PROD_INC = ['直接材料', '直接人工', '制造费用转入', '自制半成品', '委外加工费', '库存商品转入', '其他增加']
PROD_DEC = ['转入库存商品', '转入自制半成品', '转入发出商品', '转入原材料', '采购退回', '其他减少']

# ----------------------------------------------------------------------------
# 数据加载
# ----------------------------------------------------------------------------

def _split_ent_year(fn, marker):
    """从文件名解析 (ent, yr)，兼容多种导出命名：
       - 年份在尾且 4 位: '上海朗炫科目余额表2026.xlsx' -> ('上海朗炫', '2026')
       - 年份在开头 4 位: '2025东轴科目余额表.xlsx'     -> ('东轴',     '2025')
       - 年份在名中 2 位: '上分25科目余额表_导出时间戳.xlsx' -> ('上分', '2025')
       末尾用友导出时间戳 _YYYYMMDDHHMMSS / _YYYYMMDD 须剔除；实体名再剔除 4 位与
       2 位(20-29)年份片段，确保同一主体 km(带25)/gl(无25) 归并（物联合肥25↔物联合肥）。
    """
    base = fn
    for ext in ('.xlsx', '.xls', '.xlsm', '.csv'):
        if base.lower().endswith(ext):
            base = base[:-len(ext)]
            break
    base = re.sub(r'_\d{12,}$', '', base)   # 导出时间戳 YYYYMMDDHHMMSS
    base = re.sub(r'_\d{8}$', '', base)      # 导出时间戳 YYYYMMDD
    m = re.search(r'(19|20)\d{2}', base)     # 全名优先 4 位年份（含尾随 2026）
    yr = m.group(0) if m else '2025'
    prefix, _, _suffix = base.partition(marker)
    ent = re.sub(r'(19|20)\d{2}', '', prefix)        # 去 4 位年份
    ent = re.sub(r'(?<!\d)(2\d)(?!\d)', '', ent)     # 去 2 位年份 20-29
    ent = re.sub(r'年度|年\d*[-~]?\d*月', '', ent)  # 去"年度"、"年1-3月"
    ent = ent.strip(' _-')
    # 个位数前缀补0（1、FY本级公司 → 01、FY本级公司）
    _m0 = re.match(r'^(\d)([、,，.．\s])', ent)
    if _m0 and len(ent) > 2:
        ent = '0' + ent
    return ent, yr


def _discover_entities(data_dir):
    # ⚡ 2026-08-09 治本统一：audit_common.discover_entities（U8 全量/SAP adapter 单主体过滤）
    # → 转本程序嵌套结构 {E: {'km': {y: path}, 'gl': {y: path}}}
    # ⚡ 2026-08-12 治理：改用 audit_common.norm_entities（幂等兼容 U8/SAP 双结构，见
    #   audit_common；消除「SAP 结构 {年:['km','gl']} vs U8 {'km':{年}}」边界踩坑）。
    from audit_common import discover_entities as _de, norm_entities as _norm
    return _norm(_de(data_dir))




# 科目余额表读取缓存：避免空科目判定 / 审定表 / 分月表对同文件反复解析
_KM_CACHE = {}


def _read_km_cached(path, ent=None):
    # ⚡ SAP：km 是单文件 path（str）正常缓存；list 兜底（SAP gl 多文件）
    # ⚡ 2026-08-13 原则（用户定）：生成器只负责生成，数据问题前端解决——不在此区分
    #   SAP/U8，主体解析由前端 read_km（adapter 按 current_comp）完成。ent 仅用于缓存
    #   key 隔离（SAP 区间文件同一路径被多主体调用，按路径缓存会跨主体串数据）。
    _key = (ent, path) if ent else (path if isinstance(path, str) else ('sap_km',))
    if _key in _KM_CACHE:
        return _KM_CACHE[_key]
    g = _read_km(path)
    _KM_CACHE[_key] = g
    return g


def _account_has_data(ents, acc_code, acc_name=None):
    """判断某存货二级科目在所有主体/年度是否【有数据】（TB 父级行非零即视为有数据）。"""
    for ent in sorted(ents):
        for year in sorted(ents[ent]['km']):
            km_path = ents[ent]['km'].get(year)
            if not km_path:
                continue
            km = _read_km_cached(km_path, ent)
            if not km:
                continue
            rec = km.get(acc_code)
            if rec and not is_zero_amount(rec['opening'], rec['debit'], rec['credit'], rec['closing']):
                return True
            # ⚡ 2026-08-10 300 形态：TB 全 10 位子目、无 4 位父级行，且各账套子目码不同
            # （3700 原材料=1403020000/1403050000，非 1403000000）→ 精确匹配失效 → 4 位前缀扫。
            if rec is None and acc_code[:4].isdigit():
                for _c, _r in km.items():
                    if _c.startswith(acc_code[:4]) and not is_zero_amount(
                            _r['opening'], _r['debit'], _r['credit'], _r['closing']):
                        return True
    return False




# GL 读取缓存：同一文件夹内预扫描与主循环复用，避免大文件重复解析（process_folder 开头清空）
_GL_CACHE = {}


def _sap_gl_rows():
    """⚡ 2026-08-30 批次B DAO 收尾：统一逐行读取委托 ledger_backend.read_gl_rows
    （SAP 逐行 + _conv_sap_gl_std 行规范化已内聚 DAO，本函数仅转发，不再自己实现
    adapter 直读/转换；集团模式逐主体聚合防 OOM 也在 DAO 内）。
    历史背景：2026-08-10 SAP GL 行统一转换（adapter 行 {name,date,vtype,vno,cp,sm,dr,cr,month,y}
    → U8 同构 {name,date,voucher,cp,sm,debit,credit,month}）；2026-08-12 集团模式
    （current_comp=None）无参 read_gl() 返回空 → 逐主体读聚合。"""
    from ledger_backend import read_gl_rows as _bk_rows
    return _bk_rows(getattr(_adapter, '_DATA_ROOT', None))


def _conv_sap_gl(rows):
    """adapter 行 → U8 同构行（2026-08-30 批次B：实现已内聚 ledger_backend._conv_sap_gl_std，
    本函数保留为兼容别名，调用点零改动）。"""
    from ledger_backend import _conv_sap_gl_std
    return _conv_sap_gl_std(rows)


def _read_gl_cached(path):
    # ⚡ SAP：gl 是多文件 list（unhashable）→ 用标记缓存键；_read_gl 是导入时绑定的 U8 原始版
    # （patch_audit_common 只替换 AU.read_gl 属性，from-import 绑定不受影响）→ SAP 直接走 adapter。
    # ⚡ 2026-08-12 修复：缓存键含主体（SAP 集团模式每主体一份 GL，全量聚合会 OOM）。
    if _adapter is not None and not isinstance(path, str):
        _comp_iv = _adapter.current_comp()
        if not _comp_iv:
            _p0 = path[0] if isinstance(path, list) and path else path
            # ⚡ 2026-08-26 修复：主体代码在【文件名】开头（1010-1月.xlsx）。原先取路径正则
            #   [\\/](\d{4})[\\/]，AH 路径含年份目录（…/数据/2026/序时账/…）→ 误命中 '2026'
            #   年份 → read_gl('2026') 读错主体 → 存货分月明细/发生额全空。改为文件名优先。
            _m2_iv = re.match(r'^(\d{4})', os.path.basename(str(_p0)))
            if _m2_iv:
                _comp_iv = _m2_iv.group(1)
            else:
                _m_iv = re.search(r'[\\/](\d{4})[\\/]', str(_p0))
                if _m_iv:
                    _comp_iv = _m_iv.group(1)
        _key = ('sap_gl', _comp_iv)
    else:
        _key = path if isinstance(path, str) else ('sap_gl',)
    if _key in _GL_CACHE:
        return _GL_CACHE[_key]
    if _adapter is not None and not isinstance(path, str):
        _comp_iv = _key[1] if isinstance(_key, tuple) and len(_key) > 1 else _adapter.current_comp()
        g = _conv_sap_gl(_adapter.read_gl(_comp_iv) if _comp_iv else [])
    else:
        g = _read_gl(path)
    _GL_CACHE[_key] = g
    return g


def _ym(date_str):
    """'2025-12-31' / '20251231' -> (2025,12)。无法解析返回 None。
    ⚡ 2026-08-26 修复：SAP 序时账日期为 8 位无分隔符 'YYYYMMDD'（如 '20260101'），
    原仅支持 '-/.' 分隔 → 全量解析 None → 分月明细表只有全年合计、成本分析表空壳。"""
    if not date_str:
        return None
    s = str(date_str).strip()
    for sep in ('-', '/', '.'):
        if sep in s:
            parts = s.split(sep)
            if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                return (int(parts[0]), int(parts[1]))
    # SAP 'YYYYMMDD'（8 位纯数字）
    if len(s) == 8 and s.isdigit():
        y, m = int(s[:4]), int(s[4:6])
        if 1 <= m <= 12:
            return (y, m)
    return None


def _iter_comp_gl(ents, year):
    """yield (comp, gl)：逐主体提供 GL 行。
    ⚡ 2026-08-26 指令号/SAP 专属表 group 模式适配：原 `comp = current_comp(); gl =
    _read_gl_cached(None)`——AH 三集团并行 current_comp=None → 恒返回空表。现改为：
    单主体模式用 current_comp；集团模式遍历 ents 各主体，按各主体 gl 路径逐份读取。"""
    if _adapter is not None and _adapter.current_comp():
        _cc = _adapter.current_comp()
        g = _read_gl_cached(None)
        if g:
            yield (_cc, g)
        return
    for _e in sorted(ents):
        _p = (ents[_e].get('gl') or {}).get(year)
        if not _p:
            continue
        try:
            g = _read_gl_cached(_p)
        except Exception:
            continue
        if g:
            yield (_e, g)


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


def _txt(ws, r, c, v, bold=False, align=None, italic=False):
    cell = ws.cell(r, c, v)
    cell.border = SHELL_BORDER
    cell.alignment = align or SHELL_LEFT
    if bold or italic:
        cell.font = Font(name='Times New Roman', size=10,
                         bold=bold, italic=italic, color='595959' if italic else '000000')
    return cell




# ----------------------------------------------------------------------------
# 1) 存货审定表
# ----------------------------------------------------------------------------

def build_inv_by_entity_summary(wb, ents, year):
    """存货按主体余额汇总（2026-08-04 规范：标准行式『主体|期初|期末|跌价|净额』，
    供合并核对程序 det_sn 命中并逐主体提取——原分月变动表为 主体×月份 结构无法行式提取，
    曾致存货/原材料 明细表单体恒空）。期末净额=期末原值+跌价准备（跌价为负数）。"""
    ws = wb.create_sheet('存货按主体明细表')
    ws.cell(1, 1, f'存货按主体余额汇总（{year} 年度；期末净额=期末原值+跌价准备）').font = SHELL_TITLE_FONT
    _hdr(ws, 3, ['主体', '期初原值', '期末原值', '减：存货跌价准备', '期末净额'])
    r = 4
    sum_op = sum_cl = sum_bd = 0.0   # 2026-08-05 前道规范：全集团合计累计
    for ent in sorted(ents):
        km_path = ents[ent]['km'].get(year)
        if not km_path:
            continue
        km = _read_km_cached(km_path, ent)
        if not km:
            continue
        op = cl = 0.0
        for code in INV_LEVEL2:
            if code == '1471':
                continue
            # 2026-08-06 泰国兜底：标准码→泰国代码（1140.01/02/04），否则泰国存货按主体
            # 明细表全 0（泰国代码未归一化）
            _km_code = THAI_INV_MAP.get(code, code) if '泰国' in str(ent) else code
            rec = km.get(_km_code)
            if rec:
                op += rec['opening']
                cl += rec['closing']
        bd = km.get('1471')
        bd_cl = bd['closing'] if bd else 0.0
        # 按主体余额汇总：仅期初/期末原值口径（原 2 字段判断，勿误改 4 字段——本函数无 inc/dec 变量）
        if abs(op) < 0.005 and abs(cl) < 0.005:
            continue
        _txt(ws, r, 1, ent)
        _money(ws, r, 2, op)
        _money(ws, r, 3, cl)
        _money(ws, r, 4, bd_cl)
        _money(ws, r, 5, cl + bd_cl)
        sum_op += op; sum_cl += cl; sum_bd += bd_cl
        r += 1
    # 2026-08-05 前道规范：全集团合计行（核对程序取数依赖）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出全集团合计行
    if len(ents) > 1:
        _txt(ws, r, 1, '全集团合计', bold=True)
        _money(ws, r, 2, sum_op); _money(ws, r, 3, sum_cl)
        _money(ws, r, 4, sum_bd); _money(ws, r, 5, sum_cl + sum_bd)
        for c in range(1, 6):
            ws.cell(r, c).font = SHELL_BOLD
            ws.cell(r, c).fill = SHELL_TOT_FILL
    # 2026-08-05 前道规范化：写入后立即结构校验
    try:
        from audit_common import assert_sheet_standard
        assert_sheet_standard(ws, 'detail', '存货按主体明细表', raise_on_error=False)
    except Exception:
        pass
    return ws


def build_audit_schedule(ents, wb):
    ws = wb.create_sheet('存货审定表')
    ws.cell(1, 1, '存货审定表（按 2 级科目 · 含 1471 存货跌价准备）').font = SHELL_TITLE_FONT
    ws.cell(2, 1, '审定数=期末未审数+审计调整数（审计调整数留空待填）。').font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    headers = ['主体', '年度', '科目代码', '科目名称', '期初数', '本期增加数(借发)',
               '本期减少数(贷发)', '期末未审数', '审计调整数', '审定数',
               '期末未审/期初 变动比例']
    hr = _hdr(ws, 4, headers)
    r = hr
    for ent in sorted(ents):
        # 动态多期支持：按实际数据文件夹的年份读取科目余额表，取代硬编码 2025/2026
        # 2026-08-03 修复：原 sorted(ents[ent].keys()) 误遍历到 'km'/'gl' 键 → km_path 恒 None → 审定表无数据
        for year in sorted(ents[ent].get('km', {})):
            km_path = ents[ent]['km'].get(year)
            if not km_path:
                continue
            km = _read_km_cached(km_path, ent)
            if not km:
                continue
            # 账户层级拆解铁律：① 从一级(TB)出发；③ 三级标注且父=子合计不重复计；④ 0值行不显示。
            # （审定表不再列示一级↔末级勾稽说明行——用户要求勾稽说明不要。）
            start_r = r
            sum_op = sum_inc = sum_dec = sum_cl = 0.0
            wrote = False
            # 2026-08-03：1471 跌价准备不参与原值平铺（其作为备抵行列示于小计之后，
            # 报表科目=原值−跌价准备=净额；原值小计不含跌价，防双计）。
            # ⚡ 2026-08-10 SAP：按实际科目名称分类（用户要求；1406 库存产品→库存商品/产成品/
            # 半成品/在制品、1411=委托加工物资、1404/1407 成本差异；铁律58 名称主键）
            if _is_sap():
                _comp = ent
                _tb_all = _adapter.read_tb_full(getattr(_adapter, '_DATA_ROOT', '') or '',
                                                {_comp: ents.get(_comp, {})})
                _rows_by_cat = {}
                for (ee, cc, nn, yy), vv in _tb_all.items():
                    if ee != _comp or str(yy) != str(year):
                        continue
                    _cat = _sap_inv_cat(nn)
                    if _cat is None or _cat == '存货跌价准备':   # 备抵由下方 1471 块统一处理
                        continue
                    _d = _rows_by_cat.setdefault(_cat, [0.0, 0.0, 0.0, 0.0])
                    _d[0] += float(vv.get('qc') or 0.0); _d[1] += float(vv.get('jf') or 0.0)
                    _d[2] += float(vv.get('df') or 0.0); _d[3] += float(vv.get('qm') or 0.0)
                for _cat in sorted(_rows_by_cat):
                    _d = _rows_by_cat[_cat]
                    if all(abs(x) < 0.005 for x in _d):
                        continue
                    _write_audit_row(ws, r, ent, year, '', _cat,
                                     {'opening': _d[0], 'debit': _d[1],
                                      'credit': _d[2], 'closing': _d[3]}, indent=False)
                    sum_op += _d[0]; sum_inc += _d[1]; sum_dec += _d[2]; sum_cl += _d[3]
                    r += 1
                    wrote = True
            else:
                for code, name in INV_LEVEL2.items():
                    if code == '1471':
                        continue
                    # 2026-08-06 泰国主体：标准码→泰国代码映射（1140.01/02/04），否则泰国存货
                    # 空壳（泰国代码未归一化，1403/1405/1406 精确匹配落空）
                    _km_code = code
                    if '泰国' in str(ent) and code in THAI_INV_MAP:
                        _km_code = THAI_INV_MAP[code]
                    rec = km.get(_km_code)
                    if rec is None:
                        continue
                    # 铁律④：一级(4位)控制数为 0 则不显示
                    if is_zero_amount(rec['opening'], rec['debit'], rec['credit'], rec['closing']):
                        continue
                    children = INV_CHILDREN.get(code, [])
                    child_recs = [(cc, cn, km.get(cc)) for cc, cn in children if km.get(cc) is not None]
                    # 铁律③：含末级子目时标注「（含三级明细）」，父级带 TB 值、子目仅下钻不重复计
                    disp_name = name + ('（含三级明细）' if child_recs else '')
                    _write_audit_row(ws, r, ent, year, code, disp_name, rec, indent=False)
                    # 2026-08-07 修复：父级行【显示但不计入小计】（原父级+子目双加 → g 杭州铁城
                    # 1411 周转材料 480,497.06 = 141101+141102 被计 2 次 → 小计 135,502,715.32
                    # vs TB 135,022,218.26 差 480,497.06）。子目（下钻）才计入——父级值=子和，
                    # 小计=子目和=父级值，正常账套无变化，仅消除双计。
                    if not child_recs:
                        sum_op += rec['opening']; sum_inc += rec['debit']
                        sum_dec += rec['credit']; sum_cl += rec['closing']
                    r += 1
                    for cc, cn, crec in child_recs:
                        # 铁律④：末级子目为 0 也不显示
                        if is_zero_amount(crec['opening'], crec['debit'], crec['credit'], crec['closing']):
                            continue
                        _write_audit_row(ws, r, ent, year, cc, cn, crec, indent=True)
                        sum_op += crec['opening']; sum_inc += crec['debit']
                        sum_dec += crec['credit']; sum_cl += crec['closing']
                        r += 1
                    wrote = True
            if wrote:
                # 核算主体(年度)小计
                _txt(ws, r, 1, ent)
                _txt(ws, r, 2, year, align=SHELL_CEN)
                _txt(ws, r, 3, '', align=SHELL_CEN)
                _txt(ws, r, 4, '小计', bold=True)
                # 2026-08-06 协议：小计/审定数写数值（=sum_* 累计；公式 data_only 读 None 收回读不回）
                _money(ws, r, 5, round(sum_op, 2))
                _money(ws, r, 6, round(sum_inc, 2))
                _money(ws, r, 7, round(sum_dec, 2))
                _money(ws, r, 8, round(sum_cl, 2))
                _money(ws, r, 9, None)
                # 审定数 = 期末未审合计（审计调整列留空）
                ws.cell(r, 10, round(sum_cl, 2)).number_format = SHELL_NUM
                ws.cell(r, 10).border = SHELL_BORDER
                ws.cell(r, 10).alignment = SHELL_RGT
                for c in range(1, 12):
                    ws.cell(r, c).font = SHELL_BOLD
                    ws.cell(r, c).fill = SHELL_TOT_FILL
                    ws.cell(r, c).border = SHELL_BORDER
                r += 1
                # —— 备抵项：减：存货跌价准备 → 存货净额（报表科目=原值−跌价准备）——
                _rec_bd = km.get('1471')
                if _rec_bd and not is_zero_amount(_rec_bd['opening'], _rec_bd['debit'],
                                                  _rec_bd['credit'], _rec_bd['closing']):
                    _write_audit_row(ws, r, ent, year, '1471', '减：存货跌价准备', _rec_bd, indent=False)
                    r_bd = r
                    r += 1
                    # 存货净额 = 原值小计 + 跌价准备（跌价为负数，数值上=原值−跌价金额）
                    _txt(ws, r, 1, '', align=SHELL_CEN)
                    _txt(ws, r, 2, '', align=SHELL_CEN)
                    _txt(ws, r, 3, '', align=SHELL_CEN)
                    _txt(ws, r, 4, '存货净额', bold=True)
                    r_net = r
                    for c_s, c_t in [(5, 5), (6, 6), (7, 7), (8, 8)]:
                        _money(ws, r, c_s, '=%s%d+%s%d' % (get_column_letter(c_t), r - 2,
                                                           get_column_letter(c_t), r_bd))
                    _money(ws, r, 9, None)
                    ws.cell(r, 10, '=H%d' % r_net).number_format = SHELL_NUM
                    ws.cell(r, 10).border = SHELL_BORDER
                    ws.cell(r, 10).alignment = SHELL_RGT
                    for c in range(1, 12):
                        ws.cell(r, c).font = SHELL_BOLD
                        ws.cell(r, c).fill = SHELL_TOT_FILL
                        ws.cell(r, c).border = SHELL_BORDER
                    r += 1
    # 2026-08-07 修复（S 存货 877M vs 209M）：审定表缺【全集团合计】行——
    # 核对器 _extract 取最后合计行 = 最后一个主体（母公司）小计 209M，
    # 而自建TB=全部主体末级之和 877M → 口径不一致假差异。加全集团合计行：
    # 期末未审 = 全部主体小计期末之和（父级行已剔除不计入小计，此处按小计行汇总）。
    # 2026-08-08 修复：混年表（g 2025+2026 同表）原不按年度过滤 → 全集团合计=两年之和
    # （9.56亿=4.77+4.79，自建TB 2025 单年 4.77亿 → 2 倍假差异）。改为按年度分组，
    # 每个年度一行全集团合计（B 列标年度，核对器按核对年份取数）。
    grand = {}   # year -> [op, cl]
    for _i in range(hr + 1, r):
        _a4 = str(ws.cell(_i, 4).value or '')
        if _a4 == '小计':
            _yy = str(ws.cell(_i, 2).value or '')
            _v = ws.cell(_i, 8).value
            if isinstance(_v, (int, float)):
                g = grand.setdefault(_yy, [0.0, 0.0])
                g[1] += _v
                g[0] += ws.cell(_i, 5).value or 0.0
    for _yy in sorted(grand):
        _op, _cl = grand[_yy]
        # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出全集团合计行
        if len(ents) <= 1:
            break
        _txt(ws, r, 1, '全集团合计', bold=True)
        _txt(ws, r, 2, _yy, align=SHELL_CEN)
        _txt(ws, r, 3, '', align=SHELL_CEN)
        _txt(ws, r, 4, '全集团合计', bold=True)
        # 2026-08-07 修复（S 存货 1.75B 双计）：仅审定数/期末未审列填数值，
        # 期初/借发/贷发留空——finalize recalc_totals 会把含『合计』行的数值列
        # 覆盖为 =SUM(区间)，而区间含全部数据行（父级+子目）→ 双计（1.75B=2×877M
        # 近似）。审定数=全部主体小计期末之和（父级已剔除不计入小计），为核对基准。
        _money(ws, r, 5, None); _money(ws, r, 6, None); _money(ws, r, 7, None)
        _money(ws, r, 8, round(_cl, 2))
        _money(ws, r, 9, None)
        _money(ws, r, 10, round(_cl, 2))
        for c in range(1, 12):
            ws.cell(r, c).font = SHELL_BOLD
            ws.cell(r, c).fill = SHELL_TOT_FILL
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    # 列宽
    for col, w in zip(range(1, 12), [16, 8, 12, 18, 16, 16, 16, 16, 14, 16, 18]):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = 'A5'
    return ws


def _write_audit_row(ws, r, ent, year, code, name, rec, indent):
    _txt(ws, r, 1, ent)
    _txt(ws, r, 2, year, align=SHELL_CEN)
    _txt(ws, r, 3, code, align=SHELL_CEN)
    _txt(ws, r, 4, ('   ' + name) if indent else name)
    _money(ws, r, 5, rec['opening'])
    _money(ws, r, 6, rec['debit'])
    _money(ws, r, 7, rec['credit'])
    _money(ws, r, 8, rec['closing'])
    _money(ws, r, 9, None)  # 审计调整数 留空
    # 审定数 = 期末未审 + 审计调整（公式）
    ws.cell(r, 10, f'=H{r}+I{r}').number_format = SHELL_NUM
    ws.cell(r, 10).border = SHELL_BORDER
    ws.cell(r, 10).alignment = SHELL_RGT
    # 变动比例
    op = rec['opening']
    cl = rec['closing']
    if op not in (0, 0.0) and abs(op) > 1e-9:
        ratio = cl / op - 1
        cell = ws.cell(r, 11, ratio)
        cell.number_format = '0.0%'
    else:
        if cl not in (0, 0.0):
            cell = ws.cell(r, 11, '期初为0/新增')
        else:
            cell = ws.cell(r, 11, '—')
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    for c in range(1, 12):
        ws.cell(r, c).border = SHELL_BORDER




# ----------------------------------------------------------------------------
# 3) 变动分析表（最末级明细 期末 vs 期初）
# ----------------------------------------------------------------------------

def build_variation_sheet(ws, ents, title, leaf_codes):
    ws.cell(1, 1, title).font = SHELL_TITLE_FONT
    headers = ['主体', '科目代码', '科目名称', '2025期初', '2025期末', '2025变动比例',
               '2026期初', '2026期末', '2026变动比例', '变动判断']
    hr = _hdr(ws, 4, headers)
    r = hr
    for ent in sorted(ents):
        # 动态多期：取实际年份，取前两个（或一个）作为对比期
        ent_years = sorted(ents[ent].keys())
        km_by_year = {}
        for y in ent_years:
            km_path = ents[ent]['km'].get(y)
            if km_path:
                # ⚡ 2026-08-13 原则：前端解决数据问题（adapter 按 current_comp 解析主体）
                km_by_year[y] = _read_km(km_path)
        if not km_by_year:
            continue
        y1 = ent_years[0]
        y2 = ent_years[1] if len(ent_years) > 1 else ent_years[0]
        km_y1 = km_by_year.get(y1, {})
        km_y2 = km_by_year.get(y2, {})
        # 更新表头列标题为实际年份
        # 动态表头：按实际年份渲染
        _txt(ws, 1, 1, title.replace('(期末 vs 期初)', f'({y1}末 vs {y2}初)'), overwrite=True) if False else None  # noop: title already set
        # 重写表头
        hdr_texts = ['主体', '科目代码', '科目名称', f'{y1}期初', f'{y1}期末', f'{y1}变动比例',
                     f'{y2}期初', f'{y2}期末', f'{y2}变动比例', '变动判断']
        for c, h in enumerate(hdr_texts, 1):
            ws.cell(4, c).value = h
            ws.cell(4, c).font = SHELL_HFONT
            ws.cell(4, c).fill = SHELL_HFILL
            ws.cell(4, c).alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')
            ws.cell(4, c).border = SHELL_BORDER
        has = False
        for code in leaf_codes:
            rec_y1 = km_y1.get(code)
            rec_y2 = km_y2.get(code)
            if rec_y1 is None and rec_y2 is None:
                continue
            has = True
            op_y1 = rec_y1['opening'] if rec_y1 else 0.0
            cl_y1 = rec_y1['closing'] if rec_y1 else 0.0
            op_y2 = rec_y2['opening'] if rec_y2 else 0.0
            cl_y2 = rec_y2['closing'] if rec_y2 else 0.0
            name = (rec_y1 or rec_y2)['name']
            ratio_y1 = (cl_y1 / op_y1 - 1) if abs(op_y1) > 1e-9 else (None if abs(cl_y1) < 1e-9 else '期初0/新增')
            ratio_y2 = (cl_y2 / op_y2 - 1) if abs(op_y2) > 1e-9 else (None if abs(cl_y2) < 1e-9 else '期初0/新增')
            judge = _judge(ratio_y1, ratio_y2)
            _txt(ws, r, 1, ent)
            _txt(ws, r, 2, code, align=SHELL_CEN)
            _txt(ws, r, 3, name)
            _money(ws, r, 4, op_y1)
            _money(ws, r, 5, cl_y1)
            _ratio(ws, r, 6, ratio_y1)
            _money(ws, r, 7, op_y2)
            _money(ws, r, 8, cl_y2)
            _ratio(ws, r, 9, ratio_y2)
            jc = _txt(ws, r, 10, judge, align=SHELL_CEN)
            if judge == '关注':
                jc.font = Font(name='Times New Roman', bold=True, color='C00000')
            for c in range(1, 11):
                ws.cell(r, c).border = SHELL_BORDER
            r += 1
        # 衔接校验：y1末 vs y2初
        if has:
            cont = _check_continuity(km_y1, km_y2, leaf_codes)
            if cont:
                cr = _txt(ws, r, 1, f'  {ent} 衔接校验（{y1}末↔{y2}初）：{cont}')
                cr.font = Font(name='Times New Roman', size=10, italic=True, color='C00000')
                for c in range(1, 11):
                    ws.cell(r, c).border = SHELL_BORDER
                r += 1
    for col, w in zip(range(1, 11), [16, 12, 20, 16, 16, 14, 16, 16, 14, 14]):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = 'A5'
    return ws


def _ratio(ws, r, c, v):
    if v is None:
        cell = ws.cell(r, c, '—')
    elif isinstance(v, str):
        cell = ws.cell(r, c, v)
    else:
        cell = ws.cell(r, c, v)
        cell.number_format = '0.0%'
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    return cell


def _judge(r25, r26):
    for v in (r25, r26):
        if isinstance(v, str):
            return '关注'
        if isinstance(v, (int, float)) and abs(v) >= 0.3:
            return '关注'
    return '正常'


def _check_continuity(km25, km26, leaf_codes):
    msgs = []
    for code in leaf_codes:
        c25 = km25.get(code)
        c26 = km26.get(code)
        if c25 and c26:
            if abs(c25['closing'] - c26['opening']) > 0.01:
                msgs.append(f'{code} 25末{c25["closing"]:,.2f}≠26初{c26["opening"]:,.2f}')
    return '；'.join(msgs) if msgs else ''


# ----------------------------------------------------------------------------
# 4) 生产成本 成本分析表
# ----------------------------------------------------------------------------

def build_cost_analysis_sheet(ws, ents, title, year_filter=None):
    ws.cell(1, 1, title).font = SHELL_TITLE_FONT
    headers = ['主体', '年度', '月份', '直接材料', '直接人工', '制造费用转入', '自制半成品', '其他增加',
               '本月增加合计', '材料占比', '工资占比', '制造占比', '自制半成品占比', '其他占比', '占比异常']
    # ⚡ 2026-08-11 N4：双行表头分组（规则2）——行3 大类（金额/占比），表头仍行4
    _grp_row = 3
    for _g, _c0, _c1 in [('基础信息', 1, 3), ('金额（生产成本各月增加）', 4, 9),
                         ('占比（按月）', 10, 14), ('异常', 15, 15)]:
        ws.merge_cells(start_row=_grp_row, start_column=_c0, end_row=_grp_row, end_column=_c1)
        _gc = ws.cell(_grp_row, _c0, _g)
        _gc.font = SHELL_BOLD
        _gc.fill = PatternFill('solid', fgColor='DDEBF7')
        _gc.alignment = Alignment(horizontal='center', vertical='center')
        for _j in range(_c0, _c1 + 1):
            ws.cell(_grp_row, _j).border = SHELL_BORDER
    hr = _hdr(ws, 4, headers)
    r = hr
    comps = ['直接材料', '直接人工', '制造费用转入', '自制半成品', '其他增加']
    for ent in sorted(ents):
        for year in sorted(ents[ent]['km']):
            if year_filter is not None and str(year) != str(year_filter):
                continue

            gl_path = ents[ent]['gl'].get(year)
            if not gl_path:
                continue
            if _adapter is not None and not isinstance(gl_path, str):
                # ⚡ 2026-08-12 修复：原 _sap_gl_rows() 集团模式聚合全量 88 家 →
                #   每主体调用=88 倍重复 + MemoryError。此处已有 ent 上下文，直接按
                #   当前主体读（gl_path 含 {comp} 路径可解析；防 OOM 与重复累积）。
                _comp_iv = _adapter.current_comp()
                if not _comp_iv:
                    _p0 = gl_path[0] if isinstance(gl_path, list) and gl_path else gl_path
                    # ⚡ 2026-08-26 修复：主体代码在【文件名】开头；原路径正则误命中年份目录
                    _m2_iv = re.match(r'^(\d{4})', os.path.basename(str(_p0)))
                    if _m2_iv:
                        _comp_iv = _m2_iv.group(1)
                    else:
                        _m_iv = re.search(r'[\\/](\d{4})[\\/]', str(_p0))
                        if _m_iv:
                            _comp_iv = _m_iv.group(1)
                gl = _conv_sap_gl(_adapter.read_gl(_comp_iv) if _comp_iv else [])
            else:
                gl = _read_gl(gl_path)
            # 收集 生产成本 借 side 按月按类别
            monthly = {}  # ym -> {comp: amt}
            for row in gl:
                nm = row['name']
                if not (nm == PROD_NAME or nm.startswith(PROD_NAME + '-')):
                    continue
                if not row['debit']:
                    continue
                ym = _ym(row['date'])
                if ym is None:
                    continue
                cat = classify_prod_inc(nm, row.get('cp'))
                monthly.setdefault(ym, {}).setdefault(cat, 0.0)
                monthly[ym][cat] += row['debit']
            if not monthly:
                continue
            # 年度平均占比
            tot = {c: 0.0 for c in comps}
            grand = 0.0
            for ym, d in monthly.items():
                for c in comps:
                    tot[c] += d.get(c, 0.0)
                grand += sum(d.get(c, 0.0) for c in comps)
            avg = {c: (tot[c] / grand if grand else 0.0) for c in comps}
            block_first = r
            for ym in sorted(monthly):
                d = monthly[ym]
                s = sum(d.get(c, 0.0) for c in comps)
                _txt(ws, r, 1, ent)
                _txt(ws, r, 2, year, align=SHELL_CEN)
                _txt(ws, r, 3, f'{ym[0]}-{ym[1]:02d}', align=SHELL_CEN)
                _money(ws, r, 4, d.get('直接材料', 0.0))
                _money(ws, r, 5, d.get('直接人工', 0.0))
                _money(ws, r, 6, d.get('制造费用转入', 0.0))
                _money(ws, r, 7, d.get('自制半成品', 0.0))
                _money(ws, r, 8, d.get('其他增加', 0.0))
                _money(ws, r, 9, s)
                _pct(ws, r, 10, d.get('直接材料', 0.0) / s if s else 0.0)
                _pct(ws, r, 11, d.get('直接人工', 0.0) / s if s else 0.0)
                _pct(ws, r, 12, d.get('制造费用转入', 0.0) / s if s else 0.0)
                _pct(ws, r, 13, d.get('自制半成品', 0.0) / s if s else 0.0)
                _pct(ws, r, 14, d.get('其他增加', 0.0) / s if s else 0.0)
                # 异常判定（列15 占比异常；2026-08-03 修复：原误写第13列覆盖自制半成品占比）
                anom = []
                for i, c in enumerate(comps, start=9):
                    share = (d.get(c, 0.0) / s) if s else 0.0
                    if abs(share - avg[c]) >= 0.15:
                        anom.append(headers[i].replace('占比', ''))
                _txt(ws, r, 15, ('、'.join(anom) if anom else '—'), align=SHELL_CEN)
                for c in range(1, 16):
                    ws.cell(r, c).border = SHELL_BORDER
                r += 1
            # 年度合计行（2026-08-03 用户要求：年度平均占比改为合计，加计全年合计发生额）
            _txt(ws, r, 1, ent, bold=True)
            _txt(ws, r, 2, year, align=SHELL_CEN, )
            _txt(ws, r, 3, '年度合计', align=SHELL_CEN, )
            for c in range(4, 9):
                ws.cell(r, c).border = SHELL_BORDER
            for j, c in enumerate(comps):
                _money(ws, r, 4 + j, tot[c])          # 全年各成本要素合计发生额
            _money(ws, r, 9, grand, )                 # 全年增加合计
            _pct(ws, r, 10, avg['直接材料'])
            _pct(ws, r, 11, avg['直接人工'])
            _pct(ws, r, 12, avg['制造费用转入'])
            _pct(ws, r, 13, avg['自制半成品'])
            _pct(ws, r, 14, avg['其他增加'])
            _txt(ws, r, 15, '—', align=SHELL_CEN)
            for c in range(1, 16):
                if c in (1, 3):
                    ws.cell(r, c).font = SHELL_BOLD
                ws.cell(r, c).border = SHELL_BORDER
            r += 2  # 块间空行
    for col, w in zip(range(1, 16), [16, 8, 12, 14, 14, 14, 14, 12, 14, 11, 11, 11, 11, 11, 14]):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = 'A5'
    return ws


def _pct(ws, r, c, v):
    cell = ws.cell(r, c, v)
    cell.number_format = '0.0%'
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    return cell


# ----------------------------------------------------------------------------
# 生产成本 分月变动表（含上年对比）
# ----------------------------------------------------------------------------

def _agg_prod_account(ents):
    """归集生产成本(5001)全部主体/年度的 GL 发生额，按成本要素(借)与转出类别(贷)分月。"""
    result = {}
    for ent in sorted(ents):
        for year in sorted(ents[ent]['km']):
            gl_path = ents[ent]['gl'].get(year)
            km_path = ents[ent]['km'].get(year)
            if not gl_path:
                continue
            km = _read_km_cached(km_path, ent) if km_path else {}
            gl = _read_gl_cached(gl_path)
            vidx = _build_vidx(gl)
            inc = {c: {} for c in PROD_INC}
            dec = {c: {} for c in PROD_DEC}
            inc_t = {c: 0.0 for c in PROD_INC}
            dec_t = {c: 0.0 for c in PROD_DEC}
            months = set()
            for row in gl:
                nm = row['name']
                if not (nm == PROD_NAME or nm.startswith(PROD_NAME + '-')):
                    continue
                ym = _ym(row['date'])
                if ym is None:
                    continue
                months.add(ym)
                if row['debit']:
                    cp_eff = (row.get('cp') or '').strip() or _recon_cp(row, vidx)
                    cat = classify_prod_inc(nm, cp_eff)
                    inc[cat][ym] = inc[cat].get(ym, 0.0) + row['debit']
                    inc_t[cat] = inc_t.get(cat, 0.0) + row['debit']
                if row['credit']:
                    cp_eff = (row.get('cp') or '').strip() or _recon_cp(row, vidx)
                    cat = classify_prod_dec(cp_eff)
                    dec[cat][ym] = dec[cat].get(ym, 0.0) + row['credit']
                    dec_t[cat] = dec_t.get(cat, 0.0) + row['credit']
            # 2026-08-03 修复：TB 控制数按代码取，但泰国生产成本代码=5400（非中国准则5001）
            # → 按名称兜底（名含『生产成本』的父级科目），否则泰国 km_op/km_cl=0 勾稽失真。
            km_op = km_cl = 0.0
            _pc = km.get(PROD_CODE)
            if _pc:
                km_op = _pc.get('opening', 0.0); km_cl = _pc.get('closing', 0.0)
            else:
                for _c, _v in km.items():
                    if str(_v.get('name', '')).strip() == PROD_NAME and len(str(_c).replace('.', '')) <= 6:
                        km_op = _v.get('opening', 0.0); km_cl = _v.get('closing', 0.0)
                        break
            result[(ent, year)] = {'inc': inc, 'dec': dec, 'inc_t': inc_t, 'dec_t': dec_t,
                                   'months': months, 'km_op': km_op, 'km_cl': km_cl}
    return result


def build_prod_monthly_summary(ws, ents, title, year_filter=None):
    """第一张分月表：仅列示 增减变动（期初/本期增加/本期减少/期末），逐主体、逐月。"""
    agg = _agg_prod_account(ents)
    cp_empty = _prod_cp_empty(ents)
    ws.cell(1, 1, title).font = SHELL_TITLE_FONT


# ============================================================================
# 泰国存货映射 + 存货科目配置（2026-08-03 用户方法论：泰国 1140 系列纳入）
# ============================================================================
THAI_INV_CODE = {
    '原材料': '1140.01', '库存商品': '1140.02', '发出商品': '1140.04',
}
THAI_INV_GL_NAME = {
    '存货-原材料': '原材料', '存货-成品': '库存商品', '存货-发出商品': '发出商品',
}

# 存货 分月变动表/对应科目明细表 分析科目（2026-08-02 用户方法论：按科目逐张出表）
# (代码, 名称, 名称匹配词, 是否备抵)
ANALYSIS_ACCOUNTS = [
    ('1403', '原材料', ['原材料'], False),
    ('1405', '库存商品', ['库存商品'], False),
    ('1406', '发出商品', ['发出商品'], False),
    ('1408', '委托加工物资', ['委托加工物资'], False),
    ('1411', '周转材料', ['周转材料'], False),
    ('1412', '包装物', ['包装物'], False),
    ('1413', '低值易耗品', ['低值易耗品'], False),
    ('1471', '存货跌价准备', ['存货跌价准备'], True),
]


def _build_vidx(gl):
    """构建同号凭证索引 {(date, vtype, vno): [rows]}，用于空对方科目反推。"""
    vidx = {}
    for row in gl:
        key = (row.get('date'), row.get('vtype'), row.get('vno'))
        vidx.setdefault(key, []).append(row)
    return vidx


def _recon_cp(row, vidx):
    """对方科目(cp)为空时，从『同号凭证的对方行』反推真实对方科目名。

    解法：按 (日期,凭证字,凭证号) 定位同号凭证的其余行，取相反方向金额合计最大的
    非同名科目行作为反推的对方科目。2026-08-03 修复：月末结转大凭证（一贷多借）——
    旧逻辑取『首个非同名科目』会把 10.89 亿误归制造费用；改为取相反方向金额合计
    最大的科目（借找贷、贷找借），正确归到生产成本。"""
    sibs = vidx.get((row.get('date'), row.get('vtype'), row.get('vno')))
    if not sibs:
        return ''
    nm = (row.get('name') or '').strip()
    is_debit = bool(row.get('debit'))
    amt = {}
    for s in sibs:
        if s is row:
            continue
        sn = (s.get('name') or '').strip()
        if not sn or sn == nm:
            continue
        if is_debit:
            v = s.get('credit') or 0.0
        else:
            v = s.get('debit') or 0.0
        if abs(v) > 1e-9:
            amt[sn] = amt.get(sn, 0.0) + abs(v)
    if not amt:
        return ''
    return max(amt.items(), key=lambda x: x[1])[0]


def _acc_match(nm, targets):
    """科目名匹配：精确相等 或 以某 target 开头（带 '-' 分隔）。
    ⚡ 2026-08-10 SAP：子目名=父级-子级（'库存产品-库存商品'），target 在名称中间 → 包含匹配。"""
    nm = (nm or '').strip()
    if not nm:
        return False
    for t in targets:
        if nm == t or nm.startswith(t + '-') or nm.startswith(t + '　'):
            return True
        if _is_sap() and t and t in nm:
            return True
    return False


def _acc_match_thai(nm, targets):
    """泰国 GL 科目名匹配：『存货-原材料』→『原材料』等映射后命中。"""
    nm = (nm or '').strip()
    std = THAI_INV_GL_NAME.get(nm)
    if std is None:
        return False
    return _acc_match(std, targets)


def _agg_inv_account(ents, acc_code, acc_name, targets, contra):
    """归集存货类科目（如 原材料 1403）全部主体/年度的 GL 发生额，按『对方科目』语义归类。

    返回 { (ent, year): {'inc': {cat:{ym:amt}}, 'dec': {cat:{ym:amt}},
                          'inc_t': {cat:总}, 'dec_t': {cat:总},
                          'inc_raw': {ym: 借方合计}, 'dec_raw': {ym: 贷方合计},
                          'months': set, 'km_op': 期初, 'km_cl': 期末,
                          'inc_cats': [...], 'dec_cats': [...], 'contra': bool,
                          'other_inc': {ym:{cp:amt}}, 'other_dec': {ym:{cp:amt}}} }
    2026-08-03 增强：
    ①泰国 1140 系列映射（GL 名『存货-原材料』等）；
    ②备抵(contra)贷方正=计提、贷方红字(负)=转销（Z 母公司转销记贷方负数，正负拆分）；
    ③「其他增加/其他减少」额外按对方科目分月明细收集（供分月变动表段尾打开）。"""
    if contra:
        inc_cats = ['转回']
        dec_cats = ['计提']
    else:
        inc_cats = STD_INC_CATS
        dec_cats = STD_DEC_CATS
    result = {}
    for ent in sorted(ents):
        for year in sorted(ents[ent]['km']):
            gl_path = ents[ent]['gl'].get(year)
            km_path = ents[ent]['km'].get(year)
            if not gl_path:
                continue
            km = _read_km_cached(km_path, ent) if km_path else {}
            gl = _read_gl_cached(gl_path)
            vidx = _build_vidx(gl)
            inc = {c: {} for c in inc_cats}
            dec = {c: {} for c in dec_cats}
            inc_t = {c: 0.0 for c in inc_cats}
            dec_t = {c: 0.0 for c in dec_cats}
            inc_raw = {}
            dec_raw = {}
            months = set()
            other_inc = {}
            other_dec = {}
            for row in gl:
                nm = row['name']
                if not (_acc_match(nm, targets) or _acc_match_thai(nm, targets)):
                    continue
                ym = _ym(row['date'])
                if ym is None:
                    continue
                months.add(ym)
                if contra:
                    if row['credit']:
                        if row['credit'] > 0:
                            dec['计提'][ym] = dec['计提'].get(ym, 0.0) + row['credit']
                            dec_t['计提'] += row['credit']
                            dec_raw[ym] = dec_raw.get(ym, 0.0) + row['credit']
                        else:
                            inc['转回'][ym] = inc['转回'].get(ym, 0.0) + abs(row['credit'])
                            inc_t['转回'] += abs(row['credit'])
                            inc_raw[ym] = inc_raw.get(ym, 0.0) + abs(row['credit'])
                    if row['debit']:
                        inc['转回'][ym] = inc['转回'].get(ym, 0.0) + row['debit']
                        inc_t['转回'] += row['debit']
                        inc_raw[ym] = inc_raw.get(ym, 0.0) + row['debit']
                else:
                    cp_eff = (row.get('cp') or '').strip() or _recon_cp(row, vidx)
                    if row['debit']:
                        cat = _std_inv_cat(cp_eff, 'inc')
                        if cat in inc:
                            inc[cat][ym] = inc[cat].get(ym, 0.0) + row['debit']
                            inc_t[cat] += row['debit']
                            inc_raw[ym] = inc_raw.get(ym, 0.0) + row['debit']
                            if cat == '其他增加':
                                d2 = other_inc.setdefault(ym, {})
                                d2[cp_eff or '(空对方科目)'] = d2.get(cp_eff or '(空对方科目)', 0.0) + row['debit']
                    if row['credit']:
                        cat = _std_inv_cat(cp_eff, 'dec')
                        if cat in dec:
                            dec[cat][ym] = dec[cat].get(ym, 0.0) + row['credit']
                            dec_t[cat] += row['credit']
                            dec_raw[ym] = dec_raw.get(ym, 0.0) + row['credit']
                            if cat == '其他减少':
                                d2 = other_dec.setdefault(ym, {})
                                d2[cp_eff or '(空对方科目)'] = d2.get(cp_eff or '(空对方科目)', 0.0) + row['credit']
            # 控制数取【父级科目代码】权威行：TB 父=子合计；泰国 1140 兜底
            _prec = km.get(acc_code) or km.get(THAI_INV_CODE.get(acc_name, '')) or None
            km_op = _prec['opening'] if _prec else 0.0
            km_cl = _prec['closing'] if _prec else 0.0
            result[(ent, year)] = {
                'inc': inc, 'dec': dec, 'inc_t': inc_t, 'dec_t': dec_t,
                'inc_raw': inc_raw, 'dec_raw': dec_raw, 'months': months,
                'km_op': km_op, 'km_cl': km_cl,
                'inc_cats': inc_cats, 'dec_cats': dec_cats, 'contra': contra,
                'other_inc': other_inc, 'other_dec': other_dec,
            }
    return result


def build_inv_monthly_summary(ws, ents, title, acc_code, acc_name, targets, contra, year_filter=None):
    """第一张分月表（列扩展版，2026-08-03 用户方法论）：逐主体·逐月·逐对应科目类别列示增减。
    列 = 主体 | 月份 | 期初余额 | [本期增加-对应科目类别...] | [本期减少-对应科目类别...] | 期末余额。
    类别列动态生成（全集团该年有数据才建列）；「其他增加/其他减少」列后跟主体段尾
    「其他构成」块（斜体）——按对方科目打开其他，替代原《对应科目分月明细表》。
    2026-08-03 用户要求：删除第 2 行『数据源自…』说明；勾稽行删除。"""
    agg = _agg_inv_account(ents, acc_code, acc_name, targets, contra)
    ws.cell(1, 1, title).font = SHELL_TITLE_FONT

    items = [(ent, year, d) for (ent, year), d in agg.items()
             if year_filter is None or str(year) == str(year_filter)]
    if not items:
        return ws
    d0 = items[0][2]
    if contra:
        inc_cols = [c for c in d0['dec_cats'] if any(abs(d['dec_t'].get(c, 0.0)) > 0.005 for _e, _y, d in items)]
        dec_cols = [c for c in d0['inc_cats'] if any(abs(d['inc_t'].get(c, 0.0)) > 0.005 for _e, _y, d in items)]
    else:
        inc_cols = [c for c in d0['inc_cats'] if any(abs(d['inc_t'].get(c, 0.0)) > 0.005 for _e, _y, d in items)]
        dec_cols = [c for c in d0['dec_cats'] if any(abs(d['dec_t'].get(c, 0.0)) > 0.005 for _e, _y, d in items)]
    n_inc, n_dec = len(inc_cols), len(dec_cols)
    c0 = 4
    c_dec = c0 + n_inc
    c_end = c_dec + n_dec
    NC = c_end
    if contra:
        inc_defs = [(c0 + i, c, 'dec') for i, c in enumerate(inc_cols)]
        dec_defs = [(c_dec + i, c, 'inc') for i, c in enumerate(dec_cols)]
    else:
        inc_defs = [(c0 + i, c, 'inc') for i, c in enumerate(inc_cols)]
        dec_defs = [(c_dec + i, c, 'dec') for i, c in enumerate(dec_cols)]

    # —— 两行表头 ——
    hr1 = 4
    for j, h in enumerate(['主体', '月份', '期初余额'], 1):
        cc = ws.cell(hr1, j, h); cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT
        cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    if n_inc:
        ws.merge_cells(start_row=hr1, start_column=c0, end_row=hr1, end_column=c0 + n_inc - 1)
        cc = ws.cell(hr1, c0, '本期增加'); cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT
        cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    if n_dec:
        ws.merge_cells(start_row=hr1, start_column=c_dec, end_row=hr1, end_column=c_dec + n_dec - 1)
        cc = ws.cell(hr1, c_dec, '本期减少'); cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT
        cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    cc = ws.cell(hr1, c_end, '期末余额'); cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT
    cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    hr2 = hr1 + 1
    for j in range(1, 4):
        ws.cell(hr2, j).fill = SHELL_HFILL; ws.cell(hr2, j).border = SHELL_BORDER
    for i, c in enumerate(inc_cols):
        cc = ws.cell(hr2, c0 + i, c); cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT
        cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    for i, c in enumerate(dec_cols):
        cc = ws.cell(hr2, c_dec + i, c); cc.fill = SHELL_HFILL; cc.font = SHELL_HFONT
        cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    ws.cell(hr2, c_end).fill = SHELL_HFILL; ws.cell(hr2, c_end).border = SHELL_BORDER
    r = hr2 + 1

    def _col_of(cat, cols, base):
        for i, c in enumerate(cols):
            if c == cat:
                return base + i
        return None

    for ent, year, d in items:
        if not d['months']:
            _txt(ws, r, 1, ent)
            _txt(ws, r, 2, '全年合计', bold=True, align=SHELL_CEN)
            _money(ws, r, 3, d['km_op'])
            for c in range(c0, c_end):
                _money(ws, r, c, 0.0)
            _money(ws, r, c_end, d['km_cl'])
            for c in range(1, NC + 1):
                ws.cell(r, c).border = SHELL_BORDER
            r += 1
            continue
        carry = d['km_op']
        first_open = carry
        final_close = carry
        sum_cols = [0.0] * (NC + 1)
        for m in range(1, 13):
            ym = (int(year), m)
            open_here = carry
            row_vals = {}
            for col, cat, drc in inc_defs:
                row_vals[col] = (d['dec'] if drc == 'dec' else d['inc'])[cat].get(ym, 0.0)
            for col, cat, drc in dec_defs:
                row_vals[col] = (d['dec'] if drc == 'dec' else d['inc'])[cat].get(ym, 0.0)
            bs = sum(row_vals.get(c, 0.0) for c, _x, _y in inc_defs)
            cs = sum(row_vals.get(c, 0.0) for c, _x, _y in dec_defs)
            close_here = open_here + bs - cs
            mon_label = f'{year}-{m:02d}'
            if ym not in d['months'] and str(year) == '2026':
                mon_label += '(未抽取)'
            _txt(ws, r, 1, ent if m == 1 else '')
            _txt(ws, r, 2, mon_label, align=SHELL_CEN)
            _money(ws, r, 3, open_here)
            for col in range(c0, c_end):
                _money(ws, r, col, row_vals.get(col, 0.0))
            _money(ws, r, c_end, close_here)
            for c in range(1, NC + 1):
                ws.cell(r, c).border = SHELL_BORDER
            carry = close_here
            final_close = close_here
            for col in range(c0, c_end):
                sum_cols[col] += row_vals.get(col, 0.0)
            sum_cols[3] += open_here
            sum_cols[c_end] += close_here
            r += 1
        # 全年合计行
        _txt(ws, r, 1, ent)
        _txt(ws, r, 2, '全年合计', bold=True, align=SHELL_CEN)
        _money(ws, r, 3, first_open)
        for c in range(c0, c_end):
            _money(ws, r, c, sum_cols[c])
        _money(ws, r, c_end, final_close)
        for c in range(1, NC + 1):
            ws.cell(r, c).border = SHELL_BORDER
            if c == 2:
                ws.cell(r, c).font = SHELL_BOLD
        r += 1
        # 「其他」构成块（按对方科目打开）
        oi = _col_of('其他增加', inc_cols, c0)
        od = _col_of('其他减少', dec_cols, c_dec)
        oi_agg, od_agg = {}, {}
        for ym, m2 in d['other_inc'].items():
            if year_filter is None or ym[0] == int(year):
                for cp, amt in m2.items():
                    oi_agg[cp] = oi_agg.get(cp, 0.0) + amt
        for ym, m2 in d['other_dec'].items():
            if year_filter is None or ym[0] == int(year):
                for cp, amt in m2.items():
                    od_agg[cp] = od_agg.get(cp, 0.0) + amt
        if oi_agg or od_agg:
            _txt(ws, r, 2, '其他构成', bold=True, align=SHELL_CEN)
            for c in range(1, NC + 1):
                ws.cell(r, c).border = SHELL_BORDER
                ws.cell(r, c).fill = SHELL_TOT_FILL
            r += 1
            for cp, amt in sorted(oi_agg.items(), key=lambda x: -abs(x[1])):
                _txt(ws, r, 2, f'　其他增加-{cp}', italic=True)
                if oi is not None:
                    _money(ws, r, oi, amt)
                for c in range(1, NC + 1):
                    ws.cell(r, c).border = SHELL_BORDER
                r += 1
            for cp, amt in sorted(od_agg.items(), key=lambda x: -abs(x[1])):
                _txt(ws, r, 2, f'　其他减少-{cp}', italic=True)
                if od is not None:
                    _money(ws, r, od, amt)
                for c in range(1, NC + 1):
                    ws.cell(r, c).border = SHELL_BORDER
                r += 1
        # 2026-08-03 用户要求：勾稽行删除
        r += 1
    widths = [16, 14, 16] + [13] * (NC - 3)
    for c, w in zip(range(1, NC + 1), widths):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A6'
    return ws


def build_inv_monthly_detail(ws, ents, title, acc_code, acc_name, targets, contra, year_filter=None):
    """第二张分月表：按借贷方打开，逐主体·科目·分月·借贷 列示对应科目(语义类别)与金额。
    长表格式：主体 | 科目 | 月份 | 方向 | 对应科目 | 金额；每(主体,科目,月份)先列借方(增加)
    各对应科目，再列贷方(减少)各对应科目，并附借/贷小计。仅列示有发生额的组合。
    2026-08-03 起已停用（main 不再调用）；保留函数以防回退。"""
    agg = _agg_inv_account(ents, acc_code, acc_name, targets, contra)
    ws.cell(1, 1, title).font = SHELL_TITLE_FONT
    headers = ['主体', '科目', '月份', '方向', '对应科目', '金额']
    hr = _hdr(ws, 4, headers)
    r = hr
    for (ent, year), d in agg.items():
        if year_filter is not None and str(year) != str(year_filter):
            continue
        if not d['months']:
            continue
        for ym in sorted(d['months']):
            for cat in d['inc_cats']:
                v = d['inc'][cat].get(ym, 0.0)
                if abs(v) > 1e-6:
                    _txt(ws, r, 1, ent)
                    _txt(ws, r, 2, f'{year}')
                    _txt(ws, r, 3, f'{ym[0]}-{ym[1]:02d}', align=SHELL_CEN)
                    _txt(ws, r, 4, '借(增加)', align=SHELL_CEN)
                    _txt(ws, r, 5, cat)
                    _money(ws, r, 6, v)
                    for c in range(1, 7):
                        ws.cell(r, c).border = SHELL_BORDER
                    r += 1
            for cat in d['dec_cats']:
                v = d['dec'][cat].get(ym, 0.0)
                if abs(v) > 1e-6:
                    _txt(ws, r, 1, ent)
                    _txt(ws, r, 2, f'{year}')
                    _txt(ws, r, 3, f'{ym[0]}-{ym[1]:02d}', align=SHELL_CEN)
                    _txt(ws, r, 4, '贷(减少)', align=SHELL_CEN)
                    _txt(ws, r, 5, cat)
                    _money(ws, r, 6, v)
                    for c in range(1, 7):
                        ws.cell(r, c).border = SHELL_BORDER
                    r += 1
    for c, w in zip(range(1, 7), [16, 8, 12, 10, 26, 16]):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A5'
    return ws


def build_prod_monthly_detail(ws, ents, title, year_filter=None):
    """生产成本 对应科目分月明细表：借方(增加)=成本要素投入，贷方(减少)=转出类别，逐主体·月·类别展开。"""
    agg = _agg_prod_account(ents)
    ws.cell(1, 1, title).font = SHELL_TITLE_FONT
    headers = ['主体', '年度', '月份', '方向', '对应科目', '金额']
    hr = _hdr(ws, 4, headers)
    r = hr
    for (ent, year), d in agg.items():
        if year_filter is not None and str(year) != str(year_filter):
            continue
        if not d['months']:
            continue
        for ym in sorted(d['months']):
            for cat in PROD_INC:
                v = d['inc'][cat].get(ym, 0.0)
                if abs(v) > 1e-6:
                    _txt(ws, r, 1, ent)
                    _txt(ws, r, 2, f'{year}')
                    _txt(ws, r, 3, f'{ym[0]}-{ym[1]:02d}', align=SHELL_CEN)
                    _txt(ws, r, 4, '借(增加)', align=SHELL_CEN)
                    _txt(ws, r, 5, cat)
                    _money(ws, r, 6, v)
                    for c in range(1, 7):
                        ws.cell(r, c).border = SHELL_BORDER
                    r += 1
            for cat in PROD_DEC:
                v = d['dec'][cat].get(ym, 0.0)
                if abs(v) > 1e-6:
                    _txt(ws, r, 1, ent)
                    _txt(ws, r, 2, f'{year}')
                    _txt(ws, r, 3, f'{ym[0]}-{ym[1]:02d}', align=SHELL_CEN)
                    _txt(ws, r, 4, '贷(减少)', align=SHELL_CEN)
                    _txt(ws, r, 5, cat)
                    _money(ws, r, 6, v)
                    for c in range(1, 7):
                        ws.cell(r, c).border = SHELL_BORDER
                    r += 1
    for c, w in zip(range(1, 7), [16, 8, 12, 10, 26, 16]):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A5'
    return ws


def build_inventory_footnote(wb, ents, year, data_dir=None):
    """存货附注汇总（2026-08-03 用户方法论：参照上市公司存货附注结构）：
    类别×主体（原材料/库存商品/发出商品/周转材料/半成品/在途物资 动态），
    块=一期末账面余额/二存货跌价准备(正数)/三期末账面价值(余额-跌价)/
    四期初账面余额/五审计调整数/六期末审定数(账面余额=未审+调整)/
    七期末审定账面价值(=审定数−跌价)，每主体一列+全集团合计；
    八存货跌价准备变动表（集团口径：期初/计提/转回及转销/期末）。
    2026-08-03 修复：_blk 返回三元组 (下一起始行, 合计行, 类别首行)，供审定数/净值块公式引用。"""
    from audit_common import discover_entities as _acd, read_tb_full as _rtb
    ws = wb.create_sheet('存货附注汇总')
    # ⚡ 2026-08-10 SAP：按实际科目名称分类（1406 库存产品→库存商品/产成品/半成品/在制品；
    # 1411=委托加工物资；1404/1407 成本差异）——cats 换成名称关键词，_bal 走名称匹配分支
    if _is_sap():
        cats = [('原材料', ['原材料']), ('库存商品', ['库存商品']), ('产成品', ['产成品']),
                ('半成品', ['半成品']), ('在制品', ['在制品', '在产品']), ('发出商品', ['发出商品']),
                ('委托加工物资', ['委托加工物资']), ('周转材料', ['周转材料', '包装物', '低值易耗品']),
                ('材料成本差异', ['材料成本差异', '成本差异']), ('在途物资', ['在途'])]
    else:
        cats = [('原材料', ['1403']), ('库存商品', ['1405']), ('发出商品', ['1406']),
                ('委托加工物资', ['1408']), ('周转材料', ['1411']), ('包装物', ['1412']),
                ('低值易耗品', ['1413']), ('半成品', ['1404']), ('在途物资', ['1402'])]
    # 2026-08-03 修复：原 getattr(ents,'_dir') 因 ents 为 dict 无该属性→传空目录→tb 空→附注全 0。
    # 改为显式接收 data_dir（process_folder 传入）。
    ac_ents = _acd(data_dir or '')
    tb = _rtb(data_dir or '', ac_ents)
    # 简化：直接按主体 TB 取各存货代码期末/期初，泰国 1140 映射
    ent_list = sorted(ents)
    # 各主体存货余额
    def _bal(ent, codes, field):
        # ⚡ 2026-08-10 SAP 分支：codes=名称关键词列表 → 按 TB 科目名匹配（1406 库存产品
        #   子目按名称归入库存商品/产成品/半成品/在制品；成本差异单列）。末级叶子过滤防双计。
        if _is_sap():
            v = 0.0
            # ⚡ 2026-08-26 排除跌价(1471)与材料成本差异(1404/1407)：名称含类别词会误入
            #   （'库存产品成本差异-产成品' 含'产成品'→产成品双计材料成本差异）。
            #   材料成本差异类别自身（codes 含'成本差异'）不排除，独立取数。
            _cm = ' '.join(codes or [])
            _excl = ('1471',) if '成本差异' in _cm else ('1471', '1404', '1407')
            _hits = [(str(k[1]), float(r.get(field, 0.0) or 0.0))
                     for k, r in tb.items()
                     if k[0] == ent and str(k[3]) == str(year)
                     and not str(k[1]).startswith(_excl)
                     and any(kw in str(k[2] or '') for kw in codes)]
            _leafs = [a for a in _hits if not any(b != a and b[0].startswith(a[0]) for b in _hits)]
            for _c, _a in _leafs:
                v += _a   # ⚡ 2026-08-26 去 abs：保持 TB 真实余额符号（周转材料 -20032 不能变正）
            return v
        # 2026-08-06 泰国兜底：泰国代码未归一化（1140.01 原材料等），标准码精确匹配落空 →
        # 泰国主体自动映射（1403→1140.01、1405→1140.02、1406→1140.04），全部块统一生效
        # ⚡ 2026-08-10 SAP 修复：SAP TB 是 10 位码（1403000000 原材料），原 `k[1] in codes`
        #   精确匹配标准 4 位码永不命中 → 附注汇总全 0（1010 存货审定表 11.7亿 vs 附注 0）。
        #   改【代码前缀匹配 + 末级叶子过滤】（1403 前缀命中 1403000000/1403010000…，
        #   父行(有后代行)剔除只计叶子，防父+子双计）。U8 4 位码 startswith 兼容。
        v = 0.0
        _mcodes = codes
        if '泰国' in str(ent):
            _mcodes = [THAI_INV_MAP.get(c, c) for c in codes]
        _hits = [(str(k[1]), float(r.get(field, 0.0) or 0.0))
                 for k, r in tb.items()
                 if k[0] == ent and str(k[3]) == str(year)
                 and any(str(k[1]).startswith(mc) for mc in _mcodes)]
        # 末级叶子过滤：有后代行（更长同前缀）的父行剔除
        _leafs = [a for a in _hits if not any(b != a and b[0].startswith(a[0]) for b in _hits)]
        for _c, _a in _leafs:
            v += _a   # ⚡ 2026-08-26 去 abs：保持 TB 真实余额符号
        return v
    # ⚡ 2026-08-10 该类别跌价准备（附注二/五/九块用）：SAP 跌价子目按名称含『类别词』细分
    # （1471020000 跌价准备-原材料、1471030100 跌价准备-库存产品-库存商品、
    #   1471030200 跌价准备-库存产品-产成品、1471030300 跌价准备-库存产品-半成品）；
    # 原二块用 ['跌价准备'] 每行返回总数（6513 万×10 重复计 10 倍）。U8/泰国：1471 全量。
    def _bal_bd(ent, kw, field):
        if _is_sap():
            v = 0.0
            for k, r in tb.items():
                if k[0] != ent or str(k[3]) != str(year):
                    continue
                n = str(k[2] or '')
                if '跌价准备' not in n:
                    continue
                if kw and kw not in n:
                    continue
                v += abs(float(r.get(field, 0.0) or 0.0))
            return v
        # ⚡ 2026-08-26 修复：U8/泰国/集团模式（current_comp 为空）原 `return _bal(ent,['1471'])`
        #   忽略类别 → 每个科目都填 1471 全量（1010 附注 65M×9 重复计 9 倍）。
        #   改为：按名称关键词匹配 1471 子目（1471xx 跌价准备-类别）拆分；
        #   仅当 TB 无 1471 子目（只有 1471 父级）才回退全量。跌价准备按披露取正数。
        _hits = [(str(k[1]), str(k[2] or ''), float(r.get(field, 0.0) or 0.0))
                 for k, r in tb.items()
                 if k[0] == ent and str(k[3]) == str(year)
                 and str(k[1]).startswith('1471')]
        if not _hits:
            return 0.0
        _leafs = [h for h in _hits if not any(h2[0] != h[0] and h2[0].startswith(h[0]) for h2 in _hits)]
        _has_sub = any(c != '1471' for c, _, _ in _leafs)
        if _has_sub:
            if not kw:
                return 0.0
            return sum(abs(a) for _c, _n, a in _leafs if kw in _n)
        return abs(sum(abs(a) for _c, _n, a in _leafs))
    # 泰国兜底
    def _bal_thai(ent, thai_code, field):
        for k, r in tb.items():
            if k[0] == ent and k[1] == thai_code and str(k[3]) == str(year):
                return abs(float(r.get(field, 0.0) or 0.0))
        return 0.0
    # 2026-08-03：压缩"前三列空白"——主体列紧接『项目』列（原 B-D 三列空占位）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出『全集团合计』列
    _single_inv = len(ent_list) <= 1
    NC = 2 + len(ent_list) - (1 if _single_inv else 0)
    ws.cell(1, 1, f'存货附注汇总（{year}年）').font = SHELL_TITLE_FONT
    # 2026-08-03 修复：块标题原用 SHELL_HFILL(深蓝底 1F4E78)+黑字 看不清 → 改浅蓝底 DDEBF7（表头白字保留深蓝）
    _BLK = PatternFill('solid', fgColor='DDEBF7')
    # 表头
    hdr = ['项目'] + ent_list + (['全集团合计'] if not _single_inv else [])
    for i, h in enumerate(hdr, 1):
        cc = ws.cell(4, i, h); cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL
        cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    r = 5
    def _blk(r, title, getter):
        ws.cell(r, 1, title).font = SHELL_BOLD
        ws.cell(r, 1).fill = _BLK; ws.cell(r, 1).border = SHELL_BORDER
        r += 1
        first = r
        for cname, codes in cats:
            vals = []
            for ent in ent_list:
                v = getter(ent, codes, cname)
                if cname == '原材料' and v == 0:
                    v = _bal_thai(ent, '1140.01', 'qm' if '期末' in title else 'qc')
                if cname == '库存商品' and v == 0:
                    v = _bal_thai(ent, '1140.02', 'qm' if '期末' in title else 'qc')
                if cname == '发出商品' and v == 0:
                    v = _bal_thai(ent, '1140.04', 'qm' if '期末' in title else 'qc')
                vals.append(v)
            _txt(ws, r, 1, cname)
            for i, v in enumerate(vals):
                _money(ws, r, 2 + i, v)
            if not _single_inv:
                _money(ws, r, NC, sum(vals))
            for c in range(1, NC + 1):
                ws.cell(r, c).border = SHELL_BORDER
            r += 1
        # 小计（2026-08-06 协议：写数值=类别行各列和；公式 data_only 读 None 收回读不回）
        _txt(ws, r, 1, '合  计', bold=True, align=SHELL_CEN)
        for c in range(2, NC + 1):
            _sv = sum(ws.cell(x, c).value or 0.0 for x in range(first, r)
                      if isinstance(ws.cell(x, c).value, (int, float)))
            ws.cell(r, c, round(_sv, 2))
            ws.cell(r, c).number_format = SHELL_NUM
            ws.cell(r, c).alignment = SHELL_RGT
            ws.cell(r, c).border = SHELL_BORDER
        for c in range(1, NC + 1):
            ws.cell(r, c).fill = SHELL_TOT_FILL
            ws.cell(r, c).font = SHELL_BOLD
        r += 1
        return r, r - 1, first   # (下一起始行, 合计行, 类别首行)
    # 一 期末账面余额
    r, _tot_end, _end_first = _blk(r, '一、期末账面余额', lambda ent, codes, cname: _bal(ent, codes, 'qm'))
    # 二 存货跌价准备（正数）
    r, _tot_bd, _bd_first = _blk(r, '二、存货跌价准备（备抵，正数）', lambda ent, codes, cname: _bal_bd(ent, cname, 'qm'))
    # 三 期末账面价值 = 余额 - 跌价
    ws.cell(r, 1, '三、期末账面价值').font = SHELL_BOLD
    ws.cell(r, 1).fill = _BLK; ws.cell(r, 1).border = SHELL_BORDER
    r += 1
    for cname, codes in cats:
        _txt(ws, r, 1, cname)
        for i, ent in enumerate(ent_list):
            v = _bal(ent, codes, 'qm')
            # 2026-08-06 泰国兜底：泰国代码未归一化，_bal 标准码落空 → 用泰国子目码
            # （与『一、期末账面余额』块一致，否则 三、期末账面价值 全 0）
            if v == 0 and '泰国' in str(ent):
                _tc = THAI_INV_MAP.get(codes[0]) if codes else None
                if _tc:
                    v = _bal_thai(ent, _tc, 'qm')
            bd = _bal_bd(ent, cname, 'qm')
            _money(ws, r, 2 + i, v - bd)
        for c in range(1, NC + 1):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    _txt(ws, r, 1, '合  计', bold=True, align=SHELL_CEN)
    for c in range(2, NC + 1):
        # 2026-08-06 协议：合计写数值
        _sv = sum(ws.cell(x, c).value or 0.0 for x in range(r - len(cats), r)
                  if isinstance(ws.cell(x, c).value, (int, float)))
        ws.cell(r, c, round(_sv, 2))
        ws.cell(r, c).number_format = SHELL_NUM
        ws.cell(r, c).alignment = SHELL_RGT
        ws.cell(r, c).border = SHELL_BORDER
    for c in range(1, NC + 1):
        ws.cell(r, c).fill = SHELL_TOT_FILL
        ws.cell(r, c).font = SHELL_BOLD
    r += 1
    # 四 期初账面余额
    r, _q_start, _q_first = _blk(r, '四、期初账面余额', lambda ent, codes, cname: _bal(ent, codes, 'qc'))
    # —— 2026-08-10 用户要求：期初账面余额后补 期初跌价准备 / 期初账面价值（FY 模板 8 块外补充）——
    # 五 期初跌价准备（按类别）
    r, _qbd_first, _ = _blk(r, '五、期初存货跌价准备（备抵，正数）',
                            lambda ent, codes, cname: _bal_bd(ent, cname, 'qc'))
    # 六 期初账面价值 = 期初余额 − 期初跌价准备
    ws.cell(r, 1, '六、期初账面价值（=期初余额−跌价准备）').font = SHELL_BOLD
    ws.cell(r, 1).fill = _BLK; ws.cell(r, 1).border = SHELL_BORDER
    r += 1
    for i, (cname, codes) in enumerate(cats):
        _txt(ws, r, 1, cname)
        for j in range(len(ent_list)):
            _qv = ws.cell(_q_first + i, 2 + j).value if isinstance(ws.cell(_q_first + i, 2 + j).value, (int, float)) else 0.0
            _qb = _bal_bd(ent_list[j], cname, 'qc')
            ws.cell(r, 2 + j, round(_qv - _qb, 2))
            ws.cell(r, 2 + j).number_format = SHELL_NUM
            ws.cell(r, 2 + j).alignment = SHELL_RGT
        ws.cell(r, NC, round(sum(ws.cell(r, 2 + j).value or 0.0 for j in range(len(ent_list))), 2))
        ws.cell(r, NC).number_format = SHELL_NUM
        ws.cell(r, NC).alignment = SHELL_RGT
        for c in range(1, NC + 1):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    _txt(ws, r, 1, '合  计', bold=True, align=SHELL_CEN)
    for c in range(2, NC + 1):
        _sv = sum(ws.cell(x, c).value or 0.0 for x in range(r - len(cats), r)
                  if isinstance(ws.cell(x, c).value, (int, float)))
        ws.cell(r, c, round(_sv, 2))
        ws.cell(r, c).number_format = SHELL_NUM
        ws.cell(r, c).alignment = SHELL_RGT
        ws.cell(r, c).border = SHELL_BORDER
    for c in range(1, NC + 1):
        ws.cell(r, c).fill = SHELL_TOT_FILL
        ws.cell(r, c).font = SHELL_BOLD
    r += 1
    # —— 2026-08-03 五~七：审计调整数 / 期末审定数 / 期末审定账面价值（参照固定资产模板 未审→调整→审定）——
    # 五 审计调整数（期末账面余额调整，审计人员填写）
    ws.cell(r, 1, '七、审计调整数（期末账面余额）').font = SHELL_BOLD
    ws.cell(r, 1).fill = _BLK; ws.cell(r, 1).border = SHELL_BORDER
    r += 1
    _adj_first = r
    for cname, codes in cats:
        _txt(ws, r, 1, cname)
        for i in range(len(ent_list)):
            _money(ws, r, 2 + i, 0.0)
        _money(ws, r, NC, 0.0)
        for c in range(1, NC + 1):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    _txt(ws, r, 1, '合  计', bold=True, align=SHELL_CEN)
    for c in range(2, NC + 1):
        # 2026-08-06 协议：调整块合计写 0（类别行全 0 待填）
        ws.cell(r, c, 0.0)
        ws.cell(r, c).number_format = SHELL_NUM
        ws.cell(r, c).alignment = SHELL_RGT
        ws.cell(r, c).border = SHELL_BORDER
    for c in range(1, NC + 1):
        ws.cell(r, c).fill = SHELL_TOT_FILL
        ws.cell(r, c).font = SHELL_BOLD
    r += 1
    # 六 期末审定数（账面余额）= 期末未审 + 审计调整（2026-08-06 协议：写数值=期末未审，调整 0 待填）
    ws.cell(r, 1, '八、期末审定数（账面余额）').font = SHELL_BOLD
    ws.cell(r, 1).fill = _BLK; ws.cell(r, 1).border = SHELL_BORDER
    r += 1
    _aud_first = r
    for i, (cname, codes) in enumerate(cats):
        _txt(ws, r, 1, cname)
        _row_vals = []
        for j in range(len(ent_list)):
            _v = ws.cell(_end_first + i, 2 + j).value if isinstance(ws.cell(_end_first + i, 2 + j).value, (int, float)) else 0.0
            ws.cell(r, 2 + j, round(_v, 2))
            ws.cell(r, 2 + j).number_format = SHELL_NUM
            ws.cell(r, 2 + j).alignment = SHELL_RGT
            _row_vals.append(_v)
        ws.cell(r, NC, round(sum(_row_vals), 2))
        ws.cell(r, NC).number_format = SHELL_NUM
        ws.cell(r, NC).alignment = SHELL_RGT
        for c in range(1, NC + 1):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    _txt(ws, r, 1, '合  计', bold=True, align=SHELL_CEN)
    for c in range(2, NC + 1):
        _sv = sum(ws.cell(x, c).value or 0.0 for x in range(_aud_first, r)
                  if isinstance(ws.cell(x, c).value, (int, float)))
        ws.cell(r, c, round(_sv, 2))
        ws.cell(r, c).number_format = SHELL_NUM
        ws.cell(r, c).alignment = SHELL_RGT
        ws.cell(r, c).border = SHELL_BORDER
    for c in range(1, NC + 1):
        ws.cell(r, c).fill = SHELL_TOT_FILL
        ws.cell(r, c).font = SHELL_BOLD
    r += 1
    # 七 期末审定账面价值 = 审定数 − 存货跌价准备（未审，原材料扣跌价）（2026-08-06 协议：写数值）
    # 九 期末审定跌价准备（按类别；2026-08-10 用户要求：期末审定数后补审定跌价准备）
    r, _audbd_first, _ = _blk(r, '九、期末审定存货跌价准备（备抵，正数）',
                              lambda ent, codes, cname: _bal_bd(ent, cname, 'qm'))
    ws.cell(r, 1, '十、期末审定账面价值（=审定数−跌价准备）').font = SHELL_BOLD
    ws.cell(r, 1).fill = _BLK; ws.cell(r, 1).border = SHELL_BORDER
    r += 1
    _net_first = r
    for i, (cname, codes) in enumerate(cats):
        _txt(ws, r, 1, cname)
        _row_vals = []
        for j in range(len(ent_list)):
            _av = ws.cell(_aud_first + i, 2 + j).value if isinstance(ws.cell(_aud_first + i, 2 + j).value, (int, float)) else 0.0
            _bv = ws.cell(_bd_first + i, 2 + j).value if isinstance(ws.cell(_bd_first + i, 2 + j).value, (int, float)) else 0.0
            _v = _av - _bv
            ws.cell(r, 2 + j, round(_v, 2))
            ws.cell(r, 2 + j).number_format = SHELL_NUM
            ws.cell(r, 2 + j).alignment = SHELL_RGT
            _row_vals.append(_v)
        ws.cell(r, NC, round(sum(_row_vals), 2))
        ws.cell(r, NC).number_format = SHELL_NUM
        ws.cell(r, NC).alignment = SHELL_RGT
        for c in range(1, NC + 1):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    _txt(ws, r, 1, '合  计', bold=True, align=SHELL_CEN)
    for c in range(2, NC + 1):
        _sv = sum(ws.cell(x, c).value or 0.0 for x in range(_net_first, r)
                  if isinstance(ws.cell(x, c).value, (int, float)))
        ws.cell(r, c, round(_sv, 2))
        ws.cell(r, c).number_format = SHELL_NUM
        ws.cell(r, c).alignment = SHELL_RGT
        ws.cell(r, c).border = SHELL_BORDER
    for c in range(1, NC + 1):
        ws.cell(r, c).fill = SHELL_TOT_FILL
        ws.cell(r, c).font = SHELL_BOLD
    r += 1
    # —— 2026-08-03 八、存货跌价准备变动表（集团口径：期初/计提/转回及转销/期末）——
    try:
        _agg7 = _agg_inv_account(ents, '1471', '存货跌价准备', ['存货跌价准备'], True)
        _op7 = _inc7 = _dec7 = _cl7 = 0.0
        for (_e, _y), _d in _agg7.items():
            if str(_y) != str(year):
                continue
            _op7 += float(_d.get('km_op') or 0.0)
            _inc7 += float(_d.get('dec_t', {}).get('计提', 0.0) or 0.0)
            _dec7 += float(_d.get('inc_t', {}).get('转回', 0.0) or 0.0)
            _cl7 += float(_d.get('km_cl') or 0.0)
        r += 1
        ws.cell(r, 1, '十一、存货跌价准备变动表（集团口径）').font = SHELL_BOLD
        ws.cell(r, 1).fill = _BLK; ws.cell(r, 1).border = SHELL_BORDER
        r += 1
        for j, h in enumerate(['项目', '存货跌价准备'], 1):
            cc = ws.cell(r, j, h); cc.font = SHELL_BOLD
            cc.fill = SHELL_HFILL; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
        r += 1
        for lbl, val in [('期初余额', _op7), ('本期计提（贷方）', _inc7),
                         ('本期转回及转销（借方）', _dec7), ('期末余额', _cl7)]:
            _txt(ws, r, 1, lbl)
            _money(ws, r, 2, round(val, 2))
            for c in range(1, 3):
                ws.cell(r, c).border = SHELL_BORDER
            r += 1
        # 2026-08-04 用户要求：勾稽行删除（勾稽=期初+转回-计提-期末，对上了无需列示）
        r += 0
    except Exception as _ex:
        print(f'  [WARN] 存货附注-七变动表失败：{_ex}')
    ws.column_dimensions['A'].width = 26
    for c in range(2, NC + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.freeze_panes = 'B5'
    return ws


# ----------------------------------------------------------------------------
# 按指令号（WBS/订单）全年增减变动表（SAP 专属，2026-08-10 用户方法论）
# ----------------------------------------------------------------------------

def build_inv_instruction_sheet(wb, ents, year):
    """存货按指令号（WBS/订单）全年增减变动表。
    用户方法论：存货按指令号全年增减变动——有指令号且有发生额的都要有，剔除虚增虚减
    （同指令号同科目借=贷对冲行），与主营业务成本倒退核对。
    数据源：GL 行 wbs(col17)/ord(col24)/mat(col25)/code（read_gl_rows 扩展列）。
    行=指令号×物料×科目；列=主体|WBS|订单|物料|科目码|科目名|本期增加(借)|本期减少(贷)|净变动。
    对冲剔除：同键 借>0 且 贷>0 且 |借-贷|<=0.01 → 不列入（备注统计剔除额/行数）。
    SAP 专属（U8 无指令号列）→ 非 SAP 或集团模式返回空表。"""
    ws = wb.create_sheet('按指令号_全年增减变动表')
    ws.cell(1, 1, f'存货按指令号（WBS/订单）全年增减变动表（{year} 年度；单位：元；已剔除同指令号借贷对冲行）').font = SHELL_TITLE_FONT
    _hdr(ws, 3, ['主体', 'WBS', '订单', '物料', '科目码', '科目名', '本期增加(借)', '本期减少(贷)', '净变动', '备注'])
    _INV = ('1403', '1406', '1411', '1404', '5001')   # 存货原值科目（不含 1471 跌价）
    pairs = list(_iter_comp_gl(ents, year))          # ⚡ 2026-08-26 集团模式逐主体
    if not pairs:
        _txt(ws, 4, 1, '（无 GL 数据）')
        return ws
    r = 4
    sum_d = sum_c = 0.0
    n_hit = n_drop = 0
    drop_amt = 0.0
    first = True
    for comp, gl in pairs:
        if not first:
            r += 1                                     # 主体间空行
        first = False
        agg = {}
        for _rr in gl:
            code = str(_rr.get('code') or '')
            if not code.startswith(_INV):
                continue
            w = str(_rr.get('wbs') or '').strip()
            o = str(_rr.get('ord') or '').strip()
            mat = str(_rr.get('mat') or '').strip()
            # ⚡ 2026-08-10 300 适配：FBL3N 无 WBS/订单列数据（订单列空、项目列=行项目序号）
            #   → 按【物料号】降级聚合（wbs/ord/mat 任一非空即纳入）；yy 有 WBS 用 WBS。
            if not w and not o and not mat:
                continue                                # 无任何指令号行不纳入
            k = (w, o, mat[:12], code[:10], str(_rr.get('name') or '')[:30])
            d = float(_rr.get('debit') or 0.0)
            c = float(_rr.get('credit') or 0.0)
            a = agg.get(k)
            if a is None:
                agg[k] = [d, c]
            else:
                a[0] += d; a[1] += c
        # 排序：净变动绝对值降序（大额在前便于审计甄别）
        items = sorted(agg.items(), key=lambda kv: -(abs(kv[1][0] - kv[1][1]) + kv[1][0] + kv[1][1]))
        for (w, o, mat, code, nm), (d, c) in items:
            if d > 0.005 and c > 0.005 and abs(d - c) <= 0.01:
                n_drop += 1
                drop_amt += d
                continue                                # 虚增虚减对冲行剔除
            _txt(ws, r, 1, comp)
            _txt(ws, r, 2, w)
            _txt(ws, r, 3, o)
            _txt(ws, r, 4, mat)
            _txt(ws, r, 5, code)
            _txt(ws, r, 6, nm)
            _money(ws, r, 7, d)
            _money(ws, r, 8, c)
            _money(ws, r, 9, d - c)
            _txt(ws, r, 10, '借贷对冲' if (d > 0.005 and c > 0.005) else ('仅借方' if c <= 0.005 else '仅贷方'))
            n_hit += 1
            sum_d += d; sum_c += c
            r += 1
    if n_hit:
        _txt(ws, r, 1, '合计', bold=True)
        _money(ws, r, 7, sum_d)
        _money(ws, r, 8, sum_c)
        _money(ws, r, 9, sum_d - sum_c)
        for c in range(1, 10):
            ws.cell(r, c).font = SHELL_BOLD
            ws.cell(r, c).fill = SHELL_TOT_FILL
        r += 1


def build_inv_cogs_recon_sheet(wb, ents, year):
    """与主营业务成本倒退核对表（SAP 专属，2026-08-16 #702 用户方法论：存货按指令号
    全年增减变动后与主营业务成本倒退核对）。
    逻辑：6401 主营业务成本借方 应 = 存货（1403/1406/1411/1404/5001）贷方结转主营成本。
    按指令号（WBS 优先，300 用订单/物料）双向对齐：
      差异=0 → 对应正常；差异>0 → 主营成本有发生但存货无对应结转（直接采购/暂估/人工等
      非存货来源，列对方科目归因）；差异<0 → 存货结转未全部进主营成本（异常关注）。
    附：6401 借方对方科目 Top 归因（解释差异来源）。"""
    ws = wb.create_sheet('与主营业务成本_倒退核对表')
    ws.cell(1, 1, f'存货与主营业务成本倒退核对（{year} 年度；单位：元；按指令号 WBS/订单/物料双向对齐；'
                  f'差异=6401借方−存货结转主营成本）').font = SHELL_TITLE_FONT
    _hdr(ws, 3, ['主体', '指令号', '存货结转主营成本(贷)', '6401主营成本(借)', '差异', '差异方向', '差异对方科目归因'])
    _INV = ('1403', '1406', '1411', '1404', '5001')
    pairs = list(_iter_comp_gl(ents, year))          # ⚡ 2026-08-26 集团模式逐主体
    if not pairs:
        _txt(ws, 4, 1, '（无 GL 数据）')
        return ws
    r = 4
    sum_inv = sum_cog = 0.0
    n_hit = n_abn = 0
    cp_agg = {}      # 6401 对方科目 Top 归因（全主体合并）
    first = True
    for comp, gl in pairs:
        if not first:
            r += 1                                     # 主体间空行
        first = False
        def _inst(rr):
            w = str(rr.get('wbs') or '').strip()
            o = str(rr.get('ord') or '').strip()
            m = str(rr.get('mat') or '').strip()
            return w or o or m or '（无指令号）'
        inv_cogs = {}    # 指令号 -> (贷方结转主营成本, 对方科目样本)
        cogs_6401 = {}   # 指令号 -> (6401借方, 对方科目Top)
        for rr in gl:
            code = str(rr.get('code') or '')
            d = float(rr.get('debit') or 0.0)
            c = float(rr.get('credit') or 0.0)
            cp = str(rr.get('cp') or '')
            if code.startswith(_INV) and c > 0.005:
                if '主营' not in cp:
                    continue
                k = _inst(rr)
                if k == '（无指令号）':
                    continue
                a = inv_cogs.get(k)
                if a is None:
                    inv_cogs[k] = [c, cp[:40]]
                else:
                    a[0] += c
            elif code.startswith('6401') and d > 0.005:
                k = _inst(rr)
                a = cogs_6401.get(k)
                if a is None:
                    cogs_6401[k] = [d, cp[:40]]
                else:
                    a[0] += d
                # 对方科目归因（全量，非仅差异组）
                key = cp[:24] or '（无）'
                x = cp_agg.get(key)
                if x is None:
                    cp_agg[key] = [d, 1]
                else:
                    x[0] += d; x[1] += 1
        # 双向并集行
        keys = set(inv_cogs) | set(cogs_6401)
        rows_out = []
        for k in keys:
            iv = inv_cogs.get(k, [0.0, ''])
            cg = cogs_6401.get(k, [0.0, ''])
            diff = cg[0] - iv[0]
            if abs(diff) < 0.005 and iv[0] < 0.005:
                continue
            dirn = '正常对应' if abs(diff) <= 0.005 else ('成本无存货结转' if diff > 0 else '存货结转未进成本')
            attr = (cg[1] or iv[1] or '')
            rows_out.append((k, iv[0], cg[0], diff, dirn, attr))
        rows_out.sort(key=lambda t: -abs(t[3]))
        for k, iv, cg, diff, dirn, attr in rows_out:
            _txt(ws, r, 1, comp)
            _txt(ws, r, 2, k)
            _money(ws, r, 3, iv)
            _money(ws, r, 4, cg)
            _money(ws, r, 5, diff)
            jc = _txt(ws, r, 6, dirn, align=SHELL_CEN)
            if dirn != '正常对应':
                jc.font = Font(name='Times New Roman', size=10, bold=True,
                               color='C00000' if diff > 0 else 'BF8F00')
                n_abn += 1
            _txt(ws, r, 7, attr)
            sum_inv += iv; sum_cog += cg
            n_hit += 1
            r += 1
    if n_hit:
        _txt(ws, r, 1, '合计', bold=True)
        _money(ws, r, 3, sum_inv)
        _money(ws, r, 4, sum_cog)
        _money(ws, r, 5, sum_cog - sum_inv)
        for c in range(1, 7):
            ws.cell(r, c).font = SHELL_BOLD
            ws.cell(r, c).fill = SHELL_TOT_FILL
        r += 1
    # —— 6401 借方对方科目 Top 归因 ——
    r += 1
    _txt(ws, r, 1, '6401 主营业务成本借方·对方科目归因（差异来源解释；按金额降序）：', bold=True)
    r += 1
    for cp_k, (amt, n) in sorted(cp_agg.items(), key=lambda kv: -kv[1][0])[:15]:
        _txt(ws, r, 1, cp_k)
        _txt(ws, r, 2, f'{n}行', align=SHELL_RGT)
        _money(ws, r, 3, amt)
        r += 1
    r += 1
    _txt(ws, r, 1, f'说明：指令号分组 {n_hit} 个；差异组 {n_abn} 个。'
                  f'存货结转主营成本合计 {format(sum_inv, ",.2f")} 元；6401 借方合计 {format(sum_cog, ",.2f")} 元；'
                  f'净差 {format(sum_cog - sum_inv, ",.2f")} 元——为 非存货来源成本'
                  f'（直接采购确认/暂估应付款/预付/关税/人工等，见对方科目归因）；'
                  f'『存货结转未进成本』组为异常，需审计关注。', italic=True)
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 20
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 44
    ws.freeze_panes = 'B4'
    return ws


def _instr_year(s):
    """指令号年份推断（库龄分析用）：F-GF25.1150.12→2025、2511→2025、2414→2024、2611→2026。
    对指令号各段找 4 位全年份(19xx/20xx)或 2 位年份(20-26 范围)，取【最小】年份
    （最早=库龄最长，保守判断——用户方法论：指令号早→库龄可能长）。"""
    if not s:
        return None
    years = []
    for seg in re.split(r'[.\-/_]', str(s)):
        m = re.search(r'((?:19|20)\d{2})', seg)
        if m:
            years.append(int(m.group(1)))
            continue
        m = re.search(r'(\d{2})', seg)
        if m:
            yy = int(m.group(1))
            if 20 <= yy <= 26:
                years.append(2000 + yy)
    return min(years) if years else None


def build_inv_aging_sheet(wb, ents, year):
    """存货按指令号_库龄分析表（SAP 专属，2026-08-13 用户方法论：指令号含年份→库龄推断）。
    期末有留存（借-贷净额≠0）的存货指令号，按指令号年份推断库龄档（1年内/1-2年/2-3年/3年以上）。
    与「按指令号_全年增减变动表」同源（wbs/ord/mat 指令号列），剔除借贷对冲行。
    提示：指令号年份≠实际入库时间，仅提供判断线索（用户认知：指令号早→库龄可能长）。"""
    ws = wb.create_sheet('按指令号_库龄分析表')
    ws.cell(1, 1, f'存货按指令号库龄分析表（{year} 年度；单位：元；库龄=2026−指令号年份，仅判断线索非绝对库龄）').font = SHELL_TITLE_FONT
    _hdr(ws, 3, ['主体', '指令号', '指令年份', '库龄档', '科目码', '科目名', '期末留存(借-贷)', '备注'])
    _INV = ('1403', '1406', '1411', '1404', '5001')
    pairs = list(_iter_comp_gl(ents, year))          # ⚡ 2026-08-26 集团模式逐主体
    if not pairs:
        _txt(ws, 4, 1, '（无 GL 数据）')
        return ws
    _BUCKETS = [(2026, '1年以内'), (2025, '1-2年'), (2024, '2-3年')]
    _ORD = {'3年以上': 0, '2-3年': 1, '1-2年': 2, '1年以内': 3, '无年份信息': 4}

    def _bucket(y_i):
        if y_i is None:
            return '无年份信息'
        for yy, lab in _BUCKETS:
            if y_i == yy:
                return lab
        return '3年以上' if y_i < 2024 else '1年以内'

    r = 4
    n_instr_rows = 0
    n_drop_a = 0
    drop_amt_a = 0.0
    total_net = 0.0
    any_items = False
    first = True
    for comp, gl in pairs:
        if not first:
            r += 1                                     # 主体间空行
        first = False
        agg = {}
        for rr in gl:
            code = str(rr.get('code') or '')
            if not code.startswith(_INV):
                continue
            w = str(rr.get('wbs') or '').strip()
            o = str(rr.get('ord') or '').strip()
            mat = str(rr.get('mat') or '').strip()
            if not w and not o and not mat:
                continue
            instr = w or o or mat
            y_i = _instr_year(instr)
            k = (instr[:40], y_i, code[:10], str(rr.get('name') or '')[:30])
            d = float(rr.get('debit') or 0.0)
            c = float(rr.get('credit') or 0.0)
            a = agg.get(k)
            if a is None:
                agg[k] = [d, c]
            else:
                a[0] += d; a[1] += c
            n_instr_rows += 1
        items = []
        for (instr, y_i, code, nm), (d, c) in agg.items():
            if d > 0.005 and c > 0.005 and abs(d - c) <= 0.01:
                n_drop_a += 1
                drop_amt_a += d
                continue                               # 借贷对冲行剔除
            net = d - c
            if abs(net) <= 0.005:
                continue                               # 无留存不纳入
            items.append((instr, y_i, code, nm, d, c, net))
        items.sort(key=lambda t: (_ORD[_bucket(t[1])], -abs(t[4])))
        cur_bucket = None
        sum_by_bucket = {}
        for instr, y_i, code, nm, d, c, net in items:
            b = _bucket(y_i)
            if b != cur_bucket:
                if cur_bucket is not None:
                    _txt(ws, r, 1, f'{cur_bucket} 小计', bold=True)
                    _money(ws, r, 7, sum_by_bucket[cur_bucket])
                    for cc in range(1, 9):
                        ws.cell(r, cc).font = SHELL_BOLD
                        ws.cell(r, cc).fill = SHELL_TOT_FILL
                    r += 1
                cur_bucket = b
                sum_by_bucket[b] = 0.0
            _txt(ws, r, 1, comp)
            _txt(ws, r, 2, instr)
            _txt(ws, r, 3, y_i if y_i is not None else '—')
            _txt(ws, r, 4, b)
            _txt(ws, r, 5, code)
            _txt(ws, r, 6, nm)
            _money(ws, r, 7, net)
            _txt(ws, r, 8, '贷方>借方(负数)' if net < 0 else '')
            sum_by_bucket[b] += net
            total_net += net
            any_items = True
            r += 1
        if cur_bucket is not None:
            _txt(ws, r, 1, f'{cur_bucket} 小计', bold=True)
            _money(ws, r, 7, sum_by_bucket[cur_bucket])
            for cc in range(1, 9):
                ws.cell(r, cc).font = SHELL_BOLD
                ws.cell(r, cc).fill = SHELL_TOT_FILL
            r += 1
    if any_items:
        _txt(ws, r, 1, '总合计', bold=True)
        _money(ws, r, 7, total_net)
        for cc in range(1, 9):
            ws.cell(r, cc).font = SHELL_BOLD
            ws.cell(r, cc).fill = SHELL_TOT_FILL
        r += 1
    else:
        _txt(ws, r, 1, '（无留存指令号）')
    _txt(ws, r + 1, 1, '注：指令号年份≠实际入库时间（如 F-GF25.1150.12 中 25=2025 年），仅提供库龄判断线索；取指令号中最早年份（保守）。', italic=True)
    _txt(ws, r + 1, 1, '说明：有指令号发生额行 %d 条；剔除借贷对冲（虚增虚减）%d 行 / %s 元；'
                       '余额=期初+借-贷，期初指令号余额需另行取得（本年发生额口径）。'
                       % (n_instr_rows, n_drop_a, format(drop_amt_a, ',.2f')))
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 26
    for c in range(7, 10):
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.column_dimensions['J'].width = 10
    ws.freeze_panes = 'B4'
    return ws


# ----------------------------------------------------------------------------
# 存货其他增减凭证逐笔甄别（SAP 专属，2026-08-10 用户方法论）
# ----------------------------------------------------------------------------

# 核心正常业务对方科目（采购/结转销售/生产流转/差异调整/内部转库）→ 不列入甄别表
_INV_CP_WHITELIST = ('应付账款', '预付账款', '银行存款', '应付票据', '主营业务成本',
                     '其他业务成本', '生产成本', '制造费用', '库存产品', '原材料',
                     '成本差异', '在制品结转', '发出商品', '委托加工', '半成品',
                     '产成品', '在制品', '商品', '销售成本')


def build_inv_other_flow_sheet(wb, ents, year):
    """存货其他增减凭证逐笔甄别表（2026-08-10 用户方法论：分月表『其他增加/其他减少』
    空对方科目追查 + 金额较大未追查分析）。
    数据源：GL 行 cp（凭证级对方科目推断）+ wbs/ord/mat。
    范围：存货科目（1403/1406/1411/1404/5001）借贷行中，对方科目不在核心白名单
    （采购/结转销售/生产流转/差异/内部转库）的全部行——含 费用领用（劳保/机物料/防暑降温）、
    研发领料、转低值易耗品、cp 空（同科目结转）等，逐笔列出供审计甄别。
    列：日期|凭证号|科目码|科目名|对方科目|摘要|借|贷|WBS|订单|物料。"""
    ws = wb.create_sheet('存货其他增减凭证甄别表')
    ws.cell(1, 1, f'存货其他增减凭证逐笔甄别（{year} 年度；对方科目不在核心正常业务白名单的行；单位：元）').font = SHELL_TITLE_FONT
    _hdr(ws, 3, ['主体', '日期', '凭证号', '科目码', '科目名', '对方科目', '摘要', '借', '贷', 'WBS', '订单', '物料'])
    pairs = list(_iter_comp_gl(ents, year))          # ⚡ 2026-08-26 集团模式逐主体
    if not pairs:
        _txt(ws, 4, 1, '（无 GL 数据）')
        return ws
    rows_out = []
    for comp, gl in pairs:
        for rr in gl:
            code = str(rr.get('code') or '')
            if not code.startswith(('1403', '1406', '1411', '1404', '5001')):
                continue
            cp = str(rr.get('cp') or '')
            if cp and any(w in cp for w in _INV_CP_WHITELIST):
                continue
            d = float(rr.get('debit') or 0.0)
            c = float(rr.get('credit') or 0.0)
            if abs(d) < 0.005 and abs(c) < 0.005:
                continue
            rows_out.append((comp, rr, d, c))
    # 排序：金额降序
    rows_out.sort(key=lambda x: -(max(x[2], x[3])))
    r = 4
    sum_d = sum_c = 0.0
    for comp, r_gl, d, c in rows_out:
        _txt(ws, r, 1, comp)
        _txt(ws, r, 2, str(r_gl.get('date') or ''))
        _txt(ws, r, 3, str(r_gl.get('voucher') or ''))
        _txt(ws, r, 4, str(r_gl.get('code') or ''))
        _txt(ws, r, 5, str(r_gl.get('name') or '')[:30])
        _txt(ws, r, 6, str(r_gl.get('cp') or '（无对方科目，同科目结转/单边行）')[:40])
        _txt(ws, r, 7, str(r_gl.get('sm') or '')[:40])
        _money(ws, r, 8, d)
        _money(ws, r, 9, c)
        _txt(ws, r, 10, str(r_gl.get('wbs') or '')[:24])
        _txt(ws, r, 11, str(r_gl.get('ord') or '')[:12])
        _txt(ws, r, 12, str(r_gl.get('mat') or '')[:14])
        sum_d += d; sum_c += c
        r += 1
    if rows_out:
        _txt(ws, r, 1, '合计', bold=True)
        _money(ws, r, 8, sum_d)
        _money(ws, r, 9, sum_c)
        _money(ws, r, 10, sum_d - sum_c)
        for c in range(1, 11):
            ws.cell(r, c).font = SHELL_BOLD
            ws.cell(r, c).fill = SHELL_TOT_FILL
        r += 1
    _txt(ws, r + 1, 1, '说明：其他增减凭证 %d 行（借 %s / 贷 %s）。对方科目为 费用领用（劳动保护费/机物料消耗/防暑降温费）、'
                       '研发领料（研究开发费-直接投入）、转低值易耗品、应付职工薪酬等非采购/结转业务，'
                       '或 cp 空（同科目结转）——请逐笔甄别是否真实业务。' % (
                           len(rows_out), format(sum_d, ',.2f'), format(sum_c, ',.2f')))
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 26
    ws.column_dimensions['F'].width = 36
    ws.column_dimensions['G'].width = 36
    for c in range(8, 13):
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.freeze_panes = 'B4'
    return ws


def build_inv_instruction_composition_sheet(wb, ents, year):
    """指令号组成成分分析表（2026-08-10 用户方法论：按指令号进一步拆组成成分+金额占比）。
    数据源：GL 行 wbs/ord/mat/code/name/debit/credit。
    结构：行=指令号（WBS 优先，300 用订单/物料）；每指令号下按【科目名×借/贷方向】聚合
    带符号净额 + 占该指令号比；仅输出金额 Top 300 指令号（大额优先，避免几十万行）。
    列：指令号|组成科目|方向|金额|占指令号比|指令号合计。"""
    ws = wb.create_sheet('按指令号_组成成分分析表')
    ws.cell(1, 1, f'存货按指令号（WBS/订单）组成成分分析（{year} 年度；金额=带符号净额；单位：元；'
                  f'仅列金额 Top 300 指令号）').font = SHELL_TITLE_FONT
    _hdr(ws, 3, ['主体', '指令号', '组成科目', '方向', '金额', '占指令号比', '指令号合计'])
    pairs = list(_iter_comp_gl(ents, year))          # ⚡ 2026-08-26 集团模式逐主体
    if not pairs:
        _txt(ws, 4, 1, '（无 GL 数据）')
        return ws
    _INV = ('1403', '1406', '1411', '1404', '5001')
    by_inst = {}      # (comp, inst) -> { (nm, dr): amt }
    for comp, gl in pairs:
        for rr in gl:
            code = str(rr.get('code') or '')
            if not code.startswith(_INV):
                continue
            inst = str(rr.get('wbs') or '').strip() or str(rr.get('ord') or '').strip() \
                or str(rr.get('mat') or '').strip()
            if not inst:
                continue
            nm = str(rr.get('name') or '')[:26]
            d = float(rr.get('debit') or 0.0)
            c = float(rr.get('credit') or 0.0)
            amt = d - c
            if abs(amt) < 0.005:
                continue
            key = (nm, '借' if d > 0.005 else '贷')
            m = by_inst.get((comp, inst))
            if m is None:
                m = by_inst[(comp, inst)] = {}
            m[key] = m.get(key, 0.0) + amt
    insts = sorted(by_inst.items(), key=lambda kv: -abs(sum(kv[1].values())))
    r = 4
    n_out = 0
    for (comp, inst), comps in insts[:300]:
        tot = sum(comps.values())
        if abs(tot) < 0.005:
            continue
        for (nm, dr_), amt in sorted(comps.items(), key=lambda x: -abs(x[1])):
            _txt(ws, r, 1, comp)
            _txt(ws, r, 2, inst)
            _txt(ws, r, 3, nm)
            _txt(ws, r, 4, dr_)
            _money(ws, r, 5, amt)
            _money(ws, r, 6, amt / tot if tot else 0.0)
            r += 1
        _txt(ws, r, 2, inst + ' 合计', bold=True)
        _money(ws, r, 5, tot)
        for c in range(1, 7):
            ws.cell(r, c).font = SHELL_BOLD
            ws.cell(r, c).fill = SHELL_TOT_FILL
        r += 1
        n_out += 1
    if not n_out:
        _txt(ws, 4, 1, '（无指令号组成数据）')
        r = 5
    _txt(ws, r + 1, 1, '说明：共 %d 个指令号（Top 300）；组成=指令号下各科目×方向带符号净额，'
                      '占比=组成/指令号合计。' % n_out)
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 30
    ws.column_dimensions['D'].width = 6
    for c in range(5, 7):
        ws.column_dimensions[get_column_letter(c)].width = 15
    ws.column_dimensions['G'].width = 15
    ws.freeze_panes = 'B4'
    return ws


def build_inv_baddebt_detail(ws, ents, year_filter=None):
    """存货跌价准备明细表（2026-08-03 用户要求：原『存货跌价准备_分月变动表』改明细表，
    格式不变但不用分月——每主体一行：期初余额/本期计提(贷方)/本期转回及转销(借方+贷方红字)/期末余额。
    数据=GL 归集(contra) + TB 控制数(1471 期初/期末)；并附『计提数与减值损失科目核对』说明。"""
    from collections import defaultdict
    agg = _agg_inv_account(ents, '1471', '存货跌价准备', ['存货跌价准备'], True)
    ws.cell(1, 1, '存货跌价准备明细表').font = SHELL_TITLE_FONT
    hr = 4
    headers = ['主体', '期初余额', '本期计提', '本期转回及转销', '期末余额']
    for j, h in enumerate(headers, 1):
        cc = ws.cell(hr, j, h); cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL
        cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
    r = hr + 1
    tot = {'op': 0.0, 'inc': 0.0, 'dec': 0.0, 'cl': 0.0}
    for (ent, year), d in sorted(agg.items()):
        if year_filter is not None and str(year) != str(year_filter):
            continue
        op = float(d.get('km_op') or 0.0)
        inc = float(d.get('dec_t', {}).get('计提', 0.0) or 0.0)      # 贷方=计提
        dec = float(d.get('inc_t', {}).get('转回', 0.0) or 0.0)      # 借方+贷方红字=转回及转销
        cl = float(d.get('km_cl') or 0.0)
        _txt(ws, r, 1, ent)
        _money(ws, r, 2, op)
        _money(ws, r, 3, inc)
        _money(ws, r, 4, dec)
        _money(ws, r, 5, cl)
        for c in range(1, 6):
            ws.cell(r, c).border = SHELL_BORDER
        tot['op'] += op; tot['inc'] += inc; tot['dec'] += dec; tot['cl'] += cl
        r += 1
    # 合计
    _txt(ws, r, 1, '合计', bold=True, align=SHELL_CEN)
    for j, key in enumerate(('op', 'inc', 'dec', 'cl'), start=2):
        _money(ws, r, j, tot[key])
        ws.cell(r, j).fill = SHELL_TOT_FILL
        ws.cell(r, j).font = SHELL_BOLD
    for c in range(1, 6):
        ws.cell(r, c).border = SHELL_BORDER
        ws.cell(r, c).fill = SHELL_TOT_FILL
    r += 2
    # —— 计提数与减值损失科目核对（2026-08-03 用户要求）——
    recon = _recon_baddebt_impairment(ents, year_filter, '存货跌价准备',
                                      ['资产减值损失'])
    ws.cell(r, 1, '跌价计提数与资产减值损失科目核对（借：资产减值损失 / 贷：存货跌价准备；带符号口径，红字=转回）：').font = SHELL_BOLD
    r += 1
    for (ent, year), v in sorted(recon.items()):
        _txt(ws, r, 1, ent)
        _txt(ws, r, 2, year, align=SHELL_CEN)
        _txt(ws, r, 3, '跌价准备贷方发生额', align=SHELL_RGT)
        _money(ws, r, 4, v[0])
        _txt(ws, r, 5, '资产减值损失借方', align=SHELL_RGT)
        _money(ws, r, 6, v[1])
        _txt(ws, r, 7, '差异', align=SHELL_RGT)
        _money(ws, r, 8, v[0] - v[1])
        if abs(v[0] - v[1]) > 0.005:
            ws.cell(r, 8).font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
        for c in range(1, 9):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    for i, w in enumerate([16, 16, 16, 18, 16, 16, 16, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def _recon_baddebt_impairment(ents, year_filter, contra_name, loss_kws):
    """核对备抵科目(存货跌价准备/坏账准备)贷方计提数 vs 减值损失科目借方发生额（按主体×年度）。
    减值损失=GL 科目名含『信用减值损失/资产减值损失』的借方发生额（红字按负计）。
    返回 {(ent, year): (计提额, 损失借方额)}。"""
    out = {}
    for ent in sorted(ents):
        for year in sorted(ents[ent].get('km', {})):
            gl_path = ents[ent]['gl'].get(year)
            if not gl_path:
                continue
            gl = _read_gl_cached(gl_path)
            prov = 0.0
            loss = 0.0
            for row in gl:
                nm = row.get('name') or ''
                if contra_name in nm and '备抵' not in nm:
                    prov += (row.get('credit') or 0.0)     # 贷方发生额（带符号：正=计提、红字=转回）
                if any(k in nm for k in loss_kws):
                    loss += (row.get('debit') or 0.0)      # 减值损失借方发生额（带符号，红字=转回）
            out[(ent, year)] = (prov, loss)
    return out


def process_folder(data_dir, out_path=None):
    """目录模式：自动发现主体，生成 存货审计底稿_YYYY_生成.xlsx（每年度一份）。
    2026-08-03 用户方法论：删 R2 说明行、删勾稽行；跌价准备含 计提+转销。"""
    import argparse
    ents = _discover_entities(data_dir)
    if not ents:
        print(f'[SKIP] {data_dir} 无核算主体')
        return None
    years = sorted({y for e in ents for y in ents[e]['km']})
    out_files = []
    for y in years:
        out = out_path or os.path.join(data_dir, f'存货审计底稿_{y}_生成.xlsx')
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        # 1 审定表
        build_audit_schedule(ents, wb)
        print(f'[OK] 存货审定表 {y}')
        # 1b 存货按主体余额汇总（2026-08-04 规范，供合并核对单体提取）
        build_inv_by_entity_summary(wb, ents, y)
        print(f'[OK] 存货按主体明细表 {y}')
        # 1c ⚡ 2026-08-11 用户需求：存货集团审定表（每单位一行：期初/增/减/末/调整/审定 + 净额）
        try:
            build_inventory_group_audit_sheet(wb, ents, y)
        except Exception as _ex:
            print(f'  ⚠️ 存货集团审定表生成失败：{_ex}')
        # 2 各存货科目 分月变动表（列扩展版）；1471 跌价准备→明细表（不分月，2026-08-03 用户要求）
        # ⚡ 2026-08-10 SAP：按实际科目名称拆分（1406 库存产品→库存商品/产成品/半成品/在制品
        # 各自分月表，控制数取对应 10 位子目码；1404/1407 成本差异单列）
        _AA = ANALYSIS_ACCOUNTS
        if _is_sap():
            _AA = [('1403000000', '原材料', ['原材料'], False),
                   ('1406010000', '库存商品', ['库存商品'], False),
                   ('1406030000', '产成品', ['产成品'], False),
                   ('1406040000', '半成品', ['半成品'], False),
                   ('1406050000', '在制品', ['在制品', '在产品'], False),
                   ('1406', '发出商品', ['发出商品'], False),
                   ('1411000000', '委托加工物资', ['委托加工物资'], False),
                   ('1404', '材料成本差异', ['材料成本差异', '成本差异'], False),
                   ('1471', '存货跌价准备', ['跌价准备'], True)]
        for code, name, targets, contra in _AA:
            if not _account_has_data(ents, code):
                print(f'[SKIP] {name}（无数据）')
                continue
            if contra:
                build_inv_baddebt_detail(wb.create_sheet(f'{name}明细表'), ents, year_filter=y)
                print(f'[OK] {name}明细表 {y}')
                continue
            build_inv_monthly_summary(wb.create_sheet(f'{name}_分月变动表'), ents,
                                      f'{name}分月变动表', code, name, targets, contra, year_filter=y)
            print(f'[OK] {name}_分月变动表 {y}')
        # 3 生产成本 成本分析表（2026-08-03 用户要求：删除 变动分析表×3 / 生产成本分月变动表 /
        #   生产成本对应科目分月明细表，仅保留成本分析表）
        build_cost_analysis_sheet(wb.create_sheet('生产成本_成本分析表'), ents, '生产成本成本分析表', year_filter=y)
        # 3b 按指令号（WBS/订单）全年增减变动表（SAP 专属，用户方法论：剔除虚增虚减）
        if _is_sap():
            try:
                build_inv_instruction_sheet(wb, ents, y)
                print(f'[OK] 按指令号_全年增减变动表 {y}')
            except Exception as _ex:
                print(f'[WARN] 按指令号增减变动表失败：{_ex}')
            # 3c 存货其他增减凭证逐笔甄别（用户方法论：其他增减追查）
            try:
                build_inv_other_flow_sheet(wb, ents, y)
                print(f'[OK] 存货其他增减凭证甄别表 {y}')
            except Exception as _ex:
                print(f'[WARN] 其他增减甄别表失败：{_ex}')
            # 3d 指令号组成成分分析（用户方法论：指令号下组成+金额占比）
            try:
                build_inv_instruction_composition_sheet(wb, ents, y)
                print(f'[OK] 按指令号_组成成分分析表 {y}')
            except Exception as _ex:
                print(f'[WARN] 组成成分分析表失败：{_ex}')
            # 3e 指令号库龄分析（2026-08-13 用户方法论：指令号年份→库龄推断，早→可能长）
            try:
                build_inv_aging_sheet(wb, ents, y)
                print(f'[OK] 按指令号_库龄分析表 {y}')
            except Exception as _ex:
                print(f'[WARN] 库龄分析表失败：{_ex}')
            # 3f 与主营业务成本倒退核对（#702 用户方法论：存货按指令号变动后倒退核对 6401）
            try:
                build_inv_cogs_recon_sheet(wb, ents, y)
                print(f'[OK] 与主营业务成本_倒退核对表 {y}')
            except Exception as _ex:
                print(f'[WARN] 倒退核对表失败：{_ex}')
        # 5 存货附注汇总
        try:
            build_inventory_footnote(wb, ents, y, data_dir=data_dir)
            print(f'[OK] 存货附注汇总 {y}')
        except Exception as ex:
            print(f'[WARN] 存货附注汇总失败：{ex}')
        # 6 保存
        try:
            finalize_workbook(wb)
        except Exception:
            pass
        # ⚡ 2026-08-28 #875：对方科目核对（模块内集成）
        try:
            from counterparty_recon import inject_into_wb_auto
            inject_into_wb_auto(wb, data_dir, y, '存货', ents_set=set(ents))
        except Exception as _ex:
            print(f'  ⚠️ 存货对方科目核对注入失败：{_ex}')
        _safe_save(wb, out)
        wb.close()
        print(f'✅ 已生成：{out}')
        out_files.append(out)
    return out_files[0] if out_files else None


def main(argv=None):
    import argparse
    p = argparse.ArgumentParser(description='存货审计底稿生成（含泰国 1140 映射）')
    # 2026-08-04：加位置参数 data_dir——regen_all 用 [module, data_dir] 调用（sys.argv[1]=目录），
    # 原仅 --input 导致 regen_all 跑存货报 unrecognized arguments exit=2。
    p.add_argument('data_dir', nargs='?', default=None, help='数据目录（位置参数，兼容 regen_all）')
    p.add_argument('--input', '-i', default=None, help='数据目录（默认桌面 G）')
    p.add_argument('--output', '-o', default=None, help='输出文件（默认 <input>/存货审计底稿_YYYY_生成.xlsx）')
    args = p.parse_args(argv)
    data_dir = args.data_dir or args.input
    if not data_dir:
        data_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'G')
    if not os.path.isdir(data_dir):
        print(f'[ERROR] 目录不存在：{data_dir}')
        return 1
    process_folder(data_dir, args.output)
    return 0



# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None


def _is_sap():
    """判断当前数据是否为 SAP 布局。
    ⚡ 2026-08-26 修复：原全用 `_adapter.current_comp()` 判断 SAP——AH 集团模式
    （--group，current_comp=None）下恒走 U8 分支 → SAP 10 位码/名称分类全部失效：
    分月表只生 3 个（库存商品/产成品等漏列）、附注跌价重复填 65M×9、生产成本空壳。
    统一委托 sap_adapter.is_sap（按 _DATA_ROOT 布局判定），group 模式同样生效。"""
    if _adapter is None:
        return False
    try:
        return bool(_adapter.is_sap(getattr(_adapter, '_DATA_ROOT', None)))
    except Exception:
        return False


if __name__ == '__main__':
    _mp = _find_managed_python()
    if _mp and os.path.abspath(sys.executable).lower() != os.path.abspath(_mp).lower():
        import subprocess
        rc = subprocess.run([_mp, '-B', os.path.abspath(__file__)] + sys.argv[1:]).returncode
        sys.exit(rc)
    sys.exit(main())

def build_inventory_group_audit_sheet(wb, ents, year):
    """⚡ 2026-08-11 用户需求：存货集团审定表——每核算单位一行：
    主体|期初数|本期增加(借发)|本期减少(贷发)|期末数|审计调整数|审定数，尾部 合计 + 减:存货跌价准备 + 期末净额。
    数据=科目余额表(权威控制数，原值口径；1471 跌价单列于净额行)。"""
    from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_BOLD,
                             SHELL_NUM, SHELL_LEFT, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                             SHELL_TOT_FILL, SHELL_BODY)
    from openpyxl.utils import get_column_letter
    ws = wb.create_sheet(f'存货集团审定表_{year}')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    c = ws.cell(1, 1, f'存货集团审定表（{year} 年度）· 每核算单位一行 · 账面数（未审）+ 审计调整/审定数 · 净额=原值+跌价')
    c.font = SHELL_TITLE_FONT
    hdr = ['序号', '核算主体', '期初数', '本期增加(借发)', '本期减少(贷发)', '期末数', '审计调整数', '审定数']
    for j, h in enumerate(hdr, 1):
        cc = ws.cell(2, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    r = 3
    g = [0.0] * 4
    g_bd = 0.0
    g_bd_op = 0.0
    _seq = 0
    for ent in sorted(ents):
        # ⚡ 2026-08-12 修复：SAP entities 结构 {年: ['km','gl',...]}（年度为 key），
        #   U8 是 {km:{年:path}}。原 `ents[ent]['km'].get(year)` 对 SAP KeyError
        #   （'km' 不存在）→ 存货集团审定表生成失败。
        _bun = ents[ent]
        if 'km' in _bun and isinstance(_bun['km'], dict):
            km_path = _bun['km'].get(year)      # U8 结构
        elif year in _bun and isinstance(_bun[year], dict):
            km_path = _bun[year].get('km')      # ⚡ 2026-08-26 SAP 结构 {年: {'km': path, 'gl': [...]}}
        elif year in _bun and isinstance(_bun[year], list):
            km_path = _bun[year][0] if _bun[year] else None   # SAP 结构：list[0]=km
        else:
            km_path = None
        if not km_path:
            continue
        km = _read_km_cached(km_path, ent)
        if not km:
            continue
        op = inc = dec = cl = 0.0
        # ⚡⚡ 2026-08-26 修复（任务878 存货集团审定表 2519 万口径差）：SAP 存货
        #   按名称分类（SAP_INV_CATS：产成品/半成品/在制品/库存商品都挂在 1406 库存产品
        #   下），原 INV_LEVEL2 代码聚合把 1406 全量当"发出商品"、1405/1412 无代码漏计
        #   → 集团审定表 11.745亿 vs 附注/按科目审定表 11.493亿（差 2519万=产成品等细分）。
        #   对 SAP 改用【末级行按名称分类】汇总（与附注同口径），U8 仍走 INV_LEVEL2 代码。
        if _adapter is not None and _adapter.is_sap(getattr(_adapter, '_DATA_ROOT', '') or ''):
            for _c, _r in km.items():
                if len(_c) < 8 or _r.get('level', 4) < 4:
                    continue                       # 只取末级行（10 位码）
                _cat = _sap_inv_cat(_r.get('name', ''))
                if not _cat or _cat == '存货跌价准备':
                    continue
                op += _r['opening']; inc += _r['debit']; dec += _r['credit']; cl += _r['closing']
        else:
            for code in INV_LEVEL2:
                if code == '1471':
                    continue
                _km_code = THAI_INV_MAP.get(code, code) if '泰国' in str(ent) else code
                rec = km.get(_km_code)
                if rec:
                    op += rec['opening']; inc += rec['debit']; dec += rec['credit']; cl += rec['closing']
        bd = km.get('1471')
        if bd is None and _adapter is not None and _adapter.is_sap(getattr(_adapter, '_DATA_ROOT', '') or ''):
            # SAP 跌价准备：名称分类归"存货跌价准备"的末级行汇总
            bd = dict(opening=0.0, debit=0.0, credit=0.0, closing=0.0)
            for _c, _r in km.items():
                if len(_c) < 8 or _r.get('level', 4) < 4:
                    continue
                if _sap_inv_cat(_r.get('name', '')) == '存货跌价准备':
                    bd['opening'] += _r['opening']; bd['debit'] += _r['debit']
                    bd['credit'] += _r['credit']; bd['closing'] += _r['closing']
        bd_cl = bd['closing'] if bd else 0.0
        bd_op = bd['opening'] if bd else 0.0
        # ⚡ 2026-08-11：跳过判断用 4 字段（op/inc/dec/cl）全 0——期初/期末 0 但仅发生额
        #   （如 dq 2024 库存商品 借发25万/贷发25万）也应显示（曾误跳过 → 集团审定表全 0）
        if all(abs(x) < 0.005 for x in (op, inc, dec, cl)):
            continue
        _seq += 1
        row_vals = [_seq, ent, round(op, 2), round(inc, 2), round(dec, 2), round(cl, 2), 0.0, round(cl, 2)]
        for j, v in enumerate(row_vals, 1):
            cc = ws.cell(r, j, v)
            cc.border = SHELL_BORDER
            if j >= 4:
                cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
            elif j == 1:
                cc.alignment = SHELL_CEN
        # ⚡ 2026-08-12 修复：原 `g = [g[k] + v[k] ...]` 中 v 是循环残留 float
        #   （row_vals 的最后一个元素）→ 'float' object is not subscriptable，
        #   存货集团审定表生成失败。改用 op/inc/dec/cl 显式累加。
        g = [g[0] + op, g[1] + inc, g[2] + dec, g[3] + cl]
        g_bd += bd_cl
        g_bd_op += bd_op
        r += 1
    # 合计行
    row_vals = ['合计', None, round(g[0], 2), round(g[1], 2), round(g[2], 2), round(g[3], 2), 0.0, round(g[3], 2)]
    for j, v in enumerate(row_vals, 1):
        cc = ws.cell(r, j, v if j > 1 else None)
        if j == 1:
            cc = ws.cell(r, 1, '合计')
        cc.border = SHELL_BORDER
        cc.font = SHELL_BOLD
        cc.fill = SHELL_TOT_FILL
        if j >= 4:
            cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
        elif j == 1:
            cc.alignment = SHELL_CEN
    r += 1
    # 减:跌价准备 + 期初/期末净额（与报表核对）
    # ⚡ 2026-08-26 修复：原仅期末列有「减:跌价准备/期末净额」，期初列无 → 期初/期末口径
    #   不对称（期初=原值、期末=净额）。补期初跌价与期初净额，两期口径统一。
    def _fill_bd_row(label, bold, fills=None):
        nonlocal r
        for j in range(1, 9):
            ws.cell(r, j).border = SHELL_BORDER
            if fills:
                ws.cell(r, j).fill = SHELL_TOT_FILL
        cc = ws.cell(r, 2, label)
        cc.font = SHELL_BOLD if bold else SHELL_BODY
        cc.alignment = SHELL_LEFT
        return r + 1

    def _put(rr, col, val):
        if abs(val) >= 0.005:
            c3 = ws.cell(rr, col, round(val, 2))
            c3.number_format = SHELL_NUM; c3.alignment = SHELL_RGT
            c3.font = SHELL_BOLD

    # 减：存货跌价准备（期初 + 期末）
    _r0 = r
    r = _fill_bd_row('减：存货跌价准备', False)
    _put(_r0, 3, g_bd_op); _put(_r0, 6, g_bd)
    # 存货期初净额 = 期初原值 + 期初跌价
    _r1 = r
    r = _fill_bd_row('存货期初净额', True, fills=True)
    _put(_r1, 3, g[0] + g_bd_op)
    # 存货期末净额 = 期末原值 + 期末跌价
    _r2 = r
    r = _fill_bd_row('存货期末净额', True, fills=True)
    _put(_r2, 6, g[3] + g_bd)
    for i, w in enumerate([6, 16, 16, 18, 18, 16, 14, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'
    return ws
