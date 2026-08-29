# -*- coding: utf-8 -*-
# FINGERPRINT: 产出={科目代码}{科目名称}_生成.xlsx(10个利润表科目各一份) | 关键列=损益类借发/贷发/与TB控制数勾稽/税金及附加明细表含对方科目(应交税费子目)及金额勾稽/其他收益明细表含对方科目(递延收益)及与递延收益摊销额勾稽 | 职责=损益科目发生额与TB勾稽+资产减值/税金及附加/其他收益核对(独立生成器) | 抽凭排除=税金及附加与其他收益不生成凭证抽查表(摊销转入凭证不抽,改由勾稽核对表取代)
"""
pl_detail.py — 利润表科目明细底稿(含凭证抽查表) 编制小程序
=================================================================================
用法：把【账套导出文件夹】拖到同目录的 run_pl_detail.bat 上即可。
  · 自动发现文件夹内各核算主体的「科目余额表」(控制数) + 「综合查询明细表」(GL 凭证级)
  · 对 10 个利润表科目各自生成【一份】合并底稿：<科目名称>审计底稿_生成.xlsx
        （含各年明细表 ＋ 凭证抽查表，沿用中央 audit_templates/ 外壳格式）
        —— 即「每个有数据的科目 = 1 份底稿」，不再拆分单独的凭证抽查表文件
  · 凭证抽查表据综合查询明细表发生额逐笔填列；【已剔除「结转期间损益」类凭证】，
    仅保留真实业务凭证供抽凭。
  · 发生额口径：收入/利得类(rev)取【贷方发生额】；费用/损失类(exp)取【借方发生额】。
    这两类在年终结转后 TB 借发=贷发，故与 TB 控制数勾稽时直接比较 GL 自然方合计与 TB 借发。
依赖：openpyxl（WorkBuddy 托管 Python 自带）
"""
import os
import re
import sys
from collections import defaultdict, OrderedDict
import copy
# ============ 顶部引导（必须在 import openpyxl 之前）============
# 拖入/双击本 .py 时，Windows 可能用「系统 Python 3.14（无 openpyxl）」启动，
# 顶部 import openpyxl 会立即 ModuleNotFoundError 闪退。故在此先检测并在必要时
# 用 WorkBuddy 托管 Python 重新执行本文件，确保 openpyxl 可用、控制台保留。
import os as _bs_os
import sys as _bs_sys
import subprocess as _bs_sub


def _find_managed_python():
    """定位 WorkBuddy 托管 Python（含 openpyxl）。"""
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


from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
# 统一外壳样式与加载逻辑（六程序共享）
from audit_shell import (
    SHELL_HFILL, SHELL_HFONT, SHELL_SUB_FONT, SHELL_BOLD, SHELL_NUM, SHELL_THIN,
    SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT, resolve_template, AUDIT_TEMPLATES_DIR,
    finalize_workbook,
)
from audit_common import _safe_save, is_zero_amount, read_gl_rows, voucher_skip_keys, voucher_is_carryover, is_accrual_line, apply_pl_carryover_rule  # 共享库

# ---------------- 利润表 10 科目（代码, 名称, 类别） ----------------
# kind: 'rev' = 收入/利得类（发生额取【贷方】）；'exp' = 费用/损失类（发生额取【借方】）
PL_SUBJECTS = [
    # 资产处置损益 实为"利得/损失"混合科目，本数据集其真实交易(处置利得)在贷方、年末结转(借 资产处置损益 贷 本年利润)在借方。
    # 若按 exp(取借方)则会把"结转本年利润"的结账分录误当作发生额，且抽凭勾稽"实际抽凭=0"与实际可见的贷方利得矛盾。
    # 故与 投资收益/公允价值变动损益 一致取 'rev'(贷方)，使明细表"本期未审数"与抽凭勾稽均落在真实利得侧。
    # 2026-08-04 修复：编码按 FY 真实值（其他收益=6113 新准则，非 6117）；审定表 _gross_entity 另有名称兜底。
    (6115, '资产处置损益', 'rev'),
    (6101, '公允价值变动损益', 'rev'),
    (6113, '其他收益', 'rev'),
    (6301, '营业外收入', 'rev'),
    (6403, '税金及附加', 'exp'),
    (6701, '资产减值损失', 'exp'),
    (6702, '信用减值损失', 'exp'),
    (6801, '所得税费用', 'exp'),
    (6711, '营业外支出', 'exp'),
    (6111, '投资收益', 'rev'),
]
# 以下科目不生成「凭证抽查表」：其发生额由对应的勾稽核对表(见 build_subject_workbook)取代。
# 税金及附加 → 增加凭证抽查表（此前因其与应交税费计提在明细表直接核对，无单独抽查；用户要求加上）
VOUCHER_EXCLUDE = set()

# 科目名 → 发生额自然方（'rev'取贷方 / 'exp'取借方），供利润表蓝字红字规则判定
_PL_KIND = {name: kind for (_code, name, kind) in PL_SUBJECTS}

def _pl_kind(subj):
    return _PL_KIND.get(subj, 'exp')

# 名称关键词 → 科目（按「科目名称」首段匹配，跨科目体系稳健）
SUBJECT_KEYWORDS = {
    '资产处置损益': ['资产处置损益', '资产处理损益', '处置损益', '资产处置收益'],
    '公允价值变动损益': ['公允价值变动损益', '公允价值变动'],
    '其他收益': ['其他收益'],
    '营业外收入': ['营业外收入'],
    '税金及附加': ['税金及附加'],
    '资产减值损失': ['资产减值损失'],
    '信用减值损失': ['信用减值损失'],
    '所得税费用': ['所得税费用', '所得税'],
    '营业外支出': ['营业外支出'],
    '投资收益': ['投资收益', '投资损益'],
}

# 模板（格式外壳）：统一走中央 audit_templates/；可用环境变量 PL_TEMPLATE_DIR 覆盖外壳目录。
# 注：旧 VOUCHER_TEMPLATE(gu股份公司\记账凭证测试.xlsx) 已废弃——凭证抽查表现由代码直接构建。

# ---------------- 样式（统一为 audit_shell 调色板：蓝头 1F4E78 / 黄高亮小计合计） ----------------
HDR_FONT = SHELL_HFONT
HDR_FILL = SHELL_HFILL
TOT_FONT = SHELL_BOLD
TOT_FILL = PatternFill('solid', fgColor='FFF2CC')
NUM_FMT = SHELL_NUM
THIN = SHELL_THIN
BORDER = SHELL_BORDER
CENTER = SHELL_CEN
LEFT = SHELL_LEFT
RIGHT = SHELL_RGT

DETAIL_COLS = 6
DETAIL_HEADERS = ['核算主体', '二级科目名称', '上年数', '本期数',
                  '增减额', '变动率%']
VOUCHER_HEADERS = [
    '核算主体', '测试序号', '日期', '凭证字', '凭证号',
    '二级科目（或客户/供应商名称）', '摘要', '借方金额', '贷方金额',
    '对方科目', '与原始凭证相符', '原始凭证内容', '原始凭证日期',
    '会计处理正确', '所属时间无误'
]


def safe(x):
    try:
        return float(x)
    except Exception:
        return 0.0


def _subject_of(name):
    """将 GL/TB 的科目名称首段映射到利润表科目名；返回科目名或 None。
    2026-08-06 修复：匹配由 s0.startswith(kw) 改为 kw in s0——FY 账套 6403 科目名
    『营业务税金及附加-城建税』（前缀『营业务』），startswith('税金及附加') 失败 →
    GL 行全被 _subject_of 过滤 → 税金及附加底稿未生成（TB 有 197.8 万）。
    2026-08-07 修复：『递延所得税资产/负债』(1811/2901) 首段含『所得税』被误归入
    所得税费用(6801) → dq 所得税费用明细表/附注混入 DTL/DTA 行（462,792.88 /
    -592,128.17），且勾稽B def_gl_net 被污染。递延所得税资产/负债是资产负债表
    科目、非损益科目，首段以『递延』开头的科目名不参与利润表科目匹配。"""
    if not name:
        return None
    s0 = str(name).split('-')[0].strip()
    if s0.startswith('递延'):
        return None   # 递延所得税资产/负债 = 资产负债表科目，不进利润表底稿
    for subj, kws in SUBJECT_KEYWORDS.items():
        for kw in kws:
            if s0 == kw or kw in s0:
                return subj
    return None


def nat_amt(kind, debit, credit):
    """发生额（自然方）：rev 取贷方，exp 取借方。"""
    return credit if kind == 'rev' else debit


def _impair_recon_status(ie, rc_cr, rc_dr, has_seg):
    """减值损失(ie) ↔ 减值准备贷计提(rc_cr, 借转回/转销=rc_dr) 单行勾稽状态。"""
    if not has_seg:
        return '需人工核对'
    if abs(ie) < 1e-6 and abs(rc_cr) < 1e-6:
        return '—'
    diff = ie - rc_cr
    tol = max(1.0, max(abs(ie), abs(rc_cr)) * 1e-6)
    if abs(diff) <= tol:
        return '一致'
    if abs(diff - rc_dr) <= tol:
        return '一致(转销)'
    return '不一致'


# ---------------- 从文件名稳健识别会计年度 ----------------
def _extract_year(fn):
    """从文件名稳健识别会计年度。

    常见陷阱：25 年账套导出文件带 2026 时间戳（如 『…_20260723152733.xlsx』），
    若用 re.search(r'(20\\d{2})', fn) 会先命中时间戳 2026 而非会计年度 2025；
    另有一些文件直接用『25』作年份后缀且无 20xx（如 『母公司25科目余额表.xlsx』），
    旧逻辑因找不到 20xx 而整文件跳过，导致主体丢失、科目无数据。
    本函数先剥离导出时间戳（6~12 位 20xx 数字串），再匹配 4 位会计年度(20xx)，
    否则匹配独立 2 位年度(如 25 → 2025)。"""
    s = fn
    s = re.sub(r"(?<!\d)20\d{6,12}(?!\d)", "", s)   # 剥离 20+6~12 位数字串(导出时间戳)
    s = re.sub(r"(?<!\d)20\d{6}(?!\d)", "", s)      # 剥离独立 8 位 YYYYMMDD
    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", s)      # 4 位会计年度
    if m:
        return m.group(1)
    m2 = re.search(r"(?<!\d)(\d{2})(?!\d)", s)       # 独立 2 位年度 → 20xx
    if m2:
        try:
            yy = int(m2.group(1))
        except ValueError:
            return None
        if 0 <= yy <= 99:
            return f"20{yy:02d}"
    return None


# ---------------- 发现账套实体 ----------------
def _discover_entities(data_dir):
    # ⚡ 2026-08-09 治本统一：本地 discover 收敛为 audit_common.discover_entities——
    # U8 模式读原文件全量；SAP 模式 patch 后=adapter（current_comp 单主体过滤内置）。
    from audit_common import discover_entities as _de
    return _de(data_dir)


# ---------------- 探测 GL 列位置 ----------------
def detect_gl_columns(ws):
    idx = {'name': 0, 'date': 1, 'vch_type': 2, 'vch_no': 3, 'opp': 4,
           'summary': 5, 'debit': 6, 'credit': 7}
    for r in range(1, 3):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if not v:
                continue
            s = str(v)
            if '科目' in s and ('名称' in s or '编码' in s or '代码' in s):
                idx['name'] = c - 1
            elif '日期' in s or '期间' in s:
                idx['date'] = c - 1
            elif '借方' in s:
                idx['debit'] = c - 1
            elif '贷方' in s:
                idx['credit'] = c - 1
            elif '摘要' in s:
                idx['summary'] = c - 1
            elif '对方' in s:
                idx['opp'] = c - 1
            elif '凭证' in s:
                if '号' in s:
                    idx['vch_no'] = c - 1
                else:
                    idx['vch_type'] = c - 1
    return idx


# ---------------- 读 GL：发生额聚合 + 凭证行 ----------------
def read_pl_gl(data_dir, entities, years):
    # ⚡ 2026-08-10 根治 SAP 口径：SAP TB"借方1-7/贷方1-7"列是【利润中心级净额方向列】
    #   （红字冲回记负借方，如 6500003587 冲回坏账准备 3,911,685.25 在 TB 记负借、GL 记贷）
    #   → 无论 GL gross 还是净额都无法与 TB 对齐（税金及附加 gross 4041万 vs TB 3413万）。
    #   明细表/附注/勾稽的取数【改为从 TB 子目聚合】（与审定表 read_km 一级行同源，
    #   Σ子目 jf/df = 一级 = 审定控制数，天然一致）；GL 逐笔仅保留供凭证抽查。
    if _adapter is not None and _adapter.is_sap(data_dir):
        gl_agg = defaultdict(lambda: [0.0, 0.0])
        gl_rows = defaultdict(list)
        tb_full = _adapter.read_tb_full(data_dir, entities)   # {(comp, code, name, y): {qc,jf,df,qm}}
        for (e, code, nm, y), v in tb_full.items():
            subj = _subject_of(nm)
            if subj is None:
                continue
            key = (subj, e, nm, str(y))
            gl_agg[key][0] += float(v.get('jf') or 0.0)
            gl_agg[key][1] += float(v.get('df') or 0.0)
        # 凭证抽查：GL 逐笔 gross（adapter 缓存，同进程仅首调全量）
        # ⚡⚡ 2026-08-27 修复（税金及附加抽凭贷方冲回问题）：SAP 分支原样保留 GL 行借贷，
        #   计提冲回凭证（如 6500000840 冲回城建税，GL 记『税金及附加 贷方 486,604』= 红字
        #   冲减计提）在抽凭表显示在贷方列；按用户方法论，exp 科目非结转贷方蓝字应并入借方
        #   红字（抵减借方发生），不得在贷方反映。与 U8 分支一致应用 apply_pl_carryover_rule。
        for e, yd in entities.items():
            for r in _adapter.read_gl(e):
                nm = str(r.get('name') or '')
                subj = _subject_of(nm)
                if subj is None:
                    continue
                y = str(r.get('year') or '')
                _db = float(r.get('debit') or 0.0)
                _cr = float(r.get('credit') or 0.0)
                _adj = apply_pl_carryover_rule(
                    [{'dr': _db, 'cr': _cr, 'name': nm,
                      'opp': str(r.get('cp') or ''), 'summary': str(r.get('sm') or ''),
                      '_carry': False}],
                    _pl_kind(subj))[0]
                gl_rows[(subj, e, y)].append([r.get('date'), r.get('vtype'), r.get('vno'),
                                              r.get('cp'), r.get('sm'), _adj['dr'], _adj['cr'],
                                              nm])   # ⚡⚡ 2026-08-28 修复（抽查表借贷颠倒根因）：SAP 分支漏第 8 元素"本方科目名"→ write_voucher_sheet 的 subj_nm 回退成 opp（对方科目）→ 抽查表科目名全显示成对方科目（借应交税费/贷所得税费用）
        return gl_agg, gl_rows
    gl_agg = defaultdict(lambda: [0.0, 0.0])   # (subject, entity, l2, year) -> [debit, credit]
    gl_rows = defaultdict(list)               # (subject, entity, year) -> [date,字,号,对方,摘要,借,贷]
    from audit_common import voucher_is_carryover as _is_carry
    for e, yd in entities.items():
        for y, paths in yd.items():
            g = paths.get('gl')
            if not g or not os.path.exists(g):
                continue
            try:
                wb = load_workbook(g, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ 跳过 GL 读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            ci = detect_gl_columns(ws)
            # 防御：将列索引钳制到实际列范围，避免个别文件列序异常导致越界崩溃
            mc = ws.max_column
            for k in list(ci.keys()):
                if ci[k] < 0:
                    ci[k] = 0
                if ci[k] >= mc:
                    ci[k] = max(0, mc - 1)
            rows = [r for r in ws.iter_rows(min_row=3, values_only=True)]
            # ---- 凭证级结转损益识别（修复 #386 根因）----
            # 旧逻辑仅按【单行】摘要/科目名/对方判定结转，导致年终结账分录「借 本年利润 / 贷 损益科目」
            # 中损益科目的贷方行（摘要如『成本费用结转』含『结转』但无『损益/利润』、对方科目为空）
            # 被误判为非结转、其贷方蓝字被翻成借方红字、把真实费用发生额抵销为 0，使明细表≠TB审定表。
            # 现改为：同一凭证(字+号)内【任一行】含「本年利润/未分配利润/结转损益」即整笔视为年终结账，
            # 该凭证内所有损益科目的非自然方蓝字均【保留不翻转】，明细表发生额自然=TB控制数。
            carry_v = set()   # (vch_type, vch_no) 集合
            vsum_map = {}     # 凭证键(年月,字,号) -> 首个非空摘要（U8 结转账常只有首行有摘要，逐行抽凭时补全）
            for r in rows:
                if not r or len(r) <= ci['credit']:
                    continue
                nm = r[ci['name']] if 0 <= ci['name'] < len(r) else None
                if not nm:
                    continue
                opp_v = r[ci['opp']] if 0 <= ci['opp'] < len(r) else None
                summ_v = r[ci['summary']] if 0 <= ci['summary'] < len(r) else None
                vt = r[ci['vch_type']] if 0 <= ci['vch_type'] < len(r) else None
                vn = r[ci['vch_no']] if 0 <= ci['vch_no'] < len(r) else None
                if _is_carry(summary=summ_v, name=nm, opp=opp_v):
                    carry_v.add((vt, vn))
                # 摘要仅记录【非空白】值（2026-07-31 修复：GL 源首行摘要可能为纯空格" "，
                # 若记录" "会把后续行的空格当有效摘要，导致 结转兜底 被 `if not summ_v` 跳过）
                if summ_v and str(summ_v).strip():
                    d0 = r[ci['date']] if 0 <= ci['date'] < len(r) else None
                    ym = str(d0)[:7] if d0 else ''
                    vsum_map.setdefault((ym, vt, vn), summ_v)
            _whole_dup = False
            _file2 = []
            for r in rows:
                try:
                    if not r or len(r) <= ci['credit']:
                        continue
                    nm = r[ci['name']] if 0 <= ci['name'] < len(r) else None
                    if not nm:
                        continue
                    jf = safe(r[ci['debit']])
                    df = safe(r[ci['credit']])
                    if jf == 0 and df == 0:
                        continue
                    opp_v = r[ci['opp']] if 0 <= ci['opp'] < len(r) else None
                    summ_v = r[ci['summary']] if 0 <= ci['summary'] < len(r) else None
                    vt = r[ci['vch_type']] if 0 <= ci['vch_type'] < len(r) else None
                    vn = r[ci['vch_no']] if 0 <= ci['vch_no'] < len(r) else None
                    _dk = (r[ci['date']] if 0 <= ci['date'] < len(r) else None,
                           vt, vn, str(nm).strip(), jf, df,
                           str(opp_v or ''), str(summ_v or ''))
                    _file2.append((_dk, r, jf, df))
                except Exception:
                    continue
            # 文件级"整表重复"配对检测（2026-08-02）：FY深圳 2026 导出每行×2 → 全部行键计数为偶数
            # 才判定整表复制、每键留 1 行；个别合法相同分录不误删（否则勾稽/明细 双计或缺失）。
            if len(_file2) >= 4:
                from collections import Counter as _C
                _cnt = _C(k for k, *_ in _file2)
                _whole_dup = all(v >= 2 and v % 2 == 0 for v in _cnt.values())
            _seen_keys = set()
            for _dk, r, jf, df in _file2:
                try:
                    if _whole_dup:
                        if _dk in _seen_keys:
                            continue
                        _seen_keys.add(_dk)
                    nm = r[ci['name']] if 0 <= ci['name'] < len(r) else None
                    opp_v = r[ci['opp']] if 0 <= ci['opp'] < len(r) else None
                    summ_v = r[ci['summary']] if 0 <= ci['summary'] < len(r) else None
                    vt = r[ci['vch_type']] if 0 <= ci['vch_type'] < len(r) else None
                    vn = r[ci['vch_no']] if 0 <= ci['vch_no'] < len(r) else None
                    subj = _subject_of(nm)
                    if subj is None:
                        continue
                    l2 = str(nm).strip()
                    # 利润表蓝字红字规则：非结转损益的非自然方蓝字→自然方红字，使发生额=结转损益口径
                    opp_v = r[ci['opp']] if 0 <= ci['opp'] < len(r) else None
                    summ_v = r[ci['summary']] if 0 <= ci['summary'] < len(r) else None
                    vt = r[ci['vch_type']] if 0 <= ci['vch_type'] < len(r) else None
                    vn = r[ci['vch_no']] if 0 <= ci['vch_no'] < len(r) else None
                    is_carry = (vt, vn) in carry_v
                    adj = apply_pl_carryover_rule(
                        [{'dr': jf, 'cr': df, 'name': nm, 'opp': opp_v, 'summary': summ_v,
                          '_carry': is_carry}],
                        _pl_kind(subj)
                    )[0]
                    gl_agg[(subj, e, l2, y)][0] += adj['dr']
                    gl_agg[(subj, e, l2, y)][1] += adj['cr']
                    # 年终结账行（整笔凭证为结转损益）不进入凭证抽查表：既非真实交易，也避免与审定表勾稽噪声
                    # 例外：资产处置损益/营业外收支/减值损失 须"全量抽凭"，用户要求抽凭合计=明细表合计，不跳过结转行
                    if is_carry and subj not in ('资产处置损益', '营业外支出', '营业外收入', '资产减值损失', '信用减值损失'):
                        continue
                    # 摘要补全：U8 结转账常只有首行有摘要，逐行抽凭时用同凭证首个非空摘要填充
                    if not summ_v or not str(summ_v).strip():
                        d0 = r[ci['date']] if 0 <= ci['date'] < len(r) else None
                        ym = str(d0)[:7] if d0 else ''
                        summ_v = vsum_map.get((ym, vt, vn), summ_v)
                    # 兜底：结转损益凭证（对方科目=本年利润等）U8 摘要列可能全空 → 用对方科目生成可读摘要
                    if (not summ_v or not str(summ_v).strip()) and opp_v and any(k in str(opp_v) for k in ('本年利润', '损益结转', '结转')):
                        summ_v = '结转损益（对方科目：%s）' % str(opp_v).strip()[:40]
                    # 再兜底：整笔凭证判为结转损益(年终结账)且摘要仍空 → 直接标注"结转损益"
                    if (not summ_v or not str(summ_v).strip()) and is_carry:
                        summ_v = '结转损益（年终结账）'
                    gl_rows[(subj, e, y)].append([
                        r[ci['date']], vt, vn,
                        opp_v, summ_v, adj['dr'], adj['cr'], str(nm).strip()
                    ])
                except Exception as ex:
                    print(f'  ⚠️ 跳过 GL 异常行 [{e}][{y}]：{ex}')
                    continue
            wb.close()
    return gl_agg, gl_rows


def _tb_gl_fallback(data_dir, entities, years, gl_agg, gl_rows):
    """无 GL（综合查询明细表）账套的 TB 发生额兜底——2026-08-06 新增。

    JTt 等账套只有《科目余额表》（+序时账/辅助余额表），无《综合查询明细表》：
    read_pl_gl 返回空 → 全部损益科目被判"无发生额"跳过 → 底稿缺失且 coverage 报缺。
    但 JTt 的 TB 损益科目（605107\其他业务收入\… 等 6 位叶子代码）借/贷发生额完整，
    收入贷方=真实收入、费用借方=真实费用，且结转行『XX结转』只影响对方方向。
    故从 TB 叶子行聚合（铁律3：rev 取贷 / exp 取借由 nat_amt 按 kind 决定，此处借贷都存），
    排除名称含『结转』的行（JTt 结转行：收入结转行 jf 有值、成本结转行 df 有值，
    会污染对侧方向合计）。
    """
    import audit_common as A
    tb_full = A.read_tb_full(data_dir, entities)
    pl_names = {name for (_c, name, _k) in PL_SUBJECTS}
    n_hit = 0
    n_subj = set()
    for (e, c, n, y), v in tb_full.items():
        nm = str(n or '')
        # 排除结转行（『XX结转』）与本年利润承接行（『4103xx\本年利润\XX』）——
        # JTt 结转行收入侧 jf 有值/成本侧 df 有值；4103 下承接行 df/jf 有值，
        # 不排除会与 6xxx 原始行双计。
        if '结转' in nm or '本年利润' in nm:
            continue
        subj = _subject_of(nm)
        if not subj or subj not in pl_names:
            continue
        jf = float(v.get('jf') or 0.0)
        df = float(v.get('df') or 0.0)
        if abs(jf) < 0.005 and abs(df) < 0.005:
            continue
        l2 = nm.split('\\')[-1].strip()
        key = (subj, e, l2, str(y))
        gl_agg[key][0] += jf
        gl_agg[key][1] += df
        n_hit += 1
        n_subj.add(subj)
    if n_hit:
        print(f'  ⚠️ [TB-only] 无 GL 文件，从科目余额表借贷发生额兜底 {n_hit} 条 '
              f'（科目：{"、".join(sorted(n_subj))}；收入取贷/费用取借，排除结转行）')
    return gl_agg, gl_rows


# ---------------- 读 TB：一级控制数 ----------------
def read_tb_control(data_dir, entities, years):
    # SAP：从 adapter.read_km 一级行取 (借发, 贷发)
    if _adapter is not None and _adapter.is_sap(data_dir):
        tb = {}
        for e, yd in entities.items():
            km = _adapter.read_km(e)
            for code, v in km.items():
                if v.get('level') != 1:
                    continue
                subj = _subject_of(str(v.get('name') or ''))
                if subj is None:
                    continue
                for y in years:
                    key = (e, subj, str(y))
                    dj, dl = tb.get(key, (0.0, 0.0))
                    tb[key] = (dj + float(v.get('debit') or 0.0), dl + float(v.get('credit') or 0.0))
        return tb
    tb = {}   # (entity, subject, year) -> (借发, 贷发)，仅取一级(4位)
    for e, yd in entities.items():
        for y, paths in yd.items():
            km = paths.get('km')
            if not km or not os.path.exists(km):
                continue
            try:
                wb = load_workbook(km, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ 跳过 TB 读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            for r in ws.iter_rows(min_row=3, values_only=True):
                try:
                    if not r or len(r) < 6 or not r[0]:
                        continue
                    cn = re.sub(r'\D', '', str(r[0]))
                    nm = r[1] if len(r) > 1 else ''
                    subj = _subject_of(nm)
                    if subj is None:
                        continue
                    if len(cn) == 4:   # 仅取一级控制数（避免二级重复累计）
                        jf = safe(r[4]) if len(r) > 4 else 0.0
                        df = safe(r[5]) if len(r) > 5 else 0.0
                        tb[(e, subj, y)] = (jf, df)
                except Exception as ex:
                    print(f'  ⚠️ 跳过 TB 异常行 [{e}][{y}]：{ex}')
                    continue
            wb.close()
    return tb


# ---------------- 勾稽判定 ----------------
# 本项目 2026 综合查询明细表(GL)仅含 1–5 月(YTD)，而 2026 科目余额表为全年；
# 故 2026 明细表的自然方合计与 TB 全年控制数不可比，单独标注而非误报“不一致”。
YTD_YEARS = {'2026'}


def _check_recon(name, entity, year, gl_total, tb):
    if entity is None:
        ctrl = sum(v[0] for (e, s, y), v in tb.items() if s == name and y == year)
        if ctrl == 0.0:
            return '无控制数'
    else:
        v = tb.get((entity, name, year))
        if v is None:
            return '无控制数'
        ctrl = v[0]
    if year in YTD_YEARS:
        return 'YTD/全年不可比'
    diff = gl_total - ctrl
    if abs(diff) <= max(1.0, abs(ctrl) * 1e-6):
        return '一致'
    return f'不一致(差{diff:,.2f})'


# ============ 勾稽辅助：应交税费 & 减值准备 GL 聚合 ============
# 资产减值损失(6701) 仅与「非信用/非递延所得税」的资产减值准备对账；
# 坏账准备 属 信用减值损失(6702)、递延所得税资产/负债 属 所得税费用(6801)，故在此排除。
RESERVE_FIRST_SEGMENTS = {
    '坏账准备', '存货跌价准备', '固定资产减值准备', '无形资产减值准备',
    '长期股权投资减值准备', '商誉减值准备', '持有至到期投资减值准备',
    '在建工程减值准备', '投资性房地产减值准备', '生产性生物资产减值准备',
    '使用权资产减值准备', '合同资产减值准备', '长期应收款减值准备',
    '贷款损失准备', '抵债资产减值准备', '损余物资跌价准备',
    '递延所得税资产', '递延所得税负债',
}
IMPAIR_EXCLUDE_SEG = {'坏账准备', '递延所得税资产', '递延所得税负债'}


def _is_reserve_name(name):
    """判定 GL 科目名首段是否为资产减值准备类科目。"""
    s0 = str(name).split('-')[0].strip()
    if s0 in RESERVE_FIRST_SEGMENTS:
        return True
    if '减值准备' in s0 or '跌价准备' in s0:
        return True
    return False


def read_recon_gl(data_dir, entities, years):
    """聚合 GL 中『应交税费』与『各项减值准备』科目的 (借,贷)，按 (entity, full_name, year)。

    用于：① 税金及附加 ↔ 应交税费计提；② 资产减值损失 ↔ 相关资产减值准备 的勾稽。
    两笔勾稽均为 GL 内部对账（同一年度 GL 口径），2026 虽为 YTD(1-5月) 但两侧同源可比，
    不受现有『GL-vs-TB』YTD 不可比限制影响。"""
    # SAP：直接聚合 adapter 行（应交税费/减值准备名称匹配）
    if _adapter is not None and _adapter.is_sap(data_dir):
        agg = defaultdict(lambda: [0.0, 0.0])
        for e, yd in entities.items():
            for r in _adapter.read_gl(e):
                nm = str(r.get('name') or '')
                s = nm.strip()
                s0 = s.split('-')[0].strip()
                if not (s0 == '应交税费' or s.startswith('应交税费')
                        or _is_reserve_name(s)
                        or s0 == '递延收益' or s.startswith('递延收益')):
                    continue
                key = (e, s, str(r.get('year') or ''))
                agg[key][0] += float(r.get('debit') or 0.0)
                agg[key][1] += float(r.get('credit') or 0.0)
        return agg
    agg = defaultdict(lambda: [0.0, 0.0])
    for e, yd in entities.items():
        for y, paths in yd.items():
            g = paths.get('gl')
            if not g or not os.path.exists(g):
                continue
            try:
                wb = load_workbook(g, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ 跳过勾稽GL读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            ci = detect_gl_columns(ws)
            mc = ws.max_column
            for k in list(ci.keys()):
                if ci[k] < 0:
                    ci[k] = 0
                if ci[k] >= mc:
                    ci[k] = max(0, mc - 1)
            _whole_dup = False
            _file3 = []
            for r in ws.iter_rows(min_row=3, values_only=True):
                if not r or len(r) <= ci['credit']:
                    continue
                nm = r[ci['name']] if 0 <= ci['name'] < len(r) else None
                if not nm:
                    continue
                s = str(nm).strip()
                s0 = s.split('-')[0].strip()
                if not (s0 == '应交税费' or s.startswith('应交税费') or _is_reserve_name(s)
                        or s0 == '递延收益' or s.startswith('递延收益')):
                    continue
                jf = safe(r[ci['debit']])
                df = safe(r[ci['credit']])
                if jf == 0 and df == 0:
                    continue
                _dk = (r[ci['date']] if 0 <= ci['date'] < len(r) else None,
                       r[ci['vch_type']] if 0 <= ci['vch_type'] < len(r) else None,
                       r[ci['vch_no']] if 0 <= ci['vch_no'] < len(r) else None,
                       s, jf, df,
                       str(r[ci['opp']] if 0 <= ci['opp'] < len(r) else None or ''),
                       str(r[ci['summary']] if 0 <= ci['summary'] < len(r) else None or ''))
                _file3.append((_dk, s, jf, df))
            # 文件级"整表重复"配对检测（2026-08-02）：FY深圳 2026 导出每行×2 → 全部行键计数为偶数
            # 才判定整表复制、每键留 1 行；个别合法相同分录不误删（否则 recon_gl 双计与去重后
            # gl_full 剥离口径不一致 → 所得税勾稽恒差）。
            if len(_file3) >= 4:
                from collections import Counter as _C
                _cnt = _C(k for k, *_ in _file3)
                _whole_dup = all(v >= 2 and v % 2 == 0 for v in _cnt.values())
            _seen_keys = set()
            for _dk, s, jf, df in _file3:
                if _whole_dup:
                    if _dk in _seen_keys:
                        continue
                    _seen_keys.add(_dk)
                agg[(e, s, y)][0] += jf
                agg[(e, s, y)][1] += df
            wb.close()
    return agg


# ---------------- 读 TB：递延所得税资产/负债 期初/期末余额 ----------------
def read_deferred_tb(entities):
    """从各主体《科目余额表》读取 递延所得税资产(DTA,1811)/递延所得税负债(DTL,2901)
    的期初/期末(未带符号金额)，返回 {(entity, year): {'递延所得税资产':(begin,end),
    '递延所得税负债':(begin,end)}}。

    用途：余额表法勾稽 递延所得税费用 = ΔDTL − ΔDTA（用未带符号的期初/期末金额）。
    取数策略：优先取 4 位父级代码（如 1811/2901，其期初/期末已是子目汇总）；
    若无 4 位父级（仅导出了子目），则汇总所有同名首段子目。"""
    # SAP：从 adapter.read_km 按名称取 递延所得税资产/负债 期初期末（abs）
    if _adapter is not None and _adapter.is_sap(entities and next(iter(entities), None) and ''):
        res = {}
        for e, yd in entities.items():
            km = _adapter.read_km(e)
            dta_b = dta_e = dtl_b = dtl_e = 0.0
            for code, v in km.items():
                nm = str(v.get('name') or '')
                if '递延所得税资产' in nm:
                    dta_b += abs(float(v.get('opening') or 0.0))
                    dta_e += abs(float(v.get('closing') or 0.0))
                elif '递延所得税负债' in nm:
                    dtl_b += abs(float(v.get('opening') or 0.0))
                    dtl_e += abs(float(v.get('closing') or 0.0))
            for y in {yy for yd2 in entities.values() for yy in yd2}:
                res[(e, y)] = {'递延所得税资产': (dta_b, dta_e), '递延所得税负债': (dtl_b, dtl_e)}
        return res
    res = {}
    for e, yd in entities.items():
        for y, paths in yd.items():
            km = paths.get('km')
            if not km or not os.path.exists(km):
                continue
            parent = {'递延所得税资产': None, '递延所得税负债': None}   # 4位父级 -> (begin,end)
            subs = {'递延所得税资产': [], '递延所得税负债': []}        # 子目 -> [(begin,end)]
            try:
                wb = load_workbook(km, data_only=True, read_only=True)
            except Exception:
                continue
            ws = wb[wb.sheetnames[0]]
            for r in ws.iter_rows(min_row=3, values_only=True):
                if not r or len(r) < 8 or not r[0]:
                    continue
                nm = r[1] if len(r) > 1 else ''
                if not nm:
                    continue
                s0 = str(nm).split('-')[0].strip()
                if s0 not in ('递延所得税资产', '递延所得税负债'):
                    continue
                cdig = len(re.sub(r'\D', '', str(r[0])))
                begin = safe(r[3]) if len(r) > 3 and r[3] is not None else 0.0
                end = safe(r[7]) if len(r) > 7 and r[7] is not None else 0.0
                if cdig == 4:
                    parent[s0] = (begin, end)
                else:
                    subs[s0].append((begin, end))
            wb.close()
            entry = {}
            for k in ('递延所得税资产', '递延所得税负债'):
                if parent[k] is not None:
                    entry[k] = parent[k]
                elif subs[k]:
                    entry[k] = (sum(b for b, _ in subs[k]), sum(x for _, x in subs[k]))
            if entry:
                res[(e, y)] = entry
    return res


# ---------------- 所得税费用 勾稽：当期 vs 应交所得税计提；递延 vs (ΔDTL−ΔDTA) ----------------
# 2026-08-02 用户方法论升级（凭证级三块拆分）：所得税费用按凭证拆 当期/递延/以前年度，
# 使勾稽与抽凭精准化——以前年度调整（如"2024年审计调整#补提递延所得税"直接借记 DTA/DTL、
# 未过所得税费用科目）从勾稽B 右（ΔDTL−ΔDTA）中剥离，剥离后勾稽一致的主体免抽凭。
def _prior_year_kw(s):
    """摘要/科目名是否属「以前年度调整」类（补提/审计调整/上年/以前年度）。
    ⚡ 2026-08-28 加『补缴』——AH 补缴以前年度所得税（摘要『补缴23年企业所得税』
    '23年' 中文不命中 2023）此前未剥离，混入当期计提勾稽。"""
    if not s:
        return False
    ss = str(s)
    return any(k in ss for k in ('以前年度', '上年', '审计调整', '补提', '汇算清缴',
                                  '补缴', '2024', '2023', '2022', '2025'))


def _tax_payflow_kw(s):
    """摘要是否属「税负流出/冲减」类（退回/退还/退税/缴纳/清缴/付款）——非当期所得税
    计提、不应计入勾稽A 右侧应交所得税贷方。⚡ 2026-08-28 AH 1010『收深交所退回QF税费
    及境外机构税费』红字贷方 102 万此前被当计提（SAP 负数统一按贷方），致勾稽A 虚差。"""
    if not s:
        return False
    ss = str(s)
    return any(k in ss for k in ('退回', '退还', '退税', '返还', '缴纳', '清缴',
                                 '付', '退款', '退税款'))


def _withhold_kw(s):
    """摘要是否属「代扣代缴」类——非公司自身所得税（如分红派息时代扣的股息红利所得税，
    公司仅作扣缴义务人，贷方挂在 应交税费-应交企业所得税 下但不属当期费用计提勾稽范围）。"""
    if not s:
        return False
    ss = str(s)
    return any(k in ss for k in ('代扣', '代缴', '证券登记', '结算', '分红', '股息', '扣缴'))


def compute_income_tax_recon(entities, years, gl_rows, recon_gl, deferred_tb, gl_agg=None, name='所得税费用', kind='exp', gl_full=None):
    """逐(主体,年)做两笔勾稽：

    勾稽A：所得税费用-当期(借, 对方科目含『所得税』且非递延/非个人)
           ↔ 应交税费-应交所得税(贷发计提)。
    勾稽B：递延所得税费用(GL净 = Σ(借−贷)对方含『递延所得税』)
           ↔ ΔDTL − ΔDTA − 以前年度调整(余额表法, 取 TB 未带符号期初/期末)。

    2026-08-02 凭证级三块拆分（用户方法论）：所得税费用按凭证拆 当期/递延/以前年度——
    DTA/DTL 凭证中摘要含『以前年度/审计调整/补提/上年』且对方科目非所得税费用的行，
    属"以前年度调整"（直接借记 DTA/DTL、未过损益科目），从勾稽B 右 ΔDTL−ΔDTA 剥离，
    使勾稽B 只核对【当年递延所得税费用】口径；剥离后勾稽一致的主体免抽凭。

    返回 (rows, need_voucher, voucher_ey)：
      · need_voucher=True 表示任一年度无法勾稽一致(或为2026 YTD口径不可比)，
        此时应生成凭证抽查表；一致则不生成。"""
    recon_gl = recon_gl or {}
    rows = []
    need_voucher = False
    voucher_ey = set()    # 需生成凭证抽查表的 (核算主体, 年度) 集合（仅勾稽不一致者）

    def _tol(x):
        return max(1.0, abs(x) * 1e-6)

    # 预聚合：全量 GL(gl_full) 中 DTA/DTL 行的「以前年度调整」净额（(e,y) -> 净额）。
    # gl_full 行格式来自 audit_common.read_gl_rows：dict 含 name/opp/summary/dr/cr/e/y。
    prior_by_ey = {}
    if gl_full:
        for _r in gl_full:
            _nm = str(_r.get('name') or '')
            if not _nm.startswith('递延所得税'):
                continue
            if not _prior_year_kw(_r.get('summary') or _r.get('opp') or ''):
                continue
            _opp = str(_r.get('opp') or '')
            if '所得税费用' in _opp:
                continue      # 对方=所得税费用 → 已含在 def_gl_net（当年递延费用），非剥离项
            _ey = (_r.get('e'), _r.get('y'))
            prior_by_ey[_ey] = prior_by_ey.get(_ey, 0.0) + (float(_r.get('dr') or 0.0) - float(_r.get('cr') or 0.0))

    for e in sorted(entities):
        for y in sorted(years):
            # 勾稽A 左：所得税费用-当期 + 所得税费用-以前年度（取自明细表 gl_agg 二级科目）
            # 勾稽B 左：所得税费用-递延（取自明细表 gl_agg 二级科目）
            cur_gl = 0.0
            cur_gl_prior = 0.0    # 2026-08-02：二级名"当期"但摘要属以前年度（如"计提2022-2024年度
                                 # 企业所得税"挂在 6801.01 当期二级下）→ 归以前年度，从当期剥离
            def_gl_net = 0.0
            for (sb, en, l2, yr), v in (gl_agg or {}).items():
                if sb == name and en == e and yr == y:
                    a = nat_amt(kind, *v)
                    # l2=完整科目名。标准账套有二级（当期/以前年度/递延）；FY 等账套仅一级
                    # （l2==name）→ 整笔计"当期"（无递延二级，全部视为当期所得税费用）
                    if '递延' in l2:
                        def_gl_net += a
                    elif '以前年度' in l2 or '当期' in l2 or l2 == name:
                        cur_gl += a
            # 从 gl_rows 中按摘要剥离"二级当期但摘要以前年度"的行（凭证级三块拆分的一部分）：
            # gl_rows[(name,e,y)] 行格式 [date,vt,vn,opp,summ,dr,cr,name]；dr/cr 为
            # apply_pl_carryover_rule 后自然方（exp 取借）。仅当该行科目名无"递延"且摘要属以前年度。
            for _r in gl_rows.get((name, e, y), []):
                _nm = str(_r[7]) if len(_r) > 7 else ''
                if '递延' in _nm:
                    continue
                if not _prior_year_kw(_r[4]) and not _withhold_kw(_r[4]) and not _tax_payflow_kw(_r[4]):
                    continue
                _amt = float(_r[5] or 0.0) - float(_r[6] or 0.0)   # 借−贷（exp 自然方≈借）
                if _amt != 0.0:
                    cur_gl_prior += _amt
            cur_gl_cur = cur_gl - cur_gl_prior
            # 勾稽A 右：应交所得税计提(GL贷方)——2026-08-02 凭证级拆分：剔除「以前年度」类
            # 计提（摘要含 2022/2023/2024/以前年度/汇算清缴/补缴/退还 等），使勾稽A 只对
            # 【当期】应交所得税计提口径（Z母公司 应交企业所得税贷方 30.47M 中混入以前年度
            # 汇算清缴退还 4.9M、计提2023/2024年度 1.9M 等 7.2M，若不剔除与当期费用 28.89M
            # 恒差 1.58M → 只能全抽）。剥离后在 gl_rows 无对应（recon_gl 无摘要），故对
            # recon_gl 无法直接按摘要过滤——改由 gl_full 凭证级识别：凡含『应交所得税』且
            # 摘要属以前年度的贷方行，从 accr 中扣除。
            accr = 0.0
            accr_prior = 0.0
            for (ee, s, yy), (jb, dfb) in recon_gl.items():
                if ee == e and yy == y and str(s).startswith('应交税费'):
                    ss = str(s)
                    if '所得税' in ss and '递延' not in ss and '个人' not in ss and '代扣' not in ss:
                        accr += dfb
            if gl_full:
                for _r in gl_full:
                    if (_r.get('e'), _r.get('y')) != (e, y):
                        continue
                    _nm = str(_r.get('name') or '')
                    if not (_nm.startswith('应交税费') and '所得税' in _nm
                            and '递延' not in _nm and '个人' not in _nm):
                        continue
                    _cr = float(_r.get('cr') or 0.0)
                    if _cr == 0.0:
                        continue
                    _summ = str(_r.get('summary') or '')
                    if _prior_year_kw(_summ) or _withhold_kw(_summ) or _tax_payflow_kw(_summ):
                        accr_prior += _cr   # 以前年度类/代扣代缴类/退回缴税类计提：从 accr 中剥离
            accr_cur = accr - accr_prior
            # 勾稽B 右：ΔDTL − ΔDTA（余额表法）
            dt = deferred_tb.get((e, y), {})
            dta = dt.get('递延所得税资产'); dtl = dt.get('递延所得税负债')
            begin_dta, end_dta = dta if dta else (0.0, 0.0)
            begin_dtl, end_dtl = dtl if dtl else (0.0, 0.0)
            dDTA = end_dta - begin_dta
            dDTL = end_dtl - begin_dtl
            b_right = dDTL - dDTA
            # 2026-08-02 凭证级三块拆分：从全量 GL(gl_full) 预聚合的 DTA/DTL「以前年度调整」
            # 净额中取本主体本年剥离额（摘要含 以前年度/审计调整/补提/上年 且对方非所得税费用
            # 的直接借记 DTA/DTL 调整，如"2024年审计调整#补提递延所得税资产"未过损益科目），
            # 从勾稽B 右 ΔDTL−ΔDTA 剥离，使勾稽B 只核对【当年递延所得税费用】口径。
            # 符号：prior_adj = Σ(借−贷)（借增 DTA/DTL 为正）。ΔDTA/ΔDTL 已含该调整额，
            # 从变动中扣除 = b_right − prior_adj 应改为 b_right + prior_adj——
            # 验证（Z母公司）：ΔDTA=5,429,154.62、prior=+263,072.06（借补提DTA）→
            # b_right_cur = -5,429,154.62 + 263,072.06 = -5,166,082.56 = 所得税费用-递延 ✓。
            prior_adj = prior_by_ey.get((e, y), 0.0)
            b_right_cur = b_right + prior_adj
            # 判定
            a_diff = cur_gl_cur - accr_cur
            has_a = (abs(cur_gl_cur) > 1e-6 or abs(accr_cur) > 1e-6)
            a_ok = has_a and abs(a_diff) <= _tol(accr_cur)
            a_status = '一致' if a_ok else (f'不一致(差{a_diff:,.2f})' if has_a else '无数据')
            b_diff = def_gl_net - b_right_cur
            b_has = (abs(def_gl_net) > 1e-6 or abs(b_right_cur) > 1e-6)
            if y in YTD_YEARS:
                b_status = 'YTD/全年不可比'
                b_ok = False
            else:
                b_ok = b_has and abs(b_diff) <= _tol(b_right_cur)
                b_status = '一致' if b_ok else (f'不一致(差{b_diff:,.2f})' if b_has else '无数据')
            # 凭证抽查表决策（逐 核算主体/年度：仅勾稽不一致者纳入抽凭）
            ey_need = False
            if has_a and not a_ok:
                ey_need = True
            if b_has:
                if y in YTD_YEARS:
                    ey_need = True
                elif not b_ok:
                    ey_need = True
            if ey_need:
                voucher_ey.add((e, y))
            need_voucher = need_voucher or ey_need
            rows.append({
                'e': e, 'y': y,
                'cur_gl': cur_gl_cur, 'cur_gl_prior': cur_gl_prior,
                'accr': accr_cur, 'accr_prior': accr_prior,
                'a_diff': a_diff, 'a_status': a_status,
                'def_gl': def_gl_net, 'dDTL': dDTL, 'dDTA': dDTA, 'b_right': b_right,
                'prior_adj': prior_adj, 'b_right_cur': b_right_cur,
                'b_diff': b_diff, 'b_status': b_status,
                'has_a': has_a, 'b_has': b_has,
            })
    return rows, need_voucher, voucher_ey


def write_income_tax_recon_sheet(ws, rows, need_voucher, entities, years):
    ws.cell(1, 1, '所得税费用勾稽核对表（所得税费用-当期 ↔ 应交所得税当期计提；所得税费用-递延 ↔ ΔDTL−ΔDTA−以前年度调整）').font = Font(name='Times New Roman', bold=True, size=10, color='1F4E78')
    ws.cell(2, 1, '勾稽A：所得税费用-当期(借) 应 = 应交税费-应交所得税当期计提(贷发)；两侧均剔除「以前年度」类——费用侧：'
                  '二级名"当期"但摘要含 以前年度/2022-2024/汇算清缴 的凭证剥离；计提侧：应交所得税贷方中摘要属以前年度的计提剥离。'
                  '勾稽B：所得税费用-递延(借) 应 = Δ递延所得税负债 − Δ递延所得税资产 − 以前年度调整(DTA/DTL 中摘要含'
                  '以前年度/审计调整/补提且对方非所得税费用的直接调整)。2026-08-02 凭证级三块拆分。'
                  '两项均勾稽一致 → 不生成凭证抽查表；任一项无法勾稽一致(或2026 YTD口径不可比) → 仅对该核算主体/年度生成凭证抽查表(见同簿"凭证抽查表"页)。').font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    HEADERS = ['核算主体', '年度',
               '当期费用(借)', '以前年度费用(剥离)', '以前年度计提(剥离)',
               '当期计提(贷)', '差额', '是否一致',
               '递延费用(借)', 'Δ递延所得税负债', 'Δ递延所得税资产',
               '以前年度调整(剥离)', 'ΔDTL−ΔDTA−调整(应)', '差额', '是否一致',
               '综合结论']
    # ⚡ 2026-08-11 N1：双行表头分组（规则2）——行3 大类（勾稽A/勾稽B），行4 列名去前缀
    _grp_row = 3
    for _g, _c0, _c1 in [('基础信息', 1, 2), ('勾稽A：当期所得税费用 ↔ 应交所得税当期计提', 3, 8),
                         ('勾稽B：递延所得税费用 ↔ ΔDTL−ΔDTA−以前年度调整', 9, 15), ('结论', 16, 16)]:
        ws.merge_cells(start_row=_grp_row, start_column=_c0, end_row=_grp_row, end_column=_c1)
        _gc = ws.cell(_grp_row, _c0, _g)
        _gc.font = HDR_FONT; _gc.fill = HDR_FILL; _gc.alignment = CENTER; _gc.border = BORDER
        for _j in range(_c0, _c1 + 1):
            ws.cell(_grp_row, _j).border = BORDER
    hr = 4
    for c, h in enumerate(HEADERS, 1):
        cell = ws.cell(hr, c, h)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = BORDER
        cell.alignment = CENTER
    ws.freeze_panes = 'A5'
    r = hr + 1
    sc = [0.0] * 13   # 汇总 勾稽A当期/prior费/prior计/右/差, 勾稽B左/dDTL/dDTA/prior/应/差
    for row in rows:
        vals = [row['e'], row['y'],
                row['cur_gl'], row['cur_gl_prior'], row['accr_prior'], row['accr'], row['a_diff'], row['a_status'],
                row['def_gl'], row['dDTL'], row['dDTA'], row['prior_adj'], row['b_right_cur'],
                row['b_diff'], row['b_status'],
                ('一致' if (row['a_status'] == '一致' and row['b_status'] == '一致') else '未一致→已抽凭' if (row['has_a'] or row['b_has']) else '无数据')]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (1, 2, 8, 15, 16):
                cell.alignment = LEFT
            elif c in (3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14):
                cell.number_format = NUM_FMT
                cell.alignment = RIGHT
            else:
                cell.alignment = CENTER
            # ⚡ 2026-08-11 N6：是否一致 条件格式——一致绿 / 不一致红
            if c == 8:
                cell.font = Font(name='Times New Roman', color='006100') if row['a_status'] == '一致' \
                    else (Font(name='Times New Roman', color='C00000') if row['a_status'] != '一致' else None)
            if c == 15:
                cell.font = Font(name='Times New Roman', color='006100') if row['b_status'] == '一致' \
                    else (Font(name='Times New Roman', color='C00000') if row['b_status'] != '一致' else None)
        sc[0] += row['cur_gl']; sc[1] += row['cur_gl_prior']; sc[2] += row['accr_prior']
        sc[3] += row['accr']; sc[4] += row['a_diff']
        sc[5] += row['def_gl']; sc[6] += row['dDTL']; sc[7] += row['dDTA']
        sc[8] += row['prior_adj']; sc[9] += row['b_right_cur']; sc[10] += row['b_diff']
        r += 1
    # 集团合计（⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出）
    if len(entities) > 1:
        tot_vals = ['集团合计', '', sc[0], sc[1], sc[2], sc[3], sc[4], '', sc[5], sc[6], sc[7], sc[8], sc[9], sc[10], '', '']
        for c, v in enumerate(tot_vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = TOT_FONT
            cell.fill = TOT_FILL
            cell.border = BORDER
            if c in (3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14):
                cell.number_format = NUM_FMT
                cell.alignment = RIGHT
    r += 2
    # 结论
    if need_voucher:
        concl = '结论：存在任一年度勾稽未能一致(或2026为YTD口径不可比)，已生成"凭证抽查表"页，供对差异逐笔抽查。'
    else:
        concl = '结论：勾稽A(当期所得税费用↔应交所得税当期计提) 与 勾稽B(递延所得税费用↔ΔDTL−ΔDTA−以前年度调整) 全部一致，依规则不生成凭证抽查表。'
    ws.cell(r, 1, concl).font = Font(name='Times New Roman', bold=True, size=10, color='C00000' if need_voucher else '006100')
    widths = [16, 8, 16, 18, 18, 18, 14, 14, 18, 16, 16, 18, 20, 14, 14, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------- 样式辅助 ----------------
def _style_data(ws, r, num_cols):
    for c in range(1, DETAIL_COLS + 1):
        cell = ws.cell(r, c)
        cell.border = BORDER
        if c <= 2:
            cell.alignment = LEFT
        elif c in num_cols:
            cell.alignment = RIGHT
            cell.number_format = NUM_FMT
        else:
            cell.alignment = CENTER


def _style_total(ws, r, num_cols):
    for c in range(1, DETAIL_COLS + 1):
        cell = ws.cell(r, c)
        cell.border = BORDER
        cell.font = TOT_FONT
        cell.fill = TOT_FILL
        if c <= 2:
            cell.alignment = LEFT
        elif c in num_cols:
            cell.alignment = RIGHT
            cell.number_format = NUM_FMT
        else:
            cell.alignment = CENTER


# ---------------- 写「<年份>明细表」 ----------------
def write_detail_sheet(ws, name, kind, year, prev_year, entities, years, gl_agg, tb, recon_gl=None):
    # 显式写入明细表表头（单一事实来源 = DETAIL_HEADERS），不依赖模板克隆，
    # 保证新增「对应科目金额」列后表头始终正确：顺序为 对方科目(8) → 对应科目金额(9) → 是否勾稽一致(10)，
    # 即先有金额、再判勾稽一致。
    for c, h in enumerate(DETAIL_HEADERS, 1):
        cell = ws.cell(1, c, h)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = BORDER
        cell.alignment = CENTER
    ws.column_dimensions['I'].width = DETAIL_SHELL_WIDTHS[8]   # 对应科目金额
    ws.column_dimensions['J'].width = DETAIL_SHELL_WIDTHS[9]   # 是否勾稽一致
    for r in range(2, ws.max_row + 1):
        for c in range(1, DETAIL_COLS + 1):
            ws.cell(r, c).value = None
    rows_plan = []
    # 税金及附加：预构建「应交税费相关子目计提」查表（按 主体/年/规范税种），
    # 供明细表逐 L2 行填列「对方科目(应交税费子目)」并做金额勾稽（GL内部对账，不受 YTD 限制）。
    taxpay_amt = {}          # (e, y, k) -> [贷方计提额, 原始应交税费子目名]
    taxpay_base = defaultdict(float)   # (e, y, 基础税种) -> 贷方计提合计(汇总体所有属地后缀), 供属地不一致时回退
    if name == '税金及附加' and recon_gl:
        for (e, sub, y), v in recon_gl.items():
            s0 = str(sub).split('-')[0].strip()
            if s0 != '应交税费':
                continue
            if any(x in sub for x in TAXPAY_EXCLUDE):
                continue
            k = _norm_taxpay_sub(sub)
            rec = taxpay_amt.setdefault((e, y, k), [0.0, None])
            rec[0] += v[1]   # 贷方=计提
            rec[1] = sub      # 保留原始子目名作对方科目显示
            taxpay_base[(e, y, _norm_tax(sub, keep_loc=False))] += v[1]   # 基础税种汇总(剥属地后缀)
    # 其他收益：预构建「递延收益摊销(借发)」查表（按 主体/年），供明细表逐行填列
    # 「对方科目=递延收益(摊销转入)」并标注与递延收益摊销额的勾稽状态（GL内部对账，不受 YTD 限制）。
    oth_amt = defaultdict(float)     # (e, y) -> 其他收益贷发合计
    def_amt = defaultdict(float)     # (e, y) -> 递延收益借发(摊销)合计
    oth_ystatus = {}                 # (e, y) -> '一致'/'不一致'
    oth_gstatus = {}                 # (y,) -> '一致'/'不一致'（全集团）
    if name == '其他收益' and recon_gl is not None:
        for (sb, en, l2, yr), v in gl_agg.items():
            if sb == name:
                oth_amt[(en, yr)] += v[1]   # rev: 取贷方
        for (e, subn, yr), v in recon_gl.items():
            s0 = str(subn).split('-')[0].strip()
            if s0 == '递延收益' or str(subn).startswith('递延收益'):
                def_amt[(e, yr)] += v[0]    # 借发 = 摊销转出
        for (e, yr), o in oth_amt.items():
            d = def_amt.get((e, yr), 0.0)
            # Q2 修复：纯直接计入政府补助(无递延摊销)时 d=0，本就不存在『其他收益 vs 递延摊销』勾稽，
            # 不应判为『不一致』；仅在存在递延摊销(d>0)且与其他收益总额不符时才报不一致。
            oth_ystatus[(e, yr)] = ('直接计入(无需递延勾稽)'
                                    if abs(d) < 0.005
                                    else _recon_status(o - d, o, d))
        for yr in years:
            o = sum(v for (k, v) in oth_amt.items() if k[1] == yr)
            d = sum(v for (k, v) in def_amt.items() if k[1] == yr)
            oth_gstatus[yr] = ('直接计入(无需递延勾稽)'
                               if abs(d) < 0.005
                               else _recon_status(o - d, o, d))
    # 预计算「对应科目金额」按主体/全集团合计，供明细表逐行在「是否勾稽一致」前列示金额
    # （先有金额再判断是否勾稽一致）。仅本年(year)有效。
    tp_by_entity = defaultdict(float)   # e -> 应交税费计提合计(本年)
    for (e, y, k), rec in taxpay_amt.items():
        if y == year:
            tp_by_entity[e] += rec[0]
    tp_grand = sum(tp_by_entity.values())
    def_by_entity = defaultdict(float)   # e -> 递延收益摊销合计(本年)
    for (e, yr), v in def_amt.items():
        if yr == year:
            def_by_entity[e] += v
    def_grand = sum(def_by_entity.values())
    # 减值损失：预构建「对应减值准备」查表（按 主体/年/类别seg），供明细表逐 L2 行
    # 填列「对方科目(减值准备科目) + 对应科目金额(减值准备贷计提) + 是否勾稽一致」，
    # 即把『减值损失 ↔ 减值准备』勾稽直接嵌入明细表（取代原底部独立勾稽块）。
    imp_reserve = defaultdict(lambda: [0.0, 0.0, set()])   # (e, y, seg) -> [dr(借转回), cr(贷计提), raw]
    imp_tokens = imp_strip = imp_family = None
    if name in IMPAIR_PARAMS and recon_gl is not None:
        p = IMPAIR_PARAMS[name]
        imp_tokens, imp_strip, imp_family = p['tokens'], p['expense_strip'], p['reserve_family']
        for (e, rname, y), v in recon_gl.items():
            if imp_family == 'credit':
                # 仅对真正的坏账准备类科目对账；排除名称含"应收账款"但实为应交税费/重分类的科目
                if '坏账' not in rname and '信用减值' not in rname:
                    continue
                seg = _norm_asset_key(rname, ('坏账准备',), CREDIT_ASSET_TOKENS)
                if seg is None:
                    continue
            else:
                if _norm_asset_key(rname, ('坏账准备',), CREDIT_ASSET_TOKENS) is not None:
                    continue  # 坏账准备族属信用减值损失，资产减值损失侧剔除
                if not any(t in rname for t in ('准备', '减值', '跌价')):
                    continue  # 仅对真正的资产减值准备/跌价准备对账
                seg = _norm_asset_key(rname, (), ASSET_IMPAIR_TOKENS)
                if seg is None:
                    continue
            imp_reserve[(e, y, seg)][0] += v[0]
            imp_reserve[(e, y, seg)][1] += v[1]
            imp_reserve[(e, y, seg)][2].add(rname)
    all_l2 = set()          # 全集团 L2 全集 = 各主体 本年∪上年（保证跨期连续性）
    for e in sorted(entities):
        # L2 全集 = 本年 ∪ 上年（保证跨期连续性：上年有/本年无 的行也要列示）
        l2set = set()
        for (sb, en, l2, yr) in gl_agg:
            if sb == name and en == e and yr in (year, prev_year):
                l2set.add(l2)
        if not l2set:
            continue
        plan = []
        for l2 in sorted(l2set):
            key_y = (name, e, l2, year)
            amt = nat_amt(kind, *gl_agg[key_y]) if key_y in gl_agg else 0.0
            prev = None
            if prev_year is not None:
                key_p = (name, e, l2, prev_year)
                prev = nat_amt(kind, *gl_agg[key_p]) if key_p in gl_agg else 0.0
            plan.append((l2, amt, prev))
        rows_plan.append((e, plan))
        all_l2 |= l2set
    if not rows_plan:
        ws.cell(2, 1, f'（{name} 在 {year} 无综合查询明细表发生额数据）')
        return
    r = 2
    tot_cur = 0.0
    tot_prev = 0.0
    grp_l2_cur = defaultdict(float)    # 全集团按二级科目本年合计数
    grp_l2_prev = defaultdict(float)   # 全集团按二级科目上年合计数
    grp_tp = 0.0                       # 全集团 对应科目金额(匹配合计)
    for e, plan in rows_plan:
        e_cur = 0.0
        e_prev = 0.0
        e_tp_local = 0.0               # 主体内 对应科目金额(匹配合计)
        e_rc_total = 0.0               # 主体内 减值准备贷计提合计（减值损失嵌入用）
        e_rd_total = 0.0               # 主体内 减值准备借方转回合计
        for (l2, amt, prev) in plan:
            # 铁律④：先计入合计/汇总（不影响勾稽），再判断是否跳过显示
            e_cur += amt
            e_prev += (prev or 0.0)
            grp_l2_cur[l2] += amt
            grp_l2_prev[l2] += (prev or 0.0)
            if prev_year is None:
                if is_zero_amount(amt):
                    continue
            else:
                if is_zero_amount(amt, prev):
                    continue
            chg = (amt - prev) if prev is not None else None
            rate = round(chg / prev * 100, 2) if (prev is not None and abs(prev) > 0.005) else None
            ws.cell(r, 1, e)
            ws.cell(r, 2, l2)
            ws.cell(r, 3, prev)
            ws.cell(r, 4, amt)
            ws.cell(r, 5, chg)
            ws.cell(r, 6, rate)
            _style_data(ws, r, [3, 4, 5, 6])
            r += 1
        e_chg = (e_cur - e_prev) if prev_year is not None else None
        # 铁律④：主体小计全为 0 不显示
        ent_zero = is_zero_amount(e_cur) if prev_year is None else is_zero_amount(e_cur, e_prev)
        if not ent_zero:
            e_rate = round(e_chg / e_prev * 100, 2) if (e_prev is not None and abs(e_prev) > 0.005) else None
            ws.cell(r, 1, e)
            ws.cell(r, 2, '(主体小计)')
            ws.cell(r, 3, e_prev if prev_year is not None else None)
            ws.cell(r, 4, e_cur)
            ws.cell(r, 5, e_chg)
            ws.cell(r, 6, e_rate)
            _style_data(ws, r, [3, 4, 5, 6])
            _style_total(ws, r, [3, 4, 5, 6])
            r += 1
        tot_cur += e_cur
        tot_prev += (e_prev if prev_year is not None else 0.0)
    # ---- 按二级科目汇总（全集团，跨主体合计） ----
    any_l2 = False
    for l2 in sorted(all_l2):
        gc = grp_l2_cur.get(l2, 0.0)
        gp = grp_l2_prev.get(l2, 0.0) if prev_year is not None else None
        # 铁律④：全集团二级汇总为 0 不显示
        if prev_year is None:
            if is_zero_amount(gc):
                continue
        else:
            if is_zero_amount(gc, gp):
                continue
        if not any_l2:
            lab = ws.cell(r, 1, '按二级科目汇总（全集团）')
            lab.font = Font(name='Times New Roman', bold=True, size=10, color='7F6000')
            r += 1
            any_l2 = True
        gc_chg = (gc - gp) if prev_year is not None else None
        ws.cell(r, 1, '全集团')
        ws.cell(r, 2, l2)
        ws.cell(r, 3, gp)
        ws.cell(r, 4, gc)
        ws.cell(r, 7, gc_chg)
        # 税金及附加：全集团二级汇总行也填对方科目并做金额勾稽（汇总各主体应交税费计提）
        if name == '税金及附加':
            kb = _norm_tax(l2, keep_loc=False)
            tp_all = sum(taxpay_base.get((en, year, kb), 0.0) for en in entities)
            ws.cell(r, 8, f'应交税费-应交{kb}')
            ws.cell(r, 9, tp_all)                                # 对应科目金额(各主体基础税种计提合计)
            ws.cell(r, 10, _recon_status(gc - tp_all, gc, tp_all))
        # 其他收益：全集团二级汇总行填对方科目并标注全集团勾稽状态（递延收益摊销不按 L2 拆分，金额列留空）
        elif name == '其他收益':
            ws.cell(r, 8, '递延收益(摊销转入)')
            ws.cell(r, 10, oth_gstatus.get(year, '见勾稽核对表'))
        elif name in IMPAIR_PARAMS:
            # 减值损失全集团按二级汇总行也嵌入勾稽（对方科目/对应科目金额/是否勾稽一致）
            seg = _norm_asset_key(l2, imp_strip, imp_tokens)
            rc_all = sum(imp_reserve[(en, year, seg)][1] for en in entities if (en, year, seg) in imp_reserve)
            rd_all = sum(imp_reserve[(en, year, seg)][0] for en in entities if (en, year, seg) in imp_reserve)
            raw_all = set()
            for en in entities:
                if (en, year, seg) in imp_reserve:
                    raw_all |= imp_reserve[(en, year, seg)][2]
            ws.cell(r, 8, '/'.join(sorted(raw_all)) if raw_all else '—')
            ws.cell(r, 9, rc_all)                                # 对应科目金额(各主体减值准备贷计提合计)
            st = _impair_recon_status(gc, rc_all, rd_all, seg is not None)
            ws.cell(r, 10, st)
            # 差异归因说明（2026-08-01 修复 #2）：减值损失(损益) vs 减值准备(余额) 天然不等，
            # 差异主要来自 减值准备借方转回/转销(核销/处置)、期初余额调整、外币折算等；
            # 当 损失≈计提−转销 时归因"转销"，否则提示可能含期初/转回。
            if st == '不一致' and seg:
                diff = gc - rc_all
                if abs(diff - rd_all) <= max(1.0, abs(diff) * 1e-6):
                    note_txt = f'差额≈减值准备借方转销/核销 {rd_all:,.2f}，属资产处置/坏账核销，非账务差错'
                elif abs(gc) < 1e-6 and abs(rc_all) < 1e-6:
                    note_txt = ''
                else:
                    note_txt = (f'差额 {diff:,.2f} 需核实：可能含 期初余额调整/前期转回/外币折算，'
                                f'减值准备借方(转回/转销)={rd_all:,.2f}，建议对照减值准备明细')
                ws.cell(r, 11, note_txt)
            elif st == '一致(转销)':
                ws.cell(r, 11, f'损失≈计提−转销（减值准备借方转销 {rd_all:,.2f}），勾稽一致')
        _style_total(ws, r, [3, 4, 7, 9])
        r += 1
    # ---- 全集团全年合计数（最后一行） ----
    t_chg = (tot_cur - tot_prev) if prev_year is not None else None
    ws.cell(r, 1, '(全集团全年合计数)')
    ws.cell(r, 3, tot_prev if prev_year is not None else None)
    ws.cell(r, 4, tot_cur)
    ws.cell(r, 7, t_chg)
    if name == '税金及附加':
        ws.cell(r, 9, grp_tp)                                   # 对应科目金额(全集团匹配合计)
    elif name == '其他收益':
        ws.cell(r, 9, def_grand)                                # 对应科目金额(全集团递延收益摊销合计)
    ws.cell(r, 10, _check_recon(name, None, year, tot_cur, tb))
    _style_total(ws, r, [3, 4, 7, 9])

    # 资产/信用减值损失：『减值损失 ↔ 减值准备』勾稽已逐行嵌入上方明细表
    # （对方科目 / 对应科目金额 / 是否勾稽一致），不再单独开底部勾稽块（满足"上下两部分合并"需求）。


# ---------------- 写「附注明细表」（未审数 / 调整数 / 审定数 三张平行表，2026-08-01 用户需求） ----------------
def write_footnote_sheet(ws, name, kind, year, entities, gl_agg, tb=None, l2_keep=None, gl_full=None):
    """利润表科目附注明细表：三张表上下排列，行列完全一致。
      表一 未审数：第1列=二级科目，第2列起=各核算主体本年发生额（自然方，按实体名排序），末列集团合计；
      表二 调整数：同结构，初始全 0（待审计调整）；
      表三 审定数：同结构，审定 = 未审数 + 调整数（公式联动）。
    取数与「明细表」同源：gl_agg[(subj, e, l2, y)] = [dr, cr]，nat_amt(kind) 取自然方。
    l2_keep：可选，允许展示的二级名集合（2026-08-02 用户要求「附注汇总只要 2 级明细」——
    GL 归并 key 混有 3 级名（如 住房公积金/ETC通行费），须过滤为仅 2 级，避免 2级(含3级) 与 3级 同列双计）。"""
    ents = sorted(entities)
    # 全集团 L2 全集（本年；保证所有主体行一致）
    all_l2 = set()
    for (sb, en, l2, yr) in gl_agg:
        if sb == name and str(yr) == str(year):
            all_l2.add(l2)
    if l2_keep is not None:
        # 2026-08-03 修复：附注汇总 L2 列表以「TB 二级全集」（明细表同源 rows_spec）为基准，
        # 不再取 all_l2 ∩ l2_keep——GL 归并 key 常缺 2 级父行（工资/五险二金等金额全记在 3/4 级子目），
        # 取交集会把 11 个含三级二级父行丢弃 → 附注汇总合计 17.4M vs 明细表 93.7M（FY本级2025）。
        # 用户方法论：附注汇总明细与明细表二级一致，按同名匹配（个别字不一致也要归并）。
        l2_list = sorted(set(l2_keep))
    else:
        l2_list = sorted(all_l2)
    if not l2_list:
        ws.cell(1, 1, f'（{name} 在 {year} 无综合查询明细表发生额数据，未生成附注明细表）').font = \
            Font(name='Times New Roman', italic=True, color="888888")
        return
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出『集团合计』列
    _single_pl = len(ents) <= 1
    ncols = 1 + len(ents) + (0 if _single_pl else 1)   # 二级科目 + 各主体 (+集团合计, 单体省略)
    HDR_F = Font(name='Times New Roman', bold=True, size=10, color="000000")
    HDR_FL = PatternFill('solid', fgColor='DDEBF7')
    TOT_FL = PatternFill('solid', fgColor='FFF2CC')
    FN10 = Font(name='Times New Roman', size=10)

    def _val(e, l2):
        """取 (e, l2) 自然方数值。2026-08-03 修复：l2_keep 只留 2 级 key，但 GL 归并 key 混有
        3/4 级名（如 管理费用-研发费用-效益奖金）——若直接丢弃，金额主力（3/4 级）全丢，
        附注汇总合计远小于审定表（FY本级 2025 曾 15.11M vs 93.67M）。
        修复：2 级行值 = 精确 GL 发生额 + 其 3/4 级子目归并（求和）——
        父级精确行与子目行是各自独立记录的（父行=直接发生额，子目另行），相加不重叠；
        与 TB 父级含子级的口径一致（铁律30 full_path 归并后父行应含子目）。"""
        k = (name, e, l2, year)
        tot = nat_amt(kind, *gl_agg[k]) if k in gl_agg else 0.0
        pre = l2 + '-'
        for (sb, en, l2k, yr), arr in gl_agg.items():
            if sb == name and en == e and str(yr) == str(year) \
               and str(l2k).startswith(pre):
                tot += nat_amt(kind, *arr)
        return tot

    def _title_row(r, text):
        # 2026-08-02：附注汇总不设合并单元格（便于后期添加/筛选）；块标题仅 A 列加粗
        c = ws.cell(r, 1, text)
        c.font = Font(name='Times New Roman', bold=True, size=10)
        return r + 1

    def _header_row(r):
        ws.cell(r, 1, '二级科目')
        for i, e in enumerate(ents, 2):
            ws.cell(r, i, e)
        if not _single_pl:
            ws.cell(r, ncols, '集团合计')
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c); cell.font = HDR_F; cell.fill = HDR_FL
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = BORDER
        return r + 1

    def _body_rows(r, getter):
        """getter(e, l2) -> 数值。返回 (r_next, grand, col_tot)。
        合计行（2026-08-06 协议：写数值=各主体列累计；公式 data_only 读 None 收回读不回）。"""
        data_first = r
        grand = {l2: 0.0 for l2 in l2_list}
        col_tot = [0.0] * (ncols - 1)     # 各主体列累计（合计行用；下标 0=主体1 列）
        for l2 in l2_list:
            ws.cell(r, 1, l2)
            row_sum = 0.0
            for i, e in enumerate(ents, 2):
                v = getter(e, l2) or 0.0
                row_sum += v
                grand[l2] += v
                col_tot[i - 2] += v
                cell = ws.cell(r, i, round(v, 2))
                cell.border = BORDER; cell.font = FN10; cell.number_format = NUM_FMT
            if not _single_pl:
                gcell = ws.cell(r, ncols, round(row_sum, 2))
                gcell.border = BORDER; gcell.font = FN10; gcell.number_format = NUM_FMT
            cell0 = ws.cell(r, 1); cell0.border = BORDER; cell0.font = FN10
            r += 1
        # 合计行（数值；无数据行（data_first==r）时置 0）
        ws.cell(r, 1, '(合计)')
        for i, e in enumerate(ents, 2):
            cell = ws.cell(r, i, round(col_tot[i - 2], 2) if col_tot[i - 2] else 0.0)
            cell.border = BORDER; cell.font = Font(name='Times New Roman', bold=True, size=10)
            cell.fill = TOT_FL; cell.number_format = NUM_FMT
        if not _single_pl:
            gcell = ws.cell(r, ncols, round(sum(col_tot), 2) if sum(col_tot) else 0.0)
            gcell.border = BORDER; gcell.font = Font(name='Times New Roman', bold=True, size=10)
            gcell.fill = TOT_FL; gcell.number_format = NUM_FMT
        c0 = ws.cell(r, 1); c0.border = BORDER; c0.font = Font(name='Times New Roman', bold=True, size=10); c0.fill = TOT_FL
        return r + 1, grand, col_tot

    r = 1
    # 表一：未审数（本年发生额自然方）—— 数据区起始行 row1_data
    r = _title_row(r, f'{name} 附注披露 — 未审数（{year} 年，各核算主体）')
    r = _header_row(r)
    row1_data = r
    r, grand1, col_tot1 = _body_rows(r, lambda e, l2: _val(e, l2))
    r += 1
    # 表二：调整数（初始 0，待审计调整；审定数公式联动）—— 数据区起始行 row2_data
    r = _title_row(r, f'{name} 附注披露 — 审计调整数（{year} 年）')
    r = _header_row(r)
    row2_data = r
    r, grand2, col_tot2 = _body_rows(r, lambda e, l2: 0.0)
    r += 1
    # 表三：审定数 = 未审数 + 调整数（2026-08-06 协议：写数值=未审数，调整 0 待填；公式 data_only 读 None）
    r = _title_row(r, f'{name} 附注披露 — 审定数（{year} 年，未审数 + 审计调整数）')
    r = _header_row(r)
    row3_data = r
    n_l2 = len(l2_list)
    for i, l2 in enumerate(l2_list):
        rr3 = row3_data + i
        ws.cell(rr3, 1, l2)
        row_sum = 0.0
        for j, e in enumerate(ents, 2):
            v = round(_val(e, l2), 2)
            row_sum += v
            c = ws.cell(rr3, j, v)
            c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
        if not _single_pl:
            g = ws.cell(rr3, ncols, round(row_sum, 2))
            g.border = BORDER; g.font = FN10; g.number_format = NUM_FMT
        c0 = ws.cell(rr3, 1); c0.border = BORDER; c0.font = FN10
    # 表三合计行（= 表一合计 + 表二合计(0) = 表一合计，写数值）
    rr3t = row3_data + n_l2
    ws.cell(rr3t, 1, '(合计)')
    for j, e in enumerate(ents, 2):
        c = ws.cell(rr3t, j, round(col_tot1[j - 2], 2) if col_tot1[j - 2] else 0.0)
        c.border = BORDER; c.font = Font(name='Times New Roman', bold=True, size=10)
        c.fill = TOT_FL; c.number_format = NUM_FMT
    if not _single_pl:
        g = ws.cell(rr3t, ncols, round(sum(col_tot1), 2) if sum(col_tot1) else 0.0)
        g.border = BORDER; g.font = Font(name='Times New Roman', bold=True, size=10)
        g.fill = TOT_FL; g.number_format = NUM_FMT
    c0 = ws.cell(rr3t, 1); c0.border = BORDER; c0.font = Font(name='Times New Roman', bold=True, size=10); c0.fill = TOT_FL
    # 审定数块末尾：集团合计数 / 合并抵消借方 / 合并抵消贷方 / 合并报表数（2026-08-06 协议：写数值）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出
    rr = rr3t + 1   # 2026-08-10 修复：单体时 rr 也须定义（后续凭证明细块 rr += 2 引用）
    if len(ents) > 1:
        ws.cell(rr, 1, '集团合计数')
        for j in range(2, ncols + 1):
            c = ws.cell(rr, j, round(col_tot1[j - 2], 2) if col_tot1[j - 2] else 0.0)
            c.border = BORDER; c.font = Font(name='Times New Roman', bold=True, size=10)
            c.fill = TOT_FL; c.number_format = NUM_FMT
        c0 = ws.cell(rr, 1); c0.border = BORDER; c0.font = Font(name='Times New Roman', bold=True, size=10); c0.fill = TOT_FL
        rr += 1
        ws.cell(rr, 1, '合并抵消借方')
        for j in range(2, ncols + 1):
            c = ws.cell(rr, j, 0.0)
            c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
        c0 = ws.cell(rr, 1); c0.border = BORDER; c0.font = FN10
        rr += 1
        ws.cell(rr, 1, '合并抵消贷方')
        for j in range(2, ncols + 1):
            c = ws.cell(rr, j, 0.0)
            c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
        c0 = ws.cell(rr, 1); c0.border = BORDER; c0.font = FN10
        rr += 1
        ws.cell(rr, 1, '合并报表数')
        for j in range(2, ncols + 1):
            c = ws.cell(rr, j, round(col_tot1[j - 2], 2) if col_tot1[j - 2] else 0.0)
            c.border = BORDER; c.font = Font(name='Times New Roman', bold=True, size=10)
            c.fill = TOT_FL; c.number_format = NUM_FMT
        c0 = ws.cell(rr, 1); c0.border = BORDER; c0.font = Font(name='Times New Roman', bold=True, size=10); c0.fill = TOT_FL
    # 列宽
    ws.column_dimensions['A'].width = 28
    for j in range(2, ncols + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.freeze_panes = 'B3'
    # 2026-08-04：营业外收支附注——账套无二级科目（GL 二级=科目名本身）时，从凭证摘要提取
    # 明细分类（用户要求：账套未明确区分营业外收支明细，可直接提取凭证信息生成附注内容）。
    if name in ('营业外收入', '营业外支出') and gl_full:
        _subs = {}
        for _r in gl_full:
            if str(_r.get('y')) != str(year):
                continue
            _nm = str(_r.get('name') or '')
            if _nm != name and not _nm.startswith(name):
                continue
            _sm = str(_r.get('summary') or '').strip()
            if not _sm or '结转' in _sm:
                continue
            # 摘要清洗：去 到款/支付/收 等冗余尾缀，取主干作明细名
            _key = _sm[:24]
            _amt = abs(float(_r.get('dr') or 0.0)) + abs(float(_r.get('cr') or 0.0))
            _e = str(_r.get('e') or '')
            _subs.setdefault((_e, _key), 0.0)
            _subs[(_e, _key)] += _amt
        if _subs:
            rr += 2
            ws.cell(rr, 1, f'{name} 凭证明细（账套无二级科目，自 GL 摘要提取，供附注填列参考）').font = \
                Font(name='Times New Roman', bold=True, size=10)
            rr += 1
            for (j, h) in enumerate(['核算主体', '摘要（明细参考）', '金额'], 1):
                c = ws.cell(rr, j, h)
                c.font = Font(name='Times New Roman', bold=True, size=10)
                c.fill = HDR_FL
                c.border = BORDER
            rr += 1
            for (_e, _key), _amt in sorted(_subs.items(), key=lambda x: -x[1]):
                ws.cell(rr, 1, _e).border = BORDER
                ws.cell(rr, 2, _key).border = BORDER
                c = ws.cell(rr, 3, round(_amt, 2))
                c.border = BORDER; c.number_format = NUM_FMT
                rr += 1


# ---------------- 写「凭证抽查表」 ----------------
def write_voucher_sheet(ws, name, kind, entities, years, gl_rows, note=None, gl_full=None, detail_total=None, entity_year_filter=None, gl_agg=None, filter_year=None):
    """凭证抽查表：据 GL 发生额逐笔列出本 P&L 科目相关凭证。

    关键规则（2026-07-26 修订）：
      · 负数(红字)金额必须显示：旧版仅显示 >0 一侧导致红字冲减行借贷方全空白，已修正为
        按实际借方/贷方(含负)显示，负数标红，合计含负数。
      · 合计行：仅显示"有数字一侧"的合计——若全部为贷方则只列贷方合计，借方合计不列。
      · 所得税费用：仅剔除【结转类(含本年利润)】与【递延所得税】相关凭证；
        当期所得税费用凭证(借所得税费用 贷应交税费/银行存款)保留抽凭（不复用规则②纯计提剔除）。
      · 其他收益：仅抽"非递延收益摊销转入"的凭证；递延摊销转入部分在明细表/勾稽表直接核对。
      · 底部新增"抽凭发生金额勾稽说明"：抽凭自然方合计 应与 明细表本期未审数合计 − 已剔除金额 相符，
        差额为 0 表示未遗漏；若不符提示检查红字或非结转凭证是否被误剔除。
    """
    allrows = []
    n_skip = 0
    excluded_nat = 0.0   # 抽凭剔除的"自然发生金额"合计（结转类 + 计提/摊销相符 + 递延所得税），用于与明细表勾稽
    if gl_full:
        if name == '所得税费用':
            # 仅剔除结转类(含本年利润)；递延所得税凭证保留抽凭（2026-08-01 修复 #10：
            # 用户要求所得税费用抽凭包含 借所得税费用 贷递延所得税负债/资产 的递延部分）
            groups = OrderedDict()
            for r in gl_full:
                key = (r.get('e'), r.get('y'), r.get('vtype'), r.get('vno'))
                groups.setdefault(key, []).append(r)
            skip = set()
            for key, lines in groups.items():
                blob = ' '.join(' '.join(str(x) for x in (ln.get('name', '') or '', ln.get('opp', '') or '', ln.get('summary', '') or '')) for ln in lines)
                if voucher_is_carryover(blob):
                    skip.add(key)
        elif name in ('营业外支出', '营业外收入', '资产处置损益', '资产减值损失', '信用减值损失'):
            # 全部抽凭：不排除任何凭证（包括结转损益），使用户要求的"全部抽凭"生效。
            # 2026-07-29 用户要求：之前排除结转损益导致金桥信息营业外收入抽样合计与明细表差1笔。
            skip = set()
        elif name == '税金及附加':
            # GL 对方科目字段为空（U8 常用导出通病），无法通过 opp 区分计提/支付。
            # 按用户要求(2026-07-30)：全量抽凭（均抽）。
            skip = set()
        else:
            skip = voucher_skip_keys(
                gl_full,
                key_fields=lambda r: (r.get('e'), r.get('y'), r.get('vtype'), r.get('vno')),
                name_field=lambda r: r.get('name') or '',
                opp_field=lambda r: r.get('opp') or '',
                summ_field=lambda r: r.get('summary') or '',
            )
    else:
        skip = set()
    other_only_nonderred = (name == '其他收益')
    excluded_consistent = 0.0   # 因「勾稽一致」被整组剔除的(主体,年度)凭证自然方金额
    n_consistent = 0
    for (sb, e, y), rows in gl_rows.items():
        if sb != name:
            continue
        if filter_year is not None and y != filter_year:
            continue  # 2026-07-31：按实体年份重建抽凭表（emit 后逐文件调用）
        for row in rows:
            # 所得税费用：仅抽取勾稽不一致的(主体,年度)；一致的整组剔除（不纳入抽凭）
            if entity_year_filter is not None and (e, y) not in entity_year_filter:
                n_consistent += 1
                excluded_consistent += nat_amt(kind, safe(row[5]), safe(row[6]))
                continue
            # row = [date,字,号,对方,摘要,借,贷]; 按 (e, y, 字, 号) 命中整笔凭证排除集合
            if (e, y, row[1], row[2]) in skip:
                n_skip += 1
                excluded_nat += nat_amt(kind, safe(row[5]), safe(row[6]))
                continue
            # 其他收益：此前递延收益摊销转入部分不抽凭，用户要求全部纳入（含递延摊销）
            # 原有递延收益摊销与勾稽表核对逻辑保留不动，增加凭证抽查覆盖
            allrows.append([e] + row)
    allrows.sort(key=lambda x: (str(x[0]) if x[0] is not None else '', str(x[1]) if x[1] is not None else ''))
    if name in ('营业外支出', '营业外收入', '资产处置损益', '资产减值损失', '信用减值损失', '税金及附加'):
        title_suffix = '全量抽凭（含结转凭证）'
    elif name == '所得税费用':
        title_suffix = '已剔除结转损益类凭证（含递延所得税抽凭）'
    else:
        title_suffix = '已剔除结转损益类及计提/摊销相符凭证'
    title = f'{name} 凭证抽查表（据综合查询明细表发生额逐笔生成，{title_suffix}，共排除 {n_skip} 笔）'
    if note:
        title += f'  —— {note}'
    if entity_year_filter is not None:
        title += f'（仅含 {len(entity_year_filter)} 组勾稽不一致的核算主体/年度）'
    ws.cell(1, 1, title).font = Font(name='Times New Roman', bold=True, size=10)
    for c, h in enumerate(VOUCHER_HEADERS, 1):
        cell = ws.cell(2, c, h)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = BORDER
        cell.alignment = CENTER
    ws.freeze_panes = 'A3'
    r = 3
    tot_d = tot_c = 0.0
    has_d = has_c = False
    voucher_nat = 0.0
    for i, rec in enumerate(allrows, 1):
        e, date, vtype, vno, opp, summ, jf, df = rec[:8]
        subj_nm = rec[8] if len(rec) > 8 and str(rec[8]).strip() else opp  # 本方科目名（allrows=[e]+gl_rows[date,字,号,对方,摘要,借,贷,科目名]）
        # 显示实际借/贷（含负数红字），仅 0 侧留空；负数标红
        jf_v = round(jf, 2) if safe(jf) != 0 else None
        df_v = round(df, 2) if safe(df) != 0 else None
        if jf_v is not None:
            has_d = True
        if df_v is not None:
            has_c = True
        vals = [e, i, date, vtype, vno, subj_nm, summ, jf_v, df_v, opp,
                None, None, None, None, None]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (8, 9) and isinstance(v, (int, float)):  # 列8=借方金额, 列9=贷方金额（2026-07-31 修复：原 c>=9 漏掉借方列格式）
                cell.number_format = NUM_FMT
                cell.alignment = RIGHT
                if v < 0:
                    cell.font = Font(color='C00000', name='Times New Roman')
            elif c in (2, 3, 4, 5):
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
        tot_d += safe(jf)
        tot_c += safe(df)
        voucher_nat += nat_amt(kind, safe(jf), safe(df))
        r += 1
    # ---- 合计行（2026-08-06 协议：写数值=tot_d/tot_c 累计；公式 data_only 读 None 收回读不回）
    ws.cell(r, 1, '(合计)').font = TOT_FONT
    _first_data = 3
    _last_data = r - 1
    ws.cell(r, 7, '笔数：%d' % len(allrows)).font = TOT_FONT
    tot_cols = []
    if has_d:
        tot_cols.append((8, round(tot_d, 2)))  # 列8=借方金额
    if has_c:
        tot_cols.append((9, round(tot_c, 2)))  # 列9=贷方金额
    for c, val in tot_cols:
        cell = ws.cell(r, c, val)
        cell.font = TOT_FONT
        cell.fill = TOT_FILL
        cell.border = BORDER
        cell.number_format = NUM_FMT
        cell.alignment = RIGHT
    r += 2
    widths = [14, 8, 12, 8, 16, 24, 24, 30, 14, 14, 12, 12, 12, 18, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------- 默认明细表壳（模板缺失时的兜底，避免静默 0 文件） ----------------
DETAIL_SHELL_WIDTHS = [16, 28, 16, 16, 14, 14, 14, 24, 18, 18]


def _make_default_detail_wb(years, name=''):
    wb = Workbook()
    wb.remove(wb.active)
    for y in sorted(years):
        sn = f'{name}明细表_{y}' if name else f'{y}明细表'
        ws = wb.create_sheet(sn)
        for c, h in enumerate(DETAIL_HEADERS, 1):
            cell = ws.cell(1, c, h)
            cell.font = HDR_FONT
            cell.fill = HDR_FILL
            cell.border = BORDER
            cell.alignment = CENTER
        ws.freeze_panes = 'A2'
        for i, w in enumerate(DETAIL_SHELL_WIDTHS, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    return wb


# ---------------- 跨工作簿克隆表（保留模板格式） ----------------
def clone_sheet(src, dst):
    for col, dim in src.column_dimensions.items():
        dst.column_dimensions[col].width = dim.width
        dst.column_dimensions[col].hidden = dim.hidden
    for rr, dim in src.row_dimensions.items():
        dst.row_dimensions[rr].height = dim.height
    for row in src.iter_rows():
        for c in row:
            if c.value is None and not c.has_style:
                continue
            cell = dst.cell(row=c.row, column=c.column)
            cell.value = c.value
            if c.has_style:
                cell.font = copy.copy(c.font)
                cell.fill = copy.copy(c.fill)
                cell.border = copy.copy(c.border)
                cell.alignment = copy.copy(c.alignment)
                cell.number_format = c.number_format
                cell.protection = copy.copy(c.protection)
    for mc in src.merged_cells.ranges:
        dst.merge_cells(str(mc))
    dst.freeze_panes = src.freeze_panes
    dst.sheet_view.showGridLines = src.sheet_view.showGridLines


# ---------------- 勾稽表：税金及附加 ↔ 应交税费计提 ----------------
TAX_RECON_HEADERS = ['核算主体', '税种(规范)', '税金及附加(借方发生额)', '应交税费-相关子目(贷方计提)',
                     '差额', '是否勾稽一致', '备注']
# 规范税种口径（合并 城市维护建设税→城建税、地方教育附加→地方教育费附加、车船使用税→车船税 等杂乱命名）
TAX_CANON = {
    '城建税': '城建税', '城市维护建设税': '城建税',
    '教育费附加': '教育费附加',
    '地方教育费附加': '地方教育费附加', '地方教育附加': '地方教育费附加',
    '印花税': '印花税',
    '房产税': '房产税', '房屋税': '房产税',
    '土地使用税': '土地使用税', '土地使用费': '土地使用税',
    '车船税': '车船税', '车船使用税': '车船税',
    '环保税': '环保税', '环境保护税': '环保税',
    '关税': '关税',
}
# 应交税费 中不属于「税金及附加」计提口径的子目（增值税/所得税/个税/代扣/待抵扣/待转销销项等）
TAXPAY_EXCLUDE = ('增值税', '所得税', '个人所得税', '代扣代缴', '待抵扣', '待认证',
                  '未交增值税', '待转销', '销项税额', '进项税额')

_LOC_RE = re.compile(
    r'[-–](本地|预缴|本公司|异地预缴|分公司|项目部|总部|区|县|市|省|一|二|三|四|五|六|七|八|九|十|'
    r'十一|十二|十三|十四|十五|十六|十七|十八|十九|二十)$')


def _norm_tax(name, keep_loc=False):
    """税种名→规范键。keep_loc=True 保留『本地/预缴/本公司/异地预缴』等属地后缀(用于精确匹配)；
    keep_loc=False 剥属地后缀(用于按基础税种回退匹配)。"""
    s = (str(name).replace('税金及附加-', '').replace('应交税费-应交', '')
         .replace('应交税费-', '').replace('应交税费', '').strip())
    for k, v in TAX_CANON.items():
        s = s.replace(k, v)
    if not keep_loc:
        s = _LOC_RE.sub('', s).strip().rstrip('-－').strip()
    return s


def _norm_tax_l2(l2):
    return _norm_tax(l2, keep_loc=True)


def _norm_taxpay_sub(sub):
    return _norm_tax(sub, keep_loc=True)


def _recon_status(diff, base_a, base_b):
    return '一致' if abs(diff) <= max(1.0, max(abs(base_a), abs(base_b)) * 1e-6) else '不一致'


def _nz(x):
    """消除浮点残差导致的 -0.00 显示。"""
    return 0.0 if abs(x) < 1e-6 else x


def write_tax_recon_sheet(ws, recon_gl, entities, years, gl_agg):
    """税金及附加(借/费用) ↔ 应交税费相关子目(贷/计提) 逐主体·分年·分税种 勾稽。"""
    tax_exp = defaultdict(lambda: [0.0, set()])   # (e,y,k) -> [amt, raw_names]
    taxpay = defaultdict(lambda: [0.0, set()])    # (e,y,k) -> [amt, raw_names]
    for (sb, e, l2, y), v in gl_agg.items():
        if sb != '税金及附加':
            continue
        k = _norm_tax_l2(l2)
        tax_exp[(e, y, k)][0] += v[0]
        tax_exp[(e, y, k)][1].add(l2)
    for (e, name, y), v in recon_gl.items():
        s0 = str(name).split('-')[0].strip()
        if s0 != '应交税费':
            continue
        if any(x in name for x in TAXPAY_EXCLUDE):
            continue
        k = _norm_taxpay_sub(name)
        taxpay[(e, y, k)][0] += v[1]
        taxpay[(e, y, k)][1].add(name)
    # 标题 + 说明
    ws.cell(1, 1, '税金及附加 ↔ 应交税费(相关子目)计提 勾稽核对表').font = Font(name='Times New Roman', bold=True, size=10, color='1F4E78')
    ws.cell(2, 1, '左：税金及附加借方发生额(费用认定)；右：应交税费相关子目贷方计提(已剔除增值税/所得税/个税/代扣代缴)。'
                  '二者应相等。2026 为 YTD(1-5月)，但两侧同源 GL 可比。').font = Font(name='Times New Roman', italic=True, size=10, color='808080')
    hr = 3
    for c, h in enumerate(TAX_RECON_HEADERS, 1):
        cell = ws.cell(hr, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER; cell.alignment = CENTER
    ws.freeze_panes = f'A{hr + 1}'
    r = hr + 1
    tot_te = tot_tp = 0.0
    num_cols = (3, 4, 5)
    ey_rows = {}   # (e,y) -> 数据行号列表（小计公式引用，2026-08-02）
    sub_rows = []  # 各主体小计行号（全集团合计公式引用）

    def _emit(e, y, k, te, tp, note_override=None):
        nonlocal r, tot_te, tot_tp
        diff = te - tp
        status = _recon_status(diff, te, tp)
        note = note_override or ''
        if not note:
            if te == 0 and tp != 0:
                note = '税金及附加侧无发生额(可能直接缴纳未过应交税费)'
            elif tp == 0 and te != 0:
                note = '应交税费侧无对应子目(可能未通过应交税费核算)'
            raw = '/'.join(sorted(tax_exp[(e, y, k)][1] | taxpay[(e, y, k)][1]))
            if raw:
                note = (note + '；' if note else '') + '原名称:' + raw
        ws.cell(r, 1, e)
        ws.cell(r, 2, k)
        ws.cell(r, 3, te)
        ws.cell(r, 4, tp)
        ws.cell(r, 5, _nz(diff))
        ws.cell(r, 6, status)
        ws.cell(r, 7, note)
        for cc in range(1, 8):
            cell = ws.cell(r, cc)
            cell.border = BORDER
            if cc in num_cols:
                cell.number_format = NUM_FMT; cell.alignment = RIGHT
            elif cc <= 2:
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
        tot_te += te; tot_tp += tp
        ey_rows.setdefault((e, y), []).append(r)
        r += 1

    for e in sorted(entities):
        for y in sorted(years):
            keys = set()
            for (ee, yy, k) in tax_exp:
                if ee == e and yy == y:
                    keys.add(k)
            for (ee, yy, k) in taxpay:
                if ee == e and yy == y:
                    keys.add(k)
            if not keys:
                continue
            se_te = se_tp = 0.0
            for k in sorted(keys):
                te = tax_exp[(e, y, k)][0] if (e, y, k) in tax_exp else 0.0
                tp = taxpay[(e, y, k)][0] if (e, y, k) in taxpay else 0.0
                _emit(e, y, k, te, tp)
                se_te += te; se_tp += tp
            # 主体·年 小计（2026-08-06 协议：写数值=se_te/se_tp 累计；公式 data_only 读 None）
            ws.cell(r, 1, f'{e} {y} 小计').font = TOT_FONT
            ws.cell(r, 2, '(小计)')
            ws.cell(r, 3, round(se_te, 2)).number_format = NUM_FMT
            ws.cell(r, 4, round(se_tp, 2)).number_format = NUM_FMT
            ws.cell(r, 5, round(se_te - se_tp, 2)).number_format = NUM_FMT
            ws.cell(r, 6, _recon_status(se_te - se_tp, se_te, se_tp))
            for cc in range(1, 8):
                cell = ws.cell(r, cc); cell.border = BORDER; cell.fill = TOT_FILL
                if cc in num_cols:
                    cell.alignment = RIGHT
                elif cc <= 2:
                    cell.alignment = LEFT
                else:
                    cell.alignment = CENTER
            sub_rows.append(r)
            r += 1
    # 全集团合计（2026-08-06 协议：写数值；⚡ 2026-08-10 单体裁剪）
    if len(entities) > 1:
        ws.cell(r, 1, '(全集团合计)').font = TOT_FONT
        ws.cell(r, 2, '')
        ws.cell(r, 3, round(tot_te, 2)).number_format = NUM_FMT
        ws.cell(r, 4, round(tot_tp, 2)).number_format = NUM_FMT
        ws.cell(r, 5, round(tot_te - tot_tp, 2)).number_format = NUM_FMT
        ws.cell(r, 6, _recon_status(tot_te - tot_tp, tot_te, tot_tp))
        for cc in range(1, 8):
            cell = ws.cell(r, cc); cell.border = BORDER; cell.fill = TOT_FILL; cell.font = TOT_FONT
            if cc in num_cols:
                cell.alignment = RIGHT
            elif cc <= 2:
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
    widths = [16, 16, 20, 22, 16, 16, 46]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------- 勾稽表：其他收益 ↔ 递延收益摊销额 ----------------
def write_other_income_recon_sheet(ws, recon_gl, entities, years, gl_agg):
    """其他收益(贷方发生额) ↔ 递延收益摊销额(借方发生额) 逐主体·分年 勾稽。

    其他收益主要为政府补助等递延摊销转入（借 递延收益 / 贷 其他收益），属账项结转，
    不抽凭；仅就该发生额与递延收益当期借方摊销额做勾稽核对，作为唯一核对依据。
    两侧同源 GL（2026 为 YTD 但可比），不受 GL-vs-TB 的 YTD 不可比限制。"""
    OTHER_INCOME_HEADERS = ['核算主体', '年份', '其他收益(贷方发生额)', '递延收益摊销额(借方发生额)',
                            '差额', '是否勾稽一致', '备注']
    # 其他收益(贷发) 按 (e, y)
    oth = defaultdict(float)
    for (sb, e, l2, y), v in gl_agg.items():
        if sb == '其他收益':
            oth[(e, y)] += v[1]
    # 递延收益(借发=摊销转出) 按 (e, y)
    defr = defaultdict(float)
    defr_names = defaultdict(set)
    if recon_gl is not None:
        for (e, name, y), v in recon_gl.items():
            s0 = str(name).split('-')[0].strip()
            if s0 == '递延收益' or str(name).startswith('递延收益'):
                defr[(e, y)] += v[0]
                defr_names[(e, y)].add(name)
    ws.cell(1, 1, '其他收益 ↔ 递延收益摊销额 勾稽核对表').font = Font(name='Times New Roman', bold=True, size=10, color='1F4E78')
    ws.cell(2, 1, '左：其他收益贷方发生额（政府补助等递延摊销转入）；右：递延收益借方发生额（摊销转出）。'
                  '二者应相等。摊销转入凭证不再抽凭，本表为唯一核对依据。2026 为 YTD(1-5月)，两侧同源 GL 可比。'
                  ).font = Font(name='Times New Roman', italic=True, size=10, color='808080')
    hr = 3
    for c, h in enumerate(OTHER_INCOME_HEADERS, 1):
        cell = ws.cell(hr, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER; cell.alignment = CENTER
    ws.freeze_panes = f'A{hr + 1}'
    r = hr + 1
    tot_oth = tot_def = 0.0
    num_cols = (3, 4, 5)
    pairs = set(oth.keys()) | set(defr.keys())
    for e in sorted(entities):
        for y in sorted(years):
            if (e, y) not in pairs:
                continue
            o = oth.get((e, y), 0.0)
            d = defr.get((e, y), 0.0)
            diff = o - d
            status = _recon_status(diff, o, d)
            if o == 0 and d != 0:
                note = '其他收益侧无发生额'
            elif d == 0 and o != 0:
                note = '递延收益侧无对应摊销(可能系直接计入其他收益的非递延政府补助)'
            else:
                note = ''
            raw = '/'.join(sorted(defr_names.get((e, y), set())))
            if raw:
                note = (note + '；' if note else '') + '递延收益子目:' + raw
            ws.cell(r, 1, e)
            ws.cell(r, 2, y)
            ws.cell(r, 3, o)
            ws.cell(r, 4, d)
            ws.cell(r, 5, _nz(diff))
            ws.cell(r, 6, status)
            ws.cell(r, 7, note)
            for cc in range(1, 8):
                cell = ws.cell(r, cc)
                cell.border = BORDER
                if cc in num_cols:
                    cell.number_format = NUM_FMT; cell.alignment = RIGHT
                elif cc <= 2:
                    cell.alignment = LEFT
                else:
                    cell.alignment = CENTER
            tot_oth += o; tot_def += d
            r += 1
    # 全集团合计（⚡ 2026-08-10 单体裁剪）
    if len(entities) > 1:
        ws.cell(r, 1, '(全集团合计)').font = TOT_FONT
        ws.cell(r, 2, '')
        ws.cell(r, 3, tot_oth).number_format = NUM_FMT
        ws.cell(r, 4, tot_def).number_format = NUM_FMT
        ws.cell(r, 5, _nz(tot_oth - tot_def)).number_format = NUM_FMT
        ws.cell(r, 6, _recon_status(tot_oth - tot_def, tot_oth, tot_def)).font = TOT_FONT
        ws.cell(r, 7, '')
        for cc in range(1, 8):
            cell = ws.cell(r, cc); cell.border = BORDER; cell.fill = TOT_FILL; cell.font = TOT_FONT
            if cc in num_cols:
                cell.alignment = RIGHT
            elif cc <= 2:
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
    widths = [16, 10, 22, 24, 16, 16, 50]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------- 勾稽表：减值损失(信用/资产) ↔ 对应减值准备 ----------------
# 资产减值损失(6701) 对应「非金融资产」减值准备；信用减值损失(6702) 对应「金融资产」减值准备(坏账准备族)。
# 二者均为 GL 内部对账（同一年度 GL 口径），2026 虽为 YTD(1-5月) 但两侧同源可比。
# 名称归一：先剥离已知前缀(信用减值损失/坏账准备)再按关键字匹配，兼容 产成品/原材料→存货跌价准备、
# 坏账准备-应收账款→应收账款坏账准备、应收账款坏账准备→应收账款坏账准备 等杂乱命名。
ASSET_IMPAIR_TOKENS = [
    ('存货', '存货跌价准备'), ('产成品', '存货跌价准备'), ('原材料', '存货跌价准备'),
    ('周转材料', '存货跌价准备'), ('材料成本差异', '存货跌价准备'),
    ('固定资产', '固定资产减值准备'),
    ('无形资产', '无形资产减值准备'),
    ('在建工程', '在建工程减值准备'),
    ('投资性房地产', '投资性房地产减值准备'),
    ('长期股权投资', '长期股权投资减值准备'),
    ('商誉', '商誉减值准备'),
    ('持有至到期投资', '持有至到期投资减值准备'),
    ('使用权资产', '使用权资产减值准备'),
    ('合同资产', '合同资产减值准备'),
    ('生产性生物资产', '生产性生物资产减值准备'),
    ('抵债资产', '抵债资产减值准备'),
    ('损余物资', '损余物资跌价准备'),
    ('工程物资', '工程物资减值准备'),
    ('贷款', '贷款损失准备'),
]
CREDIT_ASSET_TOKENS = [
    ('应收账款', '应收账款坏账准备'), ('应收帐款', '应收账款坏账准备'),
    ('其他应收款', '其他应收款坏账准备'), ('其他应收账款', '其他应收款坏账准备'),
    ('商业承兑汇票', '商业承兑汇票坏账准备'),
    ('应收票据', '应收票据坏账准备'),
    ('预付账款', '预付账款坏账准备'), ('预付帐款', '预付账款坏账准备'),
    ('长期应收款', '长期应收款坏账准备'),
    ('应收款项融资', '应收款项融资坏账准备'),
    ('债权投资', '债权投资减值准备'),
    ('其他债权投资', '其他债权投资减值准备'),
    ('合同资产', '合同资产减值准备'),
    ('租赁应收款', '租赁应收款坏账准备'),
]
ASSET_IMPAIR_LABELS = {lab for _, lab in ASSET_IMPAIR_TOKENS}
CREDIT_ASSET_LABELS = {lab for _, lab in CREDIT_ASSET_TOKENS}


def _norm_asset_key(name, strip_prefixes, tokens):
    """从科目全称抽取『规范减值准备标签』。先剥离已知前缀(如 信用减值损失/坏账准备)，再按关键字匹配。"""
    s = str(name).strip()
    for p in strip_prefixes:
        while s.startswith(p):
            s = s[len(p):].lstrip('－-_').strip()
    for kw, lab in tokens:
        if kw in s:
            return lab
    return None

IMPAIR_RECON_HEADERS = ['核算主体', '资产类别(规范)', '减值损失(借方)',
                        '对应减值准备(贷方计提)', '对应减值准备(借方转回/转销)', '差额', '是否勾稽一致', '备注']


def _emit_impair_block(ws, r0, *, subject, expense_strip, tokens, reserve_family, title, note_prefix, entities, years, gl_agg, recon_gl, year_filter=None):
    """把『减值损失 ↔ 对应减值准备』勾稽的【表头 + 数据行(含分主体分年小计)】写入 ws 自 r0 行起，
    不含独立大标题与『全集团合计』行（由调用方决定）。供『独立勾稽表』与『嵌入明细表』两处复用。
    返回 (next_row, tot_ie, tot_rc, tot_rd)。"""
    imp_exp = defaultdict(lambda: [0.0, set()])      # (e,y,k) -> [amt, raw]; k=None 表示"其他"
    reserve = defaultdict(lambda: [0.0, 0.0, set()])  # (e,y,seg) -> [dr, cr, raw]
    for (sb, e, l2, y), v in gl_agg.items():
        if sb != subject:
            continue
        k = _norm_asset_key(l2, expense_strip, tokens)
        key = k if k else '其他(需人工核对)'
        imp_exp[(e, y, key)][0] += v[0]
        imp_exp[(e, y, key)][1].add(l2)
    for (e, name, y), v in recon_gl.items():
        if reserve_family == 'credit':
            # 仅对真正的坏账准备类科目对账；排除名称含"应收账款"但实为应交税费/重分类的科目
            if '坏账' not in name and '信用减值' not in name:
                continue
            seg = _norm_asset_key(name, ('坏账准备',), CREDIT_ASSET_TOKENS)
            if seg is None:
                continue
        else:
            if _norm_asset_key(name, ('坏账准备',), CREDIT_ASSET_TOKENS) is not None:
                continue  # 坏账准备族属信用减值损失，资产减值损失侧剔除
            if not any(t in name for t in ('准备', '减值', '跌价')):
                continue  # 仅对真正的资产减值准备/跌价准备对账
            seg = _norm_asset_key(name, (), ASSET_IMPAIR_TOKENS)
            if seg is None:
                continue
        reserve[(e, y, seg)][0] += v[0]
        reserve[(e, y, seg)][1] += v[1]
        reserve[(e, y, seg)][2].add(name)
    for c, h in enumerate(IMPAIR_RECON_HEADERS, 1):
        cell = ws.cell(r0, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.border = BORDER; cell.alignment = CENTER
    r = r0 + 1
    tot_ie = tot_rc = tot_rd = 0.0
    num_cols = (3, 4, 5, 6)
    valid_labels = CREDIT_ASSET_LABELS if reserve_family == 'credit' else ASSET_IMPAIR_LABELS

    def _emit_row(e, y, label, k, ie, note_override=None):
        nonlocal r, tot_ie, tot_rc, tot_rd
        rd, rc, raw = (reserve[(e, y, k)] if k is not None and (e, y, k) in reserve
                       else [0.0, 0.0, set()])
        diff = ie - rc
        status = _recon_status(diff, ie, rc)
        note = note_override or ''
        if not note:
            if k is None or k == '其他(需人工核对)':
                note = f'{subject}-其他，未对应具体减值准备科目，需人工核对'
                status = '需人工核对'
            elif ie == 0 and rc != 0:
                if rc > 0:
                    note = f'{subject}无发生额，减值准备有计提(可能期初/合并/处置)'
                else:
                    note = f'{subject}无发生额，减值准备为净转回/转销(贷方为负，可能资产处置或前期转回)'
            elif rc == 0 and ie != 0:
                note = f'减值准备无计提，{subject}有发生额(可能未通过减值准备科目)'
            elif abs(diff) <= max(1.0, max(abs(ie), abs(rc)) * 1e-6):
                if abs(rd) > 1e-6:
                    note = f'差额≈0；减值准备借方(转回/转销)={rd:,.2f}'
            else:
                if abs(diff - rd) <= max(1.0, abs(diff) * 1e-6):
                    note = f'差额=减值准备借方转销 {rd:,.2f}，属资产处置/坏账核销转销，正常'
                    status = '一致(转销)'
                else:
                    note = '差额需核实(可能含前期转回/处置)'
        if not note_override:
            raws = '/'.join(sorted(raw)) if raw else ''
            if raws:
                note = (note + '；' if note else '') + '减值准备原名称:' + raws
        ws.cell(r, 1, e)
        ws.cell(r, 2, label)
        ws.cell(r, 3, ie)
        ws.cell(r, 4, rc)
        ws.cell(r, 5, rd)
        ws.cell(r, 6, _nz(diff))
        ws.cell(r, 7, status)
        ws.cell(r, 8, note)
        for cc in range(1, 9):
            cell = ws.cell(r, cc)
            cell.border = BORDER
            if cc in num_cols:
                cell.number_format = NUM_FMT; cell.alignment = RIGHT
            elif cc <= 2:
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
        tot_ie += ie; tot_rc += rc; tot_rd += rd
        r += 1

    for e in sorted(entities):
        for y in sorted(years):
            if year_filter and y != year_filter:
                continue
            keys = set()
            for (ee, yy, k) in imp_exp:
                if ee == e and yy == y:
                    keys.add(k)
            for (ee, yy, seg) in reserve:
                if ee == e and yy == y and seg in valid_labels:
                    keys.add(seg)
            if not keys:
                continue
            se_ie = se_rc = se_rd = 0.0
            for k in sorted(keys, key=lambda x: (x is None, x)):
                label = '其他(人工核对)' if k is None else k
                ie = imp_exp[(e, y, k)][0] if (e, y, k) in imp_exp else 0.0
                _emit_row(e, y, label, k, ie)
                se_ie += ie
                se_rc += (reserve[(e, y, k)][1] if k is not None and (e, y, k) in reserve else 0.0)
                se_rd += (reserve[(e, y, k)][0] if k is not None and (e, y, k) in reserve else 0.0)
            ws.cell(r, 1, f'{e} {y} 小计').font = TOT_FONT
            ws.cell(r, 2, '(小计)')
            ws.cell(r, 3, se_ie).number_format = NUM_FMT
            ws.cell(r, 4, se_rc).number_format = NUM_FMT
            ws.cell(r, 5, se_rd).number_format = NUM_FMT
            ws.cell(r, 6, _nz(se_ie - se_rc)).number_format = NUM_FMT
            ws.cell(r, 7, _recon_status(se_ie - se_rc, se_ie, se_rc))
            for cc in range(1, 9):
                cell = ws.cell(r, cc); cell.border = BORDER; cell.fill = TOT_FILL
                if cc in num_cols:
                    cell.alignment = RIGHT
                elif cc <= 2:
                    cell.alignment = LEFT
                else:
                    cell.alignment = CENTER
            r += 1
    return r, tot_ie, tot_rc, tot_rd


def write_impair_recon_sheet(ws, recon_gl, entities, years, gl_agg, *,
                             subject='资产减值损失',
                             expense_strip=('资产减值损失',),
                             tokens=ASSET_IMPAIR_TOKENS,
                             reserve_family='asset',
                             title='资产减值损失 ↔ 相关资产减值准备 勾稽核对表',
                             note_prefix='资产减值损失'):
    """减值损失(借) ↔ 对应减值准备(贷计提；借=转回/转销) 逐主体·分年·分类别 勾稽（独立勾稽表封装）。

    subject='资产减值损失'(6701) 时 reserve_family='asset'：仅对『非金融资产减值准备』对账，
    剔除 坏账准备(属信用减值损失)/递延所得税(属所得税费用)；
    subject='信用减值损失'(6702) 时 reserve_family='credit'：仅对『金融资产减值准备(坏账准备族)』对账。
    名称归一：先剥离前缀再按关键字匹配，兼容 产成品/原材料→存货跌价准备、
    坏账准备-应收账款→应收账款坏账准备 等杂乱命名。"""
    ws.cell(1, 1, title).font = Font(name='Times New Roman', bold=True, size=10, color='1F4E78')
    ws.cell(2, 1, f'左：{note_prefix}(借,减值费用认定)；右：对应减值准备(贷=计提；借=转回/转销)。'
                  '二者应相等；差额恰等于"借方转回/转销"时属资产处置/坏账核销转销，正常。'
                  '已剔除另一类减值损失对应的准备科目及递延所得税。'
                  '2026 为 YTD(1-5月)，两侧同源 GL 可比。').font = Font(name='Times New Roman', italic=True, size=10, color='808080')
    r, tot_ie, tot_rc, tot_rd = _emit_impair_block(
        ws, 3, subject=subject, expense_strip=expense_strip, tokens=tokens,
        reserve_family=reserve_family, title=title, note_prefix=note_prefix,
        entities=entities, years=years, gl_agg=gl_agg, recon_gl=recon_gl)
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出
    if len(entities) > 1:
        ws.cell(r, 1, '(全集团合计)').font = TOT_FONT
        ws.cell(r, 2, '')
        ws.cell(r, 3, tot_ie).number_format = NUM_FMT
        ws.cell(r, 4, tot_rc).number_format = NUM_FMT
        ws.cell(r, 5, tot_rd).number_format = NUM_FMT
        ws.cell(r, 6, _nz(tot_ie - tot_rc)).number_format = NUM_FMT
        ws.cell(r, 7, '汇总(含各户转销/转回，以分主体分年核对为准)')
        for cc in range(1, 9):
            cell = ws.cell(r, cc); cell.border = BORDER; cell.fill = TOT_FILL; cell.font = TOT_FONT
            if cc in (3, 4, 5, 6):
                cell.alignment = RIGHT
            elif cc <= 2:
                cell.alignment = LEFT
            else:
                cell.alignment = CENTER
    widths = [16, 18, 20, 22, 22, 16, 16, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# 减值损失勾稽参数（资产减值损失 / 信用减值损失 复用 _emit_impair_block）
IMPAIR_PARAMS = {
    '资产减值损失': dict(subject='资产减值损失', expense_strip=('资产减值损失',),
                      tokens=ASSET_IMPAIR_TOKENS, reserve_family='asset',
                      title='资产减值损失 ↔ 相关资产减值准备 勾稽核对表', note_prefix='资产减值损失'),
    '信用减值损失': dict(subject='信用减值损失', expense_strip=('信用减值损失',),
                      tokens=CREDIT_ASSET_TOKENS, reserve_family='credit',
                      title='信用减值损失 ↔ 金融资产减值准备(坏账准备族) 勾稽核对表', note_prefix='信用减值损失'),
}


def _subject_detail_total(gl_agg, name, kind):
    """本 P&L 科目『本期未审数』2 年合计（自然方：rev 取贷、exp 取借），用于凭证抽查表勾稽。"""
    t = 0.0
    for (sb, e, l2, y), v in gl_agg.items():
        if sb == name:
            t += nat_amt(kind, v[0], v[1])
    return t


# ---------------- 组装单个科目工作簿 ----------------
def build_subject_workbook(data_dir, code, name, kind, entities, years, gl_agg, gl_rows, tb, recon_gl=None, gl_full=None):
    """直接按年构建（2026-08-01 改造：消除合并稿中间态，替代 emit_per_year 分拆+按年重建）。
    每年独立 wb：审定表(add_audit_summary_sheets target_year) + 明细表(write_detail_sheet 当年+上年对比)
    + 凭证抽查表(filter_year=当年) + 勾稽核对表(当年)。"""
    saved = []
    tpl = resolve_template('6111投资收益', 'PL_TEMPLATE_DIR')
    if os.path.exists(tpl):
        wb_t = load_workbook(tpl)
    else:
        print(f'  ⚠️ 模板缺失：{tpl}，改用内置默认明细表壳（建议还原模板以保留原格式）。')
        wb_t = _make_default_detail_wb(years, name)
    yrs = sorted(years)
    from audit_common import add_audit_summary_sheets as _add_pl_audit
    from audit_common import discover_entities as _de_pl, read_tb_full as _rtb_pl
    try:
        # SAP：单主体驱动时按 current_comp 过滤（防审定表混 88 家主体行）
        if _adapter is not None and _adapter.is_sap(data_dir):
            _c_pl = _adapter.current_comp()
            _e_pl = ({_c_pl: _de_pl(data_dir).get(_c_pl, {})} if _c_pl else _de_pl(data_dir))
        else:
            _e_pl = _de_pl(data_dir)
        _t_pl = _rtb_pl(data_dir, _e_pl)
    except Exception as _ex:
        _e_pl = _t_pl = None
        print(f'  ⚠️ 审定表数据准备失败：{_ex}')
    detail_total = _subject_detail_total(gl_agg, name, kind)
    voucher_ey = None
    recon_rows_all = None
    if name == '所得税费用':
        deferred_tb = read_deferred_tb(entities)
        recon_rows_all, _need_all, voucher_ey = compute_income_tax_recon(
            entities, years, gl_rows, recon_gl or {}, deferred_tb, gl_agg, name=name, kind=kind,
            gl_full=gl_full)
        # 从全量 GL(gl_full) 补充该 (主体,年) 的应交所得税借方/贷方行到抽凭数据源（仅一次，防逐年重复）
        if gl_full and voucher_ey:
            for r in gl_full:
                if (r['e'], r['y']) not in voucher_ey:
                    continue
                _nm = str(r['name'] or '')
                if _nm.startswith('应交税费') and '所得税' in _nm and '递延' not in _nm and '个人' not in _nm:
                    gl_rows.setdefault(('所得税费用', r['e'], r['y']), []).append(
                        [r['date'], r['vtype'], r['vno'], r['opp'], r['summary'], r['dr'], r['cr']])
    from audit_shell import finalize_workbook as _fw_pl
    for y in yrs:
        ys = str(y)
        out = Workbook()
        out.remove(out.active)
        # 1) 审定表（按年，置于首位）
        if name not in VOUCHER_EXCLUDE:
            try:
                # 2026-08-07 v2（会计语言优先，代码方言隔离）：审定表 codes 经账套配置
                # 解析——account_profiles.json subject_codes 有账套特例码用之（JTt 其他
                # 收益=6115、JTt 税金及附加=6071 等非标码），否则用 PL_SUBJECTS 标准码。
                _pl_reg_key = {'其他收益': 'other_income', '税金及附加': 'tax_surcharge',
                               '资产处置损益': 'asset_disposal',
                               '信用减值损失': 'credit_impair',
                               '资产减值损失': 'asset_impair'}.get(name)
                # ⚡ 2026-08-10 账套名取真实数据根（SAP 下 data_dir=_work_{comp} 独立工作目录，
                # basename 非账套名 → account_profiles 解析落空 → SAP 特例码（6730 等）失效
                # ⚡⚡ 2026-08-27 P0 修复（其他收益审定表 0 根因）：_DATA_ROOT=AH/数据/2026
                #   → basename="2026"，account_profiles 配置 key="AH" 匹配不到 → other_income
                #   特例码 6730 失效 → 用默认 6113 查 TB → 审定数 0（明细表按名称有 2086万）。
                #   改从数据根路径【向上】找第一个有配置的目录段（AH）。
                _dr = (_adapter._DATA_ROOT if _adapter is not None and getattr(_adapter, '_DATA_ROOT', None) else data_dir)
                _acct_pl = os.path.basename(str(_dr).rstrip('\\/'))
                try:
                    import subject_mapping as _sm_pl0
                    if not _sm_pl0.account_profile(_acct_pl):
                        for _seg in str(_dr).replace('\\', '/').split('/'):
                            if _seg and _sm_pl0.account_profile(_seg):
                                _acct_pl = _seg
                                break
                except Exception:
                    pass
                _pl_codes = [str(code)]
                if _pl_reg_key:
                    try:
                        import subject_mapping as _sm_pl
                        _pl_codes = _sm_pl.resolve_subject_codes(_acct_pl, _pl_reg_key, [str(code)])
                    except Exception:
                        pass
                _add_pl_audit(out, data_dir,
                              [dict(title=f'{name} 审定表', codes=_pl_codes,
                                    is_credit=(kind == 'rev'), subj_name=name,
                                    income_statement=True, by_entity=True)],
                              tb_full=_t_pl, entities=_e_pl, target_year=ys)
            except Exception as _ex:
                print(f'  ⚠️ 审定表注入失败 [{name} {y}]：{_ex}')
        # 2) 明细表（当年 + 上年对比）
        earlier = [yy for yy in yrs if yy < y]
        prev = max(earlier) if earlier else None
        if f'{y}明细表' in wb_t.sheetnames:
            src = wb_t[f'{y}明细表']
        else:
            src = wb_t[wb_t.sheetnames[0]]
        ws = out.create_sheet(f'{name}明细表_{y}')
        clone_sheet(src, ws)
        write_detail_sheet(ws, name, kind, y, prev, entities, yrs, gl_agg, tb, recon_gl)
        # 2b) 附注明细表（未审数/调整数/审定数 三张平行表，2026-08-01 用户需求）
        fws = out.create_sheet('附注汇总')
        write_footnote_sheet(fws, name, kind, y, entities, gl_agg, tb, gl_full=gl_full)
        # 3) 凭证抽查表（当年；所得税费用为条件生成）
        if name == '所得税费用':
            recon_rows_y = [r for r in (recon_rows_all or []) if str(r.get('y')) == ys]
            need_voucher_y = any(ey[1] == y for ey in (voucher_ey or []))
            rws = out.create_sheet('所得税费用勾稽核对')
            write_income_tax_recon_sheet(rws, recon_rows_y, need_voucher_y, entities, [y])
            if need_voucher_y:
                vws = out.create_sheet('凭证抽查表')
                write_voucher_sheet(vws, name, kind, entities, [y], gl_rows,
                                    note='仅含勾稽不一致的核算主体/年度凭证（已剔除结转类及递延所得税），供对差异逐笔抽查',
                                    gl_full=gl_full, detail_total=detail_total,
                                    entity_year_filter=voucher_ey, filter_year=ys)
        elif name not in VOUCHER_EXCLUDE:
            # 减值损失类：勾稽内容全部一致(含"一致(转销)") 则不生成凭证抽查表（2026-08-01 修复 #7）
            if name in IMPAIR_PARAMS and recon_gl is not None:
                _p = IMPAIR_PARAMS[name]
                _st_strip, _st_tokens, _st_family = _p.get('strip', ()), _p.get('tokens', ()), _p.get('reserve_family')
                _consistent = True
                # 聚合该年 减值损失 vs 减值准备 勾稽，全一致才跳过抽凭
                for _ey_key, (_jf, _df) in recon_gl.items():
                    if str(_ey_key[2]) != ys:
                        continue
                    _nm = _ey_key[1]
                    if _st_family == 'credit':
                        if '坏账' not in _nm and '信用减值' not in _nm:
                            continue
                    else:
                        if '坏账' in _nm or '递延所得税' in _nm:
                            continue
                        if not any(t in _nm for t in ('准备', '减值', '跌价')):
                            continue
                    if abs(_jf) > 1e-6 or abs(_df) > 1e-6:
                        _consistent = False
                        break
                if _consistent:
                    print(f'   [SKIP] {name} {ys} 凭证抽查（减值损失与减值准备勾稽一致，无需抽凭）')
                else:
                    vws = out.create_sheet('凭证抽查表')
                    write_voucher_sheet(vws, name, kind, entities, [y], gl_rows,
                                        gl_full=gl_full, detail_total=detail_total,
                                        filter_year=ys)
            else:
                vws = out.create_sheet('凭证抽查表')
                write_voucher_sheet(vws, name, kind, entities, [y], gl_rows,
                                    gl_full=gl_full, detail_total=detail_total,
                                    filter_year=ys)
        # 4) 其他收益 ↔ 递延收益摊销核对（当年）
        if recon_gl is not None and name == '其他收益':
            ows = out.create_sheet('其他收益与递延收益摊销核对')
            write_other_income_recon_sheet(ows, recon_gl, entities, [y], gl_agg)
        # 5) 审定表置首位 + 保存
        if name not in VOUCHER_EXCLUDE:
            _shs = out._sheets
            _asn = f'{name} 审定表'
            _ixa = next((i for i, s in enumerate(_shs) if s.title == _asn), None)
            if _ixa is not None and _ixa > 0:
                _shs.insert(0, _shs.pop(_ixa))
        _fw_pl(out)
        # ⚡ 2026-08-28 #875：对方科目核对（模块内集成）
        try:
            from counterparty_recon import inject_into_wb, SPECS as _CR_SPECS
            if name in _CR_SPECS:
                inject_into_wb(out, gl_full, ys, _CR_SPECS[name], ents=set(entities))
        except Exception as _ex:
            print(f'  ⚠️ {name}对方科目核对注入失败：{_ex}')
        out_path = os.path.join(data_dir, f'{name}审计底稿_{ys}_生成.xlsx')
        from audit_common import validate_workbook
        validate_workbook(out, '利润表底稿', raise_on_error=False)
        out.save(out_path)
        out.close()
        saved.append(out_path)
        print(f'  ✓ {name} {ys} 已生成')
    wb_t.close()
    return saved




# ---------------- 主流程 ----------------
def build_pl_combined(data_dir):
    print(f'>> 处理文件夹：{data_dir}')
    entities = _discover_entities(data_dir)
    if not entities:
        print('  ❌ 未发现任何账套主体（需含《科目余额表》+《综合查询明细表》）。')
        return []
    years = sorted({y for b in entities.values() for y in b})
    try:
        from mask_dict import mask_names as _mn
        _ents = _mn(sorted(entities))
    except Exception:
        _ents = sorted(entities)
    print(f'  · 发现主体 {len(entities)} 个：' + ', '.join(_ents))
    print(f'  · 年份：' + ', '.join(years))
    n_km = sum(1 for b in entities.values() for v in b.values() if v.get('km'))
    n_gl = sum(1 for b in entities.values() for v in b.values() if v.get('gl'))
    print(f'  · 含《科目余额表》主体 {n_km} 个；含《综合查询明细表》(GL) 主体 {n_gl} 个。')
    if n_gl == 0:
        print('  ❌ 未发现任何《综合查询明细表》(.xlsx)！利润表科目的"发生额"必须来自 GL。'
              '请检查：① 导出文件名是否含"综合查询明细表"字样；② 是否为 .xlsx（不是 .xls 旧格式）。'
              '两者皆不满足时，所有科目都会被判"无发生额"而跳过——表现为"没底稿"。')
    elif n_gl < n_km:
        print(f'  ⚠️ 有 {n_km - n_gl} 个主体缺 GL 文件，这些主体的利润表科目将无发生额。')
    gl_agg, gl_rows = read_pl_gl(data_dir, entities, years)
    if n_gl == 0:
        # 2026-08-06：JTt 等无 GL 账套 → TB 借贷发生额兜底（否则全部损益科目跳过）
        gl_agg, gl_rows = _tb_gl_fallback(data_dir, entities, years, gl_agg, gl_rows)
    tb = read_tb_control(data_dir, entities, years)
    recon_gl = read_recon_gl(data_dir, entities, years)
    gl_full = read_gl_rows(data_dir, entities)  # 全量 GL（含非损益科目分录），用于凭证抽查"整笔凭证"聚合判定规则②
    produced = []
    for code, name, kind in PL_SUBJECTS:
        has = any(k[0] == name for k in gl_agg) or any(k[0] == name for k in gl_rows)
        # 2026-08-04 复核：GL 无发生额时 TB 兜底生成（税金及附加等在账套 TB 有 6403 一级科目时
        # 也生成底稿，避免"试算表有数-底稿未生成"）。
        # 2026-08-06 修复：TB 兜底原用 tb.get((e, name, y))——read_tb_control 键为 (e, subj, y)，
        # subj 是 _subject_of 归一化名；且 _subject_of 原 startswith 匹配对 FY『营业务税金及附加』
        # 失败 → TB 行被过滤、兜底恒 False → 税金及附加底稿未生成（TB 有 197.8 万）。
        if not has:
            _tb_has = any(k[1] == name for k in tb
                          if k[0] in entities and str(k[2]) in {str(y) for y in years})
            if _tb_has:
                has = True
        if not has:
            print(f'  ⊘ [{code} {name}] 本文件夹综合查询明细表无发生额，跳过。')
            continue
        try:
            paths = build_subject_workbook(data_dir, code, name, kind, entities, years,
                                           gl_agg, gl_rows, tb, recon_gl, gl_full)
            produced.extend(paths)
            print(f'  ✓ [{code} {name}] 生成：' + '；'.join(os.path.basename(p) for p in paths))
        except Exception as ex:
            import traceback as _tb
            print(f'  ✗ [{code} {name}] 生成失败（已跳过该科目，不影响其他科目）：{ex}')
            print('     ' + ''.join(_tb.format_exception_only(type(ex), ex)).strip())
            continue
    return produced


def main(argv=None):
    """用法: pl_detail.py [数据文件夹] [--input 数据文件夹]
    支持位置参数与 --input/-i 旗标，便于与其他程序统一调用方式。"""
    argv = argv if argv is not None else sys.argv[1:]
    data_dir = None
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ('-h', '--help'):
            print('用法: pl_detail.py [数据文件夹路径] [--input|-i 数据文件夹路径]')
            return
        elif a in ('--input', '-i'):
            if i + 1 < len(argv):
                data_dir = argv[i + 1]; i += 2; continue
            else:
                print('ERROR: --input 需要跟一个文件夹路径'); return
        elif not a.startswith('-'):
            data_dir = a
        i += 1
    if data_dir is None:
        data_dir = input('请输入账套导出文件夹路径：').strip().strip('"')
    if not os.path.isdir(data_dir):
        print(f'ERROR: 文件夹不存在：{data_dir}')
        return
    produced = build_pl_combined(data_dir)
    print(f'\n完成，共生成 {len(produced)} 份文件：')
    for p in produced:
        print('  ', p)
    from audit_common import finalize_after_build
    finalize_after_build(data_dir)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）


def _find_managed_python():
    """定位 WorkBuddy 托管 Python（含 openpyxl），避免系统 Python 无 openpyxl 导致拖入闪退。"""
    base = os.path.join(os.environ.get('USERPROFILE', os.path.expanduser('~')),
                        '.workbuddy', 'binaries', 'python', 'versions')
    if os.path.isdir(base):
        for d in sorted(os.listdir(base), reverse=True):
            p = os.path.join(base, d, 'python.exe')
            if os.path.isfile(p):
                return p
    return None



# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None

if __name__ == '__main__':
    # 拖入/双击本 .py 时，若当前不是托管 Python（如系统 Python 3.14 无 openpyxl），
    # 自动用托管 Python 重新执行并保留控制台，避免 ImportError 黑窗闪退。
    _mp = _find_managed_python()
    if _mp and os.path.abspath(sys.executable).lower() != os.path.abspath(_mp).lower():
        import subprocess
        try:
            rc = subprocess.run([_mp, '-B', os.path.abspath(__file__)] + sys.argv[1:]).returncode
        except Exception as _e:
            print(f'[ERROR] 无法启动托管 Python：{_e}')
            rc = 1
        try:
            input('\n按回车退出…')
        except EOFError:
            pass
        sys.exit(rc)
    try:
        rc = main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        import datetime
        traceback.print_exc()
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        try:
            if os.environ.get('AUDIT_LOG') == '1':
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       f'pl_crash_{ts}.log'), 'w', encoding='utf-8') as _f:
                    _f.write('未捕获异常：\n' + traceback.format_exc())
                print(f'已记录崩溃日志：pl_crash_{ts}.log')
        except Exception:
            pass
        rc = 1
    try:
        if sys.stdin.isatty():
            input('\n按回车退出…')
    except (EOFError, Exception):
        pass
    sys.exit(rc)