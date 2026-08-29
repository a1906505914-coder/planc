# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=职工薪酬审计底稿_生成.xlsx(2025+2026合并·5sheet) | 关键列=应付职工薪酬子目贷方计提/成本费用工资借发/期初期末增减变动 | 职责=薪酬计提↔成本费用↔个税三方勾稽(独立生成器)

from audit_common import _safe_save  # 共享库：统一锁感知保存（单一来源）
"""职工薪酬明细表生成程序（桌面版，参照应交税费底稿架构）
================================================================================
用法：
  python payroll_detail.py --input <文件夹> [--output <输出.xlsx>] [--period Y]
  · 自动发现文件夹内各核算主体的「科目余额表」与「综合查询明细表」，按主体×年合并。
  · 2025 与 2026 两年合并输出一份工作簿：<数据目录>/职工薪酬审计底稿_生成.xlsx
  · 主表数字：应付职工薪酬(2211) 期初/期末取自科目余额表；计提(贷方净)/发放(借方)取自
    综合查询明细表(GL)按子目×月份；列支费用类别采用 v7 已验证的 hybrid 取值
    （GL 凭证级 + 科目余额表 leaf 级，逐明细取更接近计提者）。
  · 含 6 张表（参照应交税费底稿合并结构）：
    职工薪酬审定表(2年并列，含 集团汇总数块) / 职工薪酬明细表(原计提明细表+增减变动)
    / 分月计提明细表(原每月计提统计表) / 薪酬分配核对表(原薪酬明细表+差异判定，含 集团汇总数块)
    / 多借多贷核对表(凭证级借贷对照，含 集团汇总数块+两表差异对照) / 差异凭证清单(差异汇总)。
  · 职工薪酬明细表增列「期初金额/期末金额/增减变动额/增减变动比例/备注」，
    增减变动比例>10% 自动标记「提请注意」。

数据口径（同 职工薪酬明细表_2025定.xlsx 编制说明）：
  · 计提 = 应付职工薪酬 贷方净发生额（红字冲减以负数记入贷方自然调减）；发放 = 正数借方。
  · 列支 = 费用/成本科目中薪酬明细的借方发生额（hybrid 双路径逐明细择优）。
  · 期末(计算) = 期初 + 计提 - 发放。
  · 核对 = 列支合计 - 计提数（真实账面差异，透明列示，不塞补差桶）。
================================================================================
"""
import os
import re
import sys
import argparse

# ---- 启动诊断日志（绝对路径，纯内置库，任何崩溃前必写，用于定位"闪退无痕迹"）----
def _write_boot_log(tag):
    import os as _bl_os
    if _bl_os.environ.get('AUDIT_LOG') != '1':
        return
    try:
        import datetime
        _p = os.path.join(os.path.expanduser('~'), 'Desktop', 'payroll_boot.log')
        with open(_p, 'a', encoding='utf-8') as _f:
            _f.write('[%s] %s sys=%s argv=%s\n' % (
                datetime.datetime.now().isoformat(timespec='seconds'), tag, sys.executable, sys.argv))
    except Exception:
        pass

_write_boot_log('LOAD')
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


import openpyxl
from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_SUB_FONT,
                         SHELL_NUM, SHELL_LEFT, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                         resolve_template, apply_header_row, AUDIT_TEMPLATES_DIR, finalize_workbook)
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------- 常量 ----------------
DETAILS = ['工资', '养老保险', '医疗保险', '失业保险', '工伤保险', '住房公积金',
           '职工福利费', '工会经费', '职工教育经费', '劳务派遣费用']
CANON = ['生产成本', '制造费用', '管理费用', '研发费用', '销售费用', '主营业务成本',
         '在建工程', '工程施工', '劳务成本']
CATS = CANON + ['其他列支']
MONTHS = list(range(1, 13))

# ---------------- 明细归类（v7 已验证） ----------------
def norm_detail(name):
    if not name:
        return None
    n = str(name)
    # 2026-08-01 问题9：Z 母公司研发工资记账为「借 研发费用-计提薪酬 贷 应付职工薪酬」
    # 计提、发放时红字冲回并按人员重列「研发费用-…-工资薪金」。「计提薪酬」本质即工资计提，
    # 必须识别为工资，否则计提凭证内 28.78M 研发计提薪酬不进列支（问题9 差异主因）。
    if '计提薪酬' in n or '工资薪金' in n:
        return '工资'
    if '养老' in n or '年金' in n:
        return '养老保险'
    if '医疗' in n or '生育' in n:
        return '医疗保险'
    if '失业' in n:
        return '失业保险'
    if '工伤' in n:
        return '工伤保险'
    if '住房' in n or '公积' in n:
        return '住房公积金'
    if '劳务' in n:
        return '劳务派遣费用'
    if '工会' in n:
        return '工会经费'
    if '教育' in n:
        return '职工教育经费'
    last = n.split('-')[-1]
    if '福利' in last:
        return '职工福利费'
    # 工资薪酬主体：末段==『职工薪酬』的子目（如 应付职工薪酬-职工薪酬 / GL 的 …-职工薪酬）。
    # 末段精确匹配可避开父级『应付职工薪酬』(末段为『应付职工薪酬』) 与 『社会保险费』父级，
    # 避免把汇总行误并入工资。另覆盖费用侧 『工资薪金』(如 生产成本-…-工资薪金)。
    if last == '职工薪酬' or '工资薪金' in n:
        return '工资'
    # 辞退/解除劳动关系补偿（辞退福利）并入工资（属职工薪酬范畴；金额通常很小）。
    if '解除劳动关系补偿' in n or '辞退福利' in n or ('辞退' in n and '薪酬' in n):
        return '工资'
    if '工资' in last or '奖金' in last or '绩效' in last or '补贴' in last:
        return '工资'
    if '工资' in n or '奖金' in n or '绩效' in n or '补贴' in n:
        return '工资'
    if '福利' in n:
        return '职工福利费'
    # ⚡ 2026-08-10：『其他短期薪酬』（SAP 1010 期初 -468,894.69）——属短期薪酬范畴，
    # 归入工资桶（与审定表合计口径一致，防明细表合计≠审定表；原 None → 被过滤丢失）。
    if '其他短期薪酬' in n or '其他福利' in n:
        return '工资'
    return None

# ⚡ 2026-08-26 配置化收敛：SAP 费用特判码收口 sap_subject_map（与 expense_detail 同源）
try:
    from sap_subject_map import fee_code as _sap_fee_code, fr_category as _sap_fr_cat
except Exception:
    _sap_fee_code = lambda k: {'rd': '660006', 'finance': '6603', 'pool': '6600',
                               'fx_loss': '660303', 'fx_gain': '660304'}.get(k)
    _sap_fr_cat = lambda fr: {'6601': '销售费用', '6602': '管理费用', '5101': '制造费用',
                              '5301': '研发费用'}.get(str(fr or '').strip())
_RD_CODE = _sap_fee_code('rd') or '660006'
_FIN_CODE = _sap_fee_code('finance') or '6603'
_POOL_CODE = _sap_fee_code('pool') or '6600'


# ---------------- 费用类别归类（v7 已验证） ----------------
def cat_of(km, fr=''):
    """费用借方科目 → 列支类别（生产成本/制造费用/管理费用/研发费用/销售费用/…）。
    fr: ⚡ 2026-08-10 SAP 功能范围列（col19）——SAP 独立费用科目（『工资及奖金』6600010000/
    『工资附加费-养老保险』660002*/『研究开发费-人员人工』660006*/『保密补贴』660036* 等）
    名称无 U8『管理费用-』式前缀，原逻辑全部 None → 1010 工资列支仅 2,238 万 vs 计提 3.16 亿
    （用户指出"列支数没找全"）。按功能范围分类（与 expense_detail._sap_fee_cat 同构）。"""
    s = str(km)
    if s.startswith(('应付职工薪酬', '应付账款', '应付票据', '其他应付款', '应付利息',
                     '应付股利', '预收账款', '合同负债', '其他应收款', '预付账款')):
        return None
    if s.startswith('生产成本'):
        return '生产成本'
    if s.startswith('制造费用'):
        return '制造费用'
    if s.startswith('管理费用'):
        return '管理费用'
    if s.startswith('销售费用'):
        return '销售费用'
    if s.startswith('研发支出') or s.startswith('研发费'):
        return '研发费用'
    if s.startswith('主营业务成本') or s.startswith('营业成本'):
        return '主营业务成本'
    if s.startswith('在建工程'):
        return '在建工程'
    if s.startswith('工程施工'):
        return '工程施工'
    if s.startswith('劳务成本'):
        return '劳务成本'
    if ('费用' in s or '成本' in s or '支出' in s):
        return '其他列支'
    # ⚡ 2026-08-10 SAP 独立费用科目（6600 系列）：名称无『费用/成本』前缀但属薪酬/费用列支，
    # 按功能范围列分类（660006/研究开发费→研发；fr 6601销售/6602管理/5101制造/5301研发；缺省管理）
    if fr or s.startswith(_POOL_CODE) or s.startswith(_FIN_CODE):
        if s.startswith(_RD_CODE) or '研究开发费' in s:
            return '研发费用'
        _cat = _sap_fr_cat(fr)
        if _cat:
            return _cat
        if s.startswith(_POOL_CODE):
            return '管理费用'
    return None
    return None

# 对方科目是否指向另一成本/费用类科目（识别凭证内跨费用科目重分类行）
_CANON_CP = ['生产成本', '制造费用', '管理费用', '销售费用', '研发支出', '研发费',
             '主营业务成本', '营业成本', '在建工程', '工程施工', '劳务成本']
PAY_KW = ['职工薪酬', '工资', '福利', '社保', '养老', '医疗', '失业',
          '工伤', '住房', '公积', '工会', '教育', '劳务', '工资及福利费']


def safe(x):
    return float(x) if x is not None else 0.0


def _cp_is_reclass(cp):
    s = str(cp)
    return any(k in s for k in _CANON_CP)


def _parse_embedded_cats(cp):
    s = str(cp)
    found = []
    for c in CANON + ['其他列支']:
        if c in s and c not in found:
            found.append(c)
    return found


# ---------------- 实体发现（folder 模式，兼容 年+主体 / 主体+年 / 仅2位年 / 时间戳误标 多种命名） ----------------
def _extract_year(fn):
    """从文件名稳健识别会计年度（移植自 pl_detail.py，已验证）。

    常见陷阱：25 年账套导出文件带 2026 时间戳（如 『…_20260723152733.xlsx』），
    若用 re.search(r'(20\\d{2})', fn) 会先命中时间戳 2026 而非会计年度 2025；
    另有一些文件直接用『25』作年份后缀且无 20xx（如 『母公司25科目余额表.xlsx』），
    旧逻辑因找不到 20xx 而整文件跳过，导致主体丢失。
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


def _discover_payroll_entities(data_dir):
    # ⚡ 2026-08-09 治本统一：本地 discover 收敛为 audit_common.discover_entities——
    # U8 模式读原文件全量；SAP 模式 patch 后=adapter（current_comp 单主体过滤内置）。
    from audit_common import discover_entities as _de
    return _de(data_dir)


# ---------------- GL 读取（凭证级） ----------------
def read_gl_payroll(gl_path):
    rec = {
        'monthly_credit': {d: defaultdict(float) for d in DETAILS},
        'monthly_debit': {d: defaultdict(float) for d in DETAILS},
        'exp_gl': {d: {c: 0.0 for c in CATS} for d in DETAILS},
        'gl_accrual': defaultdict(float), 'gl_pay': defaultdict(float),
        'daifu_accrual': defaultdict(float), 'daifu_pay': defaultdict(float),
        'red_letter': [], 'direct_pay': [], 'daifu': [],
        # ⚡ 2026-08-13 差异归因（按明细 detail 汇总，供核对表「差异判定/原因」结构化归因）：
        #   attr_direct=直接付款未经2211（借费用 贷银行/其他 → 列支>计提 构成）
        #   attr_daifu=代扣代垫个人部分（发放 借2211 贷其他应收款 → 列支<计提 构成）
        #   attr_red_cnt=红字冲回重列凭证数（冲销计提+重列工资，双向已含净额，仅提示）
        #   attr_embed=嵌入式计提分流金额（计提行贷方按对方科目分配，已进 exp_gl，仅提示）
        'attr_direct': defaultdict(float), 'attr_daifu': defaultdict(float),
        'attr_red_cnt': defaultdict(int), 'attr_embed': defaultdict(float),
        # 2026-08-02 多借多贷核对表数据：凭证级借贷方（不依赖末级一借一贷匹配）
        'xdeb': defaultdict(float),   # (借方费用科目, 月) -> 金额  借方列支侧
        'xcre': defaultdict(float),   # (贷方计提明细, 月) -> 金额  贷方计提侧
    }
    # SAP：从 adapter 干净行构建 rows（字段映射，逻辑同 U8 文件读取）
    if _adapter is not None and (gl_path is None or not isinstance(gl_path, str) or _adapter.is_sap(gl_path)):
        # ⚡⚡ 2026-08-25 性能（AH 88 主体 payroll 34 分钟主因）：集团模式 read_gl() 全量
        #   455 万行，每主体都全量构造 rows dict → 88×12s 重复劳动。两遍法预过滤：
        #   第一遍只收集【含 2211/应付职工薪酬】的凭证 vkey（纯字符串判断，不构造 dict）；
        #   第二遍仅对命中凭证的行构造 rows（保留同凭证的对方费用科目/摘要行——直接付款、
        #   多借多贷归因依赖凭证级全部行）。命中行通常几千行，dict 构造开销可忽略。
        _gl_all = _adapter.read_gl()
        _vkeys = set()
        for _r in _gl_all:
            _km0 = _r.get('name')
            if _km0 and ('2211' in str(_km0) or '应付职工薪酬' in str(_km0)):
                _vkeys.add((str(_r.get('date') or ''), str(_r.get('vtype') or ''),
                            str(_r.get('vno') or '')))
        rows = []
        for _r in _gl_all:
            if (str(_r.get('date') or ''), str(_r.get('vtype') or ''),
                    str(_r.get('vno') or '')) not in _vkeys:
                continue
            km = str(_r.get('name') or '')
            dt = str(_r.get('date') or '')
            zi = str(_r.get('vtype') or '')
            hao = str(_r.get('vno') or '')
            cp = str(_r.get('cp') or '')
            jf = float(_r.get('debit') or 0.0)
            df = float(_r.get('credit') or 0.0)
            sm = str(_r.get('sm') or '')
            month = int(_r.get('month') or 0) or None
            vkey = (dt, zi, hao)
            # ⚡ 2026-08-10 保留功能范围列（col19）——SAP 独立费用科目按 FR 分类列支
            rows.append(dict(km=km, cp=cp, jf=jf, df=df, month=month, vkey=vkey,
                             dt=dt, zi=zi, hao=hao, summary=sm,
                             fr=str(_r.get('fr') or '')))
    else:
        if not gl_path or not os.path.exists(gl_path):
            return rec
        wb = openpyxl.load_workbook(gl_path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = []
        _seen = set()  # 源文件行级去重：部分账套导出时每行复制×2（如 FY深圳 2026），需精确去重
        for r in ws.iter_rows(min_row=3, values_only=True):
            if not r or len(r) < 8:
                continue
            km = str(r[0]) if r[0] is not None else ''
            dt = str(r[1]) if len(r) > 1 and r[1] is not None else ''
            zi = str(r[2]) if len(r) > 2 and r[2] is not None else ''
            hao = str(r[3]) if len(r) > 3 and r[3] is not None else ''
            cp = str(r[4]) if len(r) > 4 and r[4] is not None else ''
            jf = safe(r[6]) if len(r) > 6 else 0.0
            df = safe(r[7]) if len(r) > 7 else 0.0
            sm = str(r[5]) if len(r) > 5 and r[5] is not None else ''
            _dk = (dt, zi, hao, km, jf, df, sm)
            if _dk in _seen:
                continue
            _seen.add(_dk)
            month = None
            m = re.search(r'(\d{4})-(\d{2})', dt)
            if m:
                month = int(m.group(2))
            vkey = (dt, zi, hao)
            rows.append(dict(km=km, cp=cp, jf=jf, df=df, month=month, vkey=vkey,
                             dt=dt, zi=zi, hao=hao, summary=sm))
        wb.close()
    vgroups = defaultdict(list)
    for row in rows:
        vgroups[row['vkey']].append(row)

    for vkey, vrows in vgroups.items():
        pay_lines = [x for x in vrows if x['km'].startswith('应付职工薪酬')]
        exp_all = [x for x in vrows if cat_of(x['km'], x.get('fr', '')) and x['jf'] != 0]
        has_pay = len(pay_lines) > 0
        pay_credit = any(x['df'] != 0 for x in pay_lines)
        is_red = any((x['df'] < 0 or x['jf'] < 0) for x in vrows
                     if x['km'].startswith('应付职工薪酬') or cat_of(x['km'], x.get('fr', '')))

        # --- 计提 / 发放 / 代付 ---
        for x in pay_lines:
            d = norm_detail(x['km'])
            if not d:
                continue
            is_daifu = '其他应收款' in x['cp']
            if is_daifu:
                if x['df'] != 0:
                    rec['daifu_accrual'][d] += x['df']
                if x['jf'] != 0:
                    rec['daifu_pay'][d] += x['jf']
            rec['gl_accrual'][d] += x['df']
            if x['month'] is not None:
                rec['monthly_credit'][d][x['month']] += x['df']
            # 发放取【净】借方发生额（自然符号，红字冲回/错账调整负借方须冲减），
            # 与《科目余额表》借发(净额)勾稽一致——2026-07-31 修复（原 jf>0 只取正数，
            # 红字借方不冲减导致发放虚增，期末推算与 TB 对不上）
            if x['jf'] != 0:
                if x['month'] is not None:
                    rec['monthly_debit'][d][x['month']] += x['jf']
                rec['gl_pay'][d] += x['jf']

        # --- exp_gl 列支（计提凭证内的费用借方行） ---
        if pay_credit:
            # 2026-08-02 多借多贷核对表：凭证级借贷方汇总（不分配、不配平，展示原貌）。
            # 借方 = 费用科目(首段) 按月；贷方 = 应付职工薪酬明细 按月。两侧合计应=借贷平衡，
            # 差异即异常凭证（无法末级匹配的部分，供审计定位真正异常）。
            for x in vrows:
                if x['month'] is None:
                    continue
                if x['jf'] != 0 and cat_of(x['km'], x.get('fr', '')):
                    rec['xdeb'][(cat_of(x['km'], x.get('fr', '')), x['month'])] += x['jf']
                if x['df'] != 0 and x['km'].startswith('应付职工薪酬'):
                    _d = norm_detail(x['km']) or '其他'
                    rec['xcre'][(_d, x['month'])] += x['df']
            pay_sub = defaultdict(float)
            for x in pay_lines:
                if x['df'] != 0:
                    d = norm_detail(x['km'])
                    if d:
                        pay_sub[d] += x['df']
            pay_sub_total = sum(pay_sub.values())

            def allocate(x):
                c = cat_of(x['km'], x.get('fr', ''))
                d_line = norm_detail(x['km'])
                if d_line:
                    rec['exp_gl'][d_line][c] += x['jf']
                elif pay_sub_total > 0:
                    for d2, amt in pay_sub.items():
                        rec['exp_gl'][d2][c] += x['jf'] * amt / pay_sub_total
                else:
                    rec['exp_gl']['工资'][c] += x['jf']

            picked = [x for x in exp_all
                      if ('应付职工薪酬' in x['cp'])
                      or (norm_detail(x['km']) and _cp_is_reclass(x['cp']))]
            if picked:
                for x in picked:
                    allocate(x)
            elif exp_all:
                for x in exp_all:
                    if norm_detail(x['km']) or any(k in x['km'] for k in PAY_KW):
                        allocate(x)
            else:
                # 嵌入式记账：计提行贷方按对方科目类别分配列支
                for x in pay_lines:
                    if x['df']:
                        cats = _parse_embedded_cats(x['cp'])
                        if cats:
                            d2 = norm_detail(x['km']) or norm_detail(x['cp']) or '工资'
                            mv = x['df']
                            # ⚡ 2026-08-13 归因：嵌入式计提分流金额（按 detail，仅提示——已进 exp_gl）
                            rec['attr_embed'][d2] += mv
                            real = [c for c in cats if c != '其他列支']
                            if len(cats) == 1:
                                rec['exp_gl'][d2][cats[0]] += mv
                            elif real:
                                for c in real:
                                    rec['exp_gl'][d2][c] += mv / len(real)
                            else:
                                rec['exp_gl'][d2]['其他列支'] += mv

        # --- 差异凭证穿透（红字冲减 / 直接付款 / 代付） ---
        if is_red and has_pay:
            rl = [dict(km=x['km'], cp=x['cp'], jf=x['jf'], df=x['df'])
                  for x in vrows if (x['km'].startswith('应付职工薪酬') or cat_of(x['km'], x.get('fr', '')))]
            rec['red_letter'].append({'dt': vrows[0]['dt'], 'hao': vrows[0]['hao'],
                                      'summary': vrows[0]['summary'], 'lines': rl})
            # ⚡ 2026-08-13 归因：红字冲回重列凭证计数（按费用行 detail；提示性，不重复计金额）
            for _rl in rl:
                if cat_of(_rl['km'], _rl.get('fr', '')) or '应付职工薪酬' in _rl['km']:
                    _dl = norm_detail(_rl['km'])
                    if _dl:
                        rec['attr_red_cnt'][_dl] += 1
        if not pay_credit:
            # 对"无贷方计提"的凭证——含 ①直接列支（凭证无应付职工薪酬）与 ②发放凭证
            # （借 应付职工薪酬 贷 银行，无贷方计提，但其内常有「红字冲回计提薪酬 + 按人员重列工资」）：
            # 其薪酬费用借方行即为真实列支/冲回，按自然符号计入 exp_gl。
            # 2026-08-01 问题9修复（Z 母公司）：研发/销售等工资在发放凭证中
            # 「红字冲销计提薪酬 + 按人员重列工资薪金」，该凭证贷方为银行/其他应收款/应交税费
            # （无贷方应付职工薪酬）→ 旧逻辑仅 not has_pay 分支且仅 has_intra_ap 时补计，
            # 导致母公司研发工资 24.64M 列支缺失、计提 168.96M >> 列支 144.29M。
            # 现改为无条件计入 exp_gl（含红字负数冲减，与计提凭证正数配平）：
            #   · 正数 jf（工资薪金重列/直接列支）→ 计入列支；
            #   · 负数 jf（红字冲销计提薪酬/福利费）→ 冲减列支。
            # ⚠️ 排除【内部结转凭证】：贷方为成本费用类科目（如 借 生产成本-职工薪酬 贷 制造费用
            # -职工薪酬 的制造费用结转），此类借方非真实列支、与制造费用工资双计，须跳过。
            cr_is_cost = any(x['df'] != 0 and cat_of(x['km'], x.get('fr', '')) for x in vrows)
            if not cr_is_cost:
                for x in exp_all:
                    d_line = norm_detail(x['km'])
                    if d_line and x['jf'] != 0:
                        c = cat_of(x['km'], x.get('fr', ''))
                        if c:
                            rec['exp_gl'][d_line][c] += x['jf']
                        if x['jf'] > 0:
                            rec['direct_pay'].append({'dt': x['dt'], 'hao': x['hao'], 'summary': x['summary'],
                                                      'km': x['km'], 'cp': x['cp'], 'jf': x['jf'],
                                                      'detail': d_line})
                            # ⚡ 2026-08-13 归因：直接付款金额（按 detail，列支>计提 构成）
                            if d_line:
                                rec['attr_direct'][d_line] += x['jf']
        if pay_credit and any('其他应收款' in x['cp'] for x in pay_lines):
            for x in pay_lines:
                if x['df'] != 0 and '其他应收款' in x['cp']:
                    rec['daifu'].append({'dt': x['dt'], 'hao': x['hao'], 'summary': x['summary'],
                                         'detail': norm_detail(x['km']), 'amount': x['df'], 'km': x['km']})
                    # ⚡ 2026-08-13 归因：代扣代垫金额（按 detail，列支<计提 构成）
                    _dl = norm_detail(x['km'])
                    if _dl:
                        rec['attr_daifu'][_dl] += x['df']
    return rec


# ---------------- 科目余额表读取 ----------------
def read_tb_payroll(tb_path):
    rec = {
        'details': {d: {'open': 0.0, 'accrual': 0.0, 'pay': 0.0, 'close': 0.0} for d in DETAILS},
        'exp_tb': {d: {c: 0.0 for c in CATS} for d in DETAILS},
        'withhold': {}, 'tb_accrual': defaultdict(float), 'tb_pay': defaultdict(float),
        'parent_withhold': 0.0,
        'unalloc_open': 0.0, 'unalloc_close': 0.0,  # TB 父级 2211 未拆分明细的期初/期末余额
    }
    # SAP：从 adapter.read_km 聚合 2211 明细（期初/计提贷/发放借/期末 → details 类别）
    # ⚡ 2026-08-10 修复：read_km 只含 TB 子目级（level=4）——SAP TB 余额只在父级 2211、
    #   子目期初/期末全 0（余额不拆分子目）→ 明细表期初全 0 ≠ 审定表（_collect_pay_tb 按名称
    #   取父级余额分配）。改从 _adapter.read_tb_full 取（与审定表同源，名称匹配），
    #   父级 2211 期初/期末按名称归入标准明细（未拆分余额记 unalloc，明细表披露）。
    if _adapter is not None and (tb_path is None or not isinstance(tb_path, str)
                                 or _adapter.is_sap(tb_path)):
        _c_cur = _adapter.current_comp()
        tb_full = _adapter.read_tb_full(None, None)
        rows2211 = []
        # ⚡⚡ 2026-08-26 修复（任务878 职工薪酬 3.35 亿差异根因）：集团模式（current_comp=None）
        #   run_u8_on_sap 注入 _GROUP_COMPS（--only 主体集）——read_tb_full 现返回全量，
        #   必须在此按 _GROUP_COMPS 过滤，否则全 88 主体 TB 聚合进集团底稿（1010 工资
        #   TB 贷方 6.50 亿 vs GL 计提 3.16 亿，假差异 -3.35 亿）。
        _gc_comp = set(getattr(_adapter, '_GROUP_COMPS', None) or []) if not _c_cur else None
        for (ee, c, n, yy), v in tb_full.items():
            if _c_cur and ee != _c_cur:
                continue   # 单主体驱动：仅当前主体（防混入 88 家）
            if _gc_comp and ee not in _gc_comp:
                continue   # 集团模式：仅本集团主体
            if not (str(c).startswith('2211') or '应付职工薪酬' in str(n)):
                continue
            rows2211.append((str(c), str(n), v, str(yy)))
        if not rows2211:
            return rec
        years = sorted({yy for _, _, _, yy in rows2211})
        for yy in years:
            dset = {}
            tot1 = [0.0] * 4
            for c, n, v, yyy in rows2211:
                if yyy != yy:
                    continue
                if c == '2211' or str(n).strip() == '应付职工薪酬':
                    # ⚡⚡ 2026-08-25 统一符号（与 _collect_pay_tb 同口径）：open=-qc、close=-qm
                    #   （贷余正/借余负）。原 abs 在方向混合子目（部分借余）时 ≠ 父级净额
                    #   → 明细表 TB期末(原)/审定表期末 虚增（AH 1020 应付职工薪酬 597.4万
                    #   vs TB 586.6万，差 10.9万 = 借余子目双算）。带符号后子目合计=父级净额。
                    tot1[0] += -float(v.get('qc') or 0.0)
                    tot1[1] += float(v.get('df') or 0.0)
                    tot1[2] += float(v.get('jf') or 0.0)
                    tot1[3] += -float(v.get('qm') or 0.0)
            # 子目级（非父级 2211）：按名称归入标准明细（与审定表 _collect_pay_tb 同口径）
            for c, n, v, yyy in rows2211:
                if yyy != yy:
                    continue
                if c == '2211' or str(n).strip() == '应付职工薪酬':
                    continue   # 父级：余额不拆分子目时兜底计入 unalloc（下方处理）
                d = norm_detail(n)
                if d is None or d not in DETAILS:
                    d = '其他'
                arr = dset.setdefault(d, [0.0, 0.0, 0.0, 0.0])
                arr[0] += -float(v.get('qc') or 0.0)
                arr[1] += float(v.get('df') or 0.0)
                arr[2] += float(v.get('jf') or 0.0)
                arr[3] += -float(v.get('qm') or 0.0)
            # 父级余额在子目无对应时的兜底：父级期初/期末 - 已归子目合计 = 未拆余额 → 记入工资桶
            # （SAP 父级 2211 余额实为全部明细合计；子目发生额正常，余额挂父级）
            sub_open = sum(a[0] for a in dset.values())
            sub_close = sum(a[3] for a in dset.values())
            gap_open = max(0.0, abs(tot1[0]) - sub_open)
            gap_close = max(0.0, abs(tot1[3]) - sub_close)
            if gap_open > 0.005 or gap_close > 0.005:
                dset.setdefault('工资', [0.0, 0.0, 0.0, 0.0])
                dset['工资'][0] += gap_open
                dset['工资'][3] += gap_close
            for d, arr in dset.items():
                if d not in rec['details']:
                    continue   # '其他' 等非标准类别：仅披露用，不落 details
                rec['details'][d]['open'] += arr[0]
                rec['details'][d]['accrual'] += arr[1]
                rec['details'][d]['pay'] += arr[2]
                rec['details'][d]['close'] += arr[3]
                rec['tb_accrual'][d] += arr[1]
                rec['tb_pay'][d] += arr[2]
            # 父级未拆分余额（兜底后仍有差）记 unalloc（核对表披露）
            rec['unalloc_open'] += max(0.0, abs(tot1[0]) - sum(a[0] for a in dset.values()))
            rec['unalloc_close'] += max(0.0, abs(tot1[3]) - sum(a[3] for a in dset.values()))
        return rec
    if not tb_path or not os.path.exists(tb_path):
        return rec
    wb = openpyxl.load_workbook(tb_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    tb_rows = [r for r in ws.iter_rows(min_row=3, values_only=True) if r and len(r) >= 7]
    wb.close()
    names = [str(r[1]) for r in tb_rows if r[1] is not None]
    name_set = set(names)
    # 该主体 TB 是否展开二级明细（如 2211.01 应付职工薪酬-基本工资）？
    # 是 → 父级 2211 为汇总行（余额=子目之和），父级余额不再单独入"工资"桶（避免父+子双计）；
    # 否 → 父级期初/期末整额入"工资"桶（FY 账套未拆二级时的权宜，披露于核对表）。
    has_sub = any(nm.startswith('应付职工薪酬') and nm != '应付职工薪酬' and '-' in nm
                  for nm in names)

    def is_leaf(nm):
        prefix = nm + '-'
        for other in name_set:
            if other != nm and other.startswith(prefix):
                return False
        return True

    for r in tb_rows:
        nm = str(r[1]) if r[1] is not None else ''
        if nm.startswith('应付职工薪酬'):
            d = norm_detail(nm)
            if d and d in rec['details']:
                rec['details'][d]['open'] += safe(r[3]) if (len(r) > 3 and r[2] == '借') else (safe(r[3]) if (len(r) > 3 and r[2] == '贷') else safe(r[3]))
                rec['details'][d]['accrual'] += safe(r[5])
                rec['details'][d]['pay'] += safe(r[4])
                rec['details'][d]['close'] += safe(r[7]) if (len(r) > 7 and r[6] == '借') else (safe(r[7]) if (len(r) > 7 and r[6] == '贷') else safe(r[7]))
                rec['tb_accrual'][d] += safe(r[5])
                rec['tb_pay'][d] += safe(r[4])
            elif nm == '应付职工薪酬' and '工资' in rec['details'] and not has_sub:
                # TB 仅列父级 2211（FY 账套未拆二级明细）→ 期初/期末余额记入"工资"桶，
                # 保证 期末(计算)=期初+计提-发放 与 TB 父级期末勾稽；发生额不记
                # （GL 明细已按二级覆盖，父级借/贷发=GL 明细之和，避免双计）。
                # has_sub=True（TB 已展开二级明细）时父级余额由二级行覆盖，此处跳过防双计。
                rec['details']['工资']['open'] += safe(r[3])
                rec['details']['工资']['close'] += safe(r[7])
                rec['unalloc_open'] = rec.get('unalloc_open', 0.0) + safe(r[3])
                rec['unalloc_close'] = rec.get('unalloc_close', 0.0) + safe(r[7])
        if nm.startswith('其他应收款-应收代垫款项-'):
            rec['withhold'][nm] = safe(r[4])
        elif nm == '其他应收款-应收代垫款项':
            rec['parent_withhold'] = safe(r[4])
        c = cat_of(nm)
        if c and is_leaf(nm):
            d = norm_detail(nm)
            if d and d in rec['details']:
                rec['exp_tb'][d][c] += safe(r[4])
    if not rec['withhold'] and rec['parent_withhold']:
        rec['withhold']['其他应收款-应收代垫款项'] = rec['parent_withhold']
    return rec


# ---------------- 合并实体记录（hybrid 列支 + 四舍五入） ----------------
def build_entity_record(gl, tb):
    rec = {
        'details': tb['details'],
        'monthly_credit': gl['monthly_credit'], 'monthly_debit': gl['monthly_debit'],
        'exp_gl': gl['exp_gl'], 'exp_tb': tb['exp_tb'],
        'exp': {d: {c: 0.0 for c in CATS} for d in DETAILS},
        'exp_src': {}, 'withhold': tb['withhold'],
        'tb_accrual': tb['tb_accrual'], 'tb_pay': tb['tb_pay'],
        'gl_accrual': gl['gl_accrual'], 'gl_pay': gl['gl_pay'],
        'daifu_accrual': gl['daifu_accrual'], 'daifu_pay': gl['daifu_pay'],
        'red_letter': gl['red_letter'], 'direct_pay': gl['direct_pay'], 'daifu': gl['daifu'],
        # ⚡ 2026-08-15 修复：差异归因字段传递（2026-08-13 加字段时漏带 → _fill_merged
        #   rec['attr_direct'] KeyError，XBJ 薪酬底稿炸）
        'attr_direct': gl.get('attr_direct') or defaultdict(float),
        'attr_daifu': gl.get('attr_daifu') or defaultdict(float),
        'attr_red_cnt': gl.get('attr_red_cnt') or defaultdict(int),
        'attr_embed': gl.get('attr_embed') or defaultdict(float),
        'xdeb': gl['xdeb'], 'xcre': gl['xcre'],   # 2026-08-02 多借多贷核对表凭证级数据
    }
    for d in DETAILS:
        acc = rec['gl_accrual'].get(d, 0.0)
        gl_tot = sum(rec['exp_gl'][d].values())
        tb_tot = sum(rec['exp_tb'][d].values())
        if abs(tb_tot - acc) < abs(gl_tot - acc) - 1e-6:
            src = rec['exp_tb'][d]
            rec['exp_src'][d] = 'tb'
        else:
            src = rec['exp_gl'][d]
            rec['exp_src'][d] = 'gl'
        for c in CATS:
            rec['exp'][d][c] = src.get(c, 0.0)
    return rec


# ================= 输出工作簿 =================
HFILL = PatternFill('solid', fgColor='DDEBF7')
HFONT = Font(name='Times New Roman', bold=True, color='000000', size=10)
TFILL = PatternFill('solid', fgColor='D9E1F2')
SFILL = PatternFill('solid', fgColor='F2F2F2')
BFONT = Font(name='Times New Roman', bold=True)
TITLE_FONT = SHELL_TITLE_FONT
SUB_FONT = SHELL_SUB_FONT
RED_FONT = Font(name='Times New Roman', bold=True, color='9C0006')
WARN_FILL = PatternFill('solid', fgColor='FCE4D6')
NUM = SHELL_NUM
THIN = Side(style='thin', color='BFBFBF')
BORDER = SHELL_BORDER
CEN = SHELL_CEN
LEFT = SHELL_LEFT
RGT = SHELL_RGT


def fnum(ws, r, c, v):
    cell = ws.cell(r, c, v if v is not None else None)
    if isinstance(v, (int, float)) and c not in (1, 2):
        cell.number_format = NUM
    cell.border = BORDER
    return cell


PERIOD = {
    '2025': '2025-01 至 2025-12（完整会计年度）',
    '2026': '2026-01 至 2026-05（综合查询明细表导出期间；科目余额表为较后期快照）',
}



def _build_merged_frame(wb):
    """构建合并版职工薪酬工作簿框架（6 张 sheet，参照应交税费底稿）：
       0. 职工薪酬审定表(2025-2026)   1. 职工薪酬明细表
       2. 分月计提明细表              3. 薪酬分配核对表   4. 多借多贷核对表   5. 差异凭证清单
    样式统一采用 audit_shell 的 SHELL_* 常量。运行时 _fill_merged 只填数据。"""
    # ========== Sheet 1 职工薪酬明细表（原计提明细表 + 增减变动） ==========
    ws2 = wb.create_sheet('职工薪酬明细表')
    ws2['A1'] = '职工薪酬明细表（集团合并·按主体×明细，2025/2026）— 期初/计提/发放/期末 + 增减变动'
    ws2['A1'].font = TITLE_FONT
    head2 = ['年份', '核算主体', '明细', '期初金额', '本期计提(GL贷方净)', '本期发放(GL借方)',
             '期末金额(计算)', 'TB期末(原)', '平衡校验', 'TB本期贷方', 'GL计提', '控制差异(GL-TB)',
             '列支来源', '增减变动额', '增减变动比例', '备注']
    apply_header_row(ws2, 2, head2)
    ws2.freeze_panes = 'D3'
    for j, w in enumerate([8, 12, 14, 14, 16, 14, 14, 14, 12, 14, 14, 14, 10, 14, 12, 30], 1):
        ws2.column_dimensions[get_column_letter(j)].width = w

    # ========== Sheet 3 分月计提明细表（原每月计提统计表） ==========
    ws3 = wb.create_sheet('分月计提明细表')
    ws3['A1'] = '分月计提明细表（集团合并·按主体，2025/2026）— 分月计提/支付明细'
    ws3['A1'].font = TITLE_FONT
    head3 = ['年份', '核算主体', '明细', '期初余额'] + [f'{m}月' for m in MONTHS] + ['计提合计', '期末余额(计算)']
    apply_header_row(ws3, 2, head3)
    ws3.freeze_panes = 'D3'
    for j, w in enumerate([8, 12, 14, 14] + [9] * 12 + [14, 14], 1):
        ws3.column_dimensions[get_column_letter(j)].width = w

    # ========== Sheet 4 薪酬分配核对表（原薪酬明细表 + 差异判定/原因） ==========
    ws4 = wb.create_sheet('薪酬分配核对表')
    ws4['A1'] = '薪酬分配核对表（集团合并·按主体，2025/2026）— 薪酬计提 × 成本费用分配 + 核对'
    ws4['A1'].font = TITLE_FONT
    head4 = ['年份', '核算主体', '明细', '期初余额', '计提', '发放', '期末余额(计算)'] + \
            CANON + ['其他列支', '列支合计', '计提数', '核对', '差异判定/原因']
    apply_header_row(ws4, 2, head4)
    ws4.freeze_panes = 'D3'
    for j, w in enumerate([8, 12, 14, 14, 14, 14, 16] + [12] * len(CANON) + [14, 14, 14, 12, 50], 1):
        ws4.column_dimensions[get_column_letter(j)].width = w

    # ========== Sheet 5 多借多贷核对表（2026-08-02 用户方法论：无法末级一借一贷匹配时的借贷对照） ==========
    ws6 = wb.create_sheet('多借多贷核对表')
    ws6['A1'] = ('多借多贷核对表（集团合并·按主体×年，凭证级借贷对照）— 借方=列支费用科目、贷方=计提明细，'
                 '两侧按 分月+全年 汇总对照；不依赖末级一借一贷匹配，借贷合计之差即异常凭证线索')
    ws6['A1'].font = TITLE_FONT
    head6 = ['年份', '核算主体', '项目', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '全年合计']
    # ⚡ 2026-08-11 N3：双行表头分组（规则2）——行2 大类（月份），行3 列名；数据起始行同步改 4
    for _g, _c0, _c1 in [('基础信息', 1, 3), ('月份', 4, 15), ('全年合计', 16, 16)]:
        ws6.merge_cells(start_row=2, start_column=_c0, end_row=2, end_column=_c1)
        _gc = ws6.cell(2, _c0, _g)
        _gc.font = TITLE_FONT
        _gc.fill = PatternFill('solid', fgColor='DDEBF7')
        _gc.alignment = Alignment(horizontal='center', vertical='center')
        for _j in range(_c0, _c1 + 1):
            ws6.cell(2, _j).border = BORDER
    apply_header_row(ws6, 3, head6)
    ws6.freeze_panes = 'D4'
    for j, w in enumerate([8, 14, 22] + [10] * 12 + [14], 1):
        ws6.column_dimensions[get_column_letter(j)].width = w

    # ========== Sheet 6 差异凭证清单（按 计提 vs 列支 差异方向整理） ==========
    ws5 = wb.create_sheet('差异凭证清单')
    ws5['A1'] = '差异凭证清单（计提 vs 列支 差异分析）— 差异=计提数−列支数；计提>列支 列示原因，列支>计提 多为直接列支（2026-08-02 起不再展开凭证穿透）'
    ws5['A1'].font = TITLE_FONT
    head5 = ['年份', '核算主体', '明细', '计提数(GL贷方净)', '列支数(成本费用)', '差异(计提-列支)', '差异方向', '原因/说明']
    apply_header_row(ws5, 2, head5)
    ws5.freeze_panes = 'A3'
    for j, w in enumerate([8, 12, 14, 16, 16, 16, 12, 50, 16, 16], 1):
        ws5.column_dimensions[get_column_letter(j)].width = w


def _fill_merged(wb, years, ents, records_by_year):
    """把 2025/2026 两年数据填入已建好框架的合并工作簿（6 张 sheet）。
    仅写数据行，不重写标题/表头/列宽/冻结。"""
    issues = []

    # ========== Sheet 2 职工薪酬明细表 ==========
    ws2 = wb['职工薪酬明细表']
    r = 3
    for y in years:
        for e in ents:
            rec = records_by_year[y].get(e)
            if not rec:
                continue
            for d in DETAILS:
                openv = rec['details'][d]['open']
                acc = rec['gl_accrual'].get(d, 0.0)
                pay = rec['gl_pay'].get(d, 0.0)
                # 期末金额（计算）= 期初 + 本期计提(GL贷方净) - 本期发放(GL借方)，与分月计提表/审定表一致；
                # 末级 TB 余额（多数账套仅父级 2211 带余额、子目为 0）仅作勾稽参考列『TB期末(原)』。
                close_calc = openv + acc - pay
                close_tb = rec['details'][d]['close']
                # 平衡校验：GL 计算期末 与 TB 末级原值之差（原公式恒为 0，属 bug，已修正为真实勾稽差）
                bal = round(close_calc - close_tb, 2)
                tb_cr = rec['tb_accrual'].get(d, 0.0)
                ctrl = round(acc - tb_cr, 2)
                # 增减变动（GL 计算期末 - 期初金额），比例超 10% 提请注意
                chg = close_calc - openv
                if openv != 0:
                    pct = chg / abs(openv)
                    note = '提请注意：增减变动%.2f%%' % (pct * 100) if abs(pct) > 0.10 else '—'
                else:
                    pct = 0.0
                    note = ('期初为0，增减变动%.2f（无法计算比例）' % chg) if chg != 0 else '—'
                row = [y, e, d, openv, acc, pay, round(close_calc, 2), round(close_tb, 2), bal,
                       round(tb_cr, 2), round(acc, 2), ctrl, rec.get('exp_src', {}).get(d, ''),
                       round(chg, 2), round(pct, 4), note]
                # ⚡⚡ 2026-08-25 性能优化：ws.append 批量写行替代 fnum 逐 cell 创建
                #   （append 只写值 O(1) 追加；cell 对象已存在，后置样式定位 O(1) 设属性）。
                ws2.append(row)
                # ⚡⚡ 2026-08-25 修复：勿用 max_row 定位（O(全 cells) 遍历，大表 O(n²)）；
                #   r 为循环计数，append 连续写入同号行。
                for j, v in enumerate(row, 1):
                    cell = ws2.cell(r, j)
                    if isinstance(v, (int, float)) and j not in (1, 2):
                        cell.number_format = NUM
                    cell.border = BORDER
                    # 控制差异(GL-TB)：GL 计提 vs TB 本期贷方差异 → 提醒
                    if j == 12 and isinstance(v, (int, float)) and abs(v) > 0.005:
                        cell.fill = WARN_FILL
                    # 平衡校验：仅当 TB 末级确有余额且与计算值不符时才标红（末级 TB=0 时属正常，不预警）
                    if j == 9 and isinstance(v, (int, float)) and abs(v) > 0.005 and close_tb != 0:
                        cell.fill = WARN_FILL
                    if j == 15:
                        cell.number_format = '0.00%'
                    if j == 16:
                        cell.alignment = LEFT
                        if note.startswith('提请注意') or note.startswith('期初为0'):
                            cell.fill = WARN_FILL
                            cell.font = RED_FONT
                r += 1

    # ===== 全集团合计（Sheet2 明细表，供『明细合计 vs 审定表』勾稽核对）=====
    # ⚡⚡ 2026-08-25 规范性补充：审计程序说明/勾稽检查依赖明细表合计行，原职工薪酬
    #   明细表无合计行 → audit_checker WARN + 明细-审定核对无法执行。加"全集团合计"。
    if len(ents) > 1:
        r = ws2.max_row + 2
        ws2.cell(r, 1, '全集团合计（分年，与审定表期末核对相符）').font = SUB_FONT
        r += 1
        for y in years:
            recs = records_by_year[y]
            _es = [e for e in ents if e in recs]
            if not _es:
                continue
            _o = sum(recs[e]['details'][d]['open'] for e in _es for d in DETAILS)
            _a = sum(recs[e]['gl_accrual'].get(d, 0.0) for e in _es for d in DETAILS)
            _p = sum(recs[e]['gl_pay'].get(d, 0.0) for e in _es for d in DETAILS)
            _cl = _o + _a - _p
            _ctb = sum(recs[e]['details'][d]['close'] for e in _es for d in DETAILS)
            _tbc = sum(recs[e]['tb_accrual'].get(d, 0.0) for e in _es for d in DETAILS)
            row = [y, '全集团合计', '合计', _o, _a, _p, round(_cl, 2), round(_ctb, 2),
                   round(_cl - _ctb, 2), round(_tbc, 2), round(_a, 2),
                   round(_a - _tbc, 2), '', round(_cl - _o, 2), 0.0, '']
            ws2.append(row)
            r = ws2.max_row
            for j, v in enumerate(row, 1):
                cell = ws2.cell(r, j)
                if isinstance(v, (int, float)) and j not in (1, 2):
                    cell.number_format = NUM
                cell.border = BORDER
                cell.fill = TFILL
                cell.font = BFONT
                if j == 15:
                    cell.number_format = '0.00%'

    # ========== Sheet 3 分月计提明细表 ==========
    ws3 = wb['分月计提明细表']
    r = 3
    for y in years:
        for e in ents:
            rec = records_by_year[y].get(e)
            if not rec:
                continue
            for d in DETAILS:
                mc = rec['monthly_credit'].get(d, {})
                md = rec['monthly_debit'].get(d, {})
                s_c = sum(mc.get(m, 0.0) for m in MONTHS)
                s_d = sum(md.get(m, 0.0) for m in MONTHS)
                close_d = rec['details'][d]['open'] + s_c - s_d
                row = [y, e, d, rec['details'][d]['open']] + [mc.get(m, 0.0) for m in MONTHS] + [s_c, close_d]
                # ⚡⚡ 2026-08-25 性能优化：append 批量写行（同 Sheet2）；勿用 max_row（O(n²)）
                ws3.append(row)
                for j, v in enumerate(row, 1):
                    fnum(ws3, r, j, v)
                r += 1
            # 计提小计（与 Sheet2 本期计提=GL贷方净 一致）
            r = ws3.max_row + 1
            fnum(ws3, r, 1, y)
            fnum(ws3, r, 2, e)
            fnum(ws3, r, 3, '计提小计').font = BFONT
            fnum(ws3, r, 4, sum(rec['details'][d]['open'] for d in DETAILS))
            for j, m in enumerate(MONTHS, 5):
                fnum(ws3, r, j, sum(rec['monthly_credit'].get(d, {}).get(m, 0.0) for d in DETAILS))
            fnum(ws3, r, 5 + 12, sum(rec['gl_accrual'].get(d, 0.0) for d in DETAILS))
            fnum(ws3, r, 5 + 13, sum(rec['details'][d]['open'] + rec['gl_accrual'].get(d, 0.0) - rec['gl_pay'].get(d, 0.0) for d in DETAILS))
            for j in range(1, ws3.max_column + 1):
                ws3.cell(r, j).fill = SFILL
                ws3.cell(r, j).font = BFONT
            r += 1
            # 支付小计（与 Sheet2 本期发放=GL借方 一致）
            r = ws3.max_row + 1
            fnum(ws3, r, 1, y)
            fnum(ws3, r, 2, e)
            fnum(ws3, r, 3, '支付小计').font = BFONT
            fnum(ws3, r, 4, None)
            for j, m in enumerate(MONTHS, 5):
                fnum(ws3, r, j, sum(rec['monthly_debit'].get(d, {}).get(m, 0.0) for d in DETAILS))
            fnum(ws3, r, 5 + 12, sum(rec['gl_pay'].get(d, 0.0) for d in DETAILS))
            fnum(ws3, r, 5 + 13, None)
            for j in range(1, ws3.max_column + 1):
                ws3.cell(r, j).fill = SFILL
                ws3.cell(r, j).font = BFONT
            r += 1
    # 集团合计（全主体·分年）
    if len(ents) > 1:  # ⚡ 2026-08-10 单体裁剪（铁律81：1 主体不输出集团合计）
        r += 1
        ws3.cell(r, 1, '集团合计（全主体·分年，与职工薪酬明细表借贷方数核对相符）').font = SUB_FONT
        r += 1
        for y in years:
            recs = records_by_year[y]
            fnum(ws3, r, 1, y)
            fnum(ws3, r, 2, '集团合计')
            fnum(ws3, r, 3, '计提合计').font = BFONT
            fnum(ws3, r, 4, sum(recs[e]['details'][d]['open'] for e in ents if e in recs for d in DETAILS))
            for j, m in enumerate(MONTHS, 5):
                fnum(ws3, r, j, sum(recs[e]['monthly_credit'].get(d, {}).get(m, 0.0) for e in ents if e in recs for d in DETAILS))
            fnum(ws3, r, 5 + 12, sum(recs[e]['gl_accrual'].get(d, 0.0) for e in ents if e in recs for d in DETAILS))
            fnum(ws3, r, 5 + 13, sum(recs[e]['details'][d]['open'] + recs[e]['gl_accrual'].get(d, 0.0) - recs[e]['gl_pay'].get(d, 0.0) for e in ents if e in recs for d in DETAILS))
            for j in range(1, ws3.max_column + 1):
                ws3.cell(r, j).fill = TFILL
                ws3.cell(r, j).font = BFONT
            r += 1
            fnum(ws3, r, 1, y)
            fnum(ws3, r, 2, '集团合计')
            fnum(ws3, r, 3, '支付合计').font = BFONT
            fnum(ws3, r, 4, None)
            for j, m in enumerate(MONTHS, 5):
                fnum(ws3, r, j, sum(recs[e]['monthly_debit'].get(d, {}).get(m, 0.0) for e in ents if e in recs for d in DETAILS))
            fnum(ws3, r, 5 + 12, sum(recs[e]['gl_pay'].get(d, 0.0) for e in ents if e in recs for d in DETAILS))
            fnum(ws3, r, 5 + 13, None)
            for j in range(1, ws3.max_column + 1):
                ws3.cell(r, j).fill = TFILL
                ws3.cell(r, j).font = BFONT
            r += 1

    # ⚡⚡ 2026-08-28 口径披露（用户："审定表本期减少与分月计提的减少数不一致"）：
    #   本表计提/支付取【GL 明细口径】（凭证级分月，GL 含红字/内部结转净额）；审定表
    #   本期增加/本期减少取【科目余额表 TB 控制数】。两源存在口径差（GL 红字冲回、
    #   跨期、分类取数差异），以审定表（TB）为准，本表供分月分析。差异明细见
    #   《薪酬分配核对表/差异凭证清单》。
    r += 1
    ws3.cell(r, 1, '口径说明：本表计提/支付 = GL 明细口径（凭证级分月，含红字净额）；'
                   '《职工薪酬审定表》本期增加/减少 = 科目余额表（TB）控制数口径。'
                   '两者差异为 GL 与 TB 口径差，以审定表（TB）为准；差异明细见薪酬分配核对表/差异凭证清单。').font = SUB_FONT

    # ========== Sheet 4 薪酬分配核对表 ==========
    ws4 = wb['薪酬分配核对表']
    r = 3
    ncols = ws4.max_column
    for y in years:
        for e in ents:
            rec = records_by_year[y].get(e)
            if not rec:
                continue
            _detail_lies = []
            _detail_chks = []
            for d in DETAILS:
                acc = rec['gl_accrual'].get(d, 0.0)
                pay = rec['gl_pay'].get(d, 0.0)
                openv = rec['details'][d]['open']
                closev = openv + acc - pay
                other = round(rec['exp'][d]['其他列支'], 2)
                lie = round(sum(rec['exp'][d][c] for c in CANON) + other, 2)
                chk = round(lie - acc, 2)
                # ⚡ 2026-08-13 差异归因：差异 = 直接付款 − 代扣代垫 + 未归因（红字/嵌入式为提示项）
                _direct = rec['attr_direct'].get(d, 0.0)
                _daifu = rec['attr_daifu'].get(d, 0.0)
                _red_n = rec['attr_red_cnt'].get(d, 0)
                _embed = rec['attr_embed'].get(d, 0.0)
                _explained = round(_direct - _daifu, 2)
                _unexp = round(chk - _explained, 2)
                reasons = []
                if abs(chk) > max(1.0, abs(acc) * 0.005):
                    if abs(_direct) > 0.005:
                        _p = '（占计提%.2f%%）' % (_direct / acc * 100) if acc else '（计提为0）'
                        reasons.append(f'直接付款未经2211 {_direct:,.2f}{_p}')
                    if abs(_daifu) > 0.005:
                        _p = '（占计提%.2f%%）' % (_daifu / acc * 100) if acc else '（计提为0）'
                        reasons.append(f'代扣代垫个人部分 {_daifu:,.2f}{_p}')
                    if _red_n:
                        reasons.append(f'红字冲回重列{_red_n}张凭证')
                    if abs(_embed) > 0.005:
                        reasons.append(f'嵌入式计提分流 {_embed:,.2f}')
                    if abs(_unexp) > max(1.0, abs(acc) * 0.005):
                        reasons.append(f'另有未归因差异 {_unexp:,.2f} 需查证')
                    if not reasons:
                        reasons.append(f'差异 {chk:,.2f} 未分类，需查证')
                    pct = '（占计提%.2f%%）' % (chk / acc * 100) if acc else '（计提为0）'
                    reasons.append('合计差' + pct)
                    issues.append({'level': 'WARN', 'type': '核对', 'message':
                        f'[{e}][{y}][{d}] 列支合计{lie:.2f} 与 计提数{acc:.2f} 差{chk:.2f}{pct}'
                        f'（归因：直接付款{_direct:,.2f}/代扣代垫{_daifu:,.2f}/未归因{_unexp:,.2f}）'})
                note = '；'.join(reasons) if reasons else '—'
                row = [y, e, d, openv, acc, pay, closev] + [round(rec['exp'][d][c], 2) for c in CANON] + \
                      [other, lie, acc, chk, note]
                _detail_lies.append(lie)
                _detail_chks.append(chk)
                for j, v in enumerate(row, 1):
                    cell = fnum(ws4, r, j, v)
                    if j == ncols:
                        cell.alignment = LEFT
                        if note != '—':
                            cell.fill = WARN_FILL
                r += 1
            # 主体小计
            acc = sum(rec['gl_accrual'].get(d, 0.0) for d in DETAILS)
            pay = sum(rec['gl_pay'].get(d, 0.0) for d in DETAILS)
            openv = sum(rec['details'][d]['open'] for d in DETAILS)
            closev = openv + acc - pay
            # ⚡ 2026-08-11 修复 0.01 浮点：小计须=Σ明细行【已写入 round 后值】，勿重算
            #   （薪酬分配核对表 1010 生产成本 明细Σ=27396279.97 vs 小计 27396279.96 → audit_checker ERROR；
            #   二次修复：列支合计/核对 列小计直接引用明细行 lie/chk，防口径不一致差 0.01）
            cats_sub = {c: sum(round(rec['exp'][d][c], 2) for d in DETAILS) for c in CANON}
            other_sub = round(sum(round(rec['exp'][d]['其他列支'], 2) for d in DETAILS), 2)
            lie = round(sum(_detail_lies), 2)
            chk_sub = round(sum(_detail_chks), 2)
            row = [y, e, '小计', openv, acc, pay, closev] + [round(cats_sub[c], 2) for c in CANON] + \
                  [other_sub, lie, acc, chk_sub, '—']
            for j, v in enumerate(row, 1):
                fnum(ws4, r, j, v)
            for j in range(1, ncols + 1):
                ws4.cell(r, j).fill = SFILL
                ws4.cell(r, j).font = BFONT
            r += 1

    # 披露：TB 应付职工薪酬仅列父级 2211（未拆二级明细）的主体的期初/期末余额
    _unalloc_lines = []
    for y in years:
        for e in ents:
            rec = records_by_year[y].get(e)
            if rec and (abs(rec.get('unalloc_open', 0.0)) > 0.01 or abs(rec.get('unalloc_close', 0.0)) > 0.01):
                _unalloc_lines.append(
                    f"[{e}][{y}] 应付职工薪酬TB仅列父级2211（期初{rec['unalloc_open']:,.2f}/期末{rec['unalloc_close']:,.2f}），"
                    f"余额未拆二级明细，统一计入『工资』列；本期计提/发放仍按GL二级明细列示。")
    if _unalloc_lines:
        r += 1
        _txt4 = ws4.cell(r, 1, '注：' + '；'.join(_unalloc_lines))
        _txt4.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        r += 1

    # ========== 集团汇总数（薪酬分配核对表，2026-08-02：内容同上面各主体，值=全部主体同行之和） ==========
    # ⚡ 2026-08-10 单体裁剪（铁律81）：SAP 逐主体底稿（1 主体）不输出集团汇总数
    if len(ents) <= 1:
        r += 1
        _txt4 = ws4.cell(r, 1, '注：单体底稿（1 个核算主体），不输出集团汇总数。')
        _txt4.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        r += 1
    else:
        r += 1
        _txt4 = ws4.cell(r, 1, '集团汇总数（内容同上面各主体：各明细行+小计；值=全部核算主体之和）')
        _txt4.font = Font(name='Times New Roman', size=10, bold=True, color='000000')
        ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        r += 1
        for y in years:
            recs_y = {e: records_by_year[y].get(e) for e in ents}
            for d in DETAILS:
                g_open = sum(recs_y[e]['details'][d]['open'] for e in ents if recs_y[e])
                g_acc = sum(recs_y[e]['gl_accrual'].get(d, 0.0) for e in ents if recs_y[e])
                g_pay = sum(recs_y[e]['gl_pay'].get(d, 0.0) for e in ents if recs_y[e])
                g_close = g_open + g_acc - g_pay
                g_cats = {c: sum(round(recs_y[e]['exp'][d][c], 2) for e in ents if recs_y[e]) for c in CANON}
                g_other = round(sum(round(recs_y[e]['exp'][d]['其他列支'], 2) for e in ents if recs_y[e]), 2)
                g_lie = round(sum(g_cats.values()) + g_other, 2)
                g_chk = round(g_lie - g_acc, 2)
                reasons = []
                if abs(g_chk) > max(1.0, abs(g_acc) * 0.005):
                    if g_chk < 0:
                        reasons.append('列支<计提：含代扣代垫个人部分净额或未带薪酬明细科目')
                    else:
                        reasons.append('列支>计提：存在直接付款未经应付职工薪酬')
                    pct = '（占计提%.2f%%）' % (g_chk / g_acc * 100) if g_acc else '（计提为0）'
                    reasons[-1] += pct
                note = '；'.join(reasons) if reasons else '—'
                row = [y, '集团汇总数', d, g_open, g_acc, g_pay, g_close] + [round(g_cats[c], 2) for c in CANON] + \
                      [g_other, g_lie, g_acc, g_chk, note]
                for j, v in enumerate(row, 1):
                    cell = fnum(ws4, r, j, v)
                    if j == ncols:
                        cell.alignment = LEFT
                        if note != '—':
                            cell.fill = WARN_FILL
                r += 1
            # 集团小计（全明细）
            g_open = sum(recs_y[e]['details'][d]['open'] for e in ents if recs_y[e] for d in DETAILS)
            g_acc = sum(recs_y[e]['gl_accrual'].get(d, 0.0) for e in ents if recs_y[e] for d in DETAILS)
            g_pay = sum(recs_y[e]['gl_pay'].get(d, 0.0) for e in ents if recs_y[e] for d in DETAILS)
            g_close = g_open + g_acc - g_pay
            # ⚡ 2026-08-11 集团小计同源修复：列支合计=Σ各主体小计 lie、核对=Σ各主体小计 chk（round 后）
            _g_lies = []
            _g_chks = []
            for _e2 in ents:
                if not recs_y[_e2]:
                    continue
                _g_cats_e = {c: sum(round(recs_y[_e2]['exp'][d][c], 2) for d in DETAILS) for c in CANON}
                _g_oth_e = round(sum(round(recs_y[_e2]['exp'][d]['其他列支'], 2) for d in DETAILS), 2)
                _g_lie_e = round(sum(_g_cats_e.values()) + _g_oth_e, 2)
                _g_lies.append(_g_lie_e)
                _g_chks.append(round(_g_lie_e - sum(recs_y[_e2]['gl_accrual'].get(d, 0.0) for d in DETAILS), 2))
            g_cats = {c: sum(round(recs_y[e]['exp'][d][c], 2) for e in ents if recs_y[e] for d in DETAILS) for c in CANON}
            g_other = round(sum(round(recs_y[e]['exp'][d]['其他列支'], 2) for e in ents if recs_y[e] for d in DETAILS), 2)
            g_lie = round(sum(_g_lies), 2)
            g_chk_sub = round(sum(_g_chks), 2)
            row = [y, '集团汇总数', '小计', g_open, g_acc, g_pay, g_close] + [round(g_cats[c], 2) for c in CANON] + \
                  [g_other, g_lie, g_acc, g_chk_sub, '—']
            for j, v in enumerate(row, 1):
                fnum(ws4, r, j, v)
            for j in range(1, ncols + 1):
                ws4.cell(r, j).fill = TFILL
                ws4.cell(r, j).font = BFONT
            r += 1

# ========== Sheet 5 差异凭证清单（按 计提 vs 列支 差异方向整理） ==========
    ws5 = wb['差异凭证清单']
    r = 3
    head_sum5 = ['年份', '核算主体', '明细', '计提数(GL贷方净)', '列支数(成本费用)', '差异(计提-列支)', '差异方向', '原因/说明']
    # ---- Block 1：差异汇总（按 核算主体×明细×年） ----
    ws5.cell(r, 1, '差异汇总（按 核算主体×明细×年）：差异 = 计提数 − 列支数').font = SUB_FONT
    r += 1
    for i, h in enumerate(head_sum5, 1):
        c = ws5.cell(r, i, h)
        c.font = HFONT
        c.fill = HFILL
        c.border = BORDER
        c.alignment = CEN
    r += 1
    for y in years:
        for e in ents:
            rec = records_by_year[y].get(e)
            if not rec:
                continue
            for d in DETAILS:
                acc = rec['gl_accrual'].get(d, 0.0)
                lie = round(sum(rec['exp'][d][c] for c in CANON) + round(rec['exp'][d]['其他列支'], 2), 2)
                diff = round(acc - lie, 2)
                if abs(diff) <= 0.005:
                    direction, reason = '—', '—'
                elif diff > 0:
                    direction = '计提>列支'
                    reasons = ['应付职工薪酬保留贷方余额（计提未发放/挂账）']
                    if rec['daifu']:
                        reasons.append('含集团内部代付')
                    if rec.get('withhold'):
                        reasons.append('含代扣代垫个人款')
                    reason = '；'.join(reasons) + '（需查凭证）'
                else:
                    direction = '列支>计提'
                    reason = '直接列支（不经应付职工薪酬），预期内，暂不展开'
                row = [y, e, d, round(acc, 2), lie, diff, direction, reason]
                for j, v in enumerate(row, 1):
                    cell = fnum(ws5, r, j, v)
                    if j == 6 and isinstance(v, (int, float)) and diff > 0.005:
                        cell.fill = WARN_FILL
                    if j == 8:
                        cell.alignment = LEFT
                r += 1
    if '2026' in years:
        issues.append({'level': 'INFO', 'type': '控制数',
                       'message': '2026年度：科目余额表为较后期快照，GL 计提与 TB 本期贷方差异为真实期后差异，非取数错误。'})

    # ========== Sheet 6 多借多贷核对表（凭证级借贷对照，不依赖末级一借一贷匹配） ==========
    _fill_multidebit_credit(wb, years, ents, records_by_year)

    return issues


def _fill_multidebit_credit(wb, years, ents, records_by_year):
    """多借多贷核对表：每主体×年，行=借方列支科目(生产成本/制造费用/管理/销售/研发/其他列支)，
    列=1-12月+全年合计；下方对应行=贷方计提明细(工资/养老/医疗/失业/工伤/住房/福利/工会/教育/劳务/其他)。
    借方合计 与 贷方合计 之差 = 异常凭证线索（多借多贷无法末级配平时的净差，供审计定位）。
    2026-08-02 用户方法论：差异过大时不再强求末级一借一贷匹配，改为借贷两侧分月合计对照。"""
    ws6 = wb['多借多贷核对表']
    # 借方费用科目顺序（cat_of 归类全集）
    DEBIT_CATS = ['生产成本', '制造费用', '管理费用', '销售费用', '研发费用',
                  '主营业务成本', '在建工程', '工程施工', '劳务成本', '其他列支']
    r = 4   # ⚡ 2026-08-11 N3：表头加分组行后数据从行4（原 3）
    for y in years:
        for e in ents:
            rec = records_by_year[y].get(e)
            if not rec:
                continue
            xdeb = rec.get('xdeb') or {}
            xcre = rec.get('xcre') or {}
            # 借方行（列支科目）
            _dsum = {}
            for (c, m), v in xdeb.items():
                _dsum.setdefault(c, [0.0] * 13)
                _dsum[c][m - 1] += v
                _dsum[c][12] += v
            for c in DEBIT_CATS:
                if c not in _dsum or abs(sum(_dsum[c])) < 0.005:
                    continue
                row = [y, e, '借：' + c] + [round(v, 2) for v in _dsum[c]]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                ws6.cell(r, 3).font = Font(name='Times New Roman', bold=True)
                r += 1
            # 借方合计
            d_tot = [0.0] * 13
            for c, arr in _dsum.items():
                for m in range(13):
                    d_tot[m] += arr[m]
            if any(abs(v) > 0.005 for v in d_tot):
                row = [y, e, '借：列支合计'] + [round(v, 2) for v in d_tot]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                for j in range(1, 17):
                    ws6.cell(r, j).fill = TFILL
                    ws6.cell(r, j).font = BFONT
                r += 1
            # 贷方行（计提明细）
            _csum = {}
            for (d, m), v in xcre.items():
                _csum.setdefault(d, [0.0] * 13)
                _csum[d][m - 1] += v
                _csum[d][12] += v
            for d in DETAILS:
                if d not in _csum or abs(sum(_csum[d])) < 0.005:
                    continue
                row = [y, e, '贷：' + d] + [round(v, 2) for v in _csum[d]]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                ws6.cell(r, 3).font = Font(name='Times New Roman', bold=True)
                r += 1
            # 贷方合计
            c_tot = [0.0] * 13
            for d, arr in _csum.items():
                for m in range(13):
                    c_tot[m] += arr[m]
            if any(abs(v) > 0.005 for v in c_tot):
                row = [y, e, '贷：计提合计'] + [round(v, 2) for v in c_tot]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                for j in range(1, 17):
                    ws6.cell(r, j).fill = TFILL
                    ws6.cell(r, j).font = BFONT
                r += 1
            # 借贷差（异常线索）
            diff_row = [d_tot[m] - c_tot[m] for m in range(13)]
            if any(abs(v) > 0.005 for v in diff_row):
                row = [y, e, '差：借−贷'] + [round(v, 2) for v in diff_row]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                for j in range(1, 17):
                    ws6.cell(r, j).fill = WARN_FILL
                    ws6.cell(r, j).font = BFONT
                # 差异说明
                r += 1
                ddiff = d_tot[12] - c_tot[12]
                if abs(ddiff) <= max(1.0, abs(d_tot[12]) * 0.005):
                    note = '全年借=贷（配平）；分月差异系跨月结转/期初余额所致，非异常'
                elif ddiff > 0:
                    note = '借方列支 > 贷方计提：存在直接付款/代扣代垫个人部分（不经应付职工薪酬），查凭证'
                else:
                    note = '贷方计提 > 借方列支：计提未全额分配（保留应付职工薪酬贷方余额），查凭证'
                cell = fnum(ws6, r, 3, '→ ' + note)
                cell.alignment = LEFT
                ws6.merge_cells(start_row=r, start_column=3, end_row=r, end_column=16)
                r += 1
            r += 1   # 主体间空行

    # ========== 集团汇总数（多借多贷核对表，2026-08-02：内容同上面各主体，值=全部主体之和） ==========
    # ⚡ 2026-08-10 单体裁剪（铁律81）：SAP 逐主体底稿（1 主体）不输出集团汇总数
    if len(ents) <= 1:
        r += 1
        c = ws6.cell(r, 1, '注：单体底稿（1 个核算主体），不输出集团汇总数。')
        c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
        r += 1
    else:
        r += 1
        c = ws6.cell(r, 1, '集团汇总数（内容同上面各主体：借：各费用科目 / 借：列支合计 / 贷：各明细 / 贷：计提合计 / 差：借−贷；值=全部核算主体之和）')
        c.font = Font(name='Times New Roman', size=10, bold=True, color='000000')
        ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
        r += 1
        for y in years:
            g_dsum = {}
            g_csum = {}
            for e in ents:
                rec = records_by_year[y].get(e)
                if not rec:
                    continue
                for (cc, m), v in (rec.get('xdeb') or {}).items():
                    g_dsum.setdefault(cc, [0.0] * 13)
                    g_dsum[cc][m - 1] += v
                    g_dsum[cc][12] += v
                for (dd, m), v in (rec.get('xcre') or {}).items():
                    g_csum.setdefault(dd, [0.0] * 13)
                    g_csum[dd][m - 1] += v
                    g_csum[dd][12] += v
            # 借：各费用科目（集团）
            for cc in DEBIT_CATS:
                if cc not in g_dsum or abs(sum(g_dsum[cc])) < 0.005:
                    continue
                row = [y, '集团汇总数', '借：' + cc] + [round(v, 2) for v in g_dsum[cc]]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                ws6.cell(r, 3).font = Font(name='Times New Roman', bold=True)
                r += 1
            # 借：列支合计（集团）
            g_d_tot = [0.0] * 13
            for cc, arr in g_dsum.items():
                for m in range(13):
                    g_d_tot[m] += arr[m]
            if any(abs(v) > 0.005 for v in g_d_tot):
                row = [y, '集团汇总数', '借：列支合计'] + [round(v, 2) for v in g_d_tot]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                for j in range(1, 17):
                    ws6.cell(r, j).fill = TFILL
                    ws6.cell(r, j).font = BFONT
                r += 1
            # 贷：各明细（集团）
            for dd in DETAILS:
                if dd not in g_csum or abs(sum(g_csum[dd])) < 0.005:
                    continue
                row = [y, '集团汇总数', '贷：' + dd] + [round(v, 2) for v in g_csum[dd]]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                ws6.cell(r, 3).font = Font(name='Times New Roman', bold=True)
                r += 1
            # 贷：计提合计（集团）
            g_c_tot = [0.0] * 13
            for dd, arr in g_csum.items():
                for m in range(13):
                    g_c_tot[m] += arr[m]
            if any(abs(v) > 0.005 for v in g_c_tot):
                row = [y, '集团汇总数', '贷：计提合计'] + [round(v, 2) for v in g_c_tot]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                for j in range(1, 17):
                    ws6.cell(r, j).fill = TFILL
                    ws6.cell(r, j).font = BFONT
                r += 1
            # 差：借−贷（集团）
            g_diff = [g_d_tot[m] - g_c_tot[m] for m in range(13)]
            if any(abs(v) > 0.005 for v in g_diff):
                row = [y, '集团汇总数', '差：借−贷'] + [round(v, 2) for v in g_diff]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                for j in range(1, 17):
                    ws6.cell(r, j).fill = WARN_FILL
                    ws6.cell(r, j).font = BFONT
                r += 1
                g_ddiff = g_d_tot[12] - g_c_tot[12]
                if abs(g_ddiff) <= max(1.0, abs(g_d_tot[12]) * 0.005):
                    g_note = '全年借=贷（配平）；分月差异系跨月结转/期初余额所致，非异常'
                elif g_ddiff > 0:
                    g_note = '借方列支 > 贷方计提：存在直接付款/代扣代垫个人部分（不经应付职工薪酬），查凭证'
                else:
                    g_note = '贷方计提 > 借方列支：计提未全额分配（保留应付职工薪酬贷方余额），查凭证'
                cell = fnum(ws6, r, 3, '→ ' + g_note)
                cell.alignment = LEFT
                ws6.merge_cells(start_row=r, start_column=3, end_row=r, end_column=16)
                r += 1
            # 两表差异对照（薪酬分配核对表 vs 多借多贷核对表，全年口径）
            recs_y = {e: records_by_year[y].get(e) for e in ents}
            g_exp_acc = sum(recs_y[e]['gl_accrual'].get(d, 0.0) for e in ents if recs_y[e] for d in DETAILS)
            g_exp_lie = sum(recs_y[e]['exp'][d][c] for e in ents if recs_y[e] for d in DETAILS for c in CANON) + \
                        sum(recs_y[e]['exp'][d]['其他列支'] for e in ents if recs_y[e] for d in DETAILS)
            pay_diff = round(g_exp_lie - g_exp_acc, 2)    # 薪酬分配核对表：列支 − 计提
            xdiff = round(g_d_tot[12] - g_c_tot[12], 2)   # 多借多贷核对表：借 − 贷
            ddiff2 = round(pay_diff - xdiff, 2)
            tol2 = max(1.0, max(abs(pay_diff), abs(xdiff)) * 0.005)
            same_dir = (pay_diff >= 0) == (xdiff >= 0)
            if abs(ddiff2) <= tol2:
                verdict = '一致'
                vcolor = '006100'
            elif same_dir:
                verdict = '基本一致（方向相同，差额系两表口径差异：多借多贷=凭证级原始借方，薪酬分配=分配后列支）'
                vcolor = 'BF8F00'
            else:
                verdict = '不一致（方向相反，需查）'
                vcolor = 'C00000'
            for lab, val in [
                ('两表差异对照：薪酬分配核对表(列支−计提,全年)', pay_diff),
                ('两表差异对照：多借多贷核对表(借−贷,全年)', xdiff),
                ('两表差异对照：两表差额', ddiff2),
                ('两表差异对照：是否一致', verdict),
            ]:
                row = [y, '集团汇总数', lab] + [None] * 12 + [val]
                for j, v in enumerate(row, 1):
                    fnum(ws6, r, j, v)
                ws6.cell(r, 3).alignment = LEFT
                ws6.cell(r, 16).font = Font(name='Times New Roman', size=10, bold=True, color=vcolor)
                r += 1
    # 表尾口径说明
    r += 1
    c = ws6.cell(r, 1, '口径：借方=计提凭证内费用科目(生产成本/制造/管理/销售/研发/其他)借方发生额按月；'
                      '贷方=计提凭证内 应付职工薪酬 明细贷方发生额按月。借贷差=无法末级配平的净额，'
                      '多借多贷凭证不拆分匹配，仅以合计对照（2026-08-02）。')
    c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=16)
    return ws6


# ============================================================================
# 职工薪酬 审定表（参考《应付职工薪酬审定表.xlsx》横向格式）与 附注披露（纵向）
# 2026-08-02 用户方法论：审定表=每主体一块，列=未审数(期初/本期增加/本期减少/期末)
#   | 审计调整(借/贷) | 重分类调整(借/贷) | 期初审计调整(借/贷) | 审定数(期初/本期增加/本期减少/期末)，
#   行=标准披露层级（短期薪酬明细/设定提存明细，非各主体自身科目）；数据取自科目余额表(TB)叶子行。
#   附注披露=每主体一列、行=审定表同套标准行、段=期初数/本期增加-贷方/本期减少-借方/期末数（审定口径）。
# ============================================================================
# 审定表块内行定义：(类型, 名称)；类型: H1/H2=头, SUM=汇总, D=明细, TOT=合计, SUB=小计, GAP=勾稽
PAY_STD_BLOCK = [
    ('H1', '应付职工薪酬'),
    ('SUM', '短期薪酬'),
    ('SUM', '离职后福利—设定提存计划'),
    ('D', '辞退福利'),
    ('D', '1年内到期的其他福利'),
    ('TOT', '合  计'),
    ('GAP', None),
    ('H2', '短期薪酬'),
    ('D', '基本工资'),
    ('D', '奖金、津贴和补贴'),
    ('D', '职工福利费'),
    ('D', '社会保险费'),
    ('D', '  其中：医疗保险费及生育保险'),
    ('D', '        工伤保险费'),
    ('D', '        补充医疗'),
    ('D', '住房公积金'),
    ('D', '工会经费'),
    ('D', '职工教育经费'),
    ('D', '劳务派遣人员薪酬'),
    ('D', '劳务费（临时人员）'),
    ('SUB', '短期薪酬小计'),
    ('H2', '设定提存计划'),
    ('D', '基本养老保险'),
    ('D', '失业保险费'),
    ('D', '企业年金缴费'),
    ('D', '其他'),
    ('SUB', '设定提存计划小计'),
]
# TB 应付职工薪酬二级科目名 → 标准披露行（关键词按优先级顺序，先命中先得）
_PAY_STD_KW = [
    # 2026-08-05 修复：键必须与 PAY_STD_BLOCK/_pay_block_vals.short_keys 完全一致
    # （带缩进）。原写 '补充医疗'（无缩进）→ std.get('        补充医疗')=None →
    # 补充医疗 400 万不计入短期薪酬/合计（FY本级 审定表缺 400 万）。
    ('        补充医疗', ('补充医疗',)),
    ('基本养老保险', ('养老',)),
    ('失业保险费', ('失业',)),
    ('  其中：医疗保险费及生育保险', ('医疗', '生育')),
    ('        工伤保险费', ('工伤',)),
    ('基本工资', ('基本工资', '工资薪金', '工资')),
    ('奖金、津贴和补贴', ('奖金', '津贴', '补贴', '绩效')),
    ('辞退福利', ('辞退', '解除')),
    ('职工福利费', ('福利',)),
    ('社会保险费', ('社会保险', '社保')),
    ('住房公积金', ('住房', '公积')),
    ('工会经费', ('工会',)),
    ('职工教育经费', ('教育',)),
    ('劳务派遣人员薪酬', ('劳务派遣', '派遣')),
    ('劳务费（临时人员）', ('劳务', '临时')),
    ('企业年金缴费', ('年金',)),
    ('1年内到期的其他福利', ('一年内', '1年内', '其他福利')),
]


def _pay_std_key(name):
    """应付职工薪酬科目名 → 标准披露行名；未命中返回 None。"""
    if not name:
        return None
    n = str(name)
    for std, kws in _PAY_STD_KW:
        for k in kws:
            if k in n:
                return std
    return None


def _collect_pay_tb(tb_full, e, y):
    """收集 (主体,年) 应付职工薪酬 TB 叶子行 → ({标准行: [期初,贷发,借发,期末]}, has_l2, 一级合计)。
    负债类：期初/期末=贷余为正、借余为负（-qc/-qm）——混合方向子目 abs 加总 ≠ 父级净额
    （AYL 2026 上海朗炫：奖金贷余 850200 + 社保借余 2766，abs 合计 852966 vs 父级 847433，
    差 5532=勾稽差异）；-qm 带符号后子目合计自动=父级。增加=贷方发生额、减少=借方发生额。
    仅取叶子行（父级=子级之和，避免双计）；无二级子目时一级行作合计基准。"""
    rows = []
    for (ee, c, n, yy), v in tb_full.items():
        if ee != e or yy != y:
            continue
        if not (str(c).startswith('2211') or '应付职工薪酬' in str(n)):
            continue
        rows.append((str(c), str(n), v))
    if not rows:
        return {}, False, [0.0, 0.0, 0.0, 0.0]
    tot1 = [0.0, 0.0, 0.0, 0.0]
    # ⚡⚡ 2026-08-25 统一符号：open=-qc、close=-qm（贷余正/借余负）——覆盖主体整体借余
    #   （AQ QS 预付）与子目混合方向（AYL 上海朗炫 贷余+借余子目）两场景，子目合计=父级。
    for c, n, v in rows:
        if c == '2211' or ('.' not in c and str(n).strip() == '应付职工薪酬'):
            tot1[0] += -float(v['qc'] or 0.0); tot1[1] += v['df']; tot1[2] += v['jf']; tot1[3] += -float(v['qm'] or 0.0)
    has_l2 = any(len(c) > 4 or '-' in n for c, n, _ in rows)   # U8 二级=6位(221101) / 泰国带点(1130.01) / 名称带'-'
    out = {}

    def _is_l1(c, n):
        # 一级：code 恰为 2211（或 ≤4 位且名称==应付职工薪酬），非二级
        return c == '2211' or ('.' not in c and len(c) <= 4 and str(n).strip() == '应付职工薪酬')

    for c, n, v in rows:
        if has_l2 and _is_l1(c, n):
            continue                       # 有二级时跳过一级（父=子和）
        if any(c != c2 and ((c2.startswith(c + '.') if '.' in c else (c2.startswith(c) and len(c2) > len(c)))) for c2, _, _ in rows):
            continue                       # 有更深子级 → 父级跳过
        k = _pay_std_key(n)
        if k is None:
            k = '其他'
        a = out.setdefault(k, [0.0, 0.0, 0.0, 0.0])
        a[0] += -float(v['qc'] or 0.0); a[1] += v['df']; a[2] += v['jf']; a[3] += -float(v['qm'] or 0.0)
    return out, has_l2, tot1


def _pay_block_vals(std, has_l2, tot1):
    """按 PAY_STD_BLOCK 计算各行的 4 段值 {行名: [期初,增加,减少,期末]}（不含 H1/H2/GAP）。"""
    short_keys = ['基本工资', '奖金、津贴和补贴', '职工福利费', '社会保险费',
                  '  其中：医疗保险费及生育保险', '        工伤保险费', '        补充医疗',
                  '住房公积金', '工会经费', '职工教育经费', '劳务派遣人员薪酬', '劳务费（临时人员）']
    dfd_keys = ['基本养老保险', '失业保险费', '企业年金缴费', '其他']
    vals = {}
    for typ, nm in PAY_STD_BLOCK:
        if typ == 'SUM':
            if nm == '短期薪酬':
                v = [0.0] * 4
                for k in short_keys:
                    a = std.get(k)
                    if a:
                        for i in range(4):
                            v[i] += a[i]
                vals[nm] = v
            else:  # 离职后福利—设定提存计划
                v = [0.0] * 4
                for k in dfd_keys:
                    a = std.get(k)
                    if a:
                        for i in range(4):
                            v[i] += a[i]
                vals[nm] = v
        elif typ == 'TOT':
            v = [0.0] * 4
            for k in ('短期薪酬', '离职后福利—设定提存计划', '辞退福利', '1年内到期的其他福利'):
                a = vals.get(k)
                if a:
                    for i in range(4):
                        v[i] += a[i]
            if not has_l2:
                # 无二级明细：合计直接取一级 TB（保证与 TB 勾稽）
                v = list(tot1)
            vals[nm] = v
        elif typ == 'SUB':
            keys = short_keys if nm == '短期薪酬小计' else dfd_keys
            v = [0.0] * 4
            for k in keys:
                a = std.get(k)
                if a:
                    for i in range(4):
                        v[i] += a[i]
            vals[nm] = v
        elif typ == 'D':
            vals[nm] = list(std.get(nm, [0.0] * 4))
    return vals


def build_payroll_audit_v2(wb, data_dir, target_year=None):
    """职工薪酬 审定表（参考《应付职工薪酬审定表.xlsx》格式）：
    表头=序号|核算主体|项目|未审数(期初/本期增加/本期减少/期末)|审计调整(借/贷)|重分类调整(借/贷)
      |期初审计调整(借/贷)|审定数(期初/本期增加/本期减少/期末)；
    每核算主体一块（2 行表头 + 27 行标准披露层级），数据取自科目余额表(TB)叶子行。
    审定数=未审数（审计调整留 0 待填）。插入工作簿第一张 sheet。"""
    from audit_common import read_tb_full, discover_entities
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    # SAP：单主体驱动时按 current_comp 过滤（防审定表/附注混入 88 家主体行）
    if _adapter is not None and _adapter.is_sap(data_dir):
        _c_py = _adapter.current_comp()
        entities = ({_c_py: discover_entities(data_dir).get(_c_py, {})} if _c_py else discover_entities(data_dir))
    else:
        entities = discover_entities(data_dir)
    if not entities:
        return
    tb_full = read_tb_full(data_dir, entities)
    years = sorted({str(y) for e in entities for y in entities[e]})
    if target_year is not None:
        years = [str(target_year)]
    ent_order = sorted(entities)   # 2026-08-03：主体按 01、02… 排序（原 discover 原始序 10,11,12,01..09）
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    TITLEFILL = PatternFill('solid', fgColor='BDD7EE')
    TOTFILL = PatternFill('solid', fgColor='FCE4D6')
    SUBFILL = PatternFill('solid', fgColor='FFF2CC')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center')
    NUMFMT = '#,##0.00'
    FONT10 = Font(name='Times New Roman', size=10)
    BFONT = Font(name='Times New Roman', size=10, bold=True)
    TFONT = Font(name='Times New Roman', size=12, bold=True)
    NFONT = Font(name='Times New Roman', size=10, italic=True, color='808080')
    NC = 17

    def _money(cell, v):
        cell.value = round(v, 2) if abs(v) >= 0.005 else 0.0
        cell.number_format = NUMFMT
        cell.alignment = RGT
        cell.border = BORDER
        cell.font = FONT10

    def _txt(r, c, v, bold=False, fill=None, align=None, font=None):
        cell = ws.cell(r, c, v)
        cell.font = font or (BFONT if bold else FONT10)
        cell.alignment = align or LFT
        cell.border = BORDER
        if fill:
            cell.fill = fill

    if '职工薪酬 审定表' in wb.sheetnames:
        del wb['职工薪酬 审定表']
    ws = wb.create_sheet('职工薪酬 审定表', 0)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NC)
    t = ws.cell(1, 1, '应付职工薪酬审定表')
    t.font = TFONT; t.fill = TITLEFILL; t.alignment = CTR
    ws.row_dimensions[1].height = 22
    widths = [6, 16, 24, 14, 14, 14, 14, 11, 11, 11, 11, 11, 11, 14, 14, 14, 14]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w

    def _write_hdr(r):
        h1 = ['序号', '核算主体', '项目', '未审数', '', '', '', '审计调整', '', '重分类调整', '',
              '期初审计调整', '', '审定数', '', '', '']
        for j, h in enumerate(h1, 1):
            _txt(r, j, h or None, bold=True, fill=HFILL, align=CTR)
        for a, b in [(4, 7), (8, 9), (10, 11), (12, 13), (14, 17)]:
            ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
        r2 = r + 1
        h2 = ['', '', '', '期初数', '本期增加', '本期减少', '期末数',
              '借方', '贷方', '借方', '贷方', '借方', '贷方', '期初数', '本期增加', '本期减少', '期末数']
        for j, h in enumerate(h2, 1):
            _txt(r2, j, h or None, bold=True, fill=HFILL, align=CTR)
        ws.row_dimensions[r].height = 18
        ws.row_dimensions[r2].height = 18

    def _write_val_row(r, nm, vals, fill=None, bold=False):
        _txt(r, 3, nm, bold=bold, fill=fill, align=LFT)
        for i in range(4):
            _money(ws.cell(r, 4 + i), vals[i])          # 未审数 期初/增加/减少/期末
        for i in range(8):
            _money(ws.cell(r, 8 + i), 0.0)              # 审计调整/重分类/期初审计调整 借/贷
        for i in range(4):
            _money(ws.cell(r, 14 + i), vals[i])         # 审定数 = 未审（调整 0）
        if fill:
            for j in range(3, NC + 1):
                ws.cell(r, j).fill = fill

    r = 3
    seq = 0
    for e in ent_order:
        for y in years:
            std, has_l2, tot1 = _collect_pay_tb(tb_full, e, y)
            has_any = any(abs(x) > 1e-6 for a in std.values() for x in a) or any(abs(x) > 1e-6 for x in tot1)
            if not has_any:
                continue
            _write_hdr(r)
            r += 2
            blk = r
            vals = _pay_block_vals(std, has_l2, tot1)
            for typ, nm in PAY_STD_BLOCK:
                if typ == 'H1':
                    _txt(r, 3, nm, bold=True)
                elif typ == 'GAP':
                    _txt(r, 3, '勾稽：期初+本期增加−本期减少 应=期末；合计 应=TB一级', font=NFONT)
                    _money(ws.cell(r, 4), 0.0)
                    _money(ws.cell(r, 7), 0.0)
                    d = vals.get('合  计', [0.0] * 4)
                    roll = d[0] + d[1] - d[2] - d[3]
                    lvl = (tot1[0] + tot1[1] - tot1[2] - tot1[3]) - (d[0] + d[1] - d[2] - d[3])
                    # ⚡ 2026-08-11 勾稽不平标红（超 0.01 容差）：roll=期初+增−减−末、
                    #   lvl=TB一级−合计；不平 → 红色加粗 + 浅红底（用户要求数字标红）
                    _roll_ok = abs(roll) <= 0.01
                    _lvl_ok = abs(lvl) <= 0.01
                    _g_font = lambda ok: (NFONT if ok else RED_FONT)
                    c1 = ws.cell(r, 5, round(roll, 2)); c1.number_format = NUMFMT; c1.font = _g_font(_roll_ok)
                    c2 = ws.cell(r, 9, round(lvl, 2)); c2.number_format = NUMFMT; c2.font = _g_font(_lvl_ok)
                    if not _roll_ok or not _lvl_ok:
                        _F_RED_BG = PatternFill('solid', fgColor='FCE4EC')
                        for _c in (5, 9):
                            ws.cell(r, _c).fill = _F_RED_BG
                elif typ == 'H2':
                    _txt(r, 3, '    ' + nm, bold=True)
                elif typ == 'SUB':
                    _write_val_row(r, '    ' + nm, vals[nm], fill=SUBFILL, bold=True)
                elif typ == 'TOT':
                    _write_val_row(r, nm, vals[nm], fill=TOTFILL, bold=True)
                else:
                    _write_val_row(r, nm, vals[nm])
                r += 1
            # 序号/核算主体 列块内合并
            seq += 1
            _txt(blk, 1, seq, align=CTR)
            _txt(blk, 2, e, align=CTR)
            ws.merge_cells(start_row=blk, start_column=1, end_row=r - 1, end_column=1)
            ws.merge_cells(start_row=blk, start_column=2, end_row=r - 1, end_column=2)

    # ===== 集团汇总数（审定表，2026-08-02：内容同上面各主体标准披露行，值=全部主体之和） =====
    # ⚡ 2026-08-10 单体裁剪（铁律81）：SAP 逐主体底稿（1 主体）不输出全集团合计/合并抵消
    if len(ent_order) <= 1:
        ws.freeze_panes = 'D3'
        return ws
    for y in years:
        g_vals = {}
        g_tot1 = [0.0] * 4
        has_any = False
        for e in ent_order:
            std, has_l2, tot1 = _collect_pay_tb(tb_full, e, y)
            if any(abs(x) > 1e-6 for a in std.values() for x in a) or any(abs(x) > 1e-6 for x in tot1):
                has_any = True
            for i in range(4):
                g_tot1[i] += tot1[i]
            bv = _pay_block_vals(std, has_l2, tot1)
            for nm, arr in bv.items():
                gv = g_vals.setdefault(nm, [0.0] * 4)
                for i in range(4):
                    gv[i] += arr[i]
        if not has_any:
            continue
        _write_hdr(r)
        r += 2
        blk = r
        for typ, nm in PAY_STD_BLOCK:
            if typ == 'H1':
                _txt(r, 3, nm, bold=True)
            elif typ == 'GAP':
                _txt(r, 3, '勾稽：期初+本期增加−本期减少 应=期末；合计 应=TB一级（集团）', font=NFONT)
                _money(ws.cell(r, 4), 0.0)
                _money(ws.cell(r, 7), 0.0)
                d = g_vals.get('合  计', [0.0] * 4)
                roll = d[0] + d[1] - d[2] - d[3]
                lvl = (g_tot1[0] + g_tot1[1] - g_tot1[2] - g_tot1[3]) - (d[0] + d[1] - d[2] - d[3])
                c1 = ws.cell(r, 5, round(roll, 2)); c1.number_format = NUMFMT; c1.font = NFONT
                c2 = ws.cell(r, 9, round(lvl, 2)); c2.number_format = NUMFMT; c2.font = NFONT
            elif typ == 'H2':
                _txt(r, 3, '    ' + nm, bold=True)
            elif typ == 'SUB':
                _write_val_row(r, '    ' + nm, g_vals.get(nm, [0.0] * 4), fill=SUBFILL, bold=True)
            elif typ == 'TOT':
                _write_val_row(r, nm, g_vals.get(nm, [0.0] * 4), fill=TOTFILL, bold=True)
            else:
                _write_val_row(r, nm, g_vals.get(nm, [0.0] * 4))
            r += 1
        # 序号/核算主体 列块内合并（集团汇总数）
        seq += 1
        _txt(blk, 1, seq, align=CTR)
        _txt(blk, 2, '集团汇总数', align=CTR)
        ws.merge_cells(start_row=blk, start_column=1, end_row=r - 1, end_column=1)
        ws.merge_cells(start_row=blk, start_column=2, end_row=r - 1, end_column=2)
    ws.freeze_panes = 'D3'
    return ws


def _pay_std_vals(std):
    """std(科目→[期初,增,减,末]) → {PAY_STD_BLOCK 显示行名: [4]}。
    ⚡ 2026-08-11：SUM(短期薪酬/离职后福利)/SUB(小计)/TOT(合计) = 组内 D 行之和。
    注意：PAY_STD_BLOCK 的 TOT 行在明细 D 之前（结构如此）→ TOT 须循环后 Σ D 行计算；
    SUM 行按 SUM↔SUB 映射取值（短期薪酬←短期薪酬小计）。"""
    out = {}
    grp_name = None
    grp_val = None
    for typ, nm in PAY_STD_BLOCK:
        if nm is None:
            continue
        nm_s = nm.strip()
        if typ == 'H2':
            grp_name = nm_s
            grp_val = [0.0, 0.0, 0.0, 0.0]
        elif typ == 'D':
            v = std.get(nm, [0.0, 0.0, 0.0, 0.0])
            out[nm] = list(v)
            if grp_val is not None:
                for i in range(4):
                    grp_val[i] += v[i]
        elif typ == 'SUB':
            out[nm_s] = list(grp_val) if grp_val else [0.0, 0.0, 0.0, 0.0]
        elif typ == 'SUM':
            out[nm_s] = [0.0, 0.0, 0.0, 0.0]      # 占位（循环后按 SUB 映射填）
        # TOT 不在此填（行位在明细前，循环后统一计算）
    # 合计 = Σ 全部 D 行（std 直接科目）
    _tot = [0.0, 0.0, 0.0, 0.0]
    for _v in std.values():
        for i in range(4):
            _tot[i] += _v[i]
    out['合  计'] = _tot
    # SUM 一级行 = 对应 SUB 二级小计
    for sum_nm, sub_nm in (('短期薪酬', '短期薪酬小计'), ('离职后福利—设定提存计划', '设定提存计划小计')):
        if sum_nm in out and sub_nm in out:
            out[sum_nm] = list(out[sub_nm])
    return out


def build_payroll_extra_sheets(wb, data_dir, target_year=None):
    """⚡ 2026-08-11 用户需求（职工薪酬审定表增强）：
    ①『职工薪酬集团审定表』——每核算单位一行（仿应收账款审定表：单位|期初|本期增加|本期减少|期末|审定数，尾部合计）；
    ②『职工薪酬横展审定表』——行=二级/三级科目、列=各核算主体+合计（横向对比哪家单位工资多/少）。
    数据=TB 账面数（期初/贷发增加/借发减少/期末，未审）。集团模式（多主体）生成；单主体跳过（铁律81）。"""
    from audit_common import read_tb_full, discover_entities
    from openpyxl.utils import get_column_letter
    if _adapter is not None and _adapter.is_sap(data_dir):
        _c_py = _adapter.current_comp()
        entities = ({_c_py: discover_entities(data_dir).get(_c_py, {})} if _c_py else discover_entities(data_dir))
    else:
        entities = discover_entities(data_dir)
    ents = sorted(entities)
    if len(ents) <= 1:
        return
    tb_full = read_tb_full(data_dir, entities)
    years = sorted({str(y) for e in entities for y in entities[e]})
    if target_year is not None:
        years = [str(target_year)]
    y = years[0]
    # 每主体 TB 数据（账面数）
    ent_std = {}
    for e in ents:
        std, has_l2, tot1 = _collect_pay_tb(tb_full, e, y)
        if any(abs(x) > 1e-6 for a in std.values() for x in a) or any(abs(x) > 1e-6 for x in tot1):
            ent_std[e] = (std, has_l2, tot1, _pay_std_vals(std))
    if not ent_std:
        return
    from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_BOLD,
                             SHELL_NUM, SHELL_LEFT, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                             SHELL_TOT_FILL)
    # ===== ① 集团审定表：每单位一行 =====
    # ⚡ 2026-08-11 v2：加『审计调整数』列（0 待填）+『审定数』（=期末+调整，调整 0 时=期末），
    #   与最终报表核对（调整数在审计调整分录 sheet 过入后可刷新）。
    ws = wb.create_sheet(f'职工薪酬集团审定表_{y}')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    c = ws.cell(1, 1, f'应付职工薪酬集团审定表（{y} 年度）· 每核算单位一行 · 账面数（未审）+ 审计调整/审定数')
    c.font = SHELL_TITLE_FONT
    hdr = ['序号', '核算主体', '期初数', '本期增加', '本期减少', '期末数', '审计调整数', '审定数']
    for j, h in enumerate(hdr, 1):
        cc = ws.cell(2, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    r = 3
    g = [0.0] * 4
    for i, e in enumerate(ents, 1):
        if e not in ent_std:
            continue
        std, has_l2, tot1, _std_vals = ent_std[e]
        vals = tot1[:4]
        # 审定数 = 期末 + 审计调整数（调整列 0 待填）
        row_vals = [i, e, round(vals[0], 2), round(vals[1], 2), round(vals[2], 2),
                    round(vals[3], 2), 0.0, round(vals[3], 2)]
        for j, v in enumerate(row_vals, 1):
            cc = ws.cell(r, j, v)
            cc.border = SHELL_BORDER
            if j >= 4:
                cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
            elif j == 1:
                cc.alignment = SHELL_CEN
        g = [g[k] + vals[k] for k in range(4)]
        r += 1
    # 合计行
    row_vals = ['合计', None, round(g[0], 2), round(g[1], 2), round(g[2], 2),
                round(g[3], 2), 0.0, round(g[3], 2)]
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
    for i, w in enumerate([6, 16, 16, 16, 16, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'
    # ===== ② 横展审定表：科目×主体 =====
    ws2 = wb.create_sheet(f'职工薪酬横展审定表_{y}')
    nc = 1 + len(ents) + 1
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws2.cell(1, 1, f'应付职工薪酬横展审定表（{y} 年度）· 行=科目、列=各核算单位+合计（横向对比各单位工资规模）')
    c.font = SHELL_TITLE_FONT
    hdr2 = ['科目'] + ents + ['合计']
    for j, h in enumerate(hdr2, 1):
        cc = ws2.cell(2, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    # 行=PAY_STD_BLOCK 的显示行（H1 分组/H2 二级/SUB 三级 + 合计）；取期末数
    rows2 = []
    for typ, nm in PAY_STD_BLOCK:
        if typ in ('GAP',):
            continue
        if typ == 'H1':
            rows2.append((nm, 'H1'))
        elif typ == 'H2':
            rows2.append(('    ' + nm, 'H2'))
        elif typ == 'SUB':
            rows2.append(('    ' + nm, 'SUB'))
        elif typ == 'TOT':
            rows2.append((nm, 'TOT'))
        else:
            rows2.append((nm, 'ROW'))
    r = 3
    for nm, typ in rows2:
        _nm = nm.strip()
        if typ == 'H1':
            vals_by_e = {e: None for e in ents}
        elif _nm == '合  计':
            vals_by_e = {e: (ent_std[e][2][3] if e in ent_std else None) for e in ents}
        else:
            vals_by_e = {e: (ent_std[e][3].get(_nm, [0.0] * 4)[3] if e in ent_std else None)
                         for e in ents}
        tot = sum(v for v in vals_by_e.values() if isinstance(v, (int, float)))
        ws2.cell(r, 1, nm)
        for j, e in enumerate(ents, 2):
            v = vals_by_e[e]
            cc = ws2.cell(r, j, round(v, 2) if isinstance(v, (int, float)) and abs(v) >= 0.005 else None)
            cc.border = SHELL_BORDER
            if j >= 2:
                cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
        cc = ws2.cell(r, nc, round(tot, 2) if abs(tot) >= 0.005 else None)
        cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
        cc.font = SHELL_BOLD
        for j in range(1, nc + 1):
            ws2.cell(r, j).border = SHELL_BORDER
            if typ in ('H1', 'TOT'):
                ws2.cell(r, j).font = SHELL_BOLD
        r += 1
    ws2.column_dimensions['A'].width = 30
    for j in range(2, nc + 1):
        ws2.column_dimensions[get_column_letter(j)].width = 13
    ws2.freeze_panes = 'B3'
    return ws, ws2


def build_payroll_footnote(wb, data_dir, target_year=None):
    """职工薪酬附注披露（纵向，参考《职工薪酬附注披露.xlsx》）：
    每核算主体一列、行=审定表同套标准披露行（按审定表内容填列，非各主体自身科目格式）；
    段=期初数/本期增加-贷方/本期减少-借方/期末数（审定口径=未审，调整 0）；
    每段末 合计（该列之和）/复核 行。追加为工作簿末张 sheet。"""
    from audit_common import read_tb_full, discover_entities
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    # SAP：单主体驱动时按 current_comp 过滤（防审定表/附注混入 88 家主体行）
    if _adapter is not None and _adapter.is_sap(data_dir):
        _c_py = _adapter.current_comp()
        entities = ({_c_py: discover_entities(data_dir).get(_c_py, {})} if _c_py else discover_entities(data_dir))
    else:
        entities = discover_entities(data_dir)
    if not entities:
        return
    tb_full = read_tb_full(data_dir, entities)
    years = sorted({str(y) for e in entities for y in entities[e]})
    if target_year is not None:
        years = [str(target_year)]
    ent_order = sorted(entities)   # 2026-08-03：主体按 01、02… 排序（原 discover 原始序 10,11,12,01..09）
    # 附注披露数据行 = 审定表数据行（去 H1/H2/GAP）
    FOOT_LINES = [(typ, nm) for typ, nm in PAY_STD_BLOCK if typ in ('SUM', 'D', 'TOT', 'SUB')]
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    TITLEFILL = PatternFill('solid', fgColor='BDD7EE')
    TOTFILL = PatternFill('solid', fgColor='FCE4D6')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center')
    NUMFMT = '#,##0.00'
    FONT10 = Font(name='Times New Roman', size=10)
    BFONT = Font(name='Times New Roman', size=10, bold=True)
    TFONT = Font(name='Times New Roman', size=12, bold=True)
    n_ent = len(ent_order)
    NC = 2 + n_ent

    if '附注汇总' in wb.sheetnames:
        del wb['附注汇总']
    ws = wb.create_sheet('附注汇总')
    # 2026-08-02：附注汇总不设合并单元格（便于后期添加/筛选）；标题仅 A1 加粗
    t = ws.cell(1, 1, '应付职工薪酬附注披露（审定数口径；每主体一列；行=审定表标准披露项；段=期初/本期增加-贷方/本期减少-借方/期末）')
    t.font = TFONT; t.fill = TITLEFILL; t.alignment = CTR
    ws.row_dimensions[1].height = 22
    # 预计算各主体（当年）标准行数据
    year_now = years[0] if years else None
    _pay_cols = {}
    for e in ent_order:
        _pay_cols[e] = _collect_pay_tb(tb_full, e, year_now) if year_now else ({}, False, [0.0] * 4)
    # 主体表头
    hdr_r = 2
    ws.cell(hdr_r, 1, '项目（审定数）').font = BFONT
    ws.cell(hdr_r, 1).fill = HFILL; ws.cell(hdr_r, 1).alignment = CTR; ws.cell(hdr_r, 1).border = BORDER
    for i, ent in enumerate(ent_order, 2):
        c = ws.cell(hdr_r, i, ent)
        c.font = BFONT; c.fill = HFILL; c.alignment = CTR; c.border = BORDER
    ws.column_dimensions['A'].width = 30
    for j in range(2, NC + 1):
        ws.column_dimensions[get_column_letter(j)].width = 16
    r = 3
    seg_gc = {}   # 段 -> 集团合计数行号（2026-08-03 滚动勾稽行引用）
    for seg_label, idx in [('期初数', 0), ('本期增加-贷方', 1), ('本期减少-借方', 2), ('期末数', 3)]:
        # 段头：A 列加粗提示（不合并跨行）
        c = ws.cell(r, 1, seg_label)
        c.font = BFONT; c.fill = TOTFILL; c.alignment = CTR
        r += 1
        seg_start = r
        # 2026-08-03 公式化：SUM/TOT/SUB 行写 =SUM 公式（不再写死数值）。
        # 两遍法：①先渲染 D 明细行（数据）并记录行号；②再写 SUM/TOT/SUB/集团合计公式。
        d_rows = {}        # D 行名 → 行号
        sum_rows = {}      # SUM 行名 → 行号
        sub_rows = {}      # SUB 行名 → 行号
        tot_rows = []      # TOT 行号
        order = []         # 行渲染顺序记录：(typ, nm, r)
        # —— 第一遍：渲染全部行（SUM/TOT/SUB 先写 0 占位，D 写数据），记录行号 ——
        for typ, nm in FOOT_LINES:
            label = nm
            if typ == 'SUB':
                label = '    小  计'
            elif typ == 'TOT':
                label = '合  计'
            c = ws.cell(r, 1, label)
            c.font = BFONT if typ in ('TOT', 'SUB') else FONT10
            c.alignment = LFT; c.border = BORDER
            if typ == 'TOT':
                c.fill = TOTFILL
            elif typ == 'SUB':
                c.fill = PatternFill('solid', fgColor='FFF2CC')
            if typ == 'D':
                d_rows[nm] = r
            elif typ == 'SUM':
                sum_rows[nm] = r
            elif typ == 'SUB':
                sub_rows[nm] = r
            elif typ == 'TOT':
                tot_rows.append(r)
            order.append((typ, nm, r))
            for i, e in enumerate(ent_order, 2):
                if typ == 'D':
                    std, has_l2, tot1 = _pay_cols.get(e, (None, None, None))
                    v = 0.0
                    if std is not None:
                        bv = _pay_block_vals(std, has_l2, tot1)
                        v = bv.get(nm, [0.0] * 4)[idx]
                    cell = ws.cell(r, i, round(v, 2) if abs(v) >= 0.005 else 0.0)
                else:
                    cell = ws.cell(r, i, 0.0)   # SUM/TOT/SUB 占位，第二遍回填公式
                cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER
                cell.font = FONT10
                if typ == 'TOT':
                    cell.fill = TOTFILL
                elif typ == 'SUB':
                    cell.fill = PatternFill('solid', fgColor='FFF2CC')
            r += 1
        # —— 第二遍：计算 SUM/TOT/SUB 的数据源行号并【写数值】（2026-08-06 协议：附注矩阵
        # 合计/小计/集团行写数值而非公式——openpyxl 保存不写公式缓存，data_only 读 None →
        # 收回读不回（ga 职工薪酬附注曾 363 万 vs 正确 8.45 亿）；数值可被读取器直接读到）——
        # 归属规则（PAY_STD_BLOCK 语义）：
        #   SUB 小计 = 其上方、到上一个 SUB/H2 之间的 D 行；
        #   SUM 行（短期薪酬/离职后福利）与其后同名 SUB 配对（短期薪酬↔短期薪酬小计、
        #     离职后福利↔设定提存计划小计），SUM = 该配对小计行；
        #   TOT 合 计 = 短期薪酬SUM + 离职后福利SUM + 辞退福利D + 1年内到期D。
        seq = order

        def _colval(rr, i):
            """取某行某列已写入的数值（D 行已写；公式行算到时有值）。"""
            v = ws.cell(rr, i).value
            return float(v) if isinstance(v, (int, float)) else 0.0

        # 1) 先算各 SUB 的数据源 D 行（反向收集）
        sub_src = {}
        for typ, nm, rr in seq:
            if typ != 'SUB':
                continue
            src = []
            for t2, n2, r2 in seq:
                if r2 >= rr:
                    continue
                if t2 == 'D':
                    src.append(r2)
                elif t2 in ('SUM', 'TOT', 'SUB'):
                    src = []
            sub_src[nm] = src
            for i in range(2, NC + 1):
                if src:
                    ws.cell(rr, i, round(sum(_colval(x, i) for x in src), 2))
        # 2) SUM 行 = 配对小计行（短期薪酬↔短期薪酬小计；离职后福利↔设定提存计划小计）
        sum_pair = {
            '短期薪酬': '短期薪酬小计',
            '离职后福利—设定提存计划': '设定提存计划小计',
        }
        for typ, nm, rr in seq:
            if typ != 'SUM':
                continue
            sub_nm = sum_pair.get(nm)
            sub_r = sub_rows.get(sub_nm)
            if sub_r:
                for i in range(2, NC + 1):
                    ws.cell(rr, i, round(_colval(sub_r, i), 2))
        # 3) TOT 合 计 = 两个 SUM + 辞退福利 + 1年内到期（独立 D 行）
        tot_terms = []   # 行号列表
        for typ, nm, rr in seq:
            if typ == 'SUM':
                tot_terms.append(rr)
            elif typ == 'D' and nm in ('辞退福利', '1年内到期的其他福利'):
                tot_terms.append(rr)
        for rr in tot_rows:
            if tot_terms:
                for i in range(2, NC + 1):
                    ws.cell(rr, i, round(sum(_colval(x, i) for x in tot_terms), 2))
        # 集团合计数（= TOT 行数值；FOOT_LINES 含 SUM/TOT/SUB 不可逐行累加，防双计）
        # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出合并 4 行
        rc = r
        seg_gc[seg_label] = rc
        if len(ent_order) > 1:
            ws.cell(r, 1, '集团合计数').font = BFONT; ws.cell(r, 1).fill = TOTFILL
            ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i, e in enumerate(ent_order, 2):
                # 集团合计数 = 本段最后一个 TOT 行（数值）
                if tot_rows:
                    cell = ws.cell(r, i, round(_colval(tot_rows[-1], i), 2))
                else:
                    cell = ws.cell(r, i, 0.0)
                cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER
                cell.fill = TOTFILL; cell.font = BFONT
            r += 1
            # 合并抵消借方（待填 0）
            ws.cell(r, 1, '合并抵消借方').font = FONT10; ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i in range(2, NC + 1):
                cell = ws.cell(r, i, 0.0)
                cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER
                cell.font = FONT10
            r += 1
            # 合并抵消贷方（待填 0）
            ws.cell(r, 1, '合并抵消贷方').font = FONT10; ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i in range(2, NC + 1):
                cell = ws.cell(r, i, 0.0)
                cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER
                cell.font = FONT10
            r += 1
            # 合并报表数（= 集团合计数 + 抵消借方 − 抵消贷方，数值）
            ws.cell(r, 1, '合并报表数').font = BFONT; ws.cell(r, 1).fill = TOTFILL
            ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i in range(2, NC + 1):
                cell = ws.cell(r, i, round(_colval(rc, i) + 0.0 - 0.0, 2))
                cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER
                cell.fill = TOTFILL; cell.font = BFONT
            r += 1
        # —— 2026-08-03 用户方法论：附注须考虑 审计调整数/审定数（披露增减变动科目区分借/贷调整）——
        # 审定数 = 集团合计数(rc) + 调整借(rc+4) − 调整贷(rc+5)；调整 0 待填。
        ws.cell(r, 1, '审计调整-借方').font = FONT10; ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
        for i in range(2, NC + 1):
            cell = ws.cell(r, i, 0.0)
            cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER; cell.font = FONT10
        r += 1
        ws.cell(r, 1, '审计调整-贷方').font = FONT10; ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
        for i in range(2, NC + 1):
            cell = ws.cell(r, i, 0.0)
            cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER; cell.font = FONT10
        r += 1
        ws.cell(r, 1, '审定数').font = BFONT; ws.cell(r, 1).fill = TOTFILL
        ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
        for i in range(2, NC + 1):
            # 审定数 = 集团合计数(rc) + 调整借(rc+4) − 调整贷(rc+5)，调整 0 → 数值
            cell = ws.cell(r, i, round(_colval(rc, i) + 0.0 - 0.0, 2))
            cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER
            cell.fill = TOTFILL; cell.font = BFONT
        r += 1
    # 滚动勾稽行（2026-08-03 用户方法论：期初+本期增加-本期减少-期末=0；差异>0.01 标红；
    # 2026-08-06 写数值——公式 data_only 读 None，收回读不回）
    if len(seg_gc) == 4:
        r0 = seg_gc['期初数']; r1 = seg_gc['本期增加-贷方']
        r2 = seg_gc['本期减少-借方']; r3 = seg_gc['期末数']
        c = ws.cell(r, 1, '滚动勾稽（期初+本期增加-本期减少-期末）')
        c.font = BFONT; c.fill = TOTFILL; c.alignment = LFT; c.border = BORDER
        for i, ent in enumerate(ent_order, 2):
            cell = ws.cell(r, i, round(_colval(r0, i) + _colval(r1, i) - _colval(r2, i) - _colval(r3, i), 2))
            cell.number_format = NUMFMT; cell.alignment = RGT; cell.border = BORDER
            cell.font = Font(name='Times New Roman', size=10, bold=True)
        r += 1
    ws.freeze_panes = 'B3'
    return ws


def build_merged_workbook(entities, records_by_year, years, out_path, data_dir=None):
    """职工薪酬直接按年构建（2026-08-01 改造：消除合并稿中间态，替代 emit_per_year 分拆+按年过滤/重建）。
    每年独立 wb：框架(_build_merged_frame) + _fill_merged([y] 只填当年) + 审定表(target_year=y)。"""
    ents = list(entities.keys())
    from audit_common import add_audit_summary_sheets as _rebuild_pay_audit
    from audit_common import discover_entities as _pe_de, read_tb_full as _pe_tb
    _pe_ents = _pe_de(data_dir)
    _pe_tb_full = _pe_tb(data_dir, _pe_ents)
    base_dir = os.path.dirname(out_path) if out_path else data_dir
    issues = []
    for y in sorted(years):
        ys = str(y)
        wb = openpyxl.Workbook()
        _build_merged_frame(wb)
        if 'Sheet' in wb.sheetnames:
            del wb['Sheet']
        issues += _fill_merged(wb, [y], ents, records_by_year)
        # 审定表（按年，参考《应付职工薪酬审定表.xlsx》格式）+ 附注披露（纵向，按审定表内容填列）
        for _s in list(wb.sheetnames):
            if '审定表' in _s or '附注汇总' in _s:
                del wb[_s]
        build_payroll_audit_v2(wb, data_dir, target_year=ys)
        # ⚡ 2026-08-11 用户需求：集团审定表（每单位一行）+ 横展审定表（科目×主体）——集团模式生成
        try:
            build_payroll_extra_sheets(wb, data_dir, target_year=ys)
        except Exception as _ex:
            print(f'  ⚠️ 职工薪酬集团/横展审定表生成失败：{_ex}')
        build_payroll_footnote(wb, data_dir, target_year=ys)
        # ⚡ 2026-08-11 用户需求：多期对比表并入职工薪酬底稿（仅最新年 wb，避免重复）
        if str(ys) == str(years[-1]):
            try:
                _render_payroll_multi_period_sheet(wb, data_dir)
            except Exception as _ex:
                print(f'  ⚠️ 多期对比表并入底稿失败：{_ex}')
        # 移审定表至首位（附注披露为末张）
        _sh = wb._sheets
        _ix = next((i for i, s in enumerate(_sh) if '审定表' in s.title), None)
        if _ix is not None and _ix > 0:
            _sh.insert(0, _sh.pop(_ix))
        finalize_workbook(wb)
        # ⚡ 2026-08-28 #875：对方科目核对（模块内集成）
        try:
            from counterparty_recon import inject_into_wb_auto
            inject_into_wb_auto(wb, base_dir, ys, '职工薪酬', ents_set=set(entities))
        except Exception as _ex:
            print(f'  ⚠️ 职工薪酬对方科目核对注入失败：{_ex}')
        out = os.path.join(base_dir, f'职工薪酬审计底稿_{ys}_生成.xlsx')
        from audit_common import validate_workbook
        validate_workbook(wb, '职工薪酬底稿', raise_on_error=False)
        wb.save(out)
        wb.close()
        print(f'  ✓ 职工薪酬 {ys} 已生成')
    return issues


def build_payroll_combined(data_dir, out_dir=None, period_mode='Y'):
    entities = _discover_payroll_entities(data_dir)
    if not entities:
        print(f'❌ 目录内未找到含「科目余额表」的账套文件：{data_dir}')
        return None
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    all_years = set()
    for b in entities.values():
        all_years.update(b.keys())
    years = sorted(all_years)
    records_by_year = {}
    n_by_year = {}
    for y in years:
        records = {}
        for e, b in entities.items():
            yr = b.get(y)
            if not yr or not yr.get('km'):
                continue
            try:
                gl = read_gl_payroll(yr.get('gl'))
                tb = read_tb_payroll(yr.get('km'))
            except Exception as ex:
                # 单主体账套被 Excel 占用/文件损坏时，跳过该主体本年，不中断整体生成
                print(f'⚠️ 跳过 [{e}][{y}]：读取账套失败（可能文件被 Excel 占用）：{ex}')
                continue
            records[e] = build_entity_record(gl, tb)
        if not records:
            continue
        records_by_year[y] = records
        n_by_year[y] = len(records)
    if not records_by_year:
        print('❌ 未成功构建任何年度的实体记录，无法生成合并底稿')
        return None
    out_path = os.path.join(out_dir, '职工薪酬审计底稿_生成.xlsx')
    issues = build_merged_workbook(entities, records_by_year, years, out_path, data_dir)
    produced = [(years, out_path, n_by_year, issues)]
    # ⚡ 2026-08-11 用户需求：职工薪酬多期对比表（独立跨年文件，多期账套生成）
    try:
        if len(years) >= 2:
            build_payroll_multi_period_table(data_dir, out_dir)
    except Exception as _ex:
        print(f'  ⚠️ 多期对比表生成失败：{_ex}')
    return produced


def main(argv=None):
    raw = list(sys.argv[1:] if argv is None else argv)
    # 拖入文件夹快捷方式：若首参为文件夹路径（不以 '-' 开头），自动视为 --input
    if raw and not raw[0].startswith('-') and os.path.isdir(raw[0]):
        raw = ['--input', raw[0]] + raw[1:]
    p = argparse.ArgumentParser(description='职工薪酬明细表生成程序（多账套×多年份，参照往来款小程序）')
    p.add_argument('--input', '-i', default=None, help='含科目余额表/综合查询明细表的文件夹')
    p.add_argument('--output', '-o', default=None, help='输出目录（默认与输入同目录）')
    p.add_argument('--period', choices=['Y', 'M', 'Q'], default='Y', help='期间粒度（当前仅 Y 年维度输出）')
    args = p.parse_args(raw)

    inp = args.input
    if not inp:
        try:
            inp = input('请拖入或输入 账套导出文件夹 路径：').strip().strip('"').strip("'")
        except EOFError:
            inp = ''
    if not inp or not os.path.isdir(inp):
        print(f'❌ 路径不存在或不是文件夹：{inp}')
        return 1
    produced = build_payroll_combined(inp, args.output)
    if not produced:
        return 1
    years, out_path, n_by_year, issues = produced[0]
    errs = [i for i in issues if i['level'] == 'ERROR']
    warns = [i for i in issues if i['level'] == 'WARN']
    ndesc = '，'.join('%s:%s主体' % (y, n_by_year.get(y, '-')) for y in years)
    print(f'✅ 已生成（合并版·7张sheet）：{out_path}')
    print(f'   年度 {ndesc} | sheets=职工薪酬审定表/附注汇总/职工薪酬明细表/分月计提明细表/薪酬分配核对表/多借多贷核对表/差异凭证清单 | ERROR {len(errs)} | WARN {len(warns)}')
    from audit_common import finalize_after_build
    finalize_after_build(inp)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）
    return 0


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
                                       f'payroll_crash_{ts}.log'), 'w', encoding='utf-8') as _f:
                    _f.write('未捕获异常：\n' + traceback.format_exc())
                print(f'已记录崩溃日志：payroll_crash_{ts}.log')
        except Exception:
            pass
        rc = 1
    try:
        input('\n按回车退出…')
    except EOFError:
        pass
    sys.exit(rc)

def _render_payroll_multi_period_sheet(wb, data_dir):
    """多期对比表渲染到已建 wb（2026-08-11 用户：并入职工薪酬底稿）。
    布局：行=【科目块 × 各核算主体】、列=期间×(期初/本期增加/本期减少/期末) 动态 2-3 期。
    返回 ws（≥2 期且有多主体数据时）或 None（单期/无数据）。"""
    from audit_common import read_tb_full, discover_entities
    from openpyxl.utils import get_column_letter
    if _adapter is not None and _adapter.is_sap(data_dir):
        _c_py = _adapter.current_comp()
        entities = ({_c_py: discover_entities(data_dir).get(_c_py, {})} if _c_py else discover_entities(data_dir))
    else:
        entities = discover_entities(data_dir)
    if not entities:
        return None
    tb_full = read_tb_full(data_dir, entities)
    years = sorted({str(y) for e in entities for y in entities[e]})
    if len(years) < 2:
        return None
    ents = sorted(entities)
    ent_sv = {}
    for e in ents:
        for y in years:
            std, has_l2, tot1 = _collect_pay_tb(tb_full, e, y)
            ent_sv[(e, str(y))] = _pay_std_vals(std)
    blocks = []
    for typ, nm in PAY_STD_BLOCK:
        if nm is None:
            continue
        if typ in ('SUM', 'D', 'TOT'):
            blocks.append((typ, nm.strip()))
    if not blocks:
        return None
    from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_BOLD,
                             SHELL_NUM, SHELL_LEFT, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                             SHELL_TOT_FILL)
    ws = wb.create_sheet('职工薪酬多期对比表')
    nc = 2 + len(years) * 4
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, f'职工薪酬多期对比表（{"、".join(years)} 年度）· 行=科目×各核算主体、列=期间×(期初/本期增加/本期减少/期末) · 账面数（未审）')
    c.font = SHELL_TITLE_FONT
    ws.cell(2, 1, '科目').font = SHELL_HFONT; ws.cell(2, 1).fill = SHELL_HFILL
    ws.cell(2, 1).border = SHELL_BORDER; ws.cell(2, 1).alignment = SHELL_CEN
    ws.cell(2, 2, '核算主体').font = SHELL_HFONT; ws.cell(2, 2).fill = SHELL_HFILL
    ws.cell(2, 2).border = SHELL_BORDER; ws.cell(2, 2).alignment = SHELL_CEN
    _c0 = 3
    for y in years:
        ws.merge_cells(start_row=2, start_column=_c0, end_row=2, end_column=_c0 + 3)
        cc = ws.cell(2, _c0, f'{y} 年度')
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN
        for _j in range(_c0, _c0 + 4):
            ws.cell(2, _j).border = SHELL_BORDER
        _c0 += 4
    cols3 = ['科目', '核算主体'] + [f'{y}-{m}' for y in years for m in ('期初', '增加', '减少', '期末')]
    for j, h in enumerate(cols3, 1):
        cc = ws.cell(3, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    r = 4
    for typ, nm in blocks:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nc)
        cc = ws.cell(r, 1, nm)
        cc.font = SHELL_BOLD
        cc.fill = SHELL_TOT_FILL
        for _j in range(1, nc + 1):
            ws.cell(r, _j).border = SHELL_BORDER
        r += 1
        for e in ents:
            row_vals = []
            for y in years:
                sv = ent_sv.get((e, y), {})
                v = sv.get(nm, [0.0, 0.0, 0.0, 0.0])
                row_vals += v
            ws.cell(r, 2, e)
            for j, v in enumerate(row_vals, 3):
                cc = ws.cell(r, j, round(v, 2) if abs(v) >= 0.005 else None)
                cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
            ws.cell(r, 1).border = SHELL_BORDER
            ws.cell(r, 2).border = SHELL_BORDER
            ws.cell(r, 2).alignment = SHELL_CEN
            r += 1
        tot_row = []
        for y in years:
            _tv = [0.0, 0.0, 0.0, 0.0]
            for e in ents:
                sv = ent_sv.get((e, y), {})
                _v = sv.get(nm, [0.0, 0.0, 0.0, 0.0])
                for i in range(4):
                    _tv[i] += _v[i]
            tot_row += _tv
        ws.cell(r, 2, '全集团合计')
        ws.cell(r, 2).font = SHELL_BOLD
        for j, v in enumerate(tot_row, 3):
            cc = ws.cell(r, j, round(v, 2) if abs(v) >= 0.005 else None)
            cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
            cc.font = SHELL_BOLD
        for _j in range(1, nc + 1):
            ws.cell(r, _j).border = SHELL_BORDER
        r += 1
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 14
    for j in range(3, nc + 1):
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = 'C4'
    return ws


def build_payroll_multi_period_table(data_dir, out_dir=None):
    """职工薪酬多期对比表（独立跨年文件，2026-08-11 用户：同时并入职工薪酬底稿最新年）。
    渲染复用 _render_payroll_multi_period_sheet。单期账套不生成。"""
    from openpyxl import Workbook
    from audit_shell import finalize_workbook
    wb = Workbook()
    wb.remove(wb.active)
    ws = _render_payroll_multi_period_sheet(wb, data_dir)
    if ws is None:
        wb.close()
        return None
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, '职工薪酬多期对比表_生成.xlsx')
    finalize_workbook(wb)
    wb.save(out)
    wb.close()
    from audit_common import read_tb_full, discover_entities
    try:
        _e2 = discover_entities(data_dir)
        _y2 = sorted({str(y) for e in _e2 for y in _e2[e]})
    except Exception:
        _y2 = []
    print(f'  ✅ 职工薪酬多期对比表：{out}（{"、".join(str(x) for x in _y2) or "?"} 年度）')
    return out
