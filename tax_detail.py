# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=应交税费审计底稿_生成.xlsx(含应交增值税明细表_2026) | 关键列=税种/期末贷方/计提/缴纳/转出 | 职责=税费计提与缴纳勾稽(独立生成器) | 修正=应交税费二级科目归类(_classify_tax_name)：应交增值税222101为二级、金额=三级专栏轧差；转出未交增值税等三级专栏(含'未交增值税'子串)不再误并入未交增值税222103；企业所得税/应交所得税均归222110
"""
应交税费明细表生成程序（多账套 x 多年份，参照职工薪酬小程序 / 往来款小程序 架构）
文件夹模式：自动发现目录下各核算主体，合并输出单工作簿 应交税费审计底稿_生成.xlsx：
  - 编制说明
  - 应交税费审定表（新建：各子目期末未审数取自科目余额表，与明细表同源；置于明细表之前）
  - 应交税费明细表_YYYY（2025 改名 应交税费明细表；含新增「应交增值税222101」行 + 内嵌贷方计提对方科目勾稽）
  - 应交增值税明细表_YYYY（222101 逐月滚动：销项/进项/留抵/转出未交）
  - 税金及附加测算表（复制模板，账面计提关联明细表各税种贷方计提）
  - 借方核对_YYYY（分年独立 sheet；修正取数口径：银行缴税按对方科目含"税"归集，与应交税费借方同口径）
  - 差异凭证汇总（贷方计提勾稽差异的凭证级穿透，按主体/年度/税种分组）
数据源：各主体《科目余额表》(权威控制数) 与《综合查询明细表》(凭证级 对方科目/摘要)。
"""
import os, re, sys, argparse, fnmatch
from collections import defaultdict
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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
# 统一外壳加载逻辑（六程序共享）：中央 audit_templates/ 定位与落盘
from audit_shell import resolve_template, AUDIT_TEMPLATES_DIR, finalize_workbook
from audit_common import is_zero_amount  # 共享库：0 值行判断（与费用一致）

# ⚡ 2026-08-10 SAP 流式读取：只保留税费相关列（col0=过账日期/1=凭证号/3=类型/5=文本/
#   6=公司/10=期间/11=科目码/12=客户/21=供应商/28=凭证日期/29=金额）
_GL_TAX_COLS = {0, 1, 3, 5, 6, 10, 11, 12, 21, 28, 29}

# ================= 常量 =================
# 全税种目标科目（应交税费明细 222103-222114）
# ⚠️ 通用性说明：本表 TARGET/OPP2CODE/税率表 基于【标准用友/金蝶中国科目表 + 中国税法结构】
# （2221* 应交税费子目、680101 所得税费用、城建5%/教育3%/地方教育2% 等），覆盖绝大多数中国账套。
# 若遇【不同客户、科目表被大量自定义】的账套，需在此按新科目表重配 TARGET 子码与 OPP2CODE；
# 主表应交税费仍按【二级科目名称】归集（见 collect_tax_accounts，已不依赖代码），故主表通用，
# 仅「借方核对/计提对方科目勾稽/审定表排序」依赖本清单。曾尝试运行时由 TB 动态推导 TARGET，
# 因银行缴税归因口径(OPP2CODE 顺序)与标准清单不一致导致「借方核对」数值漂移，已回退为静态清单。
TARGET = ['222103', '222104', '222105', '222106', '222107', '222108', '222110', '222111', '222113', '222114', '222115', '222116']
TARGET_NAME = {
    '222103': '未交增值税', '222104': '印花税', '222105': '城市维护建设税', '222106': '房产税',
    '222107': '土地使用税', '222108': '个人所得税', '222110': '企业所得税', '222111': '代扣代缴所得税',
    '222113': '教育费附加', '222114': '地方教育费附加',
    '222115': '水利基金', '222116': '环境保护税'}
CAT = {'222103': 'vat', '222104': 'sur', '222105': 'sur', '222106': 'sur', '222107': 'sur',
       '222108': 'wh', '222110': 'inc', '222111': 'wh', '222113': 'sur', '222114': 'sur',
       '222115': 'sur', '222116': 'sur'}
SUR_CODES = ['222104', '222105', '222106', '222107', '222113', '222114', '222115', '222116']
WH_CODES = ['222108', '222111']

# 应交增值税（222101）叶子名 -> 二位专栏代码
LEAF2CODE = {
    '进项税': '01', '进项税额': '01', '已交税金': '02', '已交增值税': '02',
    '转出未交增值税': '03', '转出未交': '03', '销项税': '04', '销项税额': '04',
    '出口退税': '05', '进项税额转出': '06', '进项税转出': '06',
    '出口免抵税额': '07', '出口抵减内销': '07', '减免税额': '08', '减免税款': '08',
    # ⚡⚡ 2026-08-28 修复（AH 增值税专栏未进明细表）：AH 简易计税/进项加计抵扣（进项加计
    #   抵减）无映射 → 漏出增值税明细表。补 09 简易计税（贷=增加应交）、10 进项加计抵扣（借=抵减应交）。
    '简易计税': '09', '进项加计': '10', '加计抵减': '10', '加计抵减额': '10', '加计抵减税额': '10',
}

# ================= 银行缴税识别（GL 银行账户不带"银行存款"前缀） =================
# 银行存款账户以支行/分行/银行名命名（如 工行横街支行00855、建行路桥支行09443），
# 故不能用 km.startswith('银行存款') 过滤（旧版因此取到 0，导致"应交税费实缴 vs 全部银行存款减少"差异巨大）。
BANK_KW = re.compile(r'支行|分行|银行|账号|储蓄|信用社|农信|建行|工行|中行|农行|交行|招行|兴业|浦发|民生|光大|华夏|平安|邮储|杭州银行|台州银行|宁波')
# GL 对方科目(短名) -> TARGET 代码；222101 应交增值税专栏（进项/销项/出口/减免/已交税金/待认证）不计入"应交税费借方实缴"
OPP2CODE = [('未交增值税', '222103'), ('印花税', '222104'), ('城建税', '222105'), ('城市维护建设税', '222105'),
            ('房产税', '222106'), ('土地使用税', '222107'), ('个人所得税', '222108'), ('企业所得税', '222110'),
            ('代扣代缴', '222111'), ('教育费附加', '222113'), ('地方教育费附加', '222114'), ('地方教育', '222114')]
VAT_KW = re.compile(r'进项|销项|出口退税|出口免抵|减免税款|已交税金|待认证')
def tax_code_of_opp(opp):
    """GL 对方科目(短名) -> TARGET 代码；属 222101 增值税专栏返回 None。"""
    if VAT_KW.search(opp):
        return None
    for k, v in OPP2CODE:
        if k in opp:
            return v
    return None

# ================= 样式 =================
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM = '#,##0.00'
HFILL = PatternFill('solid', fgColor='DDEBF7')
HFONT = Font(name='Times New Roman', bold=True, color='000000', size=10)
TFILL = PatternFill('solid', fgColor='D9E1F2')
SFILL = PatternFill('solid', fgColor='F2F2F2')
BFONT = Font(name='Times New Roman', bold=True)
TITLE_FONT = Font(name='Times New Roman', bold=True, size=10, color='000000')
SUB_FONT = Font(name='Times New Roman', bold=True, size=10)
RED_FONT = Font(name='Times New Roman', bold=True, color='9C0006')
WARN_FILL = PatternFill('solid', fgColor='FCE4D6')
DIFF_FILL = PatternFill('solid', fgColor='FFF2CC')
OK_FILL = PatternFill('solid', fgColor='E2EFDA')
SECF = PatternFill('solid', fgColor='548235')
CRHF = PatternFill('solid', fgColor='C55A11')
TOT_FILL = PatternFill('solid', fgColor='FFF2CC')  # 黄色底色，用作小计/合计行（与 DIFF_FILL 同色）
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
RGT = Alignment(horizontal='right', vertical='center')


def fnum(ws, r, c, v):
    cell = ws.cell(r, c, v if v is not None else None)
    if isinstance(v, (int, float)) and c not in (1, 2):
        cell.number_format = NUM
    cell.border = BORDER
    return cell


def safe(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def num(v):
    return safe(v)


def sval(amt, direction):
    a = num(amt)
    d = str(direction).strip()
    if d == '贷':
        return a
    if d == '借':
        return -a
    return 0.0


# ================= 年份识别（稳健版，兼容"25"缩写+导出时间戳命名） =================
def _extract_tax_year(fn):
    """从文件名提取会计年度，兼容两种命名约定：
    ① 标准 4 位年份（其后非数字）：如 诚本2025科目余额表 / 2025东轴科目余额表 / 上海朗炫科目余额表2026；
    ② 主体简称后的 2 位年份缩写（其后可夹"外币"等再接报表关键字）：如 上分25科目余额表 / 新加坡25外币科目余额表。
    必须排除导出时间戳（_YYYYMMDD[HHMMSS] 或紧贴 .xlsx 前的 YYYYMMDD[HHMMSS]），
    否则会被误读为年度（如 上分25科目余额表_20260723152733.xlsx 的时间戳含 2026）。"""
    # 1) 去除导出时间戳（8~14 位数字，可带前导下划线，位于文件名尾）
    stripped = re.sub(r'_?\d{8}\d{0,6}(?=\.xlsx?$)', '', fn, flags=re.IGNORECASE)
    # 2) 优先：独立 4 位年份（前/后均非数字）
    m = re.search(r'(?<!\d)((?:19|20)\d{2})(?!\d)', stripped)
    if m:
        return m.group(1)
    # 3) 次选：报表关键字（科目余额表/综合查询明细表/辅助核算余额表）前的 2 位年份缩写
    m = re.search(r'(?<!\d)(\d{2})\D*(科目余额表|综合查询明细表|综合查询表|辅助核算余额表)', stripped)
    if m:
        yy = int(m.group(1))
        return ('19%02d' % yy) if yy >= 80 else ('20%02d' % yy)
    return None


# ================= 实体发现 =================
def _discover_tax_entities(data_dir):
    # ⚡ 2026-08-09 治本统一：本地 discover 收敛为 audit_common.discover_entities——
    # U8 模式读原文件全量；SAP 模式 patch 后=adapter（current_comp 单主体过滤内置）。
    from audit_common import discover_entities as _de
    return _de(data_dir)


# ================= GL 读取（凭证级） =================
def read_gl_tax(gl_path):
    rec = {
        'vat': {code: {m: {'d': 0.0, 'c': 0.0} for m in range(1, 13)} for code in
                ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10']},
        'vat_paid': {m: 0.0 for m in range(1, 13)},
        'rows': [],          # 全部 2221* 相关行（用于归因/穿透）
        'bank_tax_pay': {c: 0.0 for c in TARGET},  # 银行存款贷方·对方科目含"税"·归属 TARGET(不含222101)
    }
    # SAP：文件级流式（只保留 2221* 税科目 + 1002* 银行 + 6403* 税金及附加 + 6801* 所得税费用 行，
    #       按过账日期(col0 YYYYMMDD) 1-7 月过滤——与 TB"借方1-7"列口径一致，解决 8/5 挂 07 期间差异）
    if _adapter is not None and (gl_path is None or not isinstance(gl_path, str) or _adapter.is_sap(gl_path)):
        return _read_gl_tax_sap_stream(rec)
    if not gl_path or not os.path.exists(gl_path):
        return rec
    wb = openpyxl.load_workbook(gl_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 8:
            continue
        km = str(r[0]) if r[0] is not None else ''
        dt = str(r[1]) if len(r) > 1 and r[1] is not None else ''
        zi = str(r[2]) if len(r) > 2 and r[2] is not None else ''
        hao = str(r[3]) if len(r) > 3 and r[3] is not None else ''
        opp = str(r[4]) if len(r) > 4 and r[4] is not None else ''
        zy = str(r[5]) if len(r) > 5 and r[5] is not None else ''
        jf = safe(r[6])
        df = safe(r[7])
        month = None
        mm = re.search(r'(\d{4})-(\d{2})', dt)
        if mm:
            month = int(mm.group(2))
        vkey = (dt, zi, hao)
        # 银行缴税（GL 银行账户无"银行存款"前缀，以支行/银行等识别）：贷方且对方科目含"税"，归属 TARGET(不含222101增值税专栏)
        if BANK_KW.search(km) and df > 0.01 and '税' in opp:
            _c = tax_code_of_opp(opp)
            if _c and _c in TARGET:
                rec['bank_tax_pay'][_c] += df
        if not (km.startswith('应交税费') or km.startswith('税金及附加')
                or km.startswith('营业税金及附加')):
            continue
        rec['rows'].append(dict(km=km, dt=dt, zi=zi, hao=hao, opp=opp, zy=zy,
                                jf=jf, df=df, month=month, vkey=vkey))
        if km.startswith('应交税费-应交增值税-'):
            parts = km.split('-')
            leaf = parts[2] if len(parts) >= 3 else ''
            code = LEAF2CODE.get(leaf)
            _zy = zy or ''
            if month:
                # ⚡ 2026-08-11 DQ4 修复：专栏判定【摘要优先】——科目名=专栏名但摘要
                #   含特殊标识时按摘要归栏（DQ 进项转出凭证科目='进项税额' 摘要=
                #   '进项税额转出' → 应记 06 非 01；原只在"未划 3 级明细"时走摘要识别，
                #   漏掉此形态 → 进项转出 7 笔全记入 01，明细表 22210106 列=0）。
                if '结转' in _zy:
                    code = None          # 结转凭证（借结转销项/贷结转进项）不计专栏
                elif '转出' in _zy and '进项' in _zy:
                    code = '06'          # 进项税额转出（贷方发生）
                elif any(k in _zy for k in ('减免', '免税')):
                    code = '08'          # 减免税额（贷方）
                elif jf > 0.01 and any(k in _zy for k in ('缴', '交税', '缴税', '扣税', '入库', '实缴', '已交', '交纳')):
                    code = '02'          # 已交税金（借方）
                elif not code:
                    # 未划 3 级明细（科目名仅"应交税费-应交增值税"）且摘要无特殊标识：
                    # 贷方默认销项、借方默认进项（DQ 实测：贷方=检测费/货款(销项)+结转进项税(结转)）
                    code = '04' if df > 0.01 else ('01' if jf > 0.01 else None)
            if code and month:
                rec['vat'][code][month]['d'] += jf
                rec['vat'][code][month]['c'] += df
        elif km == '应交税费-未交增值税':
            if month:
                rec['vat_paid'][month] += jf  # 借 222103 = 实际缴纳已结转未交增值税
    wb.close()
    return rec


def _read_gl_tax_sap_stream(rec):
    """SAP 序时账 → rec（文件级流式，只保留税费相关行，避免全量 66 万行 OOM）。
    过滤：
      · 公司=current_comp（col6）
      · 过账日期（col0 YYYYMMDD）∈ 2026-01~07 —— 与 TB「借方 1-7」列口径一致
        （SAP 8/5 挂 07 期间的凭证 col0=20260805 被剔 → 税金/所得税 GL=TB，期间口径修复）
      · 科目内部码前缀：2221（应交税费）/1002（银行）/6403（税金及附加）/6801（所得税费用）
    仅对保留行做凭证级对方科目推断（_infer_opp），行量约 3.9 万（vs 全量 67 万）。"""
    # ⚡⚡ 2026-08-24 集团模式修复：原 `if not comp: return rec`（current_comp=None）
    #   直接返回空 → 应交税费/税金及附加/所得税 集团审定表全 0（AH 三集团质检发现）。
    #   改为逐主体聚合：current_comp 有值 → 单主体；否则 → _GROUP_COMPS（--only 注入）
    #   或 discover 全量。每主体 keep_rows 各自 _infer_opp（防不同主体 vno 撞号误判）。
    if not _adapter._DATA_ROOT:
        return rec
    _cur_c = _adapter.current_comp()
    if _cur_c:
        _comps = [_cur_c]
    else:
        _comps = sorted(getattr(_adapter, '_GROUP_COMPS', None) or [])
        if not _comps:
            try:
                _comps = sorted(_adapter.discover_entities(_adapter._DATA_ROOT).keys())
            except Exception:
                _comps = []
    try:
        _adapter.read_tb_full(None, None)   # 确保 code→name 映射已建
    except Exception:
        pass
    # ⚡⚡ 2026-08-25 性能优化：read_gl_rows（audit_cache GL 磁盘缓存）替代逐主体
    #   流式读原始 GL 文件——AH 单文件 10 万+ 行，tax 每主体流式读 100s+（1357 集团
    #   19 主体 ≈ 32 分钟、2468 ≈ 1-2 小时），是 AH 全套 1 小时目标的瓶颈。read_gl_rows
    #   行含 e/date/month/vtype/vno/summary/name/code/cust/supp/dr/cr，字段完全覆盖
    #   tax 过滤需求（期间 1-7 月、科目前缀、对方科目），pkl 秒级加载 + 内存过滤。
    _ents_all = {}
    try:
        _ents_all = _adapter.discover_entities(_adapter._DATA_ROOT) or {}
    except Exception:
        _ents_all = {}
    keep_rows = []
    # ⚡⚡ 2026-08-26 修复（2468 tax OOM）：原 `read_gl_rows(_DATA_ROOT, _ents_all)` 全量
    #   加载 88 主体 455 万行 GL 到内存（_gl_all 常驻），68 主体集团内存峰值 21GB+ 被杀。
    #   改【逐主体分批读】：每批读单主体 GL（磁盘 pkl 秒级加载），过滤后立即 del 释放——
    #   峰值 = 单主体行数（几十万），不再累积全集团。
    _ents_one = {_cc: _ents_all.get(_cc) or {} for _cc in _comps}
    for _cc in _comps:
        _rows_c = []
        try:
            _rows_cc = _adapter.read_gl_rows(_adapter._DATA_ROOT, {_cc: _ents_one.get(_cc)})
        except Exception:
            _rows_cc = []
        for r in _rows_cc:
            if str(r.get('e') or '') != _cc:
                continue
            _m = str(r.get('month') or '0')
            try:
                _mi = int(_m)
            except (TypeError, ValueError):
                _mi = 0
            if not (1 <= _mi <= 7):
                continue   # 期间 1-7 月（口径对齐 TB）
            code = str(r.get('code') or '')
            if not (code.startswith('2221') or code.startswith('1002')
                    or code.startswith('6403') or code.startswith('6801')):
                continue
            nm = str(r.get('name') or '') or code
            dr = float(r.get('dr') or 0.0)
            cr = float(r.get('cr') or 0.0)
            if dr == 0 and cr == 0:
                continue
            _rows_c.append({
                'e': _cc, 'y': str(r.get('y') or '2026'),
                'date': r.get('date'),
                'month': _mi,
                'vtype': str(r.get('vtype') or ''),
                'vno': str(r.get('vno') or ''),
                'summary': str(r.get('summary') or ''),
                'name': nm,
                'code': code,
                'cust': str(r.get('cust') or '').strip(),
                'supp': str(r.get('supp') or '').strip(),
                'dr': dr, 'cr': cr,
            })
        keep_rows.extend(_adapter._infer_opp(_rows_c))   # 凭证级对方科目（仅保留行，快）
        del _rows_c, _rows_cc
    for r in keep_rows:
        km = str(r.get('name') or '')
        dt = str(r.get('date') or '')
        zi = str(r.get('vtype') or '')
        hao = str(r.get('vno') or '')
        opp = str(r.get('cp') or '')
        zy = str(r.get('summary') or '')
        jf = float(r.get('dr') or 0.0)
        df = float(r.get('cr') or 0.0)
        month = int(r.get('month') or 0) or None
        vkey = (dt, zi, hao)
        if BANK_KW.search(km) and df > 0.01 and '税' in opp:
            _c = tax_code_of_opp(opp)
            if _c and _c in TARGET:
                rec['bank_tax_pay'][_c] += df
        if not (km.startswith('应交税费') or km.startswith('税金及附加')
                or km.startswith('营业税金及附加')):
            continue
        rec['rows'].append(dict(km=km, dt=dt, zi=zi, hao=hao, opp=opp, zy=zy,
                                jf=jf, df=df, month=month, vkey=vkey))
        if '应交增值税' in km:
            parts = km.split('-')
            leaf = ''
            for i, p in enumerate(parts):
                if p == '应交增值税' and i + 1 < len(parts):
                    leaf = parts[i + 1]
                    break
            code = LEAF2CODE.get(leaf)
            if code and month:
                rec['vat'][code][month]['d'] += jf
                rec['vat'][code][month]['c'] += df
        elif '进项加计' in km:
            # ⚡⚡ 2026-08-28 AH 进项加计抵扣（2221010205，借=抵减应交）归专栏 10
            if month:
                rec['vat']['10'][month]['d'] += jf
                rec['vat']['10'][month]['c'] += df
        elif '未交增值税' in km and '应交增值税' not in km:
            if month:
                rec['vat_paid'][month] += jf  # 借 未交增值税 = 实际缴纳已结转
    return rec


# ================= TB 读取 =================
def load_tb(fn):
    # ⚡ 2026-08-10 SAP 适配：SAP 科目余额表列结构（col2=科目名/col12=借1-7/col13=贷1-7/col16=累计余额）
    #   与 U8（col0=代码/col1=名称/col2-7=借贷余）完全不同 → 本地 U8 读取会全部错位（只读 1 行）。
    #   SAP 场景改走 adapter.read_tb_full（{(comp, code, name, y): {qc,jf,df,qm,level}}），
    #   转换为 U8 load_tb 同构 {code: {name, qc_dir, qc, deb, cre, qm_dir, qm}}。
    if _adapter is not None and (fn is None or not isinstance(fn, str) or _adapter.is_sap(fn)):
        return _load_tb_sap(fn)
    wb = openpyxl.load_workbook(fn, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or len(r) < 7:
            continue
        code = r[0]
        if code is None:
            continue
        code = str(code).strip()
        rows[code] = {
            'name': str(r[1] or '').strip(),
            'qc_dir': r[2], 'qc': num(r[3]),
            'deb': num(r[4]), 'cre': num(r[5]),
            'qm_dir': r[6], 'qm': num(r[7]) if len(r) > 7 else 0.0,
        }
    wb.close()
    return rows


def _load_tb_sap(fn=None):
    """SAP 科目余额表 → U8 load_tb 同构（key=10 位内部码）。
    adapter.read_tb_full 返回 {(comp, code, name, y): {qc,jf,df,qm,level}}；
    SAP 码为 10 位末级子目（无 4 位父级行）→ 按名称匹配消费（collect_tax_accounts 名称优先）。
    同时为应交增值税 222101 生成虚拟父级行（Σ222101* 子目），供 get_open_liudi/
    _find_vat_parent_code 取期初留抵与父级汇总（铁律75：SAP TB 无父级行→虚拟父级聚合）。
    ⚡ 2026-08-10 集团模式：主体从 km 路径提取（{comp}/{comp}科目余额表...），替代 current_comp
    （集团下 current_comp=None → 全滤空 → 审定表全 0）。"""
    rows = {}
    full = _adapter.read_tb_full(None, None)   # 全量（含全部主体）
    # ⚡ 2026-08-13 治本（与 _scope_comp 同构）：单主体 current_comp 优先——原实现
    #   路径提取优先（current_comp 仅兜底），区间合并文件（2250~2300科目余额表.xlsx）
    #   文件名 `re.match(r'^(\d{4})')` 取到【区间起点】2250 → 2260 底稿串成 2250 数据。
    comp_scope = None
    _cc = _adapter.current_comp()
    if _cc:
        comp_scope = _cc
    elif fn and isinstance(fn, str):
        # ⚡⚡ 2026-08-26 修复（1357 应交税费审定表全 0 根因）：原 `re.search(r'[\\/](\d{4})[\\/]')`
        #   会误匹配【年份目录】/2026/（路径 D:/.../数据/2026/科目余额表/1020.xlsx 先命中 /2026/）
        #   → comp_scope='2026' ≠ 任何主体 → rows 全空 → 审定表全 0（AH 1357 19 主体历史遗留）。
        #   改为【basename 主体码优先】——文件名前缀 4 位数字（1020.xlsx→1020），天然排除年份目录。
        _bn = os.path.basename(fn)
        m = re.match(r'^(\d{4})', _bn)
        if m:
            comp_scope = m.group(1)
        else:
            m = re.search(r'[\\/](\d{4})[\\/]', fn)
            if m:
                comp_scope = m.group(1)
    if comp_scope is None:
        comp_scope = _adapter.current_comp()
    for (comp, code, name, y), v in full.items():
        if comp != comp_scope:
            continue
        qc = float(v.get('qc') or 0.0)
        qm = float(v.get('qm') or 0.0)
        rows[str(code)] = {
            'name': str(name or ''),
            'qc_dir': '借' if qc > 0 else ('贷' if qc < 0 else '平'),
            'qc': abs(qc),
            'deb': float(v.get('jf') or 0.0),
            'cre': float(v.get('df') or 0.0),
            'qm_dir': '借' if qm > 0 else ('贷' if qm < 0 else '平'),
            'qm': abs(qm),
        }
    # 虚拟父级 222101（应交增值税）：Σ 222101* 子目（借贷/期初期末带符号汇总）
    subs = [c for c in rows if c.startswith('222101') and len(c) > 6]
    if subs:
        qc = sum(sval(rows[c]['qc'], rows[c]['qc_dir']) for c in subs)
        qm = sum(sval(rows[c]['qm'], rows[c]['qm_dir']) for c in subs)
        rows['222101'] = {
            'name': '应交税费-应交增值税',
            'qc_dir': '借' if qc > 0 else ('贷' if qc < 0 else '平'),
            'qc': abs(qc),
            'deb': sum(rows[c]['deb'] for c in subs),
            'cre': sum(rows[c]['cre'] for c in subs),
            'qm_dir': '借' if qm > 0 else ('贷' if qm < 0 else '平'),
            'qm': abs(qm),
        }
    return rows


def get_account(rows, code):
    if code not in rows:
        return {'qc': 0.0, 'inc': 0.0, 'dec': 0.0, 'qm': 0.0,
                'name': TARGET_NAME.get(code, code), 'exist': False}
    x = rows[code]
    return {'qc': sval(x['qc'], x['qc_dir']), 'inc': x['cre'], 'dec': x['deb'],
            'qm': sval(x['qm'], x['qm_dir']), 'name': x['name'], 'exist': True}


def parent_or_sum(rows, prefix):
    pdeb = num(rows[prefix]['deb']) if prefix in rows else 0.0
    pcre = num(rows[prefix]['cre']) if prefix in rows else 0.0
    cdeb = sum(num(r['deb']) for c, r in rows.items() if c.startswith(prefix) and len(c) > len(prefix))
    ccre = sum(num(r['cre']) for c, r in rows.items() if c.startswith(prefix) and len(c) > len(prefix))
    if pdeb > 0 or pcre > 0:
        return pdeb, pcre
    return cdeb, ccre


def get_parent_account(tb, prefix):
    """取父级科目(含期初/期末 signed)；父级行缺失时按子目汇总。"""
    if prefix in tb:
        return get_account(tb, prefix)
    qc = qm = inc = dec = 0.0
    for c, r in tb.items():
        if c.startswith(prefix) and len(c) > len(prefix):
            a = get_account(tb, c)
            qc += a['qc']; qm += a['qm']; inc += a['inc']; dec += a['dec']
    return {'qc': qc, 'inc': inc, 'dec': dec, 'qm': qm,
            'name': TARGET_NAME.get(prefix, prefix), 'exist': True}


# ================= 按「二级科目名称」匹配（关键：不同公司记账代码可能不一致） =================
# 应交税费二级科目名称 -> (标准代码, 标准名, 分类)
def _classify_tax_name(name, code=None):
    """按【二级科目名称】归类应交税费(2221*)各子目 -> (标准代码, 标准名, 分类)。
    ★ 关键修正：
      · 应交增值税(222101)为二级科目，其下三级专栏(进项税/销项税/转出未交增值税/减免税额等)
        由父级222101汇总(轧差)，【不单列】——故「转出未交增值税」等三级专栏返回 None，
        避免被 '未交增值税' 子串误并入 222103(未交增值税)。
      · 未交增值税(222103)是独立的二级科目，仅当科目名含 '未交增值税' 且不属于应交增值税专栏时归 222103。
      · 企业所得税/应交所得税 均归 222110（名称可能为 '企业所得税' 或 '应交所得税'）。"""
    n = (name or '').strip()
    if not n:
        return None
    # 应交增值税 三级专栏(进项/销项/转出未交/减免等)：由父级222101汇总，不单列
    if '应交增值税' in n and len(code or '') > 6:
        return None
    if '应交增值税' in n:
        return ('222101', '应交增值税', 'vat')      # 应交增值税(222101，二级，金额=三级轧差)
    if '未交增值税' in n:
        # 2026-08-05 修复：『转出未交增值税』是应交增值税三级专栏（销项转出待缴），
        # 名称仅含『未交增值税』子串、不含『应交增值税』，原排除条件(应交增值税 in n)
        # 失效 → 被误并入 222103（FY07 未交增值税期初 18,061.39 = 转出专栏 18,852.41
        # + 真未交增值税 -791.02）。显式排除：归父级 222101 轧差，不单列。
        if '转出未交增值税' in n or '减免税额' in n or '减免税款' in n:
            return None
        return ('222103', '未交增值税', 'vat')      # 未交增值税(独立二级科目，非应交增值税专栏)
    if '印花税' in n:
        return ('222104', '印花税', 'sur')
    if '城市维护建设税' in n or '城建' in n:
        return ('222105', '城市维护建设税', 'sur')
    if '房产税' in n:
        return ('222106', '房产税', 'sur')
    if '土地使用税' in n:
        return ('222107', '土地使用税', 'sur')
    if '个人所得税' in n:
        return ('222108', '个人所得税', 'wh')     # 个税(可能记为222103/222109等)
    if '代扣' in n:
        return ('222111', '代扣代缴所得税', 'wh')
    if '所得税' in n:
        return ('222110', '企业所得税', 'inc')     # 企业所得税 / 应交所得税(名称可能为应交所得税)
    if '教育费附加' in n:
        return ('222113', '教育费附加', 'sur')
    if '地方教育' in n:
        return ('222114', '地方教育费附加', 'sur')
    # ⚡⚡ 2026-08-24 修复：水利基金/环境保护税 未识别 → 审定表漏项
    #   （AFJ 水利基金 66,466.78 + 环保税 6,813.97 = 73,280.75 阶段③差额）
    if '水利' in n:
        return ('222115', '水利基金', 'sur')
    if '环境保护税' in n or '环保税' in n:
        return ('222116', '环境保护税', 'sur')
    return None


def collect_tax_accounts(tb):
    """按【二级科目名称】归集应交税费(2221*)各子目 -> 以标准代码为键的账户字典。
    不同公司记账代码可能不一致(如 未交增值税 记为 222102/222106/222108、
    个人所得税 记为 222103/222109 等，甚至 222103/222108 互换)，
    故一律以名称为准判定税种，避免代码错配/互换导致匹配错误。
    ★ 2026-07-31 修复：规避【父+子双计】——同一税种下父级名称与子级名称同时命中
      （如 FY 未交增值税 2221.02 父级 + 2221.02.01 应交未交增值税额 子级 均含'未交增值税'），
      父级借/贷发/期初期末已=子目之和，若全计则金额×2。修复：同税种组内，存在以其代码
      为前缀的其他命中代码（有子目被单独归集）时跳过父级，只取叶子。"""
    acc = {}
    for std in TARGET + ['222101']:
        acc[std] = {'qc': 0.0, 'inc': 0.0, 'dec': 0.0, 'qm': 0.0,
                     'name': TARGET_NAME.get(std, std), 'exist': False, 'code': std}
    groups = {}  # std -> [(code, r, disp)]
    for code, r in tb.items():
        if not code.startswith('2221'):
            continue
        cl = _classify_tax_name(r.get('name'), code)
        if not cl:
            continue
        std, disp, _ = cl
        groups.setdefault(std, []).append((code, r, disp))
    for std, items in groups.items():
        a = acc[std]
        codes = [c for c, _, _ in items]
        for code, r, disp in items:
            # 父级（存在以其代码为前缀的其他命中代码）→ 跳过，只取叶子（铁律6：父=子不重复计）
            if any(code != p and p.startswith(code) for p in codes):
                continue
            a['qc'] += sval(r.get('qc'), r.get('qc_dir'))
            a['inc'] += num(r.get('cre'))
            a['dec'] += num(r.get('deb'))
            a['qm'] += sval(r.get('qm'), r.get('qm_dir'))
            a['exist'] = True
            a['name'] = disp            # 以实际科目名称展示
            a['code'] = code            # 记录实际代码(备查)
    return acc


def _name_match_codes(tb, includes, excludes=None):
    res = []
    for code, r in tb.items():
        nm = r.get('name') or ''
        if excludes and any(e in nm for e in excludes):
            continue
        if any(k in nm for k in includes):
            res.append(code)
    return res


def sum_pnl_debit(tb, includes, excludes=None):
    """按【名称】归集利润表科目借方(借发)，并规避父子科目重复计列：
    若较短匹配码是较长匹配码的前缀(父级已汇总子目)，则只取子目。"""
    codes = sorted(_name_match_codes(tb, includes, excludes), key=len)
    if not codes:
        return 0.0, []
    used = [c for c in codes if not any(c != p and c.startswith(p) for p in codes)]
    tot = sum(num(tb[c]['deb']) for c in used)
    return tot, used


def _find_vat_parent_code(tb):
    """定位应交增值税父级代码：标准 222101 优先，否则按名称定位（部分公司代码非标）。"""
    if '222101' in tb:
        return '222101'
    for code, r in tb.items():
        nm = r.get('name') or ''
        if '应交增值税' in nm and '未交' not in nm:
            return code
    return '222101'


# ================= 增值税逐月滚动 =================
def get_open_liudi(tb):
    """应交增值税期初留抵（借方余额=留抵，贷方余额=应缴）。
    2026-07-31 修复：父级代码按名称定位（FY 账套为 '2221.01' 而非标准 '222101'），
    原 `if '222101' in tb` 对非标代码返回 0，导致逐月滚动全部错位。"""
    vp = _find_vat_parent_code(tb)
    if vp not in tb:
        return 0.0, '无', 0.0, False
    x = tb[vp]
    d = x['qc_dir']
    v = num(x['qc'])
    if d == '借':
        return float(v), d, float(v), False
    elif d == '贷' and abs(v) > 0.005:
        return -float(v), d, float(v), True
    return 0.0, (d or '平'), float(v), False


def build_vat_entity(gl, tb):
    open_liudi, odir, oamt, oerr = get_open_liudi(tb)
    mv = gl['vat']
    paid = gl['vat_paid']
    sxj_d = sum(mv['04'][m]['d'] for m in range(1, 13))
    sxj_c = sum(mv['04'][m]['c'] for m in range(1, 13))
    # ⚡⚡ 2026-08-28 修复（AH 应交/未交差异 3.25 亿根因）：原 `clearing = sxj_d > 1%*sxj_c`
    #   把销项税借方大（红字冲回）误判为"已交税金对冲"账套 → 用销项【贷方全额】计算应交，
    #   未减红字冲回（AH 1010 销项贷 7.78 亿中含"无票收入预提+开票确认"双计、借 5.59 亿为
    #   冲销无票收入）→ 计算应转出 5.3 亿 vs 账载转出 2.05 亿（=净销项口径）。改统一【净额】
    #   （贷-借，红字冲回天然抵减；已交税金在 AH 独立专栏 2221010202 已交，不在销项）。
    clearing = False
    rows = []
    liudi = open_liudi
    for m in range(1, 13):
        if clearing:
            n01 = mv['01'][m]['d']; n02c = mv['02'][m]['d']
            n07 = mv['07'][m]['d']; n08 = mv['08'][m]['d']
            n04 = mv['04'][m]['c']; n05 = mv['05'][m]['c']; n06 = mv['06'][m]['c']
        else:
            n01 = mv['01'][m]['d'] - mv['01'][m]['c']
            n02c = mv['02'][m]['d'] - mv['02'][m]['c']
            n07 = mv['07'][m]['d'] - mv['07'][m]['c']
            n08 = mv['08'][m]['d'] - mv['08'][m]['c']
            n04 = mv['04'][m]['c'] - mv['04'][m]['d']
            n05 = mv['05'][m]['c'] - mv['05'][m]['d']
            n06 = mv['06'][m]['c'] - mv['06'][m]['d']
        paid_222103 = paid[m]
        d03_actual = mv['03'][m]['d']
        # ⚡⚡ 2026-08-28 简易计税(09)/进项加计抵扣(10) 纳入净额（AH 专栏）：简易计税贷=增加应交，
        #   加计抵扣借=抵减应交（增值税计算口径，用户要求完整反映）。
        n09 = mv['09'][m]['c'] - mv['09'][m]['d']
        n10 = mv['10'][m]['d'] - mv['10'][m]['c']
        dr_move = n01 + n02c + n07 + n08 + n10
        cr_move = n04 + n05 + n06 + n09
        opening = liudi
        net = cr_move - dr_move - opening
        if net > 0.005:
            payable = net; ending = 0.0
        else:
            payable = 0.0; ending = -net
        rows.append(dict(m=m, opening=opening, d01=n01, d02=n02c, d07=n07, d08=n08,
                         paid=paid_222103, dr_move=dr_move, c04=n04, c05=n05, c06=n06,
                         d09=n09, d10=n10,
                         cr_move=cr_move, net=net, payable=payable, d03_actual=d03_actual, ending=ending))
        liudi = ending
    return dict(open_liudi=open_liudi, odir=odir, oamt=oamt, oerr=oerr, rows=rows,
                ending=liudi, clearing=clearing)


# ================= 应交税费全税种 =================
def build_full_entity(gl, tb):
    # ★ 全部按【二级科目名称】匹配（不同公司记账代码可能不一致，代码匹配会错配/互换）
    acc = collect_tax_accounts(tb)                          # 应交税费各子目(名称归集)
    tax_sur_deb, _ = sum_pnl_debit(tb, ['税金及附加'])   # 利润表 税金及附加(5403/6403/营业税金及附加 等)
    inc_cur = abs(sum_pnl_debit(tb, ['当期所得税', '所得税费用'], excludes=['递延'])[0])  # 所得税费用-当期(5801/680101 等)
    zc = sum_pnl_debit(tb, ['转出未交增值税', '转出未交'])[0]         # 转出未交增值税(若公司设独立专栏)
    # ⚡⚡ 2026-08-28 修复（AH 应交/未交勾稽差 2.05 亿根因）：AH 科目名『应交税费-应交税金-应交
    #   增值税-转出未交』【缺"增值税"三字】→ 原匹配词『转出未交增值税』不命中 → zc=0 → 走 GL
    #   摘要兜底只取到 187,470 → 与未交增值税(222103)贷方 2.05 亿勾稽差异 2.05 亿。
    #   匹配词加『转出未交』（覆盖 AH 简名）。
    if abs(zc) <= 0.01:
        # 无"转出未交增值税"专栏（FY 账套）→ GL 凭证级判定：未交增值税贷方行摘要含"结转"
        # （"结转X月增值税"）即账载结转数；差异=贷方-结转 为补计/税率调整/代扣等非结转部分
        zc = sum((x.get('df') or 0) for x in gl['rows']
                 if ('未交增值税' in (x.get('km') or '')) and ('应交增值税' not in (x.get('km') or ''))
                 and (x.get('df') or 0) > 0.01 and ('结转' in (x.get('zy') or '')))
    yyszc, _ = sum_pnl_debit(tb, ['营业外支出'])          # 营业外支出(名称匹配；税收罚款/滞纳金等可能计入)
    _, bank_cre = parent_or_sum(tb, '1002')
    # 222101 应交增值税（父级/汇总，按名称定位，标准码兜底）
    vat_parent = get_parent_account(tb, _find_vat_parent_code(tb))
    # 所得税计提列支核对（穿透）
    accr = acc['222110']['inc']                            # 企业所得税贷方(名称匹配)
    exp = inc_cur                                           # 当期所得税费用借方(名称匹配)
    inc_diff = accr - exp
    inc_items = []
    for x in gl['rows']:
        if '所得税' not in x['km'] and '222110' not in x['km'] and '应交所得税' not in x['km']:
            continue
        if '个人所得税' in x['km']:
            # 个人所得税/代扣代缴类(222108/222111)：对方科目为应付职工薪酬等代扣性质，
            # 已在主表「代扣代缴对方」列示，差异凭证汇总不再穿透（避免误归入企业所得税）。
            continue
        via_exp = '所得税费用' in x['opp']
        if x['df'] > 0.01 and not via_exp:
            inc_items.append((x['zi'], x['hao'], x['zy'], x['opp'], x['df']))
        if x['jf'] > 0.01 and via_exp:
            inc_items.append((x['zi'], x['hao'], x['zy'], x['opp'], x['jf']))
    # 增值税转出核对（穿透）：未交增值税直接计入（非经"转出未交增值税"专栏）
    u103 = acc['222103']['inc']                            # 未交增值税贷方(名称匹配)
    zc2 = zc
    vat_diff = u103 - zc2
    vat_items = []
    for x in gl['rows']:
        if ('未交增值税' in x['km']) and ('转出' not in x['km']) and ('应交增值税' not in x['km']):
            if abs(x['df']) > 0.01 and ('增值税' not in x['opp']):
                vat_items.append((x['zi'], x['hao'], x['zy'], x['opp'], x['df']))
    # 税金及附加 穿透（名称匹配），解释 税金及附加借 - 税金及附加类贷方 差异
    sur_items = []
    for x in gl['rows']:
        # ⚡⚡ 2026-08-28 修复（sur_items 收集错对象）：原条件含 `'税金及附加' in x['opp']`，
        #   会把『借 应交税费 / 贷 税金及附加』的红字冲回凭证（2221 借方行，opp 含税金及附加）
        #   误收成 6403 借方 → 差异凭证汇总方向错乱（6500000840 冲回城建税被当差异列出）。
        #   只认【科目本身】为 6403 税金及附加 且 借方>0 的行。
        if '税金及附加' in x['km'] or '营业税金及附加' in x['km']:
            if abs(x['jf']) > 0.01:
                sur_items.append((x['zi'], x['hao'], x['zy'], x['opp'], x['jf']))
    # ⚡⚡ 2026-08-28 修复（差异凭证汇总混入勾稽一致计提凭证，用户指出）：原把【全部】税金及附加
    #   借方凭证列进差异清单，其中大部分是"借 税金及附加 / 贷 应交税费"的计提凭证（借贷相等、
    #   勾稽一致），应只保留【真差异】凭证（直接支付车船税/房产税等无对应计提的）。按凭证级过滤：
    #   同凭证 税金及附加借 ≈ 应交税费贷 的计提凭证剔除（含多税种混合凭证 6500000840 等）。
    if sur_items:
        _vouch_tax = defaultdict(lambda: [0.0, 0.0])   # (zi,hao) -> [税金及附加借, 应交税费贷]
        for _x in gl['rows']:
            _k = (str(_x.get('zi') or ''), str(_x.get('hao') or ''))
            _km = str(_x.get('km') or ''); _opp = str(_x.get('opp') or '')
            if abs(float(_x.get('jf') or 0)) > 0.005 and ('税金及附加' in _km or '税金及附加' in _opp):
                _vouch_tax[_k][0] += float(_x.get('jf') or 0)
            if abs(float(_x.get('df') or 0)) > 0.005 and ('应交税费' in _km or '应交税费' in _opp):
                _vouch_tax[_k][1] += float(_x.get('df') or 0)
        _sur_keep = []
        for _it in sur_items:
            _de, _cr = _vouch_tax.get((str(_it[0] or ''), str(_it[1] or '')), [0.0, 0.0])
            if abs(_de - _cr) > 0.01:                  # 借贷不等 → 真差异凭证，保留
                _sur_keep.append(_it)
        sur_items = _sur_keep
    # 贷方（计提）对方科目勾稽：每税种 计提(贷方) vs 对方科目借方；差异落入「差异凭证汇总」
    recon = {}
    recon['222101'] = ('—（增值税专栏滚动，不单独核对）', 0.0, None)
    v103 = acc['222103']['inc']
    recon['222103'] = ('转出未交增值税(专栏/GL摘要结转)', zc, v103 - zc)
    s = sum(acc[c]['inc'] for c in SUR_CODES)
    recon['_sur'] = ('税金及附加(名称匹配)', tax_sur_deb, s - tax_sur_deb)
    ii = acc['222110']['inc']
    recon['222110'] = ('所得税费用-当期(名称匹配)', inc_cur, ii - inc_cur)
    recon['_yyszc'] = ('营业外支出(名称匹配)', yyszc, None)   # 信息列示，不参与税金及附加勾稽
    for c in WH_CODES:
        recon[c] = ('—（代扣代缴，不经损益）', 0.0, None)
    # 借方核对取数：A=应交税费借方合计(实缴, TB 借发 222103-222114)；B=银行存款缴税(GL)
    dec_tot = sum(acc[c]['dec'] for c in TARGET)          # 应交税费借方合计（实际缴纳税款，权威控制数）
    bank_tax_pay = gl.get('bank_tax_pay', {c: 0.0 for c in TARGET})
    return dict(acc=acc, tax_sur_deb=tax_sur_deb, inc_cur=inc_cur, zc=zc, yyszc=yyszc, bank_cre=bank_cre,
                vat_parent=vat_parent,
                inc_accr=accr, inc_exp=exp, inc_diff=inc_diff, inc_items=inc_items,
                vat_u103=u103, vat_zc=zc2, vat_diff=vat_diff, vat_items=vat_items,
                sur_items=sur_items, recon=recon,
                dec_tot=dec_tot, bank_tax_pay=bank_tax_pay)


# ================= 输出工作簿 =================
PERIOD = {
    '2025': '2025-01 至 2025-12（完整会计年度）',
    '2026': '2026-01 至 2026-05（综合查询明细表导出期间；科目余额表为较后期快照）',
}


def write_full_sheet(ws, y, entities, rec_by_year, issues):
    """应交税费明细表（模板克隆场景）。⚡ 2026-08-11 阶段二 2.1：计算产结构化行，
    渲染收敛到 audit_render.render_into（双行表头 R1/R2、数据 R3 起）。"""
    from audit_render import render_into, row as _arow
    # 清空模板残留行（模板外壳可能预置旧数据/合计行——如 FY控股 无数据主体被 skip 后残留
    # 旧"小计"数值 → 检查器误报；2026-08-02 修复）
    for _rr in range(2, (ws.max_row or 2) + 1):
        for _cc in range(1, (ws.max_column or 1) + 1):
            ws.cell(_rr, _cc).value = None
    # 双行表头分组（2026-08-11 规则2）
    groups = [('基础信息', 3), ('余额变动', 4), ('对方科目勾稽', 4), ('结论', 1)]
    heads = ['主体', '科目代码', '科目名称', '期初数', '本期增加(贷方计提)', '本期减少(借方支付)', '期末数',
             '税金及附加', '所得税费用-当期', '转出未交增值税',
             '代扣代缴(应付职工薪酬等)', '勾稽差异']
    MONEY = set(range(4, 13))
    _rows = []
    # 行顺序：未交增值税前插入 222101 应交增值税
    codes = []
    inserted = False
    for code in TARGET:
        if code == '222103' and not inserted:
            codes.append('222101'); inserted = True
        codes.append(code)
    if not inserted:
        codes.append('222101')
    g_qc = g_inc = g_dec = g_qm = 0.0
    g_h = g_i = g_j = g_k = g_df = 0.0
    sub_rows = []   # 各主体小计行号（集团合计公式引用）
    for ent in entities:
        e = rec_by_year[y].get(ent)
        if not e:
            continue
        full = e['full']
        vp = full['vat_parent']
        s_qc = s_inc = s_dec = s_qm = 0.0
        s_h = s_i = s_j = s_k = 0.0
        s_df = 0.0
        first = True
        for code in codes:
            if code == '222101':
                # 应交增值税(222101)为二级科目，金额=三级专栏轧差：直接取父级222101的借发/贷发/期初/期末
                # （不再由未交增值税推导，避免与「转出未交增值税」三级专栏混淆）
                qc = vp['qc']; qm_v = vp['qm']
                dec = vp['dec']      # 本期减少(借方) = 222101 借发(三级轧差)
                inc = vp['inc']      # 本期增加(贷方) = 222101 贷发(三级轧差)
                a = {'qc': qc, 'inc': inc, 'dec': dec, 'qm': qm_v, 'name': '应交增值税', 'exist': True}
                code_disp = '222101'; name_disp = '应交增值税'
                h_val = i_val = j_val = k_val = None; diff = 0.0
            else:
                a = full['acc'][code]
                code_disp = code; name_disp = a['name']
                if code == '222103':
                    # 未交增值税：对方科目 = 转出未交增值税(22210103借)
                    j_val = full['zc']; diff = a['inc'] - full['zc']
                    h_val = i_val = k_val = None
                elif code == '222110':
                    # 企业所得税：对方科目 = 所得税费用-当期(680101借)
                    i_val = full['inc_cur']; diff = a['inc'] - full['inc_cur']
                    h_val = j_val = k_val = None
                elif code in SUR_CODES:
                    # 税金及附加类：对方科目 = 税金及附加(6403借)，逐行填本行计提，差异在「小计」体现计提合计-6403借
                    h_val = a['inc']; diff = 0.0
                    i_val = j_val = k_val = None
                elif code in WH_CODES:
                    # 代扣代缴：不经损益，仅列示计提额
                    k_val = a['inc']; diff = 0.0
                    h_val = i_val = j_val = None
                else:
                    h_val = i_val = j_val = k_val = None; diff = 0.0
            # 0 值行不显示（与费用一致）：期初/增/减/末及勾稽列均为 0 则不输出
            if is_zero_amount(a['qc'], a['inc'], a['dec'], a['qm']) \
                    and (h_val or 0) == 0 and (i_val or 0) == 0 \
                    and (j_val or 0) == 0 and (k_val or 0) == 0:
                continue
            s_qc += a['qc']; s_inc += a['inc']; s_dec += a['dec']; s_qm += a['qm']
            g_qc += a['qc']; g_inc += a['inc']; g_dec += a['dec']; g_qm += a['qm']
            if h_val is not None:
                s_h += h_val; g_h += h_val
            if i_val is not None:
                s_i += i_val; g_i += i_val
            if j_val is not None:
                s_j += j_val; g_j += j_val
            if k_val is not None:
                s_k += k_val; g_k += k_val
            s_df += diff; g_df += diff
            # 2026-08-11 零值留空（规则4）：金额/勾稽列 0 → 空（行级已跳过全0行）
            def _zv(x):
                return None if x is None or (isinstance(x, (int, float)) and abs(x) < 0.005) else x
            _rows.append(_arow(
                [ent if first else '', code_disp, name_disp,
                 _zv(a['qc']), _zv(a['inc']), _zv(a['dec']), _zv(a['qm']),
                 _zv(h_val), _zv(i_val), _zv(j_val), _zv(k_val), _zv(diff)],
                num=MONEY, b=first, align='l',
                cell_fill=({12: DIFF_FILL} if (diff is not None and abs(diff) >= 0.01) else None),
                font=({12: RED_FONT} if (diff is not None and abs(diff) >= 0.01) else None)))
            first = False
        # 小计（2026-08-06 协议：写数值=s_* 累计；公式 data_only 读 None 收回读不回）
        s_df += (s_h - full['tax_sur_deb'])
        g_df += (s_h - full['tax_sur_deb'])
        _rows.append(_arow(
            ['%s 小计' % ent, None, None, round(s_qc, 2), round(s_inc, 2), round(s_dec, 2),
             round(s_qm, 2), round(s_h, 2), round(s_i, 2), round(s_j, 2), round(s_k, 2), s_df],
            b=True, fill='sub', num=MONEY, align='l',
            cell_fill=({12: DIFF_FILL} if abs(s_df) >= 0.01 else None),
            font=({12: RED_FONT} if abs(s_df) >= 0.01 else None)))
        sub_rows.append(len(_rows))
    if len(entities) > 1:  # 2026-08-10 单体裁剪
        _rows.append(_arow(
            ['集团合计', None, None, round(g_qc, 2), round(g_inc, 2), round(g_dec, 2),
             round(g_qm, 2), round(g_h, 2), round(g_i, 2), round(g_j, 2), round(g_k, 2), g_df],
            b=True, fill='total', num=MONEY, align='l'))
    # 清空模板残留旧列 + 渲染（render_into 全量覆盖）
    for c in range(13, (ws.max_column or 13) + 1):
        ws.cell(1, c).value = None
    render_into(ws, heads, _rows, MONEY, groups=groups, freeze='A5',
                title='应交税费明细表（%s年度，含应交增值税222101及应交税费各子目，含贷方计提对方科目勾稽）' % y)
    return ws

def write_vat_sheet(ws, y, entities, rec_by_year):
    ws.cell(1, 1, '亿利达集团 应交增值税明细表（分主体·分月·分核算明细） %s年度  单位：元' % y).font = Font(name='Times New Roman', bold=True, size=10)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=21)
    ws.cell(1, 1).alignment = CEN
    HEADS = ['主体', '月份', '期初留抵\n(借方)', '进项税额\n(22210101)', '已交税金\n(22210102)', '出口免抵税额\n(22210107)',
             '减免税额\n(22210108)', '进项加计抵扣\n(进项加计)', '借方发生额\n小计', '销项税额\n(22210104)', '出口退税\n(22210105)',
             '进项税额转出\n(22210106)', '简易计税\n(简易计税)', '贷方发生额\n小计', '本期应交/\n(留抵)', '转出未交增值税\n(计算应交)',
             '账载未交增值税\n(22210103借方)', '差异\n(计算-账载)', '期末留抵\n(借方)', '已交增值税\n(备查)', '备注']
    NCOL = len(HEADS)
    hr = 2
    for i, h in enumerate(HEADS, 1):
        c = ws.cell(hr, i, h); c.fill = HFILL; c.font = HFONT; c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[hr].height = 34
    r = hr + 1
    G = dict(op=0, d01=0, d02=0, d07=0, d08=0, d10=0, dr=0, c04=0, c05=0, c06=0, d09=0, cr=0, pay=0, d03=0, paid=0, end=0)
    ent_annual = []   # 收集每主体全年合计，供第二节"分主体不分月全年数据"
    # ===== 第一节：按主体分月计算应转出未交增值税 =====
    ws.cell(r, 1, '一、按主体分月计算应转出未交增值税（含各月转出未交增值税，逐月滚动）').font = SUB_FONT
    r += 1
    for ent in entities:
        e = rec_by_year[y].get(ent)
        if not e or not e['vat']['rows']:
            continue
        vat = e['vat']
        oerr = vat['oerr']
        s = dict(d01=0, d02=0, d07=0, d08=0, d10=0, dr=0, c04=0, c05=0, c06=0, d09=0, cr=0, pay=0, d03=0, paid=0)
        first = True
        for row in vat['rows']:
            m = row['m']
            note = ''
            if first and oerr:
                note = '期初222101为贷方%.2f，应为借方(留抵)，请核查' % vat['oamt']
            vals = [ent if first else '', '%d月' % m, row['opening'],
                    row['d01'], row['d02'], row['d07'], row['d08'], row['d10'], row['dr_move'],
                    row['c04'], row['c05'], row['c06'], row['d09'], row['cr_move'],
                    row['net'], row['payable'], row['d03_actual'], row['payable'] - row['d03_actual'],
                    row['ending'], row['paid'], note]
            for i, v in enumerate(vals, 1):
                cell = ws.cell(r, i, v); cell.border = BORDER
                if i == 1:
                    cell.alignment = LEFT
                elif i == 2:
                    cell.alignment = CEN
                elif i == NCOL:
                    cell.alignment = LEFT
                else:
                    cell.alignment = RGT
                    if isinstance(v, (int, float)):
                        cell.number_format = NUM
            if first and oerr:
                ws.cell(r, 3).fill = WARN_FILL; ws.cell(r, 3).font = RED_FONT
                ws.cell(r, NCOL).font = RED_FONT
            if abs(row['payable'] - row['d03_actual']) > 0.01:
                ws.cell(r, 18).fill = DIFF_FILL
            first = False
            for k, src in [('d01', 'd01'), ('d02', 'd02'), ('d07', 'd07'), ('d08', 'd08'), ('d10', 'd10'), ('dr', 'dr_move'),
                           ('c04', 'c04'), ('c05', 'c05'), ('c06', 'c06'), ('d09', 'd09'), ('cr', 'cr_move'),
                           ('pay', 'payable'), ('d03', 'd03_actual'), ('paid', 'paid')]:
                s[k] += row[src]
            r += 1
        # 2026-07-29 各主体月数据后加小计行
        sub_vals = [ent, '小计', vat['open_liudi'],
                    s['d01'], s['d02'], s['d07'], s['d08'], s['d10'], s['dr'],
                    s['c04'], s['c05'], s['c06'], s['d09'], s['cr'],
                    s['cr'] - s['dr'] - vat['open_liudi'], s['pay'], s['d03'],
                    s['pay'] - s['d03'], vat['ending'], s['paid'], '']
        for i, v in enumerate(sub_vals, 1):
            cell = ws.cell(r, i, v); cell.border = BORDER; cell.fill = TOT_FILL; cell.font = BFONT
            if i == 1:
                cell.alignment = LEFT
            elif i == 2:
                cell.alignment = CEN
            elif i == NCOL:
                cell.alignment = LEFT
            else:
                cell.alignment = RGT
                if isinstance(v, (int, float)):
                    cell.number_format = NUM
        r += 1
        ent_annual.append((ent, vat, s, oerr))
    # ===== 第二节：分主体不分月全年数据（各主体全年汇总，置于最后） =====
    r += 1
    ws.cell(r, 1, '二、分主体不分月全年数据（各主体全年汇总，与第一节分月数据一致；"月份"列以"全年"标识）').font = SUB_FONT
    r += 1
    for ent, vat, s, oerr in ent_annual:
        sub = [ent, '全年', vat['open_liudi'],
               s['d01'], s['d02'], s['d07'], s['d08'], s['d10'], s['dr'],
               s['c04'], s['c05'], s['c06'], s['d09'], s['cr'],
               s['cr'] - s['dr'] - vat['open_liudi'], s['pay'], s['d03'], s['pay'] - s['d03'],
               vat['ending'], s['paid'], ('' if not oerr else '期初贷方异常见上')]
        for i, v in enumerate(sub, 1):
            cell = ws.cell(r, i, v); cell.border = BORDER; cell.fill = SFILL; cell.font = BFONT
            if i == 1:
                cell.alignment = LEFT
            elif i == 2:
                cell.alignment = CEN
            elif i == NCOL:
                cell.alignment = LEFT
            else:
                cell.alignment = RGT
                if isinstance(v, (int, float)):
                    cell.number_format = NUM
        if abs(s['pay'] - s['d03']) > 0.01:
            ws.cell(r, 18).fill = DIFF_FILL
        G['op'] += vat['open_liudi']; G['end'] += vat['ending']
        for k in ['d01', 'd02', 'd07', 'd08', 'd10', 'dr', 'c04', 'c05', 'c06', 'd09', 'cr', 'pay', 'd03', 'paid']:
            G[k] += s[k]
        r += 1
    if len(entities) > 1:  # ⚡ 2026-08-10 单体裁剪
        grp = ['集团合计', '全年', G['op'],
               G['d01'], G['d02'], G['d07'], G['d08'], G['d10'], G['dr'],
               G['c04'], G['c05'], G['c06'], G['d09'], G['cr'],
               G['cr'] - G['dr'] - G['op'], G['pay'], G['d03'], G['pay'] - G['d03'],
               G['end'], G['paid'], '计算转出合计%.2f vs 账载%.2f' % (G['pay'], G['d03'])]
        for i, v in enumerate(grp, 1):
            cell = ws.cell(r, i, v); cell.border = BORDER; cell.fill = TFILL; cell.font = BFONT
            if i == 1:
                cell.alignment = LEFT
            elif i == 2:
                cell.alignment = CEN
            elif i == NCOL:
                cell.alignment = LEFT
            else:
                cell.alignment = RGT
                if isinstance(v, (int, float)):
                    cell.number_format = NUM
    widths = [13, 7, 15, 14, 13, 13, 12, 13, 14, 13, 13, 13, 13, 15, 14, 13, 13, 13, 34]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'C3'


def write_income_recon(ws, entities, rec_by_year, years):
    row = 1
    ws.cell(row, 1, '所得税计提列支核对（应交所得税222110贷方 vs 当期所得税费用680101借方）').font = Font(name='Times New Roman', bold=True, size=10); row += 1
    heads = ['主体', '应交所得税-计提(222110贷)', '当期所得税费用(680101借)', '差异', '差异原因（凭证穿透：字/号/摘要/对方科目/金额）']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row, i, h); c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    row += 1
    for y in years:
        ws.cell(row, 1, '%s年度' % y).font = BFONT; row += 1
        g1 = g2 = 0.0
        ent_rows = []   # 本年度主体数据行号（集团合计公式引用，2026-08-02）
        for ent in entities:
            e = rec_by_year[y].get(ent)
            if not e:
                continue
            full = e['full']
            accr = full['inc_accr']; exp = full['inc_exp']; diff = full['inc_diff']
            if abs(accr) < 0.01 and abs(exp) < 0.01:
                continue
            g1 += accr; g2 += exp
            if abs(diff) < 0.01:
                reason = '计提=费用，无差异'
            else:
                if full['inc_items']:
                    reason = ' ；'.join('%s%s / %s / %s / %s%.2f' % (zi, hao, zy, opp, '+' if v >= 0 else '', v)
                                       for zi, hao, zy, opp, v in full['inc_items'])
                else:
                    reason = '差异%.2f，未找到非标准计提凭证' % diff
            vals = [ent, accr, exp, diff, reason]
            for i, v in enumerate(vals, 1):
                cell = ws.cell(row, i, v); cell.border = BORDER
                if i in (2, 3, 4):
                    cell.number_format = NUM; cell.alignment = RGT
                if i == 5:
                    cell.alignment = LEFT
            if abs(diff) >= 0.01:
                for i in range(1, 6):
                    ws.cell(row, i).fill = DIFF_FILL
            ent_rows.append(row)
            row += 1
        # 集团合计（2026-08-02 公式化：=SUM 主体数据行；差异=B-C 公式）
        # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出
        if len(entities) > 1:
            ws.cell(row, 1, '%s 集团合计' % y).font = BFONT
            if ent_rows:
                _sc = lambda cl: ','.join('%s%d' % (cl, rr) for rr in ent_rows)
                ws.cell(row, 2, '=SUM(%s)' % _sc('B')); ws.cell(row, 3, '=SUM(%s)' % _sc('C'))
            else:
                ws.cell(row, 2, 0.0); ws.cell(row, 3, 0.0)
            ws.cell(row, 4, '=B%d-C%d' % (row, row))
            for i in range(1, 6):
                ws.cell(row, i).fill = TFILL; ws.cell(row, i).border = BORDER; ws.cell(row, i).font = BFONT
                if i in (2, 3, 4):
                    ws.cell(row, i).number_format = NUM; ws.cell(row, i).alignment = RGT
            row += 2
    ws.column_dimensions['A'].width = 13
    for c in ['B', 'C', 'D']:
        ws.column_dimensions[c].width = 20
    ws.column_dimensions['E'].width = 95
    for n in [
        '结论：所有主体当期正常计提分录（借 所得税费用-当期 / 贷 应交所得税）两侧完全一致；差异全部来自不经当期所得税费用的调节事项，属正常：',
        '① 汇算清缴退（补）税：借/贷 银行存款 对 应交所得税，冲减以前年度应交，不影响当期费用；',
        '② 年末审计调整冲回当期预缴所得税：借 应交所得税 / 贷 所得税费用（预缴后经测算无应纳税额全额冲回，当期费用净额归0，应交所得税贷方仍保留计提发生额）。']:
        ws.cell(row, 1, n).alignment = LEFT; row += 1


def write_vat_recon(ws, entities, rec_by_year, years):
    row = 1
    ws.cell(row, 1, '增值税转出核对（未交增值税222103贷方 = 转出未交增值税22210103借方）').font = Font(name='Times New Roman', bold=True, size=10); row += 1
    heads = ['主体', '未交增值税-贷方(222103)', '转出未交增值税-借方(22210103)', '差异', '差异来源（凭证穿透：字/号/摘要/对方科目/金额）']
    for i, h in enumerate(heads, 1):
        c = ws.cell(row, i, h); c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    row += 1
    for y in years:
        ws.cell(row, 1, '%s年度' % y).font = BFONT; row += 1
        g1 = g2 = 0.0
        ent_rows = []   # 本年度主体数据行号（集团合计公式引用，2026-08-02）
        for ent in entities:
            e = rec_by_year[y].get(ent)
            if not e:
                continue
            full = e['full']
            u103 = full['vat_u103']; zc = full['vat_zc']; diff = full['vat_diff']
            if abs(u103) < 0.01 and abs(zc) < 0.01:
                continue
            g1 += u103; g2 += zc
            if abs(diff) < 0.01:
                reason = '222103贷 = 22210103借，结转正确'
            else:
                if full['vat_items']:
                    reason = ' ；'.join('%s%s / %s / %s / %s%.2f' % (zi, hao, zy, opp, '+' if v >= 0 else '', v)
                                       for zi, hao, zy, opp, v in full['vat_items'])
                else:
                    reason = '差异%.2f，未找到直接计入未交增值税的凭证' % diff
            vals = [ent, u103, zc, diff, reason]
            for i, v in enumerate(vals, 1):
                cell = ws.cell(row, i, v); cell.border = BORDER
                if i in (2, 3, 4):
                    cell.number_format = NUM; cell.alignment = RGT
                if i == 5:
                    cell.alignment = LEFT
            if abs(diff) >= 0.01:
                for i in range(1, 6):
                    ws.cell(row, i).fill = DIFF_FILL
            ent_rows.append(row)
            row += 1
        # 集团合计（2026-08-02 公式化：=SUM 主体数据行；差异=B-C 公式）
        # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出
        if len(entities) > 1:
            ws.cell(row, 1, '%s 集团合计' % y).font = BFONT
            if ent_rows:
                _sc = lambda cl: ','.join('%s%d' % (cl, rr) for rr in ent_rows)
                ws.cell(row, 2, '=SUM(%s)' % _sc('B')); ws.cell(row, 3, '=SUM(%s)' % _sc('C'))
            else:
                ws.cell(row, 2, 0.0); ws.cell(row, 3, 0.0)
            ws.cell(row, 4, '=B%d-C%d' % (row, row))
            for i in range(1, 6):
                ws.cell(row, i).fill = TFILL; ws.cell(row, i).border = BORDER; ws.cell(row, i).font = BFONT
                if i in (2, 3, 4):
                    ws.cell(row, i).number_format = NUM; ws.cell(row, i).alignment = RGT
            row += 2
    ws.column_dimensions['A'].width = 13
    for c in ['B', 'C', 'D']:
        ws.column_dimensions[c].width = 22
    ws.column_dimensions['E'].width = 100
    for n in [
        '结论：多数主体「未交增值税(222103)贷方 ≡ 转出未交增值税(22210103)借方」完全相等，月末结转分录正确。',
        '出现差异的主体，差异全部来自直接计入未交增值税、未经"转出未交增值税"专栏的事项，属正常：',
        '① 统借统还利息增值税；② 视同销售/进项税额转出计提；③ 留抵退税返还；④ 重点人群税收减免；⑤ 年末应交增值税明细净额直接结平至未交增值税；⑥ 代收房租等价外费用增值税转出。']:
        ws.cell(row, 1, n).alignment = LEFT; row += 1


def write_diff_summary(ws, entities, rec_by_year, years):
    row = 1
    ws.cell(row, 1, '差异凭证汇总（应交税费贷方计提对方科目勾稽差异的凭证级穿透，按 主体/年度/税种 细化至具体税种与差异类型）').font = Font(name='Times New Roman', bold=True, size=10); row += 1
    heads = ['年度', '主体', '税种(细化)', '计提(贷方)', '对方科目借方', '差异', '凭证字', '凭证号', '摘要', '对方科目', '金额', '方向', '差异类型']
    # ⚡ 2026-08-11 N2：双行表头分组（规则2）——行2 大类（汇总对比/凭证明细），行3 列名
    for _g, _c0, _c1 in [('基础信息', 1, 3), ('汇总对比', 4, 6), ('凭证明细', 7, 12), ('差异类型', 13, 13)]:
        ws.merge_cells(start_row=row, start_column=_c0, end_row=row, end_column=_c1)
        _gc = ws.cell(row, _c0, _g)
        _gc.font = HFONT; _gc.fill = HFILL; _gc.alignment = CEN; _gc.border = BORDER
        for _j in range(_c0, _c1 + 1):
            ws.cell(row, _j).border = BORDER
    row += 1
    for i, h in enumerate(heads, 1):
        c = ws.cell(row, i, h); c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    row += 1

    def classify(zi, hao, zy, opp, v):
        t = '%s %s' % (zy or '', opp or '')
        if '汇算清缴' in t:
            return '汇算清缴退补税'
        if '留抵退税' in t or ('留抵' in t and '退' in t):
            return '留抵退税'
        if '统借统还' in t:
            return '统借统还利息'
        if '视同销售' in t:
            return '视同销售'
        if '减免' in t or '免税' in t:
            return '税收减免'
        if '进项转出' in t or '进项税额转出' in t:
            return '进项税转出'
        if '退税' in t:
            return '退税'
        if '补缴' in t or '补税' in t:
            return '补缴税款'
        return '其他(详见摘要)'

    any_diff = False
    for y in years:
        for ent in entities:
            e = rec_by_year[y].get(ent)
            if not e:
                continue
            full = e['full']; acc = full['acc']; recon = full['recon']
            cats = []
            # 未交增值税(222103)：其计提对方科目勾稽已在主表内嵌(H–L 列)，差异凭证汇总不再单列穿透
            # 企业所得税
            if abs(recon['222110'][2] or 0) >= 0.01:
                cats.append(('企业所得税(222110)', acc['222110']['inc'], recon['222110'][1],
                             recon['222110'][2], full['inc_items']))
            # 税金及附加：细化到各税种（逐行列示计提）+ 合计差异行（穿透6403借侧的明细）
            sr = recon.get('_sur')
            if sr and abs(sr[2] or 0) >= 0.01:
                for c in SUR_CODES:
                    if abs(acc[c]['inc']) >= 0.01:
                        cats.append((TARGET_NAME.get(c, c), acc[c]['inc'], None, None, None))
                cats.append(('税金及附加类(合计)', sum(acc[c]['inc'] for c in SUR_CODES),
                             sr[1], sr[2], full['sur_items']))
            for name, macc, opp_deb, diff, items in cats:
                any_diff = True
                ws.cell(row, 1, y); ws.cell(row, 2, ent); ws.cell(row, 3, name)
                ws.cell(row, 4, macc); ws.cell(row, 5, opp_deb if opp_deb is not None else '')
                ws.cell(row, 6, diff if diff is not None else '')
                for c in (4, 5, 6):
                    ws.cell(row, c).number_format = NUM; ws.cell(row, c).alignment = RGT
                for c in range(1, 7):
                    ws.cell(row, c).border = BORDER
                    if diff is not None:
                        ws.cell(row, c).fill = DIFF_FILL
                row += 1
                if items:
                    for zi, hao, zy, opp, v in items:
                        if abs(v) < 0.01:
                            continue
                        dc = classify(zi, hao, zy, opp, v)
                        ws.cell(row, 3, '  └ 凭证')
                        ws.cell(row, 4, ''); ws.cell(row, 5, ''); ws.cell(row, 6, abs(v))
                        ws.cell(row, 7, zi); ws.cell(row, 8, hao); ws.cell(row, 9, zy)
                        ws.cell(row, 10, opp); ws.cell(row, 11, abs(v)); ws.cell(row, 12, '借')  # 6403 借方
                        ws.cell(row, 13, dc)
                        ws.cell(row, 6).number_format = NUM; ws.cell(row, 6).alignment = RGT
                        ws.cell(row, 11).number_format = NUM; ws.cell(row, 11).alignment = RGT
                        for c in range(7, 14):
                            ws.cell(row, c).border = BORDER
                        ws.cell(row, 9).alignment = LEFT; ws.cell(row, 10).alignment = LEFT; ws.cell(row, 13).alignment = LEFT
                        row += 1
                elif diff is not None and abs(diff) >= 0.01:
                    ws.cell(row, 3, '  └ 说明')
                    ws.cell(row, 7, '差异 %.2f，未找到穿透凭证' % diff)
                    for c in range(7, 14):
                        ws.cell(row, c).border = BORDER
                    ws.cell(row, 7).alignment = LEFT
                    row += 1
    # 营业外支出（利润表，按名称匹配）：税收罚款/滞纳金等可能直接计入，供勾稽参考
    for y in years:
        for ent in entities:
            e = rec_by_year[y].get(ent)
            if not e:
                continue
            full = e['full']
            yv = full.get('yyszc') or 0.0
            if abs(yv) < 0.01:
                continue
            any_diff = True
            ws.cell(row, 1, y); ws.cell(row, 2, ent); ws.cell(row, 3, '营业外支出(名称匹配,信息)')
            ws.cell(row, 4, ''); ws.cell(row, 5, round(yv, 2)); ws.cell(row, 6, '')
            ws.cell(row, 5).number_format = NUM; ws.cell(row, 5).alignment = RGT
            for c in range(1, 7):
                ws.cell(row, c).border = BORDER
                ws.cell(row, c).fill = SFILL
            ws.cell(row, 7, '利润表营业外支出借方(按名称匹配)，税收罚款/滞纳金等可能计入，供核对参考')
            ws.cell(row, 7).alignment = LEFT
            for c in range(7, 14):
                ws.cell(row, c).border = BORDER
            row += 1
    if not any_diff:
        ws.cell(row, 1, '无贷方计提勾稽差异。').alignment = LEFT
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 8
    ws.column_dimensions['H'].width = 10
    ws.column_dimensions['I'].width = 40
    ws.column_dimensions['J'].width = 30
    ws.column_dimensions['K'].width = 16
    ws.column_dimensions['L'].width = 7
    ws.column_dimensions['M'].width = 18
    ws.freeze_panes = 'A3'




# ================= 应交税费审定表（参考『应交税费审定表.xlsx』格式，2026-08-01 用户需求） =================
# 参考格式（15 列）：
#   企业名称 | 项目 | 期末未审数 | 期初调整(借,贷) | 企业自行重分类(借,贷) | 重分类调整(借,贷) |
#   审定数(期初余额, 本期应交, 本期已交, 期末余额)
# 行 = 每主体一块（17 税种行 + 小计行），底部 = 按税种集团汇总（简单加计数）+ 合计。
# 税种顺序（参考表固定）与程序子目映射；缺失税种填 0；增值税 = 222101(vat_parent) + 222103(未交增值税)。
_TAX_AUDIT_ROWS = [  # (参考表税种名, 取数代码或 None=恒0, 是否需叠加)
    ('增值税', '222101+222103', True),
    ('城市维护建设税', '222105', False),
    ('企业所得税', '222110', False),
    ('印花税', '222104', False),
    ('消费税', None, False),
    ('水利基金', '222115', False),
    ('教育费附加', '222113', False),
    ('地方教育附加', '222114', False),
    ('应交车船税', None, False),
    ('车辆购置税', None, False),
    ('个人所得税', '222108', False),
    ('河道管理费', None, False),
    ('环境保护税', '222116', False),
    ('土地使用税', '222107', False),
    ('房产税', '222106', False),
    ('其他', '__rest__', False),
]


def _tax_audit_vals(full, code):
    """取某子目(或增值税合并) 的 (期初, 本期应交, 本期已交, 期末)。无数据返回 (0,0,0,0)。"""
    if code == '222101+222103':
        vp = full.get('vat_parent') or {}
        a = full.get('acc', {}).get('222103') or {}
        return (vp.get('qc', 0.0) + a.get('qc', 0.0),
                vp.get('inc', 0.0) + a.get('inc', 0.0),
                vp.get('dec', 0.0) + a.get('dec', 0.0),
                vp.get('qm', 0.0) + a.get('qm', 0.0))
    a = full.get('acc', {}).get(code)
    if not a:
        return (0.0, 0.0, 0.0, 0.0)
    return (a.get('qc', 0.0), a.get('inc', 0.0), a.get('dec', 0.0), a.get('qm', 0.0))


def write_tax_audit_sheet(ws, entities, y, rec_by_year):
    """应交税费审定表（2026-08-03 参照职工薪酬式 17 列）：
    序号 | 核算主体 | 项目 | 未审数(期初数/本期增加/本期减少/期末数) | 审计调整(借/贷) |
    重分类调整(借/贷) | 期初审计调整(借/贷) | 审定数(期初数/本期增加/本期减少/期末数)。
    每主体一块（15 税种行 + 小计），底部按税种集团简单加计 + 合计。数据取自 rec_by_year full 视图
    （与应交税费明细表同源，期初+本期增加-本期减少=期末 天然勾稽）；审计调整/重分类留 0 待填。"""
    N = 17
    ws.cell(1, 1, '应交税费审定表').font = Font(name='Times New Roman', bold=True, size=12)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N)
    ws.cell(1, 1).alignment = CEN
    # 两行表头（职工薪酬式）
    hdr1 = ['序号', '核算主体', '项  目', '未审数', '', '', '', '审计调整', '', '重分类调整', '',
            '期初审计调整', '', '审定数', '', '', '']
    hdr2 = ['', '', '', '期初数', '本期增加', '本期减少', '期末数',
            '借方', '贷方', '借方', '贷方', '借方', '贷方',
            '期初数', '本期增加', '本期减少', '期末数']
    for i, h in enumerate(hdr1, 1):
        c = ws.cell(3, i, h); c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    for i, h in enumerate(hdr2, 1):
        c = ws.cell(4, i, h); c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    for rng in ('D3:G3', 'H3:I3', 'J3:K3', 'L3:M3', 'N3:Q3'):
        ws.merge_cells(rng)
    r = 5
    grand = {t: [0.0, 0.0, 0.0, 0.0] for (t, _, _) in _TAX_AUDIT_ROWS}
    ent_tot = [0.0, 0.0, 0.0, 0.0]
    known = {c for (_, c, _) in _TAX_AUDIT_ROWS if c and c != '__rest__' and not isinstance(c, tuple)}
    seq = 0
    for ent in entities:
        e = rec_by_year.get(y, {}).get(ent)
        if not e or not e.get('full'):
            continue
        full = e['full']
        seq += 1
        ws.cell(r, 1, seq).font = BFONT; ws.cell(r, 1).alignment = CEN
        ws.cell(r, 2, ent).font = BFONT
        for c in range(1, N + 1):
            ws.cell(r, c).border = BORDER
        r += 1
        rest = [0.0, 0.0, 0.0, 0.0]
        for c, a in full.get('acc', {}).items():
            cs = str(c).replace('.', '')
            if cs.startswith('2221') and not any(cs == k or (k == '222101+222103' and cs in ('222101', '222103')) for k in known):
                rest[0] += a.get('qc', 0.0); rest[1] += a.get('inc', 0.0)
                rest[2] += a.get('dec', 0.0); rest[3] += a.get('qm', 0.0)
        for (tname, code, _flag) in _TAX_AUDIT_ROWS:
            if code == '__rest__':
                vals = rest
            elif code is None:
                vals = (0.0, 0.0, 0.0, 0.0)
            else:
                vals = _tax_audit_vals(full, code)
            qc_v, inc_v, dec_v, qm_v = vals
            ws.cell(r, 3, tname).alignment = LEFT
            for col, v in zip((4, 5, 6, 7), (qc_v, inc_v, dec_v, qm_v)):
                cell = ws.cell(r, col, round(v, 2)); cell.number_format = NUM
            for col in (8, 9, 10, 11, 12, 13):
                cell = ws.cell(r, col, 0); cell.number_format = NUM
            for col, v in zip((14, 15, 16, 17), (qc_v, inc_v, dec_v, qm_v)):
                cell = ws.cell(r, col, round(v, 2)); cell.number_format = NUM
            for c in range(1, N + 1):
                cell = ws.cell(r, c); cell.border = BORDER; cell.font = Font(name='Times New Roman', size=10)
                if c >= 4:
                    cell.alignment = RGT
            grand[tname][0] += qc_v; grand[tname][1] += inc_v
            grand[tname][2] += dec_v; grand[tname][3] += qm_v
            r += 1
        # 主体小计行
        s = [0.0, 0.0, 0.0, 0.0]
        for (tname, code, _flag) in _TAX_AUDIT_ROWS:
            if code == '__rest__':
                vv = rest
            elif code is None:
                vv = (0.0, 0.0, 0.0, 0.0)
            else:
                vv = _tax_audit_vals(full, code)
            for i in range(4):
                s[i] += vv[i]
        ws.cell(r, 3, '小  计').font = BFONT; ws.cell(r, 3).fill = TFILL
        for i, col in enumerate((4, 5, 6, 7, 14, 15, 16, 17)):
            cell = ws.cell(r, col, round(s[i % 4], 2)); cell.number_format = NUM
        for c in range(1, N + 1):
            cell = ws.cell(r, c); cell.border = BORDER; cell.font = BFONT; cell.fill = TFILL
            if c >= 4:
                cell.alignment = RGT
        for i in range(4):
            ent_tot[i] += s[i]
        r += 1
    # 底部：按税种集团简单加计
    for (tname, _code, _flag) in _TAX_AUDIT_ROWS:
        g = grand.get(tname, [0.0] * 4)
        ws.cell(r, 1, '集团加计').alignment = LEFT
        ws.cell(r, 3, tname).alignment = LEFT
        for i, col in enumerate((4, 5, 6, 7, 14, 15, 16, 17)):
            cell = ws.cell(r, col, round(g[i % 4], 2)); cell.number_format = NUM
        for c in range(1, N + 1):
            cell = ws.cell(r, c); cell.border = BORDER; cell.font = Font(name='Times New Roman', size=10)
            if c >= 4:
                cell.alignment = RGT
        r += 1
    # 全集团合计
    ws.cell(r, 3, '合 计').font = BFONT; ws.cell(r, 3).fill = TFILL
    for i, col in enumerate((4, 5, 6, 7, 14, 15, 16, 17)):
        cell = ws.cell(r, col, round(ent_tot[i % 4], 2)); cell.number_format = NUM
    for c in range(1, N + 1):
        cell = ws.cell(r, c); cell.border = BORDER; cell.font = BFONT; cell.fill = TFILL
        if c >= 4:
            cell.alignment = RGT
    for i, w in enumerate([6, 16, 20, 14, 14, 14, 14, 11, 11, 11, 11, 11, 11, 14, 14, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'D5'


def _group_tax_close(entities, rec_by_year, y):
    """集团汇总：各应交税费子目 期初(qc)/本期增加(inc)/本期减少(dec)/期末(qm)。
    全部取自 rec_by_year 的 full 视图（与《应交税费明细表》同源），保证 期初+增-减=末 勾稽一致。"""
    qc = defaultdict(float); qm = defaultdict(float)
    inc = defaultdict(float); dec = defaultdict(float)
    CODES = ['222101', '222103', '222104', '222105', '222106', '222107',
             '222108', '222110', '222111', '222113', '222114']
    for ent in entities:
        e = rec_by_year[y].get(ent)
        if not e:
            continue
        full = e.get('full')
        if not full:
            continue
        acc = full.get('acc', {})
        vp = full.get('vat_parent')
        for c in CODES:
            if c == '222101':
                if vp:
                    qc[c] += vp.get('qc', 0.0); qm[c] += vp.get('qm', 0.0)
                    inc[c] += vp.get('inc', 0.0); dec[c] += vp.get('dec', 0.0)
            else:
                a = acc.get(c)
                if a:
                    qc[c] += a.get('qc', 0.0); qm[c] += a.get('qm', 0.0)
                    inc[c] += a.get('inc', 0.0); dec[c] += a.get('dec', 0.0)
    return qc, qm, inc, dec


# ================= 借方核对（分年独立 sheet） =================
def write_debit_recon_sheet(ws, y, entities, rec_by_year):
    ws.cell(1, 1, '应交税费 借方核对（实际缴纳税款 = 银行存款缴税 + 非现金结算） %s年度  单位：元' % y).font = Font(name='Times New Roman', bold=True, size=10)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws.cell(1, 1).alignment = CEN
    heads = ['主体', '应交税费借方合计\n(实缴, TB借发·应交税费各子目)', '银行存款缴税\n(GL贷方·对方含税·TARGET口径)',
             '差异\n(实缴−银行缴税)', '差异率', '备注']
    hr = 2
    for i, h in enumerate(heads, 1):
        c = ws.cell(hr, i, h); c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    r = hr + 1
    g_a = g_b = g_diff = 0.0
    for ent in entities:
        e = rec_by_year[y].get(ent)
        if not e:
            continue
        full = e['full']
        a = full['dec_tot']                                  # 应交税费借方合计(实缴, TB)
        b = sum(full['bank_tax_pay'].get(c, 0.0) for c in TARGET)  # 银行存款缴税(GL, 对方含税, TARGET)
        diff = a - b
        rate = (diff / a * 100 if a else 0)
        note = '' if (a > 0 or b > 0) else '无税费缴纳及银行支出'
        vals = [ent, a, b, diff, rate, note]
        for i, val in enumerate(vals, 1):
            cell = ws.cell(r, i, val); cell.border = BORDER
            if i in (2, 3, 4):
                cell.number_format = NUM; cell.alignment = RGT
            if i == 5:
                cell.number_format = '0.00"%"'; cell.alignment = RGT
            if i == 6:
                cell.alignment = LEFT
        g_a += a; g_b += b; g_diff += diff; r += 1
    gv = ['集团合计', g_a, g_b, g_diff, (g_diff / g_a * 100 if g_a else 0),
          '差异 = 非现金结算(代扣offset/往来净额等) + 综合查询明细表非全量抽取未覆盖的银行缴税']
    for i, val in enumerate(gv, 1):
        cell = ws.cell(r, i, val); cell.border = BORDER; cell.fill = TFILL; cell.font = BFONT
        if i in (2, 3, 4):
            cell.number_format = NUM; cell.alignment = RGT
        if i == 5:
            cell.number_format = '0.00"%"'; cell.alignment = RGT
        if i == 6:
            cell.alignment = LEFT
    # ===== 差异明细：对应科目非银行存款的应交税费借方分录 =====
    r += 1
    ws.cell(r, 1, '差异明细（应交税费借方中"对应科目≠银行存款"的分录；即 应交税费借方合计 − 银行存款缴税 中对应科目非银行部分，'
            '用于定位非现金结算/代扣/往来抵消等差异来源）').font = SUB_FONT
    r += 1
    dh = ['主体', '凭证字', '凭证号', '日期', '摘要', '科目', '对方科目', '借方金额']
    for i, h in enumerate(dh, 1):
        c = ws.cell(r, i, h); c.font = HFONT; c.fill = HFILL; c.border = BORDER; c.alignment = CEN
    r += 1
    detail_total = 0.0
    for ent in entities:
        e = rec_by_year[y].get(ent)
        if not e:
            continue
        for x in e['gl'].get('rows', []):
            if not str(x.get('km', '')).startswith('应交税费'):
                continue
            if (x.get('jf') or 0) <= 0.01:
                continue
            if BANK_KW.search(str(x.get('opp') or '')):
                continue
            amt = x['jf']
            detail_total += amt
            vals = [ent, x.get('zi', ''), x.get('hao', ''), x.get('dt', ''), x.get('zy', ''),
                    x.get('km', ''), x.get('opp', ''), round(amt, 2)]
            for i, v in enumerate(vals, 1):
                c = ws.cell(r, i, v); c.border = BORDER
                if i == 8:
                    c.number_format = NUM; c.alignment = RGT
                elif i in (5, 6, 7):
                    c.alignment = LEFT
            r += 1
    ws.cell(r, 1, '非银行存款对应应交税费借方合计（不含 222101 已交税金之银行支付部分；该部分已含于上方差异但对应科目为银行）').font = BFONT
    ws.cell(r, 8, round(detail_total, 2)).number_format = NUM
    ws.cell(r, 8).font = BFONT; ws.cell(r, 8).fill = TFILL
    for c in range(1, 9):
        ws.cell(r, c).border = BORDER
    r += 2
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 24
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 42
    ws.freeze_panes = 'A3'


# ================= 税金及附加测算表（关联明细表计提数） =================
_SUR_TMPL = [
    ('城市维护建设税', '222105', 0.05), ('教育费附加', '222113', 0.03),
    ('地方教育附加', '222114', 0.02), ('资源税', None, None),
    ('矿产资源补偿费', None, None), ('房产税', '222106', None),
    ('土地使用税', '222107', None), ('车船税', None, None),
    ('印花税', '222104', None), ('合  计', '__SUM__', None),
]


def _copy_cell_style(src, dst):
    """安全复制单元格样式：openpyxl 读取模板后 cell.font/fill 等为 StyleProxy(不可哈希)，
    直接 'dst.font = cell.font' 会在加入目标工作簿样式表时抛 TypeError: unhashable type。
    故逐属性重建新对象，彻底规避 StyleProxy。"""
    try:
        f = src.font
        dst.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, u=f.u, strike=f.strike,
                        color=f.color, vertAlign=f.vertAlign, outline=f.outline, shadow=f.shadow)
    except Exception:
        pass
    try:
        fl = src.fill
        if fl is not None and getattr(fl, 'patternType', None):
            dst.fill = PatternFill(fill_type=fl.patternType, fgColor=fl.fgColor, bgColor=fl.bgColor)
    except Exception:
        pass
    try:
        a = src.alignment
        dst.alignment = Alignment(horizontal=a.horizontal, vertical=a.vertical,
                                  wrap_text=a.wrap_text, shrink_to_fit=a.shrink_to_fit,
                                  indent=a.indent, text_rotation=a.text_rotation)
    except Exception:
        pass
    try:
        b = src.border
        dst.border = Border(left=b.left, right=b.right, top=b.top, bottom=b.bottom)
    except Exception:
        pass


def write_surcharge_sheet(ws, y_target, entities, rec_by_year, tmpl_path=None):
    """税金及附加测算表——⚡ 2026-08-13 改为【内置结构】，不再依赖外部模板
    （原复制《税金及附加测算表*.xlsx》模板外壳，模板文件丢失→测算表/勾稽整体缺失；
    现直接用 _SUR_TMPL 生成外壳，数据逻辑不变）。
    账面计提(关联应交税费明细表各税种贷方) vs 应计提(计税依据×税率) → 测算差额；
    用户方法论：勾稽对上(差额≈0)不抽凭，绝大部分=计提数。"""
    t = ws.cell(1, 1, '税金及附加测算表（账面计提关联应交税费明细表各税种贷方计提；测算差额=应计提−账面计提；勾稽对上不抽凭）')
    t.font = Font(name='Times New Roman', bold=True, size=12); t.fill = TFILL; t.alignment = CEN
    ws.row_dimensions[1].height = 22
    # 表头（原模板表头行=5，列 C=计税依据/F=应计提数/G=账面计提/H=测算差额；A=税种/B=科目代码）
    hdr = ['税种', '科目代码', '计税依据金额', '', '', '应计提数', '账面计提', '测算差额']
    for j, h in enumerate(hdr, 1):
        c = ws.cell(5, j, h)
        c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 10
    for col in ('C', 'F', 'G', 'H'):
        ws.column_dimensions[col].width = 16
    # 税种行（内置 _SUR_TMPL）
    t_rows = {}
    r0 = 6
    for tname, code, rate in _SUR_TMPL:
        t_rows[tname] = r0
        ws.cell(r0, 1, tname)
        if code and code != '__SUM__':
            ws.cell(r0, 2, code)
        for j in range(1, 9):
            ws.cell(r0, j).border = BORDER
        r0 += 1
    # 增值税实缴基数(集团)：222103 借方 + 222101 已交税金(02)
    base = 0.0
    for ent in entities:
        e = rec_by_year[y_target].get(ent)
        if not e:
            continue
        full = e['full']
        base += full['acc']['222103']['dec']
        base += sum(e['gl']['vat']['02'][m]['d'] for m in range(1, 13))
    # 账面计提(集团)：各税种贷方计提(inc)，与明细表同源
    book_inc = defaultdict(float)
    for ent in entities:
        e = rec_by_year[y_target].get(ent)
        if not e:
            continue
        full = e['full']
        for code in ['222104', '222105', '222106', '222107', '222113', '222114']:
            book_inc[code] += full['acc'][code]['inc']
    # 填数列：应计提数(F=6)/账面计提(G=7)/测算差额(H=8)
    f_sum = 0.0; g_sum = 0.0
    for tname, code, rate in _SUR_TMPL:
        if tname == '合  计':
            ws.cell(t_rows.get(tname, 13), 6, round(f_sum, 2)).number_format = NUM
            ws.cell(t_rows.get(tname, 13), 7, round(g_sum, 2)).number_format = NUM
            ws.cell(t_rows.get(tname, 13), 8, round(f_sum - g_sum, 2)).number_format = NUM
            continue
        ridx = t_rows.get(tname)
        if ridx is None:
            continue
        if code in ('222104', '222105', '222106', '222107', '222113', '222114'):
            inc = book_inc.get(code, 0.0)
            ws.cell(ridx, 7, round(inc, 2)).number_format = NUM   # G 账面计提(关联明细表)
            g_sum += inc
            if rate is not None:
                f_val = base * rate
                ws.cell(ridx, 3, round(base, 2)).number_format = NUM   # C 计税依据金额
                ws.cell(ridx, 6, round(f_val, 2)).number_format = NUM  # F 应计提数
                ws.cell(ridx, 8, round(f_val - inc, 2)).number_format = NUM  # H 测算差额
                f_sum += f_val


def write_tax_footnote_sheet(ws, entities, y, rec_by_year):
    """应交税费附注汇总（应交税费审定表转置：每主体一列、行=税种、
    段=期初数/本期应交-贷方/本期已交-借方/期末数，审定口径；每段末 小计/合计/复核 行）。
    2026-08-02 用户方法论：附注汇总统一命名；内容按审定表标准行填列，非各主体自身科目。"""
    ent_list = [e for e in entities if rec_by_year.get(y, {}).get(e) and rec_by_year[y][e].get('full')]
    if not ent_list:
        return
    N = 2 + len(ent_list)
    # 2026-08-02：附注汇总不设合并单元格（便于后期添加/筛选）；标题仅 A1 加粗
    t = ws.cell(1, 1, '应交税费附注汇总（审定口径；每主体一列；行=税种；段=期初数/本期应交-贷方/本期已交-借方/期末数）')
    t.font = Font(name='Times New Roman', bold=True, size=12); t.fill = TFILL; t.alignment = CEN
    ws.row_dimensions[1].height = 22
    ws.cell(2, 1, '项目（审定数）').font = HFONT
    ws.cell(2, 1).fill = HFILL; ws.cell(2, 1).alignment = CEN; ws.cell(2, 1).border = BORDER
    for i, ent in enumerate(ent_list, 2):
        c = ws.cell(2, i, ent)
        c.font = HFONT; c.fill = HFILL; c.alignment = CEN; c.border = BORDER
    ws.column_dimensions['A'].width = 26
    for j in range(2, N + 1):
        ws.column_dimensions[get_column_letter(j)].width = 16

    def _ent_rest(full):
        rest = [0.0, 0.0, 0.0, 0.0]
        known = {c for (_, c, _) in _TAX_AUDIT_ROWS if c and c != '__rest__'}
        for c, a in full.get('acc', {}).items():
            cs = str(c).replace('.', '')
            if cs.startswith('2221') and not any(cs == k or (k == '222101+222103' and cs in ('222101', '222103')) for k in known):
                rest[0] += a.get('qc', 0.0); rest[1] += a.get('inc', 0.0)
                rest[2] += a.get('dec', 0.0); rest[3] += a.get('qm', 0.0)
        return rest

    def _ent_vals(ent):
        full = rec_by_year[y][ent]['full']
        rest = _ent_rest(full)
        out = {}
        for (tname, code, _flag) in _TAX_AUDIT_ROWS:
            if code == '__rest__':
                out[tname] = rest
            elif code is None:
                out[tname] = [0.0, 0.0, 0.0, 0.0]
            else:
                out[tname] = list(_tax_audit_vals(full, code))
        return out

    vals_by_ent = {e: _ent_vals(e) for e in ent_list}
    r = 3
    seg_gc = {}   # 段 -> 集团合计数行号（2026-08-03 滚动勾稽行引用）
    for seg_label, idx in [('期初数', 0), ('本期应交-贷方', 1), ('本期已交-借方', 2), ('期末数', 3)]:
        # 段头：A 列加粗提示（不合并跨行）
        c = ws.cell(r, 1, seg_label)
        c.font = HFONT; c.fill = TFILL; c.alignment = CEN
        r += 1
        data_first = r   # 本段税种数据首行（小计公式引用，2026-08-02）
        for (tname, _c, _f) in _TAX_AUDIT_ROWS:
            cell = ws.cell(r, 1, tname)
            cell.font = Font(name='Times New Roman', size=10); cell.alignment = LEFT; cell.border = BORDER
            for i, ent in enumerate(ent_list, 2):
                v = vals_by_ent[ent][tname][idx]
                cc = ws.cell(r, i, round(v, 2) if abs(v) >= 0.005 else 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = Font(name='Times New Roman', size=10)
            r += 1
        # 小计行（2026-08-06 协议：写数值=数据区各列和；公式 data_only 读 None 收回读不回）
        ws.cell(r, 1, '小  计').font = BFONT; ws.cell(r, 1).fill = TFILL
        ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
        for i, ent in enumerate(ent_list, 2):
            _sv = sum(ws.cell(x, i).value for x in range(data_first, r)
                      if isinstance(ws.cell(x, i).value, (int, float)))
            cc = ws.cell(r, i, round(_sv, 2))
            cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
            cc.fill = TFILL; cc.font = Font(name='Times New Roman', size=10, bold=True)
        r += 1
        # 集团合计数（=小计行值，写数值；2026-08-02 用户方法论：加合并口径 4 行）
        # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出合并 4 行
        rc = r
        seg_gc[seg_label] = rc
        if len(entities) > 1:
            ws.cell(r, 1, '集团合计数').font = BFONT; ws.cell(r, 1).fill = TFILL
            ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
            for i, ent in enumerate(ent_list, 2):
                _v = ws.cell(r - 1, i).value if isinstance(ws.cell(r - 1, i).value, (int, float)) else 0.0
                cc = ws.cell(r, i, round(_v, 2))
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.fill = TFILL; cc.font = Font(name='Times New Roman', size=10, bold=True)
            r += 1
            # 合并抵消借方（待填 0）
            ws.cell(r, 1, '合并抵消借方').font = Font(name='Times New Roman', size=10)
            ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
            for i in range(2, N + 1):
                cc = ws.cell(r, i, 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = Font(name='Times New Roman', size=10)
            r += 1
            # 合并抵消贷方（待填 0）
            ws.cell(r, 1, '合并抵消贷方').font = Font(name='Times New Roman', size=10)
            ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
            for i in range(2, N + 1):
                cc = ws.cell(r, i, 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = Font(name='Times New Roman', size=10)
            r += 1
            # 合并报表数（= 集团合计数 + 抵消借 − 抵消贷 = 集团合计，写数值）
            ws.cell(r, 1, '合并报表数').font = BFONT; ws.cell(r, 1).fill = TFILL
            ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
            for i, ent in enumerate(ent_list, 2):
                _v = ws.cell(rc, i).value if isinstance(ws.cell(rc, i).value, (int, float)) else 0.0
                cc = ws.cell(r, i, round(_v, 2))
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.fill = TFILL; cc.font = Font(name='Times New Roman', size=10, bold=True)
            r += 1
        # —— 2026-08-03 用户方法论：附注须考虑 审计调整数/审定数（披露增减变动科目区分借/贷调整）——
        # 审定数 = 集团合计数 + 调整借 − 调整贷（调整 0 待填；2026-08-06 协议：写数值=集团合计）
        ws.cell(r, 1, '审计调整-借方').font = Font(name='Times New Roman', size=10)
        ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
        for i in range(2, N + 1):
            cc = ws.cell(r, i, 0.0)
            cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
            cc.font = Font(name='Times New Roman', size=10)
        r += 1
        ws.cell(r, 1, '审计调整-贷方').font = Font(name='Times New Roman', size=10)
        ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
        for i in range(2, N + 1):
            cc = ws.cell(r, i, 0.0)
            cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
            cc.font = Font(name='Times New Roman', size=10)
        r += 1
        ws.cell(r, 1, '审定数').font = BFONT; ws.cell(r, 1).fill = TFILL
        ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
        for i, ent in enumerate(ent_list, 2):
            _v = ws.cell(rc, i).value if isinstance(ws.cell(rc, i).value, (int, float)) else 0.0
            cc = ws.cell(r, i, round(_v, 2))
            cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
            cc.fill = TFILL; cc.font = Font(name='Times New Roman', size=10, bold=True)
        r += 1
    # 滚动勾稽行（2026-08-06 协议：写数值=各段集团合计行值滚动；公式 data_only 读 None）
    if len(seg_gc) == 4:
        r0 = seg_gc['期初数']; r1 = seg_gc['本期应交-贷方']
        r2 = seg_gc['本期已交-借方']; r3 = seg_gc['期末数']
        ws.cell(r, 1, '滚动勾稽（期初+本期应交-本期已交-期末）').font = BFONT
        ws.cell(r, 1).fill = TFILL; ws.cell(r, 1).alignment = LEFT; ws.cell(r, 1).border = BORDER
        for i, ent in enumerate(ent_list, 2):
            _v0 = ws.cell(r0, i).value if isinstance(ws.cell(r0, i).value, (int, float)) else 0.0
            _v1 = ws.cell(r1, i).value if isinstance(ws.cell(r1, i).value, (int, float)) else 0.0
            _v2 = ws.cell(r2, i).value if isinstance(ws.cell(r2, i).value, (int, float)) else 0.0
            _v3 = ws.cell(r3, i).value if isinstance(ws.cell(r3, i).value, (int, float)) else 0.0
            _exp = _v0 + _v1 - _v2 - _v3
            cc = ws.cell(r, i, round(_exp, 2))
            cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
            cc.font = Font(name='Times New Roman', size=10,
                           color='FF0000' if abs(_exp) > 0.01 else '000000', bold=True)
        r += 1
    ws.freeze_panes = 'B3'


def _find_tax_tmpl(data_dir, pat):
    import fnmatch
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith('~$'):
            continue
        if fnmatch.fnmatch(fn, pat):
            return os.path.join(data_dir, fn)
    return None


# ================= 结构外壳（仅建固定 sheet，数据/表头由 write_* 填充） =================
def _build_tax_frame(wb, years):
    """创建应交税费工作簿的静态 sheet 结构：仅按固定命名建空 sheet。
    注：「税金及附加测算表」依赖数据文件夹内的模板（动态），不在此预建；
         其余固定 sheet 全部在此生成，作为可外部编辑的「外壳」持久化到 audit_templates/。"""
    for y in years:
        wb.create_sheet('应交税费审定表_%s' % y)
    for y in years:
        wb.create_sheet('附注汇总_%s' % y)
    for y in years:
        wb.create_sheet('应交税费明细表_%s' % y)
    for y in years:
        wb.create_sheet('应交增值税明细表_%s' % y)
    for y in years:
        wb.create_sheet('借方核对_%s' % y)
    wb.create_sheet('差异凭证汇总')


# ================= 未交增值税借方凭证抽查 =================
def write_vat_payment_voucher_sheet(ws, entities, rec_by_year, years):
    """其他应交税费借方凭证抽查表：从 GL 提取所有非应交增值税二级科目（222103~222114）
    的借方发生额，列示每笔凭证详情。排除应交增值税(222101)及其三级专栏。"""
    rows_data = []
    for y in sorted(years):
        for ent in entities:
            yr = rec_by_year.get(y, {}).get(ent)
            if not yr or not yr.get('gl'):
                continue
            for r in yr['gl'].get('rows', []):
                km = r.get('km', '')
                jf = r.get('jf', 0)
                if jf <= 0:
                    continue
                # 排除应交增值税(222101)及其三级专栏（如 应交税费-应交增值税-进项税额）
                if km.startswith('应交税费-应交增值税') and km != '应交税费-未交增值税':
                    continue
                # 检查是否为 TARGET 二级科目（包含税种名称即匹配，兼容 应交税费-{name} 或 应交税费-应交{name} 格式）
                if not any(tn in km for tn in TARGET_NAME.values()):
                    continue
                rows_data.append((y, ent, r['dt'], r['zi'], r['hao'], r['opp'], r['zy'], jf, km))
    if not rows_data:
        ws.cell(1, 1, '其他应交税费借方凭证抽查表（无需要抽查的凭证）').font = Font(name='Times New Roman', bold=True, size=10)
        return
    rows_data.sort(key=lambda x: (x[1], x[2] or ''))
    headers = ['核算主体', '测试序号', '日期', '凭证字', '凭证号',
               '二级科目（或客户/供应商名称）', '摘要', '借方金额', '贷方金额',
               '对方科目', '与原始凭证相符', '原始凭证内容', '原始凭证日期',
               '会计处理正确', '所属时间无误']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = HFONT; cell.fill = HFILL; cell.alignment = CEN; cell.border = BORDER
    tot = 0.0
    for i, row in enumerate(rows_data, 1):
        y, ent, dt, zi, hao, opp, zy, jf, km = row
        ws.cell(i + 1, 1, ent).border = BORDER          # 核算主体
        ws.cell(i + 1, 2, i).border = BORDER            # 测试序号
        ws.cell(i + 1, 3, dt).border = BORDER           # 日期
        ws.cell(i + 1, 4, zi).border = BORDER           # 凭证字
        ws.cell(i + 1, 5, hao).border = BORDER          # 凭证号
        ws.cell(i + 1, 6, km).border = BORDER           # 二级科目
        _c = ws.cell(i + 1, 7, zy)                      # 摘要
        _c.border = BORDER; _c.alignment = Alignment(wrap_text=True)
        _c = ws.cell(i + 1, 8, round(jf, 2))            # 借方金额
        _c.border = BORDER; _c.number_format = '#,##0.00'; _c.alignment = RGT
        ws.cell(i + 1, 9, '').border = BORDER           # 贷方金额
        ws.cell(i + 1, 10, opp).border = BORDER         # 对方科目
        for col in range(11, 16):                       # cols 11-15: empty
            ws.cell(i + 1, col, '').border = BORDER
        tot += jf
    r = len(rows_data) + 2
    ws.cell(r, 1, '合计').font = BFONT
    # 2026-08-02 公式化：借方金额 =SUM(数据区 2..r-1)
    ws.cell(r, 8, '=SUM(H2:H%d)' % (r - 1) if r > 2 else 0.0).font = BFONT
    ws.cell(r, 8).number_format = '#,##0.00'
    for c in range(1, 16):
        ws.cell(r, c).fill = TFILL; ws.cell(r, c).border = BORDER
    ws.cell(r, 1).alignment = CEN
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 26
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['H'].width = 14
    ws.column_dimensions['I'].width = 14
    ws.column_dimensions['J'].width = 24
    ws.column_dimensions['K'].width = 12
    ws.column_dimensions['L'].width = 16
    ws.column_dimensions['M'].width = 14
    ws.column_dimensions['N'].width = 12
    ws.column_dimensions['O'].width = 12


# ================= 组装 =================
def build_tax_combined(data_dir, out_dir=None, period_mode='Y'):
    entities_map = _discover_tax_entities(data_dir)
    if not entities_map:
        print('目录内未找到含「科目余额表」的账套文件：%s' % data_dir)
        return None
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    all_years = set()
    for b in entities_map.values():
        all_years.update(b.keys())
    years = sorted(all_years)
    entities = sorted(entities_map.keys())
    rec_by_year = {}
    issues = []
    for y in years:
        rec_by_year[y] = {}
        for ent, b in entities_map.items():
            yr = b.get(y)
            if not yr or not yr.get('km'):
                continue
            try:
                gl = read_gl_tax(yr.get('gl'))
                tb = load_tb(yr['km'])
                vat = build_vat_entity(gl, tb)
                full = build_full_entity(gl, tb)
            except Exception as ex:
                issues.append({'level': 'ERROR', 'type': '读取失败',
                               'message': '%s[%s] 读取 %s 失败：%s' % (ent, y, yr.get('km'), ex)})
                continue
            rec_by_year[y][ent] = {'vat': vat, 'full': full, 'gl': gl, 'tb': tb}
            if vat['oerr']:
                issues.append({'level': 'WARN', 'type': '增值税期初',
                               'message': '%s[%s] 222101 期初为贷方 %.2f（应为借方留抵），违反规则。' % (ent, y, vat['oamt'])})
            if not vat['rows']:
                issues.append({'level': 'INFO', 'type': '增值税',
                               'message': '%s[%s] 无 222101 应交增值税明细数据，不纳入增值税明细。' % (ent, y)})
    # 外壳缺失时先建一次并落盘（便于今后改排版只动外壳）
    if not resolve_template('应交税费_外壳', 'TAX_TEMPLATE'):
        try:
            _wb0 = openpyxl.Workbook(); _wb0.remove(_wb0.active)
            _build_tax_frame(_wb0, years)
            finalize_workbook(_wb0)
            os.makedirs(AUDIT_TEMPLATES_DIR, exist_ok=True)
            _wb0.save(os.path.join(AUDIT_TEMPLATES_DIR, '应交税费_外壳.xlsx'))
            _wb0.close()
        except Exception as ex:
            print('  ⚠️ 外壳落盘失败（不影响本次生成）：%s' % ex)
    from audit_shell import finalize_workbook as _fw
    outs = []
    # ===== 直接按年构建（2026-08-01 改造：消除合并稿中间态，替代 emit_per_year 分拆）=====
    for y in years:
        ys = str(y)
        shell = resolve_template('应交税费_外壳', 'TAX_TEMPLATE')
        if shell:
            wb = openpyxl.load_workbook(shell)
            def get_sheet(name):
                return wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        else:
            wb = openpyxl.Workbook(); wb.remove(wb.active)
            _build_tax_frame(wb, [y])
            def get_sheet(name):
                # ⚡ 2026-08-13：改为创建式（税金及附加测算表不再依赖外部模板预建，自动创建）
                return wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        # 剔除外壳残留的非当年分年 sheet
        _YEAR_SHEET_PREFIXES = ('应交税费审定表_', '应交税费明细表_', '应交增值税明细表_', '借方核对_')
        for _s in list(wb.sheetnames):
            for _p in _YEAR_SHEET_PREFIXES:
                if _s.startswith(_p):
                    _m = re.search(r'(20\d{2})$', _s)
                    if _m and _m.group(1) != ys:
                        del wb[_s]
                    break
        if '编制说明' in wb.sheetnames:
            del wb['编制说明']
        if '应交税费审定表' in wb.sheetnames and not any(s.startswith('应交税费审定表_') for s in wb.sheetnames):
            del wb['应交税费审定表']
        _OLD_PREFIXES = ('应交税费审定表', '应交税费明细表')
        for _s in list(wb.sheetnames):
            for _op in _OLD_PREFIXES:
                if _s.startswith(_op) and not _s.startswith(_op + '_'):
                    del wb[_s]
                    break
        # 审定表/附注汇总/明细表/应交增值税（当年）
        write_tax_audit_sheet(get_sheet('应交税费审定表_%s' % ys), entities, y, rec_by_year)
        # ⚡ 2026-08-11 用户需求（参照职工薪酬）：集团审定表（每单位一行）+ 横展审定表（税种×主体）
        try:
            build_tax_extra_sheets(wb, data_dir, rec_by_year, target_year=ys)
        except Exception as _ex:
            print(f'  ⚠️ 应交税费集团/横展审定表生成失败：{_ex}')
        # ⚡ 2026-08-11 用户需求：多期对比表并入底稿（仅最新年 wb）
        if str(ys) == str(sorted(rec_by_year.keys())[-1]) and len(rec_by_year) >= 2:
            try:
                _render_tax_multi_period_sheet(wb, data_dir, rec_by_year)
            except Exception as _ex:
                print(f'  ⚠️ 应交税费多期对比表并入底稿失败：{_ex}')
        write_tax_footnote_sheet(get_sheet('附注汇总_%s' % ys), entities, y, rec_by_year)
        write_full_sheet(get_sheet('应交税费明细表_%s' % ys), y, entities, rec_by_year, issues)
        write_vat_sheet(get_sheet('应交增值税明细表_%s' % ys), y, entities, rec_by_year)
        # 税金及附加测算表（⚡ 2026-08-13 内置结构，不再依赖外部模板——模板丢失曾致测算表/勾稽缺失）
        ws_sur = get_sheet('税金及附加测算表')
        write_surcharge_sheet(ws_sur, ys, entities, rec_by_year)
        # 差异凭证汇总（当年）
        write_diff_summary(get_sheet('差异凭证汇总'), entities, rec_by_year, [y])
        # 其他应交税费借方凭证抽查表（当年）
        write_vat_payment_voucher_sheet(get_sheet('其他应交税费借方凭证抽查'), entities, rec_by_year, [y])
        # 删除借方核对表
        for _s in list(wb.sheetnames):
            if _s.startswith('借方核对_'):
                del wb[_s]
        # 排序：审定表→明细表→差异凭证汇总→应交增值税明细表→其余→其他应交税费借方凭证抽查（最末）
        _order = ['应交税费审定表_%s' % ys, '应交税费明细表_%s' % ys, '差异凭证汇总', '应交增值税明细表_%s' % ys]
        _ordered = []
        for name in _order:
            for s in wb.worksheets:
                if s.title == name:
                    _ordered.append(s)
                    break
        _rest = [s for s in wb.worksheets if s.title not in _order]
        _vatpay = [s for s in _rest if s.title == '其他应交税费借方凭证抽查']
        _other = [s for s in _rest if s.title != '其他应交税费借方凭证抽查']
        wb._sheets = _ordered + _other + _vatpay
        _fw(wb)
        # ⚡ 2026-08-28 #875：对方科目核对（模块内集成）
        try:
            from counterparty_recon import inject_into_wb_auto
            inject_into_wb_auto(wb, data_dir, ys, '应交税费', ents_set=set(entities_map))
        except Exception as _ex:
            print(f'  ⚠️ 应交税费对方科目核对注入失败：{_ex}')
        out = os.path.join(out_dir, '应交税费审计底稿_%s_生成.xlsx' % ys)
        from audit_common import validate_workbook
        validate_workbook(wb, '应交税费底稿', raise_on_error=False)
        wb.save(out)
        wb.close()
        outs.append(out)
        print('  ✓ 应交税费 %s 已生成' % ys)
    out_path = outs[0] if outs else None
    warns = [i for i in issues if i['level'] == 'WARN']
    errs = [i for i in issues if i['level'] == 'ERROR']
    print('已生成：%s' % out_path)
    print('   核算主体 %d 个 | 年度 %s | ERROR %d | WARN %d' % (len(entities), ','.join(years), len(errs), len(warns)))
    # ⚡ 2026-08-11 用户需求：应交税费多期对比表（独立跨年文件）
    try:
        if len(rec_by_year) >= 2:
            build_tax_multi_period_table(data_dir, out_dir)
    except Exception as _ex:
        print(f'  ⚠️ 应交税费多期对比表生成失败：{_ex}')
    return out_path


def main(argv=None):
    raw = list(sys.argv[1:] if argv is None else argv)
    # 拖入文件夹快捷方式：若首参为文件夹路径（不以 '-' 开头），自动视为 --input
    if raw and not raw[0].startswith('-') and os.path.isdir(raw[0]):
        raw = ['--input', raw[0]] + raw[1:]
    p = argparse.ArgumentParser(description='应交税费明细表生成程序（多账套×多年份，参照职工薪酬小程序）')
    p.add_argument('--input', '-i', required=False, help='含科目余额表/综合查询明细表的文件夹')
    p.add_argument('--output', '-o', default=None, help='输出目录（默认与输入同目录）')
    p.add_argument('--period', choices=['Y', 'M', 'Q'], default='Y', help='期间粒度（当前仅 Y 年维度输出）')
    args = p.parse_args(raw)
    if not args.input or not os.path.isdir(args.input):
        print('❌ 用法：拖入【账套文件夹】到本程序，或指定 --input <文件夹>')
        return 1
    res = build_tax_combined(args.input, args.output)
    from audit_common import finalize_after_build
    finalize_after_build(args.input)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）
    return 0 if res else 1


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
    import sap_common as C            # 2026-08-10：流式 GL 读取（_gl_files/_safx）
    import _safx
except Exception:
    _adapter = None
    C = None
    _safx = None

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
                                       f'tax_crash_{ts}.log'), 'w', encoding='utf-8') as _f:
                    _f.write('未捕获异常：\n' + traceback.format_exc())
                print(f'已记录崩溃日志：tax_crash_{ts}.log')
        except Exception:
            pass
        rc = 1
    try:
        input('\n按回车退出…')
    except EOFError:
        pass
    sys.exit(rc)

# ============================================================================
# 应交税费三张新表（2026-08-11 用户：参照职工薪酬）
# ①集团审定表（每单位一行+审计调整/审定数）②横展审定表（税种×主体）
# ③多期对比表（税种×主体，期间×(期初/应交/已交/期末)，并入底稿最新年+独立文件）
# ============================================================================

def _tax_std_vals(full):
    """full → {税种名: [期初,本期应交,本期已交,期末]}（含『小  计』=Σ15 税种）。"""
    out = {}
    tot = [0.0, 0.0, 0.0, 0.0]
    for tname, code, _flag in _TAX_AUDIT_ROWS:
        v = _tax_audit_vals(full, code)
        out[tname] = list(v)
        for i in range(4):
            tot[i] += v[i]
    out['小  计'] = tot
    return out


def build_tax_extra_sheets(wb, data_dir, rec_by_year, target_year=None):
    """参照职工薪酬：①『应交税费集团审定表』（每核算单位一行：期初/本期增加(应交)/本期减少(已交)/
    期末/审计调整数/审定数）；②『应交税费横展审定表』（行=税种+小计、列=各核算主体期末+合计）。
    集团模式（多主体）生成；单主体跳过。"""
    from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_BOLD,
                             SHELL_NUM, SHELL_LEFT, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                             SHELL_TOT_FILL)
    from openpyxl.utils import get_column_letter
    if not rec_by_year:
        return
    years = sorted(rec_by_year.keys())
    ys = str(target_year) if target_year is not None else str(years[-1])
    # ⚡ 跨年账套（DQ：2024=dq/2025=DQ）：实体取跨年并集（职工薪酬同口径），
    #   当年无数据的实体在表中留空行
    ents = sorted({e for y in rec_by_year for e in rec_by_year[y].keys()})
    if len(ents) <= 1:
        return
    ent_std = {}
    for e in ents:
        rec = rec_by_year[ys].get(e)
        if rec:
            ent_std[e] = _tax_std_vals(rec['full'])
    if not ent_std:
        return
    # ===== ① 集团审定表：每单位一行 =====
    ws = wb.create_sheet(f'应交税费集团审定表_{ys}')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    c = ws.cell(1, 1, f'应交税费集团审定表（{ys} 年度）· 每核算单位一行 · 账面数（未审）+ 审计调整/审定数')
    c.font = SHELL_TITLE_FONT
    hdr = ['序号', '核算主体', '期初数', '本期增加(应交)', '本期减少(已交)', '期末数', '审计调整数', '审定数']
    for j, h in enumerate(hdr, 1):
        cc = ws.cell(2, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    r = 3
    g = [0.0] * 4
    for i, e in enumerate(ents, 1):
        if e not in ent_std:
            continue
        v = ent_std[e]['小  计']
        row_vals = [i, e, round(v[0], 2), round(v[1], 2), round(v[2], 2), round(v[3], 2), 0.0, round(v[3], 2)]
        for j, vv in enumerate(row_vals, 1):
            cc = ws.cell(r, j, vv)
            cc.border = SHELL_BORDER
            if j >= 4:
                cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
            elif j == 1:
                cc.alignment = SHELL_CEN
        g = [g[k] + v[k] for k in range(4)]
        r += 1
    row_vals = ['合计', None, round(g[0], 2), round(g[1], 2), round(g[2], 2), round(g[3], 2), 0.0, round(g[3], 2)]
    for j, vv in enumerate(row_vals, 1):
        cc = ws.cell(r, j, vv if j > 1 else None)
        if j == 1:
            cc = ws.cell(r, 1, '合计')
        cc.border = SHELL_BORDER
        cc.font = SHELL_BOLD
        cc.fill = SHELL_TOT_FILL
        if j >= 4:
            cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
        elif j == 1:
            cc.alignment = SHELL_CEN
    for i, w in enumerate([6, 16, 16, 18, 18, 16, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A3'
    # ===== ② 横展审定表：税种×主体 =====
    ws2 = wb.create_sheet(f'应交税费横展审定表_{ys}')
    nc = 1 + len(ents) + 1
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws2.cell(1, 1, f'应交税费横展审定表（{ys} 年度）· 行=税种、列=各核算单位期末+合计（横向对比各单位税负）')
    c.font = SHELL_TITLE_FONT
    hdr2 = ['税种'] + ents + ['合计']
    for j, h in enumerate(hdr2, 1):
        cc = ws2.cell(2, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    rows2 = [(t, t) for t, _c, _f in _TAX_AUDIT_ROWS] + [('小  计', '小  计')]
    r = 3
    for nm, _k in rows2:
        vals_by_e = {e: (ent_std[e].get(nm, [0.0] * 4)[3] if e in ent_std else None) for e in ents}
        tot = sum(v for v in vals_by_e.values() if isinstance(v, (int, float)))
        is_tot = (nm == '小  计')
        ws2.cell(r, 1, nm)
        for j, e in enumerate(ents, 2):
            v = vals_by_e[e]
            cc = ws2.cell(r, j, round(v, 2) if isinstance(v, (int, float)) and abs(v) >= 0.005 else None)
            cc.border = SHELL_BORDER
            cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
        cc = ws2.cell(r, nc, round(tot, 2) if abs(tot) >= 0.005 else None)
        cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
        cc.font = SHELL_BOLD
        for j in range(1, nc + 1):
            ws2.cell(r, j).border = SHELL_BORDER
            if is_tot:
                ws2.cell(r, j).font = SHELL_BOLD
                ws2.cell(r, j).fill = SHELL_TOT_FILL
        r += 1
    ws2.column_dimensions['A'].width = 22
    for j in range(2, nc + 1):
        ws2.column_dimensions[get_column_letter(j)].width = 13
    ws2.freeze_panes = 'B3'
    return ws, ws2


def _render_tax_multi_period_sheet(wb, data_dir, rec_by_year):
    """应交税费多期对比表渲染到已建 wb（并入职工薪酬式底稿最新年）。
    行=【税种块 × 各核算主体】、列=期间×(期初/本期增加/本期减少/期末) 动态 2-3 期。
    返回 ws（≥2 期且有数据时）或 None。"""
    from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_BOLD,
                             SHELL_NUM, SHELL_LEFT, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                             SHELL_TOT_FILL)
    from openpyxl.utils import get_column_letter
    years = sorted(rec_by_year.keys())
    if len(years) < 2:
        return None
    all_ents = sorted({e for y in years for e in rec_by_year[y].keys()})
    if not all_ents:
        return None
    ent_sv = {}
    for e in all_ents:
        for y in years:
            rec = rec_by_year.get(y, {}).get(e)
            ent_sv[(e, str(y))] = _tax_std_vals(rec['full']) if rec else {}
    blocks = [(t, t) for t, _c, _f in _TAX_AUDIT_ROWS] + [('小  计', '小  计')]
    ws = wb.create_sheet('应交税费多期对比表')
    nc = 2 + len(years) * 4
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, f'应交税费多期对比表（{"、".join(str(x) for x in years)} 年度）· 行=税种×各核算主体、列=期间×(期初/本期增加/本期减少/期末) · 账面数（未审）')
    c.font = SHELL_TITLE_FONT
    ws.cell(2, 1, '税种').font = SHELL_HFONT; ws.cell(2, 1).fill = SHELL_HFILL
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
    cols3 = ['税种', '核算主体'] + [f'{y}-{m}' for y in years for m in ('期初', '增加', '减少', '期末')]
    for j, h in enumerate(cols3, 1):
        cc = ws.cell(3, j, h)
        cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    r = 4
    for nm, _k in blocks:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nc)
        cc = ws.cell(r, 1, nm)
        cc.font = SHELL_BOLD
        cc.fill = SHELL_TOT_FILL
        for _j in range(1, nc + 1):
            ws.cell(r, _j).border = SHELL_BORDER
        r += 1
        for e in all_ents:
            row_vals = []
            for y in years:
                sv = ent_sv.get((e, str(y)), {})
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
            for e in all_ents:
                sv = ent_sv.get((e, str(y)), {})
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
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 14
    for j in range(3, nc + 1):
        ws.column_dimensions[get_column_letter(j)].width = 13
    ws.freeze_panes = 'C4'
    return ws


def build_tax_multi_period_table(data_dir, out_dir=None):
    """应交税费多期对比表（独立跨年文件，渲染复用 _render_tax_multi_period_sheet）。"""
    from openpyxl import Workbook
    from audit_shell import finalize_workbook
    entities_map = _discover_tax_entities(data_dir)
    if not entities_map:
        return None
    rec_by_year = {}
    for y in sorted({yy for b in entities_map.values() for yy in b.keys()}):
        rec_by_year[y] = {}
        for ent, b in entities_map.items():
            yr = b.get(y)
            if not yr or not yr.get('km'):
                continue
            try:
                gl = read_gl_tax(yr.get('gl'))
                tb = load_tb(yr['km'])
                vat = build_vat_entity(gl, tb)
                full = build_full_entity(gl, tb)
            except Exception:
                continue
            rec_by_year[y][ent] = {'vat': vat, 'full': full, 'gl': gl, 'tb': tb}
    wb = Workbook()
    wb.remove(wb.active)
    ws = _render_tax_multi_period_sheet(wb, data_dir, rec_by_year)
    if ws is None:
        wb.close()
        return None
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, '应交税费多期对比表_生成.xlsx')
    finalize_workbook(wb)
    wb.save(out)
    wb.close()
    print(f'  ✅ 应交税费多期对比表：{out}（{"、".join(str(x) for x in sorted(rec_by_year.keys()))} 年度）')
    return out
