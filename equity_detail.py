# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=各所有者权益科目底稿(实收资本或股本/资本公积/盈余公积/专项储备/未分配利润)审计底稿_{year}_生成.xlsx | 关键列=期初数/本期增加/本期减少/期末数/审计调整/审定数/凭证抽查(剔除当期损益结转) | 职责=所有者权益增减变动及明细+凭证抽查(通用,按科目名识别)

from audit_common import _safe_save, voucher_skip_keys  # 共享库：统一锁感知保存 + 抽凭排除规则
"""所有者权益审计底稿生成程序
================================================================================
覆盖以下所有者权益科目（均按科目【名称】自动识别，兼容不同会计科目表）：

  · 实收资本(或股本)  4001  有限责任公司称"实收资本"、股份有限公司称"股本"（互斥，按名匹配存在的那个）
  · 资本公积          4002
  · 盈余公积          4101
  · 专项储备          4102（安全生产费等；部分账套挂在盈余公积/资本公积下，按名匹配）
  · 未分配利润        4104 利润分配-未分配利润（结转损益后的留存收益余额）

数据来源（与其他底稿一致）：
  · 科目余额表 = 权威控制数（期初/借发/贷发/期末，带借贷符号）。
  · 综合查询明细表(GL) = 凭证级，用于【凭证抽查】。

输出（每个【存在该科目的主体】的年份，分别生成一份底稿）：
  · <被拖文件夹>/实收资本(或股本)审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/资本公积审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/盈余公积审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/专项储备审计底稿_{year}_生成.xlsx
  · <被拖文件夹>/未分配利润审计底稿_{year}_生成.xlsx

每张底稿含：
  ① 目录说明
  ② 增减变动及明细表：按主体分组，每个主体列明细科目(二级)及合计；
     列=期初数/本期增加/本期减少/期末数/审计调整/审定数，并设勾稽列(期初+增-减=末)。
     注：所有者权益为贷方余额科目，本期增加=贷发、本期减少=借发。
  ③ 凭证抽查：抽取该科目当年全部凭证(借/贷发生)，【剔除当期损益结转】
     （即"本年利润"结转"未分配利润"的结账分录），列示供审计抽凭；
     含主体/日期/凭证号/摘要/借贷方/对方科目/核算明细/抽查/审计结论(留白)。
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
        _p = _bs_os.path.join(_bs_os.path.expanduser('~'), 'Desktop', 'equity_boot.log')
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
import openpyxl
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TOT_FILL, SHELL_TITLE_FONT, SHELL_SUB_FONT,
                         SHELL_BOLD, SHELL_NUM, SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT,
                         finalize_workbook)
from audit_common import (discover_entities, read_tb_full, read_gl_rows, safe,
                         is_zero_amount, _recon_status)
from mask_dict import mask_names   # ⚡ console 主体列表打印脱敏（2026-08-22 复扫）

# ----------------------------------------------------------------------------
# 科目组定义（按科目【名称】识别，兼容不同科目表）
# ----------------------------------------------------------------------------
# names : 匹配用关键词（名称包含即命中）
# is_credit : 是否贷方余额科目（权益类均为 True）
GROUPS = {
    'PAIDIN': {
        'key': 'PAIDIN',
        'title': '实收资本(或股本)', 'paper': '实收资本(或股本)审计底稿',
        'main_sheet': '实收资本(或股本)明细表',
        'vouch_sheet': '实收资本(或股本)凭证抽查',
        'names': ['实收资本', '股本'], 'is_credit': True,
    },
    'CAPRES': {
        'key': 'CAPRES',
        'title': '资本公积', 'paper': '资本公积审计底稿',
        'main_sheet': '资本公积明细表',
        'vouch_sheet': '资本公积凭证抽查',
        'names': ['资本公积'], 'is_credit': True,
    },
    'SURRES': {
        'key': 'SURRES',
        'title': '盈余公积', 'paper': '盈余公积审计底稿',
        'main_sheet': '盈余公积明细表',
        'vouch_sheet': '盈余公积凭证抽查',
        'names': ['盈余公积'], 'is_credit': True,
    },
    'SPECRES': {
        'key': 'SPECRES',
        'title': '专项储备', 'paper': '专项储备审计底稿',
        'main_sheet': '专项储备明细表',
        'vouch_sheet': '专项储备凭证抽查',
        'names': ['专项储备'], 'is_credit': True,
    },
    'RETAINED': {
        'key': 'RETAINED',
        'title': '未分配利润', 'paper': '未分配利润审计底稿',
        'main_sheet': '未分配利润明细表',
        'vouch_sheet': '未分配利润凭证抽查',
        'names': ['未分配利润'], 'is_credit': True,
    },
    'TREASURY': {
        'key': 'TREASURY',
        'title': '库存股', 'paper': '库存股审计底稿',
        'main_sheet': '库存股明细表',
        'vouch_sheet': '库存股凭证抽查',
        'names': ['库存股'], 'is_credit': False,   # 借方余额科目（回购库存股记借方）
    },
}


# ----------------------------------------------------------------------------
# 样式 / 单元格助手（与长期资产底稿同源，保证视觉一致）
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


# ----------------------------------------------------------------------------
# 数据准备
# ----------------------------------------------------------------------------
def _match_name(name, g):
    """判断某 TB 科目名是否属于该组。未分配利润严格按'未分配利润'命中。"""
    n = name or ''
    if g['key'] == 'RETAINED':
        return '未分配利润' in n
    if g['key'] == 'PAIDIN':
        # 2026-08-09 修复：names=['实收资本','股本'] 会把「资本公积-资本(股本)溢价」
        # 误吸入 PAIDIN（1010 曾 matched 4 行 → pfxs 双前缀 → 聚合失败 → 实收资本漏子目）。
        # 股本命中须排除资本公积类（'资本(股本)溢价' 是资本公积子目，非实收资本）。
        return ('实收资本' in n) or ('股本' in n and '资本公积' not in n)
    return any(kw in n for kw in g['names'])


def _find_equity(tb, e, y, g):
    """在 TB 中找到该主体的该权益科目：返回 {'pcode','pname','pv','children'}。"""
    if g['key'] == 'RETAINED':
        # ⚡ 2026-08-13 修复：未分配利润=4104「利润分配」族（含未分配利润/提取盈余公积/
        #   应付股利等子目）。原仅按名称『未分配利润』命中 → 1250 等主体只有
        #   『利润分配-提取法定盈余公积』子目（名称不含"未分配利润"）→ 漏生成底稿（铁律110）。
        matched = [(c, n, v) for (ee, c, n, yy), v in tb.items()
                   if ee == e and yy == y
                   and ('未分配利润' in n or '利润分配' in n or str(c).startswith('4104'))]
    else:
        matched = [(c, n, v) for (ee, c, n, yy), v in tb.items()
                   if ee == e and yy == y and _match_name(n, g)]
    if not matched:
        return None
    # 主科目 = 代码最短者（顶层父科目）
    # ⚡ 2026-08-09 SAP 修复：SAP TB 无 4 位父级行（只有 10 位子目，如 4001010100 法人资本金）。
    # 原逻辑 min(len(code)) 会把第一个子目当父级 → pv=子目值、children 前缀匹配空 → 漏其余子目
    # （1010 实收资本 7.8亿 vs TB 9.78亿；资本公积 20.6亿 vs TB 23.4亿）。
    # 修复：matched 含多个同级子目且无 4 位父级 → 聚合虚拟父级（4 位公共前缀 + 带符号汇总）。
    _dig = lambda s: ''.join(ch for ch in str(s) if ch.isdigit())
    parent_rows = [m for m in matched if len(_dig(m[0])) == 4]
    _virtual = False
    if not parent_rows and len(matched) > 1:
        pfxs = {_dig(m[0])[:4] for m in matched}
        if len(pfxs) == 1:
            pfx = pfxs.pop()
            _v = dict(qc=0.0, jf=0.0, df=0.0, qm=0.0)
            for _c, _n, _vv in matched:
                for k in _v:
                    _v[k] += safe(_vv.get(k)) if isinstance(_vv, dict) else 0.0
            pcode, pname, pv = pfx, g['title'], _v
            _virtual = True
        else:
            primary = min(matched, key=lambda x: len(x[0]))
            pcode, pname, pv = primary
    else:
        primary = min(parent_rows or matched, key=lambda x: len(x[0]))
        pcode, pname, pv = primary
    # 2026-08-04 修复：子目按【代码前缀】从 TB 全量提取（名称无关）。
    # 原实现从 matched（名称命中结果）取子目，导致名称不含组关键词的子目全部丢失
    # （如资本公积下『资本溢价/国家拨入/股权投资准备』，FY 曾只剩 4002.04 一个明细、
    #   3.74 亿一级余额明细仅 7.5 万）。代码前缀兼容点分(4002.01)与连写(400201)。
    _pd = _dig(pcode)
    children = []
    for (ee, c, n, yy), v in tb.items():
        if ee != e or yy != y or str(c) == pcode:
            continue
        _cd = _dig(c)
        if _cd.startswith(_pd) and _cd != _pd:
            children.append((c, n, v))
    return {'pcode': pcode, 'pname': pname, 'pv': pv, 'children': children,
            'virtual': _virtual}


def _row_values(v, is_credit=True, flip_neg=False):
    """从 TB 行取 期初/借发/贷发/期末（带符号），按科目余额方向输出展示值。
    is_credit=True（默认，权益类贷方余额）：本期增加=贷方发生(df)、本期减少=借方发生(jf)；
    is_credit=False（借方余额科目，如库存股）：本期增加=借方发生(jf)、本期减少=贷方发生(df)。
    flip_neg=False（2026-08-27 新增）：单一贷余族（实收/资本公积/盈余公积/专项储备，叶子全
    qm<0）传 True → 整体取反显示贷方正（与审定表 abs 一致）；混合方向族（未分配利润=提取盈余
    公积贷余负+未分配利润借余正）传 False → 保持带符号（避免 abs 虚增合计，明细合计=审定）。
    ⚡ 2026-08-15 铁律129c：期初/期末【不再 abs】——权益类按【贷正带符号】输出
    （贷方余额为正、借方余额/亏损为负）。原 abs 会把借方余额主体（未分配利润借方=亏损、
    资本公积借方=减项）翻正 → 底稿全集团合计 Σ|主体| ≠ 试算表 Σ 带符号（XBJ 未分配利润
    底稿 32.55 亿 vs 试算表 8.94 亿根因）。审计底稿应诚实列示方向。"""
    qc = safe(v.get('qc'))
    jf = safe(v.get('jf'))
    df = safe(v.get('df'))
    qm = safe(v.get('qm'))
    if is_credit:
        # ⚡ 2026-08-26 修复（未分配利润明细符号 bug）：TB 权益贷余符号账套不一——U8 贷余记负
        #   （qm_show=-qm 转正），SAP 贷余记正（AH 4104 未分配利润 qm=+35.62亿，再取反→-35.62亿
        #   明细 vs 审定 +35.62亿 全集团 7.12亿假差异）。按数据源判断是否取反。
        # ⚡⚡ 2026-08-27 修复（AH 底稿复查——明细表 vs 审定表符号一致性 + 勾稽平）：
        #   SAP 权益科目 TB 符号不一：实收/资本公积/盈余公积/专项储备 = 贷方余额（qm<0 贷余记负）
        #   → 明细表负号 vs 审定表 abs 正号 符号相反；未分配利润族 = 混合（410401 贷余负 +
        #   410406 借余正），明细合计(带符号) = 审定。且未分配利润(借余 qm>0)按 is_credit 把
        #   借发当"减少"，实际借发使借余增大 → 勾稽不平(-587M)。
        #   修复：借余科目（qm>0）inc/dec 取反（借发=增加）；符号按 flip_neg——全贷余族取反
        #   显示贷方正（与审定 abs 一致）、混合族保持带符号（合计=审定）。
        if _adapter is not None and _adapter.is_sap(getattr(_adapter, '_DATA_ROOT', '') or ''):
            if qm < 0:
                # 贷方余额科目：贷发=增加、借发=减少
                inc, dec = df, jf
            else:
                # 借方余额科目（未分配利润借余）：借发=增加（借余增大）、贷发=减少
                inc, dec = jf, df
            if flip_neg:
                # 单一贷余族 → 整体取反显示贷方正（与审定表 abs 一致）
                qc_show = -qc
                qm_show = -qm
            else:
                # 混合方向族 → 保持带符号（明细合计=审定，不虚增）
                qc_show = qc
                qm_show = qm
        else:
            qc_show = -qc   # U8 贷正：贷方余额为正、借方余额为负
            qm_show = -qm
            inc, dec = df, jf
    else:
        qc_show = qc     # 借方余额科目（库存股）借正
        qm_show = qm
        inc = jf     # 借方发生 = 借方余额科目增加
        dec = df     # 贷方发生 = 借方余额科目减少
    return qc_show, inc, dec, qm_show


# ----------------------------------------------------------------------------
# 凭证抽查：剔除"结转损益类"结账分录（统一套用共享抽凭排除规则）
# ----------------------------------------------------------------------------
def _collect_vouchers(gl, year, g, ent=None):
    """收集该科目当年全部凭证行（剔除结转损益类凭证）。返回 (rows, excluded_count)。
    ent 给定时仅取该主体（用于逐主体统计剔除数，避免多主体重复计列）。
    排除判定统一调用 audit_common.voucher_skip_keys（2026-07-25 抽凭通用规则）：
      · 规则① 结转损益类凭证不抽（覆盖"本年利润转未分配利润""结转本年利润""结转损益"等）；
      · 规则② 相符计提/摊销不抽（权益类凭证不触发本规则）。
    先按 (e, y, 凭证字, 凭证号) 聚合为整笔凭证，再套用共享规则得到应排除凭证集合。"""

    def _group_by_key(recs, key_fields):
        from collections import OrderedDict
        groups = OrderedDict()
        for rr in recs:
            k = tuple(key_fields(rr))
            groups.setdefault(k, []).append(rr)
        return groups.items()

    matched = []
    for r in gl:
        if r.get('y') != year:
            continue
        if ent is not None and r.get('e') != ent:
            continue
        if not _match_name(r.get('name'), g):
            continue
        matched.append(r)
    if matched:
        # 未分配利润(2026-08-01 修复 #13)：通用结转规则会因「科目名含未分配利润」误伤全部凭证。
        # 仅剔除「对方科目/凭证文本含 本年利润 或 损益类结转」的真实结转凭证；
        # 提取盈余公积/补提/分配等非结转凭证保留抽凭。
        if g['key'] == 'RETAINED':
            skip = set()
            for k, lines in _group_by_key(matched, key_fields=lambda rr: (rr.get('e'), rr.get('y'), rr.get('vtype'), rr.get('vno'))):
                blob = ' '.join(' '.join(str(x) for x in (l.get('name') or '', l.get('opp') or '', l.get('summary') or ''))
                                for l in lines)
                # 仅当整笔凭证含「本年利润」或摘要含结转损益/结转未分配利润关键词才剔除
                if '本年利润' in blob or ('结转' in blob and ('损益' in blob or '未分配利润' in blob)):
                    skip.add(k)
        else:
            skip = voucher_skip_keys(
                matched,
                key_fields=lambda rr: (rr.get('e'), rr.get('y'), rr.get('vtype'), rr.get('vno')),
                name_field=lambda rr: rr.get('name') or '',
                opp_field=lambda rr: rr.get('opp') or '',
                summ_field=lambda rr: rr.get('summary') or '',
            )
    else:
        skip = set()
    rows = []
    excluded = 0
    for r in matched:
        if (r.get('e'), r.get('y'), r.get('vtype'), r.get('vno')) in skip:
            excluded += 1
            continue
        rows.append(r)
    # 按 主体→日期→凭证号 排序，稳定输出
    def _sortkey(x):
        d = x.get('date')
        ds = d.isoformat() if hasattr(d, 'isoformat') else str(d)
        return (x.get('e', ''), ds, str(x.get('vtype') or ''), str(x.get('vno') or ''))
    rows.sort(key=_sortkey)
    return rows, excluded


def _fmt_vno(r):
    vt = r.get('vtype') or ''
    vn = r.get('vno') or ''
    if vt and vn:
        return '%s-%s' % (vt, vn)
    return str(vn) if vn else ''


def _fmt_date(d):
    if d is None:
        return ''
    if hasattr(d, 'isoformat'):
        return d.isoformat()[:10]
    return str(d)


# ----------------------------------------------------------------------------
# Sheet: 目录说明
# ----------------------------------------------------------------------------
def build_cover(wb, data_dir, year, ents, recs, gkey, notes):
    g = GROUPS[gkey]
    ws = wb.create_sheet('目录说明', 0)
    ws.cell(1, 1, '%s审计底稿（%s 年度）' % (g['title'], year)).font = SHELL_TITLE_FONT
    present = [e for e, _ in recs]
    lines = [
        ('数据目录', data_dir),
        ('年度', year),
        ('主体数', '%d 个：%s' % (len(present), '、'.join(present))),
        ('数据来源', '《科目余额表》（控制数）+《综合查询明细表(GL)》（凭证级）'),
        ('科目范围', '%s（按科目名自动识别，兼容不同科目表）' % g['title']),
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
        (g['main_sheet'], '主表：按主体分组列示各明细科目的期初数/本期增加/本期减少/期末数/审计调整/审定数；'
         + ('权益为贷方余额科目，本期增加=贷发、本期减少=借发' if g['is_credit']
            else '库存股为借方余额科目，本期增加=借发、本期减少=贷发')
         + '；末列勾稽(期初+增-减=末)应=0。'),
        (g['vouch_sheet'], '该科目当年全部凭证(已剔除"当期损益结转"/本年利润结转结账分录)的抽凭清单，'
         '列示主体/日期/凭证号/摘要/借/贷/对方科目/核算明细，并留"抽查/审计结论"列供填写。'),
    ]
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
    for i, w in enumerate([26, 86, 10, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


# ----------------------------------------------------------------------------
# Sheet: 增减变动及明细表
# ----------------------------------------------------------------------------
def build_rollforward_sheet(wb, year, recs, gkey):
    g = GROUPS[gkey]
    ncols = 9
    # 2026-08-04 规范化：B 列『项目(明细科目)』→『明细科目』——含『项目』会触发核对程序
    # 组件块净值过滤（固定资产 原值/折旧/净值 用），权益明细表被误伤致单体提取恒空。
    headers = ['主体', '明细科目', '期初数', '本期增加', '本期减少', '期末数',
               '审计调整', '审定数', '勾稽(期初+增-减-末)']
    # 列号
    C_ENT, C_ITEM, C_QC, C_INC, C_DEC, C_QM, C_ADJ, C_AUD, C_CHK = range(1, ncols + 1)
    MONEY = {C_QC, C_INC, C_DEC, C_QM, C_ADJ, C_AUD, C_CHK}   # 金额列

    # ⚡ 2026-08-11 阶段二 2.1：计算层产【结构化行】(RowSpec)，渲染收敛到 audit_render.render_sheet。
    #   本函数只做计算与行组织，不含任何 openpyxl 样式调用；逐格比对旧文件验证零信息损失。
    from audit_render import render_sheet, row as _arow
    rows = []
    tot = [0.0, 0.0, 0.0, 0.0]  # 期初/增/减/末 全集团合计（一级控制数）
    recon = []                  # 每主体：一级(主体合计) ↔ 二级明细合计（期末数口径）
    ent_detail_rows = []        # 每主体明细行号列表（合计公式引用，2026-08-02）
    for ent, rec in recs:
        detail_rows = []
        # ⚡ 2026-08-27 修复（AH 底稿复查）：判断本主体该科目族的余额方向——叶子全贷余
        #   （qm<0，实收/资本公积/盈余公积/专项储备）→ flip_neg=True 取反显示贷方正（与审定
        #   abs 一致）；混合方向（未分配利润=410401 贷余负+410406 借余正）→ flip_neg=False
        #   保持带符号（明细合计=审定，避免 abs 虚增）。一级/二级共用同一 flip_neg 保证一致。
        children = rec['children']
        _flip = False
        if g['key'] != 'RETAINED':
            # 非未分配利润族（实收/资本公积/盈余公积/专项储备，Balance=Debit 贷余负取反显示贷方正）：
            # 以「主体净额方向」判断——叶子 qm 带符号合计 < 0（贷方净额）即 flip。
            # ⚡⚡ 2026-08-27 修复（AH 底稿复查）：原 all(qm<0) 判断在【部分叶子期末平】
            #   时失效（2150 实收 4001010100 qm=-1500万 + 4001020200 qm=0 → all=False 不 flip
            #   → 明细 -1500万 vs 审定 +1500万 差 3000万）。改用净额方向：sqm<0 贷余即取反。
            if children:
                _allc = [c for c, _n, _v in children]
                leaf0 = [(c, n, v) for c, n, v in children
                         if not any(o != c and o.startswith(c) for o in _allc)]
                _sqm = sum(safe(v.get('qm')) for _c, _n, v in leaf0)
                if _sqm < 0:
                    _flip = True   # 主体净贷余 → 取反显示贷方正
            elif safe(rec['pv'].get('qm')) < 0:
                # 无 children（单叶子科目，如盈余公积/专项储备）→ pv 即叶子，贷余则取反
                _flip = True
        # ⚡⚡ 2026-08-27 修复（AH 底稿复查）：未分配利润(4104)族 SAP Balance=Credit（贷方正、
        #   借余负）——qm>0=盈利、qm<0=亏损主体。审计保持带符号（盈利正/亏损负），不 flip。
        #   原逻辑按 qm<0 判断 flip 会把亏损主体（如 1357 1220 未分配利润 -1.006亿 借余）
        #   误取反显示 +1.006亿 → 明细(+1.006亿) vs 审定(-1.006亿) 符号相反。
        #   与审定表 keep_sign 同口径（RETAINED 传 keep_sign=True）。
        # 主体小计行（取父科目 = 一级）
        qc, inc, dec, qm = _row_values(rec['pv'], g['is_credit'], flip_neg=_flip)
        # 二级明细合计（仅取末级/叶子码，避免「父+子」重复计列，如 410505 法人股本 与其子 41050501 同值）
        if children:
            _allc = [c for c, _n, _v in children]
            leaf = [(c, n, v) for c, n, v in children
                    if not any(o != c and o.startswith(c) for o in _allc)]
            sqc = sinc = sdec = sqm = 0.0
            # 2026-08-07 修复：二级明细合计须【带符号求和再 abs】（净额口径）——Q 资本公积
            # 400201 借余(+78.4M)+400202 溢价贷余(-409.8M)+400203 贷余(-42.5M)，
            # 原 abs 逐行加 → 530.7M vs 一级净额 373.8M（明细表自标勾稽差异 -1.57 亿）。
            # 权益类正常全贷余（同方向）时带符号和 abs=逐行 abs，无回归。
            for _c, _n, _v in leaf:
                _qc = safe(_v.get('qc')); _qm = safe(_v.get('qm'))
                if g['is_credit']:
                    if _qm < 0:
                        _inc = safe(_v.get('df')); _dec = safe(_v.get('jf'))
                    else:
                        _inc = safe(_v.get('jf')); _dec = safe(_v.get('df'))
                else:
                    _inc = safe(_v.get('jf')); _dec = safe(_v.get('df'))
                sqc += _qc; sinc += _inc; sdec += _dec; sqm += _qm
            # ⚡ 2026-08-15 铁律129c：主体合计与 _row_values 一致——权益类【贷正带符号】
            #   （贷方余额为正、借方余额为负），不再 abs（亏损/借方减项主体不被翻正）。
            # ⚡⚡ 2026-08-26 修复（任务878 未分配利润二级合计符号 bug）：二级合计的取反
            #   约定必须与 _row_values 一致——SAP 贷余记正不取反、U8 贷余记负才取反。
            #   原逻辑无条件 -sqc/-sqm 对 SAP 把 +35.62亿 翻成 -35.62亿 → 一级(+35.62亿)
            #   vs 二级(-35.62亿) 假差异 7.12亿（410401 借方负 + 410406 贷方正 混合方向）。
            # ⚡⚡⚡ 2026-08-27：二级与一级共用 flip_neg（全贷余族取反、混合族带符号）。
            if _adapter is not None and _adapter.is_sap(getattr(_adapter, '_DATA_ROOT', '') or ''):
                if _flip:
                    sqc = -sqc
                    sqm = -sqm
            else:
                sqc = -sqc if g['is_credit'] else sqc
                sqm = -sqm if g['is_credit'] else sqm
        else:
            leaf = [(rec['pcode'], rec['pname'], rec['pv'])]
            sqc, sinc, sdec, sqm = qc, inc, dec, qm
        # 主体分组标题（整行浅蓝填充 + 加粗；A 列主体名合并 B 列）
        rows.append(_arow([ent, None, None, None, None, None, None, None, None],
                          b=True, fill='sub', merge=(C_ENT, C_ITEM)))
        # 明细行（二级明细来自 TB；仅列末级叶子，0 值行不显示）
        has_children = bool(children)
        detail = leaf
        for code, name, v in detail:
            dqc, dinc, ddec, dqm = _row_values(v, g['is_credit'], flip_neg=_flip)
            if is_zero_amount(dqc, dinc, ddec, dqm):
                continue  # 0 值行不显示（与费用一致）
            # 2026-08-06 协议：审定数/勾稽写数值（公式 data_only 读 None 收回读不回）
            rows.append(_arow([None, '%s %s' % (code, name), dqc, dinc, ddec, dqm,
                               None, dqm, dqc + dinc - ddec - dqm], num=MONEY))
            detail_rows.append(len(rows))   # 渲染后行号（旧版记录 r；此处记录 rows 索引供结构校验）
        # 主体合计行（仅当存在二级明细时，避免无子目时与唯一明细行重复计列；2026-08-06 协议：写数值=二级叶子合计）
        if has_children:
            rows.append(_arow([None, '%s 合计' % ent, sqc, sinc, sdec, sqm,
                               None, sqm, sqc + sinc - sdec - sqm], b=True, num=MONEY))
        tot[0] += qc; tot[1] += inc; tot[2] += dec; tot[3] += qm
        recon.append((ent, (qc, inc, dec, qm), (sqc, sinc, sdec, sqm)))
        ent_detail_rows.append(detail_rows)
    # 全集团合计（2026-08-06 协议：写数值=tot 累计；公式 data_only 读 None 收回读不回）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出全集团合计行
    if len(recs) > 1:
        rows.append(_arow([None, '全集团合计', tot[0], tot[1], tot[2], tot[3],
                           None, tot[3], tot[0] + tot[1] - tot[2] - tot[3]],
                          b=True, fill='sub', num=MONEY))
    # 勾稽提示块前空行（原版 r += 1，逐格对齐）
    rows.append(_arow([None] * ncols, skip_border=True))
    # ── 一级(TB) ↔ 二级明细合计(TB) 勾稽提示（从一级出发的透明化）──
    rows.append(_arow(['勾稽：一级(TB) ↔ 二级明细合计(TB)（期末数口径，差额应≈0）',
                       None, None, None, None, None, None, None, None],
                      b=True, merge=(C_ENT, C_CHK), align='l'))
    any_warn = False
    for ent, l1, l2 in recon:
        diff = l1[3] - l2[3]
        warn = _recon_status(diff, l1[3], l2[3]) == '不一致'
        if warn:
            any_warn = True
            rows.append(_arow([None, '【%s】一级(主体合计)期末' % ent,
                               None, None, None, l1[3], None, None, None], num=MONEY))
            rows.append(_arow([None, '【%s】二级明细合计期末' % ent,
                               None, None, None, l2[3], None, None, None], num=MONEY))
            rows.append(_arow([None, '【%s】差异(一级−二级)' % ent,
                               None, None, None, diff, None, None, None], b=True, red=True, num=MONEY))
    g1 = sum(x[1][3] for x in recon); g2 = sum(x[2][3] for x in recon); gdiff = g1 - g2
    rows.append(_arow([None, '全集团 一级(主体合计)期末', None, None, None, g1, None, None, None],
                      b=True, num=MONEY))
    rows.append(_arow([None, '全集团 二级明细合计期末', None, None, None, g2, None, None, None],
                      b=True, num=MONEY))
    rows.append(_arow([None, '全集团 差异(一级−二级)', None, None, None, gdiff, None, None, None],
                      b=True, red=True, num=MONEY))
    # 说明行（按科目余额方向输出取数口径）
    if g['is_credit']:
        note = '说明：本期增加=科目余额表贷方发生额；本期减少=借方发生额；勾稽列(期初+本期增加-本期减少-期末)应=0。'
    else:
        note = '说明：库存股为借方余额科目，本期增加=借方发生额；本期减少=贷方发生额；勾稽列(期初+本期增加-本期减少-期末)应=0。'
    rows.append(_arow([note, None, None, None, None, None, None, None, None],
                      skip_border=True, align='l', merge=(C_ENT, C_CHK)))

    widths = [16, 30, 16, 14, 14, 16, 12, 16, 18]
    ws = render_sheet(wb, dict(
        name=f"{g['main_sheet']}_{year}",
        title='%s增减变动及明细表（%s 年度）' % (g['title'], year),
        title_gap=1, headers=headers, rows=rows, widths=widths[:ncols],
        money_cols=MONEY, zero_blank=False,
        freeze='A4',
    ))
    # 2026-08-05 前道规范化：写入后立即结构校验（生成即规范，不靠后期检查）
    try:
        from audit_common import assert_sheet_standard
        assert_sheet_standard(ws, 'detail', g['title'], raise_on_error=False)
    except Exception:
        pass
    return ws


def _set_formula(ws, r, col, formula):
    cell = ws.cell(r, col, formula)
    cell.number_format = SHELL_NUM
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    return cell


# ----------------------------------------------------------------------------
# Sheet: 专项储备计提核对（2026-08-10 用户需求：专项储备变动=计提数，不抽凭但须计提核对）
# 口径：计提=贷方专项储备(4102)净额；费用列支=同凭证借方费用科目（SAP 1010『安全投入』
# 6600180000）——计提凭证结构『借安全投入 贷专项储备』1:1 配对，核对应≈0；
# 使用（借专项储备 贷银行/应付等）不计入列支，单独列示。
# ----------------------------------------------------------------------------
def build_zx_check_sheet(wb, year, recs, gl, gkey):
    g = GROUPS[gkey]
    ws = wb.create_sheet(f'{g["title"]}计提核对表')
    ncols = 8
    headers = ['主体', '月份', '计提数(贷专项储备)', '计提对方科目(借方·计提凭证)',
               '费用列支(借方·计提凭证)', '使用/其他(借专项储备)', '核对(计提-列支)', '差异归因']
    r = _title(ws, '%s计提核对表（%s 年度）' % (g['title'], year), ncols)
    r = _sub_note(ws, '口径：计提=贷方%s净额；计提对方科目=与计提同一凭证的借方科目（1/2级，1:1 配对'
                      '『借费用 贷专项储备』）；使用=借%s（安全投入支出，非计提）。'
                      '核对=计提-列支，≈0 即勾稽相符；非零列差异归因凭证。' % (g['title'], g['title']), ncols, r)
    hr = r + 1
    r = _hdr(ws, hr, headers)
    C = list(range(1, ncols + 1))

    # 按月聚合：计提（贷4102 净）/ 列支（计提凭证借方费用）/ 使用（借4102 净）
    from collections import defaultdict
    months = defaultdict(lambda: {'accr': 0.0, 'exp': 0.0, 'use': 0.0, 'opp': set()})  # m -> ...
    non_exp_detail = []   # 贷4102 但借方非费用的凭证（差异归因）
    vgroups = defaultdict(list)
    for rr in gl:
        if str(rr.get('y')) != str(year):
            continue
        vgroups[(rr.get('e'), rr.get('vtype'), rr.get('vno'))].append(rr)
    for vk, lines in vgroups.items():
        e, _vt, _no = vk
        cr_zx = sum(safe(x.get('cr')) for x in lines if _is_zx_name(x.get('name')))
        dr_zx = sum(safe(x.get('dr')) for x in lines if _is_zx_name(x.get('name')))
        dr_exp = 0.0
        dr_exp_names = []   # ⚡ 2026-08-28：收集计提凭证借方对方科目（1/2级，用户要求列示）
        for x in lines:
            if not _is_zx_name(x.get('name')) and safe(x.get('dr')) != 0:
                dr_exp += safe(x.get('dr'))
                _nm = str(x.get('name') or '')
                _cd = str(x.get('code') or '')
                dr_exp_names.append('%s(%s)' % (_nm[:22], _cd) if _cd else _nm[:26])
        if abs(cr_zx) < 0.005 and abs(dr_zx) < 0.005:
            continue
        m = None
        for x in lines:
            _m = x.get('month')
            if _m:
                m = int(_m)
                break
        if m is None:
            continue
        # ⚡ 2026-08-10 凭证级净额分类（SAP 1010 实测）：
        #   计提凭证 net=贷4102−借4102>0（借安全投入 贷专项储备）；使用凭证 net<0
        #   （借专项储备 贷安全投入/银行，如 6500002479）；借=贷4102（net≈0）=
        #   内部互转（6500003576 等安全费↔维简费调整），不计计提不计使用。
        net = cr_zx - dr_zx
        if abs(net) < 0.005:
            continue
        if net > 0:
            months[m]['accr'] += net
            months[m]['exp'] += dr_exp
            months[m]['opp'].update(dr_exp_names)
            if abs(dr_exp - net) > 0.01:
                non_exp_detail.append({'m': m, 'vno': _no, 'cr': net, 'exp': dr_exp,
                                       'lines': '；'.join('%s 借%s' % (str(x.get('name') or '')[:20],
                                                                       round(safe(x.get('dr')), 2))
                                                          for x in lines if safe(x.get('dr')) != 0)[:120]})
        else:
            months[m]['use'] += -net
    tot_a = tot_e = tot_u = 0.0
    all_opp = set()
    for m in sorted(months):
        d = months[m]
        chk = round(d['accr'] - d['exp'], 2)
        note = '—'
        if abs(chk) > 0.01:
            note = '非费用对方（见差异归因）'
        _txt(ws, r, C[0], recs[0][0])
        _txt(ws, r, C[1], '%02d月' % m)
        _money(ws, r, C[2], round(d['accr'], 2))
        _txt(ws, r, C[3], '；'.join(sorted(d['opp']))[:80])
        _money(ws, r, C[4], round(d['exp'], 2))
        _money(ws, r, C[5], round(d['use'], 2))
        _money(ws, r, C[6], chk)
        _txt(ws, r, C[7], note)
        tot_a += d['accr']; tot_e += d['exp']; tot_u += d['use']
        all_opp.update(d['opp'])
        r += 1
    # 合计行
    _txt(ws, r, C[0], recs[0][0], bold=True)
    _txt(ws, r, C[1], '合计', bold=True)
    _money(ws, r, C[2], round(tot_a, 2))
    _txt(ws, r, C[3], '；'.join(sorted(all_opp))[:80], bold=True)
    _money(ws, r, C[4], round(tot_e, 2)); _money(ws, r, C[5], round(tot_u, 2))
    _money(ws, r, C[6], round(tot_a - tot_e, 2))
    for j in range(1, ncols + 1):
        ws.cell(r, j).fill = SHELL_TOT_FILL
        ws.cell(r, j).font = SHELL_BOLD
    r += 1
    if non_exp_detail:
        r = _sub_note(ws, '差异归因（贷专项储备但借方非费用科目的凭证，逐笔列示）：', ncols, r)
        for d in non_exp_detail[:20]:
            _txt(ws, r, C[0], recs[0][0])
            _txt(ws, r, C[1], '%02d月' % d['m'])
            _txt(ws, r, C[2], '')
            _txt(ws, r, C[3], '')
            _money(ws, r, C[4], round(d['cr'], 2))
            _txt(ws, r, C[5], d['vno'])
            _money(ws, r, C[6], round(d['cr'] - d['exp'], 2))
            _txt(ws, r, C[7], d['lines'])
            r += 1
    for i, w in enumerate([10, 10, 16, 44, 20, 16, 14, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    r += 1
    r = _sub_note(ws, '注：本表计提/使用取自序时账(GL 凭证级净额)；明细表本期增加/减少取自科目余额表。'
                      '两表若不一致，优先核对科目余额表是否含期末日(如 7/31)过账凭证（如 AH 1010 凭证 '
                      '6500004124 借安全投入/贷专项储备 649,432.50 在科目余额表中缺失）；'
                      '确属科目余额表缺失时请企业重新导出后重跑。', ncols, r)
    return ws


def _is_zx_name(name):
    return '专项储备' in str(name or '')


# ----------------------------------------------------------------------------
# Sheet: 凭证抽查
# ----------------------------------------------------------------------------
def build_voucher_sheet(wb, year, recs, gl, gkey):
    g = GROUPS[gkey]
    # 先归集各主体凭证行；无异常(该科目当年无任何凭证交易)则不生成凭证抽查
    excl_total = 0
    per_ent = {}
    for ent, rec in recs:
        vrows, excl = _collect_vouchers(gl, year, g, ent)
        per_ent[ent] = (vrows, excl)
        excl_total += excl
    if not any(ve for ve, _ in per_ent.values()):
        return None
    ws = wb.create_sheet(g['vouch_sheet'])
    ncols = 15
    headers = ['核算主体', '测试序号', '日期', '凭证字', '凭证号',
               '二级科目（或客户/供应商名称）', '摘要', '借方金额', '贷方金额',
               '对方科目', '与原始凭证相符', '原始凭证内容', '原始凭证日期',
               '会计处理正确', '所属时间无误']
    r = _title(ws, '%s凭证抽查（%s 年度，已剔除当期损益结转）' % (g['title'], year), ncols)
    sub = ('说明：下列为本科目当年全部凭证(剔除"当期损益结转"/本年利润结转结账分录，共 %d 笔)；'
           'GL 为非全量抽取，2026 年 GL 仅 1–5 月(YTD)。请据以抽凭，并在"与原始凭证相符"等列填写。'
           % excl_total)
    r = _sub_note(ws, sub, ncols, r)
    hr = r + 1
    r = _hdr(ws, hr, headers)
    C = list(range(1, ncols + 1))
    grand_dr = grand_cr = 0.0
    # 归集所有行后按 (核算主体, 日期) 排序
    all_rows = []
    for ent, (ve, _) in per_ent.items():
        for x in ve:
            dr = safe(x.get('dr')); cr = safe(x.get('cr'))
            all_rows.append((
                ent,
                _fmt_date(x.get('date')),
                x.get('vtype') or '',
                x.get('vno') or '',
                x.get('name') or '权益类',
                x.get('summary') or '',
                round(dr, 2) if dr > 0 else None,
                round(cr, 2) if cr > 0 else None,
                x.get('opp') or '',
            ))
    all_rows.sort(key=lambda r: (str(r[0] or ''), str(r[1] or '')))
    for i, x in enumerate(all_rows, 1):
        _txt(ws, r, C[0], x[0])    # 核算主体
        _txt(ws, r, C[1], i)       # 测试序号
        _txt(ws, r, C[2], x[1])    # 日期
        _txt(ws, r, C[3], x[2])    # 凭证字
        _txt(ws, r, C[4], x[3])    # 凭证号
        _txt(ws, r, C[5], x[4])    # 二级科目（或客户/供应商名称）
        _txt(ws, r, C[6], x[5])    # 摘要
        _money(ws, r, C[7], x[6])  # 借方金额
        _money(ws, r, C[8], x[7])  # 贷方金额
        _txt(ws, r, C[9], x[8])    # 对方科目
        # col 11-15: 与原始凭证相符/原始凭证内容/原始凭证日期/会计处理正确/所属时间无误（空）
        grand_dr += x[6] or 0
        grand_cr += x[7] or 0
        r += 1
    # 合计
    _txt(ws, r, C[6], '合计', bold=True)
    _money(ws, r, C[7], grand_dr)
    _money(ws, r, C[8], grand_cr)
    for i in range(1, ncols + 1):
        ws.cell(r, i).font = SHELL_BOLD
    r += 1
    # 抽凭勾稽：借贷合计应 ≈ 该科目当年发生额（仅提示）
    _txt(ws, r, C[0], '注：上表借贷合计为该科目当年(剔除结转后)发生额，应与《增减变动及明细表》的"本期增加/减少"口径审慎核对。')
    ws.merge_cells(start_row=r, start_column=C[0], end_row=r, end_column=ncols)

    widths = [14, 10, 12, 8, 10, 24, 30, 14, 14, 24, 10, 14, 10, 10, 10]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    return ws


def _sub_note(ws, text, ncols, row):
    c = ws.cell(row, 1, text)
    c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    c.alignment = SHELL_LEFT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    return row + 1


# ----------------------------------------------------------------------------
# 主流程
# ----------------------------------------------------------------------------
def process_folder(data_dir, out_path=None):
    if not os.path.isdir(data_dir):
        print(f'[ERROR] 数据目录不存在：{data_dir}')
        return 1
    ents = discover_entities(data_dir)
    if not ents:
        print(f'[ERROR] 未在 {data_dir} 发现科目余额表/综合查询明细表')
        return 1
    print(f'[INFO] 发现主体 {len(ents)} 个：{mask_names(sorted(ents))}')

    tb = read_tb_full(data_dir, ents)
    gl = read_gl_rows(data_dir, ents)

    years = set()
    for ent in ents:
        years |= set(ents[ent].keys())  # 结构: {实体: {年份: {'km':..,'gl':..}}}
    years = sorted(years)
    if not years:
        print('[ERROR] 未发现任何年份的账套数据')
        return 1
    print(f'[INFO] 检测到年份：{years}')

    rc = 0
    for year in years:
        for gkey in GROUPS:
            rc_year = _build_year_workbook(data_dir, ents, tb, gl, year, gkey)
            if rc_year != 0:
                rc = rc_year
    return rc


def _build_year_workbook(data_dir, ents, tb, gl, year, gkey):
    g = GROUPS[gkey]
    recs = []
    for ent in sorted(ents):
        eq = _find_equity(tb, ent, year, g)
        if not eq:
            continue
        recs.append((ent, eq))
    if not recs:
        return 0  # 该科目本年无数据，跳过（不产生空底稿）

    out_path = os.path.join(data_dir, '%s_%s_生成.xlsx' % (g['paper'], year))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    notes = ['GL 为非全量抽取，GL缺口已在各表披露',
             '2026 年 GL 仅 1–5 月（YTD）' if '2026' in str(year) else '',
             '凭证抽查已剔除"当期损益结转"(本年利润↔未分配利润结账分录)']
    # 目录说明已删除（2026-07-29 清理）
    build_rollforward_sheet(wb, year, recs, gkey)
    print(f'[OK] {g["title"]}增减变动及明细表 ({year})')
    # 审定表（置于主明细表之前，数据取自科目余额表，期末未审数=TB期末，可核对）
    try:
        from audit_common import add_audit_summary_sheets as _add_audit
        # 2026-08-07 修复：泰国主体 pcode（3100 实收资本等）须映射为集团标准码——
        # add_audit 内部 normalize_thai_tb 会生成【原键 3100 + 映射键 4001】两行，
        # codes 若含 3100（泰国原键）则 4001 映射键也被 codes 命中 → 双计 2 倍
        #（Z 泰国实收资本 980M 审定成 1,960M）。映射后 codes 只含标准码，泰国只计 1 次。
        try:
            from thai_mapping import thai_code_to_std as _tcs, is_thai_entity as _ithai
            # ⚡⚡ 2026-08-24 修复（AJJ 专项储备误配根因）：thai_code_to_std 无条件把
            #   3100→4001（泰国代码表），但 AJJ 等非泰国账套 3100=专项储备 → 审定表
            #   误取 4001 生产成本（223.6万 vs 3100 专项储备 352.3万）。仅泰国主体才映射。
            # ⚡⚡ 2026-08-24 名称主键（铁律58）强化：即使泰国主体，科目名已含本组名
            #   （专项储备等）→ 名称优先，用原代码不映射（防泰国账套内同码异义）。
            _is_thai = any(_ithai(str(ent)) for (ent, eq) in recs)
            _gnames = [x for x in ([g.get('title')] + list(g.get('names') or [])) if x]

            def _maybe_tcs(code, nm):
                if any(gn and gn in str(nm or '') for gn in _gnames):
                    return code
                return _tcs(code) if _is_thai else code
            _codes = sorted({_maybe_tcs(eq['pcode'], eq.get('pname') or '') for (ent, eq) in recs})
        except Exception:
            _codes = sorted({eq['pcode'] for (ent, eq) in recs})
        # 2026-08-07：entity_codes（按主体各自 pcode）——J 实收资本东轴=4001 股本、其他
        # 主体=3001 实收资本，全局 codes 并集会精确匹配到仁旭 4001 生产成本（审定 268 万
        # =生产成本 273 万−实收资本 5 万，假差异）。按主体 codes 只匹配该主体自己的科目。
        # 2026-08-09 SAP 修复：虚拟父级（SAP TB 无 4 位父级行，_find_equity 聚合 4001）
        # 在 TB 无对应行 → BS 审定表精确匹配 c in _cl 取 0。虚拟场景 entity_codes 传
        # 子目 codes 列表（精确匹配各子目求和=父级聚合值），非虚拟仍传 pcode。
        def _ent_codes(eq):
            # ⚡⚡ 2026-08-24 仅泰国主体才 _tcs 映射 + 名称主键（铁律58）：名称已含组名不映射
            try:
                from thai_mapping import thai_code_to_std as _tcs, is_thai_entity as _ithai
                _gn2 = [x for x in ([g.get('title')] + list(g.get('names') or [])) if x]
                _do_map = _ithai(str(ent)) and not any(
                    gn and gn in str(eq.get('pname') or '') for gn in _gn2)
                if eq.get('virtual'):
                    try:
                        return sorted({(_tcs(c) if _do_map else c) for (c, _n, _v) in eq['children']}) or [eq['pcode']]
                    except Exception:
                        return [c for (c, _n, _v) in eq['children']] or [eq['pcode']]
                return [_tcs(eq['pcode'])] if _do_map else [eq['pcode']]
            except Exception:
                return [eq['pcode']]
        _ec = {}
        for ent, eq in recs:
            _ec[ent] = _ent_codes(eq)
        _items = [dict(title=f"{g['title']} 审定表", codes=_codes, is_credit=g['is_credit'],
                       before=f"{g['main_sheet']}_{year}", subj_name=g['title'], by_entity=True,
                       entity_codes=_ec,
                       # ⚡ 2026-08-27 修复（AH 底稿复查——1357 未分配利润 明细vs审定 符号差）：
                       #   SAP 未分配利润(4104)族 Balance=Credit（贷方正、借余负）——qm>0=盈利、
                       #   qm<0=亏损主体（如 1020 未分配利润 -1.1亿 借余亏损）。原审定表对 is_credit
                       #   一律 abs() 把亏损主体翻正（1357 审定 +13.42亿 vs 明细带符号 -9.85亿 虚增
                       #   23.27亿）。keep_sign=True → 审定表保持带符号（亏损显示负），集团合计=
                       #   Σ带符号=明细一级。实收/资本公积/盈余公积/专项储备 Balance=Debit（贷余负）
                       #   不传 keep_sign，仍 abs 显示贷方正（978M 等）。与明细表 flip_neg 同口径。
                       keep_sign=(g['key'] == 'RETAINED'))]
        # 2026-08-03 修复：必须传 target_year=year——否则 eff_years=全部年份，
        # BS 审定表把 2025+2026 期初/期末混年累加（FY智能 未分配利润曾显示 2026 期初 6.7M）
        _add_audit(wb, data_dir, _items, tb_full=tb, entities=ents, target_year=year)
    except Exception as ex:
        print(f"  ⚠️ 审定表注入失败 [{g['key']}][{year}]：{ex}")
    # 盈余公积：提取盈余公积分录属权益内部结转（非真实交易），不生成凭证抽查（2026-08-01 用户需求 #18）
    if g['key'] == 'SURRES':
        print('   [SKIP] 盈余公积凭证抽查（提取盈余公积属权益内部结转，用户要求不生成）')
    elif g['key'] == 'SPECRES':
        # ⚡ 2026-08-10 用户需求：专项储备变动=计提数（借费用贷专项储备），不抽凭但须【计提核对】
        try:
            build_zx_check_sheet(wb, year, recs, gl, gkey)
            print(f'[OK] {g["title"]}计提核对表 ({year})')
        except Exception as ex:
            print(f'  ⚠️ {g["title"]}计提核对表生成失败：{ex}')
    elif build_voucher_sheet(wb, year, recs, gl, gkey) is not None:
        print(f'[OK] {g["title"]}凭证抽查 ({year})')
    else:
        print(f'   [SKIP] {g["title"]}凭证抽查（该科目当年无凭证交易，跳过）')

    # 附注汇总（滚动四段：期初/本期增加/本期减少/期末；每主体一列，2026-08-02）
    try:
        from audit_common import build_footnote_generic as _bfn_eq
        _bfn_eq(wb, data_dir, '附注汇总',
                f'{g["title"]}附注汇总（审定口径；每主体一列；段=期初数/本期增加/本期减少/期末数）',
                [(g['title'], _codes)], is_credit=g['is_credit'], mode='four',
                target_year=year, tb_full=tb, entities=ents)
    except Exception as ex:
        print(f"  ⚠️ 附注汇总生成失败 [{g['key']}][{year}]：{ex}")

    # ⚡ 2026-08-11：挂 data_dir 供 finalize 统一注入（抽样配置账套覆盖等）
    wb._data_dir = data_dir
    # ⚡ 2026-08-28 #875：对方科目核对（模块内集成，内存 wb 注入，不 load+save）
    try:
        from counterparty_recon import inject_into_wb, SPECS as _CR_SPECS
        _ck = g.get('title')
        if _ck in _CR_SPECS:
            _rc = inject_into_wb(wb, gl, year, _CR_SPECS[_ck], ents=set(ents))
            print(f'[OK] {g["title"]}对方科目核对（{"已注入" if _rc == "ok" else "无数据" if _rc == "empty" else "跳过"}）')
    except Exception as _ex:
        print(f'  ⚠️ {g["title"]}对方科目核对注入失败：{_ex}')
    finalize_workbook(wb)
    out_path, _save_warn = _safe_save(wb, out_path)
    if _save_warn:
        print(f'  [WARN] {_save_warn}')
    if out_path is None:
        raise PermissionError(_save_warn or '保存失败：目标与副本均无法写入')
    wb.close()
    print(f'[DONE] 已生成：{out_path}')
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description='所有者权益审计底稿生成程序')
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


if __name__ == '__main__':
    _write_boot_log('MAIN')
    _bs_sys.exit(main())
