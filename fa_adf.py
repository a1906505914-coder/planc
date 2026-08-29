# -*- coding: utf-8 -*-
"""ADF 固定资产增减变动及折旧复核（2026-08-16 铁律131k v4，参照 GJX fa_movement）。
数据源：
  · 台账：ADF/数据/2026/固定资产台账.xlsx（SAP 资产主数据导出，Sheet1 约 5 万行/9 主体）
  · 账套：ADF/数据/2026/{主体}/{主体}科目余额表1-7月.xlsx（9 家，1601/1602 二级科目）
期间：2026 年 1-7 月（台账"本期"=1-7 月累计，TB"本期"=本年累计）。
方法（对齐 GJX 铁律131k v4）：
  · 仅取"固定资产-"分类（排除在建工程/无形资产/使用权资产/长期待摊）
  · 类别映射：台账 8 类 → 账套 TB 1601/1602 二级 5 类（房屋建筑物/机器设备/运输设备/电子设备/其他）
  · 折旧复算对齐 XBJ：月折旧=(期初原值-残值)/计划使用期间(月)；系统次月计提；本年 1-7 月应提；
    累计已提足则本期停提（封顶=剩余可提）
  · 底稿 5 sheets：①增减变动汇总（上中下三表：卡片台账/账套数/差异数，固定资产4列+累计折旧4列）
    ②折旧复算明细（主表：逐资产 27 列，含减少日期/减少方式）②b折旧反算补复算（原值增减且差异>5%资产，到期日反算）
    ③类别参数区间 ④口径说明
  · 格式对齐 XBJ：Times New Roman 10pt / 表头 DDEBF7 / thin 边框 / 金额 #,##0.00
输出：ADF/底稿/2026/固定资产增减变动复核_2026.xlsx
"""
import paths   # 数据路径中心锚点（数据绝对路径只允许出现在 paths.py）
import glob
import os
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DATA_ROOT = paths.DATA_ROOT   # 原 E 盘映射已失效，数据实际在 D 盘
ACCT = 'ADF'
YEAR = '2026'
# ⚡ 2026-08-17 期初取数修正：7月末台账「期初」列=7月初(6/30)数据，非年初 → 期初改用 12月末台账（2025.12.31）
LEDGER_END = os.path.join(DATA_ROOT, ACCT, '数据', str(YEAR), '7月固定资产台账.xlsx')
LEDGER_BEG_DIR = os.path.join(DATA_ROOT, ACCT, '数据', str(YEAR), '12月固定资产台账', '12月资产台账')
TB_DIR = os.path.join(DATA_ROOT, ACCT, '数据', str(YEAR))
COMPS = ['3100', '3200', '3300', '3400', '3500', '3700', '3900', '6100', '6900']
END_YM = (2026, 7)   # 期间截止 2026-07

# ⚡ 类别映射：台账 SAP 资产分类 → 账套 TB 1601/1602 二级（ADF 自己的数据核对确认）
CLS_MAP = {
    '固定资产-房屋': '房屋建筑物',
    '固定资产-构筑物': '房屋建筑物',
    '固定资产-机器设备': '机器设备',
    '固定资产-小型机器设备': '机器设备',
    '固定资产-运输设备': '运输设备',
    '固定资产-电子设备': '电子设备',
    '固定资产-其他': '其他',
    '固定资产-附属设施': '其他',
}
GL_CLASSES = ['房屋建筑物', '机器设备', '运输设备', '电子设备', '其他']
TB_CLASS = {'160101': '房屋建筑物', '160102': '机器设备', '160103': '运输设备',
            '160104': '电子设备', '160199': '其他'}
TB_CLASS_DEP = {'160201': '房屋建筑物', '160202': '机器设备', '160203': '运输设备',
                '160204': '电子设备', '160299': '其他'}


def _f(v):
    try:
        return float(v or 0)
    except (ValueError, TypeError):
        return 0.0


def _read_ledger_rows(fp):
    """读台账文件（12月末/7月末同构，SAP 资产主数据）→ 固定资产类行。
    返回 [{no,name,comp,cls,fa_cls,cap_date,dep_start,plan_m,used_m,residual,
    yz_end(原值期末), dep_end(累折期末), year_dep, month_dep, scrap_date, red_way, status}]"""
    wb = openpyxl.load_workbook(fp, read_only=True)
    ws = wb.worksheets[0]
    rows = []
    for r in ws.iter_rows(values_only=True):
        if not r[0] or r[0] == '资产号码':
            continue
        cl = str(r[5] or '')
        if not cl.startswith('固定资产-'):
            continue   # 排除在建/无形/使用权/长摊
        rows.append({
            'no': str(r[0]), 'name': str(r[1] or ''), 'comp': str(r[3] or ''),
            'cls': cl, 'fa_cls': CLS_MAP.get(cl, '其他'),
            'cap_date': str(r[18] or '')[:10], 'dep_start': str(r[56] or '')[:10],
            'plan_m': int(_f(r[51])), 'used_m': int(_f(r[52])), 'residual': _f(r[50]),
            'yz_end': _f(r[39]), 'dep_end': _f(r[45]),
            'year_dep': _f(r[57]), 'month_dep': _f(r[59]),
            'scrap_date': str(r[19] or '')[:10], 'red_way': str(r[32] or ''), 'status': str(r[29] or ''),
        })
    return rows


def load_ledger():
    """期初(12月末台账) + 期末(7月末台账) 跨期配对 → 每资产行。
    ⚡ 2026-08-17 修正：7月末台账「期初」列=7月初(6/30)数据，非年初 → 期初=12月末台账「期末」列（2025.12.31）。
    期初/期末按 (公司代码, 资产号码) 配对；期初无→新增、期末无→减少；存量原值增减=期末−期初净值。"""
    end_rows = _read_ledger_rows(LEDGER_END)
    begin = {}
    for fp in sorted(glob.glob(os.path.join(LEDGER_BEG_DIR, '*.XLSX'))):
        for rb in _read_ledger_rows(fp):
            begin[(rb['comp'], rb['no'])] = rb
    end_keys = set((re_['comp'], re_['no']) for re_ in end_rows)

    rows = []
    for re_ in end_rows:
        key = (re_['comp'], re_['no'])
        rb = begin.get(key)
        yz_beg = rb['yz_end'] if rb else 0.0
        dep_beg = rb['dep_end'] if rb else 0.0
        if rb:
            yz_inc = max(0.0, re_['yz_end'] - yz_beg)
            yz_dec = max(0.0, yz_beg - re_['yz_end'])
        else:
            yz_inc = re_['yz_end']          # 新增
            yz_dec = 0.0
        rows.append({
            'no': re_['no'], 'name': re_['name'], 'comp': re_['comp'],
            'cls': re_['cls'], 'fa_cls': re_['fa_cls'],
            'cap_date': re_['cap_date'], 'dep_start': re_['dep_start'],
            'plan_m': re_['plan_m'], 'used_m': re_['used_m'], 'residual': re_['residual'],
            'yz_begin': yz_beg, 'yz_inc': yz_inc, 'yz_dec': yz_dec, 'yz_end': re_['yz_end'],
            'dep_begin': dep_beg, 'dep_inc': 0.0, 'dep_dec': 0.0, 'dep_end': re_['dep_end'],
            'year_dep': re_['year_dep'], 'month_dep': re_['month_dep'],
            'scrap_date': re_['scrap_date'], 'red_way': re_['red_way'], 'status': re_['status'],
        })
    # 减少资产（期初有、期末无）
    for key, rb in begin.items():
        if key in end_keys:
            continue
        rows.append({
            'no': rb['no'], 'name': rb['name'], 'comp': rb['comp'],
            'cls': rb['cls'], 'fa_cls': rb['fa_cls'],
            'cap_date': rb['cap_date'], 'dep_start': rb['dep_start'],
            'plan_m': rb['plan_m'], 'used_m': rb['used_m'], 'residual': rb['residual'],
            'yz_begin': rb['yz_end'], 'yz_inc': 0.0, 'yz_dec': rb['yz_end'], 'yz_end': 0.0,
            'dep_begin': rb['dep_end'], 'dep_inc': 0.0, 'dep_dec': rb['dep_end'], 'dep_end': 0.0,
            'year_dep': 0.0, 'month_dep': 0.0,
            'scrap_date': rb['scrap_date'], 'red_way': rb['red_way'], 'status': rb['status'],
        })
    return rows


def load_tb_all():
    """TB 1601/1602 二级：合并 out（{fa_cls: {...}}）+ 按主体 tb_by_comp（{comp: {fa_cls: {...}}}）。
    列：0公司 1总账 2长文本 3期初方向 4期末方向 5期初金额 6本期借方 7本期贷方 10期末余额
    ⚡ SAP 累计折旧期初/期末为负（贷方）→ 取绝对值；借=转出、贷=计提"""
    def _new():
        return {'begin': 0.0, 'inc': 0.0, 'dec': 0.0, 'end': 0.0,
                'dep_begin': 0.0, 'dep_inc': 0.0, 'dep_dec': 0.0, 'dep_end': 0.0}
    out = defaultdict(_new)
    tb_by_comp = {}
    for comp in COMPS:
        # 文件名两种命名：{comp}科目余额表1-7月.xlsx / {comp}科目余额表20260731.XLSX
        fp = None
        for cand in (os.path.join(TB_DIR, comp, f'{comp}科目余额表1-7月.xlsx'),
                     os.path.join(TB_DIR, comp, f'{comp}科目余额表20260731.XLSX'),
                     os.path.join(TB_DIR, comp, f'{comp}科目余额表20260731.xlsx')):
            if os.path.exists(cand):
                fp = cand
                break
        if not fp:
            print(f'⚠️ 缺 TB: {comp}')
            continue
        wb = openpyxl.load_workbook(fp, read_only=True)
        ws = wb.worksheets[0]
        rows_ = list(ws.iter_rows(values_only=True))
        if not rows_:
            continue
        # ⚡ 按表头列名动态定位（3100 版与 3900 版列序不同）
        hdr = {str(h or '').strip(): i for i, h in enumerate(rows_[0])}
        ci = {k: hdr.get(k, -1) for k in ('期初金额', '本期借方金额', '本期贷方金额', '期末余额')}
        sub = defaultdict(_new)   # 该主体
        for r in rows_[1:]:
            if ci['期初金额'] < 0 or ci['期末余额'] < 0:
                continue
            code = str(r[1] or '')
            try:
                v_begin = _f(r[ci['期初金额']]); v_jf = _f(r[ci['本期借方金额']])
                v_df = _f(r[ci['本期贷方金额']]); v_end = _f(r[ci['期末余额']])
            except (ValueError, TypeError, IndexError):
                continue
            c6 = code[:6]   # ⚡ TB 科目代码 10 位（1601010000）→ 取前 6 位匹配二级
            if code.startswith('16019999'):
                continue    # ⚡ 排除"资产过账清算"（1601999900，内部清算借=贷，非固定资产）
            if c6 in TB_CLASS:
                cl = TB_CLASS[c6]
                v = out[cl]; v['begin'] += v_begin; v['inc'] += v_jf; v['dec'] += v_df; v['end'] += v_end
                sv = sub[cl]; sv['begin'] += v_begin; sv['inc'] += v_jf; sv['dec'] += v_df; sv['end'] += v_end
            elif c6 in TB_CLASS_DEP:
                cl = TB_CLASS_DEP[c6]
                v = out[cl]; v['dep_begin'] += abs(v_begin); v['dep_inc'] += v_df; v['dep_dec'] += v_jf; v['dep_end'] += abs(v_end)
                sv = sub[cl]; sv['dep_begin'] += abs(v_begin); sv['dep_inc'] += v_df; sv['dep_dec'] += v_jf; sv['dep_end'] += abs(v_end)
        tb_by_comp[comp] = sub
    return out, tb_by_comp


def _months_2026(dep_start, plan_m, dep_begin, yz_begin, residual):
    """2026 年 1-7 月应提月数（对齐 XBJ：折旧开始次月计提；此前已开始=满 7 月；累计已提足则 0；封顶=剩余可提）。
    返回 (应提月数, 月折旧额, 剩余可提)"""
    if plan_m <= 0:
        return 0, 0.0, 0.0
    base = max(yz_begin - residual, 0.0)
    mdep = base / plan_m
    remain = base - dep_begin                    # 剩余可提
    if remain <= 0:
        return 0, mdep, remain                   # ⚡ 累计已提足 → 本期停提
    m = re.search(r'(\d{4})-(\d{1,2})', dep_start)
    if not m:
        return 7, mdep, remain
    sy, s_m = int(m.group(1)), int(m.group(2))
    if sy > END_YM[0] or (sy == END_YM[0] and s_m > END_YM[1]):
        return 0, mdep, remain
    if sy == END_YM[0]:
        acc = END_YM[1] - s_m + 1                # 年内启用：开始月至 7 月
        acc = max(0, acc)
    else:
        acc = 7                                  # 此前已启用：1-7 月
    return max(0, min(acc, plan_m)), mdep, remain


def main():
    rows = load_ledger()
    tb, tb_by_comp = load_tb_all()
    n_fa = len(rows)
    print(f'台账固定资产类: {n_fa} 行（9 主体）', flush=True)
    # 非固定资产分类统计（口径说明用）
    wb0 = openpyxl.load_workbook(LEDGER_END, read_only=True)
    ws0 = wb0['Sheet1']
    other_cls = Counter()
    for r in ws0.iter_rows(values_only=True):
        if not r[0] or r[0] == '资产号码':
            continue
        cl = str(r[5] or '')
        if not cl.startswith('固定资产-'):
            other_cls[cl] += 1
    n_other = sum(other_cls.values())

    # ===== 折旧复算（对齐 XBJ）=====
    for x in rows:
        acc_m, mdep, remain = _months_2026(x['dep_start'], x['plan_m'], x['dep_begin'], x['yz_begin'], x['residual'])
        x['acc_m'] = acc_m
        x['mdep_calc'] = mdep
        x['dep_year_calc'] = max(0.0, min(mdep * acc_m, max(remain, 0.0)))   # 匡算本年（提足封顶）
        x['diff_year'] = x['dep_year_calc'] - x['year_dep']                  # 本年差异
        x['exp_dep'] = min(x['dep_begin'] + x['dep_year_calc'],
                           max(x['yz_begin'] - x['residual'], 0.0))          # 匡算期末累计（封顶）
        x['diff_cum'] = x['exp_dep'] - x['dep_end']                          # 累计差异

    # ⚡ 明细只列审计关注行（本期有增加/有减少/报废/本年或累计差异>5%），其余存量无变动聚合——防 5 万行写盘 OOM
    rows_disp = []
    for x in rows:
        if x['yz_inc'] > 0 or x['yz_dec'] > 0 or x['scrap_date']:
            rows_disp.append(x)
        elif x['diff_year'] and abs(x['diff_year']) / max(x['year_dep'], 1) > 0.05:
            rows_disp.append(x)
        elif x['diff_cum'] and abs(x['diff_cum']) / max(x['dep_end'], 1) > 0.05:
            rows_disp.append(x)
    n_skip = len(rows) - len(rows_disp)
    print(f'明细关注行 {len(rows_disp)} | 存量无变动聚合 {n_skip} 行', flush=True)

    # ===== 汇总（三表用）=====
    def _tot(rows_, key):
        return sum(x[key] for x in rows_)

    tot = {
        'begin': _tot(rows, 'yz_begin'), 'inc': _tot(rows, 'yz_inc'), 'dec': _tot(rows, 'yz_dec'),
        'end': _tot(rows, 'yz_end'),
        'dep_begin': _tot(rows, 'dep_begin'), 'dep_end': _tot(rows, 'dep_end'),
        'dec_dep': sum(x['dep_begin'] for x in rows if x['yz_dec'] > 0 or x['scrap_date']),  # 转出≈减少资产折旧期初
    }
    tot['dep_inc'] = tot['dep_end'] - tot['dep_begin'] + tot['dec_dep']       # 计提=期末−期初+转出
    # 类别级
    s = {}
    for cl in GL_CLASSES:
        rs = [x for x in rows if x['fa_cls'] == cl]
        dec_dep = sum(x['dep_begin'] for x in rs if x['yz_dec'] > 0 or x['scrap_date'])
        s[cl] = {
            'begin': sum(x['yz_begin'] for x in rs), 'inc': sum(x['yz_inc'] for x in rs),
            'dec': sum(x['yz_dec'] for x in rs), 'end': sum(x['yz_end'] for x in rs),
            'dep_begin': sum(x['dep_begin'] for x in rs), 'dep_end': sum(x['dep_end'] for x in rs),
            'dec_dep': dec_dep,
        }
        s[cl]['dep_inc'] = s[cl]['dep_end'] - s[cl]['dep_begin'] + s[cl]['dec_dep']

    # ⚡ 按主体汇总（台账侧 8 列 + 账套侧 8 列）——定位差异主体
    card_by_comp = {}
    for x in rows:
        c = card_by_comp.setdefault(x['comp'], {'begin': 0.0, 'inc': 0.0, 'dec': 0.0, 'end': 0.0,
                                                'dep_begin': 0.0, 'dep_inc': 0.0, 'dec_dep': 0.0, 'dep_end': 0.0})
        c['begin'] += x['yz_begin']; c['inc'] += x['yz_inc']; c['dec'] += x['yz_dec']; c['end'] += x['yz_end']
        c['dep_begin'] += x['dep_begin']; c['dep_end'] += x['dep_end']
        if x['yz_dec'] > 0 or x['scrap_date']:
            c['dec_dep'] += x['dep_begin']
    for c in card_by_comp.values():
        c['dep_inc'] = c['dep_end'] - c['dep_begin'] + c['dec_dep']

    def _tb_comp_agg(comp):
        t = tb_by_comp.get(comp, {})
        r = {'begin': 0.0, 'inc': 0.0, 'dec': 0.0, 'end': 0.0,
             'dep_begin': 0.0, 'dep_inc': 0.0, 'dep_dec': 0.0, 'dep_end': 0.0}
        for cl in GL_CLASSES:
            v = t.get(cl)
            if not v:
                continue
            for k in ('begin', 'inc', 'dec', 'end', 'dep_begin', 'dep_inc', 'dep_dec', 'dep_end'):
                r[k] += v[k]
        return r

    # ===== 写底稿 =====
    wb = openpyxl.Workbook()
    base_font = Font(name='Times New Roman', size=10)
    hdr_font = Font(name='Times New Roman', size=10, bold=True)
    total_font = Font(name='Times New Roman', size=10, bold=True)
    hdr_fill = PatternFill('solid', fgColor='DDEBF7')
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    num_fmt = '#,##0.00'
    CTR = Alignment(horizontal='center', vertical='center')

    def style_header(ws, row=1):
        for c in ws[row]:
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = CTR
            c.border = thin

    def style_cell(c, numeric=False, bold=False):
        c.font = total_font if bold else base_font
        c.border = thin
        if numeric:
            c.number_format = num_fmt

    # ① 增减变动汇总（上中下三表）
    ws = wb.active
    ws.title = '增减变动汇总'
    ws.append(['ADF 固定资产增减变动汇总（2026 年 1-7 月，全集团 9 主体合并）'])
    ws.append([])
    HDR8 = ['类别', '固定资产-期初原值', '固定资产-本期增加', '固定资产-本期减少', '固定资产-期末原值',
            '累计折旧-期初', '累计折旧-本期计提', '累计折旧-本期转出', '累计折旧-期末']

    def _dump_block(title, get8):
        ws.append([title])
        ws.append(HDR8)
        style_header(ws, ws.max_row)
        for cl in GL_CLASSES:
            v = get8(cl)
            if v is None:
                continue
            r_ = ws.max_row + 1
            ws.append([cl, v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7]])
            for ci in range(1, 9):
                style_cell(ws.cell(r_, ci + 1), numeric=True)
            style_cell(ws.cell(r_, 1))
        v = get8(None)
        r_ = ws.max_row + 1
        ws.append(['合计', v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7]])
        for c in ws[r_]:
            style_cell(c, numeric=(c.column > 1), bold=True)
        ws.append([])

    def card8(cl):
        if cl is None:
            return (tot['begin'], tot['inc'], tot['dec'], tot['end'],
                    tot['dep_begin'], tot['dep_inc'], tot['dec_dep'], tot['dep_end'])
        v = s[cl]
        return (v['begin'], v['inc'], v['dec'], v['end'],
                v['dep_begin'], v['dep_inc'], v['dec_dep'], v['dep_end'])

    def tb8(cl):
        if cl is None:
            t = tb
            return (sum(t[c]['begin'] for c in GL_CLASSES), sum(t[c]['inc'] for c in GL_CLASSES),
                    sum(t[c]['dec'] for c in GL_CLASSES), sum(t[c]['end'] for c in GL_CLASSES),
                    sum(t[c]['dep_begin'] for c in GL_CLASSES), sum(t[c]['dep_inc'] for c in GL_CLASSES),
                    sum(t[c]['dep_dec'] for c in GL_CLASSES), sum(t[c]['dep_end'] for c in GL_CLASSES))
        v = tb.get(cl)
        if not v:
            return None
        return (v['begin'], v['inc'], v['dec'], v['end'],
                v['dep_begin'], v['dep_inc'], v['dep_dec'], v['dep_end'])

    _dump_block('一、卡片台账（SAP 资产主数据，Σ类别）', card8)
    _dump_block('二、账套数（科目余额表 1601/1602 二级科目，9 家合计）', tb8)

    def diff8(cl):
        c8 = card8(cl)
        t8 = tb8(cl)
        if t8 is None:
            return c8 if cl is None else None
        return tuple(c8[i] - t8[i] for i in range(8))

    _dump_block('三、差异数（卡片台账 − 账套，差异=待查因）', diff8)

    # ④ 按主体核对（台账 vs 账套，各主体）——定位差异主体
    def _dump_comp_block(title, get8):
        ws.append([title])
        ws.append(['主体'] + HDR8[1:])
        style_header(ws, ws.max_row)
        for comp in COMPS:
            v = get8(comp)
            if v is None:
                continue
            r_ = ws.max_row + 1
            ws.append([comp, v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7]])
            for ci in range(1, 9):
                style_cell(ws.cell(r_, ci + 1), numeric=True)
            style_cell(ws.cell(r_, 1))
        v = get8(None)
        r_ = ws.max_row + 1
        ws.append(['合计', v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7]])
        for c in ws[r_]:
            style_cell(c, numeric=(c.column > 1), bold=True)
        ws.append([])

    def card_comp8(comp):
        if comp is None:
            return tuple(sum(c[k] for c in card_by_comp.values())
                         for k in ('begin', 'inc', 'dec', 'end', 'dep_begin', 'dep_inc', 'dec_dep', 'dep_end'))
        c = card_by_comp.get(comp)
        if not c:
            return (0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
        return (c['begin'], c['inc'], c['dec'], c['end'],
                c['dep_begin'], c['dep_inc'], c['dec_dep'], c['dep_end'])

    def tb_comp8(comp):
        if comp is None:
            return tuple(sum(_tb_comp_agg(c)[k] for c in COMPS)
                         for k in ('begin', 'inc', 'dec', 'end', 'dep_begin', 'dep_inc', 'dep_dec', 'dep_end'))
        c = _tb_comp_agg(comp)
        return (c['begin'], c['inc'], c['dec'], c['end'],
                c['dep_begin'], c['dep_inc'], c['dep_dec'], c['dep_end'])

    def diff_comp8(comp):
        c8 = card_comp8(comp)
        t8 = tb_comp8(comp)
        return tuple(c8[i] - t8[i] for i in range(8))

    _dump_comp_block('四、按主体核对·一 卡片台账（SAP 资产主数据，按公司代码聚合）', card_comp8)
    _dump_comp_block('四、按主体核对·二 账套数（科目余额表 1601/1602，各主体）', tb_comp8)
    _dump_comp_block('四、按主体核对·三 差异数（台账 − 账套，定位差异主体）', diff_comp8)

    r_ = ws.max_row + 1
    ws.append(['勾稽：期末原值 = 期初+增加−减少；累计折旧期末 = 期初+计提−转出（两表各自成立）'])
    style_cell(ws.cell(r_, 1))

    # ② 折旧复算明细（主表，逐资产 27 列；仅审计关注行）
    ws2 = wb.create_sheet('折旧复算明细')
    HDR2 = ['序号', '资产号码', '公司代码', '资产名称', '底稿类别', '资本化日期', '计划使用期间(月)', '已使用期间', '残值率',
            '固定资产-期初原值', '固定资产-本期增加', '固定资产-本期减少', '固定资产-期末原值',
            '累计折旧-期初', '累计折旧-本期计提', '累计折旧-本期转出', '累计折旧-期末',
            '月折旧额(计算)', '本月折旧(账面)', '匡算本年折旧', '本年折旧差异',
            '匡算期末累计折旧', '累计折旧差异', '减少日期', '减少方式', '使用状态', '备注']
    ws2.append(HDR2)
    style_header(ws2)
    NUM2 = {9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22}
    for i, x in enumerate(rows_disp, 1):
        r_ = ws2.max_row + 1
        residual_rate = x['residual'] / x['yz_begin'] if x['yz_begin'] else 0.0
        note = ''
        if x['diff_year'] and abs(x['diff_year']) / max(x['year_dep'], 1) > 0.05:
            note = '本年折旧差异率>5%（年限/残值/折旧码与账面不一致）'
        ws2.append([i, x['no'], x['comp'], x['name'], x['fa_cls'], x['cap_date'], x['plan_m'], x['used_m'],
                    residual_rate,
                    x['yz_begin'], x['yz_inc'], x['yz_dec'], x['yz_end'],
                    x['dep_begin'], x['dep_end'] - x['dep_begin'] + (x['dep_begin'] if x['yz_dec'] > 0 or x['scrap_date'] else 0.0),
                    (x['dep_begin'] if x['yz_dec'] > 0 or x['scrap_date'] else 0.0), x['dep_end'],
                    x['mdep_calc'], x['month_dep'], x['dep_year_calc'], x['diff_year'],
                    x['exp_dep'], x['diff_cum'],
                    x['scrap_date'], x['red_way'], x['status'], note])
        for ci in range(len(HDR2)):
            style_cell(ws2.cell(r_, ci + 1), numeric=(ci in NUM2), bold=False)
    # 合计（全部数值列：固定资产4列+累计折旧4列+复算列；⚡ 表头 27 列数值从 col10 起——需 9 个占位防错位）
    r_ = ws2.max_row + 1
    ws2.append(['合计', '', '', '', '', '', '', '', '',
                tot['begin'], tot['inc'], tot['dec'], tot['end'],
                tot['dep_begin'], tot['dep_inc'], tot['dec_dep'], tot['dep_end'],
                sum(x['mdep_calc'] for x in rows), sum(x['month_dep'] for x in rows),
                sum(x['dep_year_calc'] for x in rows), sum(x['diff_year'] for x in rows),
                sum(x['exp_dep'] for x in rows), sum(x['diff_cum'] for x in rows)])
    for ci in range(len(HDR2)):
        style_cell(ws2.cell(r_, ci + 1), numeric=(ci in (9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22)), bold=True)
    r_ = ws2.max_row + 1
    ws2.append(['与账套核对(Σ台账 vs TB 1601/1602, 详细见「增减变动汇总」)', '', '', '', '', '', '', '', '',
                tot['begin'], tot['inc'], tot['dec'], tot['end'],
                tot['dep_begin'], tot['dep_inc'], tot['dec_dep'], tot['dep_end']])
    for ci in range(len(HDR2)):
        style_cell(ws2.cell(r_, ci + 1), numeric=(ci in (9, 10, 11, 12, 13, 14, 15, 16)), bold=True)
    ws2.append([])
    n_big = sum(1 for x in rows_disp if x['diff_year'] and abs(x['diff_year']) / max(x['year_dep'], 1) > 0.05)
    r_ = ws2.max_row + 1
    ws2.append([f'明细共 {len(rows_disp)} 行（本期有增加/减少/报废 或 折旧差异率>5%）；其余 {n_skip} 行存量无变动，汇总见「增减变动汇总」'])
    style_cell(ws2.cell(r_, 1))

    # ②b 折旧反算补复算（原值在使用过程中增减且正向差异率>5% 的资产：从到期日反算）
    #     反算月折旧 = 剔除残值后剩余净值 ÷ 剩余月数（期末原值−残值−累计折旧期末 除以 计划使用月限−已使用期间），
    #     消除"正向按期初/期末原值均摊"对中途增减资产的大额差异。
    ws2b = wb.create_sheet('折旧反算补复算')
    HDR2B = ['序号', '资产号码', '公司代码', '资产名称', '底稿类别', '资本化日期', '使用月限(月)', '已使用期间', '剩余月数', '残值额',
             '固定资产-期初原值', '本期增加', '本期减少', '固定资产-期末原值',
             '累计折旧-期末', '账面净值', '剩余可提(剔残值)', '反算月折旧', '匡算本年折旧(反算)', '反算年末累计',
             '累计差异(反算)', '累计差异(正向)', '差异改善', '备注']
    ws2b.append(HDR2B)
    style_header(ws2b)
    rev_rows = []
    for x in rows:
        if x['yz_end'] <= 0:
            continue                       # 减少/处置（期末原值0）不计提
        if not (x['yz_inc'] > 0 or x['yz_dec'] > 0):
            continue                       # 本期无原值增减（非用户关注）
        big = ((x['diff_year'] and abs(x['diff_year']) / max(x['year_dep'], 1) > 0.05)
               or (x['diff_cum'] and abs(x['diff_cum']) / max(x['dep_end'], 1) > 0.05))
        if not big:
            continue                       # 正向差异不大，无需反算
        base_new = max(x['yz_end'] - x['residual'], 0.0)      # 期末原值剔除残值后应提总额
        remain_end = max(base_new - x['dep_end'], 0.0)        # 期末剩余可提
        left_m = max(x['plan_m'] - x['used_m'], 0)            # 剩余月数（到期日−已使用）
        mdep_rev = remain_end / left_m if left_m > 0 else 0.0
        dep_year_rev = max(0.0, min(mdep_rev * x['acc_m'], remain_end))
        exp_dep_rev = min(x['dep_begin'] + dep_year_rev, base_new)
        diff_cum_rev = exp_dep_rev - x['dep_end']
        improve = abs(x['diff_cum'] or 0) - abs(diff_cum_rev)
        if improve > 0.01:
            note = '反算后差异缩小'
        elif improve < -0.01:
            note = '反算差异反增（折旧开始日/年限字段可能失真，需人工复核）'
        else:
            note = ''
        rev_rows.append((x, base_new, remain_end, left_m, mdep_rev, dep_year_rev,
                         exp_dep_rev, diff_cum_rev, improve, note))
    NUM2B = {9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22}
    for i, t in enumerate(rev_rows, 1):
        x = t[0]
        r_ = ws2b.max_row + 1
        ws2b.append([i, x['no'], x['comp'], x['name'], x['fa_cls'], x['cap_date'], x['plan_m'], x['used_m'],
                     t[3], round(x['residual'], 2),
                     x['yz_begin'], x['yz_inc'], x['yz_dec'], x['yz_end'],
                     x['dep_end'], round(x['yz_end'] - x['dep_end'], 2),
                     round(t[1], 2), round(t[4], 2), round(t[5], 2), round(t[6], 2),
                     round(t[7], 2), round(x['diff_cum'] or 0, 2), round(t[8], 2), t[9]])
        for ci in range(len(HDR2B)):
            style_cell(ws2b.cell(r_, ci + 1), numeric=(ci in NUM2B), bold=False)
    r_ = ws2b.max_row + 1
    ws2b.append([f'共 {len(rev_rows)} 行：本期原值增减且正向复算差异率>5%，改按【到期日反算】补复算（反算月折旧=剔除残值后剩余净值÷剩余月数）；'
                 f'差异改善=正向累计差异绝对值−反算累计差异绝对值（正值=反算缩小差异）'])
    style_cell(ws2b.cell(r_, 1))

    # ③ 类别参数区间
    ws3 = wb.create_sheet('类别参数区间')
    ws3.append(['类别', '行数', '计划使用期间min(月)', '计划使用期间max(月)', '计划使用期间常见(月)',
                '残值率min', '残值率max', '残值率常见'])
    style_header(ws3)
    for cl in GL_CLASSES:
        rs = [x for x in rows if x['fa_cls'] == cl]
        if not rs:
            continue
        plan_l = [x['plan_m'] for x in rs if x['plan_m'] > 0]
        res_l = [(x['residual'] / x['yz_begin'] if x['yz_begin'] else 0.0) for x in rs if x['residual'] > 0 and x['yz_begin'] > 0]
        mcom = Counter(plan_l).most_common(1)[0][0] if plan_l else 0
        rcom = Counter(round(r, 4) for r in res_l).most_common(1)[0][0] if res_l else 0
        r_ = ws3.max_row + 1
        ws3.append([cl, len(rs), min(plan_l) if plan_l else 0, max(plan_l) if plan_l else 0, mcom,
                    min(res_l) if res_l else 0, max(res_l) if res_l else 0, rcom])
        for ci in range(1, 8):
            style_cell(ws3.cell(r_, ci + 1), numeric=True)
        style_cell(ws3.cell(r_, 1))
    r_ = ws3.max_row + 1
    ws3.append([])
    r_ = ws3.max_row + 1
    ws3.append(['注：与公司折旧政策核对（年限/残值率区间）；异常（年限超政策、残值率口径不一）需人工复核'])
    style_cell(ws3.cell(r_, 1))

    # ④ 口径说明
    ws4 = wb.create_sheet('口径说明')
    notes = [
        'ADF 固定资产增减变动复核（2026 年 1-7 月，全集团 9 主体合并）— 2026-08-16 铁律131k v4',
        '数据源：',
        f'  ① 台账：期末 {LEDGER_END} + 期初 {LEDGER_BEG_DIR}（SAP 资产主数据；固定资产类 {n_fa} 行，主体：3100/3200/3300/3400/3500/3900/6100/6900，3700 无固定资产）',
        f'  ② 账套：9 家科目余额表 1-7 月（1601/1602 二级科目），期间=2026 年 1-7 月',
        f'  ③ 已排除非固定资产分类 {n_other} 行：在建工程（新建/技改/零星/审计调整）、无形资产（土地/软件/专利/非专利）、使用权资产、长期待摊',
        '⚡ 类别映射（台账 SAP 资产分类 8 类 → 账套 TB 1601/1602 二级 5 类）：',
        '  房屋/构筑物→房屋建筑物；机器设备/小型机器设备→机器设备；运输设备→运输设备；电子设备→电子设备；其他/附属设施→其他。',
        '⚡ 结构说明（对齐 GJX 铁律131k v4）：',
        '  ① 增减变动汇总：上中下三表（卡片台账/账套数/差异数），固定资产 4 列（期初/增/减/末）+ 累计折旧 4 列（期初/计提/转出/末）不交错；',
        '  ② 折旧复算明细=主表：逐资产 27 列，增减 4 列 + 折旧 4 列 + 复算双轨 + 减少日期/减少方式（台账字段）；合计/核对仅数值列；',
        '  ②b 折旧反算补复算：原值在使用过程中存在增减且正向复算差异率>5% 的资产，另行按【到期日反算】补复算——',
        '     反算月折旧=剔除残值后剩余净值÷剩余月数（(期末原值−残值−累计折旧期末) ÷ (计划使用月限−已使用期间)），消除正向按期初原值均摊对中途增减资产的大额差异；差异改善=正向累计差异绝对值−反算累计差异绝对值（正值=反算缩小差异）。',
        '  ③ 类别参数区间（折旧政策核对）；④ 口径说明。',
        '折旧复算（对齐 XBJ）：月折旧=(期初原值−残值)÷计划使用期间(月)；系统次月计提；2026 年 1-7 月应提（年内启用=开始月至 7 月，此前启用=7 个月）；',
        '  累计已提足（期初累计折旧≥原值−残值）则本期停提；匡算本年=min(月折旧×应提月数, 剩余可提)。',
        '与账套核对：台账期初/增/减/末 vs TB 1601；累计折旧期初/计提/转出/末 vs TB 1602（SAP TB 贷方科目为负→取绝对值，借=转出、贷=计提）。',
        '⚡ 按主体核对（四）：三表同构按公司代码列示（台账 SAP 资产主数据按「公司代码」聚合，账套取各主体科目余额表）。',
        '  ⚠️ 台账公司代码仅 8 家（3100/3200/3300/3400/3500/3900/6100/6900），无 3700；3700 科目余额表亦无 1601/1602 科目（3700 无固定资产，非数据缺口）。',
        '  ⚡ 原值/累计折旧【期初·期末】各主体均 0 差（2026-08-17 修正：期初改用 12月末台账）；发生额差异=等额对冲的口径差异，详见下方查因。',
        '⚠️ 差异=待查因（若期末 0 差而期初/发生额有差，多为台账与账套期初口径/调整差异）。',
        '⚡⚡ 期初取数修正 + 发生额差异查因（2026-08-17）：',
        '  ⚡ 期初取数修正：原用 7月末台账「期初」列（=7月初 6/30 数据）非年初 → 期初固资/折旧比 TB 多 5,494 万；',
        '  已改为【12月末台账（2025.12.31）】作期初，按(公司代码,资产号码)跨期配对（期初 48,117 张 → 期末 48,389 张，新增 273 / 减少 1）；',
        '  修正后原值/累计折旧【期初·期末】全部 0 差（期初 21,652,675,340 = TB1601 期初，折旧期初 12,035,412,472 = TB1602 期初）。',
        '  发生额差异（等额对冲、净额 0）：本期原值增/减各差 3,582 万（卡片按净值差额 vs TB 全额发生额：资产拆分/调拨/清理等内部调整 TB 记增记减、卡片只记净值）；',
        '  折旧计提/转出各差 4,568 万（转出口径：卡片按减少资产期初累折全额转出，TB 借方仅记部分处置）——均为口径差异，需进一步核对。',
        '⚡⚡ GL 借方分录构成验证（2026-08-16，7 家有序时账主体 GL 筛固定资产借方 451 笔）：',
        '  TB 借发(1.24亿) = 真新增（在建工程转固≈3,600万：UF-L1/L4/L5螺杆机技改、注射系统、涤纶阳离子装置改造；',
        '  采购订单入库≈1,135万）+ 内部调整（资产拆分≈6,485万、资产清理≈536万、资产调拨≈440万、零星补单≈200万，一增一减净额 0）',
        '  → 台账「2026资本化」5,571万 ≈ 真新增部分；约 6,800 万差额=资产拆分/调拨等内部调整在 TB 借发的体现（台账不记）。',
        '⚡ 审计关注：本年折旧差异率>5% 的资产（年限/残值率/折旧码与账面计提不一致）；报废日期/减少方式完整性。',
    ]
    for i, n in enumerate(notes, 1):
        ws4.cell(row=i, column=1, value=n).font = base_font
    ws4.column_dimensions['A'].width = 130

    # 列宽
    for i, w in enumerate([16, 18, 16, 16, 16, 18, 18, 16, 16], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for i, w in enumerate([6, 14, 10, 24, 12, 12, 12, 10, 9,
                           16, 16, 16, 16, 16, 16, 16, 16,
                           14, 14, 14, 14, 16, 14, 12, 14, 10, 26], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for wsn in (ws, ws2, ws3):
        wsn.freeze_panes = 'A2'

    out_dir = os.path.join(DATA_ROOT, ACCT, '底稿', str(YEAR))
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f'固定资产增减变动复核_{YEAR}.xlsx')
    wb.save(out)
    print(f'✅ 已生成: {out}')
    print(f'  台账固定资产 {n_fa} 行 | 原值期初 {tot["begin"]:,.0f} 增 {tot["inc"]:,.0f} 减 {tot["dec"]:,.0f} 末 {tot["end"]:,.0f}')
    print(f'  折旧: 期初 {tot["dep_begin"]:,.0f} 计提 {tot["dep_inc"]:,.0f} 转出 {tot["dec_dep"]:,.0f} 末 {tot["dep_end"]:,.0f}')
    tb1 = sum(tb[c]['end'] for c in GL_CLASSES)
    tb2 = sum(tb[c]['dep_end'] for c in GL_CLASSES)
    print(f'  核对: 原值期末 卡{tot["end"]:,.0f} vs TB {tb1:,.0f} 差{tot["end"]-tb1:,.0f} | 折旧期末 卡{tot["dep_end"]:,.0f} vs TB {tb2:,.0f} 差{tot["dep_end"]-tb2:,.0f}')
    print(f'  计提: 卡{tot["dep_inc"]:,.0f} vs TB {sum(tb[c]["dep_inc"] for c in GL_CLASSES):,.0f} | 转出: 卡{tot["dec_dep"]:,.0f} vs TB {sum(tb[c]["dep_dec"] for c in GL_CLASSES):,.0f}')
    return out


if __name__ == '__main__':
    sys.exit(main())
