# -*- coding: utf-8 -*-
"""GJX 固定资产增减变动及折旧复核（2026-08-16 铁律131k，v3 底稿结构重构 19:10 用户定调）。
数据源：
  · 卡片台账（3 期间）：数据/2026/固定资产卡片/截止{2024/2025}年12月31日、2026年3月31日资产卡片.xls
  · 账套数据（GJX 自己的 GL/TB）：数据/2025/DQ2025综合查询明细表、DQ2025科目余额表
期间：底稿只出 25 全年 + 26 年 1-3 月；24 年末卡片 = 25 期初存量。
方法（用户 16:45/17:49/18:01/18:43/18:47/19:09-19:10 定调）：
  · 卡片唯一键 = 名称+原值（规避编号跨年重排误判减少）；跨期差集用多重集（同名同值多张）
  · 类别：GJX 卡片类别（8 类）→ GJX 账套 GL 二级（5 类+运输单列），由数据核对确认
  · 折旧复算对齐 XBJ（_total_months/_months_until：系统次月计提、年内启用次月至期末、累计已提足则 0）
  · 减少日期：GL 固定资产清理/1601贷方分录匹配（名称关键词+原值金额<2%），匹配不到标『待人工补』
  · ⚡ 底稿结构（v3，用户 19:09/19:10 定调）：
      ① 增减变动汇总：类别 × 【固定资产4列：期初/本期增加/本期减少/期末】+【累计折旧4列：期初/本期计提/本期转出/期末】不交错
         + 底部【与账套核对】区：卡片Σ vs TB 1601/1602（期初/借发贷发/期末），本期计提=期末−期初+转出（勾稽）
      ② 折旧复算明细（主表）：逐卡片（存量/新增/减少全覆盖）固定资产增减4列 + 累计折旧增减4列
         + 折旧复算双轨（匡算本年/累计 vs 账面）+ 减少日期/GL凭证（从账套取）+ 底部合计 + 与账套核对
      ③ 类别参数区间（折旧政策核对）  ④ 口径说明
      （已删除：新增明细/减少明细/折旧核对/与TB核对/本期计提核对——信息并入①②）
输出：GJX/底稿/{period}/固定资产增减变动复核_{period}.xlsx
"""
import paths   # 数据路径中心锚点（数据绝对路径只允许出现在 paths.py）
import os
import re
import sys
from collections import Counter, defaultdict

import xlrd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

DATA_ROOT = paths.DATA_ROOT   # 原 E 盘映射已失效，数据实际在 D 盘
ACCT = 'GJX'
YEAR = '2026'
CARD_DIR = os.path.join(DATA_ROOT, ACCT, '数据', str(YEAR), '固定资产卡片')
GL25 = os.path.join(DATA_ROOT, ACCT, '数据', '2025', 'DQ2025综合查询明细表.xlsx')
GL26 = os.path.join(DATA_ROOT, ACCT, '数据', '2026', 'DQ2026综合查询明细表_20260811175803.xlsx')
TB25 = os.path.join(DATA_ROOT, ACCT, '数据', '2025', 'DQ2025科目余额表.xlsx')
TB26 = os.path.join(DATA_ROOT, ACCT, '数据', '2026', 'DQ2026科目余额表_20260811175710.xlsx')

# ⚡ 类别映射：GJX 卡片类别(8类) → GJX 账套 GL 二级(5类+运输单列)——数据核对确认（25年净增合计5712万=GL一致）
CARD_GL_CLASS = {
    '房屋建筑物': '房屋及建筑物',
    '专用设备': '专用设备',
    '通用设备': '通用设备',
    '电子仪器设备': '电子设备',
    '电子办公设备': '电子设备',
    '家具用具': '家具用具及其他',
    '工具用具': '家具用具及其他',
    '运输设备': '运输设备',
}
GL_CLASSES = ['房屋及建筑物', '专用设备', '通用设备', '电子设备', '家具用具及其他', '运输设备']


def _num(v):
    """数值容错（xls 单元格可能为 str 带逗号/百分号）。"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').replace('%', '').strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def load_cards(fp):
    """读卡片台账 → (聚合 cards, 逐卡列表 raw_cards)。
    cards = { (名称,原值): {'count','yz','lj','cls','cids','names','residual','months','dep_months','start','mdep_book'} }（多重集键，供跨期差集）
    raw_cards = 逐张卡片 [{cid,name,yz,lj,cls,residual,months,dep_months,start,mdep_book}]（供明细逐卡展开）"""
    wb = xlrd.open_workbook(fp)
    ws = wb.sheet_by_index(0)
    cards = defaultdict(lambda: {'count': 0, 'yz': 0.0, 'lj': 0.0, 'cls': '', 'cids': [], 'names': [],
                                 'residual': None, 'months': 0, 'dep_months': 0, 'start': '', 'mdep_book': 0.0})
    raw_cards = []
    for i in range(1, ws.nrows):
        cid = str(ws.cell_value(i, 0)).strip()
        if not cid or cid.startswith('合计'):
            continue
        try:
            yz = float(str(ws.cell_value(i, 4)).replace(',', ''))
            lj = float(str(ws.cell_value(i, 8)).replace(',', ''))
        except (ValueError, TypeError):
            continue
        nm = str(ws.cell_value(i, 1)).strip()
        cls = str(ws.cell_value(i, 11)).strip()
        residual = _num(ws.cell_value(i, 6))
        months = int(_num(ws.cell_value(i, 3)) or 0)
        dep_months = int(_num(ws.cell_value(i, 9)) or 0)
        start = str(ws.cell_value(i, 2))
        mdep_book = _num(ws.cell_value(i, 12))
        raw_cards.append({'cid': cid, 'name': nm, 'yz': yz, 'lj': lj, 'cls': cls,
                          'residual': residual, 'months': months, 'dep_months': dep_months,
                          'start': start, 'mdep_book': mdep_book})
        key = (nm, round(yz, 2))
        c = cards[key]
        c['count'] += 1
        c['yz'] = yz
        c['lj'] += lj
        c['cls'] = cls
        c['mdep_book'] = c['mdep_book'] or mdep_book   # ⚡ 本月计提折旧额（账面当期口径）
        if c['residual'] is None:
            c['residual'] = residual
        c['months'] = c['months'] or months
        c['dep_months'] = c['dep_months'] or dep_months
        c['start'] = c['start'] or start
        if len(c['cids']) < 3:
            c['cids'].append(cid)
            c['names'].append(nm)
    return cards, raw_cards


def period_movement(c_old, c_new, period_label):
    """跨期差集（多重集）：old→new 的新增/减少（按 (名称,原值) 计数差）。"""
    inc = []   # 新增卡片行
    dec = []   # 减少卡片行
    for key, cn in c_new.items():
        co = c_old.get(key)
        n_old = co['count'] if co else 0
        if cn['count'] > n_old:
            for _ in range(cn['count'] - n_old):
                inc.append({'name': cn['names'][0] if cn['names'] else key[0], 'yz': cn['yz'],
                            'cls': cn['cls'], 'lj': cn['lj'] / max(cn['count'], 1), 'cids': cn['cids']})
    for key, co in c_old.items():
        cn = c_new.get(key)
        n_new = cn['count'] if cn else 0
        if co['count'] > n_new:
            for _ in range(co['count'] - n_new):
                dec.append({'name': co['names'][0] if co['names'] else key[0], 'yz': co['yz'],
                            'cls': co['cls'], 'lj': co['lj'] / max(co['count'], 1), 'cids': co['cids']})
    return inc, dec


def find_gl_dispose_dates(dec_list, gl_fp):
    """减少卡片 → GL 清理/处置分录日期匹配。
    候选分录：①固定资产清理借方 ②固定资产(1601)贷方+处置词 ③1601 贷方（普通减少，低优先级）
    匹配：名称关键词（前 4 字）in 摘要 + 金额误差<2% → 取凭证日期。匹配不到 → '待人工补'。"""
    wb = openpyxl.load_workbook(gl_fp, read_only=True)
    ws = wb.worksheets[0]
    dispose = []   # (日期, 凭证号, 摘要, 金额, 优先级0高/1低)
    for r in ws.iter_rows(values_only=True):
        km = str(r[0] or '')
        sm = str(r[5] or '')
        try:
            jf = float(r[6] or 0)
            df = float(r[7] or 0)
        except (ValueError, TypeError):
            continue
        kw_dispose = any(w in sm for w in ('报废', '清理', '处置', '出售', '盘亏', '捐赠', '拆除', '注销'))
        is_fa = '固定资产' in km and '累计折旧' not in km and '清理' not in km   # 含子目（固定资产-XX）
        if '固定资产清理' in km and (jf or df):
            amt = jf if jf else df
            dispose.append((str(r[1] or '')[:10], str(r[3] or ''), sm, amt, 0))
        elif is_fa and df and kw_dispose:
            dispose.append((str(r[1] or '')[:10], str(r[3] or ''), sm, df, 0))
        elif is_fa and df:
            dispose.append((str(r[1] or '')[:10], str(r[3] or ''), sm, df, 1))
    matched = 0
    for d in dec_list:
        kw = d['name'][:4] if len(d['name']) >= 4 else d['name']
        hit = None
        best_pri = 9
        for dt, vno, sm, amt, pri in dispose:
            if kw and kw in sm and abs(amt - d['yz']) / max(abs(d['yz']), 1) < 0.02 and pri <= best_pri:
                hit = (dt, vno, sm)
                best_pri = pri
        if hit:
            d['reduce_date'] = hit[0]
            d['reduce_vno'] = hit[1]
            d['reduce_sm'] = hit[2][:30]
            matched += 1
        else:
            d['reduce_date'] = '待人工补'
            d['reduce_vno'] = ''
            d['reduce_sm'] = ''
    return matched


def read_tb_fixed(fp):
    """TB 1601/1602 全列 + 二级类别：{'t1':1601一级, 't2':1602一级, 'classes':{类别:{固定资产/累计折旧 4列}}}。
    列：0科目代码 1科目名称 2方向 3期初金额 4本期借方 5本期贷方 6方向 7期末金额"""
    TB_CLASS_1601 = {'160101': '房屋及建筑物', '160102': '通用设备', '160103': '专用设备',
                     '160104': '运输设备', '160105': '电子设备', '160106': '家具用具及其他'}
    TB_CLASS_1602 = {'160201': '房屋及建筑物', '160202': '通用设备', '160203': '专用设备',
                     '160204': '电子设备', '160205': '运输设备', '160206': '家具用具及其他'}
    wb = openpyxl.load_workbook(fp, read_only=True)
    ws = wb.worksheets[0]
    out = {'t1': {}, 't2': {}, 'classes': {}}
    for r in ws.iter_rows(values_only=True):
        code = str(r[0] or '')
        if code not in ('1601', '1602') and code not in TB_CLASS_1601 and code not in TB_CLASS_1602:
            continue

        def _f(i):
            try:
                return float(r[i] or 0)
            except (ValueError, TypeError):
                return 0.0

        if code == '1601':
            out['t1'] = {'begin': _f(3), 'jf': _f(4), 'df': _f(5), 'end': _f(7)}
        elif code == '1602':
            out['t2'] = {'begin': _f(3), 'jf': _f(4), 'df': _f(5), 'end': _f(7)}
        elif code in TB_CLASS_1601:
            cl = TB_CLASS_1601[code]
            v = out['classes'].setdefault(cl, {'begin': 0.0, 'inc': 0.0, 'dec': 0.0, 'end': 0.0,
                                               'dep_begin': 0.0, 'dep_inc': 0.0, 'dep_dec': 0.0, 'dep_end': 0.0})
            v['begin'] += _f(3)
            v['inc'] += _f(4)          # 借=增加
            v['dec'] += _f(5)          # 贷=减少
            v['end'] += _f(7)
        elif code in TB_CLASS_1602:
            cl = TB_CLASS_1602[code]
            v = out['classes'].setdefault(cl, {'begin': 0.0, 'inc': 0.0, 'dec': 0.0, 'end': 0.0,
                                               'dep_begin': 0.0, 'dep_inc': 0.0, 'dep_dec': 0.0, 'dep_end': 0.0})
            v['dep_begin'] += _f(3)
            v['dep_inc'] += _f(5)      # 贷=计提
            v['dep_dec'] += _f(4)      # 借=转出
            v['dep_end'] += _f(7)
    return out


def summarize(c_old, inc_list, dec_list, c_new):
    """分类汇总（固定资产4列 + 累计折旧4列，勾稽：期初+计提−转出=期末）：
    固定资产: begin期初原值 / inc本期新增 / dec本期减少 / end期末原值（Σc_new）
    累计折旧: begin_dep期初(Σc_old.lj) / dep_inc本期计提(=期末−期初+转出) / dec_dep本期转出(Σ减少卡片lj) / end_dep期末(Σc_new.lj)"""
    def _v(out, cl):
        return out.setdefault(cl, {'begin': 0.0, 'inc': 0.0, 'dec': 0.0, 'end': 0.0,
                                   'begin_dep': 0.0, 'dep_inc': 0.0, 'dec_dep': 0.0, 'end_dep': 0.0})

    out = {}
    for key, c in c_old.items():
        v = _v(out, CARD_GL_CLASS.get(c['cls'], '其他'))
        v['begin'] += c['yz'] * c['count']
        v['begin_dep'] += c['lj']
    for key, c in c_new.items():
        v = _v(out, CARD_GL_CLASS.get(c['cls'], '其他'))
        v['end'] += c['yz'] * c['count']
        v['end_dep'] += c['lj']
    for d in inc_list:
        _v(out, CARD_GL_CLASS.get(d['cls'], '其他'))['inc'] += d['yz']
    for d in dec_list:
        v = _v(out, CARD_GL_CLASS.get(d['cls'], '其他'))
        v['dec'] += d['yz']
        v['dec_dep'] += d['lj']
    for cl, v in out.items():
        v['dep_inc'] = v['end_dep'] - v['begin_dep'] + v['dec_dep']   # 本期计提 = 期末 − 期初 + 转出（勾稽）
    return out


def run_period(period, out_dir=None):
    """按期间生成复核底稿（period: 2025=25全年段 / 2026=26年1-3月段；out_dir 可覆盖输出目录）。"""
    if period == '2025':
        c_old, raw_old = load_cards(os.path.join(CARD_DIR, '截止2024年12月31日资产卡片.xls'))
        c_new, raw_new = load_cards(os.path.join(CARD_DIR, '截止2025年12月31日资产卡片.xls'))
        gl_fp = GL25
        tb_fp = TB25
        seg_label = '2025 年度（期初=2024年末，期末=2025年末）'
    else:
        c_old, raw_old = load_cards(os.path.join(CARD_DIR, '截止2025年12月31日资产卡片.xls'))
        c_new, raw_new = load_cards(os.path.join(CARD_DIR, '截止2026年3月31日资产卡片.xls'))
        gl_fp = GL26
        tb_fp = TB26
        seg_label = '2026 年 1-3 月（期初=2025年末，期末=2026年3月末）'

    inc, dec = period_movement(c_old, c_new, period)
    n_matched = find_gl_dispose_dates(dec, gl_fp)
    s = summarize(c_old, inc, dec, c_new)
    tot = {k: sum(v[k] for v in s.values()) for k in ('begin', 'inc', 'dec', 'end', 'begin_dep', 'dep_inc', 'dec_dep', 'end_dep')}
    tb = read_tb_fixed(tb_fp)
    t1, t2 = tb['t1'], tb['t2']
    tb_classes = tb['classes']

    # ⚡ 折旧复算（对齐 XBJ：系统次月计提、年内启用次月至期末、此前启用满期、累计已提足则 0）
    end_y, end_m = (2025, 12) if period == '2025' else (2026, 3)
    full_months = 12 if period == '2025' else end_m   # 此前启用卡片的满期月数

    def _total_months(start):
        """自启用【次月】至期末的累计应提月数（系统次月计提，对齐 XBJ）。"""
        m = re.search(r'(\d{4})[.年\-/](\d{1,2})', str(start))
        if not m:
            return 0
        sy, s_m = int(m.group(1)), int(m.group(2))
        if sy > end_y or (sy == end_y and s_m >= end_m):
            return 0
        return max(0, (end_y - sy) * 12 + (end_m - s_m))

    def _months_until(start, total, limit):
        """期末前应提月数：期内启用→启用次月至期末；此前启用→满期月数。
        ⚡ 提足判断改由【账面期初累计折旧】剩余可提封顶负责（2026-08-17 修正）：原"推算累计满期→直接停提"会造成
        提足最后一年账面仍在计提、匡算却为 0 的 90% 假差异（对齐 ADF _months_2026）。"""
        m = re.search(r'(\d{4})[.年\-/](\d{1,2})', str(start))
        if not m or not limit or limit <= 0:
            return 0
        sy, s_m = int(m.group(1)), int(m.group(2))
        if sy > end_y or (sy == end_y and s_m >= end_m):
            return 0
        if sy == end_y:
            return max(0, min(end_m - s_m, int(limit)))
        return min(full_months, int(limit))

    # ===== 构建逐卡片行（存量/新增/减少全覆盖，固定资产4列+累计折旧4列）=====
    # reduce_info: (名称,原值) → (日期,凭证,摘要)（从 GL 匹配）
    reduce_info = {(d['name'], round(d['yz'], 2)): (d.get('reduce_date', '待人工补'),
                                                    d.get('reduce_vno', ''), d.get('reduce_sm', '')) for d in dec}
    rows = []   # 每行一张卡片
    # 期末逐卡：按 (名称,原值) 多重集匹配期初存量，前 min(n_old,n_new) 张=存量、其余=新增（一行一张，各自计算）
    key_consumed = defaultdict(int)
    raw_old_by_key = defaultdict(list)
    for rc in raw_old:
        raw_old_by_key[(rc['name'], round(rc['yz'], 2))].append(rc)
    for card in raw_new:
        if card['cls'] == '土地':
            continue
        key = (card['name'], round(card['yz'], 2))
        c = c_new[key]
        co = c_old.get(key)
        n_old = co['count'] if co else 0
        residual = card['residual'] if card['residual'] is not None else 0.05
        months = card['months'] or 1
        mdep_calc = card['yz'] * (1 - residual) / months if months else 0.0   # 单卡口径
        mdep_book = card['mdep_book']
        is_stock = key_consumed[key] < n_old
        key_consumed[key] += 1
        if is_stock:
            lj_old = co['lj'] / n_old                       # 期初单卡累计折旧
            yz_old = card['yz']; yz_inc = 0.0
            dep_old = lj_old; dep_inc = card['lj'] - lj_old; dep_end = card['lj']
            kind = '存量'
        else:
            yz_old = 0.0; yz_inc = card['yz']
            dep_old = 0.0; dep_inc = card['lj']; dep_end = card['lj']
            kind = '新增'
        base = card['yz'] * (1 - residual)                 # ⚡ 可提总额（原值−残值）
        remain = max(base - dep_old, 0.0)                  # ⚡ 剩余可提（用账面期初累计折旧判断，对齐 ADF）
        if remain < 1.0:
            remain = 0.0                                   # ⚡ 浮点尾差容差：剩余可提 <1 元视为已提足（期初累折≈封顶差几毛钱的卡）
        acc_tot = _total_months(card['start'])
        acc_m = _months_until(card['start'], acc_tot, months)
        if remain <= 0:
            acc_m = 0                                      # ⚡ 期初已提足 → 本期停提
        dep_year_book = dep_inc
        dep_year_calc = min(mdep_calc * acc_m, remain)     # ⚡ 提足封顶（补足到封顶即止）；新增老设备补录卡按本期正常口径，账面一次性补提的差异如实反映
        diff_year = dep_year_calc - dep_year_book
        exp_dep = min(dep_old + dep_year_calc, base)       # ⚡ 匡算期末累计 = 期初+本年（封顶到可提总额）
        diff_cum = exp_dep - card['lj']
        cl = CARD_GL_CLASS.get(card['cls'], '其他')
        note = ''
        if kind == '新增':
            note = '本期新增'
        elif abs(diff_year) > 1.0 and abs(diff_year) / max(dep_year_book, 1) > 0.05:
            note = '本年折旧差异率>5%（年限/残值率与账面不一致）'
        rows.append({'cid': card['cid'], 'name': card['name'], 'cl': cl, 'start': card['start'], 'months': months,
                     'dep_months': card['dep_months'], 'residual': residual,
                     'yz_old': yz_old, 'yz_inc': yz_inc, 'yz_dec': 0.0, 'yz_end': card['yz'],
                     'dep_old': dep_old, 'dep_inc': dep_inc, 'dep_dec': 0.0, 'dep_end': dep_end,
                     'mdep_calc': mdep_calc, 'mdep_book': mdep_book,
                     'dep_year_calc': dep_year_calc, 'diff_year': diff_year,
                     'exp_dep': exp_dep, 'diff_cum': diff_cum,
                     'reduce_date': '', 'reduce_vno': '', 'reduce_sm': '', 'note': note, 'kind': kind})
    # 减少卡片（只在 c_old；逐张一行，转出累计折旧，减少日期从 GL 取）
    for key, co in c_old.items():
        cn = c_new.get(key)
        n_new = cn['count'] if cn else 0
        if co['count'] <= n_new:
            continue
        n_dec = co['count'] - n_new
        cl = CARD_GL_CLASS.get(co['cls'], '其他')
        rd, rv, rs = reduce_info.get((co['names'][0] if co['names'] else key[0], round(co['yz'], 2)),
                                     ('待人工补', '', ''))
        for rc in raw_old_by_key.get(key, [])[:n_dec]:
            residual = rc['residual'] if rc['residual'] is not None else 0.05
            rows.append({'cid': rc['cid'], 'name': key[0], 'cl': cl, 'start': rc['start'], 'months': rc['months'] or 1,
                         'dep_months': rc['dep_months'], 'residual': residual,
                         'yz_old': rc['yz'], 'yz_inc': 0.0, 'yz_dec': rc['yz'], 'yz_end': 0.0,
                         'dep_old': co['lj'] / co['count'], 'dep_inc': 0.0, 'dep_dec': co['lj'] / co['count'], 'dep_end': 0.0,
                         'mdep_calc': rc['yz'] * (1 - residual) / (rc['months'] or 1),
                         'mdep_book': rc['mdep_book'],
                         'dep_year_calc': 0.0, 'diff_year': 0.0, 'exp_dep': 0.0, 'diff_cum': 0.0,
                         'reduce_date': rd, 'reduce_vno': rv, 'reduce_sm': rs,
                         'note': '本期减少', 'kind': '减少'})
    rows.sort(key=lambda r: r['kind'] != '减少')   # 减少放后面

    # ===== 写底稿 =====
    wb = openpyxl.Workbook()
    # ⚡ 格式与正式底稿一致（对齐 XBJ 折旧复核：Times New Roman 10pt / 表头 DDEBF7 加粗居中 / 全 thin 边框 / 金额 #,##0.00）
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

    def style_cell(c, numeric=False, bold=False, center=False):
        """数据格统一格式：Times New Roman 10 / thin 边框 / 金额千分位。"""
        c.font = total_font if bold else base_font
        c.border = thin
        if numeric:
            c.number_format = num_fmt
        if center:
            c.alignment = CTR

    # ① 增减变动汇总（同一 sheet 上中下三表：卡片账 / 账套数 / 差异数，同排列 固定资产4列+累计折旧4列）
    ws = wb.active
    ws.title = '增减变动汇总'
    ws.append([f'GJX 固定资产增减变动汇总（{seg_label}）'])
    ws.append([])
    HDR8 = ['类别', '固定资产-期初原值', '固定资产-本期增加', '固定资产-本期减少', '固定资产-期末原值',
            '累计折旧-期初', '累计折旧-本期计提', '累计折旧-本期转出', '累计折旧-期末']

    def _dump_block(title, get8):
        """title=表名；get8(cls或None) → 8列值（None=合计行）。"""
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

    # 【上】卡片账
    def card8(cl):
        if cl is None:
            return (tot['begin'], tot['inc'], tot['dec'], tot['end'],
                    tot['begin_dep'], tot['dep_inc'], tot['dec_dep'], tot['end_dep'])
        v = s[cl]
        return (v['begin'], v['inc'], v['dec'], v['end'],
                v['begin_dep'], v['dep_inc'], v['dec_dep'], v['end_dep'])

    _dump_block('一、卡片台账（跨期差集，Σ类别）', card8)

    # 【中】账套数（TB 1601/1602 二级科目）
    def tb8(cl):
        if cl is None:
            return (t1['begin'], t1['jf'], t1['df'], t1['end'],
                    t2['begin'], t2['df'], t2['jf'], t2['end'])
        v = tb_classes.get(cl)
        if not v:
            return None
        return (v['begin'], v['inc'], v['dec'], v['end'],
                v['dep_begin'], v['dep_inc'], v['dep_dec'], v['dep_end'])

    _dump_block('二、账套数（科目余额表 1601/1602 二级科目）', tb8)

    # 【下】差异数（卡片 − 账套）
    def diff8(cl):
        c8 = card8(cl)
        t8 = tb8(cl)
        if t8 is None:
            return (c8[0], c8[1], c8[2], c8[3], c8[4], c8[5], c8[6], c8[7]) if cl is None else None
        return tuple(c8[i] - t8[i] for i in range(8))

    _dump_block('三、差异数（卡片台账 − 账套，差异=待查因）', diff8)

    ws.append(['勾稽：期末原值 = 期初+增加−减少；累计折旧期末 = 期初+计提−转出（两表各自成立）'])

    # ③ 折旧复算明细（主表：逐卡片固定资产4列+累计折旧4列+复算双轨+减少日期）
    ws2 = wb.create_sheet('折旧复算明细')
    ws2.append(['序号', '卡片编号', '资产名称', '底稿类别', '开始使用日期', '使用月限(月)', '已计提期数', '残值率',
                '固定资产-期初原值', '固定资产-本期增加', '固定资产-本期减少', '固定资产-期末原值',
                '累计折旧-期初', '累计折旧-本期计提', '累计折旧-本期转出', '累计折旧-期末',
                '月折旧额(计算)', '本月计提(账面)', '匡算本年折旧', '本年折旧差异',
                '匡算期末累计折旧', '累计折旧差异', '减少日期', 'GL凭证号', 'GL摘要', '备注'])
    style_header(ws2)
    NUM_COLS2 = {8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21}   # 数值列（0 基）
    for i, r in enumerate(rows, 1):
        r_ = ws2.max_row + 1
        ws2.append([i, r['cid'], r['name'], r['cl'], r['start'], r['months'],
                    r['dep_months'], r['residual'],
                    r['yz_old'], r['yz_inc'], r['yz_dec'], r['yz_end'],
                    r['dep_old'], r['dep_inc'], r['dep_dec'], r['dep_end'],
                    r['mdep_calc'], r['mdep_book'], r['dep_year_calc'], r['diff_year'],
                    r['exp_dep'], r['diff_cum'], r['reduce_date'], r['reduce_vno'], r['reduce_sm'], r['note']])
        for ci in range(26):
            style_cell(ws2.cell(r_, ci + 1), numeric=(ci in NUM_COLS2), center=(ci == 0))
    # 底部合计（⚡ 仅数值列加计：固定资产4列+累计折旧4列=8-15；复算/日期/文本列不加计）
    r_ = ws2.max_row + 1
    ws2.append(['合计', '', '', '', '', '', '', '',
                sum(x['yz_old'] for x in rows), sum(x['yz_inc'] for x in rows),
                sum(x['yz_dec'] for x in rows), sum(x['yz_end'] for x in rows),
                sum(x['dep_old'] for x in rows), sum(x['dep_inc'] for x in rows),
                sum(x['dep_dec'] for x in rows), sum(x['dep_end'] for x in rows)])
    for ci in range(26):
        style_cell(ws2.cell(r_, ci + 1), numeric=(ci in (8, 9, 10, 11, 12, 13, 14, 15)), bold=True)
    # 与账套核对（Σ卡片 4列×2 vs TB，同列对齐；仅数值列）
    r_ = ws2.max_row + 1
    ws2.append(['与账套核对(Σ卡片 vs TB1601/1602, 详细见「增减变动汇总」)', '', '', '', '', '', '', '',
                tot['begin'], tot['inc'], tot['dec'], tot['end'],
                tot['begin_dep'], tot['dep_inc'], tot['dec_dep'], tot['end_dep']])
    for ci in range(26):
        style_cell(ws2.cell(r_, ci + 1), numeric=(ci in (8, 9, 10, 11, 12, 13, 14, 15)), bold=True)
    ws2.append([])
    n_year_big = sum(1 for x in rows if x['kind'] != '减少' and x['diff_year'] and abs(x['diff_year']) > 1.0 and abs(x['diff_year']) / max(x['dep_inc'], 1) > 0.05)
    n_cum_big = sum(1 for x in rows if x['kind'] != '减少' and x['diff_cum'] and abs(x['diff_cum']) > 1.0 and abs(x['diff_cum']) / max(x['dep_end'], 1) > 0.05)
    n_red_wait = sum(1 for x in rows if x['reduce_date'] == '待人工补')
    r_ = ws2.max_row + 1
    ws2.append([f'本年折旧差异率>5%: {n_year_big} 行 | 累计差异率>5%: {n_cum_big} 行 | 减少日期待人工补: {n_red_wait} 张'])
    style_cell(ws2.cell(r_, 1))

    # ④ 类别参数区间（折旧政策核对）
    ws3 = wb.create_sheet('类别参数区间')
    ws3.append(['类别', '张数', '使用月限min(月)', '使用月限max(月)', '使用月限常见(月)',
                '残值率min', '残值率max', '残值率常见'])
    style_header(ws3)
    for cl in GL_CLASSES:
        items = [x for x in rows if x['cl'] == cl]
        if not items:
            continue
        months_l = [x['months'] for x in items if x['months'] > 0]
        res_l = [x['residual'] for x in items if x['residual'] > 0]
        mcom = Counter(months_l).most_common(1)[0][0] if months_l else 0
        rcom = Counter(round(r, 4) for r in res_l).most_common(1)[0][0] if res_l else 0
        r_ = ws3.max_row + 1
        ws3.append([cl, len(items), min(months_l) if months_l else 0, max(months_l) if months_l else 0, mcom,
                    min(res_l) if res_l else 0, max(res_l) if res_l else 0, rcom])
        for ci in range(1, 8):
            style_cell(ws3.cell(r_, ci + 1), numeric=True)
        style_cell(ws3.cell(r_, 1))
    r_ = ws3.max_row + 1
    ws3.append([])
    r_ = ws3.max_row + 1
    ws3.append(['注：与公司折旧政策核对（年限/残值率区间）；异常（年限超政策、残值率口径不一）需人工复核'])
    style_cell(ws3.cell(r_, 1))

    # ⑤ 口径说明
    ws4 = wb.create_sheet('口径说明')
    notes = [
        f'GJX 固定资产增减变动复核（{seg_label}）— 2026-08-16 铁律131k，v3 结构',
        '数据源：',
        f'  ① 卡片台账（3 期间）：{CARD_DIR}',
        '     截止2024年12月31日（1476 张，=25年期初存量）/ 截止2025年12月31日（1901 张）/ 截止2026年3月31日（1926 张）',
        f'  ② 账套：{"DQ2025" if period == "2025" else "DQ2026"} 综合查询明细表 + 科目余额表（GJX 自己的数据，不参考其他账套）',
        '卡片唯一键 = 资产名称+原值（多重集，同名同值多张按计数）——仅用于【跨期差集】判断新增/减少，规避卡片编号跨年重排误判减少。',
        '⚡ 类别映射（卡片 8 类 → 账套 GL 二级 5 类+运输单列）：',
        '  房屋建筑物→房屋及建筑物；专用设备→专用设备；通用设备→通用设备；',
        '  电子仪器设备+电子办公设备→电子设备；家具用具+工具用具→家具用具及其他；运输设备→运输设备(GL 无此类,单列)。',
        '  核对验证：卡片 25 年净增 5,712 万 = GL 各二级净发生合计 5,712 万（0 差）。',
        '⚡ 25 年"编号差集减少 52 张"中 47 张名称仍在（编号重排）——减少以【名称+原值】级消失为准。',
        '⚡ 结构说明（v3，用户 19:09-19:10 定调）：',
        '  ① 增减变动汇总：固定资产 4 列（期初/本期增加/本期减少/期末）与累计折旧 4 列（期初/本期计提/本期转出/期末）分列不交错；',
        '  ② 折旧复算明细=主表：逐卡片（一行一张，同名同值多张不合并）存量/新增/减少全覆盖，固定资产增减+累计折旧增减各 4 列，与账套核对；',
        '     减少日期从 GL（固定资产清理/1601贷方分录）匹配，匹配不到标『待人工补』；',
        '  ③ 已删除独立的新增明细/减少明细/折旧核对/与TB核对/本期计提核对 sheet（信息并入①②）。',
        '折旧复算（对齐 XBJ）：月折旧=原值×(1-残值率)÷年限；系统次月计提；年内启用=次月至期末；累计已提足则本期停提。',
        '与账套核对：卡片期初/增/减/末 vs TB 1601/1602；累计折旧本期计提=期末−期初+转出 vs TB1602 贷方——差异需查因直至对上。',
        f'⚡ 审计发现点：①卡片年限/开始使用日期字段存在系统性失真（2024-2025 重建卡片参数错位，{seg_label} 本年差异 {n_year_big} 行 / 累计差异 {n_cum_big} 行，按卡片行计；明细已逐卡展开，一行一张）',
        '  ②减少日期 GL 匹配覆盖率低（GL 清理分录仅 9 笔，多数小额处置未入清理科目）→ 标『待人工补』。',
    ]
    if period == '2025':
        notes += [
            '',
            '⚡⚡ 2025 电子设备差异查因结论（2026-08-16，卡片台账 vs 账套发生额）：',
            '  差异 100% 集中在【电子办公设备】类，两端同额：卡片新增 272 张(449.2万) vs TB借发 418.8万 差 +303,554.52；',
            '  卡片减少 45 张(49.8万) vs TB贷发 19.4万 差 +303,554.52（净额 0、期末余额全 0 差）。',
            '  账套 2025 年仅有 2 笔处置：变压器报废(贷15,044.25) + "9.30报废电子办公设备36件"(贷194,300.06)，合计 209,344.31=TB贷发。',
            '  折旧差异 288,514.99 = 电子办公 293,278.99（卡片转出47.8万−账套18.5万）− 变压器 4,764（账套转出5,558>卡片794）。',
            '  结论：卡片台账内部存在一批合计 303,554.52 的【换键/重录调整】（电子办公设备，约 9 张量级）——',
            '  卡片侧记了假增假减、账套无对应分录（疑似 DLEE/DELL 并存、同设备重录编号/名称），不影响账套余额。',
            '  审计处理：向企业核实该批调整是否固定资产系统重新录入所致；若其中含真实未入账处置需补记分录。',
            '',
            '⚡⚡ 老设备补录卡复核结论（2026-08-17）：',
            '  卡片账期初累计折旧 235,902,982.20 = TB1602 期初（0 差）；原值期初 762,240,864.97 = TB1601 期初（0 差）；期末亦 0 差。',
            '  → 该批老设备补录卡期初在卡片账与账套【均不存在】，系 2025 年才入账（卡片+原值+累计折旧全部 2025 新增），不存在"卡片遗失、账面期初有折旧"的错位。',
            '  其中 60 张"期初不存在、开始使用日期 2007-2019"的老设备（全为电子办公设备，原值合计 212,394.90），账面 2025 年一次性计提折旧 208,308.24：',
            '  复算按本期正常口径（2025 当年应提 12 个月 = 69,436.02），差异 138,872.22 即账面 2025 年一次性补提的以前年度折旧（入账方式所致）——',
            '  需向企业核实入账依据；若属固定资产盘盈，应按准则通过"以前年度损益调整"处理，而非全部计入 2025 年费用。',
        ]
    elif period == '2026':
        notes += [
            '',
            '⚡ 2026 年 1-3 月与账套核对：固定资产/累计折旧 期初·增加·减少·期末 四项全部 0 差（卡片台账与账套完全一致）。',
        ]
    for i, n in enumerate(notes, 1):
        ws4.cell(row=i, column=1, value=n).font = base_font
    ws4.column_dimensions['A'].width = 130

    # 列宽
    widths1 = [16, 18, 16, 16, 16, 18, 18, 16, 16]
    for i, w in enumerate(widths1, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    widths2 = [6, 12, 22, 12, 12, 12, 10, 9,
               16, 16, 16, 16, 16, 16, 16, 16,
               14, 14, 14, 14, 16, 14, 12, 10, 26, 22]
    for i, w in enumerate(widths2, 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for wsn in (ws, ws2, ws3):
        wsn.freeze_panes = 'A2'

    out_dir = out_dir or os.path.join(DATA_ROOT, ACCT, '底稿', period)
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f'固定资产增减变动复核_{period}.xlsx')
    wb.save(out)
    print(f'✅ 已生成: {out}')
    print(f'  {seg_label}')
    print(f'  期初{sum(c["count"] for c in c_old.values())}张 | 新增{len(inc)}张({sum(d["yz"] for d in inc):,.0f}) | '
          f'减少{len(dec)}张({sum(d["yz"] for d in dec):,.0f}) | 期末{sum(c["count"] for c in c_new.values())}张')
    print(f'  减少日期 GL 匹配: {n_matched}/{len(dec)}')
    print(f'  与账套核对: 固资期末 卡{tot["end"]:,.0f} vs TB1601 {t1.get("end",0):,.0f} 差{tot["end"]-t1.get("end",0):,.0f}'
          f' | 折旧期末 卡{tot["end_dep"]:,.0f} vs TB1602 {t2.get("end",0):,.0f} 差{tot["end_dep"]-t2.get("end",0):,.0f}')
    print(f'  折旧计提: 卡{tot["dep_inc"]:,.0f} vs TB1602贷方 {t2.get("df",0):,.0f} 差{tot["dep_inc"]-t2.get("df",0):,.0f}'
          f' | 转出 卡{tot["dec_dep"]:,.0f} vs TB1602借方 {t2.get("jf",0):,.0f} 差{tot["dec_dep"]-t2.get("jf",0):,.0f}')
    return out


def main():
    periods = sys.argv[1:] if len(sys.argv) > 1 and not sys.argv[1].startswith('-') else ['2025', '2026']
    for p in periods:
        run_period(p)
    return 0


if __name__ == '__main__':
    sys.exit(main())
