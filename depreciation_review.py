# -*- coding: utf-8 -*-
"""固定资产折旧计提复核（铁律131c 升级版，2026-08-15）。

功能：
  Sheet1 折旧计提复核：明细 3011 卡，原值/累计折旧按【期初/本期增加/本期减少/期末】四段
         （卡片账仅期末时点，期初=期末-本年启用原值；增加=2025年内启用原值；减少=0无数据；
           累计期初=期末-本年折旧；累计增加=本年折旧；累计减少=0）
  Sheet2 类别汇总：按披露类别（映射表）原值/累计折旧四段归类
  Sheet3 类别vs账面差异：卡片账（映射到底稿科目）vs 固定资产审计底稿《固定资产分类汇总表》
         （全部主体按科目代码聚合）期末原值/累计折旧差异
  Sheet4 类别参数区间：按披露类别统计 使用月限 与 残值率 区间（min/max/常见值），
         供与公司折旧政策比较
  Sheet5 口径说明

用法：python depreciation_review.py XBJ [--out <路径>] [--residual 0.03] [--year 2025]
"""
import os
import re
import sys
from collections import Counter
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

import depreciation_manifest as DM

F = Font(name='Times New Roman', size=10)
FB = Font(name='Times New Roman', size=10, bold=True)
HFILL = PatternFill('solid', fgColor='DDEBF7')
TFILL = PatternFill('solid', fgColor='FCE4D6')
DFILL = PatternFill('solid', fgColor='FFF2CC')
THIN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
NUM = '#,##0.00'

HEADERS = ['序号', '资产编码', '核算主体', '资产名称', '资产类别', '开始使用日期',
           '使用月限(月)', '已计提期数', '残值率',
           '固定资产期末(原值)', '固定资产期初', '本期增加', '本期减少',
           '账面累计折旧(期末)', '累计折旧期初', '本期计提(账面)',
           '月折旧额(匡算)', '匡算期末累计折旧', '累计折旧差异',
           '匡算本年折旧', '本年折旧差异', '备注']


def _parse_date(v):
    if isinstance(v, (int, float)):
        try:
            return date(1899, 12, 30) + __import__('datetime').timedelta(days=int(v))
        except Exception:
            return None
    s = str(v or '').strip()
    m = re.match(r'(\d{4})[-/.](\d{1,2})', s)
    if m:
        return date(int(m.group(1)), int(m.group(2)), 1)
    return None


def _total_months(start, year):
    """自启用【次月】至 year 年末的累计应计提月数（系统次月计提口径，实测 2022-10 启用→38 期）。"""
    if start is None or start.year > year:
        return 0
    return max(0, (year - start.year) * 12 + (12 - start.month))


def _months_2025(start, total, limit, year):
    """year 年应计提月数：年内启用 → 启用次月至年末；此前启用 → 12 个月（累计已提足则 0）。"""
    if start is None or not limit or limit <= 0:
        return 0
    if start.year >= year:
        m = 12 - start.month
        return max(0, min(m, int(limit)))
    if total >= limit:
        return 0
    return 12


def _residual_of(yz, lj, limit, periods, mdep):
    """反推残值率：月折旧额>0 用月折旧额；否则用 累计折旧/已计提期数；无法反推返回 None。"""
    if yz > 0 and limit > 0:
        if mdep and mdep > 0.005:
            return (1 - mdep * limit / yz) * 100
        if periods > 0.5 and lj > 0.005:
            md = lj / periods
            return (1 - md * limit / yz) * 100
    return None


def _load_map(map_path):
    """资产台账映射表 → {卡片类别: (底稿科目代码, 披露类别)}。"""
    cat_map = {}
    if not map_path or not os.path.exists(map_path):
        return cat_map
    try:
        wb = openpyxl.load_workbook(map_path, read_only=True, data_only=True)
        ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 4:
                continue
            orig = str(r[1] or '').strip()
            code = str(r[2] or '').strip()
            disp = str(r[3] or '').strip()
            if orig:
                cat_map[orig] = (code, disp)
        wb.close()
    except Exception as e:
        print(f'  ⚠️ 映射表读取失败：{e}')
    return cat_map


def _load_book(audit_path):
    """固定资产审计底稿《固定资产分类汇总表》→ {科目代码: {期初,增加,减少,期末}}。

    只取 1601 原值 / 1602 累计折旧 6 位明细科目，全部主体聚合。
    """
    book = {}
    if not audit_path or not os.path.exists(audit_path):
        return book
    try:
        wb = openpyxl.load_workbook(audit_path, read_only=True, data_only=True)
        if '固定资产分类汇总表' not in wb.sheetnames:
            wb.close()
            return book
        for r in wb['固定资产分类汇总表'].iter_rows(min_row=4, values_only=True):
            if not r or len(r) < 8:
                continue
            code = str(r[2] or '').strip()
            if not re.match(r'^160[12]\d{2}$', code):
                continue
            b = book.setdefault(code, {'qc': 0.0, 'inc': 0.0, 'dec': 0.0, 'qm': 0.0})
            b['qc'] += float(r[4] or 0)
            b['inc'] += float(r[5] or 0)
            b['dec'] += float(r[6] or 0)
            b['qm'] += float(r[7] or 0)
        wb.close()
    except Exception as e:
        print(f'  ⚠️ 账面分类汇总读取失败：{e}')
    return book


def build(cfg, out_path):
    src = DM.resolve(cfg['card'])
    year = int(cfg.get('year', 2025))
    residual = float(cfg.get('residual', 0.03))
    org = cfg.get('org', '')
    land_kw = cfg.get('land_kw', ('土地',))
    has_sum = cfg.get('has_summary_row', True)
    cat_map = _load_map(DM.resolve(cfg['map']) if cfg.get('map') else None)
    book = _load_book(DM.resolve(cfg['audit']) if cfg.get('audit') else None)

    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb['卡片台账']
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[3]
    ci = {str(h): j for j, h in enumerate(hdr) if h}
    col = {k: ci[k] for k in ('资产编码', '资产名称', '资产类别', '开始使用日期',
                              '使用月限', '已计提期数', '本币原值', '累计折旧',
                              '本年折旧', '净值', '月折旧额')}
    wb.close()

    out = openpyxl.Workbook()
    w1 = out.active
    w1.title = '折旧计提复核'
    w1.append(HEADERS)
    for j in range(1, len(HEADERS) + 1):
        c = w1.cell(1, j)
        c.font = FB; c.fill = HFILL
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = THIN

    s_tot = {'原值': 0.0, '账面累计': 0.0, '账面本年': 0.0, '匡算累计': 0.0, '匡算本年': 0.0,
             '期初原值': 0.0, '增加原值': 0.0, '累计期初': 0.0}
    n_ok = n_land = n_odd = 0
    n = 0
    n_age = 0
    # ⚡ 预扫描：按披露类别统计使用月限 → 离群阈值（Q1-1.5×IQR ~ Q3+1.5×IQR，类别样本≥5）
    #   超出区间=折旧年限异常（如类别大多 60 月，个别 12/240 月）
    cat_limits = {}
    for r in rows[4:]:
        if not r or not r[0]:
            continue
        c0 = str(r[col['资产编码']] or '').strip()
        if has_sum and c0 == '合计':
            continue
        cat0 = str(r[col['资产类别']] or '').strip()
        lim0 = float(r[col['使用月限']] or 0)
        if lim0 <= 0:
            continue
        d0 = cat_map[cat0][1] if cat0 in cat_map else cat0
        cat_limits.setdefault(d0, []).append(lim0)
    cat_lim = {}
    for d0, lims in cat_limits.items():
        if len(lims) < 5:
            continue
        ls = sorted(lims)
        q1 = ls[len(ls) // 4]
        q3 = ls[3 * len(ls) // 4]
        iqr = q3 - q1
        common = Counter(ls).most_common(1)[0][0]
        cat_lim[d0] = (q1 - 1.5 * iqr, q3 + 1.5 * iqr, common)
    # 类别统计：disp -> [卡数, 原值期初,原值增加,原值减少,原值期末, 累计期初,累计增加,累计减少,累计期末,
    #                   月限list, 残值率list]
    by_cat = {}
    for r in rows[4:]:
        if not r or not r[0]:
            continue
        code = str(r[col['资产编码']] or '').strip()
        if has_sum and code == '合计':
            continue
        name = str(r[col['资产名称']] or '').strip()
        cat = str(r[col['资产类别']] or '').strip()
        start = _parse_date(r[col['开始使用日期']])
        limit = float(r[col['使用月限']] or 0)
        periods = float(r[col['已计提期数']] or 0)
        yz = float(r[col['本币原值']] or 0)
        lj = float(r[col['累计折旧']] or 0)
        bn = float(r[col['本年折旧']] or 0)
        mdep_raw = float(r[col['月折旧额']] or 0) if col.get('月折旧额') is not None else 0
        n += 1
        # 四段推算（卡片账仅期末时点）
        inc_yz = yz if (start and start.year == year) else 0.0      # 本期增加=年内启用
        qc_yz = yz - inc_yz                                          # 期初原值
        lj_qc = lj - bn                                              # 期初累计=期末-本年（无转出）
        s_tot['原值'] += yz
        s_tot['期初原值'] += qc_yz
        s_tot['增加原值'] += inc_yz
        s_tot['账面累计'] += lj
        s_tot['累计期初'] += lj_qc
        s_tot['账面本年'] += bn
        # 披露类别/底稿科目（映射表）
        disp, book_code = (cat_map[cat][1], cat_map[cat][0]) if cat in cat_map else (cat, '')
        remark = []
        is_land = any(k in cat for k in land_kw)
        if limit > 0 and yz > 0 and not is_land:
            mdep = yz * (1 - residual) / limit
            tot_m = _total_months(start, year)
            calc_lj = min(tot_m, int(limit)) * mdep
            m_cur = _months_2025(start, tot_m, limit, year)
            calc_bn = mdep * m_cur
        else:
            mdep = calc_lj = calc_bn = 0.0
            if is_land:
                remark.append('土地不计提折旧')
                n_land += 1
            elif limit <= 0:
                remark.append('使用月限缺失')
                n_odd += 1
            elif yz <= 0:
                remark.append('原值为0')
        diff_lj = lj - calc_lj
        diff_bn = bn - calc_bn
        s_tot['匡算累计'] += calc_lj
        s_tot['匡算本年'] += calc_bn
        if abs(diff_lj) > 0.005 or abs(diff_bn) > 0.005:
            remark.append('差异见列')
        if periods > limit:
            remark.append(f'已计提{int(periods)}期超月限{int(limit)}')
        # ⚡ 折旧年限异常（超出类别大多数区间）
        if disp in cat_lim:
            _lo, _hi, _common = cat_lim[disp]
            if limit > 0 and (limit < _lo or limit > _hi):
                remark.append(f'折旧年限异常({int(limit)}月，类别常见{int(_common)}月)')
                n_age += 1
        if not remark:
            n_ok += 1
        row = [n, code, org, name, cat,
               start.strftime('%Y-%m') if start else '', int(limit) if limit else '',
               int(periods) if periods else '', f'{residual:.0%}',
               round(yz, 2), round(qc_yz, 2) if qc_yz else '', round(inc_yz, 2) if inc_yz else '', '',
               round(lj, 2), round(lj_qc, 2) if lj_qc else '', round(bn, 2),
               round(mdep, 2) if mdep else '', round(calc_lj, 2) if calc_lj else '',
               round(diff_lj, 2) if abs(diff_lj) > 0.005 else '',
               round(calc_bn, 2) if calc_bn else '',
               round(diff_bn, 2) if abs(diff_bn) > 0.005 else '',
               '；'.join(remark)]
        w1.append(row)
        rr = w1.max_row
        for j in range(1, len(HEADERS) + 1):
            c = w1.cell(rr, j)
            c.font = F; c.border = THIN
            if isinstance(c.value, (int, float)) and j not in (1, 7, 8):
                c.number_format = NUM
            if j in (19, 21) and isinstance(c.value, (int, float)) and abs(c.value) > 0.005:
                c.fill = DFILL
        # 类别统计（披露类别）
        a = by_cat.setdefault(disp, [0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, [], []])
        a[0] += 1
        a[1] += qc_yz; a[2] += inc_yz; a[3] += 0; a[4] += yz
        a[5] += lj_qc; a[6] += bn; a[7] += 0; a[8] += lj
        if limit > 0:
            a[9].append(limit)
        rv = _residual_of(yz, lj, limit, periods, mdep_raw)
        if rv is not None and 0 <= rv <= 100:
            a[10].append(round(rv, 2))
    # 合计行
    w1.append(['', '', '合计', '', '', '', '', '', '',
               round(s_tot['原值'], 2), round(s_tot['期初原值'], 2), round(s_tot['增加原值'], 2), '',
               round(s_tot['账面累计'], 2), round(s_tot['累计期初'], 2), round(s_tot['账面本年'], 2),
               '', round(s_tot['匡算累计'], 2), round(s_tot['账面累计'] - s_tot['匡算累计'], 2),
               round(s_tot['匡算本年'], 2), round(s_tot['账面本年'] - s_tot['匡算本年'], 2),
               f'卡片{n}张(去合计行)；残值率{residual:.0%}直线法；土地{n_land}张不计提；异常{n_odd}张；折旧年限异常{n_age}张'])
    rr = w1.max_row
    for j in range(1, len(HEADERS) + 1):
        c = w1.cell(rr, j)
        c.font = FB; c.fill = TFILL; c.border = THIN
        if isinstance(c.value, (int, float)):
            c.number_format = NUM
    w1.freeze_panes = 'A2'
    w1.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(HEADERS))}{w1.max_row}'
    for i, wd in enumerate([6, 20, 22, 20, 14, 10, 10, 10, 8, 15, 12, 12, 12,
                            15, 12, 14, 14, 16, 13, 14, 13, 22], 1):
        w1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd

    # Sheet2 类别汇总（四段）
    w2 = out.create_sheet('类别汇总(四段)')
    w2.append(['披露类别', '卡片数', '原值期初', '原值本期增加', '原值本期减少', '原值期末',
               '累计折旧期初', '累计折旧本期增加', '累计折旧本期减少', '累计折旧期末'])
    for cat, a in sorted(by_cat.items()):
        w2.append([cat, a[0]] + [round(x, 2) if isinstance(x, float) else x for x in a[1:9]])
    for rr in w2.iter_rows(min_row=2):
        for c in rr:
            c.font = F; c.border = THIN
            if isinstance(c.value, (int, float)) and c.column > 2:
                c.number_format = NUM
    for i, wd in enumerate([18, 8, 13, 13, 13, 13, 14, 14, 14, 14], 1):
        w2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd

    # Sheet3 类别 vs 账面差异（按披露类别对齐）
    # 账面 1601xx/1602xx 科目代码第 5-6 位（明细）→ 披露类别（160101/160201→土地资产 等）
    SUB_DISP = {'01': '土地资产', '02': '房屋及建筑物', '03': '机械设备', '04': '运输设备',
                '05': '其他设备', '06': '其他设备', '07': '其他设备', '08': '其他设备',
                '99': '其他设备'}
    w3 = out.create_sheet('类别vs账面差异')
    w3.append(['披露类别', '卡片账原值期末', '账面原值期末', '原值差异',
               '卡片账累计期末', '账面累计期末', '累计差异'])
    # 卡片账按披露类别聚合（原值/累计）
    card_by_disp = {}
    for cat, a in by_cat.items():
        b = card_by_disp.setdefault(cat, [0.0, 0.0])
        b[0] += a[4]; b[1] += a[8]
    # 账面按披露类别聚合（原值 1601xx / 累计 1602xx）
    book_by_disp = {}
    for code, b in book.items():
        disp = SUB_DISP.get(code[4:6], '其他设备')
        d = book_by_disp.setdefault(disp, [0.0, 0.0])
        if code.startswith('1601'):
            d[0] += b['qm']
        else:
            d[1] += b['qm']
    disps = sorted(set(card_by_disp) | set(book_by_disp))
    for disp in disps:
        cy = card_by_disp.get(disp, [0.0, 0.0])
        by = book_by_disp.get(disp, [0.0, 0.0])
        w3.append([disp, round(cy[0], 2), round(by[0], 2), round(cy[0] - by[0], 2),
                   round(cy[1], 2), round(by[1], 2), round(cy[1] - by[1], 2)])
    for rr in w3.iter_rows(min_row=2):
        for c in rr:
            c.font = F; c.border = THIN
            if isinstance(c.value, (int, float)) and c.column > 1:
                c.number_format = NUM
    for i, wd in enumerate([18, 15, 15, 13, 15, 15, 13], 1):
        w3.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd

    # Sheet4 类别参数区间（残值率/折旧年限，供与公司政策比较）
    w4 = out.create_sheet('类别参数区间')
    w4.append(['披露类别', '卡片数', '折旧年限min(年)', '折旧年限max(年)', '折旧年限常见(年)',
               '使用月限min(月)', '使用月限max(月)', '使用月限常见(月)', '残值率min', '残值率max', '残值率常见'])
    for cat, a in sorted(by_cat.items()):
        limits = sorted(a[9])
        resvs = sorted(a[10])
        w4.append([cat, a[0],
                   round(limits[0] / 12, 1) if limits else '', round(limits[-1] / 12, 1) if limits else '',
                   round(Counter(limits).most_common(1)[0][0] / 12, 1) if limits else '',
                   int(limits[0]) if limits else '', int(limits[-1]) if limits else '',
                   int(Counter(limits).most_common(1)[0][0]) if limits else '',
                   min(resvs) if resvs else '', max(resvs) if resvs else '',
                   Counter(resvs).most_common(1)[0][0] if resvs else ''])
    for rr in w4.iter_rows(min_row=2):
        for c in rr:
            c.font = F; c.border = THIN
            if isinstance(c.value, (int, float)) and c.column > 2:
                c.number_format = NUM
    for i, wd in enumerate([18, 8, 14, 14, 15, 13, 13, 14, 11, 11, 12], 1):
        w4.column_dimensions[openpyxl.utils.get_column_letter(i)].width = wd

    # Sheet5 口径说明
    w5 = out.create_sheet('口径说明')
    notes = [
        f'固定资产折旧计提复核（{org}，卡片账查询期间 {year}-12）',
        f'1. 数据源：{DM.resolve(cfg["card"])}',
        '2. 核算主体：' + org + '（卡片账=全集团固定资产账，与试算表 1601 核对）',
        f'3. 残值率 {residual:.0%}：默认参数（卡片账可反推校准：月折旧额或累计÷期数双通道）',
        '4. 折旧方法：直线法按使用月限均摊，月折旧=原值×(1-残值率)÷使用月限',
        '5. 匡算期末累计折旧=min(自启用次月至年末累计应提月数,使用月限)×月折旧（提足封顶）',
        '6. 匡算本年折旧=月折旧×本年计提月数（系统次月计提：年内启用=启用次月至年末；此前启用=12个月，累计已提足则0）',
        '7. 土地类（' + '、'.join(land_kw) + '）不计提折旧，仅列示原值',
        '8. 【四段列示口径】卡片账仅期末时点：本期增加=2025年内启用资产原值；期初=期末-本期增加；',
        '   本期减少=0（卡片账无减少数据，待期初台账补充）；累计折旧期初=期末-本年折旧（无转出时成立）；',
        '   累计本期增加=本年折旧；累计本期减少=0',
        '9. 累计折旧差异=账面累计折旧-匡算期末累计折旧（正=账面多提）；本年折旧差异=账面本年-匡算本年',
        '10. 类别vs账面差异：卡片账按映射表归到披露类别（其他设备/机械设备/运输设备/房屋及建筑物/',
        '    土地资产等），账面侧取固定资产审计底稿《固定资产分类汇总表》全部主体按科目代码',
        '    （1601xx 原值/1602xx 累计折旧，第5-6位明细：01土地/02房屋/03机器/04运输/其余其他）',
        '    聚合到同口径披露类别后比较期末数，列示差异',
        '11. 类别参数区间：折旧年限=使用月限÷12；残值率为卡片账反推值区间，供与公司折旧政策比较',
        '12. 折旧年限异常：按披露类别使用月限统计，超出 Q1-1.5×IQR ~ Q3+1.5×IQR 区间（类别样本≥5张）',
        '    的卡片标注异常（如类别大多 60 月，个别 12/240 月），供核查卡片配置是否正确',
        '13. 备注：已计提期数>使用月限/账面累计已≈原值×(1-残值率)（提前提足）等需人工复核项',
    ]
    for i, t in enumerate(notes, 1):
        w5.cell(i, 1, t)
    w5.column_dimensions['A'].width = 130

    out.save(out_path)
    print(f'✅ [{org}] 复核表已生成: {out_path}')
    print(f'   卡片 {n} 张：无差异 {n_ok} / 土地不折旧 {n_land} / 异常 {n_odd} / 折旧年限异常 {n_age}')
    print(f'   合计 原值 {s_tot["原值"]:,.2f}（期初 {s_tot["期初原值"]:,.2f} + 增加 {s_tot["增加原值"]:,.2f}）')
    print(f'        累计折旧 {s_tot["账面累计"]:,.2f}（期初 {s_tot["累计期初"]:,.2f} + 本年 {s_tot["账面本年"]:,.2f}）')
    print(f'        账面累计 vs 匡算 {s_tot["匡算累计"]:,.2f} (差 {s_tot["账面累计"] - s_tot["匡算累计"]:,.2f})')
    print(f'        账面本年 vs 匡算 {s_tot["匡算本年"]:,.2f} (差 {s_tot["账面本年"] - s_tot["匡算本年"]:,.2f})')
    return out_path


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    accts = [a for a in argv if not a.startswith('-')] or sorted(DM.jobs())
    out_override = None
    if '--out' in argv:
        out_override = argv[argv.index('--out') + 1]
    for acct in accts:
        cfg = DM.job_for(acct)
        if not cfg:
            print(f'  ⚠️ 无配置：{acct}')
            continue
        year = cfg.get('year', 2025)
        # ⚡ 铁律131b（2026-08-15 布局规范）：底稿统一输出 底稿/{year}/
        default_out = os.path.join(DM.resolve(acct), '底稿', str(year),
                                   f'固定资产折旧计提复核_{year}.xlsx')
        build(cfg, out_override or default_out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
