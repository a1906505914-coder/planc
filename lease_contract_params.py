# -*- coding: utf-8 -*-
"""lease_contract_params.py —— 从企业测算表提取 GFY 各主体合同基础参数（供 lease_review.py 手动测算）
提取字段：name, start_date, end_date, payments[(年租金不含税, 年)], period_months, payment_freq, tax, remark
输出：GFY contracts 配置（嵌入 lease_manifest.py）
"""
import os
import re
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import paths as P

import openpyxl

DATA = os.path.join(P.DATA_DIRS['GFY'], '数据', '2026', '使用权资产租赁合同')


def _f(x):
    try:
        return float(str(x).replace(',', ''))
    except (TypeError, ValueError):
        return None


def _parse_date(s):
    """'2025/6/26' / '2025.1.1' / '2025-01-01' → 'YYYY-MM-DD'"""
    m = re.search(r'(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})', str(s))
    return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}' if m else None


def _months(d1, d2):
    """两个 YYYY-MM-DD 之间月数（含首月，近似）"""
    if not d1 or not d2:
        return None
    y1, m1 = int(d1[:4]), int(d1[5:7])
    y2, m2 = int(d2[:4]), int(d2[5:7])
    return (y2 - y1) * 12 + (m2 - m1)


def tax_rate(note):
    m = re.search(r'税率(\d+)%', str(note))
    return int(m.group(1)) / 100.0 if m else 0.05   # 默认 5%


# ---------------- 集团本级（核查过程：逐年租金 + 总期限 + 税率） ----------------
def extract_jituan(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['7-2-1 集团本级']
    rows = list(ws.iter_rows(values_only=True))
    contracts = []
    cur = None
    for r in rows[105:]:
        name = str(r[0] or '').strip()
        if name and name not in ('名称', '二、核查过程') and not name.startswith('2、') \
                and not name.startswith('3、') and '合同租赁年限' not in name:
            cur = name
            contracts.append({'name': cur, 'periods': [], 'total': str(r[6] or ''), 'note': str(r[7] or '')})
        if cur:
            p = str(r[1] or '').strip()
            amt = r[4] if r[4] not in (None, '') else r[5]
            if p and amt is not None and str(amt).strip():
                a = _f(amt)
                if a:
                    contracts[-1]['periods'].append((p, a))
    out = []
    for c in contracts:
        if not c['periods']:
            continue
        tax = tax_rate(c['note'])
        # 起止：起日=首期起日（col6 优先），止日=末期止日（含续签）
        d0 = None
        tm = re.search(r'(\d{4}[./\-]\d{1,2}[./\-]\d{1,2})\s*[-—]\s*(\d{4}[./\-]\d{1,2}[./\-]\d{1,2})', c['total'])
        if tm:
            d0 = _parse_date(tm.group(1))
        if not d0:
            d0 = _parse_date(c['periods'][0][0].split('-')[0])
        d1 = _parse_date(c['periods'][-1][0].split('-')[-1])
        if not (d0 and d1) or d1 < '2026-03-31':
            continue   # ⚡ 排除 2026-03-31 前已到期的合同（2026 年无余额）
        start_year = int(d0[:4]); end_year = int(d1[:4])
        n_total = end_year - start_year + 1
        # 逐年租金（跨年段按覆盖年份填该段年租；无租金年补 0）
        yearly_full = [0.0] * max(0, n_total)
        for p, amt in c['periods']:
            y0 = _parse_date(p.split('-')[0]); y1 = _parse_date(p.split('-')[-1])
            if not (y0 and y1):
                continue
            a = round(amt / (1 + tax), 2)
            for y in range(int(y0[:4]), int(y1[:4]) + 1):
                idx = y - start_year
                if 0 <= idx < len(yearly_full):
                    yearly_full[idx] = a
        # ⚡ 只保留 2026 年起剩余租金（账面使用权资产=剩余租期现值）
        skip = max(0, 2026 - start_year)
        yearly = yearly_full[skip:] if skip < len(yearly_full) else []
        if not yearly:
            continue
        pm = _months('2026-01-01', d1) or 0
        out.append({'name': c['name'], 'start_date': '2026-01-01', 'end_date': d1,
                    'period_months': pm, 'payments': yearly, 'payment_freq': 12,
                    'tax': round(tax, 4), 'remark': f'集团本级; 剩余租期折现; 税率{tax*100:.0f}%'})
    return out


# ---------------- 海宁皮革（核查过程：①合同租赁期限 ②合同金额/每年） ----------------
def extract_haining(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['7-2-5 方圆皮革']
    rows = list(ws.iter_rows(values_only=True))
    d0 = d1 = None
    annual = None
    tax = 0.05
    for r in rows:
        lab = str(r[0] or '').strip()
        if '租赁期限' in lab:
            m = re.search(r'(\d{4}[./\-]\d{1,2}[./\-]\d{1,2})\s*[-—]\s*(\d{4}[./\-]\d{1,2}[./\-]\d{1,2})', str(r[4] or ''))
            if m:
                d0 = _parse_date(m.group(1)); d1 = _parse_date(m.group(2))
        if '合同金额' in lab:
            annual = _f(r[4])
        if '税率' in str(r[7] or ''):
            tax = tax_rate(str(r[7]))
    if not (d0 and d1 and annual):
        return []
    pm = _months('2026-01-01', d1) or 0
    n_year_total = max(1, (int(d1[:4]) - int(d0[:4]) + 1))
    yearly_full = [round(annual / (1 + tax), 2)] * n_year_total
    skip = max(0, 2026 - int(d0[:4]))
    yearly = yearly_full[skip:] if skip < len(yearly_full) else []
    if not yearly:
        return []
    return [{'name': '海宁皮革研究院大楼', 'start_date': '2026-01-01', 'end_date': d1,
             'period_months': pm, 'payments': yearly, 'payment_freq': 12,
             'tax': round(tax, 4), 'remark': f'海宁; 剩余租期折现; 税率{tax*100:.0f}%'}]


# ---------------- 绍兴方圆（4 测算 sheet：年租金 + 期数） ----------------
def extract_shaoxing(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    out = []
    for sn in ['四楼', '一、二楼', '二楼（本年新签）', '精工广场（续签）']:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        yz = None
        dep_months = None
        periods = []   # (期限标签, 不含税年租金)
        tax = 0.05
        for r in rows:
            lab = str(r[0] or '').strip()
            if '折现系数' in str(r[2] or ''):
                pass
            if lab and '年租金' not in lab and r[1] is not None and '合计' not in lab and '-' in str(lab):
                a = _f(r[1]); b = _f(r[2])
                if a:
                    periods.append((lab, a if b is None else b))   # 不含税租金
            if '折旧月数' in str(r[8] or '') or '折旧月数' in str(r[2] or ''):
                dep_months = _f(r[9]) or _f(r[3])
        # 不含税租金在 r[2]（每期支付金额），若无则用 r[1]（年租金）
        clean = []
        for r in rows:
            lab = str(r[0] or '').strip()
            if lab.startswith('20') and '-' in lab:
                a = _f(r[2]) or _f(r[1])
                if a:
                    clean.append((lab, a))
        if not clean:
            continue
        # 期数（付款期数）
        n_pay = len(clean)
        # 租赁期：从首期起日到末期止日
        d0 = _parse_date(clean[0][0].split('-')[0])
        d1 = _parse_date(clean[-1][0].split('-')[-1])
        # ⚡ 2026 年起剩余租金
        start_year = int(d0[:4]) if d0 else 2026
        skip = max(0, 2026 - start_year)
        yearly = [round(x, 2) for _, x in clean[skip:]] if skip < len(clean) else []
        if not yearly:
            continue
        pm = _months('2026-01-01', d1) or 0
        out.append({'name': sn, 'start_date': '2026-01-01', 'end_date': d1,
                    'period_months': pm,
                    'payments': yearly,
                    'payment_freq': 12, 'tax': 0.05,
                    'remark': '绍兴测算sheet; 剩余租期折现'})
    return out


# ---------------- 方圆金属 / 龙游 / 智能（参数不全，标注待补） ----------------
def extract_fallback(fp, sn, label):
    """无核查参数 → 从汇总块取 现值/付款额/期数 反推等额年租（仅作占位，标注待补）。"""
    import lease_review_gjx as LRG
    s = LRG.load_summary(fp, sn)
    if not s:
        return []
    yz = s['yz'][0]
    fk = s['fk'][0]
    wc = s['wc'][0]
    if yz <= 0:
        return []
    n = max(1, round(fk / (fk - wc) * 12) if (fk - wc) else 12)   # 期数近似 = 付款额/净额×12
    pmt = round(yz / n * 12, 2) if n else 0
    return [{'name': label, 'start_date': None, 'end_date': None,
             'period_months': n, 'payments': [pmt] * max(1, n // 12),
             'payment_freq': 12, 'tax': 0.05,
             'remark': '⚠️ 参数待补（测算表无核查过程，按现值/付款额反推等额）'}]


def main():
    root = DATA
    results = {}
    # 集团本级
    results['集团本级'] = extract_jituan(os.path.join(root, '集团本级/集团本级测算表.xlsx'))
    # 海宁皮革
    results['海宁皮革'] = extract_haining(os.path.join(root, '海宁皮革/方圆皮革测算表.xlsx'))
    # 绍兴方圆
    results['绍兴方圆'] = extract_shaoxing(os.path.join(root, '绍兴方圆/绍兴方圆测算表.xlsx'))
    # 方圆金属 / 龙游方圆 / 方圆智能
    import lease_review_gjx as LRG
    for name, rel, sn in [('方圆金属', '方圆金属/方圆金属测算表.xlsx', '7-2-10 金属材料'),
                          ('龙游方圆', '龙游方圆/龙游方圆测算表.xlsx', '7-2-11 龙游方圆'),
                          ('方圆智能', '方圆智能/方圆智能测算表.xlsx', '7-2-6 智能技术')]:
        results[name] = extract_fallback(os.path.join(root, rel), sn, name)

    import json
    for k, v in results.items():
        print(f'===== {k}: {len(v)} 个合同 =====')
        for c in v:
            print(f'  {c["name"][:14]:<14} {c.get("start_date")} ~ {c.get("end_date")} {c.get("period_months")}月 '
                  f'年租{c["payments"][0] if c["payments"] else 0:,.0f}({len(c["payments"])}期) {c["remark"][:22]}')
    return results


if __name__ == '__main__':
    main()
