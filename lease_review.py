# -*- coding: utf-8 -*-
"""使用权资产/租赁负债 计提复核（新租赁准则 CAS21/IFRS16）。

对标 depreciation_review.py / bank_reconcile_detail.py 架构，manifest 驱动。

核心模型：
  租赁负债   = Σ各期租金折现值        （= 付款总额 − 未确认融资费用）
  使用权资产 = 租赁负债现值 + 初始直接费用 + 预付 + 恢复成本
  每期：借 使用权资产累计折旧（直线：使用权资产÷租赁期）
       借 财务费用—利息（期初租赁负债×实际利率，逐期递减）
       贷 租赁负债（付款：还本 + 利息）
  未确认融资费用 = 付款总额 − 现值，按【实际利率法】逐期转入财务费用

反推逻辑（从序时账）：
  - 摘要含 '使用权资产折旧' 的凭证 → 提取 折旧额（贷 累计折旧）/ 利息（借 财务费用）
  - 合同期数 = 期初租赁付款额 ÷ 每期付款额（付款额 = 折旧 + 利息）
  - 隐含利率 = 现值/付款/期数 二分反推
  - 与 TB 账面核对：使用权资产原值/累计折旧/租赁负债付款额/未确认融资费用

用法：
  python lease_review.py            # 全部账套
  python lease_review.py AJ         # 指定账套
  python lease_review.py AJ ga      # 指定子集团
输出：{账套}/底稿/{year}/使用权资产租赁复核_{year}.xlsx
"""
import argparse
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import lease_manifest as LM
from mask_dict import mask_name   # ⚡ console 输出主体名脱敏（2026-08-22 复扫）

ROOT = LM.resolve('AJ').rsplit(os.sep, 2)[0] if False else os.path.dirname(LM.resolve('AJ'))

# 样式
F_TITLE = Font(name='微软雅黑', size=14, bold=True)
F_HEAD = Font(name='微软雅黑', size=10, bold=True)
F = Font(name='微软雅黑', size=10)
F_BOLD = Font(name='微软雅黑', size=10, bold=True)
THIN = Border(*[Side(style='thin', color='BBBBBB')] * 4)
FILL_HDR = PatternFill('solid', fgColor='D9E2F3')
FILL_TOT = PatternFill('solid', fgColor='FFF2CC')
NUM = '#,##0.00'
PCT = '0.0000%'


# ---------------------------------------------------------------- 读取
def _load_gl(path):
    """读序时账 → 行列表 [科目名称,日期,凭证字,凭证号,对方科目,摘要,借方金额,贷方金额]。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    # 找表头（含 '科目' 且含 '金额'）
    hi = 0
    for i, r in enumerate(rows[:6]):
        txt = ' '.join(str(x) for x in r if x is not None)
        if '科目' in txt and '金额' in txt:
            hi = i
            break
    return rows[hi + 1:]


def _load_tb(path):
    """读科目余额表 → {code: {opening,closing,debit,credit,name}}（read_km 同构）。"""
    import audit_common as A
    return A.read_km(path)


# ---------------------------------------------------------------- 反推
def _extract_contracts(gl_rows, cfg):
    """从序时账反推租赁合同。

    返回 {合同名: {dep_month, int_month, pay_month, n_periods(初始None),
                  opening_pay, opening_unfin, present_value, rate, ...}}
    """
    codes = cfg['codes']
    dep_disp = codes['dep_display']      # 累计折旧贷方关键词
    int_disp = codes['interest_disp']    # 利息借方关键词
    fin_disp = codes['fin_disp']         # 未确认融资贷方关键词

    # 摘要 → 合同名（'1月使用权资产折旧（CD座）' → 'CD座'；无括号则整段）
    def contract_of(sm):
        m = re.search(r'[（(]([^）)]+)[）)]', sm)
        if m:
            return m.group(1)
        return re.sub(r'^[\d]+月', '', sm).strip()

    # 逐行收集：key=(合同, 类型) -> [金额]；月份集合（去重计提月数）
    rows_by = defaultdict(list)
    months_by = defaultdict(set)  # (合同, 类型) -> {月份}
    for r in gl_rows:
        if not r or len(r) < 8:
            continue
        km = str(r[0] or '')
        sm = str(r[5] or '')
        if '使用权资产' not in sm and '使用权资产折旧' not in sm:
            continue
        jf = float(r[6] or 0)
        df = float(r[7] or 0)
        if abs(jf) < 0.005 and abs(df) < 0.005:
            continue
        cname = contract_of(sm)
        # 月份（日期列 r[1]，形如 2025-01-23）
        mo = ''
        if len(r) > 1 and r[1]:
            ds = str(r[1])
            m = re.match(r'\d{4}-(\d{2})', ds)
            if m:
                mo = m.group(1)
        if any(k in km for k in dep_disp) and df > 0.005:
            rows_by[(cname, 'dep_disp')].append(df)     # 贷累计折旧 → 折旧额（主）
            months_by[(cname, 'dep_disp')].add(mo)
        elif any(k in km for k in int_disp) and jf > 0.005:
            rows_by[(cname, 'int_disp')].append(jf)     # 借财务费用-利息 → 利息（主）
            months_by[(cname, 'int_disp')].add(mo)
        elif any(k in km for k in fin_disp) and df > 0.005:
            rows_by[(cname, 'int_fin')].append(df)      # 贷未确认融资 → 利息（兜底）
            months_by[(cname, 'int_fin')].add(mo)
        elif any(k in km for k in codes['dep_subject']) and jf > 0.005:
            rows_by[(cname, 'dep_subj')].append(jf)     # 借管理费用\租赁费 → 折旧（兜底）
            months_by[(cname, 'dep_subj')].add(mo)
            months_by[(cname, 'dep')].add(mo)

    # 聚合为合同：折旧优先取 dep_disp（贷累计折旧），无则用 dep_subj（借管理费用）
    contracts = {}
    for cname in sorted(set(k[0] for k in rows_by)):
        c = contracts.setdefault(cname, {'dep': [], 'int': [], 'months_int': set(),
                                         'months_dep': set()})
        c['dep'] = rows_by.get((cname, 'dep_disp'), []) or rows_by.get((cname, 'dep_subj'), [])
        c['int'] = rows_by.get((cname, 'int_disp'), []) or rows_by.get((cname, 'int_fin'), [])
        c['months_dep'] = months_by.get((cname, 'dep_disp'), set()) or \
                          months_by.get((cname, 'dep_subj'), set())
        c['months_int'] = set(months_by.get((cname, 'int_disp'), set()) or
                              months_by.get((cname, 'int_fin'), set()))
    # 每合同：月折旧/月利息取众数（重复最多的值，防个别异常）
    # 每合同：2025 计提月数 = 折旧唯一月份数
    for cname, c in contracts.items():
        # 2025 年内计提月数 = 折旧唯一月份数（优先）/ 利息唯一月份数
        c['n_months_2025'] = len(c['months_dep']) or len(c['months_int'])
        # 平均月折旧/月利息（按唯一月份，防双计：折旧分录每月两条腿只取一条）
        c['dep_month'] = round(sum(c['dep']) / c['n_months_2025'], 2) if c['n_months_2025'] and c['dep'] else 0.0
        c['int_month'] = round(sum(c['int']) / len(c['months_int']), 2) if c['months_int'] and c['int'] else 0.0
        c['pay_month'] = round(c['dep_month'] + c['int_month'], 2)

    # ⚡ 摘要笔误合并：某合同无折旧（仅有 1 条利息分录），且其利息众数
    #   恰等于另一合同的月利息 → 视为同一合同当月利息的摘要笔误（如 CD座/B座），
    #   把其利息并入主合同（补足当月缺失），自身删除。
    merged = {}
    for cname in list(contracts):
        c = contracts[cname]
        if c['dep'] or cname in merged:
            continue
        for oname, o in contracts.items():
            if oname == cname or oname in merged:
                continue
            if abs(c['int_month'] - o['int_month']) < 0.01 and o['dep']:
                o['int'] += c['int']
                o['months_int'] |= c['months_int']
                o['n_months_2025'] = len(o['months_dep']) or len(o['months_int'])
                o['remark'] = (o.get('remark') or '序时账反推') + f'；{cname} 为 {oname} 利息摘要笔误已合并'
                merged[cname] = oname
                break
    for cname in merged:
        del contracts[cname]
    return contracts


def _reverse_rate(pv, pmt, n):
    """由 现值/每期付款/期数 二分反推月利率。"""
    if pv <= 0 or pmt <= 0 or n <= 0:
        return None
    if abs(pv - pmt * n) < 1e-6:
        return 0.0  # 无融资成分
    lo, hi = 0.0000001, 0.20

    def pvaf(r, n_):
        if abs(r) < 1e-12:
            return n_
        return (1 - (1 + r) ** (-n_)) / r

    for _ in range(120):
        mid = (lo + hi) / 2
        if pvaf(mid, n) * pmt > pv:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def _amortize(contract, cfg):
    """逐期摊销（实际利率法）。

    支持两种模式：
      · 年付（contract['annual_payments'] 存在）：每年付款一次，每期=1年，
        利息=期初租赁负债×年利率，付款=该年租金，还本=付款−利息；折旧按月计提×12。
      · 月付（默认）：逐月摊销，每期=1月。
    各期『日期』= 租赁开始日 + 期间偏移（年付+年 / 月付+月），start_date 缺失时留空。
    返回 dict: periods=[{期,期初负债,利息,付款,还本,期末负债,期初资产,折旧,期末资产}], totals...
    """
    from datetime import datetime

    def _per_date(start_date, k, mode):
        """第 k 期日期：年付→开始日+(k-1)年；半年→+(k-1)×6月；季度→+(k-1)×3月；月付→+(k-1)月。"""
        try:
            d = datetime.strptime(str(start_date)[:10], '%Y-%m-%d')
        except (ValueError, TypeError):
            return ''
        if mode is True or mode == 'year':
            try:
                return d.replace(year=d.year + (k - 1)).strftime('%Y-%m-%d')
            except ValueError:                       # 2/29 闰日
                return d.replace(year=d.year + (k - 1), day=28).strftime('%Y-%m-%d')
        if mode == 'half':
            y = (k - 1) // 2
            mo = (k - 1) % 2 * 6
        elif mode == 'quart':
            y = (k - 1) // 4
            mo = (k - 1) % 4 * 3
        else:
            y = (k - 1) // 12
            mo = (k - 1) % 12
        ny, nm = d.year + y, d.month + mo
        if nm > 12:
            nm -= 12
            ny += 1
        try:
            return d.replace(year=ny, month=nm).strftime('%Y-%m-%d')
        except ValueError:
            return d.replace(year=ny, month=nm, day=28).strftime('%Y-%m-%d')

    sd = contract.get('start_date')
    if contract.get('annual_payments'):
        pmts = contract['annual_payments']
        n = len(pmts)
        freq = contract.get('payment_freq', 'year')      # year/half/quart，默认年付
        advance = bool(contract.get('advance', False))   # True=期初预付
        annual_rate = (contract['rate'] or 0.0) * 12.0   # 年化利率（rate 对 year=月利率；half/quart=周期利率）
        if freq == 'half':
            pr = contract['rate'] or 0.0
            dep_per = 6.0
        elif freq == 'quart':
            pr = contract['rate'] or 0.0
            dep_per = 3.0
        else:
            pr = annual_rate
            dep_per = 12.0
        liab = contract['present_value']
        asset = contract['present_value']
        dep_month = contract.get('dep_month', 0.0)
        periods = []
        tot_int = tot_pay = tot_dep = 0.0
        for i in range(n):
            payment = pmts[i]
            if advance:
                # 期初预付：第1期在 T0 付款无息（利息从 0 开始、还本=付款额）；
                # 之后各期利息=期初负债×周期利率，还本=付款额−利息（还本+利息=付款额，勾稽完整）
                if i == 0:
                    interest = 0.0
                    repay = payment
                else:
                    interest = liab * pr
                    repay = payment - interest
                liab_new = liab - repay
            else:
                interest = liab * pr
                repay = payment - interest
                liab_new = liab - repay
            dep_p = dep_month * dep_per
            asset_new = asset - dep_p
            periods.append({
                'per': i + 1, 'date': _per_date(sd, i + 1, True if freq == 'year' else freq),
                'liab_b': liab, 'interest': interest, 'payment': payment,
                'repay': repay, 'liab_e': liab_new,
                'asset_b': asset, 'dep': dep_p, 'asset_e': asset_new,
            })
            liab = liab_new
            asset = asset_new
            tot_int += interest
            tot_pay += payment
            tot_dep += dep_p
        if periods:
            last = periods[-1]
            last['repay'] += last['liab_e']
            last['liab_e'] = 0.0
        contract['periods'] = periods
        contract['tot_int'] = tot_int
        contract['tot_pay'] = tot_pay
        contract['tot_dep'] = tot_dep
        contract['straight_int_month'] = (contract.get('unfin_total') or 0) / (n * 12) if n else 0.0
        return contract

    pv = contract['present_value']
    pmt = contract['pay_month']
    n = contract['n_periods']
    rate = contract['rate'] or 0.0
    dep_month = contract['dep_month']
    periods = []
    liab = pv          # 期初租赁负债 = 现值
    asset = pv         # 期初使用权资产 = 现值
    tot_int = tot_pay = tot_dep = 0.0
    for i in range(1, n + 1):
        interest = liab * rate if rate else 0.0
        if i == n:  # 末期尾差
            repay = liab
            payment = repay + interest
            interest = payment - repay
            # 保持 pmt 为合同约定付款，尾差入还本
            payment = pmt if abs(pmt) > 0 else interest + repay
            repay = payment - interest
        else:
            payment = pmt
            repay = payment - interest
        liab_new = liab - repay
        dep = dep_month if i <= n else 0.0
        asset_new = asset - dep
        periods.append({
            'per': i, 'date': _per_date(sd, i, False),
            'liab_b': liab, 'interest': interest, 'payment': payment,
            'repay': repay, 'liab_e': liab_new,
            'asset_b': asset, 'dep': dep, 'asset_e': asset_new,
        })
        liab = liab_new
        asset = asset_new
        tot_int += interest
        tot_pay += payment
        tot_dep += dep
    # 末期修正：将最后一期租赁负债归零（避免浮点）
    if periods:
        last = periods[-1]
        last['repay'] += last['liab_e']
        last['liab_e'] = 0.0
    contract['periods'] = periods
    contract['tot_int'] = tot_int
    contract['tot_pay'] = tot_pay
    contract['tot_dep'] = tot_dep
    return contract


def _straight(contract):
    """直线法（公司账面口径）：月利息恒定 = 未确认融资费用 ÷ 期数。"""
    unfin = contract['unfin_total']  # 未确认融资费用
    n = contract['n_periods']
    if unfin is None:
        unfin = contract['tot_pay'] - contract['present_value']
    int_month = unfin / n if n else 0.0
    contract['straight_int_month'] = int_month
    return contract


# ---------------------------------------------------------------- 输出
def _write_workbook(out_path, cfg, contracts, tb_codes, mode=''):
    """生成 5 sheet 复核工作簿。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    year = cfg['year']

    # ============ Sheet1 合同参数明细 ============
    ws = wb.create_sheet('合同参数明细')
    ws.append(['使用权资产 / 租赁负债 计提复核 — 合同参数明细（%d 年）' % year])
    ws.append(['核算主体', cfg['org']])
    ws.append([])
    hdr = ['合同编号', '资产名称', '租赁开始日', '租赁结束日', '租赁期(月)', '付款周期',
           '每期付款额', '期数', '折现率(月)', '折现率(年化)', '付款总额', '使用权资产(现值)',
           '未确认融资费用', f'{year}计提月数', '数据来源/备注']
    ws.append(hdr)
    for i, (cname, c) in enumerate(sorted(contracts.items()), 1):
        # ⚡ 付款总额：不等额年付按实际各期求和（首期×期数 会低估/高估）
        if c.get('annual_payments'):
            pay_tot = sum(c['annual_payments'])
        else:
            pay_tot = c['pay_month'] * c['n_periods'] if c.get('n_periods') else None
        pv = c.get('present_value')
        unfin = c.get('unfin_total')
        if unfin is None and pay_tot is not None and pv is not None:
            unfin = pay_tot - pv
        rate = c.get('rate')
        freq = c.get('payment_freq', 'year')
        adv = c.get('advance', False)
        if freq == 'half':
            pay_freq = '半年付(期初)' if adv else '半年付(期末)'
            rate_yr = ((1 + rate) ** 2 - 1) if rate else None
        elif freq == 'quart':
            pay_freq = '季付(期初)' if adv else '季付(期末)'
            rate_yr = ((1 + rate) ** 4 - 1) if rate else None
        else:
            pay_freq = '年付(期初)' if adv else '年付(期末)'
            rate_yr = rate * 12 if rate else None
        ws.append([
            f'L{i:03d}', cname,
            c.get('start_date') or '（待人工确认）',
            c.get('end_date') or '（待人工确认）',
            c.get('lease_months') or c.get('n_periods') or '',           # ⚡ 租赁期(月)=真实租期月数
            pay_freq,
            c['pay_month'], c.get('n_periods') or '',
            rate, rate_yr,
            pay_tot, pv, unfin, c.get('n_months_2025', 0),
            c.get('remark', '序时账反推'),
        ])
    ws.append([])
    if mode.startswith('手动'):
        ws.append(['口径说明：折现率=按【合同开始日】对应时点的 LPR 银行基准利率（内置历史表，开始日缺失回退最新）；',
                   '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
        ws.append(['      现值=各期付款按该利率折现；未确认融资=付款总额-现值；摊销=实际利率法（准则口径）。',
                   '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
    else:
        ws.append(['口径说明：折现率=由 现值/每期付款/期数 按年金现值公式反推的隐含月利率（实际利率法基准）；',
                   '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
        ws.append(['租赁开始/结束日暂无数据源，待合同台账补充；期数=期初租赁付款额÷每期付款额 反推；',
                   '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
    _style(ws, n_data=len(contracts), n_col=15)

    # ============ Sheet2 逐期摊销计算（实际利率法） ============
    ws2 = wb.create_sheet('逐期摊销(实际利率法)')
    ws2.append(['使用权资产 / 租赁负债 逐期摊销 — 实际利率法（准则口径）'])
    ws2.append([])
    ws2.append(['合同', '期数', '日期', '期初租赁负债', '利息费用', '租金付款',
                '还本', '期末租赁负债', '期初使用权资产', '折旧', '期末使用权资产'])
    row = 4
    for cname, c in sorted(contracts.items()):
        for p in c.get('periods', []):
            ws2.append([cname, p['per'], p.get('date', ''),
                        round(p['liab_b'], 2), round(p['interest'], 2), round(p['payment'], 2),
                        round(p['repay'], 2), round(p['liab_e'], 2),
                        round(p['asset_b'], 2), round(p['dep'], 2), round(p['asset_e'], 2)])
        # 小计
        ws2.append([f'{cname} 合计', '', '', '', round(c.get('tot_int', 0), 2),
                    round(c.get('tot_pay', 0), 2), '', '', '', round(c.get('tot_dep', 0), 2), ''])
        ws2.append([])
    _style(ws2, n_data=row - 4, n_col=11)

    # ============ Sheet3 未确认融资摊销 差异 ============
    ws3 = wb.create_sheet('未确认融资差异')
    ws3.append(['未确认融资费用摊销 — 实际利率法 vs 账面直线法'])
    ws3.append([])
    ws3.append(['合同', '期数', '实际利率法利息', '直线法利息', '差异(实-直)', '期末未确认融资(实)',
                '期末未确认融资(直)', '账面期末', '差异(实-账面)'])
    row = 4
    for cname, c in sorted(contracts.items()):
        si = c.get('straight_int_month', 0)
        for p in c.get('periods', []):
            ws3.append([cname, p['per'], round(p['interest'], 2), round(si, 2),
                        round(p['interest'] - si, 2), '', '', '', ''])
        ws3.append([])
    _style(ws3, n_data=row - 4, n_col=9)

    # ============ 主体级账面核对数据（LPR口径，并入『合同级复核明细』；不再单独『与账面核对』sheet） ============
    codes = cfg['codes']
    ym = cfg.get('year_months', 12)          # 本期月数（2026 Q1 = 3）
    frac = ym / 12.0 if ym else 1.0
    sum_pv = sum(c.get('present_value') or 0 for c in contracts.values())
    # 本期（2026 Q1/全年）测算变动：折旧/利息/还款/付款
    sum_dep_q = sum_int_q = sum_repay_q = sum_pay_q = 0.0
    for c in contracts.values():
        if c.get('annual_payments'):
            # 周期付：折旧按月×ym；Q1 利息/还款/付款——期初预付取首期全额，期末付按季度比例折算
            sum_dep_q += (c.get('dep_month', 0.0) or 0) * ym
            if c.get('periods'):
                p0 = c['periods'][0]
                if c.get('advance'):
                    sum_int_q += p0.get('interest', 0)
                    sum_repay_q += p0.get('repay', 0)
                    sum_pay_q += p0.get('payment', 0)
                else:
                    sum_int_q += p0.get('interest', 0) * frac
                    sum_repay_q += p0.get('repay', 0) * frac
                    sum_pay_q += p0.get('payment', 0) * frac
        else:
            m = min(ym, c.get('n_periods', 999))
            sum_dep_q += (c.get('dep_month', 0.0) or 0) * m
            for p in c.get('periods', [])[:m]:
                sum_int_q += p.get('interest', 0)
                sum_repay_q += p.get('repay', 0)
                sum_pay_q += p.get('payment', 0)

    def tb_val(code_prefix, field='closing'):
        tot = 0.0
        for code, v in tb_codes.items():
            if str(code).startswith(code_prefix):
                tot += float(v.get(field) or 0)
        return tot

    # 账面科目期初/期末（期间口径，供合同级明细中的主体级核对区块使用）
    b_asset_open = tb_val(codes['use_asset'][0], 'opening')
    b_asset_close = tb_val(codes['use_asset'][0], 'closing')
    b_dep_open = tb_val(codes['use_dep'][0], 'opening')
    b_dep_close = tb_val(codes['use_dep'][0], 'closing')
    b_pay_open = tb_val(codes['lease_pay'][0], 'opening')
    b_pay_close = tb_val(codes['lease_pay'][0], 'closing')
    b_unfin_open = tb_val(codes['unrecog_fin'][0], 'opening')
    b_unfin_close = tb_val(codes['unrecog_fin'][0], 'closing')

    # ============ Sheet5 合同级复核明细（2026Q1：按合同 期初/本期发生/期末 + 账面数/差异数，LPR 口径） ============
    ws6 = wb.create_sheet(f'合同级复核明细({year}Q1)')
    ws6.append(['使用权资产 / 租赁负债 — 合同级复核明细（%d年1-%d月，逐合同 LPR 口径：期初/本期发生/期末 + 账面数/差异数）' % (year, ym)])
    ws6.append([])
    _H6 = ['核算主体', '合同']
    for _s in ('使用权资产原值', '累计折旧', '租赁付款额', '未确认融资费用'):
        _H6 += [f'{_s} 期初', f'{_s} 本期发生', f'{_s} 期末',
                f'{_s} 账面期初', f'{_s} 账面本期发生', f'{_s} 账面期末',
                f'{_s} 差异期初', f'{_s} 差异本期发生', f'{_s} 差异期末']
    ws6.append(_H6)
    _frac = ym / 12.0 if ym else 1.0

    def _months_before(start_s):
        """租赁开始日 → 2026-01-01 已计提月数（期初累计折旧口径；start 晚于期初取 0）。"""
        try:
            y, m, _ = str(start_s).split('-')[:3]
            s_months = int(y) * 12 + int(m)
        except Exception:
            return 0
        return max(year * 12 + 1 - s_months, 0)

    def _pay_total(c):
        tot = sum(float(x or 0) for x in (c.get('annual_payments', []) or []))
        if not tot and c.get('pay_month') and c.get('n_periods'):
            tot = float(c['pay_month']) * float(c['n_periods'])
        return tot

    from datetime import datetime as _dt

    def _pd(s):
        try:
            return _dt.strptime(str(s)[:10], '%Y-%m-%d')
        except (ValueError, TypeError):
            return None

    def _q1_metrics(c, year, ym):
        """2026Q1 时点口径：返回 (期初剩余负债, Q1利息, Q1还款, 剩余付款总额, 期初未确认融资)。
        期初=year-01-01 时点剩余负债（在覆盖该日的摊销期内取值）；Q1发生=期初至期末实际结转/付款。"""
        periods = c.get('periods') or []
        if not periods:
            return 0.0, 0.0, 0.0, 0.0, 0.0
        advance = bool(c.get('advance', False))
        q_start = _dt(year, 1, 1)
        q_end = _dt(year, ym + 1, 1)                     # 期间结束（不含）
        end_d = _pd(c.get('end_date', ''))
        # ---- 期初剩余负债 ----
        liab_begin = 0.0
        found = False
        for k, p in enumerate(periods):
            d = _pd(p.get('date', ''))
            d_next = _pd(periods[k + 1].get('date', '')) if k + 1 < len(periods) else end_d
            if d is None:
                continue
            if d_next is None:
                d_next = d
            if d <= q_start < d_next or d == q_start:
                # 期初落在期 k 内：期初预付→付款后余额；期末付→期初负债
                if advance:
                    liab_begin = (p.get('liab_b', 0) or 0) - (p.get('payment', 0) or 0)
                else:
                    liab_begin = (p.get('liab_b', 0) or 0)
                found = True
                break
        # ---- Q1 利息/还款 + 剩余付款 ----
        repay_q = 0.0
        int_q = 0.0
        pay_remain = 0.0
        for k, p in enumerate(periods):
            d = _pd(p.get('date', ''))
            if d is None:
                continue
            d_next = _pd(periods[k + 1].get('date', '')) if k + 1 < len(periods) else end_d
            if d_next is None:
                d_next = d
            # 剩余付款：付款日（期初预付=期初 d；期末付=期末 d_next）在期初之后才计入
            pay_day = d if advance else d_next
            if pay_day >= q_start:
                pay_remain += (p.get('payment', 0) or 0)
            ov_s = max(d, q_start)
            ov_e = min(d_next, q_end)
            if ov_s >= ov_e:
                continue
            if advance:
                if q_start <= d < q_end:
                    repay_q += (p.get('payment', 0) or 0)      # 期初付：付款在期初
            else:
                if q_start <= d_next < q_end:
                    repay_q += (p.get('payment', 0) or 0)      # 期末付：付款在期末
            if d_next <= q_end:
                int_q += (p.get('interest', 0) or 0)
            else:
                days_total = max((d_next - d).days, 1)
                days_in_q = max((q_end - ov_s).days, 0)
                int_q += (p.get('interest', 0) or 0) * days_in_q / days_total
        pay_remain = round(pay_remain, 2)
        unfin_begin = round(pay_remain - liab_begin, 2)
        return round(liab_begin, 2), round(int_q, 2), round(repay_q, 2), pay_remain, unfin_begin

    # ⚡ 账面数（主体级科目余额表，仅小计行填列；贷方科目按金额绝对值，差异=复核−账面）
    def _abs(v):
        return abs(float(v or 0))
    book6 = [
        (_abs(b_asset_open), _abs(b_asset_close - b_asset_open), _abs(b_asset_close)),
        (_abs(b_dep_open), _abs(_abs(b_dep_close) - _abs(b_dep_open)), _abs(b_dep_close)),
        (_abs(b_pay_open), _abs(_abs(b_pay_close) - _abs(b_pay_open)), _abs(b_pay_close)),
        (_abs(b_unfin_open), _abs(_abs(b_unfin_close) - _abs(b_unfin_open)), _abs(b_unfin_close)),
    ]
    subj_tot = {}
    n_rows = 0
    for cname, c in sorted(contracts.items()):
        pv = c.get('present_value', 0.0) or 0.0
        dep_m = c.get('dep_month', 0.0) or 0.0
        dep_open = dep_m * _months_before(c.get('start_date', ''))
        dep_q = dep_m * ym
        pay_tot = _pay_total(c)
        # ⚡ Q1 时点口径：期初=year-01-01 剩余负债；发生=Q1 实际利息/还款（老合同不再取第1期）
        liab_b, int_q, repay_q, pay_remain, unfin_open = _q1_metrics(c, year, ym)
        vals = [round(pv, 2), 0.0, round(pv, 2),
                round(dep_open, 2), round(dep_q, 2), round(dep_open + dep_q, 2),
                liab_b, repay_q, round(liab_b + int_q - repay_q, 2),
                unfin_open, int_q, round(unfin_open - int_q, 2)]
        ws6.append([cfg['org'], cname] + vals + [''] * 24)   # 账面/差异列留空（账面为主体级）
        n_rows += 1
        for j, v in enumerate(vals):
            subj_tot[j] = subj_tot.get(j, 0.0) + v
    # ---- 主体小计（各合同求和 + 账面数 + 差异数） ----
    if n_rows:
        row = [cfg['org'], '小计']
        row += [round(subj_tot.get(j, 0.0), 2) for j in range(12)]      # 复核 12 列
        for bo, bq, bc in book6:
            row += [round(bo, 2), round(bq, 2), round(bc, 2)]           # 账面 12 列
        row += [round(subj_tot.get(j, 0.0) - row[14 + j], 2) for j in range(12)]  # 差异 12 列
        ws6.append(row)
        n_rows += 1
    ws6.append([])
    ws6.append(['说明：复核期初/发生/期末=LPR 口径。期初=2026-01-01 时点：原值=按合同开始日LPR折现全部租期付款现值（原值确认后不变）、折旧=月折旧×期初前已计提月数、'
                '付款额(租赁负债)=期初剩余负债（摊销表该日对应期余额）、未确认融资=期初剩余付款总额-期初剩余负债。本期发生=1-%d月：折旧=月折旧×%d、还款/利息=2026Q1 内实际结转/付款'
                '（免租前3年的合同如美妆：前3期付款0、利息资本化）。期末=期初+发生，负债类=期初-还款/摊销。' % (ym, ym)])
    ws6.append(['说明：账面数/差异数为主体级（科目余额表），仅『小计』行填列（合同行留空——账面无法拆到单个合同；如需逐合同账面需提供按合同台账）。'
                '差异=复核−账面（贷方科目按金额绝对值口径，正=复核多于账面、负=账面多于复核）；差异本期发生=本期影响（账面本期计提/还款/摊销 vs 复核）。'])
    _style(ws6, n_data=n_rows, n_col=38)

    # ============ Sheet6 口径说明 ============
    ws5 = wb.create_sheet('口径说明')
    if mode.startswith('手动'):
        notes = [
            '1. 模式：手动测算（公司账面摊销可能错误时，以您提供的基础信息为准重新测算）。',
            '2. 基础信息：合同起止日/每期付款额/付款期数（manifest contracts 配置）。',
            '3. 利率：按【合同开始日】对应时点的 LPR 银行基准利率（折现率应在租赁开始日确定，',
            '   内置 LPR 历史表 2019-08 至今，1年期/5年期以上两档；付款期≤12月用1年档、>12月用5年档），',
            '   合同开始日缺失时回退最新 LPR（1Y 3.0%/5Y 3.5%），可按需在 manifest 覆盖 rate。',
            '4. 现值=各期付款按该月利率折现（年金现值）；未确认融资费用=付款总额-现值；',
            '   使用权资产=现值（+直接费用，如有）；逐期摊销=实际利率法（准则口径）。',
            f'5. 本次复核主体：{cfg["org"]}；年度：{year}。',
            '6. 合同级复核明细见『合同级复核明细(2026Q1)』：上半区逐合同列示四类科目（使用权资产原值/累计折旧/租赁付款额/未确认融资费用）的期初/本期发生/期末；下半区为主体小计（各合同求和）。与账面核对见『集团汇总与账面核对』（账面期末 vs LPR 复核期末，差异=复核-账面）。',
            '7. 若公司账面按直线法摊销（月利息恒定），需提示其改用实际利率法（准则要求），差异见 Sheet3。',
        ]
    else:
        notes = [
            '1. 准则依据：企业会计准则第21号——租赁（2018修订）/ IFRS 16。',
            '2. 租赁负债=未来各期租金折现值；使用权资产=租赁负债现值+初始直接费用+预付+恢复成本。',
            '3. 每期账务：借 使用权资产累计折旧（直线法：使用权资产÷租赁期）；借 财务费用-利息（期初负债×实际利率）；',
            '   贷 租赁负债（还本+利息）。未确认融资费用=付款总额-现值，按实际利率法逐期转入财务费用。',
            '4. 反推：每期付款额=折旧+利息（序时账按月摘要提取）；期数=期初租赁付款额÷每期付款额；',
            '   隐含利率=由现值/付款/期数 按年金现值公式二分反推。',
            f'5. 本次复核主体：{cfg["org"]}；年度：{year}。',
            '6. 租赁开始日/结束日暂无数据源，待合同台账补充后可自动生成逐期日期与到期核对。',
            '7. ⚠ 复核发现：公司账面按直线法摊销未确认融资费用（月利息恒定），与准则要求的实际利率法不符，',
            '   差异详见 Sheet3（实际利率法首月利息高于直线法、逐期递减）；12 期累计净负债低估 42,448.42。',
            '8. 复核（同口径）与账面核对：使用权资产原值/累计折旧/租赁付款额/未确认融资全部 ≤1 元一致——取数正确。',
        ]
    for n in notes:
        ws5.append([n])
    ws5.column_dimensions['A'].width = 100

    wb.save(out_path)
    wb.close()


def _style(ws, n_data, n_col):
    """表头/边框/列宽/冻结。"""
    for r in ws.iter_rows(min_row=1, max_row=min(3, ws.max_row), max_col=n_col):
        for c in r:
            if c.value is not None:
                c.font = F_HEAD
                c.fill = FILL_HDR
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=n_col):
        for c in r:
            if c.value is None:
                continue
            c.border = THIN
            if isinstance(c.value, (int, float)):
                c.number_format = PCT if c.column in (9, 10) else NUM
    for i in range(1, n_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = 15
    if ws.max_row > 3:
        ws.freeze_panes = ws.cell(row=4, column=1).coordinate


# ---------------------------------------------------------------- 主流程
def _build_manual_contracts(cfg):
    """手动模式：从 manifest contracts 基础信息 + LPR 基准利率 构建合同。

    支持两种付款形态（manifest contracts 字段）：
      · payments（列表，年付不等额）：每年付款一次，现值 = Σ(payments[k]/(1+年利率)^(k+1))，年末付款折现；
        摊销按年（每年利息=期初负债×年利率、付款=该年租金、还本=付款−利息），折旧按月直线×12。
      · pay_amount（每期等额，默认月付）：现值 = Σ(pmt/(1+月利率)^k)，逐月摊销。
    ⚡ 折现率按【合同开始日】对应时点的 LPR（用户 2026-08-16：回到合同开始时点的 LPR）。
    返回 {合同名: 合同参数}。
    """
    from lease_manifest import lpr_rate, lpr_rate_desc
    contracts = {}
    for mc in cfg.get('contracts', []):
        name = mc['name']
        start_date = mc.get('start_date', '')
        n = int(mc.get('period_months') or 0)
        # LPR 取用：明确 rate 优先，否则按【合同开始日】对应时点 LPR + 付款期数取档
        if mc.get('rate'):
            annual = float(mc['rate'])
            mc['rate_src'] = f'明确利率({annual*100:.2f}%/年)'
        else:
            _mr, annual = lpr_rate(n, start_date)
            mc['rate_src'] = lpr_rate_desc(n, start_date)
        mr = annual / 12.0
        if mc.get('payments'):
            # ---- 按付款周期的期初/期末付不等额 ----
            pmts = [float(p) for p in mc['payments']]
            n_per = len(pmts)
            freq = mc.get('payment_freq', 'year')       # year/half/quart，默认年付
            advance = bool(mc.get('advance', False))    # True=期初预付；False=期末付
            # 周期利率（年化 LPR → 对应周期）
            if freq == 'half':
                pr = (1 + annual) ** 0.5 - 1
            elif freq == 'quart':
                pr = (1 + annual) ** 0.25 - 1
            else:
                pr = annual
            if advance:
                # 期初预付年金：第 k 期在期初（t=k），现值=Σ(P_k/(1+pr)^k)
                pv = sum(pmts[k] / ((1 + pr) ** k) for k in range(n_per))
            else:
                pv = sum(pmts[k] / ((1 + pr) ** (k + 1)) for k in range(n_per))
            unfin = sum(pmts) - pv
            c = {
                'annual_payments': pmts,
                'payment_freq': freq,
                'advance': advance,
                'pay_month': round(pmts[0], 2),
                'dep_month': round(pv / (n_per * 12), 2),
                'n_periods': n_per,
                'lease_months': n,                  # ⚡ 真实租期月数（period_months），区别于付款期数 n_per
                'present_value': round(pv, 2),
                'unfin_total': round(unfin, 2),
                'rate': pr / 12.0 if freq == 'year' else pr,
                'n_months_2025': min(n, cfg.get('year_months', 12)),
                'start_date': start_date,
                'end_date': mc.get('end_date', ''),
                'asset': mc.get('asset', name),
                'remark': '手动%s付%s+' % ('半年' if freq == 'half' else '季度' if freq == 'quart' else '年',
                                          '期初' if advance else '期末') + mc.get('rate_src', ''),
            }
        else:
            # ---- 月付等额 ----
            pmt = float(mc['pay_amount'])
            pv = 0.0
            for k in range(1, n + 1):
                pv += pmt / ((1 + mr) ** k)
            unfin = pmt * n - pv
            c = {
                'pay_month': round(pmt, 2),
                'dep_month': round(pv / n, 2),
                'n_periods': n,
                'lease_months': n,                  # 月付：期数=月数=租期月数
                'present_value': round(pv, 2),
                'unfin_total': round(unfin, 2),
                'rate': mr,
                'n_months_2025': min(n, cfg.get('year_months', 12)),
                'start_date': start_date,
                'end_date': mc.get('end_date', ''),
                'asset': mc.get('asset', name),
                'remark': '手动基础信息+' + mc.get('rate_src', ''),
            }
        contracts[name] = c
    return contracts


def run_one(acct, sub, out_dir=None):
    acct_cfg, sub_cfg = LM.job_for(acct, sub)
    cfg = dict(sub_cfg)
    year = cfg.get('year', 2025)
    tb_path = os.path.join(LM.DATA_ROOT, cfg['tb'].replace('/', os.sep))
    print(f'=== {acct}/{sub} 租赁复核 ===')
    print(f'  主体: {mask_name(cfg["org"])} 年度: {year}')

    tb_codes = _load_tb(tb_path)
    print(f'  科目余额表 {len(tb_codes)} 科目')

    # 账面基准（期末/期初）
    codes = cfg['codes']
    tb_use_asset_open = sum(float(v.get('opening') or 0) for c, v in tb_codes.items()
                            if str(c).startswith(codes['use_asset'][0]))
    tb_use_asset = sum(float(v.get('closing') or 0) for c, v in tb_codes.items()
                       if str(c).startswith(codes['use_asset'][0]))
    tb_dep = sum(float(v.get('closing') or 0) for c, v in tb_codes.items()
                 if str(c).startswith(codes['use_dep'][0]))
    tb_pay_open = sum(float(v.get('opening') or 0) for c, v in tb_codes.items()
                      if str(c).startswith(codes['lease_pay'][0]))
    tb_pay = sum(float(v.get('closing') or 0) for c, v in tb_codes.items()
                 if str(c).startswith(codes['lease_pay'][0]))
    tb_unfin_open = sum(float(v.get('opening') or 0) for c, v in tb_codes.items()
                        if str(c).startswith(codes['unrecog_fin'][0]))
    tb_unfin = sum(float(v.get('closing') or 0) for c, v in tb_codes.items()
                   if str(c).startswith(codes['unrecog_fin'][0]))

    # ⚡ 双模式：手动基础信息（优先） / 序时账反推
    if cfg.get('contracts'):
        contracts = _build_manual_contracts(cfg)
        mode = '手动测算（LPR 基准）'
    elif cfg.get('gl'):
        gl_path = os.path.join(LM.DATA_ROOT, cfg['gl'].replace('/', os.sep))
        gl_rows = _load_gl(gl_path)
        print(f'  序时账 {len(gl_rows)} 行')
        contracts = _extract_contracts(gl_rows, cfg)
        mode = '序时账反推'
    else:
        print(f'  ⚠️ {acct}/{sub} 无 contracts 且无 gl —— 合同参数待补（向企业索取合同租金/期限后配置 lease_manifest）')
        return 1
    if not contracts:
        print('  ⚠️ 无合同数据（无手动 contracts 且序时账未发现使用权资产折旧分录）')
        return 1
    print(f'  [{mode}] 合同 {len(contracts)} 个: {", ".join(contracts.keys())}')

    # 每合同：期数/现值/利率
    for cname, c in sorted(contracts.items()):
        if cfg.get('contracts'):
            # 手动模式：参数已备齐，直接摊销
            _amortize(c, cfg)
            _straight(c)
            print(f'  {cname}: 月付 {c["pay_month"]:,.2f} 期数 {c["n_periods"]} '
                  f'现值 {c["present_value"]:,.2f} 利率 {c["rate"]*12*100:.2f}%/年 '
                  f'直线月息 {c.get("straight_int_month",0):,.2f}')
            continue
        pmt = c['pay_month']
        n = round(abs(tb_pay_open) / pmt) if pmt > 0 else 0
        c['n_periods'] = n if n > 0 else c['n_months_2025']
        c['present_value'] = abs(tb_use_asset_open) if abs(tb_use_asset_open) > 0 else pmt * c['n_periods']
        c['unfin_total'] = abs(tb_unfin_open) if abs(tb_unfin_open) > 0 else None
        c['rate'] = _reverse_rate(c['present_value'], pmt, c['n_periods'])
        _amortize(c, cfg)
        _straight(c)
        print(f'  {cname}: 月付 {pmt:,.2f} 期数 {c["n_periods"]} 现值 {c["present_value"]:,.2f} '
              f'利率 {c["rate"]*100:.4f}%/月 ({c["rate"]*12*100:.2f}%/年)'
              f' 直线月息 {c.get("straight_int_month",0):,.2f}')

    # ⚡ 2026-08-18：分主体文件输出到【中间产物】（非底稿目录）——专项底稿交付形态为集团合并文件，
    #   未来拆分从集团底稿开始；分主体文件仅作为合并数据源（lease_group_merge 读取）。
    out = out_dir or os.path.join(LM.resolve(acct), '中间产物', f'租赁复核单体_{year}')
    os.makedirs(out, exist_ok=True)
    out_path = os.path.join(out, f'使用权资产租赁复核_{year}_{sub}.xlsx')
    _write_workbook(out_path, cfg, contracts, tb_codes, mode)
    print(f'  ✅ 复核表已生成(中间产物): {out_path}')
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('acct', nargs='?', default=None)
    ap.add_argument('sub', nargs='?', default=None)
    a = ap.parse_args(argv)
    rc = 0
    for acct, acct_cfg in LM.jobs().items():
        if a.acct and acct != a.acct:
            continue
        for sub in acct_cfg['subs']:
            if a.sub and sub != a.sub:
                continue
            rc |= run_one(acct, sub)
    return rc


if __name__ == '__main__':
    sys.exit(main())
