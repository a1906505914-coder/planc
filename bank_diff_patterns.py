# -*- coding: utf-8 -*-
"""银行核对差异【模式体检器】（2026-08-14 用户方法论升华：不能"碰到一次才能掌握"，
要"融会贯通"——把双向核对实战积累的差异模式固化为自动诊断，跑完核对主动识别每家
银行命中的已知模式，并给出建议，而不是等用户逐个指出"XX 也是"）。

模式库（P1-P12，全部来自实战实证，2026-08-13/14 双向核对过程）：

P1  对称双计     收差≈支差≠0 → 账侧借+贷同额双计（手续费/退押金/受托支付/资金池/
                购汇），网银记净额 → 净额化处理（_fee_net_journal/_entrust/_pool_wage）
                [3100/3300 农商 682 元、3300 建行 5.2 万购汇、农行 9.5 亿受托]
P2  口径不一致   月度对账 diff ≠ 红区净额 → 已解释项（对销/核销/归因/待报解）未同步
                [中行 3.9 亿内转]
P3  解析 bug     日期非法/金额天文/方向全反 → 表头/列映射/方向判定错误
                [浙商 2.78 亿方向、进出口 655 万列映射、华夏 1492 日期、3200 工行 6.79e21]
P4  文件缺失     账有网银无成批 → 网银只提供部分账户/月份流水 [上海/平安/外汇户]
P5  误抽非科目   出现"其他"银行 → 科目筛选关键词太宽 [DQ 177 笔应收票据/存款利息]
P6  拆组         同一银行拆两组 → 网银全称 vs 账侧简称 [民生银行/民生、光大银行/光大]
P7  红字未配     网银负金额未配 → 退款冲正账侧未记或形态不同 [一码通退款 -45]
P8  跨期错位     某月+X 相邻月-X → 跨期未达/月末滞后记账 [DQ 中行 500 万、6/7 月 2.585 亿]
P9  镜像成对     ub 同额收+支 → 资金池/跨境划转账未记 [盛虹/国望资金池、3300 建行 45 元]
P10 待报解通道   网银待报解 vs 账税费 → 代扣代缴两形态 [农商 709 万、3300 农商 5434 万]
P11 汇总记账     网银逐笔 vs 账月末汇总/区间 → 月汇总匹配/找和 [一码通/社保]
P12 浮点/尾差    微小差异 → 浮点比较/手续费尾差 [5 月 170.82 双重扣、年费 22.50]

体检器输入：banks/journal（原始）、results（核对后 ub/uj）、monthly（月度对账）。
输出：{bank: [(pid, 名称, 证据, 建议)]}——引擎/知识点在核对完成后自动调用并打印，
用户一眼看到"哪些银行命中哪些模式"，无需逐个指出。
"""
from collections import defaultdict


def _norm_amt(x):
    """红字归一化后取原始金额。"""
    if isinstance(x, dict):
        return float(x.get('_orig_amt', x.get('amount', 0)))
    return 0.0


def diagnose(banks, journal, results, monthly):
    """对每家银行做差异模式体检。返回 {bank: [(pid, name, evidence, advice)]}。"""
    report = {}
    # ⚡⚡ 2026-08-14 P6 拆组检测（数据层信号）：网银全称 vs 账侧简称被识别成两家
    #   银行 → 各自红区大量未配（DQ 民生银行/民生、光大银行/光大 实证）。
    #   判定：banks 集合内存在【包含/被包含】名称对（如 民生银行 ⊃ 民生）→ 拆组。
    _bank_set = sorted(set(banks) | set(journal))
    _pairs = []
    for _a in _bank_set:
        for _b2 in _bank_set:
            if _a == _b2:
                continue
            # 全称↔简称（含关系）：如 '民生银行' 含 '民生'。排除巧合包含
            # （'中行' 在 '民生银行' 中? 否）——需长名含短名且长度差>=2
            _long, _short = (_a, _b2) if len(_a) >= len(_b2) else (_b2, _a)
            if _short in _long and len(_long) - len(_short) >= 2:
                _pairs.append((_short, _long))
    _pairs = sorted(set(_pairs))
    for b in _bank_set:
        if b == '特殊科目':
            continue
        hits = []
        # ── P5 误抽非科目：银行名过宽（其他/特殊科目/存款利息等）→ 科目筛选漏配 ──
        _wide = ('其他', '特殊', '利息', '票据', '备用金', '理财')
        if any(w in b for w in _wide):
            hits.append(('P5', '误抽非科目',
                         f'银行名 "{b}" 含宽泛词（其他/特殊/利息/票据…）',
                         '科目筛选关键词太宽，抽入了非银行科目（如应收票据/存款利息），'
                         '应收紧筛选到真实银行主体'))
        # ── P6 拆组：本银行是某对全称/简称之一，且对方也出现 ──
        _mate = [p for p in _pairs if b in p]
        if _mate:
            _other = [x for x in _mate[0] if x != b]
            _ov = _other[0] if _other else ''
            _also = _ov in set(banks) and _ov in set(journal)
            hits.append(('P6', '银行拆组',
                         f'"{b}" 与 "{_ov}" 是全称/简称关系'
                         + ('（两侧均出现 → 同银行被拆成两组）' if _also
                            else '（仅一侧出现 → 名称可能不规范）'),
                         '网银全称 vs 账侧简称（民生银行/民生），应归一化同名后合并核对'))
        # ── P4 网银文件缺失：账侧有大量发生额但网银侧完全无该银行 ──
        #   （上海/平安/外汇户 实证：网银只提供部分账户/月份流水）
        #   ⚡ 排除 P6 拆组配对：若该银行是另一银行的全称/简称变体，属拆组非缺失
        if b not in set(banks) and not _mate:
            _jrows = journal.get(b, []) or []
            _j_amt = sum(abs(float(x.get('amount') or 0)) for x in _jrows)
            if _j_amt >= 1_000_000:  # 账侧发生额 ≥100 万
                hits.append(('P4', '网银文件缺失',
                             f'账侧"{b}" 发生额 {_j_amt:,.0f}（{len(_jrows)} 笔）但网银无该银行',
                             '网银只提供部分账户/月份流水，需补网银文件或确认该账户不在提供范围'))
        t = monthly.get(b, {}).get('TOTAL', {})
        diff_inc = float(t.get('diff_inc', 0))
        diff_exp = float(t.get('diff_exp', 0))
        r = results.get(b, {})
        ub = r.get('ub', [])
        uj = r.get('uj', [])
        # 红区净额（口径一致性的"红区"侧）
        ub_inc = sum(x['amount'] for x in ub if x['direction'] == '收')
        ub_exp = sum(x['amount'] for x in ub if x['direction'] == '支')
        uj_dr = sum(x['amount'] for x in uj if x['direction'] == '借')
        uj_cr = sum(x['amount'] for x in uj if x['direction'] == '贷')
        red_inc = ub_inc - uj_dr
        red_exp = ub_exp - uj_cr

        # ── P1 对称双计：收差≈支差 且 非零 ──
        tol = max(100.0, abs(diff_inc) * 0.0001)
        if abs(diff_inc - diff_exp) < tol and (abs(diff_inc) > 0.01 or abs(diff_exp) > 0.01):
            hits.append(('P1', '对称双计',
                         f'收差{diff_inc:,.2f}≈支差{diff_exp:,.2f}（净额平衡）',
                         '账侧借+贷同额双计（手续费/退押金/受托/资金池/购汇），'
                         '检查净额化关键词是否覆盖该银行收款/冲减名'))

        # ── P2 口径不一致：月度 diff ≠ 红区净额 ──
        # ⚡ 排除红字解释：网银负金额行（退款冲正）归一化后计入月度、还原显示在红区
        #   （设计行为，P7 已报）——若红区-月度的差全部由红字负额解释，则非口径 bug
        neg_inc = sum(_norm_amt(x) for x in ub
                      if x['direction'] == '收' and _norm_amt(x) < 0)
        adj_red_inc = red_inc - neg_inc  # 红区收净额剔除红字贡献
        if (abs(diff_inc - adj_red_inc) > 0.02 or abs(diff_exp - red_exp) > 0.02):
            hits.append(('P2', '两表口径不一致',
                         f'月度{diff_inc:,.2f}/{diff_exp:,.2f} vs 红区{red_inc:,.2f}/{red_exp:,.2f}'
                         f'（红字-{abs(neg_inc):,.2f}已解释）',
                         '已解释项（对销/核销/归因/待报解）未同步月度对账，查 expl 剔除'))

        # ── P7 红字未配：网银负金额行 ──
        neg = [x for x in ub if _norm_amt(x) < 0]
        if neg:
            amt = sum(-_norm_amt(x) for x in neg)
            hits.append(('P7', '红字未配',
                         f'网银负金额 {len(neg)} 笔 -{amt:,.2f}',
                         '退款冲正账侧未记或形态不同（收-X vs 贷X），负数归一化后应配对'))

        # ── P9 镜像成对：ub 内同额收+支（集团/资金池） ──
        amt_key = defaultdict(list)
        for x in ub:
            amt_key[round(abs(_norm_amt(x)), 2)].append(x)
        mirror_n = 0
        mirror_amt = 0.0
        for amt, xs in amt_key.items():
            recv = [x for x in xs if x['direction'] == '收']
            pay = [x for x in xs if x['direction'] == '支']
            if recv and pay:
                mirror_n += min(len(recv), len(pay))
                mirror_amt += amt * min(len(recv), len(pay))
        if mirror_n:
            hits.append(('P9', '镜像成对',
                         f'ub 内同额收+支 {mirror_n} 对 {mirror_amt:,.2f}',
                         '资金池/跨境划转账未记（净额0），已对平的银行保留红区列示，'
                         '未对平的检查是否应成对移除'))

        # ── P3 解析异常：非法日期/天文金额 ──
        bad_date = [x for x in ub + uj if not _dday_ok(x.get('date'))]
        astro = [x for x in ub + uj if abs(_norm_amt(x)) >= 1e12]
        if bad_date or astro:
            hits.append(('P3', '解析异常',
                         f'非法日期 {len(bad_date)} 笔 / 天文金额 {len(astro)} 笔',
                         '网银 parser 表头/列映射/方向判定 bug，需修正解析'))

        # ── P8 跨期错位：逐月收差符号相邻相反（成对抵消） ──
        d = monthly.get(b, {})
        months = sorted(m for m in d if m != 'TOTAL')
        swing = []
        for i in range(1, len(months)):
            p, c = months[i - 1], months[i]
            dp, dc = d[p].get('diff_inc', 0), d[c].get('diff_inc', 0)
            if dp * dc < 0 and abs(dp) > 1000 and abs(dc) > 1000:
                swing.append(f'{p}({dp:,.0f})↔{c}({dc:,.0f})')
        if swing:
            hits.append(('P8', '跨期错位',
                         '；'.join(swing[:4]),
                         '跨期未达/月末滞后记账（45 天窗口/摘要业务日期匹配）'))

        # ── P10 待报解通道：网银待报解 vs 账税费 —— 由 daikou 处理，命中则说明已覆盖，
        #    此处仅提示若红区仍有待报解残留 ──
        dai_b = [x for x in ub if any(k in (str(x.get('counterparty') or '')
                                            + str(x.get('summary') or ''))
                                      for k in ('待报解', '代理国库', '收缴'))]
        if dai_b:
            hits.append(('P10', '待报解残留',
                         f'网银待报解 {len(dai_b)} 笔 {sum(_norm_amt(x) for x in dai_b):,.2f}',
                         '代扣通道（待报解↔税费）语义匹配/月核销未完全覆盖'))

        # ── P11 汇总记账：网银逐笔 vs 账月末汇总/区间（一码通/社保 实证）──
        #   特征：红区同月内同摘要（一码通/社保/结息）多笔 → 账侧可能按月汇总入账
        _sum_by = defaultdict(list)
        for x in ub + uj:
            _m = str(x.get('date') or '')[:7]
            _t = str(x.get('summary') or '') + str(x.get('counterparty') or '')
            _sum_by[(_m, _t)].append(x)
        _sum_hits = [(k, v) for k, v in _sum_by.items()
                     if k[0] and len(v) >= 5 and sum(abs(_norm_amt(x)) for x in v) >= 1000]
        if _sum_hits:
            _s0 = _sum_hits[0]
            hits.append(('P11', '汇总记账',
                         f'红区同月同摘要 {_s0[1][0].get("date", "")[:7]}'
                         f'"{_s0[0][1][:16]}" {len(_s0[1])} 笔'
                         f'{sum(abs(_norm_amt(x)) for x in _s0[1]):,.2f}'
                         + (f'（另有 {len(_sum_hits)-1} 组同类）' if len(_sum_hits) > 1 else ''),
                         '网银逐笔 vs 账月末汇总/区间（一码通/社保）→ 月汇总匹配/找和算法'))
        if hits:
            report[b] = hits
    return report


def _dday_ok(date_str):
    """日期合法性：非法（如 6550-00-00/1492-08-64/None）返回 False。"""
    import datetime
    s = str(date_str or '')
    if not s:
        return False
    try:
        datetime.date.fromisoformat(s[:10])
        return True
    except (ValueError, TypeError):
        return False


def print_report(report, title='差异模式体检报告'):
    """打印模式体检报告（引擎/知识点调用）。"""
    print()
    print(f'===== {title} =====')
    if not report:
        print('  ✅ 全部银行无已知差异模式命中（差异均已定性或对平）')
        return
    for b in sorted(report):
        print(f'【{b}】')
        for pid, name, ev, adv in report[b]:
            print(f'  [{pid} {name}] {ev}')
            print(f'      → {adv}')
