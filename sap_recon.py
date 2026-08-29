# -*- coding: utf-8 -*-
"""SAP 企业报表 vs 自建试算表核对器（sap_recon.py，2026-08-07 v2）。

用途：把企业报表（资产负债表/利润表 .xls）与 sap_reader 读出的 TB 逐行核对，
固化「报表行 → TB 科目」映射规则，输出差异清单，作为 SAP 底稿生成核对基准。

v2 修复（3010 实测）：
  1) 报表行名匹配：关键词包含匹配（'实收资本' ↔ '实收资本（或股本）'、'一、营业收入' ↔ '营业收入'）；
  2) 科目匹配按【一级名】相等/包含（'无形资产' 不误并 '无形资产摊销' 费用科目、
     '累计折旧' 不误并 '使用权资产累计折旧'）；
  3) 备抵支持 (一级名, 子关键词) 二元组（坏账准备-应收帐款：一级名='坏账准备' 且 全名含 '应收'）；
  4) 未分配利润 = 利润分配/本年利润(期初) + 当年净利润；
     净利润 = 营业收入 - 营业成本 - 税附 - 费用表(销/管/研) - 财务净额 - 信用减值
             - 资产减值 + 其他收益 + 投资收益 + 资产处置 + 营业外收入 - 营业外支出 - 所得税；
  5) 合同负债：优先独立科目，无则回退预收账款；
  6) 费用表 {公司}.xlsx 4 sheet 合计 = 销售/管理/研发费用（权威拆分）。
"""
import paths as P
import os
import re
import sys
import time

import sap_reader as SR

# ============================================================
# 一、报表行 → TB 科目映射
# ============================================================
# 资产侧：每项 = (报表行关键词, [(一级名kw | (一级名kw, 子kw, 排除词))], 备抵 or None)
# 备抵 = (一级名kw, 子kw or None, 排除词列表 or None)；净值=原值+备抵（贷余负直接相加）
BS_ASSET_MAP = [
    ('货币资金', ['库存现金', '银行存款', '其他货币资金', '银行存款人民币', '银行存款—人民币'], None),
    ('应收票据', ['应收票据'], ('坏账准备', '应收票据', None)),
    ('应收账款', ['应收账款'], ('坏账准备', '应收帐款', ['其他应收', '合同资产', '应收票据'])),
    ('预付款项', ['预付账款', '预付帐款'], None),
    ('其他应收款', ['其他应收款'], ('坏账准备', '其他应收', None)),
    ('存货', ['原材料', '库存产品', '库存产品成本差异', '委托加工物资', '在产品', '半成品', '发出商品', '周转材料'],
     ('存货跌价准备', None, None)),
    ('合同资产', ['合同资产'], ('坏账准备', '合同资产', None)),
    ('一年内到期的非流动资产', [('委托贷款', '一年内到期', None)], None),
    ('长期股权投资', ['长期股权投资'], ('长期股权投资减值准备', None, None)),
    ('投资性房地产', ['投资性房地产'], ('投资性房地产累计折旧', None, None)),
    ('固定资产', ['固定资产'], ('累计折旧', None, ['使用权', '投资性'])),
    ('在建工程', ['在建工程'], ('在建工程减值准备', None, None)),
    ('固定资产清理', ['固定资产清理'], None),
    ('使用权资产', ['使用权资产'], ('使用权资产累计折旧', None, None)),
    ('无形资产', ['无形资产'], ('累计摊销', None, None)),
    ('开发支出', ['研发支出'], None),
    ('商誉', ['商誉'], ('商誉减值准备', None, None)),
    ('递延所得税资产', ['递延所得税资产'], None),
    ('其他非流动资产', [('委托贷款', '一年以上', None)], None),
]

# 负债权益侧：(报表行关键词, [TB 一级名 kw])
BS_LIAB_MAP = [
    ('短期借款', ['短期借款']),
    ('应付票据', ['应付票据']),
    ('应付账款', ['应付账款']),
    ('预收账款', ['预收账款', '预收帐款']),
    ('合同负债', ['合同负债']),           # 无独立科目时回退预收账款
    ('应付职工薪酬', ['应付职工薪酬']),
    ('应交税费', ['应交税费']),
    ('应付利息', ['应付利息']),
    ('应付股利', ['应付股利']),
    ('其他应付款', ['其他应付款']),
    ('一年内到期的非流动负债', ['长期借款']),  # 特殊：一年内到期部分
    ('长期借款', ['长期借款']),
    ('应付债券', ['应付债券']),
    ('租赁负债', ['租赁负债']),
    ('长期应付款', ['长期应付款']),
    ('专项应付款', ['专项应付款']),
    ('预计负债', ['预计负债']),
    ('递延收益', ['递延收益']),
    ('递延所得税负债', ['递延所得税负债']),
    ('实收资本', ['实收资本']),
    ('其他权益工具', ['其他权益工具']),
    ('资本公积', ['资本公积']),
    ('其他综合收益', ['其他综合收益']),
    ('盈余公积', ['盈余公积']),
    ('专项储备', ['专项储备']),
    ('未分配利润', ['利润分配', '本年利润']),  # 特殊
]

# 利润表：报表行关键词 → [(TB 一级名kw, 口径)]  口径: df=贷方 df_net=贷方净额 jf=借方
#   net=借-贷(财务费用带符号) loss=-(借-贷)(损失为负)
PL_MAP = [
    ('营业收入', [('主营业务收入', 'df_net'), ('其他业务收入', 'df_net')]),
    ('营业成本', [('主营业务成本', 'net'), ('其他业务成本', 'net')]),
    ('税金及附加', [('税金及附加', 'jf')]),
    ('销售费用', None),   # 费用表
    ('管理费用', None),   # 费用表
    ('研发费用', [('研究开发费', 'net')]),  # 交叉验证（net=借-贷，1010 验证）
    ('财务费用', [('财务费用', 'net')]),
    ('其他收益', [('其他收益', 'df_net')]),
    ('投资收益', [('投资收益', 'df_net')]),
    ('信用减值损失', [('信用减值损失', 'loss')]),
    ('资产减值损失', [('资产减值损失', 'loss')]),
    ('资产处置收益', [('资产处置损益', 'df_net')]),
    ('营业外收入', [('营业外收入', 'df')]),
    ('营业外支出', [('营业外支出', 'jf')]),
    ('所得税', [('所得税费用', 'jf')]),
]

EXPENSE_SHEETS = {'销售费用': '销售费用', '管理费用': '管理费用', '研发费用': '研发费用'}


def _norm(s):
    """全半角括号/冒号归一化。"""
    if s is None:
        return ''
    return str(s).replace('（', '(').replace('）', ')').replace('：', ':').replace(' ', '').strip()


def _l1(name):
    return name.split('-')[0].strip() if name else name


def _fval(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def read_expense_totals(data_dir, comp):
    """费用表两种格式 → {sheet: 合计}：
    A) 费用/{comp}.xlsx 单文件 4 sheet（销售/管理/制造/研发），每 sheet 末行『合计：』1-7月分列；
    B) 费用/{comp}/ 目录：{类别}{N}.xlsx 按月文件（Sheet1: 项目/本月数/本期累计数/上年同期累计数），
       取 7 月文件的本期累计合计（YTD）。
    """
    fp = os.path.join(data_dir, '费用', f'{comp}.xlsx')
    d = os.path.join(data_dir, '费用', comp)
    out = {}
    if os.path.exists(fp):
        import openpyxl
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        for sn in wb.sheetnames:
            ws = wb[sn]
            last = None
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or r[1] is None:
                    continue
                if '合计' in str(r[1]) and '合计数' not in str(r[1]):
                    last = sum(_fval(x) for x in r[2:9])
            if last:
                out[sn] = last
        wb.close()
    elif os.path.isdir(d):
        # 按月文件：取每个类别最后一个月（7月）文件的本期累计数
        for cat in ('销售费用', '管理费用', '研发费用', '制造费用'):
            f7 = os.path.join(d, f'{cat}7.xlsx')
            if not os.path.exists(f7):
                # 找该类别最大的月号文件
                fs = sorted(os.listdir(d), reverse=True)
                f7 = next((os.path.join(d, f) for f in fs
                           if f.startswith(cat) and f.endswith('.xlsx')), None)
            if not f7:
                continue
            import openpyxl
            wb = openpyxl.load_workbook(f7, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            last = None
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or r[0] is None:
                    continue
                if '合计' in str(r[0]) and '合计数' not in str(r[0]):
                    last = _fval(r[2])  # 本期累计数列
            wb.close()
            if last:
                out[cat] = last
    return out


def _agg(tb, comp):
    """聚合公司一级科目：{一级名: {qc,jf,df,qm}}，同时保留全名清单。"""
    l1 = {}
    for (c, _cd, name, _y), v in tb.items():
        if c != comp:
            continue
        n = _l1(name)
        a = l1.setdefault(n, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0})
        for k in ('qc', 'jf', 'df', 'qm'):
            a[k] += v[k]
    return l1


def _sub_match(l1, l1kw, subkw):
    """匹配一级名含 l1kw 的科目；若给定 subkw，须其原始全名含 subkw。
    返回带符号金额合计。"""
    tot = 0.0
    for n, a in l1.items():
        if l1kw in n:
            if subkw is not None:
                # 需要在 TB 全名中找到带 subkw 的科目，这里用一级名近似：
                # 一级名含子关键词（如 坏账准备-应收帐款 一级名=坏账准备，需全名）
                pass
            tot += a['qm']
    return tot


def _sum(tb, comp, l1, kws, mode='qm'):
    """按一级名匹配汇总。kws 元素可为：
    - str: 一级名精确匹配 l1 键（'无形资产' 不误并 '无形资产摊销'）
    - (l1kw, include, exclude): 一级名含 l1kw 且全名含 include（None=全部），排除含 exclude 词
    mode: qc/jf/df/qm（余额/发生额）；'net'=jf-df（借方净额）；'df_net'=df-jf（贷方净额）；'loss'=-(jf-df)
    """
    tot = 0.0

    def _add(v, m):
        if m == 'net':
            return v['jf'] - v['df']
        if m == 'df_net':
            return v['df'] - v['jf']
        if m == 'loss':
            return -(v['jf'] - v['df'])
        return v[m]

    for kw in kws:
        if isinstance(kw, tuple):
            l1kw, include, exclude = kw[0], kw[1], (kw[2] if len(kw) > 2 else None)
            for (c, _cd, name, _y), v in tb.items():
                if c != comp:
                    continue
                n = _l1(name)
                if l1kw in n:
                    if include is not None and include not in name:
                        continue
                    if exclude and any(e in name for e in exclude):
                        continue
                    tot += _add(v, mode)
        else:
            if kw in l1:
                tot += _add(l1[kw], mode)
            else:
                # 变体匹配：'实收资本' ↔ '实收资本(或股本)'（一级名以 kw+'(' 开头）
                for ln, a in l1.items():
                    if ln.startswith(kw + '('):
                        tot += _add(a, mode)
    return tot


def _pl_net(tb, comp, l1, exp):
    """当年净利润 = 营业收支净额 - 期间费用(费用表) - 损失 + 其他收益/投资收益/营业外。
    收入=贷方净额(df-jf)、成本/费用=借方净额(jf-df)、损失=-(借-贷)（报表损失负）。"""
    g = lambda kws, mode: _sum(tb, comp, l1, kws, mode)
    rev = g(['主营业务收入'], 'df_net') + g(['其他业务收入'], 'df_net')
    cost = g(['主营业务成本'], 'net') + g(['其他业务成本'], 'net')
    tax = g(['税金及附加'], 'jf')
    sales = exp.get('销售费用', 0.0)
    mgmt = exp.get('管理费用', 0.0)
    rd = exp.get('研发费用', 0.0)
    fin = g(['财务费用'], 'net')
    credit = g(['信用减值损失'], 'loss')
    asset = g(['资产减值损失'], 'loss')
    oth_inc = g(['其他收益'], 'df')
    inv = g(['投资收益'], 'df_net')
    disp = g(['资产处置损益'], 'df_net')
    non_in = g(['营业外收入'], 'df')
    non_out = g(['营业外支出'], 'jf')
    itax = g(['所得税费用'], 'jf')
    return (rev - cost - tax - sales - mgmt - rd - fin + credit + asset
            + oth_inc + inv + disp + non_in - non_out - itax)


def reconcile_bs(tb, comp, report_begin, report_end, exp=None):
    """report_begin/end = {归一化行名: 值}。exp=费用表合计（未分配利润需当年净利润）。"""
    l1 = _agg(tb, comp)

    def _rpt(d, rn):
        """报表值：精确 or 关键词包含匹配。"""
        if rn in d:
            return d[rn]
        for k, v in d.items():
            if rn in k:
                return v
        return None

    rows = []
    # 资产侧
    for rn, kws, bad in BS_ASSET_MAP:
        v_end = _sum(tb, comp, l1, kws)
        v_begin = _sum(tb, comp, l1, kws, 'qc')
        if bad:
            if isinstance(bad, tuple) and bad[0] == '委托贷款':
                pass  # 委托贷款拆分单独处理
            v_end += _sum(tb, comp, l1, [bad])
            v_begin += _sum(tb, comp, l1, [bad], 'qc')
        rep_b, rep_e = _rpt(report_begin, rn), _rpt(report_end, rn)
        diff = round(v_end - rep_e, 2) if isinstance(rep_e, (int, float)) and abs(rep_e) > 0.005 else None
        rows.append((rn, v_begin, rep_b, v_end, rep_e, diff))
    # 负债权益侧
    for rn, kws in BS_LIAB_MAP:
        if rn == '未分配利润':
            # ⚡ 2026-08-09 修复：去 abs()（原对借余/累计亏损公司符号错误，如利润分配+本年利润为借余时
            #    应显示负未分配利润）。带符号加总（贷余=正、借余=负），企业报表同口径。
            v_begin = _sum(tb, comp, l1, ['利润分配'], 'qc') + _sum(tb, comp, l1, ['本年利润'], 'qc')
            v_end = _sum(tb, comp, l1, ['利润分配']) + _sum(tb, comp, l1, ['本年利润']) \
                    + _pl_net(tb, comp, l1, exp or {})
        elif rn == '合同负债':
            v = _sum(tb, comp, l1, ['合同负债'])
            if abs(v) < 0.005:
                v = _sum(tb, comp, l1, ['预收账款', '预收帐款'])
            v_end, v_begin = abs(v), abs(_sum(tb, comp, l1, ['合同负债'], 'qc') or _sum(tb, comp, l1, ['预收账款', '预收帐款'], 'qc'))
        elif rn == '一年内到期的非流动负债':
            # 长期借款子科目含『一年内到期』
            v_end = 0.0
            for (c, _cd, name, _y), v in tb.items():
                if c == comp and '一年内到期' in name and ('长期借款' in name or '应付债券' in name):
                    v_end += v['qm']
            v_begin = 0.0
            for (c, _cd, name, _y), v in tb.items():
                if c == comp and '一年内到期' in name and ('长期借款' in name or '应付债券' in name):
                    v_begin += v['qc']
            v_end, v_begin = abs(v_end), abs(v_begin)
        elif rn == '长期借款':
            v = _sum(tb, comp, l1, ['长期借款'])
            # 扣除一年内到期部分
            for (c, _cd, name, _y), vv in tb.items():
                if c == comp and '一年内到期' in name and '长期借款' in name:
                    v -= vv['qm']
            v_begin = _sum(tb, comp, l1, ['长期借款'], 'qc')
            for (c, _cd, name, _y), vv in tb.items():
                if c == comp and '一年内到期' in name and '长期借款' in name:
                    v_begin -= vv['qc']
            v_end, v_begin = abs(v), abs(v_begin)
        else:
            v_end = abs(_sum(tb, comp, l1, kws))
            v_begin = abs(_sum(tb, comp, l1, kws, 'qc'))
        rep_b, rep_e = _rpt(report_begin, rn), _rpt(report_end, rn)
        # 负债侧：TB abs 与报表 abs 比较（借余 TB 正/报表负 是符号口径差异，金额一致=对平）
        diff = round(abs(v_end) - abs(rep_e), 2) if isinstance(rep_e, (int, float)) and abs(rep_e) > 0.005 else None
        rows.append((rn, v_begin, rep_b, v_end, rep_e, diff))
    return rows


def reconcile_pl(tb, comp, report, exp):
    """report = {归一化行名: {'cum': 值}}。"""
    l1 = _agg(tb, comp)

    def _rpt(rn):
        if rn in report:
            return report[rn]['cum']
        for k, v in report.items():
            if rn in k:
                return v['cum']
        return None

    rows = []
    for rn, spec in PL_MAP:
        if spec is None:
            v = exp.get(EXPENSE_SHEETS.get(rn, rn), 0.0)
        else:
            v = sum(_sum(tb, comp, l1, [kws], mode) for kws, mode in spec)
        rep = _rpt(rn)
        diff = round(v - rep, 2) if isinstance(rep, (int, float)) and abs(rep) > 0.005 else None
        rows.append((rn, v, rep, diff))
    return rows


def run(comp='3010', data_dir=None, out=None, tb=None):
    if data_dir is None:
        data_dir = P.YY
    if tb is None:  # ⚡ 2026-08-13 run_all 单次读 TB 传入，避免 88 次全量重读 OOM
        tb = SR.read_sap_tb(data_dir, None, year='2026')

    rpt_bs = SR.read_sap_report(data_dir, '资产负债表', comp)
    bs_begin = {_norm(k): v['begin'] for k, v in rpt_bs.items()}
    bs_end = {_norm(k): v['end'] for k, v in rpt_bs.items()}
    rpt_pl = SR.read_sap_report(data_dir, '利润表', comp)
    pl = {_norm(k): v for k, v in rpt_pl.items()}
    exp = read_expense_totals(data_dir, comp)
    log = (lambda s: out.write(s + '\n')) if out else print
    log(f'===== {comp} 资产负债表核对 =====')
    log('%-15s %16s %16s %16s %16s %12s %s' % (
        '行名', 'TB期末', '报表期末', 'TB年初', '报表年初', '差异', ''))
    n_diff = 0
    for rn, tb_b, rep_b, tb_e, rep_e, diff in reconcile_bs(tb, comp, bs_begin, bs_end, exp):
        flag = ''
        if diff is not None and abs(diff) > 0.5:
            flag = '<<< 差异'
            n_diff += 1
        log('%-15s %16.2f %16.2f %16.2f %16.2f %12.2f %s' % (
            rn[:15], tb_e or 0, rep_e or 0, tb_b or 0, rep_b or 0, diff or 0, flag))

    rpt_pl = SR.read_sap_report(data_dir, '利润表', comp)
    pl = {_norm(k): v for k, v in rpt_pl.items()}
    exp = read_expense_totals(data_dir, comp)
    log(f'\n===== {comp} 利润表核对 =====  费用表: { {k: round(v,2) for k,v in exp.items()} }')
    log('%-15s %16s %16s %12s %s' % ('行名', 'TB/费用表', '报表累计', '差异', ''))
    for rn, v, rep, diff in reconcile_pl(tb, comp, pl, exp):
        flag = ''
        if diff is not None and abs(diff) > 0.5:
            flag = '<<< 差异'
            n_diff += 1
        log('%-15s %16.2f %16.2f %12.2f %s' % (rn[:15], v or 0, rep or 0, diff or 0, flag))
    log(f'\n差异项数: {n_diff}')
    return n_diff


def run_all(data_dir=None, out_dir=None):
    """全量 88 家核对（#701，2026-08-13）：逐主体写差异清单，汇总报告。返回差异项总数。"""
    import os as _os
    if data_dir is None:
        data_dir = P.YY
    if out_dir is None:
        out_dir = data_dir
    comps = sorted(SR.discover_sap_entities(data_dir).keys())
    print(f'#701 全量核对：{len(comps)} 家主体 → {out_dir}', flush=True)
    tb = SR.read_sap_tb(data_dir, None, year='2026')  # 只读一次，循环共用
    rows = []
    n_zero = 0
    for c in comps:
        fp = _os.path.join(out_dir, f'SAP_{c}_试算表vs企业报表_差异清单.txt')
        try:
            with open(fp, 'w', encoding='utf-8') as f:
                f.write(f'SAP {c} 试算表 vs 企业报表核对差异清单（{time.strftime("%Y-%m-%d %H:%M")}）\n')
                f.write('=' * 80 + '\n')
                n = run(c, data_dir, out=f, tb=tb)
        except Exception as e:
            n = -1
            print(f'  {c}: ERR {e}')
        rows.append((c, n))
        if n == 0:
            n_zero += 1
        print(f'  {c}: 差异 {n} 项' + (' ✅' if n == 0 else ''), flush=True)
    # 汇总报告
    sum_fp = _os.path.join(out_dir, 'SAP_全量_试算表vs企业报表_汇总.txt')
    with open(sum_fp, 'w', encoding='utf-8') as f:
        f.write(f'SAP 试算表 vs 企业报表 全量核对汇总（{time.strftime("%Y-%m-%d %H:%M")}）\n')
        f.write('=' * 80 + '\n')
        f.write(f'主体 {len(rows)} 家；零差异 {n_zero} 家；有差异 {len(rows) - n_zero} 家\n\n')
        f.write('%-8s %8s\n' % ('主体', '差异项'))
        for c, n in rows:
            f.write('%-8s %8s\n' % (c, n if n >= 0 else 'ERR'))
        f.write(f'\n差异项合计：{sum(max(x[1], 0) for x in rows)}\n')
    print(f'汇总报告：{sum_fp}（零差异 {n_zero}/{len(rows)}）')
    return sum(max(n, 0) for n in rows)


if __name__ == '__main__':
    argv = sys.argv[1:]
    if argv and argv[0] == '--all':
        run_all(argv[1] if len(argv) > 1 else P.YY,
                argv[2] if len(argv) > 2 else None)
    else:
        comps = argv[0].split(',') if argv else ['3010']
        for c in comps:
            run(c)
