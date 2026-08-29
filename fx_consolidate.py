# -*- coding: utf-8 -*-
"""fx_consolidate.py —— 外币报表折算（2026-08-05 用户需求：Z 集团分币种底稿折算合并）。

流程（用户确认方案 A 的合并层）：
  分币种底稿（Z_split/{CNY,THB,USD}，已生成）→ 本程序对 THB/USD 组：
    ① 读取 TB 原币金额（read_tb_full，泰国经 thai_mapping 归一化）；
    ② 按科目类别 × 汇率折算为 CNY：
        · 资产负债表项目（1xx 资产 / 2xx 负债）：期初→期初汇率、期末→期末汇率；
        · 所有者权益（4xx，实收资本/资本公积等）：历史汇率（发生时点，用期初汇率近似）；
        · 成本/损益（5xx/6xx）：本期发生额→平均汇率；
        · 未分配利润：期初折算 + 本年损益折算（滚动），差异由折算差额吸收；
    ③ 折算差额 → 其他综合收益『外币报表折算差额』（保证折算后试算平衡）；
    ④ 输出：每主体折算试算表 + 折算附注汇总（全 CNY 口径）。

用法：
  python fx_consolidate.py <Z_split根目录> [输出目录]
  例：python fx_consolidate.py <DATA_ROOT>/Z_split
      → Z_split\\合并折算\\泰国_折算试算表_2025.xlsx、新加坡_折算试算表_2025.xlsx、
        折算附注汇总_2025.xlsx

汇率：fx_rates.py（期初 2024-12-31 / 期末 2025-12-31 / 平均 2025 年度，人行中间价+统计局）。
"""
import paths as P
import os
import sys
import io

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from currency_mapping import entity_currency, CURRENCY_LABEL  # noqa: E402
from fx_rates import get_rate, to_cny, needs_avg_rates  # noqa: E402

YEAR = '2025'


def _norm_tb(tb_full, entities):
    """泰国 TB 归一化（科目代码/名称 → 集团标准），幂等。
    2026-08-05：normalize_thai_tb 是『复制泰国行到标准码』，归一化后泰国主体同时存在
    原码行（1113 银行存款）与标准码行（1002 银行存款）→ 折算若全取会双计。
    处理：泰国主体行的一级码若在 THAI_CODE_MAP.keys()（原泰国码）→ 丢弃原行，仅保留
    标准码行（映射目标码 + 未映射的原泰国码科目，如 1156 政府收入未映射则保留）。"""
    try:
        from thai_mapping import normalize_thai_tb as _n
        from thai_mapping import THAI_CODE_MAP, is_thai_entity
        tb = _n(tb_full, entities)
        if tb is tb_full:
            return tb_full
    except Exception:
        return tb_full
    thai_keys = set(THAI_CODE_MAP.keys())
    out = {}
    for (e, c, n, y), v in tb.items():
        if is_thai_entity(e):
            first = str(c).split('.')[0].strip()
            if first in thai_keys:
                continue  # 原泰国码行，丢弃（标准码行已由 normalize 加入）
        out[(e, c, n, y)] = v
    return out


def classify(code):
    """按归一化标准代码首位分类：'bs'（资产负债+成本在产品，余额按期末/期初汇率）
    /'eq'（权益，历史汇率）/'pl'（损益，平均汇率）。
    2026-08-05：5xx 生产成本/制造费用期末余额=在产品（存货），属资产负债表项目 → bs；
    6xx 损益期末应无余额（已结转），按平均汇率 → pl。"""
    s = str(code).split('.')[0].strip()
    if not s:
        return 'bs'
    first = s[0]
    if first in '125':
        return 'bs'
    if first == '4':
        return 'eq'
    return 'pl'  # 6xx 损益


def convert_tb(tb_full, entities):
    """把外币 TB 折算为 CNY。返回 {entity: {code: {name, qc, jf, df, qm, cls}}} + 差额汇总。"""
    from collections import defaultdict
    out = {}
    diffs = {}
    for e in sorted(entities):
        cur = entity_currency(e)
        if cur == 'CNY':
            continue
        r_begin = get_rate(cur, 'period_begin')
        r_end = get_rate(cur, 'period_end')
        r_avg = get_rate(cur, 'avg')
        if r_begin is None or r_end is None or r_avg is None:
            diffs[e] = '缺汇率(%s)' % cur
            continue
        acc = defaultdict(lambda: {'name': '', 'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0, 'cls': 'bs'})
        # 2026-08-05：末级过滤——泰国/新加坡 TB 父级（1113）=子目之和，父+子双计会翻倍。
        # 只保留「无任何更长同前缀代码」的末级行（与审计底稿取数同策略）。
        codes_of_e = {c for (ee, c, n, y) in tb_full if ee == e}
        leaf_codes = {c for c in codes_of_e
                      if not any(c2.startswith(c) and len(c2) > len(c) for c2 in codes_of_e)}
        for (ee, c, n, y), v in tb_full.items():
            if ee != e or c not in leaf_codes:
                continue
            cls = classify(c)
            # 泰国归一化后代码可能带点，取首位分类
            a = acc[c]
            a['name'] = str(n) if str(n) else a['name']
            a['cls'] = cls
            if cls == 'bs':
                a['qc'] += v['qc'] * r_begin
                a['qm'] += v['qm'] * r_end
                a['jf'] += v['jf'] * r_avg
                a['df'] += v['df'] * r_avg
            elif cls == 'eq':
                # 权益：历史汇率近似用期初汇率（实收资本入账时点）；发生额按平均
                a['qc'] += v['qc'] * r_begin
                a['qm'] += v['qm'] * r_begin
                a['jf'] += v['jf'] * r_avg
                a['df'] += v['df'] * r_avg
            else:  # pl
                a['qc'] += v['qc'] * r_avg
                a['qm'] += v['qm'] * r_avg
                a['jf'] += v['jf'] * r_avg
                a['df'] += v['df'] * r_avg
        # 折算差额（准则19号）：期初净资产(期初汇率) + 本期损益(平均汇率) − 期末净资产(期末汇率)。
        # 期末净资产 = Σ资产qm − Σ负债qm（signed 贷为负，资产借方为正 → Σbs.qm 即净资产）；
        # 权益按历史汇率不参与差额（实收资本等）；未分配利润由损益滚动 + 差额吸收。
        # 平衡检验：Σ(折算后 qc) + Σ(折算后 jf − df) − Σ(折算后 qm) 应为 0；
        # 实际不为 0 的部分即 外币报表折算差额（计入其他综合收益）。
        net_begin = sum(a['qc'] for a in acc.values())
        net_flow = sum(a['jf'] - a['df'] for a in acc.values())
        net_end = sum(a['qm'] for a in acc.values())
        # 折算差额 = 期初净资产 + 本期净变化 − 期末净资产（signed，贷余为负）
        diff = net_begin + net_flow - net_end
        out[e] = dict(acc)
        diffs[e] = diff
    return out, diffs


def build_workbooks(out_data, diffs, out_dir, build_note=True):
    """生成每主体折算试算表；build_note=True 时另生成折算附注汇总（合并全部外币主体）。"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    os.makedirs(out_dir, exist_ok=True)
    F10 = Font(name='Times New Roman', size=10)
    BF = Font(name='Times New Roman', size=10, bold=True)
    TF = Font(name='Times New Roman', size=12, bold=True)
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    TOTFILL = PatternFill('solid', fgColor='FCE4D6')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
    NUM = '#,##0.00'

    note_paths = []
    for e, acc in out_data.items():
        cur = entity_currency(e)
        label = CURRENCY_LABEL.get(cur, cur)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '折算试算表'
        ws.cell(1, 1, '%s %s年折算试算表（%s → 人民币 CNY，资产负债表期末/期初、损益平均汇率）' % (e, YEAR, label)).font = TF
        hdr = ['科目代码', '科目名称', '类别', '期初(CNY)', '本期借方(CNY)', '本期贷方(CNY)', '期末(CNY)']
        for j, h in enumerate(hdr, 1):
            c = ws.cell(3, j, h); c.font = BF; c.fill = HFILL; c.border = BORDER; c.alignment = CTR
        r = 4
        for c in sorted(acc):
            a = acc[c]
            vals = [c, a['name'], a['cls'], a['qc'], a['jf'], a['df'], a['qm']]
            for j, v in enumerate(vals, 1):
                cell = ws.cell(r, j, v)
                cell.font = F10; cell.border = BORDER
                if j >= 4:
                    cell.number_format = NUM
            r += 1
        # 合计行
        tot = [0.0] * 4
        for a in acc.values():
            tot[0] += a['qc']; tot[1] += a['jf']; tot[2] += a['df']; tot[3] += a['qm']
        ws.cell(r, 1, '合计'); 
        for j, t in enumerate(tot, 4):
            c = ws.cell(r, j, round(t, 2)); c.font = BF; c.fill = TOTFILL; c.number_format = NUM; c.border = BORDER
        ws.cell(r, 1).font = BF; ws.cell(r, 1).fill = TOTFILL; ws.cell(r, 1).border = BORDER
        r += 1
        d = diffs.get(e, 0.0)
        ws.cell(r, 1, '外币报表折算差额（其他综合收益）').font = BF
        ws.cell(r, 2, '（折算后平衡差额，计入其他综合收益）').font = F10
        c = ws.cell(r, 7, round(-d, 2) if isinstance(d, (int, float)) else str(d))
        c.font = BF; c.number_format = NUM; c.border = BORDER
        for j, wd in enumerate([14, 34, 8, 16, 16, 16, 16], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = wd
        fp = os.path.join(out_dir, '%s_折算试算表_%s.xlsx' % (e, YEAR))
        wb.save(fp)
        note_paths.append(fp)
        print('  ✅ %s 折算试算表 → %s' % (e, fp))

    # 折算附注汇总（合并外币组，行=一级科目，列=主体×币种折算CNY）
    if build_note and out_data:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '折算附注汇总'
        ws.cell(1, 1, '外币主体折算附注汇总（%s 年，原币折算人民币 CNY）' % YEAR).font = TF
        ws.cell(3, 1, '科目代码').font = BF
        ws.cell(3, 2, '科目名称').font = BF
        col = 3
        for e in sorted(out_data):
            ws.cell(3, col, '%s（%s→CNY）' % (e, CURRENCY_LABEL.get(entity_currency(e), entity_currency(e)))).font = BF
            ws.cell(3, col + 1, '期末审定CNY').font = BF
            col += 2
        ws.cell(3, col, '合并CNY合计').font = BF
        # 汇总一级科目
        from collections import defaultdict
        agg = defaultdict(lambda: defaultdict(float))
        names = {}
        for e, acc in out_data.items():
            for c, a in acc.items():
                p = c.split('.')[0]
                agg[p][e] += a['qm']
                names[p] = a['name']
        r = 4
        for p in sorted(agg, key=lambda x: (x[0] if x[0].isdigit() else '9', x)):
            ws.cell(r, 1, p).font = F10; ws.cell(r, 1).border = BORDER
            ws.cell(r, 2, names[p][:30]).font = F10; ws.cell(r, 2).border = BORDER
            total = 0.0
            cc = 3
            for e in sorted(out_data):
                v = agg[p].get(e, 0.0)
                total += v
                c1 = ws.cell(r, cc, round(v, 2)); c1.number_format = NUM; c1.font = F10; c1.border = BORDER
                cc += 1
                c2 = ws.cell(r, cc, round(v, 2)); c2.number_format = NUM; c2.font = F10; c2.border = BORDER
                cc += 1
            c3 = ws.cell(r, col, round(total, 2)); c3.number_format = NUM; c3.font = BF; c3.fill = TOTFILL; c3.border = BORDER
            r += 1
        fp = os.path.join(out_dir, '折算附注汇总_%s.xlsx' % YEAR)
        wb.save(fp)
        note_paths.append(fp)
        print('  ✅ 折算附注汇总 → %s' % fp)
    return note_paths


def process(split_root, out_dir=None):
    out_dir = out_dir or os.path.join(split_root, '合并折算')
    missing = needs_avg_rates()
    if missing:
        print('⚠️ 平均汇率未定：%s（fx_rates.py 补充后重跑）' % missing)
        return 1
    import audit_common as ac
    total_paths = []
    all_out = {}
    all_diffs = {}
    for tag in ['THB', 'USD']:
        d = os.path.join(split_root, tag)
        if not os.path.isdir(d):
            continue
        print('===== %s 组 =====' % tag)
        ents = ac.discover_entities(d)
        tb = ac.read_tb_full(d, ents)
        tb = _norm_tb(tb, ents)
        out_data, diffs = convert_tb(tb, ents)
        if not out_data:
            try:
                from mask_dict import mask_names as _mn
                _ents = _mn(sorted(ents))
            except Exception:
                _ents = sorted(ents)
            print('  ⚠️ 无外币主体可折算（主体: %s）' % _ents)
            continue
        all_out.update(out_data)
        all_diffs.update(diffs)
        # 每主体折算试算表（独立文件）
        per_entity = {e: {c: a for c, a in acc.items()} for e, acc in out_data.items()}
        paths = build_workbooks(per_entity, diffs, out_dir, build_note=False)
        total_paths.extend(paths)
        for e, d2 in diffs.items():
            print('  [%s] 折算差额: %s' % (e, format(d2, ',.2f') if isinstance(d2, (int, float)) else d2))
    # 折算附注汇总（合并全部外币主体，一次生成）
    if all_out:
        paths = build_workbooks(all_out, all_diffs, out_dir, build_note=True)
        total_paths.extend(paths)
    print('\n✅ 折算完成：%s（%d 个文件）' % (out_dir, len(total_paths)))
    return 0


def main():
    argv = sys.argv[1:]
    split_root = argv[0] if argv else os.path.join(P.Z, 'Z_split')
    out_dir = argv[1] if len(argv) > 1 else os.path.join(split_root, '合并折算')
    if not os.path.isdir(split_root):
        print('⚠️ 目录不存在：%s' % split_root)
        return 1
    return process(split_root, out_dir)


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.exit(main())
