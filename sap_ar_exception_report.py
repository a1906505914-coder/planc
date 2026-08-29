# -*- coding: utf-8 -*-
"""sap_ar_exception_report.py —— 应收账款异常关注清单生成器（2026-08-08）。

审计重点关注（用户定义）：
  ① 大额发票冲销（冲销发票相关凭证——核查红冲依据/收入期间调节）
  ② 客户-供应商大额互抵（应收应付互抵/抵账——核查抵账协议/关联方/附注披露）
  ③ 大额客户/供应商分类调整（重分类——核查账龄与余额归类）
  ④ 其他大额 AR 调整（调整/划转/清账，金额≥阈值）
资金池上下划拨（122104 资金结算）属正常业务，不标记。

用法：python sap_ar_exception_report.py [<data_dir>] [--min 500000]
输出：{data_dir}/底稿/应收账款异常关注清单_2026.xlsx（异常汇总 + 凭证明细）
"""
import paths as P
import os, re, sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import _safx

DATA = sys.argv[1] if len(sys.argv) > 1 else P.YY
MIN_AMT = 500000.0
if '--min' in sys.argv:
    MIN_AMT = float(sys.argv[sys.argv.index('--min') + 1])

GL = os.path.join(DATA, '序时账')
# v3 修正：1123=预付账款（非应收），AR=1122 应收+1221 其他应收（排除 122104 资金结算）
AR_PFX = ('1122', '1221')
AR_EXCL = ('1221040000',)
AP_PFX = ('2202',)          # 应付账款
REV_PFX = ('6001', '6051')  # 收入

# 异常关键词
KW_INV = ('发票', '红冲', '红字')
KW_OFFSET = ('互抵', '抵账', '抵应付', '应收应付')
KW_CLASS = ('分类', '重分类')
KW_ADJ = ('调整', '划转', '冲销', '冲回', '清账')

FONT = Font(name='Times New Roman', size=10)
BOLD = Font(name='Times New Roman', size=10, bold=True)
HEAD_FONT = Font(name='Times New Roman', size=10, bold=True)
HEAD_FILL = PatternFill('solid', fgColor='DDEBF7')
TOT_FILL = PatternFill('solid', fgColor='FCE4D6')
NUMFMT = '#,##0.00'
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
RGT = Alignment(horizontal='right', vertical='center')
LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
_thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

# 审计关注点模板
AUDIT_POINT = {
    '发票冲销': '核查红冲/冲销依据、发票真伪与作废流程；关注是否跨期调节收入、应收余额冲销是否完整。',
    '应收应付互抵': '核查抵账协议/审批记录；关注互抵双方是否为关联方、金额是否真实、报表是否需附注披露净额列示。',
    '分类调整': '核查供应商/客户分类重分类依据；关注账龄分析是否被重分类影响、余额归属是否正确。',
    '其他调整': '核查调整事由、原始单据；关注是否虚增虚减应收、期末大额调整的合理性。',
}


def is_ar(s):
    return s.startswith(AR_PFX) and not s.startswith(AR_EXCL)


def gl_files(comp):
    out = []
    for fn in sorted(os.listdir(GL)):
        if not fn.lower().endswith('.xlsx'):
            continue
        base = os.path.splitext(fn)[0]
        m = re.match(r'^(\d{4})', base)
        if not m:
            continue
        m2 = re.match(r'^(\d{4})[~-](\d{4})$', base)
        if m.group(1) == comp:
            out.append(os.path.join(GL, fn))
        elif m2 and m2.group(1) <= comp <= m2.group(2):
            out.append(os.path.join(GL, fn))
    return out


def classify(vno, d, supp_act=None):
    """按用户规则分类异常：返回 (类型, AR净额, 凭证AR金额, 应付侧互抵金额, 业务性质) 或 None。
    用户定义：①大额发票冲销 ②客户-供应商大额互抵（应收应付互抵）。
    业务性质（2026-08-08 用户细化）：
      - 水电费类（连续日常业务）→ 正常
      - 暂估互抵（应付侧双向）→ 正常
      - 其他互抵：该供应商本年有其他持续发生 → 正常-持续往来；
                  本年无其他发生（仅此一笔冲抵）或非水电大额 → ⚠️ 标注核查
    排除：资金池上下划拨(122104)、系统月末评估冲销(1122999998/上评估/重置)。"""
    txt = str(d.get('_txt', ''))
    ar_net = sum(a for s, a in d.items() if s != '_txt' and is_ar(s))
    ar_abs = sum(abs(a) for s, a in d.items() if s != '_txt' and is_ar(s))
    if abs(ar_net) < MIN_AMT and ar_abs < MIN_AMT:
        return None
    # 排除系统月末评估/重置（1100 系列自动结转）
    if '1122999998' in d or '上评估' in txt or '重置' in txt:
        return None
    ap_amt = sum(abs(a) for s, a in d.items() if s != '_txt' and s.startswith(AP_PFX))
    has_ap = ap_amt > 0
    big = ar_abs >= MIN_AMT or abs(ar_net) >= MIN_AMT

    # 业务性质：互抵类（用户方法论：应付侧有对应=双向；连续日常=正常，无持续业务=标注）
    def nature_of():
        if has_ap:
            if any(k in txt for k in ('水电', '电费', '水费', '水电气')):
                return '正常-水电费结算(连续日常)'
            if '暂估' in txt:
                return '正常-暂估互抵(应付侧双向确认)'
            supp = str(d.get('_supp', '')).strip()
            act = (supp_act or {}).get(supp, 0) if supp else 0
            if act >= 3:
                return f'正常-持续往来互抵(供应商{supp}本年{act}笔)'
            if supp:
                return f'⚠️ 大额冲抵-{supp}本年无持续业务({act}笔),需核查'
            return '⚠️ 大额冲抵-无供应商信息,需核查'
        return '待核查(应付侧无对应)'

    nature = nature_of()
    # ① 发票冲销（红字发票/发票冲销/冲销…发票）
    if '发票' in txt and any(k in txt for k in ('冲销', '冲回', '红字', '红冲', '红')):
        if big:
            return ('发票冲销', ar_net, ar_abs, ap_amt, nature)
    if '红字发票' in txt and big:
        return ('发票冲销', ar_net, ar_abs, ap_amt, nature)
    # ② 应收应付互抵（客户-供应商互抵：AR 与 2202 应付同凭证大额）
    if has_ap and big:
        return ('应收应付互抵', ar_net, ar_abs, ap_amt, nature)
    if any(k in txt for k in KW_OFFSET) and big:
        return ('应收应付互抵', ar_net, ar_abs, ap_amt, nature)
    # ③ 其他大额调整（调整/串户/冲销/划转/结转，供参考）
    if any(k in txt for k in KW_ADJ) and big:
        return ('其他调整', ar_net, ar_abs, ap_amt, nature)
    return None


def analyze_one(comp):
    """单家公司异常识别 → rows 列表。"""
    v = {}
    supp_vnos = defaultdict(set)   # 供应商 → 出现过的凭证号集合（年度活跃度）
    for fp in gl_files(comp):
        try:
            for row in _safx.iter_rows(fp, keep_cols={1, 5, 6, 10, 11, 12, 21, 29}):
                if str(row.get(6, '')).strip() != comp:
                    continue
                per = str(row.get(10, '')).strip()
                if not (per[:2].isdigit() and 1 <= int(per[:2]) <= 7):
                    continue
                subj = str(row.get(11) or '')
                try:
                    amt = float(row.get(29) or 0)
                except (TypeError, ValueError):
                    continue
                vno = str(row.get(1) or '').strip()
                if not vno or not subj or amt == 0:
                    continue
                d = v.setdefault(vno, defaultdict(float))
                d[subj] += amt
                if '_txt' not in d:
                    d['_txt'] = str(row.get(5) or '').strip()
                if '_cust' not in d and row.get(12):
                    d['_cust'] = str(row.get(12)).strip()
                if '_supp' not in d and row.get(21):
                    d['_supp'] = str(row.get(21)).strip()
                if row.get(21):
                    supp_vnos[str(row.get(21)).strip()].add(vno)
        except Exception as e:
            print(f'  ⚠️ {os.path.basename(fp)}: {e}', flush=True)
    supp_act = {k: len(vn) for k, vn in supp_vnos.items()}
    rows = []
    for vno, d in v.items():
        res = classify(vno, d, supp_act)
        if res:
            typ, ar_net, ar_abs, ap_amt, nature = res
            lines = {s: a for s, a in d.items() if s not in ('_txt', '_cust', '_supp')}
            rows.append((comp, vno, typ, ar_net, ar_abs, ap_amt, nature,
                         str(d.get('_txt', ''))[:60],
                         str(d.get('_cust', ''))[:24], str(d.get('_supp', ''))[:24],
                         lines))
    del v
    return rows


def write_excel(rows):
    """rows → Excel 清单。"""
    out_dir = os.path.join(DATA, '底稿')
    os.makedirs(out_dir, exist_ok=True)
    fp = os.path.join(out_dir, '应收账款异常关注清单_2026.xlsx')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Sheet1 异常汇总
    ws = wb.create_sheet('异常汇总')
    hdrs = ['公司代码', '凭证编号', '异常类型', 'AR净额影响', 'AR发生额合计',
            '应付侧互抵金额', '业务性质', '凭证文本', '客户', '供应商', '审计关注点']
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(1, j, h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = CTR
        c.border = BORDER
    r = 2
    for comp, vno, typ, ar_net, ar_abs, ap_amt, nature, txt, cust, supp, lines in rows:
        ws.cell(r, 1, comp).border = BORDER
        ws.cell(r, 2, vno).border = BORDER
        c = ws.cell(r, 3, typ)
        c.border = BORDER
        c.fill = PatternFill('solid', fgColor='FFF2CC' if typ == '发票冲销' else 'FCE4D6')
        for j, v in [(4, ar_net), (5, ar_abs)]:
            cc = ws.cell(r, j, round(v, 2))
            cc.number_format = NUMFMT
            cc.border = BORDER
            cc.alignment = RGT
        cc = ws.cell(r, 6, round(ap_amt, 2) if ap_amt else None)
        cc.number_format = NUMFMT
        cc.border = BORDER
        cc.alignment = RGT
        c7 = ws.cell(r, 7, nature)
        c7.border = BORDER
        c7.alignment = LFT
        if nature.startswith('正常'):
            c7.fill = PatternFill('solid', fgColor='E2EFDA')
        else:
            c7.fill = PatternFill('solid', fgColor='FFC7CE')
        ws.cell(r, 8, txt).border = BORDER
        ws.cell(r, 9, cust).border = BORDER
        ws.cell(r, 10, supp).border = BORDER
        cc = ws.cell(r, 11, AUDIT_POINT.get(typ, ''))
        cc.alignment = LFT
        cc.border = BORDER
        r += 1
    from collections import Counter
    cnt = Counter(x[2] for x in rows)
    nat_cnt = Counter(x[6] for x in rows)
    r += 1
    ws.cell(r, 1, '类型统计').font = BOLD
    for i, (k, n) in enumerate(cnt.items()):
        ws.cell(r + 1 + i, 1, f'{k}: {n} 笔')
    r2 = r + 1 + len(cnt) + 1
    ws.cell(r2, 1, '业务性质统计').font = BOLD
    for i, (k, n) in enumerate(nat_cnt.items()):
        ws.cell(r2 + 1 + i, 1, f'{k}: {n} 笔')
    ws.freeze_panes = 'A2'
    for j, w in enumerate([10, 14, 12, 15, 15, 16, 30, 46, 14, 14, 46], 1):
        ws.column_dimensions[chr(64 + j)].width = w

    # Sheet2 异常凭证明细
    ws2 = wb.create_sheet('异常凭证明细')
    hdrs2 = ['公司代码', '凭证编号', '异常类型', '业务性质', '科目', '借贷方向', '金额', '文本']
    for j, h in enumerate(hdrs2, 1):
        c = ws2.cell(1, j, h)
        c.font = HEAD_FONT
        c.fill = HEAD_FILL
        c.alignment = CTR
        c.border = BORDER
    r2 = 2
    for comp, vno, typ, ar_net, ar_abs, ap_amt, nature, txt, cust, supp, lines in rows:
        for code, amt in sorted(lines.items(), key=lambda x: -abs(x[1])):
            if abs(amt) < 0.005:
                continue
            ws2.cell(r2, 1, comp).border = BORDER
            ws2.cell(r2, 2, vno).border = BORDER
            ws2.cell(r2, 3, typ).border = BORDER
            ws2.cell(r2, 4, nature).border = BORDER
            ws2.cell(r2, 5, code).border = BORDER
            ws2.cell(r2, 6, '借' if amt > 0 else '贷').border = BORDER
            ws2.cell(r2, 7, round(abs(amt), 2)).number_format = NUMFMT
            ws2.cell(r2, 7).border = BORDER
            ws2.cell(r2, 7).alignment = RGT
            ws2.cell(r2, 8, txt).border = BORDER
            r2 += 1
    ws2.freeze_panes = 'A2'
    for j, w in enumerate([10, 14, 12, 30, 16, 8, 14, 40], 1):
        ws2.column_dimensions[chr(64 + j)].width = w

    wb.save(fp)
    wb.close()
    return fp


def main(argv=None):
    if '--one' in sys.argv:
        # 单家模式（子进程调用）
        comp = sys.argv[sys.argv.index('--one') + 1]
        rows = analyze_one(comp)
        import json
        print(json.dumps([(r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7],
                           r[8], r[9], dict(r[10])) for r in rows], ensure_ascii=False))
        return
    # 全量模式：逐家子进程隔离（防内存累积 OOM）
    import json as _json
    import subprocess
    # ⚡ 2026-08-09 修复：公司集合用 TB 全量（文件名前 4 位会漏区间文件内公司 2610~5040 等）
    import sap_reader as _SR
    comps = sorted(_SR.discover_sap_entities(DATA).keys())
    print(f'扫描 {len(comps)} 家公司，异常阈值 {MIN_AMT:,.0f} ...', flush=True)
    rows = []
    for comp in comps:
        r = subprocess.run([sys.executable, os.path.abspath(__file__),
                            DATA, '--one', comp],
                           capture_output=True, text=True, encoding='utf-8',
                           errors='replace')
        if r.returncode != 0 or not r.stdout.strip():
            print(f'  ⚠️ {comp} 子进程失败: {r.stderr.strip()[:80]}', flush=True)
            continue
        try:
            arr = _json.loads(r.stdout)
        except Exception as e:
            print(f'  ⚠️ {comp} 解析失败: {e}', flush=True)
            continue
        rows += arr
        print(f'  {comp} 完成（{len(arr)} 笔）', flush=True)
    fp = write_excel(rows)
    from collections import Counter
    cnt = Counter(x[2] for x in rows)
    print(f'✅ 生成: {fp}')
    print(f'   异常凭证 {len(rows)} 笔：' + ' | '.join(f'{k}={n}' for k, n in cnt.items()))
# main 结束（write_excel 已生成文件并打印统计）


if __name__ == '__main__':
    main()
