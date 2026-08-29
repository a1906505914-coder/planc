# -*- coding: utf-8 -*-
"""截止性测试底稿（审计底稿第5层·截止性测试）。
目的：验证年结日前后的交易是否记录在正确会计期间（防跨期/截止错报）。
数据来源：各主体《综合查询明细表》(GL, 逐笔凭证)。
方法：
  · 取两个窗口的凭证 —— 基准日前 2025-12-25~2025-12-31（年结周）、基准日后 2026-01-01~2026-01-10（次年初）；
  · 按整笔凭证分组，剔除 结转损益 / 相符计提(非截止敏感)，按科目名称归为 收入 / 采购·存货 / 其他重大；
  · 收入截止：含「主营业务收入」或「应收账款(借)」的凭证（赊销确认）；审计人员据出货单日期核验是否跨期；
  · 采购·存货截止：含 存货类科目(借) / 「应付账款」(贷) / 「主营业务成本」(借) 的凭证；据入库单日期核验；
  · 其他重大：窗口内金额≥100万且非上述的交易（如大额银行收付款），供关注可能的跨期现金收支；
  · 说明与勾稽：列示各主体窗口收入/采购合计、边界凭证连续性、及数据口径限制(2026 GL 仅1-5月)。
输出：截止性测试底稿_生成.xlsx
  · 收入截止测试
  · 采购·存货截止测试
  · 其他重大交易(截止关注)
  · 截止性测试说明与勾稽
口径：审计期间 FY2025，年结日=2025-12-31；2026仅含1-5月，故基准日后窗口(1-10日)可取。
"""
import paths as P
import os, re, sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import audit_common as A

DATA_DIR = P.G
OUT = os.path.join(P.G, '截止性测试底稿_生成.xlsx')

PRE_START, PRE_END = '2025-12-25', '2025-12-31'
POST_START, POST_END = '2026-01-01', '2026-01-10'
POST_RED_END = '2026-01-31'   # ⚡ 2026-08-15 期后红冲检测窗口（次年初 1 月全月，监管映射欠账①）
THRESH = 1_000_000.0

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D9E1F2')
WARNFILL = PatternFill('solid', fgColor='FFF2CC')
BADFILL = PatternFill('solid', fgColor='FCE4D6')
OKFILL = PatternFill('solid', fgColor='E2EFDA')
HFONT = Font(bold=True, name='Times New Roman', size=10)
BFONT = Font(bold=True, name='Times New Roman', size=10)
MONEY = '#,##0.00'

INV_KW = ('原材料', '库存商品', '在途物资', '周转材料', '半成品', '发出商品',
          '委托加工物资', '存货跌价', '材料成本差异', '商品进销差价')


def in_window(date_str, kind):
    d = str(date_str or '')
    if kind == 'pre':
        return PRE_START <= d <= PRE_END
    return POST_START <= d <= POST_END


def _in_red_window(d):
    """期后红冲检测窗口：次年初 1 月全月（含 POST 窗口，为其超集）。"""
    return POST_START <= str(d or '') <= POST_RED_END


def collect_post_reversal(vouchers):
    """⚡ 2026-08-15 期后红冲检测（监管映射欠账①：期末后红冲=疑似跨期）。

    基准日后红字冲减收入/应收的凭证：
      ① 主营业务收入/其他业务收入/营业收入 贷方红字(cr<0) 或 借方蓝字(dr>0) → 冲收入；
      ② 应收账款 借方红字(dr<0) → 红冲应收（多对应销售退回/折让）。
    返回 [(voucher, amount_abs, subject_name)]。"""
    out = []
    for v in vouchers:
        if not _in_red_window(v['date']):
            continue
        # ⚡ 2026-08-15 与主流程一致：跳过结转损益/相符计提（非截止敏感，避免把
        #   年终结转损益误判为"期后红冲"——AYL 实证 1-31 结转损益被误捕）
        skip, _reason = A.voucher_should_skip(summary=v.get('summary', ''), lines=v.get('lines'))
        if skip:
            continue
        amt, subj = 0.0, ''
        for l in v.get('lines', []):
            nm = str(l.get('name') or '')
            cr = float(l.get('cr') or 0)
            dr = float(l.get('dr') or 0)
            hit = False
            if any(k in nm for k in ('主营业务收入', '其他业务收入', '营业收入')):
                if cr < 0:
                    amt += abs(cr); hit = True
                elif dr > 0:
                    amt += dr; hit = True
            elif _is_ar(nm) and dr < 0:
                amt += abs(dr); hit = True
            if hit and not subj:
                subj = nm
        if amt > 0:
            out.append((v, amt, subj))
    return out


def _is_ar(n):
    """精确匹配应收账款科目本身（排除其备抵 坏账准备-应收账款…）。"""
    return n.startswith('应收账款') and '坏账' not in n


def _is_bank_cash(n):
    """识别银行/现金类科目（户名多为'工行科苑支行'等，不总是'银行存款'）。"""
    return any(k in n for k in ('银行', '支行', '分行', '存款', '库存现金', '现金',
                                '支付宝', '微信', '财务公司', '农信', '信用社'))


def classify_voucher(v):
    """按科目名称+借贷方向归为 '收入' / '采购' / '其他'（方向感知，避免收款误判收入）。"""
    lines = v.get('lines', [])
    has_rev_cr = ar_dr = ar_cr = bank_dr = bank_cr = inv_dr = ap_cr = cogs_dr = False
    for l in lines:
        nm = str(l.get('name') or '')
        d = float(l.get('dr') or 0)
        c = float(l.get('cr') or 0)
        if '主营业务收入' in nm and c > 0:
            has_rev_cr = True
        if _is_ar(nm) and d > 0:
            ar_dr = True
        if _is_ar(nm) and c > 0:
            ar_cr = True
        if _is_bank_cash(nm):
            if d > 0:
                bank_dr = True
            if c > 0:
                bank_cr = True
        if any(k in nm for k in INV_KW) and d > 0:
            inv_dr = True
        if '应付账款' in nm and c > 0:
            ap_cr = True
        if '主营业务成本' in nm and d > 0:
            cogs_dr = True
    # 收入确认：贷主营业务收入(赊销/现销)；或借应收账款且非收款(无银行借方)
    if has_rev_cr:
        return '收入'
    if ar_dr and not bank_dr:
        return '收入'
    # 采购·存货：存货(借)/应付(贷)/成本(借)
    if inv_dr or ap_cr or cogs_dr:
        return '采购'
    return '其他'


def main_accounts(v):
    """取凭证涉及的主要科目(去重, 前6个)。"""
    seen = []
    for l in v.get('lines', []):
        n = str(l.get('name') or '').strip()
        if n and n not in seen:
            seen.append(n)
        if len(seen) >= 6:
            break
    return ' / '.join(seen)


def period_of(date_str):
    return '基准日前(2025)' if str(date_str or '') <= PRE_END else '基准日后(2026)'


def build(data_dir, out_path):
    entities = A.discover_entities(data_dir)
    all_rows = A.read_gl_rows(data_dir, entities)
    print(f'· GL 总行数 {len(all_rows)}；主体 {len(entities)} 个')

    # 仅保留窗口内行（两窗口 + 期后红冲窗口 1 月全月）
    win_rows = [r for r in all_rows if in_window(r['date'], 'pre') or _in_red_window(r['date'])]
    vouchers = A.group_gl_vouchers(win_rows)
    print(f'· 窗口内凭证(分组后) {len(vouchers)} 笔')

    # 分类
    sales, purch, other = [], [], []
    for v in vouchers:
        skip, reason = A.voucher_should_skip(summary=v.get('summary', ''), lines=v.get('lines'))
        if skip:
            continue
        dr = sum(float(l.get('dr') or 0) for l in v['lines'])
        cr = sum(float(l.get('cr') or 0) for l in v['lines'])
        amt = max(dr, cr)
        kind = classify_voucher(v)
        rec = {
            'e': v['e'], 'y': v['y'], 'date': v['date'], 'vno': f"{v['vtype']}-{v['vno']}",
            'summary': v.get('summary', ''), 'dr': dr, 'cr': cr, 'amt': amt,
            'accts': main_accounts(v), 'period': period_of(v['date']),
        }
        if kind == '收入':
            sales.append(rec)
        elif kind == '采购':
            purch.append(rec)
        else:
            if amt >= THRESH:
                other.append(rec)

    # 排序：基准日前按日期倒序(贴近年结在前)，基准日后按日期正序
    def skey2(r):
        if r['period'].startswith('基准日前'):
            return (r['e'], 0, tuple(-int(x) for x in r['date'].split('-')))
        return (r['e'], 1, tuple(int(x) for x in r['date'].split('-')))
    sales.sort(key=skey2)
    purch.sort(key=skey2)
    other.sort(key=lambda r: (r['e'], r['date'], r['vno']))

    print(f'· 收入类 {len(sales)} 笔；采购·存货类 {len(purch)} 笔；其他重大(≥100万) {len(other)} 笔')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def style_header(ws, hdr, row=3):
        for c, h in enumerate(hdr, 1):
            cell = ws.cell(row, c, h)
            cell.font = HFONT; cell.fill = HFILL; cell.border = BORDER
            cell.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')

    def write_sheet(ws, title, sub, hdr, recs, amt_cols):
        ws.cell(1, 1, title).font = Font(bold=True, name='Times New Roman', size=10)
        ws.cell(2, 1, sub).font = Font(italic=True, name='Times New Roman', size=10, color='808080')
        style_header(ws, hdr, row=3)
        r = 4
        for rec in recs:
            row = [rec['e'], rec['period'], rec['date'], rec['vno'], rec['summary'],
                   rec['dr'], rec['cr'], rec['accts'], '']
            for c, val in enumerate(row, 1):
                cell = ws.cell(r, c, val); cell.border = BORDER
                if c in amt_cols:
                    cell.number_format = MONEY
                if c == 2 and rec['period'].startswith('基准日后'):
                    cell.fill = WARNFILL
            r += 1
        # 合计
        g_dr = sum(x['dr'] for x in recs); g_cr = sum(x['cr'] for x in recs)
        ws.cell(r, 1, '合计').font = BFONT
        ws.cell(r, 3, f'{len(recs)} 笔').font = BFONT
        for cc in (1, 2, 4, 5, 8, 9):
            ws.cell(r, cc).border = BORDER
        c = ws.cell(r, 6, g_dr); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
        c = ws.cell(r, 7, g_cr); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
        for i, w in enumerate([14, 14, 12, 22, 30, 18, 18, 40, 22], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.freeze_panes = 'A4'

    # ---------- Sheet1 收入截止 ----------
    ws1 = wb.create_sheet('收入截止测试')
    write_sheet(ws1, '收入截止性测试（赊销确认，据出货单核验是否跨期）',
                f'窗口：基准日前 {PRE_START}~{PRE_END}、基准日后 {POST_START}~{POST_END}；'
                '「审计说明」填列出货单日期/是否跨期/回函或替代测试结论。「基准日后」行底色标记需重点核对。',
                ['核算主体', '期间', '日期', '凭证号', '摘要', '借方(应收)', '贷方(收入)', '主要科目', '审计说明'],
                sales, (6, 7))

    # ---------- Sheet2 采购·存货截止 ----------
    ws2 = wb.create_sheet('采购·存货截止测试')
    write_sheet(ws2, '采购·存货截止性测试（入库/应付，据入库单核验是否跨期）',
                f'窗口：基准日前 {PRE_START}~{PRE_END}、基准日后 {POST_START}~{POST_END}；'
                '含 存货类科目(借)/应付账款(贷)/主营业务成本(借)。「审计说明」填列入库单日期/是否跨期。',
                ['核算主体', '期间', '日期', '凭证号', '摘要', '借方(存货/成本)', '贷方(应付/银存)', '主要科目', '审计说明'],
                purch, (6, 7))

    # ---------- Sheet3 其他重大交易 ----------
    ws3 = wb.create_sheet('其他重大交易(截止关注)')
    write_sheet(ws3, '其他重大交易（窗口内≥100万，关注可能的跨期现金收支）',
                f'窗口：基准日前 {PRE_START}~{PRE_END}、基准日后 {POST_START}~{POST_END}；'
                '已剔除结转损益与相符计提。大额银行收付款可能跨期，需核对其实际收付/单据日期。',
                ['核算主体', '期间', '日期', '凭证号', '摘要', '借方合计', '贷方合计', '主要科目', '审计说明'],
                other, (6, 7))

    # ---------- Sheet5 期后红冲·疑似跨期（2026-08-15 监管映射欠账①：收入截止/期后红冲） ----------
    pre_revs = [(r['e'], r['date'], r['vno'], r['amt']) for r in sales if r['period'].startswith('基准日前')]
    post_revs = collect_post_reversal(vouchers)
    rows5 = []
    for v, amt, subj in post_revs:
        best = None
        for pe, pd, pv, pa in pre_revs:
            if pe != v['e']:
                continue
            dd = abs(pa - amt)
            if best is None or dd < best[0]:
                best = (dd, pd, pv, pa)
        suspect = '否'
        pre_vno = pre_date = pre_amt = None
        if best is not None and best[0] <= max(1.0, 0.01 * amt):
            suspect = '是'
            _, pre_date, pre_vno, pre_amt = best
        rows5.append({
            'e': v['e'], 'date': v['date'], 'vno': f"{v['vtype']}-{v['vno']}",
            'summary': v.get('summary', ''), 'amt': amt, 'subj': subj,
            'pre_vno': pre_vno or '—', 'pre_date': pre_date or '—', 'pre_amt': pre_amt or 0.0,
            'flag': suspect,
        })
    ws5 = wb.create_sheet('期后红冲·疑似跨期')
    ws5.cell(1, 1, '期后红冲检测（次年初红字冲减收入/应收 → 与期末前收入配对，疑似跨期确认）').font = Font(bold=True, name='Times New Roman', size=10)
    ws5.cell(2, 1, f'窗口：{POST_START}~{POST_RED_END}（1 月全月）。红字冲减金额 ≈ 期末前收入 → 疑似「期后冲回=前收入跨期确认」；否则为独立冲销/折让/红字更正。').font = Font(italic=True, name='Times New Roman', size=10, color='808080')
    style_header(ws5, ['核算主体', '日期', '凭证号', '摘要', '红字冲减金额', '冲减科目',
                       '对应期末前收入(凭证号)', '收入日期', '收入金额', '疑似跨期', '审计说明'], row=3)
    r = 4
    for rec in rows5:
        row = [rec['e'], rec['date'], rec['vno'], rec['summary'], rec['amt'], rec['subj'],
               rec['pre_vno'], rec['pre_date'], rec['pre_amt'], rec['flag'], '']
        for c, val in enumerate(row, 1):
            cell = ws5.cell(r, c, val); cell.border = BORDER
            if c in (5, 9):
                cell.number_format = MONEY
        ws5.cell(r, 10).fill = BADFILL if rec['flag'] == '是' else WARNFILL
        r += 1
    ws5.cell(r, 1, '合计').font = BFONT
    ws5.cell(r, 3, f'{len(rows5)} 笔').font = BFONT
    for cc in range(1, 12):
        ws5.cell(r, cc).border = BORDER
    c = ws5.cell(r, 5, sum(x['amt'] for x in rows5)); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
    for i, w in enumerate([12, 12, 22, 30, 16, 22, 22, 12, 16, 10, 22], 1):
        ws5.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws5.freeze_panes = 'A4'
    n_suspect = sum(1 for x in rows5 if x['flag'] == '是')
    print(f'· 期后红冲 {len(rows5)} 笔，其中疑似跨期 {n_suspect} 笔')

    # ---------- Sheet4 说明与勾稽 ----------
    ws4 = wb.create_sheet('截止性测试说明与勾稽')
    notes = [
        '截止性测试底稿 编制与勾稽说明',
        '',
        '一、测试目标与窗口（第5层·截止性测试）',
        f'  核验年结日(2025-12-31)前后交易是否记入正确期间，防跨期/截止错报。',
        f'  窗口① 基准日前：{PRE_START} ~ {PRE_END}（年结周，交易应属 FY2025）；',
        f'  窗口② 基准日后：{POST_START} ~ {POST_END}（次年初，交易应属 FY2026）。',
        '  已剔除 结转损益 与 相符计提/摊销（非截止敏感）。',
        '',
        '二、各主体窗口发生额汇总（收入/采购）',
        '  收入=主营业务收入贷方+应收账款借方类凭证；采购=存货(借)+应付账款(贷)+主营业务成本(借)类凭证。',
    ]
    # 各主体汇总
    ents_order = sorted({r['e'] for r in (sales + purch)})
    notes.append(f"  {'主体':<12}{'收入(前)':>16}{'收入(后)':>16}{'采购(前)':>16}{'采购(后)':>16}")
    for e in ents_order:
        rev_pre = sum(r['amt'] for r in sales if r['e'] == e and r['period'].startswith('基准日前'))
        rev_post = sum(r['amt'] for r in sales if r['e'] == e and r['period'].startswith('基准日后'))
        pur_pre = sum(r['amt'] for r in purch if r['e'] == e and r['period'].startswith('基准日前'))
        pur_post = sum(r['amt'] for r in purch if r['e'] == e and r['period'].startswith('基准日后'))
        notes.append(f"  {e:<12}{rev_pre:>16,.0f}{rev_post:>16,.0f}{pur_pre:>16,.0f}{pur_post:>16,.0f}")
    # 边界连续性：检查各主体 基准日前末日 与 基准日后首日 是否均有凭证
    notes += [
        '',
        '三、边界凭证连续性（抽凭范围完整性自检）',
    ]
    for e in ents_order:
        pre_v = [r for r in (sales + purch + other) if r['e'] == e and r['period'].startswith('基准日前')]
        post_v = [r for r in (sales + purch + other) if r['e'] == e and r['period'].startswith('基准日后')]
        pre_max = max((r['date'] for r in pre_v), default='—')
        post_min = min((r['date'] for r in post_v), default='—')
        notes.append(f"  {e:<12} 基准日前最晚凭证日={pre_max}  基准日后最早凭证日={post_min}")
    notes += [
        '',
        '四、数据口径限制',
        '  · 2026 GL 仅含 1-5 月(YTD)，基准日后窗口(1-10日)可取；2025 GL 为全年，基准日前窗口可取。',
        '  · 个别主体(如爱绅科技)仅2026有辅助核算/GL，2025窗口无数据属取数范围限制，非账务差错。',
        f'  · 本次共提取窗口凭证：收入 {len(sales)} 笔、采购·存货 {len(purch)} 笔、其他重大 {len(other)} 笔。',
        '',
        '五、审计建议',
        '  · 收入：抽取基准日前最后若干笔+基准日后最初若干笔赊销凭证，核出货单日期；若出货单在次年且属FOB目的地/验收后风险转移，应调至 FY2026。',
        '  · 采购：核入库单日期；年结日前已收货未入账(暂估不足)或年结日后收货提前入账，均构成截止错报。',
        '  · 现金收支：核对支票/电汇实际收付日与记账日，关注跨年开具未兑付支票。',
        '',
        '六、期后红冲检测（2026-08-15 监管映射欠账①：收入跨期/期后红冲）',
        f'  扫描次年初({POST_START}~{POST_RED_END})红字冲减收入/应收的凭证，与期末前窗口收入配对：',
        '  金额近似(差≤1%或1元) → 标记「疑似跨期」（期后冲回=期末前确认的收入跨期，需核销货/退货单据日期）；',
        '  无对应期末前收入 → 独立冲销/折让/红字更正，正常。',
    ]
    for i, t in enumerate(notes, 1):
        ws4.cell(i, 1, t)
        if i == 1:
            ws4.cell(i, 1).font = Font(name='Times New Roman', bold=True, size=10)
    ws4.column_dimensions['A'].width = 110

    path, warn = A._safe_save(wb, out_path)
    if warn:
        print('  ⚠️', warn)
    print(f'  ✓ 已保存：{path}  (收入 {len(sales)} / 采购 {len(purch)} / 其他 {len(other)})')
    return path


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else DATA_DIR
    out = sys.argv[2] if len(sys.argv) > 2 else OUT
    build(data_dir, out)


if __name__ == '__main__':
    main()
