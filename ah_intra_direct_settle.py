# -*- coding: utf-8 -*-
"""生成新底稿：关联往来未通过往来直接以银行存款结算交易明细
精确规则：
- 行科目=1002(银行存款) 且 对方公司名tc非空(集团内)
- 排除：摘要含'资金池'；对侧含往来科目(1122/1123/2202/2203/1124/1221/2241)；委贷(1301)；应收票据(1121)；资金资本类
- 保留并分类：费用类(对侧6600/6603等) / 收入类(对侧6001/6051/6402/6401)
- 按对子归集输出（Sheet1 归集对子 / Sheet2 未归集集团内主体）
- 1002行原始日期缺失 → 同凭证非空日期回填
"""
import sys
sys.path.insert(0, r'd:/底稿测试/账套取数审计小程序')
import os
import openpyxl
from collections import defaultdict, Counter
from openpyxl.styles import Font, PatternFill, Border, Side
from ah_intra_seq_extract import load_company_map
import ah_intra_diff_trace as m

DEFAULT_SEQ = r'd:/底稿测试/AH/中间产物/prepared/关联往来序时账_2026.xlsx'
DEFAULT_OUTDIR = r'd:/底稿测试/AH/中间产物/prepared/_过程稿/重建_20260821'

RECV_PAY = ('1122', '1124', '2203', '2202', '1123', '1221', '2241')
EXCLUDE_KM = ('1301', '1121')  # 委贷、应收票据
INCOME_KM = ('6001', '6051', '6402', '6401')   # 主营业务收入/其他业务收入
EXPENSE_KM = ('6600', '6601', '6602', '6603', '6604', '2211')  # 费用
CAPITAL_KM = ('2501', '2231', '2232', '2201', '1511', '4001', '6111', '1012',
              '1604', '1704', '4104')  # 借款/利息/资本类


def main(seq_fp=None, outdir=None):
    FP = seq_fp or DEFAULT_SEQ
    OUT = os.path.join(outdir or DEFAULT_OUTDIR, '关联往来未过往来直接结算交易明细_2026.xlsx')

    cmap = load_company_map()
    pairs = m._load_pairs()
    name2codes = defaultdict(set)
    for c, n in cmap.items():
        name2codes[n].add(c)

    def pair_of(comp_name, tc):
        """(本主体名, 对方名) → 对子no。名称精确/包含匹配。"""
        for p in pairs:
            x, y = p['X'], p['Y']
            if comp_name == x and tc == y: return p['no']
            if comp_name == y and tc == x: return p['no']
            if (comp_name in x or x in comp_name) and (tc in y or y in tc):
                return p['no']
            if (comp_name in y or y in comp_name) and (tc in x or x in tc):
                return p['no']
        return None

    wb = openpyxl.load_workbook(FP, read_only=True, data_only=True)
    ws = wb['关联往来序时账']
    rows = []
    for r in ws.iter_rows(values_only=True):
        if r[0] is None or r[0] == '账套':
            continue
        rows.append(r)
    wb.close()

    by_vou = defaultdict(list)
    for r in rows:
        by_vou[(str(r[0]), str(r[3]))].append(r)

    def classify(peers):
        """对侧科目分类。返回 (类别, 对侧科目串)。"""
        pkm = [str(x[4])[:4] for x in peers]
        if any(k in RECV_PAY for k in pkm): return '过往来', None
        if any(k in EXCLUDE_KM for k in pkm): return '排除', None
        if any(k in INCOME_KM for k in pkm): return '收入类', '+'.join(sorted(set(pkm)))
        if any(k in EXPENSE_KM for k in pkm): return '费用类', '+'.join(sorted(set(pkm)))
        if any(k in CAPITAL_KM for k in pkm): return '排除', None
        return '其他', '+'.join(sorted(set(pkm))[:3])

    direct_rows = []
    for r in rows:
        if str(r[4])[:4] != '1002':
            continue
        if not r[10]:
            continue
        if '资金池' in str(r[5] or ''):
            continue
        key = (str(r[0]), str(r[3]))
        peers = [x for x in by_vou.get(key, []) if str(x[4])[:4] != '1002']
        if not peers:
            continue
        cat, pkms = classify(peers)
        if cat in ('过往来', '排除'):
            continue
        comp_name = cmap.get(str(r[0]), str(r[0]))
        tc = str(r[10]).strip()
        pno = pair_of(comp_name, tc)
        ddate = r[2]
        if not ddate:
            for pr in by_vou.get(key, []):
                if pr[2]:
                    ddate = pr[2]
                    break
        direct_rows.append(dict(pno=pno, sz=str(r[0]), date=str(ddate), vou=str(r[3]),
                                km=str(r[4]), txt=str(r[5] or ''), dr=r[8], cr=r[9],
                                comp=comp_name, tc=tc, cat=cat, pkms=pkms or ''))

    print(f'直接结算分录: {len(direct_rows)}')
    print(f'分类: {dict(Counter(d["cat"] for d in direct_rows))}')
    print(f'归集到对子: {sum(1 for d in direct_rows if d["pno"] is not None)}/{len(direct_rows)}')

    # 输出Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '归集对子明细'
    F = Font(size=10); FH = Font(size=10, bold=True)
    FILL = PatternFill('solid', fgColor='DDEBF7')
    FILL2 = PatternFill('solid', fgColor='FFF2CC')
    THIN = Border(*[Side(style='thin')] * 4)
    H = ['对子', '本主体', '对方公司', '账套', '日期', '凭证', '科目', '方向', '金额',
         '对侧科目', '类别', '摘要']

    def write_sheet(ws_, rows_, fill):
        ws_.append(H)
        for c_ in ws_[1]:
            c_.font = FH; c_.fill = fill; c_.border = THIN
        for d in rows_:
            amt = d['dr'] or d['cr'] or 0
            direction = '借' if d['dr'] else '贷'
            ws_.append([d['pno'], d['comp'][:16], d['tc'][:16], d['sz'], d['date'], d['vou'],
                        d['km'], direction, amt, d['pkms'], d['cat'], d['txt'][:50]])
            for c_ in ws_[ws_.max_row]:
                c_.font = F; c_.border = THIN
                if isinstance(c_.value, (int, float)) and c_.column == 9:
                    c_.number_format = '#,##0.00'
        for col, w in zip('ABCDEFGHIJKL', [5, 18, 18, 7, 11, 12, 12, 5, 14, 12, 6, 52]):
            ws_.column_dimensions[col].width = w
        ws_.freeze_panes = 'A2'
        ws_.auto_filter.ref = f'A1:L{ws_.max_row}'

    matched = [d for d in direct_rows if d['pno'] is not None]
    unmatched = [d for d in direct_rows if d['pno'] is None]
    write_sheet(ws, matched, FILL)
    if unmatched:
        ws2 = wb.create_sheet('未归集集团内主体')
        write_sheet(ws2, unmatched, FILL2)
    wb.save(OUT)
    print(f'已归集对子: {len(matched)} 笔 | 未归集(集团内非对子主体): {len(unmatched)} 笔')
    print(f'已输出: {OUT}')
    return len(matched), len(unmatched)


if __name__ == '__main__':
    main()
