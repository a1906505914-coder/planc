# -*- coding: utf-8 -*-
"""CO 结转凭证识别统计（审阅/回归用）：各主体 CO 凭证数/金额 + 排除后 GL vs TB。
用法：python co_carryover_stats.py [--out xlsx]
"""
import sys, io, os
sys.path.insert(0, r'd:/底稿测试/账套取数审计小程序')
os.chdir(r'd:/底稿测试/账套取数审计小程序')
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import openpyxl
from collections import defaultdict
import sap_common as SC
from audit_common import discover_entities

D = r'd:/底稿测试/ADF/数据/2026'
COMPS = ['3300', '3500', '3700', '3900', '6100', '6900', '3400']


def read_tb(comp):
    ents = discover_entities(D)
    fpt = ents.get(comp, {}).get('2026', {}).get('km')
    if not fpt:
        return {}
    wb = openpyxl.load_workbook(fpt, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    hdr = None
    jf_i = None
    tb = {}
    for row in ws.iter_rows(values_only=True):
        if hdr is None:
            hdr = [str(x or '').strip() for x in row]
            try:
                jf_i = hdr.index('本期借方金额')
            except ValueError:
                return {}
            continue
        if not row or not str(row[1] or '').strip().isdigit():
            continue
        try:
            tb[str(row[1]).strip()] = float(row[jf_i] or 0)
        except (TypeError, ValueError, IndexError):
            pass
    wb.close()
    return tb


def main():
    rows_out = []
    print('主体 | GL行 | is_co行 | CO绝对额(亿) | GL借(亿) | 排除CO借(亿) | TB借(亿) | 排除后差(亿)')
    for comp in COMPS:
        rows = SC.load_comp_gl_disk(D, comp)
        co_n = sum(1 for r in rows if r.get('is_co'))
        co_amt = sum(abs(float(r['amt'])) for r in rows if r.get('is_co'))
        jf = sum(float(r['amt']) for r in rows if float(r['amt']) > 0)
        jf_ex = sum(float(r['amt']) for r in rows if float(r['amt']) > 0 and not r.get('is_co'))
        tb = read_tb(comp)
        tb_jf = sum(tb.values())
        print('%s | %d | %d | %.2f | %.2f | %.2f | %.2f | %.2f' % (
            comp, len(rows), co_n, co_amt/1e8, jf/1e8, jf_ex/1e8,
            tb_jf/1e8, (jf_ex - tb_jf)/1e8), flush=True)
        rows_out.append((comp, len(rows), co_n, co_amt/1e8, jf/1e8,
                         jf_ex/1e8, tb_jf/1e8, (jf_ex - tb_jf)/1e8))

    # 可选输出 xlsx
    out_xlsx = os.path.join(D, 'CO结转凭证统计.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'CO结转统计'
    ws.append(['主体', 'GL行数', 'CO结转凭证行', 'CO绝对额(亿)', 'GL借(亿)',
               '排除CO后借(亿)', 'TB借(亿)', '排除CO后差(亿)'])
    for r in rows_out:
        ws.append(list(r))
    # 说明
    ws.append([])
    ws.append(['说明', 'CO结转=多行(≥8)净额0 且 含704科目/文本含结转关键词。',
               '排除CO后 GL借 vs TB借 的差为剩余口径差（1-6月/红字）。'])
    wb.save(out_xlsx)
    print()
    print('已输出:', out_xlsx)


if __name__ == '__main__':
    main()
