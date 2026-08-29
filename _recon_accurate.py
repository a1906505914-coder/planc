# -*- coding: utf-8 -*-
"""精确勾稽核对：审定表合计 vs 明细表合计 vs 附注合计。
- 审定表：取最后含'合计'行 的 期末审定数/期末未审数/期末数 列
- 明细表：取最后含'合计'（含 全集团合计/年度合计）行的期末列
- 附注：取合计行
用法：python _recon_accurate.py <底稿目录>
"""
import sys, os, glob
import openpyxl

def find_end_col(ws, n_hdr=6):
    prio = ['期末审定数', '期末余额(人民币)', '期末余额', '期末未审数', '期末未审', '期末数', '期末']
    for p in prio:
        for r in range(1, n_hdr + 1):
            for c in range(1, min(ws.max_column + 1, 30)):
                v = ws.cell(r, c).value
                if v is None:
                    continue
                s = str(v).replace(' ', '')
                if s == p or (p == '期末' and '期末' in s and '期初' not in s and '减少' not in s and '增加' not in s):
                    return c
    return None

def find_total(ws, end_col, n_hdr=6):
    """找最后含 合计/小计 的行（优先 全集团合计/年度合计/合计），取 end_col 值。"""
    cand = None
    for r in range(n_hdr + 1, ws.max_row + 1):
        for c in (1, 2):
            v = ws.cell(r, c).value
            if v is not None and ('合计' in str(v) or '小计' in str(v)):
                vv = ws.cell(r, end_col).value
                if isinstance(vv, (int, float)):
                    cand = vv
    return cand

def scan_one(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets = wb.sheetnames
    aud = det = note = None
    aud_sn = det_sn = note_sn = None
    for sn in sheets:
        ws = wb[sn]
        if '审定表' in sn and aud is None and '集团' not in sn:
            ec = find_end_col(ws)
            if ec:
                aud = find_total(ws, ec)
                aud_sn = sn
        elif '明细表' in sn and '附注' not in sn and det is None and '集团' not in sn:
            ec = find_end_col(ws)
            if ec:
                det = find_total(ws, ec)
                det_sn = sn
        elif '附注' in sn and note is None:
            ec = find_end_col(ws)
            if ec:
                note = find_total(ws, ec)
                note_sn = sn
    wb.close()
    diffs = []
    if aud is not None and det is not None:
        d = det - aud
        if abs(d) > 50:
            diffs.append(('明细vs审定', d, aud, det))
    if aud is not None and note is not None:
        d = note - aud
        if abs(d) > 50:
            diffs.append(('附注vs审定', d, aud, note))
    return os.path.basename(path), aud, det, note, diffs, (aud_sn, det_sn, note_sn)

def main():
    d = sys.argv[1]
    files = sorted(glob.glob(os.path.join(d, '*审计底稿_*.xlsx')))
    if not files:
        files = sorted(glob.glob(os.path.join(d, '*.xlsx')))
    n_diff = 0
    for fp in files:
        res = scan_one(fp)
        if not res:
            continue
        name, aud, det, note, diffs, sheets = res
        if diffs:
            n_diff += 1
            print(f'== {name}')
            for kind, d, a, de in diffs:
                print(f'   {kind} {d:,.2f}  (审定={a:,.2f} 明细/附注={de:,.2f}) [{sheets[0]} / {sheets[1] if kind=="明细vs审定" else sheets[2]}]')
    print(f'---- 差异底稿数: {n_diff}/{len(files)} ----')

if __name__ == '__main__':
    main()
