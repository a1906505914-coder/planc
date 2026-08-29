# -*- coding: utf-8 -*-
"""
cleanup_subtotal_rows.py — 撤销 add_entity_subtotal.py 在真实文件上误插的行。
仅删除"主体列值恰为 '小计'"的行（本工具插入标记）。原始底稿自带的小计行把'小计'
放在名称列、主体列仍是账套名，不会被删。幂等：无则不动。
用法：python cleanup_subtotal_rows.py [folder...]   (默认全部10文件夹)
"""
import paths as P
import sys, os, glob
import openpyxl
from openpyxl.utils import get_column_letter
FOLDERS = ['ACB', 'ADZ', 'AZ', 'ATT', 'AYL', 'AS', 'AJJ', 'AQ', 'AFJ', 'ARF']
DESK = P.DATA_ROOT
SUBJ_KEYWORDS = ['账套主体', '核算主体', '往来单位名称', '主体', '客户', '公司', '单位名称', '名称']

def find_subj_col(header_cells):
    best = None; best_pri = len(SUBJ_KEYWORDS)
    for i, c in enumerate(header_cells):
        if not isinstance(c, str): continue
        for pri, kw in enumerate(SUBJ_KEYWORDS):
            if c == kw or c.strip() == kw:
                if pri < best_pri: best_pri = pri; best = i
                break
    return best

def cleanup_file(path):
    try:
        wb = openpyxl.load_workbook(path)
    except PermissionError:
        return [f'  [LOCKED skip] {os.path.basename(path)}']
    except Exception as e:
        return [f'  [load err] {os.path.basename(path)}: {e}']
    total = 0
    for ws in wb.worksheets:
        if not any(p in ws.title for p in ('明细', '分月', '汇总', '账龄', '对比', '比较')):
            continue
        # 找表头
        hdr = None; subj = None
        for ri in range(1, min(ws.max_row, 12) + 1):
            row = [ws.cell(row=ri, column=c).value for c in range(1, min(ws.max_column, 12) + 1)]
            for i, c in enumerate(row):
                if isinstance(c, str) and c in SUBJ_KEYWORDS:
                    hdr = ri; subj = i; break
            if hdr is not None: break
        if subj is None:
            continue
        # 收集要删的行（自底向上）
        del_rows = []
        for ri in range(hdr + 1, ws.max_row + 1):
            v = ws.cell(row=ri, column=subj + 1).value
            if isinstance(v, str) and v.strip() == '小计':
                del_rows.append(ri)
        if del_rows:
            for ri in sorted(del_rows, reverse=True):
                ws.delete_rows(ri, 1)
            total += len(del_rows)
    if total:
        wb.save(path)
        return [f'FILE {os.path.basename(path)} 删除 {total} 行']
    wb.close()
    return []

def main():
    args = sys.argv[1:]
    fols = args if args else FOLDERS
    for fol in fols:
        base = os.path.join(DESK, fol) if os.path.isdir(os.path.join(DESK, fol)) else fol
        if not os.path.isdir(base):
            print(f'no folder {fol}'); continue
        print(f'\n########## FOLDER {fol} ##########')
        for f in sorted(glob.glob(os.path.join(base, '*生成*.xlsx'))):
            for line in cleanup_file(f):
                print(line)

if __name__ == '__main__':
    main()
