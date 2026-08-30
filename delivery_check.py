# -*- coding: utf-8 -*-
"""交付防线脚本（delivery_check.py，2026-08-30 新建）。
用户要求：不再"给意见不落地"，把已知检查自动化，交付前一次跑完。

检查项：
  1. 命名一致性：目录中不得存在新旧两套命名并存（_生成 / _YYYY_生成 残留且已有新版）
  2. 文件占用：所有 xlsx 可独占打开（被 Excel/WPS 占用会静默顶替新文件 → 交付前拦截）
  3. 完整性：复用 wp_completeness_check（关键表空/合计不符/缺 sheet）
  4. 四表一致（可选项 --recon4）：审定表Σ主体 = 明细表合计 = 附注Σ  （费用/收入类科目）

用法：
  python delivery_check.py --dir <底稿目录> [--data <数据目录>] [--recon4]
"""
import os
import sys
import argparse

def check_naming(out_dir):
    """新旧命名并存检查：有 {base}_XBJ.xlsx（新）却残留 {base}_2025_XBJ.xlsx / {base}_生成.xlsx。"""
    issues = []
    fns = os.listdir(out_dir)
    names = {}
    for fn in fns:
        if not fn.endswith('.xlsx'):
            continue
        key = fn.replace('_2025_XBJ.xlsx', '').replace('_2026_XBJ.xlsx', '')\
                .replace('_XBJ.xlsx', '').replace('_生成.xlsx', '')\
                .replace('_2025_生成.xlsx', '').replace('_2026_生成.xlsx', '')
        names.setdefault(key, []).append(fn)
    for key, group in sorted(names.items()):
        if len(group) > 1:
            # 有非 _生成 新版（_XBJ）且又有 _生成/_YYYY_生成 旧版 → 残留
            newish = [f for f in group if '_生成' not in f]
            oldish = [f for f in group if '_生成' in f]
            if newish and oldish:
                issues.append((key, sorted(group)))
    return issues

def check_locked(out_dir):
    """文件占用检查：尝试独占打开所有 xlsx。"""
    locked = []
    for fn in sorted(os.listdir(out_dir)):
        if not fn.endswith('.xlsx'):
            continue
        p = os.path.join(out_dir, fn)
        try:
            fd = os.open(p, os.O_RDWR)
            os.close(fd)
        except (PermissionError, OSError):
            locked.append(fn)
    return locked

def recon4(data_dir, out_dir, code_names):
    """四表一致：审定表Σ主体 vs 明细表合计 vs 附注Σ。code_names: [(科目关键字, 文件名关键字)]。"""
    import openpyxl
    from sap_reader import read_sap_tb
    diffs = []
    for subj_kw, file_kw in code_names:
        for fn in sorted(os.listdir(out_dir)):
            if file_kw not in fn or not fn.endswith('.xlsx'):
                continue
            fp = os.path.join(out_dir, fn)
            try:
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
            except Exception:
                continue
            sns = wb.sheetnames
            ds_sn = [s for s in sns if '审定表' in s and '集团' not in s and '合并' not in s]
            mx_sn = [s for s in sns if '明细表' in s]
            note_sn = [s for s in sns if '附注' in s]
            if not ds_sn:
                continue
            # 审定表 Σ主体本期数（找"本期数"列或数值列）
            ws = wb[ds_sn[0]]
            ds_tot = 0.0
            hdr_row = None
            rows = list(ws.iter_rows(values_only=True))
            for i, r in enumerate(rows[:5]):
                if r and '核算主体' in [str(x) if x else '' for x in r[:6]]:
                    hdr_row = i
                    break
            if hdr_row is None:
                continue
            col = None
            for j, h in enumerate(rows[hdr_row][:6]):
                if h and ('本期' in str(h) or '审定' in str(h)):
                    col = j
                    break
            if col is None:
                col = 2
            for r in rows[hdr_row + 1:]:
                if r and r[0] and str(r[0]).strip() not in ('合计', '总计', '') and \
                        isinstance(r[col], (int, float)):
                    ds_tot += r[col]
            # 明细表合计
            mx_tot = 0.0
            for sn in mx_sn:
                for row in wb[sn].iter_rows(values_only=True):
                    if row and row[0] and '合计' in str(row[0]) and isinstance(row[3], (int, float)):
                        mx_tot += row[3]
            # 附注未审数合计（列式：Σ所有数值列）
            note_tot = 0.0
            for sn in note_sn:
                for row in wb[sn].iter_rows(values_only=True):
                    if row and row[0] and '合计' in str(row[0]):
                        for c in row[1:]:
                            if isinstance(c, (int, float)):
                                note_tot += c
            if mx_tot or note_tot:
                if abs(ds_tot - mx_tot) > 0.01 or abs(ds_tot - note_tot) > 0.01:
                    diffs.append((fn, round(ds_tot, 2), round(mx_tot, 2), round(note_tot, 2)))
            wb.close()
    return diffs

def main():
    ap = argparse.ArgumentParser(description='交付防线：交付前一次跑完命名/占用/完整性检查')
    ap.add_argument('--dir', required=True, help='底稿目录')
    ap.add_argument('--data', default=None, help='数据目录（完整性/试算表核对用）')
    ap.add_argument('--recon4', action='store_true', help='额外做四表一致核对')
    ap.add_argument('--code-names', default='财务费用,财务费用;管理费用,管理费用;销售费用,销售费用;营业收入,营业收入',
                    help='四表核对科目: 关键字,文件名关键字;...')
    args = ap.parse_args()
    out_dir = os.path.abspath(args.dir)
    if not os.path.isdir(out_dir):
        print(f'❌ 目录不存在: {out_dir}')
        return 3
    issues = []

    # 1) 命名一致性
    name_issues = check_naming(out_dir)
    if name_issues:
        issues.append('命名不一致（新旧并存）:')
        for key, group in name_issues:
            issues.append(f'  - {key}: {group}')
    else:
        print('✅ 命名一致性：无新旧并存')

    # 2) 文件占用
    locked = check_locked(out_dir)
    if locked:
        issues.append(f'文件被占用（{len(locked)} 个，交付前须关闭）:')
        for fn in locked:
            issues.append(f'  - {fn}')
    else:
        print('✅ 文件占用：全部可写')

    # 3) 完整性（复用现有检查器）
    if args.data:
        print('⏳ 完整性检查运行中…')
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from wp_completeness_check import run as _run_check
        ok, rep = _run_check(args.dir, args.data)
        print(f'   → {rep}')
        if not ok:
            issues.append(f'完整性检查未通过: {rep}')

    # 4) 四表一致（可选）
    if args.recon4:
        pairs = [tuple(p.split(',')) for p in args.code_names.split(';') if ',' in p]
        diffs = recon4(args.data or '', out_dir, pairs)
        if diffs:
            issues.append('四表不一致（审定/明细/附注）:')
            for fn, a, b, c in diffs:
                issues.append(f'  - {fn}: 审定={a:,.2f} 明细={b:,.2f} 附注={c:,.2f}')
        else:
            print('✅ 四表一致：审定=明细=附注')

    if issues:
        print('\n❌ 交付防线未通过:')
        for i in issues:
            print(i)
        return 1
    print('\n✅ 交付防线全部通过')
    return 0

if __name__ == '__main__':
    sys.exit(main())
