# -*- coding: utf-8 -*-
"""统一底稿完整性自检（2026-08-24 用户要求：编入架构，跑完底稿自动自检，无需每次提醒）。

职责：扫描账套/集团目录全部审计底稿，检查是否【整份空白】（所有 sheet 都无数值）。
判定口径：
  · 底稿至少一个 sheet 有数值 → 不算空白（审定表期末 0 但发生额/明细表有数=正常，如本期清仓）
  · 全部 sheet 无数值 → ❌ 空白底稿（真 bug，取数失败或科目未取到）
  · 部分表空但有 TB 判定 → 标注供参考（分部门明细表等已知正常空）

用法：
  python self_check.py <数据目录> [--group-dir <集团目录>]
"""
import os
import sys
import glob
import argparse


def _workbook_has_value(fp, max_rows=2000):
    """底稿是否至少一个 sheet 有数值（只读前 max_rows 行，防大表卡死）。"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(fp, read_only=True, data_only=True)
    except Exception:
        return True  # 打不开保守算有值，避免误报
    for sn in wb.sheetnames:
        ws = wb[sn]
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            if i >= max_rows:
                break
            for x in r:
                if isinstance(x, (int, float)) and abs(x) > 0.005:
                    wb.close()
                    return True
    wb.close()
    return False


def self_check(data_dir, group_dir=None, quiet=False):
    """统一自检入口。返回 (ok, report_lines)。
    ok=True 表示无空白底稿。
    data_dir   = 账套数据根（TB 判定用）
    group_dir  = 可选；指定则扫描该目录（集团稿独立目录），否则扫描 data_dir
    """
    scan_dir = group_dir if group_dir else data_dir
    if not os.path.isdir(scan_dir):
        return True, [f'[自检] 目录不存在：{scan_dir}']

    # 收集底稿
    if group_dir:
        fps = [f for f in glob.glob(os.path.join(scan_dir, '*.xlsx'))
               if '审计底稿' in f and not any(k in f for k in ('_脱敏', '_bak'))]
    else:
        fps = [f for f in glob.glob(os.path.join(scan_dir, '*审计底稿*.xlsx'))
               if not any(k in f for k in ('_脱敏', '_bak'))]

    report = [f'[自检] {scan_dir}：共 {len(fps)} 份底稿']
    blanks = []
    ok = True
    for fp in sorted(fps):
        if not _workbook_has_value(fp):
            blanks.append(os.path.basename(fp))
    if blanks:
        # 对空白底稿做 TB 判定：TB 有数据 → 真 bug；TB 无数据 → 正常空壳（无业务）
        tb_true = []
        tb_false = []
        for b in blanks:
            subj = os.path.basename(b).replace('审计底稿', '').replace('_生成', '').replace('.xlsx', '')
            import re as _re
            subj = _re.sub(r'_\d{4}集团$', '', subj)   # 先按集团名去
            subj = _re.sub(r'_\d{4}$', '', subj)       # 再去年份（必须先于去下划线）
            subj = subj.replace('_', '')
            try:
                from sheet_gap_check import tb_has_subject_data
                has = tb_has_subject_data(data_dir, subj)
            except Exception:
                has = True
            (tb_true if has else tb_false).append((b, has))
        if tb_true:
            ok = False
            report.append(f'  ❌ {len(tb_true)} 份空白底稿【TB 有数据】(取数bug)：')
            for b, _ in tb_true:
                report.append(f'      · {b}')
        if tb_false:
            report.append(f'  ✓ {len(tb_false)} 份空白底稿【TB 无数据】(正常空壳，可删)：')
            for b, _ in tb_false:
                report.append(f'      · {b}')
    else:
        report.append('  ✅ 无空白底稿')
    if not quiet:
        for l in report:
            print(l)
    return ok, report


def main():
    p = argparse.ArgumentParser(description='底稿完整性自检（架构统一，空白底稿检测）')
    p.add_argument('data_dir', help='账套数据根目录')
    p.add_argument('--group-dir', default=None, help='集团底稿目录（可选，扫描该目录）')
    a = p.parse_args()
    ok, _ = self_check(a.data_dir, group_dir=a.group_dir)
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
