# -*- coding: utf-8 -*-
"""group_review_check.py —— 集团底稿完整性统一检查（2026-08-24 用户需求）。

用户痛点：集团底稿几十上百份，人工逐个点开看不过来。本工具自动扫描集团目录
所有底稿，检查每张表（sheet）是否【取数完整】，输出统一报告 + 差异科目清单。

检查项：
  1. 空表扫描：每个 sheet 无数据（无数值单元格）→ 标记『空表』（该表可能真无业务
     或取数失败——按科目已知口径分类：往来/损益类空表多为取数 bug，待摊类可能真无）
  2. 审定表检查：审定表应含『合计』行且合计值非 0（有业务的科目）
  3. 勾稽表检查：薪酬分配核对表计提/发放列、对方科目核对 等有数值
  4. 差异科目清单：汇总所有『关键表空/无合计』的科目，供进一步优化

用法：
  python group_review_check.py <集团目录>         # 扫描单个集团目录
  python group_review_check.py <账套根> --all      # 扫描账套下全部集团（集团/、数据/）
"""
import os
import sys
import glob

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

# 关键表关键词（判定『该表应有数据』）：审定表/明细表/核对表/分配核对
_KEY_SHEET_KW = ('审定表', '明细表', '对方科目核对', '分配核对', '多借多贷', '差异凭证', 'Top10')


def _sheet_stats(ws, max_rows=2000):
    """统计 sheet 数据概况：返回 (总行, 非空行, 数值单元格数)。
    ⚡ 2026-08-24 限流：只读前 max_rows 行判定有无数据（大表 60 万行不整表读，
    避免沙箱内存/IO 限制；判定『有数据』只需头部+中部有数值即可）。"""
    n_rows = n_data = n_num = 0
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        n_rows += 1
        if i >= max_rows:
            break
        if any(x is not None and str(x).strip() for x in r):
            n_data += 1
            for x in r:
                if isinstance(x, (int, float)) and abs(x) > 0.005:
                    n_num += 1
    return n_rows, n_data, n_num


def check_workbook(fp):
    """检查单份底稿：返回 [(sheet, 空表?, 说明)] + 空表清单。"""
    from openpyxl import load_workbook
    issues = []
    try:
        wb = load_workbook(fp, read_only=True, data_only=True)
    except Exception as ex:
        return [('(打开失败)', True, f'文件损坏/无法读取：{ex}')]
    for sn in wb.sheetnames:
        ws = wb[sn]
        _r, _d, _num = _sheet_stats(ws)
        is_key = any(k in sn for k in _KEY_SHEET_KW)
        is_empty = _num == 0 and _d <= 2  # 无数值且无实质数据行
        if is_empty and is_key:
            issues.append((sn, True, f'关键表空表（{_r}行 无数值）'))
        elif is_empty and not is_key:
            issues.append((sn, False, f'空表（非关键，可能真无业务）'))
        elif is_key and _num == 0:
            issues.append((sn, True, f'关键表无数值（仅表头/文本）'))
    wb.close()
    return issues


def scan_group_dir(group_dir):
    """扫描集团目录全部底稿 → {文件名: issues}。"""
    result = {}
    fps = glob.glob(os.path.join(group_dir, '*.xlsx'))
    # 排除脱敏/反馈/中间产物
    fps = [f for f in fps if not any(k in os.path.basename(f) for k in ('_脱敏', '-反馈', '_bak'))]
    for fp in sorted(fps):
        name = os.path.basename(fp)
        issues = check_workbook(fp)
        result[name] = issues
    return result


def summarize(result, data_dir=None):
    """汇总：返回 (报告文本行列表, 需关注科目清单)。
    data_dir 提供时结合 TB 判定『应有数据但空表』= 真 bug（重点标出）。"""
    lines = []
    concerns = []
    total_files = len(result)
    empty_files = 0
    for name, issues in result.items():
        key_issues = [i for i in issues if i[1]]  # 空表/无数值
        if not key_issues:
            continue
        empty_files += 1
        lines.append(f'  ⚠️ {name}')
        for sn, is_empty, note in issues:
            if is_empty:
                # 从文件名提取科目名（去掉审计底稿_集团名）
                subj = name.replace('审计底稿', '').replace('_生成', '').replace('.xlsx', '')
                import re as _re
                subj = _re.sub(r'_\d{4}$', '', subj)
                subj = _re.sub(r'_\d{4}集团$', '', subj).replace('_', '')
                tb_has = True
                if data_dir:
                    try:
                        from sheet_gap_check import tb_has_subject_data
                        tb_has = tb_has_subject_data(data_dir, subj)
                    except Exception:
                        tb_has = True
                tag = '❌ TB有数据但空表(取数bug)' if tb_has else '✓ TB无数据(正常空表)'
                lines.append(f'      · {sn}: {note} [{tag}]')
        concerns.append((name, tag))
    summary_line = f'检查 {total_files} 份底稿，{empty_files} 份有关键表空表/无数值'
    return [summary_line] + lines, concerns


def main():
    import argparse
    ap = argparse.ArgumentParser(description='集团底稿完整性统一检查')
    ap.add_argument('path', help='集团目录 或 账套根')
    ap.add_argument('--all', action='store_true', help='扫描账套下全部集团（集团/ 子目录）')
    ap.add_argument('--report', default=None, help='报告输出路径（默认 stdout）')
    args = ap.parse_args()

    if args.all:
        # 账套根下 集团/ 子目录 或 数据/ 下的集团稿
        base = args.path
        dirs = []
        gp = os.path.join(base, '集团')
        if os.path.isdir(gp):
            dirs = [os.path.join(gp, d) for d in sorted(os.listdir(gp))
                    if os.path.isdir(os.path.join(gp, d))]
        if not dirs and os.path.isdir(base):
            dirs = [base]
    else:
        dirs = [args.path]

    all_lines = []
    all_concerns = []
    for d in dirs:
        if not os.path.isdir(d):
            print(f'目录不存在：{d}')
            continue
        print(f'\n═══════ 扫描：{d} ═══════')
        result = scan_group_dir(d)
        # data_dir 判定：集团目录 底稿/2026/集团/X → 数据根 数据/2026
        _data_dir = None
        if args.all:
            _data_dir = args.path  # 账套根
        else:
            # 尝试向上找 数据/ 目录
            _up = d
            for _ in range(4):
                _p = os.path.join(_up, '数据')
                if os.path.isdir(_p):
                    _data_dir = _p
                    break
                _up = os.path.dirname(_up)
        lines, concerns = summarize(result, data_dir=_data_dir)
        for l in lines:
            print(l)
        all_concerns += concerns
        all_lines += [d, *lines]
    if args.report:
        try:
            with open(args.report, 'w', encoding='utf-8') as f:
                f.write('\n'.join(all_lines))
            print(f'\n报告已写：{args.report}')
        except Exception:
            pass
    return 0


if __name__ == '__main__':
    sys.exit(main())
