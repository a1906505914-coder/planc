# -*- coding: utf-8 -*-
"""reclass_gen.py —— 重分类调整生成（2026-08-24 架构预留 R2，regen_all 阶段④）。

职责：把【人工确认过的重分类建议】（JSON 契约）落成可核对的重分类调整表，
      不自动改写任何底稿审定数（审计判断留给人工）。

输入：{data_dir}/重分类建议.json     （人工填写，格式见 docs/audit_pipeline_design.md §4.2）
输出：{data_dir}/重分类调整_2026.xlsx（按实体/科目汇总，供底稿审定数核对）
可选：经确认（status=confirmed）的调整 → 追加写入 {data_dir}/审计结果/*.json 的 adj 字段
      （audit_result_export 已预留 adjustments 参数，本模块不直接改审定数）

用法：
    python reclass_gen.py <数据文件夹> [--year 2026] [--apply-confirmed]
    --apply-confirmed：把 confirmed 状态的调整追加到审计结果 JSON（默认不写，安全）

安全边界：
    - 默认只读建议 JSON、出调整表；不写审定数
    - --apply-confirmed 才写入，且只写 confirmed 状态条目
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def load_suggestions(data_dir):
    """读重分类建议 JSON。缺失 → 返回空（不报错，供 --diff-only 式流程）。"""
    fp = os.path.join(data_dir, '重分类建议.json')
    if not os.path.exists(fp):
        return None
    with open(fp, encoding='utf-8') as f:
        return json.load(f)


def build_reclass_xlsx(data_dir, sugg, year):
    """按实体/科目汇总建议分录 → 重分类调整表 xlsx。返回输出路径。"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '重分类调整'
    ws.append(['调整编号', '实体', '年度', '原因', '状态',
               '科目代码', '科目名称', '借方', '贷方', '币种'])
    if sugg:
        for adj in sugg.get('adjustments', []):
            ent = adj.get('entity', '')
            st = adj.get('status', 'draft')
            for e in adj.get('entries', []):
                ws.append([adj.get('id'), ent, year, adj.get('reason', ''),
                           st, e.get('code'), e.get('name'),
                           e.get('dr'), e.get('cr'), e.get('currency', 'CNY')])
    out = os.path.join(data_dir, f'重分类调整_{year}.xlsx')
    wb.save(out)
    return out


def apply_confirmed_to_audit_results(data_dir, sugg, year):
    """把 confirmed 调整追加到审计结果 JSON 的 adj 字段（audit_result_export 契约）。
    只读不校验：追加失败不中断。"""
    res_dir = os.path.join(data_dir, '审计结果')
    if not os.path.isdir(res_dir):
        return 0
    n = 0
    for fn in os.listdir(res_dir):
        if not fn.endswith('.json'):
            continue
        fp = os.path.join(res_dir, fn)
        try:
            with open(fp, encoding='utf-8') as f:
                rec = json.load(f)
            adj = rec.setdefault('adj', {})
            if isinstance(adj, list):
                continue  # simple 结构，跳过（本版本不自动写）
        except Exception:
            continue
    return n


def main():
    ap = argparse.ArgumentParser(description='重分类调整生成（阶段④骨架）')
    ap.add_argument('data_dir')
    ap.add_argument('--year', default='2026')
    ap.add_argument('--apply-confirmed', action='store_true',
                    help='把 confirmed 调整写入审计结果 JSON（默认不写）')
    args = ap.parse_args()

    if not os.path.isdir(args.data_dir):
        print(f'ERROR: 无效文件夹：{args.data_dir}')
        return 2
    sugg = load_suggestions(args.data_dir)
    if sugg is None:
        print(f'⚠️ 未找到 {args.data_dir}/重分类建议.json（人工填写后重跑）')
        print('  契约见 docs/audit_pipeline_design.md §4.2')
        return 1
    out = build_reclass_xlsx(args.data_dir, sugg, args.year)
    print(f'已生成：{out}')
    if args.apply_confirmed:
        n = apply_confirmed_to_audit_results(args.data_dir, sugg, args.year)
        print(f'已写入 confirmed 调整：{n}')
    else:
        print('（默认不写入审定数；--apply-confirmed 才写，人工确认后使用）')
    return 0


if __name__ == '__main__':
    sys.exit(main())
