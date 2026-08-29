# -*- coding: utf-8 -*-
"""recon_ext.py —— 外部数据核对总控（2026-08-24 架构预留 R1，regen_all 阶段⑥）。

职责：把零散外部核对程序（见 recon_registry）按账套/专项统一调度，
      汇总各专项差异 JSON → 输出差异汇总表。

安全边界：
    - 只【发现差异】，不改写任何底稿审定数
    - 差异 JSON 是唯一跨阶段契约，后续人工确认 status 流转

用法：
    python recon_ext.py <数据文件夹>                  # 该账套全部已登记专项
    python recon_ext.py <文件夹> --recon=bank,cross   # 指定专项
    python recon_ext.py <文件夹> --list               # 列出已登记专项
    python recon_ext.py <文件夹> --diff-only          # 只汇总已有差异 JSON（不重新跑专项）

对应 regen_all 阶段⑥『外部试算表核对』——当前为骨架，专项逐个收编后启用。
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from recon_registry import REGISTRY, list_recons, load_recon  # noqa: E402


def _summarize_diffs(data_dir, recon_keys):
    """汇总各专项差异 JSON → 输出差异汇总表 xlsx + json。"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '外部核对差异汇总'
    ws.append(['专项', '账套', '年度', '科目', 'GL金额', '外部金额', '差异', '原因', '证据', '状态'])
    total_diff = 0.0
    n_open = 0
    for key in recon_keys:
        cfg = REGISTRY[key]
        dj = os.path.join(data_dir, cfg.get('diff_json', f'差异_{key}.json'))
        if not os.path.exists(dj):
            continue
        try:
            with open(dj, encoding='utf-8') as f:
                d = json.load(f)
        except Exception:
            continue
        for diff in d.get('diffs', []):
            ws.append([
                key, diff.get('entity', ''), diff.get('year', ''),
                diff.get('name', ''), diff.get('gl_amt'), diff.get('ext_amt'),
                diff.get('diff'), diff.get('reason', ''), diff.get('evidence', ''),
                diff.get('status', ''),
            ])
            total_diff += float(diff.get('diff') or 0)
            if diff.get('status', 'open') == 'open':
                n_open += 1
    out_xlsx = os.path.join(data_dir, '外部核对差异汇总_生成.xlsx')
    wb.save(out_xlsx)
    print(f'差异汇总：{len(recon_keys)} 专项，{n_open} 条未决，总差异 {total_diff:,.2f}')
    print(f'已生成：{out_xlsx}')
    return out_xlsx


def main():
    ap = argparse.ArgumentParser(description='外部数据核对总控（阶段⑥骨架）')
    ap.add_argument('data_dir')
    ap.add_argument('--recon', default=None, help='指定专项（逗号分隔），默认全部')
    ap.add_argument('--list', action='store_true', help='列出已登记专项')
    ap.add_argument('--diff-only', action='store_true', help='只汇总已有差异 JSON')
    args = ap.parse_args()

    if args.list:
        list_recons()
        return 0
    if not os.path.isdir(args.data_dir):
        print(f'ERROR: 无效文件夹：{args.data_dir}')
        return 2

    if args.recon:
        keys = [k.strip() for k in args.recon.split(',') if k.strip() in REGISTRY]
    else:
        keys = list(REGISTRY.keys())
    # ⚡ 2026-08-24 专项按 accts 过滤：不适用该账套的专项跳过【生成】（仍可抽取已有差异）。
    #   账套代码从 data_dir 尾段推导（d:/底稿测试/AYL → AYL）。
    _acct = os.path.basename(os.path.normpath(args.data_dir.rstrip('/\\')))
    keys = [k for k in keys
            if not REGISTRY[k].get('accts') or _acct in REGISTRY[k]['accts'] or args.recon]
    if not keys:
        print('没有可运行的专项（注册表为空或指定专项未登记/不适用该账套）。')
        return 1

    if not args.diff_only:
        for key in keys:
            try:
                from recon_registry import recon_entry
                recon_entry(key, args.data_dir)
            except Exception as ex:
                print(f'  ⚠️ {key} 执行异常：{ex}')
    _summarize_diffs(args.data_dir, keys)
    return 0


if __name__ == '__main__':
    sys.exit(main())
