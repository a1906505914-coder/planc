# -*- coding: utf-8 -*-
"""data_pipeline.py —— 外部数据流水线总控（2026-08-24 架构 R4）。

把「数据获取/预处理」与「核对差异」两个维度串成一条流水线：
    ① 数据源程序（data_source_registry）→ 产出去结构化数据/中间产物
    ② 核对专项（recon_registry）       → 从底稿抽差异 → 统一差异 JSON
    ③ 差异汇总（recon_ext）            → 差异汇总表 xlsx

安全边界：只【获取/发现】，不改写底稿审定数。

用法：
    python data_pipeline.py <账套根>                    # 全部数据源 + 全部核对
    python data_pipeline.py <账套根> --source=contract  # 只跑指定数据源
    python data_pipeline.py <账套根> --recon=bank       # 只跑指定核对
    python data_pipeline.py <账套根> --list             # 列出全部可编排项
    python data_pipeline.py <账套根> --extract-only     # 只抽取差异（不重新生成）
"""
import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from data_source_registry import REGISTRY as SRC_REG, run_source, list_sources  # noqa: E402
from recon_registry import REGISTRY as REC_REG, recon_entry, list_recons       # noqa: E402


def main():
    ap = argparse.ArgumentParser(description='外部数据流水线总控（数据获取→核对→汇总）')
    ap.add_argument('data_dir', help='账套根（如 d:/底稿测试/AZ）')
    ap.add_argument('--source', default=None, help='指定数据源（逗号分隔），默认全部')
    ap.add_argument('--recon', default=None, help='指定核对专项（逗号分隔），默认全部')
    ap.add_argument('--list', action='store_true', help='列出全部可编排项')
    ap.add_argument('--extract-only', action='store_true', help='只抽差异不重新生成')
    args = ap.parse_args()

    if args.list:
        print('=== 数据源程序 ===')
        list_sources()
        print()
        print('=== 核对专项 ===')
        list_recons()
        return 0
    if not os.path.isdir(args.data_dir):
        print(f'ERROR: 无效文件夹：{args.data_dir}')
        return 2

    # ① 数据源程序（数据获取/预处理）
    src_keys = [k for k in SRC_REG] if not args.source else \
        [k.strip() for k in args.source.split(',') if k.strip() in SRC_REG]
    if not args.extract_only:
        for key in src_keys:
            try:
                print(f'--- 数据源 [{key}] {SRC_REG[key]["name"]} ---')
                out = run_source(key, args.data_dir)
                if out:
                    print(f'  → 产出：{out}')
            except Exception as ex:
                print(f'  ⚠️ 数据源 {key} 异常：{ex}')
        print()

    # ② 核对专项（生成+抽取 或 仅抽取）
    rec_keys = [k for k in REC_REG] if not args.recon else \
        [k.strip() for k in args.recon.split(',') if k.strip() in REC_REG]
    for key in rec_keys:
        try:
            print(f'--- 核对 [{key}] {REC_REG[key]["name"]} ---')
            out = recon_entry(key, args.data_dir, extract_only=args.extract_only)
            if out:
                print(f'  → 差异 JSON：{out}')
        except Exception as ex:
            print(f'  ⚠️ 核对 {key} 异常：{ex}')
    print()

    # ③ 差异汇总
    from recon_ext import _summarize_diffs
    _summarize_diffs(args.data_dir, rec_keys)
    return 0


if __name__ == '__main__':
    sys.exit(main())
