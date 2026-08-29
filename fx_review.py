# -*- coding: utf-8 -*-
"""境外报表折算专项统一入口（2026-08-18 用户定：fx_consolidate/fx_rates/thai_mapping 提炼为一个专项）。

三者原为一个整体（fx_consolidate 调用 fx_rates 汇率 + thai_mapping 泰国映射），
现统一入口（fx_review.py）按 fx_manifest 配置路由。
新增账套：fx_manifest 加一行即可，不新建脚本。

用法：
  python fx_review.py AZ            # Z集团 境外报表折算
"""
import os
import sys

import paths as P
import fx_manifest as FM


def gen(acct, out_dir=None):
    cfg = FM.job_for(acct)
    if not cfg:
        raise KeyError(f'fx_manifest 无配置: {acct}')
    import fx_consolidate as FC
    split_root = cfg['split_root']
    if not os.path.isdir(split_root):
        print(f'⚠️ 目录不存在：{split_root}')
        return []
    return FC.process(split_root, out_dir or cfg.get('out_dir'))


def main():
    import argparse
    ap = argparse.ArgumentParser(description='境外报表折算专项')
    ap.add_argument('acct', help='账套（AZ，见 fx_manifest）')
    ap.add_argument('--out', default=None, help='输出目录（默认 split_root/合并折算）')
    a = ap.parse_args()
    rc = gen(a.acct, a.out)
    if isinstance(rc, int):
        print(f'{a.acct}: 折算完成（返回码 {rc}）')
    else:
        outs = rc if isinstance(rc, list) else [rc]
        print(f'{a.acct}: 生成 {len(outs)} 个文件')
    return 0


if __name__ == '__main__':
    sys.exit(main())
