# -*- coding: utf-8 -*-
"""tool_common.py —— 工具脚本公共参数解析（2026-08-31 建，消除 20 处 AH/2026 硬编码）。

目标：所有工具脚本统一接受 --acct/--data/--year，不再把账套名/年份写死在代码里。

用法：
    import argparse, tool_common as T
    ap = argparse.ArgumentParser()
    T.add_tool_args(ap)                     # 注入 --acct / --data / --year
    args = ap.parse_args()
    DATA, YEAR = T.resolve_data(args)       # 数据目录 + 年份（--data 优先于 --acct）

优先级：--data（显式数据目录） > --acct（账套名，paths.py 映射） > 默认账套。
年份：--year 显式 > 从数据目录路径推断（'数据/2026' → 2026）。
"""
import os
import re

import paths as P


def infer_year(data_dir):
    """从数据目录路径推断年份：'.../数据/2026' → '2026'；兜底取路径中最后 4 位数字段。"""
    m = re.search(r'[\\/]数据[\\/]?(\d{4})', data_dir)
    if m:
        return m.group(1)
    m = re.findall(r'(\d{4})', data_dir)
    return m[-1] if m else '2026'


def detect_year(acct_root):
    """从账套根下『数据』目录自动发现年份（取最新=数字最大）；无则账套根直接年份目录。"""
    years = []
    for d in (os.path.join(acct_root, '数据'), acct_root):
        if os.path.isdir(d):
            years += [fn for fn in os.listdir(d) if re.fullmatch(r'\d{4}', fn)]
        if years:
            break
    return sorted(set(years))[-1] if years else '2026'


def acct_data(acct, year=None):
    """账套名 → (数据目录, 年份)。year 缺省从账套目录自动发现。"""
    root = P.DATA_DIRS.get(acct)
    if not root:
        raise ValueError(f'未知账套: {acct}（可选: {P.ACCT_NAMES}）')
    year = str(year or detect_year(root))
    return os.path.join(root, '数据', year), year


def add_tool_args(ap, default_acct=None):
    """注入统一工具参数：--acct 账套名 / --data 数据目录 / --year 年份。"""
    ap.add_argument('--acct', default=default_acct,
                    help=f'账套名（paths.ACCT_NAMES: {P.ACCT_NAMES}）')
    ap.add_argument('--data', help='数据目录（优先于 --acct）')
    ap.add_argument('--year', help='年份（默认从数据目录路径推断）')
    return ap


def resolve_data(args, default_acct=None):
    """返回 (数据目录, 年份)。优先级：--data > --acct > default_acct。"""
    data = getattr(args, 'data', None)
    acct = getattr(args, 'acct', None) or default_acct
    if data:
        data = os.path.abspath(data)
    elif acct:
        data, _ = acct_data(acct, getattr(args, 'year', None))
    else:
        raise SystemExit('需指定 --data 或 --acct（工具脚本不再默认写死 AH/2026）')
    year = getattr(args, 'year', None) or infer_year(data)
    return data, str(year)
