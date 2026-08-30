# -*- coding: utf-8 -*-
"""GL 缓存批量预热（2026-08-26）：AH 88 主体 .gl_cache 重建。
遍历 GROUPS 全部主体，逐个 load_comp_gl_disk 触发构建（未命中则全扫落盘）。
1010 已有缓存（key 匹配则命中跳过，不匹配则重扫）。
用法：python -X utf8 gl_prewarm.py
"""
import sys, os, time, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import sap_common as C
import paths as P
from ah_parallel_run import GROUPS

DATA = os.path.join(P.DATA_DIRS['AH'], '数据', '2026')

def main():
    # ⚡⚡ 2026-08-31 参数化：--acct/--data/--year 覆盖顶部硬编码
    import argparse, tool_common as T
    global DATA
    _ap = argparse.ArgumentParser(add_help=False)
    T.add_tool_args(_ap, default_acct='AH')
    _args, _ = _ap.parse_known_args()
    _D, _Y = T.resolve_data(_args)
    _AR = os.path.dirname(os.path.dirname(_D))
    DATA = _D
    _O = os.path.join(_AR, '底稿', _Y, '集团')
    try:
        OUT_ROOT = _O
    except NameError:
        pass
    try:
        ROOT = _O
    except NameError:
        pass
    print(f'[tool_common] DATA={DATA} YEAR={_Y}', flush=True)

    comps = [c for g in GROUPS.values() for c in g]
    print(f'[预热] 共 {len(comps)} 主体：1010({len(GROUPS["1010"])}) 1357({len(GROUPS["1357"])}) 2468({len(GROUPS["2468"])})', flush=True)
    t0 = time.time()
    ok = miss = err = 0
    for i, comp in enumerate(comps, 1):
        t1 = time.time()
        try:
            rows = C.load_comp_gl_disk(DATA, comp, C._GL_CACHE_COLS)
            # 判定是否命中缓存：看缓存文件存在且 sig 匹配（load 函数已处理；这里按行数判断）
            print(f'[{i}/{len(comps)}] {comp}: {len(rows)} 行 / {time.time()-t1:.0f}s', flush=True)
            ok += 1
        except Exception as e:
            print(f'[{i}/{len(comps)}] {comp}: ERR {e}', flush=True)
            err += 1
    print(f'[预热完成] OK={ok} ERR={err} 总耗时={time.time()-t0:.0f}s', flush=True)

if __name__ == '__main__':
    main()
