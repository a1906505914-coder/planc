# -*- coding: utf-8 -*-
"""TB 父级一致性修正工具（2026-08-24 用户要求：不能以『父级陈旧』结束，需修正让核对一致）。

背景：多个账套源 TB 的【父级行 ≠ 子目合计】（父级陈旧/只含部分子目）。底稿用子目明细
（正确），recon_self 自建试算表用父级 → 阶段③核对系统性差异（AS 固定资产 8.3亿、
GJX 在建 2.76亿 等，差异精确=父级-子目差）。

修正口径：**有子目的父级行，qm 一律以子目合计为准**（叶子合计口径，与底稿一致）。
不改源文件：本工具返回修正视图（内存），供 recon_self 核对用；另可输出修正快照到中间产物。

用法：
  python tb_parent_fix.py <数据目录>            # 列出父级不一致清单（dry-run）
  python tb_parent_fix.py <数据目录> --fix      # 输出修正后 TB 快照到 中间产物/TB父级修正/
"""
import os
import sys
import argparse
from collections import defaultdict


def find_parent_mismatch(tb):
    """扫描 TB，返回父级≠子目合计 的清单 [(主体, 代码4, 名称, 父级, 子目合计, 差)]。
    仅统计【存在子目】的科目（父级即末级无子目=正常）。"""
    by = {}
    for (e, c, n, y), v in tb.items():
        cs = str(c)
        if len(cs) < 4:
            continue
        key = (e, cs[:4])
        d = by.setdefault(key, {'name': None, 'pv': 0.0, 'kv': 0.0, 'nk': 0})
        if len(cs) == 4:
            d['name'] = str(n)
            d['pv'] = float(v.get('qm') or 0)
        else:
            d['kv'] += float(v.get('qm') or 0)
            d['nk'] += 1
    out = []
    for (e, pre), d in by.items():
        if d['name'] is None or d['nk'] == 0:
            continue
        diff = d['kv'] - d['pv']
        if abs(diff) > 1000:
            out.append((e, pre, d['name'], d['pv'], d['kv'], diff, d['nk']))
    return out


def leaf_total_tb(tb):
    """返回修正视图：有子目的父级行 qm → 子目合计（叶子口径）。
    仅修正 qm；qc/jf/df 不修正（核对只看 qm 期末，发生额类科目另有口径）。
    ⚡⚡ 2026-08-25 修复（AYL 应交税费 82,012.82 vs 41,006.41 双计根因）：
    原实现把【所有 >4 位子目】都计入 kid_sum，中间级（222101 应交增值税，本身
    有子目 22210101~08）被重复累计 → 父级=中间级+叶子 双计。改为只累计
    【末级叶子】（tb 中无更长子代码的行）。"""
    # 预扫代码集合，判断末级叶子：存在同主体更长代码（以该码开头）→ 非叶子
    code_len = {}
    for (e, c, n, y), v in tb.items():
        e_key = e
        cs = str(c)
        code_len.setdefault(e_key, set()).add(cs)
    kid_sum = defaultdict(float)
    has_kid = defaultdict(bool)
    for (e, c, n, y), v in tb.items():
        cs = str(c)
        if len(cs) < 4:
            continue
        if len(cs) > 4:
            # ⚡ 末级叶子判定：无同主体更长子代码
            if not any(cs2 != cs and cs2.startswith(cs) for cs2 in code_len.get(e, ())):
                kid_sum[(e, cs[:4])] += float(v.get('qm') or 0)
            has_kid[(e, cs[:4])] = True
    out = {}
    for key, v in tb.items():
        e, c, n, y = key
        cs = str(c)
        if len(cs) == 4 and has_kid.get((e, cs)):
            v2 = dict(v)
            v2['qm'] = kid_sum[(e, cs)]
            out[key] = v2
        else:
            out[key] = v
    return out


def main():
    p = argparse.ArgumentParser(description='TB 父级一致性修正（叶子合计口径）')
    p.add_argument('data_dir')
    p.add_argument('--fix', action='store_true', help='输出修正后 TB 快照')
    a = p.parse_args()
    import audit_common
    ents = audit_common.discover_entities(a.data_dir)
    tb = audit_common.read_tb_full(a.data_dir, ents)
    mm = find_parent_mismatch(tb)
    print(f'{a.data_dir}：{len(mm)} 处父级≠子目合计（有子目）')
    for e, pre, nm, pv, kv, diff, nk in mm:
        print(f'  {e} {pre} {str(nm)[:14]}: 父级={pv:,.2f} 子目={kv:,.2f} 差={diff:,.2f} ({nk}子目)')
    if a.fix and mm:
        tb2 = leaf_total_tb(tb)
        outdir = os.path.join(a.data_dir, '..', '..', '中间产物', 'TB父级修正')
        outdir = os.path.abspath(outdir)
        os.makedirs(outdir, exist_ok=True)
        import pickle
        fp = os.path.join(outdir, f'{os.path.basename(a.data_dir)}_tb_fixed.pkl')
        with open(fp, 'wb') as f:
            pickle.dump(tb2, f)
        print(f'修正视图已导出：{fp}')
    return 0 if not mm else 1


if __name__ == '__main__':
    sys.exit(main())
