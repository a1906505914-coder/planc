# -*- coding: utf-8 -*-
"""科目覆盖审计：TB 一级科目 vs 已生成底稿 覆盖缺口扫描。

用途：消除『硬编码漏科目』——某 TB 一级科目无任何生成器底稿 → 静默缺失，
要等试算表核对才发现（其他非流动资产 1831 被预付 kw 误吸收就是这类）。

用法：
    python scan_coverage.py --data <data_dir> --dir <out_dir> [--min-amt 10000]

原理：
- 从 TB（read_tb_full 全量）聚合『一级科目』（名称首段，code 前缀 4 位去重合并）
- 从 out_dir 收集已生成底稿科目名（{科目}审计底稿*.xlsx）
- 匹配：一级科目名 被 任一底稿科目名 包含 → 覆盖；否则 → 缺口
- 输出缺口清单（按金额排序），供决定是否需要补生成器/兜底底稿
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def load_tb_l1(data_dir):
    import sap_adapter as A
    A._DATA_ROOT = data_dir
    tb = A.read_tb_full(data_dir, None)
    from collections import defaultdict
    l1 = defaultdict(float)   # 一级名 -> Σ|qm|
    l1_cnt = defaultdict(int) # 一级名 -> 主体数
    # ⚡ 一级科目按 code[:4] 聚合（SAP 子目 10 位、无 4 位父级行），
    #   一级名取组内名称公共段（多数子目 split('-')[0] 一致）。避免按名称
    #   split('-') 把『管理费用』的二级明细（防暑降温费 等无前缀）误判为一级。
    grp = defaultdict(list)   # code[:4] -> [(name, amt)]
    for (e, c, n, y), v in tb.items():
        nm = str(n or '')
        if not nm:
            continue
        cd = str(c)[:4]
        amt = abs(float(v.get('qm', 0.0)))
        grp[cd].append((nm, amt))
    for cd, items in grp.items():
        if not items:
            continue
        # 一级名：多数子目名的 '-' 首段（最常见者）
        from collections import Counter
        heads = Counter(nm.split('-')[0].strip() for nm, _ in items if nm.strip())
        if not heads:
            continue
        base = heads.most_common(1)[0][0]
        tot = sum(a for _, a in items)
        l1[base] += tot
        l1_cnt[base] += len(items)
    return l1, l1_cnt


def load_papers(out_dir):
    """收集已生成底稿科目名。命名形态：{科目}审计底稿*.xlsx"""
    papers = set()
    if not os.path.isdir(out_dir):
        return papers
    for fn in os.listdir(out_dir):
        if fn.endswith('.xlsx') and '审计底稿' in fn:
            subj = fn.split('审计底稿')[0].strip()
            if subj:
                papers.add(subj)
    return papers


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True)
    ap.add_argument('--dir', required=True)
    ap.add_argument('--min-amt', type=float, default=10000.0,
                    help='缺口金额下限（默认 1 万），低于此不提示')
    args = ap.parse_args()

    l1, l1_cnt = load_tb_l1(args.data)
    papers = load_papers(args.dir)
    print(f'TB 一级科目数: {len(l1)} | 已生成底稿科目数: {len(papers)}')
    print(f'底稿科目: {sorted(papers)}')
    print()

    # 匹配：一级名 含于 某底稿科目名 或 反向（长名包含短名）
    def covered(nm):
        for p in papers:
            if nm and (nm in p or p in nm):
                return p
        return None

    uncovered = []
    for nm, amt in sorted(l1.items(), key=lambda x: -x[1]):
        p = covered(nm)
        if p is None and amt >= args.min_amt:
            uncovered.append((nm, amt, l1_cnt[nm]))
            print(f'❌ 未覆盖  {nm:<20} Σ|qm|={amt:>18,.0f}  主体数={l1_cnt[nm]}')
    print()
    print(f'未覆盖且有余额(>={args.min_amt:,.0f})的一级科目: {len(uncovered)}')
    if not uncovered:
        print('✅ 全部一级科目均有底稿覆盖')
    return 0 if not uncovered else 1


if __name__ == '__main__':
    sys.exit(main())
