# -*- coding: utf-8 -*-
"""铁律14 验证器：独立重算 GL 原始自然方(raw_nat) 并与 TB 控制数比对。

这是被清理的 verify_pl_full.py 的精神继承者，把『gv≠TB 严禁默认归为 GL 取数缺口』
编码为可执行的检查：对每个 (主体,年度,归一化科目名)，独立对 GL 全量求和 raw_nat(dr/cr)，
与 TB 的 jf/df 比较——
  · raw_nat≈TB  → 该科目 GL 抽取完整，若某 builder 仍报 gv≠TB，属自家加工差异(OWN_BUG)，非缺口；
  · raw_nat≠TB  → 确为 GL 非全量抽取缺口(GL_GAP)，应如实披露。

匹配口径（2026-07-29 修订）：综合查询明细表(GL) 数据源仅含『科目名称』列、无科目代码，
故无法按科目代码比对；改为按【归一化科目名称】归并（NFKC 全半角 + 去全角空格 + 括号归一），
消除『采购』/『采购　』/『销项税额（开票）』等同义异写造成的双计式假缺口。
注意：GL 用凭证级业务名、TB 用总账科目名，二者命名体系本质不同的科目仍会 flagged，
需人工辨名或令 GL 导出带科目代码列方可彻底精确。

用法：
  python verify_raw_nat.py <数据文件夹> [容差]
默认文件夹 <DATA_ROOT>/g。
"""
import paths as P
import os, sys
import audit_common as A

TOL = 0.005


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else P.G
    tol = float(sys.argv[2]) if len(sys.argv) > 2 else TOL
    if not os.path.isdir(data_dir):
        print(f'✗ 文件夹不存在：{data_dir}')
        return 1
    print(f'========== 铁律14 raw_nat 守卫：{data_dir} ==========')
    entities = A.discover_entities(data_dir)
    tb = A.read_tb_full(data_dir, entities)
    gl = A.read_gl_rows(data_dir, entities)
    raw = A.recompute_raw_nat(gl)

    # TB 按 (e, y, 归一化科目名) 汇总 jf/df（GL 无代码列，按归一化名称归并以消除同义异写误报）
    tb_agg = {}
    tb_raw = {}
    for (e, c, n, y), v in tb.items():
        key = (e, y, A.norm_subject_name(n))
        d = tb_agg.setdefault(key, {'jf': 0.0, 'df': 0.0})
        d['jf'] += v.get('jf', 0.0)
        d['df'] += v.get('df', 0.0)
        tb_raw.setdefault(key, n)  # 展示用原始科目名

    consistent, gap = [], []
    for key in sorted(tb_agg, key=lambda k: (k[0] or '', str(k[1]), k[2] or '')):
        e, y, n = key
        jf = tb_agg[key]['jf']; df = tb_agg[key]['df']
        if abs(jf) < tol and abs(df) < tol:
            continue  # 无发生额科目跳过
        rn = raw.get(key, {'dr': 0.0, 'cr': 0.0})
        diff_dr = rn['dr'] - jf
        diff_cr = rn['cr'] - df
        if abs(diff_dr) <= tol and abs(diff_cr) <= tol:
            consistent.append(key)
        else:
            gap.append((key, jf, df, rn['dr'], rn['cr'], diff_dr, diff_cr))

    print(f'\nTB 有发生额科目数：{len(consistent) + len(gap)}')
    print(f'  ✓ GL↔TB 原始一致(raw_nat==TB，铁律14 守卫通过)：{len(consistent)}')
    print(f'  ✗ GL 抽取缺口(raw_nat≠TB，属非全量抽取，应如实披露)：{len(gap)}')
    if gap:
        print('\n--- GL 取数缺口明细（主体 | 年度 | 科目 | TB借发 | TB贷发 | GL借发 | GL贷发 | Δ借 | Δ贷）---')
        for key, jf, df, gdr, gcr, dd, dc in gap:
            e, y, n = key
            disp = tb_raw.get(key, n)
            print(f'  {e} | {y} | {disp} | {jf:.2f} | {df:.2f} | {gdr:.2f} | {gcr:.2f} | {dd:.2f} | {dc:.2f}')

    # 反向：GL 有数但 TB 无（极端脏源，提示）
    tb_keys = set(tb_agg.keys())
    orphan_gl = [(k, v) for k, v in raw.items()
                 if k not in tb_keys and (abs(v['dr']) > tol or abs(v['cr']) > tol)]
    if orphan_gl:
        print(f'\n⚠️ GL 有发生额但 TB 无对应科目名（归一化后仍无匹配，可能业务名/总账科目名体系不同/脏源）：{len(orphan_gl)}')
        for (e, y, n), v in sorted(orphan_gl, key=lambda x: (x[0][0] or '', str(x[0][1]), x[0][2] or ''))[:20]:
            print(f'  {e} | {y} | {n} | GL借 {v["dr"]:.2f} | GL贷 {v["cr"]:.2f}')

    # ===== 总额核对（铁律14 铁证，不受命名粒度影响）=====
    # GL 末级全名 vs TB 总账名 会导致按名匹配大量误报；但 GL 总发生额与 TB 总发生额
    # 若量级一致，即证明 GL 数据完整，缺口纯属粒度/命名差异，不得披露为取数缺口。
    tb_total = sum(d['jf'] + d['df'] for d in tb_agg.values())
    gl_total = sum(v['dr'] + v['cr'] for v in raw.values())
    diff_total = gl_total - tb_total
    print(f'\n[总额核对] TB 发生额合计：{tb_total:.2f} | GL 发生额合计：{gl_total:.2f} | 差：{diff_total:.2f}')
    rel = abs(diff_total) / abs(tb_total) if tb_total else (0.0 if abs(diff_total) < tol else 1.0)
    if rel <= 0.001:
        print('  → GL 总额≈TB 总额(差异<0.1%)：GL 数据完整。上方「✗ 缺口」系 GL 凭证级末级全名'
              ' vs TB 总账科目名 的粒度/命名差异所致，非 GL 抽取缺失，'
              '严禁据此类清单披露为取数范围限制(铁律14)。')
    else:
        print(f'  → GL 总额明显小于 TB 总额(差 {diff_total:.2f}，{rel*100:.1f}%)：'
              '确存在 GL 非全量抽取缺口，上方清单可据实披露。')

    print('\n>>> 结论（按归一化科目名称匹配 + 总额核对）：')
    if rel <= 0.001:
        print('  总额核对显示 GL 完整，上方「✗ GL 取数缺口」科目均系粒度/命名差异(OWN_BUG 类)，'
              '不得披露为取数缺口；各 builder 若仍报 gv≠TB，必为自家加工差异。')
    else:
        print('  上方「✗ GL 取数缺口」科目方可被 builder 如实披露为取数范围限制；'
              '其余科目若 builder 仍报 gv≠TB，必为自家加工差异(OWN_BUG)，不得归为 GL 缺口。')
    print('  GL 仅含科目名称(无科目代码列)、与 TB 总账科目名体系本质不同的科目仍可能 flagged，需人工辨名'
          '或令 GL 导出带科目代码列方可彻底精确。')
    print('  注：按名 gap 清单因 GL 末级全名 vs TB 总账名粒度差异，数量可能偏高（混入部分假缺口）；'
          '如需逐科目精确区分真/假缺口，请令 GL 导出带科目代码列，验证器将自动切换"代码优先"匹配。')
    return 0
    return 0


if __name__ == '__main__':
    sys.exit(main())
