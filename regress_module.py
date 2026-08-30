# -*- coding: utf-8 -*-
"""模块级回归脚本（#未来设计·分级回归体系落地）。

改某模块后，只重跑该模块（三集团）+ 指纹对比，快速确认无意外回退。
替代"改一行就全量重跑 50 分钟"的粗粒度验证。

用法：
    python regress_module.py current_account           # 跑三集团 current_account 模块
    python regress_module.py tax --group 1010,2468     # 只跑指定集团
    python regress_module.py current_account --allow fp_allow_example.txt

原理：
    1. 对每个集团调 run_u8_on_sap --subj <模块>（只生成该模块底稿，省时）
    2. 跑 working_paper_fingerprint.check 对比基线
    3. --allow 白名单过滤预期差异；剩余差异 = 意外回退，需检查
"""
import os
import sys
import time
import subprocess
import paths as P

BASE = os.path.dirname(os.path.abspath(__file__))
PY = os.environ.get('AH_PY', sys.executable)
DATA = os.path.join(P.DATA_DIRS['AH'], '数据', '2026')
OUT_ROOT = os.path.join(P.DATA_DIRS['AH'], '底稿', '2026', '集团')
GROUPS = {
    '1010': ['1010'],
    '1357': ['1020', '1030', '1040', '1050', '1060', '1070', '1080', '1090', '1100',
             '1170', '1220', '1240', '1250', '1260', '3010', '5020', '9020', '9030', '9070'],
    '2468': ['2010', '2020', '2030', '2040', '2050', '2070', '2080', '2090', '2100',
             '2110', '2130', '2150', '2160', '2200', '2210', '2220', '2230', '2240',
             '2250', '2260', '2270', '2280', '2290', '2300', '2310', '2330', '2340',
             '2350', '2360', '2370', '2380', '2390', '2400', '2410', '2420', '2430',
             '2450', '2460', '2470', '2480', '2490', '2500', '2510', '2520', '2530',
             '2540', '2550', '2570', '2580', '2590', '2600', '2610', '2630', '2640',
             '2650', '2660', '2680', '2690', '2700', '2720', '2730', '2740', '2760',
             '4030', '4040', '8020', '8030', '8040'],
}


def main():
    argv = sys.argv[1:]
    if not argv or argv[0] in ('-h', '--help'):
        print(__doc__)
        return 1
    module = argv[0]
    groups = list(GROUPS)
    allow = None
    if '--group' in argv:
        i = argv.index('--group')
        groups = argv[i + 1].split(',')
    if '--allow' in argv:
        i = argv.index('--allow')
        allow = os.path.join(BASE, argv[i + 1])
    t0 = time.time()
    ok_all = True
    for g in groups:
        comps = GROUPS.get(g, [g])
        out_dir = os.path.join(OUT_ROOT, g)
        print(f'=== [{g}] 重跑 {module} …', flush=True)
        cmd = [PY, '-X', 'utf8', os.path.join(BASE, 'run_u8_on_sap.py'),
               DATA, '--group', '--only', ','.join(comps),
               '--group-name', g, '--out-dir', out_dir, '--subj', module]
        p = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        if p.returncode != 0:
            print(f'  ❌ [{g}] {module} 运行失败 exit={p.returncode}')
            ok_all = False
            continue
        # 指纹对比（当前集团）
        check_cmd = [PY, '-X', 'utf8', os.path.join(BASE, 'working_paper_fingerprint.py'),
                     'check', '--report', os.path.join(BASE, f'_regress_{g}.txt')]
        if allow:
            check_cmd += ['--allow', allow]
        cp = subprocess.run(check_cmd, capture_output=True, text=True, encoding='utf-8')
        print(cp.stdout[-1500:] if cp.stdout else '  (无指纹输出)')
        if cp.returncode != 0:
            ok_all = False
    print(f'\n模块级回归完成：{"✅ 全部通过（无意外回退）" if ok_all else "⚠️ 存在需检查的差异"}（耗时 {time.time()-t0:.0f}s）')
    return 0 if ok_all else 1


if __name__ == '__main__':
    sys.exit(main())
