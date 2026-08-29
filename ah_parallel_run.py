# -*- coding: utf-8 -*-
"""AH 三集团并行生成调度器（2026-08-25 性能优化：多进程并行 + GL 磁盘缓存控内存）。

背景：AH 88 主体分 1010(1)/1357(19)/2468(68) 三集团。原串行跑全套约 2 小时+。
方案：3 个 python 子进程同时跑三集团（--group --only <本集团主体> --group-name <集团码>），
    每进程经 _GROUP_COMPS 过滤只 discover/加载本集团主体 → GL 只解析本集团子集 → 内存可控；
    audit_cache 磁盘缓存按 entities 签名隔离 → 各进程缓存互不冲突（跨进程复用首次解析结果）。
预计：全套 13 模块 × 3 集团 ≈ 原串行耗时的 1/3（约 40-50 分钟）。

用法：
    python ah_parallel_run.py [--subj current_account,payroll] [--out-dir 集团稿根目录] [--no-resume]
    断点续跑（默认开）：某集团崩溃后重跑会自动跳过已完成科目；--no-resume 强制全量。
    --jobs N：限制并行集团数（默认 3；内存不足/卡死时用 --jobs 1 串行）。
"""
import os
import sys
import subprocess
import time

BASE = os.path.dirname(os.path.abspath(__file__))
PY = os.environ.get('AH_PY', sys.executable)
DATA = r'd:/底稿测试/AH/数据/2026'
OUT_ROOT = r'd:/底稿测试/AH/底稿/2026/集团'

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
    subj_opt = []
    if '--subj' in argv:
        i = argv.index('--subj')
        subj_opt = ['--subj', argv[i + 1]]
    # ⚡ 2026-08-29 断点续跑透传：--no-resume 强制全量（默认 resume，崩溃后重跑自动续跑）
    if '--no-resume' in argv:
        subj_opt = ['--no-resume'] + subj_opt
    out_root = OUT_ROOT
    if '--out-dir' in argv:
        i = argv.index('--out-dir')
        out_root = argv[i + 1]
    # ⚡⚡ 2026-08-29 P0：并行内存瓶颈防护——三集团并行每进程 7GB+，内存叠加触发
    #   虚拟内存交换 → 大模块（inventory/tax）卡死数十分钟（实测 18 分钟）。支持
    #   --jobs N 限制同时运行的集团数（--jobs 1=串行），默认仍 3 并行（内存充足时快）。
    jobs = 3
    if '--jobs' in argv:
        i = argv.index('--jobs')
        try:
            jobs = max(1, min(3, int(argv[i + 1])))
        except Exception:
            jobs = 3
    t0 = time.time()
    procs = []
    tmp_dir = os.path.join(os.path.expanduser('~'), 'WorkBuddy', '2026-07-15-15-26-19', 'tmp')
    os.makedirs(tmp_dir, exist_ok=True)
    _groups = list(GROUPS.items())
    for g, comps in _groups:
        out_dir = os.path.join(out_root, g)
        os.makedirs(out_dir, exist_ok=True)
        log = open(os.path.join(tmp_dir, f'ah_parallel_{g}.log'), 'w', encoding='utf-8')
        cmd = [PY, '-X', 'utf8', os.path.join(BASE, 'run_u8_on_sap.py'),
               DATA, '--group', '--only', ','.join(comps),
               '--group-name', g, '--out-dir', out_dir] + subj_opt
        print(f'[启动] 集团 {g}（{len(comps)} 主体）-> {out_dir}', flush=True)
        if jobs < 3:
            p = subprocess.run(cmd, stdout=log, stderr=subprocess.STDOUT)
            log.close()
            print(f'[完成] 集团 {g}: exit={p.returncode}', flush=True)
        else:
            p = subprocess.Popen(cmd, stdout=log, stderr=subprocess.STDOUT)
            procs.append((g, p, log))
    # 等待全部完成
    for g, p, log in procs:
        rc = p.wait()
        log.close()
        print(f'[完成] 集团 {g}: exit={rc}', flush=True)
    print(f'\n全部完成，总耗时 {time.time() - t0:.1f}s（{time.time() - t0:.0f} 秒）', flush=True)


if __name__ == '__main__':
    main()
