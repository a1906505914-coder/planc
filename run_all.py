# -*- coding: utf-8 -*-
"""一键全套入口（#未来设计·统一入口）。

用法：
    python run_all.py ah                  # 跑 AH 三集团全套（并行）→ 自动指纹回归
    python run_all.py <数据包目录>        # 跑单数据包全套（U8/SAP 单目录，如 XBJ/ADF）
    python run_all.py ah --subj tax       # 只跑指定模块
    python run_all.py <数据包目录> --subj current_account --allow 白名单.txt

自动收尾：跑完后对输出底稿做指纹回归（--allow 白名单过滤预期差异）。
建议经 run_audit.bat 调用以固定 PYTHONHASHSEED（行序稳定）。
"""
import os
import sys
import time
import subprocess

BASE = os.path.dirname(os.path.abspath(__file__))
PY = os.environ.get('AH_PY', sys.executable)
AH_DATA = r'd:/底稿测试/AH/数据/2026'
AH_OUT = r'd:/底稿测试/AH/底稿/2026/集团'


def _run(cmd):
    print('  $ ' + ' '.join(cmd), flush=True)
    p = subprocess.run(cmd, cwd=BASE)
    return p.returncode


def main():
    argv = sys.argv[1:]
    if not argv or argv[0] in ('-h', '--help'):
        print(__doc__)
        return 1
    allow = None
    if '--allow' in argv:
        i = argv.index('--allow')
        allow = os.path.join(BASE, argv[i + 1])
    subj = None
    if '--subj' in argv:
        i = argv.index('--subj')
        subj = argv[i + 1]
    # ⚡ 2026-08-29 断点续跑透传：--no-resume 强制全量（默认 resume 跳过签名未变的已完成科目）
    no_resume = '--no-resume' in argv
    if no_resume:
        argv.remove('--no-resume')
    # ⚡ 2026-08-29 preflight 接入：跑前自动自检（磁盘空间/Excel占用/并发/数据完整性），
    #   BLOCK(2)=阻断停止、WARN(1)=警告继续；--skip-preflight 跳过。
    if '--skip-preflight' in argv:
        argv.remove('--skip-preflight')
    else:
        import preflight_check
        _pf_target = AH_DATA if target == 'ah' else os.path.abspath(target)
        _pf_rc = preflight_check.main([_pf_target])
        if _pf_rc == 2:
            print('⛔ 运行前自检阻断（磁盘/文件占用/并发等），已停止。处理后重跑，或 --skip-preflight 强行继续。', flush=True)
            return 2
        if _pf_rc == 1:
            print('⚠️ 运行前自检有警告，继续运行（可用 --skip-preflight 跳过自检）…', flush=True)
    t0 = time.time()
    target = argv[0]
    rc = 0
    if target == 'ah':
        # AH 三集团并行全套
        cmd = [PY, '-X', 'utf8', os.path.join(BASE, 'ah_parallel_run.py')]
        if subj:
            cmd += ['--subj', subj]
        if no_resume:
            cmd += ['--no-resume']
        rc = _run(cmd)
        out_root = AH_OUT
    else:
        # 单数据包：run_u8_on_sap 全模块（U8 拖拽模式；SAP 单主体由 adapter 处理）
        cmd = [PY, '-X', 'utf8', os.path.join(BASE, 'run_u8_on_sap.py'), os.path.abspath(target)]
        if subj:
            cmd += ['--subj', subj]
        if no_resume:
            cmd += ['--no-resume']
        rc = _run(cmd)
        out_root = os.path.join(target, '底稿')
    # 指纹回归（对输出目录对比基线）
    check_cmd = [PY, '-X', 'utf8', os.path.join(BASE, 'working_paper_fingerprint.py'),
                 'check', '--report', os.path.join(BASE, '_run_all_regress.txt')]
    if allow:
        check_cmd += ['--allow', allow]
    cp = subprocess.run(check_cmd, capture_output=True, text=True, encoding='utf-8')
    print(cp.stdout[-1200:] if cp.stdout else '  (无指纹输出)')
    print(f'\n一键全套完成：exit={rc}；指纹回归 {"⚠️ 有差异需检查" if cp.returncode else "✅ 无意外回退"}（耗时 {time.time()-t0:.0f}s）')
    return rc or cp.returncode


if __name__ == '__main__':
    sys.exit(main())
