# -*- coding: utf-8 -*-
"""yy 账套断点续跑脚本（2026-08-12）
挂机/关机恢复后运行：读 _yy_progress.txt 快照 → 跑剩余主体 → 日志落 E 盘。
独立进程（双击/命令行启动），不依赖 WorkBuddy 会话——agent 后台任务会被会话回收，
本脚本由用户自己启动，进程完全独立，关机/重启后可从快照续跑。
用法:
    python resume_yy.py           # 续跑全部剩余主体（自动跳过已完成）
    python resume_yy.py --all     # 忽略快照，全量重跑 88 家
"""
import paths as P
import os
import re
import subprocess
import sys

ROOT = P.YY
PROG = P.APP
PY = sys.executable
SNAP = os.path.join(ROOT, '_yy_progress.txt')
LOG = os.path.join(ROOT, '_yy_resume.log')


def read_snapshot():
    done, failed = set(), set()
    if os.path.exists(SNAP):
        with open(SNAP, encoding='utf-8') as f:
            for line in f:
                if line.startswith('已完成'):
                    done = set(re.findall(r'\d{4}', line))
                elif line.startswith('失败'):
                    failed = set(re.findall(r'\d{4}', line))
    return done, failed


def all_comps():
    sys.path.insert(0, PROG)
    import sap_adapter as A
    A._CACHE_ENT.clear()
    A.set_root(ROOT)
    return set(A.discover_entities(ROOT).keys())


def main():
    if '--all' in sys.argv:
        comps = sorted(all_comps())
        print(f'全量重跑 {len(comps)} 家')
    else:
        done, failed = read_snapshot()
        allc = all_comps()
        todo = sorted(allc - done - failed)
        print(f'快照: 已完成 {len(done)} 失败 {len(failed)} | 本次续跑 {len(todo)} 家')
        if not todo:
            print('✓ 快照内全部主体已完成（含失败补跑），无需续跑')
            return
        comps = todo
    cmd = [PY, os.path.join(PROG, '_sap_batch.py')] + comps + ['--workers', '3']
    print('启动:', ' '.join(cmd))
    print(f'日志: {LOG}')
    with open(LOG, 'w', encoding='utf-8') as f:
        r = subprocess.run(cmd, stdout=f, stderr=subprocess.STDOUT, cwd=PROG)
    print(f'退出码 {r.returncode} | 日志尾部:')
    try:
        with open(LOG, encoding='utf-8', errors='ignore') as f:
            print('\n'.join(f.readlines()[-8:]))
    except Exception:
        pass


if __name__ == '__main__':
    main()
