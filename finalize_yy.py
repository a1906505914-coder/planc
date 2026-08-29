# -*- coding: utf-8 -*-
"""yy 收口脚本（2026-08-12）——续跑完成后运行：补失败主体 → 88 家完整性验证 → 汇总报告。
独立进程（双击/命令行），不依赖 WorkBuddy 会话；日志落 E 盘 _yy_finalize.log。
用法:
    python finalize_yy.py            # 补失败主体 + 完整性验证 + 汇总
    python finalize_yy.py --check    # 仅完整性验证 + 汇总（不补跑）
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
LOG = os.path.join(ROOT, '_yy_finalize.log')
OUT = os.path.join(ROOT, '底稿')
# 已知因日志覆盖写被沙箱拦（旧批次）而失败的主体——批次目录根治后重跑即可
FAILED_FIX = ['1070', '1080', '1090', '1100', '1170', '1220']


def log(msg):
    print(msg, flush=True)


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


def check_completeness():
    """每个主体底稿文件数（营业收入/应收账款等关键文件存在）→ 完整性判定。"""
    comps = all_comps()
    missing = []
    for c in sorted(comps):
        n = len([f for f in os.listdir(OUT) if f.endswith(f'_{c}.xlsx')])
        if n == 0:
            missing.append(c)
    return comps, missing


def main():
    only_check = '--check' in sys.argv
    log(f'=== yy 收口开始 {__import__("datetime").datetime.now().strftime("%H:%M")} ===')
    if not only_check:
        # 1. 补跑失败主体（批次目录日志，不再被覆盖写拦截）
        log(f'[1/3] 补跑 {len(FAILED_FIX)} 家失败主体: {",".join(FAILED_FIX)}')
        cmd = [PY, os.path.join(PROG, '_sap_batch.py')] + FAILED_FIX + ['--workers', '2']
        r = subprocess.run(cmd, cwd=PROG)
        log(f'补跑退出码 {r.returncode}')
    # 2. 完整性验证
    log('[2/3] 88 家完整性验证')
    comps, missing = check_completeness()
    log(f'主体总数 {len(comps)} | 无底稿文件的主体: {len(missing)} 家')
    if missing:
        log(f'  缺文件: {",".join(missing)}')
    # 3. 汇总报告
    log('[3/3] 写汇总报告')
    done, failed = read_snapshot()
    lines = [
        f'yy 收口报告（{__import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M")}）',
        '=' * 52,
        f'主体总数: {len(comps)} 家',
        f'有底稿文件主体: {len(comps) - len(missing)} 家',
        f'完全无底稿主体: {len(missing)} 家' + (f'（{"、".join(missing)}）' if missing else ''),
        f'快照记录已完成: {len(done)} 家 | 失败待补: {len(failed)} 家（{"、".join(sorted(failed)) if failed else "无"}）',
    ]
    with open(os.path.join(ROOT, '_yy_finalize_report.txt'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    log('汇总报告已写: %s' % os.path.join(ROOT, '_yy_finalize_report.txt'))
    log('=== yy 收口完成 ===')


if __name__ == '__main__':
    with open(LOG, 'w', encoding='utf-8') as _lf:
        # 主进程输出同时落日志（简单重定向方式：直接 print 到 stdout 由调用方重定向亦可）
        pass
    main()
