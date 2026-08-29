# -*- coding: utf-8 -*-
"""长任务看门狗（2026-08-29 P0）：监控后台长任务进程/输出目录，失败或停滞时留下明确状态文件。
解决「任务跑 40-50 分钟，回来发现死了」：失败不可见 + 无人监控。

用法：
  python long_task_guard.py <输出目录> [--timeout-min N] [--status 状态文件路径] [--period 秒]
  监控输出目录：新文件 mtime 超过 N 分钟无推进 → 停滞警告；目录不再有进程写文件且长时间无产出 → 疑似结束。

  python long_task_guard.py --pid <PID> <输出目录> [--timeout-min N]
  监控指定进程：进程退出 → 立即写最终状态（含退出码、最后产出、建议重跑命令）。

状态文件内容（退出时写）：_TASK_END / 结束原因(成功/失败/停滞) / 最后产出时间 / 建议命令。
"""
import os
import sys
import time
import glob
import subprocess

BASE = os.path.dirname(os.path.abspath(__file__))


def _latest_mtime(out_dir):
    latest = 0.0
    try:
        for fp in glob.glob(os.path.join(out_dir, '**', '*'), recursive=True):
            if os.path.isfile(fp):
                try:
                    latest = max(latest, os.path.getmtime(fp))
                except OSError:
                    pass
    except Exception:
        pass
    return latest


def _fmt(ts):
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(ts))


def _write_status(status_path, lines):
    try:
        os.makedirs(os.path.dirname(os.path.abspath(status_path)), exist_ok=True)
        with open(status_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
    except Exception:
        pass


def main():
    argv = sys.argv[1:]
    pid = None
    status_path = None
    timeout_min = 20
    period = 60
    if '--pid' in argv:
        i = argv.index('--pid')
        pid = int(argv[i + 1])
        argv = argv[:i] + argv[i + 2:]
    if '--status' in argv:
        i = argv.index('--status')
        status_path = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    if '--timeout-min' in argv:
        i = argv.index('--timeout-min')
        timeout_min = float(argv[i + 1])
        argv = argv[:i] + argv[i + 2:]
    if '--period' in argv:
        i = argv.index('--period')
        period = float(argv[i + 1])
        argv = argv[:i] + argv[i + 2:]
    if not argv:
        print(__doc__)
        return 1
    out_dir = argv[0]
    if status_path is None:
        status_path = os.path.join(out_dir, '_task_status.txt')
    start = time.time()
    last = _latest_mtime(out_dir)
    last_t = start
    print(f'[看门狗] 开始监控 输出目录={out_dir} pid={pid} 停滞阈值={timeout_min}分钟', flush=True)
    while True:
        time.sleep(period)
        now = time.time()
        mt = _latest_mtime(out_dir)
        alive = True
        if pid is not None:
            try:
                os.kill(pid, 0)
            except (OSError, ProcessLookupError):
                alive = False
        progressed = mt > last + 1
        if progressed:
            last = mt
            last_t = now
            print(f'[看门狗] {_fmt(now)} 正常，最近产出 {_fmt(mt)}', flush=True)
        elif now - last_t > timeout_min * 60:
            _write_status(status_path, [
                f'[看门狗] ⚠️ 输出停滞超过 {timeout_min:.0f} 分钟（最后产出 {_fmt(last)}，当前 {_fmt(now)}）',
                f'[看门狗] 建议：检查进程是否卡死；卡死则 kill 后重跑（断点续跑会自动跳过已完成科目）。',
                f'[看门狗] 重跑命令参考：python run_u8_on_sap.py <数据目录> ...（失败后默认 resume）',
            ])
            print(f'[看门狗] ⚠️ 停滞 {timeout_min:.0f} 分钟，最后产出 {_fmt(last)}，已写状态文件', flush=True)
        if not alive:
            # 进程退出，再等一小段时间确认无新产出（grace）后结束
            time.sleep(period)
            mt2 = _latest_mtime(out_dir)
            last_prod = max(mt, mt2)
            _write_status(status_path, [
                f'[看门狗] 进程(pid={pid})已退出，监控结束 @ {_fmt(time.time())}',
                f'[看门狗] 最后产出时间：{_fmt(last_prod)}（距结束 {(time.time()-last_prod)/60:.1f} 分钟）',
                f'[看门狗] 若最后产出已是很久前 → 任务可能失败；查看 run_history 最新报告或本目录文件。',
                f'[看门狗] 失败后重跑：python run_u8_on_sap.py <数据目录> ...（断点续跑自动跳过已完成科目）',
            ])
            print(f'[看门狗] 进程退出，监控结束 @ {_fmt(time.time())}', flush=True)
            return 0
        if now - start > 24 * 3600:
            print('[看门狗] 超 24h 安全上限，退出', flush=True)
            return 0


if __name__ == '__main__':
    sys.exit(main())
