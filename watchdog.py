# -*- coding: utf-8 -*-
"""看门狗（2026-08-16 铁律132）：长任务卡死检测 + 【自愈重启】。
用法：
  python watchdog.py <任务名> <日志文件> [超时分钟=20] [心跳秒=60] [--auto-restart <命令文件.cmd>] [--max-restart 3]
行为：
  · 每 心跳秒 检查一次日志 mtime/行数；超 超时分钟 无新行 → 卡死
  · 卡死处置：有 --auto-restart → 【自动 kill 进程 + 重启命令】（重启次数 < max 时），
    写入 <日志>.ALERT（含已自动重启 N 次）；无 --auto-restart → 仅写 ALERT 报警
  · 日志恢复增长 → 清除 ALERT；进程消失 → 写 <日志>.DEAD
  · 心跳文件 <日志>.HB 每轮写入（证明 watchdog 自身存活）
配套：
  · run_u8_on_sap.py 科目开始打印（定位卡死科目）
  · 命令文件 .cmd：含要执行的重启命令（bash 语法，如：
      cd "<APP_ROOT>" && <PYTHON_EXE> -u run_u8_on_sap.py ... )
"""
import paths   # 数据路径中心锚点（数据绝对路径只允许出现在 paths.py）
import os
import re
import subprocess
import sys
import time


def kill_proc(pid):
    """Windows kill 进程（taskkill 优先，PowerShell 兜底）。"""
    try:
        subprocess.run(['taskkill', '/PID', str(pid), '/F'], capture_output=True, timeout=15)
    except Exception:
        pass
    time.sleep(3)


def start_cmd(cmd_fp):
    """执行命令文件（bash -c），返回新 PID。"""
    try:
        with open(cmd_fp, encoding='utf-8') as f:
            cmd = f.read().strip()
    except Exception:
        with open(cmd_fp, encoding='gbk', errors='replace') as f:
            cmd = f.read().strip()
    # bash -c 启动（继承日志重定向由命令自身负责）
    proc = subprocess.Popen(['bash', '-c', cmd], cwd=os.path.dirname(os.path.abspath(cmd_fp)))
    return proc.pid


def tail_lines(fp, n=6, enc='gbk'):
    try:
        with open(fp, 'rb') as f:
            f.seek(0, 2)
            size = f.tell()
            f.seek(max(0, size - 8192))
            raw = f.read()
        text = raw.decode(enc, errors='replace')
        lines = [l for l in text.splitlines() if l.strip()]
        return lines[-n:]
    except Exception:
        return ['<无法读取日志>']


def proc_alive(pid):
    """Windows 进程存活判断（ctypes OpenProcess，无 subprocess 线程——避免 reader thread 异常）。"""
    if not pid:
        return False
    try:
        import ctypes
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
        if h:
            ctypes.windll.kernel32.CloseHandle(h)
            return True
        return False
    except Exception:
        return False


def main():
    if len(sys.argv) < 3:
        print('用法: python watchdog.py <任务名> <日志文件> [超时分钟=20] [心跳秒=60] [--auto-restart <cmd文件>] [--max-restart N]')
        return 1
    name = sys.argv[1]
    log_fp = sys.argv[2]
    timeout_min = float(sys.argv[3]) if len(sys.argv) > 3 else 20.0
    heartbeat_s = float(sys.argv[4]) if len(sys.argv) > 4 else 60.0
    cmd_fp = None
    max_restart = 3
    if '--auto-restart' in sys.argv:
        cmd_fp = sys.argv[sys.argv.index('--auto-restart') + 1]
    if '--max-restart' in sys.argv:
        max_restart = int(sys.argv[sys.argv.index('--max-restart') + 1])
    hb_fp = log_fp + '.HB'          # 心跳文件（watchdog 自身活着）
    alert_fp = log_fp + '.ALERT'    # 卡死报警
    dead_fp = log_fp + '.DEAD'      # 进程消失
    pid_fp = log_fp + '.PID'

    last_size = -1
    last_alert_at = 0.0
    n_restart = 0
    print(f'[watchdog:{name}] 监控 {log_fp} | 超时 {timeout_min}min | 心跳 {heartbeat_s}s'
          + (f' | 自愈重启: {cmd_fp} (上限 {max_restart})' if cmd_fp else ' | 仅报警'), flush=True)
    while True:
        # ⚡⚡ 每段独立 try：任何一步异常都不能跳过【卡死检测】（2026-08-16 教训：safe-delete 异常
        #   曾使整轮跳过，卡死 23 分钟未被处理）
        mtime = size = 0.0
        idle_min = timeout_min + 1   # 异常时按"已超时"处理（宁可误重启不误卡死）
        # ① 心跳
        try:
            mtime = os.path.getmtime(log_fp)
            size = os.path.getsize(log_fp)
            now = time.time()
            idle_min = (now - mtime) / 60.0
            with open(hb_fp, 'w') as f:
                f.write(f'watchdog={name} alive={time.strftime("%H:%M:%S")} log_idle={idle_min:.1f}min size={size} restart={n_restart}\n')
        except Exception as ex:
            print(f'[watchdog:{name}] 心跳异常: {ex}', flush=True)
        # ② 恢复检测：日志增长 → ALERT 置为'已恢复'（覆盖写入，不 os.remove——safe-delete 会拦）
        try:
            if size > last_size:
                if os.path.exists(alert_fp):
                    try:
                        with open(alert_fp, 'w', encoding='utf-8') as f:
                            f.write(f'✅ 已恢复 [{name}] {time.strftime("%H:%M:%S")}，日志恢复增长\n')
                    except Exception:
                        pass
                    print(f'[watchdog:{name}] ✅ 日志恢复增长，ALERT 置为已恢复', flush=True)
                last_size = size
        except Exception as ex:
            print(f'[watchdog:{name}] 恢复检测异常: {ex}', flush=True)
        # ③ 卡死检测 → 自愈重启 or 报警（永不跳过）
        try:
            if idle_min > timeout_min:
                do_restart = cmd_fp and n_restart < max_restart
                tail = tail_lines(log_fp)
                if do_restart:
                    pid = 0
                    if os.path.exists(pid_fp):
                        try:
                            pid = int(open(pid_fp).read().strip() or 0)
                        except ValueError:
                            pid = 0
                    if pid:
                        print(f'[watchdog:{name}] 🔴 卡死({idle_min:.0f}min) → 自动 kill {pid} + 重启 #{n_restart+1}', flush=True)
                        kill_proc(pid)
                    new_pid = start_cmd(cmd_fp)
                    n_restart += 1
                    with open(pid_fp, 'w') as f:
                        f.write(str(new_pid))
                    with open(alert_fp, 'w', encoding='utf-8') as f:
                        f.write(f'🔄 卡死自动重启 [{name}] 第 {n_restart} 次\n')
                        f.write(f'  日志 {timeout_min}min 无更新（最后写入 {time.strftime("%H:%M:%S", time.localtime(mtime))}）\n')
                        f.write(f'  已 kill pid={pid}，重启新 pid={new_pid}\n')
                        f.write(f'  最后进度:\n')
                        for l in tail:
                            f.write(f'    {l[:80]}\n')
                    print(f'[watchdog:{name}] 重启完成 → 新 pid={new_pid}，ALERT 已记录', flush=True)
                    last_size = size   # 重启后日志重新增长，避免立即再报
                else:
                    if now - last_alert_at > 300:   # 每 5 分钟只报一次
                        with open(alert_fp, 'w', encoding='utf-8') as f:
                            f.write(f'🔴 卡死报警 [{name}]（已自动重启 {n_restart}/{max_restart} 次，达到上限需人工）\n' if cmd_fp else f'🔴 卡死报警 [{name}]\n')
                            f.write(f'  日志 {timeout_min} 分钟无更新（最后写入 {time.strftime("%H:%M:%S", time.localtime(mtime))}）\n')
                            f.write(f'  最后进度:\n')
                            for l in tail:
                                f.write(f'    {l[:80]}\n')
                            f.write(f'  建议: 杀 python 进程后重启任务（已完成底稿已落盘安全）\n')
                        last_alert_at = now
                        print(f'[watchdog:{name}] 🔴 卡死! 日志 {idle_min:.0f}min 无更新 → {alert_fp}', flush=True)
        except Exception as ex:
            print(f'[watchdog:{name}] 卡死检测异常: {ex}', flush=True)
        # ④ 进程消失检测
        try:
            if os.path.exists(pid_fp):
                pid = int(open(pid_fp).read().strip() or 0)
                if pid and not proc_alive(pid):
                    if cmd_fp and n_restart < max_restart:
                        print(f'[watchdog:{name}] ⚫ 进程 {pid} 消失 → 自动重启 #{n_restart+1}', flush=True)
                        new_pid = start_cmd(cmd_fp)
                        n_restart += 1
                        with open(pid_fp, 'w') as f:
                            f.write(str(new_pid))
                        with open(alert_fp, 'w', encoding='utf-8') as f:
                            f.write(f'⚫ 进程消失自动重启 [{name}] 第 {n_restart} 次 → 新 pid={new_pid}\n')
                    else:
                        with open(dead_fp, 'w', encoding='utf-8') as f:
                            f.write(f'⚫ 进程消失 [{name}] pid={pid} {time.strftime("%H:%M:%S")}\n')
                        print(f'[watchdog:{name}] ⚫ 进程 {pid} 已消失 → {dead_fp}', flush=True)
                        return 2
        except Exception as ex:
            print(f'[watchdog:{name}] 进程检测异常: {ex}', flush=True)
        time.sleep(heartbeat_s)


if __name__ == '__main__':
    sys.exit(main())
