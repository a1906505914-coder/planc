# -*- coding: utf-8 -*-
"""运行前自检脚本（2026-08-18，数据安全配套）。

用法：
  python preflight_check.py <账套名或路径> [--out <输出目录>] [--quiet]

检查项：
  1. 数据源完整性   数据源目录/文件存在（科目余额表/序时账等）
  2. 输出目录状态   输出目录存在、可写；目标 xlsx 是否被 Excel 占用（~$ 锁文件 + O_RDWR 探测）
  3. 缓存状态       .gl_cache / .cache 是否存在、大小、最近修改（异常大=提示可清）
  4. 磁盘空间       DATA_ROOT 所在盘剩余空间（<2GB 阻断，<10GB 警告）
  5. 并发检测       是否有其他 python 进程正在处理同一账套（run_u8_on_sap/小程序目录）
  6. 上次运行日志   账套目录下 *.ALERT / *.DEAD 标记（看门狗报警残留）
  7. 硬编码自检     调用 check_hardcode.py（数据绝对路径是否混入程序）

返回码：0=可运行 | 1=有警告（可继续） | 2=阻断（必须先处理）。
"""
import glob
import os
import re
import subprocess
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import paths as P

RES = []          # (level, msg)  level: OK/WARN/BLOCK/INFO
FAIL = 0          # 阻断数
WARN = 0          # 警告数


def _add(level, msg):
    RES.append((level, msg))
    globals()['FAIL'] += 1 if level == 'BLOCK' else 0
    globals()['WARN'] += 1 if level == 'WARN' else 0


# ---------------------------------------------------------------- 工具
def is_locked(path):
    """文件是否被其他进程独占（Excel/WPS 预览占用探测）。"""
    if not os.path.exists(path):
        return False
    try:
        fd = os.open(path, os.O_RDWR)
        os.close(fd)
        return False
    except PermissionError:
        return True
    except OSError:
        return False


def dir_size_mb(path):
    try:
        total = 0
        for dp, _, fns in os.walk(path):
            for fn in fns:
                try:
                    total += os.path.getsize(os.path.join(dp, fn))
                except OSError:
                    pass
        return total / 1048576
    except Exception:
        return -1


def running_related_procs(data_root):
    """检测正在处理同一账套/小程序的 python 进程（排除自身）。"""
    hits = []
    try:
        out = subprocess.run(
            ['wmic', 'process', 'where', "name='python.exe'", 'get', 'processid,commandline'],
            capture_output=True, text=True, timeout=20)
        for line in (out.stdout or '').splitlines():
            if 'python' not in line.lower():
                continue
            if str(os.getpid()) in line:
                continue          # 自身
            if APP_DIR in line or data_root in line or 'run_u8_on_sap' in line:
                hits.append(line.strip()[:120])
    except Exception:
        pass
    return hits


# ---------------------------------------------------------------- 检查项
def check_sources(data_root):
    """数据源完整性：账套根存在；常见子目录存在。"""
    if not os.path.isdir(data_root):
        _add('BLOCK', f'数据根目录不存在：{data_root}')
        return
    _add('OK', f'数据根目录存在：{data_root}')
    # 常见数据子目录/文件
    for name in ('科目余额表', '序时账', '底稿'):
        p = os.path.join(data_root, name)
        if os.path.isdir(p):
            n = len(os.listdir(p))
            _add('OK', f'  {name}/ 存在（{n} 项）')
        else:
            _add('WARN', f'  {name}/ 目录缺失（{p}）')
    # 三件套数据文件（prepared 集团模式）
    prep = os.path.join(data_root, 'prepared')
    if os.path.isdir(prep):
        _add('OK', f'  prepared/ 存在（{len(os.listdir(prep))} 项）')


def check_output(out_dir):
    if not out_dir:
        return
    if not os.path.isdir(out_dir):
        _add('WARN', f'输出目录不存在（将自动创建）：{out_dir}')
        return
    _add('OK', f'输出目录存在：{out_dir}')
    # 可写性（写成功即可写；删除被 safe-delete 拦截属正常保护，不算不可写）
    probe = os.path.join(out_dir, '.preflight_probe')
    try:
        with open(probe, 'w') as f:
            f.write('ok')
        try:
            os.remove(probe)
        except (PermissionError, OSError):
            pass   # safe-delete 拦截删除 → 忽略（保护机制）
        _add('OK', '  输出目录可写')
    except (PermissionError, OSError) as e:
        _add('BLOCK', f'  输出目录不可写：{e}')
    # Excel 锁文件
    locks = glob.glob(os.path.join(out_dir, '~$*.xlsx')) + glob.glob(os.path.join(out_dir, '**', '~$*.xlsx'), recursive=True)
    if locks:
        _add('BLOCK', f'  发现 Excel 锁文件（文件正被打开，重跑会被占用拦截）：')
        for l in locks[:5]:
            _add('BLOCK', f'    {os.path.basename(l)}')
    # 已有 xlsx 占用探测（抽查前 20 个）
    locked = []
    xls = sorted(glob.glob(os.path.join(out_dir, '*.xlsx')))[:20]
    for x in xls:
        if is_locked(x):
            locked.append(os.path.basename(x))
    if locked:
        _add('WARN', f'  已有 {len(locked)} 个 xlsx 被占用（重跑可能失败）：{locked[:3]}…')
    else:
        _add('OK', '  无 xlsx 被占用')


def check_cache(data_root):
    """缓存状态。"""
    for name, label in (('.gl_cache', '账套 GL 缓存'), ('.cache', '程序解析缓存')):
        p = os.path.join(data_root, name)
        if not os.path.exists(p):
            p = os.path.join(APP_DIR, name)
        if os.path.isdir(p):
            mb = dir_size_mb(p)
            _add('OK', f'{label} {name}/ 存在（{mb:.1f}MB）' if mb >= 0 else f'{label} {name}/ 存在')
        else:
            _add('INFO', f'{label} {name}/ 不存在（正常，首次运行自动建立）')


def check_disk():
    """数据盘剩余空间（shutil.disk_usage，跨平台可靠）。"""
    try:
        import shutil
        drive = os.path.splitdrive(P.DATA_ROOT)[0] or P.DATA_ROOT
        if not drive.endswith('\\'):
            drive += '\\'
        du = shutil.disk_usage(drive)
        free = du.free / (1024 ** 3)
        if free < 2:
            _add('BLOCK', f'磁盘剩余空间不足 2GB（当前 {free:.1f}GB），重跑大账套可能失败')
        elif free < 10:
            _add('WARN', f'磁盘剩余空间低于 10GB（当前 {free:.1f}GB）')
        else:
            _add('OK', f'磁盘剩余空间充足（{free:.1f}GB）')
    except Exception as e:
        _add('WARN', f'磁盘空间检测异常：{e}')


def check_concurrency(data_root):
    hits = running_related_procs(data_root)
    if hits:
        _add('WARN', f'检测到 {len(hits)} 个相关 python 进程在运行（并发处理同一账套可能写冲突）：')
        for h in hits[:3]:
            _add('WARN', f'    PID 命令行…{h[-70:]}')
    else:
        _add('OK', '无同账套并发进程')


def check_prior_logs(data_root):
    """上次运行看门狗报警残留。"""
    marks = glob.glob(os.path.join(data_root, '*.ALERT')) + glob.glob(os.path.join(data_root, '*.DEAD'))
    if marks:
        _add('WARN', f'上次运行有报警残留（看门狗 ALERT/DEAD）：')
        for m in marks[:5]:
            _add('WARN', f'    {os.path.basename(m)}')
    else:
        _add('OK', '无上次运行报警残留')


def check_hardcode():
    try:
        r = subprocess.run([sys.executable, os.path.join(APP_DIR, 'check_hardcode.py')],
                           capture_output=True, text=True, timeout=60)
        out = (r.stdout or '').strip()
        if '无数据路径硬编码' in out or 'OK' in out.upper():
            _add('OK', '硬编码自检通过（无数据绝对路径混入程序）')
        else:
            _add('WARN', '硬编码自检异常：' + out[:100])
    except Exception as e:
        _add('WARN', f'硬编码自检调用失败：{e}')


# ---------------------------------------------------------------- main
def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    quiet = '--quiet' in argv
    if '--quiet' in argv:
        argv.remove('--quiet')
    out_dir = None
    if '--out' in argv:
        i = argv.index('--out')
        out_dir = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]

    target = argv[0] if argv else None
    if target and os.path.isabs(target):
        data_root = os.path.abspath(target)
        acct = os.path.basename(data_root)
    elif target:
        data_root = os.path.join(P.DATA_ROOT, target)
        acct = target
    else:
        data_root = P.DATA_ROOT
        acct = '（全部账套）'

    print(f'═════ 运行前自检：{acct} ═════')
    check_sources(data_root)
    check_output(out_dir or os.path.join(data_root, '底稿'))
    check_cache(data_root)
    check_disk()
    check_concurrency(data_root)
    check_prior_logs(data_root)
    if '--skip-hardcode' not in argv:
        check_hardcode()

    print()
    print('── 汇总 ──')
    for lv, msg in RES:
        tag = {'OK': '✅', 'WARN': '⚠️', 'BLOCK': '⛔', 'INFO': 'ℹ️'}.get(lv, '?')
        print(f'  {tag} [{lv}] {msg}')
    print()
    if FAIL:
        print(f'结果：{FAIL} 项阻断 → 必须先处理后运行（返回码 2）')
        return 2
    if WARN:
        print(f'结果：{WARN} 项警告 → 可运行，但建议确认（返回码 1）')
        return 1
    print('结果：全部通过，可以安全运行（返回码 0）')
    return 0


if __name__ == '__main__':
    sys.exit(main())
