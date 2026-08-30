# -*- coding: utf-8 -*-
"""看门狗：检测长任务（AH 全量等）是否卡死。

原理：run_u8_on_sap 每科目开始/完成都会写心跳文件 `.heartbeat_{comp}`
（含时间戳 + 当前科目）。本脚本扫描心跳，若 mtime 停滞超过阈值 → 判定疑似
卡死/中断 → 打印告警 + 写 `.ALERT_{comp}` 标记（preflight_check 会读该标记，
下次运行前会提示先处理）。

用法：
    python watchdog_check.py                 # 默认扫 AH 数据目录
    python watchdog_check.py <数据目录>      # 指定目录
    python watchdog_check.py <目录> --max-age 1800   # 阈值 30 分钟（默认）
    python watchdog_check.py --clean          # 清理 ALERT/DEAD 标记

返回码：0=正常 | 1=有卡死告警（或残留标记）。
"""
import glob
import os
import sys
import time
import paths as P

DEFAULT_MAX_AGE = 1800        # 30 分钟（current_account 大表单科目可能较长）
DEFAULT_DATA = os.path.join(P.DATA_DIRS['AH'], '数据', '2026')


def scan(data_root, max_age):
    alerts = []
    marks = []
    for fp in sorted(glob.glob(os.path.join(data_root, '.heartbeat_*'))):
        comp = os.path.basename(fp).replace('.heartbeat_', '')
        try:
            st = os.stat(fp)
            age = time.time() - st.st_mtime
            with open(fp, 'r', encoding='utf-8') as f:
                content = f.read().strip()
        except Exception:
            continue
        if age > max_age:
            tag = os.path.join(data_root, f'.ALERT_{comp}')
            try:
                with open(tag, 'w', encoding='utf-8') as f:
                    f.write(f'heartbeat {comp} 停滞 {age:.0f}s > {max_age}s @ {time.strftime("%Y-%m-%d %H:%M:%S")}\n')
            except Exception:
                pass
            alerts.append((comp, age, content))
            marks.append(tag)
        else:
            print(f'  ✅ {comp}: 心跳正常（{age:.0f}s 前，{content}）')
    return alerts, marks


def clean(data_root):
    n = 0
    for pat in ('.ALERT_*', '.DEAD_*'):
        for fp in glob.glob(os.path.join(data_root, pat)):
            try:
                os.remove(fp)
                n += 1
            except Exception:
                pass
    print(f'已清理 {n} 个 ALERT/DEAD 标记')
    return 0


def main():
    argv = sys.argv[1:]
    if '--clean' in argv:
        return clean(DEFAULT_DATA)
    data_root = DEFAULT_DATA
    if argv and not argv[0].startswith('-') and os.path.isdir(argv[0]):
        data_root = os.path.abspath(argv[0])
    max_age = DEFAULT_MAX_AGE
    if '--max-age' in argv:
        i = argv.index('--max-age')
        try:
            max_age = int(argv[i + 1])
        except Exception:
            pass
    print(f'═════ 看门狗：{data_root} ═════')
    alerts, marks = scan(data_root, max_age)
    # 顺带检查残留标记（上次中断遗留）
    leftover = []
    for pat in ('.ALERT_*', '.DEAD_*'):
        for fp in glob.glob(os.path.join(data_root, pat)):
            if fp not in marks:
                leftover.append(fp)
    if leftover:
        print('⚠️ 存在上次运行残留标记：')
        for fp in leftover:
            print(f'  - {os.path.basename(fp)}')
    if alerts:
        print()
        for comp, age, content in alerts:
            print(f'  ⛔ {comp}: 心跳停滞 {age:.0f}s（{content}）→ 已写 .ALERT_{comp}')
        print(f'\n结果：{len(alerts)} 项疑似卡死/中断 → 检查 run_history 最近报告确认（返回码 1）')
        return 1
    print(f'\n结果：无卡死告警（{len(glob.glob(os.path.join(data_root, ".heartbeat_*")))} 个心跳全部正常）')
    return 0


if __name__ == '__main__':
    sys.exit(main())
