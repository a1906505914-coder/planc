# -*- coding: utf-8 -*-
"""全量备份（#单机无冗余·防丢失核心）。

WorkBuddy 代码/数据/底稿全在本地单机，无协作账号冗余。本脚本把三类资产
一键备份到独立位置（建议移动硬盘/第二磁盘/网盘目录）：
  ① 代码：git bundle（单文件、含全部历史，可随时完整还原仓库）
  ② 原始数据：各数据包 数据/ 目录（企业提供 xlsx，不可再生）
  ③ 底稿：各数据包 底稿/ 目录（含人工填写列，不可完全再生成）

用法：
    python backup_all.py                          # 备份到 D:/审计备份_<日期>/
    python backup_all.py --out E:/审计备份        # 指定备份根（建议移动硬盘）
    python backup_all.py --only code              # 只备份代码（快，常跑）
    python backup_all.py --only data              # 只备份原始数据
"""
import paths as P
import os
import sys
import time
import datetime
import subprocess
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
CODE_DIR = BASE
DATA_ROOTS = {
    'AH': P.DATA_DIRS['AH'],
    'XBJ': P.DATA_DIRS['XBJ'],
    'ADF': P.DATA_DIRS['ADF'],
}
DEFAULT_OUT = os.environ.get('AUDIT_BACKUP_ROOT', r'D:/审计备份')

_EXCLUDE_DATA = ('_work', '_archive', '.cache', '__pycache__', '.git', '_conv_before_备份')
_EXCLUDE_WP = ('_archive', '.cache', '_work')


def _ts():
    return datetime.datetime.now().strftime('%Y%m%d_%H%M%S')


def backup_code(dest_dir):
    """git bundle 备份代码仓库（单文件、含全部历史）。"""
    bundle = os.path.join(dest_dir, f'audit_code_{_ts()}.bundle')
    try:
        subprocess.run(['git', 'bundle', 'create', bundle, '--all'],
                       cwd=CODE_DIR, check=True, capture_output=True)
    except Exception as e:
        print(f'  ⚠️ git bundle 失败（无 git 历史？）：{e}')
        return 0
    size = os.path.getsize(bundle) / 1024 / 1024
    print(f'  ✅ 代码(git bundle) {size:.1f} MB → {os.path.basename(bundle)}')
    return 1


def _zip_dir(zf, root, name, excludes, arc_root):
    n = 0
    for base, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d not in excludes and not d.endswith('.bak_')]
        for fn in files:
            if fn.startswith('~$') or fn.endswith(('.pyc', '.log')):
                continue
            fp = os.path.join(base, fn)
            arc = os.path.join(arc_root, os.path.relpath(fp, root))
            try:
                zf.write(fp, arc)
                n += 1
            except Exception:
                pass
    return n


def backup_data(dest_dir, roots=DATA_ROOTS, only=None):
    """压缩各数据包 数据/ + 底稿/（可配置 --only data 只数据）。"""
    if only == 'data':
        roots = {k: v for k, v in roots.items() if k}
    made = []
    for pkg, root in roots.items():
        data_dir = os.path.join(root, '数据')
        wp_dir = os.path.join(root, '底稿')
        if not os.path.isdir(data_dir):
            print(f'  – {pkg} 无 数据/ 目录，跳过')
            continue
        zpath = os.path.join(dest_dir, f'{pkg}_data_wp_{_ts()}.zip')
        t0 = time.time()
        with zipfile.ZipFile(zpath, 'w', zipfile.ZIP_DEFLATED, allowZip64=True) as zf:
            n1 = _zip_dir(zf, data_dir, 'data', _EXCLUDE_DATA, pkg)
            n2 = 0
            if os.path.isdir(wp_dir):
                n2 = _zip_dir(zf, wp_dir, 'wp', _EXCLUDE_WP, pkg)
        size = os.path.getsize(zpath) / 1024 / 1024
        made.append(zpath)
        print(f'  ✅ {pkg} 数据+底稿 {size:.0f} MB（{n1 + n2} 文件，{time.time()-t0:.0f}s）→ {os.path.basename(zpath)}')
    return made


def main():
    argv = sys.argv[1:]
    only = None
    if '--only' in argv:
        i = argv.index('--only')
        only = argv[i + 1]
    out_root = DEFAULT_OUT
    if '--out' in argv:
        i = argv.index('--out')
        out_root = argv[i + 1]
    os.makedirs(out_root, exist_ok=True)
    t0 = time.time()
    print(f'备份开始 → {out_root}（{_ts()}）')
    if only in (None, 'code'):
        backup_code(out_root)
    if only in (None, 'data'):
        backup_data(out_root, only=only)
    print(f'\n备份完成（耗时 {time.time()-t0:.0f}s）。请确认备份目录存在且可打开；建议定期执行。')


if __name__ == '__main__':
    main()
