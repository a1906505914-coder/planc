# -*- coding: utf-8 -*-
"""底稿归档（#未来设计·正式底稿版本归档）。

正式底稿是交付物，代码有 git 但底稿文件没有版本管理。本脚本把正式底稿
按"生成批次"归档（复制到 _archive/<时间戳>/ 或打 zip），回退时可取回某批次的完整底稿。

用法：
    python archive_wp.py                    # 归档 AH 三集团正式底稿（复制到 _archive/时间戳/）
    python archive_wp.py --zip              # 打包为 zip（省空间，单文件好管理）
    python archive_wp.py --data <数据包目录>  # 归档任意数据包的底稿目录
"""
import os
import sys
import shutil
import datetime
import zipfile
import paths as P

AH_WP = os.path.join(P.DATA_DIRS['AH'], '底稿', '2026', '集团')
AH_ARCHIVE = os.path.join(P.DATA_DIRS['AH'], '底稿', '2026', '_archive')


def _ts():
    return datetime.datetime.now().strftime('%Y%m%d_%H%M%S')


def archive(data_root, use_zip=False):
    if not os.path.isdir(data_root):
        print(f'目录不存在：{data_root}')
        return 1
    ts = _ts()
    n_files = 0
    total_size = 0
    if use_zip:
        zpath = os.path.join(os.path.dirname(data_root.rstrip('/\\')), f'_archive_{ts}.zip')
        os.makedirs(os.path.dirname(zpath), exist_ok=True)
        with zipfile.ZipFile(zpath, 'w', zipfile.ZIP_DEFLATED) as zf:
            for base, _dirs, files in os.walk(data_root):
                for fn in files:
                    if not fn.endswith('.xlsx') or fn.startswith('~$'):
                        continue
                    fp = os.path.join(base, fn)
                    arc = os.path.relpath(fp, os.path.dirname(data_root.rstrip('/\\')))
                    zf.write(fp, arc)
                    n_files += 1
                    total_size += os.path.getsize(fp)
        print(f'✅ 已归档为 zip：{zpath}')
    else:
        dest = os.path.join(AH_ARCHIVE if 'AH' in data_root or '集团' in data_root
                            else os.path.dirname(data_root.rstrip('/\\')), f'_archive_{ts}')
        os.makedirs(dest, exist_ok=True)
        for base, _dirs, files in os.walk(data_root):
            for fn in files:
                if not fn.endswith('.xlsx') or fn.startswith('~$'):
                    continue
                fp = os.path.join(base, fn)
                rel = os.path.relpath(fp, data_root)
                dfp = os.path.join(dest, rel)
                os.makedirs(os.path.dirname(dfp), exist_ok=True)
                shutil.copy2(fp, dfp)
                n_files += 1
                total_size += os.path.getsize(fp)
        print(f'✅ 已归档 {n_files} 个底稿文件（{total_size/1024/1024:.1f} MB）→ {dest}')
    print(f'  批次 {ts}；回退时直接取回该批次文件即可。')
    return 0


def main():
    argv = sys.argv[1:]
    use_zip = '--zip' in argv
    data_root = None
    if '--data' in argv:
        i = argv.index('--data')
        data_root = argv[i + 1]
    if data_root is None:
        # ⚡⚡ 2026-08-31 参数化：--acct/--year 走 tool_common（替代默认 AH_WP）
        import argparse, tool_common as T
        _ap = argparse.ArgumentParser(add_help=False)
        T.add_tool_args(_ap, default_acct='AH')
        _args, _ = _ap.parse_known_args()
        DATA, YEAR = T.resolve_data(_args)
        data_root = os.path.join(os.path.dirname(DATA), '底稿', YEAR, '集团')
    sys.exit(archive(data_root, use_zip=use_zip))


if __name__ == '__main__':
    main()
