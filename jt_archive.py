# -*- coding: utf-8 -*-
"""JTt 底稿归档（2026-08-05）：生成完成后把底稿从 prepared（中间结果）归位到独立目录。
用法：python jt_archive.py <JTt目录> [--dry]
目录规划：
  JTt/prepared/{ga,jj}   = 拆分后的标准中间数据（科目余额表/综合查询明细表/辅助核算余额表）——只读数据源
  JTt/底稿/{ga,jj}       = 审计底稿输出目录（*_生成.xlsx）
生成流程：jtt_prepare.py 拆分 → regen_all.py JTt/prepared/{ga,jj} → 本脚本归档底稿
"""
import paths as P
import os, sys, io, shutil, glob
# ⚡ 2026-08-19 修复：作为模块被 import 时（无 stdout.buffer / 已关闭）不再崩溃；
#    仅直接运行时包装 stdout 为 UTF-8 输出
if hasattr(sys.stdout, 'buffer'):
    try:
        # ⚡ 用 reconfigure 原地重配，避免 TextIOWrapper 重包装导致旧对象 GC 关闭 buffer
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (ValueError, OSError, AttributeError):
        pass

def main():
    argv = sys.argv[1:]
    root = argv[0] if argv else P.JTT
    dry = '--dry' in argv
    total = 0
    for tag in ('ga', 'jj'):
        src_dir = os.path.join(root, 'prepared', tag)
        dst_dir = os.path.join(root, '底稿', tag)
        if not os.path.isdir(src_dir):
            print(f'⚠️ {src_dir} 不存在')
            continue
        os.makedirs(dst_dir, exist_ok=True)
        moved = 0
        for p in sorted(glob.glob(os.path.join(src_dir, '*_生成.xlsx'))):
            fn = os.path.basename(p)
            dp = os.path.join(dst_dir, fn)
            if os.path.exists(dp):
                # 覆盖已存在的旧底稿（同文件名为新生成）
                os.remove(dp)
            if dry:
                print(f'  [dry] {tag}/{fn} → 底稿/{tag}/')
            else:
                shutil.move(p, dp)
            moved += 1
        total += moved
        print(f'{tag}: 归档 {moved} 份底稿 → {dst_dir}')
        # 清理残留空目录（若有）
        if not dry and not glob.glob(os.path.join(src_dir, '*')) :
            os.rmdir(src_dir)
    print(f'{"[dry] " if dry else ""}归档完成，共 {total} 份。' if total else '无底稿可归档')

if __name__ == '__main__':
    main()
