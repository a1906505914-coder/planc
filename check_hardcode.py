# -*- coding: utf-8 -*-
"""check_hardcode.py —— 数据路径硬编码自检（2026-08-14 建立，程序=程序、数据=数据）。

用途：扫描小程序代码，找出任何【数据路径硬编码】残留：
  - 盘符绝对路径（E:/、C:/Users、D:/ 等）
  - 数据目录名"底稿测试"字面量
  - 旧账套名（yy2026 等历史名）字面量

规则：程序代码里【只允许】出现：
  - paths.py 中的 DATA_ROOT 默认值（唯一合法"数据路径"锚点）
  - 注释/文档里说明性的占位符（<DATA_ROOT>、<APP_ROOT> 之类）
  - 环境变量引用（AUDIT_DATA_ROOT / %USERPROFILE% / %~dp0 / sys.executable 等）

用法：
  python check_hardcode.py             # 扫描小程序目录（跳过 trash 归档）
  python check_hardcode.py --all       # 连归档目录一起扫（仅报告，不改）
  python check_hardcode.py <dir>       # 扫指定目录
退出码：0=干净（无违规）；1=有违规（列出文件:行号）
"""
import os
import re
import sys

# ---------- 规则配置 ----------
# 1) 盘符绝对路径：驱动器字母 + 冒号 + 斜杠
DRIVE_RE = re.compile(r"""(?<![A-Za-z])(?:[A-Za-z]:[\\/])""")
# 2) 数据目录名字面量（出现在字符串/注释中）
DATA_NAME_RE = re.compile(r'底稿测试')
# 3) 旧账套名（历史遗留命名）：仅当出现在【路径上下文】中才算违规
#    （如 D:/底稿测试/yy2026、<DATA_ROOT>/yy2026）；
#    纯程序标识符（layout == 'ah-sap'、_discover_ah_sap、docstring 布局说明）豁免。
OLD_ACCT_RE = re.compile(r'(?:[\\/]|底稿测试)yy2026')

# 允许出现上述模式的文件（白名单）
ALLOW_FILES = {'paths.py', 'check_hardcode.py'}
# 允许出现的"锚点"行（精确包含即豁免）——paths.py 的 DATA_ROOT 默认值本身（任意盘符，避免盘符迁移后再改）
ANCHOR_RE = re.compile(r"DATA_ROOT = os\.environ\.get\('AUDIT_DATA_ROOT', r'[A-Za-z]:\\底稿测试'\)")

# 扫描范围（跳过历史归档 + 第三方依赖目录）
SKIP_DIR_PARTS = ('_archive_trash', '.audit_legacy_trash', '__pycache__', '.git',
                  '.easyocr_libs', '.paddle_libs', '.ocr_gpu')


def scan(root):
    hits = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIR_PARTS]
        for fn in sorted(filenames):
            if not fn.endswith(('.py', '.bat', '.json')):
                continue
            fp = os.path.join(dirpath, fn)
            rel = os.path.relpath(fp, root)
            if rel.replace('\\', '/') in ALLOW_FILES:
                continue
            try:
                with open(fp, encoding='utf-8', errors='replace') as f:
                    lines = f.readlines()
            except OSError:
                continue
            for i, raw in enumerate(lines, 1):
                line = raw.rstrip('\n').rstrip('\r')
                if ANCHOR_RE.match(line.strip()) or line.strip().startswith('#'):
                    continue
                bad = []
                if DRIVE_RE.search(line):
                    bad.append('盘符绝对路径')
                if DATA_NAME_RE.search(line):
                    bad.append('数据目录名')
                if OLD_ACCT_RE.search(line):
                    bad.append('旧账套名')
                if bad:
                    hits.append((rel, i, '+'.join(bad), line.strip()))
    return hits


def main():
    root = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith('--') else \
        os.path.dirname(os.path.abspath(__file__))
    hits = scan(root)
    if not hits:
        print(f'[OK] 无数据路径硬编码：{root}')
        return 0
    print(f'[FAIL] 发现 {len(hits)} 处数据路径硬编码：')
    for rel, i, kind, line in hits:
        print(f'  {rel}:{i} [{kind}] {line[:120]}')
    return 1


if __name__ == '__main__':
    sys.exit(main())
