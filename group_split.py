# -*- coding: utf-8 -*-
"""group_split.py —— Z 归并组拆分（2026-08-05 用户确认）。

把 Z_split/CNY 的源数据按【归并主体】复制到 Z_groups/{母公司,物联,高分子,数链}：
  · 母公司组 = 母公司 + 上分 + 北分
  · 物联组   = 物联(本部) + 物联合肥 + 物联深圳 + 物联苏州
  · 高分子组 = 高分子
  · 数链组   = 数链
每组目录含各主体的科目余额表/综合查询明细表/辅助核算余额表。
组内跑 13 builder → 底稿每主体一列 + 合计列（集团合计），满足"分公司列示+合计、与单体TB核对"。

用法：python group_split.py <Z_split/CNY目录> [输出根目录]
"""
import paths as P
import os
import sys
import io
import shutil

# 归并组 → 主体文件名前缀（顺序无冲突：'物联2025'=本部，'物联合肥2025'=分公司）
GROUP_PREFIX = {
    '母公司': ['母公司2025', '上分2025', '北分2025'],
    '物联': ['物联2025', '物联合肥2025', '物联深圳2025', '物联苏州2025'],
    '高分子': ['高分子2025'],
    '数链': ['数链2025'],
}

# 非源数据文件（勿复制）
SKIP_KW = ['关联方清单', '重要性水平', '模板']


def split(src_dir, out_root):
    if not os.path.isdir(src_dir):
        print('❌ 源目录不存在: %s' % src_dir)
        return 1
    os.makedirs(out_root, exist_ok=True)
    files = sorted(f for f in os.listdir(src_dir)
                   if f.lower().endswith('.xlsx') and not any(k in f for k in SKIP_KW))
    print('源目录: %s（%d 个 xlsx）' % (src_dir, len(files)))
    total = 0
    for group, prefixes in GROUP_PREFIX.items():
        gdir = os.path.join(out_root, group)
        os.makedirs(gdir, exist_ok=True)
        n = 0
        for f in files:
            if any(f.startswith(p) for p in prefixes):
                shutil.copy2(os.path.join(src_dir, f), os.path.join(gdir, f))
                n += 1
        total += n
        print('  ✅ %s: %d 个文件 → %s' % (group, n, gdir))
    print('\n✅ 拆分完成：%d 个文件（%s）' % (total, out_root))
    return 0


def main():
    argv = sys.argv[1:]
    src = argv[0] if argv else os.path.join(P.Z, 'Z_split', 'CNY')
    out = argv[1] if len(argv) > 1 else os.path.join(P.DATA_ROOT, 'Z_groups')
    return split(src, out)


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.exit(main())
