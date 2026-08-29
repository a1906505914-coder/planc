# -*- coding: utf-8 -*-
"""currency_split.py —— 按币种拆分源数据目录（2026-08-05 用户需求：外币主体分币种处理）。

背景：Z 文件夹（浙江兆龙集团）混有人民币/泰铢/美元三种记账本位币主体：
  · 龙腾互连（泰国）→ THB（泰铢）
  · 龙腾控股（新加坡）→ USD（美元）
  · 母公司/高分子/数链/物联及分公司 → CNY（人民币）
此前各 builder 把所有主体金额当人民币直接相加 → 底稿失真。
方案 A（用户确认）：底稿按本位币分列，各币种组独立生成，组内不与跨币种主体混算；
合并层（折算成人民币）等正确合并 TB 到位后另行启用。

本脚本作用：把 data_dir 下的源数据（科目余额表/综合查询明细表/辅助核算余额表 xlsx）
按主体币种复制到 <data_dir>_split/CNY、_split/THB、_split/USD 三个子目录，
每组单独跑 regen_all 即得到分币种底稿（builder 零改动）。

用法：
  python currency_split.py <源目录> [输出根目录]
  例：python currency_split.py <DATA_ROOT>/Z
      → <DATA_ROOT>/Z/Z_split/{CNY,THB,USD}/...（复制源数据）
  然后：python regen_all.py <DATA_ROOT>/Z/Z_split/CNY <DATA_ROOT>/Z/Z_split/THB <DATA_ROOT>/Z/Z_split/USD
"""
import paths as P
import os
import re
import sys
import io
import shutil

# ⚡ 2026-08-19 修复：作为模块被 import 时（无 stdout.buffer / 已关闭）不再崩溃；
#    仅直接运行时包装 stdout 为 UTF-8 输出
if hasattr(sys.stdout, 'buffer'):
    try:
        # ⚡ 用 reconfigure 原地重配，避免 TextIOWrapper 重包装导致旧对象 GC 关闭 buffer
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (ValueError, OSError, AttributeError):
        pass

from currency_mapping import entity_currency, CURRENCY_LABEL  # noqa: E402

# 复制时跳过的文件（非标准源数据）
_SKIP_KWS = ('外币',)          # 外币科目余额表/外币辅助核算余额表：discover 本就跳过，不复制
_SKIP_EXTS = ('.oxm', '.oxt')  # SAP 原始导出


def _skip(fn):
    low = fn.lower()
    if low.endswith(_SKIP_EXTS):
        return True
    if fn.startswith('~$') or fn.startswith('.'):
        return True
    if any(k in fn for k in _SKIP_KWS):
        return True
    return False


def _entity_of_filename(fn):
    """从文件名提取主体关键词（与 discover_entities 同规则：剥年份/期间/报表关键字）。"""
    base = re.sub(r'(19|20)\d{2}', '', fn)
    base = re.sub(r'年度|年\d*[-~]?\d*月', '', base)
    E = re.sub(r'(外币)?(科目余额表|科目余额|余额表|综合查询明细表|综合查询表|辅助核算余额表|辅助核算|总账|trial).*$',
               '', base, flags=re.IGNORECASE).strip(' _-')
    _m0 = re.match(r'^(\d)([、,，.．\s])', E)
    if _m0 and len(E) > 2:
        E = '0' + E
    return E


def process(src_dir, out_root):
    if not os.path.isdir(src_dir):
        print('⚠️ 源目录不存在：%s' % src_dir)
        return 1
    groups = {'CNY': [], 'THB': [], 'USD': []}
    for fn in sorted(os.listdir(src_dir)):
        fp = os.path.join(src_dir, fn)
        if os.path.isdir(fp):
            continue  # 子目录（如 兆龙TB）不处理
        if not fn.lower().endswith('.xlsx') or _skip(fn):
            continue
        E = _entity_of_filename(fn)
        if not E:
            continue
        cur = entity_currency(E)
        groups.setdefault(cur, []).append(fn)

    if not any(groups.values()):
        print('⚠️ 未找到可拆分的 xlsx 源数据')
        return 1

    for cur, files in groups.items():
        if not files:
            continue
        out_dir = os.path.join(out_root, cur)
        os.makedirs(out_dir, exist_ok=True)
        for fn in files:
            shutil.copy2(os.path.join(src_dir, fn), os.path.join(out_dir, fn))
        print('[%s] %s 个文件 → %s' % (CURRENCY_LABEL.get(cur, cur), len(files), out_dir))
    print('\n✅ 拆分完成：%s' % out_root)
    print('后续：python regen_all.py %s\\CNY %s\\THB %s\\USD' % (out_root, out_root, out_root))
    return 0


def main():
    argv = sys.argv[1:]
    src = argv[0] if argv else P.Z
    out = argv[1] if len(argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(src)), os.path.basename(src) + '_split')
    if not os.path.isdir(src):
        print('⚠️ 源目录不存在：%s' % src)
        return 1
    return process(src, out)


if __name__ == '__main__':
    sys.exit(main())
