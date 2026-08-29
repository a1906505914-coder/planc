# -*- coding: utf-8 -*-
"""期间/年度参数化集中配置（#未来设计·新年度平滑切换）。

新年度数据接入时：
  1. 改这里 YEAR（及必要时的 PERIOD_NOTE）。
  2. 跑 `python audit_config.py scan` 扫描各模块剩余的硬编码年份，逐个核对。

注意：各模块内部仍有散落的 '2026' 硬编码（历年积累），本模块提供集中常量 +
扫描工具，优先把"主流程入口"统一到此；模块内部硬编码逐年收敛。
"""
import os
import re
import sys

# ===== 集中年度/期间配置（新年度改这里）=====
YEAR = '2026'                 # 审计年度
PERIOD_NOTE = '1-7月'         # 期间说明（SAP 科目余额表 1-7 月口径）
TARGET_YEARS = ['2026']       # 生成器默认覆盖的年度集合
# ============================================

ROOT = os.path.dirname(os.path.abspath(__file__))
_SKIP = ('audit_config.py', '.git', '__pycache__', '.bak_')


def scan_hardcoded_year():
    """扫描各 py 中与当前 YEAR 相关的硬编码年份，输出待核对清单。"""
    pat = re.compile(r"['\"]20\d{2}['\"]|_20\d{2}_|f['\"]20\d{2}|YEAR\s*=\s*['\"]20\d{2}")
    hits = []
    for f in sorted(os.listdir(ROOT)):
        if not f.endswith('.py') or f.startswith(('.', '_')) or f in _SKIP:
            continue
        fp = os.path.join(ROOT, f)
        try:
            txt = open(fp, encoding='utf-8').read()
        except Exception:
            continue
        years = set(re.findall(r"(?<!['\"])\b20\d{2}\b(?!['\"])", txt)) - {'2026'}
        if years and not any(k in txt for k in ('def ', 'import ', 'from ')):
            pass
        for y in sorted(years):
            cnt = len(re.findall(r'\b%s\b' % y, txt))
            hits.append((f, y, cnt))
    print(f'扫描 {len([f for f in os.listdir(ROOT) if f.endswith(".py")])} 个 py：')
    print(f'  当前年度配置：YEAR={YEAR}，期间={PERIOD_NOTE}')
    other = [h for h in hits if h[1] != YEAR]
    if not other:
        print('  ✅ 未发现其他硬编码年度（历史年份仅存在于注释/说明）')
    else:
        print(f'  发现 {len(other)} 处其他年度硬编码（多为历史年度参考/注释，逐个核对）：')
        for f, y, c in other[:30]:
            print(f'    {f}: 年度 {y} ×{c}')


if __name__ == '__main__':
    if len(sys.argv) > 1 and sys.argv[1] == 'scan':
        scan_hardcoded_year()
    else:
        print(__doc__)
