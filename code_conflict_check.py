# -*- coding: utf-8 -*-
"""科目代码归属冲突检测（2026-08-24 用户要求：防『修了又犯』——代码配置盲区类问题提前暴露）。

背景：科目代码在不同账套语义不同（2401 在 AS=递延收益、别处可能=一年内到期；3100 在
泰国=生产成本、中国=专项储备）。生成器靠『codes 配置 + 名称兜底』归属科目组，配置表若
同一代码被多个科目组配置（2401∈ncl+di、2601∈ll+di、2703∈ll+ncl…）→ 生成时可能误并。

本工具扫描所有生成器的 codes 配置，找出【同代码被多科目组配置】的冲突，输出清单：
  - 提示可能误并的科目组对
  - 供人工确认归属（确认某账套该代码语义后再决定去留）
纳入架构：regen_all 阶段②预检自动调用（冲突非致命，仅告警）。

用法：
  python code_conflict_check.py            # 扫描配置冲突（默认）
"""
import os
import sys
from collections import defaultdict


def scan_all_generators():
    """扫描各生成器的 codes 配置，返回 {code: [owner1, owner2...]}。"""
    code_owner = defaultdict(list)

    def _reg(module, name):
        try:
            __import__(module)
            m = sys.modules[module]
        except Exception:
            return
        # gp_other.SUBJECTS / equity_detail.GROUPS / inventory_detail.INV_LEVEL2
        for attr in ('SUBJECTS',):
            for s in getattr(m, attr, []) or []:
                for cd in (s.get('codes') or []):
                    code_owner[str(cd)].append('%s:%s' % (name, s.get('name') or '?'))
        for attr in ('GROUPS',):
            for k, g in (getattr(m, attr, {}) or {}).items():
                for cd in (g.get('codes') or []):
                    code_owner[str(cd)].append('%s:%s' % (name, g.get('title') or k))
        for attr in ('INV_LEVEL2',):
            for cd, nm in (getattr(m, attr, {}) or {}).items():
                code_owner[str(cd)].append('%s:%s' % (name, nm))

    _reg('gp_other', 'gp')
    _reg('equity_detail', 'eq')
    _reg('inventory_detail', 'inv')
    _reg('longterm_assets_detail', 'lt')
    _reg('tax_detail', 'tax')
    return code_owner


def find_conflicts(code_owner):
    """返回 [(code, [owner...])] 仅含被多科目组配置的。"""
    out = []
    for cd in sorted(code_owner):
        owners = sorted(set(code_owner[cd]))
        if len(owners) > 1:
            out.append((cd, owners))
    return out


def main():
    code_owner = scan_all_generators()
    conflicts = find_conflicts(code_owner)
    print(f'=== 科目代码归属冲突检测 ===')
    if not conflicts:
        print('  ✅ 无同代码被多科目组配置（配置干净）')
        return 0
    print(f'  ⚠️ {len(conflicts)} 处代码被多科目组配置（可能误并，需确认归属）：')
    for cd, owners in conflicts:
        print(f'    {cd}: ' + ' 与  '.join(owners))
    return 1


if __name__ == '__main__':
    sys.exit(main())
