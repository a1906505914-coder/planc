# -*- coding: utf-8 -*-
"""tb_validate.py —— 阶段三 3.2 摄取层 TB 控制数 Fail Fast（2026-08-11）

治本目标：数据问题在【摄取时】拦截（显式告警/Fail Fast），不静默生成坏底稿。
现在的问题：数据缺/表表对不上/勾稽差 → 生成完才发现（recon_self/audit_checker 事后检查）。

机制（复用 self_tb_gen.balance_check 的叶子层平衡口径）：
  1) 读原始 TB（audit_common.read_tb_full，与各 builder 同源）
  2) 每主体×年度：
     · 发生额平衡：Σ借发 == Σ贷发（差=未结转损益可容忍，pl/cost 类科目期末余额）
     · 余额平衡：Σ期末借余 == Σ期末贷余（资产=负债+权益）
  3) 输出 (ok, issues)：issue = (级别, 主体, 年度, 说明)
     · 级别 'FAIL'（硬性问题：发生额差≠未结转损益 或 期末差显著）→ fail_fast 时中止
     · 级别 'WARN'（未结转损益等口径提示）
  4) 挂 regen_all 阶段② builder 前：FAIL → 明确报错并跳过该账套（不生成坏底稿）；
     WARN → 提示不阻断。

用法：
  from tb_validate import validate_tb
  ok, issues = validate_tb(data_dir)            # 只校验不阻断
  ok, issues = validate_tb(data_dir, fail_fast=True)  # FAIL 即抛 TBValidationError
"""
from __future__ import annotations
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


class TBValidationError(Exception):
    """TB 控制数校验未通过（Fail Fast 中止）。"""
    pass


def validate_tb(data_dir, fail_fast=False, quiet=False, hard_threshold=100.0):
    """摄取层 TB 控制数校验。返回 (ok, issues)。
    issues = [(级别 'FAIL'|'WARN', 主体, 年度, 说明, 差)]。
    hard_threshold：发生额差/期末差超过该值且非未结转损益 → FAIL。"""
    import audit_common
    import self_tb_gen as S

    if not os.path.isdir(data_dir):
        return True, []
    entities = audit_common.discover_entities(data_dir)
    tb_full = audit_common.read_tb_full(data_dir, entities)
    if not tb_full:
        return True, []
    years = sorted({str(yy) for (_e, _c, _n, yy) in tb_full}) or ['2025']
    issues = []
    for ent in sorted(entities):
        for yy in years:
            items = [(c, v, n) for (e, c, n, y), v in tb_full.items()
                     if e == ent and str(y) == yy]
            if not items:
                continue
            leafs = S.leaf_codes([c for c, _v, _n in items])
            # ⚡ read_tb_full 的 v 无 name 字段（name 在键里）→ 塞入（与 recon_self 同口径）
            ent_tb = {c: dict(v, name=n) for c, v, n in items if c in leafs}
            if not ent_tb:
                continue
            bc = S.balance_check(ent_tb)
            # 发生额平衡：差应=未结转损益（pl/cost 期末余额），非则数据问题
            if not bc['ok_jf']:
                tol = abs(bc['pl_open'])
                if abs(bc['jf_df']) - tol > hard_threshold:
                    issues.append(('FAIL', ent, yy,
                                   f'借发≠贷发 差 {bc["jf_df"]:,.2f}（未结转损益仅 {tol:,.2f}，差值 {abs(bc["jf_df"]) - tol:,.2f}）',
                                   round(bc['jf_df'], 2)))
                else:
                    issues.append(('WARN', ent, yy,
                                   f'借发≠贷发 差 {bc["jf_df"]:,.2f} = 未结转损益 {tol:,.2f}（结转后应平）',
                                   round(bc['jf_df'], 2)))
            # 余额平衡：期末借余=期末贷余（资产=负债+权益）
            if not bc['ok_qm'] and abs(bc['qm_diff']) > hard_threshold:
                issues.append(('FAIL', ent, yy,
                               f'期末借余 {bc["qm_pos"]:,.2f} ≠ 贷余 {bc["qm_neg"]:,.2f} 差 {bc["qm_diff"]:,.2f}',
                               round(bc['qm_diff'], 2)))
    fails = [i for i in issues if i[0] == 'FAIL']
    if fails:
        if not quiet:
            print(f'  ❌ [TB 控制数 Fail Fast] {os.path.basename(data_dir)}：{len(fails)} 项硬性问题')
            for lv, ent, yy, msg, d in issues:
                print(f'     {lv} [{ent}][{yy}] {msg}')
        if fail_fast:
            raise TBValidationError(f'{os.path.basename(data_dir)} TB 不平衡 {len(fails)} 项')
        return False, issues
    if not quiet and issues:
        print(f'  ⚠️ [TB 控制数] {os.path.basename(data_dir)}：{len(issues)} 项口径提示（未结转损益等，不阻断）')
        for lv, ent, yy, msg, d in issues:
            print(f'     {lv} [{ent}][{yy}] {msg}')
    return True, issues


def main():
    dirs = [a for a in sys.argv[1:] if not a.startswith('-')]
    ff = '--fail-fast' in sys.argv
    if not dirs:
        print('用法：python tb_validate.py <账套目录...> [--fail-fast]')
        return 1
    total_fail = 0
    for d in dirs:
        try:
            ok, issues = validate_tb(d, fail_fast=ff)
            if not ok:
                total_fail += sum(1 for i in issues if i[0] == 'FAIL')
        except TBValidationError as ex:
            print(f'  ⛔ 中止：{ex}')
            total_fail += 1
    print(f'总 FAIL 数：{total_fail}')
    return 1 if total_fail else 0


if __name__ == '__main__':
    sys.exit(main())
