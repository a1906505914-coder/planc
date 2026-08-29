# -*- coding: utf-8 -*-
"""regress_all.py —— 全账套回归闸门（2026-08-07 用户要求：改公共层代码后，
防"适配 A 破坏 B"行为漂移）。

用法：
    python regress_all.py check          # 只核对：全账套 recon_self + 快照比对（不重跑底稿）
    python regress_all.py full           # 全量：13 账套 regen_all → recon_self → 快照比对
    python regress_all.py fast [账套...]  # 快速：只跑指定账套 recon_self + 快照（改核对器后）

输出回归报告（控制台 + regress_report.txt）：
    · 每账套 recon_self 差异数（与基线对比）
    · 快照差异（改生成器后哪些底稿值变了 → 人工确认是否预期）
    · 结论：PASS（无回归）/ REVIEW（有变化需确认）

基线文件：snapshots/snapshot_<账套>.json（snapshot_regress.py capture 生成）。
改动公共层（audit_common/subject_mapping/subjects_registry/recalc_totals）后必跑。
"""
import os
import sys
import subprocess

import paths as P   # ⚡ 2026-08-19 修复：ACCTS 顶层使用 P，原 import 在函数内导致 NameError

PY = sys.executable
BASE = os.path.dirname(os.path.abspath(__file__))
ACCTS = list(P.ACCT_NAMES)   # ⚡ 2026-08-15 全部账套（新名由 paths 唯一事实来源）

REPORT = []


def log(msg, also_print=True):
    REPORT.append(msg)
    if also_print:
        print(msg)


def run_self(acct):
    """recon_self 差异数；返回 (n_diffs, detail) 或 (None, err)。"""
    d = os.path.join(P.DATA_ROOT, acct)
    try:
        r = subprocess.run([PY, os.path.join(BASE, 'recon_self.py'), d],
                           capture_output=True, text=True, encoding='utf-8', errors='replace',
                           timeout=1800)
        out = r.stdout + r.stderr
    except Exception as e:
        return None, str(e)
    import re
    m = re.search(r'差异 (\d+) 项', out)
    if m:
        n = int(m.group(1))
        return n, out
    if '同源 0 差异' in out or '0 差异' in out:
        return 0, out
    return None, out[:200]


def run_snapshot(acct):
    """快照 check；返回 (changed, detail)。"""
    d = os.path.join(P.DATA_ROOT, acct)
    try:
        r = subprocess.run([PY, os.path.join(BASE, 'snapshot_regress.py'), 'check', d],
                           capture_output=True, text=True, encoding='utf-8', errors='replace',
                           timeout=1800)
        out = r.stdout + r.stderr
    except Exception as e:
        return None, str(e)
    import re
    m = re.search(r'差异 (\d+) 科目', out)
    if m:
        return int(m.group(1)), out
    if '无变化' in out:
        return 0, out
    return None, out[:200]


def _is_checker_acct(acct):
    """⚡ 2026-08-14 配置化：特殊回归闸门账套（走 audit_checker 而非 recon_self）。
    account_profiles.json -> accounts.<账套>.acct_type：
      'checker' = 全量扫（如 300 集团网银核对主账套）
      'sap'     = SAP 抽查（如 yy）
    账套改名只改配置/ACCT_NAMES，本函数零改动。"""
    import subject_mapping as _SM
    return _SM.feature(acct, 'acct_type')


def run_sap_check(acct):
    """⚡ 2026-08-11 P0：SAP/网银核对账套回归（U8 的 recon_self/快照不适用，SAP 底稿在 底稿/ 子目录）。
    配置驱动：acct_type='checker'（如 300）→ audit_checker 全量（ERROR 必须 0）；
    acct_type='sap'（如 yy）→ 抽查 5 家有代表性公司 × 8 个关键科目 check_file（ERROR 必须 0）。
    返回 (error_count, detail)。"""
    import audit_checker as AC
    import paths as P
    _type = _is_checker_acct(acct)
    if _type == 'checker':
        te, tw = AC.process_folder(os.path.join(P.DATA_ROOT, acct), quiet=True)
        return te, f'{acct} 全量: ERROR={te} WARN={tw}'
    # SAP 抽查
    _comps = ['1010', '1030', '1050', '1080', '2020']
    _subs = ['应收账款', '应付账款', '存货', '营业收入', '管理费用',
             '固定资产', '应交税费', '银行存款']
    errs = []
    n_files = 0
    _out = os.path.join(P.DATA_ROOT, acct, '底稿')
    for c in _comps:
        for s in _subs:
            fn = os.path.join(_out, f'{s}审计底稿_{c}.xlsx')
            if not os.path.exists(fn):
                continue
            n_files += 1
            try:
                issues = AC.check_file(fn)
                ne = sum(1 for i in issues if i[0] == 'ERROR')
                if ne:
                    errs.append(f'{c}/{s}:{ne}')
            except Exception as e:
                errs.append(f'{c}/{s}:读失败{str(e)[:30]}')
    if errs:
        return len(errs), f'{acct} 抽查 {n_files} 份 → ERROR {len(errs)} 处: ' + '; '.join(errs[:6])
    return 0, f'{acct} 抽查 {n_files} 份全部 ERROR 0 ✓'


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else 'check'
    accts = sys.argv[2:] or ACCTS

    log(f'===== 全账套回归（{mode}）=====')
    log(f'时间：{__import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
    log('')

    # ⚡ 2026-08-22 台账脱敏闸门：进对话/分析前脱敏版必须齐备（回归前置检查）
    try:
        import ledger_registry as LREG
        _ledger_miss = LREG.check()
        if _ledger_miss:
            log(f'  ⚠️ 台账脱敏不全: {_ledger_miss} → 先跑对应 mask 脚本再分析')
            log('')
    except Exception as _ex:
        _ledger_miss = []
        log(f'  ⚠️ 台账脱敏闸门异常: {_ex}')

    results = []
    for acct in accts:
        log(f'--- {acct} ---')
        # ⚡ 2026-08-11 P0：特殊账套（acct_type='checker'/'sap'，配置驱动）走 audit_checker 闸门
        if _is_checker_acct(acct):
            n, detail = run_sap_check(acct)
            log(f'  sap_check: {detail}')
            results.append((acct, n, None))
            log('')
            continue
        n, detail = run_self(acct)
        status = f'{n} 项' if n is not None else 'ERR'
        log(f'  recon_self: {status}')
        if n and n > 0:
            for line in detail.splitlines():
                if '·' in line:
                    log(f'    {line.strip()[:110]}')
        if mode == 'full':
            continue  # 全量模式快照由 regen_all 后单独跑
        sc, sdetail = run_snapshot(acct)
        if sc is None:
            log(f'  快照: ERR（{sdetail[:80]}）')
        elif sc > 0:
            log(f'  快照: {sc} 科目变化 ⚠️ 需确认')
            for line in sdetail.splitlines():
                if '✗' in line or '→' in line:
                    log(f'    {line.strip()[:110]}')
        else:
            log(f'  快照: 无变化 ✓')
        results.append((acct, n, sc))
        log('')

    log('===== 汇总 =====')
    any_review = False
    if _ledger_miss:
        any_review = True
        log(f'  台账脱敏: 不全（{_ledger_miss}）→ REVIEW')
    else:
        log(f'  台账脱敏: 齐备 → PASS')
    for acct, n, sc in results:
        if _is_checker_acct(acct):
            flag = 'PASS' if n == 0 else 'REVIEW'
        else:
            flag = 'PASS' if (n in (0, None)) and sc in (0, None) else 'REVIEW'
        if flag == 'REVIEW':
            any_review = True
        log(f'  {acct:4s}: {"sap_check=" if _is_checker_acct(acct) else "recon_self="}{n}  快照={sc}  → {flag}')
    log('')
    if any_review:
        log('结论：REVIEW —— 有变化项，请逐项确认是否预期（改生成器后快照变化正常，需 capture 更新基线）')
    else:
        log('结论：PASS —— 无回归')
    with open(os.path.join(BASE, 'regress_report.txt'), 'w', encoding='utf-8') as f:
        f.write('\n'.join(REPORT))
    print(f'\n报告已写入：{os.path.join(BASE, "regress_report.txt")}')


if __name__ == '__main__':
    main()
