# -*- coding: utf-8 -*-
"""88家往来核对 数据完整性自检（2026-08-27 建立，防漏提/防口径差滞后发现）。

背景：2026-08-27 序时账『对方公司名』(tc) 曾被污染（52,244 行误填银行账户名），
导致 _load_seq 按 (公司,tc) 匹配对子键时漏提 5.2 万行往来分录，程序补录/差异虚高。
当时是用户反复质疑后才定位，而非生成时自检发现。此脚本为『生成后必跑』的第一道闸门，
把漏提/口径差在交付前暴露。

校验项：
  ① tc 字段完整性    —— 预处理序时账『对方公司名』为空 / 含异常词(银行/年金/受托/信托/账户/账号) 的行数
  ② 对子提取覆盖      —— 核对表有期末余额的对子中，序时账按对子提取=0 的对子（红旗=疑似漏提）
  ③ GL vs TB 科目勾稽 —— 预处理序时账各往来科目借/贷发生额 vs 科目余额表(TB) 借/贷发生额，
                          差>阈值(默认 100 万) 报警并列出科目（外部单位往来属正常，需人工判断；
                          若差额集中在某对子缺失科目=漏提）
  ④ 勾稽校验表状态    —— 165 个需穿透对子 check/pool_check 归零情况

用法：
  python ah_intra_integrity_check.py [--out OUTDIR]
退出码：0=全部通过；1=存在红旗（生成结果仍有效，但需人工复核）。
"""
import os
import re
import sys
import glob
import argparse
from collections import defaultdict, Counter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)
import openpyxl  # noqa: E402

PREP = r'd:/底稿测试/AH/中间产物/prepared'
SEQ_FP = os.path.join(PREP, '关联往来序时账_2026.xlsx')
PAIR_FP = os.path.join(PREP, '88家内部往来核对_2026.xlsx')
CHECK_FP = os.path.join(PREP, '关联往来勾稽校验_2026.xlsx')
TB_DIR = r'd:/底稿测试/AH/数据/2026/科目余额表'

# 往来科目（与 ah_intra_group_recon.SUBJS 前 6 位一致）
RECV_KMS = ('1122', '1123', '1221', '2202', '2203', '2241')
# tc 异常词（SAP 账户/银行/信托等非集团主体）
BAD_TC = ('银行', '年金', '受托', '信托', '账户', '账号', '保证金', '社保', '公积金', '工会')


def check_tc_integrity():
    """① 预处理序时账 tc 完整性。"""
    if not os.path.exists(SEQ_FP):
        return {'fatal': f'缺序时账 {SEQ_FP}'}
    total = empty = bad = 0
    bad_ex = Counter()
    wb = openpyxl.load_workbook(SEQ_FP, read_only=True, data_only=True)
    ws = wb['关联往来序时账']
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        total += 1
        tc = str(r[10] or '').strip()
        if not tc:
            empty += 1
        elif any(k in tc for k in BAD_TC):
            bad += 1
            bad_ex[tc[:20]] += 1
    wb.close()
    return {'total': total, 'empty': empty, 'bad': bad,
            'bad_examples': bad_ex.most_common(5)}


def check_pair_coverage():
    """② 差异对子提取覆盖：状态≠对平 的对子应进入穿透提取范围。
    红旗 = 状态为 经营差异/资金池差异 但未进入 _load_pairs（穿透需求非『需穿透』）
           —— 可能是期初延续/本期已平（正常），需人工逐对确认。"""
    wb = openpyxl.load_workbook(PAIR_FP, read_only=True, data_only=True)
    ws = wb['内部往来逐对核对']
    diff_pairs = {}   # 序号 -> (X, Y, 状态, 穿透需求)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0] or not isinstance(r[0], (int, float)):
            continue
        st = str(r[44] or '')
        if st in ('经营差异', '资金池差异'):
            diff_pairs[int(r[0])] = (str(r[1]), str(r[2]), st, str(r[46] or ''))
    wb.close()
    import ah_intra_diff_trace as m
    pairs = m._load_pairs()
    covered = {p['no'] for p in pairs}
    not_covered = []
    for no, (x, y, st, need) in diff_pairs.items():
        if no not in covered:
            not_covered.append((no, x, y, st, need))
    return {'diff_pairs': len(diff_pairs), 'not_covered': not_covered}


def check_gl_vs_tb(threshold=1_000_000):
    """③ 预处理序时账往来科目净额 vs TB 净额。
    用净额(借-贷)对比（资金池类科目大进大出，发生额口径天然不具可比性）。"""
    # TB 汇总（三集团）
    import sap_reader
    tb = sap_reader.read_sap_tb(r'd:/底稿测试/AH/数据/2026', ['1010', '1357', '2468'], '2026')
    tb_net = defaultdict(float)
    for (e, c, n, y), v in tb.items():
        p = str(c)[:4]
        if p in RECV_KMS:
            tb_net[p] += v['jf'] - v['df']
    # 预处理序时账汇总净额
    seq_net = defaultdict(float)
    wb = openpyxl.load_workbook(SEQ_FP, read_only=True, data_only=True)
    ws = wb['关联往来序时账']
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        km = str(r[4] or '')[:4]
        if km in RECV_KMS:
            seq_net[km] += float(r[8] or 0) - float(r[9] or 0)
    wb.close()
    diffs = []
    for p in RECV_KMS:
        d = seq_net[p] - tb_net[p]
        if abs(d) > threshold:
            diffs.append((p, seq_net[p], tb_net[p], d))
    return {'diffs': diffs}


def check_recon_state():
    """④ 勾稽校验表归零状态。"""
    if not os.path.exists(CHECK_FP):
        return {'fatal': f'缺勾稽校验 {CHECK_FP}'}
    wb = openpyxl.load_workbook(CHECK_FP, read_only=True, data_only=True)
    ws = wb['勾稽校验']
    n = ok = 0
    bad = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        n += 1
        st = str(r[9] or '')
        if st.startswith('归零'):
            ok += 1
        else:
            bad.append((int(r[0]), st))
    wb.close()
    return {'n': n, 'ok': ok, 'bad': bad}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--threshold', type=float, default=1_000_000,
                    help='GL vs TB 科目级报警阈值（默认 100 万）')
    a = ap.parse_args(argv)

    flags = []
    print('=' * 60)
    print('88家往来核对 数据完整性自检')
    print('=' * 60)

    # 脱敏：公司名→账套代码（对话不出现公司名）
    try:
        from ah_intra_seq_extract import load_company_map
        _cmap = load_company_map()
        _name2code = {v: k for k, v in _cmap.items()}
    except Exception:
        _name2code = {}

    def _mask(nm, idx):
        c = _name2code.get(str(nm))
        if c:
            return c
        return 'A%03d' % idx

    # ① tc 完整性
    r1 = check_tc_integrity()
    if 'fatal' in r1:
        print(f'[✗] tc 完整性: {r1["fatal"]}')
        flags.append('fatal')
    else:
        st1 = 'OK' if (r1['empty'] == 0 and r1['bad'] == 0) else '⚠ 红旗'
        if r1['bad'] or r1['empty']:
            flags.append('tc')
        print(f'[{"✓" if st1=="OK" else "⚠"}] ① tc 完整性: {st1} | 总行 {r1["total"]:,} | '
              f'空 tc {r1["empty"]:,} | 异常 tc {r1["bad"]:,}')
        for k, v in r1['bad_examples']:
            print(f'       异常样例: {k}… ×{v}')

    # ② 对子提取覆盖
    r2 = check_pair_coverage()
    st2 = 'OK' if not r2['not_covered'] else '⚠ 红旗(需人工确认是否期初延续/本期已平)'
    if r2['not_covered']:
        flags.append('coverage')
    print(f'[{"✓" if st2=="OK" else "⚠"}] ② 差异对子提取覆盖: {st2} | '
          f'差异对子 {r2["diff_pairs"]} | 未进入穿透 {len(r2["not_covered"])}')
    for i, (no, x, y, st, need) in enumerate(r2['not_covered'], 1):
        print(f'       对子 {no} | {_mask(x,i)} ↔ {_mask(y,i+100)} | 状态={st} | 穿透需求={need}')

    # ③ GL vs TB
    r3 = check_gl_vs_tb(a.threshold)
    st3 = 'OK' if not r3['diffs'] else '⚠ 红旗(净额口径差，需人工判断是否外部单位往来/口径差)'
    if r3['diffs']:
        flags.append('gl_vs_tb')
    print(f'[{"✓" if st3=="OK" else "⚠"}] ③ GL vs TB 净额勾稽(阈值 {a.threshold:,.0f}): {st3}')
    for p, sn, tn, d in r3['diffs']:
        print(f'       科目 {p}: 序时账净额 {sn:,.0f} vs TB净额 {tn:,.0f} 差 {d:,.0f}')

    # ④ 勾稽校验
    r4 = check_recon_state()
    if 'fatal' in r4:
        print(f'[✗] ④ 勾稽校验: {r4["fatal"]}')
        flags.append('fatal')
    else:
        st4 = 'OK' if r4['ok'] == r4['n'] else '⚠'
        if r4['bad']:
            flags.append('recon')
        print(f'[{"✓" if st4=="OK" else "⚠"}] ④ 勾稽校验: {st4} | {r4["ok"]}/{r4["n"]} 归零')
        for no, st in r4['bad'][:5]:
            print(f'       对子 {no}: {st}')

    print('=' * 60)
    if flags:
        print(f'自检结论: 存在红旗 {flags} —— 生成结果仍有效，但必须复核后交付。')
        return 1
    print('自检结论: 全部通过，可交付。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
