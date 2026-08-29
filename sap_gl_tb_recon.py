# -*- coding: utf-8 -*-
"""SAP GL vs TB 核对器（sap_gl_tb_recon.py，2026-08-07 新建）。

用途：直接用序时账 GL 核对科目余额表 TB——余额+发生额双口径，检验虚增虚减剔除效果。
对每个公司：
  1. GL 按科目聚合借/贷发生额（col29 本币金额：正=借 负=贷）；
  2. TB 按科目代码聚合 jf/df/qc/qm（read_sap_tb）；
  3. 对比：GL借 vs TB借发、GL贷 vs TB贷发、GL期末(TB期初+借-贷) vs TB期末；
  4. 三种口径：全量 / 剔除stat+flow+text / 仅剔stat，看哪种最接近 TB。

结论用途：
  - 检验 GL 能否独立核对 TB（余额+发生额）；
  - 判断虚增虚减剔除规则（classify_voucher）是否准确。
"""
import paths as P
import os
import re
import sys
import glob
from collections import defaultdict

import sap_reader as SR

STAT_PK = ('99',)
FLOW_PK = ('81', '83', '86', '89', '91', '93', '96')


def _is_stat(pk):
    return str(pk or '').strip() in STAT_PK


def _is_flow(pk):
    return str(pk or '').strip() in FLOW_PK


def read_gl_by_subj(data_dir, comp, mode='all'):
    """从序时账按科目代码聚合 GL 借/贷发生额。
    mode: 'all' 全量；'normal' 剔除 stat+flow+text；'nostat' 仅剔 stat。"""
    layout = SR.detect_sap_layout(data_dir)
    gl_dir = os.path.join(data_dir, '序时账') if layout == 'ah-sap' else data_dir
    files = []
    if layout == 'ah-sap':
        for f in sorted(os.listdir(gl_dir)):
            if not f.lower().endswith(('.xlsx', '.xls')):
                continue
            base = os.path.splitext(f)[0]
            m = re.match(r'^(\d{4})', base)
            if not m:
                continue
            # ⚡ 2026-08-09 修复：区间合并文件（2610~5040 等）也须纳入（原只匹配前缀==comp 漏区间公司）
            m2 = re.search(r'[-~](\d{4})$', base)
            if m.group(1) == comp or (m2 and int(m.group(1)) <= int(comp) <= int(m2.group(1))):
                files.append(os.path.join(gl_dir, f))
    else:
        files = [os.path.join(gl_dir, f) for f in sorted(os.listdir(gl_dir))
                 if re.match(r'^\d+月\.XLSX$', f, re.I)]
    agg = defaultdict(lambda: [0.0, 0.0])  # 科目 → [借, 贷]
    import openpyxl
    for fp in files:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 30:
                continue
            c = str(r[6]).strip() if r[6] else ''
            if c != comp:
                continue
            subj = str(r[11] or '').strip()
            if not subj:
                continue
            # 只统计 TB 对应期间（1-7月 YTD；GL 可能含 8 月后续期间）
            m = str(r[10] or '').strip()
            if not (m[:2].isdigit() and 1 <= int(m[:2]) <= 7):
                continue
            amt = float(r[29] or 0)
            if mode == 'normal':
                pk = r[4]
                vc = SR.classify_voucher(pk, str(r[5] or '') + str(r[11] or ''))
                if vc != 'normal':
                    continue
            elif mode == 'nostat':
                if _is_stat(r[4]):
                    continue
            a = agg[subj]
            if amt >= 0:
                a[0] += amt
            else:
                a[1] += -amt
        wb.close()
    return agg


def read_tb_by_code(data_dir, comp, year='2026'):
    """TB 按科目代码聚合 {code: {'qc','jf','df','qm'}}。
    直接读科目余额表原始文件（read_sap_tb 丢弃内部码，无法按 code 对齐 GL）。
    只读目标公司相关文件（1010.xlsx / 2610~5040.xlsx 合并文件），避免全量读。"""
    import openpyxl
    layout = SR.detect_sap_layout(data_dir)
    tb_dir = os.path.join(data_dir, '科目余额表') if layout == 'ah-sap' else data_dir
    files = []
    if layout == 'ah-sap':
        # 单公司文件或含该公司的合并文件
        for f in sorted(os.listdir(tb_dir)):
            if not f.lower().endswith('.xlsx'):
                continue
            base = os.path.splitext(f)[0]
            m = re.match(r'^(\d{4})', base)
            if not m:
                continue
            if m.group(1) == comp:
                files.append(os.path.join(tb_dir, f))
                break  # 单公司文件优先
            if '~' in base:
                a, b = base.split('~')
                if a <= comp <= b:
                    files.append(os.path.join(tb_dir, f))
    else:
        for f in sorted(os.listdir(tb_dir)):
            if re.search(r'科目余额表.*\.XLSX$', f, re.I):
                files.append(os.path.join(tb_dir, f))
    out = {}
    for fp in files:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=3, values_only=True):
            if not r or len(r) < 17 or r[2] is None:
                continue
            comp_raw = str(r[7]).strip() if r[7] else ''
            m = re.search(r'(\d{4})\s*$', comp_raw)
            if not m or m.group(1) != comp:
                continue
            name = str(r[2]).strip()
            mm = re.match(r'^(.*?)\s{2,}([A-Z]+\d*/\d+|\d+)$', name)
            code = (mm.group(2) if mm else '').split('/')[-1]
            period = str(r[8]).strip()
            a = out.setdefault(code, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0})
            if '剩余' in period:
                a['qc'] += float(r[9] or 0)
            else:
                a['jf'] += float(r[12] or 0)
                a['df'] += float(r[13] or 0)
            if '七月' in period:
                a['qm'] += float(r[16] or 0)
        wb.close()
    return out


def recon_company(data_dir, comp, verbose=False):
    """核对单公司：GL 全量/剔除 vs TB。返回 (总科目数, 全量差异科目数, normal差异科目数, 余额差异科目数)。"""
    gl_all = read_gl_by_subj(data_dir, comp, 'all')
    gl_normal = read_gl_by_subj(data_dir, comp, 'normal')
    tb = read_tb_by_code(data_dir, comp)
    codes = set(gl_all) | set(tb)
    n_all = n_normal = n_bal = 0
    diffs = []
    for code in sorted(codes):
        g = gl_all.get(code, [0.0, 0.0])
        gn = gl_normal.get(code, [0.0, 0.0])
        t = tb.get(code, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0})
        # 发生额差异（全量）
        d_jf = g[0] - t['jf']
        d_df = g[1] - t['df']
        # 发生额差异（剔除虚增虚减）
        dn_jf = gn[0] - t['jf']
        dn_df = gn[1] - t['df']
        # 期末余额差异：GL = TB期初 + GL借 - GL贷 vs TB期末
        gl_bal = t['qc'] + g[0] - g[1]
        d_bal = gl_bal - t['qm']
        big_all = abs(d_jf) > 1 or abs(d_df) > 1
        big_normal = abs(dn_jf) > 1 or abs(dn_df) > 1
        big_bal = abs(d_bal) > 1
        if big_all:
            n_all += 1
        if big_normal:
            n_normal += 1
        if big_bal:
            n_bal += 1
        if verbose and (big_all or big_normal or big_bal):
            diffs.append((code, d_jf, d_df, dn_jf, dn_df, d_bal, t['qm']))
    return len(codes), n_all, n_normal, n_bal, diffs


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else P.YY
    comps = sys.argv[2].split(',') if len(sys.argv) > 2 and sys.argv[2] != '*' else None
    if comps is None:
        # 发现所有有 GL 的公司
        comps = set()
        for f in glob.glob(os.path.join(data_dir, '序时账', '*.xlsx')):
            m = re.match(r'^(\d{4})', os.path.basename(f))
            if m:
                comps.add(m.group(1))
        comps = sorted(comps)
    print(f'公司数: {len(comps)}')
    stats = []
    for comp in comps:
        nc, na, nn, nb, diffs = recon_company(data_dir, comp)
        status = '✅' if (na == 0 and nb == 0) else ('⚠️' if nn == 0 else '❌')
        stats.append((comp, nc, na, nn, nb))
        print(f'{status} {comp}: 科目{nc} 全量发生额差{na} 剔虚增差{nn} 余额差{nb}')
    ok = [s for s in stats if s[2] == 0 and s[4] == 0]
    print(f'\n全对（余额+发生额均 0 差异，全量口径）: {len(ok)}/{len(stats)}')
    print([s[0] for s in ok])


if __name__ == '__main__':
    main()
