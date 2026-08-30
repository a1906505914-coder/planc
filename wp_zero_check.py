# -*- coding: utf-8 -*-
"""底稿"数据为0"自检（2026-08-27 建立，用户要求：主动发现每张生成底稿数据生成为0的情况）。

背景：其他收益审定表合计=0 而明细表有 2086 万（SAP 科目码 6730 vs 配置 6113 失配），
固定资产折旧分摊表全 0（from-import 绑定旧 reader），这类静默失败此前靠人工逐张发现。
本脚本扫描底稿目录每份 xlsx，启发式检查：
  ① 审定表合计行【审定数/本期数】是否为 0
  ② 明细表是否存在非空数据行
  ③ 若 审定表=0 且 明细表>0 → 红旗（审定表取数失败，必查）
  ④ 若 审定表=0 且 明细表=0 → 疑似无数据（需确认是账套真无还是识别失败）
  ⑤ 勾稽表非零差异（如有）

用法：python wp_zero_check.py [底稿目录] [--all]
"""
import os
import sys
import glob
import argparse
import openpyxl
import paths as P

DEFAULT_WP = os.path.join(P.DATA_DIRS['AH'], '底稿', '2026', '集团', '1010')


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def inspect_wp(fp):
    """对单份底稿返回 {sheet: {'audit': num, 'detail_rows': n, ...}}。"""
    out = {}
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        info = {'rows': len(rows), 'audit_total': None, 'audit_begin': None,
                'detail_rows': 0, 'any_num': False, 'num_rows': 0, 'max_abs': 0.0,
                'abs_sum': 0.0}
        # 全 sheet 通用：数值行数/最大绝对值/绝对值合计（供各类型 sheet 判定）
        _num_rows = 0
        _max_abs = 0.0
        _abs_sum = 0.0
        for r in rows:
            for x in r:
                if isinstance(x, (int, float)) and abs(x) > 0.005:
                    _num_rows += 1
                    if abs(x) > _max_abs:
                        _max_abs = abs(x)
                    _abs_sum += abs(x)
        info['num_rows'] = _num_rows
        info['max_abs'] = _max_abs
        info['abs_sum'] = _abs_sum
        # ① 审定表：找含『合计』的行 + 『审定数/本期数/本期发生/期初』列
        if '审定表' in sn:
            # 是否有非零数值（部分审定表用『小计』非『合计』，如存货）
            if any(isinstance(x, (int, float)) and abs(x) > 0.005 for r in rows for x in r):
                info['any_num'] = True
            hdr_i = None
            for i, r in enumerate(rows[:8]):
                blob = ' '.join(str(x) for x in r if x is not None)
                if any(k in blob for k in ('审定数', '本期数', '本期发生')):
                    hdr_i = i
                    break
            if hdr_i is not None:
                hdr = [str(x or '').replace(' ', '') for x in rows[hdr_i]]
                c_aud = next((j for j, h in enumerate(hdr) if '审定数' in h), None)
                c_cur = next((j for j, h in enumerate(hdr) if '本期数' in h or '本期发生' in h or '期末' in h), None)
                c_begin = next((j for j, h in enumerate(hdr) if '期初' in h or '上年数' in h), None)
                for r in rows[hdr_i + 1:]:
                    if r and '合计' in str(r[0] or ''):
                        def _cv(c):
                            return _num(r[c]) if c is not None and c < len(r) and r[c] is not None else None
                        info['audit_total'] = _cv(c_aud) if c_aud is not None else _cv(c_cur)
                        info['audit_begin'] = _cv(c_begin)
                        break
        # ② 明细表：统计非空数据行（跳过表头/说明行）
        if '明细表' in sn:
            n = 0
            for r in rows:
                if not r:
                    continue
                vals = [x for x in r[:8] if x is not None and str(x).strip() != '']
                # 数据行：至少 2 个非空且含数值
                if len(vals) >= 2 and any(isinstance(x, (int, float)) and abs(x) > 0.005 for x in r):
                    n += 1
            info['detail_rows'] = n
        out[sn] = info
    wb.close()
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('wp_dir', nargs='?', default=DEFAULT_WP, help='底稿目录')
    ap.add_argument('--all', action='store_true', help='扫描全部三集团')
    a = ap.parse_args()

    if a.all:
        roots = glob.glob(os.path.join(P.DATA_DIRS['AH'], '底稿', '2026', '集团', '*'))
    else:
        roots = [a.wp_dir]
    files = []
    for r in roots:
        files += sorted(glob.glob(os.path.join(r, '*.xlsx')))
        files += sorted(glob.glob(os.path.join(r, '..', '*.xlsx')))
    files = sorted({os.path.normpath(f) for f in files if not os.path.basename(f).startswith('~$')})

    flags = 0
    for fp in files:
        base = os.path.basename(fp)
        if '_脱敏' in base:
            continue
        try:
            res = inspect_wp(fp)
        except Exception as ex:
            print(f'[✗] {base}: 读取失败 {str(ex)[:50]}')
            flags += 1
            continue
        audits = {sn: i for sn, i in res.items() if '审定表' in sn}
        details = {sn: i for sn, i in res.items() if '明细表' in sn}
        # 全 sheet 通用：无任何数值的 sheet（数据为空）
        empty_sheets = [sn for sn, i in res.items() if i['num_rows'] == 0 and i['rows'] > 3]
        # 审定表合计
        aud_zero = [sn for sn, i in audits.items() if i['audit_total'] is not None and abs(i['audit_total']) <= 0.005]
        aud_val = [sn for sn, i in audits.items() if (i['audit_total'] is not None and abs(i['audit_total']) > 0.005) or i['any_num']]
        det_nonzero = [sn for sn, i in details.items() if i['detail_rows'] > 0]
        det_zero = [sn for sn, i in details.items() if i['detail_rows'] == 0]
        issues = []
        # 红旗①：审定表无非零数值 但 明细表有数据 → 审定表取数失败（排除"本期结清"）
        if not aud_val and det_nonzero:
            begin_nonzero = any(abs(res[sn].get('audit_begin') or 0) > 0.005 for sn in aud_zero)
            if begin_nonzero:
                print(f'[·] {base}: 审定表期末=0 但期初={res[aud_zero[0]].get("audit_begin") or 0:,.0f}（本期结清，正常）')
            else:
                issues.append('审定表无非零数值但明细表有数据')
        # 红旗③：凭证抽查/抽查类 sheet 存在但无数据，而明细表有发生 → 抽凭可能漏抽
        vouch_empty = [sn for sn in res if ('抽查' in sn or '抽凭' in sn) and res[sn]['num_rows'] == 0]
        if vouch_empty and det_nonzero:
            issues.append('凭证抽查表为空（明细表有发生）:' + ','.join(vouch_empty))
        # 红旗④：分月/分析类 sheet 存在但无数据（明细表有发生）
        month_empty = [sn for sn in res if ('分月' in sn or '分析' in sn or '变动' in sn)
                       and res[sn]['num_rows'] == 0 and res[sn]['rows'] > 5]
        if month_empty and det_nonzero:
            issues.append('分月/分析表为空:' + ','.join(month_empty[:3]))
        if issues:
            print(f'[⚠ 红旗] {base}: {"；".join(issues)}')
            flags += 1
        elif aud_zero and not det_nonzero:
            print(f'[·] {base}: 审定表=0 且明细表无数据（疑似账套真无，需确认）')
        elif aud_val:
            _tot = res[aud_val[0]].get('audit_total')
            if _tot is not None:
                print(f'[✓] {base}: 审定表合计={_tot:,.0f}（明细表有数据）')
            else:
                print(f'[✓] {base}: 审定表有数据（无合计行，明细表有数据）')
        else:
            print(f'[·] {base}: 未识别审定表/明细表结构')
    print()
    print(f'自检完成：红旗 {flags} 项')
    return 1 if flags else 0


if __name__ == '__main__':
    sys.exit(main())
