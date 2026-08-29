# -*- coding: utf-8 -*-
"""底稿综合质量审查（2026-08-24 用户要求：对所有底稿做逻辑/格式/内容全面审查）。

对每个账套执行三层检查：
  1. 结构规范   audit_checker.process_folder → ERROR/WARN（表头/合计行/公式等结构问题）
  2. 完整性     self_check → 空白底稿（TB 有数据但整份空 = 取数 bug）
  3. 数据核对   recon_self → 底稿审定数 vs 自建试算表差异（同源核对）
输出统一报告：每账套 结构 ERROR/WARN + 空白 + 核对差异。

用法：
  python wp_review.py <数据目录>...
"""
import os
import sys
import glob
import argparse


def review_account(data_dir, quiet=True):
    """单账套三层审查，返回 dict 报告。"""
    rep = {'dir': data_dir}
    # 1. 结构规范
    try:
        import audit_checker
        _e, _w = audit_checker.process_folder(data_dir, quiet=True)
        rep['struct'] = (_e, _w)
    except Exception as ex:
        rep['struct'] = (None, str(ex)[:50])
    # 2. 完整性（空白底稿）
    try:
        import self_check
        _ok, _lines = self_check.self_check(data_dir, quiet=True)
        rep['blank_ok'] = _ok
        rep['blank'] = [l.strip() for l in _lines if '❌' in l]
    except Exception as ex:
        rep['blank_ok'] = None
        rep['blank'] = [str(ex)[:50]]
    # 3. 数据核对
    try:
        import recon_self
        _ok, _diffs = recon_self.recon_self(data_dir, quiet=True)
        rep['recon_ok'] = _ok
        rep['diffs'] = _diffs
    except Exception as ex:
        rep['recon_ok'] = None
        rep['diffs'] = [(str(ex)[:50],)]
    return rep


def report_line(rep):
    _s = rep.get('struct') or (None, '')
    _se = _s[0] if isinstance(_s, tuple) else None
    _sw = _s[1] if isinstance(_s, tuple) else None
    _b = '✅' if rep.get('blank_ok') else ('⚠️' if rep.get('blank_ok') is None else '❌')
    _r = '✅' if rep.get('recon_ok') else '❌'
    _nd = len(rep.get('diffs') or [])
    return (f"结构 E{_se}/W{_sw} | 空白 {_b} | 核对 {_r}({_nd}差异)")


def main():
    p = argparse.ArgumentParser()
    p.add_argument('dirs', nargs='+')
    p.add_argument('--out', default=None)
    a = p.parse_args()
    lines = []
    for d in a.dirs:
        rep = review_account(d)
        line = f"{rep['dir']}\n   {report_line(rep)}"
        for bl in rep.get('blank') or []:
            line += f"\n     [空白] {bl}"
        for df in rep.get('diffs') or []:
            line += f"\n     [差异] {str(df[:5])}"
        print(line, flush=True)
        lines.append(line)
    if a.out:
        with open(a.out, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))
    return 0


if __name__ == '__main__':
    sys.exit(main())
