# -*- coding: utf-8 -*-
"""科目快照回归工具（2026-08-06 用户需求：底稿生成经常回归 + 验证工作量巨大）。
对每个科目底稿固化「审定表/明细表/附注 三值 + 文件存在性」快照，
任何改动重跑后自动 diff，30 秒报出差异科目——把人工核对 40 份底稿变成自动回归。

用法：
  python snapshot_regress.py capture <文件夹> [更多文件夹...]   # 固化快照（首次/基线更新）
  python snapshot_regress.py check   <文件夹> [更多文件夹...]   # 对照快照报差异（退出码 1=有差异）
  python snapshot_regress.py clear   <文件夹> [更多文件夹...]   # 删除该文件夹快照

快照存放：小程序目录 snapshots/snapshot_<文件夹名>.json
  key = <文件夹名>|<年度>|<科目名>，值 = {file, mtime, aud, det, fn}
  aud/det/fn 为 None 表示该表未取到（与核对程序同口径：_sheet_end_total / _footnote_total）。

说明：
- 取数复用 tb_recon_detail 的函数（读底稿文件）——检测的是『底稿内容变化』，
  与核对程序同源，前后对比自洽。
- check 报告三类差异：值变化（回归）、文件缺失（底稿不再生成——追加附注后典型回归）、
  新增底稿（提示）。
"""
import sys, os, json, glob, re
import openpyxl

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
SNAP_DIR = os.path.join(TOOL_DIR, 'snapshots')
TOL = 0.01

# 明细表排除（备抵/计提/期间对比/对方科目/分摊 等非余额明细）
_DET_EXCLUDE = ('准备', '计提', '期间对比', '对方科目', '分摊', '勾稽')


def _load_recon():
    sys.path.insert(0, TOOL_DIR)
    from tb_recon_detail import _sheet_end_total, _footnote_total, _sheet_ent_values
    return _sheet_end_total, _footnote_total, _sheet_ent_values


def _extract(fp, year=None):
    """单份底稿 → (aud, det, fn)。None=未取到。
    year 提供时（2026-08-07 修复㉒）：混年底稿（c 存货审定表 2025+2026 同表，
    _sheet_end_total 取最后合计=2026）→ 用 _sheet_ent_values(ws, year=year)
    按年度取该年主体值合计，避免取错年。"""
    _end, _fn, _ent = _load_recon()
    m = re.search(r'_(\d{4})_生成\.xlsx$', os.path.basename(fp))
    subj = os.path.basename(fp).replace(f'审计底稿_{m.group(1)}_生成.xlsx', '') if m else \
        os.path.basename(fp).replace('审计底稿_生成.xlsx', '').replace('_生成.xlsx', '')
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception:
        return None, None, None, subj, (m.group(1) if m else '?')
    sheets = wb.sheetnames
    # ⚡⚡ 2026-08-24 修复（AYL 职工薪酬 280.8万 误报根因）：集团稿含
    #   『{科目} 审定表』（多主体分块、_sheet_end_total 只取某块）与
    #   『{科目}集团审定表_2025』（合计权威）。原取第一个『审定表』→
    #   AYL 取到分块表 5897.5万（正确 6178.3万）。优先选含『集团』或年份
    #   的审定表；无集团稿才回退第一个。
    aud_sn = (next((s for s in sheets if '审定表' in s and ('集团' in s or str(year) in s)), None)
              or next((s for s in sheets if '审定表' in s), None))
    det_sn = next((s for s in sheets if ('明细表' in s or '分类汇总' in s)
                   and not any(k in s for k in _DET_EXCLUDE)), None)
    fn_sn = next((s for s in sheets if '附注' in s), None)
    aud = det = fn = None
    if aud_sn:
        try:
            aud, _ = _end(wb[aud_sn])
            if year and (abs(aud or 0) < 0.005 or _is_mixed_year(wb[aud_sn])):
                _e = _ent(wb[aud_sn], year=year)
                if _e:
                    aud = sum(_e.values())
        except Exception:
            aud = None
    if det_sn:
        try:
            det, _ = _end(wb[det_sn])
        except Exception:
            det = None
        if det is None or abs(det) < 0.005:
            try:
                _e = _ent(wb[det_sn], year=year) if year else _ent(wb[det_sn])
                if _e:
                    det = sum(_e.values())
            except Exception:
                pass
    if fn_sn:
        try:
            fn = _fn(wb[fn_sn], subj)
        except Exception:
            fn = None
    wb.close()
    return aud, det, fn, subj, (m.group(1) if m else '?')


def _is_mixed_year(ws):
    """判定底稿 sheet 是否混排多年（表头含『年度』列且数据行有≥2 个年度）。"""
    try:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            return False
        hdr_i = next((i for i, r in enumerate(rows[:10])
                      if sum(1 for x in r[:40] if x is not None and str(x).strip()) >= 3
                      and not any(k in ''.join(str(x) if x is not None else '' for x in r[:40])
                                  for k in ('明细表', '审定表', '附注', '底稿', '汇总（'))), None)
        if hdr_i is None:
            return False
        ycol = next((j for j, h in enumerate(rows[hdr_i])
                     if h is not None and '年度' in str(h)), None)
        if ycol is None:
            return False
        years = {str(r[ycol]).strip() for r in rows[hdr_i + 1:]
                 if len(r) > ycol and r[ycol] is not None
                 and re.match(r'^\d{4}$', str(r[ycol]).strip())}
        return len(years) >= 2
    except Exception:
        return False


def _snap_path(folder):
    name = os.path.basename(os.path.normpath(folder))
    return os.path.join(SNAP_DIR, f'snapshot_{name}.json')


def capture(folders):
    os.makedirs(SNAP_DIR, exist_ok=True)
    for folder in folders:
        if not os.path.isdir(folder):
            print(f'⚠️ 文件夹不存在：{folder}')
            continue
        files = sorted(glob.glob(os.path.join(folder, '*审计底稿_*_生成.xlsx')))
        files = [f for f in files if not os.path.basename(f).startswith('~$')]
        snap = {}
        for fp in files:
            aud, det, fn, subj, yr = _extract(fp)
            key = f'{os.path.basename(os.path.normpath(folder))}|{yr}|{subj}'
            snap[key] = {
                'file': os.path.basename(fp), 'mtime': round(os.path.getmtime(fp), 1),
                'aud': aud, 'det': det, 'fn': fn,
            }
        out = _snap_path(folder)
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(snap, f, ensure_ascii=False, indent=1)
        print(f'✓ 快照已固化：{len(snap)} 科目底稿 → {out}')
    return 0


def check(folders):
    n_diff = 0
    for folder in folders:
        sp = _snap_path(folder)
        if not os.path.exists(sp):
            print(f'⚠️ 无快照：{sp}（先跑 capture）')
            n_diff += 1
            continue
        with open(sp, encoding='utf-8') as f:
            snap = json.load(f)
        files = sorted(glob.glob(os.path.join(folder, '*审计底稿_*_生成.xlsx')))
        files = [f for f in files if not os.path.basename(f).startswith('~$')]
        now = {}
        for fp in files:
            aud, det, fn, subj, yr = _extract(fp)
            key = f'{os.path.basename(os.path.normpath(folder))}|{yr}|{subj}'
            now[key] = {'aud': aud, 'det': det, 'fn': fn}
        name = os.path.basename(os.path.normpath(folder))
        print(f'=== {name} 快照回归 ===')
        ok = 0
        rows = []
        for key, old in sorted(snap.items()):
            cur = now.get(key)
            if cur is None:
                rows.append((key, '文件缺失（底稿不再生成！）', old, None))
                n_diff += 1
                continue
            diffs = []
            for k, lbl in (('aud', '审定'), ('det', '明细'), ('fn', '附注')):
                ov, cv = old.get(k), cur.get(k)
                if ov is None and cv is None:
                    continue
                if ov is None or cv is None:
                    diffs.append(f'{lbl}:{ov}→{cv}')
                elif abs(float(ov or 0) - float(cv or 0)) > TOL:
                    diffs.append(f'{lbl}:{ov:,.2f}→{cv:,.2f}')
            if diffs:
                rows.append((key, '；'.join(diffs), old, cur))
                n_diff += 1
            else:
                ok += 1
        new = [k for k in now if k not in snap]
        if new:
            rows.append((' + '.join(k.split('|')[-1] for k in new), '新增底稿（提示）', None, None))
        for key, msg, old, cur in rows:
            subj = key.split('|')[-1]
            print(f'  ✗ {subj}: {msg}')
        if new:
            print(f'  ✗ 新增 {len(new)} 份底稿（快照未含，确认后重新 capture）')
            n_diff += 1
        print(f'  ✓ 无变化 {ok} 科目；差异 {len(rows)} 科目')
    return 1 if n_diff else 0


def clear(folders):
    for folder in folders:
        sp = _snap_path(folder)
        if os.path.exists(sp):
            os.remove(sp)
            print(f'✓ 已删除快照：{sp}')
        else:
            print(f'（无快照可删）{sp}')
    return 0


def main():
    argv = sys.argv[1:]
    if len(argv) < 2 or argv[0] not in ('capture', 'check', 'clear'):
        print(__doc__)
        return 1
    cmd, folders = argv[0], argv[1:]
    if cmd == 'capture':
        return capture(folders)
    if cmd == 'check':
        return check(folders)
    return clear(folders)


if __name__ == '__main__':
    sys.exit(main())
