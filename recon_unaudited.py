# -*- coding: utf-8 -*-
"""未审数三表对平核对 v2（2026-08-07 用户要求：审计基础=未审口径三表一致）。
基准：自建TB 未审（叶子聚合）。
对比：①审定表『期末未审数/期末数』列 ②明细表合计 ③附注『期末数』段（合并口径，参考）。
规则：abs 对齐（贷余负 vs 正数显示）、往来重分类/净额口径单独标注、空壳（一方为0）单列。
输出分级：A=真差异(需要处理) / B=口径差(符号/重分类/净额) / C=空壳(底稿未生成)。
"""
import paths as P
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from tb_recon_detail import _sheet_end_total, _footnote_total, _sheet_ent_values


def _is_mixed_year(ws):
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


def _extract_aud(ws, year=None):
    """审定表『期末未审数/期末数』列值（未审口径）。"""
    v, why = _sheet_end_total(ws)
    if year and _is_mixed_year(ws):
        e = _sheet_ent_values(ws, year=year)
        if e:
            v = sum(e.values())
    return v, why


def _extract_det(ws, year=None):
    """明细表合计（未审口径）。"""
    v, why = _sheet_end_total(ws)
    if v is None or abs(v) < 0.005:
        e = _sheet_ent_values(ws, year=year) if (year and _is_mixed_year(ws)) else _sheet_ent_values(ws)
        if e:
            v = sum(e.values())
    return v


def recon_unaudited2(data_dir, years=None):
    import audit_common
    import self_tb_gen as S
    import subjects_registry as REG
    import openpyxl

    entities = audit_common.discover_entities(data_dir)
    tb_full = audit_common.read_tb_full(data_dir, entities)
    if not tb_full:
        return True, []
    t_years = sorted({str(yy) for (_e, _c, _n, yy) in tb_full}) or ['2025']
    if years is None:
        years = t_years
    ent_tb_all = {}
    for ent in sorted(entities):
        ent_tb_all[ent] = {}
        for yy in years:
            items = [(c, v, n) for (e, c, n, y), v in tb_full.items()
                     if e == ent and str(y) == yy]
            leafs = S.leaf_codes([c for c, _v, _n in items])
            l1_names = {}
            for c, _v, n in items:
                l1 = re.sub(r'\D', '', str(c))[:4]
                if l1 and l1 not in l1_names:
                    segs = [s for s in re.split(r'[\\/\-]', str(n))
                            if s.strip() and not re.match(r'^\d+$', s.strip())]
                    l1_names[l1] = segs[0].strip() if segs else str(n)
            ent_tb_all[ent][yy] = {c: dict(v, name=n,
                                           l1_name=l1_names.get(re.sub(r'\D', '', str(c))[:4], ''))
                                   for c, v, n in items if c in leafs}

    diffs = []
    for key, subj in REG.REGISTRY.items():
        codes = subj.get('codes') or []
        if not codes:
            continue
        if any(str(c)[:1] in '678' for c in codes):
            continue
        fk = subj['file_key']
        subj_name = subj.get('name', '')
        _excl_kw = ('清理', '待处理', '待处置', '受托', '代管')
        _contra_kw = ('累计折旧', '累计摊销', '减值准备', '跌价准备', '坏账准备',
                      '未确认融资费用', '使用权资产折旧', '折旧', '摊销')
        _contra_code_prefix = ('1231', '1471', '1472', '1602', '1603',
                               '1702', '1703', '1522', '1523', '1632')

        for yy in years:
            fp = os.path.join(data_dir, f'{fk}审计底稿_{yy}_生成.xlsx')
            if not os.path.exists(fp):
                continue
            merged = {}
            for ent in sorted(ent_tb_all):
                for c, v in ent_tb_all[ent].get(str(yy), {}).items():
                    key2 = (str(c), str(v.get('name', '')))
                    if key2 in merged:
                        merged[key2]['qm'] += v['qm']
                    else:
                        merged[key2] = dict(v)
            if not merged:
                continue

            def _name_hit(n, l1n, sname):
                def _hit(s, sn):
                    if not s or not sn:
                        return False
                    if s.startswith('提取') or '应付现金股利' in s or '应付股利' in s:
                        return False
                    if s == sn:
                        return True
                    if sn in s and (s.startswith(sn) or s.endswith(sn)) and len(s) - len(sn) <= 4:
                        return True
                    return False

                def _rev_hit(s, sn):
                    if not s or not sn:
                        return False
                    if s.startswith('提取') or '应付现金股利' in s or '应付股利' in s:
                        return False
                    if s in ('其他', '本年', '以前', '待处理', '上期', '本期', '期初', '期末'):
                        return False
                    return s in sn
                if not sname:
                    return False
                cand = [str(n)]
                if l1n:
                    cand.append(str(l1n))
                for s in cand:
                    for seg in re.split(r'[\\/\-]', s):
                        if _hit(seg.strip(), sname):
                            return True
                for s in cand:
                    for seg in re.split(r'[\\/\-]', s):
                        if _rev_hit(seg.strip(), sname):
                            return True
                return False

            s_g_n = s_c_n = 0.0
            rows_g_n, rows_c_n = [], []
            for (c, _n0), v in merged.items():
                n = str(v.get('name', ''))
                if any(k in n for k in _excl_kw):
                    continue
                if any(str(c).startswith(p) for p in _contra_code_prefix):
                    continue
                is_contra = any(k in n for k in _contra_kw)
                hit = _name_hit(n, str(v.get('l1_name', '')), subj_name)
                if hit and not is_contra:
                    s_g_n += v['qm']; rows_g_n.append(str(c))
                if codes and any(str(c).startswith(k) for k in codes) and not is_contra:
                    s_c_n += v['qm']; rows_c_n.append(str(c))

            def _pick(s_n, s_c, rn, rc):
                if not rn:
                    return s_c
                if not rc:
                    return s_n
                if len(codes) > 1 and (set(rn) & set(rc) or abs(s_c) > abs(s_n)):
                    return s_c
                return s_n
            tb_v = _pick(s_g_n, s_c_n, rows_g_n, rows_c_n)
            if abs(tb_v) <= 0.005:
                continue
            # 底稿三处
            try:
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
                sheets = wb.sheetnames
                aud_sn = next((s for s in sheets if '审定表' in s), None)
                det_sn = next((s for s in sheets if ('明细表' in s or '分类汇总' in s)
                               and not any(k in s for k in ('附注', '凭证', '账龄', 'Top'))), None)
                aud = det = None
                if aud_sn:
                    aud, _ = _extract_aud(wb[aud_sn], year=yy)
                if det_sn:
                    det = _extract_det(wb[det_sn], year=yy)
                wb.close()
            except Exception:
                aud = det = None
            # 分级判断（abs 对齐）
            at = abs(tb_v)
            problems = []
            if aud is not None and abs(abs(aud) - at) > max(0.01, at * 0.0001):
                problems.append(f'审定={aud:,.2f}')
            if det is not None and abs(abs(det) - at) > max(0.01, at * 0.0001):
                problems.append(f'明细={det:,.2f}')
            if not problems:
                continue
            # 分级：空壳=审定或明细一方为0且另一方非0
            is_empty = (aud is not None and abs(aud) < 0.005 and at > 0.005) \
                or (det is not None and abs(det) < 0.005 and at > 0.005)
            diffs.append((subj_name, fk, str(yy), round(tb_v, 2),
                          round(aud, 2) if aud is not None else None,
                          round(det, 2) if det is not None else None,
                          '；'.join(problems), '空壳' if is_empty else '差异'))
    return (not diffs), diffs


def main():
    dirs = sys.argv[1:] or [os.path.join(P.DATA_ROOT, a)
                            for a in ('ACB', 'GJX', 'ARF', 'AYL', 'AJJ', 'ADZ', 'AJ', 'AQ', 'AFJ', 'AS', 'ATT', 'AZ', 'GFY')]
    grand = {'差异': 0, '空壳': 0}
    for d in dirs:
        ok, diffs = recon_unaudited2(d)
        nc = sum(1 for x in diffs if x[7] == '空壳')
        nd = len(diffs) - nc
        grand['空壳'] += nc
        grand['差异'] += nd
        print(f'===== {os.path.basename(d.rstrip(chr(92)))} ===== 未审数对平：'
              f'{"✅ 一致" if not diffs else f"差异 {nd} 项 / 空壳 {nc} 项"}')
        for name, fk, yy, tb, aud, det, why, kind in diffs[:25]:
            print(f'  [{kind}] {name}({fk}){yy}: TB未审={tb:,.2f} | {why}')
        if len(diffs) > 25:
            print(f'    … 共 {len(diffs)} 项')
    print(f'\n总计：差异 {grand["差异"]} 项 / 空壳 {grand["空壳"]} 项')
    return 0


if __name__ == '__main__':
    sys.exit(main())
