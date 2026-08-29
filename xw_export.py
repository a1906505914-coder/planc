# -*- coding: utf-8 -*-
"""#876：xlsxwriter 版 export_combined_excel（分阶段迁移）。

对照 current_account_detail.export_combined_excel，用 xlsxwriter 渲染：
  - 明细表 sheet（流式，66 万行不 OOM）——核心优化点
  - 对方科目核对 / Top10 等小表
  - 其他 sheet 逐步补齐（追加链路 build_all_in_dir 的 xw 化见后续阶段）

数据源 = export_combined_excel 同签名参数（comb_customers/comb_summary/comb_rows/...）。
阶段1 覆盖：明细表 + 对方科目核对。其余 sheet 以标题占位（TODO 阶段2）。
"""
import os
import xw_render as X
import xw_period_sheet as XP


def build_xw_export(output, periods, comb_customers, comb_summary, issues, comb_rows,
                    comb_counterparties, mode="Y", subj=None, by_ent=None, aging_map=None,
                    bucket_type=None, km_path=None, out_dir=None, data_dir=None, subj_key=None,
                    entities_dict=None, aging_methods=None, target_year=None, prior_meta=None):
    """xw 版合并底稿导出。参数同 export_combined_excel。返回 xlsxwriter wb（调用方 close）。"""
    label = subj["label"] if subj else "往来科目"
    periods = [p for p in periods if p != "全部"] or list(periods)
    _render_periods = [target_year] if target_year is not None else periods
    wb = X.new_workbook(output)
    # ---- 明细表（核心大表，流式）----
    for pk in _render_periods:
        ws = wb.add_worksheet(f"{label}明细表_{pk}")
        XP.render_period_sheet(wb, ws, pk, comb_customers, comb_summary, label,
                               bucket_type=bucket_type, prior_meta=prior_meta)
    # ---- 对方科目核对（按核算主体·分年度，无数据不生成）----
    ent_list = sorted({x.get("entity", "") for x in comb_rows if x.get("entity")})
    _by_ent = by_ent if by_ent is not None else _recon_by_ent(comb_rows)
    _has_cp = any(
        abs(amt) >= 1.0
        for _d in _by_ent.values()
        for _cpk in ("sale_cp", "collect_cp")
        for amt in (_d.get(_cpk, {}) or {}).values()
    )
    if _has_cp:
        ws = wb.add_worksheet("对方科目核对")
        _render_cp_sheet(wb, ws, label, _by_ent, _render_periods, ent_list)
    # TODO 阶段2：票据种类拆分 / 背书贴现 / 坏账准备 / Top10 / 主营客户比 / 内部交易抵消核对
    return wb


def _recon_by_ent(rows):
    """凭证翻面还原对方科目（简化版，供 xw 对方科目核对；与 openpyxl 版口径一致）。"""
    from collections import defaultdict
    from current_account_detail import _vouch_key_of
    sub_keys = set()
    for r in rows:
        if r.get('is_sub'):
            sub_keys.add(_vouch_key_of(r))
    vg = defaultdict(lambda: {"dr": 0.0, "cr": 0.0})
    flip = defaultdict(list)
    for r in rows:
        k = _vouch_key_of(r)
        if r.get('is_sub'):
            vg[k]["dr"] += float(r.get('debit') or 0.0)
            vg[k]["cr"] += float(r.get('credit') or 0.0)
        elif k in sub_keys:
            flip[k].append(r)
    by_ent = defaultdict(lambda: {"sale_cp": {}, "collect_cp": {}})
    for k, v in vg.items():
        ent = k[0]
        d = by_ent[ent]
        for r in flip.get(k, []):
            amt = float(r.get('credit') or 0.0)
            if amt > 0:
                nm = r.get('km') or '（未知）'
                d["sale_cp"][nm] = d["sale_cp"].get(nm, 0.0) + amt
            amt2 = float(r.get('debit') or 0.0)
            if amt2 > 0:
                nm = r.get('km') or '（未知）'
                d["collect_cp"][nm] = d["collect_cp"].get(nm, 0.0) + amt2
    return dict(by_ent)


def _render_cp_sheet(wb, ws, label, by_ent, periods, ent_list):
    """对方科目核对（xw 简单版：主体×方向×对方科目 金额/笔数）。"""
    rows_out = []
    for ent in sorted(by_ent):
        d = by_ent[ent]
        for nm, amt in sorted(d.get("sale_cp", {}).items(), key=lambda x: -x[1]):
            rows_out.append((ent, '借', nm, amt))
        for nm, amt in sorted(d.get("collect_cp", {}).items(), key=lambda x: -x[1]):
            rows_out.append((ent, '贷', nm, amt))
    X.dual_header(wb, ws, 0, [('对方科目核对', 4)], ['核算主体', '方向', '对方科目', '金额'],
                  widths=[14, 6, 40, 18], freeze_rows=2)
    r = 2
    for ent, side, nm, amt in rows_out:
        ws.write_string(r, 0, str(ent))
        ws.write_string(r, 1, side)
        ws.write_string(r, 2, str(nm))
        ws.write_number(r, 3, round(amt, 2), wb.add_format({'num_format': '#,##0.00'}))
        r += 1
    X.total_row(wb, ws, r, 4, ['合计', '', '', round(sum(x[3] for x in rows_out), 2)], money_cols={3})
