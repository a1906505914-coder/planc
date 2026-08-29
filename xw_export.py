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
    # ---- P2 批次1（2026-08-29）：Top10 余额 + 内部交易抵消核对（xw，数据口径同 openpyxl 版）----
    if subj_key in ("AR", "ORA", "APR", "AP"):
        try:
            _render_top10_xw(wb, periods, comb_customers, comb_summary, label, subj_key)
        except Exception as _ex:
            print(f'  ⚠️ Top10 xw 失败：{_ex}')
    if len(ent_list) > 1:
        try:
            _render_intra_group_xw(wb, ent_list, comb_customers, comb_summary, _render_periods, label)
        except Exception as _ex:
            print(f'  ⚠️ 内部交易抵消核对 xw 失败：{_ex}')
    # TODO 阶段2：票据种类拆分 / 背书贴现 / 坏账准备 / 主营客户比 / 附注汇总 / 追加链路（审定表注入）
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


# ================= P2 批次1：Top10 余额 + 内部交易抵消核对（xw） =================
def _render_top10_xw(wb, periods, comb_customers, comb_summary, label, subj_key):
    """Top10 余额（xw 版，数据口径与 current_account_detail._write_top10_sheet 一致）。"""
    from current_account_detail import _base_counterparty, _wtype_cat, SUBJECTS
    from collections import defaultdict
    periods = sorted(periods)
    p1 = periods[-1]
    p0 = periods[-2] if len(periods) >= 2 else None
    subj_nature = SUBJECTS.get(subj_key, {}).get("nature", "asset")

    def sclose(en, pk):
        return comb_summary.get((en, pk), {}).get("s_close", 0.0)

    def aggregate(cust_list):
        agg = {}
        for en in cust_list:
            base = _base_counterparty(en[1])
            s = comb_summary.get((en, p1), {})
            wc = _wtype_cat(subj_nature)
            c1 = s.get("s_close", 0.0) or 0.0
            c0 = sclose(en, p0) if p0 else 0.0
            d = agg.get(base)
            if d is None:
                d = agg[base] = {"entity": en[0], "ents": set([en[0]]),
                                 "close_p1": 0.0, "close_p0": 0.0,
                                 "nature": s.get("nature", ""), "wtype": wc,
                                 "related": False, "maxabs": 0.0}
            d["close_p1"] += c1
            d["close_p0"] += c0
            d["ents"].add(en[0])
            d["related"] = d["related"] or bool(s.get("related"))
            aa = abs(c1)
            if aa >= d["maxabs"]:
                d["maxabs"] = aa
                d["nature"] = s.get("nature", "")
                d["entity"] = en[0]
        return agg

    hdr = (["排名", "核算主体", "往来单位名称", "往来类型", "款项性质", "是否关联方", f"期末余额({p1})"]
           + ([f"期末余额({p0})", "变动"] if p0 else []))
    ncols = len(hdr)
    money_cols = tuple(range(6, ncols))
    ws = wb.add_worksheet("Top10余额")
    X.title(wb, ws, f"{label} — 余额前10位明细表", ncols)
    r = X.dual_header(wb, ws, 1, [("", ncols)], hdr,
                      widths=[6, 22, 30, 12, 16, 10, 18] + ([18, 16] if p0 else []), freeze_rows=5)
    # 一、分主体 Top10
    ws.write_string(r, 0, "一、分核算主体 余额前10位（已按往来单位汇总同名跨子目，按期末余额绝对值）")
    r += 1
    by_ent = defaultdict(list)
    for en in comb_customers:
        by_ent[en[0]].append(en)
    for entity in sorted(by_ent.keys()):
        agg = aggregate(by_ent[entity])
        custs = sorted(agg.items(), key=lambda kv: abs(kv[1]["close_p1"]), reverse=True)[:10]
        for rank, (base, d) in enumerate(custs, 1):
            row = [rank, entity, base, _wtype_cat(subj_nature), d["nature"],
                   ("是" if d["related"] else "否"), round(d["close_p1"], 2)]
            if p0:
                row += [round(d["close_p0"], 2), round(d["close_p1"] - d["close_p0"], 2)]
            X.row(wb, ws, r, row, money_cols=money_cols, center_cols=(0, 3, 5))
            r += 1
    r += 1
    # 二、全集团 Top10
    ws.write_string(r, 0, "二、全集团 余额前10位（已按往来单位汇总同名跨子目，按期末余额绝对值）")
    r += 1
    grp_agg = aggregate(comb_customers)
    grp_top = sorted(grp_agg.items(), key=lambda kv: abs(kv[1]["close_p1"]), reverse=True)[:10]
    for rank, (base, d) in enumerate(grp_top, 1):
        ent_disp = d["entity"] if len(d["ents"]) == 1 else "/".join(sorted(d["ents"]))
        row = [rank, ent_disp, base, _wtype_cat(subj_nature), d["nature"],
               ("是" if d["related"] else "否"), round(d["close_p1"], 2)]
        if p0:
            row += [round(d["close_p0"], 2), round(d["close_p1"] - d["close_p0"], 2)]
        X.row(wb, ws, r, row, money_cols=money_cols, center_cols=(0, 3, 5))
        r += 1
    tot1 = sum(d["close_p1"] for _, d in grp_top)
    if p0:
        tot0 = sum(d["close_p0"] for _, d in grp_top)
        X.total_row(wb, ws, r, ncols, ["", "合计(Top10)", "", "", "", "", round(tot1, 2),
                                       round(tot0, 2), round(tot1 - tot0, 2)], money_cols=set(money_cols))
    else:
        X.total_row(wb, ws, r, ncols, ["", "合计(Top10)", "", "", "", "", round(tot1, 2)],
                    money_cols=set(money_cols))


def _render_intra_group_xw(wb, ent_list, comb_customers, comb_summary, periods, label):
    """内部交易抵消核对表（xw 版，集团模式专属；数据口径与 _write_intra_group_recon_sheet 一致）。"""
    _pk = periods[-1] if periods else None
    if _pk is None or not ent_list:
        return
    cust_ent = {}
    for (_e, _nm) in comb_customers:
        s = comb_summary.get(((_e, _nm), _pk))
        if s and abs(s.get('s_close', 0.0)) >= 0.005:
            cust_ent.setdefault(_nm, {})[_e] = s['s_close']
    _tot = {cc: sum(v.values()) for cc, v in cust_ent.items()}
    _ordered = sorted(cust_ent, key=lambda x: -abs(_tot[x]))[:300]
    headers = ['往来单位'] + ent_list + ['全集团合计', '抵消借', '抵消贷', '合并报表数']
    ncols = len(headers)
    ws = wb.add_worksheet("内部交易抵消核对")
    X.title(wb, ws, f'{label}—内部交易抵消核对表（{_pk}）· 同一往来单位跨主体出现=内部交易候选，抵消待填', ncols)
    r = X.dual_header(wb, ws, 1, [("", ncols)], headers,
                      widths=[28] + [14] * (ncols - 1), freeze_rows=3)
    _ent_end = 1 + len(ent_list)
    hl = wb.add_format({'bg_color': '#FFF2CC', 'border': 1})
    mid = wb.add_format({'num_format': '#,##0.00', 'bg_color': '#BDD7EE', 'border': 1})
    for cc in _ordered:
        rec = cust_ent[cc]
        n_ent = sum(1 for v in rec.values() if abs(v) >= 0.005)
        tot = sum(rec.values())
        vals = [cc] + [round(rec.get(e, 0.0), 2) if abs(rec.get(e, 0.0)) >= 0.005 else None for e in ent_list] \
               + [round(tot, 2) if abs(tot) >= 0.005 else None, None, None,
                  round(tot, 2) if abs(tot) >= 0.005 else None]
        X.row(wb, ws, r, vals, money_cols=set(range(1, ncols)), center_cols=())
        if n_ent >= 2:      # 跨主体 → 内部交易候选标黄
            for j in range(ncols):
                ws.write_string(r, j, str(vals[j]) if vals[j] is not None else '',
                                hl if j == 0 else wb.add_format({'bg_color': '#FFF2CC', 'border': 1}))
        r += 1
    ws.freeze_panes(3, 2)


# ================= P2 批次2：票据3表 + 异常清单（xw） =================
def _render_bill_type_split_xw(wb, periods, comb_customers, comb_summary, label):
    """票据种类拆分（xw，数据口径同 _write_bill_type_split）。"""
    def _type_of(s):
        return (s.get("bill_type") or "未分类") if s else "未分类"
    types = []
    for pk in periods:
        for ak in comb_customers:
            t = _type_of(comb_summary.get((ak, pk)))
            if t not in types:
                types.append(t)
    ordered = [t for t in ["银行承兑汇票", "商业承兑汇票", "财务公司承兑汇票", "未分类"] if t in types] + \
              [t for t in types if t not in ["银行承兑汇票", "商业承兑汇票", "财务公司承兑汇票", "未分类"]]
    ws = wb.add_worksheet("票据种类拆分")
    X.title(wb, ws, f"{label}—票据种类拆分（银行承兑汇票 / 商业承兑汇票 按辅助核算子目归类）", 4)
    r = X.dual_header(wb, ws, 1, [("", 4)], ["票据种类", "期间", "客户数", "本币期末余额"],
                      widths=[20, 10, 10, 20], freeze_rows=3)
    r += 1
    ws.write_string(r, 0, "按客户明细（票据种类 / 客户 / 期间 / 本币期末余额）")
    r += 1
    for j, h in enumerate(["票据种类", "客户", "期间", "本币期末余额"]):
        ws.write_string(r, j, h, wb.add_format({'bold': True, 'bg_color': '#DDEBF7'}))
    r += 1
    for t in ordered:
        for pk in periods:
            cnt = 0
            tot = 0.0
            for ak in comb_customers:
                s = comb_summary.get((ak, pk))
                if s and _type_of(s) == t:
                    cnt += 1
                    tot += s.get("s_close", 0.0)
            X.row(wb, ws, r, [t, pk, cnt, round(tot, 2)], money_cols=(3,))
            r += 1
    for ak in comb_customers:
        from current_account_detail import _disp_name
        nm = _disp_name(ak)
        for pk in periods:
            s = comb_summary.get((ak, pk))
            if not s:
                continue
            X.row(wb, ws, r, [_type_of(s), nm, pk, round(s.get("s_close", 0.0), 2)], money_cols=(3,))
            r += 1


def _render_endorse_xw(wb, rows, label):
    """背书 / 贴现 / 承兑 明细（xw，数据口径同 _write_endorse_discount）。"""
    hdr = ["序号", "日期", "字", "号", "往来单位名称", "摘要", "对方科目", "借方金额", "贷方金额",
           "业务类型", "出票日", "到期日"]
    ws = wb.add_worksheet("背书贴现明细")
    X.title(wb, ws, f"{label}—背书 / 贴现 / 承兑 明细（数据源自综合查询明细表(GL) 摘要字段）", len(hdr))
    r = X.dual_header(wb, ws, 1, [("", len(hdr))], hdr,
                      widths=[6, 12, 6, 10, 20, 30, 18, 14, 14, 8, 10, 10], freeze_rows=3)

    def _biz(t):
        t = t or ""
        if "背书" in t:
            return "背书"
        if "贴现" in t:
            return "贴现"
        if "承兑" in t:
            return "承兑"
        return "其他"
    idx = 1
    for x in rows:
        if not x.get("is_sub"):
            continue
        biz = _biz(x.get("summary", ""))
        if biz == "其他":
            continue
        d = x.get("date")
        ds = d.strftime("%Y-%m-%d") if d else ""
        X.row(wb, ws, r, [idx, ds, x.get("vtype", ""), x.get("vno", ""), x.get("cust", ""),
                          x.get("summary", ""), x.get("cp", ""),
                          round(float(x.get("debit") or 0), 2), round(float(x.get("credit") or 0), 2),
                          biz, None, None], money_cols=(7, 8), center_cols=(0, 1, 2, 3))
        r += 1
        idx += 1


def _render_note_baddebt_xw(wb, km_path, periods, label):
    """应收票据坏账准备（xw，数据口径同 _write_note_baddebt）。"""
    from current_account_detail import _read_baddebt_for_notes
    hdr = ["期间", "期初余额", "本期计提(贷)", "本期转回(借)", "期末余额", "备注"]
    ws = wb.add_worksheet("应收票据坏账准备")
    X.title(wb, ws, f"{label}—坏账准备（应收票据相关）", len(hdr))
    bd = _read_baddebt_for_notes(km_path, periods)
    if not bd:
        r = X.dual_header(wb, ws, 1, [("", len(hdr))], hdr, widths=[12, 18, 18, 18, 18, 24], freeze_rows=3)
        ws.write_string(r, 0, "（账套科目余额表中未发现与应收票据相关的坏账准备"
                              "（科目名须同时含『坏账准备』与『应收票据』）。"
                              "按审计要求，相关余额暂留空，待账套补充或手工填列。）")
        return
    r = X.dual_header(wb, ws, 1, [("", len(hdr))], hdr, widths=[12, 18, 18, 18, 18, 24], freeze_rows=3)
    for pk in periods:
        rec = bd.get(pk)
        if not rec:
            X.row(wb, ws, r, [pk] + [None] * (len(hdr) - 1), money_cols=(1, 2, 3, 4))
        else:
            X.row(wb, ws, r, [pk, round(rec["open"], 2), round(rec["credit"], 2),
                              round(-rec["debit"], 2), round(rec["close"], 2), ""], money_cols=(1, 2, 3, 4))
        r += 1


def _render_badcp_voucher_xw(wb, bad_cps, comb_rows, label):
    """异常对应科目凭证清单（xw，数据口径同 _write_bad_cp_voucher_sheet：遍历凭证明细，
    对方科目 cp 命中 bad_cps 异常科目 → 列出行，供审计员按凭证号直接定位查证）。"""
    hdr = ["核算主体", "方向", "异常/关注对应科目", "结论", "凭证号", "日期", "往来单位", "摘要", "金额(元)"]
    ws = wb.add_worksheet("异常对应科目凭证清单")
    X.title(wb, ws, f"{label}—异常对应科目凭证清单（列示『对方科目核对』标异常/关注的凭证）", len(hdr))
    r = X.dual_header(wb, ws, 1, [("", len(hdr))], hdr,
                      widths=[10, 6, 22, 20, 12, 12, 20, 34, 14], freeze_rows=3)
    n = 0
    for x in comb_rows:
        if not x.get("is_sub"):
            continue
        cpv = str(x.get("cp") or "").strip()
        if not cpv:
            continue
        dr = float(x.get("debit") or 0.0)
        cr = float(x.get("credit") or 0.0)
        for nm in cpv.split(","):
            nm = nm.strip()
            if not nm or nm not in bad_cps:
                continue
            _v = bad_cps[nm]
            reason = _v[1] if isinstance(_v, tuple) else str(_v)
            d = x.get('date')
            ds = d.strftime('%Y-%m-%d') if d else ''
            X.row(wb, ws, r, [x.get('entity', ''), '借' if dr > 0 else '贷', nm, reason,
                              x.get('vno', ''), ds, x.get('cust', ''), x.get('summary', ''),
                              round(dr if dr else cr, 2)], money_cols=(8,))
            r += 1
            n += 1
            break
    return n
