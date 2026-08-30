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
                    entities_dict=None, aging_methods=None, target_year=None, prior_meta=None,
                    tb_full=None, ac_ents=None):
    """xw 版合并底稿导出。参数同 export_combined_excel。返回 xlsxwriter wb（调用方 close）。"""
    label = subj["label"] if subj else "往来科目"
    periods = [p for p in periods if p != "全部"] or list(periods)
    _render_periods = [target_year] if target_year is not None else periods
    wb = X.new_workbook(output)
    # ---- P2 阶段3（第1步）：审定表（按核算主体；数据取 TB，置于明细表之前）----
    if tb_full and ac_ents and subj:
        try:
            import xw_audit_sheet as _XAS
            _title = f'{subj.get("label", label)} 审定表'
            _XAS.render_audit_summary_xw(
                wb, _title, subj.get("label", label), tb_full, ac_ents,
                target_year=target_year,
                # ⚡⚡ 2026-08-30 归并（用户定）：主 kw + ext_kw（应收股利→其他应收款 等）
                names=[k for k in ([subj.get("kw", "")] + list(subj.get("ext_kw") or [])) if k],
                is_credit=(subj.get("nature") == "liability"))
        except Exception as _ex:
            print(f'  ⚠️ 审定表 xw 失败：{_ex}')
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
        pass  # ⚡⚡ 2026-08-29 不再生成内部交易抵消核对（关联交易底稿由独立小程序生成，88 家核对底稿）
    # ---- P2 批次3：坏账准备计算表（AR/ORA）----
    if subj_key in ("AR", "ORA"):
        try:
            _render_baddebt_xw(wb, subj_key, label, _render_periods, comb_customers, comb_summary,
                               aging_map, bucket_type, entities_dict, target_year)
        except Exception as _ex:
            print(f'  ⚠️ 坏账准备 xw 失败：{_ex}')
    # ---- P2 批次2 接入：票据3表（ARN）+ 异常对应科目凭证清单 ----
    is_arn = subj and (subj.get("kw") == "应收票据" or subj.get("sheet") == "应收票据明细表")
    if is_arn:
        try:
            _render_bill_type_split_xw(wb, periods, comb_customers, comb_summary, label)
        except Exception as _ex:
            print(f'  ⚠️ 票据种类拆分 xw 失败：{_ex}')
        try:
            _render_endorse_xw(wb, comb_rows, label)
        except Exception as _ex:
            print(f'  ⚠️ 背书贴现 xw 失败：{_ex}')
        try:
            _render_note_baddebt_xw(wb, km_path, _render_periods, label)
        except Exception as _ex:
            print(f'  ⚠️ 应收票据坏账 xw 失败：{_ex}')
    try:
        from current_account_detail import _collect_bad_cps
        _bad_cps = _collect_bad_cps(_by_ent, subj, _render_periods)
        if _bad_cps:
            _render_badcp_voucher_xw(wb, _bad_cps, comb_rows, label)
    except Exception as _ex:
        print(f'  ⚠️ 异常凭证清单 xw 失败：{_ex}')
    # TODO 阶段3：附注汇总 / 追加链路（审定表注入/账龄/减值）
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


def _render_baddebt_xw(wb, subj_key, label, periods, comb_customers, comb_summary,
                       aging_map, bucket_type, entities_dict, target_year=None):
    """坏账准备计算表（xw，数据口径同 _write_baddebt_calc_sheet；数据全用变量计算，
    避免 openpyxl 版单元格回读——xlsxwriter 无法回读已写单元格）。"""
    from current_account_detail import AGING_BUCKETS, _BADDEBT_RATIO, _read_baddebt_tb, SUBJECTS
    buckets = AGING_BUCKETS.get(bucket_type, AGING_BUCKETS['recv'])
    ent_list = sorted(entities_dict) if entities_dict else sorted({e for (e, _n) in comb_customers})
    year = str(target_year if target_year is not None else (periods[-1] if periods else '2025'))
    kw = '其他应收' if subj_key == 'ORA' else None
    SUBJ_NM = '其他应收款' if subj_key == 'ORA' else '应收账款'
    bd = _read_baddebt_tb(entities_dict, year, kw)
    n = len(ent_list)
    NC = 4 + n   # A-D + 主体列 + 合计列

    def _aging(ent):
        out = {b: 0.0 for b in buckets}
        for (e, _nn) in comb_customers:
            if e != ent:
                continue
            am = aging_map.get(((e, _nn), year), {}) if aging_map else {}
            for b, v in am.items():
                out[b] = out.get(b, 0.0) + v
        return out

    comb_vals = [list(_aging(ent).values()) for ent in ent_list]   # [ent][bi]
    # ---- 块1 原值 ----
    orig_rows = []       # (bname, vals)
    for bi, bname in enumerate(buckets):
        orig_rows.append((bname, [comb_vals[ei][bi] for ei in range(n)]))
    combo_sub = [sum(comb_vals[ei]) for ei in range(n)]
    single_sub = [0.0] * n
    assoc_sub = [0.0] * n
    orig_tot = [combo_sub[ei] + single_sub[ei] for ei in range(n)]
    # ---- 块2 坏账准备应有 ----
    bd_rows = []
    for bi, bname in enumerate(buckets):
        ratio = _BADDEBT_RATIO[bi] if bi < len(_BADDEBT_RATIO) else 1.0
        bd_rows.append((bname, ratio, [round(comb_vals[ei][bi] * ratio, 2) for ei in range(n)]))
    bd_sub = [round(sum(bd_rows[bi][2][ei] for bi in range(len(buckets))), 2) for ei in range(n)]
    bd_single = [0.0] * n
    bd_tot = [bd_sub[ei] + bd_single[ei] for ei in range(n)]
    # ---- 块3 净值 ----
    net_rows = []
    for bi, bname in enumerate(buckets):
        net_rows.append((bname, [round(comb_vals[ei][bi] - bd_rows[bi][2][ei], 2) for ei in range(n)]))
    net_sub = [round(sum(net_rows[bi][1][ei] for bi in range(len(buckets))), 2) for ei in range(n)]
    net_single = [round(single_sub[ei] - bd_single[ei], 2) for ei in range(n)]
    net_tot = [net_sub[ei] + net_single[ei] for ei in range(n)]
    # ---- 块4 调整 ----
    unaud = [round(bd.get(ent, {}).get('close', 0.0), 2) for ent in ent_list]
    cont = [0.0] * n
    adj = [round(bd_tot[ei] - unaud[ei] - cont[ei], 2) for ei in range(n)]
    tb = [round(bd.get(ent, {}).get('close', 0.0), 2) for ent in ent_list]
    diff = [round(bd_tot[ei] - tb[ei], 2) for ei in range(n)]

    ws = wb.add_worksheet("坏账准备")
    X.title(wb, ws, f"{label}—坏账准备计算表（参照 112500-7-6 格式；账龄按 GL 交易日期 FIFO 测算）", NC)
    hdr = ['项目', '', '账   龄', '计提比例'] + ent_list + ['合计']
    r = X.dual_header(wb, ws, 1, [("", 2), ("", 2), (SUBJ_NM, n), ("", 1)], hdr,
                      widths=[22, 26, 10, 10] + [14] * n + [16], freeze_rows=4)
    mcols = set(range(4, NC))
    gray = {j: 'D9D9D9' for j in range(NC)}
    # 块1 原值
    for bi, (bname, vals) in enumerate(orig_rows):
        X.row(wb, ws, r, [SUBJ_NM + '原值' if bi == 0 else '', SUBJ_NM + '原值-按账龄组合' if bi == 0 else '',
                          bname, None] + vals + [round(sum(vals), 2)], money_cols=mcols)
        r += 1
    X.row(wb, ws, r, ['', '', '小计', None] + combo_sub + [round(sum(combo_sub), 2)], money_cols=mcols, fills=gray)
    r += 1
    for bi, bname in enumerate(buckets):
        X.row(wb, ws, r, [SUBJ_NM + '原值-单项计提' if bi == 0 else '', '', bname, None] + [0.0] * n + [0.0], money_cols=mcols)
        r += 1
    X.row(wb, ws, r, ['', '', '小计', None] + single_sub + [0.0], money_cols=mcols, fills=gray)
    r += 1
    for bi, bname in enumerate(buckets):
        X.row(wb, ws, r, [SUBJ_NM + '合并内关联方' if bi == 0 else '', '', bname, None] + [0.0] * n + [0.0], money_cols=mcols)
        r += 1
    X.row(wb, ws, r, ['', '', '小计', None] + assoc_sub + [0.0], money_cols=mcols, fills=gray)
    r += 1
    X.row(wb, ws, r, [SUBJ_NM + '原值合计', '合计', '', None] + orig_tot + [round(sum(orig_tot), 2)], money_cols=mcols, fills=gray)
    r += 1
    # 块2 坏账准备应有
    for bi, (bname, ratio, vals) in enumerate(bd_rows):
        X.row(wb, ws, r, ['坏账准备余额' if bi == 0 else '', '坏账准备应有余额-按账龄组合' if bi == 0 else '',
                          bname, ratio] + vals + [round(sum(vals), 2)], money_cols=mcols)
        r += 1
    X.row(wb, ws, r, ['', '', '小计', None] + bd_sub + [round(sum(bd_sub), 2)], money_cols=mcols, fills=gray)
    r += 1
    X.row(wb, ws, r, ['', '坏账准备-单项计提', '', None] + bd_single + [0.0], money_cols=mcols)
    r += 1
    X.row(wb, ws, r, ['坏账准备应有余额合计', '合计', '', None] + bd_tot + [round(sum(bd_tot), 2)], money_cols=mcols, fills=gray)
    r += 1
    # 块3 净值
    for bi, (bname, vals) in enumerate(net_rows):
        X.row(wb, ws, r, [SUBJ_NM + '净值' if bi == 0 else '', SUBJ_NM + '净值-按账龄组合' if bi == 0 else '',
                          bname, None] + vals + [round(sum(vals), 2)], money_cols=mcols)
        r += 1
    X.row(wb, ws, r, ['', '', '小计', None] + net_sub + [round(sum(net_sub), 2)], money_cols=mcols, fills=gray)
    r += 1
    X.row(wb, ws, r, ['', SUBJ_NM + '净值-单项计提', '', None] + net_single + [round(sum(net_single), 2)], money_cols=mcols)
    r += 1
    X.row(wb, ws, r, [SUBJ_NM + '净值合计', '合计', '', None] + net_tot + [round(sum(net_tot), 2)], money_cols=mcols, fills=gray)
    r += 2
    ws.write_string(r, 0, '坏账准备调整')
    r += 1
    X.row(wb, ws, r, ['', '坏账准备未审数', '', None] + unaud + [round(sum(unaud), 2)], money_cols=mcols)
    r += 1
    X.row(wb, ws, r, ['', '期初续调坏账准备', '', None] + cont + [0.0], money_cols=mcols)
    r += 1
    X.row(wb, ws, r, ['', '本期调整数', '', None] + adj + [round(sum(adj), 2)], money_cols=mcols, fills=gray)
    r += 1
    X.row(wb, ws, r, ['', 'TB', '', None] + tb + [round(sum(tb), 2)], money_cols=mcols)
    r += 1
    X.row(wb, ws, r, ['', '差异（应有余额-账套）', '', None] + diff + [round(sum(diff), 2)], money_cols=mcols)
    ws.freeze_panes(4, 4)
