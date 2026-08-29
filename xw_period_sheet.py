# -*- coding: utf-8 -*-
"""#876：xlsxwriter 版往来明细表渲染（对照 _cwrite_period_sheet，流式写大表）。

用法（集成后）：export_combined_excel 在 xw 模式下调用本函数替代 _cwrite_period_sheet。
本函数只负责"明细表" sheet 的渲染；不依赖 openpyxl。
"""
import xw_render as X


def render_period_sheet(wb, ws, period, customers, summary, label,
                        bucket_type=None, prior_meta=None):
    """渲染明细表（xlsxwriter 流式）。customers=[(entity,name)]，summary={((e,n),pk): rec}。"""
    # 表头结构对齐 openpyxl 版
    buckets = []
    if bucket_type == 'recv':
        buckets = ['1年以内', '1-2年', '2-3年', '3-4年', '4-5年', '5年以上']
    _has_con = any('@@' in str(n) for (_e, n) in customers)
    _off = 1 if _has_con else 0
    base_headers = ["核算主体", "序号", "往来单位编号", "往来单位名称", "款项性质", "是否关联方", "币种",
                    "最上层客户单位", "客户性质", "坏账计提方法",
                    "外币余额(期初)", "期初余额(人民币)", "本期借方发生额", "本期贷方发生额",
                    "外币余额(期末)", "期末余额(人民币)", "重分类调整", "审计调整", "审定数"]
    if _has_con:
        base_headers = base_headers[:4] + ["合同号"] + base_headers[4:]
    if buckets:
        headers = base_headers + buckets + ["交易笔数", "平衡校验(期初+借+贷-期末)", "期后收款"]
        nb = len(buckets)
    else:
        headers = base_headers + ["账龄区间", "交易笔数", "平衡校验(期初+借+贷-期末)", "期后收款"]
        nb = 0
    _base_cnt = 10 + _off
    _groups = [('基础信息', _base_cnt), ('期初', 2), ('本期发生额', 2), ('期末', 2), ('调整', 3)]
    if buckets:
        _groups.append(('账龄', nb))
    _groups.append(('其他', 3))
    data_r = X.dual_header(wb, ws, 3, _groups, headers, freeze_rows=5)   # 表头行3/4，数据从行5(0-based)
    # 列分组底色 + 金额格式
    money_cols = set(range(10 + _off, 19 + _off))
    if nb:
        money_cols |= set(range(19 + _off, 19 + _off + nb))
    money_cols.add(26 + _off)   # 平衡校验
    fills = X.period_col_fills(wb, _off, nb)
    center_cols = {1, 4 + _off, 5 + _off, 6 + _off, 7 + _off, 8 + _off, 9 + _off}
    formula_cols = {18 + _off}  # 审定数
    # 预构建每列 format（避免每格新建）：金额列 / 居中列 / 文本列
    num_f = wb.add_format({'num_format': '#,##0.00'})
    num_border = wb.add_format({'num_format': '#,##0.00', 'border': 1})
    txt_border = wb.add_format({'border': 1})
    ctr_border = wb.add_format({'border': 1, 'align': 'center'})
    col_fmt = {}
    for c in range(len(headers)):
        if c in money_cols:
            col_fmt[c] = fills.get(c, num_border)
        elif c in center_cols:
            col_fmt[c] = ctr_border
        else:
            col_fmt[c] = txt_border
    tot = {"open": 0.0, "debit": 0.0, "credit": 0.0, "close": 0.0, "count": 0, "fopen": 0.0, "fclose": 0.0}
    r = data_r
    for idx, (entity, name) in enumerate(customers, 1):
        s = summary.get(((entity, name), period))
        if s is None:
            continue
        bal = (s["s_open"] + s["s_debit"] + s["s_credit"] - s["s_close"]
               if name not in ("（无辅助核算明细）", "（未标注往来单位）") else None)
        ccy = "外币" if s.get("foreign") else "人民币"
        f_open = s.get("fcur_open", 0.0) or 0.0
        f_close = s.get("fcur_close", 0.0) or 0.0
        _nm = name
        _con = ''
        if _has_con and '@@' in str(name):
            _nm, _con = str(name).split('@@', 1)
        pm = ((prior_meta or {}).get((entity, name))
              or (prior_meta or {}).get(('*', name)) or {}) or {}
        top_unit = pm.get('top_unit', '')
        cust_nature = pm.get('nature', '')
        baddebt_method = pm.get('method', '账龄计提')
        _rc = 0.0
        if (s["s_close"] or 0.0) < -0.005:
            _rc = round(-(s["s_close"] or 0.0), 2)
        _rc_in = round(float(s.get("reclass_in", 0.0) or 0.0), 2)
        _rc_tot = _rc + _rc_in
        base_vals = [entity, idx, s["code"], _nm]
        if _has_con:
            base_vals += [_con]
        base_vals += [s.get("nature", ""), ("是" if s.get("related") else "否"), ccy,
                      top_unit, cust_nature, baddebt_method,
                      (f_open if f_open else None), s["s_open"], abs(s["s_debit"]),
                      abs(s["s_credit"]), (f_close if f_close else None), s["s_close"],
                      (_rc_tot if _rc_tot else None), None]
        # 审定数公式 =F{row}+G{row}+H{row}（0-based 列 -> xlsx 字母）
        from openpyxl.utils import get_column_letter as _gcl
        rr = r + 1   # xlsx 行号（1-based）
        _f_end = _gcl(15 + _off); _f_rc = _gcl(16 + _off); _f_adj = _gcl(17 + _off)
        base_vals += ["=%s%d+%s%d+%s%d" % (_f_end, rr, _f_rc, rr, _f_adj, rr)]
        if buckets:
            bm = {b: 0.0 for b in buckets}
            bucket_vals = [round(bm.get(b, 0.0), 2) for b in buckets]
            after = pm.get('after', None)
            vals = base_vals + bucket_vals + [s["count"], (None if bal is None else round(bal, 2)), after]
        else:
            aging = s["aging"] if s["aging"] else "（未提供）"
            after = pm.get('after', None)
            vals = base_vals + [aging, s["count"], (None if bal is None else round(bal, 2)), after]
        # 0 值金额留空
        _row = list(vals)
        for ci in money_cols:
            if ci < len(_row):
                v = _row[ci]
                if isinstance(v, (int, float)) and abs(v) <= 0.005:
                    _row[ci] = None
        # 流式写行
        for c, v in enumerate(_row):
            if v is None:
                continue
            if c in formula_cols:
                ws.write_formula(r, c, v, col_fmt[c])
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                ws.write_number(r, c, v, col_fmt[c])
            else:
                ws.write_string(r, c, str(v), col_fmt[c])
        if bal is not None and abs(bal) > 0.005:
            ws.write_blank(r, len(vals) - 1, wb.add_format({'border': 1, 'bg_color': 'FFC7CE'}))
        # 合计累计（与 openpyxl 版一致：open/debit/credit/close 带符号）
        tot["open"] += s["s_open"]; tot["debit"] += abs(s["s_debit"])
        tot["credit"] += abs(s["s_credit"]); tot["close"] += s["s_close"]
        tot["count"] += s["count"]
        tot["fopen"] += f_open; tot["fclose"] += f_close
        r += 1
    # 合计行
    tot_row = [None] * len(headers)
    tot_row[0] = '合计'
    for c, key in [(10 + _off, "fopen"), (11 + _off, "open"), (12 + _off, "debit"), (13 + _off, "credit"),
                   (14 + _off, "fclose"), (15 + _off, "close")]:
        tot_row[c] = tot[key]
    X.total_row(wb, ws, r, len(headers), tot_row, money_cols={10 + _off, 11 + _off, 12 + _off,
                                                              13 + _off, 14 + _off, 15 + _off})
    return r
