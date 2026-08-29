# -*- coding: utf-8 -*-
"""xlsxwriter 渲染层（#876 重构基础，参照 xw_proto_build.py 原型能力）。

把 current_account_detail 导出层从 openpyxl 迁移到 xlsxwriter：
   - constant_memory 流式写入（66万行不 OOM）
   - 双行表头（分组 merge）+ 列名 + 冻结 + 列宽
   - 金额列格式 / 列底色 / 公式（write_formula）
用法见 current_account_detail 的导出层迁移；本模块不依赖 openpyxl。
"""
import xlsxwriter


def new_workbook(path):
    return xlsxwriter.Workbook(path, {'constant_memory': True})


def close_workbook(wb):
    wb.close()


def add_sheet(wb, name):
    return wb.add_worksheet(name)


def _fmt(wb, **kw):
    return wb.add_format(kw)


def title(wb, ws, text, ncols, row=0):
    """标题行（合并 1..ncols）。返回下一行号。"""
    ws.merge_range(row, 0, row, ncols - 1, text,
                   _fmt(wb, **{'bold': True, 'font_size': 11}))
    return row + 1


def subtitle(wb, ws, text, ncols, row):
    ws.merge_range(row, 0, row, ncols - 1, text,
                   _fmt(wb, **{'italic': True, 'font_color': '#808080'}))
    return row + 1


def dual_header(wb, ws, group_row, groups, headers, widths=None, freeze_rows=0):
    """双行表头：group_row 行=大类分组（merge），group_row+1 行=列名。
    groups: [(分组名, 列数), ...]；widths: 每列宽列表（可选）。
    返回数据起始行（group_row+2）。"""
    grp_fmt = _fmt(wb, **{'bold': True, 'font_color': '#FFFFFF',
                        'bg_color': '#4472C4', 'align': 'center', 'border': 1})
    hdr_fmt = _fmt(wb, **{'bold': True, 'font_color': '#FFFFFF',
                        'bg_color': '#4472C4', 'align': 'center', 'border': 1})
    col = 0
    for gname, cnt in groups:
        if cnt > 1:
            ws.merge_range(group_row, col, group_row, col + cnt - 1, gname, grp_fmt)
        else:
            ws.write(group_row, col, gname, grp_fmt)
        col += cnt
    for j, h in enumerate(headers):
        ws.write(group_row + 1, j, h, hdr_fmt)
    if widths:
        for j, w in enumerate(widths):
            ws.set_column(j, j, w)
    if freeze_rows:
        ws.freeze_panes(freeze_rows, 0)
    return group_row + 2


def row(wb, ws, r, values, money_cols=(), center_cols=(), num_fmt='#,##0.00',
        fills=None, formula_cols=None):
    """写一行。values 含公式字符串（'=' 开头）时自动 write_formula。
    money_cols: 金额列（0-based）；center_cols: 居中列；fills: {col: 背景色}。"""
    for j, v in enumerate(values):
        if v is None:
            continue
        if formula_cols and j in formula_cols:
            ws.write_formula(r, j, v, _fmt(wb, **{'num_format': num_fmt}))
        elif isinstance(v, (int, float)) and not isinstance(v, bool):
            f = {'num_format': num_fmt}
            if fills and j in fills:
                f['bg_color'] = fills[j]
            ws.write_number(r, j, v, _fmt(wb, **f))
        else:
            f = {}
            if center_cols and j in center_cols:
                f['align'] = 'center'
            if fills and j in fills:
                f['bg_color'] = fills[j]
            ws.write_string(r, j, str(v), _fmt(wb, **f))


def money_fmt(wb):
    return _fmt(wb, **{'num_format': '#,##0.00'})


# ---- 列分组底色（对齐 _cwrite_period_sheet：期初蓝/发生额绿/期末蓝/调整黄/账龄灰） ----
def period_col_fills(wb, off, n_buckets=0):
    """返回 {0-based列号: format}。off=合同号列偏移（0/1）；n_buckets=账龄桶数。"""
    fills = {}
    blue = _fmt(wb, **{'num_format': '#,##0.00', 'bg_color': 'DDEBF7'})
    green = _fmt(wb, **{'num_format': '#,##0.00', 'bg_color': 'E2EFDA'})
    yellow = _fmt(wb, **{'num_format': '#,##0.00', 'bg_color': 'FFF2CC'})
    gray = _fmt(wb, **{'num_format': '#,##0.00', 'bg_color': 'EDEDED'})
    for c in (10 + off, 11 + off, 14 + off, 15 + off):   # 外币期初/期初/外币期末/期末
        fills[c] = blue
    for c in (12 + off, 13 + off):                        # 本期借/贷
        fills[c] = green
    for c in (16 + off, 17 + off, 18 + off):              # 重分类/审计调整/审定
        fills[c] = yellow
    if n_buckets:
        for c in range(19 + off, 19 + off + n_buckets):   # 账龄
            fills[c] = gray
    return fills


def total_row(wb, ws, r, ncols, values, money_cols=()):
    """合计行：粗体 + 灰底，金额列带格式。"""
    bold = _fmt(wb, **{'bold': True, 'bg_color': 'D9D9D9'})
    bold_num = _fmt(wb, **{'bold': True, 'bg_color': 'D9D9D9', 'num_format': '#,##0.00'})
    for j, v in enumerate(values[:ncols]):
        if v is None:
            continue
        f = bold_num if j in money_cols else bold
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            ws.write_number(r, j, v, f)
        else:
            ws.write_string(r, j, str(v), f)
    return r + 1
