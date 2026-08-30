# -*- coding: utf-8 -*-
"""xlsxwriter 渲染原型：验证 双行表头(merge) + 66万行流式 + freeze 兼容审计规范。
样本取 XBJ 其他应收款明细表前 N 行真实数据。"""
import paths as P
import openpyxl, xlsxwriter, time, os

SRC = os.path.join(P.DATA_DIRS['XBJ'], '数据', '2025', '其他应收款审计底稿_2025_生成.xlsx')
OUT = os.path.join(P.DATA_DIRS['AH'], 'tmp_xw', '其他应收款_原型_xlsxwriter.xlsx')
N = int(os.environ.get('PROTO_ROWS', '5000'))

HEADERS = ['核算主体', '序号', '往来单位编号', '往来单位名称', '款项性质', '是否关联方', '币种',
           '最上层客户单位', '客户性质', '坏账计提方法',
           '外币余额(期初)', '期初余额(人民币)', '本期借方发生额', '本期贷方发生额',
           '外币余额(期末)', '期末余额(人民币)', '重分类调整', '审计调整', '审定数',
           '1年以内', '1-2年', '2-3年', '3-4年', '4-5年', '5年以上',
           '交易笔数', '平衡校验(期初+借+贷-期末)', '期后收款']
GROUPS = [('基础信息', 10), ('期初', 2), ('本期发生额', 2), ('期末', 2), ('调整', 3), ('账龄', 6), ('其他', 3)]
MONEY = {11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 27, 28}

# 读源数据（R1-R5 表头跳过，R6 起数据）
ws_src = openpyxl.load_workbook(SRC, data_only=True, read_only=True)['其他应收款明细表_2025']
rows = []
for i, r in enumerate(ws_src.iter_rows(values_only=True)):
    if i < 5:
        continue
    rows.append(r)
    if len(rows) >= N:
        break
ws_src.parent.close()

os.makedirs(os.path.dirname(OUT), exist_ok=True)
t0 = time.perf_counter()
wb = xlsxwriter.Workbook(OUT, {'constant_memory': True})
ws = wb.add_worksheet('其他应收款明细表_2025')
title_fmt = wb.add_format({'bold': True, 'font_size': 11})
grp_fmt = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#4472C4',
                         'align': 'center', 'border': 1})
hdr_fmt = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#4472C4',
                         'align': 'center', 'border': 1})
num_fmt = wb.add_format({'num_format': '#,##0.00', 'border': 1})
txt_fmt = wb.add_format({'border': 1})
# ⚡ 2026-08-26 列分组底色（对齐 _cwrite_period_sheet：期初蓝/发生额绿/期末蓝/调整黄/账龄灰）
def _gfmt(color, num=False):
    f = {'border': 1, 'bg_color': color}
    if num:
        f['num_format'] = '#,##0.00'
    return wb.add_format(f)
_off = 0
_fill_blue = _gfmt('DDEBF7', True)
_fill_green = _gfmt('E2EFDA', True)
_fill_yellow = _gfmt('FFF2CC', True)
_fill_gray = _gfmt('EDEDED', True)
_col_fmt = {}
for _c in range(1, 29):
    if _c in (11 + _off, 12 + _off, 15 + _off, 16 + _off):
        _col_fmt[_c] = _fill_blue
    elif _c in (13 + _off, 14 + _off):
        _col_fmt[_c] = _fill_green
    elif _c in (17 + _off, 18 + _off, 19 + _off):
        _col_fmt[_c] = _fill_yellow
    elif 20 + _off <= _c <= 25 + _off:
        _col_fmt[_c] = _fill_gray
    elif _c in MONEY:
        _col_fmt[_c] = num_fmt
    else:
        _col_fmt[_c] = txt_fmt
# R1 标题（merge）
ws.merge_range(0, 0, 0, 27, '其他应收款明细表 — 2025', title_fmt)
# R3 分组行（merge，R2 留空）
c0 = 0
for gname, gn in GROUPS:
    ws.merge_range(2, c0, 2, c0 + gn - 1, gname, grp_fmt)
    c0 += gn
# R4 列名
for j, h in enumerate(HEADERS):
    ws.write(3, j, h, hdr_fmt)
# R5+ 数据（流式，逐 cell 带分组底 format）
for i, row in enumerate(rows):
    for j, v in enumerate(row):
        col = j + 1
        fmt = _col_fmt[col]
        if v is None:
            ws.write_blank(4 + i, j, None, fmt)
        elif isinstance(v, (int, float)):
            ws.write_number(4 + i, j, v, fmt)
        else:
            ws.write_string(4 + i, j, str(v), fmt)
ws.freeze_panes(5, 0)
wb.close()
dt = time.perf_counter() - t0
fsize = os.path.getsize(OUT) / 1048576
print('OK rows=%d write=%.1fs file=%.1fMB' % (len(rows), dt, fsize))
