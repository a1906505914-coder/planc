# -*- coding: utf-8 -*-
"""验证 xlsxwriter 原型：merge/freeze/数值/行数与源一致"""
import paths as P
import openpyxl

PROTO = os.path.join(P.DATA_DIRS['AH'], 'tmp_xw', '其他应收款_原型_xlsxwriter.xlsx')
SRC = os.path.join(P.DATA_DIRS['XBJ'], '数据', '2025', '其他应收款审计底稿_2025_生成.xlsx')

wb = openpyxl.load_workbook(PROTO, data_only=True)
ws = wb.active
print('sheet:', ws.title, ws.max_row, 'x', ws.max_column)
print('冻结:', ws.freeze_panes)
print('合并单元格:', sorted(str(r) for r in ws.merged_cells.ranges)[:10], '共', len(ws.merged_cells.ranges))
# 表头
print('R1:', ws.cell(1, 1).value)
print('R3 分组:', [ws.cell(3, c).value for c in range(1, 29) if ws.cell(3, c).value])
print('R4 列名:', [ws.cell(4, c).value for c in range(1, 9)])
# 数据抽样（源 vs 原型）
wb2 = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
ws2 = wb2['其他应收款明细表_2025']
ok = 0
for i in [0, 1, 100, 4999]:
    vp = [ws.cell(5 + i, c).value for c in (1, 2, 4, 12, 13, 28)]
    vs = [ws2.cell(6 + i, c).value for c in (1, 2, 4, 12, 13, 28)]
    same = all(str(vp[k]) == str(vs[k]) for k in range(len(vp)))
    ok += same
    print('R%d 原型:%s' % (5 + i, vp))
    print('   源 :%s %s' % (vs, 'OK' if same else 'DIFF'))
wb.close()
wb2.close()
print('抽样一致:', ok, '/4')
