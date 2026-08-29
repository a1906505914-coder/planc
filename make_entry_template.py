# -*- coding: utf-8 -*-
"""生成《审计调整分录汇总表》空模板（审计师填写，过入器 post_entry.py 读取）。
字段：年度/核算主体/科目/调整类别/借贷方向/金额/明细对象(可选)/调整说明。
"""
import paths as P
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

OUT = os.path.join(P.FY, '审计调整分录汇总表_模板.xlsx')
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "调整分录"

HFILL = PatternFill('solid', fgColor='DDEBF7')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(name='Times New Roman', size=11, bold=True)
F10 = Font(name='Times New Roman', size=10)
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
NUM = '#,##0.00'

ws.merge_cells('A1:I1')
ws.cell(1, 1, '审计调整分录汇总表 — 请按行填写，过入器将自动过入各底稿审定表/明细表').font = Font(name='Times New Roman', size=13, bold=True)
ws.cell(1, 1).alignment = CTR

note = ('填写说明：\n'
        '1. 年度：2025 / 2026（与底稿文件名年度一致）。\n'
        '2. 核算主体：01FY本级公司 / 02FY技术公司 … 12FY控股公司（与审定表"公司"列一致）。\n'
        '3. 科目：应收账款 / 应付账款 / 固定资产 / 管理费用 等（与底稿科目名一致）。\n'
        '4. 调整类别：审计调整 / 重分类调整1 / 重分类调整2（分别写入审定表 K/I/J 列）。\n'
        '5. 借贷方向：借 / 贷；金额填正数。同一调整的借方合计必须 = 贷方合计（过入器校验，不平拒绝）。\n'
        '6. 明细对象（可选）：往来单位名称。填写后同时过入明细表 Q(重分类)/R(审计调整) 对应行。\n'
        '7. 摘要/调整说明：可填凭证号、调整原因，便于复核。\n'
        '8. 程序按 主体×科目×年度 定位；未匹配到审定表行的记录会在过入报告中提示。')
ws.merge_cells('A2:I9')
ws.cell(2, 1, note).font = Font(name='Times New Roman', size=10, color='595959')
ws.cell(2, 1).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

hdr = ['序号', '年度', '核算主体', '科目', '调整类别', '借贷方向', '金额', '明细对象(可选)', '摘要/调整说明']
for j, h in enumerate(hdr, 1):
    c = ws.cell(11, j, h)
    c.font = BOLD; c.fill = HFILL; c.border = BORDER; c.alignment = CTR

example = [1, 2026, '01FY本级公司', '应收账款', '审计调整', '借', 100000, '某客户名称', '示例：坏账准备补提（删除此行后填写）']
example2 = [2, 2026, '01FY本级公司', '信用减值损失', '审计调整', '贷', 100000, '', '示例：对应贷方（借贷必须平衡）']
for i, ex in enumerate((example, example2)):
    for j, v in enumerate(ex, 1):
        c = ws.cell(11 + i, j, v)
        c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
        c.border = BORDER
        if j == 7:
            c.number_format = NUM
            c.alignment = Alignment(horizontal='right')

for j, w in enumerate([8, 10, 20, 16, 16, 10, 16, 24, 40], 1):
    ws.column_dimensions[chr(64 + j)].width = w
ws.freeze_panes = 'A12'

if __name__ == '__main__':
    # 2026-08-05 修复：顶层脚本原无 main 保护，被 post_entry import 时会立即执行
    # （在 FY 目录生成模板文件——import 副作用）；包进 __main__ 后 import 零副作用。
    wb.save(OUT)
    print("模板已生成：", OUT)
