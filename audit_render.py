# -*- coding: utf-8 -*-
"""audit_render.py —— 渲染层（阶段二 2.1 架构枢纽）

目标：生成器只产【结构化数据】（数值/文本 rows，不含样式），写 Excel 全部收敛到本渲染器。
模板 = 布局 + 样式 + 规则（双行表头/分组底色/零值留空/冻结/合并列强化 = 既有排版规则沉淀为默认样式）。

核心协议 TableSpec：
    spec = {
        'name':      sheet 名（必填）
        'title':     大标题（可选）
        'notes':     [说明行...]（可选，标题后灰色斜体，合并整行）
        'headers':   ['列1','列2',...]          单行表头
                 或  (['大类',跨列数,...], ['列1',...])  双行分组表头
        'rows':      [RowSpec, ...]
        'widths':    [列宽...]
        'freeze':    'A6' 或 None
        'money_cols': {4,5,6} 默认金额列（RowSpec 未标注时按此）
    }

RowSpec（结构化行，渲染器负责样式）：
    {
        'v':    [值...]                        # None=空
        'b':    True/False                      # 加粗
        'fill': 'total'|'group'|'sub'|None      # 合计行层级底色
        'merge': (c0,c1) 1-based                # 本行合并列（A 列起始）
        'num':  set 或 'all' 或 None            # 本行金额列（覆盖默认 money_cols）
        'align': 'l'|'c'|'r' 覆盖               # 本行对齐覆盖
        'red':  True/False                      # 金额红字（差异提示）
        'skip_border': True                     # 不加边框（纯说明行）
    }

渲染器只认数据+规则，不感知科目语义；生成器计算层与呈现层完全解耦。
2026-08-11 创建（阶段二 2.1，渐进式：equity 试点 → 逐科目推广）。
"""
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---- 样式常量：单一事实来源 = audit_shell（阶段二 2.1：渲染器只消费共享常量，
#      勿在此重复定义——曾重复定义导致标题 14号/小计中蓝 与全集团 10号/浅蓝 不一致）----
from audit_shell import (SHELL_NUM, SHELL_HFILL, SHELL_HFONT, SHELL_BOLD,
                         SHELL_BODY, SHELL_TITLE_FONT, SHELL_TOT_FILL, FILL_TOTAL_MID,
                         FILL_GROUP_DARK, SHELL_CEN, SHELL_LEFT, SHELL_RGT,
                         SHELL_BORDER)
SHELL_FONT = SHELL_BODY            # 正文默认字体（audit_shell.SHELL_BODY）
NOTE_FONT = Font(name='Times New Roman', size=9, italic=True, color='808080')  # 说明行字体（audit_shell 无此语义）
SHELL_SUB_FONT = NOTE_FONT         # 兼容别名（说明行专用，勿用于小标题）

_FILLS = {
    'sub': SHELL_TOT_FILL,
    'total': FILL_TOTAL_MID,
    'group': FILL_GROUP_DARK,
}


def _cell_style(style):
    if style == 'group':
        return Font(name='Times New Roman', size=10, bold=True, color='000000'), SHELL_CEN
    return SHELL_BOLD, SHELL_CEN


def render_sheet(wb, spec):
    """按 TableSpec 渲染一个 sheet 到工作簿。返回 ws。"""
    # ⚡ 2026-08-11：同名 sheet 已存在（生成器 wb.active 已改名）→ 复用，勿重复创建（rd_expense 曾出"明细表1"）
    if spec['name'] in wb.sheetnames:
        ws = wb[spec['name']]
    else:
        ws = wb.create_sheet(spec['name'])
    r = 1
    # ── 大标题 ──
    if spec.get('title'):
        ncols = spec.get('ncols') or _spec_ncols(spec)
        c = ws.cell(r, 1, spec['title'])
        c.font = SHELL_TITLE_FONT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(ncols, 1))
        r += 1
    # ⚡ 2026-08-11：title_gap>0 时标题后空行（equity 原版 R1 标题/R2 空/R3 表头，逐格对齐）
    r += spec.get('title_gap', 0)
    # ── 说明行 ──
    ncols = spec.get('ncols') or _spec_ncols(spec)
    for note in spec.get('notes') or []:
        c = ws.cell(r, 1, note)
        c.font = SHELL_SUB_FONT
        c.alignment = SHELL_LEFT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        r += 1
    # ── 表头（单行或双行分组）──
    hr = r
    hdr = spec['headers']
    if isinstance(hdr, tuple) and len(hdr) == 2 and isinstance(hdr[0], list) and isinstance(hdr[0][0], tuple):
        # 双行分组表头：(groups, cols)
        groups, cols = hdr
        # ⚡ 2026-08-11：group_style='light' 时第一行分组浅蓝黑字（loan 原版）；
        #   默认深蓝白字（current_account 等）。grp_gap=1 时分组行与列名行之间空 1 行
        #   （loan 原版：R1 标题/R2 分组/R3 空/R4 列名）。
        _g_light = spec.get('group_style') == 'light'
        _g_gap = spec.get('grp_gap', 0)
        c0 = 1
        for gname, gn in groups:
            for _c in range(c0, c0 + gn):
                ws.cell(r, _c).border = SHELL_BORDER
                ws.cell(r, _c).fill = PatternFill('solid', fgColor='DDEBF7') if _g_light else SHELL_HFILL
            if gn <= 1:
                cell = ws.cell(r, c0, gname)
                cell.font = SHELL_BOLD if _g_light else SHELL_HFONT
                cell.alignment = SHELL_CEN
            else:
                ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + gn - 1)
                cell = ws.cell(r, c0, gname)
                cell.font = SHELL_BOLD if _g_light else SHELL_HFONT
                cell.alignment = SHELL_CEN
            c0 += gn
        r += 1
        r += _g_gap
        for j, h in enumerate(cols, 1):
            c = ws.cell(r, j, h)
            c.fill = SHELL_HFILL
            c.font = SHELL_HFONT
            c.border = SHELL_BORDER
            c.alignment = SHELL_CEN
        r += 1
    else:
        for j, h in enumerate(hdr, 1):
            c = ws.cell(r, j, h)
            c.fill = SHELL_HFILL
            c.font = SHELL_HFONT
            c.border = SHELL_BORDER
            c.alignment = SHELL_CEN
        r += 1
    # ── 数据行 ──
    money_cols = spec.get('money_cols') or set()
    # ⚡ 2026-08-11：zero_blank=False 时 0 值也写入（equity 原版行为，逐格比对零信息损失）；
    #   默认 True（零值留空 R1，current_account 等已应用的排版规则）。
    zero_blank = spec.get('zero_blank', True)
    for rs in spec.get('rows') or []:
        vals = rs.get('v') or []
        b = rs.get('b', False)
        fill = rs.get('fill')
        merge = rs.get('merge')
        align_o = rs.get('align')
        red = rs.get('red', False)
        num = rs.get('num')
        skip_border = rs.get('skip_border', False)
        cell_fill = rs.get('cell_fill') or {}   # {col: PatternFill} 单元格级底色（如 loan 备注列黄底）
        font_ov = rs.get('font') or {}          # {col: Font} 单元格级字体覆盖（expense L3/L4 斜体灰字）
        for j, v in enumerate(vals, 1):
            cell = ws.cell(r, j)
            # 值（零值留空 R1：|v|≤0.005 不写值，仅写非零；zero_blank=False 保持原写 0 行为）
            if v is None:
                pass
            elif isinstance(v, str) and v.startswith('='):
                # 公式字符串（如 gp_other =G+H）→ 直接写入，不 round/不金额格式
                cell.value = v
                cell.alignment = SHELL_RGT
                if not skip_border:
                    cell.border = SHELL_BORDER
                continue
            elif isinstance(v, (int, float)):
                if zero_blank:
                    if abs(v) > 0.005:
                        cell.value = round(float(v), 2)
                else:
                    cell.value = round(float(v), 2)
            else:
                cell.value = v
            # 单元格级底色（优先于行级 fill）
            if j in cell_fill:
                cell.fill = cell_fill[j]
            # 对齐
            if align_o == 'l':
                cell.alignment = SHELL_LEFT
            elif align_o == 'c':
                cell.alignment = SHELL_CEN
            elif align_o == 'r':
                cell.alignment = SHELL_RGT
            elif num == 'all' or (num is None and j in money_cols):
                cell.alignment = SHELL_RGT
            else:
                cell.alignment = SHELL_LEFT
            # 金额格式
            if num == 'all' or (num is None and j in money_cols) or (isinstance(num, set) and j in num):
                cell.number_format = SHELL_NUM
            # 边框（skip_border 说明行不加）
            if not skip_border:
                cell.border = SHELL_BORDER
            # 字体/底色
            if j in font_ov:
                cell.font = font_ov[j]
            elif fill in _FILLS:
                cell.fill = _FILLS[fill]
                _f, _a = _cell_style(fill)
                cell.font = _f
            elif b:
                cell.font = SHELL_BOLD
            else:
                cell.font = SHELL_FONT
            if red and (v is not None) and (not isinstance(v, (int, float)) or abs(v) > 0.005):
                cell.font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
        if merge:
            ws.merge_cells(start_row=r, start_column=merge[0], end_row=r, end_column=merge[1])
        # 合计行层级上粗边（R5）
        if fill == 'total':
            top = Side(style='medium', color='000000')
            for j in range(1, max(len(vals), 1) + 1):
                b = ws.cell(r, j).border
                ws.cell(r, j).border = Border(left=b.left, right=b.right, top=top, bottom=b.bottom)
        r += 1
    # ── 列宽 ──
    for i, w in enumerate(spec.get('widths') or [], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # ── 冻结 ──
    if spec.get('freeze'):
        ws.freeze_panes = spec['freeze']
    return ws


def _spec_ncols(spec):
    hdr = spec.get('headers')
    if isinstance(hdr, tuple) and len(hdr) == 2 and isinstance(hdr[0], list):
        return len(hdr[1])
    return len(hdr or [])


# ============================================================
# 结构化数据模型：RowSpec 构造助手（生成器计算层专用）
# ============================================================
def row(vals, b=False, fill=None, merge=None, num=None, align=None, red=False, skip_border=False, cell_fill=None, font=None):
    """构造一行结构化数据（不含样式语义之外的呈现细节；渲染器决定样式）。"""
    return dict(v=vals, b=b, fill=fill, merge=merge, num=num,
                align=align, red=red, skip_border=skip_border, cell_fill=cell_fill, font=font)


def render_into(ws, headers, rows, money_cols, groups=None, freeze='A2', title=None):
    """把结构化 rows 渲染进【已存在的 sheet】（不新建），用于模板克隆场景。
    ⚡ 2026-08-11：render_sheet 面向新建 sheet；模板驱动的 write_* 直接写已有 ws，
    复用渲染器的行渲染逻辑（值/金额格式/边框/对齐/行样式/字体覆盖/双行表头）。
    groups: 双行表头分组（[(名, 跨列数),...]），提供时表头=分组行+列名行。
    title: 大标题，写在表头上方（R1），表头顺延。"""
    _t = 1 if title else 0
    if title:
        ws.cell(1, 1, title).font = SHELL_TITLE_FONT
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(headers), 1))
    if groups:
        # 双行表头：R(1+t) 分组（深蓝白字合并），R(2+t) 列名，数据从 R(3+t)
        c0 = 1
        for gname, gn in groups:
            for _c in range(c0, c0 + gn):
                ws.cell(1 + _t, _c).border = SHELL_BORDER
                ws.cell(1 + _t, _c).fill = SHELL_HFILL
            if gn <= 1:
                cell = ws.cell(1 + _t, c0, gname)
                cell.font = SHELL_HFONT
                cell.alignment = SHELL_CEN
            else:
                ws.merge_cells(start_row=1 + _t, start_column=c0, end_row=1 + _t, end_column=c0 + gn - 1)
                cell = ws.cell(1 + _t, c0, gname)
                cell.font = SHELL_HFONT
                cell.alignment = SHELL_CEN
            c0 += gn
        for j, h in enumerate(headers, 1):
            cell = ws.cell(2 + _t, j, h)
            cell.fill = SHELL_HFILL
            cell.font = SHELL_HFONT
            cell.border = SHELL_BORDER
            cell.alignment = SHELL_CEN
        _r0 = 3 + _t
    else:
        # 单行表头（R(1+t)，与 render_sheet 单行表头同款：深蓝白字居中）
        for j, h in enumerate(headers, 1):
            cell = ws.cell(1 + _t, j, h)
            cell.fill = SHELL_HFILL
            cell.font = SHELL_HFONT
            cell.border = SHELL_BORDER
            cell.alignment = SHELL_CEN
        _r0 = 2 + _t
    r = _r0
    for rs in rows:
        vals = rs.get('v') or []
        b = rs.get('b', False)
        fill = rs.get('fill')
        align_o = rs.get('align')
        num = rs.get('num')
        font_ov = rs.get('font') or {}
        cell_fill = rs.get('cell_fill') or {}
        for j, v in enumerate(vals, 1):
            cell = ws.cell(r, j)
            if v is None:
                pass
            elif isinstance(v, str) and v.startswith('='):
                cell.value = v
            elif isinstance(v, (int, float)):
                cell.value = round(float(v), 2)
            else:
                cell.value = v
            cell.border = SHELL_BORDER
            if j in font_ov:
                cell.font = font_ov[j]
            elif fill == 'sub':
                cell.fill = SHELL_TOT_FILL
                cell.font = SHELL_BOLD
            elif fill == 'total':
                cell.fill = FILL_TOTAL_MID
                cell.font = SHELL_BOLD
            elif fill == 'group':
                cell.fill = FILL_GROUP_DARK
                cell.font = SHELL_BOLD
            elif b:
                cell.font = SHELL_BOLD
            else:
                cell.font = SHELL_FONT
            if j in cell_fill:
                cell.fill = cell_fill[j]
            if num == 'all' or (num and j in num) or (num is None and j in money_cols):
                cell.number_format = SHELL_NUM
                cell.alignment = SHELL_RGT
            else:
                cell.alignment = SHELL_LEFT if align_o == 'l' else (SHELL_CEN if align_o == 'c' else SHELL_LEFT)
        r += 1
    ws.freeze_panes = freeze
    return ws


# ============================================================================
# 阶段二 2.2 模板体系（单体/合并/附注）
# 排版规则沉淀为模板默认样式：双行表头/分组底色/零值留空/冻结/合并列强化。
# 新增科目只需：调模板工厂 + 给数据 rows（2.1 协议），无需再写渲染。
# ============================================================================

def tpl_mono_detail(title, headers, rows, money_cols, widths=None, zero_blank=False,
                    freeze='A3', title_gap=1, notes=None):
    """单体明细模板：R1 标题 / R2(空) / R3 表头 / R4 起数据。
    equity/loan/rd 主明细表同构。rows 用 row() 构造。"""
    return dict(name=None, title=title, title_gap=title_gap, headers=headers,
                rows=rows, money_cols=set(money_cols), widths=widths,
                zero_blank=zero_blank, freeze=freeze, notes=notes or [])


def tpl_two_row_detail(title, groups, headers, rows, money_cols, widths=None,
                       zero_blank=False, freeze='A4', group_style='light',
                       grp_gap=0, notes=None):
    """双行表头明细模板：R1 标题 / R2 分组 / R3(可空) 列名 / 数据起。
    tax(深蓝)/loan(浅蓝 light) 主明细表同构。"""
    return dict(name=None, title=title, headers=(groups, headers),
                rows=rows, money_cols=set(money_cols), widths=widths,
                zero_blank=zero_blank, freeze=freeze, group_style=group_style,
                grp_gap=grp_gap, notes=notes or [])


def tpl_combined_audit(title, headers, rows, money_cols, widths=None,
                       zero_blank=False, freeze='A3', notes=None):
    """合并审定表模板（2.3 合并工作底稿主表）：项目×主体横展 + 合并4列。
    headers 含 ['项目'] + 主体列 + ['合并抵消','合并调整','合并审定数','勾稽']。"""
    return dict(name=None, title=title, title_gap=1, headers=headers,
                rows=rows, money_cols=set(money_cols), widths=widths,
                zero_blank=zero_blank, freeze=freeze, notes=notes or [])


def tpl_footnote(title, headers, rows, money_cols, widths=None, freeze='A2',
                 notes=None):
    """附注汇总模板：R1 标题 / R2 表头 / R3 数据（gp_other 审定表同构，紧凑布局）。"""
    return dict(name=None, title=title, title_gap=0, headers=headers,
                rows=rows, money_cols=set(money_cols), widths=widths,
                zero_blank=False, freeze=freeze, notes=notes or [])


# 模板注册表（按底稿类型索引；name 由生成器在调用时补全）
TEMPLATES = {
    'mono_detail': tpl_mono_detail,
    'two_row_detail': tpl_two_row_detail,
    'combined_audit': tpl_combined_audit,
    'footnote': tpl_footnote,
}


def render_tpl(wb, tpl_key, name, **kw):
    """按模板渲染：TEMPLATES[tpl_key](**kw) → spec → render_sheet(wb, spec)。
    新增科目入口：提供 标题/表头/rows/金额列 即可。"""
    spec = TEMPLATES[tpl_key](**kw)
    spec['name'] = name
    return render_sheet(wb, spec)
