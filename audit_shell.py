# -*- coding: utf-8 -*-
"""
audit_shell.py — 六个审计小程序共享的「外壳」加载与构建工具。

统一目标：
  · 所有外壳集中存放在桌面 audit_templates/ 文件夹（远离被拖的数据文件夹，杜绝误删）。
  · 所有外壳采用同一套视觉样式（蓝底白字加粗表头 / 边框 / 千分位 / 统一字体）。
  · 六个程序共用同一套加载逻辑（resolve_template / clone_sheet / make_default_shell），
    不再各写各的；缺壳时自动降级兜底，保证始终有输出、不闪退、不静默 0 文件。

外壳模式工作流（推荐范式，费用/利润表已验证）：
  1. 程序启动时用 resolve_template(name, ENV_VAR) 定位外壳 xlsx；
  2. 命中 → load_workbook(shell) 取得空格式框架（sheet/标题/表头/列宽/冻结）；
  3. 缺失 → make_default_shell(spec) 生成最简空壳（格式降级，仍有输出）；
  4. 运行时只往框架里填数据（逐格 number_format / border），结构样式全部来自外壳。
今后改排版：只改 audit_templates/ 里的 xlsx，无需动代码。
"""
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ===================== 统一视觉样式（所有外壳一致） =====================
# ⚡ 2026-08-12 可读性修复（用户要求：字体统一黑色）：深蓝底(1F4E78)改浅蓝底(DDEBF7)、
#   白字(FFFFFF)改黑字(000000)——深蓝底+白字在部分显示器/打印下对比不足，黑字更稳。
SHELL_HFILL = PatternFill('solid', fgColor='DDEBF7')          # 表头：浅蓝底（黑字清晰）
SHELL_HFONT = Font(name='Times New Roman', bold=True, color='000000', size=10)   # 表头：黑字加粗
SHELL_TOT_FILL = PatternFill('solid', fgColor='DDEBF7')       # 合计/小计/全集团合计：浅蓝底（黑字清晰可读）

SHELL_TITLE_FONT = Font(name='Times New Roman', bold=True, size=10, color='000000')  # 大标题
SHELL_SUB_FONT = Font(name='Times New Roman', bold=True, size=10)         # 小标题
SHELL_BOLD = Font(name='Times New Roman', bold=True, size=10)             # 加粗（小计/合计）
SHELL_NUM = '#,##0.00'                                        # 金额千分位（2 位小数、无货币符号）
SHELL_NUM_INT = '0'                                           # 序号/年份/月份等标识整数列（不保留小数）

# 全底稿统一字体/字号（2026-07-31 用户要求：整张底稿 Times New Roman、10 号）
FONT_NAME = 'Times New Roman'
FONT_SIZE = 10
SHELL_BODY = Font(name='Times New Roman', size=10)                        # 正文默认字体
SHELL_THIN = Side(style='thin', color='BFBFBF')
SHELL_BORDER = Border(left=SHELL_THIN, right=SHELL_THIN, top=SHELL_THIN, bottom=SHELL_THIN)
SHELL_CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
SHELL_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
SHELL_RGT = Alignment(horizontal='right', vertical='center')

# 中央外壳文件夹：默认 <脚本目录>/audit_templates，可用环境变量 AUDIT_TEMPLATES_DIR 覆盖
AUDIT_TEMPLATES_DIR = os.environ.get(
    'AUDIT_TEMPLATES_DIR',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'audit_templates')
)


def _norm_name(name):
    return name if name.endswith('.xlsx') else name + '.xlsx'


def resolve_template(name, env_var=None):
    """定位外壳 xlsx。查找顺序：环境变量 → AUDIT_TEMPLATES_DIR → 脚本目录/audit_templates → 桌面。
    返回完整路径或 None。name 可含 / 不含 .xlsx 扩展名。"""
    candidates = []
    if env_var:
        v = os.environ.get(env_var)
        if v:
            candidates.append(v)
    candidates += [
        os.path.join(AUDIT_TEMPLATES_DIR, _norm_name(name)),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'audit_templates', _norm_name(name)),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), _norm_name(name)),
        os.path.join(os.path.expanduser('~'), 'Desktop', _norm_name(name)),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return None


def clone_sheet(src, dst_wb, title=None):
    """把 src worksheet 的结构与格式克隆到 dst_wb（新建同名/title sheet），保留值/样式/列宽/冻结。"""
    ws = dst_wb.create_sheet(title or src.title)
    for row in src.iter_rows():
        for cell in row:
            nc = ws.cell(cell.row, cell.column, cell.value)
            if cell.has_style:
                nc.font = cell.font.copy()
                nc.fill = cell.fill.copy()
                nc.border = cell.border.copy()
                nc.alignment = cell.alignment.copy()
                nc.number_format = cell.number_format
    for col, dim in src.column_dimensions.items():
        ws.column_dimensions[col].width = dim.width
    ws.freeze_panes = src.freeze_panes
    ws.sheet_view.showGridLines = src.sheet_view.showGridLines
    return ws


def make_default_shell(spec):
    """缺壳兜底：spec = [(sheet_title, [表头...], [列宽...]), ...]
    生成最简空壳（含蓝头表头 + 列宽），保证始终有可用输出、不静默 0 文件。"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, headers, widths in spec:
        ws = wb.create_sheet(title)
        for j, h in enumerate(headers, 1):
            c = ws.cell(1, j, h)
            c.fill = SHELL_HFILL
            c.font = SHELL_HFONT
            c.border = SHELL_BORDER
            c.alignment = SHELL_CEN
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    return wb


def apply_header_row(ws, row, headers, start_col=1):
    """在指定行写表头（统一蓝头白字）。headers 为字符串列表。"""
    for j, h in enumerate(headers, start_col):
        c = ws.cell(row, j, h)
        c.fill = SHELL_HFILL
        c.font = SHELL_HFONT
        c.border = SHELL_BORDER
        c.alignment = SHELL_CEN
    return row + 1


# ============================================================
# 排版优化共享函数（2026-08-10 方案 R1-R7，零信息损失）
# ============================================================

# 列分组底色（R4）
FILL_GROUP_BLUE = PatternFill('solid', fgColor='DDEBF7')    # 期初/期末
FILL_GROUP_GREEN = PatternFill('solid', fgColor='E2EFDA')   # 本期发生额
FILL_GROUP_GRAY = PatternFill('solid', fgColor='F2F2F2')    # 账龄
FILL_GROUP_YELLOW = PatternFill('solid', fgColor='FFF2CC')  # 调整/审定
FILL_GROUP_RED = PatternFill('solid', fgColor='FCE4EC')     # 抵消借
FILL_GROUP_PURPLE = PatternFill('solid', fgColor='E8EAF6')  # 合并报表数
FILL_TOTAL_MID = PatternFill('solid', fgColor='BDD7EE')     # 表级合计（R5）
FILL_GROUP_DARK = PatternFill('solid', fgColor='DDEBF7')    # 集团行浅蓝（黑字，原深蓝白字可读性差）


def two_row_header(ws, row, groups, cols, start_col=1):
    """双行分组表头（R3）：第一行=大类分组（合并单元格跨列，深蓝底白字），
    第二行=具体列名（蓝头白字）。groups=[(大类名, 跨列数), ...]，cols=列名列表（len==Σ列数）。"""
    c0 = start_col
    for name, n in groups:
        for _c in range(c0, c0 + n):
            ws.cell(row, _c).border = SHELL_BORDER
            ws.cell(row, _c).fill = SHELL_HFILL
        if n <= 1:
            cell = ws.cell(row, c0, name)
            cell.font = SHELL_HFONT
            cell.alignment = SHELL_CEN
        else:
            ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c0 + n - 1)
            cell = ws.cell(row, c0, name)
            cell.font = SHELL_HFONT
            cell.alignment = SHELL_CEN
        c0 += n
    r2 = row + 1
    for j, h in enumerate(cols, start_col):
        c = ws.cell(r2, j, h)
        c.fill = SHELL_HFILL
        c.font = SHELL_HFONT
        c.border = SHELL_BORDER
        c.alignment = SHELL_CEN
    return r2 + 1


def money_cell(ws, r, c, v, fmt=SHELL_NUM):
    """零值留空写金额（R1）：|v|≤0.005 不写值（留空），仅写非零。"""
    if v is None:
        return
    if isinstance(v, (int, float)) and abs(v) > 0.005:
        cell = ws.cell(r, c, round(float(v), 2))
        cell.number_format = fmt
    elif not isinstance(v, (int, float)):
        ws.cell(r, c, v)


def group_fill(ws, r0, r1, c0, c1, fill):
    """数据区列分组底色（R4）：行 r0..r1、列 c0..c1 填充 fill 底色（保留边框）。"""
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            cell = ws.cell(r, c)
            cell.fill = fill


def style_total(ws, r, ncols, level='sub', start_col=1):
    """合计行层级样式（R5）：level=sub 块级小计(浅蓝)/total 表级合计(中蓝)/group 集团行(深蓝白字)。
    整行填充 + 加粗 + 上粗边框。"""
    fill = FILL_TOTAL_MID if level == 'total' else (FILL_GROUP_DARK if level == 'group' else SHELL_TOT_FILL)
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(r, c)
        cell.fill = fill
        cell.font = SHELL_BOLD
        cell.border = SHELL_BORDER
    top = Side(style='medium', color='000000')
    for c in range(start_col, start_col + ncols):
        b = ws.cell(r, c).border
        ws.cell(r, c).border = Border(left=b.left, right=b.right, top=top, bottom=b.bottom)


def hide_col_if_zero(ws, col, max_row, note_cell=None, note_txt=''):
    """外币列全 0 时隐藏该列（R3-A，不删除信息，可取消隐藏查看）。"""
    try:
        allzero = True
        for r in range(1, max_row + 1):
            v = ws.cell(r, col).value
            if isinstance(v, (int, float)) and abs(v) > 0.005:
                allzero = False
                break
        if allzero:
            ws.column_dimensions[ws.cell(1, col).column_letter].hidden = True
            if note_cell is not None and note_txt:
                ws.cell(*note_cell).value = note_txt
    except Exception:
        pass


def shell_path_for(name, env_var=None, auto_create=None):
    """定位外壳；若缺失且给定 auto_create（一个返回 Workbook 的 0 参函数），则生成并落盘后返回路径。"""
    p = resolve_template(name, env_var)
    if p:
        return p
    if auto_create is not None:
        wb = auto_create()
        out = os.path.join(AUDIT_TEMPLATES_DIR, _norm_name(name))
        os.makedirs(AUDIT_TEMPLATES_DIR, exist_ok=True)
        wb.save(out)
        return out
    return None


# ===================== 统一收尾：字体 + 千分位 =====================
# 标识符类表头：这些列的数字（年份/月份/序号/科目代码等）不套千分位，避免被误格式化。
_SHELL_IDENT_HEADERS = {
    '序号', 'no', '编号', '行次', '行', '页', '页次',
    '测试序号', '测试需要', '样本量', '抽样量',
    '月', '月份', '会计期间', '期间', '期',
    '科目代码', '科目编码', '代码', '科目',
    '年份', '年度', '年', '项目编号', '凭证号', '凭证字号', '凭证',
}
# 比率类表头关键字：占比/比例/增长率等小数不作金额千分位（避免 0.05 → 0.05 失真/误读）。
_SHELL_RATIO_KEYWORDS = ('率', '比例', '占比', '比重', '百分', '指数')


def _unify_sheet(ws):
    """单个 sheet 统一样式：①字体强制 Times New Roman 10 号（保留粗/斜/色）；
    ②数字单元格会计格式 '#,##0.00'（千分位、2 位小数、无货币符号），
    序号/年份/月份/凭证号等标识整数列用 '0'（不保留小数）；③数据列列宽自适应。"""
    # 收集各列表头名（前 6 行首个非空文本，用于标识列/比率列判定）
    col_header = {}
    top = min(ws.max_row, 6)
    for c in range(1, min(ws.max_column, 80) + 1):
        for r in range(1, top + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                col_header[c] = v.strip()
                break
    for row in ws.iter_rows():
        for cell in row:
            f = cell.font
            if f.name != FONT_NAME or f.size != FONT_SIZE:
                cell.font = Font(name=FONT_NAME, size=FONT_SIZE,
                                 bold=f.bold, italic=f.italic,
                                 underline=f.underline, strike=f.strike,
                                 color=f.color)
            v = cell.value
            if isinstance(v, bool) or not isinstance(v, (int, float)):
                continue
            hdr = col_header.get(cell.column, '')
            if hdr:
                hl = hdr.lower()
                if hl in _SHELL_IDENT_HEADERS:
                    # 标识整数列：序号/年份/月份/凭证号等不保留小数
                    if float(v).is_integer() and cell.number_format not in ('0', '0.00', '#,##0'):
                        cell.number_format = SHELL_NUM_INT
                    continue
                if any(k in hdr for k in _SHELL_RATIO_KEYWORDS):
                    continue
            nf = cell.number_format
            if nf not in (None, '', 'General'):
                continue  # 已设格式（百分比/日期/会计）→ 保留
            if isinstance(v, int) and (1900 <= v <= 2100 or 1 <= v <= 12):
                continue  # 年 / 月
            cell.number_format = SHELL_NUM
    auto_width(ws)


def auto_width(ws, max_width=45, min_width=8):
    """按内容自适应列宽：中文按 2 字符宽、数字按千分位显示宽（含格式占位）、
    跳过合并单元格（标题合并长文本不撑宽）。"""
    from openpyxl.utils.cell import range_boundaries
    merged = set()
    try:
        for rng in ws.merged_cells.ranges:
            mc, mr, xc, xr = range_boundaries(str(rng))
            for rr in range(mr, xr + 1):
                for cc in range(mc, xc + 1):
                    merged.add((rr, cc))
    except Exception:
        pass
    widths = {}
    for row in ws.iter_rows():
        for cell in row:
            if (cell.row, cell.column) in merged:
                continue
            v = cell.value
            if v is None:
                continue
            if isinstance(v, bool):
                w = 5
            elif isinstance(v, (int, float)):
                w = len(format(float(v), ',.2f'))
            else:
                s = str(v)
                w = sum(2 if ord(ch) > 127 else 1 for ch in s)
            if w > widths.get(cell.column, 0):
                widths[cell.column] = w
    for c, w in widths.items():
        t = min(max(w + 2, min_width), max_width)
        try:
            ws.column_dimensions[get_column_letter(c)].width = t
        except Exception:
            pass


def reorder_footnote_after_audit(wb):
    """统一排布（2026-08-02 用户方法论）：附注汇总 sheet 移到「审定表」之后、「明细表」之前。
    规则：sheet 名含『审定表』的在前、含『附注汇总』的其次、其余（明细表/检查表等）保持原相对顺序。
    无审定表（如研发支出）或 无附注汇总 → 不动。所有含附注汇总的底稿统一走本函数。"""
    sns = wb.sheetnames
    aud = [s for s in sns if '审定表' in s]
    fnt = [s for s in sns if '附注汇总' in s]
    if not aud or not fnt:
        return
    rest = [s for s in sns if s not in aud and s not in fnt]
    new_order = aud + fnt + rest
    wb._sheets = [wb[s] for s in new_order]


def _add_meta_sheets(wb, meta):
    """⚡ 2026-08-09（用户要求：审计完整性——统一抬头+审计调整分录模板）：
    每份底稿追加【仅「审计调整分录」】sheet（不破坏现有 sheet 结构）：
    审计调整载体模板（未审数→企业重分类→报表数→审计调整→审定数；
    无调整时审定数=未审数=TB）。meta={'comp','label','period'}。
    ⚡ 2026-08-10 用户要求：「底稿说明（含取数说明）」不要了，只留审计调整分录。"""
    try:
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        if not any('审计调整分录' == s for s in wb.sheetnames):
            ws = wb.create_sheet('审计调整分录')
            hdrs = ['调整编号', '调整类型', '借方科目编码', '借方科目名称', '借方金额',
                    '贷方科目编码', '贷方科目名称', '贷方金额', '调整原因及依据', '审计结论/备注']
            for j, h in enumerate(hdrs, 1):
                c = ws.cell(1, j, h)
                c.font = Font(bold=True)
                c.fill = PatternFill('solid', fgColor='DDEBF7')
                c.alignment = Alignment(horizontal='center')
                c.border = Border(*[Side(style='thin')] * 4)
            c = ws.cell(2, 1, '说明：审定数 = 未审数（TB期末） + 企业自行重分类 + 审计调整。本底稿无调整时审定数=未审数。'
                            '调整类型：重分类/差错更正/计提与转回/截止调整/其他。')
            c.font = Font(italic=True, size=9)
            c.alignment = Alignment(wrap_text=True)
            for j, w2 in enumerate([10, 12, 14, 22, 16, 14, 22, 16, 40, 24], 1):
                ws.column_dimensions[chr(64 + j)].width = w2
            ws.freeze_panes = 'A3'
    except Exception:
        pass


def finalize_workbook(wb):
    """统一收尾（各 builder 已显式设置字体与千分位，本函数作强制兜底）：
    ①字体强制 Times New Roman 10 号（覆盖模板/硬编码遗留的 Calibri/宋体/9 号）；
    ②数字统一会计格式、序号列不保留小数；③数据列列宽自适应；
    ④附注汇总 sheet 统一排布到审定表之后、明细表之前（2026-08-02）；
    ⑤若 wb._meta 存在 → 追加「审计调整分录」sheet（2026-08-09；底稿说明已按
    2026-08-10 用户要求移除）。
    每个 wb 只执行一次（_wb_style_unified 标记），避免重复全表扫描。"""
    reorder_footnote_after_audit(wb)
    meta = getattr(wb, '_meta', None)
    if meta is None:
        label = ''
        year = '2026年度'
        for s in wb.sheetnames:
            import re as _re
            m = _re.search(r'(20\d{2})', s)
            if m:
                year = m.group(1) + '年度'
            if '审定表' in s or ('汇总' in s and '表' in s):
                label = s.replace(' 审定表', '').replace('审定表', '').strip()
        if label:
            meta = {'comp': '', 'label': label, 'period': year}
            wb._meta = meta
    if meta:
        _add_meta_sheets(wb, meta)
    if getattr(wb, '_wb_style_unified', False):
        return
    wb._wb_style_unified = True
    # ⚡⚡ 2026-08-25 P0：审计程序执行说明 + 勾稽与异常检查 统一注入（改 1 处全生效）。
    #   在样式统一【之前】注入，使新 sheet 也走 _unify_sheet 统一样式。
    try:
        from audit_program import inject_program_sheets
        inject_program_sheets(wb, meta)
    except Exception:
        pass
    # ⚡⚡ 2026-08-30 批次A P0：构建时内建断言（合计/小计行 = 明细加总）。
    #   复用 audit_checker 口径（_mem_total_fix），默认 WARN 记录、AUDIT_FAIL_FAST=1 时中止；
    #   放在样式统一之前执行，保证读到的仍是 builder 写入的原始数值行。
    try:
        assert_build_totals(wb)
    except AssertionError:
        raise
    except Exception:
        pass
    for ws in wb.worksheets:
        try:
            _unify_sheet(ws)
        except Exception:
            pass
    # ⚡ 2026-08-11 方案五b：抽凭 sheet 顶部注入审计抽样规模（统一注入，改 1 处全生效）
    try:
        from audit_sampling import inject_all_sampling_sheets
        inject_all_sampling_sheets(wb)
    except Exception:
        pass


def unify_workbook_style(wb):
    """终审强制统一样式（覆盖 emit 按年拆分后重建的 sheet / 模板残留样式），
    供 recalc_totals 等生成后收尾环节复用；与 finalize_workbook 同一实现。"""
    reorder_footnote_after_audit(wb)
    for ws in wb.worksheets:
        try:
            _unify_sheet(ws)
        except Exception:
            pass


def build_merge_wide_sheet(wb, sheet_name, title, rows, ent_vals, prev_vals=None, merge_label='合并报表数', bold_rows=()):
    """横展式合并审定表（2026-08-11 阶段一 1.1 合并底稿差异化）：
    行=项目（rows: [(label, key)]），列=上年数 | 各主体(本期数) | 全集团合计 | 抵消借 | 抵消贷 | 合并报表数。
    ent_vals: {ent: {key: 数值}}（集团模式专属，单体底稿不调用）；
    prev_vals: {key: 上年数} 可选（无则上年数列留空）。
    集团合计=Σ主体（写数值）；抵消借/贷留空（待审计人员填，提示未抵消）；合并报表数=集团合计（抵消 0 时）。
    合并 4 列沿用排版强化色：集团合计中蓝/抵消借浅红/抵消贷浅绿/合并数深蓝白字。
    bold_rows：合计/小计行标签集合（2026-08-15 用户方法论：资产/负债/权益/损益段合计加粗+浅灰底）。
    返回创建的 sheet。"""
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import PatternFill
    _F_MID = PatternFill('solid', fgColor='BDD7EE')
    _F_RED = PatternFill('solid', fgColor='FCE4EC')
    _F_GRN = PatternFill('solid', fgColor='E2EFDA')
    _F_DBL = PatternFill('solid', fgColor='DDEBF7')
    _F_TOT = PatternFill('solid', fgColor='E7E6E6')
    _F_WHT = Font(name='Times New Roman', bold=True, size=10, color='000000')
    ws = wb.create_sheet(sheet_name)
    ents = sorted(ent_vals.keys())
    nc = 2 + len(ents) + 4                     # 项目 + 上年 + 主体 + 4 合并列
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, title)
    c.font = SHELL_TITLE_FONT; c.alignment = SHELL_CEN
    headers = ['项目', '上年数'] + ents + ['全集团合计', '抵消借', '抵消贷', merge_label]
    for j, h in enumerate(headers, 1):
        cell = ws.cell(2, j, h)
        cell.font = SHELL_HFONT; cell.fill = SHELL_HFILL
        cell.alignment = SHELL_CEN; cell.border = SHELL_BORDER
    _col_ent_end = 2 + len(ents)
    ws.cell(2, _col_ent_end + 1).fill = _F_MID
    ws.cell(2, _col_ent_end + 2).fill = _F_RED
    ws.cell(2, _col_ent_end + 3).fill = _F_GRN
    _mc = ws.cell(2, _col_ent_end + 4)
    _mc.fill = _F_DBL; _mc.font = _F_WHT
    r = 3
    for label, key in rows:
        tot = 0.0
        ws.cell(r, 1, label).alignment = SHELL_LEFT
        ws.cell(r, 1).border = SHELL_BORDER
        _is_tot = label in bold_rows
        if _is_tot:
            _bf = Font(name='Times New Roman', bold=True, size=10)
            ws.cell(r, 1).font = _bf
        if prev_vals and key in prev_vals:
            _pv = prev_vals[key]
            cc = ws.cell(r, 2, round(_pv, 2) if abs(_pv) >= 0.005 else None)
            cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
        for i, e in enumerate(ents):
            _v = ent_vals[e].get(key, 0.0)
            tot += _v
            cc = ws.cell(r, 3 + i, round(_v, 2) if abs(_v) >= 0.005 else None)
            cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
        cc = ws.cell(r, _col_ent_end + 1, round(tot, 2) if abs(tot) >= 0.005 else None)
        cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER; cc.fill = _F_MID
        for _off in (2, 3):                       # 抵消借/贷：留空待填（边框+底色提示）
            cc = ws.cell(r, _col_ent_end + _off, None)
            cc.border = SHELL_BORDER; cc.fill = _F_RED if _off == 2 else _F_GRN
        cc = ws.cell(r, _col_ent_end + 4, round(tot, 2) if abs(tot) >= 0.005 else None)
        cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT; cc.border = SHELL_BORDER
        cc.fill = _F_DBL; cc.font = _F_WHT
        if _is_tot:   # ⚡ 2026-08-15 段合计/小计行：加粗 + 浅灰底（用户方法论）
            for _c in range(1, nc + 1):
                _cell = ws.cell(r, _c)
                if _c == 1:
                    _cell.font = _bf
                elif _c <= _col_ent_end:
                    _cell.fill = _F_TOT
        r += 1
    ws.column_dimensions['A'].width = 26
    for j in range(2, nc + 1):
        ws.column_dimensions[get_column_letter(j)].width = 15
    ws.freeze_panes = 'C3'
    return ws


def assert_build_totals(wb, fail_fast=None):
    """构建时内建断言（2026-08-30 批次A P0）：写表前校验『合计/小计行 = 明细加总』。

    复用 audit_checker._mem_total_fix（纯内存、同口径：容差 0.005 / 公式单元格 /
    差异·勾稽列 / 集团合计 豁免）——把事后 audit_checker 的 E1 检查提前到生成时，
    避免"跑完全量才发现合计错误"。

    行为：
      - 默认 warn：发现问题记录到 wb._build_issues 并打印 [BUILD-ASSERT][WARN]（不中断量产）；
      - fail_fast=True 或环境变量 AUDIT_FAIL_FAST=1：任一 sheet 存在需修正合计 → 抛异常中止，
        防止坏底稿静默交付。
      - 大 sheet 保护：单 sheet 行数 > 上限（默认 20000，env AUDIT_ASSERT_MAX_ROWS 可调）跳过，
        交事后 audit_checker 兜底，避免大表 OOM（XBJ/current_account 10 万行场景）。

    返回 [(sheet_title, n_fix)]；异常降级返回 []（不闪退）。
    """
    import os as _os
    try:
        from audit_checker import _mem_total_fix
    except Exception:
        return []
    ff = fail_fast
    if ff is None:
        ff = _os.environ.get('AUDIT_FAIL_FAST', '') == '1'
    try:
        max_rows = int(_os.environ.get('AUDIT_ASSERT_MAX_ROWS', '20000'))
    except Exception:
        max_rows = 20000
    issues = []
    for ws in wb.worksheets:
        try:
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
        except Exception:
            continue
        if not rows or len(rows) > max_rows:
            continue
        try:
            n = _mem_total_fix(rows)
        except Exception:
            n = 0
        if n > 0:
            issues.append((ws.title, n))
    if issues:
        wb._build_issues = issues
        msg = '; '.join('%s:%d个合计单元格' % (s, n) for s, n in issues)
        if ff:
            raise AssertionError('[BUILD-ASSERT] 构建时内建断言未通过：%s' % msg)
        print('[BUILD-ASSERT][WARN] %s' % msg, flush=True)
    return issues
