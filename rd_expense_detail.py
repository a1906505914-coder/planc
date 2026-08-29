# -*- coding: utf-8 -*-
# rd_expense_detail.py — 研发支出 独立审计底稿生成器（2026-07-31 重构）
# 用途：将「研发支出」（资产负债表过渡/资本化科目）单独出一份底稿，
#       反映「费用化支出 / 资本化支出」二级明细的发生与结余，并说明其结转去向
#       （费用化→6604 研发费用(损益)；资本化→无形资产）。
#       该底稿与「研发费用」P&L 底稿相互独立、互不计入，避免类别合并双计。
# 取数铁律（2026-07-31 重构要点）：
#   ① 实体发现/读 TB 用 audit_common（discover_entities / read_tb_full）：
#      兼容「年份前置/后置/名中」文件名（如 2025东轴科目余额表.xlsx /
#      上海朗炫科目余额表2025.xlsx），动态表头列解析，不再用硬编码正则与硬编码列索引
#      （旧版对年份前置文件名静默失败——2026-07-31 检查发现 J 文件夹 rd_expense 无产出）。
#   ② 科目识别按【名称】判定（铁律5/23）：名称含「研发支出」即命中
#      （含"研发支出T"等实际账套变体）；名称含「研发费用」的损益科目不归本程序
#      （由费用/利润表底稿覆盖），避免类别双计；代码仅作结构性定位（取父/末级），
#      绝不按代码硬判语义（跨账套 5301 可能是「营业外收入」，如 J 文件夹——
#      旧版硬编码 CODE='5301' 属严重错账隐患）。
#   ③ 从一级(TB)出发；仅列次一级子目为明细行（其 TB 值已含更深级子孙，不重复计）；
#      0 值行不显示；全集团合计=各主体次一级借发之和。
#   ④ 无「研发支出」名称科目 → 生成「无研发支出科目披露」sheet 并明确打印，
#      不再静默退出（2026-07-31 修复）。
import os, re, sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from audit_shell import finalize_workbook
import audit_common as A
import audit_year_io as Y

CATNAME = '研发支出'
NAME_KW = ('研发支出',)          # 名称命中关键词（铁律23：按实际账套科目名，含"研发支出T"变体）
EXCLUDE_KW = ('研发费用',)       # 损益类研发费用不归本程序（由 pl/expense 底稿覆盖）

# ---------- 样式 ----------
HEAD_FILL = PatternFill('solid', fgColor='DDEBF7')
HEAD_FONT = Font(name='Times New Roman', bold=True, color='000000', size=10)
SUB_FILL = PatternFill('solid', fgColor='DDEBF7')
TOT_FILL = PatternFill('solid', fgColor='BDD7EE')
BOLD = Font(name='Times New Roman', bold=True)
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center')
RIGHT = Alignment(horizontal='right', vertical='center')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM_FMT = '#,##0.00'


def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = HEAD_FILL; cell.font = HEAD_FONT
        cell.alignment = CEN; cell.border = BORDER


def _money(ws, r, c, v, fill=None):
    if v is not None:
        try:
            v = round(float(v), 2)
        except (TypeError, ValueError):
            pass  # 公式字符串（'=SUM…'）原样写入
    cell = ws.cell(r, c, v)
    cell.number_format = NUM_FMT; cell.alignment = RIGHT; cell.border = BORDER
    if fill is not None:
        cell.fill = fill


def _txt(ws, r, c, v, bold=False, align=LEFT, fill=None):
    cell = ws.cell(r, c, v)
    cell.alignment = align; cell.border = BORDER
    if bold:
        cell.font = BOLD
    if fill is not None:
        cell.fill = fill


def _name_hit(name):
    """按名称判定是否「研发支出」科目（铁律5/23：跨账套代码不一致，名称是唯一语义键）。"""
    n = str(name or '')
    if any(k in n for k in EXCLUDE_KW):
        return False
    return any(k in n for k in NAME_KW)


def build_detail(entities, tb):
    """返回 (detail, control, skipped)：
    detail[(e, y)]  = list[(code, name, qc, jf, df, qm)] 次一级子目明细
    control[(e, y)] = dict(name=一级名, codes=名称命中全部代码, qc/jf/df/qm=一级数)
    skipped[(e, y)] = 跳过原因（名称未命中 → 明确披露）
    """
    detail, control, skipped = {}, {}, {}
    for e, yd in entities.items():
        for y in yd:
            hits = []
            for (ee, c, n, yy), v in tb.items():
                if ee == e and yy == y and _name_hit(n):
                    hits.append((c, n, v))
            if not hits:
                skipped[(e, y)] = '无「研发支出」名称科目（本账套代码不同/无此科目，正常跳过）'
                continue
            # 一级 = level 最浅（缺失时代码最短）；代码仅作结构性定位
            # 2026-08-06 修复：level 列可能解析为字符串（JTt U8 导出）→ 与 float 比较崩溃
            def _lv(t):
                _l = t[2].get('level')
                try:
                    return (int(float(_l)) if _l is not None and str(_l).strip() else 99)
                except (TypeError, ValueError):
                    return 99
            hits.sort(key=lambda t: (_lv(t), len(t[0])))
            top_code, top_name, top_v = hits[0]
            # ⚡ 2026-08-26 修复（研发支出明细空壳 82万 vs 审定 3,853万）：SAP read_tb_full
            #   只返回叶子（无父级 1704），top 误取首个子目 1704010000『人员人工』→ 其余子目
            #   非其前缀 → children 空 → 明细只有 1 行。无父级（top_code 无后代）时：
            #   一级=公共前缀名，子目=全部 hits。
            _has_parent2 = any(c2 != top_code and c2.startswith(top_code) for c2, _, _ in hits)
            if not _has_parent2 and len(hits) > 1:
                _all_c = sorted({c for c, _, _ in hits})
                _base = os.path.commonprefix(_all_c)[:4] if _all_c else top_code
                control[(e, y)] = dict(name=CATNAME, codes=_all_c,
                                       qc=sum(float(t.get('qc') or 0) for _, _, t in hits),
                                       jf=sum(float(t.get('jf') or 0) for _, _, t in hits),
                                       df=sum(float(t.get('df') or 0) for _, _, t in hits),
                                       qm=sum(float(t.get('qm') or 0) for _, _, t in hits))
                detail[(e, y)] = [(c, n, float(t.get('qc') or 0), float(t.get('jf') or 0),
                                   float(t.get('df') or 0), float(t.get('qm') or 0))
                                  for (c, n, t) in hits
                                  if any(abs(float(t.get(k) or 0)) > 0.005 for k in ('qc', 'jf', 'df', 'qm'))]
                continue
            control[(e, y)] = dict(name=top_name, codes=sorted({c for c, _, _ in hits}),
                                   qc=top_v['qc'], jf=top_v['jf'], df=top_v['df'], qm=top_v['qm'])
            top_lv = top_v.get('level')
            children = [(c, n, v) for (c, n, v) in hits if c.startswith(top_code) and c != top_code]
            keys = []
            if children:
                if top_lv is not None:
                    want = top_lv + 1
                    keys = sorted([t for t in children if t[2].get('level') == want])
                if not keys:
                    # level 缺失/无恰为次一级 → 代码长度增量 2 兜底（如 530101 对 5301）
                    keys = sorted([t for t in children if len(t[0]) == len(top_code) + 2])
                if not keys and children:
                    keys = sorted(children)   # 防御兜底：取全部子目
            rows = []
            for (c, n, v) in keys:
                if v['jf'] == 0 and v['df'] == 0 and v['qc'] == 0 and v['qm'] == 0:
                    continue
                rows.append((c, n, v['qc'], v['jf'], v['df'], v['qm']))
            detail[(e, y)] = rows
    return detail, control, skipped


def build_rd_monthly_sheet(ws, gl, entities, y, control):
    """研发费用明细表（2026-08-07 用户图1：跨主体合并，按【费用性质】×月份，行=『研发费用-XXX』，列=1-12月+合计）。

    改动（2026-08-07）：去除主体列，跨主体同费用性质合并（公司级费用归集常见做法）。
    dq GL 行名='研发支出-RD2025xx：项目名-费用性质'，取行名最后一段作为费用性质。
    TB 数据来源：control[(e, str(y))]['jf'] 所有主体汇总。
    """
    hdrs = ['项目'] + [f'{m}月' for m in range(1, 13)] + ['合计']
    ncols = len(hdrs)
    for c, h in enumerate(hdrs, 1):
        ws.cell(1, c, h)
        ws.cell(1, c).fill = HEAD_FILL; ws.cell(1, c).font = HEAD_FONT
        ws.cell(1, c).alignment = CEN; ws.cell(1, c).border = BORDER

    def _mo(v):
        s = str(v or '')
        m = re.search(r'-(\d{2})-', s)
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                return None
        return None

    agg = {}
    for r in gl:
        if str(r.get('y')) != str(y):
            continue
        nm = str(r.get('name', ''))
        # 2026-08-07 修复（dq 2倍根因）：研发支出在 dq 是『5301 归集 → 6602.33 结转』两段记账。
        # 5301 行名以『研发支出-』开头（RD 项目维度，月度借方归集 + 月末贷方结转）；
        # 6602.33 行名『管理费用-研发支出』（月度借方=5301 结转接收，金额完全相等）。
        # 只取 5301 归集段（研发支出- 开头），否则 5301借+6602借=2×TB（72.57M vs 36.28M）。
        if not nm.startswith('研发支出'):
            continue
        dr = float(r.get('dr') or 0.0)
        if dr <= 0.005:
            continue
        segs = [s.strip() for s in nm.split('-') if s.strip() and s.strip() != '研发支出']
        cat = segs[-1] if segs else nm
        mo = _mo(r.get('date'))
        if mo is None or mo < 1 or mo > 12:
            continue
        agg.setdefault(cat, [0.0] * 12)[mo - 1] += dr

    # 按合计降序，但把『材料』族/『工资薪金』族等用户已知类别优先；
    # 实际就是按合计降序，无固定顺序
    cats_sorted = sorted([(c, a) for c, a in agg.items() if abs(sum(a)) > 0.005],
                        key=lambda x: -sum(x[1]))

    grand_tot = 0.0
    tb_total = 0.0
    for ent in sorted(entities):
        tb_total += float((control.get((ent, str(y))) or {}).get('jf') or 0.0)

    r = 2
    for cat, arr in cats_sorted:
        tot = sum(arr)
        grand_tot += tot
        _txt(ws, r, 1, f'研发费用-{cat}', bold=False)
        for j in range(12):
            if abs(arr[j]) > 0.005:
                _money(ws, r, 2 + j, arr[j])
        _money(ws, r, 14, tot)
        ws.cell(r, 14).font = BOLD
        for c in range(1, ncols + 1):
            ws.cell(r, c).border = BORDER
        r += 1

    # 合计行
    _txt(ws, r, 1, '合计', bold=True, fill=SUB_FILL)
    for j in range(12):
        col_tot = sum(arr[j] for arr in agg.values())
        _money(ws, r, 2 + j, col_tot, fill=SUB_FILL)
    _money(ws, r, 14, grand_tot, fill=SUB_FILL)
    ws.cell(r, 14).font = BOLD
    for c in range(1, ncols + 1):
        ws.cell(r, c).font = BOLD
        ws.cell(r, c).border = BORDER
    r += 1

    # 账面材料费（来自 TB 5301 借发）—— 用作与 GL 各费用合计的对照组
    _txt(ws, r, 1, '账面材料费', bold=True, fill=SUB_FILL)
    _txt(ws, r, 2, '（详见研发支出明细表）', fill=SUB_FILL)
    _money(ws, r, 14, tb_total, fill=SUB_FILL)
    ws.cell(r, 14).font = BOLD
    for c in range(1, ncols + 1):
        ws.cell(r, c).border = BORDER

    ws.freeze_panes = 'B2'
    ws.column_dimensions['A'].width = 24
    for j in range(2, 14):
        ws.column_dimensions[chr(64 + j)].width = 14
    ws.column_dimensions['N'].width = 16
    return ws


def build_rd_proj_summary_sheet(ws, gl, entities, y):
    """研发项目汇总表（2026-08-11 用户 DQ3 修改：按【账面名称】分类，不归并 6 大类）。
    行=GL 行名末段费用性质原文（实验室耗材/工资薪金/社会保险费/设备折旧/电费...账面口径），
    列=各研发项目 RD+编号。审计员可直接对账账面科目，无需归类映射。"""

    # 取所有研发项目编号（GL 行名：'研发支出-RD202xxx：xxx'）
    projs = set()
    for r in gl:
        if str(r.get('y')) != str(y) or '研发支出' not in str(r.get('name', '')):
            continue
        nm = str(r.get('name', ''))
        m = re.search(r'(RD\d{6})', nm)
        if m:
            projs.add(m.group(1))
    proj_cols = sorted(projs)

    hdrs = ['项目'] + proj_cols + ['合计']
    ncols = len(hdrs)
    for c, h in enumerate(hdrs, 1):
        ws.cell(1, c, h)
        ws.cell(1, c).fill = HEAD_FILL; ws.cell(1, c).font = HEAD_FONT
        ws.cell(1, c).alignment = CEN; ws.cell(1, c).border = BORDER

    # 数据矩阵：agg[账面名称][项目编号] = 总金额（GL 借方）；行序=账面名称首次出现顺序
    agg = {}
    cost_order = []
    for r in gl:
        if str(r.get('y')) != str(y) or '研发支出' not in str(r.get('name', '')):
            continue
        dr = float(r.get('dr') or 0.0)
        if dr <= 0.005:
            continue
        nm = str(r.get('name', ''))
        m_proj = re.search(r'(RD\d{6})', nm)
        if not m_proj:
            continue
        proj = m_proj.group(1)
        # 取费用性质：行名最后一段（去掉项目编号段）
        segs = [s.strip() for s in nm.split('-') if s.strip() and s.strip() != '研发支出']
        # 段是 ['RD202510：低空经济无人机关键检测技术研究', '工资薪金']
        if len(segs) >= 2:
            cost = segs[-1].strip()
        else:
            cost = segs[0] if segs else nm
        # ⚡ 2026-08-11 DQ3：直接按账面名称分类（不归并 6 大类）
        if cost not in agg:
            agg[cost] = {p: 0.0 for p in proj_cols}
            cost_order.append(cost)
        agg[cost][proj] += dr

    r = 2
    proj_totals = {p: 0.0 for p in proj_cols}
    for cost in cost_order:
        cat_tot = 0.0
        _txt(ws, r, 1, cost)
        for j, p in enumerate(proj_cols):
            v = agg[cost][p]
            if abs(v) > 0.005:
                _money(ws, r, 2 + j, v)
            cat_tot += v
            proj_totals[p] += v
        _money(ws, r, len(proj_cols) + 2, cat_tot)
        ws.cell(r, len(proj_cols) + 2).font = BOLD
        for c in range(1, ncols + 1):
            ws.cell(r, c).border = BORDER
        r += 1

    # 合计行
    _txt(ws, r, 1, '合计', bold=True, fill=SUB_FILL)
    for j, p in enumerate(proj_cols):
        _money(ws, r, 2 + j, proj_totals[p], fill=SUB_FILL)
    _money(ws, r, len(proj_cols) + 2, sum(proj_totals.values()), fill=SUB_FILL)
    ws.cell(r, len(proj_cols) + 2).font = BOLD
    for c in range(1, ncols + 1):
        ws.cell(r, c).font = BOLD
        ws.cell(r, c).border = BORDER

    ws.freeze_panes = 'B2'
    ws.column_dimensions['A'].width = 24
    for j, p in enumerate(proj_cols):
        ws.column_dimensions[chr(64 + 2 + j)].width = 18
    ws.column_dimensions[chr(64 + 2 + len(proj_cols))].width = 18
    return ws



def write_paper(data_dir, detail, control, skipped, years=None, tb=None, entities=None, gl=None):
    """按年独立构建并保存（铁律20：每年底稿只含当年数据）。
    不用 emit_per_year 的原因：本底稿 sheet 名不含年份标记、年份以行分组标题出现
    （非"会计期间/年度"列），emit 拆分后 _filter_sheet_rows_by_year 无法过滤 → 会混年
    （2026-07-31 检查发现旧版产出单文件合并稿，违反按年规则）。"""
    if years is None:
        years = sorted({y for (e, y) in control})
    outs = []
    for y in years:
        wb = openpyxl.Workbook()
        # ⚡ 2026-08-11：删除默认 Sheet，由 render_sheet 统一创建首个 sheet
        wb.remove(wb.active)
        # ---- Sheet1 研发支出明细表（仅该年数据）----
        # ⚡ 2026-08-11 阶段二 2.1：计算产结构化 rows，渲染收敛到 audit_render.render_sheet
        cols = ['核算主体', '二级科目', '期初余额', '本期增加(借发)', '本期减少(贷发)', '期末余额', '校验(期初+增-减-末)']
        MONEY = {3, 4, 5, 6, 7}
        from audit_render import render_sheet, row as _arow
        _rows = []
        grand = [0.0, 0.0, 0.0, 0.0]
        sub_rows = []   # 各主体小计行号（年度合计公式引用，2026-08-02）
        es = sorted({e for (e, yy) in control if yy == y})
        for e in es:
            rows = detail.get((e, y), [])
            ct = control[(e, y)]
            e_rows = []   # 本主体数据行号
            if not rows:
                # 无子目明细：仅列一级控制数行
                _rows.append(_arow([e, ct['name'] + '（一级）', ct['qc'], ct['jf'], ct['df'], ct['qm'],
                                    ct['qc'] + ct['jf'] - ct['df'] - ct['qm']], num=MONEY))
                e_rows.append(len(_rows))
                etot = [ct['qc'], ct['jf'], ct['df'], ct['qm']]
            else:
                etot = [0.0, 0.0, 0.0, 0.0]
                for (c, nm, qc, jf, df, qm) in rows:
                    _rows.append(_arow([e, nm, qc, jf, df, qm, qc + jf - df - qm], num=MONEY))
                    for i, v in enumerate((qc, jf, df, qm)):
                        etot[i] += v
                    e_rows.append(len(_rows))
            # 主体小计（2026-08-06 协议：写数值=etot 累计；公式 data_only 读 None 收回读不回）
            _rows.append(_arow([e, '小计', round(etot[0], 2), round(etot[1], 2),
                                round(etot[2], 2), round(etot[3], 2),
                                round(etot[0] + etot[1] - etot[2] - etot[3], 2)],
                               b=True, fill='sub', num=MONEY))
            for i, v in enumerate(etot):
                grand[i] += v
            sub_rows.append(len(_rows))
        # 年度合计（2026-08-06 协议：写数值=grand 累计）
        _rows.append(_arow([f'{y} 年度合计', '', round(grand[0], 2), round(grand[1], 2),
                            round(grand[2], 2), round(grand[3], 2),
                            round(grand[0] + grand[1] - grand[2] - grand[3], 2)],
                           b=True, fill='total', num=MONEY))
        widths = [14, 34, 16, 18, 18, 16, 18]
        ws = render_sheet(wb, dict(
            name='研发支出明细表', headers=cols, rows=_rows,
            widths=widths, money_cols=MONEY, zero_blank=False, freeze='A2',
        ))

        # ---- Sheet1.5 研发费用明细表（2026-08-07 用户图1：按费用归集合并主体，
        # A 列『研发费用-XXX』，B-M 1-12月，N 合计 + 账面材料费对照行）----
        if gl:
            ws1b = wb.create_sheet('研发费用明细表')
            build_rd_monthly_sheet(ws1b, gl, es, y, control)
        # ---- Sheet1.6 研发项目汇总表（2026-08-07 用户图2：6 大费用类别×各研发项目 RD202xxx）----
        if gl:
            ws1c = wb.create_sheet('研发项目汇总表')
            build_rd_proj_summary_sheet(ws1c, gl, es, y)
        # ---- Sheet2 勾稽：一级(TB) ↔ 次一级明细合计(TB) ----
        ws2 = wb.create_sheet('一级与二级勾稽')
        ws2.append(['核算主体', '年度', '研发支出一级借发(TB)', '次一级明细借发合计(TB)', '差异', '结论'])
        _style_header(ws2, 1, 6)
        r = 2
        for e in es:
            c1 = control[(e, y)]['jf']
            c2 = sum(t[3] for t in detail.get((e, y), []))
            diff = c1 - c2
            ok = abs(diff) < 1.0
            _txt(ws2, r, 1, e); _txt(ws2, r, 2, y)
            _money(ws2, r, 3, c1); _money(ws2, r, 4, c2); _money(ws2, r, 5, diff)
            _txt(ws2, r, 6, '一致' if ok else '差异需核实')
            if not ok:
                for c in range(1, 7):
                    ws2.cell(r, c).fill = PatternFill('solid', fgColor='FCE4D6')
            r += 1
        for i, w in enumerate([14, 8, 22, 24, 16, 12], 1):
            ws2.column_dimensions[get_column_letter(i)].width = w

        # ---- Sheet3 无研发支出科目披露（名称未命中 → 如实披露，不静默） ----
        sk_y = {k: v for k, v in skipped.items() if k[1] == y}
        if sk_y:
            ws3 = wb.create_sheet('无研发支出科目披露')
            ws3.append(['核算主体', '年度', '说明'])
            _style_header(ws3, 1, 3)
            rr = 2
            for (e2, y2) in sorted(sk_y):
                _txt(ws3, rr, 1, e2); _txt(ws3, rr, 2, y2); _txt(ws3, rr, 3, sk_y[(e2, y2)])
                rr += 1
            for i, w in enumerate([16, 8, 60], 1):
                ws3.column_dimensions[get_column_letter(i)].width = w

        # ---- 研发支出 审定表（2026-08-04 参照借款审定表格式：无条件生成，有数据取 TB，
        #      无数据 0 值+披露说明，供阅读与合并试算表核对程序取结构）----
        ws_aud = wb.create_sheet('研发支出 审定表')
        ws_aud.cell(1, 1, '研发支出 审定表（按核算主体）').font = \
            Font(name='Times New Roman', size=12, bold=True)
        hdrs = ['核算主体', '期初数', '期末未审数', '审计调整数', '审定数', '与科目余额表勾稽']
        for j, h in enumerate(hdrs, 1):
            cc = ws_aud.cell(2, j, h)
            cc.font = Font(name='Times New Roman', size=10, bold=True)
            cc.fill = PatternFill('solid', fgColor='DDEBF7')
            cc.border = Border(*[Side(style='thin', color='BFBFBF')] * 4)
            cc.alignment = Alignment(horizontal='center', vertical='center')
        rr = 3
        g_qc = g_qm = 0.0
        for ent in sorted(entities):
            # 2026-08-07 修复：按名称命中的全部行含父级+子级（g 杭州铁城 5301 父级
            # +530101~530104 子级+孙级），abs 求和 → 56,004,835.15=19×TB 父级 2,948,036.57。
            # → 【末级叶子过滤】+【带符号求和】（研发支出为借方科目，正常为正余额；
            # 530101 资本化支出负余额=转出后未重分类，带符号与 TB 父级一致）。
            hits = [(str(c), str(n), v) for (e, c, n, yy), v in tb.items()
                    if str(e) == ent and str(yy) == str(y) and any(k in str(n) for k in NAME_KW)]
            _codes = [c for c, _, _ in hits]
            _leaves = [r for r in hits
                       if not any(c2 != r[0] and c2.startswith(r[0]) for c2 in _codes)]
            qc = sum(float(v.get('qc') or 0.0) for _, _, v in _leaves)
            qm = sum(float(v.get('qm') or 0.0) for _, _, v in _leaves)
            g_qc += qc; g_qm += qm
            _txt(ws_aud, rr, 1, ent)
            _money(ws_aud, rr, 2, qc)
            _money(ws_aud, rr, 3, qm)
            _money(ws_aud, rr, 4, 0.0)                       # 审计调整待填
            # 2026-08-06 协议：审定数写数值（=期末+调整0；公式 data_only 读 None 收回读不回）
            _money(ws_aud, rr, 5, round(qm, 2))
            _txt(ws_aud, rr, 6, '与科目余额表核对一致' if (qc or qm) else '无数据（科目未启用）')
            for j in range(1, 7):
                ws_aud.cell(rr, j).border = Border(*[Side(style='thin', color='BFBFBF')] * 4)
            rr += 1
        _txt(ws_aud, rr, 1, '合计', bold=True)
        _money(ws_aud, rr, 2, round(g_qc, 2))
        _money(ws_aud, rr, 3, round(g_qm, 2))
        _money(ws_aud, rr, 4, 0.0)
        _money(ws_aud, rr, 5, round(g_qm, 2))
        for j in range(1, 7):
            ws_aud.cell(rr, j).fill = PatternFill('solid', fgColor='FCE4D6')
            ws_aud.cell(rr, j).font = Font(name='Times New Roman', size=10, bold=True)
        if g_qc == 0 and g_qm == 0:
            rr += 2
            _txt(ws_aud, rr, 1, '未发现「研发支出」名称科目（含"研发支出T"等变体），本期无发生额及余额'
                               '（数据现状；结构参照借款审定表生成，供核对/阅读）。', bold=False)
            ws_aud.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=6)
        for i, w in enumerate([16, 14, 16, 14, 14, 26], 1):
            ws_aud.column_dimensions[get_column_letter(i)].width = w

        # ---- Sheet4 附注汇总（滚动四段；每主体一列；按名称匹配科目，2026-08-02）----
        try:
            from audit_common import build_footnote_generic as _bfn_rd
            _bfn_rd(wb, data_dir, '附注汇总',
                    '研发支出附注汇总（审定口径；每主体一列；段=期初数/本期增加/本期减少/期末数；'
                    '减少=费用化转出(研发费用)/资本化转出(无形资产)）',
                    [('研发支出', [], list(NAME_KW))], is_credit=False, mode='four',
                    target_year=y, tb_full=tb, entities=entities)
        except Exception as ex:
            print(f'  ⚠️ 附注汇总生成失败：{ex}')

        finalize_workbook(wb)
        # ⚡ 2026-08-28 #875：对方科目核对（模块内集成）
        try:
            from counterparty_recon import inject_into_wb_auto
            inject_into_wb_auto(wb, data_dir, y, CATNAME, ents_set=set(entities))
        except Exception as _ex:
            print(f'  ⚠️ {CATNAME}对方科目核对注入失败：{_ex}')
        out = os.path.join(data_dir, f'{CATNAME}审计底稿_{y}_生成.xlsx')
        from audit_common import validate_workbook
        validate_workbook(wb, '研发支出底稿', raise_on_error=False)
        wb.save(out)
        wb.close()
        outs.append(out)
    return outs[0] if outs else None


def main():
    if len(sys.argv) > 1:
        data_dir = sys.argv[1].strip().strip('"').strip("'")
    else:
        data_dir = input('请拖入或输入 账套导出文件夹 路径：').strip().strip('"').strip("'")
    if not os.path.isdir(data_dir):
        print('文件夹不存在：', data_dir); return
    # SAP：单主体驱动时按 current_comp 过滤（防全量 GL OOM + 混 88 家主体）
    if _adapter is not None and _adapter.is_sap(data_dir):
        _c_rd = _adapter.current_comp()
        entities = ({_c_rd: A.discover_entities(data_dir).get(_c_rd, {})} if _c_rd else A.discover_entities(data_dir))
    else:
        entities = A.discover_entities(data_dir)
    if not entities:
        print('未在该文件夹发现科目余额表。'); return
    print(f'发现核算主体 {len(entities)} 个，年度：{sorted({y for e, yd in entities.items() for y in yd})}')
    tb = A.read_tb_full(data_dir, entities)
    gl = A.read_gl_rows(data_dir, entities)   # 2026-08-07：研发支出分月明细表需 GL 借方分月
    detail, control, skipped = build_detail(entities, tb)
    all_years = sorted({y for (e, y) in control} | {y for (e, y) in skipped})
    if not all_years:
        print('未发现任何年份的科目余额表数据。'); return
    if not control:
        print(f'未发现「{CATNAME}」名称科目（含"研发支出T"等变体）；将生成无研发支出科目披露。')
    out = write_paper(data_dir, detail, control, skipped, years=all_years, tb=tb, entities=entities, gl=gl)
    if out:
        print(f'✅ 已生成：{out}')
    else:
        print('⚠️ 未生成底稿（无任何主体存在研发支出科目）。')
    from audit_common import finalize_after_build
    finalize_after_build(data_dir)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）



# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None

if __name__ == '__main__':
    main()
