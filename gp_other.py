# -*- coding: utf-8 -*-
"""其他类科目集团底稿小程序（合并版）
================================================================
由 group_paper(通用框架) + gp_generic(通用生成/辅助) + gp_other(其他类科目) 合并而来，
不再依赖独立的辅助库文件。所有 G./GP. 别名均指向本模块自身。

覆盖非 往来/税费/收入/损益/货币资金/存货/长期资产/权益/借款 的余额类科目：
  1101 交易性金融资产 / 1511 长期股权投资 / 1512 长投减值 / 1711 商誉 /
  2201 应付票据 / 2703 租赁负债 / 2711 专项应付款 / 2801 预计负债 / 2901 递延所得税负债
每个文件夹=集团，仅对 TB 有数据的科目生成底稿（无数据跳过）。
结构：审定表 + 明细表(分年度) + 科目内核对 + 凭证抽查 + 说明。
"""

import paths as P
import os, sys, glob, re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import audit_common as A

# ===================== group_paper（通用框架） =====================
# -*- coding: utf-8 -*-
"""集团审计底稿 —— 通用框架（共享层）
================================================================
落实用户 2026-07-27 重构方向：每个文件夹=集团，每个科目×年度=一份集团审计底稿。
本模块提供所有科目通用的：
  · 样式 / _txt / _money
  · 末级(叶子)TB 控制数（避免 U8 父级=子目合计 父+子双计）
  · 辅助核算余额表读取（去重 glob + 兼容表头）
  · GL 过滤 / 凭证抽查（统一调用 audit_common.voucher_should_skip）
  · 审定表（集团合计 + 各账套主体横向分列）
  · 凭证抽查（行级带「账套主体」列）
  · 说明
各科目脚本（gp_ar / gp_rd / gp_cash / gp_inventory …）import 本模块，
并提供自身的 明细表 / 科目内核对 / 分析 sheet 构建函数。

铁律复用：
  · tb_leaf_codes 仅累加末级科目（铁律3/6），返回 (期初,借,贷,期末) 顺序，与 measures 索引一致。
  · discover_aux 两 glob 模式会命中同一文件（辅助核算余额表 含 辅助核算）→ dict.fromkeys 去重。
  · voucher_should_skip 返回 (skip, reason) 元组，必须解包 sk,_ = ...。
"""

# ---------- 样式 ----------
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TITLE_FONT = Font(name='Times New Roman', bold=True, size=10, color='000000')
HDR_FILL = PatternFill('solid', fgColor='DDEBF7')
HDR_FONT = Font(name='Times New Roman', bold=True, color='000000', size=10)
SUB_FILL = PatternFill('solid', fgColor='D9E1F2')
SUB_FONT = Font(name='Times New Roman', bold=True, size=10, color='000000')
TOT_FILL = PatternFill('solid', fgColor='FCE4D6')
TOT_FONT = Font(name='Times New Roman', bold=True, size=10)
CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center')
RIGHT = Alignment(horizontal='right', vertical='center')

GENERIC_AUX = {'客户', '供应商', '科目', '现金流量', '客户/供应商', '往来单位',
               '辅助核算名称', '部门', '职员', '项目', '现金流量项目'}


def _txt(ws, r, c, v, font=None, fill=None, align=None, border=True):
    cell = ws.cell(r, c, v)
    if font: cell.font = font
    if fill: cell.fill = fill
    cell.alignment = align or LEFT
    if border: cell.border = BORDER
    return cell


def _money(ws, r, c, v, bold=False, fill=None):
    cell = ws.cell(r, c, round(v or 0.0, 2))
    cell.number_format = '#,##0.00'
    cell.alignment = RIGHT
    cell.border = BORDER
    if bold: cell.font = TOT_FONT
    if fill: cell.fill = fill
    return cell


def _norm(code):
    return re.sub(r'\D', '', str(code))


# ---------- 末级 TB 控制数 ----------
def tb_leaf_codes(tb, ent, year, prefixes, names=None):
    """返回 (期初, 借, 贷, 期末)，仅累加【末级(叶子)科目】。
    prefixes: 科目代码前缀列表（如 ['1122'] / ['5301'] / ['1001','1002','1012'] / ['14']）。
    叶子 = 同前缀族内不是任何其他代码前缀的代码。
    names（可选，名称主键 铁律58）：给定时，命中行名称须含任一 names 才计入——
    防同代码跨账套异义（AS 2401=递延收益 vs AQ 2401=一年内到期，ncl 传 ['一年内到期']）。"""
    fam = []
    for (e, code, name, y), v in tb.items():
        nc = _norm(code)
        if e == ent and y == str(year) and any(nc.startswith(p) for p in prefixes):
            if names and not any(nm and str(nm) in str(name) for nm in names):
                continue
            fam.append((nc, v))
    norm = [c for c, _ in fam]
    qc = jf = df = qm = 0.0
    for c, v in fam:
        is_leaf = not any(other != c and other.startswith(c) for other in norm)
        if is_leaf:
            qc += v['qc']; jf += v['jf']; df += v['df']; qm += v['qm']
    return qc, jf, df, qm


# ---------- 辅助核算余额表读取 ----------
def discover_aux(data_dir, entity, year):
    pats = [os.path.join(data_dir, f'*{entity}*辅助核算余额表*{year}*.xlsx'),
            os.path.join(data_dir, f'*{entity}*辅助核算*{year}*.xlsx')]
    fns = []
    for p in pats:
        fns += glob.glob(p)
    fns = list(dict.fromkeys(fns))  # 去重：两 glob 模式会命中同一文件
    if not fns:
        for fn in sorted(os.listdir(data_dir)):
            if entity in fn and '辅助核算' in fn and str(year) in fn and fn.endswith('.xlsx'):
                fns.append(os.path.join(data_dir, fn))
    return fns


def _cidx(hdr, *subs):
    for s in subs:
        for j, h in enumerate(hdr):
            if h and s in str(h):
                return j
    return None


def read_aux_generic(data_dir, entity, year, name_contains, exclude_names=()):
    """读取某账套主体某年辅助核算余额表中 指定科目 的末级逐户行。
    name_contains: 科目名称包含此串（如 '应收账款'）；exclude_names: 排除含这些串的行（如 '准备'/'坏账'）。
    返回 [{'entity','cust','open','debit','credit','close'}]，signed 余额（借+/贷-）。"""
    out = []
    for fp in discover_aux(data_dir, entity, year):
        try:
            wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception as e:
            print(f'  ⚠️ aux读取失败 [{entity}][{year}] {fp}: {e}')
            continue
        hi = None
        for i, row in enumerate(rows):
            cells = [str(c) for c in row if c is not None]
            if any('科目名称' in c for c in cells) and \
               any(('辅助核算名称' in c or '往来单位' in c or '客户名称' in c) for c in cells):
                hi = i
                break
        if hi is None:
            continue
        hdr = [str(c) for c in rows[hi]]
        i_km = _cidx(hdr, '科目名称')
        i_name = _cidx(hdr, '科目辅助核算名称', '辅助核算名称', '往来单位名称', '往来单位', '客户名称')
        i_deb = _cidx(hdr, '借方')
        i_cred = _cidx(hdr, '贷方')
        amt_cols = [j for j, h in enumerate(hdr) if str(h).strip() == '金额']
        dir_cols = [j for j, h in enumerate(hdr) if str(h).strip() == '方向']
        i_open = amt_cols[0] if amt_cols else None
        i_close = amt_cols[-1] if len(amt_cols) > 1 else (amt_cols[0] if amt_cols else None)
        od = dir_cols[0] if dir_cols else None
        cd = dir_cols[1] if len(dir_cols) > 1 else (dir_cols[0] if dir_cols else None)
        i_level = _cidx(hdr, '等级')
        if None in (i_km, i_name, i_open, i_deb, i_cred, i_close):
            continue

        def num(x):
            try:
                return float(x)
            except Exception:
                return 0.0

        def sgn(amt, dcol, row):
            if dcol is None or dcol >= len(row) or not row[dcol]:
                return amt
            d = str(row[dcol]).strip()
            if d in ('贷', '贷入', 'cr', 'credit', '-', '2', '0'):
                return -abs(amt)
            return abs(amt)

        for row in rows[hi + 1:]:
            km = str(row[i_km]) if i_km < len(row) and row[i_km] is not None else ''
            if not (name_contains in km and all(x not in km for x in exclude_names)):
                continue
            name = str(row[i_name]) if i_name < len(row) and row[i_name] is not None else ''
            name = name.strip()
            if not name or name in GENERIC_AUX:
                continue
            lvl = str(row[i_level]) if (i_level is not None and i_level < len(row) and row[i_level] is not None) else ''
            if lvl.strip() in ('1', '2') and ('科目' in km or name in GENERIC_AUX):
                continue
            o = sgn(num(row[i_open]) if i_open < len(row) else 0.0, od, row)
            cl = sgn(num(row[i_close]) if i_close < len(row) else 0.0, cd, row)
            d = num(row[i_deb]) if i_deb < len(row) else 0.0
            c = num(row[i_cred]) if i_cred < len(row) else 0.0
            out.append({'entity': entity, 'cust': name, 'open': o, 'debit': d, 'credit': c, 'close': cl})
    return out


# ---------- GL 过滤 / 凭证抽查 ----------
def gl_filter_rows(gl, name_prefix, exclude_names=()):
    return [r for r in gl
            if str(r['name']).startswith(name_prefix)
            and all(x not in str(r['name']) for x in exclude_names)]


def vouched_rows(gl, name_prefix, exclude_names=()):
    """返回应抽凭的 GL 行（已排除 计提/补提/结转/摊销 与 相符计提/结转 凭证 + 同号凭证内计提单行）。
    三层过滤：
      ① 行级：摘要含 计提/补提/补计/结转/摊销/分摊/分配 的内部归集分录直接剔除。
      ② 凭证级：按整笔凭证(e,y,凭证字,凭证号)聚合后传入 voucher_should_skip(lines=...)，
         排除 结转本年利润 与 整笔相符计提/摊销。
      ③ 行级纯计提（2026-07-27）：同号凭证混合真实付款时，计提单行(opp=应付职工薪酬等
         计提负债、无真实结算)仍逐行排除，真实付款行(opp=银行/应付票据等)保留。
    仅保留真实的外部交易分录（如 借研发支出 贷银行/供应商应付、借应收账款 贷主营业务收入）。"""
    rows = gl_filter_rows(gl, name_prefix, exclude_names)
    # 2026-07-31 移除 ALLOC_KW 摘要行级预过滤：摘要含"结转/分摊/分配"的真实业务
    # （如"分摊递延收益"确认其他收益）会被误删导致抽凭缺——现统一由
    # ②凭证级 voucher_should_skip（结转损益/整笔纯计提）与 ③行级 is_accrual_line
    # （看对方科目是否为计提负债）精确排除，不再按摘要关键词粗筛。
    from collections import OrderedDict
    groups = OrderedDict()
    for r in rows:
        key = (r['e'], r['y'], r['vtype'], r['vno'])
        groups.setdefault(key, []).append(r)
    out = []
    for grp in groups.values():
        lines = [{'name': x['name'], 'opp': x['opp'], 'dr': x['dr'], 'cr': x['cr'], 'summary': x['summary']}
                 for x in grp]
        sk, _ = A.voucher_should_skip(summary=grp[0]['summary'], name='', opp='', lines=lines)
        if sk:
            continue  # 整笔结转/纯计提 → 全排除
        # ③ 行级：剔除同号凭证中残留的"相符计提"单行
        for x in grp:
            if A.is_accrual_line({'name': x['name'], 'opp': x['opp'], 'dr': x['dr'], 'cr': x['cr'], 'summary': x['summary']}):
                continue
            out.append(x)
    return out


# ---------- 审定表（账套主体纵向排列 + 期末审计调整/审定数） ----------
def build_summary_sheet(wb, title, ctrl_fn, ent_list, years, is_credit=False, display_credit_positive=False, sheet_name=None):
    """审定表：账套主体【纵向】排列（主体多时更易读）；格式：核算主体 | 年初数 | 期末未审数 | 审计调整数 | 审定数 | 与TB勾稽。
    ⚡ 2026-08-11 阶段二 2.1：计算产结构化行，渲染收敛到 audit_render.render_sheet。"""
    ds = -1.0 if display_credit_positive else 1.0
    sn = sheet_name if sheet_name else '审定表'
    ncols = 6
    from audit_render import render_sheet, row as _arow
    MONEY = {2, 3, 4, 5}
    rows = []
    # 各主体（纵向），0 值也列示
    for ent in ent_list:
        for y in years:
            qc, jf, df, qm = ctrl_fn(ent, y)
            d_qc = ds * qc; d_qm = ds * qm   # 余额转正数显示
            # 2026-08-06 协议：审定数写数值（=期末未审+调整0），公式 data_only 读 None 收回读不回
            rows.append(_arow([ent, d_qc, d_qm, None, round(d_qm, 2), '与TB核对一致'],
                              num=MONEY, align='l'))
    # 集团合计（⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出）
    for y in years:
        if len(ent_list) <= 1:
            break
        g_qc = sum(ctrl_fn(ent, y)[0] for ent in ent_list)
        g_qm = sum(ctrl_fn(ent, y)[3] for ent in ent_list)
        rows.append(_arow(['集团合计', ds * g_qc, ds * g_qm, None, round(ds * g_qm, 2), '与TB核对一致'],
                          b=True, fill='sub', num=MONEY, align='l'))
    ws = render_sheet(wb, dict(
        name=sn, title=title, title_gap=0,
        headers=['核算主体', '年初数', '期末未审数', '审计调整数', '审定数', '与TB勾稽'],
        rows=rows, money_cols=MONEY, zero_blank=False, freeze='A3',
        widths=[16, 18, 18, 18, 18, 18],
    ))
    return ws


# ---------- 凭证抽查（行级带「账套主体」列） ----------
def build_prior_year_mark_sheet(wb, title, recs):
    """以前年度调整披露表：DTA/DTL 中对方科目属 以前年度损益调整/未分配利润/期初/年初 的凭证，
    只作标记披露、不纳入凭证抽查（2026-08-02 用户方法论：此类调整未过所得税费用科目，
    直接借记 DTA/DTL，属以前年度事项，无需抽凭）。"""
    ws = wb.create_sheet('以前年度调整披露')
    _txt(ws, 1, 1, title, TITLE_FONT, align=LEFT, border=False)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    hdr = ['核算主体', '日期', '凭证字', '凭证号', '科目', '摘要', '借方金额', '贷方金额', '对方科目（以前年度类）']
    for j, h in enumerate(hdr, 1):
        _txt(ws, 2, j, h, HDR_FONT, HDR_FILL, CEN)
    recs.sort(key=lambda r: (str(r.get('e', '') or ''), str(r.get('date', '') or '')))
    r = 3
    sd = sc = 0.0
    for rec in recs:
        _txt(ws, r, 1, rec['e'])
        _txt(ws, r, 2, str(rec['date'])[:10] if rec.get('date') else '')
        _txt(ws, r, 3, rec.get('vtype'))
        _txt(ws, r, 4, rec.get('vno'))
        _txt(ws, r, 5, rec['name'])
        _txt(ws, r, 6, rec.get('summary'))
        _money(ws, r, 7, rec.get('dr'))
        _money(ws, r, 8, rec.get('cr'))
        _txt(ws, r, 9, rec.get('opp'))
        sd += float(rec.get('dr') or 0.0); sc += float(rec.get('cr') or 0.0)
        r += 1
    _txt(ws, r, 1, '合计', TOT_FONT, TOT_FILL)
    _txt(ws, r, 2, f'{len(recs)} 笔', TOT_FONT, TOT_FILL)
    _money(ws, r, 7, sd, bold=True, fill=TOT_FILL)
    _money(ws, r, 8, sc, bold=True, fill=TOT_FILL)
    _txt(ws, r + 2, 1, '说明：以上凭证对方科目为以前年度损益调整/未分配利润/期初余额等，属以前年度事项，'
                       '已从凭证抽查中剔除（仅披露，不抽凭）。', Font(name='Times New Roman', size=10, italic=True, color='808080'), align=LEFT)
    for col, w in (('A', 16), ('B', 12), ('C', 10), ('D', 10), ('E', 24), ('F', 30), ('G', 14), ('H', 14), ('I', 30)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A3'
    return ws


# ---------- 凭证抽查（行级带「账套主体」列） ----------
def build_vouch_sheet(wb, title, vouched, ent_list, topn=2000):
    ws = wb.create_sheet('凭证抽查')
    _txt(ws, 1, 1, title, TITLE_FONT, align=LEFT, border=False)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=15)
    hdr = ['核算主体', '测试序号', '日期', '凭证字', '凭证号', '二级科目（或客户/供应商名称）', '摘要', '借方金额', '贷方金额', '对方科目', '与原始凭证相符', '原始凭证内容', '原始凭证日期', '会计处理正确', '所属时间无误']
    for j, h in enumerate(hdr, 1):
        _txt(ws, 2, j, h, HDR_FONT, HDR_FILL, CEN)
    vouched.sort(key=lambda r: (str(r.get('e', '') or ''), str(r.get('date', '') or '')))
    shown = vouched  # 2026-07-31：全量列示，不再 topn 截断（抽凭内容必须完整）
    r = 3
    sd = sc = 0.0
    for i, rec in enumerate(shown, 1):
        _txt(ws, r, 1, rec['e'])
        _txt(ws, r, 2, i)
        _txt(ws, r, 3, str(rec['date'])[:10] if rec['date'] else '')
        _txt(ws, r, 4, rec['vtype'])
        _txt(ws, r, 5, rec['vno'])
        _txt(ws, r, 6, rec['name'])
        _txt(ws, r, 7, rec['summary'])
        _money(ws, r, 8, rec['dr'])
        _money(ws, r, 9, rec['cr'])
        _txt(ws, r, 10, rec['opp'])
        # 列11-15: 留给人工填写
        _txt(ws, r, 11, None)
        _txt(ws, r, 12, None)
        _txt(ws, r, 13, None)
        _txt(ws, r, 14, None)
        _txt(ws, r, 15, None)
        sd += rec['dr']; sc += rec['cr']
        r += 1
    _txt(ws, r, 1, '合计（列示部分）', TOT_FONT, TOT_FILL)
    _txt(ws, r, 2, f'{len(shown)} 笔', TOT_FONT, TOT_FILL)
    _money(ws, r, 8, sd, bold=True, fill=TOT_FILL)
    _money(ws, r, 9, sc, bold=True, fill=TOT_FILL)
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 22
    ws.column_dimensions['G'].width = 30
    ws.column_dimensions['H'].width = 14
    ws.column_dimensions['I'].width = 14
    ws.column_dimensions['J'].width = 24
    for c in 'KLMNO':
        ws.column_dimensions[c].width = 16
    return ws


# ---------- 说明 ----------
def build_note_sheet(wb, cfg, ent_list, years, extra=None):
    ws = wb.create_sheet('说明')
    notes = [
        f'本底稿为「集团审计底稿」—— {os.path.basename(cfg["data_dir"])} 集团 / {cfg["name"]}（{cfg["code"]}）。',
        '结构：审定表 + 明细表(分年度) + 科目内核对 + 凭证抽查' + (' + 分析sheet' if cfg.get('has_analysis') else '') + '。',
        '账套主体标注：审定表 = 各账套主体【纵向】逐行 + 集团合计行（期末含 审计调整数/期末审定数 列）；明细表/凭证抽查 = 行级「账套主体」列。',
        '合并口径：文件夹内各账套主体简单求和，未做内部往来/交易抵消（符合当前模板未定阶段）。',
        '未来将另起单独程序，按「账套主体」维度拆分本集团底稿（当前标注即可支撑拆分）。',
        f'覆盖账套主体 {len(ent_list)} 个；年份 {years}。',
    ]
    if extra:
        notes += extra
    for i, n in enumerate(notes, 1):
        _txt(ws, i, 1, n, align=LEFT, border=False)
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=8)
    ws.column_dimensions['A'].width = 110
    return ws

# ===== 合并后别名：group_paper / gp_generic 的函数现已并入本模块命名空间 =====
G = sys.modules[__name__]    # group_paper 函数/常量
GP = sys.modules[__name__]   # gp_generic 函数

# ===================== gp_generic（通用生成 / 辅助函数） =====================
# -*- coding: utf-8 -*-
"""集团审计底稿 —— 通用配置驱动生成器（除 应收账款/研发支出 外的所有科目组）
================================================================
承接用户 2026-07-27 指令：每个文件夹=集团，每个科目×年度=一份集团审计底稿；
货币资金/存货不拆科目（只拆 2 个年度）。

本模块复用 group_paper.py 通用框架（审定表/凭证抽查/说明/末级TB/aux读取），
并按"科目规格 spec"统一产出：
  审定表 + 明细表_2025 + 明细表_2026 + 科目内核对 + 凭证抽查 + 说明。

spec 字段：
  key        唯一键
  name       科目名（用于文件名与标题）
  codes      TB 代码前缀列表（如 ['1123'] / ['1001','1002','1012'] / ['14']）
  is_credit  True=贷余科目（负债/权益/收入），False=借余科目（资产/费用）
  split_sub  True=明细表按 2 级科目拆行；False=不拆科目（货币资金/存货），仅按主体列示
  has_aux    True=往来科目，科目内核对用 TB↔辅助核算余额表（逐户）
  aux_name   has_aux 时 read_aux_generic 的 name_contains
  aux_excl   has_aux 时排除的科目名子串
  gl_names   GL 凭证过滤的【科目名称前缀】列表（无独立代码列，只能按名称）
  note       额外说明

取数铁律复用（见 group_paper / audit_common）：
  · tb_leaf_codes 仅累加末级，避免 U8 父+子双计（铁律3/6）。
  · 凭证抽查统一调用 A.voucher_should_skip（铁律：排除结转本年利润 / 相符计提）。
  · 货币资金/存货 split_sub=False：明细表只按账套主体列示（不拆 1001/1002/1012 或 14xx）。
"""

_norm = G._norm




# ============================ 数据读取辅助 ============================
def _tb_family(tb, ent, year, codes):
    """返回该主体/年/科目族的全部 TB 条目 [(norm_code, name, v)]。"""
    out = []
    for (e, c, n, y), v in tb.items():
        nc = _norm(c)
        if e == ent and y == str(year) and any(nc.startswith(p) for p in codes):
            out.append((nc, n, v))
    return out


def _family_total(tb, ent, year, codes):
    qc = jf = df = qm = 0.0
    for nc, n, v in _tb_family(tb, ent, year, codes):
        qc += v['qc']; jf += v['jf']; df += v['df']; qm += v['qm']
    return qc, jf, df, qm


def _family_level2(tb, ent, year, codes, gl_names=None):
    """返回 2 级(代码长6)子目 [(code,name,qc,jf,df,qm)]（已含末级，不重复计父级）。
    2026-08-07 修复（dq 租赁负债串成递延收益）：codes 是多账套并集必然撞车（dq 2601=
    递延收益[政府补助] 但 2601 同时是标准租赁负债码）→ 传入 gl_names 做【名称互斥校验】：
    全部 codes 命中行名称均与 gl_names 互斥（如『递延收益』不含『租赁』）→ 判定该代码
    在本账套是其他科目 → 返回空（账套无此科目）。"""
    fam = _tb_family(tb, ent, year, codes)
    if gl_names and fam:
        gns = [g for g in gl_names if g]
        if gns:
            _good = any(any(g in str(n) or str(n) in g for g in gns) for _, n, _ in fam)
            if not _good:
                return []
    norm = [c for c, _, _ in fam]
    # ⚡⚡ 2026-08-29 P0：gl_names 行级过滤——codes 命中族后，族内名称不匹配的行剔除。
    #   （ncl codes 加 2501 后，2501010000『一年以上长期借款』误入一年内到期明细表；
    #    族级互斥只判"族是否属于该科目"，不筛行。）
    _gns = [g for g in (gl_names or []) if g]
    def _row_hit(n):
        if not _gns:
            return True
        return any(g in str(n) or str(n) in g for g in _gns)
    rows = []
    for c, n, v in fam:
        if len(c) == 6 and _row_hit(n):  # 2 级明细（本项目标准明细粒度）
            rows.append((c, n, v['qc'], v['jf'], v['df'], v['qm']))
    # 若没有 6 位子目（个别科目仅 4 位一级），则退化为展示一级
    if not rows:
        # ⚡ 2026-08-10 SAP：TB 只有 10 位末级码（2401010000 递延收益-搬迁收益），
        # len==6/4 均不满足 → 明细表空壳。铁律73 明细到末级 → 直接展开 10 位末级。
        ten = [(c, n, v) for c, n, v in fam if len(c) == 10 and _row_hit(n)]
        if ten:
            rows = [(c, n, v['qc'], v['jf'], v['df'], v['qm']) for c, n, v in ten]
        else:
            for c, n, v in fam:
                if len(c) == 4 and _row_hit(n):
                    rows.append((c, n, v['qc'], v['jf'], v['df'], v['qm']))
    return rows


# ============================ 凭证抽查（支持多名称前缀） ============================
def _gl_name_hit(name, names):
    """GL 科目名匹配候选名称列表。铁律：科目名以【实际账套科目名】为准——
    ① 精确相等 ② 以'候选名-'开头（二级科目）③ 候选名长度>=3 且包含于实际名称。
    覆盖"销售费用/营业费用"式名称差异场景（调用方把别名一并放入 names 即可）。"""
    n = str(name or '').strip()
    for p in (names or []):
        p = (p or '').strip()
        if not p:
            continue
        if n == p or n.startswith(p + '-'):
            return True
        if len(p) >= 3 and p in n:
            return True
    return False


def _vouched_multi(gl, prefixes, excludes=()):
    """返回应抽凭 GL 行（已排除 计提/补提/结转/摊销 行级 + 相符计提/结转 凭证级 + 同行计提行级）。
    prefixes: 科目名称列表（任一命中即纳入，按实际账套科目名匹配）。

    2026-07-27 修复：除整笔凭证级排除外，追加【行级纯计提剔除】——同一凭证号混合
    真实付款行时，计提单行(opp=应付职工薪酬等计提负债、无真实结算)仍逐行排除，
    真实付款行(opp=银行/应付票据等)保留。详见 audit_common.is_accrual_line。
    """
    from collections import OrderedDict
    rows = [r for r in gl
            if _gl_name_hit(r['name'], prefixes)
            and all(x not in str(r['name']) for x in excludes)]
    # 2026-07-31 移除 ALLOC_KW 摘要行级预过滤（同 _vouched 说明：摘要含"结转/分摊/分配"
    # 的真实业务会被误删，改由凭证级/行级精确判定排除）
    groups = OrderedDict()
    for r in rows:
        key = (r['e'], r['y'], r['vtype'], r['vno'])
        groups.setdefault(key, []).append(r)
    out = []
    for grp in groups.values():
        lines = [{'name': x['name'], 'opp': x['opp'], 'dr': x['dr'], 'cr': x['cr'], 'summary': x['summary']}
                 for x in grp]
        sk, _ = A.voucher_should_skip(summary=grp[0]['summary'], name='', opp='', lines=lines)
        if sk:
            continue  # 整笔结转/纯计提 → 全排除
        # 行级：剔除同号凭证中仍残留的"相符计提"单行（真实付款行不在此列）
        for x in grp:
            if A.is_accrual_line({'name': x['name'], 'opp': x['opp'], 'dr': x['dr'], 'cr': x['cr'], 'summary': x['summary']}):
                continue
            out.append(x)
    return out


# ===================== gp_other（其他类科目） =====================
# -*- coding: utf-8 -*-
"""其他类科目集团底稿小程序
================================================================
覆盖非 往来/税费/收入/损益/货币资金/存货/长期资产/权益/借款 的余额类科目：
  1101 交易性金融资产
  1511 长期股权投资
  1512 长期股权投资减值准备（长投减值，贷余备抵）
  1711 商誉
  2201 应付票据
  2703 租赁负债
  2711 专项应付款
  2801 预计负债
  2901 递延所得税负债
每个文件夹=集团，按【科目】自动识别：仅对 TB 有数据的科目生成底稿（无数据跳过）。
明细表分 2 级明细（若有）：期初、本期增加、本期减少、期末未审、审计调整数、期末审定数（调整数留空默认0，审定数=未审数+调整数）。
结构与其他集团底稿一致：审定表 + 明细表(分年度) + 科目内核对 + 凭证抽查 + 说明。
"""

_norm = G._norm

# 辅助核算末级实名行识别：这些 科目辅助核算名称 是父级/中间汇总/维度标签，非真实末级
_GENERIC_AUX_NAMES = {"", "客户", "供应商", "科目", "个人", "单位往来", "个人往来",
                      "现金流量", "现金流量(无余额显示)", "现金流量（无余额显示）",
                      "研发项目", "项目", "银行", "银行承兑", "商业承兑"}


def _read_aux_leaf(data_dir, entity, year, names, is_credit=False):
    """读 辅助核算余额表 的【二级子目】行（银行承兑/商业承兑级），用于 应付票据 等票据科目。
    返回 [(code, disp_name, qc, jf, df, qm)]（与 GP._family_level2 同结构），金额按账户性质
    转自然符号（负债 is_credit=True → 贷余为负；aux 期初方向列常为空，缺方向时按 is_credit 兜底）。
    注：aux 的『收票单位/逐笔票据』级(level-4)存在大量重复行，合计与 TB 严重不勾稽，故只取
    干净的二级子目（与 TB 完全吻合）。无 aux 文件或无可识别二级时返回 []。"""
    fns = discover_aux(data_dir, entity, year)
    if not fns:
        return []
    name_set = set(names)
    leaf = []
    for fp in fns:
        try:
            wba = openpyxl.load_workbook(fp, data_only=True, read_only=True)
            ws = wba[wba.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
            wba.close()
        except Exception:
            continue
        hi = next((i for i, r in enumerate(rows[:8])
                   if r and any(isinstance(c, str) and '科目辅助核算代码' in c for c in r)), None)
        if hi is None:
            continue
        hdr = rows[hi]

        def col(n):
            try:
                return hdr.index(n)
            except ValueError:
                return None
        c_code = col('科目辅助核算代码'); c_name = col('科目名称'); c_aux = col('科目辅助核算名称')
        c_jf = col('本期借方'); c_df = col('本期贷方')
        amt = [i for i, c in enumerate(hdr) if c == '金额']
        c_qc = amt[0] if amt else None
        c_qm = amt[1] if len(amt) > 1 else (amt[0] if amt else None)
        dir_idxs = [i for i, c in enumerate(hdr) if c == '方向']
        c_dir0 = dir_idxs[0] if dir_idxs else None   # 期初方向
        c_dir1 = dir_idxs[1] if len(dir_idxs) > 1 else c_dir0  # 期末方向
        if None in (c_code, c_name, c_aux, c_jf, c_df, c_qc, c_qm):
            continue

        def num(x):
            try:
                return float(x)
            except (TypeError, ValueError):
                return 0.0
        # 余额符号：优先用方向列；方向缺失(aux 期初方向常空)时按账户性质(is_credit)兜底
        def sgn(v, d):
            if d == '贷':
                return -abs(v)
            if d == '借':
                return abs(v)
            return (-abs(v)) if is_credit else abs(v)

        for r in rows[hi + 1:]:
            if not r:
                continue
            kname = str(r[c_name]) if r[c_name] is not None else ''
            # 仅本科目及其二级子目(name-xxx)
            if kname not in name_set and not any(kname.startswith(n + '-') for n in name_set):
                continue
            aux = str(r[c_aux]) if r[c_aux] is not None else ''
            # 只保留汇总标签行(空/科目 等)中的二级子目，跳过收票单位级脏数据
            if aux not in _GENERIC_AUX_NAMES:
                continue
            # 仅取『name-xxx』二级子目，跳过一级父科目(name 本身)
            if not any(kname.startswith(n + '-') for n in name_set):
                continue
            code = str(r[c_code]) if r[c_code] is not None else aux
            disp = kname
            for n in name_set:
                if kname.startswith(n + '-'):
                    disp = kname[len(n) + 1:]   # 去『应付票据-』前缀 → 银行承兑/商业承兑
                    break
            qc = sgn(num(r[c_qc]) if c_qc is not None else 0.0, r[c_dir0] if c_dir0 is not None else None)
            jf = num(r[c_jf]) if c_jf is not None else 0.0
            df = num(r[c_df]) if c_df is not None else 0.0
            qm = sgn(num(r[c_qm]) if c_qm is not None else 0.0, r[c_dir1] if c_dir1 is not None else None)
            if abs(qc) < 1e-6 and abs(jf) < 1e-6 and abs(df) < 1e-6 and abs(qm) < 1e-6:
                continue  # 0 值行不显(铁律)
            leaf.append((code, disp, qc, jf, df, qm))
    return leaf


# is_credit: True=贷余(负债/备抵)，False=借余(资产)
SUBJECTS = [
    dict(key='tfa',  name='交易性金融资产',     codes=['1101', '1003'], is_credit=False, gl_names=['交易性金融资产']),
    dict(key='lti',  name='长期股权投资',       codes=['1511'], is_credit=False, gl_names=['长期股权投资']),
    dict(key='lti_imp', name='长期股权投资减值准备', codes=['1512'], is_credit=True, gl_names=['长期股权投资减值准备']),
    dict(key='gw',   name='商誉',               codes=['1711'], is_credit=False, gl_names=['商誉']),
    dict(key='apn',  name='应付票据',           codes=['2201'], is_credit=True,  gl_names=['应付票据']),
    dict(key='ll',   name='租赁负债',           codes=['2601', '2631', '2632', '2703', '2271', '2802', '2206', '220601', '220602', '2260'], is_credit=True,  gl_names=['租赁负债']),
    # ⚡⚡ 2026-08-24 恢复 2401：AQ 等账套 2401=一年内到期的非流动负债（不能删）。
    #   同代码多语义（AS 2401=递延收益）由 _family_level2 的 gl_names 名称互斥裁决：
    #   名称含『一年内到期』才归 ncl（AS 递延收益自动排除）。删代码是治标且破坏 AQ。
    dict(key='ncl',  name='一年内到期的非流动负债', codes=['2242', '2485', '2401', '2281', '2703', '2501'], is_credit=True,
         gl_names=['一年内到期的长期借款', '一年内到期的应付债券', '一年内到期的租赁负债', '一年内到期的长期应付款']),
    dict(key='di',   name='递延收益',           codes=['2401', '2601'], is_credit=True,  gl_names=['递延收益']),
    dict(key='sp',   name='专项应付款',         codes=['2711'], is_credit=True,  gl_names=['专项应付款']),
    dict(key='lpay', name='长期应付款',         codes=['2701'], is_credit=True,  gl_names=['长期应付款']),
    dict(key='bpay', name='应付债券',           codes=['2502'], is_credit=True,  gl_names=['应付债券']),
    dict(key='ep',   name='预计负债',           codes=['2801', '2301'], is_credit=True,  gl_names=['预计负债']),
    dict(key='dtl',  name='递延所得税负债',     codes=['2901'], is_credit=True,  gl_names=['递延所得税负债']),
    dict(key='dta',  name='递延所得税资产',     codes=['1811', '1901'], is_credit=False, gl_names=['递延所得税资产']),
    # ⚡ 2026-08-16 补：ga/jj 未覆盖科目（铁律：所有有数据科目均生成底稿）
    # ⚡⚡ 2026-08-24 修复：AFJ 用 1513（非 1507）→ 审定表空但附注有数（FG 48万）
    dict(key='oei',  name='其他权益工具投资',   codes=['1507', '1513', '1503'], is_credit=False, gl_names=['其他权益工具投资']),
    dict(key='apfu', name='上级拨入资金',       codes=['4200'], is_credit=True,  gl_names=['上级拨入资金']),
]

def _build_flow_recon(wb, name, codes, ent_list, years, tb, is_credit=False):
    """借贷方发生额核对表：期初余额 + 本期增加发生额 − 本期减少发生额 = 期末余额，逐主体勾稽。
    2026-07-31 用户要求恢复（递延所得税资产底稿缺少「原来的借贷方发生额核对表」）。
    - 列示【实际】本期借方/贷方发生额（TB 原始发生额）；
    - 增加/减少按科目性质映射：资产类 增加=借方发生额、减少=贷方发生额；负债类 增加=贷方发生额、减少=借方发生额；
    - 余额按科目性质转正显示（资产借余正、负债贷余正），勾稽=期初+增−减−期末，恒等 0。
    - 全主体/年度均无金额时不生成（与 _build_dt_recon 一致）。"""
    ds = -1.0 if is_credit else 1.0
    rows = []
    for ent in ent_list:
        for y in years:
            qc, jf, df, qm = tb_leaf_codes(tb, ent, y, codes)
            if abs(qc) < 0.005 and abs(jf) < 0.005 and abs(df) < 0.005 and abs(qm) < 0.005:
                continue
            d_qc = ds * qc; d_qm = ds * qm
            inc, dec = (df, jf) if is_credit else (jf, df)   # 增/减按科目性质
            chk = round(d_qc + inc - dec - d_qm, 2)
            rows.append((ent, y, d_qc, jf, df, d_qm, chk))
    if not rows:
        return
    ws = wb.create_sheet(f'{name}借贷方发生额核对')
    _txt(ws, 1, 1, f'{name}（{"/".join(codes)}）借贷方发生额核对 — 期初+本期增发生额−本期减发生额=期末', TITLE_FONT, align=LEFT, border=False)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    hdr = ['核算主体', '年份', '期初余额', '本期借方发生额', '本期贷方发生额', '期末余额', '勾稽差异(期初+增−减−末)']
    for j, h in enumerate(hdr, 1):
        _txt(ws, 2, j, h, HDR_FONT, HDR_FILL, CEN)
    r = 3
    for ent, y, d_qc, jf, df, d_qm, chk in rows:
        _txt(ws, r, 1, ent); _txt(ws, r, 2, y)
        _money(ws, r, 3, d_qc); _money(ws, r, 4, jf); _money(ws, r, 5, df); _money(ws, r, 6, d_qm)
        _money(ws, r, 7, chk)
        if abs(chk) > 0.005:
            for c in range(1, 8):
                ws.cell(r, c).fill = PatternFill('solid', fgColor='FCE4D6')
        r += 1
    for col, w in [('A', 16), ('B', 8), ('C', 16), ('D', 18), ('E', 18), ('F', 16), ('G', 22)]:
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A3'


def _build_dt_recon(wb, name, codes, ent_list, years, tb):
    """递延所得税资产/负债与对应科目核对表：ΔDTA − ΔDTL = 递延所得税费用。
    2026-07-31：若全部主体/年度均无金额数据，不生成该 sheet（用户要求"勾稽核对无金额不产生"）。"""
    rows = []
    for ent in ent_list:
        for y in years:
            _, jf_dta, df_dta, _ = G.tb_leaf_codes(tb, ent, y, codes)
            dta_chg = jf_dta - df_dta
            _, jf_dtl, df_dtl, _ = G.tb_leaf_codes(tb, ent, y, ['2901'])
            dtl_chg = df_dtl - jf_dtl
            _, jf_tax, df_tax, _ = G.tb_leaf_codes(tb, ent, y, ['6801'])
            # ⚡ 2026-08-28 修复：6801 含 当期(680101)+递延(680102)，核对列『所得税费用-递延』
            #   必须只取递延部分。原取 6801 全部 → AH 1010 把当期 25,046,483 当递延、勾稽假差
            #   （用户：所得税费用中无递延相关费用）。递延 = 6801 − 当期。
            _, jf_cur, df_cur, _ = G.tb_leaf_codes(tb, ent, y, ['680101'])
            jf_tax, df_tax = jf_tax - jf_cur, df_tax - df_cur
            # 勾稽恒等式：ΔDTA − ΔDTL = 递延所得税费用 → 差异 = 实际所得税费用 − 恒等式右端
            # 2026-08-01 修复：原 expected≠0 时强制 diff=0（掩盖差异），改显示真实差异供审计判断
            expected = -(dta_chg - dtl_chg)
            diff = round((jf_tax - df_tax) - expected, 2)
            # 按【净额】判断：借贷相抵（如所得税费用红字 jf=df）的行显示金额全 0，无核对意义，
            # 一并跳过 → 全部行均 0 时不生成核对 sheet（2026-07-31 用户要求）
            if abs(dta_chg) < 0.005 and abs(dtl_chg) < 0.005 and abs(jf_tax - df_tax) < 0.005:
                continue
            rows.append((ent, y, dta_chg, dtl_chg, jf_tax - df_tax, round(diff, 2)))
    if not rows:
        return  # 无金额数据 → 不生成核对 sheet
    ws = wb.create_sheet(f'{name}与对应科目核对')
    _txt(ws, 1, 1, f'{name} ↔ 所得税费用/递延所得税负债 勾稽核对', TITLE_FONT, align=LEFT, border=False)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    hdr = ['核算主体', '年份', f'{name}本期增减', '递延所得税负债本期增减', '所得税费用-递延', '勾稽差异']
    for j, h in enumerate(hdr, 1):
        _txt(ws, 2, j, h, HDR_FONT, HDR_FILL, CEN)
    r = 3
    for ent, y, dta_chg, dtl_chg, tax_net, diff in rows:
        _txt(ws, r, 1, ent); _txt(ws, r, 2, y)
        _money(ws, r, 3, dta_chg); _money(ws, r, 4, dtl_chg)
        _money(ws, r, 5, tax_net); _money(ws, r, 6, diff)
        r += 1
    for col, w in [('A', 16), ('B', 8), ('C', 18), ('D', 18), ('E', 18), ('F', 14)]:
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A3'


def _build_detail_year(wb, y, name, codes, is_credit, is_bill, ent_list, tb, ctrl, ds, data_dir, gl_names=None):
    """构建某科目某年度明细表（2026-08-01 按年构建提取）；返回当年勾稽是否异常。"""
    anomaly = False
    ws = wb.create_sheet(f'{name}明细表_{y}')
    lvl_label = '二级(银行承兑/商业承兑)' if is_bill else '2级'
    G._txt(ws, 1, 1, f'{name}（{"/".join(codes)}）明细表（{lvl_label}）— {os.path.basename(data_dir)} 集团 {y}', G.TITLE_FONT, align=G.LEFT, border=False)
    hdr = ['账套主体', '科目代码', '科目名称(末级)', '期初余额', '本期增加', '本期减少', '期末未审', '审计调整数', '期末审定数', '与TB勾稽']
    ncol = len(hdr)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    for j_, h in enumerate(hdr, 1):
        G._txt(ws, 2, j_, h, G.HDR_FONT, G.HDR_FILL, G.CEN)
    r = 3
    tot = [0.0, 0.0, 0.0, 0.0]
    sub_rows = []   # 主体小计行号（年度合计公式引用，2026-08-03 公式化）
    for ent in ent_list:
        etot = [0.0, 0.0, 0.0, 0.0]
        ent_rows = []   # 本主体数据行号（主体小计公式引用）
        if is_bill:
            leaf = _read_aux_leaf(data_dir, ent, y, [name], is_credit=is_credit)
            rows = leaf if leaf else GP._family_level2(tb, ent, y, codes, gl_names)
        else:
            rows = GP._family_level2(tb, ent, y, codes, gl_names)
        if not rows:
            continue
        for (c, nm, qc, jf, df, qm) in rows:
            # ⚡ 2026-08-10 0 值行隐藏（铁律4）：SAP 10 位末级含全 0 行
            # （1811000000 递延所得税资产父级位 / 2711020100 无发生额子目）→ 跳过
            if all(abs(x) < 0.005 for x in (qc, jf, df, qm)):
                continue
            inc = df if is_credit else jf
            dec = jf if is_credit else df
            d_qc = ds * qc; d_qm = ds * qm   # 负债余额转正数显示
            G._txt(ws, r, 1, ent); G._txt(ws, r, 2, c); G._txt(ws, r, 3, nm)
            G._money(ws, r, 4, d_qc); G._money(ws, r, 5, inc); G._money(ws, r, 6, dec); G._money(ws, r, 7, d_qm)
            G._txt(ws, r, 8, None)  # 审计调整数（留空待填）
            _c = ws.cell(r, 9, f'=G{r}+H{r}'); _c.number_format = '#,##0.00'; _c.alignment = G.RIGHT; _c.border = G.BORDER
            G._txt(ws, r, 10, None)
            for i_, v in enumerate((d_qc, inc, dec, d_qm)):
                etot[i_] += v
            ent_rows.append(r)
            r += 1
        # 小计条件：期初/增/减/期末任一非零即插小计（2026-08-01 修复 #8：
        # 数链递延所得税资产 期初=0/期末=0 仅发生额非零，原条件漏插小计行）
        if any(abs(v) > 0.005 for v in etot):
            G._txt(ws, r, 1, ent, G.TOT_FONT, G.SUB_FILL); G._txt(ws, r, 2, '', G.TOT_FONT, G.SUB_FILL); G._txt(ws, r, 3, '小计', G.TOT_FONT, G.SUB_FILL)
            # 2026-08-03 公式化：主体小计 = SUM(本主体数据行,逗号分隔)
            if ent_rows:
                for i_ in range(4):
                    col = chr(68 + i_)   # D/E/F/G
                    _m = ws.cell(r, 4 + i_, '=SUM(%s)' % ','.join(f'{col}{x}' for x in ent_rows))
                    _m.number_format = '#,##0.00'; _m.alignment = G.RIGHT; _m.border = G.BORDER; _m.fill = G.SUB_FILL
            else:
                for i_, v in enumerate(etot):
                    G._money(ws, r, 4 + i_, v, fill=G.SUB_FILL)
            G._txt(ws, r, 8, None, fill=G.SUB_FILL)
            _c = ws.cell(r, 9, f'=G{r}+H{r}'); _c.number_format = '#,##0.00'; _c.alignment = G.RIGHT; _c.border = G.BORDER; _c.fill = G.SUB_FILL
            G._money(ws, r, 10, round(etot[3] - ds * ctrl[(ent, y)][3], 2), fill=G.SUB_FILL)
            for i_, v in enumerate(etot):
                tot[i_] += v
            sub_rows.append(r)
            r += 1
    # 年度合计（2026-08-03 公式化 =SUM 主体小计行）
    G._txt(ws, r, 1, f'{y} 年度合计', G.TOT_FONT, G.TOT_FILL); G._txt(ws, r, 2, '', G.TOT_FONT, G.TOT_FILL); G._txt(ws, r, 3, '', G.TOT_FONT, G.TOT_FILL)
    if sub_rows:
        for i_ in range(4):
            col = chr(68 + i_)
            _m = ws.cell(r, 4 + i_, '=SUM(%s)' % ','.join(f'{col}{x}' for x in sub_rows))
            _m.number_format = '#,##0.00'; _m.alignment = G.RIGHT; _m.border = G.BORDER; _m.font = G.TOT_FONT; _m.fill = G.TOT_FILL
    else:
        for i_, v in enumerate(tot):
            G._money(ws, r, 4 + i_, v, bold=True, fill=G.TOT_FILL)
    G._txt(ws, r, 8, None, fill=G.TOT_FILL)
    _c = ws.cell(r, 9, f'=G{r}+H{r}'); _c.number_format = '#,##0.00'; _c.alignment = G.RIGHT; _c.border = G.BORDER; _c.font = G.TOT_FONT; _c.fill = G.TOT_FILL
    tb_close = sum(ctrl[(e2, y)][3] for e2 in ent_list)
    diff_close = round(tot[3] - ds * tb_close, 2)   # tot[3] 已为显示口径(负债转正)，TB 同步翻转比较
    if abs(diff_close) > 0.005:
        anomaly = True
    G._money(ws, r, 10, diff_close, bold=True, fill=G.TOT_FILL)
    r += 2
    G._txt(ws, r, 1, f'勾稽：明细表合计期末审定数 {tot[3]:,.2f}  vs  TB（{"/".join(codes)}）集团期末 {ds*tb_close:,.2f}  →  差异 {diff_close:,.2f}',
            G.TOT_FONT if abs(diff_close) > 0.005 else None)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
    for col, w in [('A', 16), ('B', 14), ('C', 34), ('D', 15), ('E', 15), ('F', 15), ('G', 15), ('H', 13), ('I', 15), ('J', 14)]:
        ws.column_dimensions[col].width = w
        return anomaly
    return anomaly

def build_other(data_dir, spec, tb, gl, ent_list, years, entities=None):
    """直接按年构建（2026-08-01 改造：消除合并稿中间态，替代 emit_per_year 分拆+按年重建）。
    每年独立 wb：明细表→审定表(audit_common 按年)→递延勾稽(按年)→凭证抽查(按年)。"""
    name = spec['name']
    # 2026-08-07 v2（会计语言优先，代码方言隔离）：codes 经账套配置解析——
    # account_profiles.json subject_codes 有账套特例码用之，否则用 spec 默认码。
    # dq 递延收益=2601、Q 租赁负债=2271、H 预计负债=2301 等特例不再进全局注册表，
    # 消除跨账套 codes 撞车（2601 在 dq=递延收益、FY=租赁负债 的历史 bug）。
    # ⚡ 2026-08-10 账套名取真实数据根（SAP 下 data_dir=_work_{comp} 独立工作目录，
    # basename 非账套名 → account_profiles 解析落空）
    _acct = os.path.basename(
        (_adapter._DATA_ROOT if _adapter is not None and getattr(_adapter, '_DATA_ROOT', None) else data_dir).rstrip('\\/'))
    try:
        import subject_mapping
        codes = subject_mapping.resolve_subject_codes(_acct, spec['key'], spec['codes'])
    except Exception:
        codes = spec['codes']
    is_credit = spec['is_credit']
    ds = -1.0 if is_credit else 1.0   # 负债(贷余)余额显示为正数

    def ctrl_fn(ent, year):
        # 2026-08-06 修复：codes 匹配不到时按 gl_names【名称】动态发现实际代码
        # （铁律5：Q 一年内到期=2401/租赁负债=227102，与标准码 2242/2601 不同 →
        # 原 codes 取数全 0 → 底稿静默缺失）。
        _fam = [c for (e, c, n, y) in tb if e == ent and str(y) == str(year)
                and any(g and g in str(n) for g in spec.get('gl_names') or [])]
        _use = _fam or codes
        # 2026-08-07 修复（dq 租赁负债串成递延收益）：codes 是多账套代码并集必然撞车
        # （dq 2601=递延收益[政府补助] 但 2601 同时是标准租赁负债码；di 与 ll 的 codes
        # 都含 2601 → ll 误取递延收益明细）。回退 codes 时【名称校验】：命中代码的行名
        # 若与 gl_names 关键词【互斥】（如『递延收益』不含『租赁』且『租赁』不含『递延』），
        # 判定该代码在本账套是其他科目 → 剔除。仅当存在 名称∩代码 非空的校验行时保留 codes。
        _g_names = [g for g in (spec.get('gl_names') or []) if g]
        if not _fam and _g_names:
            _cand = [c for (e, c, n, y) in tb if e == ent and str(y) == str(year)
                     and any(str(c).startswith(k) for k in codes)]
            if _cand:
                _good = [c for (e, c, n, y) in tb if e == ent and str(y) == str(year)
                         and any(str(c).startswith(k) for k in codes)
                         and any(g in str(n) or str(n) in g for g in _g_names)]
                _bad_only = [c for (e, c, n, y) in tb if e == ent and str(y) == str(year)
                             and any(str(c).startswith(k) for k in codes)
                             and not any(g in str(n) or str(n) in g for g in _g_names)]
                # 全部候选与名称互斥 → 该 codes 在本账套是其他科目（撞车），剔除
                if _bad_only and not _good:
                    _use = []
        return G.tb_leaf_codes(tb, ent, year, _use)

    ctrl = {(ent, y): ctrl_fn(ent, y) for ent in ent_list for y in years}
    # 全集团全年度均 0 → 跳过（该科目在此文件夹无数据）
    # 2026-08-06 修复：all_zero 判定原用 abs(sum(4元组))——(qc,jf,df,qm) 借贷平衡时
    # qc+jf+df+qm 恒等于 0（如 S 母公司 2703 一年内到期 qc=-2.04M+jf=2.04M+df=2.14M+qm=-2.14M=0）
    # → 有数据科目被误判"无数据"跳过 → 底稿缺失。改为【各分量绝对值之和】判定。
    all_zero = all(
        sum(abs(x) for x in ctrl[(ent, y)]) < 1e-6
        for ent in ent_list for y in years)
    if all_zero:
        return False

    vouched = GP._vouched_multi(gl, spec['gl_names'])
    is_bill = (codes == ['2201'])  # 应付票据：辅助核算含逐笔票据(收票单位/银行)，到末级

    # 递延所得税资产/负债：预构建全 GL 凭证级对方科目映射（各年共用）
    _gl_by_vouch = None
    if spec['key'] in ('dtl', 'dta'):
        from collections import OrderedDict
        _gl_by_vouch = OrderedDict()
        for r in gl:
            k = (r['e'], r['y'], r['vtype'], r['vno'])
            _gl_by_vouch.setdefault(k, []).append(r)

    from audit_shell import finalize_workbook as _fw
    from audit_common import add_audit_summary_sheets as _audit
    outs = []
    for y in years:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # 1) 审定表（audit_common 按年生成，与 emit 后 rebuild 格式一致；失败兜底 build_summary_sheet）
        try:
            # ⚡⚡ 2026-08-24 名称主键（铁律58）：审定表 codes 用【名称筛选后】的代码——
            #   原始 codes 是跨账套并集（ncl 含 2401），按账套实际名称过滤（AS 2401=递延收益
            #   剔除、AQ 2401=一年内到期保留）→ 防同码异义混入审定表。
            _gc = [g for g in (spec.get('gl_names') or []) if g]
            _used_codes = sorted({_norm(c) for (e, c, n, yy) in tb
                                  if str(yy) == str(y)
                                  and any(str(c).startswith(k) for k in codes)
                                  and (not _gc or any(g and (g in str(n) or str(n) in g) for g in _gc))})
            _audit_codes = _used_codes or codes
            _audit(wb, data_dir,
                   [dict(title=f'{name}审定表', codes=_audit_codes,
                         is_credit=is_credit, subj_name=name, by_entity=True)],
                   tb_full=tb, entities=entities, target_year=y)
        except Exception as ex:
            print(f'  ⚠️ 审定表生成失败 {name} {y}：{ex}')
            G.build_summary_sheet(wb, f'{name}（{"/".join(codes)}）审定表 — {os.path.basename(data_dir)} 集团 {y}',
                                  ctrl_fn, ent_list, [y], is_credit=is_credit, display_credit_positive=is_credit,
                                  sheet_name=f'{name}审定表')

        # 2) 明细表（仅当年）；返回当年勾稽是否异常
        anomaly = _build_detail_year(wb, y, name, codes, is_credit, is_bill, ent_list, tb, ctrl, ds, data_dir,
                                     gl_names=spec.get('gl_names'))

        # 3) 递延所得税资产/负债：与对应科目核对（按年）。借贷方发生额核对表 2026-08-01 用户要求删除
        if spec['key'] in ('dtl', 'dta'):
            _build_dt_recon(wb, name, codes, ent_list, [y], tb)

        # 4) 凭证抽查（仅当年）：递延所得税资产/负债全量抽凭；其余科目仅当存在勾稽异常时生成
        if spec['key'] in ('dtl', 'dta'):
            # 2026-08-02 用户方法论：DTA/DTL 凭证若对方科目属「以前年度损益调整/期初未分配利润/
            # 年初/期初」等（即以前年度调整，直接借记 DTA/DTL、未过所得税费用），只作标记披露、
            # 不纳入凭证抽查——正常情况下 DTA/DTL 不生成抽凭或仅抽当年递延确认凭证。
            vouched_raw = [r for r in gl
                           if _gl_name_hit(r['name'], spec['gl_names']) and str(r.get('y')) == str(y)]
            _prior_kws = ('以前年度损益调整', '以前年度', '未分配利润', '期初', '年初')
            prior_mark = []
            for rec in vouched_raw:
                if rec.get('opp'):
                    continue
                k = (rec['e'], rec['y'], rec['vtype'], rec['vno'])
                others = [x for x in _gl_by_vouch.get(k, [])
                          if x.get('name') != rec.get('name') and x.get('name')]
                if others:
                    opp_names = sorted(set(x['name'] for x in others))
                    rec['opp'] = '；'.join(opp_names[:5])
                    if len(opp_names) > 5:
                        rec['opp'] += f'…(共{len(opp_names)}个)'
            # 拆分：以前年度调整（对方科目命中）→ 标记披露；其余 → 抽凭
            keep_raw = []
            for rec in vouched_raw:
                _opp = str(rec.get('opp') or '')
                if any(k in _opp for k in _prior_kws):
                    prior_mark.append(rec)
                else:
                    keep_raw.append(rec)
            if prior_mark:
                G.build_prior_year_mark_sheet(wb, f'{name}以前年度调整披露（对方科目=以前年度损益/未分配利润等，'
                                                 f'已从凭证抽查中剔除） — {os.path.basename(data_dir)} 集团',
                                              prior_mark)
            if keep_raw:
                G.build_vouch_sheet(wb, f'{name}（{"/".join(codes)}）凭证抽查 — {os.path.basename(data_dir)} 集团', keep_raw, ent_list)
            else:
                print(f'   [SKIP] {name} {y} 凭证抽查（全部为以前年度调整/无凭证，跳过）')
        elif spec['key'] == 'lti' or anomaly:
            vouched_y = [r for r in vouched if str(r.get('y')) == str(y)]
            G.build_vouch_sheet(wb, f'{name}（{"/".join(codes)}）凭证抽查 — {os.path.basename(data_dir)} 集团', vouched_y, ent_list)
        else:
            print(f"   [SKIP] {name} {y} 凭证抽查（无勾稽异常，跳过）")

        # 5) 附注汇总（滚动四段/两期数；每主体一列，2026-08-02 用户方法论）
        #    滚动四段：长期股权投资/应付票据/租赁负债/专项应付款/预计负债（余额变动类）；
        #    两期数：交易性金融资产/减值准备/商誉/递延所得税资产/负债（期末/期初即可）
        try:
            from audit_common import build_footnote_generic as _bfn
            _mode = 'four' if spec['key'] in ('lti', 'apn', 'll', 'sp', 'ep') else 'two'
            _seg_desc = '段=期初数/本期增加/本期减少/期末数' if _mode == 'four' else '段=期末数/期初数'
            _rows = [(name, codes)]
            if spec['key'] == 'lti':
                # 长期股权投资：按被投资单位（TB 二级，如 长期股权投资-参股子公司/全资子公司/控股子公司）
                # 展开（上市公司附注按被投资单位列示，2026-08-02 用户方法论）；无二级时回退一级单行。
                # 2026-08-06 修复：必须剔除『父级行』——FY 1511 有 1511.01 投资成本(父)/1511.02 损益调整(父)
                # 及 1511.01.xx/1511.02.xx 末级子目行；父级金额=全部子目之和，父级+子级双收集
                # → 附注合计 804M（01FY本级）vs 审定 402M，双计 2 倍（长投附注 848M vs 审定 423M）。
                _l2_codes = []
                for (ee, cc, nn, yy), v in tb.items():
                    if yy != y or ee not in entities:
                        continue
                    cs = str(cc)
                    if cs.startswith('1511') and cs != '1511' and not cs.startswith('1512'):
                        _l2_codes.append(cs)
                # 父级=存在更长子级行的代码（1511.01 有 1511.01.01 等 → 父级，剔除）
                _parents = {c for c in _l2_codes
                            if any(c2 != c and c2.startswith(c) for c2 in _l2_codes)}
                _l2_map = {}
                for (ee, cc, nn, yy), v in tb.items():
                    if yy != y or ee not in entities:
                        continue
                    cs, ns = str(cc), str(nn)
                    if cs.startswith('1511') and cs != '1511' and not cs.startswith('1512') \
                            and '减值' not in ns and cs not in _parents:
                        _l2_map.setdefault(ns, [0.0] * 4)
                        _l2_map[ns][0] += abs(v['qc'])
                        _l2_map[ns][1] += v['jf']
                        _l2_map[ns][2] += v['df']
                        _l2_map[ns][3] += abs(v['qm'])
                _l2_named = [(n, a) for n, a in _l2_map.items()
                             if any(abs(x) > 0.005 for x in a)]
                if _l2_named:
                    _l2_named.sort(key=lambda x: -max(abs(x[1][3]), abs(x[1][0])))
                    # 二级行已覆盖全部金额：显示名去「长期股权投资-」前缀（贴近上市公司按被投资单位披露），
                    # 不再附加一级合计行（防与二级行重复求和导致集团合计数双计）
                    _rows = [(n.replace('长期股权投资-', ''), [], [n]) for n, _ in _l2_named]
            _bfn(wb, data_dir, '附注汇总',
                 f'{name}附注汇总（审定口径；每主体一列；{_seg_desc}）',
                 _rows, is_credit=is_credit, mode=_mode,
                 target_year=y, tb_full=tb, entities=entities,
                 parent_abs_keys=('租赁负债',))
        except Exception as ex:
            print(f'  ⚠️ 附注汇总生成失败 {name} {y}：{ex}')

        _fw(wb)
        # ⚡ 2026-08-28 #875：对方科目核对（模块内集成，内存 wb 注入）
        try:
            from counterparty_recon import inject_into_wb, SPECS as _CR_SPECS
            _ck = spec.get('name')
            if _ck in _CR_SPECS:
                inject_into_wb(wb, gl, y, _CR_SPECS[_ck], ents=set(ent_list))
        except Exception as _ex:
            print(f'  ⚠️ {name}对方科目核对注入失败：{_ex}')
        out = os.path.join(data_dir, f'{name}审计底稿_{y}_生成.xlsx')
        from audit_common import validate_workbook
        validate_workbook(wb, '其他类底稿', raise_on_error=False)
        wb.save(out)
        outs.append(out)
        print(f'  ✓ {name} {y} 已生成')
    return outs[0] if outs else None


def build_all(data_dir):
    print(f'\n========== 其他类集团底稿：{data_dir} ==========')
    # SAP：单主体驱动时按 current_comp 过滤（防全量 GL OOM + 审定表混 88 家主体）
    if _adapter is not None and _adapter.is_sap(data_dir):
        _c_go = _adapter.current_comp()
        entities = ({_c_go: A.discover_entities(data_dir).get(_c_go, {})} if _c_go else A.discover_entities(data_dir))
    else:
        entities = A.discover_entities(data_dir)
    ent_list = sorted(entities.keys())
    tb = A.read_tb_full(data_dir, entities)
    gl = A.read_gl_rows(data_dir, entities)
    years = sorted({y for e in entities.values() for y in e.keys()})
    print(f'账套主体 {len(ent_list)} 个；年份 {years}')
    made, skipped = [], []
    for spec in SUBJECTS:
        try:
            out = build_other(data_dir, spec, tb, gl, ent_list, years, entities=entities)
            if out:
                made.append(os.path.basename(out))
                print(f'  ✓ {spec["name"]}')
            else:
                skipped.append(spec['name'])
                print(f'  – {spec["name"]}（无数据，跳过）')
        except Exception as e:
            print(f'  ✗ {spec["name"]} 生成失败：{e}')
            import traceback; traceback.print_exc()
    print(f'其他类生成完毕：{len(made)} 份；跳过 {len(skipped)} 份无数据。')
    return made, skipped



# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None

if __name__ == '__main__':
    d = sys.argv[1] if len(sys.argv) > 1 else P.G
    build_all(d)
    from audit_common import finalize_after_build
    finalize_after_build(d)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）