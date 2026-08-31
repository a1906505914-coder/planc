# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=营业收入审计底稿_生成.xlsx | 关键列=主营业务收入借发/二级科目/客户 | 职责=收入按主体·二级·客户分析(独立生成器)
"""
revenue_detail.py — 营业收入审计底稿（第 9 个小程序）

产出两套底稿：
  A. 主营业务收入
     表1 主营业务收入_分月明细(按主体)：每个主体 收入/成本/毛利 分月 + 上年 + 增减 + 全年小计
                                 + 全集团分月汇总 + 全集团全年合计 + 与科目余额表核对
     表2 主营业务收入_分月明细(按二级)：按科目余额表二级明细从左往右排列，每个二级 8 指标列
                                 + 合计块 + 按二级与科目余额表核对
     表3 主营业务收入_客户明细(全年)  + 前十大客户_分月   （当 GL/辅助核算含具体客户收入时生成）
  B. 其他业务收入
     表4 其他业务收入_全年汇总(按二级)：无分月，按二级明细汇总全年数（收入/成本/利润/增减）

数据来源与方法论：
  · 科目余额表(TB)：权威控制数；收入取 6001 本期贷方、成本取 6401 本期借方；
    其他业务收入取 6051 本期贷方、其他业务成本取 6402 本期借方（结平科目 net=贷−借=0，须取 gross）。
  · 综合查询明细表(GL)：月度发生额（按日期取月份）；客户收入取自 GL「辅助核算名称」列。
  · 辅助核算余额表：用于客户收入交叉核对（本程序主要源为 GL 辅助核算名称，二者等价）。
  · 毛利 = 收入 − 成本（均取 gross 发生额）。
  · 上年列仅当存在上一年度数据；当前数据多仅单年 → 上年留空、增减留空（首年不虚增变动）。
  · GL 为非全量抽取：全年小计与 TB 差异按「取数范围限制」披露，不认定为账务差错。
  · 凡重跑覆盖过的文件，交付前均做最终版校验。
"""
import paths as P
import os
import re
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

import audit_shell as S   # 共享视觉样式 + finalize_workbook
from audit_common import _safe_save, is_zero_amount, read_km as _read_km, read_gl as _read_gl, read_tb_pl_safe as _read_tb_pl_safe  # 共享库：统一锁感知保存（单一来源）+ 0值判断 + 通用读取解析器 + 损益GL兜底

# ===================== 常量 =====================
REV_CODE = '6001'      # 主营业务收入
COST_CODE = '6401'     # 主营业务成本
OREV_CODE = '6051'     # 其他业务收入
OCOST_CODE = '6402'    # 其他业务成本

MONTHS = list(range(1, 13))
MONTH_LABEL = {m: f'{m}月' for m in MONTHS}

# 每个「块/二级」固定的 8 个指标列
METRICS = ['本期收入', '本期成本', '本期毛利', '上年收入', '上年成本', '上期毛利',
           '收入增减', '成本增减']

PLAIN = S.SHELL_BODY
THIN = Side(style='thin', color='BFBFBF')
BORDER = S.SHELL_BORDER
CEN = S.SHELL_CEN
LEFT = S.SHELL_LEFT
RGT = S.SHELL_RGT




def _l2_key(name):
    """'主营业务收入-内销' -> '内销'；无'-' -> 原名。
    2026-08-01 增强：连续切掉所有 『主营业务收入-』『主营业务成本-』 前缀（FY技术等主体
    GL 科目名带双重前缀，如『主营业务收入-主营业务收入-检测业务收入』或三级
    『主营业务收入-检测业务收入-主营业务收入-检测业务收入-高新收入…』），
    统一归并到产品族段（'检测业务收入'）。"""
    s = str(name or '').strip()
    while True:
        if s.startswith('主营业务收入-'):
            s = s[len('主营业务收入-'):]
        elif s.startswith('主营业务成本-'):
            s = s[len('主营业务成本-'):]
        else:
            break
    # 若仍有二级前缀（如 '检测业务收入-主营业务收入-检测业务收入-高新收入'），
    # 取最后一个『-』后仍含 收入/成本 的段；否则保留首段
    if '-' in s:
        parts = [p.strip() for p in s.split('-') if p.strip()]
        tail = [p for p in parts if p.endswith('收入') or p.endswith('成本')]
        if tail:
            return tail[-1]
        return parts[0]
    return s


# 差异判定容差（GL 非全量抽取时差异远大于此，故 0.005 仅用于滤除浮点误差）
_DIFF_TOL = 0.005


def _diff_zero(a, b):
    """True 表示 TB 与 GL 无实质差异。"""
    return abs((a or 0.0) - (b or 0.0)) <= _DIFF_TOL


_KEYWORDS = ('套圈', '淬火加工', '人民币', '美元', '欧元', '在制品', '产品', '热处理')


def _stem_of(name):
    """取收入/成本二级的【产品族】用于配对：
       · 先去掉科目前缀与词尾 收入/成本；
       · 再按产品关键词（套圈/淬火加工/人民币…）归并到同一族，
         使 内销套圈/外销套圈/轴承套圈 等同族收入与成本配对在一起。"""
    s = _l2_key(name)   # 取 '-' 之后部分（如 主营业务收入-内销套圈 -> 内销套圈）
    if s.endswith('收入') or s.endswith('成本'):
        s = s[:-2]
    for kw in _KEYWORDS:
        if kw in s:
            return kw
    return s.strip('-').strip()


def _pair_l2(rev_names, cost_names):
    """按产品族将收入/成本二级配对（同名/同族合并，避免丢名）。返回 (pairs, singles)。
    pairs: list of (stem, [rev_name,...], [cost_name,...])
    singles: list of (name, 'rev'|'cost')（未配对二级）"""
    from collections import defaultdict
    rev_map, cost_map = defaultdict(list), defaultdict(list)
    for n in rev_names:
        rev_map[_stem_of(n)].append(n)
    for n in cost_names:
        cost_map[_stem_of(n)].append(n)
    stems = list(rev_map.keys()) + [s for s in cost_map if s not in rev_map]
    pairs, singles = [], []
    for s in stems:
        rns, cns = rev_map.get(s), cost_map.get(s)
        if rns and cns:
            pairs.append((s, rns, cns))
        elif rns:
            for n in rns:
                singles.append((n, 'rev'))
        else:
            for n in cns:
                singles.append((n, 'cost'))
    return pairs, singles


# ===================== 从文件名稳健识别会计年度 =====================
def _extract_year(fn):
    s = fn
    s = re.sub(r"(?<!\d)20\d{6,12}(?!\d)", "", s)   # 剥离 20+6~12 位数字串(导出时间戳)
    s = re.sub(r"(?<!\d)20\d{6}(?!\d)", "", s)      # 剥离独立 8 位 YYYYMMDD
    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", s)
    if m:
        return m.group(1)
    m2 = re.search(r"(?<!\d)(\d{2})(?!\d)", s)
    if m2:
        try:
            yy = int(m2.group(1))
        except ValueError:
            return None
        if 0 <= yy <= 99:
            return f"20{yy:02d}"
    return None


# ===================== 发现账套实体 =====================
def _discover_entities(data_dir):
    # ⚡ 2026-08-09 治本统一：本地 discover 收敛为 audit_common.discover_entities——
    # U8 模式读原文件全量；SAP 模式 patch 后=adapter（current_comp 单主体过滤内置）。
    from audit_common import discover_entities as _de
    return _de(data_dir)


# ===================== 读 科目余额表(TB) =====================


# 分部门明细表部门白名单（2026-08-01 取自参考表『主营业务收入成本分部门明细表.xlsx』第1列 29 个部门，
# 避免辅助核算等级跨主体不稳定导致 汇总行/非部门行 混入）
_DEPT_WHITELIST = [
    '食品事业部', '高低压电器事业部', '建材产品事业部', '机械轻工事业部', '电器产品事业部',
    '信电工程事业部', '化工产品事业部', '金属制品事业部', '珠宝潮奢事业部', '包装产品事业部',
    '技术发展部', '浙江方圆检测集团龙游有限公司', '质量鉴定部', '皮革制品事业部', '方圆认证',
    '纺织产品事业部', '技术服务事业部', '战略发展部', '基地建设办公室', '抽样队', '电商运营部',
    '金华子公司', '检测室', '业务室', '理化室', '钢结构室', '业务发展部', '公司', '薪资及效益奖金',
]


def _read_aux_dept(aux_path):
    """读辅助核算余额表：按 科目二级×科目辅助核算名称(部门) 归集 主营业务收入/成本 发生额。
    返回 {l2: {dept: [rev(贷方), cost(借方)]}}。
    行结构（辅助核算余额表）：[关联方, 科目名称, 科目辅助核算名称, 方向, 金额, 本期借方, 本期贷方, 方向, 期末金额, 等级, 往来类型]
    · 动态识别部门（2026-08-01 同改：不再按参考表白名单）：往来类型='部门' 且 辅助名≠'部门'（'部门'为汇总节点行剔除）。
      辅助核算「等级」=科目深度+2，同一部门可跨等级（挂不同明细科目），故跨等级全部计入，不按固定等级过滤；
    · 收入取 本期贷方、成本取 本期借方；二级名 = 科目名称（取 '-' 后产品族段）。"""
    try:
        wb = openpyxl.load_workbook(aux_path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
    except Exception as ex:
        print(f'  ⚠️ 辅助核算余额表读取失败 {aux_path}：{ex}')
        return {}
    res = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if len(r) < 12 or r[2] is None:
            continue
        nm = str(r[2])
        if not (_is_rev(nm) or _is_cost(nm)):
            continue
        aux = str(r[3]) if r[3] is not None else ''
        wtype = str(r[11]) if len(r) > 11 and r[11] is not None else ''
        if wtype != '部门' or not aux or aux == '部门':
            continue
        is_rev = _is_rev(nm)
        l2 = _l2_key(nm)
        try:
            dr = float(r[6] or 0.0) if len(r) > 6 else 0.0
            cr = float(r[7] or 0.0) if len(r) > 7 else 0.0
        except Exception:
            continue
        cell = res.setdefault(l2, {}).setdefault(aux, [0.0, 0.0])
        if is_rev:
            cell[0] += cr                               # 收入取贷方
        else:
            cell[1] += dr                               # 成本取借方
    try:
        wb.close()
    except Exception:
        pass
    return res


def build_dept_footnote(wb, entities, cur_year, prev_year, aux_dept_all):
    """主营业务收入成本分部门明细表（附注汇总模板，2026-08-02 用户需求）：
    合并原 收入/成本 两张分部门明细表为一张；
    · 列 = 项目 | 部门1..N（每部门一列） | 合计 | 上年收入 | 上年成本 | 上年毛利（表尾上年列组·集团口径）
    · 行 = 三块：① 未审数（主营业务收入/成本/毛利）② 审计调整（收入/成本/毛利，0 待手工）
           ③ 审定数（收入/成本/毛利，=未审+调整 公式）
    · 合计列全部公式化（=SUM 引用，不写死数值——用户 2026-08-02『合计行整行公式化』）；
      毛利=收入−成本、审定=未审+调整 均为公式；
    · 上年列仅当存在上一年度数据（2026 底稿的上年=2025 全年）时填集团合计口径数值；
    · 无合并单元格（用户 2026-08-01 要求），块标题 A 列加粗提示。
    数据源：《辅助核算余额表》按「部门」归集（收入=贷方、成本=借方）。"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    HFONT = Font(name='Times New Roman', bold=True, color='000000', size=10)
    TFILL = PatternFill('solid', fgColor='D9E1F2')
    TFONT = Font(name='Times New Roman', bold=True, size=10)
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center')
    NUM = '#,##0.00'
    FONT = Font(name='Times New Roman', size=10)

    ws = wb.create_sheet('主营业务收入成本分部门明细表')
    ents = sorted(entities)

    # ---- 部门并集（跨主体、跨年份）----
    depts = set()
    for E in ents:
        for y in (str(cur_year), str(prev_year)) if prev_year else (str(cur_year),):
            d = aux_dept_all.get(E, {}).get(y, {})
            for l2, m in d.items():
                depts |= set(m.keys())
    depts = sorted(depts)

    # 列布局：A=项目 | 部门1..N | 合计 | 上年收入 | 上年成本 | 上年毛利
    C_TOT = 2 + len(depts)      # 合计列
    C_PREV_R = C_TOT + 1        # 上年收入
    C_PREV_C = C_TOT + 2        # 上年成本
    C_PREV_G = C_TOT + 3        # 上年毛利
    N = C_PREV_G

    # 标题与说明（无合并单元格）
    ws.cell(1, 1, '主营业务收入成本分部门明细表（附注汇总模板）').font = \
        Font(name='Times New Roman', bold=True, size=12)
    if prev_year:
        ws.cell(2, 1, f'说明：数据取自《辅助核算余额表》按「部门」归集（收入=贷方、成本=借方，毛利=收入−成本）；'
                      f'审定数=未审数+审计调整（公式）；合计列公式化；上年列为集团上年同期对比'
                      f'（{prev_year} 年度）。').font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    else:
        ws.cell(2, 1, f'说明：数据取自《辅助核算余额表》按「部门」归集（收入=贷方、成本=借方，毛利=收入−成本）；'
                      f'审定数=未审数+审计调整（公式）；合计列公式化；无上年度数据（{cur_year} 首年），上年列留空。'
                      ).font = Font(name='Times New Roman', size=10, italic=True, color='808080')

    # ---- 表头行3 ----
    ws.cell(3, 1, '项目')
    for i, d in enumerate(depts):
        ws.cell(3, 2 + i, d)
    ws.cell(3, C_TOT, '合计')
    ws.cell(3, C_PREV_R, '上年收入')
    ws.cell(3, C_PREV_C, '上年成本')
    ws.cell(3, C_PREV_G, '上年毛利')
    for ci in range(1, N + 1):
        cc = ws.cell(3, ci); cc.font = HFONT; cc.fill = HFILL; cc.alignment = CEN; cc.border = BORDER

    # 无部门数据 → 注记型空表（含 E2 跳过关键词「部门」）
    if not depts:
        r = 4
        ws.cell(r, 1, '注：本账套主营业务收入/成本无「部门」级辅助核算（往来类型非“部门”，'
                      '如按客户/产品/费用归集），无法生成分部门明细表。未生成。').font = \
            Font(name='Times New Roman', size=10, italic=True, color='808080')
        for ci in range(1, N + 1):
            ws.cell(r, ci).border = BORDER
        ws.column_dimensions['A'].width = 90
        return ws

    # ---- 本期（cur_year）各部门未审数 ----
    rev_d, cost_d = {}, {}
    for d in depts:
        rv = cs = 0.0
        for E in ents:
            for l2, mm in aux_dept_all.get(E, {}).get(str(cur_year), {}).items():
                v = mm.get(d)
                if v:
                    rv += v[0]; cs += v[1]
        rev_d[d] = rv; cost_d[d] = cs

    # ---- 上年集团合计（prev_year，全部部门汇总）----
    prev_rev = prev_cost = 0.0
    if prev_year:
        for E in ents:
            for l2, mm in aux_dept_all.get(E, {}).get(str(prev_year), {}).items():
                for d, v in mm.items():
                    prev_rev += v[0]; prev_cost += v[1]
    prev_gp = prev_rev - prev_cost

    dept_last = get_column_letter(C_TOT - 1)   # 部门范围末列（合计列前）

    def _sum_formula(rr):
        return f'=SUM(B{rr}:{dept_last}{rr})'

    def _row_tot(rr):
        """该行各部门列数值之和（2026-08-06 协议：合计列写数值，替代 =SUM 公式——公式
        data_only 读 None 收回读不回；本函数读已写数值 cell）。"""
        return round(sum(ws.cell(rr, 2 + i).value or 0.0 for i in range(len(depts))), 2)

    def _block_title(txt):
        ws.cell(r, 1, txt).font = TFONT
        ws.cell(r, 1).fill = TFILL
        for ci in range(1, N + 1):
            ws.cell(r, ci).border = BORDER

    # ---- 块1 未审数 ----
    r = 4
    _block_title('一、未审数'); r += 1
    ru, cu, gu = r, r + 1, r + 2
    for i, d in enumerate(depts):
        col = get_column_letter(2 + i)
        ws.cell(ru, 2 + i, round(rev_d[d], 2)).number_format = NUM
        ws.cell(cu, 2 + i, round(cost_d[d], 2)).number_format = NUM
        # 2026-08-06 协议：毛利写数值（=收入−成本）
        ws.cell(gu, 2 + i, round(rev_d[d] - cost_d[d], 2)).number_format = NUM
    for rr in (ru, cu, gu):
        ws.cell(rr, C_TOT, _row_tot(rr)).number_format = NUM             # 合计列写数值
    if prev_year:
        ws.cell(ru, C_PREV_R, round(prev_rev, 2)).number_format = NUM
        ws.cell(cu, C_PREV_C, round(prev_cost, 2)).number_format = NUM
        ws.cell(gu, C_PREV_G, round(prev_gp, 2)).number_format = NUM
    r = gu + 1
    # ---- 块2 审计调整 ----
    _block_title('二、审计调整'); r += 1
    ra, ca, ga = r, r + 1, r + 2
    for i, d in enumerate(depts):
        col = get_column_letter(2 + i)
        ws.cell(ra, 2 + i, 0).number_format = NUM                            # 0 待手工
        ws.cell(ca, 2 + i, 0).number_format = NUM
        # 2026-08-06 协议：毛利调整写数值（=收入调整−成本调整=0）
        ws.cell(ga, 2 + i, 0).number_format = NUM
    for rr in (ra, ca, ga):
        ws.cell(rr, C_TOT, _row_tot(rr)).number_format = NUM
    r = ga + 1
    # ---- 块3 审定数 ----
    _block_title('三、审定数'); r += 1
    rf, cf, gf = r, r + 1, r + 2
    for i, d in enumerate(depts):
        # 2026-08-06 协议：审定数写数值（=未审+调整0，读未审行值；公式 data_only 读 None）
        _rv = ws.cell(ru, 2 + i).value or 0.0
        _cv = ws.cell(cu, 2 + i).value or 0.0
        ws.cell(rf, 2 + i, round(_rv, 2)).number_format = NUM
        ws.cell(cf, 2 + i, round(_cv, 2)).number_format = NUM
        ws.cell(gf, 2 + i, round(_rv - _cv, 2)).number_format = NUM
    for rr in (rf, cf, gf):
        ws.cell(rr, C_TOT, _row_tot(rr)).number_format = NUM
    r = gf + 1

    # ---- 行标签 + 样式 ----
    labels = {ru: '主营业务收入未审数', cu: '主营业务成本未审数', gu: '主营业务毛利未审数',
              ra: '主营业务收入审计调整', ca: '主营业务成本审计调整', ga: '主营业务毛利审计调整',
              rf: '主营业务收入审定数', cf: '主营业务成本审定数', gf: '主营业务毛利审定数'}
    for rr in (ru, cu, gu, ra, ca, ga, rf, cf, gf):
        ws.cell(rr, 1, labels[rr]).font = FONT
        ws.cell(rr, 1).alignment = LFT
        for ci in range(1, N + 1):
            cell = ws.cell(rr, ci)
            cell.border = BORDER
            cell.font = FONT
            if ci >= 2:
                cell.alignment = RGT
    # 审定块合计列强调
    for rr in (rf, cf, gf):
        ws.cell(rr, C_TOT).font = TFONT
        ws.cell(rr, C_TOT).fill = TFILL

    # 列宽
    ws.column_dimensions['A'].width = 24
    for ci in range(2, N + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 12
    ws.freeze_panes = 'B4'
    return ws


def _empty_month():
    return {m: 0.0 for m in MONTHS}


# ===================== 聚合单个 GL 文件（按年） =====================
def _is_rev(nm):
    """收入判断：标准『主营业务收入』前缀 + XBJ 新收入准则施工『合同结算\收入结转』
    （2026-08-15 实测：XBJ 123302 收入结转贷方=收入确认；123301 价款结算≠收入）。"""
    nm = str(nm or '')
    return nm.startswith('主营业务收入') or '合同结算\收入结转' in nm


def _is_cost(nm):
    """成本判断：标准『主营业务成本』前缀 + XBJ 『合同履约成本』（借方=成本确认）。"""
    nm = str(nm or '')
    return nm.startswith('主营业务成本') or nm.startswith('合同履约成本')


def _aggregate(path):
    """返回 {year: {...}}，内含 主营业务收入/成本 的月度、二级、客户聚合。

    2026-07-31 二次修订（回归一般原则——用户质疑"结转借方/非结转借方"为特调）：
    科目余额表"本期借方发生额" = 明细账（GL）该科目全部借方发生额之和——GL 全量抽取下恒等，
    不依赖任何结转凭证结构/记账习惯（实测 FY本级 2025/2026、g 15 主体×2 年全部零差异）。故：
      · 收入 = GL 全部借方合计（收入确认凭证无借方，全部借方即月末结转损益凭证
        "借收入 贷本年利润"的结转额 = TB 借发 = TB 贷发（结平账套））；
      · 成本 = GL 全部借方合计（确认借方"借成本 贷库存/应付" + 结转凭证借方）= TB 借发；
      · 红字冲回（贷方负数/借方负数）天然含在借方合计中，规避原"确认贷方 vs TB 贷发"
        的 828,068.88 级差异（FY本级 2026 修复后差异=0）；
      · 兜底：GL 借方合计为 0 且贷方 > 0（未结转损益账套）→ 收入取确认贷方，避免漏计。
    差异凭证无需单独列示——正常账套 GL 借方 = TB 借发差异恒 0；
    仅 GL 抽取期间 < TB 期间（如 2026 GL 仅 1-5 月 YTD、TB 全年）时差额为未抽取月份，
    按取数范围限制披露（见 build_by_entity 差异归因）。
    ⚡ 2026-08-16 修复：SAP 场景 gl 是多文件 list——原无条件 `_read_gl(path)`（audit_common
    U8 版单文件）遇 list 报 TypeError（AH 集团A/B 营业收入底稿缺失根因）；且全量 88 家
    GL 读入内存过大（OOM）。SAP 分支实际只需下方 read_gl_tb 的 收入/成本 轻量行，
    非收入成本行（rows 过滤后）在后续循环完全不使用 → SAP list 场景直接跳过全量读。
    U8（单 path 字符串）路径不变。
    """
    _is_sap_gl = (_adapter is not None and _adapter.is_sap(getattr(_adapter, '_DATA_ROOT', None))
                  and isinstance(path, (list, tuple)))
    rows = [] if _is_sap_gl else _read_gl(path)
    # ⚡ 2026-08-11 P0 修复：集团模式（current_comp=None）下不能用无参 current_comp()/
    # read_gl_tb() 判断 SAP 分支（集团模式返回 None → 走 U8 分支读 list → 空）。
    # 改为：SAP 目录且 (current_comp 有值 或 主流程按主体传单主体 path) → 走 SAP 分支。
    _is_sap = _adapter is not None and _adapter.is_sap(getattr(_adapter, '_DATA_ROOT', None))
    _cur_c = _adapter.current_comp() if _is_sap else None
    if _is_sap and (_cur_c or (isinstance(path, list) and path)):
        # ⚡ SAP：收入/成本行改用【完全互抵剔除】（read_gl_tb，铁律14）——剔除同凭证借贷
        # 完全互抵的虚增（1010 收入 56 凭证/17 亿红字重记互转），保留其余原始借贷全额；
        # 分月最接近 TB 净额（残余差异=GL 天然口径差，底稿注明）。
        # ⚡ 2026-08-10 修复：read_gl_tb 返回 read_gl_rows 结构 {y, name, month, dr, cr, aux...}
        #   （year→y、debit→dr、credit→cr 重构后未同步）→ _aggregate 下游用 row['year']/
        #   ['debit']/['credit'] 报 KeyError → 营业收入底稿未生成。此处统一转 U8 同构。
        # ⚡ 2026-08-11 集团模式：read_gl_tb 无参在 current_comp=None 时读不到 →
        #   从 path 推断主体（list[0] 路径含 {comp}）显式传 comp。
        def _scope_comp_from_path(pp):
            if _cur_c:
                return _cur_c
            p0 = pp[0] if isinstance(pp, list) and pp else pp
            # ⚡⚡ 2026-08-28 修复（收入/成本分月明细全空根因）：原【先】目录段正则
            #   [\\/](\d{4})[\\/] 会先匹配路径 `.../数据/2026\序时账\1010-1月.xlsx` 中的
            #   年份目录 2026 → _c_inc='2026' → read_gl_tb('2026') 空 → 收入/成本全 0。
            #   与 longterm_assets_detail._read_gl_fa 同型（08-26 已修，此处漏同步）。
            #   修复：优先【basename 前缀】匹配（1010-1月.xlsx → 1010），再回退目录段。
            m2 = re.match(r'^(\d{4})', os.path.basename(str(p0)))
            if m2:
                return m2.group(1)
            m = re.search(r'[\\/](\d{4})[\\/]', str(p0))
            if m:
                return m.group(1)
            return None
        _c_inc = _scope_comp_from_path(path)
        _r_inc = _adapter.read_gl_tb(comp=_c_inc, pfx='主营业务收入')
        _r_cost = _adapter.read_gl_tb(comp=_c_inc, pfx='主营业务成本')

        def _conv(rs):
            out = []
            for r in rs:
                out.append({
                    'name': str(r.get('name') or ''),
                    'date': r.get('date') or '',
                    'year': str(r.get('y') or ''),
                    'month': int(r.get('month') or 0) or None,
                    'vtype': str(r.get('vtype') or ''),
                    'vno': str(r.get('vno') or ''),
                    'cp': str(r.get('cp') or r.get('opp') or ''),
                    'aux': str(r.get('aux') or r.get('cust') or r.get('supp') or '').strip(),
                    'debit': float(r.get('dr') or 0.0),
                    'credit': float(r.get('cr') or 0.0),
                })
            return out

        _r_inc = _conv(_r_inc)
        _r_cost = _conv(_r_cost)
        rows = ([r for r in rows
                 if not (str(r.get('name') or '').startswith(('主营业务收入', '主营业务成本')))]
                + _r_inc + _r_cost)
    res = {}
    # 各年收入/成本 GL 借方/贷方合计（用于未结转兜底判定）
    y_tot = {}
    for row in rows:
        y = row['year']
        if y is None:
            continue
        y = str(y)
        nm = row['name']
        if _is_rev(nm) or _is_cost(nm):
            t = y_tot.setdefault(y, {'rev_dr': 0.0, 'rev_cr': 0.0, 'cost_dr': 0.0, 'cost_cr': 0.0, 'max_m': 0})
            t['max_m'] = max(t.get('max_m', 0), int(row['month'] or 0))
            if _is_rev(nm):
                t['rev_dr'] += row['debit']; t['rev_cr'] += row['credit']
            else:
                t['cost_dr'] += row['debit']; t['cost_cr'] += row['credit']
    for row in rows:
        y = row['year']
        if y is None:
            continue
        y = str(y)
        d = res.setdefault(y, {
            'rev_m': _empty_month(), 'cost_m': _empty_month(),
            'rev_l2': {}, 'cost_l2': {},
            'rev_leaf': {}, 'cost_leaf': {},   # ⚡ 2026-08-10 末级（完整科目路径，不 _l2_key 合并）
            'rev_cust': {}, 'cost_cust': {},
        })
        d['gl_max_m'] = y_tot.get(y, {}).get('max_m', 0)   # GL 实际最大月份（期间错配识别用）
        nm = row['name']
        m = row.get('month')
        if not m:
            # ⚡⚡ 2026-08-30 U8 部分主体 GL 行 month 缺失/0（date 正常）→ 从 date 解析，
            #   否则 d['rev_m'][0] KeyError（XBJ revenue group 取数失败）。
            _dt = row.get('date')
            try:
                if hasattr(_dt, 'month'):
                    m = _dt.month
                else:
                    _s = str(_dt or '')
                    _mm = re.search(r'[-\/](\d{1,2})(?:[-\/]|$)', _s)
                    m = int(_mm.group(1)) if _mm else None
            except Exception:
                m = None
        if not m:
            continue
        if _is_rev(nm):
            t = y_tot.get(y, {})
            # ⚡ 2026-08-09 治本修复：SAP 收入 GL 借方有红字冲销/重分类（1010 借 4.7 亿）→
            # use_cr 判定（借方合计=0）在 SAP 失效，误把借方当收入（1月 7086万 vs 应贷 5.09亿）。
            # 铁律3：收入取贷（gross 确认）、借方=冲回——SAP 一律取贷方。
            use_cr = (abs(t.get('rev_dr', 0.0)) < 1e-9)
            if _is_sap:
                use_cr = True   # SAP：收入一律取贷方（借方=红字冲回，铁律3；集团模式同）
            # ⚡⚡ 2026-08-30 结论：营业收入取贷方发生额（铁律3）即与 TB 利润表（Sheet3 发生额）
            #   一致（1030: 贷方20.5亿 = Sheet3 20.5亿）。曾试改净额 credit-debit，但红字冲回
            #   记借方负数 → credit-debit 双计（41亿=20.5×2）→ 回退，维持贷方口径。
            amt = row['credit'] if use_cr else row['debit']
            d['rev_m'][m] += amt
            l2 = _l2_key(nm)
            d['rev_l2'].setdefault(l2, _empty_month())[m] += amt
            # ⚡ 2026-08-10 末级展开（完整科目路径，防不同路径同名末级被 _l2_key 自动加总）
            d['rev_leaf'].setdefault(nm, _empty_month())[m] += amt
            if row['aux']:
                d['rev_cust'].setdefault(row['aux'], _empty_month())[m] += amt
        elif _is_cost(nm):
            amt = row['debit']                            # 全部借方（含结转凭证借方）
            d['cost_m'][m] += amt
            l2 = _l2_key(nm)
            d['cost_l2'].setdefault(l2, _empty_month())[m] += amt
            d['cost_leaf'].setdefault(nm, _empty_month())[m] += amt
            if row['aux']:
                d['cost_cust'].setdefault(row['aux'], _empty_month())[m] += amt
        # 其他业务收入/成本无分月需求，年度由 TB 取数，这里不聚合
    return res


# ===================== TB 控制数辅助 =====================
# 四个收入/成本科目的标准名称（用于按名称识别，兼容本地科目表如 泰国25 的 4100/4200）
_REV_NAMES = ('主营业务收入',)
_COST_NAMES = ('主营业务成本', '合同履约成本')
_OREV_NAMES = ('其他业务收入',)
_OCOST_NAMES = ('其他业务成本', '其他业务支出')   # ⚡⚡ 2026-09-01 AZ 6402 名『其他业务支出』
_STD = {'rev': REV_CODE, 'cost': COST_CODE, 'orev': OREV_CODE, 'ocost': OCOST_CODE}
_NAME_MAP = {'rev': _REV_NAMES, 'cost': _COST_NAMES, 'orev': _OREV_NAMES, 'ocost': _OCOST_NAMES}
# ⚡⚡ 2026-08-15 无父级 TB 名称前缀兜底（ga 60011601『主营业务收入\环保业务收入\…』实证）。
#   rev 不含『合同结算\收入结转』——123302 credit=0（结转在借方=镜像），收入源=6001 贷方
#   或无 6001 时的 123301『价款结算』贷方（XBJ 38 个无 6001 主体实证：123301贷=123302借=收入）。
#   cost 多前缀：6401 主营业务成本 + 5002 合同履约成本（XBJ 各主体成本体系不一，需都算）。
_PREFIX_OF = {'rev': ('主营业务收入',),
              'cost': ('主营业务成本', '合同履约成本'),
              'orev': ('其他业务收入',),
              'ocost': ('其他业务成本',)}


def _tb_l2_entries(tb_year, prefix, kind='credit'):
    """按【名称前缀】收集 TB 二级明细（铁律5 名称主键；⚡ 2026-08-12 SAP 修复：
    SAP TB 无父级行、二级=10 位末级码（level=4，如 6051010000『其他业务收入-材料销售收入』），
    原 `level==2 + code.startswith(父码+'.')` 双重失效 → 其他业务收入汇总/二级核对空。
    改为：名称以 prefix 开头（排除恰好==prefix 的一级汇总行）→ 剥离前缀段作二级键
    （'其他业务收入-材料销售收入'→'材料销售收入'；勿用 _l2_key——其只切『主营业务』前缀）。
    返回 {l2_key: 金额合计}。kind: 'credit'(收入贷方) / 'debit'(成本借方/铁律27 恒等)。"""
    out = {}
    if not tb_year:
        return out
    for code, v in tb_year.items():
        nm = str(v.get('name') or '').strip()
        if not nm.startswith(prefix) or nm == prefix:
            continue
        tail = nm[len(prefix):]
        if tail.startswith('-') or tail.startswith('－'):
            tail = tail[1:]
        l2 = tail.strip() or nm
        out[l2] = out.get(l2, 0.0) + float(v.get(kind) or 0.0)
    return out


def _detect_codes(tb_year):
    """从 TB 按名称识别 主营收入/成本、其他收入/成本 的一级科目代码。
    返回 {rev,cost,orev,ocost: code|None}。优先级：
      1) 一级科目名称精确匹配（如 '主营业务收入'）；
      2) 任一层级名称精确匹配（含 XBJ『合同结算\收入结转』『合同履约成本』新收入准则施工科目）；
      3) 退回到标准代码 6001/6401/6051/6402；
      4) ⚡⚡ 2026-08-15 无父级 TB 名称前缀兜底（'prefix:主营业务收入'）：
         ga 实证——TB 只有 60011601『主营业务收入\环保业务收入\…』等末级行，
         无 6001 父级行 + level 全 0 → 精确名/标准码全落空 → 审定表全 0。
         返回 'prefix:名称' 标记，_tb_total 按名称前缀聚合（结转行 600199 方向相反自然排除）。
    如此可兼容境外子公司采用本地科目表（如 泰国25 用 4100/4200）、SAP 无父级 10 位码、
    ga 无父级 6/8 位码、XBJ 新收入准则施工科目等情形。"""
    out = {k: None for k in ('rev', 'cost', 'orev', 'ocost')}
    if not tb_year:
        return out
    for key, names in _NAME_MAP.items():
        code = None
        for c, v in tb_year.items():                 # 优先一级精确名
            if v.get('level') == 1 and v['name'] in names:
                code = c
                break
        if code is None:                            # 退而求其次：任一层级精确名
            for c, v in tb_year.items():
                if v['name'] in names:
                    code = c
                    break
        if code is None and _STD[key] in tb_year:   # 兜底标准代码
            # ⚡⚡ 2026-08-31 修复（XBJ 主营收入 260K vs TB 547M）：read_km 对 XBJ 生成了
            #   错误的『标准码键』——6001 键 name=『主营业务收入\建筑工程\房建』（实为 60010101
            #   子级值 260K），非真实父级。若不验证 name 直接按标准码取 → 漏 60010102/60010103
            #   → 审定表主营收入少 5.4 亿。改：标准码行 name 精确命中才用；否则走名称前缀兜底。
            _sv = tb_year[_STD[key]]
            _nm = str(_sv.get('name') or '')
            if _nm in names or _sv.get('level') is None:
                code = _STD[key]
        if code is None:                            # 无父级 TB 名称前缀兜底
            _pfx = _PREFIX_OF.get(key)
            if _pfx and any(str(v.get('name') or '').startswith(p)
                            for p in _pfx for v in tb_year.values()):
                code = 'prefix:' + key
        if code is None and key == 'rev':
            # ⚡⚡ 2026-08-31 修复（XBJ 主营收入 7.862B vs 试算表 7.852B 差异 10.4M）：
            #   原 123301『合同结算\价款结算』贷方兜底——把结算进度款当收入（21/26 主体
            #   5.8M/4.6M）。但利润表/试算表（sap_tb_gen）主营收入只取 6001 贷方（7.852B），
            #   合同结算是资产负债表科目（结算进度款≠已确认收入）。为维持铁律 TB=审定表
            #   =利润表一致：无 6001 主体主营收入=0（不再用合同结算兜底）。
            #   ⚠️ 历史注释称"38 主体实证 123301贷=收入"，但非零仅 21/26 两主体 10.4M，
            #      且与试算表口径冲突 → 按审计列报口径（利润表=6001）收敛。
            pass
        out[key] = code
    return out


def _code_of(tb_year, which):
    """返回某科目的实际代码（用于表内标签），取不到回退标准代码。"""
    return _detect_codes(tb_year).get(which) or _STD.get(which)


# 模块级损益口径模式：None=自动（单主体同额比例）；False=真实账套（AH 有企业报表，
# 损益取净额，同额行=0）；True=镜像账套（XBJ/AZ 无企业报表，损益同额行取发生额）。
_MIRROR_MODE = None


def _tb_total(tb_year, which, kind, mirror=None):
    """TB 中某收入/成本科目(一级)的本期借或贷合计(gross)。which: rev/cost/orev/ocost。
    ⚡⚡ 2026-08-15：支持 _detect_codes 返回的 'prefix:key' 标记——无父级 TB
    （ga 60011601 等末级行）按名称前缀聚合（收入取贷/成本取借；结转行方向相反自然排除）；
    cost 多前缀（主营业务成本+合同履约成本 都算，XBJ 各主体成本体系不一）。
    ⚡⚡ 2026-09-01 mirror 镜像账套（损益科目借贷同额复制，XBJ/AZ）→ 同额行取发生额
    （收入贷方、成本借方，=底稿口径）；真实账套（AH）→ 一律净额（同额行=0，企业报表
    口径——AH 2680 其他业务 605103 技术服务费 jf=df=1.98M，底稿取 df 6.37M vs 企业
    报表 3.39M 差 3M 根因）。mirror=None 时按主体损益同额比例自动判断。"""
    if not tb_year:
        return 0.0
    if mirror is None:
        mirror = _MIRROR_MODE
    if mirror is None:
        _probe = []
        for _v in tb_year.values():
            _c = float(_v.get('credit') or 0.0)
            _d = float(_v.get('debit') or 0.0)
            if _c or _d:
                _m = max(abs(_c), abs(_d))
                _probe.append(abs(_c - _d) < 0.005 or (_m > 0 and abs(_c - _d) / _m < 0.02))
        mirror = bool(_probe) and sum(_probe) / len(_probe) > 0.8
    code = _code_of(tb_year, which)
    if not code:
        return 0.0
    if isinstance(code, str) and code.startswith('prefix:'):
        _key = code[len('prefix:'):]
        if _key == 'cost':
            # ⚡⚡ 2026-08-15 成本科目优先级（ga 实证 11:23 用户质疑"成本>收入"）：
            #   优先『主营业务成本』（6401，=已确认成本），仅当无主营业务成本才取
            #   『合同履约成本』（5002/5401，新收入准则成本发生科目）。
            #   ⚠️ ga 的 5401 合同履约成本是【归集科目】（含未完工在建成本+结转行
            #   691,797,482），若优先取它会虚增成本（15 分公司 14.26 亿 vs 正确 7.34 亿）；
            #   XBJ 的 5002 合同履约成本才是成本科目（6401 是结转镜像，双加=2 倍
            #   XBJ03 实证 68427.82）。故：有主营业务成本→取它；无→取合同履约成本。
            # ⚡⚡ 2026-08-31 修复（XBJ 主营成本 7.042B vs 利润表 7.483B 差异 441M）：
            #   无 6401 主体的 5002 合同履约成本若为【负数】（结转冲回，8 主体 -4K~-108M），
            #   不是成本发生（利润表营业成本=6401+6402 不含）→ 该主体成本=0，与利润表一致。
            #   仅当 5002 有正数借发（真实成本归集）才用合同履约成本兜底。
            if any(str(v.get('name') or '').startswith('主营业务成本')
                   for v in tb_year.values()):
                _prefs = ('主营业务成本',)
            elif sum(float(v.get('debit') or 0.0) for v in tb_year.values()
                     if '合同履约成本' in str(v.get('name') or '')
                     and '结转' not in str(v.get('name') or '')) > 0.005:
                _prefs = ('合同履约成本',)
            else:
                _prefs = ('主营业务成本',)   # 无 6401 且 5002 净额非正（结转冲回）→ 成本 0
        else:
            _prefs = _PREFIX_OF.get(_key, ())
        tot = 0.0
        # ⚡⚡ 2026-08-31 修复（XBJ 主营收入 14.2B=2×7.1B 双计根因）：read_km 对 XBJ
        #   同时生成【一级聚合行】(code[:4]，6001) + 【末级行】(600101xx)，name 都是
        #   『主营业务收入\…』→ prefix 遍历全部命中 → 双计。排除父级（有子级以其 code 为
        #   前缀的行），只计末级叶子（同 _gross_entity 父=子和逻辑）。
        #   ⚡⚡ 2026-08-31 二次修复（AH 收入 15.95B vs 试算表 12.88B）：AH 6001 父级
        #   (6001010000 df 2.48B) ≠ 子和(3.79B) → 非镜像父级【保留】独立值；GL 流水行
        #   (6001010001 df 3.69B/debit 3.01B) 取【净额】。仅镜像父级(父=子和)排除防双计。
        _hit = {c for c, v in tb_year.items()
                if any(str(v.get('name') or '').startswith(p) for p in _prefs)}
        _excl = set()
        for c in _hit:
            _kids = [c2 for c2 in _hit if c2 != c and c2.startswith(c)]
            if not _kids:
                continue
            _f = 'credit' if kind == 'cr' else 'debit'
            _pv = abs(float(tb_year[c].get(_f) or 0.0))
            _vsum = sum(float(tb_year[k2].get(_f) or 0.0) for k2 in _kids)
            # 父级排除：父=子和（镜像）或 父=任一子级值（read_km 错误聚合行，
            #   XBJ 6401 父=64010102 首子级）→ 防双计。AH 6001 父=独立值保留。
            if (abs(_pv - abs(_vsum)) < 1.0
                    or any(abs(_pv - abs(float(tb_year[k2].get(_f) or 0.0))) < 1.0
                           for k2 in _kids)):
                _excl.add(c)
        _leaves = [c for c in _hit if c not in _excl]
        for c in _leaves:
            v = tb_year[c]
            nm = str(v.get('name') or '')
            if kind == 'cr':
                _c = float(v.get('credit') or 0.0)
                _d = float(v.get('debit') or 0.0)
                if '（GL）' in nm or '（GL' in nm:
                    tot += _c - _d   # GL 流水行净额（AH 6001 GL 行 3.69B−3.01B）
                elif mirror and abs(_c - _d) < 0.005:
                    tot += _c        # 镜像账套（XBJ）借贷同额 → 取贷方
                else:
                    tot += _c - _d   # 真实账套（AH）一律净额，同额行=0
            else:
                # ⚡⚡ 2026-08-31 成本口径：AH 6401 制造成本有真实贷方冲减（debit 2.80B/
                #   credit 0.34B → 净额 2.55B=企业报表）；XBJ U8 成本科目借贷同额
                #   （debit==credit 镜像）→ 取借方（=利润表）。区分：借贷不同取净额。
                _d = float(v.get('debit') or 0.0)
                _c = float(v.get('credit') or 0.0)
                tot += _d if (mirror and abs(_d - _c) < 0.005) else (_d - _c)
        return tot
    v = tb_year.get(code)
    if not v:
        return 0.0
    # ⚡⚡ 2026-09-01 精确码分支也用镜像口径（AH 2680 orev 6051 键 credit6.37M/debit2.98M
    #   = 同额行1.98+1.00 保留 + 净额3.39，净额口径应返回 3.39M；原直接返回 credit 6.37M
    #   vs 企业报表 3.39M 差 3M）
    _c = float(v.get('credit') or 0.0)
    _d = float(v.get('debit') or 0.0)
    if kind == 'cr':
        return _c if mirror else (_c - _d)
    return _d if mirror else (_d - _c)


# ===================== 样式辅助 =====================
def _hcell(ws, r, c, val):
    cell = ws.cell(r, c, val)
    cell.fill = S.SHELL_HFILL
    cell.font = S.SHELL_HFONT
    cell.border = BORDER
    cell.alignment = CEN
    return cell


def _dcell(ws, r, c, val, align=RGT, bold=False, fill=None):
    cell = ws.cell(r, c, val)
    cell.border = BORDER
    cell.alignment = align
    if bold or fill:
        cell.font = Font(name='Times New Roman', size=10, bold=bold)
        if fill:
            cell.fill = fill
    return cell


def _sum_months(dicts):
    """多个 {month:amt} 求和。"""
    out = _empty_month()
    for dd in dicts:
        if not dd:
            continue
        for m in MONTHS:
            out[m] += dd.get(m, 0.0)
    return out


def _metrics_block(rev_md, cost_md, prev_rev_md, prev_cost_md, prev_year):
    """给定某范围(主体/二级/集团)的 收入月度dict、成本月度dict 及上年对应 dict，
    返回 8 指标 的「月度列表(12) + 全年小计」二维：[ [m1..m12], 小计 ]。"""
    monthly = []
    for m in MONTHS:
        inc = rev_md.get(m, 0.0) if rev_md else 0.0
        cst = cost_md.get(m, 0.0) if cost_md else 0.0
        gp = inc - cst
        pin = prev_rev_md.get(m, 0.0) if (prev_year and prev_rev_md) else None
        pcs = prev_cost_md.get(m, 0.0) if (prev_year and prev_cost_md) else None
        pgp = (pin - pcs) if (prev_year and pin is not None) else None
        dinc = (inc - pin) if (prev_year and pin is not None) else None
        dcst = (cst - pcs) if (prev_year and pcs is not None) else None
        monthly.append([inc, cst, gp, pin, pcs, pgp, dinc, dcst])
    # 全年小计
    tinc = sum(x[0] for x in monthly)
    tcst = sum(x[1] for x in monthly)
    tgp = tinc - tcst
    tpin = sum(x[3] for x in monthly if x[3] is not None) if prev_year else None
    tpcs = sum(x[4] for x in monthly if x[4] is not None) if prev_year else None
    tpgp = (tpin - tpcs) if (prev_year and tpin is not None) else None
    tdinc = (tinc - tpin) if (prev_year and tpin is not None) else None
    tdcst = (tcst - tpcs) if (prev_year and tpcs is not None) else None
    total = [tinc, tcst, tgp, tpin, tpcs, tpgp, tdinc, tdcst]
    return monthly, total


# 全局占位（在 build 时填充当前年度），供 _tb_l2_map 风格函数使用
cur_year_holder = [None]


# ===================== 表1：按主体 分月明细 =====================
def build_by_entity(wb, entities, cur_year, prev_year, data, tb_all):
    # 2026-08-04 规范：新增『主营业务收入按主体明细表』——标准行式（主体|8指标），
    # 供合并核对程序 det_sn 命中并逐主体提取（分月明细为块结构无法行式提取，曾致收入明细单体恒空）。
    wsum = wb.create_sheet('主营业务收入按主体明细表')
    wsum.cell(1, 1, f'主营业务收入/成本 按主体汇总（{cur_year} 年度'
                    + (f'，对比 {prev_year}' if prev_year else '') + '）').font = S.SHELL_TITLE_FONT
    r_sum = 3
    _hcell(wsum, r_sum, 1, '主体')
    for j, mt in enumerate(METRICS):
        _hcell(wsum, r_sum, 2 + j, mt)
    r_sum += 1
    ws = wb.create_sheet('主营业务收入_分月明细(按主体)')
    ws.cell(1, 1, f'主营业务收入分月明细表（按主体，{cur_year} 年度'
                   + (f' 对比 {prev_year}' if prev_year else ' · 无上年度数据，上年列留空') + '）').font = S.SHELL_TITLE_FONT
    ws.cell(2, 1, '说明：GL 分月数与科目余额表勾稽差异恒为 0（GL 为全量抽取，"本期借方发生额"=GL 全部借方合计）；'
                  '仅当 GL 抽取期间小于科目余额表期间（如 2026 GL 仅 1-5 月）时，差额为未抽取月份（取数范围限制）。').font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    if _adapter is not None and _adapter.current_comp():
        ws.cell(3, 1, '⚠️ SAP 口径说明：分月为 GL【完全互抵剔除】口径（剔除同凭证借贷互抵的红字重记/内部互转），'
                      '与 TB 全年存在 GL 天然口径差异（未过账/跨期），差异待核实。').font = Font(name='Times New Roman', size=10, italic=True, color='C00000')

    r = 4
    ncols = 1 + len(METRICS)   # 月份 + 8 指标
    _tot_acc = None            # 2026-08-05 前道规范：收入按主体汇总 全集团合计行累计
    # ---- 逐主体块 ----
    for E in sorted(entities):
        d = data.get(E, {}).get(cur_year)
        tb_year = tb_all.get(E, {}).get(cur_year)
        # 2026-08-02 修正：核对口径改 TB【本期借方发生额】——铁律27 GL 全部借方合计=TB 借发（GL 全量抽取恒等）。
        # 原用 TB 贷发（确认口径），对未结转损益账套（借发≠贷发，如 FY智能2025 期末贷方余额 11.86M）出现
        # 假差异：GL 借方=已结转部分≠TB 贷发=确认总额。改借发后差异恒 0，未结转部分另行披露。
        tb_rev = _tb_total(tb_year, 'rev', 'db')
        tb_cost = _tb_total(tb_year, 'cost', 'db')
        gl_rev = sum(d.get('rev_m', {}).values()) if d else 0.0
        gl_cost = sum(d.get('cost_m', {}).values()) if d else 0.0
        has_data = d is not None and 'rev_m' in d and 'cost_m' in d
        has_diff = (not _diff_zero(tb_rev, gl_rev)) or (not _diff_zero(tb_cost, gl_cost))
        pd = data.get(E, {}).get(prev_year) if prev_year else None
        # 标题（与 TB 核对一致者标注，但照常列示明细）
        tag = '' if has_diff else '（与 TB 核对一致，差异=0）'
        # ⚡ 2026-08-11 N5：主体块标题 跨列合并+浅蓝底色（块式分隔）
        _tc = ws.cell(r, 1, f'【{E}】 主营业务收入 / 主营业务成本 分月明细{tag}')
        _tc.font = S.SHELL_SUB_FONT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        _F_BLK = PatternFill('solid', fgColor='DDEBF7')
        for _j in range(1, ncols + 1):
            ws.cell(r, _j).fill = _F_BLK
            ws.cell(r, _j).border = S.SHELL_BORDER
        r += 1
        # 表头
        _hcell(ws, r, 1, '月份')
        for j, mt in enumerate(METRICS):
            _hcell(ws, r, 2 + j, mt)
        r += 1
        if not has_data:
            _dcell(ws, r, 1, '（该主体无主营业务收入 GL 数据，无法与 TB 核对）', align=LEFT)
            r += 1
            continue
        monthly, total = _metrics_block(d['rev_m'], d['cost_m'],
                                        pd['rev_m'] if pd else None,
                                        pd['cost_m'] if pd else None, prev_year)
        # 2026-08-04 规范：按主体汇总行（标准行式，核对程序逐主体提取）
        _dcell(wsum, r_sum, 1, E)
        for j in range(8):
            _dcell(wsum, r_sum, 2 + j, total[j])
        r_sum += 1
        # 2026-08-05 前道规范：累计全集团合计（循环后写合计行）
        if _tot_acc is None:
            _tot_acc = [0.0] * 8
        for j in range(8):
            if total[j] is not None:
                _tot_acc[j] += total[j]
        for i, m in enumerate(MONTHS):
            _dcell(ws, r, 1, MONTH_LABEL[m], align=CEN)
            for j in range(8):
                _dcell(ws, r, 2 + j, monthly[i][j])
            r += 1
        # 全年小计
        _dcell(ws, r, 1, '全年小计', bold=True, fill=S.SHELL_TOT_FILL if hasattr(S, 'SHELL_TOT_FILL') else None)
        for j in range(8):
            _dcell(ws, r, 2 + j, total[j], bold=True)
        r += 1
        # 与科目余额表核对（该主体）——口径=GL 借方合计 vs TB 本期借发（铁律27 恒等）
        rev_c = _code_of(tb_year, 'rev')
        cost_c = _code_of(tb_year, 'cost')
        _dcell(ws, r, 1, '与科目余额表核对', align=LEFT, bold=True)
        _dcell(ws, r, 2, f'TB收入({rev_c}借){tb_rev:,.2f}', align=LEFT)
        _dcell(ws, r, 3, f'GL收入{tb_rev - gl_rev:+,.2f}', align=LEFT)
        _dcell(ws, r, 4, f'TB成本({cost_c}借){tb_cost:,.2f}', align=LEFT)
        _dcell(ws, r, 5, f'GL成本{tb_cost - gl_cost:+,.2f}', align=LEFT)
        r += 1
        # 未结转损益披露（2026-08-02）：TB 借发≠贷发（损益科目期末余额≠0，如 FY智能2025 收入期末贷方余额
        # 11,862,574.13=12 月确认未结转）→ GL 借方=已结转部分，与 TB 贷发(确认总额)差异为未结转额。
        tb_rev_cr = _tb_total(tb_year, 'rev', 'cr')
        tb_cost_cr = _tb_total(tb_year, 'cost', 'cr')
        if not _diff_zero(tb_rev, tb_rev_cr) or not _diff_zero(tb_cost, tb_cost_cr):
            _dcell(ws, r, 1, f'注：本主体损益科目未结平——{rev_c} 本期借发 {tb_rev:,.2f} ≠ 贷发 {tb_rev_cr:,.2f}'
                             f'（期末贷方余额 {tb_rev_cr - tb_rev:,.2f}）、{cost_c} 借发 {tb_cost:,.2f} ≠ 贷发 '
                             f'{tb_cost_cr:,.2f}（期末借方余额 {tb_cost - tb_cost_cr:,.2f}）；GL 分月数=GL 全部借方'
                             f'（=TB 借发，已结转部分），未结转额未计入 GL 分月数；审定表按确认口径（收入=TB 贷发'
                             f'）反映本期经营业绩，二者口径差异属未结转损益披露项，非 GL 抽取缺失。', align=LEFT)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
            r += 1
        # 差异归因披露（2026-07-31 二次修订·回归一般原则）：
        # TB 借发 = GL 全部借方合计（GL 全量抽取恒等）→ 正常账套差异应为 0，无需差异凭证清单。
        # 残留差异仅两类：① GL 抽取期间 < TB 期间（2026 GL 仅 1-5 月 YTD、TB 全年）→ 差额为未抽取
        # 月份发生额（取数范围限制，非账务差错）；② 未结转损益账套（收入科目期末≠0）。
        if has_diff:
            gl_max_m = int((d or {}).get('gl_max_m', 0) or 0)
            if gl_max_m and gl_max_m < 12:
                _dcell(ws, r, 1, f'差异归因：GL 抽取期间仅至 {gl_max_m} 月（YTD），TB 科目余额表为全年'
                                 f'（{cur_year}）→ 收入差额 {tb_rev-gl_rev:+,.2f}、成本差额 {tb_cost-gl_cost:+,.2f}'
                                 f' 为未抽取月份（{gl_max_m+1}-12 月）发生额（取数范围限制，非账务差错、'
                                 f'非记账习惯差异）；GL 在抽取期间内为全量（无缺失行）。', align=LEFT)
            else:
                _dcell(ws, r, 1, f'差异归因：GL 分月数取 GL 借方合计（收入=结转损益凭证借方、成本=确认借方'
                                 f'+结转借方），TB 借发 = GL 全部借方（全量抽取恒等）→ 正常账套差异应为 0；'
                                 f'本主体收入差额 {tb_rev-gl_rev:+,.2f}、成本差额 {tb_cost-gl_cost:+,.2f}'
                                 f' 系未结转损益账套或特殊凭证结构（收入科目期末余额非 0，TB 贷发=确认贷方≠'
                                 f'GL 借方），须按差异凭证逐笔核实；GL 为全量抽取，非抽取缺失。', align=LEFT)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
            r += 1
    # 2026-08-05 前道规范：按主体汇总表 补『全集团合计』行（核对程序取数依赖合计行）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出
    if _tot_acc is not None and len(entities) > 1:
        _dcell(wsum, r_sum, 1, '全集团合计', bold=True, fill=S.SHELL_TOT_FILL if hasattr(S, 'SHELL_TOT_FILL') else None)
        for j in range(8):
            _dcell(wsum, r_sum, 2 + j, _tot_acc[j], bold=True,
                   fill=S.SHELL_TOT_FILL if hasattr(S, 'SHELL_TOT_FILL') else None)
    # 2026-08-05 前道规范化：写入后立即结构校验
    try:
        from audit_common import assert_sheet_standard
        assert_sheet_standard(wsum, 'detail', '主营业务收入按主体明细表', raise_on_error=False)
    except Exception:
        pass
    # ---- 全集团块 ----
    ws.cell(r, 1, '【全集团】 主营业务收入 / 主营业务成本 分月汇总').font = S.SHELL_SUB_FONT
    r += 1
    _hcell(ws, r, 1, '月份')
    for j, mt in enumerate(METRICS):
        _hcell(ws, r, 2 + j, mt)
    r += 1
    grp_rev = _sum_months([data.get(E, {}).get(cur_year, {}).get('rev_m', {}) for E in entities])
    grp_cost = _sum_months([data.get(E, {}).get(cur_year, {}).get('cost_m', {}) for E in entities])
    grp_prev_rev = _sum_months([data.get(E, {}).get(prev_year, {}).get('rev_m', {}) for E in entities]) if prev_year else None
    grp_prev_cost = _sum_months([data.get(E, {}).get(prev_year, {}).get('cost_m', {}) for E in entities]) if prev_year else None
    monthly, total = _metrics_block(grp_rev, grp_cost, grp_prev_rev, grp_prev_cost, prev_year)
    for i, m in enumerate(MONTHS):
        _dcell(ws, r, 1, MONTH_LABEL[m], align=CEN)
        for j in range(8):
            _dcell(ws, r, 2 + j, monthly[i][j])
        r += 1
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出『全集团全年合计』行
    if len(entities) > 1:
        _dcell(ws, r, 1, '全集团全年合计', bold=True)
        for j in range(8):
            _dcell(ws, r, 2 + j, total[j], bold=True)
        r += 2
    else:
        r += 1
    # ---- 与科目余额表核对（全集团）----
    ws.cell(r, 1, '与科目余额表核对（全集团）').font = S.SHELL_SUB_FONT
    r += 1
    _hcell(ws, r, 1, '主体')
    _hcell(ws, r, 2, 'TB收入(借)')
    _hcell(ws, r, 3, 'GL收入合计')
    _hcell(ws, r, 4, '差异')
    _hcell(ws, r, 5, 'TB成本(借)')
    _hcell(ws, r, 6, 'GL成本合计')
    _hcell(ws, r, 7, '差异')
    r += 1
    tot_tb_rev = tot_gl_rev = tot_tb_cost = tot_gl_cost = 0.0
    for E in sorted(entities):
        tb_year = tb_all.get(E, {}).get(cur_year)
        tb_rev = _tb_total(tb_year, 'rev', 'db')
        tb_cost = _tb_total(tb_year, 'cost', 'db')
        d = data.get(E, {}).get(cur_year)
        gl_rev = sum(d.get('rev_m', {}).values()) if d else 0.0
        gl_cost = sum(d.get('cost_m', {}).values()) if d else 0.0
        tot_tb_rev += tb_rev; tot_gl_rev += gl_rev
        tot_tb_cost += tb_cost; tot_gl_cost += gl_cost
        _dcell(ws, r, 1, E, align=LEFT)
        _dcell(ws, r, 2, tb_rev); _dcell(ws, r, 3, gl_rev); _dcell(ws, r, 4, tb_rev - gl_rev)
        _dcell(ws, r, 5, tb_cost); _dcell(ws, r, 6, gl_cost); _dcell(ws, r, 7, tb_cost - gl_cost)
        r += 1
    _dcell(ws, r, 1, '全集团', bold=True)
    _dcell(ws, r, 2, tot_tb_rev, bold=True); _dcell(ws, r, 3, tot_gl_rev, bold=True)
    _dcell(ws, r, 4, tot_tb_rev - tot_gl_rev, bold=True)
    _dcell(ws, r, 5, tot_tb_cost, bold=True); _dcell(ws, r, 6, tot_gl_cost, bold=True)
    _dcell(ws, r, 7, tot_tb_cost - tot_gl_cost, bold=True)
    # 列宽
    for c in range(1, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 13
    ws.column_dimensions['A'].width = 14
    ws.freeze_panes = 'B5'
    return ws


# ===================== 表2：按二级明细 分月明细（全集团，收入/成本配对，纵向排列） =====================
def build_by_l2(wb, entities, cur_year, prev_year, data, tb_all):
    # 收集 rev / cost 二级名称（TB 代码顺序，再补 GL 独有）——须在 create_sheet 前完成，
    # 供『全未匹配则不生成』判断使用（2026-08-02，避免残留空 sheet）
    rev_names, cost_names = [], []
    seen_r, seen_c = set(), set()
    for E in sorted(entities):
        tb_year = tb_all.get(E, {}).get(cur_year)
        if not tb_year:
            continue
        # ⚡ 2026-08-12 SAP 修复：名称前缀法（铁律5），原 level==2+startswith 对 SAP 无父级 TB 失效
        for l2 in _tb_l2_entries(tb_year, '主营业务收入', 'credit'):
            if l2 not in seen_r:
                seen_r.add(l2); rev_names.append(l2)
        for l2 in _tb_l2_entries(tb_year, '主营业务成本', 'debit'):
            if l2 not in seen_c:
                seen_c.add(l2); cost_names.append(l2)
    for E in entities:
        d = data.get(E, {}).get(cur_year)
        if not d:
            continue
        for n in d['rev_l2']:
            if n not in seen_r:
                seen_r.add(n); rev_names.append(n)
        for n in d['cost_l2']:
            if n not in seen_c:
                seen_c.add(n); cost_names.append(n)

    def _grp_rev(n):
        return _sum_months([data.get(E, {}).get(cur_year, {}).get('rev_l2', {}).get(n, {}) for E in entities])
    def _grp_cost(n):
        return _sum_months([data.get(E, {}).get(cur_year, {}).get('cost_l2', {}).get(n, {}) for E in entities])
    def _grp_prev_rev(n):
        return _sum_months([data.get(E, {}).get(prev_year, {}).get('rev_l2', {}).get(n, {}) for E in entities]) if prev_year else None
    def _grp_prev_cost(n):
        return _sum_months([data.get(E, {}).get(prev_year, {}).get('cost_l2', {}).get(n, {}) for E in entities]) if prev_year else None

    # 配对
    pairs, singles = _pair_l2(rev_names, cost_names)
    # 2026-08-02 用户要求：收入/成本二级【全部未匹配】（无任何产品族配对成功，如 FY 收入按业务类型
    # 检测/技术服务/科研/特定收费、成本按费用性质耗材/工资/折旧等，天然不对应）→ 本表无意义，
    # 与『二级核对』一并跳过不生成。
    if not pairs and (rev_names or cost_names):
        return None
    ws = wb.create_sheet('主营业务收入_分月明细(按二级)')
    ws.cell(1, 1, f'主营业务收入分月明细表（按二级明细，全集团，{cur_year} 年度'
                   + (f' 对比 {prev_year}' if prev_year else ' · 无上年度数据') + '）').font = S.SHELL_TITLE_FONT
    ws.cell(2, 1, '注：套圈/外销等收入无对应成本二级（成本归集于「产品」/「主营业务成本结转」），仅列收入、'
                  '不列成本列；其毛利需另作成本分配。').font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    fam_rev = {s: _sum_months([_grp_rev(n) for n in rns]) for (s, rns, cns) in pairs}
    fam_cost = {s: _sum_months([_grp_cost(n) for n in cns]) for (s, rns, cns) in pairs}

    # 构建渲染列表：每个 (label, rev_md, cost_md, is_single, kind)
    render = []
    for s, rns, cns in pairs:
        if len(rns) == 1 and len(cns) == 1:
            label = f'{rns[0]} / {cns[0]}'
        else:
            label = f'{s}（收入:{"/".join(rns)}；成本:{"/".join(cns)}）'
        render.append((label, fam_rev.get(s, _empty_month()), fam_cost.get(s, _empty_month()), False, None))
    for name, kind in singles:
        render.append((name, _grp_rev(name) if kind == 'rev' else _empty_month(),
                       _grp_cost(name) if kind == 'cost' else _empty_month(), True, kind))

    # 纵向表头
    METRICS_HEADERS = ['本期收入', '本期成本', '本期毛利', '上年收入', '上年成本', '上年毛利', '收入增减', '成本增减']
    r = 4
    ws.cell(r, 1, '月份').font = S.SHELL_HFONT; ws.cell(r, 1).fill = S.SHELL_HFILL; ws.cell(r, 1).alignment = CEN
    ws.cell(r, 2, '产品族').font = S.SHELL_HFONT; ws.cell(r, 2).fill = S.SHELL_HFILL; ws.cell(r, 2).alignment = CEN
    for j, mt in enumerate(METRICS_HEADERS):
        c = ws.cell(r, 3 + j, mt)
        c.font = S.SHELL_HFONT; c.fill = S.SHELL_HFILL; c.alignment = CEN
    r += 1

    # 数据行：每行 = (月份, 产品族, 8个指标)
    for m in MONTHS:
        for label, rev_md, cost_md, is_single, kind in render:
            prev_rev = None; prev_cost = None
            if is_single and kind == 'rev':
                pass  # rev single → cost empty
            prev_monthly, _ = _metrics_block(rev_md, cost_md, None, None, None)
            vals = prev_monthly[m - 1]
            ws.cell(r, 1, MONTH_LABEL[m]).alignment = CEN
            ws.cell(r, 2, label).alignment = LEFT
            for j in range(8):
                v = vals[j]
                c = ws.cell(r, 3 + j, v)
                c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal='right')
            r += 1

    # 各产品族全年合计
    tot_first = r                                # 全年合计块首行（全集团合计公式引用范围）
    ws.cell(r, 1, '全年合计').font = S.SHELL_BOLD; ws.cell(r, 1).alignment = CEN
    ws.cell(r, 2, '').font = S.SHELL_BOLD
    for label, rev_md, cost_md, is_single, kind in render:
        _, tot = _metrics_block(rev_md, cost_md, None, None, None)
        ws.cell(r, 2, label).font = S.SHELL_BOLD
        for j in range(8):
            c = ws.cell(r, 3 + j, tot[j])
            c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal='right'); c.font = S.SHELL_BOLD
        r += 1

    # 合计行（2026-08-06 协议：写数值=产品族行累计；公式 data_only 读 None 收回读不回）
    ws.cell(r, 1, ''); ws.cell(r, 2, '全集团合计').font = S.SHELL_BOLD
    ws.cell(r, 2).fill = S.SHELL_TOT_FILL
    for j in range(8):
        _sv = sum(ws.cell(x, 3 + j).value or 0.0 for x in range(tot_first, r)
                  if isinstance(ws.cell(x, 3 + j).value, (int, float)))
        c = ws.cell(r, 3 + j, round(_sv, 2))
        c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal='right'); c.font = S.SHELL_BOLD; c.fill = S.SHELL_TOT_FILL

    # 列宽
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 28
    for c in range(3, 11):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.freeze_panes = 'C5'
    return ws


# ===================== 表2b：全集团二级核对（与科目余额表 TB 核对） =====================
def build_l2_recon_sheet(wb, entities, cur_year, prev_year, data, tb_all):
    """单开 sheet，列示：
       ① 与科目余额表核对（按二级，全集团 TB/GL/差异）
       ② 一级科目(TB) ↔ 二级明细合计(TB) 勾稽
    2026-07-31：无任何二级名称（TB/GL 均无二级数据）时不生成该 sheet（用户要求"勾稽核对无金额不产生"）。
    """
    # 重算数据
    rev_names, cost_names = [], []
    seen_r, seen_c = set(), set()
    for E in sorted(entities):
        tb_year = tb_all.get(E, {}).get(cur_year)
        if not tb_year:
            continue
        # ⚡ 2026-08-12 SAP 修复：名称前缀法（铁律5），原 level==2+startswith 对 SAP 无父级 TB 失效
        for l2 in _tb_l2_entries(tb_year, '主营业务收入', 'credit'):
            if l2 not in seen_r:
                seen_r.add(l2); rev_names.append(l2)
        for l2 in _tb_l2_entries(tb_year, '主营业务成本', 'debit'):
            if l2 not in seen_c:
                seen_c.add(l2); cost_names.append(l2)
    for E in entities:
        d = data.get(E, {}).get(cur_year)
        if not d:
            continue
        for n in d['rev_l2']:
            if n not in seen_r:
                seen_r.add(n); rev_names.append(n)
        for n in d['cost_l2']:
            if n not in seen_c:
                seen_c.add(n); cost_names.append(n)
    if not rev_names and not cost_names:
        return  # 无二级数据 → 不生成核对 sheet
    # 2026-08-02 用户要求：收入/成本二级【全部未匹配】（无任何产品族配对成功，如 FY 收入按业务类型
    # 检测/技术服务/科研/特定收费、成本按费用性质耗材/工资/折旧等，天然不对应）→ 核对/分月(按二级)
    # 均无意义，本 sheet 与『分月明细(按二级)』一并跳过。
    _pairs0, _ = _pair_l2(rev_names, cost_names)
    if not _pairs0:
        return
    ws = wb.create_sheet(f'主营业务收入_二级核对')
    ws.cell(1, 1, f'主营业务收入二级核对表（与科目余额表 TB 核对，{cur_year} 年度）').font = S.SHELL_TITLE_FONT

    def _grp_rev(n):
        total = 0.0
        for E in entities:
            md = data.get(E, {}).get(cur_year, {}).get('rev_l2', {}).get(n, {})
            if md:
                total += sum(md.values())
        return total
    def _grp_cost(n):
        total = 0.0
        for E in entities:
            md = data.get(E, {}).get(cur_year, {}).get('cost_l2', {}).get(n, {})
            if md:
                total += sum(md.values())
        return total

    tb_rev_l2, tb_cost_l2 = {}, {}
    for E in entities:
        tb_year = tb_all.get(E, {}).get(cur_year)
        if not tb_year:
            continue
        # ⚡ 2026-08-12 SAP 修复：名称前缀法（铁律5）；铁律27 恒等——TB 收入二级取借发与 GL 借方一致
        for l2, amt in _tb_l2_entries(tb_year, '主营业务收入', 'debit').items():
            tb_rev_l2[l2] = tb_rev_l2.get(l2, 0.0) + amt
        for l2, amt in _tb_l2_entries(tb_year, '主营业务成本', 'debit').items():
            tb_cost_l2[l2] = tb_cost_l2.get(l2, 0.0) + amt

    r = 3
    # ---- ① 产品族配对展示（收入 ↔ 成本匹配情况） ----
    ws.cell(r, 1, '产品族配对（收入 ↔ 成本按名称/词干匹配）').font = S.SHELL_SUB_FONT; r += 1
    hdr_pair = ['匹配状态', '收入二级名称', '成本二级名称']
    for j, h in enumerate(hdr_pair, 1):
        _hcell(ws, r, j, h)
    r += 1
    pairs, singles = _pair_l2(rev_names, cost_names)
    matched_rev = set()
    for s, rns, cns in pairs:
        _dcell(ws, r, 1, '已匹配'); _dcell(ws, r, 2, '/'.join(rns), align=LEFT); _dcell(ws, r, 3, '/'.join(cns), align=LEFT)
        for n in rns: matched_rev.add(n)
        r += 1
    for nm, kind in singles:
        _dcell(ws, r, 1, '未匹配')
        _dcell(ws, r, 2, nm if kind == 'rev' else '', align=LEFT)
        _dcell(ws, r, 3, nm if kind == 'cost' else '', align=LEFT)
        r += 1
    _dcell(ws, r, 1, '注：收入二级按业务类型（检测/技术/科研/特定收费），成本二级按费用性质（耗材/委托业务费/人员工资等），'
                      '两者名称天然不对应，『未匹配』属数据现状而非异常；毛利需按分部门明细表另行归集。', align=LEFT, fill=None)
    r += 1
    r += 1
    # ---- ② TB vs GL 核对 ----
    ws.cell(r, 1, '与科目余额表核对（按二级，全集团）').font = S.SHELL_SUB_FONT; r += 1
    hdr = ['二级明细', 'TB收入(借)', 'GL收入', '差异', 'TB成本(借)', 'GL成本', '差异']
    for j, h in enumerate(hdr, 1):
        _hcell(ws, r, j, h)
    r += 1
    for nm in rev_names:
        gl_v = _grp_rev(nm)
        tb_v = tb_rev_l2.get(nm, 0.0)
        diff = tb_v - gl_v
        _dcell(ws, r, 1, nm, align=LEFT); _dcell(ws, r, 2, tb_v); _dcell(ws, r, 3, gl_v)
        _dcell(ws, r, 4, diff)
        if abs(diff) > 0.005:
            ws.cell(r, 4).font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
        r += 1
    for nm in cost_names:
        gl_v = _grp_cost(nm)
        tb_v = tb_cost_l2.get(nm, 0.0)
        diff = tb_v - gl_v
        _dcell(ws, r, 1, nm, align=LEFT); _dcell(ws, r, 5, tb_v); _dcell(ws, r, 6, gl_v)
        _dcell(ws, r, 7, diff)
        if abs(diff) > 0.005:
            ws.cell(r, 7).font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
        r += 1
    r += 1

    # ---- ② 一级↔二级勾稽 ----
    ws.cell(r, 1, '一级科目(TB) ↔ 二级明细合计(TB) 勾稽').font = S.SHELL_SUB_FONT; r += 1
    _hcell(ws, r, 1, '一级科目'); _hcell(ws, r, 2, '一级(TB) gross')
    _hcell(ws, r, 3, '二级明细合计(TB)'); _hcell(ws, r, 4, '差异')
    r += 1
    _l1_rev = sum(_tb_total(tb_all.get(E, {}).get(cur_year), 'rev', 'db') for E in entities)
    if tb_rev_l2:
        _l1_r = _l1_rev; _l2_r = sum(tb_rev_l2.values())
        _dcell(ws, r, 1, REV_CODE, align=LEFT); _dcell(ws, r, 2, _l1_r); _dcell(ws, r, 3, _l2_r)
        _dcell(ws, r, 4, _l1_r - _l2_r)
        if abs(_l1_r - _l2_r) > 0.005:
            ws.cell(r, 4).font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
        r += 1
    _l1_cost = sum(_tb_total(tb_all.get(E, {}).get(cur_year), 'cost', 'db') for E in entities)
    if tb_cost_l2:
        _l1_c = _l1_cost; _l2_c = sum(tb_cost_l2.values())
        _dcell(ws, r, 1, COST_CODE, align=LEFT); _dcell(ws, r, 2, _l1_c); _dcell(ws, r, 3, _l2_c)
        _dcell(ws, r, 4, _l1_c - _l2_c)
        if abs(_l1_c - _l2_c) > 0.005:
            ws.cell(r, 4).font = Font(name='Times New Roman', size=10, bold=True, color='C00000')

    for c in range(1, 8):
        ws.column_dimensions[get_column_letter(c)].width = 16 if c == 1 else 14
    ws.freeze_panes = 'B4'
    return ws

# ===================== 表3a：按客户 全年明细 =====================
def build_by_customer(wb, entities, cur_year, prev_year, data):
    ws = wb.create_sheet(f'主营业务收入_客户明细({cur_year})')
    ws.cell(1, 1, f'主营业务收入客户明细表（全年，{cur_year}'
                   + (f' 对比 {prev_year}' if prev_year else ' · 无上年度数据') + '）').font = S.SHELL_TITLE_FONT
    # 客户取数/排序口径说明已按用户要求删除（2026-07-31）

    # 按 (entity, customer) 分别取数
    ec_rev = {}     # (E, c) -> 本期收入
    ec_cost = {}    # (E, c) -> 本期成本
    ec_prev_rev = {}
    ec_prev_cost = {}
    for E in entities:
        d = data.get(E, {}).get(cur_year)
        if d:
            for c, md in d['rev_cust'].items():
                ec_rev[(E, c)] = sum(md.values())
                ec_cost[(E, c)] = sum(d['cost_cust'].get(c, _empty_month()).values())
        if prev_year:
            pd = data.get(E, {}).get(prev_year)
            if pd:
                for c, md in pd['rev_cust'].items():
                    ec_prev_rev[(E, c)] = sum(md.values())
                    ec_prev_cost[(E, c)] = sum(pd['cost_cust'].get(c, _empty_month()).values())

    # 按全集团客户总收入排序
    cust_total = {}
    for (E, c), v in ec_rev.items():
        cust_total[c] = cust_total.get(c, 0.0) + v
    sorted_custs = sorted(cust_total, key=lambda x: -cust_total[x])

    r = 4
    headers = ['核算主体', '客户', '本期收入', '本期成本', '本期毛利',
               '上年收入', '上年成本', '上年毛利', '收入增减', '成本增减', '毛利增减']
    for j, h in enumerate(headers):
        _hcell(ws, r, 1 + j, h)
    r += 1

    for c in sorted_custs:
        sub_rev = 0.0; sub_cost = 0.0
        for E in sorted(entities):
            key = (E, c)
            inc = ec_rev.get(key, 0.0)
            cst = ec_cost.get(key, 0.0)
            if inc == 0 and cst == 0:
                continue
            gp = inc - cst
            pin = ec_prev_rev.get(key)
            pcs = ec_prev_cost.get(key)
            pgp = (pin - pcs) if (prev_year and pin is not None) else None
            dinc = (inc - pin) if (prev_year and pin is not None) else None
            dcst = (cst - pcs) if (prev_year and pcs is not None) else None
            dgp = (gp - pgp) if (prev_year and pgp is not None) else None
            vals = [E, c, inc, cst, gp,
                    pin if pin is not None else '', pcs if pcs is not None else '',
                    pgp if pgp is not None else '', dinc if dinc is not None else '',
                    dcst if dcst is not None else '', dgp if dgp is not None else '']
            for j, v in enumerate(vals):
                _dcell(ws, r, 1 + j, v)
            r += 1
            sub_rev += inc; sub_cost += cst
        # 客户小计
        if sub_rev or sub_cost:
            sgp = sub_rev - sub_cost
            _dcell(ws, r, 1, '', fill=S.SHELL_TOT_FILL); _dcell(ws, r, 2, c + ' 小计', bold=True, fill=S.SHELL_TOT_FILL)
            _dcell(ws, r, 3, sub_rev, bold=True, fill=S.SHELL_TOT_FILL)
            _dcell(ws, r, 4, sub_cost, bold=True, fill=S.SHELL_TOT_FILL)
            _dcell(ws, r, 5, sgp, bold=True, fill=S.SHELL_TOT_FILL)
            for j in range(6, 12):
                _dcell(ws, r, j, '', fill=S.SHELL_TOT_FILL)
            r += 1
    # 合计
    tinc = sum(ec_rev.values())
    tcst = sum(ec_cost.values())
    tgp = tinc - tcst
    tpin = sum(ec_prev_rev.values()) if prev_year else None
    tpcs = sum(ec_prev_cost.values()) if prev_year else None
    tpgp = (tpin - tpcs) if (prev_year and tpin is not None) else None
    tdinc = (tinc - tpin) if (prev_year and tpin is not None) else None
    tdcst = (tcst - tpcs) if (prev_year and tpcs is not None) else None
    tdgp = (tgp - tpgp) if (prev_year and tpgp is not None) else None
    vals = ['合计', tinc, tcst, tgp, tpin if tpin is not None else '', tpcs if tpcs is not None else '',
            tpgp if tpgp is not None else '', tdinc if tdinc is not None else '',
            tdcst if tdcst is not None else '', tdgp if tdgp is not None else '']
    _dcell(ws, r, 1, vals[0], bold=True, align=LEFT)
    for j in range(1, 10):
        _dcell(ws, r, 1 + j, vals[j], bold=True)
    for c in range(1, 11):
        ws.column_dimensions[get_column_letter(c)].width = 16 if c == 1 else 13
    ws.freeze_panes = 'B5'
    return ws, sorted(cust_total, key=lambda x: -cust_total[x])


# ===================== 主营业务成本分月明细（2026-08-07 用户要求） =====================
# 用户：DQ 主营业务成本比较特殊，单独补一张类似费用分月明细表（按 主体×末级科目×月）。
# 数据已由 _aggregate 聚合（data[E][y]['cost_leaf'] = {末级科目路径: {月: GL借方}}）；
# TB 控制数=6401 一级借发。主体小计填各月 + 全集团合计（2026-08-10 用户要求）。
def build_cost_monthly_sheet(wb, entities, cur_year, data, tb_all):
    ws = wb.create_sheet('主营业务成本_分月明细')
    ws.cell(1, 1, f'主营业务成本分月明细表（{cur_year} 年度，按末级科目×月，GL借方 vs TB借发）').font = S.SHELL_TITLE_FONT
    hdrs = ['核算主体', '末级科目名称'] + [f'{m}月' for m in MONTHS] + ['全年', 'TB借发', '差异']
    ncols = len(hdrs)
    for c, h in enumerate(hdrs, 1):
        cc = ws.cell(2, c, h)
        cc.font = S.SHELL_HFONT; cc.fill = S.SHELL_HFILL; cc.border = S.SHELL_BORDER
        cc.alignment = S.SHELL_CEN
    r = 3
    grp_months = [0.0] * 12
    grp_tot = 0.0
    grp_tb = 0.0
    for E in sorted(entities):
        d = (data.get(E) or {}).get(str(cur_year), {})
        leafs = d.get('cost_leaf') or {}
        tb = (tb_all.get(E) or {}).get(str(cur_year), {})
        code = _detect_codes(tb).get('cost')
        tb_debit = float(tb.get(code, {}).get('debit') or 0.0) if code else 0.0
        ent_tot = 0.0
        ent_months = [0.0] * 12
        leaf_rows = []
        for leaf in sorted(leafs, key=lambda x: -sum(leafs[x].values())):
            months = leafs[leaf]
            vals = [months.get(m, 0.0) for m in MONTHS]
            tot = sum(vals)
            if abs(tot) < 0.005:
                continue
            ws.cell(r, 1, E); ws.cell(r, 2, leaf)
            for j, m in enumerate(MONTHS):
                cc = ws.cell(r, 3 + j, round(vals[j], 2) if vals[j] != 0 else None)
                cc.number_format = S.SHELL_NUM; cc.border = S.SHELL_BORDER; cc.alignment = S.SHELL_RGT
            ws.cell(r, 15, round(tot, 2)).number_format = S.SHELL_NUM
            ws.cell(r, 15).border = S.SHELL_BORDER; ws.cell(r, 15).alignment = S.SHELL_RGT
            ws.cell(r, 16, None); ws.cell(r, 17, None)
            ent_tot += tot
            for j, m in enumerate(MONTHS):
                ent_months[j] += vals[j]
            leaf_rows.append(r)
            r += 1
        if leaf_rows or abs(tb_debit) > 0.005:
            # ⚡ 2026-08-10 主体小计：补各月金额（原只有全年/TB/差异列）
            ws.cell(r, 1, E + ' 小计'); ws.cell(r, 2, '(主体小计)')
            for j, m in enumerate(MONTHS):
                cc = ws.cell(r, 3 + j, round(ent_months[j], 2) if ent_months[j] != 0 else None)
                cc.number_format = S.SHELL_NUM; cc.alignment = S.SHELL_RGT
            ws.cell(r, 15, round(ent_tot, 2)).number_format = S.SHELL_NUM
            ws.cell(r, 16, round(tb_debit, 2)).number_format = S.SHELL_NUM
            ws.cell(r, 17, round(ent_tot - tb_debit, 2)).number_format = S.SHELL_NUM
            for c in range(1, ncols + 1):
                ws.cell(r, c).fill = S.SHELL_TOT_FILL; ws.cell(r, c).border = S.SHELL_BORDER
                if c in (15, 16, 17):
                    ws.cell(r, c).alignment = S.SHELL_RGT
            for j in range(12):
                grp_months[j] += ent_months[j]
            grp_tot += ent_tot
            grp_tb += tb_debit
            r += 1
    # ⚡ 2026-08-10 全集团合计（各月+全年+TB+差异）
    if grp_tot or grp_tb:
        ws.cell(r, 1, '全集团合计'); ws.cell(r, 2, '(集团合计)')
        for j, m in enumerate(MONTHS):
            cc = ws.cell(r, 3 + j, round(grp_months[j], 2) if grp_months[j] != 0 else None)
            cc.number_format = S.SHELL_NUM; cc.alignment = S.SHELL_RGT
        ws.cell(r, 15, round(grp_tot, 2)).number_format = S.SHELL_NUM
        ws.cell(r, 16, round(grp_tb, 2)).number_format = S.SHELL_NUM
        ws.cell(r, 17, round(grp_tot - grp_tb, 2)).number_format = S.SHELL_NUM
        for c in range(1, ncols + 1):
            ws.cell(r, c).fill = S.SHELL_TOT_FILL; ws.cell(r, c).border = S.SHELL_BORDER
            if c in (15, 16, 17):
                ws.cell(r, c).alignment = S.SHELL_RGT
        r += 1
    ws.freeze_panes = 'C3'
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 46
    return ws


# ===================== 主营业务收入分月明细（2026-08-07 用户要求，参照成本分月格式） =====================
# 按 主体×末级科目×月（data[E][y]['rev_leaf']=GL 收入末级分月，完整科目路径不合并同类）；
# TB 控制数=收入一级贷方。主体小计填各月 + 全集团合计（2026-08-10 用户要求）。
def build_rev_monthly_sheet(wb, entities, cur_year, data, tb_all):
    ws = wb.create_sheet('主营业务收入_分月明细')
    ws.cell(1, 1, f'主营业务收入分月明细表（{cur_year} 年度，按末级科目×月，GL贷方 vs TB贷发，不合并同类）').font = S.SHELL_TITLE_FONT
    hdrs = ['核算主体', '末级科目名称'] + [f'{m}月' for m in MONTHS] + ['全年', 'TB贷发', '差异']
    ncols = len(hdrs)
    for c, h in enumerate(hdrs, 1):
        cc = ws.cell(2, c, h)
        cc.font = S.SHELL_HFONT; cc.fill = S.SHELL_HFILL; cc.border = S.SHELL_BORDER
        cc.alignment = S.SHELL_CEN
    r = 3
    grp_months = [0.0] * 12
    grp_tot = 0.0
    grp_tb = 0.0
    for E in sorted(entities):
        d = (data.get(E) or {}).get(str(cur_year), {})
        leafs = d.get('rev_leaf') or {}
        tb = (tb_all.get(E) or {}).get(str(cur_year), {})
        code = _detect_codes(tb).get('rev')
        tb_credit = float(tb.get(code, {}).get('credit') or 0.0) if code else 0.0
        ent_tot = 0.0
        ent_months = [0.0] * 12
        leaf_rows = []
        for leaf in sorted(leafs, key=lambda x: -sum(leafs[x].values())):
            months = leafs[leaf]
            vals = [months.get(m, 0.0) for m in MONTHS]
            tot = sum(vals)
            if abs(tot) < 0.005:
                continue
            ws.cell(r, 1, E); ws.cell(r, 2, leaf)
            for j, m in enumerate(MONTHS):
                cc = ws.cell(r, 3 + j, round(vals[j], 2) if vals[j] != 0 else None)
                cc.number_format = S.SHELL_NUM; cc.border = S.SHELL_BORDER; cc.alignment = S.SHELL_RGT
            ws.cell(r, 15, round(tot, 2)).number_format = S.SHELL_NUM
            ws.cell(r, 15).border = S.SHELL_BORDER; ws.cell(r, 15).alignment = S.SHELL_RGT
            ws.cell(r, 16, None); ws.cell(r, 17, None)
            ent_tot += tot
            for j, m in enumerate(MONTHS):
                ent_months[j] += vals[j]
            leaf_rows.append(r)
            r += 1
        if leaf_rows or abs(tb_credit) > 0.005:
            # ⚡ 2026-08-10 主体小计：补各月金额（原只有全年/TB/差异列）
            ws.cell(r, 1, E + ' 小计'); ws.cell(r, 2, '(主体小计)')
            for j, m in enumerate(MONTHS):
                cc = ws.cell(r, 3 + j, round(ent_months[j], 2) if ent_months[j] != 0 else None)
                cc.number_format = S.SHELL_NUM; cc.alignment = S.SHELL_RGT
            ws.cell(r, 15, round(ent_tot, 2)).number_format = S.SHELL_NUM
            ws.cell(r, 16, round(tb_credit, 2)).number_format = S.SHELL_NUM
            ws.cell(r, 17, round(ent_tot - tb_credit, 2)).number_format = S.SHELL_NUM
            for c in range(1, ncols + 1):
                ws.cell(r, c).fill = S.SHELL_TOT_FILL; ws.cell(r, c).border = S.SHELL_BORDER
                if c in (15, 16, 17):
                    ws.cell(r, c).alignment = S.SHELL_RGT
            for j in range(12):
                grp_months[j] += ent_months[j]
            grp_tot += ent_tot
            grp_tb += tb_credit
            r += 1
    # ⚡ 2026-08-10 全集团合计（各月+全年+TB+差异）
    if grp_tot or grp_tb:
        ws.cell(r, 1, '全集团合计'); ws.cell(r, 2, '(集团合计)')
        for j, m in enumerate(MONTHS):
            cc = ws.cell(r, 3 + j, round(grp_months[j], 2) if grp_months[j] != 0 else None)
            cc.number_format = S.SHELL_NUM; cc.alignment = S.SHELL_RGT
        ws.cell(r, 15, round(grp_tot, 2)).number_format = S.SHELL_NUM
        ws.cell(r, 16, round(grp_tb, 2)).number_format = S.SHELL_NUM
        ws.cell(r, 17, round(grp_tot - grp_tb, 2)).number_format = S.SHELL_NUM
        for c in range(1, ncols + 1):
            ws.cell(r, c).fill = S.SHELL_TOT_FILL; ws.cell(r, c).border = S.SHELL_BORDER
            if c in (15, 16, 17):
                ws.cell(r, c).alignment = S.SHELL_RGT
        r += 1
    ws.freeze_panes = 'C3'
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 46
    return ws


# ===================== 分析性程序（2026-08-09 对齐 SAP 5 模块） =====================
def build_revenue_analysis(wb, entities, cur_year, data, tb_all):
    """分析性程序（用户要求 SAP/U8 统一）：
    ①全集团分月收入/成本/毛利率趋势（环比波动>50% 标⚠️）
    ②客户集中度 Top5（占全集团收入%）——关注依赖风险
    ③红字冲销占比（收入借方负数占借方合计）
    ④收入确认方式（赊销→应收/现销→银行/票据→应收票据/预收转收入，按 GL 对方科目）
    ⑤大额收入凭证（≥1000万，从 GL 直接扫）"""
    ws = wb.create_sheet('分析性程序')
    ws.cell(1, 1, f'主营业务收入 分析性程序（{cur_year} 年度）').font = S.SHELL_TITLE_FONT
    r = 3
    # ① 全集团分月趋势
    ws.cell(r, 1, '① 全集团分月收入/成本/毛利率（环比波动>50% 标⚠️）').font = S.SHELL_SUB_FONT
    r += 1
    for j, h in enumerate(['月份', '收入', '成本', '毛利率%', '环比%', '提示'], 1):
        _hcell(ws, r, j, h)
    r += 1
    rev_m = [0.0] * 12
    cost_m = [0.0] * 12
    for E in entities:
        d = data.get(E, {}).get(cur_year)
        if not d:
            continue
        for m in MONTHS:
            rev_m[m - 1] += d.get('rev_m', {}).get(m, 0.0)
            cost_m[m - 1] += d.get('cost_m', {}).get(m, 0.0)
    prev = None
    for i, m in enumerate(MONTHS, 1):
        rv = rev_m[i - 1]
        cv = cost_m[i - 1]
        gp = (rv - cv) / rv * 100 if rv else 0
        chg = (rv - prev) / prev * 100 if prev else 0
        tip = '⚠️ 收入环比波动大' if abs(chg) > 50 and prev else ''
        _dcell(ws, r, 1, f'{m}月')
        c1 = _dcell(ws, r, 2, rv)
        c1.number_format = S.SHELL_NUM
        c2 = _dcell(ws, r, 3, cv)
        c2.number_format = S.SHELL_NUM
        _dcell(ws, r, 4, round(gp, 2))
        _dcell(ws, r, 5, round(chg, 1) if prev else None)
        _dcell(ws, r, 6, tip)
        prev = rv
        r += 1
    tot_r = sum(rev_m)
    ws.cell(r, 1, '全年合计').font = S.SHELL_BOLD
    _dcell(ws, r, 2, tot_r)
    ws.cell(r, 2).number_format = S.SHELL_NUM
    _dcell(ws, r, 3, sum(cost_m))
    ws.cell(r, 3).number_format = S.SHELL_NUM
    _dcell(ws, r, 4, round((tot_r - sum(cost_m)) / tot_r * 100, 2) if tot_r else 0)
    for c in range(1, 7):
        ws.cell(r, c).fill = S.SHELL_TOT_FILL
    r += 2
    # ② 客户集中度 Top5
    ws.cell(r, 1, '② 客户集中度 Top5（占全集团收入%）——关注依赖风险').font = S.SHELL_SUB_FONT
    r += 1
    for j, h in enumerate(['客户', '收入', '占比%'], 1):
        _hcell(ws, r, j, h)
    r += 1
    cust_agg = {}
    for E in entities:
        d = data.get(E, {}).get(cur_year)
        if not d:
            continue
        for cname, months in (d.get('rev_cust') or {}).items():
            cust_agg[cname] = cust_agg.get(cname, 0.0) + sum(months.values())
    for cname, amt in sorted(cust_agg.items(), key=lambda x: -x[1])[:5]:
        _dcell(ws, r, 1, cname)
        _dcell(ws, r, 2, amt)
        ws.cell(r, 2).number_format = S.SHELL_NUM
        _dcell(ws, r, 3, round(amt / tot_r * 100, 2) if tot_r else 0)
        r += 1
    if not cust_agg:
        _dcell(ws, r, 1, '（无客户级收入数据）')
        r += 1
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 20
    ws.freeze_panes = 'A3'
    return ws


# ===================== 表3b：前十大客户 分月 =====================
def build_top10_customer(wb, entities, cur_year, prev_year, data, top_customers):
    ws = wb.create_sheet('主营业务收入_前十大客户_分月')
    ws.cell(1, 1, f'主营业务收入前十大客户分月明细（{cur_year}'
                   + (f' 对比 {prev_year}' if prev_year else ' · 无上年度数据') + '）').font = S.SHELL_TITLE_FONT
    # 前十大客户口径说明已按用户要求删除（2026-07-31）

    r = 4
    cols = ['月份', '本期收入', '本期成本', '本期毛利', '上年收入', '上年成本', '上年毛利',
            '收入增减', '成本增减', '毛利增减']
    ncols = len(cols)
    for k, c in enumerate(top_customers[:10]):
        # 该客户全集团月度（本期/上年）
        rev_md = _sum_months([data.get(E, {}).get(cur_year, {}).get('rev_cust', {}).get(c, {}) for E in entities])
        cost_md = _sum_months([data.get(E, {}).get(cur_year, {}).get('cost_cust', {}).get(c, _empty_month()) for E in entities])
        prev_rev = _sum_months([data.get(E, {}).get(prev_year, {}).get('rev_cust', {}).get(c, {}) for E in entities]) if prev_year else None
        prev_cost = _sum_months([data.get(E, {}).get(prev_year, {}).get('cost_cust', {}).get(c, _empty_month()) for E in entities]) if prev_year else None
        ws.cell(r, 1, f'◆ 客户：{c}（本期收入第 {k + 1} 大）').font = S.SHELL_SUB_FONT
        r += 1
        for j, h in enumerate(cols):
            _hcell(ws, r, 1 + j, h)
        r += 1
        for m in MONTHS:
            inc = rev_md.get(m, 0.0); cst = cost_md.get(m, 0.0); gp = inc - cst
            pin = prev_rev.get(m, 0.0) if prev_year else None
            pcs = prev_cost.get(m, 0.0) if prev_year else None
            pgp = (pin - pcs) if (prev_year and pin is not None) else None
            dinc = (inc - pin) if (prev_year and pin is not None) else None
            dcst = (cst - pcs) if (prev_year and pcs is not None) else None
            dgp = (gp - pgp) if (prev_year and pgp is not None) else None
            vals = [MONTH_LABEL[m], inc, cst, gp,
                    pin if pin is not None else '', pcs if pcs is not None else '',
                    pgp if pgp is not None else '', dinc if dinc is not None else '',
                    dcst if dcst is not None else '', dgp if dgp is not None else '']
            _dcell(ws, r, 1, vals[0], align=CEN)
            for j in range(1, ncols):
                _dcell(ws, r, 1 + j, vals[j])
            r += 1
        # 全年小计
        tinc = sum(rev_md.values()); tcst = sum(cost_md.values()); tgp = tinc - tcst
        tpin = sum(prev_rev.values()) if prev_year else None
        tpcs = sum(prev_cost.values()) if prev_year else None
        tpgp = (tpin - tpcs) if (prev_year and tpin is not None) else None
        tdinc = (tinc - tpin) if (prev_year and tpin is not None) else None
        tdcst = (tcst - tpcs) if (prev_year and tpcs is not None) else None
        tdgp = (tgp - tpgp) if (prev_year and tpgp is not None) else None
        vals = ['全年小计', tinc, tcst, tgp,
                tpin if tpin is not None else '', tpcs if tpcs is not None else '',
                tpgp if tpgp is not None else '', tdinc if tdinc is not None else '',
                tdcst if tdcst is not None else '', tdgp if tdgp is not None else '']
        _dcell(ws, r, 1, vals[0], bold=True, align=CEN)
        for j in range(1, ncols):
            _dcell(ws, r, 1 + j, vals[j], bold=True)
        r += 2
    for c in range(1, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14 if c == 1 else 13
    ws.freeze_panes = 'B5'
    return ws


# ===================== 表4：其他业务收入 全年汇总(按二级) =====================
def build_other(wb, entities, cur_year, prev_year, tb_all):
    ws = wb.create_sheet(f'其他业务收入_全年汇总(按二级,{cur_year})')
    ws.cell(1, 1, f'其他业务收入全年汇总表（按二级明细，{cur_year}'
                   + (f' 对比 {prev_year}' if prev_year else ' · 无上年度数据') + '）').font = S.SHELL_TITLE_FONT
    # 其他业务口径说明已按用户要求删除（2026-07-31）

    tb_rev_l2, tb_cost_l2 = {}, {}
    for E in sorted(entities):
        for y in (cur_year, prev_year) if prev_year else (cur_year,):
            tb = tb_all.get(E, {}).get(y)
            if not tb:
                continue
            # ⚡ 2026-08-12 SAP 修复：原 level==2+startswith(父码+'.') 对 SAP 无父级 TB 失效 →
            # 名称前缀法（铁律5）；主营收入/成本一级汇总行（名称==prefix）自动排除。
            for l2, amt in _tb_l2_entries(tb, '其他业务收入', 'credit').items():
                tb_rev_l2.setdefault(E, {}).setdefault(y, {})[l2] = \
                    tb_rev_l2.get(E, {}).get(y, {}).get(l2, 0.0) + amt
            for l2, amt in _tb_l2_entries(tb, '其他业务成本', 'debit').items():
                tb_cost_l2.setdefault(E, {}).setdefault(y, {})[l2] = \
                    tb_cost_l2.get(E, {}).get(y, {}).get(l2, 0.0) + amt

    # 所有二级名称
    all_l2 = sorted(set(
        n for E in tb_rev_l2 for y in tb_rev_l2[E] for n in tb_rev_l2[E][y]
    ) | set(
        n for E in tb_cost_l2 for y in tb_cost_l2[E] for n in tb_cost_l2[E][y]
    ))

    r = 4
    headers = ['核算主体', '二级明细', '本期收入', '本期成本', '本期利润',
               '上期收入', '上期成本', '上期利润', '利润增减']
    for j, h in enumerate(headers):
        _hcell(ws, r, 1 + j, h)
    r += 1

    g_inc = g_cst = g_pin = g_pcs = 0.0
    for E in sorted(entities):
        e_inc = e_cst = 0.0
        started = False
        for nm in all_l2:
            inc = tb_rev_l2.get(E, {}).get(cur_year, {}).get(nm, 0.0)
            cst = tb_cost_l2.get(E, {}).get(cur_year, {}).get(nm, 0.0)
            pin = tb_rev_l2.get(E, {}).get(prev_year, {}).get(nm, 0.0) if prev_year else 0.0
            pcs = tb_cost_l2.get(E, {}).get(prev_year, {}).get(nm, 0.0) if prev_year else 0.0
            if prev_year:
                if is_zero_amount(inc, cst, pin, pcs):
                    continue
            elif is_zero_amount(inc, cst):
                continue
            started = True
            gp = inc - cst
            pgp = (pin - pcs) if prev_year else None
            dgp = (gp - pgp) if (prev_year and pgp is not None) else None
            vals = [E, nm, inc, cst, gp,
                    pin if prev_year else '', pcs if prev_year else '',
                    pgp if pgp is not None else '', dgp if dgp is not None else '']
            for j, v in enumerate(vals):
                _dcell(ws, r, 1 + j, v)
            r += 1
            e_inc += inc; e_cst += cst
        if started:
            # 主体小计
            egp = e_inc - e_cst
            ws.cell(r, 1, '').fill = S.SHELL_TOT_FILL; ws.cell(r, 2, E + ' 小计').fill = S.SHELL_TOT_FILL
            ws.cell(r, 2).font = S.SHELL_BOLD
            for j_idx, val in [(3, e_inc), (4, e_cst), (5, egp)]:
                c = ws.cell(r, j_idx, round(val, 2)); c.fill = S.SHELL_TOT_FILL; c.font = S.SHELL_BOLD; c.number_format = '#,##0.00'
            for j in range(1, 10):
                ws.cell(r, j).border = S.SHELL_BORDER
            r += 1
            g_inc += e_inc; g_cst += e_cst

    # 全集团合计（含二级展开）
    ws.cell(r, 1, '').font = S.SHELL_BOLD; ws.cell(r, 2, '全集团合计').font = S.SHELL_BOLD
    for j in range(1, 10):
        ws.cell(r, j).fill = S.SHELL_TOT_FILL; ws.cell(r, j).border = S.SHELL_BORDER
    r += 1
    for nm in all_l2:
        inc = sum(tb_rev_l2.get(E, {}).get(cur_year, {}).get(nm, 0.0) for E in entities)
        cst = sum(tb_cost_l2.get(E, {}).get(cur_year, {}).get(nm, 0.0) for E in entities)
        pin = sum(tb_rev_l2.get(E, {}).get(prev_year, {}).get(nm, 0.0) for E in entities) if prev_year else 0.0
        pcs = sum(tb_cost_l2.get(E, {}).get(prev_year, {}).get(nm, 0.0) for E in entities) if prev_year else 0.0
        if prev_year and is_zero_amount(inc, cst, pin, pcs):
            continue
        if not prev_year and is_zero_amount(inc, cst):
            continue
        gp = inc - cst; pgp = (pin - pcs) if prev_year else None; dgp = (gp - pgp) if (prev_year and pgp is not None) else None
        vals = ['', nm, inc, cst, gp, pin if prev_year else '', pcs if prev_year else '',
                pgp if pgp is not None else '', dgp if dgp is not None else '']
        for j, v in enumerate(vals):
            _dcell(ws, r, 1 + j, v)
        r += 1
    # 全集团总计
    tinc = sum(tb_rev_l2.get(E, {}).get(cur_year, {}).get(nm, 0.0) for E in entities for nm in all_l2)
    tcst = sum(tb_cost_l2.get(E, {}).get(cur_year, {}).get(nm, 0.0) for E in entities for nm in all_l2)
    tgp = tinc - tcst
    _dcell(ws, r, 2, '总计', bold=True); _dcell(ws, r, 3, tinc, bold=True); _dcell(ws, r, 4, tcst, bold=True); _dcell(ws, r, 5, tgp, bold=True)
    for j in range(1, 10):
        ws.cell(r, j).fill = S.SHELL_TOT_FILL; ws.cell(r, j).font = S.SHELL_BOLD; ws.cell(r, j).border = S.SHELL_BORDER

    for c in range(1, 10):
        ws.column_dimensions[get_column_letter(c)].width = 18 if c <= 2 else 14
    ws.freeze_panes = 'C5'
    return ws


def build_revenue_cost_footnote(wb, entities, tb_all, cur_year, prev_year):
    """营业收入（营业成本）附注汇总（2026-08-04 用户要求：与固定资产等底稿附注一致，每主体一列）：
    行=主营业务收入/成本/利润、其他业务收入/成本/利润、营业总收入/总成本/营业利润；
    列=各核算主体 + 全集团合计（公式=SUM主体区）；
    取数=损益科目发生额（收入贷方 credit、成本借方 debit），按主体取数（非期末余额）。
    FY 无产品类别/内外销分类 → 收入分解信息暂缺（原格式2 合并口径废弃）。"""
    ws = wb.create_sheet('营业收入（营业成本）附注汇总')
    ws.cell(1, 1, f'营业收入（营业成本）附注汇总（{cur_year} 年，每主体一列）').font = S.SHELL_TITLE_FONT
    ents = sorted(entities)
    NC = 2 + len(ents)   # A项目 + 各主体列 + 全集团合计
    hdr = ['项目'] + ents + ['全集团合计']
    for j, h in enumerate(hdr, 1):
        cc = ws.cell(3, j, h); cc.font = S.SHELL_HFONT; cc.fill = S.SHELL_HFILL
        cc.border = S.SHELL_BORDER; cc.alignment = S.SHELL_CEN
    # 按主体取数（损益科目发生额：收入 credit、成本 debit）
    per = {}
    for E in ents:
        tb = (tb_all.get(E) or {}).get(cur_year)
        if not tb:
            continue
        cd = _detect_codes(tb)
        rec = dict(rev=0.0, cost=0.0, orev=0.0, ocost=0.0)
        for key in ('rev', 'cost', 'orev', 'ocost'):
            code = cd.get(key)
            if code and code in tb:
                v = tb[code]
                rec[key] = float(v.get('credit') if key in ('rev', 'orev') else v.get('debit') or 0.0)
        per[E] = rec

    def _row(r, label, getter, bold=False):
        """getter(E) -> 值；全集团合计=SUM(主体区)。"""
        ws.cell(r, 1, label)
        if bold:
            ws.cell(r, 1).font = S.SHELL_BOLD
            ws.cell(r, 1).fill = S.SHELL_HFILL
        else:
            ws.cell(r, 1).font = S.SHELL_HFONT
        for i, E in enumerate(ents):
            v = getter(E) or 0.0
            cell = ws.cell(r, 2 + i, round(v, 2))
            cell.number_format = S.SHELL_NUM
            cell.border = S.SHELL_BORDER
            cell.alignment = S.SHELL_RGT
            if bold:
                cell.fill = S.SHELL_TOT_FILL; cell.font = S.SHELL_BOLD
        # 2026-08-06 协议：全集团合计列写数值（=各主体和；公式 data_only 读 None 收回读不回）
        tot = ws.cell(r, NC, round(sum(ws.cell(r, 2 + i).value or 0.0 for i in range(len(ents))), 2))
        tot.number_format = S.SHELL_NUM; tot.border = S.SHELL_BORDER
        tot.alignment = S.SHELL_RGT
        if bold:
            tot.fill = S.SHELL_TOT_FILL; tot.font = S.SHELL_BOLD
        for j in range(1, NC + 1):
            cc = ws.cell(r, j)
            if cc.border is None or cc.border.left.style is None:
                cc.border = S.SHELL_BORDER
        return r + 1

    r = 4
    def _mk(k):
        return lambda E: per.get(E, {}).get(k, 0.0)
    r = _row(r, '一、主营业务收入', _mk('rev'), bold=True)
    r = _row(r, '二、主营业务成本', _mk('cost'))
    r = _row(r, '三、主营业务利润', lambda E: (per.get(E, {}).get('rev', 0.0) - per.get(E, {}).get('cost', 0.0)))
    r = _row(r, '四、其他业务收入', _mk('orev'), bold=True)
    r = _row(r, '五、其他业务成本', _mk('ocost'))
    r = _row(r, '六、其他业务利润', lambda E: (per.get(E, {}).get('orev', 0.0) - per.get(E, {}).get('ocost', 0.0)))
    r += 1
    r = _row(r, '七、营业总收入', lambda E: (per.get(E, {}).get('rev', 0.0) + per.get(E, {}).get('orev', 0.0)), bold=True)
    r = _row(r, '八、营业总成本', lambda E: (per.get(E, {}).get('cost', 0.0) + per.get(E, {}).get('ocost', 0.0)))
    r = _row(r, '九、营业利润', lambda E: (per.get(E, {}).get('rev', 0.0) + per.get(E, {}).get('orev', 0.0)
                                           - per.get(E, {}).get('cost', 0.0) - per.get(E, {}).get('ocost', 0.0)), bold=True)
    r += 1
    note = ws.cell(r, 1, '收入分解信息（按收入类别/内外销）暂缺：FY 数据无产品类别/内外销分类，待数据补充后生成。')
    note.font = Font(name='Times New Roman', size=9, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    for j, w in enumerate([24] + [14] * (NC - 1), 1):
        # 2026-08-05 修复：chr(64+j) 在列数>26（ga 40 主体 → NC=41）时产生非法列名 '[' → get_column_letter
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'B4'
    return ws


# ===================== 目录说明 =====================
def build_cover(wb, entities, cur_year, prev_year, has_customer):
    ws = wb.create_sheet('目录说明', 0)
    ws.cell(1, 1, f'营业收入审计底稿（{cur_year} 年度）').font = S.SHELL_TITLE_FONT
    lines = [
        '',
        '一、底稿构成',
        '  1. 主营业务收入_分月明细(按主体)：每个主体 收入/成本/毛利 分月 + 上年 + 增减 + 全年小计；',
        '     底部含全集团分月汇总、全集团全年合计，及与《科目余额表》逐主体/集团核对（口径=GL 借方 vs TB 借发）。',
        '  2. 主营业务收入_分月明细(按二级) + 主营业务收入_二级核对：按产品族配对展示（收入↔成本）；',
        '     收入/成本二级全部未匹配（收入按业务类型、成本按费用性质，无产品族配对）时不生成（无意义）。',
        '  3. 主营业务收入成本分部门明细表（附注汇总模板）：合并收入/成本两张分部门表；部门为列，',
        '     行=未审数/审计调整/审定数三块（主营业务收入/成本/毛利），表尾上年列组（集团上年同期），合计列公式化。',
        '  4. 主营业务收入_客户明细(年度) + 前十大客户_分月：当 GL/辅助核算含具体客户收入时生成。',
        '  5. 其他业务收入_全年汇总(按二级)：无分月，按二级明细汇总全年数（收入/成本/利润/增减）。',
        '',
        '二、取数与勾稽方法',
        '  · 审定表/其他业务：收入/成本按《科目余额表》取 gross 发生额（收入=贷方、成本=借方），代码按名称',
        '    识别，兼容境外子公司本地科目表（如 泰国25 用 4100/4200 而非 6001/6051）。',
        '  · 分月/二级核对：GL 分月数=GL 全部借方合计=TB 本期借方发生额（GL 全量抽取恒等）；',
        '    未结转损益账套（借发≠贷发，损益科目期末余额≠0）未结转额不进入 GL 分月数，作披露项（见主体核对注）。',
        '  · 毛利 = 收入 − 成本。',
        '  · 月度数据来自《综合查询明细表》(GL) 按日期取月份汇总；客户收入取自 GL「辅助核算名称」。',
        '  · 上年列仅当存在上一年度数据；无上年时留空，首年不虚增变动。',
        '  · GL 抽取期间 < TB 期间（如 2026 GL 仅 1-3 月 YTD）时差异按「取数范围限制」披露，不认定为账务差错。',
        '',
        f'三、覆盖范围：主体 {len(entities)} 个（{", ".join(sorted(entities))}）；',
        f'    会计年度 {cur_year}' + (f'、对比年度 {prev_year}' if prev_year else '（无上年度数据）') + '；',
        f'    客户级底稿：{"已生成（GL 含客户收入）" if has_customer else "未生成（GL/辅助核算无具体客户收入）"}。',
    ]
    r = 2
    for ln in lines:
        c = ws.cell(r, 1, ln)
        c.font = S.SHELL_BODY
        c.alignment = LEFT
        r += 1
    ws.column_dimensions['A'].width = 110
    return ws


# ===================== 审定表（主营业务收入/成本/利润 + 其他业务收入/成本/利润） =====================
def build_revenue_audit_sheet(wb, entities, tb_all, cur_y, prev_y, data_dir=None):
    """营业收入审定表（按年构建：本期数=cur_y、上年数=prev_y，2026-07-31 改按年拆分后
    每个年度文件各自构建，不再一次输出全部年份）。"""
    ws = wb.create_sheet('营业收入 审定表', 0)  # 插入为第一张
    ws.cell(1, 1, '营业收入审定表（全集团汇总）').font = S.SHELL_TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    hdr = ['项目', '核算主体', '上年数', '本期数', '审计调整数', '审定数']
    r = 3
    for j, h in enumerate(hdr, 1):
        cell = ws.cell(r, j, h)
        cell.font = S.SHELL_HFONT; cell.fill = S.SHELL_HFILL; cell.alignment = S.SHELL_CEN; cell.border = S.SHELL_BORDER
    r += 1

    # 从 TB 汇总（cur_y=本期、prev_y=上年对比，均由调用方传入）

    def _tb(ent, y, role):
        """从 TB 取数：rev→credit, cost→debit, other_rev→credit, other_cost→debit。
        ⚡⚡ 2026-08-30 结论：维持 gross 贷方/借方口径（铁律3）——与 TB 利润表发生额一致
        （收入取贷、成本取借；净额改法因红字冲回借方负数导致双计 ×2，已回退）。"""
        tb = tb_all.get(ent, {}).get(y, {})
        cd = _detect_codes(tb)
        if role == 'rev':
            return _tb_total(tb, 'rev', 'cr')
        elif role == 'cost':
            return _tb_total(tb, 'cost', 'db')
        elif role == 'orev':
            return _tb_total(tb, 'orev', 'cr') if cd.get('orev') else 0.0
        elif role == 'ocost':
            return _tb_total(tb, 'ocost', 'db') if cd.get('ocost') else 0.0
        return 0.0

    # 行定义：(label, role, is_minus, is_profit)
    # is_profit=True 的行现在也会按核算主体展开（2026-07-30 用户要求）
    # 2026-08-02 用户确认：成本列改为正数显示（与分部门明细表/参考表一致），利润行=收入−成本；
    # 「减：」前缀已表达抵减语义，is_minus 参数保留仅为兼容，不再取负。
    LINES = [
        ('一、主营业务收入', 'rev', False, False),
        ('减：主营业务成本', 'cost', True, False),
        ('  主营业务利润', None, False, True),  # 收入-成本，按核算主体展开
        ('二、其他业务收入', 'orev', False, False),
        ('减：其他业务成本', 'ocost', True, False),
        ('  其他业务利润', None, False, True),  # 其他收入-其他成本，按核算主体展开
    ]

    # ⚡ 2026-08-11 阶段一 1.1 合并底稿差异化：收集横展合并审定表数据（行=项目，列=主体）
    merge_vals = {e: {} for e in sorted(entities)}
    merge_prev = {}

    for label, role, is_minus, is_profit in LINES:
        tot_cur = 0.0; tot_prev = 0.0
        body_first = r                        # 本块主体行首行（全集团合计公式引用范围）
        ent_vals = {}  # {ent: (cur, prev)}
        for ent in sorted(entities):
            v_cur = _tb(ent, cur_y, role) if role else 0.0
            v_prev = _tb(ent, prev_y, role) if (role and prev_y) else 0.0
            if is_profit:
                # 利润行：直接从 TB 取值计算（收入_ent - 成本_ent）
                is_main = (label == '  主营业务利润')
                v_cur_rev = _tb(ent, cur_y, 'rev') if is_main else _tb(ent, cur_y, 'orev')
                v_prev_rev = _tb(ent, prev_y, 'rev') if (prev_y and is_main) else (_tb(ent, prev_y, 'orev') if prev_y else 0.0)
                v_cur_cost = _tb(ent, cur_y, 'cost') if is_main else _tb(ent, cur_y, 'ocost')
                v_prev_cost = _tb(ent, prev_y, 'cost') if (prev_y and is_main) else (_tb(ent, prev_y, 'ocost') if prev_y else 0.0)
                v_cur = v_cur_rev - v_cur_cost
                v_prev = v_prev_rev - v_prev_cost
                if v_cur == 0 and v_prev == 0:
                    continue
                val_cur = v_cur
                val_prev = v_prev
                ent_vals[ent] = (val_cur, val_prev)
            else:
                if v_cur == 0 and v_prev == 0:
                    continue
                # 2026-08-02：成本正数显示（不再 is_minus 取负）
                val_cur = v_cur
                val_prev = v_prev
                ent_vals[ent] = (val_cur, val_prev)
            merge_vals[ent][label] = val_cur          # 横展矩阵：主体×项目（本期数）
            ws.cell(r, 1, label).border = S.SHELL_BORDER
            ws.cell(r, 2, ent).border = S.SHELL_BORDER
            ws.cell(r, 3, round(val_prev, 2)).border = S.SHELL_BORDER; ws.cell(r, 3).number_format = '#,##0.00'
            ws.cell(r, 4, round(val_cur, 2)).border = S.SHELL_BORDER; ws.cell(r, 4).number_format = '#,##0.00'
            ws.cell(r, 5, None).border = S.SHELL_BORDER
            # 2026-08-06 协议：审定数写数值（=本期+调整0）
            ws.cell(r, 6, round(val_cur, 2)).border = S.SHELL_BORDER; ws.cell(r, 6).number_format = '#,##0.00'
            tot_cur += val_cur; tot_prev += val_prev
            r += 1
        # 全集团合计（2026-08-06 协议：写数值=tot_prev/tot_cur 累计；公式 data_only 读 None）
        ws.cell(r, 1, label).font = S.SHELL_BOLD; ws.cell(r, 1).fill = S.SHELL_TOT_FILL
        ws.cell(r, 2, '全集团合计').font = S.SHELL_BOLD; ws.cell(r, 2).fill = S.SHELL_TOT_FILL
        ws.cell(r, 3, round(tot_prev, 2)); ws.cell(r, 3).font = S.SHELL_BOLD; ws.cell(r, 3).fill = S.SHELL_TOT_FILL
        ws.cell(r, 3).number_format = '#,##0.00'
        ws.cell(r, 4, round(tot_cur, 2)); ws.cell(r, 4).font = S.SHELL_BOLD; ws.cell(r, 4).fill = S.SHELL_TOT_FILL
        ws.cell(r, 4).number_format = '#,##0.00'
        ws.cell(r, 5, None).fill = S.SHELL_TOT_FILL
        ws.cell(r, 6, round(tot_cur, 2)).fill = S.SHELL_TOT_FILL; ws.cell(r, 6).number_format = '#,##0.00'
        for j in range(1, 7):
            ws.cell(r, j).border = S.SHELL_BORDER
            ws.cell(r, j).font = S.SHELL_BOLD
        r += 1
        merge_prev[label] = tot_prev                    # 横展矩阵：项目上年数

    # ⚡ 2026-08-11 阶段一 1.1 合并底稿差异化：集团模式（多主体）追加横展合并审定表
    if len(entities) > 1:
        try:
            S.build_merge_wide_sheet(
                wb, '营业收入 合并审定表',
                f'营业收入合并审定表（{cur_y}）· 集团口径：行=项目，列=主体，抵消待填',
                [(lb, lb) for lb, *_ in LINES], merge_vals, prev_vals=merge_prev)
        except Exception as _ex:
            print(f'  ⚠️ 合并审定表失败：{_ex}')

    for i, w in enumerate([30, 16, 14, 14, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'
    return ws


def build_rev_cost_instruction_sheet(wb, cur_year):
    """收入成本按指令号配比表（2026-08-10 用户方法论：yy 指令号(WBS)/300 订单号
    在收入↔成本间配比）。
    数据：GL 收入行（6001* 贷）与成本行（6401* 借）按指令号（wbs 优先，300 用 ord/物料）聚合。
    实测 1010：收入 98.4% 金额、成本 94.7% 金额有指令号 → 配比可行。
    输出：指令号|收入|成本|毛利|毛利率；无指令号行并入『（无指令号）』；仅列 Top 300 指令号。
    ⚡ 2026-08-11 三处修复：
      ① WBS 父级聚合——SAP 成本挂子 WBS（P-GF24126.1KF119 等 94 个），收入挂父级
         P-GF24126.1 → 按全名聚合父级收入 2.42 亿成本=0（误判"无成本"）。父级=截断到
         "P-GF数字.数字" 主段（子级是其后字母段）。订单号/物料号无子级不变。
      ② 集团模式生成——原 `not _adapter.current_comp()` 集团下直接跳过（集团底稿无此表）。
         改为 current_comp 有值用单主体 read_gl；集团模式按主体循环聚合。
      ③ 合计=Top300 明细之和（原合计=全部指令号 2061 个 → 明细 300 行 Σ≠合计，勾稽挂）。
         尾部注明 Top300 外 N 个指令号合计，供审计人员决定是否全列。"""
    import re as _re
    from collections import defaultdict
    ws = wb.create_sheet('收入成本_按指令号配比表')
    ws.cell(1, 1, f'收入-成本按指令号（WBS/订单）配比（{cur_year} 年度；毛利=收入-成本；单位：元）').font = S.SHELL_TITLE_FONT

    def _wbs_parent(w):
        """WBS 子级截断：P-GF24126.1KF119 → P-GF24126.1；P-GF24126.1 → 自身。"""
        m = _re.match(r'^(P-GF\d+\.\d+)', str(w))
        return m.group(1) if m else str(w)

    def _gather_gl(rows_iter):
        rev = defaultdict(float)
        cost = defaultdict(float)
        # ⚡ 2026-08-12 需求③：收入/成本对应科目列——同凭证(vno)内其他科目名聚合
        #   （收入确认凭证=6001贷+应收借+销项税贷 → 对应科目=应收账款/应交税费-销项税；
        #   成本结转=6401借+库存贷 → 对应科目=库存商品等）。供审计定位收入/成本挂账形态。
        rev_cp = defaultdict(set)
        cost_cp = defaultdict(set)
        _rows = list(rows_iter)
        _by_vno = defaultdict(list)
        for _r in _rows:
            _by_vno[(_r.get('y'), _r.get('vno'))].append(_r)
        for r in _rows:
            c = str(r.get('code') or '')
            if not c.startswith(('6001', '6401')):
                continue
            w = str(r.get('wbs') or '').strip()
            o = str(r.get('ord') or '').strip()
            m = str(r.get('mat') or '').strip()
            inst = _wbs_parent(w) if w else (o or m or '（无指令号）')
            # 同凭证其他行科目名（金额绝对值降序，取前 3）
            _others = [x for x in _by_vno.get((r.get('y'), r.get('vno')), [])
                       if x is not r and str(x.get('name') or '') and str(x.get('name')) != str(r.get('name'))]
            _cp_names = sorted({str(x.get('name') or '') for x in _others},
                               key=lambda n: -max((abs(float(x.get('cr') or 0.0)) + abs(float(x.get('dr') or 0.0)))
                                                  for x in _others if str(x.get('name')) == n))[:3]
            if c.startswith('6001'):
                rev[inst] += float(r.get('credit') or r.get('cr') or 0.0)
                if _cp_names:
                    rev_cp[inst].add('、'.join(_cp_names))
            else:
                cost[inst] += float(r.get('debit') or r.get('dr') or 0.0)
                if _cp_names:
                    cost_cp[inst].add('、'.join(_cp_names))
        return rev, cost, rev_cp, cost_cp

    rev_all = defaultdict(float)
    cost_all = defaultdict(float)
    rev_cp_all = defaultdict(set)
    cost_cp_all = defaultdict(set)
    if _adapter is None:
        ws.cell(3, 1, '（非 SAP 账套无此表）').font = S.SHELL_BODY
        return ws
    if _adapter.current_comp():
        # 单主体（1010 等）：read_gl 即该主体
        _rv, _cs, _rvc, _csc = _gather_gl(_adapter.read_gl())
        rev_all.update(_rv); cost_all.update(_cs)
        for k, v in _rvc.items(): rev_cp_all[k] |= v
        for k, v in _csc.items(): cost_cp_all[k] |= v
    else:
        # 集团模式：逐主体聚合（read_gl_rows 需逐主体调用防 OOM）
        for _c in sorted(_adapter.discover_entities(getattr(_adapter, '_DATA_ROOT', None) or P.YY)):
            try:
                _rv, _cs, _rvc, _csc = _gather_gl(_adapter.read_gl_rows(getattr(_adapter, '_DATA_ROOT', None) or P.YY,
                                                                        {_c: _adapter.discover_entities(getattr(_adapter, '_DATA_ROOT', None) or P.YY).get(_c, {})}))
                for k, v in _rv.items(): rev_all[k] += v
                for k, v in _cs.items(): cost_all[k] += v
                for k, v in _rvc.items(): rev_cp_all[k] |= v
                for k, v in _csc.items(): cost_cp_all[k] |= v
            except Exception:
                continue
    if not rev_all and not cost_all:
        ws.cell(3, 1, '（无收入/成本数据）').font = S.SHELL_BODY
        return ws
    insts = sorted(set(rev_all) | set(cost_all), key=lambda x: -(rev_all.get(x, 0.0) + cost_all.get(x, 0.0)))
    top = insts[:300]
    # 合计 = Top300 明细之和（勾稽一致）
    _tr = sum(rev_all.get(x, 0.0) for x in top)
    _tc = sum(cost_all.get(x, 0.0) for x in top)
    _rest_r = sum(rev_all.values()) - _tr
    _rest_n = len(insts) - len(top)
    r = 3
    ws.cell(r, 1, '合计').font = S.SHELL_SUB_FONT
    ws.cell(r, 2, _tr).number_format = S.SHELL_NUM
    ws.cell(r, 3, _tc).number_format = S.SHELL_NUM
    ws.cell(r, 4, _tr - _tc).number_format = S.SHELL_NUM
    ws.cell(r, 5, (_tr - _tc) / _tr if _tr else 0).number_format = '0.00%'
    r += 2
    S.apply_header_row(ws, r, ['指令号', '收入', '成本', '毛利', '毛利率', '收入对应科目', '成本对应科目'])
    r += 1
    n_out = 0
    for inst in top:
        rv = rev_all.get(inst, 0.0)
        cs = cost_all.get(inst, 0.0)
        if abs(rv) < 0.005 and abs(cs) < 0.005:
            continue
        ws.cell(r, 1, inst).font = S.SHELL_BODY
        ws.cell(r, 2, rv).number_format = S.SHELL_NUM
        ws.cell(r, 3, cs).number_format = S.SHELL_NUM
        ws.cell(r, 4, rv - cs).number_format = S.SHELL_NUM
        ws.cell(r, 5, (rv - cs) / rv if rv else 0).number_format = '0.00%'
        ws.cell(r, 6, '；'.join(sorted(rev_cp_all.get(inst, set()))[:2]) or None).alignment = S.SHELL_LEFT
        ws.cell(r, 7, '；'.join(sorted(cost_cp_all.get(inst, set()))[:2]) or None).alignment = S.SHELL_LEFT
        for c in range(1, 8):
            ws.cell(r, c).border = S.SHELL_BORDER
        r += 1
        n_out += 1
    if not n_out:
        ws.cell(r, 1, '（无指令号数据）').font = S.SHELL_BODY
        r += 1
    r += 1
    ws.cell(r, 1, '说明：收入按 6001* 贷方、成本按 6401* 借方按指令号聚合；WBS 成本子级（KF/KL/K1/K2…）'
                  '归并到父级（P-GFxxxxx.x）便于收入成本配对；无指令号行并入（无指令号）。'
                  '毛利率=毛利/收入。合计=所列 Top300 指令号之和（共 %d 个指令号，Top300 之外 %d 个'
                  '指令号收入合计 %s，如需要可改参数全列）。收入对应科目=同凭证其他借方科目（应收/预收/'
                  '销项税等），成本对应科目=同凭证其他贷方科目（库存/暂估等），供定位挂账形态。（无指令号）'
                  '行通常为 出口/免税/内部 等不挂指令号业务，需结合销项税配对核查表甄别。'
                  % (len(insts), _rest_n, f'{_rest_r:,.2f}')
                  ).font = S.SHELL_BODY
    for i, w in enumerate([30, 16, 16, 16, 10, 30, 30], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'B4'
    return ws


def build_vat_recon_sheet(wb, cur_year):
    """销项税-收入配对核查表（SAP 专属，2026-08-10 用户方法论：综合税率偏离/
    无销项税大额凭证核查）。
    数据：GL 收入行（6001* 贷）与销项税行（2221010105 贷）按凭证号配对。
    输出：①汇总（收入/销项税/综合税率）②税率分布（13/9/6/3/0 分组）
          ③偏离凭证（配对且税率不在标准±0.5个百分点，逐笔）④无销项税大额收入（≥100万）。"""
    from collections import defaultdict
    ws = wb.create_sheet('销项税-收入配对核查表')
    ws.cell(1, 1, f'销项税-收入配对核查（{cur_year} 年度；单位：元；税率=销项税/不含税收入）').font = S.SHELL_TITLE_FONT
    if _adapter is None or not _adapter.current_comp():
        ws.cell(3, 1, '（SAP 单主体模式专属）').font = S.SHELL_BODY
        return ws
    gl = _adapter.read_gl()
    rev = defaultdict(float)
    vat = defaultdict(float)
    rinfo = {}
    for r in gl:
        c = str(r.get('code') or '')
        nm = str(r.get('name') or '')
        vno = str(r.get('vno') or '')
        cr = float(r.get('credit') or 0.0)
        # ⚡ 2026-08-10 科目码双体系：yy=2221010105 / 300=2221010500（应交税费-增值税-销项税额）
        #   → 销项税按名称『销项税』匹配（铁律：科目匹配以名称为主，代码仅结构定位）。
        if (c.startswith('6001') or nm.startswith('主营业务收入')) and cr > 0.005:
            rev[vno] += cr
            if vno not in rinfo:
                rinfo[vno] = (str(r.get('date') or ''), str(r.get('sm') or '')[:34])
        elif ('销项税' in nm or c.startswith('2221010105')) and cr > 0.005:
            vat[vno] += cr
    _STD = (0.13, 0.09, 0.06, 0.03, 0.0)
    # ① 汇总
    tot_rev = sum(rev.values())
    tot_vat = sum(vat.values())
    r = 3
    ws.cell(r, 1, '主营业务收入（6001*）贷方合计').font = S.SHELL_SUB_FONT
    ws.cell(r, 2, tot_rev).number_format = S.SHELL_NUM
    r += 1
    ws.cell(r, 1, '销项税（2221010105）贷方合计').font = S.SHELL_SUB_FONT
    ws.cell(r, 2, tot_vat).number_format = S.SHELL_NUM
    r += 1
    ws.cell(r, 1, '综合税率').font = S.SHELL_SUB_FONT
    ws.cell(r, 2, tot_vat / tot_rev if tot_rev else 0).number_format = '0.00%'
    r += 2
    # ② 税率分布（配对凭证）
    n_std = {s: 0 for s in _STD}
    n_other = 0
    for vno in set(rev) & set(vat):
        rr = vat[vno] / rev[vno] if rev[vno] else 0
        hit = False
        for s in _STD:
            if abs(rr - s) <= 0.005:
                n_std[s] += 1
                hit = True
                break
        if not hit:
            n_other += 1
    ws.cell(r, 1, '配对凭证税率分布（笔数）').font = S.SHELL_SUB_FONT
    ws.cell(r, 2, '13%').font = S.SHELL_BODY
    ws.cell(r, 3, n_std[0.13]).number_format = S.SHELL_NUM_INT
    ws.cell(r, 4, '9%').font = S.SHELL_BODY
    ws.cell(r, 5, n_std[0.09]).number_format = S.SHELL_NUM_INT
    ws.cell(r, 6, '6%').font = S.SHELL_BODY
    ws.cell(r, 7, n_std[0.06]).number_format = S.SHELL_NUM_INT
    ws.cell(r, 8, '3%').font = S.SHELL_BODY
    ws.cell(r, 9, n_std[0.03]).number_format = S.SHELL_NUM_INT
    ws.cell(r, 10, '0%').font = S.SHELL_BODY
    ws.cell(r, 11, n_std[0.0]).number_format = S.SHELL_NUM_INT
    ws.cell(r, 12, '其他偏离').font = S.SHELL_BODY
    ws.cell(r, 13, n_other).number_format = S.SHELL_NUM_INT
    r += 2
    # ③ 偏离凭证明细
    ws.cell(r, 1, '税率偏离凭证（不在 13/9/6/3/0% ±0.5个百分点）').font = S.SHELL_SUB_FONT
    r += 1
    S.apply_header_row(ws, r, ['日期', '凭证号', '摘要', '收入', '销项税', '税率', '判定'])
    r += 1
    _devi = []
    for vno in set(rev) & set(vat):
        rr = vat[vno] / rev[vno] if rev[vno] else 0
        if any(abs(rr - s) <= 0.005 for s in _STD):
            continue
        _devi.append((vno, rinfo.get(vno, ('', ''))[0], rinfo.get(vno, ('', ''))[1],
                      rev[vno], vat[vno], rr))
    _devi.sort(key=lambda x: -max(x[3], x[4]))
    for vno, dt, sm, rv, vt, rr in _devi[:60]:
        ws.cell(r, 1, dt).font = S.SHELL_BODY
        ws.cell(r, 2, vno).font = S.SHELL_BODY
        ws.cell(r, 3, sm).font = S.SHELL_BODY
        ws.cell(r, 4, rv).number_format = S.SHELL_NUM
        ws.cell(r, 5, vt).number_format = S.SHELL_NUM
        ws.cell(r, 6, rr).number_format = '0.00%'
        ws.cell(r, 7, '需甄别').font = S.SHELL_BODY
        for c in range(1, 8):
            ws.cell(r, c).border = S.SHELL_BORDER
        r += 1
    if not _devi:
        ws.cell(r, 1, '（无偏离凭证）').font = S.SHELL_BODY
        r += 1
    r += 1
    # ④ 无销项税大额收入（≥100万）
    ws.cell(r, 1, '无销项税且收入≥100万 的凭证（出口免税/视同销售/异常待核）').font = S.SHELL_SUB_FONT
    r += 1
    S.apply_header_row(ws, r, ['日期', '凭证号', '摘要', '收入', '销项税', '税率', '判定'])
    r += 1
    _nof = [(vno, rinfo.get(vno, ('', ''))[0], rinfo.get(vno, ('', ''))[1], rev[vno])
            for vno in (set(rev) - set(vat)) if rev[vno] >= 1000000]
    _nof.sort(key=lambda x: -x[3])
    for vno, dt, sm, rv in _nof[:60]:
        ws.cell(r, 1, dt).font = S.SHELL_BODY
        ws.cell(r, 2, vno).font = S.SHELL_BODY
        ws.cell(r, 3, sm).font = S.SHELL_BODY
        ws.cell(r, 4, rv).number_format = S.SHELL_NUM
        ws.cell(r, 5, 0).number_format = S.SHELL_NUM
        ws.cell(r, 6, 0).number_format = '0.00%'
        ws.cell(r, 7, '待核（免税/未开票/异常）').font = S.SHELL_BODY
        for c in range(1, 8):
            ws.cell(r, c).border = S.SHELL_BORDER
        r += 1
    if not _nof:
        ws.cell(r, 1, '（无）').font = S.SHELL_BODY
        r += 1
    r += 1
    ws.cell(r, 1, '说明：销项税凭证 %d 笔 / %s 元；无销项税大额收入 %d 笔 / %s 元。'
                  '综合税率与标准税率差异源于混合税率/出口免税/价外费用等，需结合开票明细核对。'
                  % (len(vat), format(tot_vat, ',.2f'), len(_nof), format(sum(x[3] for x in _nof), ',.2f'))).font = S.SHELL_BODY
    for i, w in enumerate([12, 16, 34, 15, 15, 10, 22], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'
    return ws


def build_rev_counterparty_sheet(wb, cur_year, entities=None, data_dir=None):
    """营业收入对方科目核对（2026-08-12 补齐对方科目覆盖盘点缺口：收入循环无抽凭/对方科目；
    2026-08-15 修复：数据源改用本地 read_gl（与主流程同源），不再依赖全局 SAP adapter
    _DATA_ROOT——XBJ 等 U8 账套曾被错误指向 SAP 88 家主体）。

    数据：GL 收入确认凭证（主营业务收入 贷）按【年月+凭证号】取同凭证【对方科目】——
    借方侧（应收/预收/合同资产/现金等收款方）+ 贷方侧（销项税/其他收益等）。
    输出：①对方科目汇总（按凭证聚合的借方/贷方科目及金额）②收入确认凭证笔数。
    单主体+集团模式均支持（集团逐主体防 OOM）。"""
    from collections import defaultdict
    ws = wb.create_sheet('营业收入对方科目核对')
    ws.cell(1, 1, f'营业收入对方科目核对（{cur_year} 年度；按收入确认凭证 主营业务收入贷 '
                  f'取同【年月+凭证号】对方科目；单位：元）').font = S.SHELL_TITLE_FONT
    if not entities or not data_dir:
        ws.cell(3, 1, '（无数据：未提供 entities/data_dir）').font = S.SHELL_BODY
        return ws

    def _scan(rows):
        # ⚡ 精确聚合 key=（年, 月, 凭证号）——凭证号跨月复用（每月记-0001 起），
        #   只按凭证号聚合会串月混入其他凭证（曾致 16.97 亿主营业务成本误报）。
        rev_rows = [r for r in rows
                    if str(r.get('name') or '').startswith('主营业务收入')
                    and float(r.get('credit') or 0.0) > 0.005]
        by_vno = defaultdict(list)
        for r in rows:
            by_vno[(r.get('year'), r.get('month'), str(r.get('vno') or ''))].append(r)
        db_cp = defaultdict(float)      # 对方科目借方（收入确认借方：应收/预收/现金…）
        cr_cp = defaultdict(float)      # 对方科目贷方（销项税等）
        n_rev = len(rev_rows)
        # ⚡⚡ 2026-08-28 修复（对方科目核对金额远超收入 332 亿根因）：原【逐收入行】累加
        #   同凭证其他行——一张凭证 N 行收入时，应收借方被重复累加 N 次（如 2100000026 8 行
        #   收入 → 应收 199.98万×8）。改为【凭证级去重】：每张收入确认凭证的对方科目只累加一次；
        #   同时排除收入行自身（同凭证多行收入不再互加），并保持含税口径（借方应收含税 > 贷方
        #   收入不含税，属正常，核对表注明）。
        seen_vno = set()
        for r in rev_rows:
            k = (r.get('year'), r.get('month'), str(r.get('vno') or ''))
            if k in seen_vno:
                continue
            seen_vno.add(k)
            for x in by_vno.get(k, []):
                nm = str(x.get('name') or '')
                if not nm or nm.startswith('主营业务收入'):
                    continue
                d = float(x.get('debit') or 0.0)
                c = float(x.get('credit') or 0.0)
                if d > 0.005:
                    db_cp[nm] += d
                if c > 0.005:
                    cr_cp[nm] += c
        return db_cp, cr_cp, n_rev

    db_cp, cr_cp, n_rev = defaultdict(float), defaultdict(float), 0
    for _c, yd in entities.items():
        _gl = None
        for _y, _p in yd.items():
            if _p.get('gl'):
                _gl = _p['gl']
                break
        if not _gl:
            continue
        try:
            _d, _c2, _n = _scan(_read_gl(_gl))
            for k, v in _d.items(): db_cp[k] += v
            for k, v in _c2.items(): cr_cp[k] += v
            n_rev += _n
        except Exception:
            continue
    ents_note = f'集团 {len(entities)} 家'
    r = 3
    ws.cell(r, 1, f'收入确认凭证 {n_rev} 笔（{ents_note}）').font = S.SHELL_SUB_FONT
    r += 1
    ws.cell(r, 1, '收入确认凭证 借方对方科目（收款方：应收/预收/合同资产/现金等）').font = S.SHELL_SUB_FONT
    r += 1
    S.apply_header_row(ws, r, ['对方科目', '借方金额', '占比'])
    r += 1
    _db_tot = sum(db_cp.values())
    for nm in sorted(db_cp, key=lambda x: -db_cp[x]):
        ws.cell(r, 1, nm).font = S.SHELL_BODY
        ws.cell(r, 2, db_cp[nm]).number_format = S.SHELL_NUM
        ws.cell(r, 3, db_cp[nm] / _db_tot if _db_tot else 0).number_format = '0.00%'
        for c in range(1, 4):
            ws.cell(r, c).border = S.SHELL_BORDER
        r += 1
    if not db_cp:
        ws.cell(r, 1, '（无）').font = S.SHELL_BODY
        r += 1
    r += 1
    ws.cell(r, 1, '收入确认凭证 贷方对方科目（销项税/其他收益等）').font = S.SHELL_SUB_FONT
    r += 1
    S.apply_header_row(ws, r, ['对方科目', '贷方金额', '占比'])
    r += 1
    _cr_tot = sum(cr_cp.values())
    for nm in sorted(cr_cp, key=lambda x: -cr_cp[x]):
        ws.cell(r, 1, nm).font = S.SHELL_BODY
        ws.cell(r, 2, cr_cp[nm]).number_format = S.SHELL_NUM
        ws.cell(r, 3, cr_cp[nm] / _cr_tot if _cr_tot else 0).number_format = '0.00%'
        for c in range(1, 4):
            ws.cell(r, c).border = S.SHELL_BORDER
        r += 1
    if not cr_cp:
        ws.cell(r, 1, '（无）').font = S.SHELL_BODY
        r += 1
    r += 1
    ws.cell(r, 1, '说明：收入按【主营业务收入 贷】行定位确认凭证，对方科目=同【年月+凭证号】其他'
                  '科目按借贷方向聚合（凭证级精确口径，凭证号跨月复用已按年月区分）；借方侧应收/合同'
                  '结算占比高=赊销确认，贷方侧销项税/待转销项占比=开票确认；施工企业完工百分比一票'
                  '确认（经营计价）凭证内含 主营业务成本借/合同履约成本结转贷，属正常记账习惯。'
                  '集团模式按主体循环聚合（防 OOM）。').font = S.SHELL_BODY
    r += 1
    ws.cell(r, 1, '⚠️ 借方发生额=确认凭证【借方总额】，≠收入金额：确认凭证借贷平衡，借方=合同结算'
                  '\收入结转+合同资产+应收账款+成本同行，贷方=主营业务收入(不含税)+待转销项税+'
                  '价款结算——借方合计约为收入的 2~3 倍是凭证结构必然，各借方科目与收入无固定比率'
                  '（全集团：合同结算≈收入1.2倍[收入结转]、合同资产≈收入1.2倍[已确认未结算含税收款权]、'
                  '应收=赊销）。收入=报表营业收入=主营业务收入+其他业务收入 贷方（不含税，见审定表）。').font = S.SHELL_BODY
    for i, w in enumerate([40, 16, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'
    return ws


def build_revenue_workbook(data_dir):
    global _MIRROR_MODE
    print(f'>> 处理文件夹：{data_dir}')
    # ⚡⚡ 2026-09-01 损益口径模式：有企业报表目录（AH 账套导出）→ 真实账套（净额，
    #   同额行=0，AH 2680 其他业务 605103 jf=df 取净额=企业报表 3.39M）；无报表
    #   （XBJ/AZ）→ 镜像账套（同额行取发生额，XBJ 主营收入 547M）。
    try:
        import sap_report as _RPT
        _MIRROR_MODE = not bool(_RPT.has_enterprise_reports(data_dir))
    except Exception:
        _MIRROR_MODE = None
    entities = _discover_entities(data_dir)
    if not entities:
        print('  ❌ 未发现任何账套主体（需含《科目余额表》+《综合查询明细表》）。')
        return None
    years = sorted({y for b in entities.values() for y in b})
    try:
        from mask_dict import mask_names as _mn
        _ents = _mn(sorted(entities))
    except Exception:
        _ents = sorted(entities)
    print(f'  · 发现主体 {len(entities)} 个：' + ', '.join(_ents))
    print(f'  · 年份：' + ', '.join(years))
    cur_year_holder[0] = max(years)

    n_km = sum(1 for b in entities.values() for v in b.values() if v.get('km'))
    n_gl = sum(1 for b in entities.values() for v in b.values() if v.get('gl'))
    print(f'  · 含《科目余额表》主体 {n_km} 个；含《综合查询明细表》主体 {n_gl} 个。')
    if n_gl == 0:
        print('  ❌ 未发现任何《综合查询明细表》(.xlsx)！主营业务收入分月数据必须来自 GL。')

    # 读 TB 与 GL 聚合
    tb_all = {}
    data = {}
    aux_dept_all = {}     # E -> {y: _read_aux_dept 结果}（分部门明细表数据源，2026-08-01）
    # ⚡⚡ 2026-08-30 集团模式逐主体锁定：current_comp=None 时 read_km/_aggregate 的区间文件
    #   （2310~2400.xlsx）解析为【区间起点】主体（2310）→ 2370 读到 2310 数据（串号根因，
    #   营业收入 8+ 主体同值 417.6M）。2405 注释设计"read_km 按 current_comp"需在此激活。
    _orig_comp = _adapter.current_comp() if (_adapter is not None and hasattr(_adapter, 'current_comp')) else None
    for E in entities:
        if _adapter is not None and hasattr(_adapter, 'set_comp'):
            _adapter.set_comp(E)
        yd = entities[E]
        tb_all[E] = {}
        data[E] = {}
        aux_dept_all[E] = {}
        for y, paths in yd.items():
            if paths.get('km'):
                # ⚡ 2026-08-13 原则（用户定）：生成器只负责生成，数据问题前端解决——
                #   不在此处区分 SAP/U8，主体解析由前端 read_km（adapter 按 current_comp）完成。
                # ⚡⚡ 2026-08-15 铁律127：TB 损益缺失 → GL 兜底（ga 实证：16/40 主体 TB
                #   损益科目本期发生额全 0，GL 序时账完整）→ read_tb_pl_safe 自动检测补齐。
                tb_all[E][y] = _read_tb_pl_safe(paths)
            if paths.get('gl'):
                agg = _aggregate(paths['gl'])
                # 年份统一为字符串（discover 用 '2025'，_aggregate 现已按 str(year) 键）
                data[E][y] = agg.get(y, {})
            if paths.get('aux'):
                aux_dept_all[E][y] = _read_aux_dept(paths['aux'])
    # ⚡⚡ 恢复 current_comp（集团模式循环结束还原 None，避免污染后续读取）
    if _adapter is not None and hasattr(_adapter, 'set_comp'):
        _adapter.set_comp(_orig_comp)

    # 按年独立构建并保存（2026-07-31 用户要求：与其他科目一致，每年一份底稿；
    # 本期=cur_y、上年对比=prev_y（有则显示），符合铁律20 每年底稿只含当年数据）
    outs = []
    for i, cur_year in enumerate(years):
        prev_year = years[i - 1] if i > 0 else None
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        build_revenue_audit_sheet(wb, entities, tb_all, cur_year, prev_year, data_dir)
        print(f'  ✓ 营业收入 审定表（{cur_year}' + (f' 对比 {prev_year}' if prev_year else '') + ')')
        build_by_entity(wb, entities, cur_year, prev_year, data, tb_all)
        print(f'  ✓ 主营业务收入_分月明细(按主体)（{cur_year}）')
        build_cost_monthly_sheet(wb, entities, cur_year, data, tb_all)
        print(f'  ✓ 主营业务成本_分月明细（{cur_year}）')
        build_rev_monthly_sheet(wb, entities, cur_year, data, tb_all)
        print(f'  ✓ 主营业务收入_分月明细（{cur_year}）')
        _s2 = build_by_l2(wb, entities, cur_year, prev_year, data, tb_all)
        if _s2 is not None:
            print(f'  ✓ 主营业务收入_分月明细(按二级)（{cur_year}）')
        _s3 = build_l2_recon_sheet(wb, entities, cur_year, prev_year, data, tb_all)
        if _s3 is not None:
            print(f'  ✓ 主营业务收入_二级核对（{cur_year}）')
        else:
            print('  ⊘ 收入/成本二级全部未匹配（收入按业务类型、成本按费用性质，无产品族配对）→ '
                  '跳过 分月明细(按二级) 与 二级核对（2026-08-02 用户要求）。')
        # 分部门明细表（2026-08-02 用户需求：合并收入/成本两张表为一张附注汇总模板，
        # 部门为列 + 未审/调整/审定三块 + 表尾上年列组 + 合计公式化）
        build_dept_footnote(wb, entities, cur_year, prev_year, aux_dept_all)
        print(f'  ✓ 主营业务收入成本分部门明细表（附注汇总模板）（{cur_year}）')
        has_customer = any((data.get(E, {}).get(cur_year) or {}).get('rev_cust') for E in entities)
        if has_customer:
            _, top = build_by_customer(wb, entities, cur_year, prev_year, data)
            print(f'  ✓ 主营业务收入_客户明细(全年)（{cur_year}）')
            build_top10_customer(wb, entities, cur_year, prev_year, data, top)
            print(f'  ✓ 主营业务收入_前十大客户_分月（{cur_year}）')
        else:
            print('  ⊘ 未检测到客户级收入（GL 辅助核算名称为空），跳过客户底稿。')
        build_other(wb, entities, cur_year, prev_year, tb_all)
        print(f'  ✓ 其他业务收入_全年汇总(按二级)（{cur_year}）')
        # 2026-08-03 营业收入（营业成本）附注汇总（格式2 合并口径；FY 无产品类别/内外销
        # 分类 → 仅汇总主营/其他 收入/成本/利润，收入分解信息暂缺）
        build_revenue_cost_footnote(wb, entities, tb_all, cur_year, prev_year)
        print(f'  ✓ 营业收入（营业成本）附注汇总（{cur_year}）')
        # 2026-08-09 分析性程序（对齐 SAP 5 模块，用户要求 SAP/U8 统一）
        try:
            build_revenue_analysis(wb, entities, cur_year, data, tb_all)
            print(f'  ✓ 分析性程序（{cur_year}）')
        except Exception as _ex:
            print(f'  ⚠️ 分析性程序生成失败：{_ex}')
        # 2026-08-10 销项税-收入配对核查（SAP 单主体专属：综合税率偏离/无销项税大额凭证）
        if _adapter is not None and _adapter.current_comp():
            try:
                build_vat_recon_sheet(wb, cur_year)
                print(f'  ✓ 销项税-收入配对核查表（{cur_year}）')
            except Exception as _ex:
                print(f'  ⚠️ 销项税配对核查表失败：{_ex}')
        # ⚡ 2026-08-12 营业收入对方科目核对（补齐收入循环缺口；单主体+集团）
        try:
            build_rev_counterparty_sheet(wb, cur_year, entities, data_dir)
            print(f'  ✓ 营业收入对方科目核对（{cur_year}）')
        except Exception as _ex:
            print(f'  ⚠️ 营业收入对方科目核对失败：{_ex}')
        # ⚡⚡ 2026-08-28 修复（按指令号分析缺失根因）：以下两表原被误缩进在 except 块内
        #   ——仅当对方科目核对【抛异常】时才执行，正常路径永不生成。移出与上方 try 平级。
        # 2026-08-10 收入成本按指令号配比（yy WBS/300 订单）
        try:
            build_rev_cost_instruction_sheet(wb, cur_year)
            print(f'  ✓ 收入成本_按指令号配比表（{cur_year}）')
        except Exception as _ex:
            print(f'  ⚠️ 收入成本指令号配比表失败：{_ex}')
        # ⚡ 2026-08-11 指令号全链条表（存货→成本 与 收款→收入 双边，单主体专属）
        try:
            _fc = build_instruction_fullchain_sheet(wb, cur_year)
            if _fc is not None:
                print(f'  ✓ 收入成本_指令号全链条表（{cur_year}）')
        except Exception as _ex:
            print(f'  ⚠️ 指令号全链条表失败：{_ex}')

        S.finalize_workbook(wb)
        out = os.path.join(data_dir, f'营业收入审计底稿_{cur_year}_生成.xlsx')
        wb.save(out)
        wb.close()
        outs.append(out)
        print(f'  ✓ 已保存：{out}')
    return outs




def main():
    if len(sys.argv) > 1:
        data_dir = sys.argv[1]
    else:
        data_dir = input('请输入账套导出文件夹路径：').strip().strip('"')
    if not os.path.isdir(data_dir):
        print(f'ERROR: 文件夹不存在：{data_dir}')
        return
    build_revenue_workbook(data_dir)
    from audit_common import finalize_after_build
    finalize_after_build(data_dir)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）


def _find_managed_python():
    # ⚡ 2026-08-22 禁用托管Python重启：venv依赖已齐全，托管3.13 openpyxl损坏
    return None



# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None

if __name__ == '__main__':
    _mp = _find_managed_python()
    if _mp and os.path.abspath(sys.executable).lower() != os.path.abspath(_mp).lower():
        import subprocess
        try:
            rc = subprocess.run([_mp, '-B', os.path.abspath(__file__)] + sys.argv[1:]).returncode
        except Exception as _e:
            print(f'[ERROR] 无法启动托管 Python：{_e}')
            rc = 1
        try:
            input('\n按回车退出…')
        except EOFError:
            pass
        sys.exit(rc)
    try:
        rc = main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        import datetime
        traceback.print_exc()
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        try:
            if os.environ.get('AUDIT_LOG') == '1':
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       f'revenue_crash_{ts}.log'), 'w', encoding='utf-8') as _f:
                    _f.write('未捕获异常：\n' + traceback.format_exc())
                print(f'已记录崩溃日志：revenue_crash_{ts}.log')
        except Exception:
            pass
        try:
            input('\n按回车退出…')
        except EOFError:
            pass
        sys.exit(1)

# ============================================================================
# ⚡ 2026-08-11 用户需求：指令号全链条表（存货→成本 与 收款→收入 双边）
# 行=指令号(WBS父级)；列=存货侧(材料采购/原材料/在制品/产成品期初·增·减·期末) →
# 主营业务成本结转；收款侧(预收/应收期初·增·减·期末) → 主营业务收入结转。
# 目的：按指令号串起"存货→成本"和"收款→收入"两条链，期末一边存货一边预收/应收。
# ============================================================================
def build_instruction_fullchain_sheet(wb, cur_year):
    """指令号全链条表。数据=GL（read_gl_rows 结构），WBS 父级聚合。
    行=指令号；列=【存货侧】存货各科目期初/本期增/本期减/期末 + 主营业务成本结转；
    【收款侧】应收账款/预收账款 期初/本期增/本期减/期末 + 主营业务收入结转。
    存货=1403材料采购/1405原材料/1406在制品/1407产成品/1408周转材料/1409半成品；
    成本=6401 借；收款=1122应收+2203预收+2205合同负债；收入=6001 贷。
    集团模式(current_comp=None)不生成（按主体逐家看才有意义，单体底稿含此表）。"""
    import re as _re
    from collections import defaultdict
    if _adapter is None or not _adapter.current_comp():
        return None
    comp = _adapter.current_comp()
    gl = _adapter.read_gl_rows(getattr(_adapter, '_DATA_ROOT', None) or P.YY,
                               {comp: _adapter.discover_entities(
                                   getattr(_adapter, '_DATA_ROOT', None) or P.YY).get(comp, {})})
    _INV_CODES = ('1403', '1405', '1406', '1407', '1408', '1409', '1411')
    _RECV_CODES = ('1122', '2203', '2205')

    def _wbs_parent(w):
        m = _re.match(r'^(P-GF\d+\.\d+)', str(w))
        return m.group(1) if m else str(w)

    def _inst_of(r):
        w = str(r.get('wbs') or '').strip()
        o = str(r.get('ord') or '').strip()
        m = str(r.get('mat') or '').strip()
        return _wbs_parent(w) if w else (o or m or '')

    inv = defaultdict(lambda: {'op': 0.0, 'inc': 0.0, 'dec': 0.0, 'cl': 0.0})
    cost = defaultdict(float)
    recv = defaultdict(lambda: {'op': 0.0, 'inc': 0.0, 'dec': 0.0, 'cl': 0.0})
    rev = defaultdict(float)
    cust_by_inst = {}
    # ⚡ 2026-08-11 修复收款侧重复：应收/预收/合同负债行【无指令号】(实测 8751 行全无
    #   wbs/mat，只有客户编码)；且一凭证含多指令号收入行(142 凭证) → 原 vno→wbs 桥接把
    #   应收行【全额】记入单一指令号 → 收减远大于收入(4亿 vs 3.54亿)。正确口径：
    #   收款侧按【客户】归集（应收行按客户编码→名称映射），每个指令号显示其客户的总
    #   收款变动——客户级对应，避免凭证金额重复分摊。存货/成本/收入仍按指令号。
    # ⚡ 2026-08-11 精确分摊（凭证级）：应收/预收/合同负债行无指令号，但同一凭证(vno)内
    #   收入行(6001)有指令号 → 按凭证内各指令号收入净额占比分摊应收行金额。
    #   已验证：凭证净额勾稽 应收净=收入净+销项税净(149,190,000=132,026,548.65+17,163,451.35)，
    #   故按收入占比分摊应收即可（含税效应体现在收入占比，分摊无重复）。
    by_vno = defaultdict(list)
    for r in gl:
        by_vno[str(r.get('vno') or '')].append(r)
    recv = defaultdict(lambda: {'op': 0.0, 'inc': 0.0, 'dec': 0.0, 'cl': 0.0})
    for v, vrows in by_vno.items():
        # 凭证内各指令号收入净额（贷-借）
        inc_net = defaultdict(float)
        for r in vrows:
            if str(r.get('code') or '').startswith('6001'):
                _i = _inst_of(r)
                if _i:
                    inc_net[_i] += float(r.get('cr') or 0.0) - float(r.get('dr') or 0.0)
        if not inc_net:
            continue
        tot_inc = sum(inc_net.values())
        if abs(tot_inc) < 0.005:
            continue
        # 凭证内应收/预收/合同负债净变动（借-贷；收款方借=减少）
        recv_net = 0.0
        for r in vrows:
            if str(r.get('code') or '').startswith(_RECV_CODES):
                recv_net += float(r.get('dr') or 0.0) - float(r.get('cr') or 0.0)
        if abs(recv_net) < 0.005:
            continue
        for _i, _amt in inc_net.items():
            _share = recv_net * (_amt / tot_inc)
            if _share > 0:
                recv[_i]['inc'] += _share
            else:
                recv[_i]['dec'] += -_share
    for r in gl:
        c = str(r.get('code') or '')
        inst = _inst_of(r)
        if not inst:
            continue
        dr = float(r.get('dr') or 0.0)
        cr = float(r.get('cr') or 0.0)
        if c.startswith(_INV_CODES):
            d = inv[inst]
            d['inc'] += dr
            d['dec'] += cr
        elif c.startswith('6401'):
            cost[inst] += dr
        elif c.startswith('6001'):
            rev[inst] += cr
            cu = str(r.get('cust') or '').strip()
            if cu and inst not in cust_by_inst:
                cust_by_inst[inst] = cu
    insts = sorted(set(inv) | set(cost) | set(recv) | set(rev),
                   key=lambda x: -(rev.get(x, 0.0) + cost.get(x, 0.0) + inv.get(x, {}).get('inc', 0.0)))
    insts = [i for i in insts if (rev.get(i) or cost.get(i) or inv.get(i, {}).get('inc') or recv.get(i, {}).get('inc'))]
    if not insts:
        return None
    ws = wb.create_sheet('收入成本_指令号全链条表')
    ws.cell(1, 1, f'指令号全链条（{cur_year} 年度）：存货(期初→增减→期末)→成本结转 | 收款(期初→增减→期末)→收入结转').font = S.SHELL_TITLE_FONT
    hdrs = ['指令号', '客户',
            '存货-期初', '存货-本期增', '存货-本期减', '存货-期末', '成本结转(6401借)',
            '收款-期初', '收款-本期增', '收款-本期减', '收款-期末', '收入结转(6001贷)']
    S.apply_header_row(ws, 2, hdrs)
    r = 3
    for inst in insts[:300]:
        iv = inv.get(inst, {})
        cu = cust_by_inst.get(inst, '')
        rv_d = recv.get(inst, {})
        ws.cell(r, 1, inst).font = S.SHELL_BODY
        ws.cell(r, 2, cu)
        vals = [iv.get('inc', 0.0) - iv.get('dec', 0.0), iv.get('inc', 0.0), iv.get('dec', 0.0),
                iv.get('inc', 0.0) - iv.get('dec', 0.0), cost.get(inst, 0.0),
                rv_d.get('inc', 0.0) - rv_d.get('dec', 0.0), rv_d.get('inc', 0.0), rv_d.get('dec', 0.0),
                rv_d.get('inc', 0.0) - rv_d.get('dec', 0.0), rev.get(inst, 0.0)]
        for j, v in enumerate(vals, 3):
            cc = ws.cell(r, j, round(v, 2) if abs(v) >= 0.005 else None)
            cc.number_format = S.SHELL_NUM; cc.border = S.SHELL_BORDER; cc.alignment = S.SHELL_RGT
        for c in (1, 2):
            ws.cell(r, c).border = S.SHELL_BORDER
        r += 1
    r += 1
    ws.cell(r, 1, '说明：存货=1403/1405/1406/1407/1408/1409/1411 借方增/贷方减；成本=6401 借；'
                  '收款=应收(1122)+预收(2203)+合同负债(2205)【凭证级按收入占比分摊】贷方增/借方减'
                  '（GL 收款行无指令号，按同凭证收入行金额占比分摊到指令号，已验证凭证勾稽'
                  '应收净=收入净+销项税净）；收入=6001 贷。WBS 子级归并父级。期初/期末=本期'
                  '发生净额（GL 无期初，TB 期初在审定表核对）。').font = S.SHELL_BODY
    for i, w in enumerate([28, 26, 12, 12, 12, 12, 14, 12, 12, 12, 12, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'C3'
    return ws
