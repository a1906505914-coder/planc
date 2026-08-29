# -*- coding: utf-8 -*-
"""AH 88 家内部往来核对（配对法，2026-08-18 用户需求修正版）。

数据源：
  · AH/prepared/公司代码.xlsx              —— 账套代码→公司全名（SAP 导出 `|代码|名称>|`）
  · AH/prepared/{2468,1010,1357}/        —— 三集团 6 往来科目审计底稿明细表

识别"内部往来"：往来单位名称与 137 家公司全名做包含匹配（长名优先），命中=对方为集团内主体。

对平逻辑（配对法，正确口径）：
  对集团内每对主体 (X, Y)：
    · X 应收 Y（X 记录 Y 欠 X） 应 ≈ Y 应付 X（Y 记录欠 X）       → 对账点①
    · X 应付 Y（X 记录欠 Y）   应 ≈ Y 应收 X（Y 记录 X 欠 Y）       → 对账点②
  双向净额 = (X应收Y − X应付Y) + (Y应收X − Y应付X) ≈ 0

输出（AH/prepared/88家内部往来核对_2026.xlsx）：
  ① 内部往来逐对核对（核心）：X/Y 各往来科目期初/期末金额展开（含镜像差）+ 状态
  ② 说明
  （2026-08-27 精简：删除 内部往来逐科目明细/内部往来汇总/内部交易核对 三张汇总 sheet，主表已含全部科目信息）
"""
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import paths as P

PREP = os.path.join(P.DATA_DIRS['AH'], '中间产物', 'prepared')
CODE_FP = os.path.join(PREP, '公司代码.xlsx')
GROUPS = ['2468', '1010', '1357']
SUBJS = ['应收账款', '应付账款', '其他应收款', '其他应付款', '预付账款', '预收账款',
         '合同资产', '合同负债']
# 经营组：应收侧=应收/预收/合同资产/合同负债（销售类），应付侧=应付/预付（采购类）
BIZ_RECV = ('应收账款', '预收账款', '合同资产', '合同负债')
BIZ_PAY = ('应付账款', '预付账款')
ASSET = ('应收账款', '预付账款', '其他应收款', '合同资产')   # 资产科目
LIAB = ('应付账款', '预收账款', '其他应付款', '合同负债')     # 负债科目

F = Font(name='微软雅黑', size=10)
FH = Font(name='微软雅黑', size=10, bold=True)
FILL = PatternFill('solid', fgColor='DDEBF7')
FILL_BAD = PatternFill('solid', fgColor='FFC7CE')
FILL_OK = PatternFill('solid', fgColor='E2EFDA')
THIN = Border(*[Side(style='thin', color='BBBBBB')] * 4)
NUM = '#,##0.00'


def load_company_map():
    """公司代码.xlsx → {代码: 全名}。
    2026-08-18 归一：『一分/二分』分公司并入主名（如 杭州杭氧物资有限公司一分 → 杭州杭氧物资有限公司）。"""
    out = {}
    wb = openpyxl.load_workbook(CODE_FP, read_only=True, data_only=True)
    for r in wb['公司代码'].iter_rows(values_only=True):
        for cell in r:
            m = re.match(r'\|(\d{4})\|([^|]+)', str(cell or '').strip())
            if m:
                nm = m.group(2).strip().rstrip('>').strip()
                # ⚡ 分公司后缀归一：一分/二分 并入主公司
                nm = re.sub(r'(一分|二分)$', '', nm).strip()
                if nm:
                    out[m.group(1)] = nm
    return out


def load_detail(fp):
    """往来科目明细表 → [{comp, name, begin, end, dr, cr}]。"""
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception:
        return []
    sn = next((s for s in wb.sheetnames if '明细' in s and '2026' in s),
              next((s for s in wb.sheetnames if '明细' in s), None))
    if not sn:
        return []
    ws = wb[sn]
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    for i, r in enumerate(rows[:12]):
        if any(str(x or '').strip() == '往来单位名称' for x in r):
            hi = i
            break
    if hi is None:
        return []
    hdr = [str(x or '') for x in rows[hi]]

    def ci(name):
        return hdr.index(name) if name in hdr else None

    c_comp, c_name = ci('核算主体'), ci('往来单位名称')
    c_begin = ci('期初余额(人民币)')          # ⚡ 2026-08-19 期初核对
    c_end = ci('期末余额(人民币)')
    c_dr = ci('本期借方发生额')
    c_cr = ci('本期贷方发生额')
    if c_comp is None or c_name is None or c_end is None:
        return []
    out = []
    for r in rows[hi + 1:]:
        nm = str(r[c_name] or '').strip()
        if not nm or nm in ('小计', '合计'):
            continue
        try:
            end = float(r[c_end] or 0)
        except (TypeError, ValueError):
            continue
        try:
            begin = float(r[c_begin] or 0) if c_begin is not None else 0.0
        except (TypeError, ValueError):
            begin = 0.0
        try:
            dr = float(r[c_dr] or 0) if c_dr is not None else 0.0
        except (TypeError, ValueError):
            dr = 0.0
        try:
            cr = float(r[c_cr] or 0) if c_cr is not None else 0.0
        except (TypeError, ValueError):
            cr = 0.0
        out.append({'comp': str(r[c_comp] or '').strip(),
                    'name': nm, 'begin': begin, 'end': end, 'dr': dr, 'cr': cr})
    return out


def main():
    cmap = load_company_map()
    name2code = {v: k for k, v in cmap.items()}
    full = sorted(name2code, key=len, reverse=True)

    def match_code(nm):
        for c in full:
            if len(c) >= 6 and (c in nm or nm in c):
                return c
        return None

    # pair[(my_name, their_name)] = {end: {subj: val}, begin: {subj: val}, dr: {...}, cr: {...}}
    # 键用公司全名（同一公司多账套自动归并，如 1010/1270 均=杭氧集团）；跳过自我匹配与零金额
    pair = defaultdict(lambda: {'end': defaultdict(float), 'begin': defaultdict(float),
                                'dr': defaultdict(float), 'cr': defaultdict(float)})
    n_inner = 0
    for g in GROUPS:
        for subj in SUBJS:
            # ⚡ 合同资产/合同负债底稿的 2468 目录是错误重复（混入 1010/1030/1050/1060/1100/3010，
            #    与 1357 目录内容重复）。按集团定义：2468=2/4/6/8开头气体公司，1010=仅1010，1357=其余。
            #    合同资产数据实际归属 1010目录(1010) + 1357目录(1030/1050/1060/1100/3010)，故跳过 2468。
            if subj in ('合同资产', '合同负债') and g == '2468':
                continue
            fp = os.path.join(PREP, g, f'{subj}审计底稿_{g}.xlsx')
            if not os.path.exists(fp):
                continue
            for x in load_detail(fp):
                tc = match_code(x['name'])
                if tc is None:
                    continue
                mn = cmap.get(x['comp'], '')
                if not mn or mn == tc:          # 我方代码无映射 / 同公司自我匹配
                    continue
                p = pair[(mn, tc)]
                p['end'][subj] = round(p['end'][subj] + x['end'], 2)
                p['begin'][subj] = round(p['begin'][subj] + x['begin'], 2)
                if x['dr']:
                    p['dr'][subj] = round(p['dr'][subj] + x['dr'], 2)
                if x['cr']:
                    p['cr'][subj] = round(p['cr'][subj] + x['cr'], 2)
                n_inner += 1
    # 过滤零金额对子
    pair = {k: v for k, v in pair.items()
            if any(abs(x) > 0.005 for x in v['end'].values())
            or any(abs(x) > 0.005 for x in v['dr'].values())
            or any(abs(x) > 0.005 for x in v['cr'].values())}
    if not pair:
        print('未识别到内部往来')
        return 1
    print('内部往来行:', n_inner, '| 对子:', len(pair))

    def _split(bal):
        """按方向拆分归类（2026-08-18 修正符号语义）。
        ⚡ 底稿明细表『期末余额(人民币)』列：资产科目 正=借余(对方欠我)、负=贷余(我欠对方)；
           负债科目 正=借余(对方欠我)、负=贷余(我欠对方) —— 即『借余为正、贷余为负』，
           与科目余额表贷方余额带负号的口径一致（期末=期初+借-贷）。
        对方欠我 = 资产正值 + 负债正值（借余=对方欠我）
        我欠对方 = 资产负值(绝对值) + 负债负值(绝对值)（贷余=我欠对方）"""
        they_owe = owe = 0.0
        for s in ASSET:
            v = bal.get(s, 0.0)
            they_owe += max(v, 0.0)          # 资产借余 → 对方欠我
            owe += max(-v, 0.0)              # 资产贷余 → 我欠对方
        for s in LIAB:
            v = bal.get(s, 0.0)
            owe += max(-v, 0.0)              # 负债贷余(负) → 我欠对方
            they_owe += max(v, 0.0)          # 负债借余(正) → 对方欠我
        return round(they_owe, 2), round(owe, 2)

    wb = openpyxl.Workbook()

    # ===== Sheet1 内部往来逐对核对 =====
    ws = wb.active
    ws.title = '内部往来逐对核对'
    # 列：主体 + 期初/期末 X/Y 8 个具体往来科目展开（2026-08-27 用户要求：先把 X、Y 包含的
    #     各科目金额列示出来，再给镜像差核对）+ 状态/一致性/穿透。列序被 ah_intra_diff_trace
    #     _load_pairs 依赖：期初差 col18-21、期末差 col38-41、净额 col42、状态 col43、
    #     一致性 col44、穿透 col45。
    SUBJ_ORDER = ['应收账款', '预收账款', '合同资产', '合同负债',
                  '应付账款', '预付账款', '其他应收款', '其他应付款']
    # ⚡⚡ 2026-08-27 对子序号（第 0 列）：全表统一索引，供差异分录/核对一致/勾稽校验等
    #   输出表通过『对子序号』串联（ah_intra_diff_trace._load_pairs 读 r[0] 作为 no）。
    H1 = ['对子序号', '我方主体', '对方主体']
    for s in SUBJ_ORDER:
        H1.append(f'期初 {s}(X)')
    for s in SUBJ_ORDER:
        H1.append(f'期初 {s}(Y)')
    H1 += ['期初经营差①', '期初经营差②', '期初资金池差①', '期初资金池差②']
    for s in SUBJ_ORDER:
        H1.append(f'期末 {s}(X)')
    for s in SUBJ_ORDER:
        H1.append(f'期末 {s}(Y)')
    H1 += ['经营镜像差①', '经营镜像差②', '资金池镜像差①', '资金池镜像差②',
           '双向净额', '状态', '差异一致性(期初vs期末)', '穿透需求']
    ws.append(H1)
    for c in ws[1]:
        c.font = FH
        c.fill = FILL
        c.border = THIN
    # 汇总每对（含双向）：按 (min,max) 合并
    rows_out = []
    keys = sorted(pair)
    done = set()
    for (mx, tx) in keys:
        if (tx, mx) in done:
            continue
        done.add((mx, tx))
        p_xy = pair.get((mx, tx))
        p_yx = pair.get((tx, mx))
        bx = {s: p_xy['end'].get(s, 0) for s in SUBJS} if p_xy else {s: 0 for s in SUBJS}
        by = {s: p_yx['end'].get(s, 0) for s in SUBJS} if p_yx else {s: 0 for s in SUBJS}
        # ⚡ 2026-08-18 期初核对：明细底稿『期初余额(人民币)』列，与期末同口径（借余为正/贷余为负）
        bx0 = {s: p_xy['begin'].get(s, 0) for s in SUBJS} if p_xy else {s: 0 for s in SUBJS}
        by0 = {s: p_yx['begin'].get(s, 0) for s in SUBJS} if p_yx else {s: 0 for s in SUBJS}
        # ---- 经营组：应收预收合同资产合同负债 ↔ 应付预付（销售/采购类镜像）----
        # ⚡⚡ 2026-08-27 修复（用户指正经营差异期初数需重新核对——经营组与资金池同类符号 bug）：
        #   底稿余额符号口径 = 资产『借正贷负』、负债『贷正借负』。
        #   BIZ_RECV 里的预收/合同负债为贷正=我欠对方，在『对方欠我方』应收侧应取负；
        #   BIZ_PAY 里的预付为借正=对方欠我方，在『我方欠对方』应付侧应取负，应付贷正=对方欠我方（正）。
        #   '对方欠我方'净额：应收+合同资产−预收−合同负债；应付−预付。
        #   镜像差 = X应收侧净额 − Y应付侧净额（同为『Y欠X』，同向相减；原相加虚增一倍——
        #   临汾↔吕梁 临汾应付234.75万=吕梁应收234.75万 完全对平，原公式 d2=469.5万 假差异）。
        def _recv_net(bal):
            return sum(bal[s] for s in ('应收账款', '合同资产')) \
                 - sum(bal[s] for s in ('预收账款', '合同负债'))
        def _pay_net(bal):
            return bal['应付账款'] - bal['预付账款']
        a_xy = round(_recv_net(bx), 2)      # X 对 Y：对方欠我方净额（应收侧）
        l_yx = round(_pay_net(by), 2)       # Y 对 X：对方欠我方净额（应付侧）
        l_xy = round(_pay_net(bx), 2)       # X 对 Y：我方欠对方净额（应付侧）
        a_yx = round(_recv_net(by), 2)      # Y 对 X：我方欠对方净额（应收侧）
        d1 = round(a_xy - l_yx, 2)
        d2 = round(l_xy - a_yx, 2)
        # ---- 资金池组：其他应收 ↔ 其他应付（集团资金归集镜像）----
        # ⚡⚡ 2026-08-27 修复（用户指正"期初资金池差异不可能那么多"）：
        #   底稿余额符号口径 = 资产『借正贷负』、负债『贷正借负』（current_account_detail.py 统一）。
        #   其他应收款（资产）借余=正、其他应付款（负债）贷余=正——两侧方向本就相反，
        #   镜像差必须【相减】。原公式相加把两个正数叠加，虚增一倍
        #   （化医↔集团 期初资金池差 9.35亿 = 化医其他应收 4.68亿 + 集团其他应付 4.68亿，
        #    实际镜像仅差 20 万）。与下方 net 公式「- 其他应付款」的贷正口径保持一致。
        p1 = round(bx['其他应收款'] - by['其他应付款'], 2)
        p2 = round(bx['其他应付款'] - by['其他应收款'], 2)
        net = round((a_xy - l_xy + bx['其他应收款'] - bx['其他应付款'])
                    + (a_yx - l_yx + by['其他应收款'] - by['其他应付款']), 2)
        # ---- 期初镜像差（同公式，用期初余额）----
        a_xy0 = round(_recv_net(bx0), 2)
        l_yx0 = round(_pay_net(by0), 2)
        l_xy0 = round(_pay_net(bx0), 2)
        a_yx0 = round(_recv_net(by0), 2)
        d10 = round(a_xy0 - l_yx0, 2)
        d20 = round(l_xy0 - a_yx0, 2)
        p10 = round(bx0['其他应收款'] - by0['其他应付款'], 2)
        p20 = round(bx0['其他应付款'] - by0['其他应收款'], 2)
        # ⚡⚡ 2026-08-27 用户指正：资金池镜像差①≈②（p1-p2≈0 = 双向资金池净额≈0）时，
        #   资金池【实质对平】——p1=p2≠0 是科目重分类：同一笔资金池在 X/Y 挂不同科目方向
        #   （如化医↔集团：化医应收集团 5.454亿 = 集团应付化医 5.452亿 + 集团其他应收贷余 20万，
        #    两边总额一致，集团把 20万 挂在其他应收贷方而非应付里）。此时不再标"资金池差异"，
        #    仅当 p1、p2 方向不一致（真实单边，如制氧机 -2312万）才算资金池差异。
        pool_ok = (abs(p1) <= 1.0 and abs(p2) <= 1.0) or abs(p1 - p2) <= 1.0
        biz_ok = abs(d1) <= 1.0 and abs(d2) <= 1.0
        st = '对平' if (pool_ok and biz_ok) else ('资金池差异' if not pool_ok else '经营差异')

        # ---- 差异一致性 + 穿透需求（2026-08-18 用户方法论）----
        #  期初差异 vs 期末差异：一致 → 极大概率期初差异延续到期末（也可能巧合，标注人工判断）；
        #  不一致 → 本期发生造成差异 → 需穿透序时账找单边凭证。
        #  两边往来全程对平 → 无需穿透。
        biz_b0 = abs(d10) > 1.0 or abs(d20) > 1.0        # 期初经营有差
        biz_bt = abs(d1) > 1.0 or abs(d2) > 1.0          # 期末经营有差
        if not biz_b0 and not biz_bt:
            consistency = '期初平/期末平'
            need = '无需穿透（全程对平）'
        elif biz_b0 and biz_bt and abs(d10 - d1) <= 1.0 and abs(d20 - d2) <= 1.0:
            consistency = '一致（期初延续）'
            need = '无需穿透（期初差异延续，可能巧合）'
        elif biz_b0 and biz_bt:
            consistency = '不一致（本期变化）'
            need = '需穿透（本期发生形成差异）'
        elif not biz_b0 and biz_bt:
            consistency = '期初平/期末不平'
            need = '需穿透（本期新增差异）'
        else:
            consistency = '期初不平/期末平'
            need = '本期已平（可选穿透验证）'
        row = [mx, tx]
        for s in SUBJ_ORDER:
            row.append(round(bx0[s], 2))            # 期初 X 侧 8 科目
        for s in SUBJ_ORDER:
            row.append(round(by0[s], 2))            # 期初 Y 侧 8 科目
        row += [d10, d20, p10, p20]                 # 期初镜像差
        for s in SUBJ_ORDER:
            row.append(round(bx[s], 2))             # 期末 X 侧 8 科目
        for s in SUBJ_ORDER:
            row.append(round(by[s], 2))             # 期末 Y 侧 8 科目
        row += [d1, d2, p1, p2]                     # 期末镜像差
        row += [net, st, consistency, need]
        rows_out.append(row)
    # 期末镜像差列（rows_out 内部 col38-41）取最大绝对值排序；序号=排序后行号（全表统一索引）
    rows_out.sort(key=lambda r: -max(abs(r[38]), abs(r[39]), abs(r[40]), abs(r[41])))
    for i, row in enumerate(rows_out):
        ws.append([i + 1] + row)            # ⚡ 2026-08-27 第 0 列=对子序号（全表唯一，供串联）
        for c in ws[ws.max_row]:
            c.font = F
            c.border = THIN
            if isinstance(c.value, (int, float)) and c.column > 3:
                c.number_format = NUM
        ST = 43  # rows_out 内部状态列索引（主体2+期初科目16+期初差4+期末科目16+期末差4+净额1 → 43）
        if row[ST] == '对平':
            for c in ws[ws.max_row]:
                c.fill = FILL_OK
        elif row[ST] == '资金池差异':
            for c in ws[ws.max_row]:
                c.fill = PatternFill('solid', fgColor='FFD966')   # 黄=资金池差异
        else:
            for c in ws[ws.max_row]:
                c.fill = FILL_BAD                              # 红=经营差异
    ws.freeze_panes = 'D2'   # 冻结序号+两主体列与表头（序号列 C? 实际序号在A，主体B/C → 冻结D2）
    # 列宽 + 分组表头着色（期初块=浅蓝、期末块=浅绿，便于区分；A=序号 B/C=主体）
    from openpyxl.utils import get_column_letter as _gcl
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 26
    _BLK1 = PatternFill('solid', fgColor='DDEBF7')   # 期初 X/Y 科目（D~S）
    _BLK2 = PatternFill('solid', fgColor='E2EFDA')   # 期末 X/Y 科目（X~AN）
    for col in range(4, ws.max_column + 1):
        ws.column_dimensions[_gcl(col)].width = 13
        if 4 <= col <= 19:
            ws.cell(1, col).fill = _BLK1
        elif 24 <= col <= 39:
            ws.cell(1, col).fill = _BLK2
    n_ok = sum(1 for r in rows_out if r[ST] == '对平')
    n_pool_diff = sum(1 for r in rows_out if r[ST] == '资金池差异')
    n_biz_diff = sum(1 for r in rows_out if r[ST] == '经营差异')
    n_need_trace = sum(1 for r in rows_out if r[45].startswith('需穿透'))
    ws.append([])
    ws.append([f'内部往来对子 {len(rows_out)} 对；对平 {n_ok} 对，资金池差异 {n_pool_diff} 对，经营差异 {n_biz_diff} 对；需穿透 {n_need_trace} 对'])

    # ===== Sheet4 说明 =====
    ws4 = wb.create_sheet('说明')
    notes = [
        '【88 家内部往来核对说明（配对法 · 方向归类）】',
        '1. 识别：往来单位名称与《公司代码.xlsx》137 家公司全名做包含匹配（长名优先），命中=对方为集团内主体；九江萍钢/山西晋南等外部单位不进入。',
        '2. 所有往来科目打包：6 个科目全部纳入（资产：应收账款/预付账款/其他应收款；负债：应付账款/预收账款/其他应付款），X对Y 与 Y对X 两侧逐科目列示（带符号）。',
        '3. 正负号约定：底稿明细表『期末余额(人民币)』列=资产『借正贷负』、负债『贷正借负』（负债贷方余额显示为正）。资产科目 正=对方欠我方、负=反向(我方欠对方)；负债科目 正=贷余(我方欠对方)、负=借余(对方欠我方)。',
        '4. 【经营组】应收/预收/合同资产/合同负债 打包（销售类）↔ 应付/预付（采购类）镜像核对（2026-08-27 修正符号口径）：资产借正=对方欠我方（+）、负债贷正=我方欠对方（−）。应收侧净额=应收+合同资产−预收−合同负债；应付侧净额=应付−预付（同为『对方欠我方』正口径）。经营镜像差=应收侧净额−应付侧净额≈0=经营对平（原公式相加把方向相反的贷正/借正叠加，虚增一倍）。',
        '5. 【资金池组】其他应收 ↔ 其他应付 镜像核对：集团资金归集（资金下拨/上收）体现在这两个科目。其他应收借正(对方欠X)、其他应付贷正(Y欠对方)——两侧方向相反，镜像差=其他应收−其他应付≈0=资金池对平（2026-08-27 修正：原公式相加虚增一倍）。⚡ 资金池镜像差①≈②（p1−p2≈0=双向净额≈0）即资金池实质对平：p1=p2≠0 是科目重分类（同一笔资金池在 X/Y 挂不同科目方向，如集团记其他应收贷方 vs 子公司记其他应收借方），不标差异。',
        '6. 状态：资金池与经营均对平=对平（绿）；仅资金池不平=资金池差异（黄）；仅经营不平=经营差异（红，审计重点关注）。',
        '7. 双向净额 = 全部 6 科目(X对Y + Y对X)带符号之和，应≈0。',
        '8. 交易口径：销售=应收本期借方发生额；采购=应付本期贷方发生额（含税）。',
        '9. 数据源：AH/prepared/{2468,1010,1357} 三集团 6 往来科目审计底稿明细表 + 公司代码.xlsx。',
    ]
    for n in notes:
        ws4.append([n])
        ws4[ws4.max_row][0].font = F

    out = os.path.join(PREP, '88家内部往来核对_2026.xlsx')
    wb.save(out)
    print(f'已生成：{out}')
    print(f'  对子 {len(rows_out)} 对 | 对平 {n_ok} | 资金池差异 {n_pool_diff} | 经营差异 {n_biz_diff}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
