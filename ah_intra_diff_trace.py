# -*- coding: utf-8 -*-
"""AH88家 关联往来差异穿透底稿 v3（2026-08-18 用户定：分层核对 + 期初纳入 + 单表 + 自检）。

核心逻辑（用户方法论）：
  ① 往来明细找双边数据 → ② 区分期初核对与期末核对 → ③ 期初差=期末差=期初延续(可能巧合，不穿透)，
     不一致=本期变化(需穿透) → ④ 序时账匹配(借鉴网银，无15天硬窗口)留单边 → ⑤ 勾稽 期初+本期=期末
     → ⑥ 全程对平不穿透。

结果评判（用户定）：差异凭证数量应非常有限（一般不超过 10 多笔）；
  若某对单边笔数偏多（>MAX_UNIL_PER_PAIR）或两边仍有多笔金额，说明匹配不充分，需进一步核对。
  → 自检报告（Sheet6）自动挑出：单边过多 / 勾稽差≠0 的对子。

输出（单表交付）：AH/prepared/关联往来差异穿透底稿_2026.xlsx
  Sheet1 差异对对冲汇总   每对：期初/期末镜像差 + 差异一致性 + 穿透需求 + 单边 + 勾稽差额 + 自检
  Sheet2 单边入账明细     未来调整候选（凭证级，按对子分组，标侧+类型）
  Sheet3 双边对上明细     备查（X行 ↔ Y行 配对，含日期差）
  Sheet4 序时账命中明细   原独立《关联往来序时账》并入单表
  Sheet5 口径说明
  Sheet6 自检报告         单边笔数>阈值 / 勾稽差≠0 的对子清单（自动质量把关）
"""
import os
import sys
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import paths as P

PREP = os.path.join(P.DATA_DIRS['AH'], '中间产物', 'prepared')
RECON_FP = os.path.join(PREP, '88家内部往来核对_2026.xlsx')
SEQ_FP = os.path.join(PREP, '关联往来序时账_2026.xlsx')
OUT_FP = os.path.join(PREP, '关联往来差异穿透底稿_2026.xlsx')

# ⚡ 匹配窗口（天）：None=不限制（关联往来跨期常见，按企业内部核对要求调整）
MATCH_WIN_DAYS = None
AMT_TOL = 1.0      # 金额容忍 ±1 元（含税/尾差）
# ⚡ 找和窗口：拆分行/分期支付可跨期数月（透平收杭氧货款 560万 3-30 = 集团 200万 3-31 + 300万 + 60万 5月），
#    用户强调关联往来时间跨度远长于网银月度对账，找和窗口放宽到 180 天
SUM_WIN_DAYS = 180
MAX_SUM_CAND = 80    # ⚡ 找和候选数上限（性能保护，防大对子 DFS 组合爆炸 OOM）
COMBO_LIMIT = 3000000  # ⚡ 单次找和 DFS 迭代上限（超限返回 None，防卡死）

# ⚡⚡ 结果评判阈值（2026-08-18 用户定：差异凭证数量应非常有限，一般不超过 10 多笔）
#    单边笔数超过该阈值 → 自检标记『匹配不充分，需进一步核对』（两边仍有多笔金额=匹配层漏配）
MAX_UNIL_PER_PAIR = 15

# ⚡ 底稿范围参数（2026-08-20 新增）：
#    辅助核算底稿数据截至 2026-07-31，序时账为全年（含 8 月）。
#    8 月凭证 = 期后事项，不在底稿 Δd 口径内 → 提取时按 CUTOFF_DATE 剔除，避免制造假残差。
#    设为 None 表示不过滤（默认保留全年，由调用方决定）。
CUTOFF_DATE = None   # 例: '20260801' 即剔除 2026-08-01 及之后凭证

# ⚡ 往来科目 4 位前缀白名单（SAP 代码，AH 科目余额表核实）
# ⚡⚡ 2026-08-18 用户定：应收票据(1121)/应付票据(2201)【不参与】往来差异查找——
#    票据是结算工具：持票人借票据贷往来、开票人借往来贷票据，票据本身不是往来余额，
#    最终往来增减体现在应收/应付等科目上，由那些科目参与核对。
# ⚡ 2026-08-19 1123 预付属"应付侧"（预付=我方先付给对方，镜像于应付/预收，核对表 BIZ_PAY=应付预付）——
#    原误放 RECV_CODES，导致带符号单边净额方向虚增（对子4 泽州 1123 收设备发票 2859万×3 被按应收算，
#    勾稽差 -1.11亿 中近半来自此）。已移到 PAY_CODES。
# ⚡⚡ 2026-08-19 科目归类对齐核对表（用户要求"考虑期初数影响后全量重跑"）：
#    核对表口径 ah_intra_group_recon.py：BIZ_RECV=应收+预收+合同资产+合同负债；BIZ_PAY=应付+预付；
#    资金池组=其他应收(1221)+其他应付(2241)；利息/股利科目(1131/1132/2231/2232) SUBJS 不含。
#    原序时账把预收(2203)归"应付侧"、把利息/股利归经营，与核对表相反 → 单边净额与期末差口径
#    不一致 → 勾稽永远差（#4 集团预收泽州 2.69亿 被镜像校验拆开后单边净额虚增 5.38亿）。
#    修复：预收(2203)移入应收侧 RECV_CODES；利息/股利科目移出经营（核对表不含，不参与穿透）。
RECV_CODES = ('1122', '1124', '2203')            # 应收/合同资产/预收（核对表 BIZ_RECV）
PAY_CODES = ('2202', '1123')                     # 应付/预付（核对表 BIZ_PAY）
RECV_PAY = RECV_CODES + PAY_CODES


def _km_typ(km):
    """科目类型：R=应收侧 / P=应付侧 / O=其他。"""
    k = str(km)[:4]
    if k in RECV_CODES:
        return 'R'
    if k in PAY_CODES:
        return 'P'
    return 'O'


def _mirror_ok(ka, kb):
    """镜像校验：配对的两科目必须方向相反（应收↔应付 或 预收↔预付），禁止同方向配对。
    ⚡ 2026-08-19 用户要求验证"应勾稽一致没对上"：match_pool/_ms_dir 原纯金额配对，
       X借应收(确认) 会被配到 Y贷应收(回收)——同方向科目错误镜像，配后消耗单边但方向错
       （Y应收变化不计入差），产生勾稽差。双边对上 3104 行中 249 行为此类(应收↔应收128+
       应付↔应付121)，对应 B型 17 个对子勾稽差。
       ⚡ 特例：预收(2203)↔预付(1123) 是合法镜像（我收对方预收 = 对方预付我），
       两者虽同为 PAY 归类但方向相反，必须允许（#48 化医收电子气进度款 1062万 曾误拦）。"""
    ka4, kb4 = str(ka)[:4], str(kb)[:4]
    ta, tb = _km_typ(ka), _km_typ(kb)
    if {ta, tb} == {'R', 'P'}:
        return True
    if {ka4, kb4} == {'2203', '1123'}:
        return True
    return False
# ⚡ 资金池科目（集团资金归集：X 其他应收 ↔ Y 其他应付），按发生净额镜像核对不逐笔对冲
POOL_X = '1221'
POOL_Y = '2241'


def pn(v):
    """数字千分位显示。"""
    return f'{v:,.0f}' if isinstance(v, (int, float)) else str(v)

F = Font(name='微软雅黑', size=10)
FH = Font(name='微软雅黑', size=10, bold=True)
FILL = PatternFill('solid', fgColor='DDEBF7')
FILL_RED = PatternFill('solid', fgColor='FCE4EC')
THIN = Border(*[Side(style='thin', color='BBBBBB')] * 4)
NUM = '#,##0.00'


def _wc(ws, value, fill=None):
    """write_only 表头单元格。"""
    from openpyxl.cell import WriteOnlyCell
    c = WriteOnlyCell(ws, value=value)
    c.font = FH
    c.fill = fill or FILL
    c.border = THIN
    return c


def _pd(s):
    try:
        return datetime.strptime(str(s)[:8], '%Y%m%d')
    except Exception:
        return None


def _is_internal(rows):
    """凭证是否【内部科目结转】：所有行科目都在往来白名单内 且 借贷平衡。
    如 借合同资产(1124)/贷应收(1122)『本期开票结转合同资产』——往来科目间互转，
    不产生新的对外往来，集团侧本就不应有对应分录，应从单边判定中排除。
    返回 True=内部结转。"""
    kms = [str(r['km'])[:4] for r in rows]
    if not kms or not all(k in RECV_PAY for k in kms):
        return False
    dr = sum(r['dr'] for r in rows)
    cr = sum(r['cr'] for r in rows)
    return abs(dr - cr) < 0.01


def _drop_internal(rows):
    """按凭证聚合剔除内部科目结转凭证的行。返回 (保留行, 剔除行)。"""
    keep, dropped = [], []
    by_vou = defaultdict(list)
    for r in rows:
        by_vou[(r['sz'], r['date'], r['vou'])].append(r)
    for k, rs in by_vou.items():
        if _is_internal(rs):
            dropped.extend(rs)
        else:
            keep.extend(rs)
    return keep, dropped


def _is_recv(km):
    k = str(km or '')
    # ⚡ 应付暂估(220206)不参与：暂估是单方入账/后续转正冲回，非真实往来余额
    #    （凭证=借暂估486.7万+借税63.3万+贷应付550万，借暂估是内部冲销，透平侧无对应）
    if k.startswith('220206'):
        return False
    return k[:4] in RECV_PAY


def _squeeze(rows, opp_names, cmap):
    """⚡⚡ 挤水分（2026-08-19 用户定：借应付需区分性质）：
    ① 供应商/客户代码过滤：行 cust/supp 为数字代码且非【本方】也非【对方主体】→ 凭证级聚合误归
       （如 2500001763 借220201=贷220201 供应商1070 非物资），剔除。
    ② 应付内部冲销：同凭证 借2202 = 贷2202（±1元）→ 应付内部调整/重分类净额0，非往来差异，剔除。
    返回挤水后行。"""
    opp_codes = {k for k, v in cmap.items() if v in opp_names}
    out = []
    for r in rows:
        sz = str(r['sz'])
        my_code = sz[:4]
        bad = False
        for c in (str(r.get('cust') or ''), str(r.get('supp') or '')):
            c = c.strip()
            if c.isdigit() and len(c) >= 4:
                code = c[:4]
                if code != my_code and code not in opp_codes:
                    bad = True
                    break
        if not bad:
            out.append(r)
    rows = out
    # ② 同凭证 借=贷 内部闭环剔除（±1元）
    #    ⚡ 2026-08-19 扩展科目：原只 2202/2231/2232/1131/1132 → 全部往来科目（RECV_CODES+PAY_CODES）。
    #    对子4 集团 凭证6500001291 借1122=贷1122 4450万（结转开票↔冲无票收入）同凭证借=贷对称
    #    = 内部应收重分类（无票收入转开票应收），净额 0 非往来差异，剔除（用户 20:04 确认）。
    #    ⚠ 仅限"同一凭证内"借=贷对称——跨凭证/跨期的 1122 确认应收↔回收应收（对子10 教训：
    #    借9000万确认 ↔ 贷4800+4200万回收）是真实业务，不同凭证，本层不剔除。
    CLOSE_CODES = RECV_CODES + PAY_CODES
    by_vou = defaultdict(lambda: {'dr': 0.0, 'cr': 0.0, 'rows': []})
    for r in rows:
        if str(r['km'])[:4] in CLOSE_CODES:
            k = (r['sz'], r['vou'])
            by_vou[k]['dr'] += r['dr'] or 0
            by_vou[k]['cr'] += r['cr'] or 0
            by_vou[k]['rows'].append(r)
    drop = set()
    for k, v in by_vou.items():
        if v['dr'] > 0 and v['cr'] > 0 and abs(v['dr'] - v['cr']) <= 1.0:
            # ⚡ 2026-08-20 修复: 同凭证借=贷对称剔除必须排除"跨主体重分类"——
            #    服务费确认凭证 6500003551-3572"调整空分操作技术培训费":
            #      借1122 M=气体公司代码(+10,000) / 贷1122 M=1010(-10,000)
            #    借方M=对侧、贷方M=自身 → 集团把应收从自身重分类到气体公司 = 真实对子往来，
            #    不应按"内部应收重分类净0"剔除（此前误剔导致13个服务费对子残差）。
            cross = False
            for r in v['rows']:
                dr_row = r.get('dr') or 0
                cr_row = r.get('cr') or 0
                mv = str(r.get('cust') or r.get('supp') or '').strip()
                sz = str(r.get('sz') or '')
                my_code = sz[:4]
                if dr_row > 0 and mv and mv != my_code and mv in opp_codes:
                    # 借方挂对侧代码 → 真实跨主体应收，保留
                    cross = True
                    break
                if cr_row > 0 and mv and mv != my_code and mv in opp_codes:
                    cross = True
                    break
            if not cross:
                for r in v['rows']:
                    drop.add(id(r))
    # ⚡ 2026-08-19 跨凭证聚合：股利/利息类科目（非经营）同账套+同科目 借贷合计对称 → 计提↔发放闭环剔除。
    #    对子138 2021-06-10 两笔（6200000457 借2232=2080万 发放 + 6500000250 贷2232=2080万 计提）
    #    是同一股利业务跨凭证闭环（凭证日期=业务原始日，记账在2026），同凭证层剔不到 → 聚合层剔。
    #    仅限非经营科目；经营科目（1122 等）不作跨凭证聚合（防真实业务误剔）。
    by_szkm = defaultdict(lambda: {'dr': 0.0, 'cr': 0.0, 'rows': []})
    for r in rows:
        if id(r) in drop:
            continue
        km4 = str(r['km'])[:4]
        if km4 in ('2231', '2232', '1131', '1132'):
            k = (r['sz'], km4)
            by_szkm[k]['dr'] += r['dr'] or 0
            by_szkm[k]['cr'] += r['cr'] or 0
            by_szkm[k]['rows'].append(r)
    for k, v in by_szkm.items():
        if v['dr'] > 0 and v['cr'] > 0 and abs(v['dr'] - v['cr']) <= 1.0:
            for r in v['rows']:
                drop.add(id(r))
    # ⚡ 2026-08-19 权益调账分录剔除：股改清算/转未分配利润/分红款等非真实往来调账。
    #    对子22 江西 凭证6500000687 借410406(股改清算期损益)4000万 / 贷122101(其他应收)4000万
    #    ——权益重组调账，往来科目只有单侧（借方410406 非往来科目），同凭证借=贷对称规则
    #    （410406 不计入 by_vou）覆盖不到 → 资金池镜像差虚增 4000万。
    #    同业务集团端：224101 借4000万"调整结转24年收到江氧分红款"（冲其他应付）也非真实资金池。
    #    按文本关键词精确剔除，避免误剔真实业务。
    ADJ_KW = ('股改清算', '清算期损益', '转未分配利润', '分红款')
    for r in rows:
        if id(r) in drop:
            continue
        km4 = str(r.get('km') or '')[:4]
        if (km4 in CLOSE_CODES or km4 in (POOL_X, POOL_Y)) and any(k in str(r.get('txt') or '') for k in ADJ_KW):
            drop.add(id(r))
    rows = [r for r in rows if id(r) not in drop]
    return rows


def _load_pairs():
    """核对表差异对。
    ⚡ 2026-08-18 分层核对后：只穿透『需穿透』对子（期初差异≠期末差异=本期变化，或本期新增差异）。
       期初延续（差异一致）与全程对平的对子不穿透——期初延续极大概率是期初结转拖到期末（可能巧合），无需回序时账。"""
    # ⚡ 2026-08-20 保护: 核对表被归档到 _过程稿 时, 从 pair_diff_confirm.json 恢复对子列表(只读场景)
    if not os.path.exists(RECON_FP):
        alt = None
        for cand in ('tmp/pair_diff_confirm.json',
                     os.path.join(os.getcwd(), 'tmp', 'pair_diff_confirm.json'),
                     os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'tmp', 'pair_diff_confirm.json')):
            if os.path.exists(cand):
                alt = cand
                break
        if alt:
            import json as _json
            rows = _json.load(open(alt, encoding='utf-8'))
            out = []
            for r in rows:
                out.append({'no': r['no'], 'X': r['X'], 'Y': r['Y'], 'st': '',
                            'op1': 0, 'op2': 0, 'pool1': 0, 'pool2': 0,
                            'op1_0': 0, 'op2_0': 0, 'pool1_0': 0, 'pool2_0': 0,
                            'consist': '', 'need': '', 'net': 0})
            return out
        raise FileNotFoundError(f'核对表缺失: {RECON_FP}')
    wb = openpyxl.load_workbook(RECON_FP, read_only=True, data_only=True)
    ws = wb['内部往来逐对核对']
    out = []
    for r in ws.iter_rows(values_only=True):
        if r[0] is None or r[1] is None or str(r[1]).strip() == '我方主体':
            continue
        no = r[0]                                   # ⚡⚡ 2026-08-27 对子序号=核对表全局行号（4 张表串联索引）
        st = str(r[44] or '').strip()
        if st == '对平':
            continue
        need = str(r[46] or '').strip()
        if not need.startswith('需穿透'):
            continue  # ⚡ 只穿透需要穿透的对子（期初延续/本期已平/全程对平跳过）
        # ⚡⚡ 2026-08-27 列序对齐（主表含对子序号列）：新表头
        #   [0对子序号 1主体 2主体 3-18期初X/Y 8科目 19-22期初差 23-38期末X/Y 8科目
        #    39-42期末差 43净额 44状态 45一致性 46穿透]
        out.append({'no': no, 'X': str(r[1]).strip(), 'Y': str(r[2]).strip(), 'st': st,
                    'op1': r[39] or 0, 'op2': r[40] or 0, 'pool1': r[41] or 0, 'pool2': r[42] or 0,
                    'op1_0': r[19] or 0, 'op2_0': r[20] or 0,     # 期初经营差①/②
                    'pool1_0': r[21] or 0, 'pool2_0': r[22] or 0,  # 期初资金池差①/②
                    'consist': r[45] or '', 'need': need,
                    'net': r[43] or 0})
    wb.close()
    return out


def _load_seq(keys, with_opp=False):
    """流式读序时账，只保留 差异对键∩往来科目 的分录。返回 {(公司,对方): [行...]}。
    ⚡ 2026-08-27 新增 with_opp=True：为每条往来分录附带『对方科目名』（同凭证非往来科目，
       取自科目余额表 code→name 映射），供核对一致/差异明细输出——帮助审计判断该笔往来对应
       主营业务收入/其他业务收入/主营业务成本/存货/费用/长期资产等，为合并抵消分录提供依据。"""
    wb = openpyxl.load_workbook(SEQ_FP, read_only=True, data_only=True)
    ws = wb['关联往来序时账']
    all_rows = [r for r in ws.iter_rows(values_only=True)]
    wb.close()
    idx = {}
    vk = defaultdict(set)          # (comp, vou) → 全科目代码集合
    if with_opp:
        for r in all_rows:
            if r[0] == '账套':
                continue
            comp = str(r[1] or '').strip()
            vou = str(r[3] or '')
            km = str(r[4] or '')
            if km:
                vk[(comp, vou)].add(km)
    for r in all_rows:
        if r[0] == '账套':
            continue
        km = str(r[4] or '')
        if not _is_recv(km):
            continue
        date = str(r[2] or '')[:10]
        # ⚡ 不做日期过滤：序时账全部为 2026 年记账（用户 2026-08-19 明确"没有2026年之前的分录"）。
        #    凭证日期可能是业务原始日期（如股利决议 2021 年），但记账期间在 2026，属本期发生，保留。
        # ⚡ 2026-08-20: 若设置了 CUTOFF_DATE，剔除底稿范围外（期后）分录
        if CUTOFF_DATE is not None and len(date) >= 8 and date >= CUTOFF_DATE:
            continue
        comp = str(r[1] or '').strip()
        tc = str(r[10] or '').strip()
        k = (comp, tc)
        if k not in keys:
            continue
        rec = {'sz': str(r[0]), 'date': date, 'vou': str(r[3] or ''),
               'km': km, 'txt': str(r[5] or '')[:60], 'cust': str(r[6] or ''),
               'supp': str(r[7] or ''), 'dr': float(r[8] or 0), 'cr': float(r[9] or 0),
               'tc': tc, 'note': str(r[11] or '')}
        if with_opp:
            rec['opp'] = sorted(vk.get((comp, str(r[3] or '')), set()) - {km})
        idx.setdefault(k, []).append(rec)
    return idx


def match_pool(A, B):
    """对冲匹配：A(我方借方向) ↔ B(对方贷方向) 或反向。
    金额±AMT_TOL 配对（⚠ 2026-08-19 修复：原按 round(amt,2) 精确到分分组，X/Y 记账含 0.04 元
    舍入差（如 265,379.40 vs 265,379.44，气体购销月度购销镜像）即配对失败 → 大量气体购销对子
    单边虚增。改按"元"分组 + |差|≤AMT_TOL 容差，覆盖 ±1 元尾差）。
    同金额多笔按日期差近者优先；win_days=None 不限制日期。
    返回 (paired, A_left, B_left)。"""
    bya = defaultdict(list)
    for a in A:
        bya[round(a['amt'])].append(a)          # 按元分组（容差内就近）
    used = set()
    paired = []
    for k, alist in bya.items():
        cand_idx = [i for i, b in enumerate(B) if i not in used and abs(b['amt'] - k) <= AMT_TOL]
        if not cand_idx:
            continue
        for a in alist:
            ad = _pd(a['date'])
            best, best_d = None, None
            for i in cand_idx:
                b = B[i]
                # ⚡ 2026-08-19 镜像校验：X借应收↔Y贷应付 / X借应付↔Y贷应收（禁止应收↔应收、应付↔应付）
                if not _mirror_ok(a['km'], b['km']):
                    continue
                bd = _pd(b['date'])
                if MATCH_WIN_DAYS is not None and (ad and bd) and abs((ad - bd).days) > MATCH_WIN_DAYS:
                    continue
                d = abs((ad - bd).days) if (ad and bd) else 0
                if best_d is None or d < best_d:
                    best, best_d = i, d
            if best is not None:
                used.add(best)
                cand_idx.remove(best)
                paired.append((a, B[best]))
    used_a = {id(a) for a, _ in paired}
    A_left = [a for a in A if id(a) not in used_a]
    B_left = [b for i, b in enumerate(B) if i not in used]
    return paired, A_left, B_left


def _find_combo(target, cands, max_n=5, tol=1.0, limit=None):
    """在 cands(已按金额降序)中找 2..max_n 笔之和 ≈ target(±tol)。
    cands 为 dict 列表（每项含 'amt'）。返回选中元素列表，找不到 None。DFS+剪枝。
    ⚡ 迭代上限保护：组合爆炸时返回 None（防 OOM），宁可漏配不可卡死。
    ⚡ 2026-08-19 增强：①回溯式共享 picks（避免递归反复拷贝列表，防内存爆）
       ②集中结算/拆分行需 8-15 笔（如集团 4610万 付款 = 物资 15 笔应收之和），
         max_n 与 limit 按 target 金额动态传入（见 _ms_dir）。"""
    if limit is None:
        limit = COMBO_LIMIT
    lim = [limit]                              # 递归节点计数（可变，非local）
    if cands:
        min_amt = min(c['amt'] for c in cands)
    else:
        min_amt = 0.0
    if len(cands) > MAX_SUM_CAND:
        cands = cands[:MAX_SUM_CAND]           # 只保留金额最大的前 N 笔
    picks = []                                 # 共享选列表（回溯式）

    def rec(idx, need):
        lim[0] -= 1
        if lim[0] <= 0:
            return 'LIMIT'                     # 超限信号（区别于 None）
        if len(picks) >= 2 and abs(need) <= tol:
            return list(picks)
        if need < 0 or need < min_amt - tol:
            return None
        if len(picks) >= max_n:
            return None
        for i in range(idx, len(cands)):
            if cands[i]['amt'] > need + tol:
                continue
            picks.append(cands[i])
            r = rec(i + 1, need - cands[i]['amt'])
            if r:
                return r
            picks.pop()
        return None
    r = rec(0, target)
    return r if r != 'LIMIT' else None


def _ms_dir(targets, srcs):
    """单方向找和：targets 中单笔 = srcs 中 2..N 笔之和（±1元，窗口 SUM_WIN_DAYS）。
    ⚡ 2026-08-19 增强（材料款集中结算验证）：集团大额付款/应付 = 物资 8-15 笔项目应收之和，
       找和层按 target 金额动态 max_n / 迭代上限：
         target ≥ 1000万 → max_n=15, 上限 8000万次（覆盖 4610万=15笔）
         target ≥ 100万  → max_n=10, 上限 2000万次（覆盖 2800万=10笔/2200万=8笔）
         其他            → max_n=6,  上限 300万次
    ⚡ 性能保护：srcs 候选按金额排序后取前 MAX_SUM_CAND；超限返回 None 防卡死。
    返回 (pairs_target2src, targets_left, srcs_left)。"""
    pairs, used_t, used_s = [], set(), set()
    for a in targets:
        if id(a) in used_t:
            continue
        ad = _pd(a['date'])
        cands = []
        for b in srcs:
            if id(b) in used_s or b['amt'] > a['amt'] + 1:
                continue
            # ⚡ 2026-08-19 镜像校验：找和同样禁止应收↔应收/应付↔应付配对
            if not _mirror_ok(a['km'], b['km']):
                continue
            bd = _pd(b['date'])
            if SUM_WIN_DAYS is not None:
                # ⚡ 2026-08-19 日期缺失(ad/bd 为 None)不跳过——对子20 达州预付设备款 3000万 日期
                #    为空、集团预收 2203 有日期，原 `not (ad and bd)` 直接跳过致预收↔预付镜像配不上、
                #    单边虚增 6000万。仅双方都有日期时才校验窗口（金额约束 + 审计复核兜底）。
                if (ad and bd) and abs((ad - bd).days) > SUM_WIN_DAYS:
                    continue
            cands.append(b)
        if len(cands) < 2:
            continue
        cands.sort(key=lambda b: -b['amt'])
        amt = a['amt']
        if amt >= 10000000:
            max_n, lim = 15, 80000000
        elif amt >= 1000000:
            max_n, lim = 10, 20000000
        else:
            max_n, lim = 6, 3000000
        combo = _find_combo(amt, cands, max_n=max_n, limit=lim)
        if combo:
            used_t.add(id(a))
            used_s.update(id(c) for c in combo)
            pairs.append((a, combo))
    targets_left = [t for t in targets if id(t) not in used_t]
    srcs_left = [s for s in srcs if id(s) not in used_s]
    return pairs, targets_left, srcs_left


def match_sum(A_left, B_left):
    """找和层（双向）：A 单笔 = B 多笔之和；或 B 单笔 = A 多笔之和。
    ⚡ 2026-08-18 修复：宁夏宝丰 透平 192.5万×3(借应收) = 集团 577.5万(贷应付)——
       多笔在 A 侧、单笔在 B 侧，原单向实现配不上。返回 (paired, A_left2, B_left2)。"""
    out1, A_l1, B_l1 = _ms_dir(A_left, B_left)          # A 单笔 = B 多笔
    out2, B_l2, A_l2 = _ms_dir(B_l1, A_l1)              # B 单笔 = A 多笔（用剩余）
    paired = list(out1)
    for b, alist in out2:                                # 展开：每笔 A 对应 B 单笔
        for a in alist:
            paired.append((a, [b]))
    return paired, A_l2, B_l2


def classify(r):
    """单边分录类型标注（辅助审计判断）。"""
    km = str(r['km'])[:4]
    txt = str(r['txt'] or '')
    if km in ('1121', '2201') and ('承兑' in txt or '托收' in txt or '背书' in txt):
        return '票据到期变现/转让'
    if txt.startswith('销') or '销杭氧' in txt:
        return '销售应收(销集团)'
    if '货款' in txt or '采购' in txt:
        return '采购/货款'
    if '资金' in txt or '上收' in txt or '下拨' in txt:
        return '资金池残留'
    if '期初' in txt or '结转' in txt or str(r['date'])[:4] != '2026':
        return '期初/跨期'
    return '其他'


def _side_recover(rows):
    """⚡⚡ 同侧确认↔回收配对（2026-08-19 用户定：代垫/材料款应收确认后必然回收）。
    同一侧(如物资)借应收(确认) ↔ 贷应收(收到款/回收) 找和配对——如物资确认餐费应收 564.5万、
    回收 549.7万，净应收仅 14.8万，大部分是拆分行闭环（确认+回收都记了），非真实单边。
    在异侧配对【前】对原始经营分录执行，配掉的从单边剔除。
    轻量找和(max_n=5)：大额集中结算(8-15笔)由异侧增强找和处理，此处避免 DFS 卡死。
    返回 (剩余行, 配对行数)。"""
    drs = [dict(r, amt=r['dr'], dir='借', _oid=id(r)) for r in rows if r['dr']]
    crs = [dict(r, amt=r['cr'], dir='贷', _oid=id(r)) for r in rows if r['cr']]
    p1, drs_l, crs_l = match_pool(drs, crs)
    # 轻量双向找和（max_n=5 固定，防大额 DFS 卡死）
    p2, drs_l2, crs_l2 = _ms_lite(drs_l, crs_l)
    p3, crs_l3, drs_l3 = _ms_lite(crs_l, drs_l)   # 反向：回收单笔 = 确认多笔
    used = set()
    for a, b in p1:
        used.add(a['_oid']); used.add(b['_oid'])
    for a, bl in p2:
        used.add(a['_oid'])
        for b in bl:
            used.add(b['_oid'])
    for b, al in p3:                             # b=回收单笔, al=确认多笔
        used.add(b['_oid'])
        for a in al:
            used.add(a['_oid'])
    left = [r for r in rows if id(r) not in used]
    return left, len(used)


def _ms_lite(targets, srcs):
    """轻量单向找和：target 单笔 = srcs 2..5 笔之和（±1元，固定 max_n=5 防卡死）。
    返回 (pairs, targets_left, srcs_left)。"""
    pairs, used_t, used_s = [], set(), set()
    for a in targets:
        if id(a) in used_t:
            continue
        cands = [b for b in srcs if id(b) not in used_s and b['amt'] <= a['amt'] + 1]
        cands.sort(key=lambda b: -b['amt'])
        combo = _find_combo(a['amt'], cands, max_n=5, limit=500000)
        if combo:
            used_t.add(id(a))
            used_s.update(id(c) for c in combo)
            pairs.append((a, combo))
    targets_left = [t for t in targets if id(t) not in used_t]
    srcs_left = [s for s in srcs if id(s) not in used_s]
    return pairs, targets_left, srcs_left


def _self_net(rows):
    """同侧借贷对冲（支持 1:1 同额 + 找和，如 集团借2202 9000万 = 贷4800万 + 贷4200万）。
    同一侧(如 Y 侧)借贷同额/找和抵消 = 集团自身确认应付+支付的自平（净额 0），非往来差异。
    返回 (剩余行, 对冲行数)。"""
    drs = [r for r in rows if r.get('dir') == '借']
    crs = [r for r in rows if r.get('dir') == '贷']
    p1, drs_l, crs_l = match_pool(drs, crs)          # 1:1 同额
    p2, drs_l, crs_l = match_sum(drs_l, crs_l)       # 找和（借多笔=贷单笔 或 反向）
    used = set()
    for a, b in p1:
        used.add(id(a))
        used.add(id(b))
    for a, bl in p2:
        used.add(id(a))
        for b in bl:
            used.add(id(b))
    left = [r for r in rows if id(r) not in used]
    return left, len(used)


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    pair_no = None
    if '--pair' in argv:
        pair_no = int(argv[argv.index('--pair') + 1])
        print(f'单对模式：仅处理对子 #{pair_no}（先抽该对涉及账套序时账：--only 过滤）', flush=True)
    pairs = _load_pairs()
    if pair_no is not None:
        m = [p for p in pairs if p['no'] == pair_no]
        if not m:
            print(f'⚠️ 对子 #{pair_no} 不存在或非需穿透对子（可用范围见核对表穿透需求列）')
            return 1
        pairs = m
    keys = set()
    for p in pairs:
        keys.add((p['X'], p['Y']))
        keys.add((p['Y'], p['X']))
    idx = _load_seq(keys)
    print(f'差异对 {len(pairs)} 对 | 往来科目分录 {sum(len(v) for v in idx.values())} 行')

    # ⚡ 公司代码映射（供 _squeeze 供应商过滤/内部冲销挤水分）
    from ah_intra_seq_extract import load_company_map
    cmap = load_company_map()

    from openpyxl.cell import WriteOnlyCell  # noqa
    wb = openpyxl.Workbook(write_only=True)

    # ---- Sheet1 对冲汇总 ----
    ws = wb.create_sheet('差异对对冲汇总')
    H1 = ['序号', '我方', '对方', '状态',
          '经营镜像差①', '经营镜像差②', '资金池镜像差①', '资金池镜像差②', '双向净额',
          '期初经营差①', '期初经营差②', '期初资金池差①', '期初资金池差②',
          '差异一致性', '穿透需求',
          '双方往来分录数', '双边对上(笔)', '双边对上金额',
          'X侧单边(笔)', 'X侧单边金额', 'Y侧单边(笔)', 'Y侧单边金额', '单边净额',
          '净额镜像闭环(笔)', '勾稽差额',
          '自检', '结论']
    ws.append([_wc(ws, v) for v in H1])

    # ---- Sheet2 单边入账明细（未来调整候选）----
    ws2 = wb.create_sheet('单边入账明细')
    H2 = ['序号', '我方', '对方', '单边侧', '账套', '日期', '凭证编号', '科目', '文本', '客户', '供应商',
          '借方金额', '贷方金额', '类型']
    ws2.append([_wc(ws2, v, FILL_RED) for v in H2])

    # ---- Sheet3 双边对上明细 ----
    ws3 = wb.create_sheet('双边对上明细')
    H3 = ['序号', '我方', '对方', '方向', 'X侧账套', 'X侧日期', 'X侧凭证', 'X侧科目', 'X侧文本', 'X侧金额',
          'Y侧账套', 'Y侧日期', 'Y侧凭证', 'Y侧科目', 'Y侧文本', 'Y侧金额', '日期差(天)']
    ws3.append([_wc(ws3, v) for v in H3])

    n_unil = 0
    n_bil = 0
    n_pool_pair = 0
    n_int_total = 0
    n_self_total = 0
    n_side_total = 0   # ⚡ 同侧确认↔回收配对行数（代垫/材料款闭环）
    self_check_list = []   # ⚡ 自检警告集合（单边过多 / 勾稽差≠0）
    for i, p in enumerate(pairs, 1):
        print(f'  [{i}/{len(pairs)}] 对子#{p["no"]} {p["X"][:12]}↔{p["Y"][:12]} ...', flush=True)
        rows_x2y = idx.get((p['X'], p['Y']), [])
        rows_y2x = idx.get((p['Y'], p['X']), [])
        # ⚡⚡ 挤水分（2026-08-19 用户定）：①供应商代码非本方/对方主体 → 凭证级聚合误归剔除
        #    ②同凭证 借2202=贷2202 应付内部冲销（净额0非往来差异）剔除
        rows_x2y = _squeeze(rows_x2y, {p['Y']}, cmap)
        rows_y2x = _squeeze(rows_y2x, {p['X']}, cmap)
        # ⚡ 内部科目结转剔除（如 借合同资产1124/贷应收1122『开票结转』——往来科目间互转，
        #    不产生新对外往来，集团侧不应有对应，原被误判为单边）
        rows_x2y, drop_x2y = _drop_internal(rows_x2y)
        rows_y2x, drop_y2x = _drop_internal(rows_y2x)
        n_int = len(drop_x2y) + len(drop_y2x)
        n_int_total += n_int
        # ⚡ 资金池分轨：其他应收(1221) ↔ 其他应付(2241) 按【发生净额镜像】核对
        #    （用户17:25定：资金池是集团单向归集，逐笔借贷方向可能相同/含利息尾差，无需回序时账逐笔对冲）
        #    ⚡ 2026-08-19 方向自适应：不硬编码"X=1221/Y=2241"——对子反向时（X记2241/Y记1221，
        #    如 A086↔A094 资金池对子）原逻辑两侧 pool 皆空，资金池全落经营单边 → 单边净额虚增 32,142 万。
        #    修复：任一侧有 1221 + 任一侧有 2241 即构成资金池镜像，按科目汇总两侧全部 1221/2241。
        #    ⚡ 2026-08-19 委贷利息应收（1221010000 其他应收-委贷利息）非资金池归集 → 排除出资金池，
        #    归经营应收利息（与对方 2231 应付利息镜像）。用户 20:23 确认。对子4 集团委贷利息应收
        #    原被资金池吸走、泽州 2231 在经营轨，两边不同轨致经营勾稽差 102万。
        #    ⚡⚡ 2026-08-19 科目归类对齐核对表后撤销：利息/股利科目(1131/1132/2231/2232)已移出经营
        #    （核对表 SUBJS 不含），委贷利息不再参与经营核对；122101 属"其他应收款"（核对表资金池组），
        #    归资金池分轨，与核对表 p1/p2 口径一致。此前对子4 的 102万 勾稽差随利息科目整体移出而消失。
        #    ⚡ 2026-08-19 资金下拨（1122030000 其他应收-资金下拨）纳入资金池分轨——集团↔成员资金调拨，
        #    成员应收集团资金，属资金池类。对子77 滕州 112203 借=贷3000万（6100001959/1960 同业务跨凭证
        #    对称净0）原落经营轨，与设备应收 3000万 match_sum 金额巧合误配 → 设备应收 1.836亿被假配对
        #    消耗。纳入 pool_x 后净0不影响镜像差，且不再误配经营应收。（#27 亦曾见 112203 300万×2 对称）
        def _is_pool_x(km):
            k = str(km)
            return k[:4] == POOL_X or k.startswith('112203')
        x1221 = [r for r in rows_x2y if _is_pool_x(r['km'])]
        x2241 = [r for r in rows_x2y if str(r['km'])[:4] == POOL_Y]
        y1221 = [r for r in rows_y2x if _is_pool_x(r['km'])]
        y2241 = [r for r in rows_y2x if str(r['km'])[:4] == POOL_Y]
        pool_x = x1221 + y1221          # 全部 1221（应收资金池，除委贷利息应收）+ 112203 资金下拨
        pool_y = x2241 + y2241          # 全部 2241（应付资金池）
        # 经营：其余往来科目行（排除 1221/2241 两类资金池科目与 112203 资金下拨）
        def _is_pool_km(km):
            k = str(km)
            return k[:4] in (POOL_X, POOL_Y) or k.startswith('112203')
        biz_x2y = [r for r in rows_x2y if not _is_pool_km(r['km'])]
        biz_y2x = [r for r in rows_y2x if not _is_pool_km(r['km'])]
        # ⚡ 2026-08-19 曾加『同侧确认↔回收配对』(_side_recover)：对子10 回归 4→19 笔误配——
        #    提前消耗 X 贷应收(回收)，破坏原本成功的异侧配对(X贷应收↔Y借应付)。已回滚。
        #    该方法仅适用于"受益方直接支付(贷银行存款/预付)不挂往来"的特殊场景(如代垫费用)，
        #    不能推广到通过往来结算的对子(用户定)。代垫确认应收留单边=真实未收，保留为调整候选。
        n_x_side = n_y_side = 0

        # ---- 资金池净额镜像 ----
        pool_x_dr = sum(r['dr'] for r in pool_x)
        pool_x_cr = sum(r['cr'] for r in pool_x)
        pool_y_dr = sum(r['dr'] for r in pool_y)
        pool_y_cr = sum(r['cr'] for r in pool_y)
        pool_x_net = pool_x_dr - pool_x_cr      # X 其他应收 本期净变化
        pool_y_net = pool_y_cr - pool_y_dr      # Y 其他应付 本期净变化（贷余）
        pool_mirror = pool_x_net - pool_y_net   # 镜像差（应收净变化 - 应付净变化，应≈0）
        n_pool_x = len(pool_x)
        n_pool_y = len(pool_y)
        if n_pool_x or n_pool_y:
            n_pool_pair += 1

        # ---- 经营逐笔对冲（异向 + 找和）----
        x2y_dr = [dict(r, amt=r['dr'], dir='借') for r in biz_x2y if r['dr']]
        x2y_cr = [dict(r, amt=r['cr'], dir='贷') for r in biz_x2y if r['cr']]
        y2x_dr = [dict(r, amt=r['dr'], dir='借') for r in biz_y2x if r['dr']]
        y2x_cr = [dict(r, amt=r['cr'], dir='贷') for r in biz_y2x if r['cr']]
        p1, x2y_dr_left, y2x_cr_left = match_pool(x2y_dr, y2x_cr)
        p2, x2y_cr_left, y2x_dr_left = match_pool(x2y_cr, y2x_dr)
        s1, x2y_dr_left, y2x_cr_left = match_sum(x2y_dr_left, y2x_cr_left)
        s2, x2y_cr_left, y2x_dr_left = match_sum(x2y_cr_left, y2x_dr_left)
        paired = p1 + p2 + s1 + s2
        x_left = x2y_dr_left + x2y_cr_left
        y_left = y2x_cr_left + y2x_dr_left
        # ⚡ 2026-08-19 用户定：『先不能同侧抵消』——核对原则是异侧配对：
        #    X借应收(确认) ↔ Y贷应付(确认)；X贷应收(回收) ↔ Y借应付(支付)。
        #    _self_net(同侧借贷对冲)会把 Y侧"借应付9000万=贷4800+4200万"同侧抵消，
        #    掩盖 X侧真实应收/回收 → 已禁用。仅"受益方直接支付(贷银行存款)"场景才允许同侧抵消（未启用）。
        n_x_self = n_y_self = 0
        n_self_total += 0

        # ⚡ 2026-08-18 曾加『月度净额镜像层』：实测拆分行对子每月净额也严重不对齐（真实差异，
        #    非记账粒度差），降级会掩盖差异、且效果微弱（单边 3587→3416、警告 135→138）→ 已回滚。
        #    单边全部保留为调整候选，由自检报告标注"单边过多需进一步核对"。
        n_net_closed = 0
        n_bil += len(paired)
        bil_amt = sum(a['amt'] for a, _ in paired)
        n_x, n_y = len(x_left), len(y_left)
        amt_x = sum((r['amt'] if r['amt'] > 0 else -r['amt']) for r in x_left)   # 绝对量（量级展示）
        amt_y = sum((r['amt'] if r['amt'] > 0 else -r['amt']) for r in y_left)
        # ⚡⚡ 2026-08-19 带符号单边净额（用户定：绝对量差使勾稽失真——对子7 曾显示 -5216万 实为口径假象，
        #    带符号后仅 -137万）。差①方向 = X应收净变 − Y应付净变；差②方向 = X应付净变 − Y应收净变。
        x_recv = sum((r['dr'] or 0) - (r['cr'] or 0) for r in x_left if str(r['km'])[:4] in RECV_CODES)
        x_pay = sum((r['cr'] or 0) - (r['dr'] or 0) for r in x_left if str(r['km'])[:4] in PAY_CODES)
        y_recv = sum((r['dr'] or 0) - (r['cr'] or 0) for r in y_left if str(r['km'])[:4] in RECV_CODES)
        y_pay = sum((r['cr'] or 0) - (r['dr'] or 0) for r in y_left if str(r['km'])[:4] in PAY_CODES)
        d1_unil = x_recv - y_pay          # 差①方向单边贡献
        d2_unil = x_pay - y_recv          # 差②方向单边贡献
        unil_net = d1_unil + d2_unil      # 带符号经营单边净额

        parts = []
        if n_int:
            parts.append(f'内部科目结转剔除 {n_int} 行')
        if n_x_self + n_y_self:
            parts.append(f'同侧借贷对冲 {n_x_self + n_y_self} 行')
        if n_pool_x or n_pool_y:
            parts.append(f'资金池: X侧{pn(pool_x_net)} vs Y侧{pn(pool_y_net)} 镜像差{pn(pool_mirror)} ({n_pool_x}+{n_pool_y}行)')
        if n_net_closed:
            parts.append(f'净额镜像闭环 {n_net_closed} 行（拆分行已核对）')
        if n_x_side + n_y_side:
            parts.append(f'同侧确认回收 {n_x_side + n_y_side} 行')
        if n_x + n_y == 0:
            if not parts:
                concl = '序时账无双方发生分录：疑期初结转，或凭证文本/客户未含对方公司名'
            else:
                parts.append('经营往来全部双边对上')
                concl = '；'.join(parts)
        else:
            parts.append(f'经营单边 {n_x}笔(X)+{n_y}笔(Y) 净额{pn(unil_net)}')
            concl = '；'.join(parts)

        # ⚡ 勾稽校验：期末经营差 - 期初经营差 ≈ 本期经营单边净（双边对上镜像抵消不改变差）
        #    期初差来自核对表明细底稿『期初余额(人民币)』列，非科目余额表（明细粒度可到对子，更精确）
        check = round((p['op1'] - p['op1_0']) + (p['op2'] - p['op2_0']) - unil_net, 2)
        parts.append(f'勾稽: 期末经营差{round(p["op1"]-p["op1_0"]+p["op2"]-p["op2_0"],2):,.0f}'
                     f'-期初{(p["op1_0"]+p["op2_0"]):,.0f}-本期单边{unil_net:,.0f}'
                     f'=差{check:,.0f}')

        # ⚡ 自检（用户定：差异凭证一般不超过 10 多笔；两边仍多笔金额=匹配不充分需进一步核对）
        warns = []
        if n_x + n_y > MAX_UNIL_PER_PAIR:
            warns.append(f'单边{n_x + n_y}笔>阈值{MAX_UNIL_PER_PAIR}')
        if abs(check) > 1.0:
            warns.append(f'勾稽差{check:,.0f}')
        self_check = '通过' if not warns else '⚠️ ' + '；'.join(warns)
        if warns:
            self_check_list.append([p['no'], p['X'], p['Y'], n_x + n_y, round(unil_net, 2), check, self_check])

        ws.append([p['no'], p['X'], p['Y'], p['st'],
                   round(p['op1'], 2), round(p['op2'], 2), round(p['pool1'], 2), round(p['pool2'], 2),
                   round(p['net'], 2),
                   round(p['op1_0'], 2), round(p['op2_0'], 2), round(p['pool1_0'], 2), round(p['pool2_0'], 2),
                   p['consist'], p['need'],
                   len(rows_x2y) + len(rows_y2x), len(paired), round(bil_amt, 2),
                   n_x, round(amt_x, 2), n_y, round(amt_y, 2), round(unil_net, 2),
                   n_net_closed, check,
                   self_check, concl])

        # Sheet2 经营单边（未来调整候选）
        for r in x_left:
            n_unil += 1
            ws2.append([p['no'], p['X'], p['Y'], 'X侧', r['sz'], r['date'], r['vou'], r['km'], r['txt'],
                        r['cust'], r['supp'],
                        round(r['dr'], 2) if r['dr'] else None, round(r['cr'], 2) if r['cr'] else None,
                        classify(r)])
        for r in y_left:
            n_unil += 1
            ws2.append([p['no'], p['X'], p['Y'], 'Y侧', r['sz'], r['date'], r['vou'], r['km'], r['txt'],
                        r['cust'], r['supp'],
                        round(r['dr'], 2) if r['dr'] else None, round(r['cr'], 2) if r['cr'] else None,
                        classify(r)])
        # Sheet3 双边对上（精确配对 b 是 dict；找和配对 b 是 [b1,b2] 展开）
        for a, b in paired:
            blist = b if isinstance(b, list) else [b]
            tag = 'X借↔Y贷' if (a, b) in p1 or (a, b) in s1 else 'X贷↔Y借'
            for bi in blist:
                ad, bd = _pd(a['date']), _pd(bi['date'])
                diff = abs((ad - bd).days) if (ad and bd) else ''
                ws3.append([p['no'], p['X'], p['Y'], tag,
                            a['sz'], a['date'], a['vou'], a['km'], a['txt'], round(a['amt'], 2),
                            bi['sz'], bi['date'], bi['vou'], bi['km'], bi['txt'], round(bi['amt'], 2), diff])

    for col, w in zip('ABCDEFGHIJKLMNOPQRSTUVWXYZAA', [5, 20, 20, 8, 13, 13, 13, 13, 14, 13, 13, 13, 13,
                                                       16, 26, 11, 10, 12, 10, 14, 10, 14, 14, 12, 12, 16, 42]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:AA{len(pairs) + 1}'
    for col, w in zip('ABCDEFGHIJKLMN', [5, 20, 20, 6, 10, 11, 13, 14, 46, 10, 10, 14, 14, 14]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = 'A2'
    ws2.auto_filter.ref = f'A1:N{n_unil + 1}'
    for col, w in zip('ABCDEFGHIJKLMNOPQ', [5, 20, 20, 9, 9, 11, 12, 13, 40, 13, 9, 11, 12, 13, 40, 13, 8]):
        ws3.column_dimensions[col].width = w
    ws3.freeze_panes = 'A2'
    ws3.auto_filter.ref = f'A1:Q{n_bil + 1}'

    # ---- Sheet4 序时账命中明细（原独立《关联往来序时账_2026.xlsx》并入单表）----
    ws4 = wb.create_sheet('序时账命中明细')
    src = openpyxl.load_workbook(SEQ_FP, read_only=True, data_only=True)
    ssrc = src['关联往来序时账']
    it_src = ssrc.iter_rows(values_only=True)
    try:
        hdr_src = next(it_src)
    except StopIteration:
        hdr_src = ['账套', '公司', '日期', '凭证编号', '科目', '文本', '客户', '供应商',
                   '借方金额', '贷方金额', '对方公司名', '备注']
    ws4.append([_wc(ws4, v) for v in hdr_src])
    n_seq = 0
    for r in it_src:
        ws4.append(r)
        n_seq += 1
    src.close()
    for col, w in zip('ABCDEFGHIJKL', [10, 18, 11, 13, 14, 50, 12, 12, 15, 15, 22, 16]):
        ws4.column_dimensions[col].width = w
    ws4.freeze_panes = 'A2'
    ws4.auto_filter.ref = f'A1:L{n_seq + 1}'

    # ---- Sheet5 口径说明 ----
    ws5 = wb.create_sheet('口径说明')
    notes = [
        '关联往来差异穿透底稿 v3（2026-08-18）—— 分层核对 + 期初纳入 + 单表交付',
        '1. 数据源：核对表《88家内部往来核对_2026.xlsx》『需穿透』对子 + 序时账《关联往来序时账_2026.xlsx》全量命中明细（已并入本表 Sheet4）。',
        '2. 【分层核对】核对表区分期初核对（明细底稿『期初余额』列）与期末核对（『期末余额』列）两个口径：',
        '   · 期初差=期末差（一致）→ 极大概率期初差异一直拖到期末，但也可能巧合，标注『期初延续』不穿透（无需回序时账）；',
        '   · 期初差≠期末差（不一致）→ 本期发生造成差异，『需穿透』；',
        '   · 期初平/期末不平 → 本期新增差异，『需穿透』；期初不平/期末平 → 本期已消除，『本期已平』不穿透；',
        '   · 全程对平 → 无需穿透。',
        '3. 只取【往来科目】分录参与对冲（应收/应付/预收/预付/其他应收/其他应付/合同资产/合同负债），票据(1121/2201)是结算工具非往来余额、应付暂估(220206)、内部科目结转(往来科目互转)、同侧借贷自平均剔除。',
        '4. 对冲匹配：X记借↔Y记贷（或反向）金额±1元；找和层支持 A单笔=B多笔 / B单笔=A多笔（跨期数月），窗口 SUM_WIN_DAYS=180 天。',
        '5. 双边对上剔除后【剩余=单边入账】（Sheet2），列为未来审计调整候选。',
        '6. 【勾稽校验】期末经营差 - 期初经营差 ≈ 本期单边净（Sheet1『勾稽差额』列，=0 则勾稽成立）；双边对上镜像抵消不改变差，故期末差=期初差+本期单边。',
        '7. 期初差取自明细底稿『期初余额(人民币)』列（对子级，比科目余额表更精确）；勾稽差额≠0 的对子表示口径未完全对齐（如对子3），需结合期初往来余额逐对核查。',
        '8. 序时账仅为发生额；期初结转差异在序时账无分录（如化医↔宏裕 7,464 万为期初），已在期初核对列单独体现。',
    ]
    for n in notes:
        ws5.append([_wc(ws5, n, None)])
    ws5.column_dimensions['A'].width = 130

    # ---- Sheet6 自检报告（自动质量把关）----
    ws6 = wb.create_sheet('自检报告')
    H6 = ['对子号', '我方', '对方', '单边笔数', '单边净额', '勾稽差额', '自检结果', '处理建议']
    ws6.append([_wc(ws6, v, FILL_RED) for v in H6])
    for row in self_check_list:
        i, X, Y, nunil, unet, check, sc = row
        if '单边' in sc:
            advice = '匹配层未充分（可能含税口径/找和/同侧未配），需增强匹配后重跑该对'
        else:
            advice = '期初与本期口径未对齐，需核查期初科目归属与本期发生'
        ws6.append([i, X, Y, nunil, round(unet, 2), round(check, 2), sc, advice])
    for col, w in zip('ABCDEFGH', [7, 20, 20, 10, 14, 12, 34, 42]):
        ws6.column_dimensions[col].width = w
    ws6.freeze_panes = 'A2'
    ws6.auto_filter.ref = f'A1:H{len(self_check_list) + 1}'
    ws6.append([])
    ws6.append([f'穿透对子 {len(pairs)} 对；自检通过 {len(pairs) - len(self_check_list)} 对；'
                f'警告 {len(self_check_list)} 对（单边>阈值{MAX_UNIL_PER_PAIR}笔 或 勾稽差≠0）'])

    wb.save(OUT_FP)
    print(f'已生成：{OUT_FP}')
    print(f'  差异对 {len(pairs)} | 双边对上 {n_bil} 笔 | 单边入账 {n_unil} 笔（未来调整候选）| 序时账明细 {n_seq} 行')
    print(f'  自检：警告 {len(self_check_list)} 对（单边过多/勾稽差≠0），详见 Sheet6')
    return 0


if __name__ == '__main__':
    sys.exit(main())
