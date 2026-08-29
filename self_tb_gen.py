# -*- coding: utf-8 -*-
"""自建试算表生成器（2026-08-06，用户方法论：流程倒过来）

拿到账套数据后，程序先从各主体单体科目余额表【直出】一版试算表：
  · 完全按 TB 原值汇总，不做任何重分类/审计调整/合并抵消
  · 叶子过滤（U8 父级行=子级行同金额，父+子双计 → 只取末级）
  · 平衡校验：每主体 借发=贷发、期末借=期末贷（叶子层）
  · 未结转损益识别：损益类科目期末余额 ≠ 0 → 单独列示，归入权益（报表勾稽用）
  · 标准报表：资产负债表（资产=负债+权益，含未结转损益）与利润表（发生额，
    收入取贷/费用取借，铁律3）勾稽相符校验
  · 集团合计列 = 全部主体之和；母分汇总/集团合并后续模块基于本输出扩展

用法：
  python self_tb_gen.py <folder> [--out <path>] [--year 2025]
"""
import os
import re
import sys
import argparse
from collections import OrderedDict, defaultdict

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import audit_common as A

F = Font(name='Times New Roman', size=10)
FB = Font(name='Times New Roman', size=10, bold=True)
FR = Font(name='Times New Roman', size=10, color='C00000', bold=True)
FG = Font(name='Times New Roman', size=10, color='006100')
FILL_H = PatternFill('solid', fgColor='DDEBF7')
THIN = Border(*[Side(style='thin', color='BFBFBF')] * 4)
NUM = '#,##0.00'


# ---------------------------------------------------------------- 叶子过滤
def leaf_codes(code_list):
    """剔除有子级的父级：返回仅末级代码集合（U8 父级=子级同金额，父+子双计）。"""
    codes = sorted(code_list)
    return {c for i, c in enumerate(codes)
            if not (i + 1 < len(codes) and codes[i + 1].startswith(c))}


def code_class(code):
    """1=资产 2=负债 3=共同 4=权益 5=成本 6/7/8=损益"""
    c = re.sub(r'\D', '', str(code))[:1]
    return {'1': 'asset', '2': 'liab', '3': 'equity', '4': 'equity',
            '5': 'cost', '6': 'pl', '7': 'pl', '8': 'pl'}.get(c, 'other')


def name_first(nm):
    """科目名取首段（去反斜杠层级：'管理费用\\研发费用' → '管理费用'）。"""
    nm = str(nm or '')
    nm = re.sub(r'^\d+[\\/.-]\s*', '', nm)  # 去"1002\银行存款"前缀编码
    for sep in ('\\', '/', '-'):
        if sep in nm:
            return nm.split(sep)[0].strip()
    return nm.strip()


# ---------------------------------------------------------------- 报表行映射
BS_ASSET_ROWS = OrderedDict([
    ('货币资金',        {'kw': ['库存现金', '银行存款', '其他货币资金', '内部结算中心存款']}),
    ('应收票据',        {'kw': ['应收票据']}),
    ('应收账款',        {'kw': ['应收账款']}),
    ('其他应收款',      {'kw': ['其他应收款', '应收暂付款', '应收利息', '应收股利']}),
    ('合同资产',        {'kw': ['合同资产']}),
    ('存货',            {'kw': ['库存商品', '发出商品', '原材料', '包装物', '低值易耗品', '周转材料',
                                '在产品', '材料成本差异', '开发成本', '工程施工', '存货', '委托调拨物资']}),
    ('其他流动资产',    {'kw': ['待摊', '预付款', '预付账款']}),
    ('长期股权投资',    {'kw': ['长期股权']}),
    ('投资性房地产',    {'kw': ['投资性房地产']}),
    ('固定资产',        {'kw': ['固定资产', '临时设施']}),
    ('在建工程',        {'kw': ['在建工程']}),
    ('使用权资产',      {'kw': ['使用权资产']}),
    ('无形资产',        {'kw': ['无形资产']}),
    ('长期待摊费用',    {'kw': ['长期待摊']}),
    ('长期应收款',      {'kw': ['长期应收款']}),
    ('递延所得税资产',  {'kw': ['递延所得税资产']}),
    ('其他非流动资产',  {'kw': ['抵债资产', '持有待售资产', '贷款', '其他非流动']}),
])

BS_LIAB_ROWS = OrderedDict([
    ('短期借款',        {'kw': ['短期借款']}),
    ('应付票据',        {'kw': ['应付票据']}),
    ('应付账款',        {'kw': ['应付账款']}),
    ('应付职工薪酬',    {'kw': ['应付职工薪酬']}),
    ('应交税费',        {'kw': ['应交税费', '应交税金']}),
    ('其他应付款',      {'kw': ['其他应付款', '应付利息', '应付股利', '预提费用']}),
    # 2026-08-07 修复：预收账款独立成行（原挂其他应付款 kw 含『预收』→ dq 预收账款
    # 65.7M 被并入其他应付款 → 报表其他应付款 -157.55M vs TB 91.85M，差=预收 65.7M）
    ('预收账款',        {'kw': ['预收账款', '预收']}),
    ('合同负债',        {'kw': ['合同负债']}),
    ('一年内到期的非流动负债', {'kw': ['一年内到期']}),
    ('长期借款',        {'kw': ['长期借款']}),
    ('租赁负债',        {'kw': ['租赁负债']}),
    ('长期应付款',      {'kw': ['长期应付款']}),
    ('专项应付款',      {'kw': ['专项应付款']}),
    ('递延收益',        {'kw': ['递延收益']}),
    ('递延所得税负债',  {'kw': ['递延所得税负债']}),
])

BS_EQUITY_ROWS = OrderedDict([
    ('实收资本',        {'kw': ['实收资本', '股本']}),
    ('资本公积',        {'kw': ['资本公积']}),
    ('盈余公积',        {'kw': ['盈余公积']}),
    ('未分配利润',      {'kw': ['未分配利润', '利润分配', '本年利润']}),
    ('专项储备',        {'kw': ['专项储备']}),
    ('未结转损益',      {}),
])

# 兜底行：未匹配科目收纳（保证勾稽恒平，同时暴露待补映射科目）
BS_OTHER_ROWS = OrderedDict([
    ('其他（未匹配）', {}),
])

# 备抵科目（贷余负值，自动抵减原值科目）
CONTRA_MAP = {
    '坏账准备': '应收账款',
    '存货跌价准备': '存货',
    '累计折旧': '固定资产',
    '固定资产减值准备': '固定资产',
    '累计摊销': '无形资产',
    '无形资产减值准备': '无形资产',
    '使用权资产累计折旧': '使用权资产',
    '未确认融资费用': '租赁负债',
    '合同资产减值准备': '合同资产',
    '临时设施摊销': '固定资产',
    '抵债资产跌价准备': '其他非流动资产',
    '贷款损失准备': '其他非流动资产',
    '持有待售资产减值': '其他非流动资产',
}

IS_ROWS = OrderedDict([
    ('一、营业收入',            {'credit': ['6001', '6051'], 'kw': ['主营业务收入', '其他业务收入']}),
    ('减：营业成本',            {'debit': ['6401', '6402'], 'kw': ['主营业务成本', '其他业务成本']}),
    ('税金及附加',              {'debit': ['6403'], 'kw': ['税金及附加']}),
    ('销售费用',                {'debit': ['6601'], 'kw': ['销售费用']}),
    ('管理费用',                {'debit': ['6602', '6600'], 'kw': ['管理费用']}),
    # ⚡ 2026-08-15 ga 特例：研发费用=6604（费用化，ga 全集团 2.24 亿；不单列会漏计
    #   进利润总额 → 净利润虚高）。SAP/其他账套无 6604 → 恒 0 行，零影响。
    ('研发费用',                {'debit': ['6604'], 'kw': ['研发']}),
    ('财务费用',                {'debit': ['6603'], 'kw': ['财务费用']}),
    ('加：其他收益',            {'credit': ['6117', '6115', '6113', '6121'], 'kw': ['其他收益']}),
    ('投资收益',                {'credit': ['6101', '6111'], 'kw': ['投资收益']}),
    ('营业外收入',              {'credit': ['6301'], 'kw': ['营业外收入']}),
    ('减：营业外支出',          {'debit': ['6711'], 'kw': ['营业外支出']}),
    ('资产减值损失',            {'debit': ['6701'], 'kw': ['资产减值损失']}),
    ('信用减值损失',            {'debit': ['6702'], 'kw': ['信用减值损失']}),
    ('四、利润总额',            {'calc': True}),
    ('减：所得税费用',          {'debit': ['6801'], 'kw': ['所得税费用']}),
    ('五、净利润',              {'calc': True}),
])


# ---------------------------------------------------------------- 主构建
_SAP_MODE = False    # 模块级标志：SAP 账套（AH 等）build_is 按名称匹配（铁律130b，2026-08-15）


def _is_sap(data_dir):
    """SAP 账套检测（detect_sap_layout 命中 → SAP 数据形态，走 read_sap_tb 分支）。"""
    try:
        import sap_reader as _SR
        return bool(_SR.detect_sap_layout(data_dir))
    except Exception:
        return False


def build_entity_tb(data_dir, years=None, comps=None):
    """读 folder 全部主体单体 TB → 叶子过滤 → {ent: {yy: {code: 值字典}}}。

    ⚡⚡ 铁律127b（2026-08-15）：TB 损益发生额全 0 且该主体存在 GL → read_tb_pl_safe 重读
    （TB+GL 兜底）。ga 实证：40 主体中 16 个 TB 损益科目本期发生额缺失（收入/成本/费用
    系统性少计，18/29/30 主体收入 0 vs GL 序时账完整）——与损益类生成器 read_tb_pl_safe
    同源 → 自建试算表利润表与底稿天然一致（合并稿对平前提）。正常账套（TB=GL）零影响。

    ⚡⚡ 铁律130b（2026-08-15）：SAP 账套（AH 等）走 read_sap_tb 分支——SAP 内部码 10 位、
    名称以 - 分层、无 U8 文件名结构。l1=code[:4]（SAP 内部码前 4 位对齐标准科目码）、
    name=名称首段。comps 白名单（集团分组）在此过滤。"""
    global _SAP_MODE
    if _is_sap(data_dir):
        _SAP_MODE = True
        # ⚡ 2026-08-24 架构阶段1：统一 read_tb 入口（ledger_backend 按形态分发，替代直接 read_sap_tb）
        from ledger_backend import read_tb as _lrtb
        yy = str((years or ['2026'])[0])
        _tb = _lrtb(data_dir)
        _all = sorted({c for (c, _cd, _n, _y) in _tb})
        if comps:
            comps = [c for c in _all if c in comps]
        else:
            comps = _all
        out = {}
        for _comp in comps:
            ent_tb = {}
            for (c, cd, n, y), v in _tb.items():
                if c != _comp:
                    continue
                _code = str(cd).replace('.', '').strip()
                if not _code:
                    continue
                ent_tb[_code] = {'name': name_first(n), 'l1': _code[:4],
                                 'qc': float(v.get('qc') or 0.0),
                                 'jf': float(v.get('jf') or 0.0),
                                 'df': float(v.get('df') or 0.0),
                                 'qm': float(v.get('qm') or 0.0)}
            out[_comp] = {str(yy): ent_tb}
        return out, [str(yy)]
    _SAP_MODE = False
    entities = A.discover_entities(data_dir)
    tb_full = A.read_tb_full(data_dir, entities)
    if years is None:
        years = sorted({str(yy) for (e, c, n, yy) in tb_full}) or ['2025']
    out = {}
    for ent in sorted(entities):
        out[ent] = {}
        for yy in years:
            rows = {c: v for (e, c, n, y), v in tb_full.items()
                    if e == ent and str(y) == yy}
            # ⚡ 铁律127b：TB 损益发生额缺失/异常 → GL 兜底重读（仅损益科目，余零影响）。
            #   read_tb_pl_safe 内部 pl_gl_fallback 自动检测：a) 损益全 0 → 全量填回；
            #   b) 单科目 TB 与 GL 差>50%（TB 减半/部分缺失，ga 18 主体 TB 0.90 亿 vs GL
            #   4.64 亿实证）→ GL 覆盖。正常账套（TB=GL）不触发，值零变化。
            _pl_codes = [c for c in rows if str(c)[:1] in ('6', '7', '8')]
            _paths = entities.get(ent, {}).get(yy, {})
            if _pl_codes and _paths.get('gl') and _paths.get('km'):
                try:
                    _tb2 = A.read_tb_pl_safe({'km': _paths['km'], 'gl': _paths['gl']})
                    _n2 = 0
                    for _c, _v in _tb2.items():
                        _k = str(_c).replace('.', '')
                        if _k in rows:
                            rows[_k]['jf'] = float(_v.get('debit') or 0.0)
                            rows[_k]['df'] = float(_v.get('credit') or 0.0)
                            _n2 += 1
                    if _n2:
                        print(f'  ⚡ [GL兜底] {ent[:18]}：TB 损益发生额缺失/异常，已按序时账重读 {_n2} 科目')
                except Exception as _ex:
                    print(f'  ⚠️ [GL兜底失败] {ent[:18]}：{_ex}')
            names = {c: n for (e, c, n, y) in tb_full if e == ent and str(y) == yy}
            # 一级名称：优先取父级行（level=1）名称（如 1002『银行存款』），
            # 叶子名是子级名（1002.01『人民币』）不能用于报表映射
            l1_names = {}
            for (e, c, n, y), v in tb_full.items():
                # ⚡⚡ 2026-08-23 一级名来源兼容：level==1 或 代码无点（U8 点分级如
                #   AFJ『1012.02 承兑保证金』，read_tb_full 可能未标 level）——否则
                #   l1_names 空 → 一级行名取到子目名（1012 显示『承兑保证金』错）。
                _c = str(c)
                if e == ent and str(y) == yy and (v.get('level') == 1 or '.' not in _c):
                    l1 = re.sub(r'\D', '', _c)[:4]
                    l1_names.setdefault(l1, name_first(n))
            leafs = leaf_codes(rows.keys())
            ent_tb = {}
            for c in leafs:
                v = rows[c]
                l1 = re.sub(r'\D', '', str(c))[:4]
                # ⚡⚡ 2026-08-23 子目显示真实子目名（1012.02『承兑保证金』而非父级
                #   『其他货币资金』）：一级（无点）用父级行名，子目（有点）用自身名。
                _nm = name_first(names.get(c, c)) or ''
                _has_dot = '.' in str(c)
                ent_tb[c] = {'name': (_nm if _has_dot else l1_names.get(l1)) or _nm,
                             # ⚡⚡ 2026-08-23 父级名（一级科目名）：l1_agg/试算表用它
                             #   显示一级行名——原 l1_agg 取第一个叶子名（1012→『承兑
                             #   保证金』错，应为父级『其他货币资金』）。
                             'l1_name': l1_names.get(l1) or _nm,
                             'l1': l1,
                             'qc': float(v.get('qc') or 0.0),
                             'jf': float(v.get('jf') or 0.0),
                             'df': float(v.get('df') or 0.0),
                             'qm': float(v.get('qm') or 0.0)}
            out[ent][yy] = ent_tb
    return out, years


def l1_agg(ent_tb):
    """按一级代码（4位）汇总叶子行 → {l1code: 值字典}（名称优先用父级行名 l1_name）。"""
    agg = OrderedDict()
    for c, v in sorted(ent_tb.items()):
        l1 = v.get('l1') or re.sub(r'\D', '', c)[:4]
        a = agg.setdefault(l1, {'name': v.get('l1_name') or v['name'],
                                'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0})
        a['qc'] += v['qc']; a['jf'] += v['jf']; a['df'] += v['df']; a['qm'] += v['qm']
    return agg


def balance_check(ent_tb):
    """叶子层平衡校验。返回 dict。"""
    l1 = l1_agg(ent_tb)
    jf = sum(v['jf'] for v in l1.values())
    df = sum(v['df'] for v in l1.values())
    qc_pos = sum(v['qc'] for v in l1.values() if v['qc'] >= 0)
    qc_neg = sum(v['qc'] for v in l1.values() if v['qc'] < 0)
    qm_pos = sum(v['qm'] for v in l1.values() if v['qm'] >= 0)
    qm_neg = sum(v['qm'] for v in l1.values() if v['qm'] < 0)
    # ⚡⚡ 2026-08-15 铁律128：未结转损益只含【损益类】(6/7/8 开头)——成本类(5 开头)
    #   期末余额是资本化资产（在产品/开发成本/合同履约成本），不是未结转损益！
    #   XBJ 实证：5006 房地产开发成本 3.95 亿已归入存货（资产），若再算进 pl_open 并入
    #   未分配利润 → 资产权益双计 → 资产负债表勾稽不平 3.95 亿。
    pl_open = sum(v['qm'] for c, v in l1.items() if code_class(c) == 'pl')
    return {
        'jf': jf, 'df': df, 'jf_df': jf - df,
        'qc_pos': qc_pos, 'qc_neg': qc_neg, 'qc_diff': qc_pos + qc_neg,
        'qm_pos': qm_pos, 'qm_neg': qm_neg, 'qm_diff': qm_pos + qm_neg,
        'pl_open': pl_open,
        'ok_jf': abs(jf - df) < 0.01,
        'ok_qm': abs(qm_pos + qm_neg) < 0.01,
    }


def check_retained_earnings(ent_tb, ent_is=None):
    """⚡ 铁律124（2026-08-15）：期初未分配 + 本期净利润 = 期末未分配 勾稽。

    若不平 = 存在【直接调整期初未分配利润】的分录（审计底稿体系无『期初未分配』
    『以前年度损益调整』科目，期初数只能通过调整分录整体调整）。
    注意：项目账（XBJ 等）损益科目期末不结转 → 未结转损益也是权益的一部分，
    且可能跨期遗留（期初就有）→ 公式：
        期初未分配 + 期初未结转损益 + 本期净利润 = 期末未分配 + 期末未结转损益
    返回 dict：qc_end(期初未分配带符号)/qm_end(期末未分配带符号)/
    qc_pl(期初未结转损益带符号)/pl_open(期末未结转损益带符号)/
    net(本期净利润)/diff(勾稽差，>0.005 即需查因)。
    符号约定：未分配利润贷余为负（带符号），本期净利润为正=盈利。
    """
    l1 = l1_agg(ent_tb)
    # 期初/期末未分配 = 利润分配/未分配利润/本年利润 类科目（名称含"利润"的权益类）
    qc_end = sum(float(v.get('qc') or 0.0) for c, v in l1.items()
                 if code_class(c) == 'equity' and '利润' in str(v.get('name') or ''))
    qm_end = sum(float(v.get('qm') or 0.0) for c, v in l1.items()
                 if code_class(c) == 'equity' and '利润' in str(v.get('name') or ''))
    # 期初/期末未结转损益 = 损益类科目（6/7/8 开头）期初/期末余额（贷余负=盈利留存）。
    # ⚡ 铁律128：成本类(5 开头)期末余额=资本化资产（在产品/开发成本），非未结转损益。
    qc_pl = sum(float(v.get('qc') or 0.0) for c, v in l1.items()
                if code_class(c) == 'pl')
    pl_open = sum(float(v.get('qm') or 0.0) for c, v in l1.items()
                  if code_class(c) == 'pl')
    net = float(ent_is.get('五、净利润') or 0.0) if ent_is else 0.0
    # 勾稽（贷正口径：期初未分配 + 期初未结转 + 净利 = 期末未分配 + 期末未结转）：
    #   diff = (期末-期初)权益变动 - 净利润 = 直接调整期初未分配的金额（贷正=调增期初）
    diff = (qc_end + qc_pl) - (qm_end + pl_open) - net
    # 待查分录：利润分配/未分配利润类科目 本期发生额明细（正常=结转净利/提取盈余公积/分配股利；
    #   非正常=直接调期初，需审计判断）→ {code: {name, qc, jf, df, qm}}
    det = {}
    for c, v in l1.items():
        if code_class(c) == 'equity' and '利润' in str(v.get('name') or ''):
            if abs(float(v.get('qc') or 0)) > 0.005 or abs(float(v.get('jf') or 0)) > 0.005 \
                    or abs(float(v.get('df') or 0)) > 0.005 or abs(float(v.get('qm') or 0)) > 0.005:
                det[c] = {'name': v['name'], 'qc': v['qc'], 'jf': v['jf'],
                          'df': v['df'], 'qm': v['qm']}
    return {'qc_end': qc_end, 'qm_end': qm_end, 'qc_pl': qc_pl,
            'pl_open': pl_open, 'net': net, 'diff': diff, 'det': det}


def build_bs(ent_tb_all, yy):
    """资产负债表：按报表行映射汇总（备抵带符号抵减原值），未结转损益入权益。"""
    ent_rows = {}
    ent_unmatched = {}
    for ent, tb in ent_tb_all.items():
        ent_tb = tb.get(yy, {})
        l1 = l1_agg(ent_tb)
        all_cfg = dict(list(BS_ASSET_ROWS.items()) + list(BS_LIAB_ROWS.items()) + list(BS_EQUITY_ROWS.items()))
        assign = {}
        unmatched = []
        for c, v in l1.items():
            nm = v['name']
            rn = None
            for rname, cfg in all_cfg.items():
                if any(k in nm for k in cfg.get('kw', [])):
                    rn = rname
                    break
            if rn is None:
                for kw, target in CONTRA_MAP.items():
                    if kw in nm:
                        rn = target
                        break
            if rn:
                assign[c] = rn
            else:
                unmatched.append((c, nm))
        ent_unmatched[ent] = unmatched
        rows = defaultdict(float)
        for c, v in l1.items():
            rn = assign.get(c)
            if rn:
                rows[rn] += v['qm']   # 备抵贷余为负 → 自动抵减
            elif code_class(c) == 'pl':
                continue              # 损益类期末余额由「未结转损益」行统一处理（不入兜底）
            else:
                rows['其他（未匹配）'] += v['qm']   # 兜底行（含成本类资本化科目期末余额，按符号收纳，勾稽恒平）
        # 未结转损益 → 权益（2026-08-15 铁律124b：未结转损益【并入未分配利润】列示——
        #   报表口径「未分配利润」= 利润分配+本年利润+未结转损益（未结转损益=损益科目期末
        #   余额，实质是未分配利润的过渡态；单列会导致未分配利润数≠报表口径，被误读）。
        #   pl_open 仍保留在勾稽行「其中：未结转损益」供核对。
        bc = balance_check(ent_tb)
        if abs(bc['pl_open']) > 0.005:
            rows['未分配利润'] += bc['pl_open']   # 并入未分配利润（带符号）
        ent_rows[ent] = rows
    return ent_rows, ent_unmatched


def build_equity_sheet_gl(ent_tb_all, yy, data_dir):
    """⚡ 所有者权益变动表数据（2026-08-15 铁律124c：上交利润/盈余公积=利润分配，权益变动表列示，
    非期初调整）。未分配利润口径（贷正）：
      期初未分配(含未结转) + 净利润 − 提取盈余公积 − 上交利润分配 + 直接调期初 = 期末未分配(含未结转)
    其中 直接调期初 = 勾稽差 diff − (盈余公积 + 上交利润)（反推，保证勾稽恒平且精确反映期初调整）。
    返回 {label: {ent: 值}}，label 同上。"""
    import glob as _g, re as _re, os as _os, openpyxl as _xl
    out = {'期初未分配': {}, '净利润': {}, '提取盈余公积': {}, '上交利润分配': {},
           '直接调期初': {}, '期末未分配': {}}
    ents = sorted(ent_tb_all)
    ent2file = {}
    for e in ents:
        fs = _g.glob(_os.path.join(data_dir, _re.escape(e[:2]) + '*' + '综合查询明细表*.xlsx'))
        ent2file[e] = fs[0] if fs else None
    for e in ents:
        ent_tb = ent_tb_all[e].get(yy, {})
        l1 = l1_agg(ent_tb)
        # 期初/期末未分配（含未结转损益，贷正口径）
        qc_end = sum(float(v.get('qc') or 0.0) for c, v in l1.items()
                     if code_class(c) == 'equity' and '利润' in str(v.get('name') or ''))
        qm_end = sum(float(v.get('qm') or 0.0) for c, v in l1.items()
                     if code_class(c) == 'equity' and '利润' in str(v.get('name') or ''))
        qc_pl = sum(float(v.get('qc') or 0.0) for c, v in l1.items()
                    if code_class(c) == 'pl')
        pl_open = sum(float(v.get('qm') or 0.0) for c, v in l1.items()
                      if code_class(c) == 'pl')
        out['期初未分配'][e] = -(qc_end + qc_pl)
        out['期末未分配'][e] = -(qm_end + pl_open)
        is_rows = build_is({e: ent_tb_all[e]}, yy)
        out['净利润'][e] = float(is_rows.get(e, {}).get('五、净利润') or 0.0)
        # GL 利润分配分录分类（盈余公积 / 上交利润）
        sr = gj = 0.0
        fp = ent2file.get(e)
        if fp:
            try:
                wb = _xl.load_workbook(fp, read_only=True, data_only=True)
                ws = wb[wb.sheetnames[0]]
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if not r or not r[0]:
                        continue
                    nm = str(r[0])
                    if '利润' not in nm:
                        continue
                    jf = float(r[6] or 0); df = float(r[7] or 0)
                    if abs(jf) < 0.005 and abs(df) < 0.005:
                        continue
                    sm = str(r[5]); net = df - jf
                    if '盈余公积' in nm:
                        sr += net
                    elif ('上交利润' in nm or '所属上交' in nm or '管理费' in sm or '上交' in sm
                          or '退还' in sm or '补交' in sm or '代付' in sm or '代缴' in sm):
                        gj += net
                wb.close()
            except Exception:
                pass
        out['提取盈余公积'][e] = sr
        out['上交利润分配'][e] = gj
        # 直接调期初 = diff − (盈余公积 + 上交利润)（反推，勾稽恒平）
        chk = check_retained_earnings(ent_tb, is_rows.get(e))
        out['直接调期初'][e] = chk['diff'] - (sr + gj)
    return out


def build_is(ent_tb_all, yy):
    """利润表：收入取贷方发生额/费用取借方发生额（铁律3：损益 gross）。
    ⚡ 铁律130b（SAP 模式）：SAP 内部码 ≠ 标准码（6010收入/64xx成本/66xx费用），
    IS_ROWS 的 code 前缀对 SAP 无效 → 按 cfg.kw 名称关键词匹配（ga 之外 AH 实证）。"""
    global _SAP_MODE
    ent_is = {}
    for ent, tb in ent_tb_all.items():
        ent_tb = tb.get(yy, {})
        l1 = l1_agg(ent_tb)
        rows = {}
        for rn, cfg in IS_ROWS.items():
            if cfg.get('calc'):
                continue
            tot = 0.0
            for c, v in l1.items():
                _nm = str(v.get('name') or '')
                _hit = False
                for k in cfg.get('credit', []):
                    if c.startswith(k):
                        tot += v['df']
                        _hit = True
                for k in cfg.get('debit', []):
                    if c.startswith(k):
                        tot += v['jf']
                        _hit = True
                # ⚡ SAP 模式：code 前缀未命中 → 名称关键词匹配（行级一次，防双计）
                if not _hit and _SAP_MODE and cfg.get('kw') \
                        and any(_k in _nm for _k in cfg['kw']):
                    if cfg.get('credit'):
                        tot += v['df']
                    if cfg.get('debit'):
                        tot += v['jf']
            rows[rn] = tot
        # ⚡⚡ 2026-08-15 XBJ 无 6001 主体收入兜底（与 revenue_detail 一致）：
        #   收入科目只有『合同结算\价款结算』(123301) 的主体，其【贷方正数】=本期收入确认
        #   （38 个无 6001 主体实证：123301贷=123302借=收入）；负贷方=停用/红字冲销（非收入）。
        #   ⚠️ 仅当该主体无 6001 系科目才启用——有 6001 主体误用 123301 会把跨期结算当收入。
        #   ⚠️ read_tb_full 读出的 123301 name 只有父级『合同结算』（子目名丢失），故按 code 判断。
        if '一、营业收入' in rows:
            _has_6001 = any(str(c).startswith('6001') for c in ent_tb)
            if not _has_6001:
                for c, v in ent_tb.items():
                    if str(c) == '123301' and float(v.get('df') or 0.0) > 0.005:
                        rows['一、营业收入'] += float(v.get('df') or 0.0)
                        break
        # 利润总额 = 收入合计 - 费用合计（除所得税）
        income = 0.0
        expense = 0.0
        for rn, cfg in IS_ROWS.items():
            if cfg.get('calc'):
                continue
            if rn == '减：所得税费用':
                continue
            if 'credit' in cfg:
                income += rows[rn]
            if 'debit' in cfg:
                expense += rows[rn]
        rows['四、利润总额'] = income - expense
        rows['五、净利润'] = rows['四、利润总额'] - rows.get('减：所得税费用', 0.0)
        ent_is[ent] = rows
    return ent_is


# ---------------------------------------------------------------- 输出
def _write_rows(ws, header, rows, bold_rows=()):
    ws.append(header)
    for i, row in enumerate(rows):
        ws.append(row)
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = F
            cell.border = THIN
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
        if i - 1 in bold_rows:
            for cell in row:
                cell.font = FB


def write_workbook(out_path, ent_tb_all, years):
    wb = openpyxl.Workbook()
    ents = sorted(ent_tb_all)
    # ⚡⚡ 2026-08-15 修复：主展示年度=【数据最全年度】而非 years[0]（字母序 2020 排前）。
    #   修复前 _extract_year 误判使 years 含 2020/2022 空年度 → years[0]='2020' → Sheet2 全 0。
    #   按各主体 TB 行数最多的年度选主表年度（正常账套单年度不受影响）。
    def _rows_of(y):
        return sum(len(ent_tb_all[e].get(y, {})) for e in ents)
    yy = max(years, key=lambda y: _rows_of(y))

    # ============ Sheet1 平衡校验 ============
    ws = wb.active
    ws.title = '平衡校验'
    header = ['主体', '年度', '借发合计', '贷发合计', '发生额差',
              '期初借余', '期初贷余', '期初差', '期末借余', '期末贷余', '期末差',
              '未结转损益', '状态']
    ws.append(header)
    for e in ents:
        for y in years:
            bc = balance_check(ent_tb_all[e].get(y, {}))
            st = []
            st.append('借发=贷发✓' if bc['ok_jf'] else '借发≠贷发(差=未结转损益)!')
            st.append('期末平✓' if bc['ok_qm'] else '期末差%.2f(未结转损益)!' % bc['qm_diff'])
            if abs(bc['pl_open']) > 0.005:
                st.append('未结转损益%.2f' % bc['pl_open'])
            ws.append([e, y, round(bc['jf'], 2), round(bc['df'], 2), round(bc['jf_df'], 2),
                       round(bc['qc_pos'], 2), round(bc['qc_neg'], 2), round(bc['qc_diff'], 2),
                       round(bc['qm_pos'], 2), round(bc['qm_neg'], 2), round(bc['qm_diff'], 2),
                       round(bc['pl_open'], 2), '；'.join(st)])
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = F
            cell.border = THIN
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
        row[-1].font = FR if '!' in str(row[-1].value) else FG
    widths = {'A': 14}
    for col in 'BCDEFGHIJKL':
        widths[col] = 14
    widths['M'] = 40
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ============ Sheet2 试算表（一级科目 × 主体，期末余额） ============
    ws2 = wb.create_sheet('试算表')
    l1_names = {}
    for e in ents:
        for c, v in l1_agg(ent_tb_all[e].get(yy, {})).items():
            l1_names.setdefault(c, v['name'])
    ws2.append(['科目代码', '科目名称'] + ents + ['集团合计'])
    grand_tot = 0.0
    # ⚡⚡ 2026-08-15 用户方法论：按 资产→负债→权益→损益 报表科目顺序排列 + 各段合计行。
    #   试算表是实际一级科目（原材料/主营业务收入…），段分类按【科目代码首位】优先、
    #   标准名清单兜底（名称匹配不上的实际科目码分类更可靠）。
    _SEG_LABEL = {'asset': '资产合计', 'liab': '负债合计', 'equity': '所有者权益合计',
                  'pl_income': '收入合计', 'pl_cost': '成本费用合计', 'other': '其他合计'}
    _REPORT_CAT_ORDER = {'asset': 0, 'liab': 1, 'equity': 2, 'pl_income': 3, 'pl_cost': 4, 'other': 5}
    seg_tot = {s: 0.0 for s in _SEG_LABEL}
    seg_bye = {s: [0.0] * len(ents) for s in _SEG_LABEL}

    def _seg_of(code, name):
        c = str(code or '')[:1]
        nm = str(name or '')
        if c == '1':
            return 'asset'
        if c == '2':
            return 'liab'
        if c in ('3', '4'):
            return 'equity'
        if c in ('5',):
            return 'pl_cost'
        if c == '6':
            return 'pl_income' if any(k in nm for k in ('收入', '收益', '利得')) else 'pl_cost'
        cat, _ = A.report_cat(nm)
        return cat if cat != 'other' else 'other'

    # ⚡⚡ 2026-08-23 引入 2-3 级科目：一级汇总行 + 子目明细行，按 段→一级→级次→code 排序
    #   （用户需求：试算表引入对应 2-3 级科目、按报表顺序排列，方便核对各科目底稿数据）。
    #   一级行用 l1_agg 汇总（跨子目），子目行用叶子值；段合计只按一级行累加防重复。
    all_codes = set()
    for e in ents:
        all_codes |= set(ent_tb_all[e].get(yy, {}))

    def _nm_of(code):
        for e in ents:
            v = ent_tb_all[e].get(yy, {}).get(code)
            if v and v.get('name'):
                return v['name']
        return str(code)

    def _l1_of(code):
        return re.sub(r'\D', '', str(code))[:4] or str(code)

    def _pos(v):
        return -v if v < 0 else v

    l1_order = sorted(l1_names,
                      key=lambda cc: (_REPORT_CAT_ORDER[_seg_of(cc, l1_names[cc])], str(cc)))
    for c in l1_order:
        # —— 一级汇总行 ——
        row = [c, l1_names[c]]
        tot = 0.0
        bye = []
        for e in ents:
            v = l1_agg(ent_tb_all[e].get(yy, {})).get(c, {}).get('qm', 0.0)
            _d = _pos(v)
            row.append(round(_d, 2))
            tot += _d
            bye.append(round(_d, 2))
        row.append(round(tot, 2))
        grand_tot += tot
        sg = _seg_of(c, l1_names[c])
        seg_tot[sg] += tot
        for i, v in enumerate(bye):
            seg_bye[sg][i] += v
        ws2.append(row)
        # —— 子目明细行（2-3 级，紧跟其一级汇总行之后）——
        subs = sorted([cc for cc in all_codes if _l1_of(cc) == c and cc != c],
                      key=lambda cc: (cc.count('.'), str(cc)))
        for sub in subs:
            srow = [sub, _nm_of(sub)]
            stot = 0.0
            for e in ents:
                v = ent_tb_all[e].get(yy, {}).get(sub, {}).get('qm', 0.0)
                _d = _pos(v)
                srow.append(round(_d, 2))
                stot += _d
            srow.append(round(stot, 2))
            ws2.append(srow)
    # 段合计行（加粗）
    for sg in ('asset', 'liab', 'equity', 'pl_income', 'pl_cost', 'other'):
        if seg_tot.get(sg, 0.0) != 0.0 or any(abs(v) > 0.005 for v in seg_bye.get(sg, [])):
            ws2.append(['', _SEG_LABEL[sg]]
                       + [round(v, 2) for v in seg_bye[sg]] + [round(seg_tot[sg], 2)])
    # 平衡行（转正口径合计；带符号勾稽由资产负债表承担）
    ws2.append(['', '期末余额合计（贷余转正口径）'] + [''] * len(ents) + [round(grand_tot, 2)])
    for i, row in enumerate(ws2.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = F
            cell.border = THIN
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
            if row[1] and ('合计' in str(row[1])):
                cell.font = FB
    ws2.column_dimensions['A'].width = 10
    ws2.column_dimensions['B'].width = 22

    # ============ Sheet3 资产负债表 ============
    ws3 = wb.create_sheet('资产负债表')
    ent_rows, ent_unmatched = build_bs(ent_tb_all, yy)
    ws3.append(['项目'] + ents + ['集团合计'])
    blocks = [('资 产', BS_ASSET_ROWS), ('负 债', BS_LIAB_ROWS), ('所有者权益', BS_EQUITY_ROWS)]
    for sec, rows_cfg in blocks:
        ws3.append([sec])
        # 2026-08-07 修复：报表展示口径——负债/权益贷余转正（用户 TB 负债正数，
        # 原输出贷余负 → 『负债用负数表示』质疑）。资产侧保持借余正；勾稽用带符号
        # 合计在下方另行标注（资产=负债+权益，均正数）。
        _sign = -1.0 if sec in ('负 债', '所有者权益') else 1.0
        for rn in rows_cfg:
            row = [rn]
            tot = 0.0
            for e in ents:
                v = ent_rows[e].get(rn, 0.0) * _sign
                row.append(round(v, 2))
                tot += v
            row.append(round(tot, 2))
            ws3.append(row)
    # 总计与勾稽（带符号：资产借余正/负债·权益贷余负；资产+负债+权益+其他=0，差额=未结转损益）
    def _tot_of(rows_cfg):
        return [sum(ent_rows[e].get(rn, 0.0) for rn in rows_cfg) for e in ents]
    asset_tot = _tot_of(BS_ASSET_ROWS)
    liab_tot = _tot_of(BS_LIAB_ROWS)
    eq_tot = _tot_of(BS_EQUITY_ROWS)
    other_tot = _tot_of(BS_OTHER_ROWS)
    pl_opens = [balance_check(ent_tb_all[e].get(yy, {})).get('pl_open', 0.0) for e in ents]

    def _totrow(label, vals):
        return [label] + [round(v, 2) for v in vals] + [round(sum(vals), 2)]
    ws3.append([])
    ws3.append(_totrow('资产总计', asset_tot))
    # 2026-08-07：负债/权益总计转正（与上方明细行转正口径一致，报表格式正数）
    ws3.append(_totrow('负债总计', [-v for v in liab_tot]))
    ws3.append(_totrow('权益总计', [-v for v in eq_tot]))
    ws3.append(_totrow('其他（未匹配）', other_tot))
    ws3.append(_totrow('勾稽（资产+负债+权益+其他，=0）',
                       [asset_tot[i] + liab_tot[i] + eq_tot[i] + other_tot[i] for i in range(len(ents))]))
    ws3.append(_totrow('其中：未结转损益', pl_opens))
    for i, row in enumerate(ws3.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = F
            cell.border = THIN
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
    ws3.column_dimensions['A'].width = 24

    # ============ Sheet4 利润表 ============
    ws4 = wb.create_sheet('利润表')
    ent_is = build_is(ent_tb_all, yy)
    ws4.append(['项目'] + ents + ['集团合计'])
    for rn in IS_ROWS:
        row = [rn]
        tot = 0.0
        for e in ents:
            v = ent_is[e].get(rn, 0.0)
            row.append(round(v, 2))
            tot += v
        row.append(round(tot, 2))
        ws4.append(row)
    for i, row in enumerate(ws4.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = F
            cell.border = THIN
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
    ws4.column_dimensions['A'].width = 24

    # ============ Sheet5 期初未分配勾稽（铁律124：期初未分配+净利=期末未分配）============
    ws5 = wb.create_sheet('期初未分配勾稽')
    ws5.append(['期初未分配勾稽（铁律124：期初未分配 + 本期净利润 = 期末未分配 + 未结转损益）',
                '勾稽差≠0 = 存在【直接调整期初未分配利润】的分录，需找出分录形成底稿做调整期初分录'])
    ws5.append(['主体', '期初未分配', '期初未结转', '本期净利润', '期末未分配', '期末未结转',
                '勾稽差', '待查分录（利润分配发生额）'])
    ent_is = build_is(ent_tb_all, yy)
    _chk_n = 0
    for e in ents:
        chk = check_retained_earnings(ent_tb_all[e].get(yy, {}), ent_is.get(e))
        if abs(chk['diff']) <= 0.005:
            continue
        _chk_n += 1
        _det = '；'.join(f"{d['name']}(借{round(d['jf'],2)}/贷{round(d['df'],2)})"
                         for d in chk['det'].values())
        ws5.append([e, round(chk['qc_end'], 2), round(chk['qc_pl'], 2), round(chk['net'], 2),
                    round(chk['qm_end'], 2), round(chk['pl_open'], 2),
                    round(chk['diff'], 2), _det])
    for i, row in enumerate(ws5.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = F
            cell.border = THIN
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
    ws5.column_dimensions['A'].width = 40
    for j, w in enumerate([16, 16, 14, 16, 16, 14, 14, 60], 2):
        ws5.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    print(f'  ⚠️ 期初未分配勾稽不平（直接调期初候选）: {_chk_n} 主体，见 Sheet5 期初未分配勾稽')

    # ============ Sheet6 所有者权益变动表（铁律124c：上交利润/盈余公积=利润分配列示）============
    ws6 = wb.create_sheet('权益变动表')
    ws6.append(['所有者权益变动表（2025 年度）· 期初未分配 + 净利润 + 提取盈余公积(带符号) + 上交利润分配(带符号) + 直接调期初 = 期末未分配'])
    ws6.append(['项目', '期初未分配', '净利润', '提取盈余公积', '上交利润分配', '直接调期初', '期末未分配', '勾稽校验'])
    _eqv = build_equity_sheet_gl(ent_tb_all, yy, os.path.dirname(os.path.abspath(out_path)))
    _eq_labels = ['期初未分配', '净利润', '提取盈余公积', '上交利润分配', '直接调期初', '期末未分配']
    for e in ents:
        row = [e]
        for lb in _eq_labels:
            row.append(round(_eqv[lb].get(e, 0.0), 2))
        # 勾稽：期初 + 净利 + 盈余公积(带符号) + 上交利润(带符号) + 直接调期初 = 期末（全带符号）
        chk_eq = (_eqv['期初未分配'].get(e, 0.0) + _eqv['净利润'].get(e, 0.0)
                  + _eqv['提取盈余公积'].get(e, 0.0) + _eqv['上交利润分配'].get(e, 0.0)
                  + _eqv['直接调期初'].get(e, 0.0) - _eqv['期末未分配'].get(e, 0.0))
        row.append(round(chk_eq, 2))
        ws6.append(row)
    for i, row in enumerate(ws6.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.font = F
            cell.border = THIN
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
    ws6.column_dimensions['A'].width = 40
    for j, w in enumerate([14, 16, 14, 14, 14, 14, 16, 12], 2):
        ws6.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w

    wb.save(out_path)
    wb.close()
    return out_path


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument('folder')
    ap.add_argument('--out')
    ap.add_argument('--year', default=None)
    ap.add_argument('--comps', default=None, help='主体白名单（逗号分隔，集团分组用）')
    a = ap.parse_args(argv)
    data_dir = a.folder
    comps = [x.strip() for x in a.comps.split(',') if x.strip()] if a.comps else None
    ent_tb_all, years = build_entity_tb(data_dir, [a.year] if a.year else None, comps=comps)
    out = a.out or os.path.join(data_dir, '自建试算表_%s.xlsx' % years[0])
    write_workbook(out, ent_tb_all, years)
    print('自建试算表已生成: %s' % out)
    print('主体 %d 个 | 年度 %s' % (len(ent_tb_all), ','.join(years)))
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
