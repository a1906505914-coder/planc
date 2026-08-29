# -*- coding: utf-8 -*-
"""项目制核算 收入底稿（铁律131f，2026-08-16 用户需求；降级模式：无合同台账）。

按【项目/合同】编制收入底稿。项目维度键（manifest project_key）：
  · aux_unit = 辅助核算「往来单位名称」（XBJ/U8：业主名；主体=项目部=项目）
  · sap_wbs  = SAP WBS 元素（col17，AH 迁移后启用）
  · sap_contract = SAP 合同号（col15）

数据源（无台账降级模式——账务侧全可程序取）：
  收入      GL 贷方（主营业务收入/其他业务收入/合同结算\收入结转）→ 按项目键聚合
  成本      GL 借方（合同履约成本/主营业务成本）→ 按项目键聚合（XBJ 成本无单位→主体级）
  毛利率    (收入-成本)/收入
  收款      GL 银行存款借方/应收贷方 → 按项目键聚合
  期末应收  应收账款辅助核算（往来单位=业主）
  合同要素  客户/内容/周期/约定金额/预算成本 → 留空（待台账，manifest contracts 配置后填）

产出：底稿/{year}/项目制收入底稿_{year}.xlsx（8 sheets）：
  ①项目合同台账  ②按项目收入成本确认表(时段法)  ③与账面核对  ④收款与应收款核对
  ⑤毛利异常分析  ⑥项目-成本构成  ⑦口径说明  ⑧审计调整

用法：python project_review.py [acct] [sub]   （缺省=全部账套）
"""
import os
import sys
import re
from collections import defaultdict

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import project_manifest as PM
from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_BOLD,
                         SHELL_NUM, SHELL_CEN, SHELL_RGT, SHELL_BORDER,
                         SHELL_TOT_FILL, finalize_workbook)
from openpyxl.styles import Font

try:
    import sap_adapter as A
except Exception:
    A = None

# 施工企业收入/成本 名称关键词（科目名列=名称，铁律5）
REV_KW = ['主营业务收入', '其他业务收入', '合同结算\\收入结转']
COST_KW = ['合同履约成本', '主营业务成本', '合同结算\\工程施工']
RECV_KW = ['应收账款']
BANK_KW = ['银行存款', '内部结算中心存款']
COST_CAT_KW = ['直接材料', '直接人工', '机械使用', '分包', '间接费用', '其他费用']


def _f(v):
    try:
        return float(v or 0.0)
    except (TypeError, ValueError):
        return 0.0


def _txt(ws, r, c, v, bold=False, align=None, italic=False, color=None):
    cell = ws.cell(r, c, v)
    cell.border = SHELL_BORDER
    if bold:
        cell.font = SHELL_BOLD
    elif italic:
        cell.font = Font(name='Times New Roman', size=10, italic=True)
    if align:
        cell.alignment = align
    if color:
        cell.font = Font(name='Times New Roman', size=10, bold=True, color=color)
    return cell


def _money(ws, r, c, v):
    cell = ws.cell(r, c, round(_f(v), 2) if v is not None else None)
    cell.number_format = SHELL_NUM
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    return cell


def _hdr(ws, row, headers):
    for j, h in enumerate(headers, 1):
        c = ws.cell(row, j, h)
        c.fill = SHELL_HFILL
        c.font = SHELL_HFONT
        c.border = SHELL_BORDER
        c.alignment = SHELL_CEN
    return row + 1


def _read_aux(data_dir, ent):
    """读主体辅助核算余额表 → {'科目名': {'unit': {方向,期初,借,贷,期末}}}。
    文件=<数据目录>/{主体名}...辅助核算余额表.xlsx。列=科目名称/往来单位/期初方向/期初金额/
    本期借方/本期贷方/期末方向/期末金额。"""
    import glob
    pat = os.path.join(data_dir, f'*{ent}*辅助核算余额表*.xlsx')
    files = glob.glob(pat)
    if not files:
        return {}
    try:
        wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception:
        return {}
    hi = 0
    for i, r in enumerate(rows[:6]):
        if r and any('科目' in str(x) for x in r):
            hi = i
            break
    out = {}
    for r in rows[hi + 1:]:
        if not r or len(r) < 8:
            continue
        km = str(r[0] or '').strip()
        unit = str(r[1] or '').strip()
        qc_dir = str(r[2] or '')
        qc = _f(r[3])
        jf = _f(r[4])
        df = _f(r[5])
        qm_dir = str(r[6] or '')
        qm = _f(r[7])
        d = out.setdefault(km, {})
        u = d.setdefault(unit, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0})
        u['qc'] += qc
        u['jf'] += jf
        u['df'] += df
        u['qm'] += qm
    return out


def _read_gl(data_dir, ent):
    """读主体 GL（综合查询明细表）→ [{name, date, jf, df, sm}]。
    列=科目名称/日期/凭证字/凭证号/对方科目/摘要/借方金额/贷方金额。"""
    import glob
    pat = os.path.join(data_dir, f'*{ent}*综合查询明细表*.xlsx')
    files = glob.glob(pat)
    if not files:
        return []
    try:
        wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception:
        return []
    hi = 0
    for i, r in enumerate(rows[:6]):
        if r and any('科目' in str(x) for x in r) and any('金额' in str(x) for x in r):
            hi = i
            break
    out = []
    for r in rows[hi + 1:]:
        if not r or len(r) < 8:
            continue
        out.append({'name': str(r[0] or ''),
                    'date': str(r[1] or '')[:10],
                    'sm': str(r[5] or ''),
                    'jf': _f(r[6]), 'df': _f(r[7])})
    return out


def _read_tb(data_dir, ent):
    """读主体科目余额表 → {科目编码: {name, opening, debit, credit, closing}}。
    列=科目编码/科目名称/期初方向/期初金额/本期借方/本期贷方/期末方向/期末金额/等级。"""
    import glob
    pat = os.path.join(data_dir, f'*{ent}*科目余额表*.xlsx')
    files = glob.glob(pat)
    if not files:
        return {}
    try:
        wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception:
        return {}
    hi = 0
    for i, r in enumerate(rows[:6]):
        if r and any('科目' in str(x) for x in r) and any('金额' in str(x) for x in r):
            hi = i
            break
    # 列位置（表头 r0：科目编码/科目名称/期初方向/期初金额/本期借方/本期贷方/期末方向/期末金额）
    hdr = rows[hi] if hi < len(rows) else ()
    i_code = i_name = i_qc = i_jf = i_df = i_qm = 0
    for i, h in enumerate(hdr):
        hs = str(h or '').strip()
        if '编码' in hs or '代码' in hs:
            i_code = i
        elif '名称' in hs and '科目' in hs:
            i_name = i
        if '期初' in hs and '金额' in hs:
            i_qc = i
        elif '本期借方' in hs or '借方' == hs:
            i_jf = i
        elif '本期贷方' in hs or '贷方' == hs:
            i_df = i
        elif '期末' in hs and '金额' in hs:
            i_qm = i
    out = {}
    for r in rows[hi + 1:]:
        if not r or len(r) < 8:
            continue
        code = str(r[i_code] or '').strip()
        nm = str(r[i_name] or '').strip()
        if not code and not nm:
            continue
        out[code] = {'name': nm,
                     'opening': _f(r[i_qc]) if i_qc < len(r) else 0.0,
                     'debit': _f(r[i_jf]) if i_jf < len(r) else 0.0,
                     'credit': _f(r[i_df]) if i_df < len(r) else 0.0,
                     'closing': _f(r[i_qm]) if i_qm < len(r) else 0.0}
    return out


def _project_rows(acct_cfg, sub_cfg, data_dir):
    """聚合全部主体 → 项目行列表。
    行=项目（XBJ：主体=项目部=项目；aux_unit 时客户列取辅助核算业主名）。
    收入/成本=**主体级全量**（GL 同源，与 TB 勾稽）；业主仅作客户维度与应收拆分。
    返回 [dict]：ent/unit/customer/rev_prev/rev_cur/cost_prev/cost_cur/recv_qc/recv_qm/
                  settle(累计结算)/rev_cum(累计已确认收入)。"""
    import glob
    ents = sorted({os.path.basename(f).split('年')[0] for f in
                   glob.glob(os.path.join(data_dir, '*科目余额表*.xlsx'))})
    rows = []
    rev_kw = REV_KW
    cost_kw = COST_KW
    for ent in ents:
        aux = _read_aux(data_dir, ent)
        gl = _read_gl(data_dir, ent)
        # —— 账务侧聚合（主体级全量，保证与 TB 勾稽）——
        rev_cur = 0.0   # 本期确认收入（GL 贷方：主营/其他/收入结转）
        cost_cur = 0.0  # 本期成本（GL 借方：合同履约成本/主营成本）
        for r in gl:
            nm = r['name']
            if any(k in nm for k in rev_kw):
                rev_cur += r['df']          # 收入取贷方（铁律3）
            elif any(k in nm for k in cost_kw):
                cost_cur += r['jf']         # 成本取借方
        # 无收入无成本且无辅助核算往来 → 跳过（空壳主体）
        if abs(rev_cur) < 0.005 and abs(cost_cur) < 0.005 and not aux:
            continue
        # —— 业主维度（客户列/应收/结算）——
        units = {}
        for km, uinfo in aux.items():
            if km.startswith('合同结算') and '价款结算' in km:
                for u, v in uinfo.items():
                    if u:
                        d = units.setdefault(u, {'settle': 0.0, 'recv_qm': 0.0, 'recv_qc': 0.0,
                                                 'rev': 0.0})
                        # 累计已结算额=期末余额（贷方为结算额；借方正数=完工对冲红字，带符号）
                        d['settle'] += _f(v['qm']) if _f(v['qm']) else (_f(v['df']) - _f(v['jf']))
            if km.startswith('合同结算') and '收入结转' in km:
                for u, v in uinfo.items():
                    if u:
                        d = units.setdefault(u, {'settle': 0.0, 'recv_qm': 0.0, 'recv_qc': 0.0,
                                                 'rev': 0.0})
                        d['rev'] += _f(v['qm']) + _f(v['df']) - _f(v['jf'])  # 累计已确认收入
            if any(k in km for k in RECV_KW):
                for u, v in uinfo.items():
                    if u:
                        d = units.setdefault(u, {'settle': 0.0, 'recv_qm': 0.0, 'recv_qc': 0.0,
                                                 'rev': 0.0})
                        d['recv_qm'] += _f(v['qm'])
                        d['recv_qc'] += _f(v['qc'])
        if not units:
            units = {'（未识别业主）': {'settle': 0.0, 'recv_qm': 0.0, 'recv_qc': 0.0, 'rev': 0.0}}
        # 主体级收入/成本 全量计入每个业主行会造成重复——正确做法：
        #   Σ业主行 收入=主体收入（一个主体一行业主聚合），成本只在主体维度列示。
        # 简化：主体→业主 1:N 时，收入/成本/毛利率在【主体汇总行】体现，业主行仅客户/应收/结算。
        if len(units) == 1:
            u = next(iter(units))
            d = units[u]
            rows.append({
                'ent': ent, 'unit': u,
                'customer': u,
                'content': '', 'start': '', 'end': '',
                'contract_amount': d.get('settle', 0.0),
                'budget_cost': '',
                'rev_prev': 0.0, 'rev_cur': rev_cur,
                'cost_prev': 0.0, 'cost_cur': cost_cur,
                'settle': d.get('settle', 0.0),
                'recv_qc': d.get('recv_qc', 0.0), 'recv_qm': d.get('recv_qm', 0.0),
                'rev_cum': d.get('rev', 0.0),
            })
        else:
            # 多业主：主体汇总行 + 业主明细行（业主行 收入/成本=0 避免重复）
            rows.append({
                'ent': ent, 'unit': '（多业主）',
                'customer': '、'.join(sorted(units)),
                'content': '', 'start': '', 'end': '',
                'contract_amount': sum(_f(d.get('settle', 0.0)) for d in units.values()),
                'budget_cost': '',
                'rev_prev': 0.0, 'rev_cur': rev_cur,
                'cost_prev': 0.0, 'cost_cur': cost_cur,
                'settle': sum(_f(d.get('settle', 0.0)) for d in units.values()),
                'recv_qc': sum(_f(d.get('recv_qc', 0.0)) for d in units.values()),
                'recv_qm': sum(_f(d.get('recv_qm', 0.0)) for d in units.values()),
                'rev_cum': sum(_f(d.get('rev', 0.0)) for d in units.values()),
            })
            for u, d in units.items():
                rows.append({
                    'ent': ent, 'unit': u,
                    'customer': u,
                    'content': '', 'start': '', 'end': '',
                    'contract_amount': _f(d.get('settle', 0.0)),
                    'budget_cost': '',
                    'rev_prev': 0.0, 'rev_cur': 0.0,
                    'cost_prev': 0.0, 'cost_cur': 0.0,
                    'settle': _f(d.get('settle', 0.0)),
                    'recv_qc': _f(d.get('recv_qc', 0.0)),
                    'recv_qm': _f(d.get('recv_qm', 0.0)),
                    'rev_cum': _f(d.get('rev', 0.0)),
                })
    return rows


def _sheet1_contracts(ws, rows, cfg):
    """①项目合同台账（合同要素；无台账时按项目键列出账务侧可识别项目，要素留空）。"""
    ws.cell(1, 1, '项目合同台账（合同要素需台账补充；未提供前留空——账务侧已识别项目列示）').font = SHELL_TITLE_FONT
    r = _hdr(ws, 3, ['核算主体', '项目/合同键', '客户名称', '合同内容', '合同开始', '合同结束',
                     '合同约定收入(含税?)', '预算总成本', '确认方法',
                     '结算周期', '质保期(月)', '付款条件说明'])
    contracts = {c.get('project', ''): c for c in cfg.get('contracts', [])}
    for row in rows:
        key = row['unit']
        c = contracts.get(key, {})
        _txt(ws, r, 1, row['ent'])
        _txt(ws, r, 2, key)
        _txt(ws, r, 3, c.get('customer', row['customer']) or '')
        _txt(ws, r, 4, c.get('content', row.get('content', '')) or '')
        _txt(ws, r, 5, c.get('start', '') or '')
        _txt(ws, r, 6, c.get('end', '') or '')
        _money(ws, r, 7, c.get('contract_amount', 0.0) or 0.0)
        _money(ws, r, 8, c.get('budget_cost', 0.0) or 0.0)
        _txt(ws, r, 9, c.get('method', '投入法(完工百分比)') if c else '投入法(完工百分比)')
        # ⚡ 付款条件（2026-08-16）：结算周期/质保期 → 合同资产账龄判断依据
        _txt(ws, r, 10, c.get('settle_cycle', '') if c else '')
        _txt(ws, r, 11, c.get('quality_months', '') if c else '')
        _txt(ws, r, 12, c.get('pay_terms', '') if c else '')
        r += 1
    _txt(ws, r + 1, 1, '注：合同要素（内容/周期/约定金额/预算成本/付款条件）需合同台账补充——'
                       '暂无台账，账务侧已识别项目 %d 个；接入台账后（manifest contracts 配置）自动填列；'
                       '「结算周期/质保期」为合同资产账龄判断的关键（月结=1月/季结=3月/竣工结算=竣工时点；'
                       '质保金按质保期未到期）。' % len(rows),
         italic=True)
    for i, w in enumerate([16, 26, 30, 24, 12, 12, 18, 16, 16, 12, 12, 30], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'
    return ws


def _sheet2_confirm(ws, rows, cfg):
    """②按项目收入成本确认表（时段法：前期/本期 收入/成本/毛利率 + 完工进度）。"""
    ws.cell(1, 1, '按项目收入成本确认表（时段法=完工百分比法；单位：元）').font = SHELL_TITLE_FONT
    r = _hdr(ws, 3, ['核算主体', '项目/合同键', '前期确认收入', '前期确认成本', '前期毛利率',
                     '本期确认收入', '本期确认成本', '本期毛利率',
                     '累计已确认收入', '累计结算额(合同约定近似)', '完工进度(累计/预算)'])
    tot = defaultdict(float)
    for row in rows:
        rev_prev, cost_prev = _f(row['rev_prev']), _f(row['cost_prev'])
        rev_cur, cost_cur = _f(row['rev_cur']), _f(row['cost_cur'])
        gm_prev = (rev_prev - cost_prev) / rev_prev if abs(rev_prev) > 1e-9 else None
        gm_cur = (rev_cur - cost_cur) / rev_cur if abs(rev_cur) > 1e-9 else None
        rev_cum = rev_prev + rev_cur
        cost_cum = cost_prev + cost_cur
        _txt(ws, r, 1, row['ent'])
        _txt(ws, r, 2, row['unit'])
        _money(ws, r, 3, rev_prev)
        _money(ws, r, 4, cost_prev)
        _money(ws, r, 5, gm_prev if gm_prev is not None else '')
        _money(ws, r, 6, rev_cur)
        _money(ws, r, 7, cost_cur)
        _money(ws, r, 8, gm_cur if gm_cur is not None else '')
        _money(ws, r, 9, rev_cum)
        _money(ws, r, 10, row['settle'])
        _txt(ws, r, 11, '')
        for c in ('rev_cur', 'cost_cur'):
            tot[c] += locals()[c]
        for c in range(1, 12):
            ws.cell(r, c).border = SHELL_BORDER
        r += 1
    _txt(ws, r, 1, '合计', bold=True, align=SHELL_CEN)
    _money(ws, r, 6, tot['rev_cur'])
    _money(ws, r, 7, tot['cost_cur'])
    for c in range(1, 12):
        ws.cell(r, c).font = SHELL_BOLD
        ws.cell(r, c).fill = SHELL_TOT_FILL
    _txt(ws, r + 2, 1, '注：时段法=完工百分比法（投入法：完工进度=累计实际成本÷合同预计总成本）。'
                       '前期数=以前年度累计确认（无台账时=0，需上年度数据/台账补充）；'
                       '本期收入=GL 贷方（主营业务收入/合同结算\\收入结转）；本期成本=GL 借方（合同履约成本）；'
                       '毛利率=(收入-成本)/收入；累计结算额=合同结算\\价款结算 贷方余额（合同约定收入近似，含税口径待台账确认）。',
         italic=True)
    for i, w in enumerate([16, 24, 16, 16, 12, 16, 16, 12, 16, 22, 16], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'
    return ws


def _sheet3_tb(ws, rows, cfg, tb_rev, tb_cost):
    """③与账面核对（Σ项目 vs 主营业务收入/合同履约成本 TB）。"""
    ws.cell(1, 1, '与账面核对（Σ项目收入/成本 vs 科目余额表；单位：元）').font = SHELL_TITLE_FONT
    r = _hdr(ws, 3, ['口径', '项目Σ（账务侧）', '科目余额表(TB)', '差异', '说明'])
    sum_rev = sum(_f(x['rev_cur']) + _f(x['rev_prev']) for x in rows)
    sum_cost = sum(_f(x['cost_cur']) + _f(x['cost_prev']) for x in rows)
    _txt(ws, r, 1, '本期确认收入(累计)')
    _money(ws, r, 2, sum_rev)
    _money(ws, r, 3, tb_rev)
    _money(ws, r, 4, sum_rev - tb_rev)
    _txt(ws, r, 5, 'TB=主营业务收入 贷方发生额；项目Σ=GL 贷方同源（差异应为 0 或取数口径差）')
    r += 1
    _txt(ws, r, 1, '本期确认成本')
    _money(ws, r, 2, sum_cost)
    _money(ws, r, 3, tb_cost)
    _money(ws, r, 4, sum_cost - tb_cost)
    _txt(ws, r, 5, 'TB=合同履约成本 借方发生额；项目Σ=GL 借方同源')
    r += 1
    _txt(ws, r + 1, 1, '注：项目Σ与 TB 同为 GL 取数（同源应 0 差）；若差异大=取数不完整（铁律：明细与 TB 不一致=取数 bug）。',
         italic=True)
    for i, w in enumerate([24, 18, 18, 18, 60], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


def _sheet4_cash(ws, rows, cfg):
    """④收款与应收款核对（前期应收/期末应收 vs 应收辅助核算；收款=应收贷方）。"""
    ws.cell(1, 1, '收款与应收款核对（单位：元）').font = SHELL_TITLE_FONT
    r = _hdr(ws, 3, ['核算主体', '项目/合同键', '期初应收', '期末应收', '本期收款(≈期初+确认-期末)', '期末应收(辅助核算)'])
    tot = defaultdict(float)
    for row in rows:
        recv_qc, recv_qm = _f(row['recv_qc']), _f(row['recv_qm'])
        rev_cur = _f(row['rev_cur'])
        recv_pay = recv_qc + rev_cur - recv_qm  # 本期收款≈期初应收+本期确认-期末应收
        _txt(ws, r, 1, row['ent'])
        _txt(ws, r, 2, row['unit'])
        _money(ws, r, 3, recv_qc)
        _money(ws, r, 4, recv_qm)
        _money(ws, r, 5, recv_pay)
        _money(ws, r, 6, recv_qm)
        tot['qc'] += recv_qc; tot['qm'] += recv_qm; tot['pay'] += recv_pay
        r += 1
    _txt(ws, r, 1, '合计', bold=True, align=SHELL_CEN)
    _money(ws, r, 3, tot['qc'])
    _money(ws, r, 4, tot['qm'])
    _money(ws, r, 5, tot['pay'])
    _money(ws, r, 6, tot['qm'])
    for c in range(1, 7):
        ws.cell(r, c).font = SHELL_BOLD
        ws.cell(r, c).fill = SHELL_TOT_FILL
    _txt(ws, r + 2, 1, '注：期初/期末应收=应收账款辅助核算（往来单位=业主）；本期收款为倒推值'
                       '（期初应收+本期确认-期末应收），未含 预收/质保金 等，精确收款需 银行流水+预收辅助 补充。',
         italic=True)
    for i, w in enumerate([16, 26, 16, 16, 20, 18], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'
    return ws


def _sheet5_margin(ws, rows, cfg):
    """⑤毛利异常分析（毛利率异常项目/亏损合同预警）。"""
    ws.cell(1, 1, '毛利异常分析（毛利率<0 或异常波动预警；单位：元）').font = SHELL_TITLE_FONT
    r = _hdr(ws, 3, ['核算主体', '项目/合同键', '本期收入', '本期成本', '本期毛利率', '预警'])
    n = 0
    for row in rows:
        rev_cur, cost_cur = _f(row['rev_cur']), _f(row['cost_cur'])
        gm = (rev_cur - cost_cur) / rev_cur if abs(rev_cur) > 1e-9 else None
        warn = ''
        if gm is None:
            warn = '无收入'
        elif gm < 0:
            warn = '亏损项目'
        elif gm > 0.5:
            warn = '毛利率异常高(>50%)'
        if not warn:
            continue
        _txt(ws, r, 1, row['ent'])
        _txt(ws, r, 2, row['unit'])
        _money(ws, r, 3, rev_cur)
        _money(ws, r, 4, cost_cur)
        _money(ws, r, 5, gm if gm is not None else '')
        jc = _txt(ws, r, 6, warn, align=SHELL_CEN)
        jc.font = Font(name='Times New Roman', size=10, bold=True,
                       color='C00000' if '亏损' in warn else 'BF8F00')
        r += 1
        n += 1
    if not n:
        _txt(ws, r, 1, '（无异常项目）')
    _txt(ws, r + 1, 1, '注：施工企业成本=合同履约成本（无单位→主体级分摊）；亏损合同需判断减值（准则15号 第二十七条）。',
         italic=True)
    for i, w in enumerate([16, 26, 16, 16, 12, 20], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


def _sheet6_cost(ws, rows, cfg):
    """⑥项目-成本构成（与存货按指令号衔接；成本类别占比）。"""
    ws.cell(1, 1, '项目-成本构成（合同履约成本 借方 按类别；单位：元——与存货按指令号同源衔接）').font = SHELL_TITLE_FONT
    r = _hdr(ws, 3, ['核算主体', '项目/合同键', '直接材料', '直接人工', '机械使用', '分包', '间接费用', '其他', '合计'])
    # 简化：成本类别需逐主体 GL 再聚合——此处按主体级成本拆分（无明细时合计列示）
    from collections import Counter
    cat_tot = Counter()
    for row in rows:
        cost_cur = _f(row['cost_cur'])
        if abs(cost_cur) < 0.005:
            continue
        _txt(ws, r, 1, row['ent'])
        _txt(ws, r, 2, row['unit'])
        _money(ws, r, 3, '')
        _money(ws, r, 4, '')
        _money(ws, r, 5, '')
        _money(ws, r, 6, '')
        _money(ws, r, 7, '')
        _money(ws, r, 8, '')
        _money(ws, r, 9, cost_cur)
        cat_tot['合计'] += cost_cur
        r += 1
    _txt(ws, r, 1, '合计', bold=True, align=SHELL_CEN)
    _money(ws, r, 9, cat_tot.get('合计', 0.0))
    for c in range(1, 10):
        ws.cell(r, c).font = SHELL_BOLD
        ws.cell(r, c).fill = SHELL_TOT_FILL
    _txt(ws, r + 2, 1, '注：成本类别拆分需逐主体 GL 按科目名聚合（直接材料/人工/机械/分包/间接费用）——'
                       '与存货按指令号（WBS/订单/项目编号）同源衔接；无台账降级版先列主体级合计，'
                       '类别明细在明细版补齐。', italic=True)
    for i, w in enumerate([16, 26, 14, 14, 14, 14, 14, 14, 16], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


def _sheet7_note(ws, rows, cfg):
    """⑦口径说明。"""
    ws.cell(1, 1, '口径说明（项目制核算收入底稿）').font = SHELL_TITLE_FONT
    notes = [
        '1. 项目制核算：施工/工程类企业按【项目/合同】归集收入与成本；XBJ 每个核算主体=一个项目部=一个项目，'
        '辅助核算「往来单位名称」=业主（客户）。',
        '2. 收入确认（时段法/完工百分比法，准则15号）：结果可可靠估计时，本期收入=合同总收入×完工进度−前期累计已确认；'
        '完工进度=累计实际成本÷合同预计总成本（投入法）。',
        '3. 本期收入取数：GL 贷方（主营业务收入/其他业务收入/合同结算\\收入结转）；'
        '合同结算\\价款结算 贷方余额=累计已结算额（合同约定收入近似，含税口径待台账确认）。',
        '4. 本期成本取数：GL 借方（合同履约成本/主营业务成本）；XBJ 成本无辅助单位→主体级（项目级）。',
        '5. 收款/应收：期初/期末应收=应收账款辅助核算（往来单位=业主）；本期收款=倒推值（期初应收+本期确认−期末应收），'
        '精确收款需 银行流水+预收辅助 补充。',
        '6. 合同台账要素（合同内容/周期/约定金额/预算成本）暂缺——留空待补；接入台账后（manifest contracts 或读台账 xlsx）'
        '按项目键自动匹配填列，并可计算完工进度=累计成本/预算成本。',
        '7. 项目维度键：XBJ=辅助核算往来单位（业主）；AH(SAP)=WBS(col17)/合同号(col15)/订单(col24)，'
        '与存货按指令号同源——项目-成本构成 sheet 与存货底稿衔接。',
        '8. 与账面核对：项目Σ（GL 同源）与 TB 应 0 差；差异大=取数不完整（铁律：明细与 TB 不一致=取数 bug）。',
        '9. 毛利异常/亏损合同：毛利率<0 或异常高需查因；亏损合同按准则15号 第二十七条 计提预计损失。',
        '10. 本版=降级模式（无合同台账）；后续完善：台账接入+完工进度计算+两年毛利率对比+与存货指令号深度衔接。',
    ]
    r = 3
    for n in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(r, 1, n)
        c.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical='top')
        c.border = SHELL_BORDER
        ws.row_dimensions[r].height = 32
        r += 1
    for i, w in enumerate([18, 18, 18], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


def _sheet8_adjust(ws, rows, cfg):
    """⑧审计调整（占位，待审计判断后填）。"""
    ws.cell(1, 1, '审计调整分录（待审计判断后填列）').font = SHELL_TITLE_FONT
    _hdr(ws, 3, ['核算主体', '项目/合同键', '调整类型', '借：科目', '贷：科目', '金额', '说明'])
    _txt(ws, 5, 1, '（暂无调整）', italic=True)
    for i, w in enumerate([16, 26, 14, 24, 24, 16, 30], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


def _write_workbook(out_path, cfg, rows, tb_rev, tb_cost):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _sheet1_contracts(wb.create_sheet('项目合同台账'), rows, cfg)
    _sheet2_confirm(wb.create_sheet('按项目收入成本确认'), rows, cfg)
    _sheet3_tb(wb.create_sheet('与账面核对'), rows, cfg, tb_rev, tb_cost)
    _sheet4_cash(wb.create_sheet('收款与应收款核对'), rows, cfg)
    _sheet5_margin(wb.create_sheet('毛利异常分析'), rows, cfg)
    _sheet6_cost(wb.create_sheet('项目-成本构成'), rows, cfg)
    _sheet7_note(wb.create_sheet('口径说明'), rows, cfg)
    _sheet8_adjust(wb.create_sheet('审计调整'), rows, cfg)
    try:
        finalize_workbook(wb)
    except Exception:
        pass
    wb.save(out_path)
    wb.close()


def run_one(acct, sub, out_dir=None):
    acct_cfg, sub_cfg = PM.job_for(acct, sub)
    cfg = dict(sub_cfg)
    year = cfg['year']
    data_dir = os.path.join(PM.DATA_ROOT, cfg['data_dir'].replace('/', os.sep))
    print(f'=== {acct}/{sub} 项目制核算收入底稿 ===')
    print(f'  数据目录: {data_dir}')

    rows = _project_rows(acct_cfg, sub_cfg, data_dir)
    print(f'  识别项目 {len(rows)} 个')
    if not rows:
        print('  ⚠️ 无项目数据')
        return 1

    # TB 控制数（收入/成本 借方贷方；复用 _read_tb）
    tb_rev = 0.0
    tb_cost = 0.0
    import glob
    for f in glob.glob(os.path.join(data_dir, '*科目余额表*.xlsx')):
        ent_hint = os.path.basename(f)
        ent_hint = re.sub(r'(外币)?(科目余额表|科目余额|余额表).*$', '', ent_hint)
        tb = _read_tb(data_dir, ent_hint)
        for code, v in tb.items():
            nm = v['name']
            if any(k in nm for k in REV_KW):
                tb_rev += _f(v['credit'])   # 收入取贷方（铁律3）
            if any(k in nm for k in COST_KW):
                tb_cost += _f(v['debit'])   # 成本取借方
    print(f'  TB: 收入贷发 {tb_rev:,.2f} / 成本借发 {tb_cost:,.2f}')

    out = os.path.join(os.path.dirname(PM.resolve(acct)), acct, '底稿', str(year))
    if sub != '_root':
        out = os.path.join(out, sub)
    os.makedirs(out, exist_ok=True)
    out_path = os.path.join(out, f'项目制收入底稿_{year}.xlsx')
    _write_workbook(out_path, cfg, rows, tb_rev, tb_cost)
    print(f'  ✅ 项目制收入底稿已生成: {out_path}')
    return 0


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('acct', nargs='?', default=None)
    ap.add_argument('sub', nargs='?', default=None)
    a = ap.parse_args(argv)
    rc = 0
    for acct, acct_cfg in PM.jobs().items():
        if a.acct and acct != a.acct:
            continue
        for sub in acct_cfg['subs']:
            if a.sub and sub != a.sub:
                continue
            rc |= run_one(acct, sub)
    return rc


if __name__ == '__main__':
    sys.exit(main())
