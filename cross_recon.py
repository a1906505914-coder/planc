# -*- coding: utf-8 -*-
"""科目间勾稽核对（通用跨科目勾稽引擎，2026-07-29 新增）。

背景：此前 13 个 builder 的勾稽各自内联实现、部分硬编码配对（expense 只查 制造费用5101贷方↔
生产成本5001.04借方、inventory 依赖硬编码外部 xlsx 且缺失即跳过），存在"找的不完整"漏配风险。
本模块提供【单一来源】的通用跨科目勾稽：自动枚举常见的科目间勾稽链，并以
『控制科目 GL 净发生 ↔ TB 净发生』的干净口径逐条核对，供审计人员整体复核勾稽链是否闭合。

勾稽配对（均为"控制科目 GL 净发生 vs TB 净发生"，避免按业务族粗分导致的伪差异）：
  1. 累计折旧/累计摊销/长期待摊费用摊销：GL 净(贷−借) ↔ TB 净(贷发−借发)；
  2. 坏账准备：GL 净(贷−借) ↔ TB 净(贷发−借发)；
  3. 存货跌价准备：GL 净(贷−借) ↔ TB 净(贷发−借发)；
  4. 税金及附加：GL 借发 ↔ TB 借发（费用侧）；应交税费(附加税费) GL 贷发 ↔ TB 贷发（参考）；
  5. 递延收益：GL 净(贷−借) ↔ TB 净(贷发−借发)；
  6. 应付职工薪酬：GL 净(贷−借) ↔ TB 净(贷发−借发)；
  7. 收入（主营/其他业务收入）：GL 贷发 ↔ TB 贷发；
  8. 存货采购 / 收入确认：方向性参考（借 存货 / 贷 收入 对应 应付/应收/银行），无 TB 直接对照。

数据口径：
  · TB（科目余额表）= 权威控制数（期初/借发/贷发/期末）。
  · GL（综合查询明细表）= 方向性归集。
  · 2025 GL 全量(1–12月)；2026 GL 仅 1–5 月(YTD)，差额=凭证抽取不完整，非账务错误，须披露。
  · GL 数据源仅含"科目名称"列、无科目代码，故控制科目按名称关键字匹配（铁律5 延伸）。

输出：科目间勾稽核对_生成.xlsx
  · 科目间勾稽汇总（逐配对：主体/年度/勾稽关系/控制科目/TB净发生/GL净发生/差异/状态/说明）
  · 说明
用法：python cross_recon.py <文件夹> [输出xlsx]   （默认输出到 文件夹/科目间勾稽核对_生成.xlsx）
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from audit_common import (discover_entities, read_gl_rows, read_tb_full,
                          norm_subject_name, _safe_save,
                          _recon_status, finalize_workbook)

# 视觉（与全程序统一，re-export 自 audit_shell）
from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT, SHELL_SUB_FONT,
                         SHELL_BOLD, SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT)
from openpyxl.utils import get_column_letter

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOT_FILL = PatternFill('solid', fgColor='E2EFDA')
WARN_FILL = PatternFill('solid', fgColor='FCE4D6')
REF_FILL = PatternFill('solid', fgColor='DDEBF7')
MONEY = '#,##0.00'


# ---------------------------------------------------------------------------
# 数据准备
# ---------------------------------------------------------------------------
def _gl_by_subject(rows, e, y):
    """返回 {归一化科目名: {'dr':借发合计, 'cr':贷发合计}}，仅该主体该年度。"""
    agg = defaultdict(lambda: {'dr': 0.0, 'cr': 0.0})
    for r in rows:
        if r.get('e') != e or r.get('y') != y:
            continue
        nm = norm_subject_name(r.get('name'))
        if not nm:
            continue
        agg[nm]['dr'] += float(r.get('dr', 0.0) or 0.0)
        agg[nm]['cr'] += float(r.get('cr', 0.0) or 0.0)
    return agg


def _tb_leaf(tb, e, y):
    """返回 {归一化科目名: {'jf':借发, 'df':贷发}}，仅该主体该年度末级。"""
    agg = defaultdict(lambda: {'jf': 0.0, 'df': 0.0})
    for (te, code, name, ty), v in tb.items():
        if te != e or ty != y:
            continue
        nm = norm_subject_name(name)
        if not nm:
            continue
        agg[nm]['jf'] += float(v.get('jf', 0.0) or 0.0)
        agg[nm]['df'] += float(v.get('df', 0.0) or 0.0)
    return agg


def _match_keys(pool, keywords):
    """从 pool(归一化名->数值) 中按关键字匹配所有命中的科目名（返回列表）。"""
    out = []
    for nm in pool:
        if any(k in nm for k in keywords):
            out.append(nm)
    return out


def _mk(e, y, relation, ctrl, tb_val, gl_val, status_override=None, note='', ref=False):
    diff = (gl_val - tb_val) if (tb_val is not None and gl_val is not None) else None
    if status_override:
        status = status_override
    elif diff is None:
        status = '参考'
    else:
        status = _recon_status(diff, gl_val, tb_val)
    return {'e': e, 'y': y, 'relation': relation, 'ctrl': ctrl,
            'tb_val': tb_val, 'gl_val': gl_val, 'diff': diff, 'status': status,
            'note': note, 'ref': ref}


# ---------------------------------------------------------------------------
# 各勾稽配对（返回 list of record dict）
# ---------------------------------------------------------------------------
def recon_depreciation(e, y, tb_leaf, gl_subj):
    recs = []
    keys = ['累计折旧', '累计摊销', '长期待摊费用']
    for nm in _match_keys(gl_subj, keys):
        gl_net = gl_subj[nm]['cr'] - gl_subj[nm]['dr']   # 贷−借 净增加
        tb_net = tb_leaf.get(nm, {}).get('df', 0.0) - tb_leaf.get(nm, {}).get('jf', 0.0)
        recs.append(_mk(e, y, '折旧/摊销计提', nm, tb_net, gl_net,
                         note='GL 净(贷−借) ↔ TB 净(贷发−借发)；差额非账务错误(2026 YTD/取数缺口)须披露'))
    return recs


def recon_baddebt(e, y, tb_leaf, gl_subj):
    recs = []
    for nm in _match_keys(gl_subj, ['坏账准备']):
        gl_net = gl_subj[nm]['cr'] - gl_subj[nm]['dr']
        tb_net = tb_leaf.get(nm, {}).get('df', 0.0) - tb_leaf.get(nm, {}).get('jf', 0.0)
        recs.append(_mk(e, y, '坏账准备', nm, tb_net, gl_net,
                         note='GL 净(贷−借) ↔ TB 净(贷发−借发)；含计提与核销'))
    return recs


def recon_inv_impair(e, y, tb_leaf, gl_subj):
    recs = []
    for nm in _match_keys(gl_subj, ['存货跌价准备']):
        gl_net = gl_subj[nm]['cr'] - gl_subj[nm]['dr']
        tb_net = tb_leaf.get(nm, {}).get('df', 0.0) - tb_leaf.get(nm, {}).get('jf', 0.0)
        recs.append(_mk(e, y, '存货跌价准备', nm, tb_net, gl_net,
                         note='GL 净(贷−借) ↔ TB 净(贷发−借发)；含计提与转回'))
    return recs


def recon_tax_surcharge(e, y, tb_leaf, gl_subj):
    recs = []
    for nm in _match_keys(gl_subj, ['税金及附加']):
        gl_dr = gl_subj[nm]['dr']
        tb_jf = tb_leaf.get(nm, {}).get('jf', 0.0)
        recs.append(_mk(e, y, '税金及附加(费用侧)', nm, tb_jf, gl_dr,
                         note='GL 借发 ↔ TB 借发'))
    # 应交税费(附加税费族) 仅作参考：GL 贷发 vs TB 贷发，但应交税费含增值税等，无法单独剥离附加
    for nm in _match_keys(gl_subj, ['应交税费']):
        gl_cr = gl_subj[nm]['cr']
        tb_df = tb_leaf.get(nm, {}).get('df', 0.0)
        recs.append(_mk(e, y, '应交税费(参考)', nm, tb_df, gl_cr, ref=True,
                         note='参考：应交税费含增值税等，附加税费无法单独剥离，差异含增值税'))
    return recs


def recon_deferred_income(e, y, tb_leaf, gl_subj):
    recs = []
    for nm in _match_keys(gl_subj, ['递延收益']):
        gl_net = gl_subj[nm]['cr'] - gl_subj[nm]['dr']
        tb_net = tb_leaf.get(nm, {}).get('df', 0.0) - tb_leaf.get(nm, {}).get('jf', 0.0)
        recs.append(_mk(e, y, '递延收益', nm, tb_net, gl_net,
                         note='GL 净(贷−借) ↔ TB 净(贷发−借发)；政府补助等递延'))
    return recs


def recon_payroll(e, y, tb_leaf, gl_subj):
    recs = []
    for nm in _match_keys(gl_subj, ['应付职工薪酬']):
        gl_net = gl_subj[nm]['cr'] - gl_subj[nm]['dr']
        tb_net = tb_leaf.get(nm, {}).get('df', 0.0) - tb_leaf.get(nm, {}).get('jf', 0.0)
        recs.append(_mk(e, y, '应付职工薪酬', nm, tb_net, gl_net,
                         note='GL 净(贷−借) ↔ TB 净(贷发−借发)；含计提与支付'))
    return recs


def recon_revenue(e, y, tb_leaf, gl_subj):
    recs = []
    for nm in _match_keys(gl_subj, ['主营业务收入', '其他业务收入', '营业收入']):
        gl_cr = gl_subj[nm]['cr']
        tb_df = tb_leaf.get(nm, {}).get('df', 0.0)
        recs.append(_mk(e, y, '收入确认', nm, tb_df, gl_cr,
                         note='GL 贷发 ↔ TB 贷发'))
    return recs


def recon_inventory_purchase(e, y, tb_leaf, gl_subj):
    """方向性参考：借 存货 总额 对应 贷 应付/银行(含进项税)。无 TB 直接对照，仅列示参考。"""
    recs = []
    inv_keys = ['原材料', '库存商品', '周转材料', '在途物资', '委托加工物资', '材料采购',
                '商品采购', '低值易耗品', '包装物', '发出商品', '半成品', '生产成本']
    for nm in _match_keys(gl_subj, inv_keys):
        if '生产成本' in nm:
            continue
        gl_dr = gl_subj[nm]['dr']
        if abs(gl_dr) < 0.005:
            continue
        recs.append(_mk(e, y, '存货采购(参考)', nm, None, gl_dr, ref=True,
                         note='参考：借 存货 对应 贷 应付/应付票据/银行存款(含进项税)，方向性提示'))
    return recs


RECON_FUNCS = [
    recon_depreciation,
    recon_baddebt,
    recon_inv_impair,
    recon_tax_surcharge,
    recon_deferred_income,
    recon_payroll,
    recon_revenue,
    recon_inventory_purchase,
]


# ---------------------------------------------------------------------------
# 构建工作簿
# ---------------------------------------------------------------------------
def build(data_dir, out_path):
    ents = discover_entities(data_dir)
    if not ents:
        print('❌ 未发现任何账套主体'); return None
    tb = read_tb_full(data_dir, ents)
    gl_rows = read_gl_rows(data_dir, ents)
    years = sorted({y for e, yd in ents.items() for y in yd})
    print(f'· 主体 {len(ents)} 个，年份 {years}，TB 行 {len(tb)}，GL 行 {len(gl_rows)}')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet('科目间勾稽汇总')
    ws.cell(1, 1, '科目间勾稽核对（通用跨科目勾稽，控制科目 GL 净发生 ↔ TB 净发生）').font = SHELL_TITLE_FONT
    ws.cell(2, 1, '配对按"控制科目名称"精确匹配（跨账套代码不一致仍按名称，铁律5延伸）；状态=一致/不一致(容差内)/参考。'
                  '「参考」类为方向性提示、无 TB 直接对照数。2026 GL 仅 1-5 月(YTD)，差额非账务错误，须披露。').font = \
        Font(name='Times New Roman', italic=True, size=10, color='808080')
    hdr = ['主体', '年度', '勾稽关系', '控制科目', 'TB 净发生', 'GL 净发生', '差异', '状态', '说明']
    hr = 4
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(hr, c, h); cell.font = SHELL_HFONT; cell.fill = SHELL_HFILL
        cell.border = BORDER; cell.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')
    r = hr + 1
    all_recs = []
    for e in sorted(ents):
        for y in sorted(ents[e]):
            tb_leaf_m = _tb_leaf(tb, e, y)
            gl_subj_m = _gl_by_subject(gl_rows, e, y)
            for fn in RECON_FUNCS:
                recs = fn(e, y, tb_leaf_m, gl_subj_m)
                all_recs.extend(recs)
                for rec in recs:
                    vals = [rec['e'], rec['y'], rec['relation'], rec['ctrl'],
                            rec['tb_val'], rec['gl_val'], rec['diff'], rec['status'], rec['note']]
                    for c, v in enumerate(vals, 1):
                        cell = ws.cell(r, c, ('' if v is None else v))
                        cell.border = BORDER
                        if c in (5, 6, 7) and isinstance(v, (int, float)):
                            cell.number_format = MONEY
                        if c == 9:
                            cell.alignment = Alignment(wrap_text=True, vertical='center')
                        if c == 8:
                            cell.alignment = SHELL_CEN
                            if v == '不一致':
                                cell.fill = WARN_FILL
                            elif v == '一致':
                                cell.fill = TOT_FILL
                            elif v == '参考':
                                cell.fill = REF_FILL
                    r += 1
    if not all_recs:
        ws.cell(r, 1, '未发现任何可勾稽的跨科目配对（或无 GL/TB 数据）。').font = Font(name='Times New Roman', italic=True)
    for i, w in enumerate([14, 7, 18, 24, 16, 16, 16, 10, 52], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)

    # 说明 sheet
    ws2 = wb.create_sheet('说明')
    n_consistent = sum(1 for x in all_recs if x['status'] == '一致')
    n_inconsistent = sum(1 for x in all_recs if x['status'] == '不一致')
    n_ref = sum(1 for x in all_recs if x['status'] == '参考')
    notes = [
        '科目间勾稽核对 — 编制与口径说明',
        '',
        '一、目的（补全各科目底稿勾稽"找的不完整"的盲区）',
        '  本表以统一引擎自动枚举并核对科目间常见的勾稽链，弥补各 builder 内联勾稽的漏配：',
        '  · 折旧/摊销计提：    累计折旧/累计摊销/长期待摊费用摊销 净发生；',
        '  · 坏账准备 / 存货跌价准备：计提与核销后的净发生；',
        '  · 税金及附加：      费用侧 借发 ↔ TB 借发（应交税费仅作参考）；',
        '  · 递延收益 / 应付职工薪酬：净发生；',
        '  · 收入确认：        主营/其他业务收入 贷发 ↔ TB 贷发；',
        '  · 存货采购 / 收入确认：方向性参考（无 TB 直接对照）。',
        '',
        '二、方法（单一来源，杜绝重复实现）',
        '  · 每个"控制科目"按名称关键字匹配 TB 末级与 GL 全量，取 GL 净(贷−借) 与 TB 净(贷发−借发) 对照；',
        '  · 状态：差异在容差内(≤max(1,较大基准*1e-6))为"一致"；否则"不一致"；无 TB 直接对照的为"参考"。',
        '  · 科目归类以【科目名称】判定（跨账套代码体系不一致时仍按名称，铁律5延伸）。',
        '',
        '三、口径提示',
        '  · 应交税费/收入等含多子目，匹配按名称关键字；若某账套末级命名差异大可能漏配，须人工补充。',
        '  · 2026 GL 仅含 1-5 月(YTD)，凭证抽取不完整 → 勾稽差额可能非账务错误，须作为取数范围限制披露。',
        f'  · 本次共生成勾稽记录 {len(all_recs)} 条（一致 {n_consistent} / 不一致 {n_inconsistent} / 参考 {n_ref}）。',
    ]
    for i, t in enumerate(notes, 1):
        ws2.cell(i, 1, t)
        if i == 1:
            ws2.cell(i, 1).font = Font(name='Times New Roman', bold=True, size=10)
    ws2.column_dimensions['A'].width = 120

    finalize_workbook(wb)
    path, warn = _safe_save(wb, out_path)
    if warn:
        print('  ⚠️', warn)
    print(f'  ✓ 已保存：{path}  (一致 {n_consistent} / 不一致 {n_inconsistent} / 参考 {n_ref})')
    return path


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else None
    out = sys.argv[2] if len(sys.argv) > 2 else None
    if not data_dir:
        print('用法：python cross_recon.py <文件夹> [输出xlsx]')
        return
    if out is None:
        out = os.path.join(data_dir, '科目间勾稽核对_生成.xlsx')
    build(data_dir, out)


if __name__ == '__main__':
    main()
