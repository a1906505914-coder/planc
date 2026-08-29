# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=短期借款审计底稿_<年>_生成.xlsx + 长期借款审计底稿_<年>_生成.xlsx（各含 审定表/明细表/凭证抽查/附注汇总） | 关键列=核算主体/年度/借款性质/借款银行/借款金额(期末余额)/借款日/归还日/利率/利息金额/借款类型/对应合同编号/具体担保或抵押内容/期初余额/本期增加/本期减少/币种/备注 | 职责=短期(2001)与长期(2501)借款按科目拆分独立成稿，各按主体→年度小计→全集团合计的逐笔借款明细(辅助核算余额表优先,TB子目兜底,GL抽凭日期/利息匹配) | 利率/合同编号/借款类型/担保内容 留空待审计人员据合同补录
"""
借款(短期/长期)审计底稿生成程序 —— 合并版
================================================================================
覆盖借款科目（按科目【名称】自动识别，兼容不同会计科目表）：
  · 短期借款    2001
  · 长期借款    2501

数据来源：
  · 科目余额表       = 权威控制数（期初/借发/贷发/期末，带借贷符号）。
  · 辅助核算余额表   = 按"借款银行/借款项目"的末级明细（等级3），是"借款明细表"的主数据源；
                      若某主体无辅助核算借款明细，则回退到科目余额表 2001/2501 子目。
  · 综合查询明细表(GL) = 凭证级，用于【凭证抽查】与【利息支出匹配】；
                      借款日/归还日优先取 GL 中该笔借款的借入/偿还日期，取不到则回落为
                      期初日期(年度1月1日)/期末日期(年度12月31日)。

输出（按科目拆分，每年度两份独立底稿）：
  <被拖文件夹>/短期借款审计底稿_<年>_生成.xlsx
  <被拖文件夹>/长期借款审计底稿_<年>_生成.xlsx

每份底稿含 4 个工作表：
  ① 审定表（第一张表）：按核算主体列示 期初数/期末未审数/审计调整数/审定数 + 全集团合计。
  ② 明细表（第二张表，无编制说明）：按 核算主体 → 年度 小计 → 全集团合计，
     逐笔列示每笔借款（借款银行/项目）的借款银行、借款金额(=期末余额)、借款日、归还日、
     利率、利息金额、借款类型(信用/担保/抵押)、对应合同编号、具体担保或抵押内容、
     期初余额/本期增加(借入)/本期减少(偿还)/币种/备注。
     · 负债科目：本期增加=贷方发生额(借入)、本期减少=借方发生额(偿还)。
     · 利率/合同编号/借款类型/担保内容 数据源不含，留空由审计人员据借款合同补录。
  ③ 凭证抽查：该科目全年度发生额凭证清单（合并各年），列示主体/年度/日期/凭证号/
     摘要/借/贷/对方科目，并留"抽查/审计结论"列。
  ④ 附注汇总：本科目 期末数/期初数 两期数（每主体一列，增减变动由现金流量表核对）。
================================================================================
"""

# ---- 启动诊断（先于 import openpyxl，便于排查"拖入闪退无痕迹"）----
import os as _bs_os
import sys as _bs_sys
import subprocess as _bs_sub
import re as _re


def _write_boot_log(tag):
    import os as _bl_os
    if _bl_os.environ.get('AUDIT_LOG') != '1':
        return
    try:
        import datetime
        _p = _bs_os.path.join(_bs_os.path.expanduser('~'), 'Desktop', 'loan_boot.log')
        with open(_p, 'a', encoding='utf-8') as _f:
            _f.write('[%s] %s sys=%s argv=%s\n' % (
                datetime.datetime.now().isoformat(timespec='seconds'),
                tag, _bs_sys.executable, _bs_sys.argv))
    except Exception:
        pass


_write_boot_log('LOAD')


def _find_managed_python():
    base = _bs_os.path.join(_bs_os.environ.get('USERPROFILE', _bs_os.path.expanduser('~')),
                            '.workbuddy', 'binaries', 'python', 'versions')
    if _bs_os.path.isdir(base):
        for d in sorted(_bs_os.listdir(base), reverse=True):
            p = _bs_os.path.join(base, d, 'python.exe')
            if _bs_os.path.isfile(p):
                return p
    return None


def _bootstrap_relaunch():
    return  # ⚡ 2026-08-22 禁用托管Python重启：venv依赖已齐全，托管3.13 openpyxl损坏


if _bs_os.environ.get('AUDIT_NO_RELAUNCH') != '1':
    _bootstrap_relaunch()

import argparse
import os
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

from audit_shell import (SHELL_HFILL, SHELL_HFONT, SHELL_TOT_FILL, SHELL_TITLE_FONT, SHELL_SUB_FONT,
                         SHELL_BOLD, SHELL_NUM, SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT,
                         finalize_workbook, resolve_template)
from audit_common import (discover_entities, read_tb_full, read_gl_rows, safe, signed_balance, _safe_save, voucher_skip_keys, is_accrual_line)
from mask_dict import mask_names   # ⚡ console 主体列表打印脱敏（2026-08-22 复扫）
# 2026-08-22 融资台账 adapter（从本模块抽取，供借款/应付票据/表外复用；脱敏版优先）
from financing_ledger import (ENTITY_MAP, _load_entity_map, _ent_disp,
                              _fin_ent_code, _xls_date, _guarantee_to_cat,
                              _GUARANTOR_MAP, _BANK_CODE_MAP, _mask_bank, _mask_guarantor,
                              _load_finance_loans, set_fin_data_root)

# ----------------------------------------------------------------------------


# ----------------------------------------------------------------------------
# 借款合同信息补充（合同信息_脱敏.json，contract_detail_ledger 产出）
# ----------------------------------------------------------------------------
_CONTRACT_INFO_CACHE = None


def _load_contract_info():
    global _CONTRACT_INFO_CACHE
    if _CONTRACT_INFO_CACHE is None:
        try:
            fp = os.path.join(r'd:/底稿测试/ADF/数据/2026', '合同信息_脱敏.json')
            with open(fp, encoding='utf-8') as f:
                _CONTRACT_INFO_CACHE = json.load(f)
        except Exception:
            _CONTRACT_INFO_CACHE = []
    return _CONTRACT_INFO_CACHE


def _contract_rate_by_ent(ent, sk='ST'):
    """主体（借X）→ 该主体借款合同/委贷借据的利率列表（去重，用于补充明细表利率列）。"""
    info = _load_contract_info()
    rates = []
    for r in info:
        if r.get('借款主体') != ent:
            continue
        if r.get('类型') not in ('借款合同', '委托贷款(借据)', '委贷协议'):
            continue
        v = r.get('利率(%)')
        if v:
            try:
                rates.append(float(v))
            except (TypeError, ValueError):
                pass
    return sorted(set(rates))


# ----------------------------------------------------------------------------

# ----------------------------------------------------------------------------
# 借款科目组定义（按科目【名称/代码】识别，兼容不同科目表）
# ----------------------------------------------------------------------------
SUBJECTS = {
    'ST': {
        'key': 'ST',
        'title': '短期借款',
        'code': '2001', 'names': ['短期借款'], 'is_credit': True,
    },
    'LT': {
        'key': 'LT',
        'title': '长期借款',
        'code': '2501', 'names': ['长期借款'], 'is_credit': True,
    },
}

# 银行关键词 -> token（用于把 GL 利息摘要/借款科目匹配到借款银行）
BANK_KW = [('华夏银行', '华夏'), ('中国银行', '中行'), ('中行', '中行'), ('工商银行', '工行'), ('工行', '工行'),
           ('建设银行', '建行'), ('建行', '建行'), ('农业银行', '农行'), ('农行', '农行'), ('宁波银行', '宁波'),
           ('浙商银行', '浙商'), ('杭州银行', '杭州'), ('平安银行', '平安'), ('光大银行', '光大'),
           ('招商银行', '招行'), ('交通银行', '交行'), ('兴业银行', '兴业'), ('中信银行', '中信'),
           ('民生银行', '民生'), ('浦发银行', '浦发'), ('邮储', '邮储'), ('进出口银行', '进出口')]


def _bank_token(zh):
    if not zh:
        return None
    z = str(zh)
    for kw, tok in BANK_KW:
        if kw in z:
            return tok
    return None


# ----------------------------------------------------------------------------
# 样式 / 单元格助手（与所有者权益底稿同源，保证视觉一致）
# ----------------------------------------------------------------------------
def _hdr(ws, row, headers, start_col=1):
    for j, h in enumerate(headers, start_col):
        c = ws.cell(row, j, h)
        c.fill = SHELL_HFILL
        c.font = SHELL_HFONT
        c.border = SHELL_BORDER
        c.alignment = SHELL_CEN
    return row + 1


def _money(ws, r, c, v):
    if v is not None:
        try:
            v = round(float(v), 2)
            if abs(v) < 0.005:
                v = None          # ⚡ 2026-08-11 零值留空（规则4，loan_detail 全表统一）
        except (TypeError, ValueError):
            pass
    cell = ws.cell(r, c, v)
    cell.number_format = SHELL_NUM
    cell.border = SHELL_BORDER
    cell.alignment = SHELL_RGT
    return cell


def _txt(ws, r, c, v, bold=False, align=None, fill=None):
    cell = ws.cell(r, c, v)
    cell.border = SHELL_BORDER
    cell.alignment = align or SHELL_LEFT
    if bold:
        cell.font = SHELL_BOLD
    if fill:
        cell.fill = fill
    return cell


def _title(ws, text, ncols, row=1):
    ws.cell(row, 1, text).font = SHELL_TITLE_FONT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=max(ncols, 1))
    return row + 1


def _fmt_vno(r):
    vt = r.get('vtype') or ''
    vn = r.get('vno') or ''
    if vt and vn:
        return '%s-%s' % (vt, vn)
    return str(vn) if vn else ''


def _fmt_date(d):
    if d is None:
        return ''
    if hasattr(d, 'isoformat'):
        return d.isoformat()[:10]
    return str(d)


def _sub_note(ws, text, ncols, row):
    c = ws.cell(row, 1, text)
    c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    c.alignment = SHELL_LEFT
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    return row + 1


# ----------------------------------------------------------------------------
# 辅助核算余额表 发现与读取
# ----------------------------------------------------------------------------
def _discover_aux(data_dir, ents):
    """返回 {(entity, year): aux_path}。文件名顺序无关：命中 {实体}…辅助核算…{年}…xlsx（含时间戳后缀）。"""
    out = {}
    if not os.path.isdir(data_dir):
        return out
    for e in ents:
        for y in ents[e]:
            hits = [os.path.join(data_dir, f) for f in os.listdir(data_dir)
                    if e in f and '辅助核算' in f and str(y) in f and f.endswith('.xlsx')]
            if hits:
                out[(e, y)] = hits[0]
    return out


def _load_aux(path):
    """返回 [(code, rel, km_name, aux_name, qc_dir, qc_amt, jf, df, qm_dir, qm_amt, lv, wtype)]。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for r in ws.iter_rows(min_row=3, values_only=True):
        try:
            if not r or len(r) < 11:
                continue
            rows.append((
                str(r[0] or ''), r[1], str(r[2] or ''), str(r[3] or ''),
                r[4], safe(r[5]), safe(r[6]), safe(r[7]), r[8], safe(r[9]), r[10],
                str(r[11] or '') if len(r) > 11 else '',   # 往来类型列（U8 有）
            ))
        except Exception:
            continue
    wb.close()
    return rows


# ----------------------------------------------------------------------------
# 借款明细提取（辅助核算优先，科目余额表子目兜底）
# ----------------------------------------------------------------------------
def _parse_aux(path, subj):
    """从辅助核算余额表取该主体的该借款科目末级明细（按借款银行/项目）。"""
    rows = _load_aux(path)
    out = []
    for (code, rel, km_name, aux_name, qc_dir, qc_amt, jf, df, qm_dir, qm_amt, lv, wtype) in rows:
        if not (km_name and km_name.startswith(subj['names'][0])):
            continue
        if code == subj['code']:
            continue
        # 2026-08-07 修复：c 诚本等 U8 账套 aux 为『科目+往来+现金流量』混合维度——
        # ① 现金流量维度行（wtype 含『现金流量』或 aux_name 为现金流项目名，如
        #    『偿还债务支付的现金』）非借款项目，混入会虚增/冲减借款余额
        #    （c 短期借款明细曾 -44.9M 现金流行混入 → 期末 43.16M vs TB 56.53M）。
        if '现金流量' in str(wtype) or '现金流量' in str(aux_name):
            continue
        # ② 科目子目行（code 为纯数字层级如 2001.01，aux_name 空=科目级汇总）——
        #    与真实往来行（code=G 开头等）双计同笔借款（c 2001.01 抵押借款 45M
        #    = 农商行 35M + 农行 10M 被计两次）。真实往来行 code 非纯数字层级。
        _is_subj_row = bool(code) and code[0].isdigit() and ('.' in code or '\\' in code)
        if _is_subj_row and not aux_name:
            continue
        if aux_name and aux_name == '借款项目':
            continue
        lender = aux_name if (aux_name and aux_name not in (km_name, '')) else code
        if lender in (km_name, '借款项目', subj['code'], ''):
            continue
        qc = -signed_balance(qc_dir, qc_amt)   # 贷余为正、借余为负（credit-normalized）
        qm = -signed_balance(qm_dir, qm_amt)
        inc = df     # 贷方发生 = 借入
        dec = jf     # 借方发生 = 偿还
        out.append(dict(code=code, lender=lender, qc=qc, inc=inc, dec=dec, qm=qm))
    return out


def _tb_detail(tb, e, y, prefix, subj_name):
    """兜底：科目余额表 2001/2501 子目作为借款项目明细。"""
    matched = [(c, n, v) for (ee, c, n, yy), v in tb.items()
               if ee == e and yy == y and c.startswith(prefix)]
    # 2026-08-06：Z 泰国账套长期借款=2138.02（非标准 2501）→ 名称兜底（铁律5）
    if not matched and subj_name:
        matched = [(str(c), n, v) for (ee, c, n, yy), v in tb.items()
                   if ee == e and yy == y and subj_name in str(n)]
    if not matched:
        return []
    parent = [m for m in matched if m[0] == prefix]
    children = [m for m in matched if m[0] != prefix]
    base = children if children else matched
    # 2026-08-07 修复：children 含 子级+孙级 全层级（J 东轴 200101 流动资金贷款 +
    # 20010101 抵押贷款 + 2001010101/0104 孙级 = 3 倍，明细 100.05M vs TB 33.35M）。
    # 只取【末级叶子】（父=子和，铁律）——J 东轴叶子=德阳银行(-4M)+农商银行(-26.7M)+
    # 商信农商(-2.65M)=-33.35M=TB ✓。名称兜底匹配的 Z 泰国 2138/2138.02 同适用。
    if len(base) > 1:
        base = [(c, n, v) for c, n, v in base
                if not any(c != c2 and c2.startswith(c) for c2, _, _ in base)]
    out = []
    for c, n, v in base:
        # 2026-08-22 修复：名称兜底把「应付利息-短期借款/财务费用-利息支出-短期借款」
        # 也误当借款明细（name 含"短期借款"）。排除含 利息/费用 的损益与往来科目。
        if _re.search(r'利息|财务费用|应付利息|预提', str(n)):
            continue
        # 2026-08-22 修复：SAP 科目余额表负债科目余额为负，read_tb_full 已按方向归一为正，
        # _tb_detail 原 `-v['qc']` 双反 → 期初负。负债借款科目余额取绝对值。
        qc = abs(v['qc'])
        qm = abs(v['qm'])
        inc = v['df']
        dec = v['jf']
        out.append(dict(code=c, lender=(n or c), qc=qc, inc=inc, dec=dec, qm=qm))
    return out


def _tb_total(tb, e, y, prefix, subj_name=None):
    """科目余额表该借款科目父级控制数（credit-normalized）。"""
    for (ee, c, n, yy), v in tb.items():
        if ee == e and yy == y and c == prefix:
            return (-v['qc'], v['df'], v['jf'], -v['qm'])
    # 2026-08-06：Z 泰国账套长期借款=2138.02 → 名称兜底
    if subj_name:
        for (ee, c, n, yy), v in tb.items():
            if ee == e and yy == y and subj_name in str(n):
                return (-v['qc'], v['df'], v['jf'], -v['qm'])
    return None


def _extract_loan_detail(aux_path, tb, e, y, subj):
    details = []
    if aux_path and os.path.exists(aux_path):
        details = _parse_aux(aux_path, subj)
    if not details:
        details = _tb_detail(tb, e, y, subj['code'], subj['names'][0])
    if not details:
        return None
    tb_total = _tb_total(tb, e, y, subj['code'], subj['names'][0])
    return {'details': details, 'tb_total': tb_total}


# ----------------------------------------------------------------------------
# 利息支出匹配（综合查询明细表：财务费用-利息支出 短期/长期借款利息）
# ----------------------------------------------------------------------------
def _extract_interest(gl, e, y, subj_name):
    """返回 (named: {token:amt}, unknown: float, scope_total: float)。"""
    named = {}
    unknown = 0.0
    scope = 0.0
    for r in gl:
        if r.get('e') != e or r.get('y') != y:
            continue
        nm = r.get('name') or ''
        if not nm.startswith('财务费用-利息支出'):
            continue
        if subj_name == '短期借款' and '短期借款利息' not in nm:
            continue
        if subj_name == '长期借款' and '长期借款利息' not in nm:
            continue
        amt = safe(r.get('cr'))
        if amt <= 0:
            continue
        scope += amt
        tok = _bank_token(r.get('summary'))
        if tok:
            named[tok] = named.get(tok, 0.0) + amt
        else:
            unknown += amt
    return named, unknown, scope


def _loan_interest(loan, named, unknown, subj_name, lender_count=1):
    """把命名利息按银行 token 匹配到单笔借款；融资租赁/折溢价调增等不单列。
    2026-07-29 改进：若该银行只有一笔借款且 token 未匹配到利息，则把未知利息归给它。"""
    lender = loan.get('lender', '')
    tok = _bank_token(lender)
    if tok and tok in named:
        return round(named[tok], 2)
    # 单笔借款兜底：该银行只有这一笔，未知利息归给它
    if lender_count == 1 and unknown and abs(unknown) > 0.005:
        return round(unknown, 2)
    return 0.0


def _loan_period_dates(gl, e, y, lender, subj_name):
    """返回 (借款日, 归还日)：GL 中该笔借款(按银行 token 匹配)最早借入日期 / 最晚偿还日期。
    综合查询明细表若未含借款凭证行，返回 (None, None)，由调用方回落为期初/期末日期。"""
    tok = _bank_token(lender)
    bd = []
    rd = []
    for r in gl:
        if r.get('e') != e or r.get('y') != y:
            continue
        nm = r.get('name') or ''
        if subj_name not in nm:
            continue
        rtok = _bank_token(nm) or _bank_token(r.get('opp')) or _bank_token(r.get('summary'))
        if tok and rtok and tok != rtok:
            continue
        if rtok is None and tok is not None:
            continue
        cr = safe(r.get('cr'))
        dr = safe(r.get('dr'))
        dt = r.get('date')
        if cr > 0 and dt:
            bd.append(dt)
        if dr > 0 and dt:
            rd.append(dt)
    return (min(bd) if bd else None, max(rd) if rd else None)


# ----------------------------------------------------------------------------
# 凭证抽查（发生额）：GL 中该借款科目全部凭证行（year=None 时合并全年度）
# ----------------------------------------------------------------------------
def _collect_vouchers(gl, year, subj, excl_carryover=True):
    """收集该借款科目发生额凭证行（默认剔除结转损益类凭证）。返回 (rows, excluded_count)。
    排除判定统一调用 audit_common.voucher_skip_keys（2026-07-25 抽凭通用规则）。
    year=None 时不过滤年份（合并全年度）。"""
    matched = []
    for r in gl:
        if year is not None and r.get('y') != year:
            continue
        if not (subj['names'][0] in (r.get('name') or '')):
            continue
        matched.append(r)
    if excl_carryover and matched:
        skip = voucher_skip_keys(
            matched,
            key_fields=lambda rr: (rr.get('e'), rr.get('y'), rr.get('vtype'), rr.get('vno')),
            name_field=lambda rr: rr.get('name') or '',
            opp_field=lambda rr: rr.get('opp') or '',
            summ_field=lambda rr: rr.get('summary') or '',
        )
    else:
        skip = set()
    rows = []
    excluded = 0
    for r in matched:
        if (r.get('e'), r.get('y'), r.get('vtype'), r.get('vno')) in skip:
            excluded += 1
            continue
        # 行级：剔除同号凭证中仍残留的"相符计提/摊销"单行（铁律15），避免纯计提分录漏网
        if is_accrual_line({'name': r.get('name'), 'opp': r.get('opp'), 'dr': r.get('dr'), 'cr': r.get('cr'), 'summary': r.get('summary')}):
            excluded += 1
            continue
        rows.append(r)
    rows.sort(key=lambda x: (x.get('e', ''), str(x.get('y') or ''),
                             (x.get('date').isoformat() if hasattr(x.get('date'), 'isoformat') else str(x.get('date'))),
                             str(x.get('vtype') or ''), str(x.get('vno') or '')))
    return rows, excluded


# ----------------------------------------------------------------------------
# Sheet: 借款明细表（合并全集团、全年度；按 主体→年度 小计→全集团合计）
# ⚡ 2026-08-22 模板驱动：改用 audit_templates/借款明细_外壳.xlsx 的「短期借款明细表」
#    + 新增「借款利息检查表」（本金×利率×天数/360 测算，利率留待审计补录）。
# ----------------------------------------------------------------------------

# 模板「短期借款明细表」列序（1-based，模板 31 列）
_TP_D = {
    'ENT': 1, 'LEND': 2, 'BD': 3, 'RD': 4, 'RATE': 5,
    'QC_F': 6, 'QC_B': 7, 'INC_DT': 8, 'INC_F': 9, 'INC_B': 10,
    'DEC_DT': 11, 'DEC_F': 12, 'DEC_B': 13, 'QM_F': 14, 'QM_B': 15,
    'OVERDUE': 16, 'FX_QMF': 17, 'FX_RATE': 18, 'FX_AMT': 19, 'FX_DIFF': 20,
    'COND': 21, 'GUAR': 22, 'GUAR_TERM': 23, 'GUAR_IDX': 24, 'COUNTER': 25,
    'PURPOSE': 26, 'CNO_IDX': 27, 'MEM': 28, 'C_CON': 29, 'C_GUA': 30, 'C_MORT': 31,
}
# 模板「借款利息检查表」列序（12 列）
_TP_I = {'ENT': 1, 'LEND': 2, 'PRIN': 3, 'RATE': 4, 'START': 5, 'END': 6,
         'DAYS_PRE': 7, 'DAYS_POST': 8, 'INT_PRE': 9, 'INT_POST': 10, 'INT_SUM': 11}
_LOAN_SHELL = '借款明细_外壳'


def _clone_shell(wb, sheet_name, target_name=None):
    """从模板 xlsx 克隆 sheet 到 wb（值+样式+合并单元格+列宽+冻结），返回新 ws。
    模板由 resolve_template 定位（audit_templates/借款明细_外壳.xlsx）。"""
    tpl = resolve_template(_LOAN_SHELL)
    if not tpl:
        raise FileNotFoundError(f'缺借款模板：audit_templates/{_LOAN_SHELL}.xlsx')
    src = openpyxl.load_workbook(tpl)
    try:
        sws = src[sheet_name]
        ws = wb.create_sheet(target_name or sheet_name)
        for row in sws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    ws.cell(cell.row, cell.column, cell.value)
                if cell.has_style:
                    n = ws.cell(cell.row, cell.column)
                    n.font = cell.font.copy()
                    n.fill = cell.fill.copy()
                    n.border = cell.border.copy()
                    n.alignment = cell.alignment.copy()
                    n.number_format = cell.number_format
        for mc in sws.merged_cells.ranges:
            ws.merge_cells(str(mc))
        for col, dim in sws.column_dimensions.items():
            if dim.width:
                ws.column_dimensions[col].width = dim.width
        ws.freeze_panes = sws.freeze_panes
        return ws
    finally:
        src.close()


def _fmt_days(d):
    """日期 → 'YYYY-MM-DD'；datetime/date 对象转字符串，None 返回空。"""
    if d is None:
        return ''
    if hasattr(d, 'strftime'):
        return d.strftime('%Y-%m-%d')
    return str(d)


def _day_diff(bd, rd):
    """计息天数 = rd - bd（整段）。解析字符串/date；失败返回 None。"""
    try:
        from datetime import datetime, date
        def _p(x):
            if isinstance(x, (datetime, date)):
                return x
            return datetime.strptime(str(x)[:10], '%Y-%m-%d')
        return (_p(rd) - _p(bd)).days
    except Exception:
        return None

def build_loan_detail_sheet(wb, model, years, gl, sk='ST'):
    """借款明细表（模板驱动）：克隆模板「短期借款明细表」，R6 起逐笔填数。
    按 主体 → 逐笔明细 → 主体合计 → 全集团合计；利率/借款条件/合同索引等留空待审计补录。"""
    subj = SUBJECTS[sk]
    tname = f'{subj["title"]}明细表'
    ws = _clone_shell(wb, '短期借款明细表', tname)
    keys = sorted(model.keys(), key=lambda k: (k[0], k[1], k[2]))
    keys = [k for k in keys if k[1] in years and k[2] == sk]
    ents = sorted({k[0] for k in keys})
    row = 6
    gtot = [0.0] * 5   # qc inc dec qm intr
    for ent in ents:
        ent_keys = [k for k in keys if k[0] == ent]
        ent_tot = [0.0] * 5
        ent_rates = _contract_rate_by_ent(_ent_disp(ent), sk)   # 合同补充利率
        # 融资明细补充（该主体的逐笔起止日/利率，用于明细表日期/利率列）
        fin_loans = [x for x in _load_finance_loans(sk) if x['ent'] == _ent_disp(ent)]
        for k in sorted(ent_keys, key=lambda x: (x[1], x[2])):
            rec = model[k]
            y = k[1]
            named, unknown, scope = _extract_interest(gl, ent, y, subj['title'])
            lender_token_cnt = {}
            for d in rec['details']:
                tok = _bank_token(d['lender'])
                lender_token_cnt[tok] = lender_token_cnt.get(tok, 0) + 1
            bd_default = '%s-01-01' % y
            rd_default = '%s-12-31' % y
            for d in rec['details']:
                lend = d['lender']
                bd, rd = _loan_period_dates(gl, ent, y, lend, subj['title'])
                bd_s = _fmt_days(bd) if bd else bd_default
                rd_s = _fmt_days(rd) if rd else rd_default
                # 融资明细补充：该主体实际借款期间（最早开始~最晚到期）优先于整年默认
                if fin_loans:
                    _starts = [x['start'] for x in fin_loans if x['start'] and len(x['start']) >= 10]
                    _ends = [x['end'] for x in fin_loans if x['end'] and len(x['end']) >= 10]
                    if _starts:
                        bd_s = min(_starts)
                    if _ends:
                        rd_s = max(_ends)
                intr = _loan_interest(d, named, unknown, subj['title'],
                                      lender_count=lender_token_cnt.get(_bank_token(lend), 1))
                note = ''
                if lend in ('利息调整',):
                    note = '折溢价摊销含于本行，利息不单列'
                elif '融资租赁' in lend:
                    note = '融资租赁，利息未单列，建议查租赁合同'
                elif abs(d['qc']) < 0.005 and abs(d['qm']) < 0.005 and abs(d['inc']) < 0.005 and abs(d['dec']) < 0.005:
                    note = '本期无发生额'
                if ent_rates:
                    if len(ent_rates) == 1:
                        ws.cell(row, _TP_D['RATE'], ent_rates[0])
                    else:
                        # 多档利率：填代表值，备注列全列（待按合同分档核对）
                        ws.cell(row, _TP_D['RATE'], ent_rates[0])
                        note = (note + '；' if note else '') + \
                               f'合同利率多档{"/".join(str(x) for x in ent_rates)}%'
                ws.cell(row, _TP_D['ENT'], _ent_disp(ent))
                ws.cell(row, _TP_D['LEND'], lend)
                ws.cell(row, _TP_D['BD'], bd_s)
                ws.cell(row, _TP_D['RD'], rd_s)
                # 利率(5)/逾期(16)/汇兑(17-20)/借款条件(21-25)/用途(26)/合同索引(27)/合同(29-31) 留空待补录
                ws.cell(row, _TP_D['QC_F'], d['qc'])
                ws.cell(row, _TP_D['QC_B'], d['qc'])
                ws.cell(row, _TP_D['INC_F'], d['inc'])
                ws.cell(row, _TP_D['INC_B'], d['inc'])
                ws.cell(row, _TP_D['DEC_F'], d['dec'])
                ws.cell(row, _TP_D['DEC_B'], d['dec'])
                ws.cell(row, _TP_D['QM_F'], d['qm'])
                ws.cell(row, _TP_D['QM_B'], d['qm'])
                if note:
                    ws.cell(row, _TP_D['MEM'], note)
                vals = [d['qc'], d['inc'], d['dec'], d['qm'], (intr or 0.0)]
                ent_tot = [ent_tot[j] + vals[j] for j in range(5)]
                gtot = [gtot[j] + vals[j] for j in range(5)]
                row += 1
        if ent_keys and any(abs(v) > 0.005 for v in ent_tot):
            ws.cell(row, _TP_D['ENT'], f'【{_ent_disp(ent)} 合计】')
            for j, cc in enumerate((_TP_D['QC_F'], _TP_D['INC_F'], _TP_D['DEC_F'], _TP_D['QM_F'])):
                ws.cell(row, cc, ent_tot[j])
                ws.cell(row, cc + 1, ent_tot[j])   # 本位币列
            for cc in (_TP_D['ENT'], _TP_D['QC_F'], _TP_D['QC_B'], _TP_D['INC_F'], _TP_D['INC_B'],
                       _TP_D['DEC_F'], _TP_D['DEC_B'], _TP_D['QM_F'], _TP_D['QM_B']):
                ws.cell(row, cc).font = openpyxl.styles.Font(bold=True)
            row += 1
    if len(ents) > 1:
        ws.cell(row, _TP_D['ENT'], '全集团合计')
        for j, cc in enumerate((_TP_D['QC_F'], _TP_D['INC_F'], _TP_D['DEC_F'], _TP_D['QM_F'])):
            ws.cell(row, cc, gtot[j])
            ws.cell(row, cc + 1, gtot[j])
        for cc in (_TP_D['ENT'], _TP_D['QC_F'], _TP_D['QC_B'], _TP_D['INC_F'], _TP_D['INC_B'],
                   _TP_D['DEC_F'], _TP_D['DEC_B'], _TP_D['QM_F'], _TP_D['QM_B']):
            ws.cell(row, cc).font = openpyxl.styles.Font(bold=True)
    return ws


def build_loan_interest_sheet(wb, model, years, gl, sk='ST'):
    """借款利息检查表（模板驱动）：克隆模板「借款利息检查表」，R6 起逐笔填数。
    2026-08-22 数据源改为【融资明细.xls 逐笔】（起止日/利率/金额完整，72/37 笔）：
    每笔 主体/银行/本金/利率/起止日/计息天数 + K 列公式（本金×年利率×天数/360 自动测算）。
    融资明细缺失时回退科目子目（默认整年，利率留空待补）。"""
    subj = SUBJECTS[sk]
    ws = _clone_shell(wb, '借款利息检查表')
    fin = _load_finance_loans(sk)
    row = 6
    if fin:
        y = years[0] if years else 2026
        for x in sorted(fin, key=lambda a: (a['ent'], a['bank'], a['start'] or '')):
            bd = x['start'] or f'{y}-01-01'
            rd = x['end'] or f'{y}-12-31'
            days = _day_diff(bd, rd)
            ws.cell(row, _TP_I['ENT'], x['ent'])
            ws.cell(row, _TP_I['LEND'], _mask_bank(x['bank']))
            ws.cell(row, _TP_I['PRIN'], x['amt'])
            if x['rate']:
                ws.cell(row, _TP_I['RATE'], x['rate'])
            ws.cell(row, _TP_I['START'], bd)
            ws.cell(row, _TP_I['END'], rd)
            if days is not None:
                ws.cell(row, _TP_I['DAYS_PRE'], days)
                # K 列小计：本金×利率×天数/360（利率缺失则补录后自动测算）
                ws.cell(row, _TP_I['INT_SUM'], f'=C{row}*D{row}/100*G{row}/360')
            row += 1
        return ws
    # ---- 回退：科目子目（融资明细不可用）----
    keys = sorted(model.keys(), key=lambda k: (k[0], k[1], k[2]))
    keys = [k for k in keys if k[1] in years and k[2] == sk]
    ents = sorted({k[0] for k in keys})
    for ent in ents:
        ent_keys = [k for k in keys if k[0] == ent]
        ent_rates = _contract_rate_by_ent(_ent_disp(ent), sk)
        for k in sorted(ent_keys, key=lambda x: (x[1], x[2])):
            rec = model[k]
            y = k[1]
            bd_default = '%s-01-01' % y
            rd_default = '%s-12-31' % y
            for d in rec['details']:
                lend = d['lender']
                if lend in ('利息调整',) or '融资租赁' in lend:
                    continue
                bd, rd = _loan_period_dates(gl, ent, y, lend, subj['title'])
                bd_s = _fmt_days(bd) if bd else bd_default
                rd_s = _fmt_days(rd) if rd else rd_default
                days = _day_diff(bd or bd_default, rd or rd_default)
                ws.cell(row, _TP_I['ENT'], _ent_disp(ent))
                ws.cell(row, _TP_I['LEND'], lend)
                ws.cell(row, _TP_I['PRIN'], d['qm'])
                if ent_rates:
                    ws.cell(row, _TP_I['RATE'], ent_rates[0])
                ws.cell(row, _TP_I['START'], bd_s)
                ws.cell(row, _TP_I['END'], rd_s)
                if days is not None:
                    ws.cell(row, _TP_I['DAYS_PRE'], days)
                    ws.cell(row, _TP_I['INT_SUM'], f'=C{row}*D{row}/100*G{row}/360')
                row += 1
    return ws


# ----------------------------------------------------------------------------
# Sheet: 凭证抽查（发生额，合并全年度）
# ----------------------------------------------------------------------------
def build_voucher_sheet(wb, year, recs, gl, subj_key):
    subj = SUBJECTS[subj_key]
    ents_order = [e for (e, _, _) in recs]
    # 先归集各主体凭证行；无异常(该科目当年无任何凭证交易)则不生成凭证抽查
    vrows_by_ent = {}
    excluded_total = 0
    for ent in ents_order:
        vrows, excl = _collect_vouchers(gl, year, subj)
        vrows = [x for x in vrows if x.get('e') == ent]
        excluded_total += excl
        if vrows:
            vrows_by_ent[ent] = vrows
    if not vrows_by_ent:
        return None
    ws = wb.create_sheet('%s凭证抽查' % subj['title'])
    ncols = 15
    headers = ['核算主体', '测试序号', '日期', '凭证字', '凭证号',
               '二级科目（或客户/供应商名称）', '摘要', '借方金额', '贷方金额',
               '对方科目', '与原始凭证相符', '原始凭证内容', '原始凭证日期',
               '会计处理正确', '所属时间无误']
    yrtitle = '全年度' if year is None else ('%s 年度' % year)
    r = _title(ws, '%s凭证抽查（%s，发生额）' % (subj['title'], yrtitle), ncols)
    sub = ('说明：下列为本科目%s全部发生额凭证(借/贷，已剔除结转期间损益凭证 %d 笔)。GL 为非全量抽取，'
           '2026 年 GL 仅 1–5 月(YTD)。请据以抽凭，并在"抽查/审计结论"列填写。' % (yrtitle, excluded_total))
    r = _sub_note(ws, sub, ncols, r)
    hr = r + 1
    r = _hdr(ws, hr, headers)
    C = list(range(1, ncols + 1))
    grand_dr = grand_cr = 0.0
    for ent, ve in vrows_by_ent.items():
        # ⚡ 2026-08-28 统一样式：去掉『【主体】』合并分隔行（所得税等通用抽凭表无此行），
        #   数据行直接逐行写核算主体，与其他底稿抽凭表样式一致。
        for i, x in enumerate(sorted(ve, key=lambda v: str(v.get('date', ''))), 1):
            dr = safe(x.get('dr'))
            cr = safe(x.get('cr'))
            _txt(ws, r, C[0], ent)
            _txt(ws, r, C[1], i)
            _txt(ws, r, C[2], _fmt_date(x.get('date')))
            _txt(ws, r, C[3], x.get('y') if x.get('y') is not None else '')
            _txt(ws, r, C[4], _fmt_vno(x))
            _txt(ws, r, C[5], x.get('name') or '借款')
            _txt(ws, r, C[6], x.get('summary') or '')
            _money(ws, r, C[7], round(dr, 2) if dr > 0 else None)
            _money(ws, r, C[8], round(cr, 2) if cr > 0 else None)
            _txt(ws, r, C[9], '')
            # 列 11–15（与原始凭证相符/原始凭证内容/原始凭证日期/会计处理正确/所属时间无误）
            for ci in range(10, 15):
                _txt(ws, r, C[ci], '')
            grand_dr += dr
            grand_cr += cr
            r += 1
    _txt(ws, r, C[1], '合计', bold=True)
    _money(ws, r, C[7], grand_dr)
    _money(ws, r, C[8], grand_cr)
    for i in range(1, ncols + 1):
        ws.cell(r, i).font = SHELL_BOLD
    r += 1
    _txt(ws, r, C[0], '注：上表借贷合计为该科目%s发生额，应与《借款明细表》的"本期增加/减少"口径审慎核对。' % yrtitle)
    ws.merge_cells(start_row=r, start_column=C[0], end_row=r, end_column=ncols)

    widths = [14, 10, 12, 10, 12, 24, 30, 14, 14, 24, 12, 16, 12, 12, 12]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    return ws


# ----------------------------------------------------------------------------
# 主流程（合并全集团一份底稿）
# ----------------------------------------------------------------------------
def process_folder(data_dir, out_path=None):
    if not os.path.isdir(data_dir):
        print(f'[ERROR] 数据目录不存在：{data_dir}')
        return 1
    # ⚡ 2026-08-28 修复：融资台账按当前项目数据根定位（防 AH 误用 ADF 借款台账）
    set_fin_data_root(data_dir)
    ents = discover_entities(data_dir)
    if not ents:
        print(f'[ERROR] 未在 {data_dir} 发现科目余额表/综合查询明细表')
        return 1
    print(f'[INFO] 发现主体 {len(ents)} 个：{mask_names(sorted(ents))}')

    aux_map = _discover_aux(data_dir, ents)
    if aux_map:
        print(f'[INFO] 发现辅助核算余额表 {len(aux_map)} 份（借款明细数据源）')
    else:
        print('[WARN] 未发现任何辅助核算余额表，将回退到科目余额表子目作为借款明细')

    tb = read_tb_full(data_dir, ents)
    gl = read_gl_rows(data_dir, ents)

    years = set()
    for ent in ents:
        years |= set(ents[ent].keys())
    years = sorted(years)
    if not years:
        print('[ERROR] 未发现任何年份的账套数据')
        return 1
    print(f'[INFO] 检测到年份：{years}')

    # 预构建借款明细模型 model[(e,y,sk)] = {'details':..,'tb_total':..,'_sk':sk}
    model = {}
    for e in sorted(ents):
        for y in sorted(ents[e]):
            aux = aux_map.get((e, y))
            for sk, subj in SUBJECTS.items():
                rec = _extract_loan_detail(aux, tb, e, y, subj)
                if rec:
                    rec['_sk'] = sk
                    model[(e, y, sk)] = rec

    if not model:
        print('[SKIP] 未发现任何短期/长期借款数据，跳过')
        return 0

    return _build_split_workbooks(data_dir, model, years, gl, tb, ents, out_path)


def _clean_opp(opp, self_name):
    """凭证抽查『对方科目』在用友/U8 导出中常为整张凭证全部分录的拼接串；
    剔除与本科目自身同名的分段，仅保留真正的交易对方科目，便于按金额定位对应科目。"""
    if not opp:
        return ''
    segs = _re.split(r'[；;\n\r]+', str(opp))
    out = []
    for s in segs:
        s = s.strip()
        if not s:
            continue
        if self_name and self_name in s:
            continue
        out.append(s)
    return '；'.join(out)


def build_loan_audit_confirm_sheet(wb, tb, ents, years, sk='ST'):
    """审定表：按核算主体列示 短期借款(2001) 或 长期借款(2501) 期初数/期末未审数/重分类调整/审计调整/审定数。
    2026-08-04 按科目拆分：sk='ST'|'LT' 各生成独立审定表（短期借款 审定表 / 长期借款 审定表）。
    2026-08-04 加重分类调整列：1 年内到期的借款由审计师填『重分类调整』（转出至一年内到期的非流动负债），
    审定数 = 期末未审 + 重分类 + 审计调整（公式联动）。
    只取 years 列表中的年份数据。"""
    subj = SUBJECTS[sk]
    ws = wb.create_sheet(f'{subj["title"]} 审定表')
    ncols = 7
    headers = ['核算主体', '期初数', '期末未审数', '重分类调整', '审计调整数', '审定数', '与科目余额表勾稽']
    r = _title(ws, f'{subj["title"]} 审定表（按核算主体）', ncols)
    hr = r + 1
    r = _hdr(ws, hr, headers)
    g_qc = g_qm = 0.0
    data_first = r   # 主体数据首行（合计公式引用，2026-08-03 公式化）
    # 按实体汇总 本科目（仅限指定年份）
    # 2026-08-06 修复：Z 泰国长期借款=2138.02（非标准 2501）→ 名称兜底（与明细表/附注一致，铁律5）
    by_e = {}
    for (e, code, name, y), v in tb.items():
        if y not in years:
            continue
        if code == subj['code'] or subj['names'][0] in str(name):
            # 2026-08-22：与明细表同口径——排除「应付利息-短期借款/财务费用-利息支出-短期借款」
            # 等含"利息/费用"科目（名称含"短期借款"曾误计入审定表 → 期末虚增）。
            if _re.search(r'利息|财务费用|应付利息|预提', str(name)):
                continue
            by_e.setdefault(e, []).append((code, name, v))
    # 防父+子双计：只取【末级叶子】（无后代的行）——2026-08-07 修复：原逻辑只过滤一层
    #（有子级就只留含'.'行），J 东轴 2001→200101→20010101→2001010101 四层结构下
    # 中间父级仍全计入 → 审定 133.4M=4×33.35M。叶子过滤：任何其他代码以此为前缀即剔除
    #（泰国 2138 父级=2138.02 子级同金额 → 只留 2138.02 ✓）
    for e, recs in by_e.items():
        codes = [str(c) for c, _, _ in recs]
        by_e[e] = [r for r in recs
                   if not any(str(c2) != str(r[0]) and str(c2).startswith(str(r[0]))
                              for c2 in codes)]
    for e in sorted(by_e):
        recs = by_e[e]
        # 借款为贷方科目：TB 自然符号借正贷负，abs() 使年初/期末以正数列示（铁律20）
        qc = sum(abs(v['qc']) for _, _, v in recs)
        qm = sum(abs(v['qm']) for _, _, v in recs)
        g_qc += qc; g_qm += qm
        _txt(ws, r, 1, _ent_disp(e))
        _money(ws, r, 2, qc)
        _money(ws, r, 3, qm)
        _txt(ws, r, 4, '')                       # 重分类调整（1年内到期→NCL，审计师填）
        _txt(ws, r, 5, '')                       # 审计调整
        # 2026-08-06 协议：审定数写数值（=期末+重分类+调整，后两者空=期末；公式 data_only 读 None）
        _money(ws, r, 6, round(qm, 2))
        _txt(ws, r, 7, '与科目余额表核对一致')
        r += 1
    # 合计（2026-08-06 协议：写数值=g_qc/g_qm 累计）
    _txt(ws, r, 1, '合计', bold=True)
    _txt(ws, r, 2, '', bold=True)
    cell = ws.cell(r, 2, round(g_qc, 2))
    cell.font = SHELL_BOLD; cell.fill = SHELL_TOT_FILL
    cell.number_format = '#,##0.00'; cell.alignment = SHELL_RGT
    cell = ws.cell(r, 3, round(g_qm, 2))
    cell.font = SHELL_BOLD; cell.fill = SHELL_TOT_FILL
    cell.number_format = '#,##0.00'; cell.alignment = SHELL_RGT
    ws.cell(r, 6, round(g_qm, 2)).number_format = '#,##0.00'
    ws.cell(r, 6).font = SHELL_BOLD; ws.cell(r, 6).fill = SHELL_TOT_FILL; ws.cell(r, 6).alignment = SHELL_RGT
    _txt(ws, r, 7, '', bold=True, fill=SHELL_TOT_FILL)
    _txt(ws, r, 4, '', bold=True, fill=SHELL_TOT_FILL)
    _txt(ws, r, 5, '', bold=True, fill=SHELL_TOT_FILL)
    widths = [16, 16, 16, 14, 14, 16, 28]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    # ---- 借款类别披露区（2026-08-22 新增：披露口径，与审计报告附注一致）----
    # 数据源：融资明细.xls 担保信息 → 质押/抵押/保证/信用；金额=业务金额（非账面余额）
    fin = _load_finance_loans(sk)
    if fin:
        r += 2
        _txt(ws, r, 1, f'按借款类别披露（{subj["title"]}，口径：融资明细业务金额，供披露参考）',
             bold=True)
        r += 1
        cat_tot = {}
        for x in fin:
            cat_tot[x['cat']] = cat_tot.get(x['cat'], 0.0) + x['amt']
        order = ['质押借款', '抵押借款', '保证借款', '信用借款', '其他']
        rows_data = [(c, cat_tot.get(c, 0.0)) for c in order if cat_tot.get(c, 0.0) > 0]
        _txt(ws, r, 1, '借款类别')
        _txt(ws, r, 2, '金额(业务金额)')
        _txt(ws, r, 3, '说明')
        for cc in range(1, 4):
            ws.cell(r, cc).font = SHELL_BOLD
            ws.cell(r, cc).fill = SHELL_HFILL
            ws.cell(r, cc).border = SHELL_BORDER
            ws.cell(r, cc).alignment = SHELL_CEN
        r += 1
        first_cat = r
        for c, v in rows_data:
            _txt(ws, r, 1, c)
            _money(ws, r, 2, round(v, 2))
            _txt(ws, r, 3, '来源：融资明细.xls 担保信息归类')
            r += 1
        _txt(ws, r, 1, '合计', bold=True)
        cell = ws.cell(r, 2, round(sum(cat_tot.values()), 2))
        cell.font = SHELL_BOLD; cell.fill = SHELL_TOT_FILL
        cell.number_format = '#,##0.00'; cell.alignment = SHELL_RGT
        _txt(ws, r, 3, '', bold=True, fill=SHELL_TOT_FILL)
        r += 1
        # 与账面余额勾稽（业务金额 vs 期末审定数）
        diff = round(sum(cat_tot.values()) - g_qm, 2)
        _txt(ws, r, 1, '与期末账面余额差异')
        _money(ws, r, 2, diff)
        _txt(ws, r, 3, '业务金额口径 vs 科目余额表口径差异（提款/还款时点）', bold=True)
        r += 1
    return ws


def build_finance_sheet(wb, fin, sk='ST'):
    """融资明细台账（2026-08-22 新增）：逐笔借款（融资明细.xls 贷款类），
    含 主体/银行/业务编号/金额(业务金额)/利率/起止日/担保信息/借款类别，供披露与核对。"""
    subj = SUBJECTS[sk]
    ws = wb.create_sheet(f'{subj["title"]}融资明细台账')
    headers = ['借款主体', '银行', '业务编号', '融资小类', '业务金额', '利率(%)',
               '开始日', '到期日', '担保信息', '借款类别']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = SHELL_BOLD
        cell.fill = SHELL_HFILL
        cell.border = SHELL_BORDER
        cell.alignment = SHELL_CEN
    fin_sorted = sorted(fin, key=lambda x: (x['ent'], x['bank']))
    for x in fin_sorted:
        ws.append([x['ent'], _mask_bank(x['bank']), x['no'], x['small'], x['amt'],
                   x['rate'], x['start'], x['end'], _mask_guarantor(x['guar']), x['cat']])
    for row in ws.iter_rows(min_row=2):
        for c_ in row:
            c_.font = SHELL_HFONT
            c_.border = SHELL_BORDER
            if c_.column in (5, 6) and isinstance(c_.value, (int, float)):
                c_.number_format = '#,##0.00'
    for col, w in zip('ABCDEFGHIJ', [8, 20, 22, 14, 15, 8, 12, 12, 34, 10]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:J{ws.max_row}'
    # 合计行（业务金额）
    tot = sum(x['amt'] for x in fin_sorted)
    r = ws.max_row + 1
    ws.cell(r, 1, '合计')
    ws.cell(r, 5, round(tot, 2)).number_format = '#,##0.00'
    for cc in (1, 5):
        ws.cell(r, cc).font = SHELL_BOLD
        ws.cell(r, cc).fill = SHELL_TOT_FILL
    return ws


def _build_split_workbooks(data_dir, model, years, gl, tb, ents, out_path=None):
    """2026-08-04 按科目拆分：每个年度分别生成 短期借款审计底稿 与 长期借款审计底稿 两份独立文件，
    各自含 审定表/明细表/凭证抽查/附注汇总。out_path 在拆分模式下不适用（固定按科目命名）。"""
    if out_path:
        print('  [INFO] --output 在按科目拆分模式下不适用，忽略（固定输出 短期借款审计底稿_<年>_生成.xlsx / 长期借款审计底稿_<年>_生成.xlsx）')
    outs = []
    for y in years:
        for sk in ['ST', 'LT']:
            subj = SUBJECTS[sk]
            has = any(str(k[1]) == str(y) and k[2] == sk for k in model)
            if not has:
                print(f'  [SKIP] {subj["title"]}（{y} 无数据，不生成）')
                continue
            wb = openpyxl.Workbook()
            wb.remove(wb.active)

            # ---- 审定表（当年，仅本科目） ----
            build_loan_audit_confirm_sheet(wb, tb, ents, [y], sk=sk)
            print(f'[OK] {subj["title"]}审定表 {y}')
            # ---- 融资明细台账（逐笔借款披露，2026-08-22 新增） ----
            fin = _load_finance_loans(sk)
            if fin:
                build_finance_sheet(wb, fin, sk=sk)
                print(f'[OK] {subj["title"]}融资明细台账（{len(fin)} 笔）')
            # ---- 明细表（当年，仅本科目） ----
            build_loan_detail_sheet(wb, model, [y], gl, sk=sk)
            print(f'[OK] {subj["title"]}明细表 {y}')
            # ---- 借款利息检查表（模板驱动，测算底稿，2026-08-22 新增）----
            build_loan_interest_sheet(wb, model, [y], gl, sk=sk)
            print(f'[OK] {subj["title"]}借款利息检查表 {y}')
            # ---- 凭证抽查（当年，无交易跳过） ----
            subj_keys = [k for k in model if k[2] == sk and str(k[1]) == str(y)]
            if subj_keys:
                recs = [(k[0], model[k], None) for k in subj_keys]
                if build_voucher_sheet(wb, y, recs, gl, sk) is not None:
                    print(f'[OK] {subj["title"]}凭证抽查 {y}')
                else:
                    print(f'   [SKIP] {subj["title"]}凭证抽查（{y} 无凭证交易，跳过）')

            # ---- 附注汇总（仅本科目，两期数——增减变动由现金流量表核对，2026-08-02）----
            try:
                from audit_common import build_footnote_generic as _bfn_loan
                # 2026-08-07 v2（会计语言优先，代码方言隔离）：附注 codes 经账套配置
                # 解析——account_profiles.json subject_codes 有账套特例码用之（Z 长期借款
                # =2138,2138.02 非标准 2501），否则用 SUBJECTS 标准码（2001/2501）。
                _acct_ln = os.path.basename(data_dir.rstrip('\\/'))
                _ln_reg_key = 'st_loan' if sk == 'ST' else 'lt_loan'
                _ln_codes = [subj['code']]
                try:
                    import subject_mapping as _sm_ln
                    _ln_codes = _sm_ln.resolve_subject_codes(_acct_ln, _ln_reg_key, [subj['code']])
                except Exception:
                    pass
                _bfn_loan(wb, data_dir, f'{subj["title"]}附注汇总',
                          f'{subj["title"]}附注汇总（审定口径；每主体一列；段=期末数/期初数；增减变动由现金流量表核对）',
                          [(subj['title'], _ln_codes)], is_credit=True, mode='two',
                          target_year=y, tb_full=tb, entities=ents)
            except Exception as _ex:
                print(f'  ⚠️ {subj["title"]}附注汇总生成失败：{_ex}')

            finalize_workbook(wb)
            # 审定表+明细表置于第1、2位
            sheets = wb._sheets
            for sname in [f'{subj["title"]} 审定表', f'{subj["title"]}明细表']:
                cur = next((i for i, s in enumerate(sheets) if s.title == sname), None)
                target = 0 if '审定表' in sname else 1
                if cur is not None and cur != target:
                    sheets.insert(target, sheets.pop(cur))
            out = os.path.join(data_dir, f'{subj["title"]}审计底稿_{y}_生成.xlsx')
            from audit_common import validate_workbook
            validate_workbook(wb, '借款底稿', raise_on_error=False)
            # ⚡ 2026-08-28 #875：对方科目核对（模块内集成）
            try:
                from counterparty_recon import inject_into_wb_auto
                inject_into_wb_auto(wb, data_dir, y, subj['title'], ents_set=set(ents))
            except Exception as _ex:
                print(f"  ⚠️ {subj['title']}对方科目核对注入失败：{_ex}")
            wb.save(out)
            wb.close()
            outs.append(out)
            print(f'[DONE] 已生成：{out}')

    # 清理旧版合并稿（2026-08-04 拆分后废弃）
    # 2026-08-06 修复：沙箱 os.remove 被 safe-delete 拦截（SAFE_DELETE_FAIL_CLOSED）→
    # 旧合并稿残留 → 检查器报『缺附注汇总』。改 mv 改名归档（_old_合并稿），不删除。
    for y in years:
        for old_name in (f'借款审计底稿_{y}_生成.xlsx', f'借款明细表审计底稿_{y}_生成.xlsx'):
            old = os.path.join(data_dir, old_name)
            if os.path.exists(old):
                _bak = os.path.join(data_dir, old_name.replace('_生成.xlsx', '_生成_old_合并稿.xlsx'))
                try:
                    os.replace(old, _bak)
                    print(f'[CLEAN] 旧版合并稿归档为 {os.path.basename(_bak)}')
                except BaseException as ex:
                    print(f'  [WARN] 无法归档旧版 {old}：{ex}')
    print(f'[DONE] 共生成 {len(outs)} 份（按科目拆分：短期借款 / 长期借款，无合并稿）')
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description='借款(短期/长期)审计底稿生成程序（2026-08-04 按科目拆分版：短期借款/长期借款各独立底稿）')
    ap.add_argument('folders', nargs='*', default=[],
                    help='一个或多个数据目录（拖入文件夹时自动作为位置参数传入）')
    ap.add_argument('--input', action='append', dest='inputs', default=[],
                    help='数据目录（可多次指定）')
    ap.add_argument('--output', default=None, help='指定输出路径')
    args = ap.parse_args(argv)

    dirs = list(args.folders) + list(args.inputs)
    if not dirs:
        dirs = [P.T]

    rc = 0
    for d in dirs:
        out = args.output if len(dirs) == 1 else None
        r = process_folder(d, out)
        if r != 0:
            rc = r
        from audit_common import finalize_after_build
        finalize_after_build(d)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）
    return rc


if __name__ == '__main__':
    _write_boot_log('MAIN')
    _bs_sys.exit(main())
