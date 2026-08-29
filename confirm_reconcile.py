# -*- coding: utf-8 -*-
"""函证回函与底稿科目核对（AZ 2025，#14 v2）。

数据：
  回函解析：{回函目录}/回函_解析结果_全量.json（confirm_parser 全量解析 91 份）
  底稿：{账套}/底稿/2025/银行存款审计底稿_2025_生成.xlsx + 各往来科目审计底稿_2025_生成.xlsx
        （交易额核对直接取底稿『本期借方/贷方发生额』，不从账套再取数）

核对口径：
  发函金额    = 函证列示金额（从回函 OCR 提取；企业询证函『往来款项列示』为发函方填列的账面数）
  回函确认金额 = 文件名结论『相符』→ 发函金额；『不符/未标注』→ 空（差异在结论区，待人工）
  底稿余额    = 科目明细表期末余额(人民币)（账面）
  差异        = 回函确认金额 − 底稿余额
  交易额      = 函证『交易列示』销售/采购发生额（不含税），与底稿该单位应收借方/应付贷方发生额
                （含税）并列对照——口径不同，标注提示，不做差额判定

输出：
  {账套}/底稿/2025/函证回函_核对表_2025.xlsx（3 sheets：往来核对/银行核对/未提取清单）

用法：
  python confirm_reconcile.py AZ [--year 2025]
"""
import json
import os
import re
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import paths as P

import openpyxl
from openpyxl.styles import Font, PatternFill

# 主体名映射：文件名被审计单位全称 → 底稿核算主体简称
SUBJ_MAP = [
    ('杭州兆龙物联技术有限公司合肥分公司', '物联合肥'),
    ('杭州兆龙物联技术有限公司深圳分公司', '物联深圳'),
    ('杭州兆龙物联技术有限公司苏州分公司', '物联苏州'),
    ('杭州兆龙物联技术有限公司', '物联'),
    ('浙江兆龙互连科技股份有限公司上海分公司', '上分'),
    ('浙江兆龙互连科技股份有限公司北京分公司', '北分'),
    ('浙江兆龙互连科技股份有限公司', '母公司'),
    ('浙江兆龙数链科技有限公司', '数链'),
    ('浙江兆龙高分子材料有限公司', '高分子'),
    ('LONGTEK INTERCONNECT (THAILAND) CO., LTD.', '泰国'),
    ('LONGTEK HOLDING GROUP PTE. LTD.', '新加坡'),
    ('LONGTEK SINGAPORE TRADING PTE. LTD.', '新加坡'),
    ('LONGTEK', '泰国'),
]

# 科目 → (底稿文件名前缀, 明细表 sheet 名)
SUBJ_FILES = {
    '应收账款': ('应收账款', '应收账款明细表_2025'),
    '应付账款': ('应付账款', '应付账款明细表_2025'),
    '预付账款': ('预付账款', '预付账款明细表_2025'),
    '预收账款': ('预收账款', '预收账款明细表_2025'),
    '其他应收款': ('其他应收款', '其他应收款明细表_2025'),
    '其他应付款': ('其他应付款', '其他应付款明细表_2025'),
}

# 交易方向 → 取底稿发生额列（销售→应收借方发生额=新增应收；采购→应付贷方发生额=新增应付）
TRADE_WP = {
    '销售': ('应收账款', 'dr'),
    '采购': ('应付账款', 'cr'),
}


def map_subj(auditee_full):
    for k, v in SUBJ_MAP:
        if k in auditee_full:
            return v
    return auditee_full


def load_ar_detail(fp, sheet):
    """明细表 → {(主体, 往来单位名称): {'bal':期末余额, 'dr':本期借方, 'cr':本期贷方}}。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        return {}
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    hdr_i = None
    for i, r in enumerate(rows[:6]):
        if any(str(x) == '往来单位名称' for x in r):
            hdr_i = i
            break
    if hdr_i is None:
        return {}
    hdr = [str(x or '') for x in rows[hdr_i]]

    def _ci(name):
        return hdr.index(name) if name in hdr else None

    ci_comp = _ci('核算主体') or 0
    ci_name = _ci('往来单位名称')
    ci_bal = _ci('期末余额(人民币)') or (len(hdr) - 1)
    ci_dr = _ci('本期借方发生额')
    ci_cr = _ci('本期贷方发生额')
    out = {}
    for r in rows[hdr_i + 1:]:
        if len(r) <= max(ci_comp, ci_name, ci_bal):
            continue
        comp = str(r[ci_comp] or '').strip()
        name = str(r[ci_name] or '').strip()
        if not comp or not name or '合计' in comp or '小计' in comp or comp.isdigit():
            continue

        def _f(i):
            try:
                return float(r[i] or 0)
            except (TypeError, ValueError):
                return 0.0

        rec = out.setdefault((comp, name), {'bal': 0.0, 'dr': 0.0, 'cr': 0.0})
        rec['bal'] = round(rec['bal'] + _f(ci_bal), 2)
        if ci_dr is not None:
            rec['dr'] = round(rec['dr'] + _f(ci_dr), 2)
        if ci_cr is not None:
            rec['cr'] = round(rec['cr'] + _f(ci_cr), 2)
    return out


def match_amount(detail, subj, name):
    """在明细中匹配 (主体, 对方名) → (记录 or None, 匹配方式)。"""
    key = (subj, name)
    if key in detail:
        return detail[key], '精确'
    cands = [(k, v) for (s, n), v in detail.items()
             if s == subj and (name in n or n in name)]
    if cands:
        cands.sort(key=lambda x: -abs(x[1]['bal']))
        return cands[0][1], '包含'
    return None, '未找到'


def load_bank_detail(fp):
    """银行存款明细表2025 → {主体: 本币期末余额合计}。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['银行存款明细表2025']
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(x or '') for x in rows[0]]
    ci_comp = hdr.index('核算主体')
    ci_amt = hdr.index('期末余额(本币)') if '期末余额(本币)' in hdr else None
    out = {}
    for r in rows[1:]:
        comp = str(r[ci_comp] or '').strip()
        if not comp or '合计' in comp or comp.isdigit():
            continue
        try:
            v = float(r[ci_amt] or 0)
        except (TypeError, ValueError):
            continue
        out[comp] = round(out.get(comp, 0.0) + v, 2)
    return out


def confirm_conclusion_from_filename(base):
    m = re.search(r'_(相符|不符|未标注)(?:\.pdf)?$', base)
    return m.group(1) if m else '未标注'


def main(acct='AZ', year='2025'):
    root = P.DATA_DIRS.get(acct)
    if not root:
        print(f'账套不存在：{acct}')
        return 1
    redo_dir = os.path.join(root, '数据', year, '回函')
    json_fp = os.path.join(redo_dir, '回函_解析结果_全量.json')
    if not os.path.exists(json_fp):
        print(f'缺解析结果：{json_fp}')
        return 1
    rows = json.load(open(json_fp, encoding='utf-8'))

    wp_dir = os.path.join(root, '底稿', year)
    details = {}
    for subj, (prefix, sheet) in SUBJ_FILES.items():
        fp = os.path.join(wp_dir, f'{prefix}审计底稿_{year}_生成.xlsx')
        details[subj] = load_ar_detail(fp, sheet)
    bank_acct = load_bank_detail(os.path.join(wp_dir, f'银行存款审计底稿_{year}_生成.xlsx'))

    ar_rows = []
    for r in rows:
        if r.get('type') != '企业':
            continue
        base = r['file']
        parts = base.split('_')
        auditee_full = parts[0] if parts else ''
        counter = parts[1] if len(parts) > 1 else ''
        subj_short = map_subj(auditee_full)
        concl = confirm_conclusion_from_filename(base)

        def _confirm_val(amt):
            """回函确认金额：相符→发函金额；不符/未标注→None（待人工）。"""
            return amt if concl == '相符' else None

        for it in r.get('ar_items', []):
            subj = it['item']
            if subj not in details:
                continue
            rec, match = match_amount(details[subj], subj_short, counter)
            wp_bal = rec['bal'] if rec else None
            diff = None if (wp_bal is None or _confirm_val(it['amount']) is None) else round(_confirm_val(it['amount']) - wp_bal, 2)
            flag = ''
            if diff is not None and abs(diff) > 0.01:
                flag = '差异'
            if wp_bal and it['amount'] > 0 and wp_bal > 0:
                ratio = it['amount'] / wp_bal
                if ratio > 5 or ratio < 0.2:
                    flag = '差异·疑似OCR错位'
            ar_rows.append({
                'ref': r.get('ref_no', ''), 'auditee': subj_short, 'counter': counter,
                'conclusion': concl, 'subject': subj,
                'send_amt': it['amount'], 'confirm_amt': _confirm_val(it['amount']),
                'wp_amt': wp_bal, 'diff': diff, 'match': match, 'flag': flag,
            })
        # 交易额：发函金额=交易额；回函确认=相符→交易额；底稿=对应发生额（含税，仅对照）
        for it in r.get('trade_items', []):
            key, col = TRADE_WP.get(it['item'], (None, None))
            wp_val = None
            wp_subj = ''
            if key and key in details:
                rec, match = match_amount(details[key], subj_short, counter)
                if rec:
                    wp_val = rec[col]
                    wp_subj = key
            flag = '交易额(不含税)vs底稿发生额(含税)'
            if wp_val is None:
                flag = '交易额·底稿未找到对应发生额'
            ar_rows.append({
                'ref': r.get('ref_no', ''), 'auditee': subj_short, 'counter': counter,
                'conclusion': concl, 'subject': f"{it['item']}(交易)",
                'send_amt': it['amount'], 'confirm_amt': _confirm_val(it['amount']),
                'wp_amt': wp_val, 'diff': None, 'match': wp_subj or '-', 'flag': flag,
            })

    # ========== 银行函证核对（逐银行明细 + 主体汇总） ==========
    bank_detail_rows = []
    for r in rows:
        if r.get('type') != '银行':
            continue
        base = r['file']
        parts = base.split('_')
        auditee_full = parts[0] if parts else ''
        bank_name = parts[1] if len(parts) > 1 else ''
        subj_short = map_subj(auditee_full)
        concl = confirm_conclusion_from_filename(base)
        send = round(r.get('bank_total', 0), 2)
        confirm = send if (concl == '相符' and send) else None   # 余额未提取≠确认0，置空待人工
        loan = round(r.get('loan_total', 0), 2) or None
        ob = [f"{it['no']}{it['name'][:8]}" for it in r.get('bank_items', [])
              if it.get('no', 0) >= 3 and it.get('conclusion') == '有明细']
        wp = bank_acct.get(subj_short)
        note = ''
        if not r.get('bank_total'):
            note = '余额未提取'
        if loan:
            note = (note + ';' if note else '') + f'银行借款{loan:,.2f}'
        bank_detail_rows.append({
            'auditee': subj_short, 'bank': bank_name, 'ref': r.get('ref_no', ''),
            'conclusion': concl, 'send': send, 'confirm': confirm,
            'accts': r.get('bank_count', 0), 'loan': loan,
            'off_balance': ';'.join(ob), 'wp': wp, 'note': note,
        })

    by_subj = {}
    for r in bank_detail_rows:
        by_subj.setdefault(r['auditee'], []).append(r)
    bank_rows = []
    for subj, rs in by_subj.items():
        n_bank = len(rs)
        n_extracted = sum(1 for r in rs if r['send'] > 0)
        send = round(sum(r['send'] for r in rs), 2)
        all_match = all(r['conclusion'] == '相符' for r in rs)
        confirm = send if (all_match and send) else None
        wp = bank_acct.get(subj)
        diff = None if (wp is None or confirm is None) else round(confirm - wp, 2)
        not_ext = [r['bank'] for r in rs if not r['send']]
        bank_rows.append({
            'auditee': subj, 'n_bank': n_bank, 'n_extracted': n_extracted,
            'send': send, 'confirm': confirm, 'wp': wp, 'diff': diff,
            'not_extracted': ';'.join(not_ext[:6]),
        })

    no_rows = []
    for r in rows:
        if r.get('type') == '企业' and not r.get('ar_items') and not r.get('trade_items'):
            no_rows.append({'file': r['file'], 'type': '企业往来', 'reason': '金额未提取（英文函证/特殊格式/OCR质量差）'})
        if r.get('type') == '银行' and not r.get('bank_total'):
            no_rows.append({'file': r['file'], 'type': '银行', 'reason': '银行存款余额未提取'})

    # ========== 写 Excel ==========
    out_fp = os.path.join(wp_dir, f'函证回函_核对表_{year}.xlsx')
    wb = openpyxl.Workbook()
    hdr_font = Font(bold=True, color='FFFFFF')
    hdr_fill = PatternFill('solid', fgColor='4472C4')
    diff_fill = PatternFill('solid', fgColor='FFC7CE')

    ws = wb.active
    ws.title = '往来函证核对'
    ws.append(['函证编号', '核算主体', '对方', '文件名结论', '科目',
               '发函金额(函证列示)', '回函确认金额', '底稿余额(期末)', '差异(回函-底稿)', '匹配方式', '标注'])
    for h in ws[1]:
        h.font = hdr_font
        h.fill = hdr_fill
    for r in ar_rows:
        ws.append([r['ref'], r['auditee'], r['counter'], r['conclusion'], r['subject'],
                   r['send_amt'], r['confirm_amt'], r['wp_amt'], r['diff'], r['match'], r['flag']])
        if r['flag'] and '差异' in str(r['flag']) and 'OCR' not in str(r['flag']):
            for c in ws[ws.max_row]:
                c.fill = diff_fill
    ws.freeze_panes = 'A2'

    ws1b = wb.create_sheet('银行函证回函明细')
    ws1b.append(['核算主体', '银行名称', '函证编号', '文件名结论',
                 '发函金额(回函提取)', '回函确认金额', '账户数', '银行借款(余额)',
                 '表外事项(有明细)', '账面银行存款(主体)', '备注'])
    for h in ws1b[1]:
        h.font = hdr_font
        h.fill = hdr_fill
    for r in bank_detail_rows:
        ws1b.append([r['auditee'], r['bank'], r['ref'], r['conclusion'],
                     r['send'], r['confirm'], r['accts'], r['loan'],
                     r['off_balance'], r['wp'], r['note']])
        if r['conclusion'] != '相符':
            for c in ws1b[ws1b.max_row]:
                c.fill = PatternFill('solid', fgColor='FFF2CC')
    ws1b.freeze_panes = 'A2'

    ws2 = wb.create_sheet('银行函证核对')
    ws2.append(['核算主体', '回函银行数', '已提取余额银行数',
                '发函金额(回函提取)', '回函确认金额', '账面银行存款余额', '差异(回函-账面)', '未提取余额银行'])
    for h in ws2[1]:
        h.font = hdr_font
        h.fill = hdr_fill
    for r in bank_rows:
        ws2.append([r['auditee'], r['n_bank'], r['n_extracted'],
                    r['send'], r['confirm'], r['wp'], r['diff'], r['not_extracted']])
        if r['diff'] is not None and abs(r['diff']) > 0.01:
            for c in ws2[ws2.max_row]:
                c.fill = diff_fill
    ws2.freeze_panes = 'A2'

    ws3 = wb.create_sheet('未提取金额清单')
    ws3.append(['文件', '类型', '原因'])
    for h in ws3[1]:
        h.font = hdr_font
        h.fill = hdr_fill
    for r in no_rows:
        ws3.append([r['file'], r['type'], r['reason']])
    ws3.freeze_panes = 'A2'

    wb.save(out_fp)
    print(f'已生成：{out_fp}')
    print(f'  往来核对 {len(ar_rows)} 行 | 银行明细 {len(bank_detail_rows)} 行 | 银行汇总 {len(bank_rows)} 行 | 未提取 {len(no_rows)} 份')
    n_diff = sum(1 for r in ar_rows if r['flag'] == '差异')
    n_trade = sum(1 for r in ar_rows if '(交易)' in r['subject'])
    print(f'  往来余额差异 {n_diff} 行 | 交易额对照 {n_trade} 行')
    return 0


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    _year = '2025'
    if '--year' in sys.argv:
        _year = sys.argv[sys.argv.index('--year') + 1]
    sys.exit(main(args[0] if args else 'AZ', _year))
