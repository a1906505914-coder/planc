# -*- coding: utf-8 -*-
"""函证回函全量解析（AZ 2025，#13）：扫描 OCR txt（递归，含文件夹型回函）→ confirm_parser 结构化
→ 脱敏（confirm_mask：发函公司→A代码，回函方→1-2字）
→ 银行回函 1-14 项逐项展开
→ 输出 {回函目录}/回函_解析结果_脱敏.json + 回函_脱敏明细表.xlsx（脱敏版，可外发）。

用法：
  python confirm_parse_all.py AZ [--year 2025]
"""
import glob
import json
import os
import re
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import paths as P
import confirm_parser as CP
import confirm_mask as CM

BANK_ITEM_NAMES = {1: '银行存款', 2: '银行借款', 3: '委托贷款(委托)', 4: '委托贷款(借款)',
                   5: '担保(提供)', 6: '担保(贵行提供)', 7: '银行承兑汇票', 8: '已贴现商票',
                   9: '托收商票', 10: '信用证', 11: '外汇买卖合约', 12: '托管证券',
                   13: '理财产品', 14: '其他'}

# 自有格式银行回函（境外/自有模板，无 1-14 编号）的逐项关键词判定
_ITEM_KW = {
    1: ['存款', '账户', 'balance', 'account', 'current account'],
    2: ['借款', '贷款', 'loan', 'facility', 'borrow'],
    3: ['委托贷款', 'entrusted loan', 'trust loan'],
    4: ['委托贷款', 'entrusted loan'],
    5: ['担保', '保证', 'guarantee', 'guarantor'],
    6: ['保函', '备用信用证', 'standby', 'letter of guarantee'],
    7: ['承兑', 'acceptance', '承兑汇票'],
    8: ['贴现', 'discount'],
    9: ['托收', 'collection'],
    10: ['信用证', 'letter of credit', 'L/C', 'documentary credit'],
    11: ['外汇', '远期', '衍生', 'swap', 'forward', 'foreign exchange'],
    12: ['托管', 'custody', '证券', 'share', 'bond'],
    13: ['理财', '基金', '结构性存款', 'investment', 'wealth'],
    14: ['其他', 'other'],
}

# 境外自有格式回函：实际列示项目解析
_FOREIGN_CUR = ('USD', 'THB', 'SGD', 'CNH', 'CNY', 'EUR', 'HKD', 'GBP', 'JPY', 'AUD')
_FOREIGN_AMT_RE = re.compile(r'([A-Z]{3})\s*([\d][\d,，]*\.\d{1,2})')
_FOREIGN_KW = [
    ('贷款/借款', ['loan', 'borrow', 'facility', 'overdraft', 'lending', '借款', '贷款']),
    ('担保', ['guarantee', 'guarantor', 'collateral', 'surety', '担保']),
    ('票据/承兑', ['acceptance', 'bill', 'discount', '承兑', '票据']),
    ('信用证', ['letter of credit', 'documentary credit', '信用证']),
    ('外汇/衍生', ['foreign exchange', 'forward', 'swap', 'derivative', '外汇']),
    ('托管/证券', ['custody', 'security', 'share', 'bond', '托管', '证券']),
    ('理财/投资', ['investment', 'wealth', 'deposit product', '理财']),
]


def _clean_amt(s):
    """OCR 金额清洗：'10.449.79' → 10449.79（中间点为千分位）。"""
    parts = s.split('.')
    if len(parts) > 2:
        s = ','.join(parts[:-1]) + '.' + parts[-1]
    return float(s.replace(',', '').replace('，', ''))


def _parse_foreign(text):
    """境外自有格式银行回函 → 实际列示项目。
    返回 [{'item','conclusion','summary'}]：存款(金额) / 中泰ITEM / 关键词提及项。"""
    lines = text.splitlines()
    tl = text.lower()
    out = []

    # ① 存款：币种+金额（同行或紧邻上一行；币种白名单；排除 0/小额/利率）
    def _find_cur(line):
        """行内查找白名单币种（兼容 '3SGD'/'THB2.7' 数字紧贴）。"""
        if not line:
            return None
        for m in re.finditer(r'([A-Z]{3})(?![A-Z])', line):
            if m.group(1) in _FOREIGN_CUR:
                return m.group(1)
        return None

    _AMT_IN_LINE_RE = re.compile(r'\b([\d][\d,，.]*\.\d{1,2})')
    amts = []
    for i, ln in enumerate(lines):
        cur = _find_cur(ln)
        if cur:
            for m in _AMT_IN_LINE_RE.finditer(ln):
                try:
                    v = _clean_amt(m.group(1))
                    if v >= 1:
                        amts.append(f'{cur} {m.group(1)}')
                except ValueError:
                    pass
        else:
            # 单独金额：紧邻上一行有白名单币种
            cm = _find_cur(lines[i - 1]) if i > 0 else None
            if cm:
                for m in _AMT_IN_LINE_RE.finditer(ln):
                    try:
                        v = _clean_amt(m.group(1))
                        if v >= 1 and 'NIL' not in ln:
                            amts.append(f'{cm} {m.group(1)}')
                    except ValueError:
                        pass
    amts = list(dict.fromkeys(amts))[:6]
    if amts:
        out.append({'item': '银行存款(账户余额)', 'conclusion': '有明细',
                    'summary': '；'.join(amts)})
    else:
        out.append({'item': '银行存款', 'conclusion': '未列示', 'summary': ''})

    # ② 中泰 ITEM1-14 结构
    item_ms = list(re.finditer(r'\bITEM\s*(\d{1,2})\b', text))
    if item_ms:
        segs = []
        for k, m in enumerate(item_ms):
            end = item_ms[k + 1].start() if k + 1 < len(item_ms) else m.end() + 250
            seg = re.sub(r'\s+', ' ', text[m.end():end]).strip()[:60]
            has_nil = 'NIL' in seg
            has_amt = bool(re.search(r'\d[\d,，.]*\.\d{1,2}', seg))
            tag = '有数据' if has_amt else ('NIL' if has_nil else '待核实')
            segs.append(f'ITEM{m.group(1)}[{tag}]{seg}')
        out.append({'item': '中泰ITEM结构', 'conclusion': '见明细',
                    'summary': ' | '.join(segs[:10])})
    else:
        # ③ 其他项目关键词（仅非中泰格式）
        for label, kws in _FOREIGN_KW:
            if any(k.lower() in tl for k in kws):
                out.append({'item': label, 'conclusion': '提及', 'summary': ''})
    return out


def _classify_bank_items(text, items):
    """自有格式回函（无 1-14 编号）→ 按含金额行（含上下文）归属项目。
    只对【金额行所在项】判有明细，其余项 → 待核实/未提及。"""
    raw = text.splitlines()
    matched = {}
    for i, ln in enumerate(raw):
        if not re.search(r'\d[\d,，]{3,}\.\d{1,2}', ln) \
                and not re.search(r'\b\d{8,}\b', ln):
            continue
        ctx = ' '.join(raw[max(0, i - 1):i + 2]).lower()
        for no, kws in _ITEM_KW.items():
            if any(k.lower() in ctx for k in kws):
                matched[no] = ln.strip()[:60]
                break
    for it in items:
        if it['conclusion'] in ('有明细', '无'):
            continue
        if it['no'] in matched:
            it['conclusion'] = '有明细'
            it['summary'] = matched[it['no']]
        elif any(k.lower() in text.lower() for k in _ITEM_KW.get(it['no'], [])):
            it['conclusion'] = '待核实'
        else:
            it['conclusion'] = '未提及'
    return items


def _split_base(base):
    """文件名 → (发函方, 回函方, 编号, 结论)。"""
    parts = base.split('_')
    if len(parts) >= 4:
        return parts[0].strip(), parts[1].strip(), parts[2].strip(), parts[3].strip()
    if len(parts) >= 2:
        return parts[0].strip(), parts[1].strip(), '', ''
    return base, '', '', ''


def _is_foreign_text(text):
    """文本是否以英文/泰文为主（境外自由格式）→ 需补充实际列示项目解析。
    中文占比 < 50% 视为境外（与银行简称编码无关，纯文本特征判断）。"""
    if not text:
        return False
    zh = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
    alpha = sum(1 for c in text if c.isalpha())
    total = zh + alpha
    if total == 0:
        return False
    return zh / total < 0.5


def _write_excel(out_xlsx, rows):
    """脱敏明细表：Sheet1 汇总 / Sheet2 银行 1-14 项展开。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = '回函汇总'
    hdr = ['序号', '发函公司(代码)', '回函方(简称)', '类型', '编号', '结论',
           '存款合计(元)', '存款笔数', '借款合计(元)', '借款笔数',
           '不符差异', '回函日期', '备注']
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='DDEBF7')
    for i, r in enumerate(rows, 1):
        ws.append([i, r.get('auditee', ''), r.get('reply', ''), r.get('type', ''),
                   r.get('ref_no', ''), r.get('conclusion', ''),
                   r.get('bank_total', 0) or r.get('deposit_total', 0),
                   r.get('bank_count', 0) or r.get('deposit_accts', 0),
                   r.get('loan_total', 0), r.get('loan_count', 0),
                   r.get('discrepancy', ''), r.get('reply_date', ''), ''])
    # 列宽
    for col, w in zip('ABCDEFGHIJKLMN', (6, 22, 20, 6, 14, 10, 16, 8, 16, 8, 30, 12, 12)):
        ws.column_dimensions[col].width = w

    # Sheet2 银行 1-14 项展开（长表）
    ws2 = wb.create_sheet('银行14项展开')
    hdr2 = ['序号', '发函公司(代码)', '回函方(代码)', '编号', '项目', '结论', '摘要']
    ws2.append(hdr2)
    for c in ws2[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='E2EFDA')
    n = 0
    for r in rows:
        if r.get('type') != '银行':
            continue
        for it in r.get('bank_items', []):
            n += 1
            ws2.append([n, r.get('auditee', ''), r.get('reply', ''),
                        f"{it.get('no', '')}", it.get('name', ''), it.get('conclusion', ''),
                        it.get('summary', '')])
    for col, w in zip('ABCDEFG', (6, 22, 20, 6, 46, 10, 60)):
        ws2.column_dimensions[col].width = w

    # Sheet3 境外回函实际列示项目
    ws3 = wb.create_sheet('境外回函实际项目')
    hdr3 = ['序号', '发函公司(代码)', '回函方(代码)', '项目', '结论', '摘要/金额']
    ws3.append(hdr3)
    for c in ws3[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='FFF2CC')
    m = 0
    for r in rows:
        if r.get('type') != '银行' or not r.get('foreign_items'):
            continue
        for it in r['foreign_items']:
            m += 1
            ws3.append([m, r.get('auditee', ''), r.get('reply', ''),
                        it.get('item', ''), it.get('conclusion', ''),
                        it.get('summary', '')])
    for col, w in zip('ABCDEF', (6, 22, 20, 34, 10, 80)):
        ws3.column_dimensions[col].width = w

    wb.save(out_xlsx)


def main(acct='AZ', year='2025'):
    root = P.DATA_DIRS.get(acct)
    if not root:
        print(f'账套不存在：{acct}')
        return 1
    redo_root = os.path.join(root, '数据', year, '回函')
    # ⚡ 只读【本地脱敏版】（先运行 回函处理.bat 脱敏），不接触原始含主体名数据
    ocr_root = os.path.join(redo_root, '回函_脱敏版')
    if not os.path.isdir(ocr_root):
        print(f'缺脱敏版目录：{ocr_root}')
        print('请先双击运行 回函处理.bat 生成脱敏版，再执行本程序。')
        return 1
    # 脱敏过期检测：OCR 文本比脱敏版新 → 强制提示先重新脱敏
    def _newest_mtime(d):
        mt = 0
        for dp, _, fns in os.walk(d):
            for fn in fns:
                mt = max(mt, os.path.getmtime(os.path.join(dp, fn)))
        return mt
    src = os.path.join(redo_root, '_ocr_识别结果')
    if os.path.isdir(src):
        src_mt = _newest_mtime(src)
        mask_mt = _newest_mtime(ocr_root)
        if src_mt > mask_mt + 1:
            print(f'⚠️ 检测到 _ocr_识别结果 比脱敏版新，请先运行 回函处理.bat 重新脱敏！')
            return 1
    rows = []
    for sub in ('企业询证函回函', '银行询证函回函'):
        kind = '企业' if sub.startswith('企业') else '银行'
        root2 = os.path.join(ocr_root, sub)
        txs = sorted(glob.glob(os.path.join(root2, '**', '*.txt'), recursive=True))
        seen = set()
        n = 0
        for fp in txs:
            rel = os.path.relpath(fp, root2)
            # 文件夹型回函：跳过发函件（Audit Request Letter），只取回函
            low = rel.lower()
            if 'request' in low or 'request letter' in low or '发函' in rel:
                continue
            base = rel.split(os.sep)[0][:-4] if os.sep in rel else rel[:-4]
            if base in seen:
                continue
            seen.add(base)
            n += 1
            text = open(fp, encoding='utf-8', errors='replace').read()
            aud, reply, ref, concl = _split_base(base)
            # 脱敏
            aud_m = CM.mask_auditee(aud) or CM.mask_text(aud) or aud
            reply_m = CM.mask_reply(reply) or CM.mask_text(reply) or reply
            try:
                r = CP.parse_confirm(text)
            except Exception as e:
                rows.append({'file': base, 'auditee': aud_m, 'reply': reply_m,
                             'type': kind, 'ref_no': ref, 'parse_error': str(e)})
                continue
            items = CP._find_bank_items(text) if kind == '银行' else []
            foreign_items = []
            if kind == '银行':
                # 标准 1-14 编号格式（识别到 ≥5 个编号）→ 已展开；
                # 自有格式（境外/自由模板，编号 <5）→ 关键词补判定 + 境外实际项目
                recognized = sum(1 for it in items
                                 if it['conclusion'] != '未识别')
                if recognized < 5:
                    items = _classify_bank_items(text, items)
                    foreign_items = _parse_foreign(text)
                elif _is_foreign_text(text):
                    # 境外回函即使有编号结构，也补充实际列示项目（存款金额等）
                    foreign_items = _parse_foreign(text)
                # 强约束：金额提取到 → 对应项有明细（金额提取比分段判定更可靠）
                dep_total, dep_n = CP._find_bank_deposit(text)
                loan_total, loan_n = CP._find_bank_loan(text)
                for it in items:
                    if it['no'] == 1 and dep_total > 0:
                        it['conclusion'] = '有明细'
                    elif it['no'] == 2 and loan_total > 0:
                        it['conclusion'] = '有明细'
            row = {
                'file': base,
                'auditee': aud_m, 'reply': reply_m, 'type': kind,
                'ref_no': r.get('ref_no') or ref,
                'conclusion': r.get('conclusion', ''),
                'discrepancy': CM.mask_text(r.get('discrepancy', ''))[:200],
                'reply_date': r.get('reply_date', ''),
                'deposit_total': r.get('deposit_total', 0),
                'deposit_accts': r.get('deposit_accts', 0),
                'loan_total': r.get('loan_total', 0),
                'loan_count': r.get('loan_count', 0),
                'bank_total': (CP._find_bank_deposit(text)[0] if kind == '银行' else 0),
                'bank_count': (CP._find_bank_deposit(text)[1] if kind == '银行' else 0),
                'ar_items': r.get('ar_amounts', []),
                'trade_items': r.get('trade_amounts', []),
                'bank_items': items,          # 银行 1-14 项逐项
                'foreign_items': foreign_items,  # 境外回函实际列示项目
            }
            rows.append(row)
        print(f'  [{kind}] {n} 份')
    out_json = os.path.join(redo_root, '回函_解析结果_脱敏.json')
    out_xlsx = os.path.join(redo_root, '回函_脱敏明细表.xlsx')
    json.dump(rows, open(out_json, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    _write_excel(out_xlsx, rows)
    nb = sum(1 for r in rows if r['type'] == '银行')
    nba = sum(1 for r in rows if r['type'] == '银行' and r.get('bank_total', 0) > 0)
    naa = sum(1 for r in rows if r['type'] == '企业' and r.get('ar_items'))
    nat = sum(1 for r in rows if r['type'] == '企业' and r.get('trade_items'))
    print(f'总 {len(rows)} | 银行 {nb}（有余额 {nba}）| 企业有往来 {naa} 有交易 {nat}')
    print(f'已写入：{out_json}')
    print(f'已写入：{out_xlsx}')
    return 0


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    _year = '2025'
    if '--year' in sys.argv:
        _year = sys.argv[sys.argv.index('--year') + 1]
    sys.exit(main(args[0] if args else 'AZ', _year))
