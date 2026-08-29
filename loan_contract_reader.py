# -*- coding: utf-8 -*-
"""loan_contract_reader.py —— 借款合同台账（脱敏）v2
双轨：
  ① 文件名/目录解析（92份全覆盖）：借款主体/贷款方/合同类型/文件名金额/起止日(文件名)
  ② 文本型内容补充：合同编号/大写金额/利率/起止日(OCR/文本)
输出：脱敏台账 Excel（双方名称→代号 借X/贷X）
"""
import os, re, sys, glob
from collections import OrderedDict
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

BORROWER_MAP = {'国望': '借01', '苏州纤维': '借02', '中鲈': '借03',
                '港虹': '借04', '苏震': '借05', '新视界': '借06'}
LENDER_MAP = [('中国银行', '贷01'), ('中行', '贷01'),
              ('农业银行', '贷02'), ('农行', '贷02'),
              ('工商银行', '贷03'), ('工行', '贷03'),
              ('建设银行', '贷04'), ('建行', '贷04'),
              ('江苏银行', '贷05'),
              ('招银金租', '贷06'),
              ('委贷', '贷07')]
CN_NUM = {'零': 0, '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6,
          '柒': 7, '捌': 8, '玖': 9}


def cn_amount_to_num(s):
    """中文大写金额 → 数值。正确处理 壹仟肆佰万元整 / 伍亿贰仟万元整 等。"""
    total = 0.0
    seg = 0.0        # 当前段（万以下/亿以下）累计
    cur = 0.0        # 当前个位数字
    mult = {'拾': 10, '佰': 100, '仟': 1000}
    for ch in s:
        if ch in CN_NUM:
            cur = CN_NUM[ch]
        elif ch in mult:
            seg += cur * mult[ch] if cur else mult[ch]
            cur = 0
        elif ch == '万':
            seg += cur
            total += seg * 10000
            seg = 0; cur = 0
        elif ch == '亿':
            seg += cur
            total = (total + seg) * 100000000
            seg = 0; cur = 0
        elif ch in ('元', '圆'):
            total += seg + cur
            seg = 0; cur = 0
        elif ch == '角':
            total += cur * 0.1; cur = 0
        elif ch == '分':
            total += cur * 0.01; cur = 0
    if cur:
        total += seg + cur
    elif seg:
        total += seg
    return round(total, 2) if total else None


def parse_filename_info(rel):
    """从路径/文件名解析：借款主体/贷款方/合同类型/金额/起止日。"""
    parts = rel.split(os.sep)
    first = parts[0]
    borrower = re.sub(r'^\d+', '', first)          # 3100国望→国望
    fn = os.path.basename(rel)
    full = ' '.join(parts) + ' ' + fn              # 目录名+文件名（金额/银行常在目录名）
    # 贷款方
    lender = None; lender_code = None
    for kw, code in LENDER_MAP:
        if kw in full:
            lender, lender_code = kw, code
            break
    if not lender_code and '委贷' in full:
        lender, lender_code = '委贷方(非银行)', '贷07'
    # 合同类型（目录/文件名关键词）
    ctype = '其他'
    if re.search(r'保字|保证|担保|（保）|最高额保证', full): ctype = '保证'
    elif re.search(r'押字|抵押|（抵）|最高额抵押', full): ctype = '抵押'
    elif re.search(r'租赁协议|融资租赁|金租', full): ctype = '租赁'
    elif re.search(r'借据', full): ctype = '借据'
    elif re.search(r'委贷|委托借据|委托贷款', full): ctype = '委贷'
    elif re.search(r'借字|借款|流贷|LDZJ|（流贷）|流动资金', full): ctype = '借款'
    # 金额（整个路径搜索：港虹0.8亿/5000万/21354万/1.6亿/0.7亿）
    amt = None
    mm = re.search(r'(\d+(?:\.\d+)?)\s*(亿|万)', full)
    if mm:
        v = float(mm.group(1))
        amt = v * 100000000 if mm.group(2) == '亿' else v * 10000
    # 起止日（文件名 2025.02.26-2030.02.26 / 2025.12.30-2027.06.29）
    period = None
    mm = re.search(r'(\d{4}[./]\d{1,2}[./]\d{1,2})\s*[-~至]\s*(\d{4}[./]\d{1,2}[./]\d{1,2})', full)
    if mm:
        period = f'{mm.group(1)}~{mm.group(2)}'
    return dict(borrower=borrower, lender=lender, lender_code=lender_code,
                ctype=ctype, amt=amt, period=period)


def _ocr_fix(s):
    """OCR 常见误识纠错。"""
    if not s: return ''
    s = s.replace('吉', '壹').replace('从民币', '人民币').replace('从民市', '人民币')
    s = s.replace('Q', '0').replace('O', '0')
    s = s.replace('木', '不').replace('风', '，').replace('口', '（')
    s = s.replace('任', '仟').replace('侣', '佰').replace('位', '亿')
    s = s.replace('红苏', '江苏').replace('苏州长', '苏州长').replace('江办', '江苏')
    return s


def _cn_amt_ctx(t):
    """在文本中找大写金额附近的小写数字金额（优先格式：大写后跟小写）。"""
    # 找 (小写) 后紧跟的数字金额
    mm = re.search(r'[（(]?小写[）)]?\s*[¥￥]?\s*([0-9][0-9,，.]{2,})', t)
    if mm:
        try:
            return float(mm.group(1).replace(',', '').replace('，', ''))
        except ValueError:
            return None
    return None


def extract_content(rel, text):
    """文本/OCR内容补充：合同编号/金额/利率/期限/起止日/担保抵押信息。返回dict(仅命中字段)。"""
    out = OrderedDict()
    t = _ocr_fix(text or '')
    # 合同类型（文本标题为准，细化保证/抵押/质押/借款）
    if re.search(r'最高额保证合同|保证合同|（保）|保字', t): out['合同类型'] = '保证'
    elif re.search(r'最高额抵押合同|抵押合同|押字', t): out['合同类型'] = '抵押'
    elif re.search(r'质押合同|质押|质字', t): out['合同类型'] = '质押'
    elif re.search(r'流动资金借款合同|借款合同|借字', t): out['合同类型'] = '借款'
    # 合同编号
    mm = re.search(r'合同编号\s*[：:]\s*([A-Za-z0-9\-]{6,})', t)
    if not mm:
        mm = re.search(r'(JK\d{10,}|HTZ[A-Za-z0-9]{8,}|LDZJ[A-Za-z0-9]{6,}|示范区[借保押]字\d+[- ]?\d*号?|'
                       r'32010\d{9,}|3210\d{9,}|0110200016-\d{4}年（\S+）字\d+号|WJ\d{6}|BZ\d{10,}|CC\d+[A-Z0-9]+)', t)
    if mm:
        out['合同编号'] = mm.group(1)
    # 担保人 / 抵押人 / 出质人（保证/抵押合同主体）——脱敏只判身份类型
    for kw in ('保证人', '担保人', '抵押人', '出质人'):
        mm = re.search(rf'{kw}\s*[（(]?全称[）)]?\s*[：:]\s*([^\n]{{2,30}}?)\s*(?:统一社会信用代码|法定代表人|负责人|住所地|住所|电话|传真|$)', t)
        if not mm:
            mm = re.search(rf'{kw}\s*[：:]\s*([^\n]{{2,30}}?)\s*(?:统一社会信用代码|法定代表人|负责人|住所地|住所|电话|传真|$)', t)
        if not mm:
            mm = re.search(rf'{kw}\s*[（(]?全称[）)]?\s*([^\n]{{2,30}}?)(?:\s*为了|，)', t)
        if mm:
            nm = mm.group(1).strip()
            if len(nm) >= 2 and not re.match(r'^\d+$', nm):
                out['担保人'] = nm
                break
    # 被担保最高债权额（保证/抵押合同的担保金额：大写+小写）
    if re.search(r'被担保最高债权额|最高本金余额|担保债权|被担保的主债权|本金数额', t):
        amt = None
        mm = re.search(r'（大写）\s*[（(]?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分]{5,}整?)[）)]?', t)
        if mm:
            amt = cn_amount_to_num(_ocr_fix(mm.group(1)))
        if not amt:
            mm = re.search(r'（币种及大写金额）\s*为\s*人民币?\s*([零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分]{5,}整?)', t)
            if mm:
                amt = cn_amount_to_num(_ocr_fix(mm.group(1)))
        if not amt:
            amt = _cn_amt_ctx(t)
        if amt:
            # OCR 错位防护：大写缺失而小写异常（>10亿且非整百/整万）→ 标记待核实，不入账
            if amt > 1000000000 and not re.search(r'（大写）\s*[（(]?[零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分]{5,}整?', t):
                out['_amt_warn'] = f'担保金额OCR可疑({amt:,.0f}，大写缺失)'
                amt = None
            if amt:
                out['担保金额(元)'] = round(amt, 2)
    # 保证方式（连带责任保证/一般保证）——OCR可能把"责任"识别为"责仟/货任"
    mm = re.search(r'保证方式\s*为下列第\s*(\d)\s*项', t)
    if mm:
        out['保证方式'] = '连带责任保证' if mm.group(1) == '1' else '一般保证'
    else:
        mm = re.search(r'保证方式[^。]{0,40}(连带(?:责仟|责任|货任)保证|一般保证)', t, re.S)
        if mm:
            out['保证方式'] = '连带责任保证' if '连带' in mm.group(1) else '一般保证'
    # 抵押物 / 质押物（附清单描述；正文提取简短描述）
    if re.search(r'抵押物', t):
        mm = re.search(r'抵押物\s*[：:]?\s*([\u4e00-\u9fffA-Za-z0-9（）()]{2,40}?)(?:。|；|，|$)', t)
        if mm:
            out['抵押物'] = mm.group(1).strip()
        else:
            out['抵押物'] = '见抵押物清单'
    if re.search(r'质押物|质押财产', t) and '抵押' not in out:
        mm = re.search(r'质押物\s*[：:]?\s*([\u4e00-\u9fffA-Za-z0-9（）()]{2,40}?)(?:。|；|，|$)', t)
        if mm:
            out['质押物'] = mm.group(1).strip()
        else:
            out['质押物'] = '见质押物清单'
    # 抵押物价值/评估价值
    if re.search(r'评估价值|抵押物价值|担保价值', t):
        amt = _cn_amt_ctx(t) or None
        mm = re.search(r'(?:评估价值|抵押物价值)\s*[：:]?\s*[¥￥]?\s*([0-9][0-9,，.]{2,})', t)
        if mm:
            try:
                amt = float(mm.group(1).replace(',', '').replace('，', ''))
            except ValueError:
                pass
        if amt:
            out['抵押物价值(元)'] = round(amt, 2)
    # 借款金额：大写 + 小写
    amt = None
    mm = re.search(r'借款(?:币种及)?金额\s*[：:]?\s*[（(]?大写[）)]?\s*[：:]?\s*[（(]?人民币?[）)]?\s*[（(]?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分]{5,}整?)[）)]?', t)
    if mm:
        amt = cn_amount_to_num(mm.group(1))
    if not amt:
        mm = re.search(r'[（(]?小写[）)]?\s*[¥￥]?\s*([0-9][0-9,，.]{3,})', t)
        if mm:
            try:
                amt = float(mm.group(1).replace(',', '').replace('，', ''))
            except ValueError:
                amt = None
    if not amt:
        mm = re.search(r'借款(?:币种及)?金额\s*[：:]?\s*[（(]?大写[）)]?\s*[：:]?\s*[（(]?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分]{5,})', t)
        if mm:
            amt = cn_amount_to_num(mm.group(1))
    if amt:
        out['金额(元)'] = round(amt, 2)
    # 利率
    rate = None
    mm = re.search(r'(?:执行)?年利率\s*[（(]?[：:]?\s*([0-9]+(?:\.[0-9]+)?)\s*%', t)
    if not mm:
        mm = re.search(r'利率\s*[：:]?\s*([0-9]+(?:\.[0-9]+)?)\s*%', t)
    if not mm:
        mm = re.search(r'LPR\s*[+-]\s*([0-9]+(?:\.[0-9]+)?)\s*%', t)
    if mm:
        rate = float(mm.group(1))
    if rate is not None:
        out['利率(%)'] = rate
    # 期限
    mm = re.search(r'借款期限\s*[：:]?\s*[（(]?[：:]?\s*(\d+)\s*个月', t)
    if not mm:
        mm = re.search(r'期限\s*[：:]?\s*(\d+)\s*个月', t)
    if not mm:
        mm = re.search(r'(\d+)\s*个月', t)
    if mm:
        out['期限'] = f'{mm.group(1)}个月'
    # 起止日
    mm = re.search(r'自\s*(\d{4}年\d{1,2}月\d{1,2}日)\s*至\s*(\d{4}年\d{1,2}月\d{1,2}日)', t)
    if not mm:
        mm = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日(?:至|到|起至|止于|一)(\d{4})年(\d{1,2})月(\d{1,2})日', t)
        if mm:
            out['起止日'] = f'{mm.group(1)}年{mm.group(2)}月{mm.group(3)}日~{mm.group(4)}年{mm.group(5)}月{mm.group(6)}日'
    else:
        out['起止日'] = f'{mm.group(1)}~{mm.group(2)}'
    return out


def _locate_ocr_root(contract_dir):
    """从合同目录向上定位「借款合同_ocr/识别文本_脱敏」，找不到返回 None。"""
    cur = os.path.abspath(contract_dir)
    for _ in range(5):
        cand = os.path.join(cur, '借款合同_ocr', '识别文本_脱敏')
        if os.path.isdir(cand):
            return cand
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return None


def main(contract_dir, out_path, skip_ocr=True):
    pdfs = sorted(glob.glob(os.path.join(contract_dir, '**', '*.pdf'), recursive=True))
    print(f'PDF 总数: {len(pdfs)}')
    rows = []
    for i, fp in enumerate(pdfs, 1):
        rel = os.path.relpath(fp, contract_dir)
        f = parse_filename_info(rel)
        b = BORROWER_MAP.get(f['borrower'], '借??')
        rec = OrderedDict()
        rec['借款方'] = b
        rec['贷款方'] = f['lender_code'] or ''
        rec['合同类型'] = f['ctype']
        rec['文件名金额(元)'] = f['amt']
        rec['内容金额(元)'] = None
        rec['担保金额(元)'] = None
        rec['抵押物价值(元)'] = None
        rec['利率(%)'] = None
        rec['期限'] = None
        rec['起止日'] = f['period']
        rec['合同编号'] = None
        rec['担保人'] = None
        rec['保证方式'] = None
        rec['抵押物/质押物'] = None
        # 文本/OCR内容：优先读本地脱敏版识别文本（contract_mask_local 已脱敏），否则 pymupdf 提取
        ocr_txt = None
        ocr_root = _locate_ocr_root(contract_dir)
        if ocr_root:
            try:
                import contract_mask_local as CML
                ocr_rel = CML.mask_path(rel + '.txt')
            except Exception:
                ocr_rel = rel + '.txt'
            ocr_fp = os.path.join(ocr_root, ocr_rel)
            if os.path.exists(ocr_fp):
                ocr_txt = open(ocr_fp, encoding='utf-8').read()
        if os.path.exists(ocr_fp):
            ocr_txt = open(ocr_fp, encoding='utf-8').read()
        if not ocr_txt or len(ocr_txt.strip()) < 30:
            try:
                import pymupdf as fitz
                doc = fitz.open(fp)
                ocr_txt = ''.join(p.get_text() for p in doc[:20])
                doc.close()
            except Exception:
                ocr_txt = ''
        if ocr_txt and len(ocr_txt.strip()) > 80:
            c = extract_content(rel, ocr_txt)
            if '合同类型' in c: rec['合同类型'] = c['合同类型']
            if '合同编号' in c: rec['合同编号'] = c['合同编号']
            if '金额(元)' in c: rec['内容金额(元)'] = c['金额(元)']
            if '担保金额(元)' in c: rec['担保金额(元)'] = c['担保金额(元)']
            if '抵押物价值(元)' in c: rec['抵押物价值(元)'] = c['抵押物价值(元)']
            if '利率(%)' in c: rec['利率(%)'] = c['利率(%)']
            if '期限' in c: rec['期限'] = c['期限']
            if '起止日' in c: rec['起止日'] = c['起止日']
            if '保证方式' in c: rec['保证方式'] = c['保证方式']
            if '抵押物' in c: rec['抵押物/质押物'] = c['抵押物']
            if '质押物' in c: rec['抵押物/质押物'] = c['质押物']
            if '担保人' in c:
                # 脱敏：担保人/抵押人名称 → 代号或"外部主体"
                g = c['担保人']
                mapped = None
                for name, code in BORROWER_MAP.items():
                    if name in g or g in name:
                        mapped = code
                        break
                if mapped:
                    rec['担保人'] = mapped + '(自担保方)' if mapped == rec['借款方'] else mapped
                else:
                    rec['担保人'] = '外部主体'
            rec['_src'] = '文本'
        else:
            rec['_src'] = '待识别'
        rows.append(rec)
        print(f'  [{i}/{len(pdfs)}] {rec["借款方"]}/{rec["贷款方"]} {rec["合同类型"]} '
              f'文件金额={rec["文件名金额(元)"]} 内容金额={rec["内容金额(元)"]} 担保金额={rec["担保金额(元)"]} '
              f'利率={rec["利率(%)"]} 担保人={rec["担保人"]} 保证方式={rec["保证方式"]} {rec["起止日"]} [{rec["_src"]}]', flush=True)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb = openpyxl.Workbook(); ws = wb.active
    ws.title = '借款合同台账(脱敏)'
    F = Font(size=10); FH = Font(size=10, bold=True)
    FILL = PatternFill('solid', fgColor='DDEBF7')
    THIN = Border(*[Side(style='thin')] * 4)
    headers = ['借款方', '贷款方', '合同类型', '合同编号',
               '文件名金额(元)', '内容金额(元)', '担保金额(元)', '抵押物价值(元)',
               '利率(%)', '期限', '起止日', '担保人', '保证方式', '抵押物/质押物', '信息来源']
    ws.append(headers)
    for c_ in ws[1]:
        c_.font = FH; c_.fill = FILL; c_.border = THIN
    for r in rows:
        ws.append([r['借款方'], r['贷款方'], r['合同类型'], r['合同编号'],
                   r['文件名金额(元)'], r['内容金额(元)'], r['担保金额(元)'], r['抵押物价值(元)'],
                   r['利率(%)'], r['期限'], r['起止日'], r['担保人'], r['保证方式'],
                   r['抵押物/质押物'], r['_src']])
        for c_ in ws[ws.max_row]:
            c_.font = F; c_.border = THIN
            if isinstance(c_.value, (int, float)) and c_.column in (5, 6, 7, 8, 9):
                c_.number_format = '#,##0.00'
    for col, w in zip('ABCDEFGHIJKLMNO', [9, 9, 8, 20, 14, 14, 14, 14, 8, 8, 30, 9, 12, 22, 8]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:O{ws.max_row}'
    wb.save(out_path)
    print(f'\n台账已输出: {out_path} ({len(rows)} 份)')
    # 汇总
    from collections import Counter
    print('合同类型分布:', dict(Counter(r['合同类型'] for r in rows)))
    print('信息来源:', dict(Counter(r['_src'] for r in rows)))


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('contract_dir')
    ap.add_argument('--out', default='借款合同台账_脱敏.xlsx')
    a = ap.parse_args()
    main(a.contract_dir, a.out)
