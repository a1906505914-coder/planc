# -*- coding: utf-8 -*-
"""contract_detail_ledger.py —— 表外事项明细底稿生成器（脱敏版输入）
从 借款合同_ocr/识别文本_脱敏 提取合同明细，分四类输出：
  Sheet1 对外担保（保证合同）：担保方/被担保方/债权人/担保金额/保证方式/主债权期间
  Sheet2 抵质押（抵押/质押）：抵押人/债务人/抵押权人/抵押物/金额/期间
  Sheet3 委托贷款（委贷借据）：委托人/借款人/金额/利率/起息/到期/用途
  Sheet4 借款合同：借款人/贷款人/金额/利率/期限/担保方式
同时输出合同信息 JSON（供 loan_detail 借款底稿补充利率/担保列）。
只读脱敏版，输出全部为 借X/贷X 代码。
"""
import paths as P
import os, re, sys, glob, json
from collections import OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from loan_contract_reader import cn_amount_to_num, BORROWER_MAP, LENDER_MAP

OCR_ROOT = os.path.join(P.DATA_DIRS['ADF'], '数据', '2026', '借款合同_ocr', '识别文本_脱敏')
OUT_XLSX = os.path.join(P.DATA_DIRS['ADF'], '数据', '2026', '表外事项明细底稿.xlsx')
OUT_JSON = os.path.join(P.DATA_DIRS['ADF'], '数据', '2026', '合同信息_脱敏.json')

BORROWER_DIR = {'借01', '借02', '借03', '借04', '借05', '借06'}


def _ocr_fix(s):
    if not s:
        return ''
    s = s.replace('吉', '壹').replace('从民币', '人民币').replace('从民市', '人民币')
    s = s.replace('Q', '0').replace('O', '0')
    s = s.replace('木', '不').replace('风', '，').replace('口', '（')
    s = s.replace('任', '仟').replace('侣', '佰').replace('位', '亿')
    s = s.replace('红苏', '江苏').replace('江办', '江苏')
    return s


def classify(fname, rel, text):
    """合同类型：文件名优先，文本标题补充。"""
    full = rel.replace(os.sep, ' ') + ' ' + fname
    t = _ocr_fix(text or '')
    # ① 委贷/租赁（文件名强信号，优先）
    if re.search(r'委贷总协议|委托贷款总协议|委托贷款业务总协议|业务合作总协议', full): return '委贷协议'
    if re.search(r'委贷|委托借据|电子委贷借据|电子委托借据|借据', full): return '委贷借据'
    if re.search(r'租赁协议|融资租赁|金租', full): return '租赁'
    # ② 文件名精确类型词（目录名/文件名）
    if re.search(r'保字|（保）|最高额保证|保证合同|BZ\d', full): return '保证'
    if re.search(r'押字|（抵）|最高额抵押|抵押合同|抵字', full): return '抵押'
    if re.search(r'质字|（质）|最高额质押|质押合同|质字', full): return '质押'
    if re.search(r'借字|（借）|借款合同|流贷|LDZJ|流动资金|JK\d', full): return '借款'
    # ③ 文本标题（前 4000 字，覆盖第1-2页合同标题）
    title = t[:4000]
    if re.search(r'最高额保证合同|保证合同', title): return '保证'
    if re.search(r'最高额抵押合同|抵押合同', title): return '抵押'
    if re.search(r'最高额质押合同|质押合同', title): return '质押'
    if re.search(r'流动资金借款合同|借款合同|流贷', title): return '借款'
    return '其他'


def _borrower_of(rel):
    """路径第一层 → 借X"""
    p0 = rel.split(os.sep)[0]
    return p0 if p0 in BORROWER_DIR else '借??'


def _map_entity(name):
    """文本中的主体名 → 借X/贷X/外部主体。含 借/贷+数字 → 对应代码；含公司/银行等主体特征 → 外部主体（脱敏泛化）。"""
    if not name:
        return None
    m = re.search(r'(借\d{2}|贷\d{2})', name)
    if m:
        return m.group(1)
    # 外部主体泛化（防真实名残留）
    if re.search(r'(有限公司|有限责任公司|股份有限|集团公司|公司|集团|银行|厂|事务所|'
                 r'合伙企业|分公司|总厂|支行|营业部)', name):
        return '外部主体'
    return None


def _sanitize_amt(amt, t):
    """金额可信度防护：大写缺失 + 数值非整万/整亿 → 可疑（OCR 错位），返回 None。"""
    if not amt:
        return None
    if amt > 100000000:
        has_big = bool(re.search(r'[（(]?大写[）)]?\s*[（(]?[零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分壹贰壹]{5,}整?', t))
        if not has_big and amt % 1000000 != 0 and amt % 10000 != 0:
            return None
    return round(amt, 2)


def _cn_amt_ctx(t):
    mm = re.search(r'[（(]?小写[）)]?\s*[¥￥]?\s*([0-9][0-9,，.]{3,})', t)
    if mm:
        try:
            return float(mm.group(1).replace(',', '').replace('，', ''))
        except ValueError:
            return None
    return None


def _extract_big_amt(t, kws):
    """从文本找大写金额（OCR 容错，支持"（币种及大写金额）为人民币X"格式）。"""
    # 全文大写金额模式（不限关键词位置）
    pats = [
        r'（币种及大写金额）\s*为\s*人民币?\s*[（(]?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分壹贰壹]{5,}整?)[）)]?',
        r'（大写）\s*[（(]?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分壹贰壹]{5,}整?)[）)]?',
        r'[（(]?大写[）)]?\s*[（(]?人民币?[）)]?\s*[（(]?([零壹贰叁肆伍陆柒捌玖拾佰仟万亿元角分壹贰壹]{5,}整?)[）)]?',
    ]
    # 优先在关键词后 300 字内找
    for kw in kws:
        idx = t.find(kw)
        if idx < 0:
            continue
        seg = t[idx:idx + 300]
        for pat in pats:
            mm = re.search(pat, seg)
            if mm:
                a = cn_amount_to_num(_ocr_fix(mm.group(1)))
                if a:
                    return a
    # 全文兜底
    for pat in pats:
        mm = re.search(pat, t)
        if mm:
            a = cn_amount_to_num(_ocr_fix(mm.group(1)))
            if a:
                return a
    return None


def parse_guarantee(rel, text, fname):
    """保证合同 → 对外担保明细"""
    t = _ocr_fix(text or '')
    out = {'类型': '对外担保(保证)', '合同编号': None, '担保方': None, '被担保方': None,
              '债权人': None, '担保金额(元)': None, '保证方式': None, '主债权期间': None, '来源': rel}
    mm = re.search(r'编号[：:]\s*([^\s]{4,40})', t)
    if mm:
        out['合同编号'] = mm.group(1).strip()
    # 保证人（担保方）
    for kw in ('保证人', '担保人'):
        mm = re.search(rf'{kw}\s*[：:]\s*([^\n]{{2,30}}?)\s*(?:统一社会信用代码|法定代表人|负责人|住所地|电话|传真|$)', t)
        if mm:
            out['担保方'] = _map_entity(mm.group(1).strip())
            break
    # 债务人（被担保方）：文本"债权人与债务人XXX之间"
    mm = re.search(r'债务人\s*[（(]?[：:]?\s*([\u4e00-\u9fffA-Za-z\d借贷]{2,30}?)\s*(?:之间|_之间|，|$)', t)
    if mm:
        ent = _map_entity(mm.group(1).strip())
        if ent:
            out['被担保方'] = ent
    # 债权人
    mm = re.search(r'债权人\s*[：:]\s*([^\n]{2,30}?)\s*(?:法定代表人|负责人|住所地|电话|传真|$)', t)
    if mm:
        out['债权人'] = _map_entity(mm.group(1).strip())
    # 担保金额
    amt = _extract_big_amt(t, ['被担保最高债权额', '被担保的主债权', '最高本金余额'])
    if not amt:
        amt = _cn_amt_ctx(t)
    amt = _sanitize_amt(amt, t)
    if amt:
        out['担保金额(元)'] = round(amt, 2)
    else:
        out['担保金额(元)'] = '见原函'
    # 保证方式
    mm = re.search(r'保证方式\s*为下列第\s*(\d)\s*项', t)
    if mm:
        out['保证方式'] = '连带责任保证' if mm.group(1) == '1' else '一般保证'
    else:
        mm = re.search(r'保证方式[^。]{0,40}(连带(?:责仟|责任|货任)保证|一般保证)', t, re.S)
        if mm:
            out['保证方式'] = '连带责任保证' if '连带' in mm.group(1) else '一般保证'
    # 主债权期间（自X至X / 文件名日期）
    mm = re.search(r'自\s*(\d{4}年?\d{1,2}月?\d{1,2}日?|[一二三四五六七八九十]{1,3}年[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,3}日)'
                   r'\s*至\s*([\d一二三四五六七八九十]{4}年?[\d一二三四五六七八九十]{1,2}月?[\d一二三四五六七八九十]{1,2}日?)', t)
    if mm:
        out['主债权期间'] = f'{mm.group(1)}~{mm.group(2)}'
    else:
        mm = re.search(r'(\d{4}[./]\d{1,2}[./]\d{1,2})\s*[-~至]\s*(\d{4}[./]\d{1,2}[./]\d{1,2})', fname)
        if mm:
            out['主债权期间'] = f'{mm.group(1)}~{mm.group(2)}'
    return out


def parse_mortgage(rel, text, fname):
    """抵押/质押合同 → 抵质押明细"""
    kind = '质押' if '质' in fname else '抵押'
    t = _ocr_fix(text or '')
    out = {'类型': kind, '合同编号': None, '抵押(出质)人': None, '债务人': None,
              '抵押权(质权)人': None, '抵押(质)物': None, '金额(元)': None, '期间': None, '来源': rel}
    mm = re.search(r'编号[：:]\s*([^\s]{4,40})', t)
    if mm:
        out['合同编号'] = mm.group(1).strip()
    kw = '抵押人' if kind == '抵押' else '出质人'
    mm = re.search(rf'{kw}\s*[：:]\s*([^\n]{{2,30}}?)\s*(?:统一社会信用代码|法定代表人|负责人|住所地|电话|传真|$)', t)
    if mm:
        out['抵押(出质)人'] = _map_entity(mm.group(1).strip())
    mm = re.search(r'债务人\s*[（(]?名称?[）)]?[：:]?\s*([^\n]{2,30}?)(?:\s*之间|，|$)', t)
    if mm:
        out['债务人'] = _map_entity(mm.group(1).strip())
    kw2 = '抵押权人' if kind == '抵押' else '质权人'
    mm = re.search(rf'{kw2}\s*[：:]\s*([^\n]{{2,30}}?)\s*(?:法定代表人|负责人|住所地|电话|传真|$)', t)
    if mm:
        out['抵押权(质权)人'] = _map_entity(mm.group(1).strip())
    # 抵押物：文件名（借02 描述性文件）优先
    mm = re.search(r'(土地房产|机器设备|房地产|厂房|设备|土地使用权|房产|股权|应收账款|存货)\S{0,15}', fname)
    if mm:
        out['抵押(质)物'] = mm.group(0).strip()
    else:
        mm = re.search(r'抵押物\s*[：:]?\s*([\u4e00-\u9fffA-Za-z0-9（）()]{2,40}?)(?:。|；|，|$)', t)
        v = mm.group(1).strip() if mm else None
        if v:
            # 排除条款性描述（"如抵押物为房屋的"/"抵押物毁损"等）
            if re.match(r'^(如|若|为|毁|灭|被|已|未|其)', v) or '抵押物' in v:
                v = None
            out['抵押(质)物'] = v
        else:
            out['抵押(质)物'] = '见抵押物清单'
    amt = _extract_big_amt(t, ['被担保最高债权额', '最高本金余额'])
    if not amt:
        mm = re.search(r'(\d+(?:\.\d+)?)\s*(亿|万)', fname)
        if mm:
            amt = float(mm.group(1)) * (1e8 if mm.group(2) == '亿' else 1e4)
    if not amt:
        amt = _cn_amt_ctx(t)
    amt = _sanitize_amt(amt, t)
    out['金额(元)'] = round(amt, 2) if amt else '见原函'
    mm = re.search(r'(\d{4}[./]\d{1,2}[./]\d{1,2})\s*[-~至]\s*(\d{4}[./]\d{1,2}[./]\d{1,2})', fname)
    if mm:
        out['期间'] = f'{mm.group(1)}~{mm.group(2)}'
    return out


def parse_ledger(rel, text, fname):
    """委贷借据 → 委托贷款明细"""
    t = _ocr_fix(text or '')
    out = {'类型': '委托贷款(借据)', '借据编号': None, '协议编号': None, '委托人': None,
              '借款人': None, '金额(元)': None, '利率(%)': None, '起息日': None, '到期日': None,
              '用途': None, '来源': rel}
    # 借据编号：借据编号标签后数字；或标题行"电子委贷借据\n004"后的独立编号
    mm = re.search(r'借据编号[：:]\s*([A-Za-z0-9\-]{2,20})', t)
    if mm and not re.match(r'^\d{4}-\d{2}-\d{2}', mm.group(1)) \
            and '打印日期' not in mm.group(1):
        out['借据编号'] = mm.group(1).strip()
    if not out['借据编号']:
        mm = re.search(r'(?:电子委贷借据|电子委托借据|借据)\s*\n\s*(\d{2,4})\s*$', t, re.M)
        if mm:
            out['借据编号'] = mm.group(1)
    # 协议编号：8-15 位数字（排除 16+ 位账号、排除利率小数）
    mm = re.search(r'(?<!\d)(\d{8,15})(?!\d)', t)
    if mm:
        out['协议编号'] = mm.group(1)
    # 借款人恒为本主体（委贷借据借款人=路径第一层借X），委托人为文本中另一个借X
    out['借款人'] = _borrower_of(rel)
    others = re.findall(r'(借\d{2})', t)
    for ent in others:
        if ent != out['借款人']:
            out['委托人'] = ent
            break
    # 金额：小写金额 RMB
    mm = re.search(r'小写金额[：:]?\s*RMB\s*([\d][\d,，.]{3,})', t)
    if not mm:
        mm = re.search(r'([\d][\d,，.]{3,}\.\d{2})', t)
    if mm:
        try:
            out['金额(元)'] = round(float(mm.group(1).replace(',', '').replace('，', '')), 2)
        except ValueError:
            pass
    # 利率：数字 + (%) 标签（前后都试，含全角括号）；数值须在 0-100
    rate_v = None
    for pat in (r'(\d+(?:\.\d+)?)\s*\n?\s*[（(]%[）)]',
                r'[（(]%[）)]\s*\n?\s*(\d+(?:\.\d+)?)',
                r'(\d+\.\d{2,7})'):
        for m in re.finditer(pat, t):
            v = float(m.group(1))
            if 0 < v < 100:
                rate_v = v
                break
        if rate_v is not None:
            break
    if rate_v is not None:
        out['利率(%)'] = rate_v
    # 起息/到期：以"起息/到期"标签定位，各取其后的日期；若同取下一个
    def _date_after(keyword, skip=None):
        m = re.search(keyword + r'[^0-9]{0,20}?(\d{4}-\d{2}-\d{2})', t)
        if not m:
            return None
        d = m.group(1)
        if d == skip:
            # 找该标签后的下一个不同日期
            m2 = re.search(keyword + r'[^0-9]{0,20}?' + d + r'[^0-9]{0,20}?(\d{4}-\d{2}-\d{2})', t)
            if m2:
                return m2.group(1)
        return d
    qd = _date_after(r'起息') or _date_after(r'起\n息')
    dd = _date_after(r'到期', skip=qd)
    if qd:
        out['起息日'] = qd
    if dd:
        out['到期日'] = dd
    mm = re.search(r'用途\s*([^\n]{2,20}?)$', t, re.M)
    if mm:
        out['用途'] = mm.group(1).strip()
    return out


def parse_loan(rel, text, fname):
    """借款/流贷合同 → 借款合同明细"""
    t = _ocr_fix(text or '')
    out = {'类型': '借款合同', '合同编号': None, '借款人': None, '贷款人': None,
              '金额(元)': None, '利率(%)': None, '期限': None, '起止日': None, '担保方式': None, '来源': rel}
    mm = re.search(r'编号[：:]\s*([^\s]{4,40})', t)
    if mm:
        out['合同编号'] = mm.group(1).strip()
    mm = re.search(r'借款人[：:]\s*([^\n]{2,30}?)\s*(?:统一社会信用代码|法定代表人|负责人|住所地|电话|传真|$)', t)
    if mm:
        out['借款人'] = _map_entity(mm.group(1).strip())
    mm = re.search(r'贷款人[：:]\s*([^\n]{2,30}?)\s*(?:法定代表人|负责人|住所地|电话|传真|$)', t)
    if mm:
        out['贷款人'] = _map_entity(mm.group(1).strip())
    amt = _extract_big_amt(t, ['借款金额', '借款额度金额'])
    if not amt:
        amt = _cn_amt_ctx(t)
    if amt:
        out['金额(元)'] = round(amt, 2)
    rate = None
    mm = re.search(r'(?:固定利率，)?(?:年利率|执行利率)\s*[（(]?[：:]?\s*([0-9]+(?:\.[0-9]+)?)\s*%', t)
    if not mm:
        mm = re.search(r'LPR\s*[+-]\s*([0-9]+(?:\.[0-9]+)?)\s*%', t)
    if mm:
        rate = float(mm.group(1))
    if rate is not None:
        out['利率(%)'] = rate
    mm = re.search(r'借款期限\s*[：:]?\s*(\d+)\s*个月', t)
    if not mm:
        mm = re.search(r'期限\s*[：:]?\s*(\d+)\s*个月', t)
    if mm:
        out['期限'] = f'{mm.group(1)}个月'
    mm = re.search(r'自\s*(\d{4}年\d{1,2}月\d{1,2}日)\s*至\s*(\d{4}年\d{1,2}月\d{1,2}日)', t)
    if mm:
        out['起止日'] = f'{mm.group(1)}~{mm.group(2)}'
    # 担保方式（本合同的担保安排）
    if re.search(r'保证人[：:]', t) or re.search(r'连带责任保证|保证方式', t):
        out['担保方式'] = '保证'
    elif re.search(r'抵押人[：:]|最高额抵押', t):
        out['担保方式'] = '抵押'
    elif re.search(r'出质人[：:]|质押', t):
        out['担保方式'] = '质押'
    elif re.search(r'信用', t):
        out['担保方式'] = '信用'
    return out


def parse_one(rel, text, fname):
    ctype = classify(fname, rel, text)
    if ctype == '保证':
        return parse_guarantee(rel, text, fname)
    if ctype in ('抵押', '质押'):
        return parse_mortgage(rel, text, fname)
    if ctype == '委贷借据':
        return parse_ledger(rel, text, fname)
    if ctype == '委贷协议':
        return {'类型': '委贷协议', '合同编号': None, '委托人': None, '借款人': None,
                '金额(元)': None, '利率(%)': None, '来源': rel}
    if ctype == '借款':
        return parse_loan(rel, text, fname)
    return {'类型': '其他', '来源': rel}


def main():
    files = sorted(glob.glob(os.path.join(OCR_ROOT, '**', '*.txt'), recursive=True))
    print(f'扫描合同文本: {len(files)} 份')
    rows = []
    for fp in files:
        rel = os.path.relpath(fp, OCR_ROOT)
        fname = os.path.basename(rel)
        with open(fp, encoding='utf-8', errors='replace') as f:
            text = f.read()
        rec = parse_one(rel, text, fname)
        rec['借款主体'] = _borrower_of(rel)
        rows.append(rec)
        print(f'  [{rec["类型"]}] {rec["借款主体"]} {rel}', flush=True)

    # 保存 JSON（供 loan_detail 补充用）
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(rows, f, ensure_ascii=False, indent=1)
    print(f'合同信息 JSON → {OUT_JSON}')

    # 分组输出 Excel
    wb = openpyxl.Workbook()
    F = Font(size=10); FH = Font(size=10, bold=True)
    FILL = PatternFill('solid', fgColor='DDEBF7')
    THIN = Border(*[Side(style='thin')] * 4)
    groups = {'对外担保': ('对外担保(保证)', ['类型', '借款主体', '合同编号', '担保方', '被担保方', '债权人',
                                              '担保金额(元)', '保证方式', '主债权期间', '来源']),
              '抵质押': ('抵押', ['类型', '借款主体', '合同编号', '抵押(出质)人', '债务人', '抵押权(质权)人',
                                     '抵押(质)物', '金额(元)', '期间', '来源']),
              '委托贷款': ('委托贷款(借据)', ['类型', '借款主体', '借据编号', '协议编号', '委托人', '借款人',
                                                '金额(元)', '利率(%)', '起息日', '到期日', '用途', '来源']),
              '借款合同': ('借款合同', ['类型', '借款主体', '合同编号', '借款人', '贷款人', '金额(元)',
                                            '利率(%)', '期限', '起止日', '担保方式', '来源'])}
    first = True
    for sheet_title, (tkey, cols) in groups.items():
        ws = wb.active if first else wb.create_sheet()
        first = False
        ws.title = sheet_title
        ws.append(cols)
        for c_ in ws[1]:
            c_.font = FH; c_.fill = FILL; c_.border = THIN
        for r in rows:
            if r.get('类型') == tkey or (tkey == '抵押' and r.get('类型') in ('抵押', '质押')):
                ws.append([r.get(k) for k in cols])
                for c_ in ws[ws.max_row]:
                    c_.font = F; c_.border = THIN
                    if isinstance(c_.value, (int, float)):
                        c_.number_format = '#,##0.00'
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = f'A1:{chr(64 + len(cols))}{ws.max_row}'
        for col, w in zip('ABCDEFGHIJKL', [12, 8, 20, 10, 10, 10, 14, 12, 24, 10, 16, 40]):
            ws.column_dimensions[col].width = w
    wb.save(OUT_XLSX)
    print(f'表外事项明细底稿 → {OUT_XLSX}')
    from collections import Counter
    print('合同类型分布:', dict(Counter(r.get('类型') for r in rows)))


if __name__ == '__main__':
    main()
