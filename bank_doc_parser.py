# -*- coding: utf-8 -*-
"""银行文档解析器（2026-08-16，铁律131f-ocr 扩展）：
从 OCR 识别文本中抽取结构化信息——为函证回函/银行清单/对账单识别打基础。

parse_bank_list(text)  → 已开立银行结算账户清单
  {'depositor': 存款人名称, 'accounts': [{bank, account, acct_type, approval, open_date, status, close_date}]}
parse_bank_statement(text) → 银行对账单/月结单
  {'bank', 'account', 'currency', 'holder', 'period', 'open_balance', 'close_balance', 'txns': [{date, debit, credit, balance, counter, summary}]}
"""
import paths   # 数据路径中心锚点（数据绝对路径只允许出现在 paths.py）
import re
from mask_dict import mask_name   # ⚡ console 输出主体/户名脱敏（2026-08-22 复扫）
from mask_engine import _mask_num as _msk   # ⚡ 账号/编号掩码（S2）

_BANK_KW = ['中国工商银行', '中国农业银行', '中国银行', '中国建设银行', '交通银行',
            '中国邮政储蓄银行', '招商银行', '浦发银行', '中信银行', '兴业银行', '民生银行',
            '光大银行', '华夏银行', '广发银行', '平安银行', '浙商银行', '恒丰银行',
            '北京银行', '上海银行', '江苏银行', '宁波银行', '杭州银行', '渤海银行']
_ACCT_TYPE = ['基本存款账户', '一般存款账户', '专用存款账户', '临时存款账户']


def _find_bank(text):
    """文本段 → 标准银行名（排除标题/表头词）。"""
    for bad in ('已开立', '单位存款人', '开户银行名称', '银行结算账户清单', '银行账户清单'):
        text = text.replace(bad, '')
    for b in _BANK_KW:
        if b in text:
            return b
    for p in _BANK_PREFIX:
        if p in text:
            return p
    m = re.search(r'([\u4e00-\u9fa5]{2,12}银行[^\s\n]{0,8})', text)
    return m.group(1) if m else ''


_BANK_PREFIX = ('中国工商', '中国农业', '中国建设', '中国银行', '交通银行', '中国邮政', '招商银行',
                 '浦发银行', '中信银行', '兴业银行', '民生银行', '光大银行', '华夏银行', '广发银行',
                 '平安银行', '浙商银行', '恒丰银行', '北京银行', '上海银行', '江苏银行', '宁波银行',
                 '杭州银行', '渤海银行', '徽商银行', '广州银行')


def _is_bank_line(line):
    """行是否为开户银行行（含标准行名/银行名前缀/分行支行后缀）。
    排除：标题/表头/提示文字（'已开立银行结算账户清单'/'开户银行名称'/'填写所有'/'人民银行'/'单位存款人'等）。"""
    for bad in ('已开立银行结算账户清单', '开立银行账户清单', '开户银行名称', '填写所有',
                '人民银行', '银行结算账户', '单位存款人', '序开户银行', '开户日期', '账户状态',
                '存款人名称', '核准号', '账户性质', '销户日', '久悬日'):
        if bad in line:
            return False
    for b in _BANK_KW:
        if b in line:
            return True
    for p in _BANK_PREFIX:
        if p in line:
            return True
    return bool(re.search(r'[\u4e00-\u9fa5]{2,12}银行(?:股份)?(?:有限公司)?(?:分行|支行|营业部|办事处|网点)', line))


def _seg_account(seg, approval=''):
    """从单个账户文本段提取 {bank, account, acct_type, approval, open_date, status}。"""
    acct = {'bank': _find_bank(seg), 'account': '', 'acct_type': '',
            'approval': approval, 'open_date': '', 'status': '', 'close_date': ''}
    # 账号：10-20 位数字（排除核准号主体/年份）
    m = re.search(r'(?<!\d)\d{10,20}(?!\d)', seg.replace(' ', ''))
    if m:
        acct['account'] = m.group(0)
    # 账户类型（首个命中）
    for t in _ACCT_TYPE:
        if t in seg:
            acct['acct_type'] = t
            break
    # 开户日期
    m = re.search(r'(20\d{2})[年\-/](\d{1,2})[月\-/](\d{1,2})', seg)
    if m:
        acct['open_date'] = f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    # 状态
    for s in ['正常', '久悬', '销户']:
        if s in seg:
            acct['status'] = s
            break
    # 销户日期
    m = re.search(r'销户[日期：:]*\s*(20\d{2})[年\-/](\d{1,2})[月\-/](\d{1,2})', seg)
    if m:
        acct['close_date'] = f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    return acct


def parse_bank_list(text):
    """已开立银行结算账户清单 → 结构化（⚡ 逐账户解析：N 个账户 → N 行）。"""
    out = {'depositor': '', 'accounts': []}
    # 存款人名称：关键词后首行【公司主体词】行（排除 银行/基本/核准/开户 等干扰词）
    _stop = ('银行', '基本存款', '一般存款', '专用存款', '临时存款', '核准', '开户', '账户', '单位存款', '中国')
    m = re.search(r'存款人名称\s*[：:（(]?\s*([^\n]{2,40})', text)
    if m:
        cand = m.group(1).strip()
        if not any(s in cand for s in _stop):
            out['depositor'] = cand
    if not out['depositor']:
        # 回退：选含主体词的【最长完整候选】（'江苏富丽华通用设备股份有限公司' 优于残缺'行股份有限'）；
        # 排除含长数字的行（账号/核准号混入，如'股份有限510019816800'）
        best = ''
        for line in text.splitlines():
            line = line.strip()
            if 3 <= len(line) <= 50 and ('有限公司' in line or '股份' in line or '厂' in line
                                         or '中心' in line or '事务所' in line or '集团' in line):
                if any(s in line for s in _stop):
                    continue
                if re.search(r'\d{8,}', line):
                    continue
                if len(line) > len(best):
                    best = line
        if best:
            out['depositor'] = best
    # ⚡⚡ 逐账户解析：按银行行分段（银行清单常含多家银行多个账户）
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    bank_idx = [i for i, l in enumerate(lines) if _is_bank_line(l)]
    approvals = re.findall(r'[JQ][A-Z]?\d{10,14}', text)
    if bank_idx:
        for k, bi in enumerate(bank_idx):
            end = bank_idx[k + 1] if k + 1 < len(bank_idx) else len(lines)
            seg = '\n'.join(lines[bi:end])
            a = _seg_account(seg, approvals[k] if k < len(approvals) else '')
            if a['account'] or a['bank']:
                out['accounts'].append(a)
    if not out['accounts']:
        # 回退：单账户整文提取
        a = _seg_account(text, approvals[0] if approvals else '')
        if a['account'] or a['bank']:
            out['accounts'].append(a)
    return out


def parse_bank_statement(text):
    """银行对账单/月结单（含汇丰结单）→ 结构化（户名/账号/币种/期间/承前余额/交易明细）。"""
    out = {'bank': _find_bank(text), 'holder': '', 'account': '', 'currency': '',
           'period': '', 'open_balance': None, 'close_balance': None, 'txns': []}
    # 户名/存款人（含繁体"户名"）
    m = re.search(r'(?:户名|客戶名称|客户名称)[：:\s]*([^\n]{2,40})', text)
    if m:
        out['holder'] = m.group(1).strip()
    # 账号：数字-数字（汇丰 817-839848-838）或纯数字
    m = re.search(r'(?:账户号码|账号|帳戶號碼)[：:\s]*([0-9A-Z\-]{6,30})', text)
    if m:
        out['account'] = m.group(1).strip()
    if not out['account']:
        m = re.search(r'\b\d{3,4}-\d{6,12}(?:-\d{2,4})?\b', text)
        if m:
            out['account'] = m.group(0)
    if not out['account']:
        m = re.search(r'(?<!\d)\d{10,20}(?!\d)', text)
        if m:
            out['account'] = m.group(0)
    # 币种（中英）
    m = re.search(r'(?:币种|幣種|Currency)[：:\s]*([A-Z]{3})', text)
    if m:
        out['currency'] = m.group(1)
    if not out['currency']:
        m = re.search(r'\b(USD|HKD|RMB|CNY|EUR|GBP|JPY)\b', text)
        if m:
            out['currency'] = m.group(1)
    # 期间：2025年01月 / January 2025
    m = re.search(r'(20\d{2})年(\d{1,2})月', text)
    if m:
        out['period'] = f'{m.group(1)}-{int(m.group(2)):02d}'
    else:
        m = re.search(r'([A-Z][a-z]{2,8})\s+(20\d{2})', text)
        if m:
            _EN_MON = {'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05',
                       'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
                       'November': '11', 'December': '12'}
            out['period'] = f'{m.group(2)}-{_EN_MON.get(m.group(1), "??")}'
    # 承前余额（中文/英文 B/F）
    m = re.search(r'(?:承前余额|承前轉結|B/F\s*BALANCE)[：:\s]*([\d,，.]+)', text)
    if m:
        out['open_balance'] = float(m.group(1).replace(',', '').replace('，', ''))
    # 交易行：中文日期 20250103 / 英文 30 Dec / 30 December
    pat = re.compile(r'(20\d{2})(\d{2})(\d{2})\s*([+\-]?[\d,，.]+)\s*([\d,，.]+)')
    for m in pat.finditer(text):
        date = f'{m.group(1)}-{m.group(2)}-{m.group(3)}'
        amt = float(m.group(4).replace(',', '').replace('，', ''))
        bal = float(m.group(5).replace(',', '').replace('，', ''))
        out['txns'].append({'date': date, 'amt': amt, 'balance': bal})
    if out['txns']:
        out['close_balance'] = out['txns'][-1]['balance']
    return out


def parse(text, kind='auto'):
    """自动路由：银行清单 or 对账单/结单。"""
    if '已开立银行结算账户清单' in text or '开立银行账户清单' in text:
        return parse_bank_list(text)
    if ('月结单' in text or '对账单' in text or '承前余额' in text or '承前轉結' in text
            or '结单' in text or 'Statement' in text or 'B/F' in text):
        return parse_bank_statement(text)
    return {'raw_len': len(text), 'hint': '未识别文档类型'}


if __name__ == '__main__':
    import sys, glob, os, json
    d = sys.argv[1] if len(sys.argv) > 1 else os.path.join(paths.DATA_ROOT, '银行资料', '_ocr_识别结果')
    for f in sorted(glob.glob(os.path.join(d, '*.txt'))):
        try:
            text = open(f, encoding='utf-8').read()
        except Exception:
            continue
        r = parse(text)
        name = os.path.basename(f)
        print(f'=== {name} ===')
        if 'accounts' in r:
            print(f'  存款人: {mask_name(r.get("depositor"))}')
            for a in r['accounts']:
                print(f'    银行={mask_name(a["bank"])} 账号={_msk(a["account"])} 类型={a["acct_type"]} 核准号={_msk(a["approval"])} 开户={a["open_date"]} 状态={a["status"]}')
        elif 'txns' in r:
            print(f'  {mask_name(r["bank"])} | 户名={mask_name(r["holder"])} 账号={_msk(r["account"])} 币种={r["currency"]} {r["period"]} 承前={r["open_balance"]} 期末={r["close_balance"]} 交易{r["txns"].__len__()}笔')
        else:
            print(f'  {r}')
