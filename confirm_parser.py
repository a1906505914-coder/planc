# -*- coding: utf-8 -*-
"""函证回函解析器（2026-08-16，铁律131f-ocr 扩展）。
从 OCR 识别文本抽取函证回函结构化信息——银行询证函回函 / 往来询证函回函。

parse_confirm(text) → {
  'type': '银行询证函回函' / '往来询证函回函',
  'ref_no': 函证编号, 'auditee': 被审计单位, 'to_auditor': 致会计师事务所,
  'conclusion': '相符' / '不相符' / '未标注',
  'discrepancy': 差异说明, 'amounts': [{(项目, 金额)}], 'stamp_unit': 回函盖章单位,
  'reply_date': 回函日期, 'matched_items': 相符项目数, 'mismatch_items': 不相符项目数,
}
"""
import re


def _find_date(text):
    """回函日期：优先'回函日期'显式字段；否则取文本【最后一个】日期（回函日期通常在末尾）。"""
    m = re.search(r'(?:回函日期|答复日期|复函日期)[：:\s]*(\d{4})年?\s*(\d{1,2})月\s*(\d{1,2})日?', text)
    if m:
        return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    dates = re.findall(r'(20\d{2})[年\-/](\d{1,2})[月\-/](\d{1,2})', text)
    if dates:
        y, mo, d = dates[-1]
        return f'{y}-{int(mo):02d}-{int(d):02d}'
    return ''


def _find_auditee(text):
    m = re.search(r'(?:被审计单位|贵单位|委托方)[：:\s]*([^\n]{2,50})', text)
    if m:
        return m.group(1).strip()
    return ''


def _find_ref(text):
    m = re.search(r'(?:编号|函证编号|询证函编号)[：:\s]*([^\n]{3,40})', text)
    if m:
        return m.group(1).strip()
    m = re.search(r'([A-Z]{0,5}\d{4}年[第]?\d{1,6}号)', text)
    return m.group(1) if m else ''


def _find_stamp(text):
    """回函盖章单位：'XX银行XX支行' 或 企业名 + 盖章/签章。"""
    for kw in ('回函单位', '回函机构', '银行签章', '经办行'):
        m = re.search(kw + r'[：:\s]*([^\n]{2,40})', text)
        if m:
            return m.group(1).strip()
    m = re.search(r'([\u4e00-\u9fa5]{2,20}(?:银行|有限公司|股份|事务所)[^\n]{0,12}?)\s*[（(]?\s*(?:盖章|签章|公章)', text)
    if m:
        return m.group(1).strip().rstrip('（(')
    return ''


def _find_conclusion(text):
    """相符/不相符结论（银行/企业回函核心）。

    ⚡ 优先：回函函证"以下由被询证单位填列"区的手写勾选只能视觉判定（OCR 文本无法识别打勾），
    因此结论判定的**权威源是文件名标记**（审计师/录入时已标注），文本关键词仅作辅助参考。
    """
    # 文本优先级：先"不相符"（含"存在以下不符之处"），再"相符"
    if re.search(r'不\s*相\s*符|信息不符|与函证内容不符|存在差异|存在以下不符之处', text):
        return '不相符'
    if re.search(r'相\s*符|核对无误|与函证内容一致|数据相符|信息一致', text):
        return '相符'
    return '未标注'


def _find_discrepancy(text):
    """往来函证差异说明：'存在以下不符之处'后的手写差异文字/金额。
    重点提取：金额、原因、日期等手写数字与文字（OCR 对手写识别率低，但印刷体差异说明可提取）。
    返回 {'discrepancy_text': str, 'discrepancy_amount': float, 'discrepancy_date': str}。
    """
    out = {'discrepancy_text': '', 'discrepancy_amount': 0.0, 'discrepancy_date': ''}
    m = re.search(r'存在以下不符之处[：:\s]*?([\s\S]{2,300}?)(?:签章|经办人|日期|第\s*\d+\s*页|\Z)', text)
    if not m:
        return out
    seg = m.group(1).strip()
    out['discrepancy_text'] = re.sub(r'\s+', ' ', seg)[:200]
    # 提取金额（手写常见格式：金额：X元 / 差X元 / 多X元 / 少X元 / 欠X元）
    m_amt = re.search(r'(?:金额|差额|差|多|少|欠)[：:\s]*([\d,]+\.\d{1,2})\s*元', seg) or \
            re.search(r'([\d,]+\.\d{1,2})\s*元[（(]?(?:截至|截|至).*?[）)]?', seg) or \
            re.search(r'(?:上数|本次|本次核对|差|多|少|欠)\s*([\d,]+\.\d{1,2})\s*元', seg)
    if m_amt:
        try:
            out['discrepancy_amount'] = float(m_amt.group(1).replace(',', ''))
        except ValueError:
            pass
    # 提取日期
    m_dt = re.search(r'(?:截至|截|至|日期)?\s*(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})日?', seg)
    if m_dt:
        out['discrepancy_date'] = f'{m_dt.group(1)}-{int(m_dt.group(2)):02d}-{int(m_dt.group(3)):02d}'
    return out


def extract_conclusion_region(pdf_path, out_png, dpi=300, region_ratio=(0.50, 0.95)):
    """截取函证回函'以下由被询证单位填列'结论区（结论勾选/签章/日期/差异说明）→ 高清 PNG。
    ⚡ 函证结论/差异/签章均为手写，文本 OCR 几乎无法识别，必须人工看图核对——本函数生成专项图。
    定位策略：①pymupdf 文本层搜索'以下由被询证单位填列'位置（电子函证）；②失败 → 固定比例裁剪（扫描件）。
    返回 True 成功 / False 失败。
    """
    import fitz
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return False
    hit = False
    for page in doc:
        page_h = page.rect.height
        page_w = page.rect.width
        # 策略 1：文本层搜索
        try:
            rects = page.search_for('以下由被询证单位填列')
            if rects:
                y0 = rects[0].y1 + 4
                y1 = min(page_h, y0 + (page_h - y0) * 0.40)
                clip = fitz.Rect(0, y0, page_w, y1)
                pix = page.get_pixmap(dpi=dpi, clip=clip)
                pix.save(out_png)
                hit = True
                break
        except Exception:
            pass
        # 策略 2：固定比例（扫描件 fallback）
        if not hit:
            y0 = page_h * region_ratio[0]
            y1 = page_h * region_ratio[1]
            clip = fitz.Rect(0, y0, page_w, y1)
            pix = page.get_pixmap(dpi=dpi, clip=clip)
            pix.save(out_png)
            hit = True
            break
    doc.close()
    return hit


def _find_amounts(text):
    """函证金额项：完整项目名 + 金额（'银行存款 余额 X'/'应收账款余额 X'/'银行借款 X'）。"""
    _ITEMS = ['银行存款', '银行借款', '应收账款', '应付账款', '其他应收款', '其他应付款',
              '预付账款', '预收账款', '短期借款', '长期借款', '应付票据', '应收票据',
              '货币资金', '担保', '往来款']
    amounts = []
    for it in _ITEMS:
        for m in re.finditer(re.escape(it) + r'[^0-9\n]{0,10}?([\d,，.]+)\s*(?:元)?', text):
            try:
                val = float(m.group(1).replace(',', '').replace('，', ''))
            except ValueError:
                continue
            if val > 0:
                amounts.append({'item': it, 'amount': val})
                break   # 每项目取首个
    return amounts


def _is_date_fragment(s):
    """金额候选是否为日期/年份碎片（2025 / 202512 / 20251231 / 2025-12-31 等 OCR 拆解）。"""
    t = re.sub(r'[-/.,。]', '', s)          # 去分隔符
    if re.fullmatch(r'20\d{2}', t) or re.fullmatch(r'19\d{2}', t):
        return True                          # 年份
    # 日期碎片需带合法月份（20YYMM / 20YYMMDD），避免误伤 20503436 之类金额
    if re.fullmatch(r'20\d{2}(?:0[1-9]|1[0-2])(?:\d{2})?', t):
        return True
    return False


def _deconfuse(s):
    """OCR 字符混淆还原（数字上下文中 O→0、G→6、I/l/|→1）。"""
    s = s.replace('O', '0').replace('Ｏ', '0').replace('o', '0')
    s = s.replace('G', '6').replace('Ｇ', '6')
    s = s.replace('I', '1').replace('l', '1').replace('|', '1')
    return s


def _clean_money_str(s):
    """OCR 金额字符串矫正 → (数值, 置信 高/中/低)。
    ⚡ 千分位错位/字符混淆还原：
      7,328,863,73 → 7,328,863.73（末逗号误读小数点）
      1,172,000,00 → 1,172,000.00
      561,0O0.O0 / 1,116,G00.O0 → 字符混淆（O→0、G→6）
    """
    t = re.sub(r'[ \u3000]', '', s)
    t = t.replace('O', '0').replace('Ｏ', '0').replace('o', '0')
    t = t.replace('G', '6').replace('Ｇ', '6')
    t = t.replace('I', '1').replace('l', '1').replace('|', '1')
    t = t.replace('，', ',')
    # 1) 末位小数误读：\d{1,3}(,\d{3})+,<2位> → 末逗号改小数点
    m = re.fullmatch(r'(\d{1,3}(?:,\d{3}){1,}),(\d{2})', t)
    if m:
        return float(m.group(1).replace(',', '') + '.' + m.group(2)), '高'
    # 2) 标准千分位带小数点（每组3位）
    m = re.fullmatch(r'(\d{1,3}(?:,\d{3}){1,})\.(\d{1,3})', t)
    if m:
        return float(m.group(1).replace(',', '') + '.' + m.group(2)), '高'
    # 3) 千分位整数（每组3位）
    m = re.fullmatch(r'\d{1,3}(?:,\d{3}){1,}', t)
    if m:
        return float(t.replace(',', '')), '高'
    # 4) 普通数字（无千分位）
    m = re.fullmatch(r'-?\d[\d,]*\.?\d*', t)
    if m:
        return float(t.replace(',', '')), '中'
    return None, '低'


def _find_ar_amounts(text):
    """企业往来询证函：往来款项明细表（款项内容 + 金额）。
    ⚡ 表格 OCR 后为竖排单值行（项目名 / 截至日期 / 币种 / 金额 各占一行），金额与项目名分离，
    故按『项目名 → 其后首个金额』关联。方向由款项内容语义隐含：
    应收/预付/其他应收/应收票据=对方欠本公司；应付/预收/其他应付/应付票据=本公司欠对方。
    返回 [{'item', 'amount', 'direction'}]。
    """
    seg = re.search(r'[往求]?款[项顷]列示如下[：:\s]*?(.*?)(?:本函仅为复核|交易列示如下|被审计单位盖章|以下由被询证单位|结论[：:]|$)', text, re.S)
    if not seg:
        return []
    _AR_ITEMS = ['应收账款', '应付账款', '预付账款', '预收账款', '其他应收款', '其他应付款',
                 '应收票据', '应付票据']
    _RECV = ('应收账款', '预付账款', '其他应收款', '应收票据')   # 对方欠本公司
    _DATE = re.compile(r'(\d{4})[-/.]\d{1,2}[-/.]\d{1,2}')
    _AMT = re.compile(r'(\d[\d,，]*\.\d{1,3})|(\d[\d,，]{1,})')
    out = []
    cur = None
    for ln in seg.group(1).splitlines():
        ln = re.sub(r'[ \u3000]', '', ln).strip()   # ⚡ 清理数字内空格（OCR '1, 940, 081. 00'）
        if not ln:
            continue
        ln = _deconfuse(ln)                          # ⚡ 字符混淆还原（O→0 等），供金额正则匹配
        # 表头/说明行跳过
        if ln in ('币种', '款项内容', '截至日期', '备注', '贵公司欠本公司', '本公司欠贵公司',
                  '人民币', '美元', '欧元', '港币', '日元'):
            continue
        hit = None
        for it in _AR_ITEMS:
            # 项目名可带前导编号（'1.应收账款'）与表格后缀（'应付账款-暂估'/'应收账款一己开票'）
            if re.match(r'^\d{0,2}\s*[.、．\-]?\s*' + re.escape(it) + r'(?:[^\d]{0,8})?$', ln):
                hit = it
                break
        if hit:
            cur = hit
            continue
        if cur is None:
            continue
        if _DATE.search(ln) or '%' in ln:
            continue
        m = _AMT.search(ln)
        if m:
            raw = (m.group(1) or m.group(2))
            if _is_date_fragment(raw):
                continue
            v, conf = _clean_money_str(raw)
            if v is None or not (0 < v < 1e13):
                continue
            if conf == '低':
                continue
            if 0 < v < 1e13:
                # 同科目同方向累加（可能分 开票/暂估 多行）
                for ex in out:
                    if ex['item'] == cur:
                        ex['amount'] = round(ex['amount'] + v, 2)
                        break
                else:
                    out.append({'item': cur, 'amount': v,
                                'direction': '对方欠本公司' if cur in _RECV else '本公司欠对方'})
    return out


def _find_trade_amounts(text):
    """企业询证函『交易列示』段（销售/采购发生额，非往来余额）：
    2.本公司与贵公司之间的交易列示如下: 本公司向贵公司销售金额 / 本公司从贵公司采购金额 / 货款 / 2025年1-12月 / 人民币 / 金额。
    返回 [{'item':'销售'|'采购', 'amount'}]；方向由款项内容语义隐含。
    """
    seg = re.search(r'交易列示如下[：:\s]*?(.*?)(?:本函仅为复核|被审计单位盖章|以下由被询证单位|结论[：:]|$)', text, re.S)
    if not seg:
        return []
    body = seg.group(1)
    _DATE = re.compile(r'(\d{4})\s*[年\-/.]\s*\d{0,2}\s*[月\-/.]?\s*\d{0,2}')
    _AMT = re.compile(r'(\d[\d,，]*\.\d{1,3})|(\d[\d,，]{1,})')
    cur = None
    amt = {}
    for ln in body.splitlines():
        ln = re.sub(r'[ \u3000]', '', ln).strip()   # ⚡ 清理数字内空格
        if not ln:
            continue
        ln = _deconfuse(ln)
        if '销售金额' in ln or '向贵公司销售' in ln:
            cur = '销售'
            continue
        if '采购金额' in ln or '从贵公司采购' in ln:
            cur = '采购'
            continue
        if cur is None or _DATE.search(ln) or '%' in ln or '不含税' in ln or '币种' in ln:
            continue
        m = _AMT.search(ln)
        if m:
            raw = (m.group(1) or m.group(2))
            if _is_date_fragment(raw):
                continue
            v, conf = _clean_money_str(raw)
            if v is None or not (0 < v < 1e13):
                continue
            if conf == '低':
                continue
            if cur not in amt:
                amt[cur] = v
    return [{'item': k, 'amount': v} for k, v in amt.items()]


def _find_bank_deposit(text):
    """银行询证函回函：银行存款余额合计（'1．银行存款' 段内各账户余额求和）。
    返回 (余额合计, 账户数)。⚡ 表格 OCR 竖排：账号/币种/利率/余额各占一行，余额必带小数点；
    段头兼容带/不带编号前缀（建行'1.银行存款' / 中行'银行存款'）。优先取段内带小数金额行
    （排除利率%/日期/表头/币种行），无带小数行时以整数大数兜底（排除≥11位账号、前导0编号）。
    账户数=带小数金额行数（含0余额账户）。
    """
    seg = re.search(
        r'^\s*(?:[1１]\s*[.、．。,]\s*)?银行存款\s*\n(.*?)(?:^\s*[2２]\s*[.、．。,]\s*银行借款|^\s*[2２]\s*[.、．。,]\s*|^\s*[3３]\s*[.、．。,]\s*|^\s*证明编号\s*[:：])',
        text, re.S | re.M)
    if not seg:
        return 0.0, 0
    body = seg.group(1)
    _DATE = re.compile(r'(\d{4})[-/.]\d{1,2}[-/.]\d{1,2}')
    _CCY = ('人民币', '美元', '欧元', '港币', '日元', '英镑', '泰铢', '新加坡元')
    dec = []      # 带小数金额
    ints = []     # 整数大数兜底
    for ln in body.splitlines():
        ln = ln.strip()
        if not ln or _DATE.search(ln) or '%' in ln:
            continue
        if any(k in ln for k in ('账号', '账户名称', '币种', '利率', '起止日期')):
            continue
        if any(c in ln for c in _CCY):
            continue                       # 利率行常与币种同行（'人民币 0.05'），排除
        ln = _deconfuse(ln)                # ⚡ 字符混淆还原（O→0 等）
        m = re.search(r'(\d[\d,，]*\.\d{1,4})', ln)
        if m:
            raw = m.group(1)
            if _is_date_fragment(raw):
                continue
            v, conf = _clean_money_str(raw)
            if v is None or abs(v) >= 1e13:
                continue
            if conf == '低':
                continue
            dec.append(v)              # 0 余额账户仍计入账户数（金额0不计合计）
            continue
        # 整数兜底：排除≥11位账号、前导0编号、年份/日期碎片
        m = re.search(r'(\d[\d,，]{1,})', ln)
        if m:
            raw = m.group(1)
            if len(raw) >= 11 or raw.startswith('0'):
                continue
            v, conf = _clean_money_str(raw)
            if v is None or not (100 <= v < 1e13):
                continue
            if conf == '低':
                continue
            ints.append(v)
    if dec:
        total = round(sum(v for v in dec if v > 0), 2)
        return total, len(dec)
    return round(sum(ints), 2), len(ints)


def _find_bank_loan(text):
    """银行询证函『2.银行借款』段：返回 (余额合计, 借款笔数)。
    格式与银行存款类似（借款人名称/账号/币种/余额/利率 竖排），无借款时回函列『无』。
    """
    seg = re.search(r'^\s*[2２]\s*[.、．。,]\s*银行借款\s*\n(.*?)(?:^\s*[3３]\s*[.、．。,]\s*|^\s*证明编号\s*[:：])', text, re.S | re.M)
    if not seg:
        return 0.0, 0
    body = seg.group(1)
    if '无此事项' in body or re.search(r'^\s*无\s*$', body, re.M):
        return 0.0, 0
    _DATE = re.compile(r'(\d{4})[-/.]\d{1,2}[-/.]\d{1,2}')
    _HDR = ('名称', '账号', '币种', '余额', '日期', '利率', '押品', '担保', '借款日', '到期日', '编号', '序号')
    total = 0.0
    n = 0
    for ln in body.splitlines():
        ln = _deconfuse(ln.strip())
        if not ln or _DATE.search(ln) or '%' in ln or any(k in ln for k in _HDR):
            continue
        m = re.search(r'(\d[\d,，]*\.\d{1,4})', ln)
        if m:
            v, conf = _clean_money_str(m.group(1))
            if v and 0 < v < 1e13 and conf != '低':
                total += v
                n += 1
    return round(total, 2), n


_BANK_STD_14 = {
    1: '银行存款',
    2: '银行借款',
    3: '作为委托人的委托贷款',
    4: '本公司作为借款人的委托贷款',
    5: '担保-本公司为其他单位提供的担保',
    6: '担保-贵行向本公司提供的担保',
    7: '本公司为出票人且由贵行承兑而尚未支付的银行承兑汇票',
    8: '本公司向贵行已贴现而尚未到期的商业汇票',
    9: '本公司为持票人且由贵行托收（或由本公司提示付款）的商业汇票',
    10: '本公司为申请人由贵行开具的未履行完毕的不可撤销信用证',
    11: '本公司与贵行之间未履行完毕的外汇买卖合约',
    12: '本公司存放于贵行托管的证券或其他产权文件',
    13: '本公司购买的由贵行发行的未到期银行理财产品',
    14: '其他',
}


_BANK_STD_14 = {
    1: '银行存款',
    2: '银行借款',
    3: '作为委托人的委托贷款',
    4: '本公司作为借款人的委托贷款',
    5: '担保-本公司为其他单位提供的担保',
    6: '担保-贵行向本公司提供的担保',
    7: '本公司为出票人且由贵行承兑而尚未支付的银行承兑汇票',
    8: '本公司向贵行已贴现而尚未到期的商业汇票',
    9: '本公司为持票人且由贵行托收（或由本公司提示付款）的商业汇票',
    10: '本公司为申请人由贵行开具的未履行完毕的不可撤销信用证',
    11: '本公司与贵行之间未履行完毕的外汇买卖合约',
    12: '本公司存放于贵行托管的证券或其他产权文件',
    13: '本公司购买的由贵行发行的未到期银行理财产品',
    14: '其他',
}


_BANK_STD_14 = {
    1: '银行存款',
    2: '银行借款',
    3: '作为委托人的委托贷款',
    4: '本公司作为借款人的委托贷款',
    5: '担保-本公司为其他单位提供的担保',
    6: '担保-贵行向本公司提供的担保',
    7: '本公司为出票人且由贵行承兑而尚未支付的银行承兑汇票',
    8: '本公司向贵行已贴现而尚未到期的商业汇票',
    9: '本公司为持票人且由贵行托收（或由本公司提示付款）的商业汇票',
    10: '本公司为申请人由贵行开具的未履行完毕的不可撤销信用证',
    11: '本公司与贵行之间未履行完毕的外汇买卖合约',
    12: '本公司存放于贵行托管的证券或其他产权文件',
    13: '本公司购买的由贵行发行的未到期银行理财产品',
    14: '其他',
}


_BANK_STD_14 = {
    1: '银行存款',
    2: '银行借款',
    3: '作为委托人的委托贷款',
    4: '本公司作为借款人的委托贷款',
    5: '担保-本公司为其他单位提供的担保',
    6: '担保-贵行向本公司提供的担保',
    7: '本公司为出票人且由贵行承兑而尚未支付的银行承兑汇票',
    8: '本公司向贵行已贴现而尚未到期的商业汇票',
    9: '本公司为持票人且由贵行托收（或由本公司提示付款）的商业汇票',
    10: '本公司为申请人由贵行开具的未履行完毕的不可撤销信用证',
    11: '本公司与贵行之间未履行完毕的外汇买卖合约',
    12: '本公司存放于贵行托管的证券或其他产权文件',
    13: '本公司购买的由贵行发行的未到期银行理财产品',
    14: '其他',
}


def _find_bank_items(text):
    """银行询证函**完整事项切段（1-14 项逐项展开）**。
    兼容编号独立一行 / 编号+名称同行 / 子项(6(1))，按标准 14 项名称表补全缺失项。
    返回 [{'no','name','conclusion','summary'}]，no 恒为 1-14。
    """
    lines = text.splitlines()
    # 编号行：`7.名称` / `7、名称` / `6.(1)…` / `3。名称`（全角句号）/ 纯编号 `3`
    #        / `7该公司…`（编号直接跟中文，无分隔符）
    _NO_RE = re.compile(
        r'^([1-9]|1[0-4])\s*(?:[.、．,，。]\s*(.*)|([\u4e00-\u9fa5].*))?$')
    # 说明性行（页脚/提示，不是询证事项）→ 跳过
    _SKIP_NAME_KW = ('核验', '此证明', '如需', '加盖', '复印', '涂改', '有效',
                     '多页', '缺页', '本函', '本回函', '证明编号', '总页', '电子印章')
    _FIELD_KW = ('序号', '类型', '编号', '账号', '账户', '币种', '利率', '余额', '日期', '号码',
                 '金额', '名称', '笔号', '到期', '类别', '数量', '净值', '份额', '汇率', '交收',
                 '结算', '起止', '产品')
    _NOTE_KW = ('手机银行', '融e联', '可验证询证函内容', '文档唯一号', '证明编号',
                '银行经办人姓名', '银行复核人姓名')

    def _is_header(ln):
        if not ln:
            return False
        if ln.startswith(('序号', '类别', '银行承兑汇票号码', '商业汇票号码', '信用证号码')):
            return True
        return sum(1 for k in _FIELD_KW if k in ln) >= 2

    found = {}          # no -> {'name', 'start'}
    seq = []
    for i, ln in enumerate(lines):
        s = ln.strip()
        m = _NO_RE.match(s)
        if not m:
            continue
        no = int(m.group(1))
        name = ((m.group(2) or m.group(3)) or '').strip()
        # 编号独立一行 → 从下 1-3 行取名称
        if not name:
            for j in range(i + 1, min(i + 4, len(lines))):
                nxt = lines[j].strip()
                if (nxt and not _NO_RE.match(nxt)
                        and not _is_header(nxt)
                        and not any(k in nxt for k in _NOTE_KW)
                        and re.search(r'[\u4e00-\u9fa5]', nxt)):
                    name = nxt
                    break
        # 名称清洗：去 (1) 子项前缀、栏位说明，取第一段
        name = re.sub(r'^[（(]\d+[）)]?', '', name)
        name = re.split(r'[（(]', name)[0].strip()[:30]
        # 表格数据行（编号后跟金额/数字，如 '2.309.964否'）非询证事项 → 跳过
        if re.match(r'^[\d.%]+', name):
            continue
        # 表格账户序号（活期5/定期3/序号N）非询证事项编号 → 跳过
        if re.match(r'^(活期|定期|序号)\s*\d+', name):
            continue
        # 说明性/干扰行 → 不是询证事项，跳过
        if any(k in name for k in _SKIP_NAME_KW):
            continue
        if no not in found:
            found[no] = {'name': name, 'start': i}
            seq.append(no)

    # —— 编号反查：OCR 漏识别编号的项，用标准名称关键词补位（从上一个已知项之后查找）——
    _BANK_KW = {
        2: ['银行借款'],
        3: ['注销的银行存款账户', '注销的银行'],
        4: ['作为委托方的委托贷款', '作为委托人的委托贷款'],
        5: ['作为借款方的委托贷款', '作为借款人的委托贷款'],
        6: ['担保受益人的担保', '担保受益人的', '为其他单位提供的'],
        7: ['为出票人', '承兑而尚未支付'],
        8: ['已贴现'],
        9: ['为持票人', '托收'],
        10: ['为申请人', '不可撤销信用证'],
        11: ['外汇买卖合约'],
        12: ['托管'],
        13: ['理财'],
        14: ['其他'],
    }

    def _char_of_line(line_no):
        if line_no <= 0:
            return 0
        return len('\n'.join(lines[:line_no])) + 1

    search_from = 0
    present_nos = set(found.keys())
    for no in range(1, 15):
        if no in present_nos:
            search_from = max(search_from, _char_of_line(found[no]['start']))
            continue
        for kw in _BANK_KW.get(no, []):
            idx = text.find(kw, search_from)
            if idx > 0:
                line_no = text[:idx].count('\n')
                found[no] = {'name': kw, 'start': line_no}
                seq.append(no)
                search_from = idx
                break

    out = []
    for k, no in enumerate(seq):
        it = found[no]
        end = found[seq[k + 1]]['start'] if k + 1 < len(seq) else len(lines)
        body = [ln.strip() for ln in lines[it['start'] + 1:end]
                if ln.strip() and not ln.strip().startswith(('----', '---'))]
        # 表格列名/勾选栏名（大华/建行表格表头，OCR 拆行）→ 不作为实际内容
        _TABLE_HEAD_KW = ('是否存在', '是否属于', '是否被', '账户名称', '账户类型',
                          '银行账号', '借款账号', '起始日期', '终止日期', '币种',
                          '账户余额', '使用限制', '备注', '账号', '抵押', '质押',
                          '担保', '序号', '借款日', '到期日', '类型', '编号',
                          '资金归集', '冻结', '银行', '账户')

        def _is_tabular(ln):
            if not ln:
                return False
            if ln.startswith(('是否存在', '是否属于', '是否被', '账户', '银行账号',
                              '借款账号', '起始', '终止', '抵押', '质押')):
                return True
            if sum(1 for k in _TABLE_HEAD_KW if k in ln) >= 2:
                return True
            # 无数字的短中文碎片（拆行表头）→ 表头
            if not re.search(r'\d', ln) and len(ln) <= 12 \
                    and any(k in ln for k in _TABLE_HEAD_KW):
                return True
            return False

        clean = [ln for ln in body
                 if ln and not _is_header(ln) and not _is_tabular(ln)
                 and not any(k in ln for k in _NOTE_KW)
                 and '=====' not in ln                        # 页标记
                 and not re.match(r'^第\s*\d+\s*页', ln)      # 第N页
                 and not re.search(r'共\s*\d+\s*页', ln)      # 共N页
                 and not re.match(r'^询证函编号', ln)
                 and not re.match(r'^编号[:：]', ln)          # 页脚证明编号
                 and '业务专用章' not in ln
                 and not re.match(r'^([1-9]|1[0-4])\s*(?:[.、．。]\s*)?[\u4e00-\u9fa5]', ln)]  # 重复编号标题行

        def _has_amount(ln):
            """明确金额特征：大金额、余额/金额/面额/份额/净值/汇率+数字、账号+金额。"""
            if re.search(r'\d[\d,，]{4,}\.\d{1,2}', ln):
                return True
            if re.search(r'(?:余额|金额|面额|份额|净值|汇率)[：:：\s]*[\d][\d,，]*\.?\d*', ln):
                return True
            if re.search(r'利率[：:：\s]*[\d.]+%', ln):
                return True
            if re.search(r'\b\d{10,}\b', ln) and re.search(r'[\d,，]+\.\d{1,2}', ln):
                return True
            return False

        def _is_real_content(ln):
            """实际内容：含数字（金额/账号/日期）或 ≥20 字长句；短中文碎片/子项标题视为表头/无。"""
            if len(ln) < 4:
                return False
            if any(k in ln for k in ('无此事项', '未找到', '不存在', '特此回复',
                                     '证明编号', '以下空白', '本页无正文')):
                return False
            if ln in ('无', '无。', '/', '—', '-'):
                return False
            if re.match(r'^[（(]\d+[）)]', ln):   # 子项标题（(1)担保…）非内容
                return False
            if re.search(r'\d', ln):
                return True
            if len(ln) >= 20:
                return True
            return False

        has_amount = any(_has_amount(ln) for ln in clean)
        has_wu = any('无此事项' in ln or '未找到' in ln or '不存在' in ln
                     or ln in ('无', '无。', '/', '—', '-')
                     or (ln and all(c in '无 ' for c in ln))
                     or re.search(r'除上述列示的?[^。]{0,40}?(并无|并未)', ln)
                     or '注销其他账户' in ln
                     or '无其他' in ln
                     for ln in clean)
        has_real = any(_is_real_content(ln) for ln in clean)
        if has_amount:
            conclusion = '有明细'
        elif has_wu or not has_real:
            conclusion = '无'
        else:
            conclusion = '待核实'
        summary = '；'.join(clean[:4])[:120]
        out.append({'no': no, 'name': (it['name'] or _BANK_STD_14.get(no, ''))[:30],
                    'conclusion': conclusion, 'summary': summary})

    # 补全 1-14 缺失项（OCR 漏识别编号时用标准名兜底）
    present = {it['no'] for it in out}
    for no in range(1, 15):
        if no not in present:
            out.append({'no': no, 'name': _BANK_STD_14[no],
                        'conclusion': '未识别', 'summary': ''})
    out.sort(key=lambda x: x['no'])
    return out

def parse_confirm(text):
    """函证回函 OCR 文本 → 结构化。"""
    is_bank = ('银行询证函' in text or '银行存款' in text and '询证' in text)
    is_ar = (not is_bank and '询证' in text)   # ⚡ 企业询证函放宽判定（'往来款项'/'贵公司' 均可），银行排除
    out = {
        'type': '银行询证函回函' if is_bank else ('往来询证函回函' if is_ar else '函证回函'),
        'ref_no': _find_ref(text),
        'auditee': _find_auditee(text),
        'to_auditor': '',
        'conclusion': _find_conclusion(text),
        'discrepancy': '',
        'discrepancy_amount': 0.0,    # ⚡ 手写差异金额
        'discrepancy_date': '',        # ⚡ 截至日期（手写或印刷）
        'amounts': _find_amounts(text),
        'ar_amounts': _find_ar_amounts(text),
        'trade_amounts': [],       # 交易列示（销售/采购发生额）
        'deposit_total': 0.0,
        'deposit_accts': 0,
        'loan_total': 0.0,
        'loan_count': 0,
        'bank_items': [],       # 银行询证函全部事项（含表外）
        'off_balance': [],      # 表外事项（有明细的），如担保/汇票/信用证/理财/委托贷款等
        'stamp_unit': _find_stamp(text),
        'reply_date': _find_date(text),
        'matched_items': 0,
        'mismatch_items': 0,
    }
    if is_bank:
        out['deposit_total'], out['deposit_accts'] = _find_bank_deposit(text)
        out['loan_total'], out['loan_count'] = _find_bank_loan(text)
        out['bank_items'] = _find_bank_items(text)
        out['off_balance'] = [it for it in out['bank_items']
                              if it['conclusion'] == '有明细' and it['no'] >= 3]
    m = re.search(r'(?:致|寄至|致送)[：:\s]*([\u4e00-\u9fa5]{2,30}?(?:会计师事务所|事务所))', text)
    if m:
        out['to_auditor'] = m.group(1).strip()
    m = re.search(r'(?:差异说明|不符事项|说明)[：:\s]*([^\n]{2,100})', text)
    if m:
        out['discrepancy'] = m.group(1).strip()
    # 往来函证差异说明专项提取（手写差异金额/日期）
    if is_ar:
        out['trade_amounts'] = _find_trade_amounts(text)
        d = _find_discrepancy(text)
        if d['discrepancy_text'] and not out['discrepancy']:
            out['discrepancy'] = d['discrepancy_text']
        if d['discrepancy_amount']:
            out['discrepancy_amount'] = d['discrepancy_amount']
        if d['discrepancy_date']:
            out['discrepancy_date'] = d['discrepancy_date']
    out['matched_items'] = len(re.findall(r'相符', text))
    out['mismatch_items'] = len(re.findall(r'不相符', text))
    return out


if __name__ == '__main__':
    import sys, glob, os, json
    if len(sys.argv) > 1:
        for f in sys.argv[1:]:
            import contract_reader as CR
            text, fmt = CR.extract_text(f)
            r = parse_confirm(text)
            print(f'=== {os.path.basename(f)} ({fmt}) ===')
            print(json.dumps(r, ensure_ascii=False, indent=1))
