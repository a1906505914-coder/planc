# -*- coding: utf-8 -*-
"""
合同 OCR 结构化摘要提取 + 汇总
================================
输入：合同 OCR 输出目录（contract_ocr_batch.py 生成，`_ocr_识别结果/<相对路径>/<文件名>.pdf.txt`）
输出（写入 合同摘要/ 子目录，与 OCR 相对路径一致）：
  1. 每份合同一个 Markdown 摘要文件（<文件名>.md）——含 出租方/承租方/标的/面积/租期/租金/递增/支付频率/免租期/押金
  2. 合同摘要汇总.xlsx —— 全部合同一张总表
  3. 合同摘要汇总.md —— 全部合同汇总文档（便于快速浏览）

用法：
  python contract_summary.py <合同目录> [--out <输出目录>]
  # <合同目录> 指向含 _ocr_识别结果 的目录（如 使用权资产租赁合同）
"""
import argparse
import glob
import os
import re
import sys
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------------- 文本归一
_NORM_MAP = {
    '＇': "'", '‘': "'", '’': "'", '“': '"', '”': '"', '，': ',', '。': '.',
    '；': ';', '：': ':', '（': '(', '）': ')', '【': '[', '】': ']', '、': ',',
    '—': '-', '－': '-', '‐': '-', '－': '-', '·': '.',
    '㎡': '平方米', 'm2': '平方米', 'm²': '平方米',
    '仟': '千', '任': '千', '佰': '百', '壹': '一', '贰': '二', '叁': '三',
    '参': '三', '肆': '四', '伍': '五', '陆': '六', '柒': '七', '捌': '八',
    '玖': '九', '拾': '十', '佰': '百',
    '圆': '元', '整': '',
}
_CN_DIGITS = {'零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
              '六': 6, '七': 7, '八': 8, '九': 9}
# 中文大写金额候选字符（含 OCR 常见大写/错字），用于正则限定
_CN_CHARS = '零一二三四五六七八九十百千万亿拾佰仟壹贰叁肆伍陆柒捌玖任参'


def _norm(s):
    """OCR 常见错字/全角归一（便于正则匹配）。"""
    s = str(s or '')
    for a, b in _NORM_MAP.items():
        s = s.replace(a, b)
    return s


def _cn_money(s):
    """中文大写金额 → 数值（万元）。支持 亿/万/千/百/十 混合、缺位。"""
    s = _norm(s).replace('元', '')
    m = re.search(r'[零一二三四五六七八九十百千万亿]+', s)
    if not m:
        return None
    units = {'十': 10, '百': 100, '千': 1000, '万': 10000, '亿': 100000000}
    tot = 0.0
    cur = 0
    sec = 0
    for ch in m.group(0):
        if ch in _CN_DIGITS:
            sec = _CN_DIGITS[ch]
        elif ch in units:
            u = units[ch]
            if u >= 10000:                       # 万/亿：封段
                cur = (cur + (sec or 1)) * u
                tot += cur
                cur = 0
                sec = 0
            else:                                 # 十/百/千：累计当前段
                cur += (sec or 1) * u
                sec = 0
    cur += sec
    return (tot + cur) / 10000.0                 # 元 → 万元


def _cn_int(s):
    """中文数字 → int（支持 一~十、十N、N十、N十N）。"""
    t = _norm(s)
    m = re.search(r'[零一二三四五六七八九十]+', t)
    if not m:
        return None
    t = m.group(0)
    if t == '十':
        return 10
    if '十' in t:
        a, b = t.split('十')
        return int((_CN_DIGITS.get(a, 1) if a else 1) * 10 + (_CN_DIGITS.get(b, 0) if b else 0))
    if len(t) == 1:
        return _CN_DIGITS.get(t, 0)
    return None


# ---------------------------------------------------------------- 日期
def _date(s):
    """'2025年4月1日' / '2024.3.31' / '2025-12-31' / '2025.4' → 'YYYY-MM-DD' 或 ''。"""
    s = (s or '').strip().strip('_').strip()
    m = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?', s)
    if m:
        return f'{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    m = re.search(r'(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})', s)
    if m:
        return f'{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
    return ''


def _months(d0, d1):
    try:
        from datetime import date
        a = date(*[int(x) for x in d0.split('-')])
        b = date(*[int(x) for x in d1.split('-')])
        return max(0, (b.year - a.year) * 12 + b.month - a.month)
    except Exception:
        return 0


def _fnum(s):
    """'806,872.32' → float。"""
    try:
        return float(str(s).replace(',', '').replace('_', '').strip())
    except (ValueError, TypeError):
        return None


# 日期捕获（含 月/日/年/数字/点/斜杠/下划线/连字符；贪婪直到遇到 止/至/起/、/。/）等终止符）
_D = r'([\d._年\-/月日]+)'
# ---------------------------------------------------------------- 租期提取
_PERIOD_PATS = [
    r'(?:租赁期为|租期为)(?:\d+|[一二三四五六七八九十]+)\s*年\s*[，,]\s*自\s*' + _D + r'\s*至\s*' + _D + r'(?:\s*止)?',
    r'本合同项下的租赁期自\s*' + _D + r'\s*起至\s*' + _D + r'(?:\s*止)?',
    r'租赁期自\s*' + _D + r'\s*起至\s*' + _D + r'(?:\s*止)?',
    r'租期从\s*' + _D + r'\s*起至\s*' + _D + r'(?:\s*止)?',
    r'(?:房屋)?租赁期限为\s*' + _D + r'\s*至\s*' + _D + r'(?:\s*止)?',
    r'房屋租赁期为\s*' + _D + r'\s*至\s*' + _D + r'(?:\s*止)?',
    r'合同期限[:：]?\s*自\s*' + _D + r'\s*至\s*' + _D,
    r'租赁期限[:：]?\s*' + _D + r'\s*至\s*' + _D + r'(?:\s*止)?',
    r'房屋租赁期自\s*' + _D + r'\s*[（(][^）)0]{0,30}[）)0]?\s*至\s*' + _D + r'(?:\s*止)?',
    r'续租\s*[（(]\s*([\d.\-/]+?)\s*[-—]\s*([\d.\-/]+?)\s*[）)]',
]


def _extract_period(full, fname):
    """从全文/文件名提取租期起止 (d0, d1)。"""
    for p in _PERIOD_PATS:
        m = re.search(p, full)
        if m:
            d0, d1 = _date(m.group(1)), _date(m.group(2))
            if d0 and d1:
                return d0, d1
    m = re.search(r'(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})\s*[-—]\s*(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})', fname)
    if m:
        return (f'{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}',
                f'{int(m.group(4)):04d}-{int(m.group(5)):02d}-{int(m.group(6)):02d}')
    return '', ''


# ---------------------------------------------------------------- 租金提取
def _extract_rent(full, lines):
    """返回 (首年租金描述, 万元, 口径)。提取失败 → ('', None, '')。"""
    # 1. 金属格式：首年（...）租金为人民币...小写_584784.00_元
    m = re.search(r'首年[（(][^)）]*[）)]?租金为人民币.{0,40}?小写[_\s]*([\d,.]+)[_\s]*元', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元', v / 10000, '首年(合同原文)'
    # 2. 集团本级：第一年至第三年年租金为XXX万元
    m = re.search(r'第一年至第三年年租金为\s*([\d,.]+)\s*万元', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.2f} 万元', v, '第1-3年年租金'
    # 3. 绍兴：第一年：48.16万元（其中含税9.63万元）
    m = re.search(r'第一年[：:]\s*([\d,.]+)\s*万元\s*(?:（其中含税\s*([\d,.]+)\s*万元）)?', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            tax = f'，含税{_fnum(m.group(2)):,.2f}万' if m.group(2) and _fnum(m.group(2)) else ''
            return f'{v:,.2f} 万元{tax}', v, '第一年'
    # 4. 集团本级/智能（七格模板）：年租金为XXX元 / 年租金XXX元
    m = re.search(r'年租金\s*为?\s*([\d,.]+)\s*元', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元/年', v / 10000, '年租金'
    # 5. 海宁：租金合计为XXX元/年
    m = re.search(r'租金合计为\s*([\d,.]+)\s*元/年', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元/年', v / 10000, '年租金'
    # 6. 龙游场地：年租金合计XXX元（含税）
    m = re.search(r'年租金合计\s*([\d,.]+)\s*元\s*（含税）', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元/年(含税)', v / 10000, '年租金(含税)'
    # 7. 亨德利萧山：年租金额（含税/合税）为人民币：105984元/年
    m = re.search(r'年租金额[（(][含合]税[）)]为人民币[:：]?\s*([\d,.]+)\s*元/年', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元/年(含税)', v / 10000, '年租金(含税)'
    # 7b. 建友：自X至Y，该房屋租金为2410800元
    m = re.search(r'该房屋租金为\s*([\d,.]+)\s*元', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元/年', v / 10000, '首年(逐年租金)'
    # 8. 金华房屋：房屋租赁费按每年含税价￥1204900.00元
    m = re.search(r'房屋租赁费按每年含税价[¥￥]\s*([\d,.]+)\s*元', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元/年(含税)', v / 10000, '年租金(含税)'
    # 9. 金华设备：首个租赁年度租金为含税人民币930762.00元
    m = re.search(r'首个租赁年度租金为含税人民币\s*([\d,.]+)\s*元', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元/年(含税)', v / 10000, '首年(含税)'
    # 10. 下沙：租金总额为人民币（大写...）整（￥8582453.00）→ 中文大写优先
    m = re.search(r'租金总额为人民币([' + _CN_CHARS + r']+)元?整', full)
    if m:
        v = _cn_money(m.group(1))
        if v:
            return f'{v:,.2f} 万元', v, '租金总额(中文大写)'
    # 11. 亨德利二楼：首年度租金（含税）为：人民币￥4716&元（OCR 数字损坏 → 放弃，不落入通用￥）
    m = re.search(r'首年度租金[（(]含税[）)]为[:：]?人民币[¥￥]\s*[\d,.]+\s*[&＆]', full)
    if m:
        return '', None, 'OCR金额损坏待人工补'
    # 11b. 龙游设备：总计￥XXX（10年总价，非年租金）→ 标注为总价
    m = re.search(r'总计[¥￥]\s*([\d,.]+)\s*元', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元（总价）', v / 10000, '合同总价(含税)'
    # 12. 通用：￥185000元（精工等）
    m = re.search(r'[¥￥]\s*([\d,.]+)(?:\s*元)?', full)
    if m:
        v = _fnum(m.group(1))
        if v:
            return f'{v:,.0f} 元', v / 10000, '￥金额'
    # 13. 中文大写兜底
    m = re.search(r'(?:年)?租金为人民币([零一二三四五六七八九十百千万亿]+)元', full)
    if m:
        v = _cn_money(m.group(1))
        if v:
            return f'{v:,.2f} 万元', v, '年租金(中文大写)'
    return '', None, ''


# ---------------------------------------------------------------- 主提取
def extract_summary(txt_path):
    """单份合同 OCR txt → 结构化摘要 dict。"""
    with open(txt_path, encoding='utf-8') as f:
        raw = f.read()
    text = re.sub(r'={5,}\s*第\s*\d+\s*页\s*={5,}', '\n', raw)   # 去页码分隔头
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    full = re.sub(r'\s+', '', text)                                # 去空白全文
    fname = os.path.basename(txt_path)[:-4]

    # ---- 合同名称：首行 / 文件名
    name = fname
    for ln in lines[:8]:
        if 4 <= len(ln) <= 40 and not re.search(r'合同编号|出租方|承租方|甲方|乙方|续租', ln):
            name = ln
            break

    # ---- 出租方 / 承租方（全文收集候选 → 已知库模糊纠正 + 取最优）
    lessor_cands, lessee_cands = [], []
    for ln in lines[:80]:
        t = re.search(r'(?:出租方|甲方|申方|出租人)\s*(?:[（(][^）)（(]*[）)])?\s*[:：]\s*(.+)$', ln)
        if t:
            lessor_cands.append(_clean_party(t.group(1)))
        t = re.search(r'(?:承租方|乙方|承租人)\s*(?:[（(][^）)（(]*[）)])?\s*[:：]\s*(.+)$', ln)
        if t:
            lessee_cands.append(_clean_party(t.group(1)))

    def _pick(cands):
        if not cands:
            return '', 0.0
        scored = [(_match_known(c), c) for c in cands]
        best = max(scored, key=lambda x: x[0][1])
        if best[0][1] >= 0.45:
            return best[0][0], best[0][1]      # 用已知库纠正后的完整名
        return max(cands, key=len), 0.0        # 无法匹配 → 取最长候选

    lessor, _r = _pick(lessor_cands)
    lessee, _r = _pick(lessee_cands)

    # ---- 租赁标的 / 面积
    subject = ''
    m = re.search(r'(?:坐落于|位于|坐落在)(.{2,60}?)[，,。(（；;]', full)
    if m:
        subject = m.group(1)
    if not subject and name and name != fname:
        subject = name                      # 标的兜底：合同名称（如"美妆大厦写字楼租赁合同"）
    area = ''
    # 优先：一层+二层 多层面积求和（绍兴安昌格式）
    m = re.search(r'(?:一层|第一层|1层).{0,20}?([\d,.]+)\s*平方米.{0,40}?(?:二层|第二层|2层).{0,20}?([\d,.]+)\s*平方米', full)
    if m:
        v1, v2 = _fnum(m.group(1)), _fnum(m.group(2))
        if v1 and v2:
            area = f'{v1 + v2:,.2f}'
    if not area:
        m = re.search(r'(?:建筑面积|出租面积为|出租面积|租赁面积|面积约|面积为)\s*[约为_]*\s*([\d,.]+)_?\s*平方米', full)
        if m:
            area = m.group(1)
    if not area:
        m = re.search(r'([\d,.]+)_?\s*平方米', full)
        if m:
            area = m.group(1)

    # ---- 租期起止
    d0, d1 = _extract_period(full, fname)
    years = round(_months(d0, d1) / 12.0, 1) if d0 and d1 else ''

    # ---- 首年租金
    rent, rent_wan, rent_src = _extract_rent(full, lines)

    # ---- 单价（支持 OCR 下划线/全角斜杠；排除"物业费"上下文的单价）
    unit = ''
    _UNIT_PAT = re.compile(
        r'([\d.]+\s*_?\s*元[／/]平方米\s*[／/]\s*月|[\d.]+\s*_?\s*元[／/]m\s*[／/]\s*月|'
        r'[\d.]+\s*_?\s*元[／/]平方\s*[／/]\s*月|[\d.]+\s*_?\s*元\s*[／/]m[／/]\s*天|'
        r'[\d.]+\s*_?\s*元[／/]天[／/]m|[\d.]+\s*_?\s*元[／/]㎡²?\s*[／/]\s*天|'
        r'[\d.]+\s*_?\s*元[／/]平方米\s*[／/]\s*天|[\d.]+\s*_?\s*元[／/]平方米\s*[／/]\s*年)')
    for m in _UNIT_PAT.finditer(full):
        pre = full[max(0, m.start() - 12):m.start()]
        if '物业' in pre:
            continue                       # 物业费单价（非租金），跳过
        unit = m.group(1).replace(' ', '')
        break

    # ---- 租金递增
    inc = ''
    m = re.search(r'第([一二三四五六七八九十\d]+)\s*年起[^。]{0,30}?递增\s*([\d.]+)%', full)
    if m:
        n = _cn_int(m.group(1)) or m.group(1)
        inc = f'第{n}年起递增{m.group(2)}%'

    # ---- 支付频率
    freq = ''
    m = re.search(r'每\s*(\d+)\s*个月?支付', full)
    if m:
        freq = f'每{m.group(1)}个月一次'
    elif '每半年支付' in full:
        freq = '每半年一次'
    elif '每季度为1期' in full or '分4期支付' in full or '按季度支付' in full or '每季度支付' in full:
        freq = '每季度一次'
    elif '壹年壹付' in full or '租金每年一付' in full or '每年5月1日前一次性缴清' in full \
            or '应于每年' in full or '每年需支付' in full:
        freq = '每年一付'
    elif '按年度核算' in full or '按年度支付' in full or '按年支付' in full or '每一租赁年度' in full:
        freq = '每年一付'
    elif '年度租金' in full and '一次性缴清' in full:
        freq = '每年一付'

    # ---- 免租期
    free = ''
    m = re.search(r'装修免租期为[^。]{0,30}?([一\d]+\s*个月|[一\d]+月)', full)
    if m:
        free = f'装修免租 {m.group(1)}'
    elif not free:
        m = re.search(r'免租期[:：]?[^。]{0,15}?(\d+\s*天|\d+\s*个月)', full)
        if m:
            free = f'免租期 {m.group(1)}'
    if not free and (('前三年' in full or '三年内' in full) and '免租' in full):
        free = '有免租期（见备注）'
    if not free and '免交租金' in full:
        free = '有免租/补助政策（见备注）'

    # ---- 押金/保证金
    deposit = ''
    m = re.search(r'(?:保证金|押金)[^。]{0,25}?([\d,.]+)\s*万元', full)
    if m:
        deposit = f'{m.group(1)} 万元'
    if not deposit:
        m = re.search(r'人民币([' + _CN_CHARS + r']+)元作为履约保证金', full)
        if m:
            v = _cn_money(m.group(1))
            if v:
                deposit = f'{v:,.2f} 万元'
    if not deposit and ('免收押金' in full or '免收保证金' in full):
        deposit = '免收'

    # ---- 备注
    remark = ''
    if '三免三减半' in full:
        m = re.search(r'(.{0,20}三免三减半.{0,60})', full)
        if m:
            remark = m.group(1)
    # 续租判断：仅文件名或文本标题区（前 8 行）出现"续租"才认定（排除正文"可续租"条款误判）
    head_text = ''.join(lines[:8])
    if '续租' in fname or ('续租' in head_text and '若需续租' not in head_text):
        remark = (remark + '；' if remark else '') + '续租合同'
    if '租金支付表' in full and not rent:
        remark = (remark + '；' if remark else '') + '年租金金额见合同附件《租金支付表》'
    if not rent and 'OCR金额损坏' in rent_src:
        remark = (remark + '；' if remark else '') + '租金金额 OCR 识别损坏，需人工核对该页'
    if lessee and len(lessee) < 4:
        remark = (remark + '；' if remark else '') + '承租方名称 OCR 残缺，需人工核对'
    if not lessor:
        remark = (remark + '；' if remark else '') + '出租方名称合同未载明，需人工核实'

    return OrderedDict([
        ('合同名称', name),
        ('出租方', lessor),
        ('承租方', lessee),
        ('租赁标的', subject),
        ('面积(㎡)', area),
        ('租期起', d0),
        ('租期止', d1),
        ('租期(年)', years),
        ('首年租金', rent),
        ('首年租金(万元)', rent_wan),
        ('租金口径', rent_src),
        ('单价', unit),
        ('租金递增', inc),
        ('支付频率', freq),
        ('免租期', free),
        ('押金/保证金', deposit),
        ('备注', remark),
        ('源文件', os.path.basename(txt_path)),
    ])


def _clean_party(s):
    """清理主体名：去尾部杂质（地址/法人/（以下简称X方）/孤立'一'）。"""
    s = s.split('，')[0].split('地址')[0].split('法定代表人')[0].split('法人代表')[0].strip()
    s = re.sub(r'[一]?（以下简称[甲乙]方）\s*$', '', s)      # 尾部（以下简称甲方）等
    s = re.sub(r'[一]?\s*（以下简称[甲乙]方）\s*$', '', s)
    s = re.sub(r'\s+', '', s)
    return s[:30]


# 已知方圆系主体（含各子公司名），用于 OCR 残缺名模糊纠正
KNOWN_PARTIES = [
    '浙江方圆检测集团股份有限公司',
    '浙江方圆智能技术检测有限公司',
    '浙江方圆金属材料检测有限公司',
    '浙江方圆皮革轻纺检测认证有限公司',
    '绍兴方圆检测科技有限公司',
    '浙江金华方圆检测计量有限公司',
    '浙江方圆检测集团龙游有限公司',
    '杭州七格股份经济合作社',
    '杭州市钱塘区下沙街道七格集体经济组织',
    '杭州建友物流科技有限公司',
    '杭州亨德利文化创意有限公司',
    '杭州亨德利实业有限公司',
    '绍兴市柯桥区诗韵纺织绣品有限公司',
    '浙江海宁鹃湖科技城开发投资有限责任公司',
    '金华市计量质量科学研究院',
    '湖州市吴兴区埭溪镇人民政府',
    '浙江省市场监督管理局',
    '龙游新北综合能源开发利用有限公司',
    '龙游道一科技有限公司',
    '王祖勇',
]


def _match_known(name):
    """OCR 残缺主体名 → 已知主体库模糊纠正。返回 (纠正后的名, 匹配度)。匹配度<0.45 保持原样。"""
    if not name or len(name) < 3:
        return name, 0.0
    from difflib import SequenceMatcher
    best, best_r = name, 0.0
    for k in KNOWN_PARTIES:
        r = SequenceMatcher(None, name, k).ratio()
        if r > best_r:
            best, best_r = k, r
    return (best if best_r >= 0.45 else name), best_r


# ---------------------------------------------------------------- 异常条款检测
def _fnum2(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def detect_anomalies(all_rows):
    """同类合同对比 + 异常条款检测。all_rows: [主体, 名称, 出租方, 承租方, 标的, 面积,
    租期起, 租期止, 租期(年), 首年租金, 首年租金(万元), 口径, 单价, 递增, 频率, 免租, 押金, 备注, 源文件]
    返回 [(主体, 合同, 检测类型, 检测值, 参照, 提示)]。"""
    from collections import defaultdict
    out = []

    # 1) 单位年租金：同主体+同出租方组内偏离（元/㎡/年）
    groups = defaultdict(list)
    for r in all_rows:
        area = _fnum2(r[5])
        rw = r[10] if isinstance(r[10], (int, float)) else _fnum2(r[10])
        unit = rw * 10000 / area if (area and rw) else None
        groups[(r[0], r[2])].append({'row': r, 'unit': unit})
    for (comp, lessor), items in groups.items():
        units = sorted(i['unit'] for i in items if i['unit'])
        if len(units) >= 3:
            med = units[len(units) // 2]
            for i in items:
                u = i['unit']
                if u and (u > med * 2.5 or u < med * 0.4):
                    out.append((comp, i['row'][1], '单位年租金异常',
                                f'{u:,.0f} 元/㎡/年', f'同组中位 {med:,.0f}',
                                '显著偏离同类合同，核对面积/首年租金口径'))

    # 2) 租期：同主体组内偏离
    g2 = defaultdict(list)
    for r in all_rows:
        g2[r[0]].append((r, _fnum2(r[8])))
    for comp, items in g2.items():
        ys = sorted(y for _, y in items if y)
        if len(ys) >= 3:
            med = ys[len(ys) // 2]
            for r, y in items:
                if y and abs(y - med) > 0.5 and (y > med * 2 or y < med * 0.4):
                    out.append((comp, r[1], '租期偏离', f'{y} 年', f'同主体中位 {med} 年',
                                '与同类合同租期差异明显，关注资产是否同一处'))

    # 3) 递增幅度 ≥10%
    for r in all_rows:
        m = re.search(r'递增\s*([\d.]+)%', str(r[13] or ''))
        if m and _fnum2(m.group(1)) and float(m.group(1)) >= 10:
            out.append((r[0], r[1], '租金递增幅度大', str(r[13]), '多数合同无递增',
                        '年递增≥10%，关注商业条款合理性'))

    # 4) 免租期 ≥6 个月
    for r in all_rows:
        m = re.search(r'(\d+)\s*个月', str(r[15] or ''))
        if m and int(m.group(1)) >= 6:
            out.append((r[0], r[1], '免租期较长', str(r[15]), '—', '免租期≥6个月，关注是否合理'))

    # 5) 押金 ≥ 半年租金
    for r in all_rows:
        m = re.search(r'([\d,.]+)\s*万元', str(r[16] or ''))
        rw = r[10] if isinstance(r[10], (int, float)) else _fnum2(r[10])
        if m and rw:
            dv = float(m.group(1).replace(',', ''))
            if dv / rw >= 0.5:
                out.append((r[0], r[1], '押金比例高', f'{dv:,.0f} 万', f'首年租金 {rw:,.0f} 万',
                            '押金≥半年租金，关注资金占用'))

    # 6) 特殊商业条款
    for r in all_rows:
        note = str(r[17] or '')
        if any(k in note for k in ('三免三减半', '免租', '无偿', '补助')):
            out.append((r[0], r[1], '特殊条款', note[:50], '—', '存在免租/补助等特殊商业安排'))
    return out


# ---------------------------------------------------------------- 输出
HEADERS = ['主体', '合同名称', '出租方', '承租方', '租赁标的', '面积(㎡)', '租期起', '租期止',
           '租期(年)', '首年租金', '首年租金(万元)', '租金口径', '单价', '租金递增',
           '支付频率', '免租期', '押金/保证金', '备注', '源文件']

FONT = Font(name='Times New Roman', size=10)
HEAD_FILL = PatternFill('solid', fgColor='DDEBF7')
THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_summary_md(summary, md_path):
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write(f"# {summary['合同名称']}\n\n")
        f.write(f"- 源文件：`{summary['源文件']}`\n")
        f.write(f"- 出租方：{summary['出租方'] or '—'}\n")
        f.write(f"- 承租方：{summary['承租方'] or '—'}\n")
        f.write(f"- 租赁标的：{summary['租赁标的'] or '—'}\n")
        if summary['面积(㎡)']:
            f.write(f"- 面积：{summary['面积(㎡)']} ㎡\n\n")
        else:
            f.write("\n")
        f.write("## 租赁期与租金\n\n")
        f.write(f"- 租期：{summary['租期起'] or '—'} 至 {summary['租期止'] or '—'}（{summary['租期(年)']} 年）\n")
        f.write(f"- 首年租金：{summary['首年租金'] or '—'}（口径：{summary['租金口径'] or '—'}）\n")
        if summary['单价']:
            f.write(f"- 单价：{summary['单价']}\n")
        if summary['租金递增']:
            f.write(f"- 租金递增：{summary['租金递增']}\n")
        if summary['支付频率']:
            f.write(f"- 支付频率：{summary['支付频率']}\n")
        if summary['免租期']:
            f.write(f"- 免租期：{summary['免租期']}\n")
        if summary['押金/保证金']:
            f.write(f"- 押金/保证金：{summary['押金/保证金']}\n")
        if summary['备注']:
            f.write(f"- 备注：{summary['备注']}\n")
    return md_path


def write_excel(all_rows, xlsx_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '合同摘要汇总'
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(name='Times New Roman', size=10, bold=True)
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for row in all_rows:
        ws.append(list(row))
        for c in ws[ws.max_row]:
            c.font = FONT
            c.border = BORDER
            if c.column == 11:                       # 首年租金(万元) 数值列
                c.number_format = '#,##0.00'
            c.alignment = Alignment(vertical='center', wrap_text=True)
    widths = [12, 30, 24, 26, 30, 10, 12, 12, 9, 15, 14, 14, 14, 14, 12, 14, 12, 30, 24]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:{openpyxl.utils.get_column_letter(len(HEADERS))}{ws.max_row}'

    # ---- 合同异常检测 sheet ----
    ws2 = wb.create_sheet('合同异常检测')
    hdr2 = ['主体', '合同名称', '检测类型', '检测值', '同类参照', '提示']
    ws2.append(hdr2)
    for c in ws2[1]:
        c.font = Font(name='Times New Roman', size=10, bold=True)
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    anomalies = detect_anomalies(all_rows)
    for a in anomalies:
        ws2.append(list(a))
        for c in ws2[ws2.max_row]:
            c.font = FONT
            c.border = BORDER
            c.alignment = Alignment(vertical='center', wrap_text=True)
    n_contracts = len({(x[0], x[1]) for x in anomalies})
    ws2.append([])
    ws2.append([f'共检测 {len(all_rows)} 份合同，{len(anomalies)} 条提示（涉及 {n_contracts} 份合同）。其余合同无显著异常。'])
    for c in ws2[ws2.max_row]:
        c.font = FONT
    for i, w in enumerate([10, 26, 18, 20, 24, 44], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2.freeze_panes = 'A2'

    wb.save(xlsx_path)
    return xlsx_path


def write_summary_md_all(all_rows, md_path):
    with open(md_path, 'w', encoding='utf-8') as f:
        f.write('# 使用权资产租赁合同摘要汇总\n\n')
        f.write(f'共 {len(all_rows)} 份合同。\n\n')
        cur = None
        for row in all_rows:
            comp, name = row[0], row[1]
            if comp != cur:
                f.write(f'\n## {comp}\n\n')
                cur = comp
            f.write(f'1. **{name}**\n')
            if row[2]:
                f.write(f'   - 出租方：{row[2]}\n')
            f.write(f'   - 租期：{row[6] or "—"} 至 {row[7] or "—"}（{row[8]} 年）\n')
            if row[9]:
                f.write(f'   - 首年租金：{row[9]}（{row[11]}）\n')
            if row[12]:
                f.write(f'   - 单价：{row[12]}\n')
            if row[13]:
                f.write(f'   - 租金递增：{row[13]}\n')
            if row[14]:
                f.write(f'   - 支付频率：{row[14]}\n')
            if row[15]:
                f.write(f'   - 免租期：{row[15]}\n')
            if row[16]:
                f.write(f'   - 押金/保证金：{row[16]}\n')
            if row[17]:
                f.write(f'   - 备注：{row[17]}\n')
            f.write(f'   - 源文件：`{row[18]}`\n')
    return md_path


# ---------------------------------------------------------------- 主流程
def main(argv=None):
    ap = argparse.ArgumentParser(description='合同 OCR 结构化摘要 + 汇总')
    ap.add_argument('root', help='合同目录（含 _ocr_识别结果 子目录）')
    ap.add_argument('--out', default=None, help='输出目录（默认 <root>/合同摘要）')
    a = ap.parse_args(argv)

    ocr_dir = os.path.join(a.root, '_ocr_识别结果')
    if not os.path.isdir(ocr_dir):
        print(f'未找到 OCR 目录: {ocr_dir}')
        return 1
    txts = sorted(glob.glob(os.path.join(ocr_dir, '**', '*.pdf.txt'), recursive=True))
    if not txts:
        print('未找到合同 OCR 文本')
        return 1

    out_root = a.out or os.path.join(a.root, '合同摘要')
    all_rows = []
    for i, txt in enumerate(txts, 1):
        rel = os.path.relpath(txt, ocr_dir)
        comp = rel.split(os.sep)[0]                   # 主体 = 顶层目录
        s = extract_summary(txt)
        md_rel = rel[:-4] + '.md'                     # .pdf.txt → .pdf.md
        md_path = os.path.join(out_root, '每份合同摘要', md_rel)
        os.makedirs(os.path.dirname(md_path), exist_ok=True)
        write_summary_md(s, md_path)
        all_rows.append([comp] + list(s.values()))
        flag = '' if s['租期起'] and s['首年租金'] else ' ⚠缺失'
        print(f'  [{i}/{len(txts)}] {comp}/{s["合同名称"][:24]:<24} 租期 {s["租期起"]}~{s["租期止"]:<20} 首年 {s["首年租金"] or "?"}{flag}')
    all_rows.sort(key=lambda r: (r[0], r[6] or ''))

    xlsx_path = write_excel(all_rows, os.path.join(out_root, '合同摘要汇总.xlsx'))
    md_all_path = write_summary_md_all(all_rows, os.path.join(out_root, '合同摘要汇总.md'))
    print(f'\n完成：{len(txts)} 份合同')
    print(f'  Excel 汇总: {xlsx_path}')
    print(f'  Markdown 汇总: {md_all_path}')
    print(f'  每份合同摘要: {out_root}/每份合同摘要/')
    return 0


if __name__ == '__main__':
    sys.exit(main())
