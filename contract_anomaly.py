# -*- coding: utf-8 -*-
"""contract_anomaly.py —— 借款合同非格式条款/异常条款识别

基于全量 OCR 识别文本，扫描每份合同的特殊约定、异常要素、手写标注等，
输出关注清单（审计视角）。

用法:
  python contract_anomaly.py [识别文本目录] [输出报告路径]
"""
import os
import re
import sys
import glob

OCR_DIR = sys.argv[1] if len(sys.argv) > 1 else \
    r'd:/底稿测试/ADF/数据/2026/借款合同_ocr/识别文本_脱敏'
OUT_PATH = sys.argv[2] if len(sys.argv) > 2 else \
    r'd:/底稿测试/ADF/数据/2026/借款合同_异常条款清单.md'

# —— 异常/非格式条款关键词（按关注度分组）——
KEYWORDS = [
    ('特别约定/附加', ['特别约定', '特约', '附加', '额外约定', '另行约定', '双方另行', '补充约定', '补充协议']),
    ('展期/续贷/变更', ['展期', '续贷', '延期', '贷款展期', '合同变更', '协议变更', '要素调整', '重新签订']),
    ('减免/豁免/优惠', ['减免', '豁免', '免除', '减免罚息', '不上浮', '利率优惠', '优惠利率', '让利']),
    ('担保异常', ['追加担保', '更换担保', '保证人变', '反担保', '担保范围变更', '无限连带', '最高额担保之外', '补充担保']),
    ('资金监管/放款条件', ['专款专用', '资金监管', '受托支付', '放款条件', '用款计划', '资金流向', '账户监管']),
    ('违约/罚息/提前还款', ['提前还款违约金', '提前还款', '罚息', '复利', '逾期利息', '扣收', '划扣', '强制扣款']),
    ('特殊商务安排', ['保证金', '服务费', '手续费', '顾问费', '财务顾问', '中间业务', '代扣代缴', '对赌', '回购']),
    ('关联/三方', ['关联企业', '关联方', '第三方', '共同借款', '联合借款', '第三人提供']),
]

# —— 借款要素异常阈值 ——
RATE_LO, RATE_HI = 2.0, 9.0        # 年利率%正常参考区间
AMOUNT_LO = 1e6                     # 低于100万（小额）关注

# 利率提取：支持 5.85%、年利率5.85、千分之x、万分之x、LPR+xxbp
_RATE_RE = re.compile(
    r'(?:年利率|利率|执行利率|借款利率)[^0-9%]*?'
    r'(\d+(?:\.\d+)?)\s*(%|％|‰|‰)')
_RATE2_RE = re.compile(r'(\d+(?:\.\d+)?)\s*(%|％)\s*[^。]{0,6}(?:年|一年|贷款期限)')
_AMT_RE = re.compile(r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*(万元|万|元|亿元|亿)')


def _norm_amt(text, num, unit):
    """金额数字→元"""
    try:
        n = float(num.replace(',', ''))
    except ValueError:
        return None
    if unit in ('万元', '万'):
        return n * 10000
    if unit in ('亿元', '亿'):
        return n * 1e8
    return n


def analyze(txt, common_kws=None):
    """单份合同 → 异常/关注项列表
    common_kws: 全库高频（格式条款）关键词，命中只标※不展开。"""
    common_kws = common_kws or set()
    notes = []

    # 1) 关键词段落扫描
    for label, kws in KEYWORDS:
        hits, common_n = [], 0
        for kw in kws:
            for m in re.finditer(re.escape(kw), txt):
                if kw in common_kws:
                    common_n += 1
                    continue
                s = max(0, m.start() - 30)
                e = min(len(txt), m.end() + 50)
                snippet = re.sub(r'\s+', '', txt[s:e])
                hits.append(f'「{kw}」…{snippet[:70]}')
        if hits:
            notes.append((label, list(dict.fromkeys(hits))[:5]))
        elif common_n:
            notes.append((label + '※', [f'（格式条款，全库高频出现 {common_n} 处）']))

    # 2) 利率异常（排除罚息/加收/上浮/挪用等加收比例语境）
    rates = []
    for m in _RATE_RE.finditer(txt):
        ctx = txt[max(0, m.start() - 12):m.end() + 8]
        if re.search(r'罚息|加收|上浮|挪用|逾期|基点', ctx):
            continue
        num = float(m.group(1))
        sym = m.group(2)
        if sym in ('%', '％'):
            rates.append(num)          # 已是年利率%口径
        elif sym == '‰':                # 千分之 → 除以10
            rates.append(num / 10)
    for m in _RATE2_RE.finditer(txt):
        ctx = txt[max(0, m.start() - 12):m.end() + 8]
        if re.search(r'罚息|加收|上浮|挪用|逾期|基点', ctx):
            continue
        rates.append(float(m.group(1)))
    rates = [r for r in rates if 0.1 < r < 60]
    if rates:
        rmax, rmin = max(rates), min(rates)
        if rmax > RATE_HI:
            notes.append(('利率偏高', [f'出现年利率约 {rmax}% (> {RATE_HI}%) 需关注']))
        if rmin < RATE_LO:
            notes.append(('利率偏低', [f'出现年利率约 {rmin}% (< {RATE_LO}%) 需关注']))
        if len(set(round(r, 2) for r in rates)) >= 3:
            notes.append(('利率多样', [f'同一合同出现 {len(set(round(r,2) for r in rates))} 种利率，核对分段计息']))

    # 3) 金额异常（小额/超大额）
    amts = []
    for m in _AMT_RE.finditer(txt):
        v = _norm_amt(txt, m.group(1), m.group(2))
        if v:
            amts.append(v)
    if amts:
        amin, amax = min(amts), max(amts)
        if amin < AMOUNT_LO:
            notes.append(('小额借款', [f'单笔低至 {amin/10000:.2f} 万元，关注资金用途']))
        if amax >= 1e8:
            notes.append(('大额借款', [f'单笔达 {amax/1e8:.2f} 亿元，关注授信/审批']))

    # 4) 特殊符号/手写标注痕迹
    marks = []
    for pat in [r'（[^）]{0,20}(?:注|说明|补充|手写)[^）]{0,20}）',
                r'\([^)]{0,20}(?:注|说明|补充)[^)]{0,20}\)',
                r'(?:以下空白|无正文|本页无正文|以此为准)']:
        for m in re.finditer(pat, txt):
            marks.append(m.group(0))
    if marks:
        notes.append(('标注/批注痕迹', list(dict.fromkeys(marks))[:5]))

    return notes


def main():
    files = sorted(glob.glob(os.path.join(OCR_DIR, '**', '*.txt'),
                             recursive=True))
    print(f'扫描识别文本: {len(files)} 份')
    texts = {}
    for fp in files:
        with open(fp, 'r', encoding='utf-8', errors='replace') as f:
            texts[fp] = f.read()

    # —— 第一遍：统计关键词命中合同数，判定格式条款（出现率 ≥ 40%）——
    n_files = max(len(files), 1)
    kw_freq = {}
    for txt in texts.values():
        for _, kws in KEYWORDS:
            for kw in kws:
                if kw in txt:
                    kw_freq[kw] = kw_freq.get(kw, 0) + 1
    common_kws = {kw for kw, n in kw_freq.items()
                  if n / n_files >= 0.40}

    lines = ['# 借款合同 非格式条款/异常条款 关注清单',
             f'\n> 基于 {OCR_DIR} 全量识别文本 | 格式条款(※)为全库高频项，已折叠',
             '',
             '| # | 合同 | 异常项数 | 要点（※=格式条款） |',
             '|---|------|--------|------|']
    details = ['', '---', '', '## 明细（仅含非格式/异常项）']
    total_flag = 0
    for idx, fp in enumerate(files, 1):
        rel = os.path.relpath(fp, OCR_DIR).replace('.txt', '')
        notes = analyze(texts[fp], common_kws)
        if not notes:
            continue
        total_flag += 1
        brief = '; '.join(f'{l}({len(v)})' for l, v in notes)
        lines.append(f'| {total_flag} | {rel} | {len(notes)} | {brief} |')
        details.append(f'\n### {rel}')
        for label, vals in notes:
            details.append(f'\n- **{label}**')
            for v in vals:
                details.append(f'  - {v}')
    lines.append(f'\n**共 {total_flag} 份合同存在关注条款（全量 {len(files)} 份）**')
    os.makedirs(os.path.dirname(OUT_PATH) or '.', exist_ok=True)
    with open(OUT_PATH, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines + details))
    print(f'完成: {total_flag}/{len(files)} 份含异常项 → {OUT_PATH}')
    print(f'格式条款关键词(出现率≥40%): {sorted(common_kws)}')


if __name__ == '__main__':
    main()
