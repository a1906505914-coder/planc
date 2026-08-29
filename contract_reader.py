# -*- coding: utf-8 -*-
"""合同识别模块（铁律131f-contract，2026-08-16 用户需求："我只要提供合同即可"）。

目标：把【合同文件】自动识别为结构化台账（与 project_manifest.contracts 同构），
输入合同 → 输出 JSON → 项目制收入底稿 + 合同资产账龄专项自动生效。

管线（分层，各层可单独调用/替换）：
  ① 文本提取 extract_text(path)
       · PDF   → pypdf 文本层；无文本层（扫描件）→ pymupdf 渲染页面 → RapidOCR
       · 图片  → RapidOCR（照片/扫描件，中文识别率高、CPU 快）
       · Word  → python-docx（段落+表格）
       · Excel → openpyxl（合同台账表）
       · txt/md → 直接读
  ② 规则抽取 extract_by_rules(text) —— 正则+关键词，输出确定性字段
       · 客户名称（发包人/甲方/业主/采购方）
       · 合同内容/名称（标题/工程名称）
       · 工期起止（开工/竣工/工期）
       · 合同金额（含税识别：'含税'/'含增值税'/'税率'）
       · 结算周期（月结/季结/进度款/验收后…）→ settle_cycle
       · 质保期（质保/保修/质量保证金…）→ quality_months
       · 预算总成本/成本（预算价/投标价/成本…）
  ③ AI 辅助兜底 ai_enhance(text, partial) —— 规则未命中字段走大模型（预留接口，
     用户配置 AI 回调后启用；未配置时返回规则结果 + 提示缺字段）
  ④ 台账生成 build_ledger(entries) —— 输出与 project_manifest.contracts 同构 JSON
     {'project': 业主名, 'customer':…, 'content':…, 'start':…, 'end':…,
      'contract_amount':…, 'tax_inclusive':…, 'budget_cost':…, 'method':…,
      'settle_cycle':…, 'quality_months':…, 'pay_terms':…}

用法：
  python contract_reader.py <合同文件...> [--out 台账.json] [--merge]
  python contract_reader.py --scan <目录> [--out 台账.json]   # 批量识别
"""
import argparse
import json
import os
import re
import sys

# ============================================================
# ① 文本提取
# ============================================================

# OCR 识别文本缓存目录（⚠️ 只读【本地脱敏版】，先运行 contract_mask_local.py 生成）
_OCR_CACHE_DIR = r'd:/底稿测试/ADF/数据/2026/借款合同_ocr/识别文本_脱敏'
_CONTRACT_BASE_DIR = r'd:/底稿测试/ADF/数据/2026/借款合同/化纤本年度新增借款合同'


def _read_ocr_cache(path):
    """从已有 OCR 识别文本读取（避免重复 OCR）；无命中返回 None。
    ⚠️ 脱敏版目录已替换为 借X/贷X（contract_mask_local.mask_path），路径需映射。"""
    try:
        ap = os.path.abspath(path)
        base = os.path.abspath(_CONTRACT_BASE_DIR)
        if ap.startswith(base):
            rel = os.path.relpath(ap, base)
            try:
                import contract_mask_local as CML
                cache_rel = CML.mask_path(rel + '.txt')
            except Exception:
                cache_rel = rel + '.txt'
            cache = os.path.join(_OCR_CACHE_DIR, cache_rel)
            if os.path.isfile(cache):
                with open(cache, 'r', encoding='utf-8', errors='replace') as f:
                    return f.read()
    except Exception:
        pass
    return None


# OCR 形近字纠错词典（扫描件常见错字——宁可错纠为已知词，不保留错字进台账）
_OCR_FIXES = [
    ('若美县', '若羌县'), ('若羌县', '若羌县'),          # 若羌→若美（最典型）
    ('已美县', '且末县'), ('铁门美', '铁门关'), ('铁门类', '铁门关'),
    ('巴音郭愣', '巴音郭楞'), ('巴音郭楞蒙古自治州', '巴音郭楞蒙古自治州'),
    ('图木舒美', '图木舒克'), ('图木舒克市', '图木舒克市'),
    ('阿拉而', '阿拉尔'), ('阿拉尔市', '阿拉尔市'),
    ('昆玉市', '昆玉市'), ('昆五市', '昆玉市'),
    ('胡杨河市', '胡杨河市'), ('新星市', '新星市'),
    ('水电利', '水利'), ('水水利', '水利'), ('水电水', '水利水电'),
    ('若差', '若羌'), ('若姜', '若羌'),
    ('壹阡', '壹仟'), ('壹任', '壹仟'), ('壹千', '壹仟'),
    ('贰佰', '贰佰'), ('陆伯', '陆佰'), ('叁佰', '叁佰'), ('肆佰', '肆佰'), ('伍佰', '伍佰'),
    ('柒佰', '柒佰'), ('捌佰', '捌佰'), ('玖佰', '玖佰'), ('拾亿', '拾亿'),
]


def _ocr_fix_text(text):
    """OCR 文本纠错：应用形近字词典（顺序替换，长词优先）。"""
    for bad, good in sorted(_OCR_FIXES, key=lambda x: -len(x[0])):
        if bad in text:
            text = text.replace(bad, good)
    return text


_OCR_ENGINE = None          # RapidOCR 懒加载单例
_IMG_EXTS = ('.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp')


class _EasyOCREngine:
    """easyocr 引擎包装（兼容 rapidocr 接口）：__call__(img) → (result, None)。
    ⚡ 由环境变量 CONTRACT_OCR_EASYOCR=1 触发；模型在 ~/.EasyOCR/model（已手动下载）。"""

    def __init__(self):
        _lib = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.easyocr_libs')
        if os.path.isdir(_lib) and _lib not in sys.path:
            sys.path.insert(0, _lib)
        import easyocr
        self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False, verbose=False)

    def __call__(self, img):
        res = self.reader.readtext(img)
        return [[list(map(list, r[0])), r[1], float(r[2])] for r in res], None


def _ocr_engine():
    """懒加载 OCR 引擎。
    ⚡ 默认优先 easyocr（CPU 单页约 4.7s，比 RapidOCR 快约 30%，模型已就绪；批量场景引擎只加载一次）；
       CONTRACT_OCR_RAPIDOCR=1 → 强制 RapidOCR（GPU 加速或需要它时）；
    ⚡ RapidOCR 分支：GPU（.ocr_gpu/）优先，CONTRACT_OCR_CPU=1 强制 CPU（多进程并行 GPU CUDA 冲突卡死）。"""
    global _OCR_ENGINE
    if _OCR_ENGINE is None:
        if not os.environ.get('CONTRACT_OCR_RAPIDOCR'):
            try:
                _OCR_ENGINE = _EasyOCREngine()
                return _OCR_ENGINE
            except Exception as ex:
                print(f'  ⚠️ easyocr 加载失败({ex})，回退 RapidOCR')
        use_gpu = not os.environ.get('CONTRACT_OCR_CPU')
        _gpu_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.ocr_gpu')
        if use_gpu and os.path.isdir(_gpu_dir) and _gpu_dir not in sys.path:
            sys.path.insert(0, _gpu_dir)   # 优先加载 onnxruntime-gpu（CUDA）
        from rapidocr_onnxruntime import RapidOCR
        if use_gpu:
            try:
                _OCR_ENGINE = RapidOCR(Global={'use_cuda': True})   # GPU 优先
            except Exception:
                _OCR_ENGINE = RapidOCR()                            # GPU 不可用 → CPU 兜底
        else:
            _OCR_ENGINE = RapidOCR()                                # 多进程并行：强制 CPU
    return _OCR_ENGINE


def ocr_image(img_path, max_pages=20):
    """图片/扫描页 → 文本（RapidOCR + 形近字纠错）。返回识别文本；失败返回 ''。"""
    try:
        engine = _ocr_engine()
        result, _ = engine(img_path)
        if not result:
            return ''
        lines = [ln[1] for ln in result if len(ln) > 1 and ln[1]]
        return _ocr_fix_text('\n'.join(lines))
    except Exception as ex:
        print(f'  ⚠️ OCR 失败 {os.path.basename(img_path)}: {ex}')
        return ''


def ocr_image_array(img_arr):
    """numpy BGR/RGB 数组 → 文本（内存版，不落盘；绕开 safe-delete 文件删除失败）。"""
    try:
        engine = _ocr_engine()
        result, _ = engine(img_arr)
        if not result:
            return ''
        lines = [ln[1] for ln in result if len(ln) > 1 and ln[1]]
        return _ocr_fix_text('\n'.join(lines))
    except Exception as ex:
        print(f'  ⚠️ OCR 数组失败: {ex}')
        return ''


def extract_text(path):
    """合同文件 → 纯文本。返回 (text, fmt)。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.pdf',):
        return _extract_pdf(path), 'pdf'
    if ext in _IMG_EXTS:
        return ocr_image(path), 'image'
    if ext in ('.docx', '.doc'):
        return _extract_docx(path), 'docx'
    if ext in ('.xlsx', '.xls'):
        return _extract_xlsx(path), 'xlsx'
    if ext in ('.txt', '.md', '.json'):
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            return f.read(), 'txt'
    return '', ext


def _extract_pdf(path):
    """PDF 文本提取：先读 OCR 缓存；无缓存→pypdf 文本层；空→扫描件内存 OCR。"""
    # ① 缓存优先：识别文本已存在 → 直接复用（秒级，不重复 OCR）
    cached = _read_ocr_cache(path)
    if cached:
        return cached
    # ② pypdf 文本层（文本型 PDF）
    try:
        from pypdf import PdfReader
        r = PdfReader(path)
        joined = '\n'.join((pg.extract_text() or '') for pg in r.pages)
    except Exception:
        joined = ''
    # 文本层有效（合同文本一般 >100 字符）→ 直接用
    if joined and len(joined.strip()) > 80:
        return joined
    # ③ 扫描件 → 内存渲染 OCR（不落盘，绕开 safe-delete）
    try:
        try:
            import fitz  # pymupdf
        except ImportError:
            import pymupdf as fitz  # PyMuPDF 1.24+ 模块名改为 pymupdf
        import numpy as np
        doc = fitz.open(path)
        n = min(len(doc), 20)
        ocr_parts = []
        for i in range(n):
            pix = doc[i].get_pixmap(dpi=200)
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n)
            img = arr[:, :, :3] if pix.n >= 3 else arr
            t = ocr_image_array(img)
            if t.strip():
                ocr_parts.append(t)
        return '\n'.join(ocr_parts)
    except Exception as ex:
        return joined or f'[扫描PDF OCR失败: {ex}]'


def _extract_docx(path):
    try:
        import docx
        d = docx.Document(path)
        parts = [p.text for p in d.paragraphs if p.text.strip()]
        for tbl in d.tables:
            for row in tbl.rows:
                parts.append(' | '.join(c.text.strip() for c in row.cells))
        return '\n'.join(parts)
    except Exception as ex:
        return f'[Word提取失败: {ex}]'


def _extract_xlsx(path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        parts = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if cells:
                    parts.append(' | '.join(cells))
        wb.close()
        return '\n'.join(parts)
    except Exception as ex:
        return f'[Excel提取失败: {ex}]'


# ============================================================
# ② 规则抽取
# ============================================================

# 客户名称：甲方/发包人/业主/采购方/委托人（冒号后取值，去'公司/集团'后缀截断）
_CUST_PATTERNS = [
    r'(?:发包人|甲方|业主|采购方|委托方|定作方|建设方)\s*[：:]\s*([^\n，。;；]{2,60})',
    r'客户名称\s*[：:]\s*([^\n，。;；]{2,60})',
]
# 工程/合同名称（标题行或 工程名称：）
_NAME_PATTERNS = [
    r'工程名称\s*[：:]\s*([^\n，。;；]{2,80})',
    r'项目名称\s*[：:]\s*([^\n，。;；]{2,80})',
    r'合同名称\s*[：:]\s*([^\n，。;；]{2,80})',
]
# 日期：2024年3月1日 / 2024-03-01 / 2024.03.01
_DATE = r'(\d{4})\s*[年.\-/]\s*(\d{1,2})\s*[月.\-/]\s*(\d{1,2})\s*日?'
_START_PATTERNS = [
    r'(?:开工日期|开工时间|计划开工|合同开工)\s*[：:]\s*' + _DATE,
    r'(?:自)\s*' + _DATE + r'\s*(?:起|开工)',
]
_END_PATTERNS = [
    r'(?:竣工日期|完工日期|计划竣工|合同竣工|完工时间)\s*[：:]\s*' + _DATE,
    r'(?:至|到)\s*' + _DATE + r'\s*(?:止|竣工|完工)',
    r'(?:工期)\s*[：:]\s*.*?' + _DATE + r'\s*(?:止|结束|完工)',
]
# 金额：人民币 X 元 / ¥X / X万元 / 合同价 X（数字或中文大写）；单位=捕获组2
# ⚡ 中文大写金额自身可含'万/亿'（壹仟贰佰万元=12,000,000），故 数字串 与 中文串 分开处理：
#   数字串 后单位独立捕获；中文串 贪婪含万/亿/元
_CN_MONEY = r'([零壹贰叁肆伍陆柒捌玖拾佰仟万亿]+)(?:元|圆|整)?'
_MONEY = r'([0-9][0-9,，.]*)\s*((?:万元|亿|元|圆)?)'
_AMT_PATTERNS = [
    r'(?:合同金额|合同价款|合同总价|合同价|工程总价|合同总额|合同造价)\s*[：:为是]?\s*[（(]?\s*(?:人民币|人民币（大写）)?[（(]?\s*(?:' + _MONEY + '|' + _CN_MONEY + ')',
    r'(?:人民币|¥|￥|RMB)\s*(?:' + _MONEY + '|' + _CN_MONEY + ')',
]
# 中文大写数字 → 元（壹仟贰佰万→1200万→12000000）
_CN_DIGITS = {'零': 0, '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6,
              '柒': 7, '捌': 8, '玖': 9, '拾': 10, '佰': 100, '仟': 1000,
              '万': 10000, '亿': 100000000}


def _cn_money_to_num(s):
    """中文大写金额 → 元。壹仟贰佰万→12,000,000；贰仟叁佰伍拾万零陆佰→23,500,600。"""
    if not s:
        return None
    total = 0
    section = 0
    num = 0
    for ch in s:
        if ch in '零壹贰叁肆伍陆柒捌玖':
            num = _CN_DIGITS[ch]
        elif ch in '拾佰仟':
            section += (num if num else 1) * _CN_DIGITS[ch]
            num = 0
        elif ch == '万':
            section = (section + num) * 10000
            total += section
            section = 0
            num = 0
        elif ch == '亿':
            section = (section + num) * 100000000
            total += section
            section = 0
            num = 0
    total += section + num
    return total if total else None
# 含税识别
_TAX_PATTERNS = [r'含税', r'含增值税', r'含(?:发票|税款)', r'税率']
# 结算周期（⚠ 按"结算/支付"优先，'进度款按月计量'≠结算周期——后者只是计量频率）
_CYCLE_PATTERNS = [
    (r'(?:合同价款|工程款|结算款).{0,8}按季(?:结算|支付)', '季结'),
    (r'按季(?:结算|支付)', '季结'),
    (r'半年(?:结算|支付)', '半年结'),
    (r'按年(?:结算|支付)', '年结'),
    (r'(?:竣工|完工).{0,6}(?:结算|支付)', '竣工结算'),
    (r'(?:合同价款|工程款|结算款).{0,8}按月(?:结算|支付)', '月结'),
    (r'按月(?:结算|支付)', '月结'),
    (r'月度结算|月结算', '月结'),
]
# 质保期
_QUALITY_PATTERNS = [
    (r'质保期\s*[：:为]?\s*(\d+)\s*(?:个?月)', 'm'),
    (r'质保期\s*[：:为]?\s*(\d+)\s*年', 'y'),
    (r'保修期\s*[：:为]?\s*(\d+)\s*(?:个?月)', 'm'),
    (r'保修期\s*[：:为]?\s*(\d+)\s*年', 'y'),
    (r'质量保证金.*?(\d+)\s*%', 'pct'),
]
# 预算成本（复用 _MONEY/_CN_MONEY 双分支）
_BUDGET_PATTERNS = [
    r'(?:预算(?:总)?成本|预计总成本|投标价|中标价|成本(?:预算)?)\s*[：:为是]?\s*[（(]?\s*(?:人民币)?[（(]?\s*(?:' + _MONEY + '|' + _CN_MONEY + ')',
]
# 确认方法
_METHOD_PATTERNS = [
    (r'完工百分比|投入法|成本法', '投入法(完工百分比)'),
    (r'工作量法|产出法', '产出法(工作量)'),
]


def _cn2num(s):
    """中文数字 → int（用于质保期'一年'）。"""
    m = {'一': 1, '两': 2, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6,
         '七': 7, '八': 8, '九': 9, '十': 10, '半': 0.5}
    if s in m:
        return m[s]
    if '十' in s:
        a, _, b = s.partition('十')
        return (m.get(a, 1) * 10 + m.get(b, 0)) if b else m.get(a, 1) * 10
    return None


def _money_val(m):
    """兼容两种结构：数字串 (group1 数字, group2 单位) / 中文大写 (group3)。"""
    g3 = m.group(3) if m.re.groups >= 3 and m.group(3) else None
    if g3:
        n = _cn_money_to_num(g3)
        return float(n) if n else None
    raw = m.group(1).replace(',', '').replace('，', '').replace('。', '.')
    n = float(raw)
    try:
        unit = m.group(2) or ''
    except IndexError:
        unit = ''
    if '万' in unit:
        n *= 1e4
    elif '亿' in unit:
        n *= 1e8
    return n


def _norm_date(ymd):
    """(y, m, d) → 'YYYY-MM-DD'。"""
    y, m, d = int(ymd[0]), int(ymd[1]), int(ymd[2])
    return f'{y:04d}-{m:02d}-{d:02d}'


def extract_by_rules(text):
    """规则抽取 → 部分字段 dict（未命中=空）。"""
    out = {}
    t = text.replace('\u3000', ' ')
    # 客户名称（首个命中；去掉常见尾缀）
    for p in _CUST_PATTERNS:
        m = re.search(p, t)
        if m:
            v = m.group(1).strip().rstrip('。，,;；')
            # 截断到'公司/集团/中心/局/处/所'后的标点
            out['customer'] = re.split(r'[，,;；。]', v)[0].strip()
            break
    # 合同内容
    for p in _NAME_PATTERNS:
        m = re.search(p, t)
        if m:
            out['content'] = re.split(r'[，,;；。]', m.group(1).strip())[0].strip()
            break
    # 工期起止
    for p in _START_PATTERNS:
        m = re.search(p, t)
        if m:
            out['start'] = _norm_date(m.groups())
            break
    for p in _END_PATTERNS:
        m = re.search(p, t)
        if m:
            out['end'] = _norm_date(m.groups())
            break
    # 合同金额（含税识别）
    for p in _AMT_PATTERNS:
        m = re.search(p, t)
        if m:
            out['contract_amount'] = round(_money_val(m), 2)
            break
    out['tax_inclusive'] = bool(re.search('|'.join(_TAX_PATTERNS), t))
    # 结算周期
    for pat, label in _CYCLE_PATTERNS:
        if re.search(pat, t):
            out['settle_cycle'] = label
            break
    # 质保期
    for pat, kind in _QUALITY_PATTERNS:
        m = re.search(pat, t)
        if m:
            if kind == 'm':
                out['quality_months'] = int(m.group(1))
            elif kind == 'y':
                out['quality_months'] = int(m.group(1)) * 12
            elif kind == 'pct':
                out['quality_pct'] = float(m.group(1))
            break
    if 'quality_months' not in out:
        m = re.search(r'质保期\s*[：:为]?\s*(一|二|两|三|半)?\s*年', t)
        if m:
            v = _cn2num(m.group(1) or '一') or 1
            out['quality_months'] = int(v * 12)
    # 预算成本
    for p in _BUDGET_PATTERNS:
        m = re.search(p, t)
        if m:
            out['budget_cost'] = round(_money_val(m), 2)
            break
    # 确认方法
    for pat, label in _METHOD_PATTERNS:
        if re.search(pat, t):
            out['method'] = label
            break
    return out


# ============================================================
# ③ AI 辅助兜底（预留接口）
# ============================================================

def ai_enhance(text, partial, ai_call=None):
    """规则未命中字段 → AI 补充。ai_call: func(contract_text, missing_fields) -> dict。
    未配置 ai_call → 返回原结果 + 提示缺字段（诚实，不伪造）。"""
    if ai_call is None:
        return partial, [k for k, v in partial.items() if v in (None, '', 0)]
    missing = [k for k, v in partial.items() if v in (None, '', 0)]
    if not missing:
        return partial, []
    try:
        extra = ai_call(text, missing) or {}
        partial.update({k: extra[k] for k in missing if k in extra})
    except Exception as ex:
        print(f'  ⚠️ AI 抽取失败: {ex}')
    return partial, [k for k, v in partial.items() if v in (None, '', 0)]


# ============================================================
# ④ 台账生成
# ============================================================

def build_ledger(entries):
    """entries=[{文件, 规则字段, 缺字段, ...}] → 与 project_manifest.contracts 同构 JSON。
    project 键=业主名（可后续按账套覆盖规则）。"""
    out = []
    for e in entries:
        r = e.get('rule', {})
        rec = {
            'project': r.get('customer', '') or e.get('file', ''),
            'customer': r.get('customer', ''),
            'content': r.get('content', ''),
            'start': r.get('start', ''),
            'end': r.get('end', ''),
            'contract_amount': r.get('contract_amount', 0) or None,
            'tax_inclusive': r.get('tax_inclusive', False),
            'budget_cost': r.get('budget_cost', 0) or None,
            'method': r.get('method', ''),
            'settle_cycle': r.get('settle_cycle', ''),
            'quality_months': r.get('quality_months', 0) or None,
            'pay_terms': e.get('pay_terms', ''),
        }
        # 去掉全空字段（保持 JSON 干净）
        out.append({k: v for k, v in rec.items() if v not in (None, '', False, 0)})
    return out


def process_files(paths, ai_call=None, verbose=True):
    """识别一组合同文件 → (ledger, report)。"""
    entries = []
    for p in paths:
        if not os.path.exists(p):
            print(f'  ⚠️ 文件不存在: {p}')
            continue
        text, fmt = extract_text(p)
        if not text or text.startswith('['):
            print(f'  ⚠️ {os.path.basename(p)} 文本提取失败/无文本层（{fmt}）')
            continue
        rule = extract_by_rules(text)
        rule, missing = ai_enhance(text, rule, ai_call)
        entries.append({'file': os.path.basename(p), 'rule': rule,
                        'missing': missing, 'fmt': fmt})
        if verbose:
            print(f'  ✅ {os.path.basename(p)} [{fmt}] 识别: '
                  f'客户={rule.get("customer", "")[:16] or "?"} '
                  f'金额={rule.get("contract_amount", "?")} '
                  f'周期={rule.get("settle_cycle", "?")} '
                  f'质保={rule.get("quality_months", "?")}月'
                  + (f' 缺:{",".join(missing)}' if missing else ''))
    return build_ledger(entries), entries


def main(argv=None):
    ap = argparse.ArgumentParser(description='合同识别：合同文件 → 结构化台账 JSON')
    ap.add_argument('paths', nargs='*', help='合同文件或目录')
    ap.add_argument('--scan', action='store_true', help='批量识别目录内合同文件')
    ap.add_argument('--out', default=None, help='输出台账 JSON 路径')
    ap.add_argument('--merge', action='store_true', help='与现有台账合并')
    a = ap.parse_args(argv)

    files = []
    if a.scan:
        for d in a.paths:
            if os.path.isdir(d):
                for root, _, fs in os.walk(d):
                    for f in fs:
                        if f.lower().endswith(('.pdf', '.docx', '.doc', '.xlsx', '.txt', '.md', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp')):
                            files.append(os.path.join(root, f))
    else:
        for p in a.paths:
            if os.path.isdir(p):
                for f in os.listdir(p):
                    if f.lower().endswith(('.pdf', '.docx', '.doc', '.xlsx', '.txt', '.md', '.jpg', '.jpeg', '.png', '.bmp', '.tif', '.tiff', '.webp')):
                        files.append(os.path.join(p, f))
            else:
                files.append(p)
    if not files:
        print('未找到合同文件。用法：python contract_reader.py <合同.pdf/docx> [--out 台账.json]')
        return 1

    print(f'识别 {len(files)} 个合同文件…')
    ledger, entries = process_files(files)
    print(f'\n识别完成：{len(ledger)} 条。')

    out = a.out
    if not out:
        out = os.path.join(os.path.dirname(os.path.abspath(__file__)), '合同台账_识别结果.json')
    if a.merge and os.path.exists(out):
        try:
            old = json.load(open(out, encoding='utf-8'))
            old.extend(ledger)
            ledger = old
        except Exception:
            pass
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(ledger, f, ensure_ascii=False, indent=2)
    print(f'台账已写出: {out}')
    print('\n缺字段汇总（需人工补充或接入 AI）：')
    miss = {}
    for e in entries:
        for k in e.get('missing', []):
            miss[k] = miss.get(k, 0) + 1
    for k, n in sorted(miss.items(), key=lambda x: -x[1]):
        print(f'  {k}: {n} 条')
    return 0


if __name__ == '__main__':
    sys.exit(main())
