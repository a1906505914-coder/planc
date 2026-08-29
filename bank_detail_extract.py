# -*- coding: utf-8 -*-
"""银行账户资料提取程序。
扫描指定文件夹（默认桌面 亿利达/银行资料），从 .xlsx / .pdf 中提取
开户银行名称、账号、账户名称、账户性质、开户日期、账户状态、销户日期、久悬日期，
输出统一格式的 Excel 汇总表。

用法：
  拖文件夹到程序上  或
  python bank_detail_extract.py "<银行资料目录>"
"""

import paths as P
import os, re, sys
from collections import OrderedDict

# ==================== 输出列定义 ====================
OUTPUT_COLS = [
    ('核算主体', 18),
    ('开户银行', 22),
    ('账号', 24),
    ('账户名称', 24),
    ('币种', 10),
    ('账户性质', 18),
    ('开户日期', 14),
    ('账户状态', 10),
    ('销户日期', 14),
    ('久悬日期', 14),
    ('数据来源', 12),
]

# ==================== XLSX 解析（爱绅科技格式） ====================
def _parse_xlsx(path, ent_name):
    """解析 xlsx 银行账户清单。如 爱绅科技.xlsx 格式：
       组织, 银行, 账户号码, 账户名称, 开户银行, 币种, 账户性质
    """
    import openpyxl
    rows = []
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[wb.sheetnames[0]]
        header = None
        for r in ws.iter_rows(values_only=True):
            if not r or all(v is None for v in r):
                continue
            if header is None:
                # 探测表头
                txt = '|'.join(str(c or '') for c in r)
                if any(kw in txt for kw in ['账户', '账号', '开户', '组织', '银行']):
                    header = [str(c or '') for c in r]
                    continue
            if header:
                vals = [str(c or '') for c in (r + (None,) * 20)[:len(header)]]
                row = {'ent': ent_name}
                for i, h in enumerate(header):
                    h = h.strip()
                    v = vals[i].strip()
                    if '组织' in h or '主体' in h or '单位' in h:
                        row['ent'] = v or ent_name
                    if '开户银行' in h:
                        row['bank_detail'] = v
                    elif '银行' in h and '开户' not in h and '账号' not in h and '币种' not in h:
                        row['bank'] = v
                    elif '账户号码' in h or '账号' in h:
                        row['acct_no'] = v
                    elif '账户名称' in h:
                        row['acct_name'] = v
                    elif '币种' in h:
                        row['currency'] = v
                    elif '账户性质' in h or ('性质' in h and '账' in h):
                        row['type'] = v
                    elif '开户日期' in h or '开户日' in h:
                        row['open_date'] = v
                    elif '状态' in h:
                        row['status'] = v
                    elif '销户' in h:
                        row['close_date'] = v
                    elif '久悬' in h:
                        row['dormant_date'] = v
                if row.get('acct_no'):
                    rows.append(row)
        wb.close()
    except Exception as ex:
        print(f'  ⚠️ xlsx 解析失败 {os.path.basename(path)}：{ex}')
    return rows


# ==================== PDF 解析 ====================
_OCR_READER = None  # 延迟初始化

def _get_ocr_reader():
    global _OCR_READER
    if _OCR_READER is None:
        try:
            # ⚡ GPU 加速：若存在独立安装的 easyocr+torch（.easyocr_libs/），优先注入并启用 CUDA
            _gpu_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.easyocr_libs')
            if os.path.isdir(_gpu_dir) and _gpu_dir not in sys.path:
                sys.path.insert(0, _gpu_dir)
            import easyocr
            print('  ⏳ 正在加载 OCR 模型（GPU，首次需下载模型约 1-2 分钟）…')
            _OCR_READER = easyocr.Reader(['ch_sim', 'en'], gpu=True)
            print('  ✅ OCR 模型加载完成（GPU）')
        except Exception as ex:
            print(f'  ⚠️ OCR 模型加载失败：{ex}')
    return _OCR_READER


def _ocr_pdf(path):
    """用 OCR 从扫描件 PDF 中提取 (文字, y_center, x_center, y0, x0, y1, x1) 元组列表。
       保留包围盒坐标以便后续按空间位置重建表格行。
       使用 CLAHE 预处理 + 250 DPI 平衡速度和识别率。"""
    try:
        import fitz, cv2, numpy as np
        from PIL import Image
        import io
        doc = fitz.open(path)
        all_blocks = []
        for pi in range(doc.page_count):
            page = doc[pi]
            pix = page.get_pixmap(dpi=250)
            img = pix.tobytes('png')
            pil_img = Image.open(io.BytesIO(img)).convert('L')
            arr = np.array(pil_img)
            clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(arr)
            reader = _get_ocr_reader()
            if not reader:
                return None
            # detail=1 返回 (bbox, text, confidence)
            raw = reader.readtext(enhanced, detail=1, paragraph=False)
            for bbox, text, conf in raw:
                if conf < 0.3:
                    continue
                t = text.strip()
                if not t:
                    continue
                x0, y0 = bbox[0][0], bbox[0][1]
                x1, y1 = bbox[2][0], bbox[2][1]
                y_c = (y0 + y1) / 2
                x_c = (x0 + x1) / 2
                all_blocks.append((t, y_c, x_c, y0, x0, y1, x1))
        doc.close()
        return all_blocks
    except Exception as ex:
        print(f'  ⚠️ OCR 处理失败：{ex}')
        return None


def _ocr_image(path):
    """从图片文件（JPG/PNG）直接 OCR，返回同 _ocr_pdf 格式。"""
    try:
        reader = _get_ocr_reader()
        if not reader:
            return None
        raw = reader.readtext(path, detail=1, paragraph=False)
        blocks = []
        for bbox, text, conf in raw:
            if conf < 0.3:
                continue
            t = text.strip()
            if not t:
                continue
            x0, y0 = bbox[0][0], bbox[0][1]
            x1, y1 = bbox[2][0], bbox[2][1]
            y_c = (y0 + y1) / 2
            x_c = (x0 + x1) / 2
            blocks.append((t, y_c, x_c, y0, x0, y1, x1))
        return blocks
    except Exception as ex:
        print(f'  ⚠️ 图片 OCR 失败：{ex}')
        return None


def _reconstruct_table_rows(blocks, page_height=1200):
    """将 OCR 块按 Y 坐标分组重建表格行。返回 [{row_y, cells:[{text, x, w, y}]}]。"""
    if not blocks:
        return []
    # 计算文字高度：排除噪声块（>80px 大概率是OCR拼合错误）后取中位数
    raw_h = [b[6] - b[3] for b in blocks if 5 < (b[6] - b[3]) <= 80]
    median_h = sorted(raw_h)[len(raw_h)//2] if raw_h else 30
    Y_TOLERANCE = max(median_h * 1.3, 20)  # 行高 1.3 倍
    sorted_b = sorted(blocks, key=lambda x: x[1])  # 按 y_center 排序
    rows = []
    cur_row = []
    cur_y = None
    for item in sorted_b:
        t, y_c, x_c, y0, x0, y1, x1 = item
        if cur_y is None:
            cur_y = y_c
        if abs(y_c - cur_y) > Y_TOLERANCE:
            # 换行
            if cur_row:
                rows.append(cur_row)
            cur_row = [(t, x_c, x1 - x0, y0, x0)]
            cur_y = y_c
        else:
            cur_row.append((t, x_c, x1 - x0, y0, x0))
    if cur_row:
        rows.append(cur_row)

    # 2) 每行内按 X 排序，合并水平相邻块
    out_rows = []
    for cells in rows:
        cells.sort(key=lambda c: c[1])  # 按 x_center 排序
        merged = []
        i = 0
        while i < len(cells):
            t, x_c, w, y0, x0 = cells[i]
            # 尝试向右合并相邻的块（间距小于半字宽）
            while i + 1 < len(cells):
                t2, x_c2, w2, y0_2, x0_2 = cells[i + 1]
                gap = x0_2 - (x0 + w)
                if gap < 0.4 * w:  # 间距小于 0.4 个字符宽度 → 合并
                    t += t2 if t2 else ''
                    w = (x0_2 + w2) - x0
                    i += 1
                else:
                    break
            merged.append((t, x_c, w, y0, x0))
            i += 1
        if merged:
            out_rows.append(merged)
    return out_rows


def _classify_cell(text):
    """智能判断 OCR 单元格属于哪个字段。返回字段名或 None。"""
    t = text.strip()
    if not t:
        return None
    # 账号：纯数字或含字母数字混排，8-30 位
    if re.match(r'^[A-Za-z0-9\-/]{8,30}$', t) and re.search(r'\d{5,}', t):
        return 'acct_no'
    # 日期：YYYYMMDD 或 YYYY-MM-DD 或 YYYY/MM/DD
    if re.match(r'^\d{4}[-/]?\d{1,2}[-/]?\d{1,2}$', t):
        return 'open_date'
    # 账户状态：常见的几个词
    if t in ('正常', '销户', '久悬', '冻结', '睡眠', '正常销户'):
        return 'status'
    # ���行名：包含"银行"或"支行"或"信用社"(含OCR片段)
    if any(kw in t for kw in ['银行', '支行', '分行', '营业部', '信用社',
                               '行股份', '农村商业', '农商', '村镇银行',
                               '商业银行', '银行股']):
        return 'bank'
    # 币种：RMB/CNY/USD/港��/美元/人民币 等
    if t in ('RMB', 'CNY', 'USD', 'HKD', 'EUR', 'GBP', 'AUD', 'JPY',
             '人民币', '美元', '港币', '欧元', '英镑', '澳元', '日元',
             '人民币元'):
        return 'currency'
    return None


def _fuzzy_contains(text, keywords):
    """模糊匹配：检查 text 是否包含 keywords 中的任意一个（容忍 OCR 错别字）。"""
    for kw in keywords:
        # 逐字匹配（容忍 1-2 个字符差异）
        matches = 0
        for ch in kw:
            if ch in text:
                matches += 1
        if matches >= len(kw) - 1 and len(kw) >= 2:
            return True
        if kw in text:
            return True
    return False


def _parse_scanned_by_patterns(blocks, ent_name):
    """从扫描件的 OCR 块中按行重建→每行独立匹配账号及其他字段。"""
    if not blocks:
        return []

    rows = _reconstruct_table_rows(blocks)
    records = []

    for cells in rows:
        full_text = ''.join(c[0] for c in cells).replace('|', '')
        # 跳过噪声行
        if not full_text.strip():
            continue
        if any(kw in full_text for kw in ['已开立', '存款人', '扫描', '创建', '序号', '账户性质',
                                           '销户曰', '曰期', '账号']):
            continue
        if re.match(r'^\s*\d{1,3}\s*$', full_text):
            continue

        # 在该行所有单元格中找最长数字串（≥8位）→ 账号
        all_digits = re.findall(r'\d{8,}', full_text)
        if not all_digits:
            continue
        acct_val = max(all_digits, key=len)
        if len(acct_val) > 25:
            acct_val = acct_val[:25]  # 太长的可能是两列拼合

        bank_val = ''
        name_val = ''
        type_val = ''
        date_val = ''
        status_val = ''

        for t, x_c, w, y0, x0 in cells:
            text = t.strip('| ').strip()
            if not text or text.replace('-', '').replace('/', '').isdigit():
                continue  # 纯数字跳过
            # 银行名
            if any(kw in text for kw in ['银行', '支行', '分行', '营业部', '信用社',
                                          '行股份', '农村商业', '农商', '村镇银行',
                                          '商业银行', '银行股']):
                bank_val = text if not bank_val else bank_val
                continue
            # 状态
            if text in ('正常', '销户', '久悬', '撤销'):
                status_val = text
                continue
            # 类型
            if any(kw in text for kw in ['基本存款', '一般存款', '专用存款', '临时存款',
                                          '结算户', '般存款', '存款账']):
                tv = text.replace('般存款账', '一般存款账户').replace('用存款账', '一般存款账户')
                type_val = tv if not type_val else type_val
                continue
            # 日期
            dt_match = re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', text)
            if dt_match and not date_val:
                date_val = dt_match.group(1)
                continue
            # 长中文文本 → 户名
            cn_cnt = sum(1 for c in text if '\u4e00' <= c <= '\u9fff')
            if cn_cnt >= 6 and not name_val:
                name_val = text.lstrip('| ').strip()

        if acct_val:
            records.append({
                'ent': ent_name, 'bank': bank_val.lstrip('|').strip(),
                'acct_no': acct_val,
                'acct_name': name_val.lstrip('|').strip(),
                'currency': '', 'type': type_val,
                'open_date': date_val, 'status': status_val,
                'close_date': '', 'dormant_date': ''
            })

    # 去重
    seen = set()
    deduped = []
    for r in records:
        key = r['acct_no']
        if key not in seen:
            seen.add(key)
            deduped.append(r)
    return deduped


def _parse_pdf(path, ent_name):
    """提取 PDF 中的银行账户表格。
    先试 markitdown（仅对文字型PDF有效）；
    失败则用 OCR 重建表格行 → 列值匹配。"""
    rows = []
    text = ''

    # 先试 markitdown
    try:
        from markitdown import MarkItDown
        md = MarkItDown()
        result = md.convert(path)
        text = (result.text_content or '').strip()
    except Exception:
        pass

    lines = [l.strip() for l in text.split('\n') if l.strip() and '扫描全能王' not in l and '创建' not in l]
    if lines and len(''.join(lines)) >= 30:
        # 文字型 PDF → 通过空白分隔解析
        for line in lines:
            if re.match(r'^\d{1,3}$', line):
                continue
            if any(kw in line for kw in ['序号', '开户银行', '账号', '账户名称']):
                continue
            cols = re.split(r'\s{2,}|\t', line)
            if len(cols) >= 3:
                row = {'ent': ent_name, 'bank': '', 'acct_no': '', 'acct_name': '',
                       'currency': '', 'type': '', 'open_date': '', 'status': '',
                       'close_date': '', 'dormant_date': ''}
                if len(cols) >= 9:
                    row['bank'] = cols[1].strip()
                    row['acct_no'] = cols[2].strip()
                    row['acct_name'] = cols[3].strip()
                    for idx, key in [(4, 'type'), (5, 'open_date'), (6, 'status'),
                                     (7, 'close_date'), (8, 'dormant_date')]:
                        if idx < len(cols) and cols[idx].strip():
                            row[key] = cols[idx].strip()
                elif len(cols) >= 5:
                    row['bank'] = cols[0].strip()
                    row['acct_no'] = cols[1].strip()
                    row['acct_name'] = cols[2].strip()
                    row['type'] = cols[3].strip()
                    row['status'] = cols[4].strip()
                if row.get('acct_no') and len(row['acct_no'].strip()) >= 5:
                    rows.append(row)
        if rows:
            return rows

    # markitdown 失败或无文本 → OCR
    print(f'  🔍 扫描件/图片，启动 OCR（约 2-5 分钟）…')
    blocks = _ocr_pdf(path)
    if not blocks:
        print(f'  ⚠️ OCR 也失败')
        rows.append({'ent': ent_name, 'bank': '', 'acct_no': '（OCR失败，需人工录入）',
                     'acct_name': '', 'currency': '', 'type': '', 'open_date': '',
                     'status': '', 'close_date': '', 'dormant_date': ''})
        return rows

    # 重建表格行
    import fitz
    doc = fitz.open(path)
    page_height = doc[0].rect.height if doc.page_count > 0 else 1200
    # 转换为以页面高度为参考的 Y 坐标
    # fitz 用 72 DPI => 点坐标；OCR 用像素坐标需比例归一化
    # 但我们统一在 OCR 像素坐标系（dpi=200）下聚类
    # 估算页面在 200 DPI 下的像素高度
    pix_h = round(page_height / 72 * 200)  # points -> pixels at 200dpi
    doc.close()
    table_rows = _reconstruct_table_rows(blocks, page_height=pix_h)

    recs = _parse_scanned_by_patterns(blocks, ent_name)
    if recs:
        print(f'  ✅ OCR 解析到 {len(recs)} 条银行账户记录')
        return recs

    print(f'  ⚠️ 无法从扫描件中提取结构化表格数据')
    rows.append({'ent': ent_name, 'bank': '', 'acct_no': '（扫描件，需人工录入）',
                 'acct_name': '', 'currency': '', 'type': '', 'open_date': '',
                 'status': '', 'close_date': '', 'dormant_date': ''})
    return rows


def _parse_image(path, ent_name):
    """从图片文件（JPG/PNG）提取银行账户表格。"""
    print(f'  🔍 图片 OCR（约 1-2 分钟）…')
    blocks = _ocr_image(path)
    if not blocks:
        print(f'  ⚠️ 图片 OCR 失败')
        return [{'ent': ent_name, 'bank': '', 'acct_no': '（OCR失败，需人工录入）',
                 'acct_name': '', 'currency': '', 'type': '', 'open_date': '',
                 'status': '', 'close_date': '', 'dormant_date': ''}]
    # 图像高 = 最大 y 边界
    max_y = max(b[6] for b in blocks)  # y1
    recs = _parse_scanned_by_patterns(blocks, ent_name)
    if recs:
        print(f'  ✅ OCR 解析到 {len(recs)} 条银行账户记录')
        return recs
    print(f'  ⚠️ 无法从图片中提取结构化表格数据')
    return [{'ent': ent_name, 'bank': '', 'acct_no': '（扫描件，需人工录入）',
             'acct_name': '', 'currency': '', 'type': '', 'open_date': '',
             'status': '', 'close_date': '', 'dormant_date': ''}]


# ==================== 实体名推测 ====================
def _guess_ent(fn):
    """从文件名推测核算主体名称。"""
    base = fn.replace('.xlsx', '').replace('.pdf', '').replace('.PDF', '')
    # 去掉常见前缀
    base = re.sub(r'^[三二][.、＋+ \s]*4[.、＋+ \s]*', '', base)
    base = re.sub(r'^\d+[.、]', '', base)
    # 策略1：取最后一个 -/— 之后的内容
    for sep in ['—', '-', '－', '——', '--']:
        if sep in base:
            return base.rsplit(sep, 1)[-1].strip('-_—. ()（）、\t\n\r ')
    # 策略2：取 "已开立" 之前的内容（如 三、4、三进已开立...）
    m = re.search(r'已开立', base)
    if m:
        before = base[:m.start()].strip('-_—. ()（）、\t\n\r ')
        if before:
            return before
        # 已开立在开头 → 取之后的内容（如 已开立银行账户清单 铁城信息 → 铁城信息）
        after = base[m.end():].strip('-_—. ()（）、\t\n\r ')
        after = re.sub(r'银行[^ ]*(结算|账户|开户)?[^ ]*清单', '', after)
        after = re.sub(r'\d{6,}', '', after)
        after = after.strip('-_—. ()（）、\t\n\r ')
        if after:
            return after
    # 策略3：去掉已知模板词
    base = re.sub(r'已开立.*', '', base)
    base = re.sub(r'银行结算(开户|账户)清单', '', base)
    base = re.sub(r'开立银行户清单', '', base)
    base = re.sub(r'\d{6,}', '', base)
    base = re.sub(r'20\d{2}', '', base)
    base = re.sub(r'(?<!\d)[01]\d月?(?!\d)', '', base)
    base = base.strip('-_—. ()（）、\t\n\r ')
    return base if base and len(base) >= 2 else fn.split('.')[0]


# ==================== 主流程 ====================
def process_folder(data_dir, out_path=None):
    """扫描 data_dir 下的银行资料文件，生成汇总 Excel。"""
    if not os.path.isdir(data_dir):
        print(f'ERROR: 文件夹不存在：{data_dir}')
        return

    if out_path is None:
        out_path = os.path.join(data_dir, '银行账户汇总表_生成.xlsx')

    # 发现所有数据文件
    files = []
    img_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif')
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith('~$') or '银行账户汇总表' in fn:
            continue
        lo = fn.lower()
        if lo.endswith('.xlsx'):
            files.append(('xlsx', fn))
        elif lo.endswith('.pdf'):
            files.append(('pdf', fn))
        elif any(lo.endswith(e) for e in img_exts):
            files.append(('img', fn))

    all_rows = []
    seen_ents = set()
    for ftype, fn in files:
        path = os.path.join(data_dir, fn)
        ent = _guess_ent(fn)
        seen_ents.add(ent)
        print(f'[{ftype.upper():5s}] {fn} → {ent}')

        if ftype == 'xlsx':
            rows = _parse_xlsx(path, ent)
        elif ftype == 'pdf':
            rows = _parse_pdf(path, ent)
        else:
            rows = _parse_image(path, ent)
        all_rows.extend(rows)
        print(f'  → {len(rows)} 条记录')

    # 输出 Excel
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '银行账户汇总'

    # 标题
    ws.cell(1, 1, f'银行账户资料汇总（共 {len(seen_ents)} 个核算主体，{len(all_rows)} 条记录）')
    ws.cell(1, 1).font = Font(name='Times New Roman', bold=True, size=10, color='000000')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(OUTPUT_COLS))
    ws.row_dimensions[1].height = 22

    # 表头
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    BFONT = Font(name='Times New Roman', bold=True, size=10)
    THIN = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for j, (col_name, col_w) in enumerate(OUTPUT_COLS, 1):
        c = ws.cell(3, j, col_name)
        c.font = BFONT; c.fill = HFILL; c.alignment = Alignment(horizontal='center')
        c.border = BORDER
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = col_w

    # 数据
    NUMFMT = '#,##0.00'
    for i, row in enumerate(all_rows, 4):
        ws.cell(i, 1, row.get('ent', '')).border = BORDER
        ws.cell(i, 2, row.get('bank', '')).border = BORDER
        ws.cell(i, 3, row.get('acct_no', '')).border = BORDER
        ws.cell(i, 4, row.get('acct_name', '')).border = BORDER
        ws.cell(i, 5, row.get('currency', '')).border = BORDER
        ws.cell(i, 6, row.get('type', '')).border = BORDER
        ws.cell(i, 7, row.get('open_date', '')).border = BORDER
        ws.cell(i, 8, row.get('status', '')).border = BORDER
        ws.cell(i, 9, row.get('close_date', '')).border = BORDER
        ws.cell(i, 10, row.get('dormant_date', '')).border = BORDER
        ws.cell(i, 11, 'xlsx' if row.get('acct_no', '').startswith('（') else 'PDF提取').border = BORDER

    ws.freeze_panes = 'A4'

    # 防锁处理
    try:
        wb.save(out_path)
        print(f'\n✅ 已生成：{out_path}')
        print(f'   共 {len(seen_ents)} 个核算主体，{len(all_rows)} 条银行账户记录')
        for ent in sorted(seen_ents):
            cnt = sum(1 for r in all_rows if r.get('ent') == ent)
            print(f'   - {ent}：{cnt} 条')
    except PermissionError:
        # 被腾讯文档锁定则写隔壁
        alt = out_path.replace('.xlsx', '_新.xlsx')
        wb.save(alt)
        print(f'\n✅ 已生成（原文件被锁定，另存为）：{alt}')
    wb.close()
    return out_path


def main():
    if len(sys.argv) > 1:
        data_dir = sys.argv[1]
    else:
        data_dir = input('请输入银行资料文件夹路径：').strip().strip('"')
    process_folder(data_dir)


if __name__ == '__main__':
    main()
