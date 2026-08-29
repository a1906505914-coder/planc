# -*- coding: utf-8 -*-
"""SSS 适配器（2026-08-06）：把兵团水利水电 3003 账套的 U8 集团级导出
拆分成 FY 标准结构（每主体一年一文件：科目余额表 / 综合查询明细表 / 辅助核算余额表），
13 个 builder 零改动直接跑。

输入：<SSS 目录>（或指定数据目录）
  3003科目辅助余额表.xlsx    —— 科目×辅助核算×核算账簿 的 期初/本期借/本期贷/借累计/贷累计/期末（原币+本币），
                               一份文件同时承担 TB（科目合计行）+ aux（辅助核算行），201 个核算账簿混排
  3003序时账1.xlsx / 3003序时账2.xlsx —— 账套级 GL（双层表头，月|日|核算账簿名称|凭证号|分录号|摘要|科目编码|科目名称|辅助项|币种|借原币|借本币|贷原币|贷本币）
  公司标识表0130(1).xlsx     —— 账簿代码→公司映射（可选，用于主体名可读化）

输出：<输出目录>/ 下每主体三个标准文件
  NN{主体名}2025年科目余额表.xlsx     （表头 r1：科目编码|科目名称|期初方向|期初金额|本期借方|本期贷方|期末方向|期末金额|等级）
  NN{主体名}2025年综合查询明细表.xlsx （表头 r1：科目名称|日期|凭证字|凭证号|对方科目|摘要|借方金额|贷方金额）
  NN{主体名}2025年辅助核算余额表.xlsx （表头 r1：科目名称|往来单位名称|期初方向|期初金额|本期借方|本期贷方|期末方向|期末金额|往来类型）

清洗规则（同 jtt_prepare 的 U8 惯例）：
- 科目名：去『编码\\』前缀（1002\\银行存款 → 银行存款），保留中文层级
- 辅助项：『【客商：XXX】』→ 往来单位名称=XXX；无客商（项目/部门）→ 往来单位名称为空
- 对方科目：序时账无该列 → 留空（regen_all 终审 fill_voucher_counterparty 会按凭证翻面补全）
- 借/贷：取【本币】列；日期由 2025-月-日 拼接
- 等级：按编码长度推断（4位=1 / 6位=2 / 8位=3 / 10位=4）

用法：
  python sss_prepare.py <SSS 目录> [<输出目录>] [--books 账簿1 账簿2 ...]
  不指定 --books 时默认拆分全部账簿（201 个）；试点可只指定 2-3 个。
"""
import openpyxl, os, re, sys, io, argparse

YEAR = '2025'

# ---------- 通用清洗 ----------

def _num(v):
    """金额清洗：None/空/'None' → 0.0。"""
    if v is None:
        return 0.0
    if isinstance(v, str):
        s = v.strip()
        if s == '' or s.lower() == 'none':
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def clean_subj(raw):
    """科目名：去『编码\\』前缀，保留中文层级（1002\\银行存款\\结算中心 → 银行存款\\结算中心）。"""
    s = str(raw or '').strip()
    s = re.sub(r'^\d+\\', '', s)
    return s.strip()


def clean_book(raw):
    """核算账簿名清洗：去换行/去『-兵团第十一师账簿』等后缀，保留主体名。"""
    s = str(raw or '').strip().replace('\n', '').replace('\r', '').replace('\u3000', '')
    return s


def _dir_char(v):
    """方向列：借/贷/平；空 → 平。"""
    s = str(v or '').strip()
    return s if s in ('借', '贷', '平') else '平'


def _level_of(code):
    """等级 = 编码数字位数映射（4=1级 6=2级 8=3级 10=4级）。"""
    cn = re.sub(r'\D', '', str(code))
    n = len(cn)
    if n <= 4:
        return 1
    if n <= 6:
        return 2
    if n <= 8:
        return 3
    return 4


def _find_header_row(ws, kws, max_scan=8):
    """在表头区找含全部 kws 的行（双层表头取第一层）。返回行号(0基)或 None。"""
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True)):
        if not r:
            continue
        cells = [str(c) for c in r if c is not None]
        if cells and all(any(k in c for c in cells) for k in kws):
            return i
    return None


# ---------- 读源文件 ----------

def load_aux_balance(path):
    """3003科目辅助余额表 → (表头 idx, rows 列表)。
    双层表头：r0 标题 / r1 核算账簿信息 / r2 空 / r3-r4 表头。数据从 r5。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    # 表头行 = 含『科目编码』且『辅助核算』（r3/r4 双层，取含"科目编码"的上一层 r3）
    hdr = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True)):
        cells = [str(c) for c in r if c is not None]
        if cells and '科目编码' in cells[0] and '辅助核算' in cells:
            hdr = i
            break
    if hdr is None:
        hdr = 3
    rows = list(ws.iter_rows(min_row=hdr + 2, values_only=True))
    wb.close()
    return hdr, rows


def load_gl(path):
    """3003序时账 → rows 列表。双层表头 r3/r4，数据从 r5。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hdr = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True)):
        cells = [str(c) for c in r if c is not None]
        if cells and '凭证号' in cells and '科目编码' in cells:
            hdr = i
            break
    if hdr is None:
        hdr = 3
    rows = list(ws.iter_rows(min_row=hdr + 2, values_only=True))
    wb.close()
    return rows


# ---------- 主体发现 ----------

def discover_books(aux_rows):
    """从科目辅助余额表收集全部核算账簿（col3=核算账簿名称）。
    ⚡ 2026-08-15 修复：表头行『核算账簿名称』被误收为账簿（生成 176核算账簿名称 伪主体）——
    过滤表头字面量与空值。"""
    books = []
    seen = set()
    for r in aux_rows:
        if not r or len(r) < 4:
            continue
        b = clean_book(r[3])
        if not b or b == '核算账簿名称':
            continue
        if b not in seen:
            seen.add(b)
            books.append(b)
    return sorted(books)


# ---------- 输出标准三件套 ----------

def write_tb(path, book, aux_rows):
    """按账簿输出标准科目余额表：只取【科目合计行】（col1 名称为空、col0='编码\\名称科目合计'），
    避免与辅助明细行双计（合计行=该科目全部辅助行之和）。无合计行的科目退回明细行聚合。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '科目余额表'
    ws.append(['科目编码', '科目名称', '期初方向', '期初金额', '本期借方', '本期贷方',
               '期末方向', '期末金额', '等级'])
    sum_rows = {}   # code -> 合计行
    det_rows = {}   # code -> 明细行列表（无合计行时兜底聚合）
    for r in aux_rows:
        if not r or len(r) < 19:
            continue
        if clean_book(r[3]) != book:
            continue
        code_raw = str(r[0]).strip() if r[0] is not None else ''
        name_raw = str(r[1]).strip() if r[1] is not None else ''
        if not code_raw or not code_raw[0].isdigit():
            continue
        if name_raw:
            # 明细行：col0=纯编码、col1=名称、col2=辅助核算 → 归 aux，不参与 TB 合计
            code = code_raw
            det_rows.setdefault(code, []).append(r)
        else:
            # 合计行：col0='1002\银行存款科目合计' → 拆分编码+名称
            m = re.match(r'^(\d+)\\(.+?)(科目合计)?$', code_raw)
            if not m:
                continue
            code = m.group(1)
            sum_rows[code] = (r, clean_subj(m.group(2)))
    agg = {}
    for code, (r, name) in sum_rows.items():
        agg[code] = {
            'name': name, 'qc_fx': _dir_char(r[5]), 'qc': _num(r[7]),
            'jf': _num(r[9]), 'df': _num(r[11]), 'qm_fx': _dir_char(r[16]), 'qm': _num(r[18]),
        }
    # 无合计行的科目：明细行聚合兜底（保证 TB 覆盖）
    for code, rows in det_rows.items():
        if code in agg:
            continue
        a = {'name': '', 'qc_fx': '平', 'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm_fx': '平', 'qm': 0.0}
        for r in rows:
            if not a['name']:
                a['name'] = clean_subj(r[1])
            a['qc'] += _num(r[7]); a['jf'] += _num(r[9]); a['df'] += _num(r[11]); a['qm'] += _num(r[18])
            a['qc_fx'] = _dir_char(r[5]) or a['qc_fx']; a['qm_fx'] = _dir_char(r[16]) or a['qm_fx']
        agg[code] = a
    for code in sorted(agg, key=lambda c: (len(re.sub(r'\D', '', c)), c)):
        a = agg[code]
        ws.append([code, a['name'], a['qc_fx'], round(a['qc'], 2), round(a['jf'], 2),
                   round(a['df'], 2), a['qm_fx'], round(a['qm'], 2), _level_of(code)])
    wb.save(path)
    wb.close()
    return len(agg)


def write_aux(path, book, aux_rows):
    """按账簿输出标准辅助核算余额表：仅辅助核算行（有【客商：】提取往来单位；项目/部门行保留空往来单位）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '辅助核算余额表'
    ws.append(['科目名称', '往来单位名称', '期初方向', '期初金额', '本期借方', '本期贷方',
               '期末方向', '期末金额', '往来类型'])
    n = 0
    for r in aux_rows:
        if not r or len(r) < 19:
            continue
        if clean_book(r[3]) != book:
            continue
        code = str(r[0]).strip() if r[0] is not None else ''
        if not code or not code[0].isdigit():
            continue
        aux_raw = str(r[2] or '').strip()
        cust = ''
        m = re.search(r'【客商：(.+?)】', aux_raw)
        if m:
            cust = m.group(1).strip()
        # 无辅助核算（科目合计行）→ 不进 aux（TB 已承载）；有【项目/部门】但无客商 → 保留空往来单位行
        # （current_account 按『空保留』处理，供内部往来核对）
        if not aux_raw:
            continue
        name = clean_subj(r[1])
        qc_fx = _dir_char(r[5])
        qc = _num(r[7])
        jf = _num(r[9])
        df = _num(r[11])
        qm_fx = _dir_char(r[16])
        qm = _num(r[18])
        ws.append([name, cust, qc_fx, round(qc, 2), round(jf, 2), round(df, 2),
                   qm_fx, round(qm, 2), ''])
        n += 1
    wb.save(path)
    wb.close()
    return n


def write_gl(path, book, gl_rows):
    """按账簿输出标准综合查询明细表：取本币列；对方科目留空（终审翻面补全）。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '综合查询明细表'
    ws.append(['科目名称', '日期', '凭证字', '凭证号', '对方科目', '摘要', '借方金额', '贷方金额'])
    n = 0
    for r in gl_rows:
        if not r or len(r) < 16:
            continue
        if clean_book(r[2]) != book:
            continue
        code = str(r[6]).strip() if r[6] is not None else ''
        if not code or not code[0].isdigit():
            continue
        mon = str(r[0] or '').strip().zfill(2)
        day = str(r[1] or '').strip().zfill(2)
        date = f'{YEAR}-{mon}-{day}'
        vno = str(r[3] or '').strip()
        sm = str(r[5] or '').strip()
        name = clean_subj(r[7])
        jf = _num(r[11])   # col11 = 借方本币
        df = _num(r[13])   # col13 = 贷方本币
        if jf == 0 and df == 0:
            continue
        ws.append([name, date, '', vno, '', sm, round(jf, 2), round(df, 2)])
        n += 1
    wb.save(path)
    wb.close()
    return n


# ---------- 主流程 ----------

def prepare(src_dir, out_dir, books=None):
    aux_path = None
    gl_paths = []
    for fn in sorted(os.listdir(src_dir)):
        if not fn.endswith('.xlsx'):
            continue
        if '科目辅助余额表' in fn:
            aux_path = os.path.join(src_dir, fn)
        elif '序时账' in fn:
            gl_paths.append(os.path.join(src_dir, fn))
    if not aux_path:
        print(f'[ERROR] 未找到 科目辅助余额表：{src_dir}')
        return 1
    print(f'[INFO] 科目辅助余额表：{os.path.basename(aux_path)}；序时账 {len(gl_paths)} 个')
    hdr, aux_rows = load_aux_balance(aux_path)
    all_books = discover_books(aux_rows)
    print(f'[INFO] 核算账簿总数：{len(all_books)}')
    if books:
        sel = [b for b in all_books if b in books]
        print(f'[INFO] 指定账簿 {len(sel)} 个（共 {len(all_books)} 个可用）')
    else:
        sel = all_books
    if not sel:
        print('[ERROR] 无匹配账簿')
        return 1
    os.makedirs(out_dir, exist_ok=True)
    gl_all = []
    for gp in gl_paths:
        gl_all.extend(load_gl(gp))
    print(f'[INFO] 序时账合计 {len(gl_all)} 行')
    n_ok = 0
    for i, book in enumerate(sel, 1):
        ent = f'{i:02d}{re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", book)[:40]}'
        base = os.path.join(out_dir, f'{ent}{YEAR}年')
        n_km = write_tb(base + '科目余额表.xlsx', book, aux_rows)
        n_ax = write_aux(base + '辅助核算余额表.xlsx', book, aux_rows)
        n_gl = write_gl(base + '综合查询明细表.xlsx', book, gl_all)
        print(f'  [{i:3d}] {book[:36]:38s} TB={n_km:4d} aux={n_ax:5d} GL={n_gl:6d}')
        n_ok += 1
    print(f'[DONE] 已输出 {n_ok} 个主体标准三件套 -> {out_dir}')
    return 0


def main(argv=None):
    raw = list(sys.argv[1:] if argv is None else argv)
    ap = argparse.ArgumentParser(description='SSS 账套适配器：U8 集团级导出 → FY 标准三件套')
    ap.add_argument('src', help='SSS 原始数据目录')
    ap.add_argument('out', nargs='?', default=None, help='输出目录（默认 <src>/prepared）')
    ap.add_argument('--books', nargs='*', default=[], help='限定账簿（默认全部 201 个）')
    args = ap.parse_args(raw)
    src = args.src.strip().strip('"').strip("'")
    out = args.out or os.path.join(src, 'prepared')
    return prepare(src, out, books=set(args.books) or None)


if __name__ == '__main__':
    sys.exit(main())
