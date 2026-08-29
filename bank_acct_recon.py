# -*- coding: utf-8 -*-
"""银行开户清单 vs 账面银行存款账户核对（2026-08-16，铁律131f-ocr 扩展）。
数据源：
  ① 开户清单：银行资料 OCR 识别结果（_全量识别.json 或 txt 重新解析）
  ② 账面：AYL 科目余额表 1002 二级科目（银行网点名+尾号）
匹配：按【主体名】关联 + 【银行名规范化】比对（工行/农行/中行/建行…）。
输出：AYL\\底稿\\{year}\\银行开户账户核对_{year}.xlsx
  ① 按主体核对汇总  ② 逐账户明细（匹配/清单独有/账面独有）③ 口径说明
"""
import paths   # 数据路径中心锚点（数据绝对路径只允许出现在 paths.py）
import openpyxl
import glob
import json
import os
import re
import sys
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment

DATA_ROOT = paths.DATA_ROOT
ACCT = 'AYL'
YEAR = '2025'
OCR_DIR = os.path.join(DATA_ROOT, '银行资料', '_ocr_识别结果')
OUT = os.path.join(DATA_ROOT, ACCT, '底稿', str(YEAR))
os.makedirs(OUT, exist_ok=True)

# 银行名规范化：开户清单(全名) → 短名（与账面网点名比对用）
_BANK_SHORT = [
    ('中国工商银行', '工行'), ('工商银行', '工行'),
    ('中国农业银行', '农行'), ('农业银行', '农行'),
    ('中国建设银行', '建行'), ('建设银行', '建行'),
    ('中国银行', '中行'), ('中国邮政储蓄银行', '邮储'),
    ('交通银行', '交行'), ('招商银行', '招行'), ('浦发银行', '浦发'),
    ('中信银行', '中信'), ('兴业银行', '兴业'), ('民生银行', '民生'),
    ('光大银行', '光大'), ('华夏银行', '华夏'), ('广发银行', '广发'),
    ('平安银行', '平安'), ('浙商银行', '浙商'), ('恒丰银行', '恒丰'),
    ('宁波银行', '宁波'), ('杭州银行', '杭州'), ('渤海银行', '渤海'),
    ('汇丰银行', '汇丰'), ('渣打银行', '渣打'), ('上海浦东发展银行', '浦发'),
    ('农村商业银行', '农商'), ('泰隆商业银行', '泰隆'),
]


def norm_bank(b):
    """银行名 → 短名（'中国工商银行'→'工行'；OCR 截断'中国工商'→'工行'）。"""
    if not b:
        return ''
    # 完整匹配优先
    for full, short in _BANK_SHORT:
        if full in b:
            return short
    # OCR 截断前缀（'中国工商'/'中国农业'/'中国建设'/'中国银行'…）
    for pref, short in [('中国工商', '工行'), ('中国农业', '农行'), ('中国建设', '建行'),
                        ('中国银行', '中行'), ('中国邮政', '邮储'), ('中国交通', '交行'),
                        ('中国招商', '招行'), ('上海浦东', '浦发'), ('宁波银行', '宁波'),
                        ('杭州银行', '杭州'), ('浙商银行', '浙商'), ('农村商业', '农商')]:
        if b.startswith(pref) or pref in b:
            return short
    return b


def norm_ent(name):
    """主体名规范化（OCR 存款人全名 → AYL 主体短名匹配键）。"""
    if not name:
        return ''
    # 去公司后缀与地域括号
    n = re.sub(r'(股份|有限|责任|公司|集团|控股|管理|贸易|科技|材料|机械|风机|通风机)', '', name)
    n = re.sub(r'[（(].*?[)）]', '', n)
    n = re.sub(r'新型|新型材料|有限公司', '', n)
    return n.strip()


def load_book_accounts():
    """账面：AYL 科目余额表 1002 二级科目 → {主体: [(科目名, 期末)]}"""
    book = defaultdict(list)
    d = os.path.join(DATA_ROOT, ACCT, '数据', str(YEAR))
    for f in sorted(glob.glob(os.path.join(d, '*科目余额表' + str(YEAR) + '.xlsx'))):
        ent = os.path.basename(f).replace('科目余额表' + str(YEAR) + '.xlsx', '')
        wb = openpyxl.load_workbook(f, read_only=True)
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        for r in rows[2:]:
            code = str(r[0] or '')
            nm = str(r[1] or '')
            if code.startswith('1002') and len(code) > 4:
                try:
                    qm = float(r[-1] or 0)
                except (ValueError, TypeError):
                    qm = 0.0
                book[ent].append((nm, qm))
    return dict(book)


def load_ocr_accounts():
    """开户清单：OCR 识别结果 → [(存款人, 银行, 账号, 性质, 状态, 销户日期, 来源文件)]"""
    out = []
    for f in sorted(glob.glob(os.path.join(OCR_DIR, '*.txt'))):
        if os.path.basename(f).startswith('_'):
            continue
        try:
            text = open(f, encoding='utf-8').read()
        except Exception:
            continue
        if '清单' not in text:
            continue
        import bank_doc_parser as BP
        r = BP.parse_bank_list(text)
        # 文本级状态增强：清单全文含 '久悬'/'销户' 而账户级未识别 → 补标
        has_jx = '久悬' in text
        has_xh = '销户' in text
        for a in r['accounts']:
            if not a['account']:
                continue
            st = a['status']
            if not st:
                if has_xh:
                    st = '销户'
                elif has_jx:
                    st = '久悬'
            out.append({'depositor': r.get('depositor', ''), 'bank': a['bank'],
                        'account': a['account'], 'type': a['acct_type'],
                        'status': st, 'close_date': a.get('close_date', ''),
                        'src': os.path.basename(f)})
    return out


def match_main(text, bank):
    """账面科目名是否与清单银行匹配：账面科目名含 短名 或 银行全名关键词。"""
    short = norm_bank(bank)
    if not short:
        return False
    return short in text


def run_recon(acct='AYL', year='2025', ocr_dir=None, out_dir=None):
    """开户清单 vs 账面核对主程序（doc_router 可复用入口）。
    返回核对底稿路径。"""
    global DATA_ROOT, ACCT, YEAR, OCR_DIR, OUT
    ACCT = acct
    YEAR = str(year)
    if ocr_dir:
        OCR_DIR = ocr_dir
    if out_dir:
        OUT = out_dir
    else:
        OUT = os.path.join(DATA_ROOT, ACCT, '底稿', str(year))
    os.makedirs(OUT, exist_ok=True)
    book = load_book_accounts()
    ocr = load_ocr_accounts()
    # 主体映射：OCR 存款人 → 账面主体（含短名匹配）
    ents_book = list(book.keys())

    def find_book_ent(depositor, src=''):
        for be in ents_book:
            if be and (be in depositor or (norm_ent(be) and norm_ent(be) in depositor)):
                return be
        # 文件名辅助（OCR 存款人残缺时）：主体名片段在来源文件名（'富丽华'∈'…-富丽华.txt'）
        if src:
            for be in ents_book:
                if not be:
                    continue
                segs = [be[2:] if len(be) > 2 else be, be[:2], be[:4]]
                for seg in segs:
                    if len(seg) >= 2 and seg in src:
                        return be
        return None

    # 开户清单去重（同主体同账号保留首条）
    seen = set()
    ocr_uniq = []
    for a in ocr:
        k = (a['depositor'], a['account'])
        if k in seen:
            continue
        seen.add(k)
        ocr_uniq.append(a)
    ocr = ocr_uniq

    # ⚡ 多时点清单对比：同主体同账号在【旧清单】有、在【最新清单】无 → 疑似已注销
    #   按主体收集账号→来源文件年份（文件名含 2025/20260709）
    import re as _re
    acct_srcs = defaultdict(set)   # (depositor, account) → {年份/时点}
    for a in ocr:
        m = _re.search(r'(20\d{6})', a['src']) or _re.search(r'(2025|2026)', a['src'])
        acct_srcs[(a['depositor'], a['account'])].add(m.group(1) if m else '?')
    closed_hint = set()
    for k, srcs in acct_srcs.items():
        if len(srcs) >= 2:   # 出现在多份清单
            latest = max(srcs, key=lambda s: (len(s), s))
            if any(s != latest for s in srcs):
                # 该账号未出现在最新清单 → 已注销候选
                closed_hint.add(k)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '按主体核对汇总'
    hdr = ['主体', '开户清单账户数', '账面账户数', '匹配', '清单独有(账面无)', '账面独有(清单无)', '备注']
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='D9E2F3')
        c.alignment = Alignment(horizontal='center')

    ws2 = wb.create_sheet('逐账户明细')
    hdr2 = ['主体', '开户清单-银行', '开户清单-账号', '开户清单-性质', '开户清单-状态',
            '开户清单-销户日期', '账面-银行存款科目', '匹配结果', '来源文件']
    ws2.append(hdr2)
    for c in ws2[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='D9E2F3')
        c.alignment = Alignment(horizontal='center')

    # 按主体分组核对
    by_ent = defaultdict(list)
    for a in ocr:
        be = find_book_ent(a['depositor'], a['src'])
        a['book_ent'] = be
        by_ent[be or ('(未匹配主体)' + a['depositor'])].append(a)

    detail_rows = []
    total_match = total_only_ocr = total_only_book = 0
    for be in sorted(set(list(book.keys()) + list(by_ent.keys())), key=lambda x: (x is None, x)):
        ocr_list = by_ent.get(be, [])
        book_list = book.get(be, [])
        if be.startswith('(未匹配主体)'):
            # 清单有但找不到账面主体 → 全部清单独有
            n_ocr = len(ocr_list)
            total_only_ocr += n_ocr
            for a in ocr_list:
                detail_rows.append([a['depositor'], a['bank'], a['account'], a['type'], a['status'],
                                    '', '⚠️ 清单主体未匹配账面', a['src']])
            ws.append([be, n_ocr, 0, 0, n_ocr, 0, '开户清单主体在 AYL 账面无对应'])
            continue
        matched = 0
        only_ocr = []
        only_book = []
        for a in ocr_list:
            # 已注销/久悬标注
            flags = []
            if a.get('status') in ('销户', '久悬'):
                flags.append(f"{a['status']}")
            if a.get('close_date'):
                flags.append(f"销户日期{a['close_date']}")
            if (a['depositor'], a['account']) in closed_hint:
                flags.append('疑似已注销(最新清单已无)')
            flag_str = '/'.join(flags)
            hit = False
            for bnm, _ in book_list:
                if match_main(bnm, a['bank']):
                    hit = True
                    break
            if hit:
                matched += 1
                detail_rows.append([be, a['bank'], a['account'], a['type'], a['status'],
                                    a.get('close_date', ''),
                                    next((bnm for bnm, _ in book_list if match_main(bnm, a['bank'])), ''),
                                    ('✅ 匹配' + (f' [{flag_str}]' if flag_str else '')), a['src']])
            else:
                only_ocr.append(a)
                detail_rows.append([be, a['bank'], a['account'], a['type'], a['status'],
                                    a.get('close_date', ''),
                                    '', ('⚠️ 清单独有(账面未发现)' + (f' [{flag_str}]' if flag_str else '')), a['src']])
        for bnm, qm in book_list:
            hit = False
            for a in ocr_list:
                if match_main(bnm, a['bank']):
                    hit = True
                    break
            if not hit:
                only_book.append(bnm)
                detail_rows.append([be, '', '', '', '', '', bnm, '⚠️ 账面独有(清单未识别)', ''])
        n_ocr = len(ocr_list)
        n_book = len(book_list)
        total_match += matched
        total_only_ocr += len(only_ocr)
        total_only_book += len(only_book)
        ws.append([be, n_ocr, n_book, matched, len(only_ocr), len(only_book), ''])

    # 汇总行
    ws.append(['合计', '', '', total_match, total_only_ocr, total_only_book, ''])
    for r in ws.iter_rows(min_row=2):
        if r[0].value == '合计':
            for c in r:
                c.font = Font(bold=True)

    for row in detail_rows:
        ws2.append(row)
    # 口径说明
    ws3 = wb.create_sheet('口径说明')
    notes = [
        '银行开户清单 vs 账面银行存款账户核对（2026-08-16 铁律131f-ocr 扩展）',
        '数据源：',
        '  ① 开户清单 = 银行资料 17 份扫描件/照片 OCR 识别（bank_doc_parser 逐账户解析）',
        '  ② 账面 = AYL 科目余额表 1002 银行存款二级科目（银行网点名+尾号）',
        '匹配规则：主体名关联（OCR 存款人 ↔ AYL 主体）+ 银行名规范化比对（工行/农行/中行/建行/宁波…）',
        '说明：',
        '  · 匹配 = 清单账户银行在账面有对应银行存款二级科目',
        '  · 清单独有 = 开户清单有、账面无对应科目（可能销户未及时清理/新开未入账，需查证）',
        '  · 账面独有 = 账面有、开户清单未识别（清单可能未覆盖该行/OCR 漏识别）',
        '  · 账面科目名不带完整账号（仅网点名+尾号），清单有账号——本版按银行名核对，',
        '    后续若账面能提供账号（序时账/辅助核算）可升级为账号级精确匹配',
        '  · OCR 识别存在形近字错别字，核对结果需人工复核后使用',
    ]
    for i, n in enumerate(notes, 1):
        ws3.cell(row=i, column=1, value=n)
    ws3.column_dimensions['A'].width = 120

    for wsn, widths in [(ws, [22, 14, 12, 8, 12, 12, 30]), (ws2, [20, 22, 22, 12, 10, 34, 18, 24])]:
        for i, w in enumerate(widths, 1):
            wsn.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for wsn in (ws, ws2):
        wsn.freeze_panes = 'A2'

    out_path = os.path.join(OUT, f'银行开户账户核对_{YEAR}.xlsx')
    wb.save(out_path)
    print(f'✅ 已生成: {out_path}')
    print(f'   开户清单账户 {len(ocr)} | 账面账户 {sum(len(v) for v in book.values())} | '
          f'匹配 {total_match} | 清单独有 {total_only_ocr} | 账面独有 {total_only_book}')
    return out_path


if __name__ == '__main__':
    run_recon()
