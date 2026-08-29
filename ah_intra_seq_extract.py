# -*- coding: utf-8 -*-
"""AH 88家 关联往来序时账预抽取（2026-08-18）。

背景：差异穿透检查时每次都要回到全量序时账（88账套×7月）里翻凭证，
耗时且容易漏。本脚本预先扫描全部序时账，把『往来单位命中集团内公司名』
的行抽取出来，落成独立台账《关联往来序时账_2026.xlsx》，
后续差异定位直接查该表。

识别规则：
  · 序时账『文本』或『客户/供应商』列含集团内公司名（公司代码.xlsx 归一后长名优先包含匹配）
  · 排除 同公司自我匹配（如 1010 文本"杭氧集团"且本公司=1010 → 可能是内部凭证但非往来，保留但标注）
输出字段：公司/日期/凭证编号/科目/文本/客户/供应商/借方金额/贷方金额/对方公司名/备注

性能安全：序时账 sheet1.xml 解压后可达 247MB，openpyxl 读取会触发沙箱 OOM 强杀（静默 exit 1）。
故用 zipfile + ElementTree.iterparse 流式解析，逐行处理、用完即弃。
"""
import os
import re
import glob
import sys
import zipfile
import gc
from collections import defaultdict

import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)
import paths as P

PREP = os.path.join(P.DATA_DIRS['AH'], '中间产物', 'prepared')
CODE_FP = os.path.join(PREP, '公司代码.xlsx')
SEQ_DIR = os.path.join(P.DATA_DIRS['AH'], '数据', '2026', '序时账')
LOG_FP = os.path.join(APP_DIR, 'seq_extract.log')

F = Font(name='微软雅黑', size=10)
FH = Font(name='微软雅黑', size=10, bold=True)
FILL = PatternFill('solid', fgColor='DDEBF7')
THIN = Border(*[Side(style='thin', color='BBBBBB')] * 4)
NUM = '#,##0.00'

_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def _wc(ws, value, font=None, fill=None, border=None):
    """write_only 模式创建带样式的单元格（表头用）。"""
    from openpyxl.cell import WriteOnlyCell
    c = WriteOnlyCell(ws, value=value)
    if font:
        c.font = font
    if fill:
        c.fill = fill
    if border:
        c.border = border
    return c


def log(msg):
    with open(LOG_FP, 'a', encoding='utf-8', buffering=1) as f:
        f.write(msg + '\n')


def iter_sheet_rows(xlsx_path):
    """流式读取 sheet1.xml，逐行 yield {列字母: 值}。
    数字单元格返回字符串（保持精度），sharedString 返回文本，空为 None。
    """
    with zipfile.ZipFile(xlsx_path) as zf:
        # 1) sharedStrings 全量载入（几 MB 内，可接受）
        sst = []
        try:
            with zf.open('xl/sharedStrings.xml') as f:
                for ev, el in ET.iterparse(f, events=('end',)):
                    if el.tag == _NS + 'si':
                        sst.append(''.join(el.itertext()))
                        el.clear()
        except KeyError:
            pass
        # 2) sheet1 流式
        with zf.open('xl/worksheets/sheet1.xml') as f:
            for ev, el in ET.iterparse(f, events=('end',)):
                if el.tag != _NS + 'row':
                    continue
                cells = {}
                for ci, c in enumerate(el):
                    ref = c.get('r') or ''
                    col = ref.rstrip('0123456789') or chr(65 + ci)
                    t = c.get('t')
                    v = None
                    for ch in c:
                        if ch.tag == _NS + 'v':
                            raw = ch.text or ''
                            if t == 's':
                                try:
                                    v = sst[int(raw)]
                                except (IndexError, ValueError):
                                    v = ''
                            elif t == 'inlineStr':
                                v = ''.join(ch.itertext())
                            elif t == 'b':
                                v = raw == '1'
                            else:
                                v = raw  # 数字保留字符串
                            break
                        elif ch.tag == _NS + 'is' and t == 'inlineStr':
                            v = ''.join(ch.itertext())
                            break
                    cells[col] = v
                yield cells
                el.clear()


def load_company_map():
    out = {}
    wb = openpyxl.load_workbook(CODE_FP, read_only=True, data_only=True)
    for r in wb['公司代码'].iter_rows(values_only=True):
        for cell in r:
            m = re.match(r'\|(\d{4})\|([^|]+)', str(cell or '').strip())
            if m:
                nm = m.group(2).strip().rstrip('>').strip()
                nm = re.sub(r'(一分|二分)$', '', nm).strip()
                if nm:
                    out[m.group(1)] = nm
    wb.close()
    return out


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    only_codes = None
    if '--only' in argv:
        only_codes = set(argv[argv.index('--only') + 1].split(','))
        print(f'单账套模式：仅扫描 {sorted(only_codes)}', flush=True)
    # 覆盖写日志（沙箱会拦截 os.remove，safe-delete 保护）
    with open(LOG_FP, 'w', encoding='utf-8'):
        pass
    cmap = load_company_map()
    name2code = {v: k for k, v in cmap.items()}
    full = sorted(name2code, key=len, reverse=True)
    log('公司名映射: %d 家' % len(cmap))

    def match_code(nm):
        # ⚡ 2026-08-18 修复：空字符串 `'' in c` 恒为 True → 误匹配第一个公司名（透平）
        if not nm or not nm.strip():
            return None
        for c in full:
            if len(c) >= 6 and (c in nm or nm in c):
                return c
        return None

    # 正则预过滤（c in nm 方向），避免每行对 80+ 个名字逐条 in
    pat = re.compile('|'.join(re.escape(c) for c in full if len(c) >= 6))
    pat_code = re.compile(r'\b(\d{4})\b')  # 客户/供应商列常存公司代码（如 2230=宏裕）

    def file_codes(base):
        """文件名 → 潜在主体代码集合。
        支持三种命名：单主体 1030.xlsx → {1030}；按月拆分 1010-1月.xlsx → {1010}；
        区间合并 2620-2660.xlsx → {2620..2660}（实际主体以文件内公司列为准，区间仅做候选判断）。"""
        mm = re.match(r'(\d{4})-(\d{4})', base)
        if mm:
            a, b = int(mm.group(1)), int(mm.group(2))
            return {str(c) for c in range(a, b + 1)}
        mm = re.match(r'(\d{4})', base)
        return {mm.group(1)} if mm else set()

    # 扫描序时账全部文件（--only 时仅涉及账套，单对调试快速）
    seq_files = sorted(glob.glob(os.path.join(SEQ_DIR, '*.xlsx')))
    if only_codes:
        seq_files = [f for f in seq_files
                     if file_codes(os.path.basename(f)[:-5]) & only_codes]
    log('序时账文件数: %d' % len(seq_files))

    rows_out = []
    n_hit = 0
    for i, sf in enumerate(seq_files):
        base = os.path.basename(sf)[:-5]
        # 首行做表头定位
        it = iter_sheet_rows(sf)
        try:
            hdr = next(it)
        except StopIteration:
            log('  skip %s 空文件' % base)
            continue
        col_of = {}
        for col, h in hdr.items():
            hs = str(h or '')
            # ⚡⚡ 2026-08-28 修复（对方科目未取完整根因）：日期列【优先"凭证日期"】，"分配"列
            #   仅作最后兜底。SAP 序时账 col0『分配』对往来行是日期、对收入/存货行却是物料文本
            #   （"备品备件"等）→ 若优先"分配"，同凭证 key=(文件,公司,分配,凭证号) 被拆成多组
            #   → 收入/成本行单独成组无命中而丢失 → 对方科目只剩"应收+销项税"（缺主营业务收入）。
            for key in ('凭证日期', '过账日期', '分配'):
                if key in hs and 'date' not in col_of:
                    col_of['date'] = col
            for key in ('凭证编号',):
                if key in hs:
                    col_of['vou'] = col
            for key in ('文本',):
                if key in hs:
                    col_of['txt'] = col
            for key in ('科目',):
                if key in hs:
                    col_of['km'] = col
            for key in ('客户',):
                if key in hs:
                    col_of['cust'] = col
            for key in ('供应商',):
                if key in hs:
                    col_of['supp'] = col
            for key in ('本币金额',):
                if key in hs:
                    col_of['amt'] = col
            for key in ('公司',):
                if key in hs:
                    col_of['comp'] = col
        c_txt = col_of.get('txt')
        c_amt = col_of.get('amt')
        if c_amt is None or c_txt is None:
            log('  skip %s 缺列(%s)' % (base, col_of))
            continue
        n_file = 0
        cur_key = None
        cur_rows = []

        def _row_tc(r):
            """单行命中集团内公司 → 返回 tc（对方公司全名），否则 None。
            ⚡ 2026-08-21 修复：tc 归属以 M/V列(客户/供应商辅助核算) 优先，文本辅助。
               核对表辅助核算按 M/V 列挂账；原文本优先会把 M/V=对方但文本含其他
               集团内公司名的行归错对子（如集团借2202、M/V=1040、文本空/含南昌杭氧）。"""
            txt = str(r[col_of['txt']] or '').strip() if col_of.get('txt') is not None else ''
            cust = str(r[col_of['cust']] or '').strip() if col_of.get('cust') is not None else ''
            supp = str(r[col_of['supp']] or '').strip() if col_of.get('supp') is not None else ''
            # ① M/V列 4位集团代码精确（优先于文本，与核对表辅助核算一致）
            for v in (cust, supp):
                if v and re.fullmatch(r'\d{4}', v) and v in cmap:
                    return cmap[v]
            # ② M/V列 名称匹配（简称/全名，双向包含）
            for v in (cust, supp):
                if v:
                    tc = match_code(v)
                    if tc:
                        return tc
            # ③ 文本辅助（名称优先，其次代码）——仅当 M/V 列未命中集团
            tc = match_code(txt)
            if tc:
                return tc
            ch = pat_code.search(txt)
            if ch and ch.group(1) in cmap:
                return cmap[ch.group(1)]
            return None

        def _flush(rows_):
            """凭证聚合：凭证任一行命中 → 输出凭证全部行（含贷方应付等文本未命中的行）。
            ⚡ 2026-08-18 用户定：匹配必须看完整凭证借贷（'看另一边贷记什么'），
               只抽命中行会漏掉'借费用/贷应付'的贷方应付行。
            ⚡ 2026-08-21 修复：每行输出【自己的 tc】（行级 M/V 归属），
               未命中行跟随凭证级 tc（第一行命中者）——保证 M/V 列归属精确，同时不丢凭证行。"""
            nonlocal n_file
            tcs = [t for t in (_row_tc(x) for x in rows_) if t]
            if not tcs:
                return
            tc_cred = tcs[0]
            comp = ''
            if col_of.get('comp') is not None:
                comp = str(rows_[0][col_of['comp']] or '').strip()
            my = cmap.get(comp, comp) if comp else ''
            if not my:
                mm = re.match(r'(\d{4})', base)
                if mm:
                    my = cmap.get(mm.group(1), mm.group(1))
            for x in rows_:
                # ⚡ 2026-08-18 修复：未命中行若客户/供应商为【外部编号】（数字且非集团 4 位代码），
                #    该行是凭证内对外业务（如资金池代付外部供应商货款 10008875），不属集团往来，跳过。
                row_tc = _row_tc(x)
                if row_tc is None:
                    _ext = False
                    for _c in (str(x[col_of.get('cust')] or '').strip(),
                               str(x[col_of.get('supp')] or '').strip()):
                        if _c.isdigit() and _c not in cmap:
                            _ext = True
                            break
                    if _ext:
                        continue
                    row_tc = tc_cred   # 未命中行跟随凭证级 tc
                amt_raw = x[col_of['amt']]
                try:
                    amt = float(amt_raw or 0)
                except (TypeError, ValueError):
                    continue
                if abs(amt) < 0.005:
                    continue
                txt = str(x[col_of['txt']] or '')
                cust = str(x[col_of.get('cust')] or '') if col_of.get('cust') is not None else ''
                supp = str(x[col_of.get('supp')] or '') if col_of.get('supp') is not None else ''
                dt = str(x[col_of.get('date')] or '')[:10] if col_of.get('date') is not None else ''
                vou = str(x[col_of.get('vou')] or '') if col_of.get('vou') is not None else ''
                km = str(x[col_of.get('km')] or '') if col_of.get('km') is not None else ''
                # ⚡ 账套列输出行级公司代码（合并文件内可精确区分主体），公司列为空回退文件名
                sz = comp if comp else base
                note = '同公司自我（内部科目结转）' if (my and row_tc and my == row_tc) else ''
                rows_out.append([sz, my, dt, vou, km, txt[:80], cust, supp,
                                 amt if amt > 0 else '', -amt if amt < 0 else '',
                                 row_tc, note])
            n_file += len(rows_)
            nonlocal n_hit
            n_hit += len(rows_)

        for r in it:
            vou = str(r[col_of['vou']] or '') if col_of.get('vou') is not None else ''
            dt = str(r[col_of.get('date')] or '') if col_of.get('date') is not None else ''
            # ⚡ 合并文件（如 2620-2660.xlsx）一个文件含多个主体：
            #    --only 时按公司列过滤到目标主体，凭证 key 含主体代码防跨主体串凭证
            comp_code = str(r[col_of.get('comp')] or '').strip() if col_of.get('comp') is not None else ''
            if only_codes and comp_code and comp_code not in only_codes:
                continue
            key = (base, comp_code, dt, vou)
            if cur_key is not None and key != cur_key:
                _flush(cur_rows)
                cur_rows = []
            cur_key = key
            cur_rows.append(r)
        if cur_key is not None:
            _flush(cur_rows)
        log('  [%d/%d] %s 命中凭证 %d 行' % (i + 1, len(seq_files), base, n_file))
        del it
        gc.collect()
    log('命中行总计: %d' % n_hit)

    # ⚡ 6.6万行普通模式逐格样式会卡死（沙箱 OOM/超慢）→ write_only 流式写入
    #    表头用 WriteOnlyCell 保留样式，数据行免样式（中间底稿，供差异定位筛选）
    from openpyxl.cell import WriteOnlyCell
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet('关联往来序时账')
    H = ['账套', '公司', '日期', '凭证编号', '科目', '文本', '客户', '供应商',
         '借方金额', '贷方金额', '对方公司名', '备注']
    ws.append([_wc(ws, v, FH, FILL) for v in H])
    for r in rows_out:
        ws.append(r)
    for col, w in zip('ABCDEFGHIJKL', [10, 18, 11, 13, 14, 50, 12, 12, 15, 15, 22, 16]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = 'A1:L%d' % (len(rows_out) + 1)

    # Sheet2 汇总（行数少，保留数据行样式）
    ws2 = wb.create_sheet('按对汇总')
    H2 = ['我方', '对方', '行数', '借方合计', '贷方合计', '净额']
    ws2.append([_wc(ws2, v, FH, FILL) for v in H2])
    agg = defaultdict(lambda: {'n': 0, 'dr': 0.0, 'cr': 0.0})
    for r in rows_out:
        if r[11] == '同公司自我（内部科目结转）':
            continue
        k = (r[1], r[10])
        agg[k]['n'] += 1
        agg[k]['dr'] += float(r[8] or 0)
        agg[k]['cr'] += float(r[9] or 0)
    for (mx, tx), v in sorted(agg.items(), key=lambda kv: -abs(kv[1]['dr'] - kv[1]['cr'])):
        # write_only 模式不可逐格设样式，数值已 round(2) 直接写入
        ws2.append([mx, tx, v['n'], round(v['dr'], 2), round(v['cr'], 2),
                    round(v['dr'] - v['cr'], 2)])
    for col, w in zip('ABCDEF', [22, 22, 8, 16, 16, 16]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = 'A2'

    out = os.path.join(PREP, '关联往来序时账_2026.xlsx')
    wb.save(out)
    log('已生成: %s' % out)
    log('  关联往来序时账 %d 行 | 对子 %d 组' % (len(rows_out), len(agg)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
