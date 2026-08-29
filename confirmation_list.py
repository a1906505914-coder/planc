# -*- coding: utf-8 -*-
"""函证清单（审计底稿第5层·函证：应收函证 + 应付函证）。
数据来源：各主体《辅助核算余额表》(含 1122/2202 按客商末级余额) + 《科目余额表》(TB, 控制数勾稽)。
方法：
  · 从辅助核算余额表提取 等级=4 的末级客商行（科目辅助核算名称=真实客商名称，非'客户/供应商'分组标签），
    会计账户以「科目名称」列判定（应收/应付），不能用 code 前缀（国外应收客户代码以2202开头会误判）；
    取其期末方向+金额(带符号)作为该客商期末往来余额；
  · 应收账款函证清单：按客户列示 2025期末余额 / 2026期初余额 / 是否关联方；
  · 应付账款函证清单：按供应商列示 2025期末余额 / 2026期初余额 / 是否关联方；
  · 与 TB(1122/2202) 一级控制数勾稽（明细合计应=控制数），差异列示。
输出：函证清单_生成.xlsx
  · 应收账款函证清单
  · 应付账款函证清单
  · 勾稽说明
口径：审计期间 FY2025，函证基准日=2025-12-31（取 2025 期末；2026期初作衔接校验）。
"""
import paths as P
import os, sys, re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import audit_common as A

DATA_DIR = P.G
OUT = os.path.join(P.G, '函证清单_生成.xlsx')

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D9E1F2')
WARNFILL = PatternFill('solid', fgColor='FFF2CC')
BADFILL = PatternFill('solid', fgColor='FCE4D6')
HFONT = Font(bold=True, name='Times New Roman', size=10)
BFONT = Font(bold=True, name='Times New Roman', size=10)
MONEY = '#,##0.00'

# 仅「精确等于」这些分组标签的行才是汇总分组行（余额0，应排除）；
# 形如"个人XXX/单位XXX"的真实客商名须保留（此前误按前缀排除导致应收缺漏）。
GROUP_LABELS = ('客户', '供应商', '项目', '部门', '员工', '客商', '核算', '单位', '个人')

# 其他往来科目（非函证对象 应收/应付）：出现即把账户上下文清空，避免其末级客商被错归应收/应付。
EXCLUDE_KW = ('应收票据', '应付票据', '应收利息', '应付利息', '其他应收', '其他应付',
              '预收', '预付', '坏账', '应收股利', '应付职工薪酬', '应收股息')


def _dg(code):
    return re.sub(r'\D', '', str(code))


def _aux_ent_name(fn):
    """实体名抽取（与 audit_common.discover_entities 一致：剥年份+科余表关键字+尾随杂质）。"""
    base = re.sub(r'(19|20)\d{2}', '', fn)
    base = re.sub(r'年度|年\d*[-~]?\d*月', '', base)  # 剔除"年度"、"年1-3月"等期间后缀
    E = re.sub(r'(外币)?(科目余额表|科目余额|余额表|综合查询明细表|综合查询表|辅助核算余额表|辅助核算|总账|trial).*$',
                  '', base, flags=re.IGNORECASE).strip(' _-')
    # 个位数前缀补0（1、XX → 01、XX）
    _m0 = re.match(r'^(\d)([、,，.．\s])', E)
    if _m0 and len(E) > 2:
        E = '0' + E
    return E


def discover_aux(data_dir):
    """返回 {(ent, year): path}，仅取 辅助核算余额表。
    年份判定稳健化（2026-07-27）：用 A._extract_year 先剥导出时间戳再取数据年；
    若 aux 文件名漏标数据年(仅含时间戳)，则继承该主体 TB(科目余额表)的数据年，
    避免「金桥亦法」类导出漏标导致 aux 被错挂到导出年份(2026)。
    ent 命名与 audit_common.discover_entities 对齐，保证与 TB 勾稽匹配。"""
    aux = {}
    if not os.path.isdir(data_dir):
        return aux
    files = [fn for fn in sorted(os.listdir(data_dir))
             if fn.endswith('.xlsx') and not fn.startswith('~') and not fn.startswith('.')]
    # Pass 1：各主体 TB 数据年（权威）
    tb_year = {}
    for fn in files:
        if not (('科目余额表' in fn) and '辅助核算' not in fn
                and '综合查询' not in fn and '外币' not in fn):
            continue
        E = _aux_ent_name(fn)
        y = A._extract_year(fn)
        if E and y:
            tb_year.setdefault(E, y)
    # Pass 2：aux 年份（优先文件名去时间戳；缺 token 继承 TB 年）
    for fn in files:
        if '辅助核算余额表' not in fn:
            continue
        E = _aux_ent_name(fn)
        y = A._extract_year(fn)
        if y is None:
            y = tb_year.get(E)
        if y and E:
            aux[(E, y)] = os.path.join(data_dir, fn)
    return aux


def signed_bal(direction, amount):
    a = float(amount or 0.0)
    if str(direction).strip() == '贷':
        return -a
    return a


def classify_name(name):
    """依科目名称判定应收/应付/排除。返回 '1122' / '2202' / 'EXCLUDE' / None(无法判定, 继承)。"""
    n = str(name or '').strip()
    if not n:
        return None
    if '应收账款' in n and '应收票据' not in n and '坏账' not in n \
            and '应收利息' not in n and '其他应收' not in n:
        return '1122'
    if '应付账款' in n and '应付票据' not in n and '应付利息' not in n \
            and '应付职工' not in n and '其他应付' not in n:
        return '2202'
    if any(k in n for k in EXCLUDE_KW):
        return 'EXCLUDE'
    return None


def parse_aux(path):
    """返回 { '1122':[...], '2202':[...] }，每项 {'party','related','end'}。
    判定账户上下文的**权威信号是「科目名称」列**（末级客商行的该列有值，如'应收账款-国内应收账款'）；
    code 前缀不可靠——国外应收的客户辅助代码形如 220222/2852003，会以'2202'开头而被误判应付。
    规则：
      · level==1 一级科目头：按 name 设 cur(1122/2202)，其他科目(含收入/费用)重置 cur=None；
        name 为空(合并单元格)时按 code 前缀兜底(仅表头行，绝不用客户代码)；
      · level>=2：name 可判定应收/应付则更新 cur；命中 EXCLUDE_KW(票据/利息/其他应收/坏账等)则 cur=None；
        否则继承 cur（处理空 name 的末级客商行）；
      · 仅 level==4 末级且 科目辅助核算名称 为真实客商(非空、非精确分组标签)、非零余额 才提取。
    """
    out = {'1122': [], '2202': []}
    if not path or not os.path.exists(path):
        return out
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    cur = None
    for r in rows:
        if not r or len(r) < 12:
            continue
        related = r[1]                         # 关联方
        name = str(r[2] or '').strip()        # 科目名称
        pname = str(r[3] or '').strip()       # 科目辅助核算名称（末级=真实客商名）
        qc_amt = r[5]                          # 期初金额
        fdir = r[8]                            # 期末方向
        famt = r[9]                            # 期末金额
        level = r[10]                          # 等级
        if level == 1:
            res = classify_name(name)
            if res in ('1122', '2202'):
                cur = res
            elif not name:
                g = _dg(r[0])
                cur = '1122' if g.startswith('1122') else ('2202' if g.startswith('2202') else None)
            else:
                cur = None
            continue
        # level >= 2
        res = classify_name(name)
        if res in ('1122', '2202'):
            cur = res
        elif res == 'EXCLUDE':
            cur = None
        # else 继承 cur
        if cur is None:
            continue
        if level != 4:
            continue
        if not pname or pname in GROUP_LABELS:
            continue
        if abs(float(famt or 0.0)) < 0.005 and abs(float(qc_amt or 0.0)) < 0.005:
            continue  # 零余额末级（分组行）跳过
        out[cur].append({
            'party': pname,
            'related': bool(related),
            'end': signed_bal(fdir, famt),
        })
    return out


def build(data_dir, out_path):
    aux = discover_aux(data_dir)
    if not aux:
        print('❌ 未发现任何辅助核算余额表'); return None
    entities = A.discover_entities(data_dir)
    tb = A.read_tb_full(data_dir, entities)
    print(f'· 辅助核算文件 {len(aux)} 个；主体 {len(entities)} 个')

    # 归集 AR / AP 末级客商
    ar_rows = []   # (ent, year, party, related, end)
    ap_rows = []
    for (ent, year), path in sorted(aux.items()):
        parsed = parse_aux(path)
        for g, lst in parsed.items():
            for d in lst:
                if g == '1122':
                    ar_rows.append((ent, year, d['party'], d['related'], d['end']))
                else:
                    ap_rows.append((ent, year, d['party'], d['related'], d['end']))

    # TB 控制数（一级）
    tb_ar = {}   # (e,y) -> qm signed
    tb_ap = {}
    for (e, code, name, y), v in tb.items():
        g = _dg(code)
        if g == '1122':
            tb_ar[(e, y)] = tb_ar.get((e, y), 0.0) + v['qm']
        elif g == '2202':
            tb_ap[(e, y)] = tb_ap.get((e, y), 0.0) + v['qm']

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def style_header(ws, hdr, row=3):
        for c, h in enumerate(hdr, 1):
            cell = ws.cell(row, c, h); cell.font = HFONT; cell.fill = HFILL; cell.border = BORDER
            cell.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')

    # ---------- Sheet1 应收账款函证清单 ----------
    ws = wb.create_sheet('应收账款函证清单')
    ws.cell(1, 1, '应收账款函证清单（按客户，基准日 2025-12-31）').font = Font(bold=True, name='Times New Roman', size=10)
    ws.cell(2, 1, '来源=辅助核算余额表(1122 末级)；期末余额=2025期末(审计基准)；2026期初作衔接校验。'
                  '「函证结论」由审计人员填列（回函相符/不符/未回函替代测试）。').font = \
        Font(italic=True, name='Times New Roman', size=10, color='808080')
    hdr = ['核算主体', '年度', '客户名称', '是否关联方', '期末余额(2025-12-31)', '期初余额(2026-01-01)',
           '余额方向', '函证结论']
    style_header(ws, hdr, row=3)
    r = 4
    # 按 (ent, year, party) 聚合（同客户可能多条末级，合计）
    ar_agg = {}
    for ent, year, party, rel, end in ar_rows:
        k = (ent, year, party)
        d = ar_agg.setdefault(k, {'rel': rel, 'end': 0.0, 'n': 0})
        d['end'] += end; d['rel'] = d['rel'] or rel; d['n'] += 1
    ar_items = sorted(ar_agg.items(), key=lambda x: (x[0][0], x[0][1], -abs(x[1]['end'])))
    for (ent, year, party), d in ar_items:
        end = d['end']
        dirn = '借' if end >= 0 else '贷'
        # 2026期初：同客户在 2026 辅助核算的余额（若有）
        next_end = ar_agg.get((ent, '2026', party), {}).get('end', None)
        row = [ent, year, party, '是' if d['rel'] else '否', abs(end),
               abs(next_end) if next_end is not None else '', dirn, '']
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val); cell.border = BORDER
            if c in (5, 6):
                cell.number_format = MONEY
            if c == 4 and d['rel']:
                cell.fill = WARNFILL
        r += 1
    # 合计
    g_end = sum(d['end'] for d in ar_agg.values())
    ws.cell(r, 1, '合计').font = BFONT
    ws.cell(r, 3, f'{len(ar_agg)} 个客户').font = BFONT
    c = ws.cell(r, 5, abs(g_end)); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
    for cc in (1, 2, 3, 4, 7, 8):
        ws.cell(r, cc).border = BORDER
    for i, w in enumerate([16, 7, 40, 10, 20, 20, 9, 16], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'

    # ---------- Sheet2 应付账款函证清单 ----------
    ws2 = wb.create_sheet('应付账款函证清单')
    ws2.cell(1, 1, '应付账款函证清单（按供应商，基准日 2025-12-31）').font = Font(bold=True, name='Times New Roman', size=10)
    ws2.cell(2, 1, '来源=辅助核算余额表(2202 末级)；期末余额=2025期末(审计基准)；2026期初作衔接校验。').font = \
        Font(italic=True, name='Times New Roman', size=10, color='808080')
    style_header(ws2, hdr, row=3)
    r = 4
    ap_agg = {}
    for ent, year, party, rel, end in ap_rows:
        k = (ent, year, party)
        d = ap_agg.setdefault(k, {'rel': rel, 'end': 0.0, 'n': 0})
        d['end'] += end; d['rel'] = d['rel'] or rel; d['n'] += 1
    ap_items = sorted(ap_agg.items(), key=lambda x: (x[0][0], x[0][1], -abs(x[1]['end'])))
    for (ent, year, party), d in ap_items:
        end = d['end']
        dirn = '借' if end >= 0 else '贷'
        next_end = ap_agg.get((ent, '2026', party), {}).get('end', None)
        row = [ent, year, party, '是' if d['rel'] else '否', abs(end),
               abs(next_end) if next_end is not None else '', dirn, '']
        for c, val in enumerate(row, 1):
            cell = ws2.cell(r, c, val); cell.border = BORDER
            if c in (5, 6):
                cell.number_format = MONEY
            if c == 4 and d['rel']:
                cell.fill = WARNFILL
        r += 1
    g_end2 = sum(d['end'] for d in ap_agg.values())
    ws2.cell(r, 1, '合计').font = BFONT
    ws2.cell(r, 3, f'{len(ap_agg)} 个供应商').font = BFONT
    c = ws2.cell(r, 5, abs(g_end2)); c.font = BFONT; c.number_format = MONEY; c.border = BORDER
    for cc in (1, 2, 3, 4, 7, 8):
        ws2.cell(r, cc).border = BORDER
    for i, w in enumerate([16, 7, 40, 10, 20, 20, 9, 16], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws2.freeze_panes = 'A4'

    # ---------- Sheet3 勾稽说明 ----------
    ws3 = wb.create_sheet('勾稽说明')
    notes = [
        '函证清单 编制与勾稽说明',
        '',
        '一、方法（第5层·函证）',
        '  1. 应收/应付函证对象取自《辅助核算余额表》末级客商行（科目辅助核算名称=真实客商，等级=3）；',
        '  2. 函证基准日=2025-12-31，取 2025 期末余额；2026 期初余额作跨期衔接校验；',
        '  3. 是否关联方取自辅助核算「关联方」列，函证时需特别关注关联交易完整性。',
        '',
        '二、与 TB 控制数勾稽（明细合计 应= 一级控制数）',
        '  各主体 应收(1122)/应付(2202) 末级客商余额合计 与 TB 一级期末余额核对；',
        '  差异>容差提示可能存在未达/遗漏客商或辅助核算与总账不符，需追查。',
        '',
    ]
    # 勾稽明细（区分「真实余额差异」与「辅助核算文件缺失=取数范围限制」）
    def _flag(diff, ctrl, key):
        if abs(diff) <= max(1.0, abs(ctrl) * 1e-6):
            return ''
        if key not in aux:
            return '  <== 辅助核算文件缺失(取数范围限制)'
        return '  <== 差异!'
    notes.append('  · 应收账款 明细合计 vs TB(1122)：')
    for (e, y) in sorted(tb_ar):
        sub = sum(d['end'] for (ent, year, party), d in ar_agg.items() if ent == e and year == y)
        diff = sub - tb_ar[(e, y)]
        flag = _flag(diff, tb_ar[(e, y)], (e, y))
        notes.append(f'    {e} {y}: 明细合计={sub:,.2f}  TB控制={tb_ar[(e,y)]:,.2f}  差={diff:,.2f}{flag}')
    notes.append('  · 应付账款 明细合计 vs TB(2202)：')
    for (e, y) in sorted(tb_ap):
        sub = sum(d['end'] for (ent, year, party), d in ap_agg.items() if ent == e and year == y)
        diff = sub - tb_ap[(e, y)]
        flag = _flag(diff, tb_ap[(e, y)], (e, y))
        notes.append(f'    {e} {y}: 明细合计={sub:,.2f}  TB控制={tb_ap[(e,y)]:,.2f}  差={diff:,.2f}{flag}')
    n_missing = sum(1 for k in set(list(tb_ar) + list(tb_ap)) if k not in aux
                    and (abs(tb_ar.get(k, 0.0)) > 1.0 or abs(tb_ap.get(k, 0.0)) > 1.0))
    if n_missing:
        notes.append(f'  · 注：{n_missing} 个(主体,年度)有 TB 余额但缺失辅助核算余额表，'
                      '属取数范围限制（非账务差错），函证对象以 2026 或其他可取得的辅助核算数据替代/追查。')
    notes += [
        '',
        '三、建议',
        '  · 应收：对余额重大/关联方/长期挂账客户执行积极式函证；未回函实施替代测试（检查期后收款/销售合同发票出货单）。',
        '  · 应付：对余额重大/关联方供应商函证，并核查暂估应付款完整性（已入库未收票）。',
        '  · 2026 仅含 1-5 月数据，2026期初仅作衔接，函证基准仍以 2025 期末为准。',
        f'  · 应收共 {len(ar_agg)} 个客户末级，应付共 {len(ap_agg)} 个供应商末级。',
    ]
    for i, t in enumerate(notes, 1):
        ws3.cell(i, 1, t)
        if i == 1:
            ws3.cell(i, 1).font = Font(name='Times New Roman', bold=True, size=10)
    ws3.column_dimensions['A'].width = 120

    path, warn = A._safe_save(wb, out_path)
    if warn:
        print('  ⚠️', warn)
    print(f'  ✓ 已保存：{path}  (应收 {len(ar_agg)} 客户 / 应付 {len(ap_agg)} 供应商)')
    return path


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else DATA_DIR
    out = sys.argv[2] if len(sys.argv) > 2 else OUT
    build(data_dir, out)


if __name__ == '__main__':
    main()
