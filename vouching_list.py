# -*- coding: utf-8 -*-
"""统一抽凭清单（审计底稿第4层：对金额较大或有疑问的交易抽凭，检查原始凭证）。
数据来源：各主体《综合查询明细表》(GL) + 《科目余额表》(TB, 控制数，仅作附注)。
方法：
  1) 按凭证聚合 GL 全部分录；套用统一抽凭排除规则（结转损益 / 相符计提摊销 不抽）；
  2) 对每笔凭证做「大额 / 异常」标记：
     · 大额：单笔凭证任一分录 借或贷 ≥ 100万；
     · 异常：跨科目的一般会计处理之外分录（资金对手方非典型、应付贷方对手方非采购/费用、
             应收借方对手方非客户、收入/费用对手方为关联/股东/个人/投资/借款等）；
  3) 输出统一抽凭清单（带凭证字+号、日期、摘要、科目、对方、借/贷合计、标记），
     供审计人员逐笔抽取原始凭证核对。
输出：抽凭清单_生成.xlsx
  · 抽凭清单(按金额降序)
  · 按科目分类统计
  · 说明
口径：2026 GL 仅 1-5 月(YTD)，抽凭范围同步受限，须作为取数范围限制披露。
"""
import paths as P
import os, sys, re
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import audit_common as A

DATA_DIR = P.G
OUT = os.path.join(P.G, '抽凭清单_生成.xlsx')
THRESH = 1_000_000.0

THIN = Side(style='thin')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill('solid', fgColor='D9E1F2')
WARNFILL = PatternFill('solid', fgColor='FFF2CC')
BADFILL = PatternFill('solid', fgColor='FCE4D6')
HFONT = Font(bold=True, name='Times New Roman', size=10)
BFONT = Font(bold=True, name='Times New Roman', size=10)
MONEY = '#,##0.00'

INV_KW = ('库存商品', '原材料', '周转材料', '在途物资', '委托加工物资', '材料采购',
          '商品采购', '低值易耗品', '包装物', '发出商品', '半成品', '生产成本', '制造费用')
EXP_KW = ('管理费用', '销售费用', '研发费用', '主营业务成本', '其他业务成本', '财务费用',
          '税金及附加', '营业外支出', '信用减值损失', '资产减值损失', '所得税费用', '业务成本')
SUS_KW = ('关联', '股东', '个人', '投资', '借款', '分红', '股权', '实收资本', '资本公积',
          '捐赠', '备用金', '保证金', '资金拆', '往来款', '基金', '理财', '长投', '长期股权')
NORMAL_OPP_KW = ('应收账款', '应付账款', '预收账款', '预付账款', '库存商品', '原材料', '周转材料',
                 '委托加工', '在途物资', '应付职工薪酬', '应交税费', '管理费用', '销售费用',
                 '制造费用', '研发费用', '主营业务成本', '其他业务成本', '其他应收款', '其他应付款',
                 '合同负债', '固定资产', '在建工程', '无形资产', '长期待摊费用', '使用权资产',
                 '库存现金', '银行存款', '信用减值损失', '资产减值损失', '营业外', '手续费', '利息',
                 '社保', '公积金', '工资', '税金', '成本', '费用', '应收票据', '应付票据')


def _dg(code):
    return re.sub(r'\D', '', str(code))


def is_bank_account(s):
    if not s:
        return False
    s2 = s.strip()
    if '利息' in s2 or '手续费' in s2 or '费用' in s2:
        return False
    if s2.startswith('银行存款'):
        return True
    if re.search(r'(支行|分行|营业部|信用社|农商|邮储|银行)', s2):
        return True
    if re.search(r'^\d{6,}$', s2):
        return True
    return False


def is_ap(name):
    return name and '应付账款' in name


def is_ar(name):
    return name and ('应收账款' in name or '应收票据' in name)


def is_rev(name):
    return name and ('主营业务收入' in name or '其他业务收入' in name or '收入' in name or name.startswith('6001'))


def is_exp(name):
    return name and any(k in name for k in EXP_KW)


def classify_ap_opp(opp):
    if not opp:
        return '异常(空对方)'
    if '应付账款' in opp:
        return '应付内部重分类'
    if any(k in opp for k in INV_KW):
        return '存货采购'
    if any(k in opp for k in EXP_KW):
        return '费用'
    if '预付' in opp:
        return '预付冲销'
    if '应付票据' in opp or '票据' in opp:
        return '应付票据'
    return '异常'


def flag_voucher(v):
    """返回 (tags, trigger_name, trigger_opp, vdr, vcr)。tags 为空表示该凭证不需抽凭。"""
    lines = v.get('lines', [])
    vdr = sum(l.get('dr', 0.0) or 0.0 for l in lines)
    vcr = sum(l.get('cr', 0.0) or 0.0 for l in lines)
    max_amt = max([(l.get('dr', 0.0) or 0.0) for l in lines] +
                  [(l.get('cr', 0.0) or 0.0) for l in lines] + [0.0])
    tags = []
    trig_name = trig_opp = ''
    # 大额
    if max_amt >= THRESH:
        tags.append('大额(>=100万)')
        if abs(max_amt % 1_000_000) < 0.5:
            tags.append('整数大额')
    # 逐行异常判定
    for l in lines:
        nm = (l.get('name') or '').strip()
        op = (l.get('opp') or '').strip()
        dr = l.get('dr', 0.0) or 0.0
        cr = l.get('cr', 0.0) or 0.0
        # 1) 资金对手方非典型：一侧为银行账户，另侧非「银行/典型经营科目」
        bank_side = is_bank_account(nm) or is_bank_account(op)
        if bank_side:
            other = op if is_bank_account(nm) else nm
            # 另侧也是银行账户 → 银行间互转，属正常内部划转，不标异常（仅按大额处理）
            if other and not (is_bank_account(other) or any(k in other for k in NORMAL_OPP_KW)):
                if '异常(资金对手方非典型)' not in tags:
                    tags.append('异常(资金对手方非典型)')
                    if not trig_name:
                        trig_name, trig_opp = nm, op
        # 2) 应付贷方对手方异常/可疑
        if is_ap(nm) and cr > 0:
            cat = classify_ap_opp(op)
            if cat == '异常':
                tags.append('异常(应付贷方对手方非采购/费用)')
                if not trig_name:
                    trig_name, trig_opp = nm, op
            if any(k in op for k in SUS_KW):
                if '可疑对手方(' not in ''.join(tags):
                    tags.append('可疑对手方(应付)')
                if not trig_name:
                    trig_name, trig_opp = nm, op
        # 3) 应收借方对手方异常（正常赊销 借应收 贷主营业务收入/应收票据 不标异常）
        if is_ar(nm) and dr > 0:
            if op:
                if is_bank_account(op) or '库存现金' in op:
                    tags.append('异常(应收借方对手方为资金)')
                    if not trig_name:
                        trig_name, trig_opp = nm, op
                elif any(k in op for k in ('实收资本', '资本公积', '盈余公积', '长期股权投资',
                                          '长期股权', '投资收益', '费用', '借款')):
                    tags.append('异常(应收借方对手方非经营)')
                    if not trig_name:
                        trig_name, trig_opp = nm, op
        # 4) 收入对手方可疑
        if is_rev(nm) and cr > 0:
            if any(k in op for k in SUS_KW):
                if '可疑对手方(' not in ''.join(tags):
                    tags.append('可疑对手方(收入)')
                if not trig_name:
                    trig_name, trig_opp = nm, op
        # 5) 费用对手方可疑
        if is_exp(nm) and dr > 0:
            if any(k in op for k in SUS_KW):
                if '可疑对手方(' not in ''.join(tags):
                    tags.append('可疑对手方(费用)')
                if not trig_name:
                    trig_name, trig_opp = nm, op
        # 6) 通用可疑对手方（任一行为关联/股东/个人/投资/借款等）
        if op and any(k in op for k in SUS_KW):
            if '可疑对手方(' not in ''.join(tags):
                tags.append('可疑对手方(通用)')
            if not trig_name:
                trig_name, trig_opp = nm, op
    # 触发科目/对方科目兜底：取金额最大行（避免纯大额凭证留空）
    if not trig_name:
        best = max(lines, key=lambda l: max(l.get('dr', 0.0) or 0.0, l.get('cr', 0.0) or 0.0))
        trig_name = (best.get('name') or '').strip()
        trig_opp = (best.get('opp') or '').strip()
    # 去重 tags 保序
    seen = set(); uniq = []
    for t in tags:
        if t not in seen:
            seen.add(t); uniq.append(t)
    return uniq, trig_name, trig_opp, vdr, vcr


def subject_of(name):
    if not name:
        return '其他'
    if is_bank_account(name) or '库存现金' in name:
        return '资金类'
    if is_ap(name):
        return '应付账款'
    if is_ar(name):
        return '应收款项'
    if is_rev(name):
        return '收入'
    if is_exp(name):
        return '费用'
    if any(k in name for k in INV_KW):
        return '存货/成本'
    if '固定资产' in name or '在建工程' in name or '无形资产' in name or '长期待摊' in name or '使用权资产' in name:
        return '长期资产'
    if '应付职工薪酬' in name or '应交税费' in name or '其他应付款' in name or '其他应收' in name:
        return '往来/负债'
    if '实收资本' in name or '资本公积' in name or '盈余公积' in name or '利润分配' in name or '本年利润' in name or '未分配利润' in name:
        return '权益'
    return '其他'


def build(data_dir, out_path):
    entities = A.discover_entities(data_dir)
    if not entities:
        print('❌ 未发现任何账套主体'); return None
    rows = A.read_gl_rows(data_dir, entities)
    years = sorted({y for e, yd in entities.items() for y in yd})
    print(f'· 主体 {len(entities)} 个，年份 {years}，GL 行数 {len(rows)}')

    vouchers = A.group_gl_vouchers(rows)
    print(f'· 聚合凭证 {len(vouchers)} 笔')

    recs = []          # 抽凭记录
    cat_stat = {}      # 科目分类 -> (笔数, 金额)
    skip_n = 0
    for v in vouchers:
        sm = v.get('summary') or ''
        lines = [{'name': (l.get('name') or ''), 'opp': (l.get('opp') or ''),
                  'dr': l.get('dr', 0.0) or 0.0, 'cr': l.get('cr', 0.0) or 0.0} for l in v.get('lines', [])]
        skip, reason = A.voucher_should_skip(sm, lines[0]['name'] if lines else '',
                                             lines[0]['opp'] if lines else '', lines)
        if skip:
            skip_n += 1
            continue
        tags, tn, top, vdr, vcr = flag_voucher(v)
        if not tags:
            continue
        amt = max(vdr, vcr)
        recs.append({
            'e': v.get('e'), 'y': v.get('y'), 'vtype': v.get('vtype'), 'vno': v.get('vno'),
            'date': v.get('date'), 'summary': sm, 'tn': tn, 'top': top,
            'vdr': vdr, 'vcr': vcr, 'tags': tags, 'amt': amt,
        })
        # 以触发科目的大类计入统计
        subj = subject_of(tn) if tn else '其他'
        d = cat_stat.setdefault(subj, {'n': 0, 'amt': 0.0})
        d['n'] += 1; d['amt'] += amt
        # 同时按"异常/大额"计数
        if any(t.startswith('异常') for t in tags):
            d2 = cat_stat.setdefault('（其中异常类）', {'n': 0, 'amt': 0.0})
            d2['n'] += 1; d2['amt'] += amt
        if any(t.startswith('大额') for t in tags):
            d3 = cat_stat.setdefault('（其中大额类）', {'n': 0, 'amt': 0.0})
            d3['n'] += 1; d3['amt'] += amt

    recs.sort(key=lambda x: -x['amt'])
    print(f'· 排除(结转损益/相符计提){skip_n} 笔；需抽凭 {len(recs)} 笔')

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    def style_header(ws, hdr, row=3):
        for c, h in enumerate(hdr, 1):
            cell = ws.cell(row, c, h); cell.font = HFONT; cell.fill = HFILL; cell.border = BORDER
            cell.alignment = Alignment(horizontal='center', wrap_text=True, vertical='center')

    # ---------- Sheet1 抽凭清单 ----------
    ws = wb.create_sheet('抽凭清单(按金额降序)')
    ws.cell(1, 1, '统一抽凭清单（金额较大或有疑问的交易，抽凭检查原始凭证）').font = Font(bold=True, name='Times New Roman', size=10)
    ws.cell(2, 1, '已排除 结转损益/相符计提摊销 凭证；标记=大额(>=100万)/异常/可疑对手方；'
                  '「抽凭结论」列由审计人员填列。2026 GL 仅 1-5 月(YTD)，抽凭范围受限。').font = \
        Font(italic=True, name='Times New Roman', size=10, color='808080')
    hdr = ['核算主体', '年度', '凭证字', '凭证号', '日期', '摘要', '触发科目', '对方科目',
           '借方(凭证合计)', '贷方(凭证合计)', '标记', '抽凭结论']
    style_header(ws, hdr, row=3)
    r = 4
    for x in recs:
        row = [x['e'], x['y'], x['vtype'], x['vno'], x['date'], x['summary'], x['tn'], x['top'],
               x['vdr'], x['vcr'], '；'.join(x['tags']), '']
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val); cell.border = BORDER
            if c in (9, 10):
                cell.number_format = MONEY
            if c == 11:
                cell.fill = BADFILL if any(t.startswith('异常') for t in x['tags']) else WARNFILL
        r += 1
    if not recs:
        ws.cell(r, 1, '无达到抽凭阈值的交易。').font = Font(name='Times New Roman', italic=True)
    for i, w in enumerate([14, 7, 8, 10, 12, 34, 24, 30, 16, 16, 34, 14], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'

    # ---------- Sheet2 按科目分类统计 ----------
    ws2 = wb.create_sheet('按科目分类统计')
    ws2.cell(1, 1, '抽凭笔数按科目大类统计').font = Font(bold=True, name='Times New Roman', size=10)
    hdr2 = ['科目大类', '抽凭笔数', '涉及金额(借/贷较大方)']
    style_header(ws2, hdr2, row=3)
    r = 4
    for subj in sorted(cat_stat, key=lambda k: -cat_stat[k]['amt']):
        d = cat_stat[subj]
        ws2.cell(r, 1, subj).border = BORDER
        ws2.cell(r, 2, d['n']).border = BORDER
        c = ws2.cell(r, 3, round(d['amt'], 2)); c.border = BORDER; c.number_format = MONEY
        r += 1
    t_n = sum(d['n'] for d in cat_stat.values() if not d['n'] is None)
    # 合计（排除汇总行）
    base = [k for k in cat_stat if not k.startswith('（')]
    ws2.cell(r, 1, '合计').font = BFONT
    ws2.cell(r, 2, sum(cat_stat[k]['n'] for k in base)).font = BFONT
    c = ws2.cell(r, 3, round(sum(cat_stat[k]['amt'] for k in base), 2)); c.font = BFONT; c.number_format = MONEY
    for cc in range(1, 4):
        ws2.cell(r, cc).border = BORDER
    for i, w in enumerate([22, 12, 26], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # ---------- Sheet3 说明 ----------
    ws3 = wb.create_sheet('说明')
    notes = [
        '统一抽凭清单 编制与抽凭规则说明',
        '',
        '一、目的（第4层：对金额较大或有疑问的交易抽凭，检查原始凭证）',
        '  本清单将各科目分析中识别的「大额 / 异常 / 可疑」交易汇总为一本统一抽凭清单，',
        '  供审计人员逐笔抽取记账凭证及原始单据（发票、合同、银行回单、出入库单等）核对。',
        '',
        '二、抽凭排除规则（与全集团小程序一致）',
        '  · 规则① 结转损益类凭证（摘要含 结转本年利润/期间损益结转 等）不抽；',
        '  · 规则② 相符计提/摊销（借费用贷应付职工薪酬/累计折旧/递延收益 等，借贷相等且无资金结算）不抽；',
        '  · 聚焦例外：直接付款/代付/红字冲减/跨科目异常/大额交易。',
        '',
        '三、标记规则（阈值 100万元）',
        '  · 大额(>=100万)：单笔凭证任一分录 借或贷 ≥ 100万（整数大额另行标注）；',
        '  · 异常(资金对手方非典型)：银行/现金 一侧的对手方为 权益/长投/其他往来/关联 等非经营科目；',
        '  · 异常(应付贷方对手方非采购/费用)：贷 应付账款 对手方为 银行/现金/应收/权益/长投/个人 等（非存货采购/费用）；',
        '  · 异常(应收借方对手方非客户)：借 应收账款/应收票据 对手方非 客户/银行/应收票据；',
        '  · 可疑对手方(应付/收入/费用/通用)：对手方含 关联/股东/个人/投资/借款/分红/备用金/保证金/往来款 等。',
        '  · 上述均为「应抽凭核查」的筛选提示，是否构成会计差错需结合原始凭证与业务实质判断。',
        '',
        '四、与既有分析的关系',
        '  · 本清单与 银行存款发生额异常分析、应付账款贷方发生勾稽 中已列的大额/异常分录口径一致，',
        '    此处统一归集并补全 应收/收入/费用/长期资产 等科目，形成一本完整抽凭底稿。',
        '',
        '五、口径提示',
        '  · 2026 综合查询明细表仅含 1-5 月(YTD)，抽凭范围受限，列示分录非全年完整，须作为取数范围限制披露。',
        f'  · 本次共识别需抽凭凭证 {len(recs)} 笔（已排除 {skip_n} 笔 结转损益/相符计提）。',
    ]
    for i, t in enumerate(notes, 1):
        ws3.cell(i, 1, t)
        if i == 1:
            ws3.cell(i, 1).font = Font(name='Times New Roman', bold=True, size=10)
    ws3.column_dimensions['A'].width = 120

    path, warn = A._safe_save(wb, out_path)
    if warn:
        print('  ⚠️', warn)
    print(f'  ✓ 已保存：{path}  (需抽凭 {len(recs)} 笔)')
    return path


def main():
    data_dir = sys.argv[1] if len(sys.argv) > 1 else DATA_DIR
    out = sys.argv[2] if len(sys.argv) > 2 else OUT
    build(data_dir, out)


if __name__ == '__main__':
    main()
