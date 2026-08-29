# -*- coding: utf-8 -*-
"""GFY 使用权资产租赁复核——集团合并（5 主体同名 sheet 合并 + 集团汇总与账面核对）。

把 {主体}.xlsx 各 sheet（合同参数明细/逐期摊销/未确认融资差异/与账面核对/口径说明）
合并为 使用权资产租赁复核_2026.xlsx，**样式与分主体一致**（微软雅黑10、表头浅蓝 D9E2F3、
灰边框、数字 #,##0.00，复制源单元格样式），并重建『集团汇总与账面核对』。

用法：
  python lease_group_merge.py
"""
import os
import sys
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import paths as P

SUBS = ['01FY本级', '05FY皮革', '09FY绍兴', '10FY金属', '11FY龙游']
ACCT_ROOT = P.DATA_DIRS['GFY']
WP_DIR = os.path.join(ACCT_ROOT, '底稿', '2026')

HDR_KW = {
    '合同参数明细': '租赁开始日',
    '逐期摊销(实际利率法)': '期初租赁负债',
    '未确认融资差异': '实际利率法利息',
    '合同级复核明细(2026Q1)': '使用权资产原值 期初',
}
SUMMARY_ITEMS = [
    ('使用权资产原值', '使用权资产原值'),
    ('使用权资产累计折旧', '使用权资产累计折旧'),
    ('租赁负债-租赁付款额', '租赁负债-租赁付款额'),
    ('租赁负债-未确认融资费用', '租赁负债-未确认融资费用'),
]

# 与 lease_review 一致的样式常量
F_TITLE = Font(name='微软雅黑', size=14, bold=True)
F_HEAD = Font(name='微软雅黑', size=10, bold=True)
F = Font(name='微软雅黑', size=10)
F_BOLD = Font(name='微软雅黑', size=10, bold=True)
THIN = Border(*[Side(style='thin', color='BBBBBB')] * 4)
FILL_HDR = PatternFill('solid', fgColor='D9E2F3')
FILL_TOT = PatternFill('solid', fgColor='FFF2CC')
NUM = '#,##0.00'
PCT = '0.0000%'


def _copy_style(src_cell, dst_cell):
    """复制源单元格样式到目标单元格。"""
    dst_cell.font = copy(src_cell.font)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.border = copy(src_cell.border)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format


def _load_rows(fp, sheet):
    """非 read_only 加载：保留样式。返回 worksheet 或 None。"""
    if not os.path.exists(fp):
        return None
    wb = openpyxl.load_workbook(fp)
    return wb[sheet] if sheet in wb.sheetnames else None


def _hdr_idx(ws, kw):
    for i in range(min(6, ws.max_row)):
        for c in list(ws[i + 1])[:6]:
            if c.value is not None and str(c.value).strip() == kw:
                return i
    return None


def _sub_fp(sub):
    """分主体文件路径：lease_review 输出到【中间产物/租赁复核单体_{year}】；glob 兼容各年份/旧归档。"""
    import glob as _g
    for base in (os.path.join(ACCT_ROOT, '中间产物', '租赁复核单体_*'),
                 os.path.join(ACCT_ROOT, '中间产物', '租赁复核单体_误生成20260818')):
        hits = _g.glob(os.path.join(base, f'使用权资产租赁复核_*_{sub}.xlsx'))
        if hits:
            return hits[0]
    return os.path.join(ACCT_ROOT, '中间产物', '租赁复核单体_2026',
                        f'使用权资产租赁复核_2026_{sub}.xlsx')


def _extract_summary(ws_src):
    """从『合同级复核明细』的『主体小计』行提取复核值：
    列布局=核算主体/合同/原值期初/发生/期末/折旧期初/发生/期末/付款额期初/发生/期末/未确认期初/发生/期末。
    返回 {科目: {'end': 复核期末(正数累计口径), 'qm': 本期发生(正)}}；账面期初/期末由 _book_end 读科目余额表。"""
    data = {}
    for r in ws_src.iter_rows():
        vals = [c.value for c in r]
        if len(vals) < 14:
            continue
        if str(vals[1] or '').strip() != '小计':
            continue

        def _f(i):
            try:
                return float(vals[i] or 0)
            except (TypeError, ValueError):
                return 0.0

        data = {
            '使用权资产原值': {'end': _f(4), 'qm': _f(3)},
            '使用权资产累计折旧': {'end': _f(7), 'qm': _f(6)},
            '租赁负债-租赁付款额': {'end': _f(10), 'qm': _f(9)},
            '租赁负债-未确认融资费用': {'end': _f(13), 'qm': _f(12)},
        }
        return data
    return data


def _book_end(sub, acct='GFY'):
    """从科目余额表读主体账面期初/期末（4 科目 opening/closing），供集团汇总。"""
    import lease_manifest as LM
    from lease_review import _load_tb
    try:
        acct_cfg, sub_cfg = LM.job_for(acct, sub)
    except Exception:
        return {}
    cfg = dict(sub_cfg)
    tb_path = os.path.join(LM.DATA_ROOT, cfg['tb'].replace('/', os.sep))
    if not os.path.exists(tb_path):
        return {}
    try:
        tb = _load_tb(tb_path)
    except Exception:
        return {}
    codes = cfg['codes']

    def val(pre, field):
        return sum(float(v.get(field) or 0) for c, v in tb.items()
                   if str(c).startswith(pre))

    return {
        '使用权资产原值': {'open': val(codes['use_asset'][0], 'opening'),
                        'close': val(codes['use_asset'][0], 'closing')},
        '使用权资产累计折旧': {'open': val(codes['use_dep'][0], 'opening'),
                          'close': val(codes['use_dep'][0], 'closing')},
        '租赁负债-租赁付款额': {'open': val(codes['lease_pay'][0], 'opening'),
                          'close': val(codes['lease_pay'][0], 'closing')},
        '租赁负债-未确认融资费用': {'open': val(codes['unrecog_fin'][0], 'opening'),
                            'close': val(codes['unrecog_fin'][0], 'closing')},
    }


def main(acct='GFY', year=2026):
    global ACCT_ROOT, WP_DIR, SUBS
    ACCT_ROOT = P.DATA_DIRS.get(acct, ACCT_ROOT)
    WP_DIR = os.path.join(ACCT_ROOT, '底稿', str(year))
    from lease_manifest import jobs
    _subs = jobs().get(acct, {}).get('subs', {})
    if _subs:
        SUBS = list(_subs.keys())
    wb = openpyxl.Workbook()

    # ============ 集团汇总与账面核对 ============
    ws0 = wb.active
    ws0.title = '集团汇总与账面核对'
    hdr = ['主体',
           '使用权资产原值-账面期末', '使用权资产原值-复核期末', '使用权资产原值-差异',
           '使用权资产累计折旧-账面期末', '使用权资产累计折旧-复核期末', '使用权资产累计折旧-差异',
           '租赁负债-租赁付款额-账面期末', '租赁负债-租赁付款额-复核期末', '租赁负债-租赁付款额-差异',
           '租赁负债-未确认融资费用-账面期末', '租赁负债-未确认融资费用-复核期末', '租赁负债-未确认融资费用-差异']
    subj_rows = {}
    subj_book = {}
    for sub in SUBS:
        ws_src = _load_rows(_sub_fp(sub), '合同级复核明细(2026Q1)')
        if ws_src is None:
            continue
        subj_rows[sub] = _extract_summary(ws_src)
        subj_book[sub] = _book_end(sub, acct)
    # 表头
    for j, h in enumerate(hdr, 1):
        c = ws0.cell(1, j, h)
        c.font = F_HEAD
        c.fill = FILL_HDR
        c.border = THIN
    # 数据
    r_i = 2
    for sub in SUBS:
        data = subj_rows.get(sub, {})
        book = subj_book.get(sub, {})
        vals = [sub]
        for item, _ in SUMMARY_ITEMS:
            d = data.get(item, {})
            close = book.get(item, {}).get('close', 0)
            if item == '使用权资产原值':
                rev = d.get('end', 0)          # 原值复核期末 = LPR 现值
            else:
                # 贷方科目复核期末 = 账面期初 - 本期复核发生（与账面同口径带符号）
                rev = book.get(item, {}).get('open', 0) - d.get('qm', 0)
            vals += [round(close, 2), round(rev, 2), round(rev - close, 2)]
        for j, v in enumerate(vals, 1):
            c = ws0.cell(r_i, j, v)
            c.font = F
            c.border = THIN
            if isinstance(v, (int, float)):
                c.number_format = NUM
        r_i += 1
    # 集团合计
    n_item = len(SUMMARY_ITEMS)
    tot = [0.0] * (n_item * 3)
    for sub in subj_rows:
        data = subj_rows[sub]
        book = subj_book.get(sub, {})
        for k, (item, _) in enumerate(SUMMARY_ITEMS):
            d = data.get(item, {})
            close = book.get(item, {}).get('close', 0)
            if item == '使用权资产原值':
                rev = d.get('end', 0)
            else:
                rev = book.get(item, {}).get('open', 0) - d.get('qm', 0)
            tot[k * 3] += close
            tot[k * 3 + 1] += rev
            tot[k * 3 + 2] += rev - close
    for j, v in enumerate(['集团合计'] + [round(x, 2) for x in tot], 1):
        c = ws0.cell(r_i, j, v)
        c.font = F_BOLD
        c.fill = FILL_TOT
        c.border = THIN
        if j > 1:
            c.number_format = NUM
    ws0.append([])
    ws0.append(['说明：复核期末=按合同基础信息（起止日/剩余租期租金）+ LPR 银行基准利率（按合同开始日对应时点）测算；差异=复核-账面。'])
    ws0.append(['口径：租赁期(月)=合同真实租期；2026计提=账面实际计提月数（Q1=3）。'])
    for r in ws0.iter_rows(min_row=r_i + 1, max_row=ws0.max_row):
        for c in r:
            if c.value:
                c.font = F
    for j in range(1, 14):
        ws0.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 18
    ws0.freeze_panes = 'A2'

    # ============ 各 sheet 合并（样式统一：微软雅黑10/浅蓝表头/细边框/数字格式，插入核算主体列） ============
    for sheet, kw in HDR_KW.items():
        ws = wb.create_sheet(sheet)
        wrote_hdr = False
        max_cols = 1
        skip_first = False
        for sub in SUBS:
            ws_src = _load_rows(_sub_fp(sub), sheet)
            if ws_src is None:
                continue
            hi = _hdr_idx(ws_src, kw)
            if hi is None:
                continue
            # ---- 表头（仅第一主体写入，前插『核算主体』列；源已有主体列时跳过） ----
            if not wrote_hdr:
                src_hdr = [c.value for c in ws_src[hi + 1]]
                skip_first = bool(src_hdr) and str(src_hdr[0]).strip() in ('核算主体', '主体')
                if skip_first:
                    src_hdr = src_hdr[1:]
                for j, v in enumerate(['核算主体'] + src_hdr, 1):
                    c = ws.cell(1, j, v)
                    c.font = F_HEAD
                    c.fill = FILL_HDR
                    c.border = THIN
                max_cols = max(max_cols, len(src_hdr) + 1)
                wrote_hdr = True
            else:
                ws.cell(ws.max_row + 1, 1)          # 主体间空行分隔（append([])不推进行号）
            # ---- 数据（强制样式；口径说明/说明行过滤） ----
            skip_active = skip_first
            for src_r in ws_src.iter_rows(min_row=hi + 2):
                raw0 = str(src_r[0].value or '') if src_r else ''
                # ⚡ 合同级明细的『主体级账面核对』区块行/说明行保留第一列（项目/科目名/标题）
                if skip_active and ('主体级账面核对' in raw0 or raw0.strip() == '项目'
                                    or raw0.strip().startswith('——')
                                    or raw0.strip().startswith('说明：')):
                    skip_active = False
                vals = [c.value for c in src_r]
                if skip_active and vals:
                    vals = vals[1:]                 # ⚡ 跳过源文件自带主体列
                if not any(v is not None and str(v).strip() for v in vals):
                    continue
                r0 = str(vals[0] or '')
                if sheet == '合同参数明细' and any(k in r0 for k in ('口径说明', '现值=各期付款', '未确认融资=付款总额')):
                    continue
                dst_r = ws.max_row + 1
                max_cols = max(max_cols, len(vals) + 1)
                c0 = ws.cell(dst_r, 1, sub)          # 核算主体列
                c0.font = F
                c0.border = THIN
                for j, v in enumerate(vals, 2):
                    c = ws.cell(dst_r, j, v)
                    c.font = F
                    c.border = THIN
                    if isinstance(v, (int, float)):
                        c.number_format = NUM
        if not wrote_hdr:
            ws.append(['核算主体'])
        ws.freeze_panes = 'A2'
        for j in range(1, max_cols + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 15

    # ============ 口径说明（每主体分隔，独立 sheet） ============
    ws = wb.create_sheet('口径说明')
    for sub in SUBS:
        ws_src = _load_rows(_sub_fp(sub), '口径说明')
        if ws_src is None:
            continue
        ws.append([f'—— {sub} ——'])
        for c in ws[ws.max_row]:
            c.font = F_BOLD
        for src_r in ws_src.iter_rows():
            if not any(c.value is not None and str(c.value).strip() for c in src_r):
                continue
            dst_r = ws.max_row + 1
            for j, src_c in enumerate(src_r, 1):
                if src_c.value is not None and str(src_c.value).strip():
                    c = ws.cell(dst_r, j, src_c.value)
                    c.font = F
        ws.cell(ws.max_row + 1, 1)                   # 主体间空行分隔
    ws.column_dimensions['A'].width = 100

    out = os.path.join(WP_DIR, '使用权资产租赁复核_2026.xlsx')
    if '--out' in sys.argv:
        out = os.path.abspath(sys.argv[sys.argv.index('--out') + 1])
    wb.save(out)
    print(f'已生成：{out}')
    print('  合同参数明细 %d 行 | 逐期摊销 %d 行 | 未确认融资 %d 行 | 合同级复核明细 %d 行 | 口径说明 %d 行'
          % (wb['合同参数明细'].max_row, wb['逐期摊销(实际利率法)'].max_row,
             wb['未确认融资差异'].max_row, wb['合同级复核明细(2026Q1)'].max_row,
             wb['口径说明'].max_row))
    return 0


if __name__ == '__main__':
    sys.exit(main())
