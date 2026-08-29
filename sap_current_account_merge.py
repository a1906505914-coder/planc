# -*- coding: utf-8 -*-
"""sap_current_account_merge.py —— SAP 往来款集团合并底稿 v2（2026-08-09 用户要求）。

v2 变更（用户反馈"之前集团底稿不是这样的"）：按【单体底稿样式】对齐——
不再用「6957 往来单位 × 88 主体列」大矩阵，改为与单体底稿相同的 sheet 结构，
88 个核算主体【竖排、每主体一块】：

    Sheet1 审定表     公司/期初/增/减/期末（88 主体每主体一行 + 集团合计）
    Sheet2 增减构成   增减类别/金额（每主体一块竖排）
    Sheet3 异常凭证   凭证级异常（每主体一块竖排）
    Sheet4 明细表     核算主体/序号/单位/款项性质/期初/借/贷/期末/平衡（每主体一块竖排 + 小计）
    Sheet5 账龄分布   账龄区间/期末余额（每主体一块竖排）
    Sheet6 Top10余额  每主体期末 Top10（每主体一块竖排）
    Sheet7 对方科目核对 对方科目/借/贷/笔数（每主体一块竖排）
    Sheet8 底稿说明 + Sheet9 审计调整分录（finalize 自动）

数据源与单体一致：应收/{comp}.xlsx（客户）、应付/{comp}.xlsx（供应商）；
铁律73：应付表贷正借负 → qc/qm 取负还原；TB 兜底占位行承接 aux≠TB 差额。

用法：
    python sap_current_account_merge.py [--out 输出目录]
    # 默认输出 D:/底稿测试/yy/往来款合并底稿/
"""
import paths as P
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import sap_common as C
import sap_reader as SR
import sap_current_account_detail as SCA

DATA_DIR = P.YY
OUT_DIR = os.path.join(DATA_DIR, '往来款合并底稿')

F10 = Font(name='Times New Roman', size=10)
BF = Font(name='Times New Roman', size=10, bold=True)
TF = Font(name='Times New Roman', size=12, bold=True)
HFILL = PatternFill('solid', fgColor='DDEBF7')
TITLEFILL = PatternFill('solid', fgColor='BDD7EE')
TOTFILL = PatternFill('solid', fgColor='FCE4D6')
BLOCKFILL = PatternFill('solid', fgColor='D9E2F3')
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
RGT = Alignment(horizontal='right', vertical='center')
NUM = '#,##0.00'

# 往来科目（与单体 SUBJECTS 一致）
SUBJECTS = [
    ('ar',      ('1122',), '应收账款',   'asset', 'cust'),
    ('arn',     ('1121',), '应收票据',   'asset', 'cust'),
    ('ca',      ('1124',), '合同资产',   'asset', 'cust'),
    ('ora',     ('1221',), '其他应收款', 'asset', 'cust'),
    ('petty',   ('1222',), '备用金',     'asset', 'cust'),
    ('prepay',  ('1123',), '预付账款',   'asset', 'supp'),
    ('ap',      ('2202',), '应付账款',   'liab',  'supp'),
    ('apn',     ('2201',), '应付票据',   'liab',  'supp'),
    ('orp',     ('2241',), '其他应付款', 'liab',  'supp'),
    ('advrecv', ('2203',), '预收账款',   'liab',  'cust'),
    ('cl',      ('2205',), '合同负债',   'liab',  'supp'),
]


def build_one(data_dir, out_dir, key, pfx, name, nature, src, year='2026'):
    """单科目集团合并底稿（子进程隔离调用）。"""
    want_asset = (src == 'cust')
    comps = sorted(SR.discover_sap_entities(data_dir).keys())
    tb = C.load_tb(data_dir)
    os.makedirs(out_dir, exist_ok=True)

    # ⚡ 预计算全部主体数据（只扫一次 GL/aux/行项目，各 sheet 复用缓存防 OOM）
    is_liab = (nature == 'liab')
    pdata = {}   # comp -> dict(stat, inc, dec, exc, begin, tb_end, aux, lines, opp)
    for comp in comps:
        stat, inc, dec, exc_rows = SCA.analyze_subject(data_dir, comp, pfx[0], is_liab)
        tb_end = tb.get(name, {}).get(comp, 0.0)
        qc_tb, _qm_tb, qc_exist = C.tb_begin_end(data_dir, comp, [name])
        begin = qc_tb if qc_exist else (tb_end - inc + dec)
        if abs(tb_end) < 0.5 and not stat:
            continue   # 无数据主体跳过
        aux = [a for a in SCA._aux_rows(data_dir, comp, want_asset)
               if a[0].startswith(pfx[0])]
        lines = [l for l in SCA._line_items(data_dir, comp, want_asset)
                 if l[0].startswith(pfx[0])]
        opp = SCA._opp_pairs(data_dir, comp, pfx[0]) if aux else []
        pdata[comp] = dict(stat=stat, inc=inc, dec=dec, exc=exc_rows,
                           begin=begin, tb_end=tb_end, aux=aux, lines=lines, opp=opp)
    if not pdata:
        wb = openpyxl.Workbook()
        wb.close()
        return None

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ============ Sheet1 审定表（88 主体每主体一行） ============
    ws1 = wb.create_sheet('审定表')
    _hdr(ws1, ['公司', '期初余额', '本期增加', '本期减少', '期末余额', '期末(TB)'], f'{name}（{pfx[0]}）')
    r1 = 3
    grp1 = [0.0] * 5
    for comp, d in pdata.items():
        _txt(ws1, r1, 1, comp)
        _money(ws1, r1, 2, d['begin'])
        _money(ws1, r1, 3, d['inc'])
        _money(ws1, r1, 4, d['dec'])
        _money(ws1, r1, 5, d['tb_end'])
        _money(ws1, r1, 6, d['tb_end'])
        for i in range(5):
            grp1[i] += [d['begin'], d['inc'], d['dec'], d['tb_end'], d['tb_end']][i]
        r1 += 1
    _txt(ws1, r1, 1, '集团合计数', bold=True, fill=TOTFILL)
    for i in range(5):
        _money(ws1, r1, 2 + i, grp1[i], TOTFILL, bold=True)
    _fmt(ws1, [10, 16, 16, 16, 16, 16])
    ws1.freeze_panes = 'A2'

    # ============ Sheet2 增减构成（每主体一块） ============
    ws2 = wb.create_sheet('增减构成')
    _hdr(ws2, ['核算主体', '增减类别', '金额'], f'{name}（{pfx[0]}）')
    r2 = 3
    for comp, d in pdata.items():
        if not d['stat']:
            continue
        _block_hdr(ws2, r2, 3, f'【{comp}】')
        r2 += 1
        for k, v in sorted(d['stat'].items(), key=lambda x: -abs(x[1])):
            _txt(ws2, r2, 1, comp)
            _txt(ws2, r2, 2, k)
            _money(ws2, r2, 3, v)
            r2 += 1
    _fmt(ws2, [12, 30, 16])

    # ============ Sheet3 异常凭证（每主体一块） ============
    ws3 = wb.create_sheet('异常凭证')
    _hdr(ws3, ['核算主体', '凭证编号', '科目净额影响', '凭证文本', '供应商', '业务性质'], f'{name}（{pfx[0]}）')
    r3 = 3
    for comp, d in pdata.items():
        if not d['exc']:
            continue
        _block_hdr(ws3, r3, 6, f'【{comp}】')
        r3 += 1
        for vno, netv, txtv, suppv, naturev in d['exc']:
            _txt(ws3, r3, 1, comp)
            _txt(ws3, r3, 2, vno)
            _money(ws3, r3, 3, netv)
            _txt(ws3, r3, 4, txtv)
            _txt(ws3, r3, 5, suppv)
            c = _txt(ws3, r3, 6, naturev)
            c.fill = PatternFill('solid', fgColor='E2EFDA' if naturev.startswith('正常') else 'FFC7CE')
            r3 += 1
    _fmt(ws3, [12, 14, 16, 40, 14, 34])

    # ============ Sheet4 明细表（每主体一块：单位×期初/借/贷/期末/平衡） ============
    ws4 = wb.create_sheet(f'{name}明细表')
    hdrs4 = ['核算主体', '序号', '往来单位编号', '往来单位名称', '款项性质', '期初余额',
             '本期借方发生额', '本期贷方发生额', '期末余额', '平衡校验(期初+借-贷-期末)']
    _hdr(ws4, hdrs4, f'{name}（{pfx[0]}）')
    r4 = 3
    grp4 = [0.0] * 4
    for comp, d in pdata.items():
        aux_rows = d['aux']
        if not aux_rows:
            continue
        _block_hdr(ws4, r4, 10, f'【{comp}】 {name} 明细（单位×合同号）')
        r4 += 1
        seq = 0
        sub4 = [0.0] * 4
        for code, nm, cc, cn, con, qc, jf, df, qm in aux_rows:
            seq += 1
            bal = round(qc + jf - df - qm, 2)
            vals = [comp, seq, cc, cn, con or '—', nm, qc, jf, df, qm, bal]
            for j, v in enumerate(vals, 1):
                c = ws4.cell(r4, j, v)
                if isinstance(v, (int, float)):
                    c.number_format = NUM
                c.border = BORDER
                c.alignment = RGT if j >= 7 else LFT
            sub4[0] += qc; sub4[1] += jf; sub4[2] += df; sub4[3] += qm
            r4 += 1
        # 主体小计
        _txt(ws4, r4, 1, comp, bold=True)
        _txt(ws4, r4, 2, '小计', bold=True)
        for j, v in zip((7, 8, 9, 10), sub4):
            _money(ws4, r4, j, v, TOTFILL, bold=True)
        _txt(ws4, r4, 11, round(sub4[0] + sub4[1] - sub4[2] - sub4[3], 2), bold=True)
        for j in range(1, 12):
            ws4.cell(r4, j).fill = TOTFILL
        r4 += 1
        # TB 兜底占位（aux≠TB）
        d_qc = d['begin'] - sub4[0]
        d_qm = d['tb_end'] - sub4[3]
        if abs(d_qc) > 0.005 or abs(d_qm) > 0.005:
            _txt(ws4, r4, 1, comp, bold=True)
            _txt(ws4, r4, 3, '（无辅助核算明细）', bold=True)
            _txt(ws4, r4, 5, 'TB差额：该部分无单位辅助核算明细', bold=True)
            for j, v in zip((7, 8, 9, 10), (round(d_qc, 2), 0.0, 0.0, round(d_qm, 2))):
                _money(ws4, r4, j, v, bold=True)
            r4 += 1
        for i in range(4):
            grp4[i] += sub4[i]
        r4 += 1   # 主体间空行
    # 集团合计
    _txt(ws4, r4, 1, '集团合计数', bold=True, fill=TOTFILL)
    for j, v in zip((7, 8, 9, 10), grp4):
        _money(ws4, r4, j, v, TOTFILL, bold=True)
    _fmt(ws4, [10, 6, 14, 26, 16, 18, 14, 14, 14, 14, 22])
    ws4.freeze_panes = 'A2'

    # ============ Sheet5 账龄分布（每主体一块） ============
    # ⚡ 2026-08-16 旧『账龄分布』已废弃：由专项 aging_review.py（应收+合同资产联动）替代，
    #   SAP（AH）迁移后接入专项；此处不再输出账龄分布 sheet

    # ============ Sheet6 Top10余额（每主体一块） ============
    ws6 = wb.create_sheet('Top10余额')
    _hdr(ws6, ['核算主体', '序号', '往来单位编号', '往来单位名称', '合同号', '款项性质',
               '期初余额', '本期借方', '本期贷方', '期末余额'], f'{name}（{pfx[0]}）')
    r6 = 3
    for comp, d in pdata.items():
        aux_rows = d['aux']
        if not aux_rows:
            continue
        _block_hdr(ws6, r6, 10, f'【{comp}】 期末 Top10')
        r6 += 1
        for i, (code, nm, cc, cn, con, qc, jf, df, qm) in enumerate(
                sorted(aux_rows, key=lambda x: -abs(x[8]))[:10], 1):
            vals = [comp, i, cc, cn, con or '—', nm, qc, jf, df, qm]
            for j, v in enumerate(vals, 1):
                c = ws6.cell(r6, j, v)
                if isinstance(v, (int, float)):
                    c.number_format = NUM
                c.border = BORDER
                c.alignment = RGT if j >= 7 else LFT
            r6 += 1
        r6 += 1
    _fmt(ws6, [12, 6, 14, 26, 16, 18, 14, 14, 14, 14])
    ws6.freeze_panes = 'A2'

    # ============ Sheet7 对方科目核对（每主体一块） ============
    ws7 = wb.create_sheet('对方科目核对')
    _hdr(ws7, ['核算主体', '对方科目编码', '对方科目名称', '借方金额', '贷方金额', '笔数', '说明'], f'{name}（{pfx[0]}）')
    r7 = 3
    for comp, d in pdata.items():
        opp = d['opp']
        if not opp:
            continue
        _block_hdr(ws7, r7, 7, f'【{comp}】')
        r7 += 1
        for ocode, oname, oj, oc, n in opp[:20]:
            _txt(ws7, r7, 1, comp)
            _txt(ws7, r7, 2, ocode)
            _txt(ws7, r7, 3, oname)
            _money(ws7, r7, 4, oj)
            _money(ws7, r7, 5, oc)
            _txt(ws7, r7, 6, n)
            _txt(ws7, r7, 7, '凭证级配对（同凭证本科目净额）')
            r7 += 1
    _fmt(ws7, [12, 14, 30, 14, 14, 8, 34])

    # finalize：底稿说明 + 审计调整分录
    wb._meta = dict(comp='集团', label=name, period='2026年度1-7月')
    try:
        from audit_shell import finalize_workbook as _fw
        _fw(wb)
    except Exception:
        pass
    fp = os.path.join(out_dir, f'{name}往来款集团合并底稿.xlsx')
    wb.save(fp)
    wb.close()
    print(f'  ✅ {name}: 主体数={len(pdata)} 期末集团合计={sum(grp1[3:4]):,.2f}')
    return name


def _hdr(ws, hdrs, code_name=None):
    """表头；code_name 给定 → 首行标注科目代码/名称。"""
    if code_name:
        c = ws.cell(1, 1, code_name)
        c.font = BF
        c.fill = TITLEFILL
        c.alignment = LFT
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(2 if code_name else 1, j, h)
        c.font = C.HEAD_FONT
        c.fill = HFILL
        c.alignment = CTR
        c.border = BORDER


def _block_hdr(ws, r, ncol, label):
    """主体块标题行（横跨前 ncol 列，浅蓝底）。"""
    c = ws.cell(r, 1, label)
    c.font = BF
    c.fill = BLOCKFILL
    for j in range(1, ncol + 1):
        ws.cell(r, j).fill = BLOCKFILL
        ws.cell(r, j).border = BORDER
    return r + 1


def _txt(ws, r, c, v, bold=False, fill=None):
    cell = ws.cell(r, c, v)
    cell.font = BF if bold else F10
    if fill:
        cell.fill = fill
    cell.border = BORDER
    cell.alignment = LFT
    return cell


def _money(ws, r, c, v, fill=None, bold=False):
    cell = ws.cell(r, c, round(v, 2) if isinstance(v, (int, float)) else v)
    cell.number_format = NUM
    cell.font = BF if bold else F10
    cell.alignment = RGT
    cell.border = BORDER
    if fill:
        cell.fill = fill
    return cell


def _fmt(ws, widths):
    for j, w in enumerate(widths, 1):
        n = j - 1
        letters = ''
        while True:
            letters = chr(65 + n % 26) + letters
            n = n // 26 - 1
            if n < 0:
                break
        ws.column_dimensions[letters].width = w


def main(argv=None):
    argv = argv if argv is not None else sys.argv
    out_dir = None
    if '--out' in argv:
        out_dir = argv[argv.index('--out') + 1]
    out_dir = out_dir or OUT_DIR
    os.makedirs(out_dir, exist_ok=True)
    import subprocess
    PY = sys.executable
    ok = 0
    for key, pfx, name, nature, src in SUBJECTS:
        r = subprocess.run([PY, '-u', '-c',
                            f"import sys; sys.path.insert(0, {P.APP!r});"
                            "import sap_current_account_merge as M;"
                            f"M.build_one(r'{DATA_DIR}', r'{out_dir}', '{key}', {pfx!r}, '{name}', '{nature}', '{src}')"],
                           cwd=HERE, capture_output=True, text=True,
                           encoding='utf-8', errors='replace', timeout=3600)
        out = (r.stdout or '') + (r.stderr or '')
        if r.returncode == 0 and '✅' in out:
            ok += 1
            print(f'  ✅ {name}', flush=True)
        else:
            print(f'  ❌ {name}: {out[-400:]}', flush=True)
    print(f'完成 {ok}/{len(SUBJECTS)} 科目 → {out_dir}')
    return 0 if ok == len(SUBJECTS) else 2


if __name__ == '__main__':
    main()
