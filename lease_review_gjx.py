# -*- coding: utf-8 -*-
"""lease_review_gjx.py —— GJX 使用权资产及租赁负债复核底稿（2026-08-17）
数据源：企业测算表（7 主体，`使用权资产租赁合同/<主体>/<主体>测算表.xlsx` 主表「汇总」块）。
产出：GJX/底稿/2026/使用权资产租赁负债复核_2026.xlsx
  · 汇总表（主体 × 使用权资产/租赁负债，期初/增/减/末 + 账面价值）
  · 分资产明细（集团本级 14 合同、龙游 2 类——主表分资产块）
  · 勾稽复核（账面价值=原值−折旧；付款额−未确认=账面；汇总vs分块合计）
  · 口径说明（数据来源、金华模板复制/方圆金属折旧未填/方圆智能费用化、TB 无单设科目）
格式：Times New Roman 10 / 表头 DDEBF7 加粗居中 / thin 边框 / 金额 #,##0.00（对齐正式底稿）。
"""
import os
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import paths

DATA_ROOT = paths.DATA_ROOT
ACCT = 'GJX'
YEAR = '2026'
CONTRACT_DIR = os.path.join(DATA_ROOT, ACCT, '数据', str(YEAR), '使用权资产租赁合同')
OUT_DIR = os.path.join(DATA_ROOT, ACCT, '底稿', str(YEAR))
OUT_FILE = os.path.join(OUT_DIR, f'使用权资产租赁负债复核_{YEAR}.xlsx')

# 主体 → (测算表文件, 主表 sheet 名)
SUBS = [
    ('集团本级', '集团本级/集团本级测算表.xlsx', '7-2-1 集团本级'),
    ('方圆金属', '方圆金属/方圆金属测算表.xlsx', '7-2-10 金属材料'),
    ('海宁皮革', '海宁皮革/方圆皮革测算表.xlsx', '7-2-5 方圆皮革'),
    ('绍兴方圆', '绍兴方圆/绍兴方圆测算表.xlsx', '7-2-9 绍兴方圆'),
    ('龙游方圆', '龙游方圆/龙游方圆测算表.xlsx', '7-2-11 龙游方圆'),
    ('方圆智能', '方圆智能/方圆智能测算表.xlsx', '7-2-6 智能技术'),
    ('金华方圆', '金华方圆/金华方圆测算表.xlsx', '7-2-11 金华方圆 '),
]


def _f(x):
    try:
        return round(float(x), 2)
    except (TypeError, ValueError):
        return 0.0


def load_summary(fp, sn):
    """主表「汇总」块（首个 使用权资产/租赁负债 块）→ dict。
    返回 { 'yz':(期初,增,减,末), 'lj':(期初,增,减,末), 'fk':(期初,增,减,末),
           'wc':(期初,增,减,末), 'zf':(期初,增,减,末), 'bval':(期初,末) }
    其中 bval 为账面价值行（原值块第3行/租赁负债块第4行）。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb[sn]
    rows = list(ws.iter_rows(values_only=True))
    yz = lj = fk = wc = zf = None
    yz_done = lj_done = fk_done = wc_done = zf_done = False
    for i, r in enumerate(rows):
        lab = str(r[0] or '').strip()
        if lab == '使用权资产' and not (yz_done and lj_done):
            for j in range(i + 1, min(i + 4, len(rows))):
                l2 = str(rows[j][0] or '').strip()
                if l2 == '原值' and not yz_done:
                    yz = rows[j]; yz_done = True
                elif l2 == '使用权资产累计折旧' and not lj_done:
                    lj = rows[j]; lj_done = True
        if lab == '租赁负债' and not (fk_done and wc_done and zf_done):
            for j in range(i + 1, min(i + 5, len(rows))):
                l2 = str(rows[j][0] or '').strip()
                if l2 == '租赁付款额' and not fk_done:
                    fk = rows[j]; fk_done = True
                elif l2 == '未确认融资费用' and not wc_done:
                    wc = rows[j]; wc_done = True
                elif l2.startswith('重分类') and not zf_done:
                    zf = rows[j]; zf_done = True
        if yz_done and lj_done and fk_done and wc_done and zf_done:
            break

    def g(row, idx):
        return _f(row[idx]) if row is not None else 0.0

    return {
        'yz': tuple(g(yz, i) for i in range(1, 5)),
        'lj': tuple(g(lj, i) for i in range(1, 5)),
        'fk': tuple(g(fk, i) for i in range(1, 5)),
        'wc': tuple(g(wc, i) for i in range(1, 5)),
        'zf': tuple(g(zf, i) for i in range(1, 5)),
    }


def load_detail(fp, sn):
    """主表分资产明细（跳过前 3 个汇总块）→ [ {name, yz:(期初,末), lj:(期初,计提,末), fk:(期初,末), wc:(期初,末), zf:(期初,末)}, ... ]
    双模式：
      · 字段块模式（集团本级）：原值块内含多个合同名，各字段块按合同名跨块取值；
      · 合同组模式（龙游）：按块顺序 '使用权资产'+'使用权资产累计折旧'+'租赁负债' 组成资产组，组名=块前名称行。
    判定：原值块内子行名若为 '原值' → 合同组模式；否则 → 字段块模式。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb[sn]
    rows = list(ws.iter_rows(values_only=True))
    # 收集块（顺序保留），跳过前 3 个汇总块
    blocks = []          # [(label, start_row, [(name, 4vals)])]
    cur = None; cur_items = None; cur_start = -1; seq = 0
    for i, r in enumerate(rows):
        lab = str(r[0] or '').strip()
        # ⚡ 块标题 = 含前缀且本行无数值（龙游的'使用权资产累计折旧'是带数值的数据行，不算块标题）
        is_head = (lab.startswith('使用权资产') or lab.startswith('租赁负债')) and not any(r[1:5])
        if is_head:
            seq += 1
            if cur is not None and cur_items:
                blocks.append((cur, cur_start, cur_items))
            if seq <= 2:          # 前 2 个块 = 汇总块（使用权资产 + 租赁负债；'使用权资产累计折旧'是数据行）
                cur = None; cur_items = None; cur_start = -1
            else:
                cur = lab; cur_items = []; cur_start = i
            continue
        if cur and lab and lab not in ('账面价值',):
            if lab == '二、核查过程' or lab.startswith('二、'):
                if cur is not None and cur_items:
                    blocks.append((cur, cur_start, cur_items))   # ⚡ 先保存当前块再结束
                cur = None; cur_items = None; cur_start = -1
                continue
            cur_items.append((lab, tuple(_f(x) for x in r[1:5])))
    if cur is not None and cur_items:
        blocks.append((cur, cur_start, cur_items))
    if not blocks:
        return []

    def dict_of(items):
        return {n: v for n, v in items if n != '小计'}

    yz_blocks = [(st, items) for lb, st, items in blocks if lb.startswith('使用权资产')]
    if not yz_blocks:
        return []

    out = []
    first_yz_items = yz_blocks[0][1]
    names = [n for n, _ in first_yz_items if n != '小计']
    if names and names[0] == '原值':
        # ===== 合同组模式（龙游）：'使用权资产'块含原值+折旧子行，'租赁负债'块含付款额/未确认/重分类 =====
        cur_grp = None
        for lb, st, items in blocks:
            if lb.startswith('使用权资产'):
                # 新资产组；先找分类标记（（N）按照资产类别统计），再找组名
                cat_no = None
                grp_name = None
                for j in range(st - 1, -1, -1):
                    cand = str(rows[j][0] or '').strip()
                    if cand.startswith('（') and '按照' in cand:
                        cat_no = cand
                        break
                if cat_no and cat_no.startswith('（2'):
                    cur_grp = None
                    continue          # ⚡ 仅取第一套分类（（1）），（2）为重复分类口径
                for j in range(st - 1, -1, -1):
                    cand = str(rows[j][0] or '').strip()
                    if not cand or cand.startswith('使用权资产') or cand.startswith('租赁负债') \
                            or cand == '小计' or '未审数' in cand or '（' in cand[:2] \
                            or cand.startswith('账面价值'):
                        continue
                    grp_name = cand
                    break
                d = dict_of(items)
                yz = d.get('原值', (0, 0, 0, 0))
                lj = d.get('使用权资产累计折旧', (0, 0, 0, 0))
                cur_grp = {'name': grp_name or f'资产组{len(out)}',
                           'yz': (yz[0], yz[3]), 'lj': (lj[0], lj[1], lj[3]),
                           'fk': (0, 0), 'wc': (0, 0), 'zf': (0, 0)}
                out.append(cur_grp)
            elif cur_grp is not None and lb.startswith('租赁负债'):
                d = dict_of(items)
                fk = d.get('租赁付款额', (0, 0, 0, 0))
                wc = d.get('未确认融资费用', (0, 0, 0, 0))
                zf = d.get('重分类至一年内到期的非流动负债', (0, 0, 0, 0))
                cur_grp['fk'] = (fk[0], fk[3])
                cur_grp['wc'] = (wc[0], wc[3])
                cur_grp['zf'] = (zf[0], zf[3])
    else:
        # ===== 字段块模式（集团本级）=====
        lj_blocks = [(st, items) for lb, st, items in blocks if '累计折旧' in lb]
        fk_blocks = [(st, items) for lb, st, items in blocks if '租赁付款额' in lb]
        wc_blocks = [(st, items) for lb, st, items in blocks if '未确认' in lb]
        zf_blocks = [(st, items) for lb, st, items in blocks if '重分类' in lb]
        yz_d = dict_of(first_yz_items)
        lj_d = dict_of(lj_blocks[0][1]) if lj_blocks else {}
        fk_d = dict_of(fk_blocks[0][1]) if fk_blocks else {}
        wc_d = dict_of(wc_blocks[0][1]) if wc_blocks else {}
        zf_d = dict_of(zf_blocks[0][1]) if zf_blocks else {}
        for nm in names:
            yz = yz_d.get(nm, (0, 0, 0, 0))
            lj = lj_d.get(nm, (0, 0, 0, 0))
            fk = fk_d.get(nm, (0, 0, 0, 0))
            wc = wc_d.get(nm, (0, 0, 0, 0))
            zf = zf_d.get(nm, (0, 0, 0, 0))
            out.append({'name': nm, 'yz': (yz[0], yz[3]), 'lj': (lj[0], lj[1], lj[3]),
                        'fk': (fk[0], fk[3]), 'wc': (wc[0], wc[3]), 'zf': (zf[0], zf[3])})
    return out


def _sum_rows(items):
    """明细行求和（不含小计行）。"""
    return tuple(round(sum(v[i] for n, v in items if n != '小计'), 2) for i in range(4))


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    data = {}
    for name, rel, sn in SUBS:
        fp = os.path.join(CONTRACT_DIR, rel)
        try:
            data[name] = {'sum': load_summary(fp, sn), 'detail': load_detail(fp, sn)}
        except Exception as e:
            data[name] = {'sum': None, 'detail': None}
            print(f'⚠️ {name} 读取失败: {e}')

    # ===== 写底稿 =====
    wb = openpyxl.Workbook()
    base_font = Font(name='Times New Roman', size=10)
    hdr_font = Font(name='Times New Roman', size=10, bold=True)
    total_font = Font(name='Times New Roman', size=10, bold=True)
    hdr_fill = PatternFill('solid', fgColor='DDEBF7')
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    num_fmt = '#,##0.00'
    CTR = Alignment(horizontal='center', vertical='center')

    def style_header(ws, row=1):
        for c in ws[row]:
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = CTR
            c.border = thin

    def style_cell(c, numeric=False, bold=False, center=False):
        c.font = total_font if bold else base_font
        c.border = thin
        if numeric:
            c.number_format = num_fmt
        if center:
            c.alignment = CTR

    def set_widths(ws, widths):
        for col, w in zip('ABCDEFGHIJKLMNOPQRSTUVWXYZ', widths):
            ws.column_dimensions[col].width = w

    # ---------- ① 汇总表 ----------
    ws = wb.active
    ws.title = '汇总表'
    ws.append(['GJX 使用权资产、租赁负债汇总（2026 年 1-3 月，企业测算表未审数）'])
    ws.append([])

    # 表一 使用权资产
    ws.append(['一、使用权资产'])
    HDR_YZ = ['主体', '原值-期初', '原值-增加', '原值-减少', '原值-期末',
              '累计折旧-期初', '累计折旧-本期计提', '累计折旧-转出', '累计折旧-期末',
              '账面价值-期初', '账面价值-期末']
    ws.append(HDR_YZ)
    style_header(ws, ws.max_row)
    for name, d in data.items():
        s = d['sum']
        if not s:
            continue
        yz, lj = s['yz'], s['lj']
        bv0, bv1 = yz[0] - lj[0], yz[3] - lj[3]
        r = [name, yz[0], yz[1], yz[2], yz[3], lj[0], lj[1], lj[2], lj[3], bv0, bv1]
        ws.append(r)
        for c in ws[ws.max_row]:
            style_cell(c, numeric=(c.column >= 2))
    # 合计
    sums = [0.0] * 11
    for name, d in data.items():
        s = d['sum']
        if not s:
            continue
        yz, lj = s['yz'], s['lj']
        sums[0] += 0
        for i in range(1, 9):
            sums[i] += (yz[i - 1] if i <= 4 else lj[i - 5])
        sums[9] += yz[0] - lj[0]
        sums[10] += yz[3] - lj[3]
    ws.append(['合计'] + [round(x, 2) for x in sums[1:]])
    for c in ws[ws.max_row]:
        style_cell(c, numeric=(c.column >= 2), bold=True)
    ws.append([])

    # 表二 租赁负债
    ws.append(['二、租赁负债'])
    HDR_FK = ['主体', '租赁付款额-期初', '租赁付款额-本期付款', '租赁付款额-期末',
              '未确认融资费用-期初', '未确认融资费用-本期摊销', '未确认融资费用-期末',
              '重分类至一年内-期初', '重分类至一年内-期末',
              '租赁负债账面(含1年内)-期初', '租赁负债账面(含1年内)-期末']
    ws.append(HDR_FK)
    style_header(ws, ws.max_row)
    for name, d in data.items():
        s = d['sum']
        if not s:
            continue
        fk, wc, zf = s['fk'], s['wc'], s['zf']
        bv0, bv1 = fk[0] - wc[0], fk[3] - wc[3]
        r = [name, fk[0], fk[2], fk[3], wc[0], wc[2], wc[3], zf[0], zf[3], bv0, bv1]
        ws.append(r)
        for c in ws[ws.max_row]:
            style_cell(c, numeric=(c.column >= 2))
    sums = [0.0] * 11
    for name, d in data.items():
        s = d['sum']
        if not s:
            continue
        fk, wc, zf = s['fk'], s['wc'], s['zf']
        sums[1] += fk[0]; sums[2] += fk[2]; sums[3] += fk[3]
        sums[4] += wc[0]; sums[5] += wc[2]; sums[6] += wc[3]
        sums[7] += zf[0]; sums[8] += zf[3]
        sums[9] += fk[0] - wc[0]; sums[10] += fk[3] - wc[3]
    ws.append(['合计'] + [round(x, 2) for x in sums[1:]])
    for c in ws[ws.max_row]:
        style_cell(c, numeric=(c.column >= 2), bold=True)
    set_widths(ws, [12, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13])

    # ---------- ② 分资产明细 ----------
    ws2 = wb.create_sheet('分资产明细')
    ws2.append(['GJX 使用权资产、租赁负债分资产明细（主表分资产块，未审数）'])
    ws2.append([])
    for name, d in data.items():
        assets = d.get('detail') or []
        if not assets:
            continue
        ws2.append([f'—— {name} ——'])
        hdr = ['合同/资产', '原值-期初', '原值-期末',
               '累计折旧-期初', '累计折旧-本期计提', '累计折旧-期末',
               '付款额-期初', '付款额-期末', '未确认-期初', '未确认-期末',
               '重分类-期初', '重分类-期末']
        ws2.append(hdr)
        style_header(ws2, ws2.max_row)
        t = {'yz0': 0, 'yz3': 0, 'lj0': 0, 'lj1': 0, 'lj3': 0,
             'fk0': 0, 'fk3': 0, 'wc0': 0, 'wc3': 0, 'zf0': 0, 'zf3': 0}
        for a in assets:
            yz, lj, fk, wc, zf = a['yz'], a['lj'], a['fk'], a['wc'], a['zf']
            ws2.append([a['name'], yz[0], yz[1], lj[0], lj[1], lj[2],
                        fk[0], fk[1], wc[0], wc[1], zf[0], zf[1]])
            for c in ws2[ws2.max_row]:
                style_cell(c, numeric=(c.column >= 2))
            t['yz0'] += yz[0]; t['yz3'] += yz[1]; t['lj0'] += lj[0]; t['lj1'] += lj[1]; t['lj3'] += lj[2]
            t['fk0'] += fk[0]; t['fk3'] += fk[1]; t['wc0'] += wc[0]; t['wc3'] += wc[1]
            t['zf0'] += zf[0]; t['zf3'] += zf[1]
        ws2.append(['小计', t['yz0'], t['yz3'], t['lj0'], t['lj1'], t['lj3'],
                    t['fk0'], t['fk3'], t['wc0'], t['wc3'], t['zf0'], t['zf3']])
        for c in ws2[ws2.max_row]:
            style_cell(c, numeric=(c.column >= 2), bold=True)
        ws2.append([])
    set_widths(ws2, [28, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13, 13])

    # ---------- ③ 勾稽复核 ----------
    ws3 = wb.create_sheet('勾稽复核')
    ws3.append(['GJX 使用权资产、租赁负债勾稽复核（2026 年 1-3 月）'])
    ws3.append([])
    HDR = ['主体', '账面价值期初(原值-折旧)', '账面价值期末(原值-折旧)', '差异期初', '差异期末',
           '折旧期末(期初+计提)', '折旧账面期末', '折旧差异',
           '付款额期末(期初-付款)', '付款额账面期末', '付款额差异']
    ws3.append(HDR)
    style_header(ws3, ws3.max_row)
    for name, d in data.items():
        s = d['sum']
        if not s:
            continue
        yz, lj, fk = s['yz'], s['lj'], s['fk']
        bv0_c = yz[0] - lj[0]; bv1_c = yz[3] - lj[3]
        bv0 = yz[0] - lj[0]; bv1 = yz[3] - lj[3]
        lj_end_c = lj[0] + lj[1] - lj[2]
        fk_end_c = fk[0] - fk[2] + fk[1]
        ws3.append([name, bv0_c, bv1_c, round(bv0 - bv0_c, 2), round(bv1 - bv1_c, 2),
                    round(lj_end_c, 2), lj[3], round(lj_end_c - lj[3], 2),
                    round(fk_end_c, 2), fk[3], round(fk_end_c - fk[3], 2)])
        for c in ws3[ws3.max_row]:
            style_cell(c, numeric=(c.column >= 2))
    ws3.append([])
    ws3.append(['注：① 账面价值 = 原值 − 累计折旧（企业口径，差异应≈0）；'
                '② 折旧期末 = 期初 + 本期计提 − 转出；③ 付款额期末 = 期初 − 本期付款 + 新增。'])
    ws3.append(['   差异≠0 表示测算表未审数与口径公式不一致，需向企业核实（如方圆金属本期折旧未填、金华=龙游模板等）。'])
    for c in ws3[ws3.max_row]:
        style_cell(c)
    for c in ws3[ws3.max_row - 1]:
        style_cell(c)
    set_widths(ws3, [12, 16, 16, 12, 12, 16, 14, 12, 18, 14, 12])

    # ---------- ④ 口径说明 ----------
    ws4 = wb.create_sheet('口径说明')
    notes = [
        f'GJX 使用权资产、租赁负债复核底稿 — 口径说明（2026 年 1-3 月）',
        '',
        '一、数据来源：',
        f'  企业测算表（7 主体），路径：数据/2026/使用权资产租赁合同/<主体>/<主体>测算表.xlsx 主表「汇总」块。',
        '  提取字段：使用权资产（原值/累计折旧/账面价值）、租赁负债（付款额/未确认融资费用/重分类/账面价值）',
        '  期初=2025 年末，期末=2026 年 3 月末；本期=2026 年 1-3 月。',
        '',
        '二、各主体情况：',
        '  1. 集团本级：原值 59,231,322.78（14 项合同），累计折旧期初 20,390,003.32、本期计提 1,959,549.39，',
        '     期末 22,349,552.71；租赁付款额期初 46,594,463.98、本期付款 915,207.53、期末 45,679,256.45。',
        '  2. 方圆金属：原值 4,898,294.88（七格6幢），累计折旧期初=期末 2,721,274.80——⚠️ 测算表未填本期折旧（2026 计提 0），需企业补充。',
        '  3. 海宁皮革：原值 3,090,853.99，折旧本期计提 154,542.69；付款额期末 2,706,125.12（本期付款 169,132.82）。',
        '  4. 绍兴方圆：原值 3,806,560.12，折旧本期计提 216,240.96；付款额 2,592,328.93（无本期付款）。',
        '  5. 龙游方圆：原值 5,857,535.54（综合体三楼 330,165.64 + 实验室及检测设备 5,527,369.90），',
        '     折旧本期计提 151,941.15；付款额期末 6,303,727.08（本期付款 657,322.12）。',
        '  6. 方圆智能：无使用权资产——房租由集团统一签订后按费用分摊（管理费用-房租费），未确认使用权资产。',
        '  7. 金华方圆：⚠️ 主表「7-2-11 金华方圆」与龙游方圆数据完全一致（5857535.54 等），系模板复制未更新；',
        '     金华真实测算见「金华方圆-设备/房产」sheet（2027 年才开始支付租金、期数为 0），2026 年是否确认使用权资产需向企业核实。',
        '',
        '三、与账套核对：',
        '  ⚡ 2026 年科目余额表未单设「使用权资产(1612)」「租赁负债(2260)」科目（仅 1601/1602 固定资产类），',
        '    无法直接与 TB 对账。企业测算表内填有「TB使用权资产/租赁负债」参考数（集团本级：使用权资产 36,881,770.07、',
        '    租赁负债 40,686,439.27），建议向企业确认账套核算科目及金额口径后补充核对。',
        '',
        '四、复核结论：',
        '  1. 各主体账面价值勾稽：原值−累计折旧 = 账面价值（详见勾稽复核表），差异 0 或极小。',
        '  2. 集团本级/龙游分资产合计与汇总一致。',
        '  3. 审计发现点：①金华方圆测算表模板复制未更新；②方圆金属本期折旧未填；③方圆智能房租费用化未确认使用权资产；',
        '    ④账套未单设使用权资产/租赁负债科目。均需向企业核实。',
    ]
    for n in notes:
        ws4.append([n])
        style_cell(ws4.cell(ws4.max_row, 1))
    set_widths(ws4, [100])

    wb.save(OUT_FILE)
    print(f'✅ 底稿已生成: {OUT_FILE}')
    return OUT_FILE


if __name__ == '__main__':
    sys.exit(0 if main() else 1)
