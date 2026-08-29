# -*- coding: utf-8 -*-
"""sap_gl_diff_report.py —— SAP GL vs TB 差异清单生成器（三口径取最优）

方法论（2026-08-08 凌晨验证固化）：
  - 余额/净额口径可靠：GL（期间01-07）期末 = TB 期初+借-贷，34+ 家 0 差异
  - 发生额差异 = 镜像虚增虚减，无法自动剔除干净；PK04/02（未清项+冲销冲回）仅在部分公司有效
  - **每科目三口径比较取最优**：全量GL / 剔PK04-02后 / TB，取差异更小者；仍差异 → 列凭证级清单
  - 凭证级聚合键必须 (comp, vno)（SAP 各公司凭证号独立编号，会串味）

用法：
  python sap_gl_diff_report.py <data_dir> <comp[,comp...]|*> [--out <xlsx>]
  python sap_gl_diff_report.py <DATA_ROOT>/yy 3010,2200
  单公司单进程跑（沙箱内存限制：1 进程读 1 批文件），批量由外层循环调度。

输出 xlsx sheets：
  ① 科目三口径汇总  全量GL借/贷、剔后借/贷、TB借/贷、采用口径、状态
  ② 被剔凭证清单    PK04/02 成对凭证（凭证级：期间/凭证号/借/贷/文本/PK/系统词）
  ③ 差异科目明细    三口径均不对平的科目差额 + 涉及凭证行
"""
import paths as P
import os
import re
import sys
from collections import defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import _safx
from sap_gl_tb_recon import read_tb_by_code

# 系统结转特征词（标注用，不用于剔除）
SYS_KW = ('未清项', '清账', '清帐', '冲销', '冲回', '资金', '上收', '下拨',
          '结转', '暂估', '调整', '评估', '重置', '计提', '划款')

# 序时账所需列
KEEP = {1, 4, 5, 6, 10, 11, 12, 21, 29}


def _gl_files(data_dir, comp):
    """返回该公司的序时账文件列表（单公司前缀优先，否则 4 位区间合并文件）。"""
    gl_dir = os.path.join(data_dir, '序时账')
    if not os.path.isdir(gl_dir):
        gl_dir = data_dir
    single, merged = [], []
    for f in sorted(os.listdir(gl_dir)):
        if not f.lower().endswith(('.xlsx', '.xls')):
            continue
        base = os.path.splitext(f)[0]
        m = re.match(r'^(\d{4})', base)
        if not m:
            continue
        if m.group(1) == comp:
            single.append(os.path.join(gl_dir, f))
        m2 = re.match(r'^(\d{4})[~-](\d{4})$', base)
        if m2 and m2.group(1) <= comp <= m2.group(2):
            merged.append(os.path.join(gl_dir, f))
    return single if single else merged


def _norm_code(name):
    """科目余额表名称尾部内部码 → 纯科目代码。"""
    mm = re.match(r'^(.*?)\s{2,}([A-Z]+\d*/\d+|\d+)$', name)
    return (mm.group(2).split('/')[-1] if mm else ''), (mm.group(1).strip() if mm else name)


def scan_gl(data_dir, comp):
    """流式扫描该公司全部序时账 → 科目级借/贷聚合 + 凭证级聚合 + 凭证文本缓存。
    返回 (subj_agg, vno_agg, vno_txt)
      subj_agg: {subj: [jf, df]}            全量行级
      vno_agg:  {vno: [tot_jf, tot_df, pk04, pk02]}   凭证级（判剔除）
      vno_txt:  {vno: (period, first_txt, cust, supp)}
    """
    subj_agg = defaultdict(lambda: [0.0, 0.0])
    vno_agg = {}
    vno_txt = {}
    for fp in _gl_files(data_dir, comp):
        for row in _safx.iter_rows(fp, keep_cols=KEEP):
            if str(row.get(6, '')).strip() != comp:
                continue
            per = str(row.get(10, '')).strip()
            if not (per[:2].isdigit() and 1 <= int(per[:2]) <= 7):
                continue
            subj = str(row.get(11) or '')
            if not subj:
                continue
            amt = float(row.get(29) or 0)
            vno = str(row.get(1) or '').strip()
            if not vno:
                continue
            a = subj_agg[subj]
            if amt >= 0:
                a[0] += amt
            else:
                a[1] += -amt
            v = vno_agg.get(vno)
            if v is None:
                v = [0.0, 0.0, 0, 0]
                vno_agg[vno] = v
            if amt >= 0:
                v[0] += amt
            else:
                v[1] += -amt
            pk = str(row.get(4) or '').strip()
            if pk == '04':
                v[2] = 1
            elif pk == '02':
                v[3] = 1
            if vno not in vno_txt:
                txt = str(row.get(5) or '').strip()
                cust = str(row.get(12) or '').strip()
                supp = str(row.get(21) or '').strip()
                vno_txt[vno] = (per, txt[:40], cust, supp)
    return subj_agg, vno_agg, vno_txt


def drop_vnos(vno_agg):
    """凭证级净额=0 且含 PK04/02 → 被剔凭证。"""
    return {v for v, (jf, df, p4, p2) in vno_agg.items()
            if abs(jf - df) < 0.01 and jf > 0.01 and (p4 or p2)}


def _sys_flags(txt):
    hit = [kw for kw in SYS_KW if kw in txt]
    return hit


# ============================================================
# v2：精确版（被剔凭证按科目拆行需要行级数据，做两遍扫描：
#      第一遍 科目全量 + 凭证判定；第二遍 只收集被剔凭证中差异科目的行）
# ============================================================
def recon_company_v2(data_dir, comp, threshold=100.0):
    """两遍扫描精确版：
    pass1: 科目全量借/贷 + 凭证级判定被剔凭证
    pass2: 收集被剔凭证中每科目的借/贷（行级），得到剔后科目借/贷
    """
    subj_agg, vno_agg, vno_txt = scan_gl(data_dir, comp)
    drops = drop_vnos(vno_agg)
    tb = read_tb_by_code(data_dir, comp)

    # pass2：被剔凭证中按科目行级聚合
    drop_subj = defaultdict(lambda: [0.0, 0.0])
    drop_rows = defaultdict(list)   # subj -> [(per, vno, jf, df, txt, pk, flags)]
    for fp in _gl_files(data_dir, comp):
        for row in _safx.iter_rows(fp, keep_cols=KEEP):
            if str(row.get(6, '')).strip() != comp:
                continue
            per = str(row.get(10, '')).strip()
            if not (per[:2].isdigit() and 1 <= int(per[:2]) <= 7):
                continue
            vno = str(row.get(1) or '').strip()
            if vno not in drops:
                continue
            subj = str(row.get(11) or '')
            if not subj:
                continue
            amt = float(row.get(29) or 0)
            a = drop_subj[subj]
            txt = str(row.get(5) or '').strip()
            pk = str(row.get(4) or '').strip()
            if amt >= 0:
                a[0] += amt
                drop_rows[subj].append((per, vno, amt, 0.0, txt, pk))
            else:
                a[1] += -amt
                drop_rows[subj].append((per, vno, 0.0, -amt, txt, pk))

    rows = []
    diff_detail = []
    all_subjs = set(subj_agg) | set(tb) | set(drop_subj)
    for subj in sorted(all_subjs):
        g = subj_agg.get(subj, [0.0, 0.0])
        d = drop_subj.get(subj, [0.0, 0.0])
        t = tb.get(subj, {'jf': 0.0, 'df': 0.0})
        after_jf = g[0] - d[0]
        after_df = g[1] - d[1]
        df_full = g[0] - t['jf']
        df_after = after_jf - t['jf']
        # 采用口径：取借差更小者；状态判定
        if abs(df_full) <= threshold and abs(df_after) <= threshold:
            adopt, status = '对平', '✅ 对平'
        elif abs(df_full) < abs(df_after):
            adopt, status = '全量', '差异'
        else:
            adopt, status = '剔后', '差异'
        rows.append((subj, g[0], g[1], d[0], d[1], after_jf, after_df,
                     t['jf'], t['df'], df_full, df_after, adopt, status))
        # 差异科目 → 收集凭证行（被剔行 + 全量行里非成对的大额行）
        if status != '✅ 对平' and abs(df_full) > threshold:
            for r in drop_rows.get(subj, []):
                per, vno, jf, df_, txt, pk = r
                diff_detail.append((subj, per, vno, jf, df_, txt, pk,
                                    '剔', ','.join(_sys_flags(txt))))
    return rows, drops, vno_txt, vno_agg, diff_detail


def _load_subj_names(data_dir, comp):
    """从科目余额表构建 代码→名称 映射。"""
    import openpyxl
    names = {}
    tb_dir = os.path.join(data_dir, '科目余额表')
    if not os.path.isdir(tb_dir):
        return names
    for f in sorted(os.listdir(tb_dir)):
        if not f.lower().endswith('.xlsx'):
            continue
        try:
            wb = openpyxl.load_workbook(os.path.join(tb_dir, f), read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=3, values_only=True):
            if not r or len(r) < 9 or r[2] is None:
                continue
            cr = str(r[7]).strip() if r[7] else ''
            m = re.search(r'(\d{4})\s*$', cr)
            if not m or m.group(1) != comp:
                continue
            code, n = _norm_code(str(r[2]).strip())
            if code and code not in names:
                names[code] = n
        wb.close()
    return names


def write_report(data_dir, comp, out_path=None):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    rows, drops, vno_txt, vno_agg, diff_detail = recon_company_v2(data_dir, comp)
    names = _load_subj_names(data_dir, comp)
    if not rows:
        print(f'  ⚠️ [{comp}] 无 GL 数据，跳过。')
        return None

    if out_path is None:
        out_dir = os.path.join(data_dir, '底稿')
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, f'GL差异清单_{comp}.xlsx')

    BOLD = Font(bold=True)
    HDR_FILL = PatternFill('solid', fgColor='D9E1F2')
    TOT_FILL = PatternFill('solid', fgColor='FCE4D6')
    BAD_FILL = PatternFill('solid', fgColor='FFC7CE')
    RGT = Alignment(horizontal='right')
    CTR = Alignment(horizontal='center')
    THIN = Border(*[Side(style='thin')] * 4)

    wb = openpyxl.Workbook()

    # ① 科目三口径汇总
    ws = wb.active
    ws.title = '科目三口径汇总'
    hdrs = ['科目代码', '科目名称', 'GL全量借', 'GL全量贷', '剔PK04/02借', '剔PK04/02贷',
            '剔后借', '剔后贷', 'TB借', 'TB贷', '全量借差', '剔后借差',
            '采用口径', '状态']
    ws.append(hdrs)
    for c in ws[1]:
        c.font = BOLD
        c.fill = HDR_FILL
        c.alignment = CTR
        c.border = THIN
    n_diff = 0
    for r in rows:
        subj, gj, gd, dj0, dd0, aj, ad, tj, td, dff, dfa, adopt, status = r
        ws.append([subj, names.get(subj, ''), gj, gd, dj0, dd0, aj, ad, tj, td, dff, dfa, adopt, status])
        row_i = ws.max_row
        for c in ws[row_i]:
            c.border = THIN
            c.number_format = '#,##0.00'
        if status != '✅ 对平':
            n_diff += 1
            for c in ws[row_i]:
                c.fill = BAD_FILL
    for col, w in zip('ABCDEFGHIJKLMN', [14, 26, 15, 15, 15, 15, 15, 15, 15, 15, 15, 15, 9, 9]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'

    # ② 被剔凭证清单
    ws2 = wb.create_sheet('被剔凭证清单')
    ws2.append(['期间', '凭证号', '借', '贷', 'PK', '文本', '系统词'])
    for c in ws2[1]:
        c.font = BOLD
        c.fill = HDR_FILL
        c.alignment = CTR
        c.border = THIN
    for vno in sorted(drops):
        v = vno_txt.get(vno, ('', '', '', ''))
        per, txt, cust, supp = v
        vg = vno_agg.get(vno, [0.0, 0.0, 0, 0])
        ws2.append([per, vno, vg[0], vg[1], '04/02', txt, ','.join(_sys_flags(txt))])
    ws2.append(['合计', '', '', '', '', '', ''])
    for col, w in zip('ABCDEFG', [7, 16, 15, 15, 7, 46, 18]):
        ws2.column_dimensions[col].width = w

    # ③ 差异科目明细
    ws3 = wb.create_sheet('差异科目明细')
    ws3.append(['科目代码', '科目名称', '期间', '凭证号', '借', '贷', '文本', 'PK', '类别', '系统词'])
    for c in ws3[1]:
        c.font = BOLD
        c.fill = HDR_FILL
        c.alignment = CTR
        c.border = THIN
    for r in diff_detail:
        subj, per, vno, jf, df_, txt, pk, cat, flags = r
        ws3.append([subj, names.get(subj, ''), per, vno, jf, df_, txt, pk, cat, flags])
    for col, w in zip('ABCDEFGHIJ', [14, 26, 7, 16, 15, 15, 40, 7, 7, 18]):
        ws3.column_dimensions[col].width = w

    wb.save(out_path)
    print(f'  ✅ [{comp}] 差异清单: {os.path.basename(out_path)}  科目{len(rows)} 差异{n_diff} 剔凭证{len(drops)}')
    return out_path


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    out = None
    if '--out' in sys.argv:
        out = sys.argv[sys.argv.index('--out') + 1]
    if len(args) < 2:
        print(__doc__)
        sys.exit(1)
    data_dir = args[0]
    comps_arg = args[1]
    comps = sorted({c for c in comps_arg.split(',') if c}) if comps_arg != '*' else \
        sorted({c for c in os.listdir(os.path.join(data_dir, '序时账')) if c[:4].isdigit()})
    for comp in comps:
        try:
            write_report(data_dir, comp, out)
        except Exception as e:
            import traceback
            print(f'  ❌ [{comp}] {type(e).__name__}: {e}')
            traceback.print_exc()
