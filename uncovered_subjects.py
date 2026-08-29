# -*- coding: utf-8 -*-
"""未覆盖科目清单（2026-08-07 用户需求：账套中有的数据都可以生成各类明细表，
使用者自行判断用不用——先把【注册表未覆盖的科目】全部列出来）。

扫描账套 TB 全部一级科目（4 位码或层级根），与 subjects_registry 的 codes + 名称
关键词比对，输出未覆盖科目清单（代码/名称/期末余额/发生额/建议归属 builder），
写入 账套目录/未覆盖科目清单.xlsx 并在控制台打印。

用途：跑完常规 13 类后，把 TB 里存在但未匹配任何 builder 的科目主动暴露，
使用者直接看清单决定要不要补——比靠 coverage_check 被动发现更彻底。

用法：
    python uncovered_subjects.py <账套目录>
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import audit_common as A
import subjects_registry as REG

# 注册表全部 codes（多账套并集）→ 判断某科目是否已被 builder 覆盖
REG_CODES = set()
for subj in REG.REGISTRY.values():
    for c in (subj.get('codes') or []):
        c = str(c)
        if c.isdigit():
            REG_CODES.add(c)

# builder 层另有名称关键词覆盖（current_account/pl_detail 等按名称匹配），
# 此处把注册表名称关键词也收集，做名称级判定
REG_NAMES = [s.get('name', '') for s in REG.REGISTRY.values() if s.get('name')]

# 已知无需生成底稿的科目族（损益结转过渡/权益内部/备抵已随主科目）
SKIP_PREFIX = {
    '4103',  # 本年利润（结转过渡）
    '4104',  # 利润分配（equity_detail 已生成未分配利润底稿，子目不再单列）
}
SKIP_NAME_KW = ('结转', '过渡', '辅助')

# ⚡ 2026-08-12 修复：并入主科目底稿的科目（current_account.RELATED_MAP 机制）——
#   应付股利/应付利息 并入其他应付款、应收股利/应收利息 并入其他应收款、坏账准备 并入应收，
#   它们没有独立底稿文件但已随主科目底稿覆盖。uncovered_subjects 原只查注册表 →
#   RELATED_MAP 并入科目被误报"未覆盖"（DQ 2232 应付股利误报 1 亿）。
RELATED_COVERED_CODES = ('2232', '2231', '1131', '1132', '1231')
RELATED_COVERED_KW = ('应付股利', '应付利息', '应收股利', '应收利息', '坏账准备')


def _norm(c):
    return re.sub(r'\D', '', str(c))


def _is_covered(code, name):
    """判断该一级科目是否已被注册表 codes 覆盖（前缀命中）。"""
    cc = _norm(code)
    # ⚡ 2026-08-12：并入主科目底稿的科目（RELATED_MAP）视为已覆盖
    for rc in RELATED_COVERED_CODES:
        if cc == rc or cc.startswith(rc):
            return True
    for kw in RELATED_COVERED_KW:
        if kw in name:
            return True
    # 精确/前缀命中
    for rc in REG_CODES:
        if cc == rc or cc.startswith(rc):
            return True
    # 名称关键词命中
    for nm in REG_NAMES:
        if nm and (nm in name or name in nm):
            return True
    return False


def uncovered_subjects(data_dir, years=None):
    """返回 [(code, name, entity, year, qc, jf, df, qm)] 未覆盖科目行。"""
    if not os.path.isdir(data_dir):
        return []
    entities = A.discover_entities(data_dir)
    tb_full = A.read_tb_full(data_dir, entities)
    if years is None:
        years = sorted({str(yy) for (_e, _c, _n, yy) in tb_full}) or ['2025']
    out = []
    seen = set()
    for (e, c, n, yy), v in tb_full.items():
        if str(yy) not in years:
            continue
        cc = _norm(c)
        if not cc or len(cc) < 4:
            continue
        # 只取一级（4 位码或是最短前缀根）——U8 父级=子和，父级行才是科目级
        # 判定：该 code 不是任何其他 code 的子级（即不存在更短前缀的同族根）
        is_root = True
        # 简化：仅当 4 位码（len==4）或该码是族内最短
        if len(cc) > 4:
            is_root = False
            # 检查是否有 4 位前缀行存在（父级）
            for (e2, c2, n2, yy2), v2 in tb_full.items():
                if str(yy2) != str(yy) or e2 != e:
                    continue
                if _norm(c2) and _norm(c2) == cc[:4]:
                    is_root = False
                    break
        if not is_root:
            continue
        qc = float(v.get('qc') or 0.0)
        jf = float(v.get('jf') or 0.0)
        df = float(v.get('df') or 0.0)
        qm = float(v.get('qm') or 0.0)
        if abs(qc) < 0.005 and abs(jf) < 0.005 and abs(df) < 0.005 and abs(qm) < 0.005:
            continue
        if cc in SKIP_PREFIX:
            continue
        if any(k in str(n) for k in SKIP_NAME_KW):
            continue
        if _is_covered(cc, str(n)):
            continue
        key = (cc, str(n))
        if key in seen:
            continue
        seen.add(key)
        out.append((cc, str(n), e, str(yy), qc, jf, df, qm))
    out.sort(key=lambda x: (x[3], x[0]))
    return out


def write_uncovered_xlsx(data_dir, rows):
    """写入 账套\未覆盖科目清单.xlsx（根目录，便于用户直接看到）。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    ws = wb.active
    ws.title = '未覆盖科目清单'
    hdr = ['科目代码', '科目名称', '核算主体', '年度', '期初余额', '借方发生额',
           '贷方发生额', '期末余额', '建议']
    HDR_FILL = PatternFill('solid', fgColor='DDEBF7')
    HDR_FONT = Font(name='Times New Roman', bold=True, color='000000', size=10)
    THIN = Side(style='thin')
    BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for j, h in enumerate(hdr, 1):
        c = ws.cell(1, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal='center')
        c.border = BORDER
    r = 2
    for (cc, nm, e, yy, qc, jf, df, qm) in rows:
        # 建议归属：按科目代码首字符粗判
        if cc.startswith(('1',)):
            sug = '资产类（若属往来/长期资产请并入对应 builder）'
        elif cc.startswith(('2',)):
            sug = '负债类（若属借款/往来请并入对应 builder）'
        elif cc.startswith(('4',)):
            sug = '权益类（equity_detail 已覆盖主要科目）'
        elif cc.startswith(('5', '6')):
            sug = '损益类（pl_detail 已覆盖主要科目）'
        else:
            sug = ''
        vals = [cc, nm, e, yy, qc, jf, df, qm, sug]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.border = BORDER
            if j >= 5:
                c.number_format = '#,##0.00'
        r += 1
    for col, w in [('A', 12), ('B', 30), ('C', 14), ('D', 8),
                   ('E', 16), ('F', 16), ('G', 16), ('H', 16), ('I', 44)]:
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    out = os.path.join(data_dir, '未覆盖科目清单.xlsx')
    try:
        wb.save(out)
    except PermissionError:
        out = os.path.join(data_dir, 'prepared') if os.path.isdir(os.path.join(data_dir, 'prepared')) else data_dir
        out = os.path.join(out, '未覆盖科目清单.xlsx')
        wb.save(out)
    return out


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print('用法：python uncovered_subjects.py <账套目录>')
        return
    data_dir = argv[0]
    rows = uncovered_subjects(data_dir)
    print(f'\n===== 未覆盖科目清单（{os.path.basename(data_dir.rstrip(chr(92)+chr(47)))}）=====')
    if not rows:
        print('  ✅ 全部有数据一级科目均已覆盖（无未覆盖科目）')
        return
    print(f'  共 {len(rows)} 个一级科目未被注册表/builder 覆盖：')
    for (cc, nm, e, yy, qc, jf, df, qm) in rows:
        print(f'    {cc} | {nm[:22]:22s} | {e:8s} | {yy} | 期末={qm:,.2f} | 借发={jf:,.2f}')
    out = write_uncovered_xlsx(data_dir, rows)
    print(f'  ✓ 清单已写入：{out}')


if __name__ == '__main__':
    main()
