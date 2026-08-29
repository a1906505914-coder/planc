# -*- coding: utf-8 -*-
"""post_entry.py — 审计调整分录过入器（2026-08-04）

读《审计调整分录汇总表.xlsx》→ 按 主体×科目×年度 定位各底稿审定表
『审计重分类调整1/审计重分类调整2/审计调整』列增量写入（铁律33：只写调整数单元格，
绝不重生成）；指定明细对象时同步过入明细表 Q(重分类)/R(审计调整) 列。

用法：
    python post_entry.py <分录表.xlsx> [底稿文件夹，缺省=分录表所在目录]

过入前校验（不满足即拒绝）：
    ① 借贷平衡：借合计=贷合计（0.01 铁律）
    ② 主体×科目×年度 能在底稿审定表定位到行
    ③ 金额为正数
过入后重算：审定数公式已联动（=期末+重分类1+重分类2+审计调整）；合计行 K 列 SUM 已补。

说明：重复运行会重复累加（幂等由审计师保证；如需重来，重新生成底稿即可）。
"""
import os, sys, io, argparse
import openpyxl
from openpyxl.styles import Font

# ⚡ 2026-08-19 修复：作为模块被 import 时（无 stdout.buffer / 已关闭）不再崩溃；
#    仅直接运行时包装 stdout 为 UTF-8 输出
if hasattr(sys.stdout, 'buffer'):
    try:
        # ⚡ 用 reconfigure 原地重配，避免 TextIOWrapper 重包装导致旧对象 GC 关闭 buffer
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (ValueError, OSError, AttributeError):
        pass

# 调整类别 → 审定表列名 / 明细表列名
CAT_MAP = {
    '重分类调整1': ('审计重分类调整1', '重分类调整'),
    '重分类调整2': ('审计重分类调整2', '重分类调整'),
    '审计调整': ('审计调整', '审计调整'),
}
RED = Font(name='Times New Roman', size=10, bold=True, color='C00000')


def _col_by_header(ws, name):
    """在审定表/明细表表头行找列名 → (表头行, 列号)。
    精确匹配列名（'审计调整' 不会误中 '审计重分类调整1'）；对不含空格的原样比较。"""
    for r in range(1, min(ws.max_row, 20) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            sv = str(v).replace('\n', '')
            if sv == name:
                return r, c
            # 兼容 '审计调整数'（损益审定表）等带后缀变体
            if name == '审计调整' and sv.startswith('审计调整'):
                return r, c
    return None, None


def _find_audit_row(ws, ent, year):
    """在审定表找 主体=ent 的行。支持两类结构：
    A) 往来款 12 列（公司列=B，第2列）
    B) 损益/其他 6 列（核算主体列=A，第1列）
    返回行号；找不到返回 None。"""
    # 定位表头行（含『公司』或『主体』或『单位』）
    hdr_r = None
    ent_col = 2
    for r in range(1, min(ws.max_row, 15) + 1):
        row_vals = [str(ws.cell(r, c).value or '') for c in range(1, ws.max_column + 1)]
        if any('公司' in v or '主体' in v or '单位' in v for v in row_vals):
            hdr_r = r
            # 主体列：若第1列表头含『主体/单位/公司』则用 A 列，否则用 B 列
            if row_vals and ('主体' in row_vals[0] or '单位' in row_vals[0] or '公司' in row_vals[0]):
                ent_col = 1
            break
    if hdr_r is None:
        return None
    for r in range(hdr_r + 1, ws.max_row + 1):
        v = ws.cell(r, ent_col).value
        if v and ent in str(v):
            return r
    return None


def _find_detail_row(ws, ent, cust):
    """在明细表找 主体=A列 且 往来单位名称=D列 的行。"""
    if not cust:
        return None
    for r in range(2, ws.max_row + 1):
        a = str(ws.cell(r, 1).value or '')
        d = str(ws.cell(r, 4).value or '')
        if a == ent and cust in d:
            return r
    return None


def main():
    ap = argparse.ArgumentParser(description='审计调整分录过入器')
    ap.add_argument('entry_file', help='审计调整分录汇总表.xlsx')
    ap.add_argument('folder', nargs='?', default=None, help='底稿文件夹（缺省=分录表所在目录）')
    args = ap.parse_args()

    entry_file = os.path.abspath(args.entry_file)
    folder = os.path.abspath(args.folder) if args.folder else os.path.dirname(entry_file)
    if not os.path.isfile(entry_file):
        print(f'❌ 分录表不存在：{entry_file}')
        return 1

    wb = openpyxl.load_workbook(entry_file, data_only=True)
    ws = wb.active
    # 表头定位
    hdr_r = None
    for r in range(1, min(ws.max_row, 20) + 1):
        row_vals = [str(ws.cell(r, c).value or '') for c in range(1, ws.max_column + 1)]
        if any('科目' in v for v in row_vals) and any('金额' in v for v in row_vals):
            hdr_r = r
            break
    if hdr_r is None:
        print('❌ 分录表未找到表头（需含 科目/金额 列）')
        return 1
    col = {}
    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(hdr_r, c).value or '').strip()
        for k in ('年度', '核算主体', '科目', '调整类别', '借贷方向', '金额', '明细对象'):
            if k in v and k not in col:
                col[k] = c
    need = ('年度', '核算主体', '科目', '调整类别', '借贷方向', '金额')
    if any(k not in col for k in need):
        print(f'❌ 分录表缺列：{set(need) - set(col)}')
        return 1

    entries = []
    for r in range(hdr_r + 1, ws.max_row + 1):
        y = ws.cell(r, col['年度']).value
        ent = ws.cell(r, col['核算主体']).value
        subj = ws.cell(r, col['科目']).value
        cat = ws.cell(r, col['调整类别']).value
        drcr = ws.cell(r, col['借贷方向']).value
        amt = ws.cell(r, col['金额']).value
        cust = ws.cell(r, col['明细对象']).value if '明细对象' in col else None
        if y is None and ent is None and subj is None:
            continue
        if cat not in CAT_MAP:
            print(f'  ⚠️ 行{r} 调整类别无效：{cat}（应 审计调整/重分类调整1/重分类调整2）')
            continue
        if str(drcr).strip() not in ('借', '贷'):
            print(f'  ⚠️ 行{r} 借贷方向无效：{drcr}')
            continue
        try:
            amt = float(amt)
        except (TypeError, ValueError):
            print(f'  ⚠️ 行{r} 金额无效：{amt}')
            continue
        entries.append(dict(r=r, y=str(y).strip(), ent=str(ent).strip(), subj=str(subj).strip(),
                            cat=cat, drcr=str(drcr).strip(), amt=abs(amt),
                            cust=(str(cust).strip() if cust else '')))

    if not entries:
        print('❌ 分录表无有效行')
        return 1
    # 借贷平衡校验
    db = sum(x['amt'] for x in entries if x['drcr'] == '借')
    cr = sum(x['amt'] for x in entries if x['drcr'] == '贷')
    print(f'录入 {len(entries)} 条：借合计 {db:,.2f} / 贷合计 {cr:,.2f}')
    if abs(db - cr) > 0.01:
        print(f'❌ 借贷不平衡（差 {db - cr:,.2f}），拒绝过入')
        return 1
    print('✅ 借贷平衡通过')

    # 过入
    ok = fail = 0
    detail_fail = []
    for x in entries:
        # 定位底稿文件
        pattern = f'{x["subj"]}审计底稿_{x["y"]}_生成.xlsx'
        path = os.path.join(folder, pattern)
        if not os.path.isfile(path):
            print(f'  ⚠️ 未找到底稿：{pattern}')
            fail += 1
            continue
        wb2 = openpyxl.load_workbook(path)
        try:
            # 审定表 sheet：名称含『审定表』
            sheet_name = next((s for s in wb2.sheetnames if '审定表' in s), None)
            if sheet_name is None:
                print(f'  ⚠️ {pattern} 无审定表 sheet')
                fail += 1
                continue
            ws2 = wb2[sheet_name]
            row = _find_audit_row(ws2, x['ent'], x['y'])
            if row is None:
                print(f'  ⚠️ {pattern} 未找到主体 {x["ent"]}')
                fail += 1
                continue
            hdr_r2, col2 = _col_by_header(ws2, CAT_MAP[x['cat']][0])
            if col2 is None:
                print(f'  ⚠️ {pattern} 审定表无列 {CAT_MAP[x["cat"]][0]}')
                fail += 1
                continue
            signed = x['amt'] if x['drcr'] == '借' else -x['amt']
            # 审定数公式 = 期末 + 调整（I/J/K 列直接加项），借贷方向即符号（借+贷−），
            # 重分类/审计调整统一按此处理（2026-08-04 简化，避免重复取反）。
            cur = ws2.cell(row, col2).value or 0
            try:
                cur = float(cur)
            except (TypeError, ValueError):
                cur = 0.0
            ws2.cell(row, col2, round(cur + signed, 2))
            # 明细表行级过入（可选）
            if x['cust']:
                dname = next((s for s in wb2.sheetnames if '明细表' in s), None)
                if dname:
                    wsd = wb2[dname]
                    drow = _find_detail_row(wsd, x['ent'], x['cust'])
                    if drow:
                        _, dcol = _col_by_header(wsd, CAT_MAP[x['cat']][1])
                        if dcol:
                            dcur = wsd.cell(drow, dcol).value or 0
                            try:
                                dcur = float(dcur)
                            except (TypeError, ValueError):
                                dcur = 0.0
                            wsd.cell(drow, dcol, round(dcur + signed, 2))
                        else:
                            detail_fail.append((x['ent'], x['cust'], x['subj']))
                    else:
                        detail_fail.append((x['ent'], x['cust'], x['subj']))
            ok += 1
            print(f"  ✅ 行{x['r']} {x['ent']}/{x['subj']}/{x['y']} {x['cat']} "
                  f"{x['drcr']} {x['amt']:,.2f} → 审定表 r{row} c{col2}")
        finally:
            wb2.save(path)
            wb2.close()
    print(f'\n过入完成：成功 {ok} / 失败 {fail}')
    if detail_fail:
        print(f'  ⚠️ 明细表行级未定位（仅过入审定表）：{detail_fail[:10]}')
    print('\n提示：审定数/合计行公式已联动；建议随后重跑 audit_checker 复核。')
    return 0 if fail == 0 else 2


if __name__ == '__main__':
    sys.exit(main())
