# -*- coding: utf-8 -*-
"""currency_annotate.py —— 底稿币种标注（2026-08-05 分币种生成配套）。

对【外币目录】的已生成底稿，在首 sheet 标题单元格（A1）追加币种标注，
如『应收账款2025审定表（泰铢 THB）』。幂等：已含币种关键词则不重复追加。
人民币目录（全部主体 CNY）不标注（默认人民币，无需标注）。

设计说明：
- 目录即币种（currency_split.py 按币种拆分），但打开单个底稿文件时目录不可见，
  故在文件内标题行显式标注，防止泰铢/美元金额被误当人民币。
- 只标注首 sheet 的标题行，不触碰数据区/公式，不影响 audit_checker 与 tb_recon 取数。
- 混币种目录（拆分前）不标注（应先用 currency_split 拆分）。

用法：
  python currency_annotate.py <文件夹>   # 对文件夹内全部 *_生成.xlsx 标注
"""
import paths as P
import os
import sys
import io

# 注意：不在模块级包装 sys.stdout（被 audit_finalize import 时再次包装会关闭底层 buffer），
# 仅在作为主程序运行时（__main__）包装。

CURRENCY_CN = {'CNY': '人民币', 'THB': '泰铢', 'USD': '美元'}
CURRENCY_KW = ('人民币', '泰铢', '美元', 'CNY', 'THB', 'USD', 'RMB')


def _annotate_file(fp, currency):
    """对单个底稿文件：首 sheet A1 标题行追加『（币种：XX）』。幂等。返回是否修改。"""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(fp)
    except Exception:
        return False
    modified = False
    ws = wb[wb.sheetnames[0]]
    a1 = ws.cell(1, 1).value
    if a1 is None or not str(a1).strip():
        a1 = ws.cell(2, 1).value
        row_i = 2
    else:
        row_i = 1
    if a1 is None:
        wb.close()
        return False
    s = str(a1)
    if any(k in s for k in CURRENCY_KW):
        wb.close()
        return False  # 已标注过（幂等）
    label = CURRENCY_CN.get(currency, currency)
    ws.cell(row_i, 1, '%s（%s %s）' % (s, label, currency))
    try:
        wb.save(fp)
        modified = True
    except Exception:
        modified = False
    wb.close()
    return modified


def process_folder(folder):
    """标注文件夹内全部 *_生成.xlsx。返回标注文件数。非外币目录返回 0。"""
    try:
        from currency_mapping import entity_currency
        from audit_common import discover_entities
    except Exception:
        return 0
    ents = discover_entities(folder)
    if not ents:
        return 0
    curs = {entity_currency(e) for e in ents}
    if curs <= {'CNY'}:
        return 0  # 全人民币：默认币种，不标注
    if len(curs) > 1:
        print('  ⚠️ [币种标注] 目录含多币种主体 %s——请先用 currency_split.py 拆分' % sorted(curs))
        return 0
    currency = curs.pop()
    n = 0
    for fn in sorted(os.listdir(folder)):
        if not (fn.endswith('_生成.xlsx') and not fn.startswith('~$') and '_bak' not in fn):
            continue
        if _annotate_file(os.path.join(folder, fn), currency):
            n += 1
    if n:
        print('  [币种标注] %s：%d 份底稿标注（%s %s）' % (
            os.path.basename(folder), n, CURRENCY_CN.get(currency, currency), currency))
    return n


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    d = sys.argv[1] if len(sys.argv) > 1 else os.path.join(P.Z, 'Z_split', 'THB')
    n = process_folder(d)
    print('标注文件数：%d' % n)
