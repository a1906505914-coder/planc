# -*- coding: utf-8 -*-
"""_safx.py —— 流式 xlsx 解析器（SAP 系公共依赖，2026-08-19 重建）。

⚠️ 2026-08-19 整理文件夹时被误删，按调用方契约重建：
  · iter_rows(fp, keep_cols=None, min_row=None, sheet=None) → 逐行 yield {列索引: 值}
  · _all_sheet_paths(z) → zip 内全部 worksheet XML 路径（按 sheet 序号排序）
  · 基于 zipfile + ElementTree.iterparse 流式解析（不整表加载，SAP 大文件不 OOM）；
    sharedStrings 用 iterparse 流式构建（数十万条也不爆内存）。

调用方：sap_common / sap_reader / sap_adapter / sap_gl_diff_report /
         sap_ar_exception_report / sap_current_account_detail / tax_detail
"""
import re
import zipfile
import xml.etree.ElementTree as ET

_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
_RE_SHEET = re.compile(r'xl/worksheets/sheet(\d+)\.xml$')
_RE_COL = re.compile(r'^([A-Z]+)')


def _all_sheet_paths(z):
    """返回 zip 内全部 worksheet XML 路径（按 sheet 序号升序）。"""
    paths = [n for n in z.namelist() if _RE_SHEET.match(n)]
    paths.sort(key=lambda p: int(_RE_SHEET.search(p).group(1)))
    return paths


def _col_idx(ref):
    """'AB12' → 27（0-based 列索引）。"""
    m = _RE_COL.match(ref or '')
    if not m:
        return 0
    idx = 0
    for ch in m.group(1):
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _load_sst(z):
    """流式构建共享字符串表（iterparse，不整文件入内存）。"""
    sst = []
    with z.open('xl/sharedStrings.xml') as f:
        for event, el in ET.iterparse(f, events=('end',)):
            if el.tag == _NS + 'si':
                sst.append(''.join(el.itertext()))
                el.clear()
    return sst


def _cell_value(c, sst):
    """按类型取单元格值（数字保留为字符串，与既有解析器一致）。"""
    t = c.get('t')
    if t == 's':
        v = c.find(_NS + 'v')
        if v is not None and v.text is not None:
            try:
                return sst[int(v.text)]
            except (IndexError, ValueError):
                return ''
        return ''
    if t == 'inlineStr':
        return ''.join(c.itertext())
    if t == 'str':
        v = c.find(_NS + 'v')
        return v.text if v is not None and v.text is not None else ''
    if t == 'b':
        v = c.find(_NS + 'v')
        return (v.text == '1') if v is not None and v.text is not None else None
    # 默认数字：保留字符串（SAP 读取器以 str 处理，金额列可 float()）
    v = c.find(_NS + 'v')
    return v.text if v is not None and v.text is not None else ''


def iter_rows(fp, keep_cols=None, min_row=None, sheet=None):
    """逐行 yield {列索引: 值}。

    keep_cols: 限制读取列（None=全部）；min_row: 跳过前 N 行（0-based）；
    sheet: 指定 worksheet XML 路径（默认第一个 sheet）。
    """
    z = zipfile.ZipFile(fp)
    try:
        paths = _all_sheet_paths(z)
        if not paths:
            return
        sp = sheet or paths[0]
        sst = None
        if 'xl/sharedStrings.xml' in z.namelist():
            sst = _load_sst(z)
        skip = min_row or 0
        row_no = -1
        with z.open(sp) as f:
            for event, el in ET.iterparse(f, events=('end',)):
                if el.tag != _NS + 'row':
                    continue
                row_no += 1
                if row_no < skip:
                    el.clear()
                    continue
                cells = {}
                for c in el:
                    col = _col_idx(c.get('r', ''))
                    if keep_cols is not None and col not in keep_cols:
                        continue
                    cells[col] = _cell_value(c, sst)
                yield cells
                el.clear()
    finally:
        z.close()


if __name__ == '__main__':
    import sys
    for fp in sys.argv[1:]:
        z = zipfile.ZipFile(fp)
        print(fp, '→ sheets:', _all_sheet_paths(z))
        z.close()
        for i, r in enumerate(iter_rows(fp, keep_cols={0, 1, 2}, min_row=1)):
            print('  ', r)
            if i >= 3:
                break
