# -*- coding: utf-8 -*-
"""merge_user_tb.py —— 2026-08-06 新增：把账套内【分主体科目余额表】合并成一张
标准格式合并试算表（科目编码|科目名称|方向|期初金额|本期借方|本期贷方|方向|期末金额|等级），
供 tb_recon_detail 做「用户试算表 vs 程序底稿」核对（FY 有用户手工合并 TB，其他账套
只有分主体版 → 需要程序合并）。

用法：python merge_user_tb.py <数据文件夹> [输出路径]
依赖：audit_common.read_tb_full（已兼容各账套表头：科目编码/名称/方向/金额/等级）。
"""
import paths as P
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import audit_common as A
import subject_mapping  # 2026-08-07 任务683：科目名归一化接入（H 递延资产→长期待摊费用等）

F = Font(name='Times New Roman', size=10)
THIN = Side(style='thin', color='000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CTR = Alignment(horizontal='center', vertical='center')
RGT = Alignment(horizontal='right', vertical='center')
NUM = '#,##0.00'


def _fx(v):
    """自然借贷符号金额 → (方向, 金额绝对值)。"""
    v = float(v or 0.0)
    return ('借', abs(v)) if v >= 0 else ('贷', abs(v))


def merge_tb(data_dir, out_path=None):
    entities = A.discover_entities(data_dir)
    if not entities:
        print(f'⚠️ {data_dir} 未发现主体')
        return None
    tb = A.read_tb_full(data_dir, entities)
    if not tb:
        print(f'⚠️ {data_dir} TB 解析为空')
        return None
    # 2026-08-06 修复③：只聚合【目标年度】（首个出现年度，与 self_tb_gen 一致）——
    # c 等账套有 2025+2026 两套 TB 文件，聚合键含年度会拆成两套同码行（53→98 科目）。
    _years_all = sorted({str(y) for (e, c, n, y) in tb})
    _tg = _years_all[0] if _years_all else '2025'
    # 2026-08-06 改进：按【一级科目（4位码）】聚合【末级叶子】——对齐底稿一级口径，
    # 避免二级/三级行（如 dq『宁波银行嘉兴分行』）被核对器误报"试算表有数-底稿未生成"。
    # 叶子判定：同 (e, y) 内不存在其他代码以其为前缀。一级码=去点前 4 位。
    # 2026-08-06 修复②：字母后缀代码（J 2241A4=2241 其他应付款子科目）→ 取数字前缀 2241；
    # 非数字开头（R『核算维度编码』表头行/JTt『总计』汇总行）→ 排除（非科目行）。
    import re as _re

    def _l1(cs):
        d = str(cs).replace('.', '')
        m = _re.match(r'^\d{1,4}', d)
        if m:
            return m.group(0)
        return None   # 非数字代码 → 过滤

    _SKIP_NAMES = ('核算账簿', '核算维度', '总计', '合计', '小计')

    def _ok_name(nn):
        if not nn:
            return True
        return not any(k in str(nn) for k in _SKIP_NAMES)

    # 先收集每 (e, y) 的代码集合判叶子
    codes_of = {}
    l1_names = {}   # (l1, y) -> 一级科目名（取父级行名称；无父级行→该码下最常见名）
    name_cnt = {}
    for (e, c, n, y) in tb:
        if str(y) != _tg:
            continue
        cs = str(c)
        codes_of.setdefault((e, str(y)), set()).add(cs)
        _d = cs.replace('.', '')
        _l1v = _re.match(r'^\d{1,4}', _d)
        _l1v = _l1v.group(0) if _l1v else None
        if _l1v is None:
            continue
        k = (_l1v, str(y))
        name_cnt[k] = name_cnt.get(k, 0) + 1
        if len(_d) == 4 and _d.isdigit():   # 4 位纯数字 = 一级父级行
            if _ok_name(n):
                l1_names.setdefault(k, str(n or ''))
    for k, _cnt in name_cnt.items():
        l1_names.setdefault(k, '')   # 无父级行则留空（下方兜底取最长名）
    for k, _cnt in name_cnt.items():
        l1_names.setdefault(k, '')   # 无父级行则留空（下方兜底取最长名）
    _l1_seg_cnt = {}
    for (e, c, n, y) in tb:
        if str(y) != _tg:
            continue
        cs = str(c)
        _d = cs.replace('.', '')
        _m = _re.match(r'^\d{1,4}', _d)
        if _m is None:
            continue
        _l1v = _m.group(0)
        k = (_l1v, str(y))
        if not l1_names.get(k):
            # 2026-08-06：无父级行（JTt 末级导出）→ 一级名取【路径第二段】
            # （'640303\税金及附加\城市维护建设税' → '税金及附加'）；无路径则取末段名。
            _parts = str(n or '').split('\\')
            _seg = _parts[1].strip() if len(_parts) >= 2 else _parts[-1].strip()
            if _ok_name(_seg):
                _cnt = _l1_seg_cnt.get(k, {})
                _cnt[_seg] = _cnt.get(_seg, 0) + 1
                _l1_seg_cnt[k] = _cnt
    for k, _cnt in _l1_seg_cnt.items():
        # 2026-08-06 修复④：直接赋值覆盖预置空值（setdefault 对已存在键（''）不覆盖
        # → JTt 无父级行账套 l1_names 恒空 → 每叶子末段名独立成行 → 558 科目暴增）
        l1_names[k] = max(_cnt, key=_cnt.get)
    agg = {}
    for (e, c, n, y), v in tb.items():
        if str(y) != _tg:
            continue
        cs = str(c)
        fam = codes_of.get((e, str(y)), set())
        is_leaf = not any(o != cs and o.startswith(cs) for o in fam)
        if not is_leaf:
            continue   # 父级行跳过（金额=子级之和，避免双计）
        l1 = _l1(cs)
        if l1 is None or not _ok_name(str(n or '')):
            continue   # 非数字代码 / 汇总行（JTt 总计、R 核算维度）过滤
        k = (l1, str(y))
        name = l1_names.get(k) or str(n or '').split('\\')[-1].split('-')[0].strip()
        # 2026-08-07 任务683：账套科目名归一化（account_profiles.json common/专属 name_map
        # ——H『递延资产』→『长期待摊费用』、通用『股本』→『实收资本(或股本)』等），
        # 合并试算表输出统一标准名，后续核对/底稿匹配更顺。
        name = subject_mapping.norm_name(os.path.basename(data_dir.rstrip('\\/')), name)
        key = (l1, name, str(y))
        d = agg.setdefault(key, {'qc': 0.0, 'qm': 0.0, 'jf': 0.0, 'df': 0.0, 'name': name})
        d['qc'] += float(v.get('qc') or 0.0)
        d['qm'] += float(v.get('qm') or 0.0)
        d['jf'] += float(v.get('jf') or 0.0)
        d['df'] += float(v.get('df') or 0.0)
    if not agg:
        print(f'⚠️ {data_dir} 聚合为空')
        return None
    # 按代码排序（数字序）
    def _ck(k):
        cc = k[0]
        try:
            return (0, int(cc.replace('.', '')))
        except ValueError:
            return (1, cc)
    items = sorted(agg.items(), key=lambda kv: _ck(kv[0]))
    out_path = out_path or os.path.join(data_dir, f'合并试算表_{_tg}.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'TB表'
    ws.append(['科目编码', '科目名称', '方向', '期初金额', '本期借方', '本期贷方', '方向', '期末金额', '等级'])
    for (l1, nm, yy), d in items:
        # 2026-08-06：方向列留空（None）——parse_trial_balance signed=True 只在方向列
        # 含『贷』时翻负；FY 用户合并TB 方向列即空 → 期末列=正数报表口径，与底稿审定
        # （负债/权益类正数）天然对齐。若输出『贷』方向 → 负债类翻负 → 核对差 2 倍。
        ws.append([l1, d['name'] or nm, None, round(abs(d['qc']), 2), round(d['jf'], 2),
                   round(d['df'], 2), None, round(abs(d['qm']), 2), 1])
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = F
            cell.border = BORDER
            if cell.column in (2, 3, 7):
                cell.alignment = CTR
            elif isinstance(cell.value, (int, float)):
                cell.number_format = NUM
                cell.alignment = RGT
    for col, w in zip('ABCDEFGHI', (12, 30, 6, 14, 14, 14, 6, 14, 6)):
        ws.column_dimensions[col].width = w
    try:
        wb.save(out_path)
    except BaseException as ex:
        print(f'⚠️ 保存失败（可能被 Excel 占用）：{ex}')
        return None
    print(f'✓ 合并试算表已生成：{out_path}（{len(items)} 个科目 × {len(entities)} 主体）')
    return out_path


if __name__ == '__main__':
    d = sys.argv[1] if len(sys.argv) > 1 else P.C
    o = sys.argv[2] if len(sys.argv) > 2 else None
    merge_tb(d, o)
