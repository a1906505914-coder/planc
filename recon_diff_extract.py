# -*- coding: utf-8 -*-
"""recon_diff_extract.py —— 统一差异抽取器（2026-08-24 架构 R1）。

职责：从专项核对底稿 xlsx 中，按【差异标记】抽取差异行 → 统一差异 JSON 契约
（见 docs/audit_pipeline_design.md §6.1）。

设计：
    - 只【读】专项产出的 xlsx，不改动任何内容（安全：不侵入各专项子模块）
    - 每个专项在 recon_registry.REGISTRY 配置 diff_sheets：
        文件 glob + sheet 名 + 差异标记关键词列
    - 匹配到的行 → diff JSON；未匹配到的差异行被记录为 evidence 引用

差异标记约定（各专项通用）：
    '⚠️' / '✗' / '差异' / '不符' / '未发现' / '独有' / '无对应' / '待查'
"""
import glob
import json
import os
import re

# 通用差异标记（按优先级排列，命中即视为差异行）
_DIFF_MARKS = ('⚠️', '✗', '不符', '未发现', '独有', '无对应', '待查', '差异')


def find_diff_files(data_dir, patterns):
    """按 glob 模式在 data_dir 下找专项产出 xlsx（优先非脱敏版）。"""
    files = []
    for pat in patterns:
        p = os.path.join(data_dir, pat)
        files += [f for f in glob.glob(p) if not f.endswith('_脱敏.xlsx')]
    # 去重保序
    seen = set()
    out = []
    for f in files:
        if f not in seen:
            seen.add(f)
            out.append(f)
    return out


def _is_diff_cell(v):
    s = str(v or '')
    return any(m in s for m in _DIFF_MARKS)


def extract_workbook(fp, sheet_name, mark_cols, extra_cols=None, num_mark=False,
                     skip_rows=0, mark_text_cols=None):
    """从单 xlsx 指定 sheet 抽取差异行。
    参数：
        fp              xlsx 路径
        sheet_name      目标 sheet 名（None=全部 sheet）
        mark_cols       差异标记所在列索引（0-based；任一列含标记即差异行）
        extra_cols      额外要带出的列索引（0-based；默认全部非空列）
        num_mark        True 时差异判定为【数值非 0】（mark_cols 取数值列）
        skip_rows       表头前跳过行数
        mark_text_cols  附加文本标记列（如函证『不符』结论；与数值列 OR 判定）
    返回：差异行列表 [{col0: val, ...}]（含来源文件标记）。
    """
    from openpyxl import load_workbook
    diffs = []
    wb = load_workbook(fp, read_only=True, data_only=True)
    sheets = [sheet_name] if sheet_name else wb.sheetnames
    for sn in sheets:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        for ri, r in enumerate(ws.iter_rows(values_only=True)):
            if not r:
                continue
            if ri < skip_rows:
                continue
            # 差异判定
            if num_mark:
                hit = False
                for c in mark_cols:
                    if c < len(r):
                        v = r[c]
                        # 兼容 openpyxl 数值以 str 存储（'50'）与真实 float
                        try:
                            nv = float(v) if v not in (None, '') else 0.0
                        except (TypeError, ValueError):
                            nv = 0.0
                        if abs(nv) > 0.0001:
                            hit = True
                            break
                # 附加文本标记列（如函证『不符』结论）：任一含标记即差异
                if not hit:
                    for c in (mark_text_cols or []):
                        if c < len(r) and _is_diff_cell(r[c]):
                            hit = True
                            break
            else:
                hit = any(_is_diff_cell(r[c]) for c in mark_cols if c < len(r))
            if not hit:
                continue
            rec = {'_src': os.path.basename(fp), '_sheet': sn}
            for i, v in enumerate(r):
                if v is None:
                    continue
                rec[str(i)] = str(v).strip()
            diffs.append(rec)
    wb.close()
    return diffs


def extract_all(data_dir, cfg):
    """按专项配置抽取全部差异 → 统一差异 JSON 结构。
    cfg 字段（recon_registry.REGISTRY[key]）：
        diff_sheets: [{file: '底稿/**/银行开户账户核对_*.xlsx', sheet: '逐账户明细',
                       marks: [7], cols: [0,1,2,6]}, ...]
    返回统一差异 dict：{'recon','diffs':[{entity?,gl_amt?,ext_amt?,diff?,...}]}。
    """
    recon_name = cfg.get('name', '')
    diffs = []
    for ds in cfg.get('diff_sheets', []):
        files = find_diff_files(data_dir, [ds['file']])
        for fp in files:
            rows = extract_workbook(fp, ds.get('sheet'), ds.get('marks', [0]),
                                    ds.get('cols'), num_mark=ds.get('num_mark', False),
                                    skip_rows=ds.get('skip_rows', 0),
                                    mark_text_cols=ds.get('mark_text_cols'))
            for r in rows:
                diffs.append(_to_diff_contract(r, cfg, ds))
    return {'recon': recon_name, 'diffs': diffs}


def _to_diff_contract(row, cfg, ds):
    """xlsx 行 → 统一差异契约字段（尽量映射 entity/code/gl_amt/ext_amt/diff；
    无金额信息的专项只留 evidence 文本，金额留人工确认）。"""
    # 配置的列语义映射（可选）：{契约字段: 列索引}
    colmap = ds.get('colmap', {})
    rec = {
        'reason': '', 'evidence': '', 'status': 'open',
        '_src': row.get('_src', ''), '_sheet': row.get('_sheet', ''),
    }
    for field, ci in colmap.items():
        if ci in row:
            rec[field] = row[ci]
    # 把整行文本拼成 evidence（便于人工核对；差异标记前的文本即差异原因）
    vals = [row[k] for k in sorted(row) if not k.startswith('_') and row[k]]
    if not rec.get('evidence'):
        rec['evidence'] = ' | '.join(vals)
    return rec


def write_diffs(data_dir, cfg, diff_out=None):
    """抽取并写差异 JSON（data_dir 下）。返回 JSON 路径。"""
    d = extract_all(data_dir, cfg)
    fp = diff_out or os.path.join(data_dir, cfg.get('diff_json', '差异.json'))
    os.makedirs(os.path.dirname(fp) if os.path.dirname(fp) else '.', exist_ok=True)
    with open(fp, 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    return fp
