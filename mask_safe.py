# -*- coding: utf-8 -*-
"""mask_safe.py —— 读取层「脱敏版优先」统一 helper（2026-08-22）。

铁律：进对话/分析的数据必须是脱敏版。本模块提供统一的"脱敏版优先"取路径/
读取入口，新台账/专项接入时使用，避免各程序各自实现、靠自觉。

用法：
  from mask_safe import prefer_masked, require_masked
  fp = prefer_masked(r'D:/底稿测试/ADF/数据/2026/融资明细.xls')
        # → 存在 融资明细_脱敏.xlsx 则返回脱敏版，否则原路径
  fp = require_masked(path)          # strict：无脱敏版 → 返回 None 并告警（拒绝分析）
  wb = read_masked_workbook(path)    # 按脱敏版优先读 openpyxl workbook（已关闭原对象）
"""
import os


def _paired_mask(root, fname):
    """找配对脱敏版（返回绝对路径或 None），支持三种布局：
    ① 同目录 {stem}_脱敏.{ext}
    ② 同级/父级 {dirname}_脱敏 或 {dirname}_脱敏版 目录 + 相对子路径
       （合同/回函等目录级脱敏：识别文本_脱敏/、回函_脱敏版/）"""
    base_nm, _ = os.path.splitext(fname)
    for e in ('.xlsx', '.xls'):
        p = os.path.join(root, base_nm + '_脱敏' + e)
        if os.path.exists(p):
            return p
    d = root
    while True:
        parent = os.path.dirname(d)
        if parent == d:
            break
        dname = os.path.basename(d)
        rel_sub = os.path.relpath(root, d)
        for suf in ('_脱敏', '_脱敏版'):
            mask_root = os.path.join(parent, dname + suf)
            if os.path.isdir(mask_root):
                cand = os.path.join(mask_root, rel_sub, fname) if rel_sub != '.' \
                    else os.path.join(mask_root, fname)
                if os.path.exists(cand):
                    return cand
        d = parent
    return None


def prefer_masked(path):
    """原始路径 → 优先返回【脱敏版】路径；无脱敏版回退原路径（本机审计用）。"""
    if not path:
        return path
    root = os.path.dirname(path)
    fname = os.path.basename(path)
    p = _paired_mask(root, fname)
    return p if p else path


def require_masked(path):
    """strict：无脱敏版 → 返回 None 并告警（调用方应拒绝分析原始文件）。"""
    if not path:
        return None
    root = os.path.dirname(path)
    fname = os.path.basename(path)
    p = _paired_mask(root, fname)
    if p:
        return p
    print(f'  ⚠️ [脱敏闸门] {fname} 无脱敏版（{path}）——原始含主体名，禁止进对话/分析')
    return None


def read_masked_workbook(path, **kw):
    """按脱敏版优先读 openpyxl workbook（data_only=True）。无脱敏版时仍可读（本机审计），
    但调用方须自行保证不进对话。"""
    import openpyxl
    fp = prefer_masked(path)
    kw.setdefault('data_only', True)
    return openpyxl.load_workbook(fp, **kw), fp
