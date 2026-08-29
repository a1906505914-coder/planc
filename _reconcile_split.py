# -*- coding: utf-8 -*-
"""_reconcile_split.py —— 拆分整合模块（2026-08-22 重建）。

⚠️ 原版 2026-08-19 整理文件夹时被误删且无备份，本文件按 special_reviews.py
调用点语义重建：把专项复核表（银行核对/折旧复核/账龄复核）的 sheets 注入
对应科目审计底稿（以 {prefix}{sheet} 命名），原底稿先备份。整合失败不崩。

接口（special_reviews 使用）：
  bank_integrate(acct)                         → 注入成功的银行存款底稿数
  bank_group_summary(acct)                     → 集团银行底稿数（= bank_integrate）
  dep_integrate(acct, year, src_fp, subj_fp, prefix=None, label=None)
                                               → 注入成功的 sheet 数
"""
import glob
import os
import shutil

DATA_ROOT = os.environ.get('AUDIT_DATA_ROOT', r'D:\底稿测试')


def _inject_sheets(target_fp, src_fp, prefix=''):
    """把 src_fp 的全部 sheets 复制进 target_fp（{prefix}{sheet} 命名，替换已存在）。
    返回注入 sheet 数；任一步失败返回 0。"""
    import openpyxl
    if not os.path.exists(target_fp) or not os.path.exists(src_fp):
        print(f'  ⚠️ 整合跳过：目标/来源缺失 ({os.path.basename(target_fp)} / {os.path.basename(src_fp)})')
        return 0
    bak = target_fp.replace('.xlsx', '_bak_before_inject.xlsx')
    if not os.path.exists(bak):
        try:
            shutil.copy2(target_fp, bak)
        except Exception as _ex:
            print(f'  ⚠️ 备份失败 {target_fp}：{_ex}')
    try:
        src_wb = openpyxl.load_workbook(src_fp)
        dst_wb = openpyxl.load_workbook(target_fp)
    except Exception as ex:
        print(f'  ⚠️ 整合失败（读取）：{ex}')
        return 0
    added = 0
    for s in src_wb.sheetnames:
        name = f'{prefix}{s}'
        if name in dst_wb.sheetnames:
            del dst_wb[name]
        src_ws = src_wb[s]
        ws = dst_wb.create_sheet(title=name)
        for row in src_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    ws.cell(row=cell.row, column=cell.column, value=cell.value)
        added += 1
    try:
        dst_wb.save(target_fp)
    except Exception as ex:
        print(f'  ⚠️ 整合失败（保存）：{ex}')
        return 0
    print(f'  ✅ 注入 {added} 个 sheets（{prefix}…）→ {os.path.basename(target_fp)}')
    return added


def _acct_dir(acct):
    return os.path.join(DATA_ROOT, acct, '底稿')


def bank_integrate(acct):
    """银行核对/双向核对/开户核对产物 → 注入对应银行存款审计底稿。"""
    root = _acct_dir(acct)
    if not os.path.isdir(root):
        print(f'  ⚠️ {acct}: 无底稿目录')
        return 0
    # 银行核对产物（各年份、各主体目录）
    src_cands = []
    for pat in ('*银行*核对*.xlsx', '*双向核对*.xlsx', '*对账*核对*.xlsx', '*开户核对*.xlsx'):
        src_cands += glob.glob(os.path.join(root, '**', pat), recursive=True)
    # 排除已注入的（前缀含"核对-"的复制版不入源）
    src_cands = [f for f in src_cands if not any(k in os.path.basename(f) for k in ('_bak', '_old'))]
    total = 0
    for src_fp in src_cands:
        sdir = os.path.dirname(src_fp)
        # 目标：同目录（或同级子目录）的银行存款审计底稿
        targets = [os.path.join(sdir, f) for f in os.listdir(sdir)
                   if '银行存款' in f and '审计底稿' in f and f.endswith('_生成.xlsx')]
        if not targets and os.path.basename(sdir) != '':
            parent = os.path.dirname(sdir)
            targets = [os.path.join(parent, f) for f in os.listdir(parent)
                       if '银行存款' in f and '审计底稿' in f and f.endswith('_生成.xlsx')]
        for t in targets:
            total += _inject_sheets(t, src_fp, prefix='银行核对-')
    return total


def bank_group_summary(acct):
    """集团银行核对底稿数（统计口径，同 bank_integrate 注入数）。"""
    return bank_integrate(acct)


def dep_integrate(acct, year, src_fp, subj_fp, prefix='折旧复核-', label='折旧复核'):
    """折旧/账龄等复核表 → 注入科目底稿（prefix 前缀命名）。"""
    if not os.path.exists(subj_fp):
        print(f'  ⚠️ {acct}: 目标科目底稿不存在：{os.path.basename(subj_fp)}')
        return 0
    if not os.path.exists(src_fp):
        print(f'  ⚠️ {acct}: 复核表不存在：{os.path.basename(src_fp)}')
        return 0
    return _inject_sheets(subj_fp, src_fp, prefix=prefix or f'{label}-')
