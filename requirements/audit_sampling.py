# -*- coding: utf-8 -*-
"""审计抽样 —— 审计要求知识点。

audit_sampling 是【注入器】（往已生成的抽凭 sheet 注入重要性水平/样本量表头，
由 finalize_workbook 自动调用）。知识点化后：
    ① 独立执行：对指定底稿文件（params.files）或 data_dir 下自动发现的底稿，
       批量注入抽样头——用于"底稿已生成但抽样头缺失/需要重算样本量"的场景；
    ② 也可先用 compute_min_sample_size 试算样本量（params.om/pm/sad）。
默认处理：data_dir 下文件名含『底稿』的 .xlsx。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import audit_sampling as SAM  # noqa: E402


def _discover_workbooks(data_dir):
    """自动发现 data_dir 下的底稿工作簿（文件名含『底稿』）。"""
    if not os.path.isdir(data_dir):
        return []
    return [os.path.join(data_dir, f) for f in sorted(os.listdir(data_dir))
            if f.endswith('.xlsx') and '底稿' in f and not f.startswith('~')]


@register_requirement(
    name='audit_sampling',
    desc='审计抽样（重要性水平/样本量表头注入抽凭底稿 + 样本量试算）',
    subjects=('收入', '往来款', '存货', '费用'),
    params=('files', 'om', 'pm', 'sad'),
)
def audit_sampling(ctx):
    import openpyxl
    data_dir = ctx['data_dir']
    files = ctx.get('files') or _discover_workbooks(data_dir)
    if not files:
        print('  [audit_sampling] 未找到待注入底稿（可用 params.files 指定）')
        return []
    n_ok = 0
    for fp in files:
        try:
            wb = openpyxl.load_workbook(fp)
            wb._data_dir = data_dir
            SAM.inject_all_sampling_sheets(wb, data_dir)
            wb.save(fp)
            n_ok += 1
        except Exception as e:
            print(f'  [audit_sampling] {os.path.basename(fp)} 注入失败：{e}')
    if ctx.get('verbose'):
        print(f'  [audit_sampling] 注入 {n_ok}/{len(files)} 个底稿')
    return files
