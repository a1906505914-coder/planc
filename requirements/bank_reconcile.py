# -*- coding: utf-8 -*-
"""银行网银双向核对 —— 第一个审计要求知识点（2026-08-14 架构试点）。

包装 bank_reconcile_detail.build_bank_reconcile（内部 reconcile 逻辑不改，实战已验证：
3100/3200/3300/3400/3000/DQ 全账套口径一致），对外暴露知识点接口：
声明输入参数 → 引擎按清单注入 → 执行 → 返回输出文件。

⚡ 2026-08-14 架构收口：核对完成后自动调用 diff_intelligence.run_intelligence
（P1-P12 体检 + 层2 知识 + 学习器候选统一报告）——build 内部已集成，此处仅透出
知识点的"差异智能"能力说明。

数据依赖（声明式）：
    data_dir        : 账套目录（含 TB/GL 数据，或 sap_adapter 可 discover 的根）
    stmt_folder     : 网银流水文件夹（如 '3300网银明细2026年1-7月'）
    journal_file    : 银行日记账文件（SAP 多主体 XLSX 或 U8 表单）
    comp            : 主体代码（缺省时从网银文件夹名提取）
    year            : 年度（缺省 2026）
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import bank_reconcile_detail as BR  # noqa: E402


@register_requirement(
    name='bank_reconcile',
    desc='银行网银双向核对（网银流水 vs 银行日记账，含月度对账/差异清单/跨账户归因/'
         '双计剔除/差异智能体检 P1-P12/学习器候选）',
    subjects=('银行存款',),
    params=('stmt_folder', 'journal_file'),
)
def bank_reconcile(ctx):
    import paths as P
    data_dir = ctx['data_dir']
    # ⚡ 2026-08-14 params 路径字段支持相对 DATA_ROOT（清单相对化约定）
    stmt_folder = P.resolve_rel(ctx.get('stmt_folder'))
    journal_file = P.resolve_rel(ctx.get('journal_file'))
    out = BR.build_bank_reconcile(
        data_dir,
        stmt_folder=stmt_folder,
        comp=ctx.get('comp'),
        year=ctx.get('year', '2026'),
        journal_file=journal_file,
    )
    if ctx.get('verbose') and out:
        print(f'  [bank_reconcile] 输出：{out}')
    return [out] if out else []
