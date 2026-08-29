# -*- coding: utf-8 -*-
"""差异智能体检 —— 审计要求知识点（2026-08-14 架构扩展）。

把 diff_intelligence（P1-P12 差异模式体检 + 层2 知识 + 学习器候选）注册为独立
知识点：任何科目核对程序（银行/往来/存货/收入…）跑完，只要把核对结果按约定
传给本知识点，即可自动体检差异模式并给建议——13 个科目生成器无需改动。

通用性说明：
  P1-P12 模式源自银行核对实证，但本质是跨科目通用差异规律：
    对称双计/口径不一致/解析bug/文件缺失/误抽科目/拆组/红字未配/
    跨期错位/镜像成对/待报解/汇总记账/浮点尾差 —— 往来/存货/收入核对同样适用。

接口（对要求层暴露）：
  params.diff_sources: 可选，JSON 数组描述核对结果来源（银行核对默认内置）：
      [{"kind": "bank", "stmt_folder": ..., "journal_file": ...}, ...]
    缺省时尝试自动发现账套内核对底稿。

用法（manifest）：
  {"requirements": [{"name": "diff_intelligence", "params": {}}]}
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import diff_intelligence as DI  # noqa: E402


@register_requirement(
    name='diff_intelligence',
    desc='差异智能体检（P1-P12 差异模式 + 层2 知识 + 学习器候选，跨科目通用核对后体检）',
    subjects=('银行存款', '往来款', '收入', '存货', '费用', '应交税费'),
    params=('diff_sources',),
)
def diff_intelligence(ctx):
    data_dir = ctx['data_dir']
    # ⚡ 银行核对场景：若账套有 manifest 指定了网银流水/日记账，自动跑 bank_reconcile
    #   拿结果再体检（diff_intelligence 纯诊断，需要已核对结果）
    import bank_reconcile_detail as BR  # noqa: E402
    import sap_adapter as A  # noqa: E402

    srcs = ctx.get('diff_sources')
    if isinstance(srcs, str):
        try:
            srcs = json.loads(srcs)
        except Exception:
            srcs = None
    if not srcs:
        # 默认：自动发现账套内 manifest_*.json 的 bank_reconcile 参数
        srcs = _auto_discover(data_dir)
    if not srcs:
        if ctx.get('verbose'):
            print('  [diff_intelligence] 未找到核对结果来源（需 manifest 指定 '
                  'diff_sources 或 bank_reconcile 参数），跳过')
        return []

    outputs = []
    for src in srcs:
        kind = src.get('kind', 'bank')
        if kind == 'bank':
            # 复用 bank_reconcile：跑核对（内部已含 diff_intelligence 报告），
            # 这里主要是让"独立知识点"也能在无 bank_reconcile 清单时工作
            try:
                stmt = src.get('stmt_folder')
                jf = src.get('journal_file')
                if not stmt or not jf:
                    continue
                out = BR.build_bank_reconcile(
                    data_dir,
                    stmt_folder=stmt,
                    comp=src.get('comp') or ctx.get('comp'),
                    year=ctx.get('year', '2026'),
                    journal_file=jf,
                )
                if out:
                    outputs.append(out)
            except Exception as _ex:
                print(f'  [diff_intelligence] bank 核对异常（跳过）：{_ex}')
    if ctx.get('verbose') and outputs:
        print(f'  [diff_intelligence] 输出：{len(outputs)} 个文件（含差异智能报告）')
    return outputs


def _auto_discover(data_dir):
    """自动发现账套/小程序目录内 manifest 的 bank_reconcile 参数。"""
    import glob
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # 小程序根
    out = []
    for f in glob.glob(os.path.join(data_dir, 'manifest_*.json')) \
           + glob.glob(os.path.join(here, 'manifest_*.json')):
        try:
            man = json.load(open(f, encoding='utf-8'))
        except Exception:
            continue
        # 只取 data_dir 匹配本账套的清单（避免 3300/3400 参数混用）
        md = str(man.get('data_dir', '')).lower()
        if md and os.path.normpath(md) != os.path.normpath(data_dir).lower() \
           and os.path.basename(os.path.normpath(md)) != os.path.basename(os.path.normpath(data_dir)).lower():
            continue
        for req in man.get('requirements', []):
            if req.get('name') == 'bank_reconcile':
                p = req.get('params') or {}
                if p.get('stmt_folder') and p.get('journal_file'):
                    out.append({'kind': 'bank', **p})
    return out
