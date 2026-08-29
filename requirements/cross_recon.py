# -*- coding: utf-8 -*-
"""科目间交叉核对 —— 审计要求知识点。

包装 cross_recon.build（跨科目/跨表勾稽核对，发现科目间不一致）。
默认输出：{data_dir}/交叉核对_生成.xlsx（可用 params.out_path 覆盖）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import sap_adapter as A  # noqa: E402
import cross_recon as CR  # noqa: E402


@register_requirement(
    name='cross_recon',
    desc='科目间交叉核对（跨科目/跨表勾稽，发现不一致）',
    subjects=('银行存款', '往来款', '收入', '存货', '费用'),
    params=('out_path',),
)
def cross_recon(ctx):
    data_dir = ctx['data_dir']
    A.set_root(data_dir)  # 知识点包装负责初始化数据上下文（build 内部不 set_root，读数据统一走 _DATA_ROOT）
    out = ctx.get('out_path') or os.path.join(data_dir, '交叉核对_生成.xlsx')
    CR.build(data_dir, out)
    if ctx.get('verbose'):
        print(f'  [cross_recon] 输出：{out}')
    return [out] if os.path.exists(out) else []
