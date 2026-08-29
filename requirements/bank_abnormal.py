# -*- coding: utf-8 -*-
"""银行存款发生额异常分析 —— 审计要求知识点。

包装 bank_abnormal.build（银行存款发生额异常/大额异常波动检测）。
默认输出：{data_dir}/银行存款发生额异常分析_生成.xlsx（可用 params.out_path 覆盖）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import sap_adapter as A  # noqa: E402
import bank_abnormal as BA  # noqa: E402


@register_requirement(
    name='bank_abnormal',
    desc='银行存款发生额异常分析（大额异常波动检测）',
    subjects=('银行存款',),
    params=('out_path',),
)
def bank_abnormal(ctx):
    data_dir = ctx['data_dir']
    A.set_root(data_dir)  # 知识点包装负责初始化数据上下文（build 内部不 set_root，读数据统一走 _DATA_ROOT）
    out = ctx.get('out_path') or os.path.join(data_dir, '银行存款发生额异常分析_生成.xlsx')
    BA.build(data_dir, out)
    if ctx.get('verbose'):
        print(f'  [bank_abnormal] 输出：{out}')
    return [out] if os.path.exists(out) else []
