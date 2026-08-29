# -*- coding: utf-8 -*-
"""存货监盘支持表 —— 审计要求知识点。

包装 inventory_count.build（存货数量/金额监盘支持数据）。
默认输出：{data_dir}/存货监盘支持表_生成.xlsx（可用 params.out_path 覆盖）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import sap_adapter as A  # noqa: E402
import inventory_count as IC  # noqa: E402


@register_requirement(
    name='inventory_count',
    desc='存货监盘支持表（存货数量/金额监盘数据）',
    subjects=('存货',),
    params=('out_path',),
)
def inventory_count(ctx):
    data_dir = ctx['data_dir']
    A.set_root(data_dir)  # 知识点包装负责初始化数据上下文（build 内部不 set_root，读数据统一走 _DATA_ROOT）
    out = ctx.get('out_path') or os.path.join(data_dir, '存货监盘支持表_生成.xlsx')
    IC.build(data_dir, out)
    if ctx.get('verbose'):
        print(f'  [inventory_count] 输出：{out}')
    return [out] if os.path.exists(out) else []
