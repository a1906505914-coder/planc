# -*- coding: utf-8 -*-
"""毛利率分析 —— 审计要求知识点。

包装 gross_margin.build（读 TB 收入/成本二级科目，按主体/期间算毛利率，异常波动标红）。
默认输出：{data_dir}/毛利率分析表_生成.xlsx（可用 params.out_path 覆盖）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import sap_adapter as A  # noqa: E402
import gross_margin as GM  # noqa: E402


@register_requirement(
    name='gross_margin',
    desc='毛利率分析（收入/成本二级科目逐主体逐期，异常波动核查）',
    subjects=('收入', '存货', '利润表'),
    params=('out_path',),
)
def gross_margin(ctx):
    data_dir = ctx['data_dir']
    A.set_root(data_dir)  # 知识点包装负责初始化数据上下文（build 内部不 set_root，读数据统一走 _DATA_ROOT）
    out = ctx.get('out_path') or os.path.join(data_dir, '毛利率分析表_生成.xlsx')
    GM.build(data_dir, out)
    if ctx.get('verbose'):
        print(f'  [gross_margin] 输出：{out}')
    return [out] if os.path.exists(out) else []
