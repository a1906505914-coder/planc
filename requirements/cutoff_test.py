# -*- coding: utf-8 -*-
"""收入/采购截止性测试 —— 审计要求知识点。

包装 cutoff_test.build（读 GL 窗口期凭证，分类销售/采购/其他，生成截止测试底稿）。
默认输出：{data_dir}/截止性测试底稿_生成.xlsx（可用 params.out_path 覆盖）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import sap_adapter as A  # noqa: E402
import cutoff_test as CUT  # noqa: E402


@register_requirement(
    name='cutoff_test',
    desc='收入/采购截止性测试（凭证日期窗口期扫描，防跨期确认收入/费用）',
    subjects=('收入', '存货', '费用'),
    params=('out_path',),
)
def cutoff_test(ctx):
    data_dir = ctx['data_dir']
    A.set_root(data_dir)  # 知识点包装负责初始化数据上下文（build 内部不 set_root，读数据统一走 _DATA_ROOT）
    out = ctx.get('out_path') or os.path.join(data_dir, '截止性测试底稿_生成.xlsx')
    CUT.build(data_dir, out)
    if ctx.get('verbose'):
        print(f'  [cutoff_test] 输出：{out}')
    return [out] if os.path.exists(out) else []
