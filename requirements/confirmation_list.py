# -*- coding: utf-8 -*-
"""函证清单 —— 审计要求知识点。

包装 confirmation_list.build（读辅助核算应收/应付末级客商余额，生成函证清单底稿）。
默认输出：{data_dir}/函证清单_生成.xlsx（可用 params.out_path 覆盖）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from requirements import register_requirement  # noqa: E402

import sap_adapter as A  # noqa: E402
import confirmation_list as CONF  # noqa: E402


@register_requirement(
    name='confirmation_list',
    desc='函证清单（应收/应付末级客商余额汇总，关联方标注）',
    subjects=('往来款', '银行存款'),
    params=('out_path',),
)
def confirmation_list(ctx):
    data_dir = ctx['data_dir']
    A.set_root(data_dir)  # 知识点包装负责初始化数据上下文（build 内部不 set_root，读数据统一走 _DATA_ROOT）
    out = ctx.get('out_path') or os.path.join(data_dir, '函证清单_生成.xlsx')
    CONF.build(data_dir, out)
    if ctx.get('verbose'):
        print(f'  [confirmation_list] 输出：{out}')
    return [out] if os.path.exists(out) else []
