# -*- coding: utf-8 -*-
"""境外报表折算专项 manifest（2026-08-18 用户定：fx_consolidate/fx_rates/thai_mapping 提炼为一个专项）。

三者原为整体：fx_consolidate（外币报表折算合并主流程，process 已参数化）调用
fx_rates（汇率表 get_rate/to_cny）与 thai_mapping（泰国科目映射 normalize_thai_tb）。
现显式打包为专项（fx_review.py），本文件管配置。
新增账套 = 加一行（split_root + 币种），不新建脚本。

数据源：
  · Z 集团（paths.Z=AZ）分币种底稿目录 Z_split/{THB,USD}（currency_split 生成）
  · 汇率：fx_rates.py（期初 2024-12-31 / 期末 2025-12-31 / 平均 2025，人行中间价+统计局）
"""
import os

import paths as P


def _p(*parts):
    return os.path.join(*parts)


def jobs():
    return {
        'AZ': {   # 浙江兆龙互连（泰国三币种）
            'name': '境外报表折算（Z集团）',
            'split_root': os.path.join(P.Z, '数据', '2025'),   # ⚡ 分币种源数据目录 {CNY,THB,USD}（process 遍历 THB/USD）
            'out_dir': os.path.join(P.Z, '数据', '2025', '合并折算'),  # 输出（历史产物位置）
            'currencies': ['THB', 'USD'],
            'thai': True,        # 启用泰国科目映射（thai_mapping.normalize_thai_tb）
        },
    }


def job_for(acct):
    return jobs().get(acct)


if __name__ == '__main__':
    for acct, cfg in jobs().items():
        print(f'{acct:6s} {cfg["name"]} split_root={cfg["split_root"]}')
