# -*- coding: utf-8 -*-
"""折旧计提复核 数据源清单（对标 bank_reconcile_manifest，2026-08-15）。

铁律119：路径一律相对 DATA_ROOT（唯一锚点 paths.DATA_ROOT / AUDIT_DATA_ROOT）。

字段说明：
  card       卡片台账路径（相对 DATA_ROOT；16 个重复文件任取 1 份）
  org        核算主体（卡片台账"财务组织"）
  residual   残值率（默认 3%，可由卡片账反推校准）
  year       复核年度（查询期间）
  has_summary_row  是否含"合计"汇总行（True 需排除，防双计）
  land_kw    土地类关键词（不计提折旧，仅列示）
  map        资产台账映射表路径（相对 DATA_ROOT；卡片类别→底稿科目/披露类别，Sheet3 类别归类用）
  audit      固定资产审计底稿路径（相对 DATA_ROOT；Sheet3 与财务账面类别比较的数据源）
"""
import os

from paths import DATA_ROOT


def jobs():
    """返回 {账套: 折旧复核配置}。"""
    return {
        'XBJ': {
            'card': os.path.join('XBJ', '数据', '2025', '固定资产卡片台账', '卡片台账1.xlsx'),
            'org': '新疆兵团水利水电工程集团有限公司',
            'residual': 0.03,
            'year': 2025,
            'has_summary_row': True,
            'land_kw': ('土地',),
            'map': os.path.join('XBJ', '数据', '2025', '资产台账映射表0121.xlsx'),
            'audit': os.path.join('XBJ', '底稿', '2025', '固定资产审计底稿_2025_生成.xlsx'),
        },
    }


def job_for(acct):
    jobs_ = jobs()
    return jobs_.get(acct)


def resolve(rel):
    return os.path.join(DATA_ROOT, rel)
