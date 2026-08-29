# -*- coding: utf-8 -*-
"""应收+合同资产 联动账龄 数据源清单（铁律119：相对路径，配置化）。

对标 lease_manifest / depreciation_manifest。每个账套/子集团一配置。

字段说明：
  orgs_dir     数据目录（数据/{year}/ 或 数据/{year}/{sub}/）
  year         复核年度
  codes        科目码方言（账套差异必须配置化）：
                 ar        应收账款科目码前缀（1122）
                 ca        合同资产科目码前缀（XBJ 1481 / AJ-ga 1125 / AH 1124）
                 baddebt   坏账准备科目码前缀（1231）
                 ca_qual   质保金子目关键词（判断合同资产类别用）
  ca_cats      合同资产类别：名称关键词 → 未到期比例（类别优先，后续可逐合同细化）
                 {'质保': 1.0, '审计': 1.0, '工程': 0.8, '_default': 0.8}
                 未到期比例=该类合同资产在期末时点尚未达到无条件收款权、不进账龄桶的比例
  buckets      账龄桶（六档，与专项账龄复核六档一致）
  provision    坏账准备计提比例（应收 vs 合同资产 各自账龄对应比例，2026-08-16 用户要求）：
                 {'ar': [0.05, 0.10, 0.20, 0.30, 0.50, 1.0],   # 应收：1年以内~5年以上
                  'ca': [0.01, 0.05, 0.10, 0.15, 0.20, 0.50]}  # 合同资产：通常低于应收
                 桶索引 0~5 = 1年以内/1-2/2-3/3-4/4-5/5年以上；用户提供公司政策后覆盖，
                 未提供前为默认占位（Sheet3 标注"待公司政策确认"）。
  ⚡ 互转剔除（2026-08-16 用户方法论）：先按凭证号识别 应收借方=合同资产贷方 的互转，
    合并抵消 → 统一分龄 → 按合同拆回两科（各自账龄）→ 各自计提比例测算坏账准备。
  ⚡⚡ 付款条件→合同资产账龄（2026-08-16 用户要求，铁律131f 衔接）：
    合同资产=已确认收入尚未到结算条件的部分；未到期判断应按【每个合同的具体付款条件】——
    结算周期（月结=1月/季结=3月/半年=6月/竣工结算=竣工时点）+ 质保期（quality_months 质保金未到期）。
    合同台账在 project_manifest.contracts 配置（与项目制收入底稿同源），
    按 project 键（业主名/WBS）匹配 → 逐合同判断未到期金额，替代固定比例。
    无台账项目 → 回退 ca_cats 类别比例（默认 0.8）。
  说明：账龄=按 GL 交易日期 FIFO（先进先出）测算；合同资产先按未到期比例拆
        分"未到期"(单独列示) 与"已到期"(转应收进账龄)。
"""
import os

from paths import DATA_ROOT

# ⚡ 坏账准备计提比例默认占位（用户提供公司政策后覆盖；桶 0~5 = 1年以内~5年以上）
DEFAULT_PROVISION = {
    'ar': [0.05, 0.10, 0.20, 0.30, 0.50, 1.00],
    'ca': [0.01, 0.05, 0.10, 0.15, 0.20, 0.50],
}


def resolve(rel):
    return os.path.join(DATA_ROOT, rel)


def jobs():
    return {
        'XBJ': {
            'subs': {'_root': {
                'org': '新疆兵团水利水电工程集团（XBJ）',
                'data_dir': os.path.join('XBJ', '数据', '2025'),
                'year': 2025,
                'codes': {'ar': ['1122'], 'ca': ['1481'], 'baddebt': ['1231']},
                'ca_cats': [
                    {'kw': '质保金', 'unexpired': 1.0},
                    {'kw': '审计金', 'unexpired': 1.0},
                    {'kw': '合同款', 'unexpired': 0.8},
                    {'kw': '_default', 'unexpired': 0.8},
                ],
            }},
        },
        'AJ': {
            'subs': {
                'ga': {
                    'org': '浙江省工业设备安装集团（ga）',
                    'data_dir': os.path.join('AJ', '数据', '2025', 'ga'),
                    'year': 2025,
                    'codes': {'ar': ['1122'], 'ca': ['1125'], 'baddebt': ['1231']},
                    'ca_cats': [
                        {'kw': '质保金', 'unexpired': 1.0},
                        {'kw': '_default', 'unexpired': 0.8},
                    ],
                },
                'jj': {
                    'org': '浙江省工业设备安装集团（jj）',
                    'data_dir': os.path.join('AJ', '数据', '2025', 'jj'),
                    'year': 2025,
                    'codes': {'ar': ['1122'], 'ca': ['1125'], 'baddebt': ['1231']},
                    'ca_cats': [
                        {'kw': '质保金', 'unexpired': 1.0},
                        {'kw': '_default', 'unexpired': 0.8},
                    ],
                },
            },
        },
        # AH：SAP 账套尚未迁移到 数据/{yy}/ 布局（_work_ 平铺+序时账/ 分月），
        # 账龄接入待 AH 迁移后补（需 SAP 专用读取器，与 lease 同思路）
        # 'AH': {
        #     'subs': {'_root': {
        #         'org': 'AH SAP 集团',
        #         'data_dir': os.path.join('AH', '数据', '2026'),
        #         'year': 2026,
        #         'codes': {'ar': ['1122'], 'ca': ['1124'], 'baddebt': ['1231']},
        #         'ca_cats': [
        #             {'kw': '质保', 'unexpired': 1.0},
        #             {'kw': '_default', 'unexpired': 0.8},
        #         ],
        #     }},
        # },
    }


def job_for(acct, sub):
    aj = jobs()[acct]
    return aj, aj['subs'][sub]
