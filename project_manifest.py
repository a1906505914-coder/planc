# -*- coding: utf-8 -*-
"""项目制核算 收入底稿 数据源清单（铁律119：相对路径，配置化；铁律131f 2026-08-16）。

背景（用户 2026-08-16 需求）：
  项目制核算（施工/工程类）收入底稿按【项目/合同】编制——客户名称/合同内容/合同周期/
  合同约定收入(可能含税)/预算总成本/时段法确认(前期+本期 收入/成本/毛利率)/前期已收款/
  本期收款/期末应收款。
  暂无合同台账 → 降级模式：账务侧按【项目维度键】聚合（XBJ=辅助核算往来单位名=业主、
  AH SAP=WBS(col17)/合同号(col15)/订单(col24)），合同要素列留空待补；
  有台账后（manifest contracts 或读台账 xlsx）自动匹配填列。

字段说明：
  data_dir    数据目录（数据/{year}/）
  year        复核年度
  project_key 项目维度键规则：
               'aux_unit'  = 辅助核算「往来单位名称」（XBJ/U8，业主名）
               'sap_wbs'   = SAP WBS 元素（col17）
               'sap_contract' = SAP 合同号（col15）
  rev_codes   收入科目码前缀（主营业务收入/其他业务收入；XBJ 123301 价款结算兜底）
  cost_codes  成本科目码前缀（合同履约成本/主营业务成本）
  recv_codes  应收账款科目码前缀（辅助核算按业主取期末应收）
  bank_codes  银行存款科目码前缀（收款）
  contracts   可选合同台账（留空 = 无台账降级模式；配置后按 project 键匹配填列）：
               [{'project': '业主名/WBS', 'customer': '客户名称',
                 'content': '合同内容', 'start': '2024-03-01', 'end': '2026-06-30',
                 'contract_amount': 10000000, 'tax_inclusive': True,
                 'budget_cost': 8000000, 'method': '投入法',
                 'settle_cycle': '月结',      # ⚡ 付款条件-结算周期：月结/季结/半年结/竣工结算（合同资产到期时点判断）
                 'quality_months': 12}]       # ⚡ 付款条件-质保期（月）：质保金未到期时长（默认 0=无质保金）
  ⚡ 付款条件 → 合同资产账龄判断（2026-08-16 用户要求）：合同资产=已确认收入尚未到结算条件的部分；
    按 结算周期（月结=1月/季结=3月/半年=6月/竣工=竣工时点）推算每笔合同资产应转应收的时点，
    超过时点未转=逾期进账龄；质保金按质保期（quality_months）判断未到期。
"""
import os

from paths import DATA_ROOT


def resolve(rel):
    return os.path.join(DATA_ROOT, rel)


def jobs():
    return {
        'XBJ': {
            'subs': {'_root': {
                'org': '新疆兵团水利水电工程集团（XBJ）',
                'data_dir': os.path.join('XBJ', '数据', '2025'),
                'year': 2025,
                'project_key': 'aux_unit',
                'rev_codes': ['6001', '6051', '1233'],
                'cost_codes': ['5002', '5401', '6401'],
                'recv_codes': ['1122', '1123'],
                'bank_codes': ['1002', '1004'],
                'contracts': [],   # 暂无合同台账，留空降级
            }},
        },
        # AH：SAP 账套（迁移后接入，需 SAP 专用读取器；项目键=WBS/合同号）
        # 'AH': {
        #     'subs': {'_root': {
        #         'org': 'AH SAP 集团',
        #         'data_dir': os.path.join('AH', '数据', '2026'),
        #         'year': 2026,
        #         'project_key': 'sap_wbs',
        #         'rev_codes': ['6001', '6051'],
        #         'cost_codes': ['6401', '5001'],
        #         'recv_codes': ['1122'],
        #         'bank_codes': ['1002'],
        #         'contracts': [],
        #     }},
        # },
    }


def job_for(acct, sub):
    aj = jobs()[acct]
    return aj, aj['subs'][sub]
