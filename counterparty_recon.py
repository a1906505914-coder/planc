# -*- coding: utf-8 -*-
"""通用对方科目核对生成器（#875 试点：损益类科目）。

背景：对方科目核对目前仅往来科目（current_account）与银行存款有。其余 46 个科目
（资产/负债/权益/损益类）缺失。本脚本用 GL 序时账按【凭证翻面】还原每个科目的
对方科目，按 主体×方向×对方科目 聚合，并按期望/关注/异常映射分类，输出
『{科目}对方科目核对_{集团}.xlsx』（独立文件，不触碰正式底稿 → 零回归）。

用法：
    python counterparty_recon.py            # 跑配置内全部试点科目 × 三集团
    python counterparty_recon.py 税金及附加 # 只跑指定科目
    python counterparty_recon.py --list     # 列出已配置科目
输出：D:/底稿测试/AH/中间产物/counterparty_recon/{集团}/{科目}对方科目核对_{集团}.xlsx

GL 行字段（sap_adapter.read_gl_rows）：e,y,date,month,vtype,vno,name,code,dr,cr,cp,...
"""
import os, sys, glob, json, time
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import sap_adapter

DATA = r'D:/底稿测试/AH/数据/2026'
OUT_ROOT = r'D:/底稿测试/AH/中间产物/counterparty_recon'
ROOT = r'D:/底稿测试/AH/底稿/2026/集团'   # 正式底稿目录（inject 用）
GROUPS = {
    '1010': ['1010'],
    '1357': ['1020', '1030', '1040', '1050', '1060', '1070', '1080', '1090', '1100',
             '1170', '1220', '1240', '1250', '1260', '3010', '5020', '9020', '9030', '9070'],
    '2468': ['2010', '2020', '2030', '2040', '2050', '2070', '2080', '2090', '2100',
             '2110', '2130', '2150', '2160', '2200', '2210', '2220', '2230', '2240',
             '2250', '2260', '2270', '2280', '2290', '2300', '2310', '2330', '2340',
             '2350', '2360', '2370', '2380', '2390', '2400', '2410', '2420', '2430',
             '2450', '2460', '2470', '2480', '2490', '2500', '2510', '2520', '2530',
             '2540', '2550', '2570', '2580', '2590', '2600', '2610', '2630', '2640',
             '2650', '2660', '2680', '2690', '2700', '2720', '2730', '2740', '2760',
             '4030', '4040', '8020', '8030', '8040'],
}

# 科目配置：kw=名称关键词（命中即目标科目），expected_dr/expected_cr=该侧期望对方科目关键词
# （期望命中=正常；未命中=关注；note_map 带『★异常』=异常）
SPECS = {
    '税金及附加': dict(
        kw='税金及附加', codes=['6403'],
        expected_dr=['应交税费', '银行存款', '库存现金'],
        expected_cr=['应交税费', '本年利润', '税金及附加'],
        note_map={'营业外': '损益科目串户，核查', '管理费用': '费用串户，核查',
                  '其他应付款': '税金暂挂往来，关注结转时点'},
    ),
    '其他收益': dict(
        kw='其他收益', codes=['6117'],
        expected_cr=['递延收益', '银行存款', '其他应收款', '其他应付款', '应交税费'],
        expected_dr=['本年利润', '其他收益', '应交税费'],
        note_map={'营业外': '与营业外收入串户，核查', '应交税费': '含税费重分类，关注',
                  '应收账款': '政府补助挂应收，核查补助到账'},
    ),
    '营业外收入': dict(
        kw='营业外收入', codes=['6301'],
        expected_cr=['银行存款', '库存现金', '应收账款', '应付账款', '固定资产清理',
                     '递延收益', '长期应付款', '其他应付款', '应交税费', '预收账款'],
        expected_dr=['本年利润', '营业外收入', '应交税费'],
        note_map={'未分配利润': '★异常直接进权益，核查是否应计损益',
                  '营业成本': '成本费用串户，核查', '制造费用': '费用串户，核查',
                  '管理费用': '费用串户，核查', '销售费用': '费用串户，核查'},
    ),
    '营业外支出': dict(
        kw='营业外支出', codes=['6711'],
        expected_dr=['银行存款', '库存现金', '固定资产清理', '存货', '应付账款',
                     '其他应付款', '应收账款', '预付账款', '无形资产', '应交税费'],
        expected_cr=['本年利润', '营业外支出', '应交税费'],
        note_map={'未分配利润': '★异常直接进权益，核查是否应计损益',
                  '营业外收入': '收支对冲异常，核查', '应收账款': '坏账核销需核查依据',
                  '制造费用': '费用串户，核查'},
    ),
    '信用减值损失': dict(
        kw='信用减值损失', codes=['6702'],
        expected_dr=['坏账准备', '应收票据', '其他应收款', '应收账款'],
        expected_cr=['坏账准备', '本年利润', '信用减值损失'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '损益串户，核查'},
    ),
    '资产减值损失': dict(
        kw='资产减值损失', codes=['6701'],
        expected_dr=['存货跌价准备', '固定资产减值准备', '无形资产减值准备',
                     '长期股权投资', '合同资产', '在建工程', '投资性房地产'],
        expected_cr=['存货跌价准备', '固定资产减值准备', '无形资产减值准备',
                     '本年利润', '资产减值损失'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '损益串户，核查'},
    ),
    '所得税费用': dict(
        kw='所得税费用', codes=['6801'],
        expected_dr=['应交税费', '递延所得税资产', '递延所得税负债'],
        expected_cr=['应交税费', '递延所得税资产', '递延所得税负债', '本年利润', '所得税费用'],
        note_map={'未分配利润': '★异常直接进权益，核查递延/当期口径',
                  '营业外': '损益串户，核查'},
    ),
    '投资收益': dict(
        kw='投资收益', codes=['6111'],
        expected_dr=['长期股权投资', '银行存款', '应收股利', '交易性金融资产',
                     '其他权益工具投资', '公允价值变动', '其他货币资金'],
        expected_cr=['长期股权投资', '银行存款', '应收股利', '交易性金融资产',
                     '本年利润', '投资收益', '其他货币资金'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '损益串户，核查',
                  '主营业务': '收入与投资收益串户，核查'},
    ),
    '资产处置损益': dict(
        kw='资产处置损益', codes=['6115'],
        expected_dr=['固定资产清理', '无形资产', '在建工程', '银行存款', '其他应付款'],
        expected_cr=['固定资产清理', '无形资产', '银行存款', '本年利润', '资产处置损益'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '与营业外收支口径，核查',
                  '累计折旧': '处置结转不完整，核查'},
    ),
    '研发费用': dict(
        kw='研发费用', codes=['6604', '660406', '6600'], fr={'5301'},
        expected_dr=['银行存款', '应付职工薪酬', '累计折旧', '累计摊销', '原材料',
                     '其他应付款', '库存现金', '应付账款'],
        expected_cr=['银行存款', '本年利润', '研发费用', '应付职工薪酬'],
        note_map={'研发支出': '资本化/费用化口径，核查', '营业外': '损益串户，核查',
                  '未分配利润': '★异常直接进权益'},
    ),
    '管理费用': dict(
        kw='管理费用', codes=['6602', '6600'], fr={'6602'},
        expected_dr=['银行存款', '库存现金', '应付职工薪酬', '累计折旧', '累计摊销',
                     '其他应付款', '应付账款', '应交税费', '预付账款', '待摊费用'],
        expected_cr=['银行存款', '本年利润', '管理费用', '应付职工薪酬', '其他应付款'],
        note_map={'营业外': '损益串户，核查', '制造费用': '费用串户，核查',
                  '研发费用': '研发/管理口径，核查', '未分配利润': '★异常直接进权益'},
    ),
    '销售费用': dict(
        kw='销售费用', codes=['6601', '6600'], fr={'6601'},
        expected_dr=['银行存款', '库存现金', '应付职工薪酬', '累计折旧', '累计摊销',
                     '其他应付款', '应付账款', '应交税费', '预付账款'],
        expected_cr=['银行存款', '本年利润', '销售费用', '应付职工薪酬', '其他应付款'],
        note_map={'营业外': '损益串户，核查', '制造费用': '费用串户，核查',
                  '管理费用': '费用串户，核查', '未分配利润': '★异常直接进权益'},
    ),
    '财务费用': dict(
        kw='财务费用', codes=['6603'],
        expected_dr=['银行存款', '应付利息', '长期借款', '短期借款', '应付债券',
                     '其他应付款', '应交税费', '库存现金'],
        expected_cr=['银行存款', '本年利润', '财务费用', '应付利息', '其他应付款'],
        note_map={'营业外': '损益串户，核查', '投资收益': '利息收入与投资收益口径，核查',
                  '未分配利润': '★异常直接进权益'},
    ),
    '制造费用': dict(
        kw='制造费用', codes=['5101', '6600'], fr={'5101'},
        expected_dr=['银行存款', '库存现金', '应付职工薪酬', '累计折旧', '累计摊销',
                     '原材料', '其他应付款', '应付账款', '应交税费', '辅助生产成本'],
        expected_cr=['生产成本', '制造费用', '银行存款', '本年利润'],
        note_map={'营业外': '损益串户，核查', '管理费用': '费用串户，核查',
                  '销售费用': '费用串户，核查', '未分配利润': '★异常直接进权益'},
    ),
    '研发支出': dict(
        kw='研发支出', codes=['5301', '1704'],
        expected_dr=['银行存款', '应付职工薪酬', '累计折旧', '累计摊销', '原材料',
                     '其他应付款', '库存现金', '应付账款', '长期待摊费用'],
        expected_cr=['管理费用', '无形资产', '研发支出', '银行存款', '本年利润'],
        note_map={'营业外': '损益串户，核查', '未分配利润': '★异常直接进权益',
                  '资本化': '资本化/费用化口径，核查'},
    ),
    '应收票据': dict(
        kw='应收票据', codes=['1121'],
        expected_dr=['应收账款', '主营业务收入', '其他业务收入', '应交税费', '银行存款'],
        expected_cr=['银行存款', '应收账款', '应付账款', '短期借款', '应收票据', '主营业务收入'],
        note_map={'营业外': '非常规结算，核查', '在建工程': '票据结算工程款，关注',
                  '未分配利润': '★异常直接进权益'},
    ),
    '投资性房地产': dict(
        kw='投资性房地产', codes=['1521'],
        expected_dr=['银行存款', '固定资产', '无形资产', '在建工程', '应付账款', '长期借款'],
        expected_cr=['固定资产', '无形资产', '公允价值变动', '营业外收入', '银行存款',
                     '累计折旧', '投资性房地产'],
        note_map={'未分配利润': '★异常直接进权益', '主营业务': '与主业收入串户，核查'},
    ),
    '递延所得税资产': dict(
        kw='递延所得税资产', codes=['1811'],
        expected_dr=['所得税费用', '资本公积', '其他综合收益'],
        expected_cr=['所得税费用', '资本公积', '其他综合收益', '递延所得税资产'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '损益串户，核查'},
    ),
    '其他权益工具投资': dict(
        kw='其他权益工具投资', codes=['1501'],
        expected_dr=['银行存款', '资本公积', '其他综合收益', '其他货币资金'],
        expected_cr=['银行存款', '资本公积', '其他综合收益', '留存收益', '其他权益工具投资'],
        note_map={'投资收益': '公允价值变动进损益，核查指定依据', '未分配利润': '★异常直接进权益'},
    ),
    '长期股权投资': dict(
        kw='长期股权投资', codes=['1511'],
        expected_dr=['银行存款', '固定资产', '无形资产', '其他货币资金', '资本公积',
                     '投资收益', '应收股利', '长期应付款'],
        expected_cr=['银行存款', '投资收益', '应收股利', '资本公积', '长期股权投资', '其他货币资金'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '处置损益口径，核查'},
    ),
    '长期股权投资减值准备': dict(
        kw='长期股权投资减值准备', codes=['1512'],
        expected_dr=['资产减值损失', '投资收益'],
        expected_cr=['长期股权投资', '资产减值损失', '投资收益'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '损益串户，核查'},
    ),
    '一年内到期的非流动负债': dict(
        kw='一年内到期的非流动负债', codes=['2241'],
        expected_dr=['银行存款', '长期借款', '长期应付款', '应付债券'],
        expected_cr=['长期借款', '长期应付款', '应付债券', '银行存款', '一年内到期的非流动负债'],
        note_map={'营业外': '重分类异常，核查', '未分配利润': '★异常直接进权益'},
    ),
    '合同负债': dict(
        kw='合同负债', codes=['2241', '2203'],
        expected_dr=['主营业务收入', '其他业务收入', '应交税费', '银行存款', '预收账款'],
        expected_cr=['银行存款', '应收账款', '主营业务收入', '合同负债'],
        note_map={'营业外': '非常规结转，核查', '营业成本': '与收入配比，核查',
                  '未分配利润': '★异常直接进权益'},
    ),
    '短期借款': dict(
        kw='短期借款', codes=['2001'],
        expected_dr=['银行存款', '其他货币资金', '财务费用'],
        expected_cr=['银行存款', '其他货币资金', '短期借款', '财务费用'],
        note_map={'营业外': '借款费用异常，核查', '长期借款': '长短期串户，核查',
                  '未分配利润': '★异常直接进权益'},
    ),
    '长期借款': dict(
        kw='长期借款', codes=['2501'],
        expected_dr=['银行存款', '其他货币资金', '财务费用', '应付利息'],
        expected_cr=['银行存款', '其他货币资金', '长期借款', '财务费用', '应付利息', '一年内到期的非流动负债'],
        note_map={'营业外': '借款费用异常，核查', '短期借款': '长短期串户，核查',
                  '未分配利润': '★异常直接进权益'},
    ),
    '长期应付款': dict(
        kw='长期应付款', codes=['2701'],
        expected_dr=['银行存款', '固定资产', '在建工程', '财务费用', '未确认融资费用'],
        expected_cr=['银行存款', '固定资产', '在建工程', '长期应付款', '未确认融资费用'],
        note_map={'营业外': '融资租赁处理异常，核查', '未分配利润': '★异常直接进权益'},
    ),
    '应付债券': dict(
        kw='应付债券', codes=['2502'],
        expected_dr=['银行存款', '财务费用', '应付利息'],
        expected_cr=['银行存款', '应付债券', '财务费用', '应付利息', '其他货币资金'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '债券发行费用异常，核查'},
    ),
    '应付票据': dict(
        kw='应付票据', codes=['2201'],
        expected_dr=['银行存款', '应付账款', '其他货币资金', '短期借款'],
        expected_cr=['应付账款', '银行存款', '原材料', '固定资产', '应付票据', '主营业务成本'],
        note_map={'应收票据': '票据背书，关注', '营业外': '非常规结算，核查',
                  '未分配利润': '★异常直接进权益'},
    ),
    '应交税费': dict(
        kw='应交税费', codes=['2221'],
        expected_dr=['银行存款', '税金及附加', '所得税费用', '应收账款', '其他应收款', '营业外收入'],
        expected_cr=['银行存款', '税金及附加', '所得税费用', '主营业务收入', '其他业务收入',
                     '管理费用', '财务费用', '营业外支出', '应交税费'],
        note_map={'未分配利润': '★异常直接进权益，核查税金', '制造费用': '税金计入成本，关注',
                  '应付职工薪酬': '个税/社保代扣，关注'},
    ),
    '职工薪酬': dict(
        kw='职工薪酬', codes=['2211'],
        expected_dr=['银行存款', '库存现金', '其他应付款', '应交税费'],
        expected_cr=['生产成本', '制造费用', '管理费用', '销售费用', '研发费用', '研发支出',
                     '在建工程', '应付职工薪酬'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '薪酬挂账异常，核查',
                  '应收账款': '代垫薪酬，关注', '其他应收款': '代垫薪酬，关注'},
    ),
    '专项储备': dict(
        kw='专项储备', codes=['4102'],
        expected_dr=['银行存款', '应付账款', '其他应付款', '安全投入', '制造费用', '管理费用'],
        expected_cr=['制造费用', '管理费用', '生产成本', '专项储备', '银行存款'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '安全费使用异常，核查'},
    ),
    '专项应付款': dict(
        kw='专项应付款', codes=['2711'],
        expected_dr=['银行存款', '其他应收款', '营业外收入', '递延收益', '管理费用'],
        expected_cr=['银行存款', '其他应收款', '专项应付款', '营业外支出'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '专项拨款结转，关注',
                  '递延收益': '与递延收益口径，核查'},
    ),
    '递延收益': dict(
        kw='递延收益', codes=['2401'],
        expected_dr=['银行存款', '其他应收款', '其他收益', '营业外收入'],
        expected_cr=['银行存款', '其他应收款', '递延收益', '营业外支出'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '政府补助结转口径，核查',
                  '应交税费': '补助涉税，关注'},
    ),
    '预计负债': dict(
        kw='预计负债', codes=['2801'],
        expected_dr=['银行存款', '其他应付款', '营业外支出', '管理费用'],
        expected_cr=['营业外支出', '管理费用', '销售费用', '预计负债', '银行存款'],
        note_map={'未分配利润': '★异常直接进权益', '主营业务': '产品质量保证金，关注'},
    ),
    '固定资产': dict(
        kw='固定资产', codes=['1601'],
        expected_dr=['银行存款', '应付账款', '在建工程', '其他应付款', '资本公积',
                     '固定资产清理', '长期应付款'],
        expected_cr=['固定资产清理', '银行存款', '应收账款', '营业外支出', '营业外收入',
                     '固定资产', '累计折旧'],
        note_map={'未分配利润': '★异常直接进权益', '主营业务': '资产购置与收入串户，核查',
                  '累计折旧': '处置转固不完整，核查'},
    ),
    '无形资产': dict(
        kw='无形资产', codes=['1701'],
        expected_dr=['银行存款', '应付账款', '研发支出', '资本公积', '其他应付款', '长期应付款'],
        expected_cr=['无形资产', '营业外支出', '营业外收入', '银行存款', '累计摊销', '长期待摊费用'],
        note_map={'未分配利润': '★异常直接进权益', '主营业务': '资产购置与收入串户，核查',
                  '研发支出': '资本化转入，关注'},
    ),
    '在建工程': dict(
        kw='在建工程', codes=['1604'],
        expected_dr=['银行存款', '应付账款', '原材料', '应付职工薪酬', '库存商品',
                     '累计折旧', '资本公积', '其他应付款'],
        expected_cr=['固定资产', '银行存款', '营业外支出', '在建工程', '投资性房地产'],
        note_map={'未分配利润': '★异常直接进权益', '主营业务': '工程款与收入串户，核查',
                  '管理费用': '费用化与资本化口径，核查'},
    ),
    '使用权资产': dict(
        kw='使用权资产', codes=['1611'],
        expected_dr=['租赁负债', '银行存款', '长期应付款'],
        expected_cr=['租赁负债', '累计折旧', '银行存款', '使用权资产', '营业外支出'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '租赁终止处理，核查'},
    ),
    '存货': dict(
        kw='存货', codes=['1405', '1403', '1406', '1411', '1412'],
        expected_dr=['银行存款', '应付账款', '生产成本', '原材料', '库存商品',
                     '主营业务成本', '在建工程', '其他应付款', '制造费用'],
        expected_cr=['主营业务成本', '生产成本', '存货', '营业外支出', '银行存款',
                     '应付账款', '原材料', '库存商品', '在建工程'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '盘亏/报废处理，核查',
                  '固定资产': '存货与资产互转，关注', '管理费用': '存货管理费串户，核查'},
    ),
    '实收资本(或股本)': dict(
        kw='实收资本', codes=['4001'],
        expected_dr=['银行存款', '资本公积', '库存股', '其他货币资金'],
        expected_cr=['银行存款', '资本公积', '库存股', '其他货币资金', '实收资本'],
        note_map={'未分配利润': '★异常直接进权益，核查增减资', '营业外': '权益与损益串户，核查',
                  '应付账款': '债转股，关注'},
    ),
    '资本公积': dict(
        kw='资本公积', codes=['4002'],
        expected_dr=['实收资本', '银行存款', '长期股权投资', '库存股', '资本公积'],
        expected_cr=['实收资本', '银行存款', '长期股权投资', '库存股', '资本公积', '其他权益工具投资'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '权益与损益串户，核查',
                  '其他综合收益': '与 OCI 口径，核查'},
    ),
    '盈余公积': dict(
        kw='盈余公积', codes=['4101'],
        expected_dr=['利润分配', '实收资本', '盈余公积'],
        expected_cr=['利润分配', '盈余公积', '资本公积'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '权益与损益串户，核查',
                  '本年利润': '盈余公积直接结转损益，核查'},
    ),
    '未分配利润': dict(
        kw='未分配利润', codes=['4104'],
        expected_dr=['本年利润', '利润分配', '盈余公积', '应付股利', '银行存款'],
        expected_cr=['本年利润', '利润分配', '盈余公积', '未分配利润', '银行存款'],
        note_map={'营业外': '损益直接进权益，核查', '主营业务': '损益直接进权益，核查',
                  '管理费用': '损益直接进权益，核查', '资本公积': '权益内部结转，关注'},
    ),
    '库存股': dict(
        kw='库存股', codes=['4201'],
        expected_dr=['银行存款', '其他货币资金'],
        expected_cr=['银行存款', '资本公积', '利润分配', '实收资本', '库存股'],
        note_map={'未分配利润': '★异常直接进权益', '营业外': '权益与损益串户，核查'},
    ),
}


def _vkey(r):
    return (r.get('e'), r.get('y'), r.get('month'), r.get('vtype'), r.get('vno'))


def _is_target(r, spec):
    nm = str(r.get('name') or '')
    cd = str(r.get('code') or '')
    hit = False
    if spec.get('codes'):
        hit = any(cd.startswith(c) for c in spec['codes'])
    if not hit and spec.get('kw'):
        hit = spec['kw'] in nm
    if hit and spec.get('fr'):
        # AH/SAP：费用科目在 6600 总池，按功能范围列（fr）区分制造/管理/销售/研发
        if str(r.get('fr') or '') not in spec['fr']:
            return False
    return hit


def _norm_opp(nm):
    """对方科目名规范化：去代码/去层级前缀，取一级名。"""
    nm = str(nm or '')
    for sep in ('-', '—', ' '):
        if sep in nm:
            return nm.split(sep)[0]
    return nm


def _classify(opp, spec, side):
    opp_n = _norm_opp(opp)
    exps = spec.get('expected_dr' if side == '借' else 'expected_cr', [])
    for e in exps:
        if e and (e in opp or opp.startswith(e) or opp_n.startswith(e)):
            return '正常', ''
    note = ''
    for k, v in spec.get('note_map', {}).items():
        if k in opp:
            note = v
            break
    if note:
        if '★异常' in note:
            return '异常', note
        if '核查' in note or '关注' in note:
            return '关注', note
        return '关注', note
    return '关注', '非期望对应科目，需核实业务实质与凭证'


def aggregate(gl_rows, year, spec, ents=None):
    """按凭证翻面还原对方科目 → {(ent, side, opp): {amt, cnt}}。ents=None=全部主体。"""
    sub = [r for r in gl_rows
           if r.get('y') == str(year)
           and (ents is None or r.get('e') in ents) and _is_target(r, spec)]
    if not sub:
        return {}
    target_keys = set()
    for r in sub:
        target_keys.add(_vkey(r))
    non_sub = defaultdict(list)
    for r in gl_rows:
        if r.get('y') != str(year):
            continue
        if ents is not None and r.get('e') not in ents:
            continue
        k = _vkey(r)
        if k in target_keys and not _is_target(r, spec):
            non_sub[k].append(r)
    agg = defaultdict(lambda: {'amt': 0.0, 'cnt': 0})
    for k, trows in _group_by(sub):
        ent = trows[0].get('e', '')
        D = sum(float(r['dr'] or 0) for r in trows)
        C = sum(float(r['cr'] or 0) for r in trows)
        oth = non_sub.get(k, [])
        if D > 0:
            for r in oth:
                if float(r['cr'] or 0) > 0:
                    key = (ent, '借', r.get('name') or r.get('cp') or '（未知）')
                    agg[key]['amt'] += float(r['cr'] or 0)
                    agg[key]['cnt'] += 1
        if C > 0:
            for r in oth:
                if float(r['dr'] or 0) > 0:
                    key = (ent, '贷', r.get('name') or r.get('cp') or '（未知）')
                    agg[key]['amt'] += float(r['dr'] or 0)
                    agg[key]['cnt'] += 1
    return dict(agg)


def write_sheet(ws, agg, year, spec):
    """把 agg 写入 ws（对方科目核对）。返回数据行数。"""
    ncols = 9
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.cell(1, 1, f'{spec["kw"]}（{"/".join(spec.get("codes") or [])}）对方科目核对（{year} 年度）')
    ws.cell(2, 1, '说明：按凭证翻面还原对方科目，按 核算主体×方向×对应科目 归集；'
                  '红色=异常（★），黄色=关注，未标=正常。')
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    hdr = ['核算主体', '年度', '方向', '对应科目（科目名称）', '金额', '笔数', '占该侧%', '是否异常', '异常说明']
    for j, h in enumerate(hdr, 1):
        c = ws.cell(3, j, h)
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='DDEBF7')
        c.alignment = Alignment(horizontal='center')
    side_tot = defaultdict(float)
    for (ent, side, opp), v in agg.items():
        side_tot[(ent, side)] += v['amt']
    r = 4
    for (ent, side, opp), v in sorted(agg.items(), key=lambda x: (x[0][0], x[0][1], -x[1]['amt'])):
        verdict, note = _classify(opp, spec, side)
        pct = v['amt'] / side_tot[(ent, side)] * 100 if side_tot.get((ent, side)) else 0.0
        ws.cell(r, 1, ent); ws.cell(r, 2, year); ws.cell(r, 3, side)
        ws.cell(r, 4, opp)
        ws.cell(r, 5, round(v['amt'], 2))
        ws.cell(r, 5).number_format = '#,##0.00'   # 2026-08-29 千分位
        ws.cell(r, 6, v['cnt'])
        ws.cell(r, 7, round(pct, 4))
        ws.cell(r, 8, verdict)
        ws.cell(r, 9, note)
        if verdict == '异常':
            for j in range(1, ncols + 1):
                ws.cell(r, j).fill = PatternFill('solid', fgColor='F4CCCC')
        elif verdict == '关注':
            for j in range(1, ncols + 1):
                ws.cell(r, j).fill = PatternFill('solid', fgColor='FFF2CC')
        r += 1
    for (ent, side) in sorted(side_tot):
        ws.cell(r, 1, ent); ws.cell(r, 3, side + '方小计')
        ws.cell(r, 5, round(side_tot[(ent, side)], 2))
        ws.cell(r, 5).number_format = '#,##0.00'   # 2026-08-29 千分位
        r += 1
    for j, w in enumerate([14, 8, 6, 34, 16, 8, 10, 8, 46], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = 'A4'
    return r - 4


def build_one(gl_rows, year, ent, spec, out_path):
    agg = aggregate(gl_rows, year, spec, ents={ent})
    if not agg:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '对方科目核对'
    write_sheet(ws, agg, year, spec)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)
    return len(agg)


def inject_into_wb(wb, gl, year, spec, ents=None):
    """在【内存】wb 追加『对方科目核对』sheet（模块内集成用，不 load+save → 零破坏）。
    已在生成器保存前调用，重跑自动带出。返回 'ok'/'skip'/'empty'。"""
    if any('对方科目核对' in s for s in wb.sheetnames):
        return 'skip'
    agg = aggregate(gl, year, spec, ents=ents)
    if not agg:
        return 'empty'
    ws = wb.create_sheet('对方科目核对')
    write_sheet(ws, agg, year, spec)
    return 'ok'


def inject_into_wb_auto(wb, data_dir, year, spec_key, ents_set=None):
    """模块内集成统一入口：内部读全量 GL（sap_adapter + 磁盘缓存）后注入。
    ents_set=主体集合（None=全部）。用于 tax/payroll 等模块内 gl 非全量的场景。"""
    spec = SPECS.get(spec_key)
    if spec is None:
        return 'skip'
    if any('对方科目核对' in s for s in wb.sheetnames):
        return 'skip'
    try:
        import sap_adapter as SA
        if not SA._DATA_ROOT:
            SA._DATA_ROOT = data_dir
        root = SA._DATA_ROOT or data_dir
        full = SA.discover_entities(root)
        if ents_set:
            full = {c: full.get(str(c), {}) for c in ents_set}
        gl = SA.read_gl_rows(root, full)
    except Exception:
        from audit_common import read_gl_rows, discover_entities
        full = discover_entities(data_dir)
        if ents_set:
            full = {c: full.get(str(c), {}) for c in ents_set}
        gl = read_gl_rows(data_dir, full)
    return inject_into_wb(wb, gl, year, spec, ents=set(full))


def _fp_to_spec(fp, group):
    """由底稿文件名推导 spec 键；无法映射返回 None。"""
    stem = os.path.basename(fp)[:-5]   # 去 .xlsx
    suf = f'审计底稿_{group}'
    if not stem.endswith(suf):
        return None
    key = stem[: -len(suf)]
    return SPECS.get(key)


def inject_all(gl_by_group, year='2026'):
    """把对方科目核对 sheet 注入正式底稿（每个底稿文件追加同名 sheet）。
    返回 (ok, skip, empty, err) 计数。load+save 完整模式，逐文件处理。"""
    from collections import Counter
    stat = Counter()
    for g, gl in gl_by_group.items():
        d = os.path.join(ROOT, g)
        if not os.path.isdir(d):
            continue
        ents = set(GROUPS.get(g, []))
        for fp in sorted(glob.glob(os.path.join(d, '*.xlsx'))):
            fn = os.path.basename(fp)
            if fn.startswith('~$'):
                continue
            spec = _fp_to_spec(fp, g)
            if spec is None:
                continue
            try:
                wb = openpyxl.load_workbook(fp)
            except Exception:
                stat['err'] += 1
                continue
            if any('对方科目核对' in s for s in wb.sheetnames):
                wb.close()
                stat['skip'] += 1
                continue
            agg = aggregate(gl, year, spec, ents=ents)
            if not agg:
                wb.close()
                stat['empty'] += 1
                continue
            ws = wb.create_sheet('对方科目核对')
            write_sheet(ws, agg, year, spec)
            wb.save(fp)
            stat['ok'] += 1
            print(f'  ✓ {g}/{fn}', flush=True)
    return stat


def _group_by(rows):
    out = defaultdict(list)
    for r in rows:
        out[_vkey(r)].append(r)
    return out.items()


def main():
    argv = sys.argv[1:]
    if '--list' in argv:
        print('已配置科目：')
        for k, s in SPECS.items():
            print(f'  {k}（{"/".join(s.get("codes") or [])}）kw={s["kw"]}')
        return
    only = None
    if argv and not argv[0].startswith('-'):
        only = argv[0]
    sap_adapter._DATA_ROOT = DATA
    # 读 GL 缓存（按集团）
    gl_cache = {}
    for g, comps in GROUPS.items():
        ents = {c: sap_adapter.discover_entities(DATA).get(c, {}) for c in comps}
        gl_cache[g] = sap_adapter.read_gl_rows(DATA, ents)
        print(f'  [GL] {g} {len(gl_cache[g])} 行', flush=True)
    t0 = time.time()
    if only == 'inject':
        # 注入正式底稿：每个底稿文件追加『对方科目核对』sheet
        stat = inject_all(gl_cache, year='2026')
        print(f'注入完成：OK {stat["ok"]} / 已存在跳过 {stat["skip"]} / 无数据 {stat["empty"]} / 失败 {stat["err"]}'
              f'（耗时 {time.time()-t0:.0f}s）')
        return
    made = 0
    for name, spec in SPECS.items():
        if only and name != only:
            continue
        for g, comps in GROUPS.items():
            gl = gl_cache[g]
            for ent in comps:
                out = os.path.join(OUT_ROOT, g, f'{name}对方科目核对_{g}.xlsx')
                n = build_one(gl, '2026', ent, spec, out)
                if n:
                    made += 1
    print(f'完成：{made} 个 主体×科目 文件 → {OUT_ROOT}（耗时 {time.time()-t0:.0f}s）')


if __name__ == '__main__':
    main()
