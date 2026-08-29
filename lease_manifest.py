# -*- coding: utf-8 -*-
"""租赁合同复核 数据源清单（铁律119：相对路径，唯一锚点=paths.py DATA_ROOT）。

对标 bank_reconcile_manifest / depreciation_manifest 架构。
每账套一配置；subs 内每个子集团/目录一配置。

字段说明：
  org        核算主体（发现主体用）
  gl         综合查询明细表（序时账，反推合同参数的主源）
  tb         科目余额表（账面核对基准）
  year       复核年度
  codes      租赁科目码映射（账套方言）：
                use_asset      使用权资产原值科目码（如 1509）
                use_dep        使用权资产累计折旧（如 1510）
                lease_pay      租赁负债-租赁付款额（如 220601）
                unrecog_fin    租赁负债-未确认融资费用（如 220602）
                dta_lease      递延所得税资产\\租赁负债（如 181104）
                interest_fee   财务费用-利息支出 名称关键词（摊销借贷识别用）
                dep_subject    折旧借方科目名称关键词（'管理费用\\租赁费'）
                dep_display    累计折旧贷方科目名称关键词（'使用权资产累计折旧'）
                interest_disp  利息借方科目名称关键词（'利息支出'）
                fin_disp       未确认融资贷方名称关键词（'未确认融资费用'）
  months_in_year  年付款次数（默认12）
  round_dec       金额取整位数（默认2）

  ⚡ 手动模式（账面摊销可能错误时，用户提供基础信息，利率按 LPR 基准）：
    lpr              LPR 基准利率表（期限年 → 年利率）：
                      {'1': 0.030, '5': 0.035}  # 1年期 3.0% / 5年期以上 3.5%（2026-08 现行）
    lpr_method       LPR 取用规则：'period'=按合同付款期数对应期限（默认）
    contracts        [{name, asset, start_date, end_date, pay_amount, period_months,
                        payment_freq(月数,默认1), rate(可选,缺省按LPR), remark}]
                      — 提供后进入【手动测算模式】：现值=付款折现、摊销=实际利率法，
                        与账面核对（TB 科目码仍从 codes 取）
  注意：科目余额表字段=opening/closing/debit/credit（read_km），序时账列=
[科目名称,日期,凭证字,凭证号,对方科目,摘要,借方金额,贷方金额]。
"""
import os

from paths import DATA_ROOT

# ⚡⚡ LPR 历史基准利率（年化）——按【合同开始日】对应时点取用（用户 2026-08-16 要求）
#   折现率应在租赁开始日确定，用当时市场的 LPR（增量借款利率），而非当前值。
#   数据来源：全国银行间同业拆借中心（中行/农行官网 LPR 报价历史，2019-08 起）。
#   格式：(生效日期 YYYY-MM-DD, 1年期, 5年期以上)
LPR_HISTORY = [
    ('2019-08-20', 0.0425, 0.0485),
    ('2019-09-20', 0.0420, 0.0485),
    ('2019-11-20', 0.0415, 0.0480),
    ('2020-02-20', 0.0405, 0.0475),
    ('2020-04-20', 0.0385, 0.0465),
    ('2022-01-20', 0.0370, 0.0460),
    ('2022-05-20', 0.0370, 0.0445),
    ('2022-08-22', 0.0365, 0.0430),
    ('2023-06-20', 0.0355, 0.0420),
    ('2023-08-21', 0.0345, 0.0420),
    ('2024-02-20', 0.0345, 0.0395),
    ('2024-07-22', 0.0335, 0.0385),
    ('2024-10-21', 0.0310, 0.0360),
    ('2025-05-20', 0.0300, 0.0350),   # 之后一直保持至今（2026-08）
]
# 默认（无历史匹配/合同开始日缺失时）→ 最新 LPR
LPR_LATEST = (LPR_HISTORY[-1][1], LPR_HISTORY[-1][2])


def lpr_rate(period_months, start_date=None):
    """按付款期数（月）取 LPR 年化基准 → 返回月利率。

    ⚡ start_date=租赁开始日（折现率应在开始日确定，用户 2026-08-16 要求）：
       · 提供 → 取 start_date 当日生效的 LPR（LPR_HISTORY 按日期定位）
       · 缺失 → 取最新 LPR（标注待补开始日）
    期数≤12个月 → 1年期LPR；>12个月 → 5年期以上LPR（保守口径）。"""
    annual1, annual5 = LPR_LATEST
    if start_date:
        sd = str(start_date)[:10]
        for eff, a1, a5 in reversed(LPR_HISTORY):
            if sd >= eff:
                annual1, annual5 = a1, a5
                break
    annual = annual1 if period_months <= 12 else annual5
    return annual / 12.0, annual


def lpr_rate_desc(period_months, start_date=None):
    """返回利率来源描述（供底稿备注/说明列）。"""
    annual1, annual5 = LPR_LATEST
    eff_label = '最新'
    if start_date:
        sd = str(start_date)[:10]
        for eff, a1, a5 in reversed(LPR_HISTORY):
            if sd >= eff:
                annual1, annual5 = a1, a5
                # 显示合同开始日时点（而非该档生效日，更直观）
                eff_label = f'{sd} 时点'
                break
    annual = annual1 if period_months <= 12 else annual5
    return f'LPR@{eff_label}({annual*100:.2f}%/年, 付款期{period_months}月)'


def resolve(acct):
    return os.path.join(DATA_ROOT, acct)


def jobs():
    return {
        'AJ': {
            'subs': {
                'ga': {
                    'org': '03浙江省工业设备安装集团有限公司上海分公司',
                    'gl': os.path.join('AJ', '数据', '2025', 'ga',
                                       '03浙江省工业设备安装集团有限公司上海分公司2025年综合查询明细表.xlsx'),
                    'tb': os.path.join('AJ', '数据', '2025', 'ga',
                                       '03浙江省工业设备安装集团有限公司上海分公司2025年科目余额表.xlsx'),
                    'year': 2025,
                    'codes': {
                        'use_asset': ['1509'],
                        'use_dep': ['1510'],
                        'lease_pay': ['220601'],
                        'unrecog_fin': ['220602'],
                        'dta_lease': ['181104'],
                        'dep_subject': ['管理费用', '租赁费'],
                        'dep_display': ['使用权资产累计折旧'],
                        'interest_disp': ['利息支出'],
                        'fin_disp': ['未确认融资费用'],
                    },
                    'months_in_year': 12,
                    # ⚡ 手动模式（账面摊销可能错误时启用）：取消注释并填基础信息，利率按 LPR。
                    #   启用后程序用 contracts 测算现值/摊销，与账面核对（不再从序时账反推）。
                    # 'contracts': [
                    #     {'name': 'CD座', 'asset': 'CD座办公室',
                    #      'start_date': '2024-01-01', 'end_date': '2026-12-31',
                    #      'pay_amount': 45233.33, 'period_months': 36, 'payment_freq': 1,
                    #      # 'rate': 0.03,   # 可选：明确年化利率；缺省按付款期数对应 LPR
                    #      },
                    # ],
                },
            },
        },
                'GFY': {
            'subs': {
        '01FY本级': {
            'org': '01FY本级（集团本级）',
            'tb': os.path.join('GFY', '数据', '2026', '01FY本级公司' + '2026年1-3月科目余额表.xlsx'),
            'year': 2026, 'year_months': 3,
            'codes': {
                'use_asset': ['1704'], 'use_dep': ['1705'],
                'lease_pay': ['2601.02'], 'unrecog_fin': ['2601.01'],
                'dta_lease': [], 'dep_subject': ['管理费用', '租赁费'],
                'dep_display': ['使用权资产折旧'], 'interest_disp': ['利息支出'],
                'fin_disp': ['未确认融资费用'],
            },
            'contracts': [
                {
                    'name': '七格13幢', 'start_date': '2025-04-01', 'end_date': '2030-03-31',
                    'period_months': 60,
                    'payment_freq': 'half', 'advance': True,
                    'payments': [740324.57, 740324.57, 740324.57, 740324.57, 784744.05, 784744.05, 784744.05, 784744.05, 784744.05, 784744.05],
                },
                {
                    'name': '建友', 'start_date': '2021-06-15', 'end_date': '2031-06-14',
                    'period_months': 120,
                    'payment_freq': 'half', 'advance': True,
                    # 10年20期半年付；年租金(免租47天已内含首年从2021-08-01起算)2,410,800/2,755,200/2,892,960/3,037,608/3,189,488/3,348,962/3,516,410/3,692,231/3,876,843/4,070,685；÷1.05不含税
                    'payments': [1148000.00, 1148000.00,
                                 1312000.00, 1312000.00,
                                 1377600.00, 1377600.00,
                                 1446480.00, 1446480.00,
                                 1518803.81, 1518803.81,
                                 1594743.81, 1594743.81,
                                 1674480.95, 1674480.95,
                                 1758205.24, 1758205.24,
                                 1846115.71, 1846115.71,
                                 1938421.43, 1938421.43],
                },
                {
                    'name': '建友103.203', 'start_date': '2025-05-12', 'end_date': '2035-05-11',
                    'period_months': 120,
                    'payment_freq': 'half', 'advance': True,
                    # 合同金额(不含税含物业费)第1-3年1,392,900/第4-6年1,504,332/第7-10年1,727,196，租金占70%剥离物业费→×0.7；半年付
                    'payments': [487515.00, 487515.00, 487515.00, 487515.00, 487515.00, 487515.00,
                                 526516.20, 526516.20, 526516.20, 526516.20, 526516.20, 526516.20,
                                 604518.60, 604518.60, 604518.60, 604518.60, 604518.60, 604518.60, 604518.60, 604518.60],
                },
                {
                    'name': '七格5幢', 'start_date': '2025-06-26', 'end_date': '2030-06-25',
                    'period_months': 60,
                    'payment_freq': 'half', 'advance': True,
                    'payments': [384224.91, 384224.91, 384224.91, 384224.91, 384224.91, 384224.91, 384224.91, 384224.91, 384224.91, 384224.91],
                },
                {
                    'name': '七格6幢', 'start_date': '2020-03-10', 'end_date': '2030-03-09',
                    'period_months': 120,
                    'payment_freq': 'half', 'advance': True,
                    # 10年20期半年付；年租金按合同：第1年1,159,325(免租22天已内含)、第2-3年1,232,197、第4-6年×1.07、第7-10年×1.10；÷1.05不含税
                    'payments': [552059.52, 552059.52,
                                 586760.48, 586760.48, 586760.48, 586760.48,
                                 627833.71, 627833.71, 627833.71, 627833.71, 627833.71, 627833.71,
                                 690617.08, 690617.08, 690617.08, 690617.08, 690617.08, 690617.08, 690617.08, 690617.08],
                },
                {
                    'name': '千邻-达峰', 'start_date': '2022-03-15', 'end_date': '2027-03-15',
                    'period_months': 60,
                    'payment_freq': 'year', 'advance': True,
                    # 5年期年付：第1年333,975扣免租60天=279,075(合同已写抵扣后金额)、第2-5年356,240/378,505/400,770/423,035；÷1.05不含税
                    'payments': [265785.71, 339276.19, 360480.95, 381685.71, 402890.48],
                },
                {
                    'name': '亨德利', 'start_date': '2025-02-08', 'end_date': '2028-02-07',
                    'period_months': 36,
                    'payment_freq': 'half', 'advance': True,
                    'payments': [50468.57, 50468.57, 50468.57, 50468.57, 50468.57, 50468.57],
                },
                {
                    'name': '美妆大厦', 'start_date': '2021-09-01', 'end_date': '2026-08-31',
                    'period_months': 60,
                    'payment_freq': 'year', 'advance': True,
                    # 前3年免租(2021.9-2024.8)，第4、5年各付458,455(含税→÷1.05=436,623.81)；付款实际延后到第4年起
                    'payments': [0, 0, 0, 436623.81, 436623.81],
                },
                {
                    'name': '七格9幢一层', 'start_date': '2025-08-07', 'end_date': '2030-08-06',
                    'period_months': 60,
                    'payment_freq': 'half', 'advance': True,
                    'payments': [91428.57, 91428.57, 91428.57, 91428.57, 91428.57, 91428.57, 91428.57, 91428.57, 91428.57, 91428.57],
                },
                {
                    'name': '珠宝城2楼', 'start_date': '2026-06-28', 'end_date': '2029-06-27',
                    'period_months': 36,
                    'payment_freq': 'half', 'advance': True,
                    'payments': [70080.00, 70080.00, 70080.00, 70080.00, 70080.00, 70080.00],
                },
                {
                    'name': '珠宝城3楼', 'start_date': '2026-06-28', 'end_date': '2029-06-27',
                    'period_months': 36,
                    'payment_freq': 'half', 'advance': True,
                    'payments': [98862.86, 98862.86, 98862.86, 98862.86, 98862.86, 98862.86],
                },
                {
                    'name': '军承设备', 'start_date': '2026-01-01', 'end_date': '2032-07-31',
                    'period_months': 78,
                    'payment_freq': 'year', 'advance': True,
                    'payments': [551670.8, 551670.8, 551670.8, 551670.8, 551670.8, 551670.8, 551670.8],
                },
                {
                    'name': '下沙大院', 'start_date': '2026-01-01', 'end_date': '2026-12-31',
                    'period_months': 12,
                    'payment_freq': 'quart', 'advance': True,
                    'payments': [2145613.25, 2145613.25, 2145613.25, 2145613.25],
                },
            ],
        },
        '09FY绍兴': {
            'org': '09FY绍兴（绍兴方圆）',
            'tb': os.path.join('GFY', '数据', '2026', '09FY绍兴公司' + '2026年1-3月科目余额表.xlsx'),
            'year': 2026, 'year_months': 3,
            'codes': {
                'use_asset': ['1704'], 'use_dep': ['1705'],
                'lease_pay': ['2601.02'], 'unrecog_fin': ['2601.01'],
                'dta_lease': [], 'dep_subject': ['管理费用', '租赁费'],
                'dep_display': ['使用权资产折旧'], 'interest_disp': ['利息支出'],
                'fin_disp': ['未确认融资费用'],
            },
            'contracts': [
                {
                    'name': '四楼', 'start_date': '2024-03-31', 'end_date': '2029-03-30',
                    'period_months': 38,
                    'payments': [486571.43, 501238.1, 516190.48],
                },
                {
                    'name': '一、二楼', 'start_date': '2024-11-01', 'end_date': '2029-10-31',
                    'period_months': 45,
                    'payments': [209238.1, 215523.81, 222000.0],
                },
                {
                    'name': '二楼（本年新签）', 'start_date': '2025-06-10', 'end_date': '2030-06-09',
                    'period_months': 53,
                    'payments': [54952.38, 56571.43, 58285.71, 60095.24],
                },
                {
                    'name': '精工广场（续签）', 'start_date': '2025-05-01', 'end_date': '2027-04-30',
                    'period_months': 15,
                    'payments': [189440.0],
                },
            ],
        },
        '05FY皮革': {
            'org': '05FY皮革（海宁皮革）',
            'tb': os.path.join('GFY', '数据', '2026', '05FY皮革公司' + '2026年1-3月科目余额表.xlsx'),
            'year': 2026, 'year_months': 3,
            'codes': {
                'use_asset': ['1704'], 'use_dep': ['1705'],
                'lease_pay': ['2601.02'], 'unrecog_fin': ['2601.01'],
                'dta_lease': [], 'dep_subject': ['管理费用', '租赁费'],
                'dep_display': ['使用权资产折旧'], 'interest_disp': ['利息支出'],
                'fin_disp': ['未确认融资费用'],
            },
            'contracts': [
                {
                    'name': '海宁皮革研究院大楼', 'start_date': '2025-01-01', 'end_date': '2029-12-31',
                    'period_months': 47,
                    'payments': [676531.29, 676531.29, 676531.29, 676531.29],
                },
            ],
        },
        '06FY智能': {
            'org': '06FY智能（方圆智能）',
            'tb': os.path.join('GFY', '数据', '2026', '06FY智能公司' + '2026年1-3月科目余额表.xlsx'),
            'year': 2026, 'year_months': 3,
            'codes': {
                'use_asset': ['1704'], 'use_dep': ['1705'],
                'lease_pay': ['2601.02'], 'unrecog_fin': ['2601.01'],
                'dta_lease': [], 'dep_subject': ['管理费用', '租赁费'],
                'dep_display': ['使用权资产折旧'], 'interest_disp': ['利息支出'],
                'fin_disp': ['未确认融资费用'],
            },
            'contracts': []
        },
        '10FY金属': {
            'org': '10FY金属（方圆金属）',
            'tb': os.path.join('GFY', '数据', '2026', '10FY金属公司' + '2026年1-3月科目余额表.xlsx'),
            'year': 2026, 'year_months': 3,
            'codes': {
                'use_asset': ['1704'], 'use_dep': ['1705'],
                'lease_pay': ['2601.02'], 'unrecog_fin': ['2601.01'],
                'dta_lease': [], 'dep_subject': ['管理费用', '租赁费'],
                'dep_display': ['使用权资产折旧'], 'interest_disp': ['利息支出'],
                'fin_disp': ['未确认融资费用'],
            },
            'contracts': [
                {
                    'name': '七格厂房(金属)', 'start_date': '2019-12-10', 'end_date': '2029-12-09',
                    'period_months': 47, 'payments': [859662, 945628, 1040191, 1144210],
                },
            ],

        },
        '11FY龙游': {
            'org': '11FY龙游（龙游方圆）',
            'tb': os.path.join('GFY', '数据', '2026', '11FY龙游公司' + '2026年1-3月科目余额表.xlsx'),
            'year': 2026, 'year_months': 3,
            'codes': {
                'use_asset': ['1704'], 'use_dep': ['1705'],
                'lease_pay': ['2601.02'], 'unrecog_fin': ['2601.01'],
                'dta_lease': [], 'dep_subject': ['管理费用', '租赁费'],
                'dep_display': ['使用权资产折旧'], 'interest_disp': ['利息支出'],
                'fin_disp': ['未确认融资费用'],
            },
            'contracts': [
                {
                    'name': '龙游产业创新服务综合体三楼', 'start_date': '2024-05-18', 'end_date': '2030-05-17',
                    'period_months': 53, 'payments': [0, 129276, 129276, 129276, 48479],
                },
            ],

        },
            },
        },
    }


def job_for(acct, sub):
    """返回 (账套配置, 子集团配置)。"""
    aj = jobs()[acct]
    cfg = aj['subs'][sub]
    # ⚡ 租赁复核合同配置优先取《租赁合同台账》Excel（数据驱动，不再硬编码）；无台账/读失败回退 manifest
    if acct == 'GFY':
        try:
            from lease_ledger import contracts_for
            year = cfg.get('year', 2026)
            rows = contracts_for(acct, sub, year=year)
            if rows:
                cfg = dict(cfg)
                cfg['contracts'] = rows
        except Exception:
            pass
    return aj, cfg
