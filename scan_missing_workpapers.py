# -*- coding: utf-8 -*-
"""scan_missing_workpapers.py —— 期望矩阵扫描器（2026-08-13 用户方法论：#737 质检收口核心）
「系统从未定义『应该有哪些底稿』→ 缺失静默存在」的治本工具：
① discover 全部主体 + 读 TB 全量（控制数）
② 对每主体 × 每科目族：TB 该族有无数据（qc/qm/jf/df 任一非零，铁律 110）→ 有=应有底稿
③ 实际底稿文件（*审计底稿_{comp}.xlsx）比对 → 缺失即报
用法：python scan_missing_workpapers.py <账套根目录> [--out <报告路径>]
"""
import os
import sys
import glob
import time
import warnings
from mask_dict import mask_names   # ⚡ console 主体列表打印脱敏（2026-08-22 复扫）

warnings.filterwarnings('ignore')

# ⚡ 科目族定义：文件名词干 + [(TB代码前缀, TB名称关键词)]（任一匹配且有发生/余额=应有底稿）
#   代码前缀按 SAP 方言（2026-08-09 实测），名称关键词兜底（防代码错位）
SUBJECTS = [
    # 长期资产（longterm 6 组）
    ('固定资产', [('1601', '固定资产')]),
    ('在建工程', [('1604', '在建工程')]),
    ('无形资产', [('1701', '无形资产')]),
    ('使用权资产', [('1608', '使用权资产')]),
    ('投资性房地产', [('1521', '投资性房地产')]),
    ('长期待摊费用', [('1801', '长期待摊费用')]),
    # 权益（equity）
    ('实收资本', [('4001', '实收资本')]),
    ('盈余公积', [('4101', '盈余公积')]),
    ('未分配利润', [('4104', '未分配利润')]),
    ('专项储备', [('4102', '专项储备')]),
    ('资本公积', [('4002', '资本公积')]),
    ('库存股', [('4201', '库存股')]),
    # gp_other（SAP 无独立「一年内到期」科目——空代码前缀=仅按名称匹配，
    # 防 2401 递延收益/2402 预提费用误判）
    ('一年内到期', [('', '一年内到期')]),
    ('递延收益', [('2401', '递延收益')]),
    ('递延所得税资产', [('1811', '递延所得税资产')]),
    ('递延所得税负债', [('1812', '递延所得税负债')]),
    # 存货
    ('存货', [('1403', '原材料'), ('1405', '库存商品'), ('1406', '发出商品'),
              ('1412', '包装物'), ('1408', '委托加工物资'), ('1461', '存货跌价准备')]),
    # 收入/成本/费用
    ('营业收入', [('6001', '营业收入'), ('6002', '营业收入')]),
    ('营业成本', [('6003', '营业成本')]),
    ('销售费用', [('6601', '销售费用')]),
    ('管理费用', [('6602', '管理费用')]),
    ('研发费用', [('6602', '研发费用')]),
    ('财务费用', [('6603', '财务费用')]),
    ('税金及附加', [('6403', '税金及附加')]),
    ('其他收益', [('6730', '其他收益')]),
    ('营业外收入', [('6301', '营业外收入')]),
    ('营业外支出', [('6711', '营业外支出')]),
    ('信用减值损失', [('6741', '信用减值损失')]),
    ('资产减值损失', [('6701', '资产减值损失')]),
    ('所得税费用', [('6801', '所得税费用')]),
    ('资产处置损益', [('6751', '资产处置损益')]),
    ('投资收益', [('6111', '投资收益')]),
    # 薪酬/税/银行
    ('职工薪酬', [('2211', '应付职工薪酬')]),
    ('应交税费', [('2221', '应交税费')]),
    ('银行存款', [('1002', '银行存款')]),
    # 借款/债券
    ('短期借款', [('2001', '短期借款')]),
    ('长期借款', [('2501', '长期借款')]),
    ('应付债券', [('2502', '应付债券')]),
    ('租赁负债', [('2260', '租赁负债')]),
    ('长期应付款', [('2701', '长期应付款')]),
    ('预计负债', [('2601', '预计负债')]),
    ('专项应付款', [('2702', '专项应付款')]),
    ('交易性金融资产', [('1101', '交易性金融资产')]),
    ('长期股权投资', [('1511', '长期股权投资')]),
    ('应收票据', [('1121', '应收票据')]),
    ('应收账款', [('1122', '应收账款')]),
    ('预付账款', [('1123', '预付账款')]),
    ('其他应收款', [('1221', '其他应收款')]),
    ('应付账款', [('2202', '应付账款')]),
    ('预收账款', [('2203', '预收账款')]),
    ('其他应付款', [('2241', '其他应付款')]),
    ('合同资产', [('1223', '合同资产')]),
    ('合同负债', [('2204', '合同负债')]),
    ('应付票据', [('2201', '应付票据')]),
]


def _has_data(rows, pfx_list, kw_list):
    """TB 行中该科目族有无数据（末级叶子净额任一非零；防父子互抵误判：
    8030 递延所得税负债 父 18907 + 子 -18907 = 0 → 真无数据不误报）。
    pfx 空串=仅按名称匹配（如一年内到期）。"""
    if not pfx_list or pfx_list == ['']:
        for (code, name, val) in rows:
            if any(k in str(name) for k in kw_list) \
                    and any(str(val.get(k) or 0) not in ('', '0', '0.0', 0, 0.0)
                            for k in ('qc', 'jf', 'df', 'qm')):
                return True
        return False
    for p in pfx_list:
        fam = [(c, n, v) for (c, n, v) in rows if str(c).startswith(p)]
        if not fam:
            continue
        # 叶子 = 族内不被其他代码前缀的（10 位末级子目）
        leaves = [(c, n, v) for (c, n, v) in fam
                  if not any(c != c2 and c2.startswith(c) for (c2, _n2, _v2) in fam)]
        tot = [0.0] * 4
        for _c, _n, v in leaves:
            for i, k in enumerate(('qc', 'jf', 'df', 'qm')):
                tot[i] += float(v.get(k) or 0)
        if any(abs(x) > 1e-6 for x in tot):
            return True
        # 名称兜底：按名称过滤的「族」同样算叶子净额（防互抵误判：
        # 1010 租赁负债 未确认融资 -143452 + 未支付付款额 143452 = 0 → 无数据）
        name_rows = [(c, n, v) for (c, n, v) in fam if any(k in str(n) for k in kw_list)]
        if name_rows:
            nleaves = [x for x in name_rows
                       if not any(x[0] != c2 and c2.startswith(x[0]) for (c2, _n2, _v2) in name_rows)]
            nt = [0.0] * 4
            for _c, _n, v in nleaves:
                for i, k in enumerate(('qc', 'jf', 'df', 'qm')):
                    nt[i] += float(v.get(k) or 0)
            if any(abs(x) > 1e-6 for x in nt):
                return True
    return False


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if not args:
        print('用法：python scan_missing_workpapers.py <账套根目录> [--out 报告路径]')
        return 1
    root = os.path.abspath(args[0])
    out_path = None
    if '--out' in sys.argv:
        out_path = sys.argv[sys.argv.index('--out') + 1]

    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import sap_adapter as A
    A.set_root(root)

    # ⚡ 支持子目录主体形态（300/6100）：子目录 discover 空 → 上溯父目录并取子目录名=comp
    single_comp = None
    ents = A.discover_entities(None)
    if not ents:
        up = os.path.dirname(root)
        A.set_root(up)
        ents = A.discover_entities(None)
        if ents:
            single_comp = os.path.basename(root)
            root = up
            print(f'识别为 {up} 下的主体目录，主体={single_comp}')
        else:
            print(f'[ERROR] 无法识别账套结构：{root}')
            return 1

    print(f'[1/3] discover 主体…')
    comps = [single_comp] if single_comp else sorted(ents)
    print(f'  主体 {len(comps)} 家: {mask_names(comps, 10)}')

    print('[2/3] 读 TB 全量（控制数）…')
    tb = A.read_tb_full(root, None)
    tb_by_comp = {}
    for (e, c, n, y), v in tb.items():
        tb_by_comp.setdefault(e, []).append((c, n, v))
    print(f'  TB 主体 {len(tb_by_comp)} 家 / 行 {len(tb)}')

    print('[3/3] 底稿文件比对…')
    found = set()
    for pat in (os.path.join(root, '底稿', '*审计底稿*.xlsx'),
                os.path.join(root, '*审计底稿*.xlsx'),
                os.path.join(root, '*', '*审计底稿*.xlsx')):
        for f in glob.glob(pat):
            base = os.path.basename(f)
            for comp in comps:
                if f'_审计底稿_{comp}.' in base or f'审计底稿_{comp}.' in base:
                    stem = base.split('审计底稿')[0]
                    found.add((comp, stem))
    print(f'  实际底稿文件对 (主体,科目族) 共 {len(found)}')

    missing, skip = [], []
    for comp in comps:
        rows = tb_by_comp.get(comp, [])
        stems_comp = {fs for (c, fs) in found if c == comp}
        for stem, rules in SUBJECTS:
            pfx_list = [r[0] for r in rules]
            kw_list = [r[1] for r in rules]
            if not _has_data(rows, pfx_list, kw_list):
                continue  # TB 无该族数据 → 正常无底稿
            # ⚡ 模糊匹配：文件名 stem 可能带后缀变体（『实收资本(或股本)』vs『实收资本』）
            if not any(fs.startswith(stem) or stem in fs for fs in stems_comp):
                missing.append((comp, stem))
    # SKIP 反向：有底稿但 TB 无数据（残留/串户嫌疑）
    for comp, stem in sorted(found):
        rules = next((s for s in SUBJECTS if s[0] == stem), None)
        if rules is None:
            continue
        rows = tb_by_comp.get(comp, [])
        if not _has_data(rows, [r[0] for r in rules[1]], [r[1] for r in rules[1]]):
            skip.append((comp, stem))

    print()
    print('══════════ 期望矩阵扫描结果 ══════════')
    print(f'主体 {len(comps)} 家 × 科目族 {len(SUBJECTS)} 组 = 理论组合 {len(comps) * len(SUBJECTS)}')
    print(f'❌ 缺失（TB 有数据但无底稿）：{len(missing)}')
    for comp, stem in missing:
        print(f'    {comp} {stem}审计底稿')
    print(f'⚠️ 残留/串户嫌疑（有底稿但 TB 无该族数据）：{len(skip)}')
    for comp, stem in skip[:30]:
        print(f'    {comp} {stem}审计底稿')
    if len(skip) > 30:
        print(f'    … 其余 {len(skip) - 30} 项')

    report = (f'期望矩阵扫描 {time.strftime("%Y-%m-%d %H:%M")} — 缺失 {len(missing)} / 残留嫌疑 {len(skip)}\n'
              + f'主体 {len(comps)} 家，科目族 {len(SUBJECTS)} 组\n'
              + '【缺失】\n' + '\n'.join(f'{c} {s}审计底稿' for c, s in missing)
              + '\n【残留/串户嫌疑】\n' + '\n'.join(f'{c} {s}审计底稿' for c, s in skip))
    if out_path:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f'\n报告已写：{out_path}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
