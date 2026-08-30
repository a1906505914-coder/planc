# -*- coding: utf-8 -*-
"""financing_review.py —— 统一融资专项（2026-08-22）。

把「借款/表外/应付票据(承兑)/信用证」等融资相关底稿整合为一个专项入口，
注册进 special_reviews.REGISTRY（'financing'）。

产出（ADF/数据/2026）：
  ① 短期借款审计底稿_<年>_生成.xlsx / 长期借款...   ← loan_detail（账套数据）
  ② 表外事项明细底稿.xlsx                           ← contract_detail_ledger（合同脱敏版）
  ③ 融资全景_脱敏.xlsx                              ← 本模块（融资明细_脱敏.xlsx 按融资大类统计）
                                                       覆盖贷款/贸易融资/信用证/承兑(应付票据)/其他

安全：所有数据源用脱敏版（融资明细_脱敏.xlsx / 合同脱敏版），输出为 借X/贷X 代码。
用法：
  python financing_review.py                     # ADF 统一融资专项
  python financing_review.py --acct ADF          # 指定账套
"""
import paths as P
import argparse
import os
import sys

sys.stdout.reconfigure(encoding='utf-8', errors='replace')

DATA = os.path.join(P.DATA_DIRS['ADF'], '数据', '2026')
FIN_MASKED = os.path.join(DATA, '融资明细_脱敏.xlsx')
OUT_OVERVIEW = os.path.join(DATA, '融资全景_脱敏.xlsx')


def build_financing_overview(src=FIN_MASKED, out=OUT_OVERVIEW):
    """融资全景：主体×融资大类 业务金额/笔数/利率区间（读脱敏版）。返回输出路径。"""
    from openpyxl import Workbook, load_workbook
    from collections import defaultdict
    if not os.path.exists(src):
        print('  ⚠️ 融资明细脱敏版不存在：%s（先跑 融资台账脱敏.bat）' % src)
        return None
    wb = load_workbook(src, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    bigs = ['贸易融资', '流动资金贷款', '信用证', '项目贷款', '承兑', '委托贷款', '其他']
    cells = defaultdict(float)      # (ent, big) -> 金额
    cnt = defaultdict(int)
    rates = defaultdict(list)
    for row in ws.iter_rows(values_only=True):
        if not row:
            continue
        ent_raw = str(row[3] or '').strip()
        if ent_raw in ('', '融资单位'):        # 表头/空行跳过
            continue
        ent = str(row[3] or '').strip()          # 融资单位（借X）
        big = str(row[6] or '').strip()          # 融资大类
        try:
            amt = float(row[13] or 0)
        except (TypeError, ValueError):
            amt = 0.0
        try:
            rate = float(row[18]) if row[18] else None
        except (TypeError, ValueError):
            rate = None
        big = big if big in bigs else '其他'
        cells[(ent, big)] += amt
        cnt[(ent, big)] += 1
        if rate:
            rates[(ent, big)].append(rate)
    wb.close()

    wout = Workbook()
    wso = wout.active
    wso.title = '融资全景'
    wso.append(['融资全景（ADF 2026；单位：元；口径=融资明细业务金额；来源=融资明细_脱敏.xlsx）'])
    hdr = ['主体'] + bigs + ['合计']
    wso.append(hdr)
    ents = sorted({e for (e, _) in cells})
    for ent in ents:
        row = [ent]
        tot = 0.0
        for b in bigs:
            v = cells.get((ent, b), 0.0)
            row.append(round(v, 2) if v else 0.0)
            tot += v
        row.append(round(tot, 2))
        wso.append(row)
    # 全集团合计
    tot_row = ['全集团合计']
    tot_all = 0.0
    for b in bigs:
        v = sum(cells.get((e, b), 0.0) for e in ents)
        tot_row.append(round(v, 2) if v else 0.0)
        tot_all += v
    tot_row.append(round(tot_all, 2))
    wso.append(tot_row)
    wso.append([])
    wso.append(['说明：本表为融资业务口径（台账），与账套短期/长期借款账面口径差异=未入账/重分类待查。'])
    for col in wso.columns:
        wso.column_dimensions[col[0].column_letter].width = 16
    wout.save(out)
    print('  ✅ 融资全景 → %s（%d 主体 × %d 大类）' % (out, len(ents), len(bigs)))
    return out


def gen(acct='ADF'):
    """统一融资专项入口。返回产出文件列表。"""
    if acct != 'ADF':
        print(f'  ⏭️ 融资专项暂仅 ADF 有数据（{acct} 未配置）')
        return []
    outs = []
    # ① 借款底稿（loan_detail，账套数据）
    print('  ① 借款底稿（loan_detail）…')
    try:
        import loan_detail
        rc = loan_detail.process_folder(DATA)
        if rc == 0:
            outs += [os.path.join(DATA, f)
                     for f in ('短期借款审计底稿_2026_生成.xlsx', '长期借款审计底稿_2026_生成.xlsx')
                     if os.path.exists(os.path.join(DATA, f))]
    except Exception as ex:
        print(f'  ⚠️ 借款底稿生成失败：{ex}')
    # ② 表外事项（contract_detail_ledger，合同脱敏版）
    print('  ② 表外事项（contract_detail_ledger）…')
    try:
        import contract_detail_ledger as CDL
        CDL.main()
        fp = os.path.join(DATA, '表外事项明细底稿.xlsx')
        if os.path.exists(fp):
            outs.append(fp)
    except Exception as ex:
        print(f'  ⚠️ 表外事项生成失败：{ex}')
    # ③ 融资全景（本模块）
    fp = build_financing_overview()
    if fp:
        outs.append(fp)
    return outs


def main(argv=None):
    ap = argparse.ArgumentParser(description='统一融资专项（借款+表外+融资全景）')
    ap.add_argument('--acct', default='ADF', help='账套（暂仅 ADF）')
    a = ap.parse_args(argv)
    outs = gen(a.acct)
    print('\n产出 %d 个文件：' % len(outs))
    for o in outs:
        print('  · %s' % os.path.basename(o))
    return 0


if __name__ == '__main__':
    sys.exit(main())
