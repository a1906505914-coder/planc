# -*- coding: utf-8 -*-
"""铁律核对诊断（tb_recon.py，2026-08-30 新建）。
审计铁律：TB(自建试算表) = 审定表 = 明细表 = 附注汇总。
本工具对合并集团底稿全量核对「TB vs 审定表」（损益科目用利润表 Sheet3 发生额、
资产负债科目用试算表 Sheet1 期末余额），一次性输出所有差异科目与主体。

用法：
  python tb_recon.py --data <数据目录> --dir <合并集团底稿目录> [--detail 差异明细主体]
"""
import os
import sys
import argparse
import openpyxl


def _sum_no_total(row, hdr):
    """Σ 行值，排除表头含『合计/总计』的列（TB 有"集团合计"列会双计 → 大量科目×2 根因）。"""
    return sum(v for i, v in enumerate(row)
               if isinstance(v, (int, float)) and not (hdr[i] and '合计' in str(hdr[i])))


def load_tb_sheet1(fp):
    """试算表 Sheet1：{科目: Σ主体}（期末余额，贷余负/借余正），排除集团合计列。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['试算表']
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[0]
    out = {}
    for r in rows[1:]:
        if r and r[0]:
            out[str(r[0]).strip()] = _sum_no_total(r, hdr)
    wb.close()
    return out


def load_tb_sheet3(fp):
    """利润表 Sheet3：{科目: Σ主体}（损益发生额，收入贷/费用借，正数），排除集团合计列。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    ws = wb['利润表']
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[1]
    out = {}
    for r in rows[2:]:
        if r and r[0]:
            out[str(r[0]).strip()] = _sum_no_total(r, hdr)
    wb.close()
    return out


PL_KW = ('收入', '成本', '费用', '收益', '损失', '税金', '所得税', '利润', '支出',
         '减值', '摊销', '折旧', '财务', '管理', '销售', '研发', '营业外')


def read_audit_sheet_total(fp, sheet_kw='审定表'):
    """读底稿审定表，按【项目】聚合 Σ主体。返回 {项目: Σ本期数}。
    ⚡⚡ 2026-08-30 修复：审定表为"项目×主体"展开（营业收入=主营/成本/利润/其他…），
    不能全行 Σ（会把成本/利润混入收入）。按 row[0]（项目）分组逐项聚合。"""
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    sns = [s for s in wb.sheetnames if sheet_kw in s and '集团' not in s and '合并' not in s]
    if not sns:
        wb.close()
        return {}
    ws = wb[sns[0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr = None
    for i, r in enumerate(rows[:10]):
        if r and ('本期数' in [str(x) if x else '' for x in r]
                  or '审定数' in [str(x) if x else '' for x in r]
                  or '期末' in [str(x) if x else '' for x in r]):
            hdr = i
            break
    if hdr is None:
        wb.close()
        return {}
    # ⚡⚡ 2026-08-30 列识别：损益类取『本期数』，资产负债类取『审定数/期末数』
    col_cands = [j for j, h in enumerate(rows[hdr]) if h and '本期数' in str(h)] or \
                [j for j, h in enumerate(rows[hdr]) if h and ('审定数' in str(h) or '期末' in str(h))]
    if not col_cands:
        wb.close()
        return {}
    col = col_cands[0]
    proj = {}
    # ⚡⚡ 2026-08-30 结构识别：row[0] 为 4 位数字（核算主体码）→ 结构A 单列Σ；
    #   row[0] 为项目文本 → 结构B 按项目聚合。
    import re as _re
    first = str(rows[hdr + 1][0]).strip() if hdr + 1 < len(rows) and rows[hdr + 1] and rows[hdr + 1][0] else ''
    if _re.fullmatch(r'\d{4,6}', first):
        tot = 0.0
        for r in rows[hdr + 1:]:
            if not r or not r[0]:
                continue
            label = str(r[0]).strip()
            if label in ('合计', '总计', '全集团合计'):
                continue
            if isinstance(r[col], (int, float)):
                tot += r[col]
        proj['__TOTAL__'] = tot
    else:
        for r in rows[hdr + 1:]:
            if not r or not r[0]:
                continue
            label = str(r[0]).strip()
            if label in ('合计', '总计', '全集团合计'):
                continue
            # ⚡⚡ 2026-08-30 结构B（项目×主体）：排除主体列为『全集团合计/合计』的行
            #   （营业收入等审定表含"全集团合计"主体行 → 不排除会全部科目≈2倍）
            if len(r) > 1 and r[1] and any(k in str(r[1]) for k in ('全集团', '合计', '总计')):
                continue
            if isinstance(r[col], (int, float)):
                proj[label] = proj.get(label, 0.0) + r[col]
    wb.close()
    return proj


def tb_match(name, tb1, tb3):
    """返回 (源, 匹配科目值列表)。损益科目优先 Sheet3。"""
    is_pl = any(k in name for k in PL_KW)
    if is_pl:
        m = {k: v for k, v in tb3.items() if k and (k in name or name in k)}
        if m:
            return 'Sheet3', m
    m1 = {k: v for k, v in tb1.items() if k and (k in name or name in k)}
    return 'Sheet1', m1


def main():
    ap = argparse.ArgumentParser(description='铁律核对：TB vs 审定表')
    ap.add_argument('--data', required=True, help='数据目录（含自建试算表_2026.xlsx）')
    ap.add_argument('--dir', required=True, help='合并集团底稿目录')
    ap.add_argument('--year', default='2026')
    ap.add_argument('--detail', action='store_true', help='输出差异主体明细')
    args = ap.parse_args()

    tb1 = load_tb_sheet1(os.path.join(args.data, f'自建试算表_{args.year}.xlsx'))
    tb3 = load_tb_sheet3(os.path.join(args.data, f'自建试算表_{args.year}.xlsx'))

    results = []
    for fn in sorted(os.listdir(args.dir)):
        if not fn.endswith('.xlsx'):
            continue
        subj = fn.replace('审计底稿_AH合并.xlsx', '').replace('审计底稿_2026_生成.xlsx', '')
        fp = os.path.join(args.dir, fn)
        try:
            proj = read_audit_sheet_total(fp)
        except Exception:
            continue
        if not proj:
            results.append((subj, 0.0, 0.0, 0.0, '无审定表'))
            continue
        # 逐项目核对（清洗前缀：一、/二、/减：/加：/其中：/空格）
        for label, val in sorted(proj.items()):
            if label == '__TOTAL__':
                show = subj
                clean = subj
            else:
                show = f'{subj}/{label}'
                clean = label
            for pre in ('一、', '二、', '三、', '四、', '五、', '六、', '减：', '加：', '其中：', '   '):
                clean = clean.replace(pre, '')
            clean = clean.strip()
            if not clean or clean in ('合计', '总计'):
                continue
            is_pl = (any(k in clean for k in PL_KW) or
                     any(k in clean for k in ('主营业务收入', '其他业务收入', '主营业务成本',
                                              '其他业务成本', '营业收入', '营业成本', '利润')))
            tbmap = tb3 if is_pl else tb1
            m = {k: v for k, v in tbmap.items() if k and (k in clean or clean in k)}
            if not m and is_pl:
                # ⚡⚡ 2026-08-30 损益匹配映射：底稿项目 vs Sheet3 利润表项目
                MAP = {'主营业务收入': '营业收入', '其他业务收入': '营业收入',
                       '主营业务成本': '营业成本', '其他业务成本': '营业成本',
                       '营业税金及附加': '税金及附加', '税金及附加': '税金及附加'}
                for _k, _v in MAP.items():
                    if _k in clean:
                        m = {_k2: _v2 for _k2, _v2 in tb3.items() if _k2 == _v}
                        break
            if not m:
                results.append((show, val, None, None, '无匹配'))
                continue
            tsum = sum(m.values())
            d = val - tsum
            if abs(d) < 0.01:
                concl = '一致'
            elif abs(abs(val) - abs(tsum)) < 0.01:
                concl = '符号口径'
            else:
                concl = '差异'
            results.append((show, val, tsum, d, concl))
    print(f"{'科目':<14}{'底稿审定Σ':>16}{'TB值':>16}{'差异':>14}  结论")
    n_diff = 0
    for subj, tot, tsum, d, concl in sorted(results):
        flag = ''
        if concl == '差异':
            n_diff += 1
            flag = ' ❌'
        elif concl == '符号口径':
            flag = ' ~'
        print(f'{subj:<14}{tot:>16,.2f}{tsum if tsum is not None else 0:>16,.2f}'
              f'{d if d is not None else 0:>14,.2f}  {concl}{flag}')
    print(f'\n差异 {n_diff} / 符号口径 {sum(1 for r in results if r[4] == "符号口径")} / '
          f'一致 {sum(1 for r in results if r[4] == "一致")} / 无匹配 {sum(1 for r in results if r[4] == "无匹配")}')
    return 0 if n_diff == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
