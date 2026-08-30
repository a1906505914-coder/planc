# -*- coding: utf-8 -*-
"""最终铁律核对：底稿审定表 逐主体 vs 自建试算表（Sheet1 期末余额）。

口径：
- 底稿取『审定数/期末未审数』列（优先审定数；无则本期数——损益类）
- TB Sheet1 逐主体期末余额（qm）
- 逐主体比绝对值（排除贷余为负的符号口径），差异>阈值 记真实差异
- 输出：每科目 真实差异主体数 / 符号口径一致主体数 / 差异额合计
用法：python recon_final.py --data <data_dir> --tb <自建试算表.xlsx> --dir <底稿目录>
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def find_val_col(rows):
    """优先『审定数』列（资产负债表科目期末口径）；无则『期末未审数/本期数』。"""
    for kw in ('审定数', '期末未审数', '本期数', '期末数'):
        for i, r in enumerate(rows[:8]):
            hs = [str(x) if x else '' for x in r]
            if kw in hs:
                return i, hs.index(kw)
    return None, None


# 名称含这些字样但【非】该科目本身（递延所得税-xxx / 处置利得 等），核对 TB 时排除
TB_EXCLUDE = ('递延所得税', '利得', '损失', '处置', '清理', '折旧', '摊销', '减值',
              '坏账', '准备', '收入', '费用', '营业外', '所得税')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', required=True, help='数据目录（含 科目余额表）')
    ap.add_argument('--tb', required=True, help='自建试算表 xlsx 路径')
    ap.add_argument('--dir', required=True, help='底稿目录')
    ap.add_argument('--threshold', type=float, default=1.0)
    args = ap.parse_args()

    import openpyxl
    import sap_adapter as A
    A._DATA_ROOT = args.data

    # TB Sheet1 逐主体
    wb = openpyxl.load_workbook(args.tb, read_only=True, data_only=True)
    ws = wb['试算表']
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[0]
    tb_rows = {}
    for r in rows[1:]:
        if r[0]:
            tb_rows[str(r[0]).strip()] = {
                str(hdr[i]): v for i, v in enumerate(r)
                if i > 0 and hdr[i] and '合计' not in str(hdr[i]) and isinstance(v, (int, float))}
    wb.close()

    out = args.dir
    print(f"{'科目':<14}{'真实差异':>8}{'符号一致':>8}{'差异额合计':>18}  判定")
    n_real = n_ok = 0
    for fn in sorted(os.listdir(out)):
        if not fn.endswith('.xlsx'):
            continue
        subj = fn.split('审计底稿')[0].strip()   # 通用命名（AH合并/XBJ/2025_生成 全适配）
        try:
            wb2 = openpyxl.load_workbook(os.path.join(out, fn), read_only=True, data_only=True)
            sns = [s for s in wb2.sheetnames if '审定表' in s and '集团' not in s and '合并' not in s]
            if not sns:
                wb2.close()
                continue
            rws = list(wb2[sns[0]].iter_rows(values_only=True))
            wb2.close()
            hi, ci = find_val_col(rws)
            if hi is None:
                continue
            # 结构判断：数据行 row[1] 是否为『项目』（原值/累计折旧/减值/净值）→ 结构B（主体×项目）
            # ⚡ 跳过表头后的空行（审定表标题/说明/空行）找首个数据行
            is_struct_b = False
            for _r0 in rws[hi + 1:]:
                if _r0 and _r0[0] and str(_r0[0]).strip() not in ('合计', '总计', '全集团合计', '集团加计'):
                    if _r0[1] and any(k in str(_r0[1]) for k in ('原值', '累计折旧', '减值准备', '净值', '账面')):
                        is_struct_b = True
                    break
            wp = {}
            for r in rws[hi + 1:]:
                if not r or not r[0] or str(r[0]).strip() in ('合计', '总计', '全集团合计', '集团加计'):
                    continue
                e = str(r[0]).strip()
                if is_struct_b:
                    # 只取『原值』项目行
                    if r[1] and '原值' in str(r[1]) and len(r) > ci and isinstance(r[ci], (int, float)):
                        wp[e] = wp.get(e, 0.0) + r[ci]
                elif len(r) > ci and isinstance(r[ci], (int, float)):
                    wp[e] = r[ci]
            if not wp:
                continue
            # TB 匹配：精确名优先；否则 名称包含（排除 递延所得税/处置利得 从属科目、
            #   及『其他应付账款』等含前缀但不同科目标签的误配）
            tbk = None
            if subj in tb_rows:
                tbk = subj
            else:
                for k in tb_rows:
                    if subj and subj in k and len(k) <= len(subj) + 4 \
                            and not any(x in k for x in TB_EXCLUDE) \
                            and not (k.startswith('其他') and not subj.startswith('其他')):
                        tbk = k
                        break
            if tbk is None:
                for k in tb_rows:
                    if k and k in subj and len(subj) <= len(k) + 4:
                        tbk = k
                        break
            if tbk is None:
                continue
            tbm = dict(tb_rows[tbk])
            # ⚡⚡ 2026-08-30 归并（审计列报口径）：其他应收/应付 需并入 应收/应付股利利息（TB 分开）
            MERGE = {'其他应收款': ['应收股利', '应收利息'],
                     '其他应付款': ['应付股利', '应付利息']}
            for _mk in MERGE.get(subj, []):
                if _mk in tb_rows:
                    for _e, _v in tb_rows[_mk].items():
                        tbm[_e] = tbm.get(_e, 0.0) + _v
            real = []
            sign = 0
            diff_tot = 0.0
            for e, wv in wp.items():
                tv = tbm.get(e, 0.0)
                if abs(abs(wv) - abs(tv)) > args.threshold:
                    real.append((e, wv, tv))
                    diff_tot += abs(abs(wv) - abs(tv))
                else:
                    sign += 1
            mark = '❌' if real else '✓'
            if real:
                n_real += 1
            else:
                n_ok += 1
            print(f'{subj:<14}{len(real):>8}{sign:>8}{diff_tot:>18,.0f}  {mark}')
            for e, wv, tv in real[:3]:
                print(f'    {e}: 底稿={wv:,.2f} TB={tv:,.2f}')
        except Exception:
            pass
    print(f'\n一致/符号口径 {n_ok} 个 / 真实差异 {n_real} 个')
    return 0 if n_real == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
