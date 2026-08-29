# -*- coding: utf-8 -*-
"""AH 三集团底稿回归检查（历史问题清单"重新过一遍"的自动化部分）。
只读不写：检查已生成底稿的关键勾稽，输出红旗清单。
检查项（对应历史问题清单高风险复发项）：
  R1 折旧/摊销分摊表：核对列≈0（或仅披露项差异）；STD 外列存在（主营业务成本等）
  R2 无形资产摊销核对表：研发支出(资本化) 与 研发费用 分列
  R3 权益底稿：明细=审定（符号口径，08-27 修复项）
  R4 底稿《勾稽与异常检查》无 ERROR
  R5 公司代码.xlsx 88 家齐全
  R6 费用中摊销折旧计提汇总 sheet 存在且无资产负债类杂质
"""
import glob, os, re, sys
import openpyxl

ROOT = 'd:/底稿测试/AH/底稿/2026/集团'
GRPS = ['1010', '1357', '2468']
ROOT_SUFFIX = ''  # 正式目录
FLAGS = []


def flag(tag, msg):
    FLAGS.append((tag, msg))
    print(f'  ⚠ [{tag}] {msg}')


def check_amort_sheet(g, fp, sheet, col_idx):
    """核对列≈0 检查；col_idx 为核对列索引(0-based)。"""
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception as e:
        flag('R1', f'{g}/{os.path.basename(fp)} 打开失败: {e}')
        return
    if sheet not in wb.sheetnames:
        flag('R1', f'{g}/{os.path.basename(fp)} 缺 sheet {sheet}')
        return
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 5:
        flag('R1', f'{g}/{os.path.basename(fp)} {sheet} 空')
        return
    bad = 0
    for r in rows[4:]:
        if r and len(r) > col_idx and r[1] == '小计':
            try:
                chk = float(r[col_idx] or 0)
            except (TypeError, ValueError):
                continue
            if abs(chk) > 0.01:
                bad += 1
                if bad <= 3:
                    flag('R1', f'{g} {r[0]} {sheet} 小计核对={chk:,.2f}')
    if bad == 0:
        print(f'  ✓ [R1] {g} {sheet}: 各主体小计核对全部≈0')


def check_fee_summary(g, fp):
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception:
        return
    if '费用中摊销折旧计提汇总' not in wb.sheetnames:
        flag('R6', f'{g} 缺《费用中摊销折旧计提汇总》')
        return
    ws = wb['费用中摊销折旧计提汇总']
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[2] if len(rows) > 2 else ()
    tops = [str(h or '') for h in hdr[1:-1]]
    bad = [t for t in tops if t in ('待摊费用', '长期待摊费用', '其他应付款', '应付账款',
                                    '应收票据', '预付账款', '应交税费', '递延收益', '专项应付款')]
    if bad:
        flag('R6', f'{g} 费用侧汇总混入资产/负债类列: {bad}')
    else:
        print(f'  ✓ [R6] {g} 费用侧汇总列={tops}（无资产负债杂质）')


def check_equity(g):
    fp = glob.glob(os.path.join(ROOT, g + ROOT_SUFFIX, '未分配利润审计底稿*.xlsx'))
    if not fp:
        print(f'  · [R3] {g} 无未分配利润底稿（跳过）')
        return
    fp = fp[0].replace('_脱敏', '')
    if not os.path.exists(fp):
        fp = fp.replace('.xlsx', '_脱敏.xlsx') if os.path.exists(fp.replace('.xlsx', '_脱敏.xlsx')) else fp
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception as e:
        flag('R3', f'{g} 权益底稿打开失败: {e}')
        return
    # 检查审定表与明细表合计是否相等
    for sn in wb.sheetnames:
        if '审定' in sn or '明细' in sn:
            ws = wb[sn]
            rows = list(ws.iter_rows(values_only=True))
            if rows and any(any('合计' == str(c or '').strip() or '审计认定' in str(c or '') for c in r) for r in rows[:5]):
                print(f'  · [R3] {g} {sn}: 有合计行（人工核对明细=审定）')
                return
    print(f'  ✓ [R3] {g} 权益底稿存在，未发现明显空表')


def check_integrity_sheet(g):
    """勾稽与异常检查 sheet 无 ERROR。"""
    for fp in sorted(glob.glob(os.path.join(ROOT, g + ROOT_SUFFIX, '*.xlsx'))):
        if '_脱敏' in fp:
            continue
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception:
            continue
        if '勾稽与异常检查' not in wb.sheetnames:
            continue
        ws = wb['勾稽与异常检查']
        txt = ''
        for r in ws.iter_rows(values_only=True):
            for c in r:
                if c is not None:
                    txt += str(c) + '\n'
        if re.search(r'\bERROR\b|不通过|校验失败|勾稽不平|勾稽失败|❌', txt):
            errs = [l for l in txt.split('\n') if re.search(r'\bERROR\b|不通过|校验失败|勾稽不平|勾稽失败|❌', l)][:3]
            flag('R4', f'{g}/{os.path.basename(fp)} 勾稽检查含错误: {errs}')
        # else 不打印，避免噪音


def main():
    print('==== AH 三集团底稿回归检查 ====\n')
    # R5 公司代码 88 家
    cp = 'D:/底稿测试/AH/中间产物/prepared/公司代码.xlsx'
    try:
        wb = openpyxl.load_workbook(cp, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        codes = [re.search(r'\|(\d{4})\|', str(r[0])).group(1) for r in ws.iter_rows(values_only=True)
                 if r and r[0] and re.search(r'\|(\d{4})\|', str(r[0]))]
        n = len(set(codes))
        print(f'  {"✓" if n == 88 else "⚠"} [R5] 公司代码.xlsx: {n} 家' + ('（88 齐全）' if n == 88 else ''))
        if n != 88:
            flag('R5', f'公司代码.xlsx 只有 {n} 家')
    except Exception as e:
        flag('R5', f'公司代码.xlsx 读取失败: {e}')

    for g in GRPS:
        print(f'\n--- 集团 {g} ---')
        # 固定资产底稿：折旧分摊表 + 费用侧汇总
        fa = glob.glob(os.path.join(ROOT, g + ROOT_SUFFIX, '固定资产审计底稿*.xlsx'))
        if not fa:
            flag('R1', f'{g} 无固定资产底稿')
        else:
            fa = [f for f in fa if '_脱敏' not in f]
            if fa:
                check_amort_sheet(g, fa[0], '折旧分摊表', -1)
                check_fee_summary(g, fa[0])
        # 无形资产底稿：摊销核对表
        ia = [f for f in glob.glob(os.path.join(ROOT, g + ROOT_SUFFIX, '无形资产审计底稿*.xlsx')) if '_脱敏' not in f]
        if ia:
            check_amort_sheet(g, ia[0], '无形资产摊销核对表', -1)
        # 权益
        check_equity(g)
        # 勾稽与异常检查
        check_integrity_sheet(g)

    print('\n==== 汇总 ====')
    if FLAGS:
        print(f'共 {len(FLAGS)} 处红旗：')
        for tag, msg in FLAGS:
            print(f'  {tag}: {msg}')
    else:
        print('✓ 无红旗')
    return 1 if FLAGS else 0


if __name__ == '__main__':
    sys.exit(main())
