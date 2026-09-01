"""附注汇总 vs 合并审定表核对工具 v2（2026-09-01）
用法：python footnote_recon.py <底稿目录> <附注汇总xlsx> <合并工作底稿xlsx>

三项目标：
  ① 科目名映射：附注 sheet 名（银行存款/职工薪酬/营业收入…）↔ 合并审定表科目名（货币资金/应付职工薪酬/一、营业收入…）
  ② 符号统一：合并审定表负债/权益为负数，附注披露为正数 → 按绝对值比较
  ③ 4 种附注格式统一提取：
     A 审定数格式（长期借款/资本公积/货币资金…）：取最后一个『审定数』行（期末块）
     B 账龄格式（应收/应付/预收…）：取『审定数』列中 账面价值合计/账面价值/合 计 行
     C 明细格式（固定资产/存货/营业收入…）：取『合并报表数』/『集团合计数』行；存货取期末审定账面价值块；营业收入取营业总收入行
     D 二级科目格式（管理费用/销售费用…）：取最后『集团合计数』行
输出差异清单（绝对值口径），供审计判断参考。"""
import sys
import os
import openpyxl

# 附注 sheet 名 → 合并审定表科目名
NOTE_TO_MERGE = {
    '银行存款': '货币资金',
    '职工薪酬': '应付职工薪酬',
    '实收资本(或股本)': '实收资本',
    '营业收入': '一、营业收入',
    '其他收益': '加：其他收益',
    '营业外支出': '减：营业外支出',
    '所得税费用': '减：所得税费用',
}
# 附注有、合并审定表横展刻意不列（非披露科目/汇总行），不做差异报告
MERGE_SUMMARY_ROWS = {
    '资产合计', '负债合计', '所有者权益合计', '收入合计', '成本费用合计',
    '净利润', '未结转损益', '其他（未匹配）', '其他流动资产', '长期应收款',
}


def _nums(r):
    return [x for x in r if isinstance(x, (int, float)) and abs(x) > 0.005]


def _sum_nums(r):
    return sum(_nums(r))


def _pick_val(r):
    """每主体一列格式（A/C/D）取『全集团合计』= 行内最后一个数值。"""
    ns = _nums(r)
    return ns[-1] if ns else 0.0


def _label_rows(rows, start, key):
    return [i for i in range(start, len(rows))
            if rows[i] and rows[i][0] and key in str(rows[i][0])]


def _hdr_info(rows):
    for i, r in enumerate(rows[:6]):
        if not r or not r[0]:
            continue
        s = str(r[0])
        if '项目（审定数）' in s:
            return 'A', i
        if '科目筛选列' in s:
            return 'B', i
        if '二级科目' in s:
            return 'D', i
        if s.strip() == '项目' or '项目' in s:
            return 'C', i
    return None, None


def extract_note(rows):
    typ, hi = _hdr_info(rows)
    if typ is None:
        return 0.0

    # ---- A 审定数格式：取『期末数』块对应的审定数 ----
    if typ == 'A':
        # 1) 期末审定数标签行（货币资金：期末审定数→明细行）
        ej = _label_rows(rows, hi + 1, '期末审定数')
        if ej:
            tot = 0.0
            stop = ('期初', '本期', '审计', '滚动', '注：', '口径', '勾稽',
                    '期末数', '本期计提', '附注', '项目', '集团合计数', '合并报表数')
            for k in range(ej[-1] + 1, len(rows)):
                r = rows[k]
                if not r or not r[0]:
                    continue
                t = str(r[0]).strip()
                if any(x in t for x in stop):
                    break
                tot += _pick_val(r)
            if abs(tot) > 0.005:
                return tot
        # 2) 『期末数』块标签 → 其后首个『审定数』行
        qm = _label_rows(rows, hi + 1, '期末数')
        if qm:
            for k in range(qm[-1] + 1, len(rows)):
                if rows[k] and rows[k][0] and '审定数' in str(rows[k][0]):
                    return _sum_nums(rows[k])
        # 3) 回退：最后一个『审定数』行
        aud = _label_rows(rows, hi + 1, '审定数')
        if aud:
            i = aud[-1]
            s = _sum_nums(rows[i])
            if abs(s) > 0.005:
                return s
            tot = 0.0
            stop = ('期初', '本期', '审计', '滚动', '注：', '口径', '勾稽',
                    '期末数', '本期计提', '附注', '项目', '集团合计数', '合并报表数')
            for k in range(i + 1, len(rows)):
                r = rows[k]
                if not r or not r[0]:
                    continue
                t = str(r[0]).strip()
                if any(x in t for x in stop):
                    break
                tot += _pick_val(r)
            return tot
        for key in ('合并报表数', '集团合计数'):
            idxs = _label_rows(rows, hi + 1, key)
            if idxs:
                return _sum_nums(rows[idxs[-1]])
        return 0.0

    # ---- B 账龄格式：审定数列，优先 账面价值合计→账面价值→合 计 ----
    if typ == 'B':
        hdr = rows[hi]
        j_aud = next((j for j, v in enumerate(hdr) if v and '审定数' in str(v)), None)
        if j_aud is None:
            return 0.0
        for pat in ('账面价值合计', '账面价值', '合  计', '合计'):
            for r in rows[hi + 1:]:
                label = ' '.join(str(x) for x in (r[0], r[1] if len(r) > 1 else '') if x)
                if (pat in label and len(r) > j_aud
                        and isinstance(r[j_aud], (int, float)) and abs(r[j_aud]) > 0.005):
                    return r[j_aud]
        for r in reversed(rows[hi + 1:]):
            if r and len(r) > j_aud and isinstance(r[j_aud], (int, float)) and abs(r[j_aud]) > 0.005:
                return r[j_aud]
        return 0.0

    # ---- D 二级科目格式：最后『集团合计数』行（每主体一列，无合计列 → sum）----
    if typ == 'D':
        idxs = _label_rows(rows, hi + 1, '集团合计数')
        if idxs:
            return _sum_nums(rows[idxs[-1]])
        for r in reversed(rows[hi + 1:]):
            if r and r[0] and '(合计)' in str(r[0]):
                return _sum_nums(r)
        return 0.0

    # ---- C 明细格式 ----
    # 存货：期末审定账面价值块（跳过合 计 行避免双计）
    idxs = _label_rows(rows, hi + 1, '期末审定账面价值')
    if idxs:
        tot = 0.0
        for k in range(idxs[-1] + 1, len(rows)):
            r = rows[k]
            if not r or not r[0]:
                continue
            t = str(r[0]).strip()
            if '跌价准备变动表' in t or '十一' in t or '勾稽' in t or '滚动' in t:
                break
            if '合' in t and '计' in t:
                continue
            tot += _pick_val(r)
        if abs(tot) > 0.005:
            return tot
    # 营业收入：营业总收入行（= 主营 + 其他业务收入）
    idxs = _label_rows(rows, hi + 1, '营业总收入')
    if idxs:
        return _pick_val(rows[idxs[-1]])
    # 一般：最后『合并报表数』/『集团合计数』
    for key in ('合并报表数', '集团合计数'):
        idxs = _label_rows(rows, hi + 1, key)
        if idxs:
            return _pick_val(rows[idxs[-1]])
    return 0.0


def read_merged(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    rows = list(wb['合并审定表（横展）'].iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows[:4]) if r and r[0] and str(r[0]).strip() == '项目'), None)
    if hi is None:
        wb.close()
        return {}
    hdr = rows[hi]
    c_m = next((j for j, h in enumerate(hdr) if h and '合并报表数' in str(h)), None)
    if c_m is None:
        c_m = len(hdr) - 1
    out = {}
    for r in rows[hi + 1:]:
        if r and r[0]:
            v = r[c_m] if len(r) > c_m and isinstance(r[c_m], (int, float)) else 0
            out[str(r[0]).strip()] = v
    wb.close()
    return out


def read_note(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    out = {}
    for sn in wb.sheetnames:
        rows = list(wb[sn].iter_rows(values_only=True))
        out[sn] = extract_note(rows)
    wb.close()
    return out


def main(src_dir, note_fp, merge_fp):
    merged = read_merged(merge_fp)
    notes = read_note(note_fp)
    ok, diff, note_only = 0, [], []
    for sn, note in notes.items():
        m_key = NOTE_TO_MERGE.get(sn, sn)
        m = merged.get(m_key, 0)
        # ⚡⚡ 2026-09-01 长期股权投资净额口径：附注按被投资单位披露【原值】+单独披露
        #   『长期股权投资减值准备』；合并审定表（横展）按资产负债表口径列示【净值】=
        #   原值-减值（AH 差异 157.3M=减值 1512）。核对时附注原值扣减值后对比横展净值。
        if sn == '长期股权投资' and '长期股权投资减值准备' in notes:
            note -= notes['长期股权投资减值准备']
        if m == 0 and note == 0:
            continue
        if m == 0:
            note_only.append((sn, note))
            continue
        if abs(abs(note) - abs(m)) > 1:
            diff.append((sn, note, m, m_key))
        else:
            ok += 1
    print(f'=== 附注汇总 vs 合并审定表（{src_dir}）===')
    print(f'一致 {ok} 个 / 差异 {len(diff)} 个 / 附注有而横展无 {len(note_only)} 个')
    print('\n[差异清单]（绝对值口径）')
    for sn, n, m, mk in sorted(diff, key=lambda x: -abs(abs(x[1]) - abs(x[2]))):
        print(f'  {sn[:12]:<14} 附注={n:>18,.0f}  合并({mk[:10]})={m:>18,.0f}  |差|={abs(abs(n)-abs(m)):,.0f}')
    if note_only:
        print('\n[仅附注披露、合并审定表未列]')
        for sn, n in note_only:
            print(f'  {sn[:14]:<16} 附注={n:,.0f}')
    return 0


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print('用法: python footnote_recon.py <底稿目录> <附注汇总xlsx> <合并工作底稿xlsx>')
        sys.exit(1)
    main(sys.argv[1], sys.argv[2], sys.argv[3])
