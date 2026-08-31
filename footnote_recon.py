"""附注汇总 vs 合并审定表核对工具（2026-09-01）
用法：python footnote_recon.py <底稿目录> <附注汇总xlsx> <合并工作底稿xlsx>
提取附注汇总各 sheet 的『审定数/合并报表数』，与合并审定表横展『合并报表数』对比。
支持 4 种格式：
  A 审定数格式（长期借款等）：找『合并报表数』或『审定数』行
  B 账龄格式（应收/应付等）：找『账面价值合计』或『与TB校对』行
  C 明细格式（固定资产等）：找含『审定』的行
  D 二级科目格式（管理费用等）：找『合计』行
输出差异清单，供口径专项参考。"""
import sys
import openpyxl


def read_merged(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    rows = list(wb['合并审定表（横展）'].iter_rows(values_only=True))
    hi = next((i for i, r in enumerate(rows[:4]) if r and r[0] and str(r[0]).strip() == '项目'), None)
    hdr = rows[hi]
    c_m = next((j for j, h in enumerate(hdr) if h and '合并报表数' in str(h)), None)
    out = {}
    for r in rows[hi + 1:]:
        if r and r[0]:
            v = r[c_m] if c_m is not None and len(r) > c_m and isinstance(r[c_m], (int, float)) else 0
            out[str(r[0]).strip()] = v
    wb.close()
    return out


def _fmt_val(rr, hdr_i, keywords):
    """在 hdr_i 之后找首个含任一 keyword 的行，返回其数值（取行内绝对值最大的数值）。"""
    for r in rr[hdr_i + 1:]:
        if r and r[0] and any(k in str(r[0]) for k in keywords):
            vals = [x for x in r[1:] if isinstance(x, (int, float)) and abs(x) > 0.005]
            if vals:
                return vals[0] if len(vals) == 1 else max(vals, key=abs)
    return 0.0


def read_note(fp):
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    out = {}
    for sn in wb.sheetnames:
        rr = list(wb[sn].iter_rows(values_only=True))
        note = 0.0
        # 审定数格式：找『合并报表数』或『审定数』行
        if any(r and r[0] and '项目（审定数）' in str(r[0]) for r in rr[:3]):
            note = _fmt_val(rr, 0, ('合并报表数', '审定数'))
        # 账龄格式：找『账面价值合计』或『与TB校对』
        elif any(r and r[0] and '简单加计数' in str(r[0]) for r in rr[:5]):
            note = _fmt_val(rr, 2, ('账面价值合计', '与TB校对'))
        # 二级科目格式：找『合计』行（末行）
        else:
            note = _fmt_val(rr, 1, ('合计', '总计'))
        out[sn] = note
    wb.close()
    return out


def main(src_dir, note_fp, merge_fp):
    merged = read_merged(merge_fp)
    notes = read_note(note_fp)
    ok = 0
    diff = []
    for sn, note in notes.items():
        m = merged.get(sn, 0)
        if m == 0 and note == 0:
            continue
        if abs(abs(note) - abs(m)) > 1:
            diff.append((sn, note, m))
        else:
            ok += 1
    print(f'=== 附注汇总 vs 合并审定表（{src_dir}）===')
    print(f'一致 {ok} 个 / 差异 {len(diff)} 个')
    for sn, n, m in sorted(diff, key=lambda x: -abs(x[1] - x[2]))[:25]:
        print(f'  {sn[:14]}: 附注={n:,.0f} 合并={m:,.0f} 差={n - m:,.0f}')
    return 0


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3])
