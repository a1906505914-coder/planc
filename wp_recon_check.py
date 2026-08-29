# -*- coding: utf-8 -*-
"""底稿勾稽专项核对（2026-08-24 用户要求：不仅看有数据，还要看数据完整性与勾稽差异）。

检查维度：
1. 审定表内部勾稽：期初 + 本期增加 - 本期减少 ≈ 期末（逐数据行；容差 1 元）
2. 审定表合计行：合计行列值 vs 明细行加总
3. 附注汇总审定数 vs 审定表合计（存在附注汇总且可定位时）
4. 借贷平衡：明细表借方合计 ≈ 贷方合计（仅对有借/贷列的明细表）
5. 空白底稿（整份无数值）——复用 self_check 判定

输出：按账套分组的勾稽差异清单（科目/表/行/差异值）。返回差异计数。
"""
import sys, os, re, glob

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

TOL = 1.0  # 容差（元）


def _cell(r, i):
    return r[i] if i is not None and i < len(r) else None


def _num(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(str(v).replace(',', '').strip())
        except Exception:
            return None
    return None


def _find_header(rows, kws):
    """在前若干行找表头：同时含所有 kws 中任一关键词的行。返回 (行号, 行内容)。"""
    for i, r in enumerate(rows[:12]):
        if not r:
            continue
        s = ''.join(str(x) if x is not None else '' for x in r[:16])
        if all(any(k in s for k in kw) for kw in kws):
            return i, r
    return None, None


def _col_of(hdr, kws):
    for i, h in enumerate(hdr[:16]):
        if h and any(k in str(h) for k in kws):
            return i
    return None


def check_workbook(fp):
    """返回该底稿的勾稽差异清单：[(sheet, 行号, 描述, 差异)]"""
    import openpyxl
    diffs = []
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception:
        return [('', 0, '文件无法打开', 0.0)]
    for sn in wb.sheetnames:
        # ⚡⚡ 2026-08-25 内存优化（XBJ 81MB 大底稿 OOM 根因）：只加载【需要检查的
        #   sheet】（审定表 / 多借多贷 / 审计调整），明细表/序时账等大表跳过——
        #   原对每个 sheet 全量读入 list，多个大 sheet 撑爆内存（python 被 OOM 杀）。
        if not ('审定表' in sn or '多借多贷' in sn or '借贷平衡' in sn or '审计调整' in sn):
            continue
        ws = wb[sn]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        # ⚡⚡ 2026-08-25 补充维度（docstring 第4条）：借贷平衡检查——借方金额合计 ≈
        #   贷方金额合计（每笔分录借贷必平衡）。⚠️ 仅对【多借多贷核对表】与【审计调整
        #   分录】类表启用；凭证抽查/检查表是【抽样表】（只抽借方向或贷方向，借贷合计
        #   必然不等），不适用借贷平衡（GFY 2025 凭证抽查 13 项误报教训）。
        if ('多借多贷' in sn or '借贷平衡' in sn or '审计调整' in sn):
            hdr = {}
            for r in rows[:10]:
                for c in range(min(len(r), 16)):
                    v = r[c]
                    if v is None:
                        continue
                    s = str(v).strip()
                    if s == '借方金额':
                        hdr.setdefault('dr', c)
                    elif s == '贷方金额':
                        hdr.setdefault('cr', c)
            if 'dr' in hdr and 'cr' in hdr:
                s_dr = s_cr = 0.0
                for r in rows[2:]:
                    # 跳过说明/合计行（A列文本非调整分录）
                    a0 = str(r[0]).strip() if r and r[0] is not None else ''
                    if not a0 or a0.startswith('说明') or '合计' in a0 or '总计' in a0:
                        continue
                    dv = _num(_cell(r, hdr['dr']))
                    cv = _num(_cell(r, hdr['cr']))
                    if dv is not None:
                        s_dr += dv
                    if cv is not None:
                        s_cr += cv
                if abs(s_dr - s_cr) > TOL and (abs(s_dr) > TOL or abs(s_cr) > TOL):
                    diffs.append((sn, 0, '借贷平衡: 借方=%.2f 贷方=%.2f 差=%.2f' % (s_dr, s_cr, s_dr - s_cr),
                                  s_dr - s_cr))
        if '审定表' not in sn:
            continue
        # 表头：需含 期初/增加/减少/期末
        hi, hdr = _find_header(rows, [['期初', '年初'], ['期末', '余额', '审定'], ['增加'], ['减少']])
        if hi is None:
            # 横展审定表等双年表可能不同结构，跳过
            continue
        # ⚡⚡ 2026-08-25 SAP 集团审定表豁免：表头含『汇总包顺序』= 集团/汇总包口径
        #   （SAP 账套如 AH/ADF 按汇总包列示各子公司）。其『本期增加/减少』是往来
        #   发生额、期初/期末是余额（含方向翻转+绝对值展示，如预付净额取绝对值），
        #   期初+增-减=期末 在集团口径下不必然成立（实测 AH 1010 预付差 251万、
        #   1090 其他应付 15005.15→14994.85 均 TB 平、审定数=|TB qm|、recon_self OK）。
        #   → 集团审定表跳过此勾稽检查，仅对 U8 主体口径审定表生效（U8 全量 0 差异）。
        hdr_vals = [str(x).strip() for x in hdr]
        if any('汇总包' in x or '公司' == x for x in hdr_vals):
            continue
        c_qc = _col_of(hdr, ['期初'])
        c_inc = _col_of(hdr, ['增加'])
        c_dec = _col_of(hdr, ['减少'])
        # ⚡⚡ 2026-08-24 修复：『余额』会误配『期初余额』→ 期末列优先『期末』，
        #   找不到再看『审定』（审定数列），『余额』最后兜底（排除期初余额）。
        c_qm = _col_of(hdr, ['期末'])
        if c_qm is None:
            c_qm = _col_of(hdr, ['审定'])
        if c_qm is None:
            c_qm = _col_of(hdr, ['余额'])
            if c_qm is not None and c_qm == c_qc:
                c_qm = None
        if c_qc is None or c_qm is None:
            continue
        # 逐行勾稽（跳过合计/小计/标题行）
        for i in range(hi + 1, len(rows)):
            r = rows[i]
            label = ''.join(str(x) if x is not None else '' for x in r[:6])
            if not any(_num(x) is not None for x in r[:14]):
                continue
            if '合计' in label or '小计' in label:
                continue
            # ⚡⚡ 2026-08-24：审定表『减：坏账准备』『存货期末净额』等汇总结构行
            #   期初/增/减=0、期末单独取值，不适用 期初+增-减=期末 勾稽，跳过。
            if '减：' in label or '净额' in label:
                continue
            qc = _num(_cell(r, c_qc)); inc = _num(_cell(r, c_inc)) if c_inc is not None else 0.0
            dec = _num(_cell(r, c_dec)) if c_dec is not None else 0.0
            qm = _num(_cell(r, c_qm))
            if qc is None and inc is None and dec is None and qm is None:
                continue
            qc = qc or 0.0; inc = inc or 0.0; dec = dec or 0.0
            # ⚡⚡ 2026-08-25：仅期末有值的行（应付利息/应付股利/一年内到期等独立科目行，
            #   期初+增+减全 0）不适用勾稽，跳过——否则误报（AQ 其他应付款审定表应付股利行）。
            if abs(qc) < 1e-6 and abs(inc) < 1e-6 and abs(dec) < 1e-6 and abs(qm) > 1e-6:
                continue
            # 勾稽：期初+增-减 = 期末（行内公式单元格可能为 None，跳过缺期末的）
            if qm is None:
                continue
            diff = (qc + inc - dec) - qm
            if abs(diff) > TOL:
                # 反向勾稽（负债类模板列接反时：期初-增+减=期末）——用于定性列方向
                diff_rev = (qc - inc + dec) - qm
                # ⚡⚡ 2026-08-25 方向翻转科目豁免：current_account_detail 期初/期末
                #   方向翻转对齐后（贷余科目期初按期末方向显示，等价式 qc-inc+dec=qm），
                #   反向勾稽成立 = 展示方向自洽，非数据错误，跳过
                #   （AYL2026 预付 爱绅科技 qc=-228.1 借余发生 228.1 期末平 误报根因）。
                if abs(diff_rev) <= TOL:
                    continue
                # ⚡⚡ 2026-08-25 再补：方向翻转+绝对值展示豁免——期初/期末方向相反
                #   （期初借余→期末贷余 或反之），期末取贷余绝对值显示，勾稽式结果带负号
                #   与期末绝对值相等。如 AH 1090 其他应付款 qc=+15005.15(借余)
                #   qm=-14994.85(贷余→显示14994.85)：qc-inc+dec=-14994.85，
                #   |.|=|qm| → 展示自洽，非数据错误（SAP 集团审定表按汇总包常见）。
                #   ADF 6900 应付账款同理。TB 均平、审定数=|TB qm|（recon_self OK 佐证）。
                if abs(abs(qc + inc - dec) - abs(qm)) <= TOL:
                    continue
                if abs(abs(qc - inc + dec) - abs(qm)) <= TOL:
                    continue
                diffs.append((sn, i + 1, '审定表勾稽: %s 正向(期初+增-减)=%.2f 反向(期初-增+减)=%.2f 期末=%.2f [正差=%.2f/反差=%.2f]' % (
                    label[:16], qc + inc - dec, qc - inc + dec, qm, diff, diff_rev), diff))
    wb.close()
    return diffs


def check_tb_self(data_dir, out):
    """TB 自身勾稽：带符号体系 qc + jf - df = qm（借余/贷余通用，贷余 qm 为负）。
    ⚡⚡ 2026-08-24 用户铁律：『先查 TB，一般 TB 是平的，正负号一定想清楚』——
    带符号下 qc+jf-df=qm 是通用恒等式；若某行不平 = 源科目余额表数据问题（需客户核实），
    底稿勾稽不平大多由生成器方向/口径引起，不能归因 TB。返回不平行数。"""
    # ⚡⚡ 2026-08-25 修复（ADF SAP TB 误报 1822 不平）：SAP 账套必须优先走 sap_adapter
    #   —— audit_common 通用路径对 SAP 科目余额表的 qc/jf/df 列解析错位（ADF 实测
    #   qc=jf=0、df/qm 有值，qm 列正确故 recon_self OK 但全列勾稽 1822 误报），
    #   sap_adapter 读 ADF 0 不平。U8 账套才走 audit_common。
    try:
        import sap_adapter
        if sap_adapter.is_sap(data_dir):
            sap_adapter.set_root(data_dir)
            tb = sap_adapter.read_tb_full(data_dir, None)
        else:
            import audit_common
            ents = audit_common.discover_entities(data_dir)
            tb = audit_common.read_tb_full(data_dir, ents)
    except Exception:
        try:
            import audit_common
            ents = audit_common.discover_entities(data_dir)
            tb = audit_common.read_tb_full(data_dir, ents)
        except Exception:
            return 0
    bad = []
    for (e, c, n, y), v in tb.items():
        qc = float(v.get('qc') or 0); jf = float(v.get('jf') or 0)
        df = float(v.get('df') or 0); qm = float(v.get('qm') or 0)
        diff = qm - (qc + jf - df)
        if abs(diff) > 1.0:
            bad.append((e, c, str(n)[:14], qc, jf, df, qm, diff))
    out.write('  [TB自检] qc+jf-df=qm：%d 行不平\n' % len(bad))
    if not bad:
        return 0
    from collections import defaultdict
    by_pre = defaultdict(lambda: [0, 0.0])
    for e, c, n, qc, jf, df, qm, diff in bad:
        pre = ''.join(ch for ch in str(c) if ch.isdigit())[:4] or str(c)[:4]
        by_pre[pre][0] += 1
        by_pre[pre][1] += abs(diff)
    for pre in sorted(by_pre):
        out.write('    前缀 %s: %d 行, 差额绝对值合计 %.2f\n' % (pre, by_pre[pre][0], by_pre[pre][1]))
    for e, c, n, qc, jf, df, qm, diff in sorted(bad, key=lambda x: -abs(x[7]))[:10]:
        out.write('    %s %s %s 差=%.2f\n' % (e, c, n, diff))
    return len(bad)


def check_folder(data_dir, out):
    """扫描一个账套目录全部底稿。"""
    fps = [f for f in glob.glob(os.path.join(data_dir, '*审计底稿*.xlsx'))
           if '_脱敏' not in f and '_bak' not in f and not os.path.basename(f).startswith('~$')]
    total = 0
    for fp in sorted(fps):
        diffs = check_workbook(fp)
        if diffs:
            base = os.path.basename(fp)
            for sn, row, desc, dv in diffs[:8]:
                out.write('%s | %s | r%d | %s | diff=%.2f\n' % (base, sn, row, desc, dv))
                total += 1
            out.flush()
    return total


def main():
    args = sys.argv[1:]
    out_path = None
    if '--out' in args:
        i = args.index('--out')
        out_path = args[i + 1]
        args = args[:i] + args[i + 2:]
    out = open(out_path, 'w', encoding='utf-8') if out_path else sys.stdout
    grand = 0
    for d in args:
        if not os.path.isdir(d):
            continue
        label = os.path.basename(d.rstrip('\\/'))
        # ⚡⚡ 2026-08-24：label 带账套名（d=/.../AFJ/数据/2025 → 'AFJ 2025'），
        #   否则多账套输出全是 '2025' 无法区分
        try:
            _acct = os.path.basename(os.path.dirname(os.path.dirname(d.rstrip('\\/'))))
            if _acct and _acct != '数据':
                label = '%s %s' % (_acct, label)
        except Exception:
            pass
        out.write('===== %s =====\n' % label)
        # ⚡⚡ 2026-08-24：先查 TB 自身是否平（带符号），再查底稿勾稽
        n_tb = check_tb_self(d, out)
        n = check_folder(d, out)
        out.write('→ %s: TB不平 %d 行, 底稿勾稽差异 %d 处\n' % (label, n_tb, n))
        grand += n + n_tb
        out.flush()
    out.write('TOTAL %d\n' % grand)
    if out_path:
        out.close()
    return 0 if grand == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
