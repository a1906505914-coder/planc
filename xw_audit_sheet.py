# -*- coding: utf-8 -*-
"""#876 P2 阶段3（第1步）：xw 版审定表渲染（资产负债科目 · 按核算主体）。

对标 audit_common.add_audit_summary_sheets 的 by_entity + 资产负债分支：
  列 = 核算主体 | 年初数 | 期末未审数 | 审计调整数 | 审定数 | 与科目余额表勾稽
  - 每主体按 codes（前缀+末级叶子）或 names（精确）从 TB 聚合 qc/qm；
  - 方向：is_credit（负债/权益贷方正显示）取 abs；keep_sign 保持带符号；
  - 审定数 = 期末未审（调整数留空待手填）；勾稽列标注『与科目余额表核对一致』；
  - 合计行 = 各主体累计。

不修改 openpyxl 版（AH 正式流程零影响）；仅服务 CA_XW 模式。
"""
import xw_render as X


def _aggregate_bs(tb_full, entities, eff_years, codes, names, is_credit, keep_sign=False):
    """按主体聚合 期初/期末（含方向处理）。返回 [(ent, qc_disp, qm_disp)]。"""
    out = []
    for ent in sorted(entities):
        qc = qm = 0.0
        for (e, c, n, yy), v in tb_full.items():
            if e != ent or yy not in eff_years:
                continue
            if names:
                if not any(nm and (str(n) == nm or str(n).startswith(nm + '-') or nm in str(n))
                           for nm in names):
                    continue
                # ⚡⚡ 2026-08-30 修复：排除坏账/备抵科目（应收 18.27亿→18.21亿 差 592万 根因）
                #   names 匹配 '应收账款' in '坏账准备-其他应收账款坏帐准备'=True 误命中 123102；
                #   且繁简不一致（123101'应收帐款'帐≠账 不命中、123102'应收账款'命中）→ 只扣 592万。
                #   坏账准备由坏账准备 sheet 单独列示，审定表必须取原值（与 openpyxl _amt_raw 跳过
                #   1231/1471/1602/1703 备抵码一致）。
                # ⚡⚡ 2026-08-30 修复②：同类从属科目误吸收——
                #   · '递延所得税资产-合同资产减值准备'(1811) 被合同资产误算入（1050 差 628K）
                #   · '其他应付账款-其他应付账款调整'(2241) 被应付账款误算入（1100 差 40K）
                #   排除：名称含 递延所得税/跌价，或 code 1811/2901（递延所得税），或
                #   名称以『其他』开头但非本组（其他应收/其他应付属独立组）。
                _nm = str(n)
                if ('坏账' in _nm) or ('坏帐' in _nm) or ('递延所得税' in _nm) or ('跌价' in _nm) \
                        or str(c).startswith(('1231', '1471', '1602', '1703', '1811', '2901')) \
                        or (_nm.startswith('其他') and not any(nm and _nm.startswith(nm) for nm in names)):
                    continue
            elif codes:
                if not any(str(c).startswith(p) for p in codes):
                    continue
                # 末级叶子过滤：本行是父级（有同族后代）→ 剔除，防父+子双计
                _hit = {str(cc) for (ee, cc, _nn, _yy), _v in tb_full.items()
                        if (yy is None or _yy == yy) and ee == ent
                        and any(str(cc).startswith(p) for p in codes)}
                if any(str(c) != cc2 and cc2.startswith(str(c)) for cc2 in _hit):
                    continue
            else:
                continue
            qc += v['qc']
            qm += v['qm']
        if keep_sign:
            d_qc, d_qm = round(qc, 2), round(qm, 2)
        elif is_credit:
            d_qc, d_qm = round(abs(qc), 2), round(abs(qm), 2)
        else:
            d_qc = round(-qc if qc < 0 else qc, 2)
            d_qm = round(-qm if qm < 0 else qm, 2)
        out.append((ent, d_qc, d_qm))
    return out


def render_audit_summary_xw(wb, title, subj_name, tb_full, entities, target_year=None,
                            codes=None, names=None, is_credit=False, keep_sign=False):
    """xw 版审定表（按核算主体）。返回生成的 sheet 名（无数据返回 None）。"""
    years = sorted({y for (e, c, n, y) in tb_full})
    eff_years = [target_year] if target_year else years
    rows = _aggregate_bs(tb_full, entities, eff_years, codes, names, is_credit, keep_sign)
    # 全部 0 也生成（审定表是审计交付结构，不因无数据省略）
    ws = wb.add_worksheet(title)
    X.title(wb, ws, f'{subj_name} 审定表（按核算主体）', 6)
    r = X.dual_header(wb, ws, 1, [('', 6)],
                      ['核算主体', '年初数', '期末未审数', '审计调整数', '审定数', '与科目余额表勾稽'],
                      widths=[16, 18, 18, 14, 18, 28], freeze_rows=3)
    tot_qc = tot_qm = 0.0
    for ent, d_qc, d_qm in rows:
        X.row(wb, ws, r, [ent, d_qc, d_qm, None, d_qm, '与科目余额表核对一致'],
              money_cols=(1, 2, 3, 4), center_cols=(0,))
        tot_qc += d_qc
        tot_qm += d_qm
        r += 1
    X.row(wb, ws, r, ['合计', round(tot_qc, 2), round(tot_qm, 2), None, round(tot_qm, 2), None],
          money_cols=(1, 2, 3, 4), center_cols=(0,), fills={j: 'FCE4D6' for j in range(6)})
    ws.freeze_panes(3, 0)
    return title
