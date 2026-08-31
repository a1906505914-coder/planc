# -*- coding: utf-8 -*-
"""阶段二 2.3 单体/合并自动分表 —— 合并工作底稿汇总模块（2026-08-11）

机制（集团模式多主体专属；单体模式跳过）：
  1) 遍历 subjects_registry.REGISTRY 全部科目
  2) 读已生成的 <科目>审计底稿_<年>_生成.xlsx，用 tb_recon_detail._sheet_ent_values
     提取【逐主体审定数】（与核对程序同源）
  3) 生成合并工作底稿：
     · 横展审定表（科目×主体 + 上年 + 全集团合计 + 抵消借/贷 + 合并报表数）——
       复用 audit_shell.build_merge_wide_sheet（阶段一 1.1 已实现的合并4列强化）
     · 内部交易抵消核对：往来类科目自动附（current_account 已内置），此处汇总表
       提示跨主体候选（同往来单位出现于 ≥2 主体）
     · 合并调整分录：空模板 sheet（待审计人员填）
  4) 单体模式（≤1 主体）不生成。

用法：python audit_merge.py <账套目录> [输出目录]
输出：<输出目录>/合并工作底稿_<年>_生成.xlsx（每年一份）
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import audit_common  # noqa: E402  discover_entities 等（原依赖调用方 import，补显式）


def build_merge_workbook(data_dir, out_dir=None, years=None, quiet=False, entity_whitelist=None):
    """生成合并工作底稿（横展审定表 + 内部交易提示 + 合并调整分录）。
    entity_whitelist：可选主体白名单（如 XBJ 仅合并层级清单内 168 个主体，2026-08-15）；
    None=合并账套全部主体。返回 (ok, out_paths, issues)。issues=[(科目, 年度, 说明)]。"""
    import audit_common as A2
    import subjects_registry as REG
    import audit_shell as S
    from tb_recon_detail import _sheet_ent_values

    if not os.path.isdir(data_dir):
        return True, [], []
    # ⚡ 2026-08-15 铁律130b：SAP 账套（AH 等）主体清单与年度来自 read_sap_tb（无 U8 文件名结构）
    _sap = False
    try:
        import sap_reader as _SR0
        _sap = bool(_SR0.detect_sap_layout(data_dir))
    except Exception:
        _sap = False
    if entity_whitelist:
        entities = sorted(set(entity_whitelist))
    else:
        # ⚡ 2026-08-24 架构阶段2：统一 backend.discover（SAP/U8 形态内聚，替代 read_sap_tb 取主体）
        from ledger_backend import get_backend as _gb
        entities = sorted(_gb(data_dir).discover())
    # 单体模式（SAP 逐主体底稿）：不生成合并工作底稿
    if len(entities) <= 1:
        if not quiet:
            print(f'  ⏭️ [合并工作底稿] 单体模式（{len(entities)} 主体），跳过')
        return True, [], []
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    if years is None:
        if _sap:
            years = ['2026']
        else:
            years = sorted({str(yy) for (_e, _c, _n, yy)
                            in A2.read_tb_full(data_dir, A2.discover_entities(data_dir))}) \
                or ['2025']
    issues = []
    out_paths = []
    acct = os.path.basename(data_dir.rstrip('\\/'))
    # ⚡⚡ 2026-08-15：账套名从路径匹配 ACCT_NAMES（data_dir 可能是 <账套>/prepared 子目录，
    #   basename 取到 'prepared' → merge_subjects 配置匹配失败 → 补充科目全丢）
    try:
        import paths as _P
        _dp = os.path.normcase(os.path.abspath(data_dir))
        for _n in _P.ACCT_NAMES:
            if os.path.normcase(os.path.join(_P.DATA_ROOT, _n)) in _dp:
                acct = _n
                break
    except Exception:
        pass
    for yy in years:
        # ---- ⚡⚡ 2026-08-15 merge_source='self_tb'：合并稿与自建试算表【严格同源】 ----
        #   ga 实证：40 主体底稿审定表结构杂（双层表头/项目×主体长表/块式组件表），
        #   _sheet_ent_values 仅 20/51 科目提取成功；且 ga TB 损益发生额部分主体缺失
        #   （GL 兜底后才有数）→ 底稿提取路径不可靠。self_tb 模式直接复用
        #   self_tb_gen.build_bs/build_is（已含 GL 兜底）逐主体结果 → 合并稿与
        #   自建试算表同函数同数据 → 天然对平。配置：account_profiles.json
        #   accounts.<acct>.merge_source='self_tb'。其余账套仍走底稿提取+merge_subjects。
        import subject_mapping as _SM2
        _msrc = _SM2.feature(acct, 'merge_source')
        if _msrc == 'self_tb':
            wb = __import__('openpyxl').Workbook()
            wb.remove(wb.active)
            import self_tb_gen as _STG2
            _ent_tb2, _yrs2 = _STG2.build_entity_tb(data_dir, [yy], comps=entity_whitelist)
            _ent_bs2, _unmatched2 = _STG2.build_bs(_ent_tb2, yy)
            _ent_is2 = _STG2.build_is(_ent_tb2, yy, data_dir=data_dir)   # ⚡ 6600 费用总池拆分
            _BS2 = [rn for rn in _STG2.BS_ASSET_ROWS] + \
                   [rn for rn in _STG2.BS_LIAB_ROWS] + \
                   [rn for rn in _STG2.BS_EQUITY_ROWS]
            _inc2 = [rn for rn in _STG2.IS_ROWS if _STG2.IS_ROWS[rn].get('credit')]
            _cost2 = [rn for rn in _STG2.IS_ROWS if _STG2.IS_ROWS[rn].get('debit')]
            # rows：资产段+合计 → 负债段+合计 → 权益段+合计 → 其他(未匹配) →
            #       收入类+收入合计 → 成本费用类+成本费用合计 → 净利润
            rows = []
            tot_labels = []
            for rn in _STG2.BS_ASSET_ROWS:
                rows.append((rn, rn))
            rows.append(('资产合计', '资产合计')); tot_labels.append('资产合计')
            for rn in _STG2.BS_LIAB_ROWS:
                rows.append((rn, rn))
            rows.append(('负债合计', '负债合计')); tot_labels.append('负债合计')
            for rn in _STG2.BS_EQUITY_ROWS:
                rows.append((rn, rn))
            rows.append(('所有者权益合计', '所有者权益合计')); tot_labels.append('所有者权益合计')
            if any(_unmatched2.get(e) for e in entities):
                rows.append(('其他（未匹配）', '其他（未匹配）'))
            for rn in _inc2:
                rows.append((rn, rn))
            rows.append(('收入合计', '收入合计')); tot_labels.append('收入合计')
            for rn in _cost2:
                rows.append((rn, rn))
            rows.append(('成本费用合计', '成本费用合计')); tot_labels.append('成本费用合计')
            rows.append(('净利润', '净利润')); tot_labels.append('净利润')
            # 逐主体值（build_bs 带符号 / build_is 口径；与试算表 Sheet3/4 完全同源）
            ent_vals_t = {}
            for e in entities:
                d = {}
                _b = _ent_bs2.get(e, {})
                _i = _ent_is2.get(e, {})
                for rn in _BS2:
                    d[rn] = float(_b.get(rn, 0.0) or 0.0)
                d['其他（未匹配）'] = float(_b.get('其他（未匹配）', 0.0) or 0.0)
                for rn in (_inc2 + _cost2):
                    d[rn] = float(_i.get(rn, 0.0) or 0.0)
                d['资产合计'] = sum(d.get(rn, 0.0) for rn in _STG2.BS_ASSET_ROWS)
                d['负债合计'] = sum(d.get(rn, 0.0) for rn in _STG2.BS_LIAB_ROWS)
                d['所有者权益合计'] = sum(d.get(rn, 0.0) for rn in _STG2.BS_EQUITY_ROWS)
                d['收入合计'] = sum(d.get(rn, 0.0) for rn in _inc2)
                d['成本费用合计'] = sum(d.get(rn, 0.0) for rn in _cost2)
                d['净利润'] = float(_i.get('五、净利润', 0.0) or 0.0)
                ent_vals_t[e] = d
            if any(abs(ent_vals_t[e].get('净利润', 0.0)) > 0.005 for e in entities):
                S.build_merge_wide_sheet(wb, '合并审定表（横展）',
                                         f'全科目合并审定表（{yy} 年度）· 横展 科目×主体 + 合并4列'
                                         f'（与自建试算表同源）',
                                         rows, ent_vals_t, prev_vals=None,
                                         merge_label='合并报表数',
                                         bold_rows=tuple(tot_labels))
            else:
                wb.close()
                continue
            # ---- 内部交易抵消核对（跨主体候选提示）----
            ws2 = wb.create_sheet('内部交易抵消核对')
            ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
            c = ws2.cell(1, 1, f'内部交易抵消核对（{yy} 年度）· 同一往来单位出现于 ≥2 主体 = 内部交易候选，抵消待填')
            c.font = S.SHELL_TITLE_FONT
            for j, h in enumerate(['科目', '往来单位', '出现主体数', '主体列表', '期末余额合计', '抵消状态'], 1):
                cc = ws2.cell(2, j, h)
                cc.font = S.SHELL_HFONT; cc.fill = S.SHELL_HFILL
                cc.alignment = S.SHELL_CEN; cc.border = S.SHELL_BORDER
            r = 3
            _cands = _interco_candidates(data_dir, yy, entities)
            for cand in _cands:
                for j, v in enumerate(cand, 1):
                    cc = ws2.cell(r, j, v)
                    cc.border = S.SHELL_BORDER
                    cc.alignment = S.SHELL_LEFT if j in (2, 4) else S.SHELL_CEN
                    if j == 6:
                        cc.fill = __import__('openpyxl').styles.PatternFill('solid', fgColor='FFF2CC')
                r += 1
            if r == 3:
                cc = ws2.cell(3, 2, '（无跨主体往来候选）')
                cc.font = S.SHELL_SUB_FONT
            for i, w in enumerate([16, 30, 10, 40, 18, 12], 1):
                ws2.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width = w
            # ---- 合并调整分录（空模板）----
            ws3 = wb.create_sheet('合并调整分录')
            ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
            c = ws3.cell(1, 1, f'合并调整分录（{yy} 年度）· 审计人员填列')
            c.font = S.SHELL_TITLE_FONT
            for j, h in enumerate(['序号', '调整事项', '借方科目', '贷方科目', '金额', '调整说明'], 1):
                cc = ws3.cell(2, j, h)
                cc.font = S.SHELL_HFONT; cc.fill = S.SHELL_HFILL
                cc.alignment = S.SHELL_CEN; cc.border = S.SHELL_BORDER
            for i, w in enumerate([8, 30, 14, 14, 16, 30], 1):
                ws3.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width = w
            S.finalize_workbook(wb)
            out_p = os.path.join(out_dir, f'合并工作底稿_{yy}_生成.xlsx')
            wb.save(out_p)
            out_paths.append(out_p)
            if not quiet:
                print(f'  ✅ [合并工作底稿] {os.path.basename(out_p)}（self_tb 同源：'
                      f'{len(rows) - len(tot_labels)} 科目 × {len(entities)} 主体）')
            continue
        # ---- 逐科目提取逐主体审定数 ----
        ent_vals = {}       # {科目中文名: {ent: val}}
        prev_vals = {}      # {科目中文名: 上年数}（同年度文件无上年，留空）
        for key, subj in REG.REGISTRY.items():
            fk = subj.get('file_key')
            name = subj.get('name', '')
            if not fk:
                continue
            fp = os.path.join(data_dir, f'{fk}审计底稿_{yy}_生成.xlsx')
            if not os.path.exists(fp):
                continue
            try:
                wb = __import__('openpyxl').load_workbook(fp, read_only=True, data_only=True)
            except Exception:
                continue
            aud_sn = next((s for s in wb.sheetnames if '审定表' in s), None)
            if aud_sn is None:
                wb.close()
                continue
            try:
                # ⚡⚡ 2026-08-15 XBJ 收入翻倍修复：收入审定表为『项目×主体』长表
                #   （A 列=项目名、B 列=主体名），无 proj_kw 会把『全集团合计』行+全部
                #   项目行都当主体提取 → 营业收入 157.90 亿=2×78.95 亿（双计）。
                #   ⚠️ proj_kw 仅对『收入』科目生效——其他科目审定表同为『科目×主体』长表
                #   （货币资金/往来等 A 列=科目名），proj_kw 过滤会滤掉 A 列不含科目名的行
                #   → 货币资金只剩 30 万、应收/应交税费丢失（资产合计 229 亿→112 亿）。
                #   收入科目 A 列『一、主营业务收入』含『收入』，SUBJ_ROW_KW 展开匹配安全。
                _kw = name if '收入' in name else None
                ev = _sheet_ent_values(wb[aud_sn], proj_kw=_kw, year=yy) or {}
            except Exception:
                ev = {}
            wb.close()
            # 仅保留与账套主体交集（防跨年度实体混入）
            ev = {k: v for k, v in ev.items() if k in entities}
            if not ev:
                continue
            ent_vals[name] = ev
        # ⚡⚡ 2026-08-15 补充科目（merge_subjects，配置驱动）：账套有 TB 数据但无底稿的
        #   科目（XBJ 施工特有：合同资产/合同负债/应付账款/内部结算中心存款等）→
        #   从 TB 直取（复用 self_tb_gen 的 build_entity_tb/l1_agg，与自建试算表同源 → 天然一致；
        #   一级名称匹配 + qm 转正口径）。配置：account_profiles.json accounts.<acct>.merge_subjects
        #   ⚡ 2026-08-15 扩展：
        #     - side='debit'/'credit'：损益类取发生额（成本/费用=借方发生、收入=贷方发生），
        #       期末余额 qm=0（已结转）→ 默认 qm 取数对损益类无效（主营成本 57 亿缺失根因）；
        #     - signed=True：带符号（不 abs 转正）——未分配利润等须带符号求和（底稿 abs 展示值
        #       Σ|主体|≠Σ 带符号净额，XBJ 32.55 亿 vs 试算表 8.94 亿根因）。
        import subject_mapping as _SM
        _extras = _SM.feature(acct, 'merge_subjects') or []
        if _extras:
            import self_tb_gen as _STG
            _ent_tb, _yrs2 = _STG.build_entity_tb(data_dir, [yy])
            for _cfg in _extras:
                _nm = str(_cfg.get('name') or '')
                if not _nm:
                    continue
                # 显示名（横展科目名）缺省=name；匹配名支持多名称（未分配利润=利润分配+本年利润）
                _disp = str(_cfg.get('display') or _nm)
                _names = [str(x) for x in (_cfg.get('names') or [_nm]) if x]
                _side = str(_cfg.get('side') or 'qm')
                _signed = _cfg.get('signed', False)
                _ev = {}
                for _e in entities:
                    _agg = _STG.l1_agg(_ent_tb.get(_e, {}).get(yy, {}))
                    _tot = 0.0
                    for _c, _v in _agg.items():
                        if str(_v.get('name') or '') in _names:
                            if _side == 'debit':
                                _tot += float(_v.get('jf') or 0.0)
                            elif _side == 'credit':
                                _tot += float(_v.get('df') or 0.0)
                            else:
                                _q = float(_v.get('qm') or 0.0)
                                if _signed == 'credit':
                                    _tot += -_q    # 贷正（权益类：贷方余额为正，未分配利润用）
                                elif _signed is True:
                                    _tot += _q    # 带符号（备抵类：贷方余额为负=资产减项）
                                else:
                                    _tot += _q if _q >= 0 else -_q   # abs 转正
                    _ev[_e] = round(_tot, 2)
                if any(abs(v) > 0.005 for v in _ev.values()):
                    ent_vals[_disp] = _ev
        if not ent_vals:
            continue
        # ---- 横展审定表（科目×主体 + 合并4列）----
        wb = __import__('openpyxl').Workbook()
        wb.remove(wb.active)
        # ⚡⚡ 2026-08-15 用户方法论：科目按 资产→负债→权益→损益 报表顺序排列，
        #   各段合计行（资产合计/负债合计/所有者权益合计/收入合计/成本费用合计）**紧跟在各自段末科目之后**，
        #   净利润最后（=收入合计−成本费用合计）。
        _names = sorted(ent_vals)
        names_sorted = A2.sort_report_names(_names)
        # ⚡ build_merge_wide_sheet 期望 ent_vals={主体: {项目key: 值}}（列=主体）
        #   → 把 {科目: {主体: 值}} 转置
        ent_vals_t = {ent: {nm: ev.get(ent, 0.0) for nm, ev in ent_vals.items()}
                      for ent in entities}
        # 段合计与利润表小计（2026-08-15）：按主体逐主体计算
        _tot_vals = A2.merge_segment_totals(ent_vals_t, _names)
        _SEG_TAIL = {'asset': '资产合计', 'liab': '负债合计', 'equity': '所有者权益合计',
                     'pl_income': '收入合计', 'pl_cost': '成本费用合计'}
        rows = []            # (label, key)
        tot_labels = []      # 出现过的段尾小计 label（按出现顺序）
        _cur_cat = None
        for nm in names_sorted:
            _cat, _ = A2.report_cat(nm)
            # 段切换：上一段结束时，把上一段的合计行紧跟其后
            if _cur_cat is not None and _cat != _cur_cat and _cur_cat in _SEG_TAIL:
                _lbl = _SEG_TAIL[_cur_cat]
                rows.append((_lbl, _lbl))
                tot_labels.append(_lbl)
            rows.append((nm, nm))
            _cur_cat = _cat
        if _cur_cat in _SEG_TAIL:          # 末段科目后补该段合计
            _lbl = _SEG_TAIL[_cur_cat]
            rows.append((_lbl, _lbl))
            tot_labels.append(_lbl)
        rows.append(('净利润', '净利润'))   # 净利润恒在最后
        tot_labels.append('净利润')
        for ent in entities:
            for lbl in tot_labels:
                ent_vals_t[ent][lbl] = _tot_vals[lbl].get(ent, 0.0)
        S.build_merge_wide_sheet(wb, '合并审定表（横展）',
                                 f'全科目合并审定表（{yy} 年度）· 横展 科目×主体 + 合并4列（按报表科目顺序）',
                                 rows, ent_vals_t, prev_vals=None,
                                 merge_label='合并报表数',
                                 bold_rows=tuple(tot_labels))
        # ---- 内部交易抵消核对（跨主体候选提示）----
        ws2 = wb.create_sheet('内部交易抵消核对')
        ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        c = ws2.cell(1, 1, f'内部交易抵消核对（{yy} 年度）· 同一往来单位出现于 ≥2 主体 = 内部交易候选，抵消待填')
        c.font = S.SHELL_TITLE_FONT
        for j, h in enumerate(['科目', '往来单位', '出现主体数', '主体列表', '期末余额合计', '抵消状态'], 1):
            cc = ws2.cell(2, j, h)
            cc.font = S.SHELL_HFONT; cc.fill = S.SHELL_HFILL
            cc.alignment = S.SHELL_CEN; cc.border = S.SHELL_BORDER
        r = 3
        _cands = _interco_candidates(data_dir, yy, entities)
        for cand in _cands:
            for j, v in enumerate(cand, 1):
                cc = ws2.cell(r, j, v)
                cc.border = S.SHELL_BORDER
                cc.alignment = S.SHELL_LEFT if j in (2, 4) else S.SHELL_CEN
                if j == 6:
                    cc.fill = __import__('openpyxl').styles.PatternFill('solid', fgColor='FFF2CC')
            r += 1
        if r == 3:
            cc = ws2.cell(3, 2, '（无跨主体往来候选）')
            cc.font = S.SHELL_SUB_FONT
        for i, w in enumerate([16, 30, 10, 40, 18, 12], 1):
            ws2.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width = w
        # ---- 合并调整分录（空模板）----
        ws3 = wb.create_sheet('合并调整分录')
        ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        c = ws3.cell(1, 1, f'合并调整分录（{yy} 年度）· 审计人员填列')
        c.font = S.SHELL_TITLE_FONT
        for j, h in enumerate(['序号', '调整事项', '借方科目', '贷方科目', '金额', '调整说明'], 1):
            cc = ws3.cell(2, j, h)
            cc.font = S.SHELL_HFONT; cc.fill = S.SHELL_HFILL
            cc.alignment = S.SHELL_CEN; cc.border = S.SHELL_BORDER
        for i, w in enumerate([8, 30, 14, 14, 16, 30], 1):
            ws3.column_dimensions[__import__('openpyxl').utils.get_column_letter(i)].width = w
        S.finalize_workbook(wb)
        out_p = os.path.join(out_dir, f'合并工作底稿_{yy}_生成.xlsx')
        wb.save(out_p)
        out_paths.append(out_p)
        if not quiet:
            print(f'  ✅ [合并工作底稿] {os.path.basename(out_p)}（{len(ent_vals)} 科目 × {len(entities)} 主体）')
    return True, out_paths, issues


def _interco_candidates(data_dir, yy, entities=None):
    """跨主体往来候选：扫描各往来科目底稿明细表，同一往来单位出现于 ≥2 主体 → 候选。
    返回 [(科目, 单位, 主体数, 主体列表, 余额合计, '待抵消')]。"""
    import audit_common
    import subjects_registry as REG
    import openpyxl
    if entities is None:
        entities = sorted(audit_common.discover_entities(data_dir))
    out = []
    # 往来类科目（明细表含『往来单位名称』列）
    ca_keys = ('ar', 'ap', 'ar_other', 'ap_other', 'advance_recv', 'prepay',
               'contract_asset', 'contract_liab', 'note_recv', 'apn')
    for key in ca_keys:
        subj = REG.REGISTRY.get(key)
        if not subj:
            continue
        fk = subj.get('file_key')
        fp = os.path.join(data_dir, f'{fk}审计底稿_{yy}_生成.xlsx')
        if not os.path.exists(fp):
            continue
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception:
            continue
        det_sn = next((s for s in wb.sheetnames if '明细表' in s and str(yy) in s), None)
        if det_sn is None:
            wb.close()
            continue
        ws = wb[det_sn]
        # 找表头（含『往来单位名称』）
        hdr_row = None
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            if row and '往来单位名称' in [str(x) if x else '' for x in row]:
                hdr_row = i
                break
        if hdr_row is None:
            wb.close()
            continue
        rows = list(ws.iter_rows(values_only=True))
        hdr = rows[hdr_row - 1]
        c_ent = hdr.index('核算主体') if '核算主体' in hdr else 0
        c_nm = hdr.index('往来单位名称') if '往来单位名称' in hdr else 3
        c_end = next((j for j, h in enumerate(hdr) if h == '期末余额(人民币)'), None)
        ent_unit = {}      # {unit: set(ent)}
        ent_sum = {}       # {unit: 期末合计}
        for row in rows[hdr_row:]:
            if not row or not row[c_nm]:
                continue
            ent = row[c_ent] if c_ent is not None else ''
            unit = str(row[c_nm])
            ent_unit.setdefault(unit, set()).add(ent)
            if c_end is not None and isinstance(row[c_end], (int, float)):
                ent_sum[unit] = ent_sum.get(unit, 0.0) + row[c_end]
        wb.close()
        for unit, ents in ent_unit.items():
            if len(ents) >= 2:
                out.append((subj.get('name', fk), unit, len(ents),
                            '、'.join(sorted(ents)), round(ent_sum.get(unit, 0.0), 2), '待抵消'))
    out.sort(key=lambda x: (-x[2], x[0]))
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    # ⚡ 2026-08-15 集团分组：--comps 主体白名单（逗号分隔，SAP 分组试算表同源用）
    comps = None
    if '--comps' in sys.argv:
        _i = sys.argv.index('--comps')
        if _i + 1 < len(sys.argv):
            comps = [x.strip() for x in sys.argv[_i + 1].split(',') if x.strip()]
    if not args:
        print('用法：python audit_merge.py <账套目录> [输出目录] [--comps 主体白名单]')
        return 1
    data_dir = args[0]
    out_dir = args[1] if len(args) > 1 else data_dir
    ok, paths, issues = build_merge_workbook(data_dir, out_dir, entity_whitelist=comps)
    print(f'{"OK" if ok else "FAIL"}：{paths or "（单体模式跳过）"}')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
