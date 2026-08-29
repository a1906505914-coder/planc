# -*- coding: utf-8 -*-
"""audit_result_export.py —— 审计结果中间层（2026-08-02 架构铺垫 P0-②）。

目标：把 13 个 builder 生成的底稿中"审定数"沉淀为结构化 JSON，
供未来大动作直接读取（不解析 xlsx——避免公式缓存/列漂移/合并单元格脆弱性）：
  1. 全科目附注汇总表（按科目枚举、抽审定数）；
  2. Word 附注自动填入；
  3. 上年审定数（期初过入方式 audited 直接采用上年审定数）。

设计原则：
  · 与附注生成器【解耦】：直接从注册表科目 + TB 取数，不依赖各 builder 内部逻辑，
    因此 34 个科目全覆盖（含自带附注生成器的 pl/tax/payroll/longterm 与暂缓附注的往来/存货/收入）；
  · 审定数 = 未审数 + 审计调整数（本期调整默认 0，未来由"调整分录台账"过入后刷新）；
  · 期初/上年数按注册表 opening 方式标记：audit_py（期初未审 vs 上年审定核对）
    / audited（直接采用上年审定数）；
  · 输出目录：<data_dir>/审计结果/科目名_年份.json
"""
from __future__ import annotations
import paths as P
import os
import json
import re

import subjects_registry as REG
import subject_mapping  # 2026-08-07 科目映射配置化（account_profiles.json）

_OUT_DIR = '审计结果'

OPENING_AUDITED = REG.OPENING_AUDITED


# ---------- 导出：按科目从 TB 取未审数 ----------

def _tb_leafs(tb_full, entity, year, codes, names=None):
    """取某主体某年某科目的【叶子级】TB 行（不重复计父+子）。"""
    leafs = []
    rows = []
    for (e, c, n, yy), v in tb_full.items():
        if e != entity or str(yy) != str(year):
            continue
        cs = str(c)
        hit = False
        if names:
            hit = any(k in str(n) for k in names)
        else:
            hit = any(cs == k or (cs.startswith(k) and len(cs) > len(k)) for k in codes)
        if hit:
            rows.append((cs, v))
    for cs, v in rows:
        is_parent = any(cs != c2 and c2.startswith(cs) and len(c2) > len(cs) for c2, _ in rows)
        if not is_parent:
            leafs.append((cs, v))
    return leafs


def _subject_unaudit(tb_full, entity, year, subj):
    """按注册表科目取未审数 4 值：[期初, 本期增加, 本期减少, 期末]。
    取数口径与 audit_common.add_audit_summary_sheets 一致（铁律：TB 权威控制数；
    qc/qm 取绝对值正数展示，借/贷按 is_credit 方向；增减=借发/贷发）。
    """
    codes = subj['codes']
    is_credit = subj['is_credit']
    leafs = _tb_leafs(tb_full, entity, year, codes)
    qc = qm = jf = df = 0.0
    for cs, v in leafs:
        qc += abs(v['qc']); qm += abs(v['qm'])
        jf += v['jf']; df += v['df']
    # 增减按科目性质映射：资产类 增加=借发、减少=贷发；负债/权益类 增加=贷发、减少=借发
    inc, dec = (df, jf) if is_credit else (jf, df)
    return [round(qc, 2), round(inc, 2), round(dec, 2), round(qm, 2)]


def export_subject(data_dir, subj_key, years=None, tb_full=None, entities=None,
                   out_dir=None, adjustments=None):
    """导出单个科目到 JSON。返回输出文件路径（无数据也导出，含空主体）。
    adjustments: {(entity, year): [期初,增,减,末]} 审计调整数（默认全 0）。
    """
    import audit_common
    if entities is None:
        entities = audit_common.discover_entities(data_dir)
    if tb_full is None:
        tb_full = audit_common.read_tb_full(data_dir, entities)
    subj = REG.REGISTRY[subj_key]
    if years is None:
        years = sorted({str(y) for e in entities for y in entities[e]})
    out_dir = out_dir or os.path.join(data_dir, _OUT_DIR)
    os.makedirs(out_dir, exist_ok=True)

    adj = adjustments or {}
    rec = {
        'subject': subj['name'],
        'file_key': subj['file_key'],
        'builder': subj['builder'],
        'footnote_mode': subj['footnote'],
        'opening': subj['opening'],          # audit_py / audited
        'adj_structure': subj['adj'],        # simple / split
        'is_credit': subj['is_credit'],
        'codes': subj['codes'],
        'years': {},
    }
    for year in years:
        yrec = {'entities': {}, 'note': subj['note']}
        for entity in sorted(entities):
            u = _subject_unaudit(tb_full, entity, year, subj)
            a = [0.0, 0.0, 0.0, 0.0]
            if adj.get((entity, year)):
                a = adj[(entity, year)]
            # 审定数 = 未审 + 调整（期初/增/减/末 四段）
            aud = [round(u[i] + a[i], 2) for i in range(4)]
            # 损益类（opening='audited'）补充发生额口径：期末审定=本期发生额（借/贷按性质）
            flow = None
            if subj['opening'] == OPENING_AUDITED:
                _jf = sum(l[1]['jf'] for l in _tb_leafs(tb_full, entity, year, subj['codes']))
                _df = sum(l[1]['df'] for l in _tb_leafs(tb_full, entity, year, subj['codes']))
                flow = round(_df if subj['is_credit'] else _jf, 2)
            yrec['entities'][entity] = {
                'unaudit': u,     # [期初, 本期增加, 本期减少, 期末]（未审数）
                'adj': a,         # [期初, 本期增加, 本期减少, 期末]（审计调整，默认 0）
                'audited': aud,   # [期初, 本期增加, 本期减少, 期末]（审定数）
                'flow': flow,     # 本期发生额（损益类；余额类 None）
            }
        rec['years'][str(year)] = yrec

    out = os.path.join(out_dir, f"{subj['file_key']}_{year if len(years) == 1 else 'all'}.json")
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(rec, f, ensure_ascii=False, indent=1)
    return out


def export_folder(data_dir, years=None, tb_full=None, entities=None, out_dir=None,
                  only_keys=None, quiet=False):
    """导出全部分子科目（或 only_keys 子集）到 <data_dir>/审计结果/。返回输出文件列表。"""
    import audit_common
    if entities is None:
        entities = audit_common.discover_entities(data_dir)
    if tb_full is None:
        tb_full = audit_common.read_tb_full(data_dir, entities)
    out_dir = out_dir or os.path.join(data_dir, _OUT_DIR)
    os.makedirs(out_dir, exist_ok=True)
    outs = []
    for key, subj in REG.REGISTRY.items():
        if only_keys and key not in only_keys:
            continue
        # 未生成底稿的科目（账套无此科目）不导出：无该科目 TB 数据 → 跳过
        if not _subject_has_data(tb_full, entities, subj):
            if not quiet:
                print(f'  ⊘ {subj["name"]}：账套无此科目数据，跳过')
            continue
        try:
            out = export_subject(data_dir, key, years=years, tb_full=tb_full,
                                 entities=entities, out_dir=out_dir)
            outs.append(out)
            if not quiet:
                print(f'  ✓ {subj["name"]} -> {os.path.basename(out)}')
        except Exception as ex:
            print(f'  ⚠️ {subj["name"]} 导出失败：{ex}')
    return outs


def _subject_has_data(tb_full, entities, subj):
    """任一主体任一年有该科目任一 codes 前缀的 TB 行即视为有数据。"""
    codes = subj['codes']
    for (e, c, n, yy), v in tb_full.items():
        cs = str(c)
        if any(cs == k or cs.startswith(k) for k in codes):
            return True
    return False


# ---------- 读取 ----------

def load_subject(data_dir, subj_key, year=None):
    """读单科目 JSON。返回 dict 或 None。year=None 读 all 文件。"""
    subj = REG.REGISTRY[subj_key]
    out_dir = os.path.join(data_dir, _OUT_DIR)
    fn = os.path.join(out_dir, f"{subj['file_key']}_{year or 'all'}.json")
    if not os.path.exists(fn):
        return None
    with open(fn, encoding='utf-8') as f:
        return json.load(f)


def load_folder(data_dir, year=None):
    """读全部分子科目 JSON。返回 {subj_key: rec}。"""
    recs = {}
    for key in REG.REGISTRY:
        rec = load_subject(data_dir, key, year=year)
        if rec:
            recs[key] = rec
    return recs


# ---------- 汇总：全科目附注审定数（供全科目附注汇总/Word） ----------

def summarize_all(data_dir, year=None):
    """全科目审定数汇总表（分主体 × 科目 × 期末审定数）。
    返回 list[dict]：{subject, file_key, entity, year, audited_qm, footnote_mode, opening}。
    """
    rows = []
    for key, rec in load_folder(data_dir, year=year).items():
        for yy, yrec in rec['years'].items():
            for ent, vals in yrec['entities'].items():
                u, a, aud = vals['unaudit'], vals['adj'], vals['audited']
                rows.append({
                    'subject': rec['subject'], 'file_key': rec['file_key'],
                    'entity': ent, 'year': yy,
                    'unaudit_qm': u[3], 'adj_qm': a[3], 'audited_qm': aud[3],
                    'flow': vals.get('flow'),
                    'footnote_mode': rec['footnote_mode'],
                    'opening': rec['opening'],
                })
    return rows


def export_summary_xlsx(data_dir, out_path=None, year=None):
    """全科目附注审定数汇总表（大动作②雏形）：行=主体×科目、列=未审/调整/审定。
    out_path 默认 <data_dir>/全科目附注审定数汇总_<year or all>.xlsx。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    rows = summarize_all(data_dir, year=year)
    if not rows:
        return None
    out_path = out_path or os.path.join(data_dir, f'全科目附注审定数汇总_{year or "all"}.xlsx')
    wb = Workbook(); ws = wb.active; ws.title = '全科目审定数汇总'
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr = ['核算主体', '年度', '科目', '期初未审', '本期增加未审', '本期减少未审', '期末未审',
           '期末调整', '期末审定数', '本期发生额', '期初过入方式', '附注模板']
    for j, h in enumerate(hdr, 1):
        c = ws.cell(1, j, h); c.font = Font(name='Times New Roman', size=10, bold=True)
        c.fill = HFILL; c.alignment = Alignment(horizontal='center'); c.border = BORDER
    r = 2
    for row in sorted(rows, key=lambda x: (x['year'], x['entity'], x['subject'])):
        vals = [row['entity'], row['year'], row['subject'],
                None, None, None, row['unaudit_qm'],
                row['adj_qm'], row['audited_qm'],
                row['flow'], row['opening'], row['footnote_mode']]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.font = Font(name='Times New Roman', size=10)
            c.border = BORDER
            if 4 <= j <= 10 and v is not None:
                c.number_format = '#,##0.00'; c.alignment = Alignment(horizontal='right')
        r += 1
    for j, w in enumerate([22, 8, 14, 14, 14, 14, 14, 12, 14, 14, 12, 10], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'A2'
    wb.save(out_path)
    return out_path


if __name__ == '__main__':
    import sys
    data_dir = sys.argv[1] if len(sys.argv) > 1 else P.FY
    print(f'=== 导出审计结果中间层：{data_dir} ===')
    outs = export_folder(data_dir)
    print(f'\n共导出 {len(outs)} 个科目 JSON 到 {os.path.join(data_dir, _OUT_DIR)}')
    print('\n=== 汇总样例（前 8 行）===')
    for r in summarize_all(data_dir)[:8]:
        print(f"  {r['subject']:10s} {r['entity'][:12]:14s} {r['year']} 期末审定={r['audited_qm']:,.2f}")
    print('\n=== 全科目附注审定数汇总表 ===')
    x = export_summary_xlsx(data_dir)
    print('  ✓', x)


# ===================== 产物完整性自检（2026-08-06） =====================

def coverage_check(data_dir, years=None, quiet=False):
    """产物完整性自检：注册表科目【TB 有数据】→ 对应底稿文件必须存在。

    目的：杜绝"builder 静默跳过/识别失败"导致的缺底稿（jtt 曾因 read_km 表头
    不兼容整表 0 科目 → 6 组资产底稿全缺且 regen_all 仍报 13/13 成功）。
    规则：_subject_has_data(TB 有该科目任一 codes 前缀) 且 该年度底稿文件缺失 → 报缺口。
    返回 (ok, missing_list)；missing 元素 = (科目名, file_key, 年度)。
    """
    import audit_common
    if not os.path.isdir(data_dir):
        return True, []
    entities = audit_common.discover_entities(data_dir)
    tb_full = audit_common.read_tb_full(data_dir, entities)
    if years is None:
        years = sorted({str(yy) for (e, c, n, yy) in tb_full}) or ['2025']
    missing = []
    for key, subj in REG.REGISTRY.items():
        # 2026-08-06：按【金额非零】判定有数据（科目存在但 qc/qm/jf/df 全 0 不算——
        # jj 单体 2711/2901 科目存在但余额 0，builder 正确跳过，不应报缺）。
        has_amt = False
        nm_kw = subj.get('name', '')

        def _name_hit(_n, nm_kw):
            """铁律5/23 名称匹配：①正向包含（标准名⊂实际名，如『税金及附加』⊂『营业税金及附加』）；
            ②反向包含仅限【标准名括号成分】（如『股本』⊂『实收资本(或股本)』的『或股本』）——
            2026-08-06 修复：H 账套 223304『费用』⊂『销售费用』、22210104『所得税』⊂
            『递延所得税负债/资产』被反向包含误判有数据 → 完整性自检假缺失。"""
            if not nm_kw or not _n:
                return False
            if nm_kw in _n:
                return True
            if _n in nm_kw:
                # 反向：仅当 _n 是 nm_kw 括号内成分（或X / (X) / X）才认
                for _part in re.split(r'[（()）]', nm_kw):
                    _p = _part.strip('或')
                    if _p and _p != nm_kw and _n == _p:
                        return True
            return False

        for (e, c, n, yy), v in tb_full.items():
            # 2026-08-07 配置化：科目名先按账套元数据归一化（H『递延资产』→『长期待摊费用』、
            # Q『投资损益』→『投资收益』）——跨账套别名统一走 account_profiles.json，不进代码。
            _n = subject_mapping.norm_name(os.path.basename(data_dir.rstrip('\\/')), n)
            _hit = _name_hit(_n, nm_kw)
            if not _hit and subj.get('codes'):
                _cs = str(c)
                if any(_cs == k or _cs.startswith(k) for k in subj['codes']):
                    # 2026-08-06：codes 兜底须【名称含 subj 名词根】才认——
                    # S 账套 4104=利润分配 撞 专项储备标准码 4104 → 名称无『储备』词根 → 不算有数据。
                    _parts = [p for p in re.split(r'[/（()）]', nm_kw or '') if len(p) >= 2]
                    if _parts and any(p in _n for p in _parts):
                        _hit = True
            if not _hit:
                continue
            if any(abs(float(v.get(f) or 0.0)) > 0.005 for f in ('qc', 'qm', 'jf', 'df')):
                has_amt = True
                break
        if not has_amt:
            continue
        fk = subj['file_key']
        for yy in years:
            fn = f'{fk}审计底稿_{yy}_生成.xlsx'
            if os.path.exists(os.path.join(data_dir, fn)):
                continue
            # 个别科目文件名非"file_key审计底稿"模式（如 借款拆分后 短期借款/长期借款 已各自登记）
            alt = [x for x in os.listdir(data_dir)
                   if x.endswith(f'审计底稿_{yy}_生成.xlsx') and fk in x]
            if not alt:
                missing.append((subj['name'], fk, yy))
    if missing:
        print('  ⚠️ [完整性自检] 以下科目 TB 有数据但底稿文件缺失：')
        for name, fk, yy in missing:
            print(f'     · {name}（{fk}）{yy} —— 请检查该 builder 是否识别失败/静默跳过')
        return False, missing
    print('  ✅ [完整性自检] 全部有数据科目均已生成底稿（0 缺失）')
    return True, []


def run_coverage(data_dir):
    """regen_all 末尾调用：不中断流程，只报告缺口。"""
    try:
        ok, _ = coverage_check(data_dir)
        return ok
    except Exception as ex:
        print(f'  ⚠️ [完整性自检] 执行异常：{ex}')
        return False
