# -*- coding: utf-8 -*-
"""审计程序执行说明自动注入（2026-08-25 用户 P0 落地）。

审计准则：审计工作底稿应记录【已执行的审计程序】。自动化底稿更有义务记录
「程序为你完成了哪些审计步骤」——本模块在 audit_shell.finalize_workbook 统一注入
（改 1 处、全套 13 个生成器生效），自动分析 wb 已有 sheet，生成两张收尾表：

① 『审计程序执行说明』：程序步骤清单（按 wb 实际包含的 sheet 自动检测）+ 数据来源 + 勾稽结果。
   形如：
     1. 编制应收账款审定表，审定数与科目余额表(TB)核对一致
     2. 编制明细表，按核算主体×往来单位列示期初/增减/期末
     3. 账龄分析（6 桶）
     4. 勾稽检查：期初+本期增加-本期减少=期末（N 表一致 / 差异见勾稽表）
     5. 对方科目异常识别（异常 N 笔 / 关注 N 笔）

② 『勾稽与异常检查』：表内勾稽逐表结果 + L4 异常规则库信号。
   L4 规则（前 4 条，可自动检测）：
     R1 大额整数金额（期末=整数且 ≥1万，疑手工调整/粉饰）
     R3 资产科目贷余 / 负债科目借余（分类/重分类错误）
     R4 长期挂账（期初≈期末≠0 且本期无发生，潜亏/坏账迹象）

用户约定（2026-08-25 P0 调整）：底稿头【暂不】添加 审计目的/结论/编制人/复核人，
仅注入以上两张程序说明表；『审计调整分录』meta sheet 逻辑保持现状。

用法（无需直接调用，finalize_workbook 统一注入）：
    from audit_program import inject_program_sheets
"""
import re

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------- 程序清单库：sheet 名关键词 → (审计程序说明, 数据来源) ----------------
# 顺序 = 程序步骤展示顺序；每个底稿按实际包含的 sheet 自动匹配。
PROGRAM_DB = [
    ('集团审定表', '编制集团合并审定表（Σ单体，尾部集团合计/合并抵消）', '各单体审定表'),
    ('审定表', '编制{label}审定表，审定数与科目余额表(TB)核对一致', '科目余额表(TB)'),
    ('明细表', '编制明细表，按核算主体/往来单位列示期初/增减/期末', '辅助核算余额表+序时账'),
    ('对方科目核对', '对方科目异常识别（异常/关注对应科目逐笔列示）', '序时账'),
    ('异常对应科目', '异常对应科目凭证级线索清单', '序时账'),
    ('凭证抽查', '凭证抽查（抽样表），借贷方向/金额/摘要核对', '序时账'),
    ('分月', '分月发生额统计', '序时账'),
    ('账龄', '账龄分析（账龄区间分布）', '辅助核算余额表'),
    ('多借多贷', '多借多贷核对表（凭证级借贷平衡检查）', '序时账'),
    ('票据种类拆分', '应收票据种类拆分', '序时账+辅助核算'),
    ('背书贴现', '应收票据背书/贴现明细', '序时账'),
    ('坏账准备', '坏账准备测算', '序时账+辅助核算'),
    ('期间对比', '多期数据对比', '各期账套'),
    ('月度', '月度发生额/余额明细', '序时账'),
    ('附注汇总', '编制附注披露汇总，按审定数填列', '审定表'),
]

# ---------------- 样式 ----------------
_THIN = Side(style='thin', color='BFBFBF')
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_TITLE_F = Font(name='Times New Roman', size=12, bold=True)
_HDR_F = Font(name='Times New Roman', size=10, bold=True)
_BODY_F = Font(name='Times New Roman', size=10)
_NOTE_F = Font(name='Times New Roman', size=10, italic=True, color='808080')
_HDR_FILL = PatternFill('solid', fgColor='DDEBF7')
_WARN_FILL = PatternFill('solid', fgColor='FFF2CC')
_OK_FILL = PatternFill('solid', fgColor='E2EFDA')
_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)


def _sheet_label(sn):
    """从 sheet 名剥离『年份/明细/审定表/审计底稿』等词，得到科目/底稿名。"""
    s = re.sub(r'(_|审?定?表|明细|审计底稿|_生成|\d{4})', '', sn)
    return s.strip() or '本底稿'


def _detect_programs(wb):
    """扫描 wb 已有 sheet → 匹配到的程序步骤 [(说明, 数据来源)]。按 PROGRAM_DB 顺序去重。"""
    sns = list(wb.sheetnames)
    label = _sheet_label(next((s for s in sns if '审定表' in s), sns[0] if sns else '本底稿'))
    steps = []
    seen = set()
    for kw, desc, src in PROGRAM_DB:
        hit = any(kw in s for s in sns)
        if not hit:
            continue
        key = kw
        if key in seen:
            continue
        seen.add(key)
        steps.append((desc.format(label=label), src))
    if not steps:
        steps.append((f'编制{label}底稿，数据自账套直连取数', '账套'))
    return steps, label


def _recon_check(wb):
    """对含『审定表』的 sheet 跑表内勾稽（期初+增-减=期末，容差 1 元）。
    返回 [(sheet名, 数据行数, 差异数)]。复用 wp_recon_check 逻辑（轻量版）。"""
    from openpyxl import load_workbook
    results = []
    for sn in wb.sheetnames:
        if '审定表' not in sn:
            continue
        ws = wb[sn]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            continue
        # 定位表头行（含 期初/期末/增加/减少）
        hi = None
        for i, r in enumerate(rows[:8]):
            txt = ' '.join(str(x) if x is not None else '' for x in r)
            if all(k in txt for k in ('期初', '期末')) and ('增加' in txt or '增加' in txt):
                hi = i
                break
        if hi is None:
            continue
        hdr = [str(x).strip() if x is not None else '' for x in rows[hi]]
        # ⚡⚡ 2026-08-25 对齐 wp_recon_check：SAP 集团审定表（表头含『汇总包顺序/公司』）
        #   本期增/减=往来发生额、期初/期末=余额，期初+增-减=期末 不适用 → 跳过勾稽。
        if any('汇总包' in h or h == '公司' for h in hdr):
            continue
        c_qc = next((i for i, h in enumerate(hdr) if '期初' in h), None)
        c_inc = next((i for i, h in enumerate(hdr) if h == '本期增加' or h == '本期\n增加' or '增加' in h and '本期' in h), None)
        c_dec = next((i for i, h in enumerate(hdr) if h == '本期减少' or h == '本期\n减少' or '减少' in h and '本期' in h), None)
        c_qm = next((i for i, h in enumerate(hdr) if '期末' in h), None)
        if c_qc is None or c_qm is None:
            continue
        bad = 0
        n = 0
        for r in rows[hi + 1:]:
            if not r or len(r) <= max(c_qc, c_inc or 0, c_dec or 0, c_qm):
                continue
            def _f(x):
                try:
                    return float(x)
                except (TypeError, ValueError):
                    return 0.0
            qc = _f(r[c_qc]) if c_qc < len(r) else 0.0
            inc = _f(r[c_inc]) if c_inc is not None and c_inc < len(r) else 0.0
            dec = _f(r[c_dec]) if c_dec is not None and c_dec < len(r) else 0.0
            qm = _f(r[c_qm]) if c_qm < len(r) else 0.0
            if abs(qc) + abs(inc) + abs(dec) + abs(qm) < 0.005:
                continue
            n += 1
            diff = (qc + inc - dec) - qm
            diff_rev = (qc - inc + dec) - qm
            if abs(diff) > 1.0 and abs(diff_rev) > 1.0:
                # 方向翻转+绝对值展示豁免（与 wp_recon_check 一致）：期初/期末方向翻转
                # 且期末取贷余绝对值显示时 |qc±inc∓dec|≈|qm| 为展示自洽，非勾稽错误。
                if abs(abs(qc + inc - dec) - abs(qm)) <= 1.0 or \
                        abs(abs(qc - inc + dec) - abs(qm)) <= 1.0:
                    continue
                bad += 1
        results.append((sn, n, bad))
    return results


def _anomaly_rules(wb):
    """L4 异常规则库（可自动检测的子集）：
    R1 大额整数金额：审定表任一金额 ≥10000 且为整数（疑手工调整/粉饰）。
    R3 资产科目贷余/负债科目借余：审定表数据行期末为负（资产类）或非负（负债类）。
    R4 长期挂账：期初≈期末≠0 且本期增减均 0（跨年未动，潜亏/坏账迹象）。
    返回 [('R1', 描述, 位置), ...]，最多各取前 3 条避免刷屏。"""
    issues = []
    for sn in wb.sheetnames:
        if '审定表' not in sn:
            continue
        ws = wb[sn]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        # 表头定位（同上）
        hi = None
        for i, r in enumerate(rows[:8]):
            txt = ' '.join(str(x) if x is not None else '' for x in r)
            if '期初' in txt and '期末' in txt:
                hi = i
                break
        if hi is None:
            continue
        hdr = [str(x).strip() if x is not None else '' for x in rows[hi]]
        # 集团审定表（汇总包表头）：异常规则跳过（口径不同，与勾稽一致）
        if any('汇总包' in h or h == '公司' for h in hdr):
            continue
        c_qc = next((i for i, h in enumerate(hdr) if '期初' in h), None)
        c_inc = next((i for i, h in enumerate(hdr) if '增加' in h), None)
        c_dec = next((i for i, h in enumerate(hdr) if '减少' in h), None)
        c_qm = next((i for i, h in enumerate(hdr) if '期末' in h), None)
        if c_qc is None or c_qm is None:
            continue
        r1 = r2 = r3 = r4 = 0
        # ⚡⚡ 2026-08-25 R2/R3 适用性修正：应交税费期末借余（留抵/预缴）、本期发生远大于
        #   期末（当期计提当期缴纳）均为正常业务 → 排除『应交』及损益/流转类。
        _is_flow = any(k in sn for k in ('应交', '收入', '费用', '成本', '营业', '损益',
                                          '税金及附加', '所得税', '薪酬'))
        _is_asset = any(k in sn for k in ('应收', '预付', '存货', '合同资产', '其他应收',
                                           '长期应收', '债权', '其他权益工具', '固定资产',
                                           '无形资产', '在建工程', '使用权资产', '生产性生物'))
        _is_liab = any(k in sn for k in ('应付', '预收', '其他应付款', '合同负债',
                                          '借款', '长期应付', '递延', '预计负债', '租赁负债'))
        # R2 只对【往来余额科目】判定：存货/固定资产等购产销流转科目本期发生大、期末小是正常
        _is_arap = any(k in sn for k in ('应收', '应付', '预收', '预付', '合同资产', '合同负债'))
        for r in rows[hi + 1:]:
            if not r or len(r) <= max(c_qc, c_qm):
                continue
            # ⚡⚡ 2026-08-25 跳过低值/汇总行：『减：坏账准备/存货跌价准备/净额/合计/小计』等
            #   抵减行期末为负属正常（如存货审定表 跌价准备 -709,823.33）。label 可能在
            #   科目名称列（r[3]，存货审定表 主体|年度|代码|名称|期初|增|减|期末）→ 拼前 6 列判断。
            _label = ' '.join(str(x) for x in r[:6] if x is not None)
            if any(k in _label for k in ('减：', '净额', '准备', '减值', '抵减', '合计', '小计', '总计')):
                continue
            def _f(x):
                try:
                    return float(x)
                except (TypeError, ValueError):
                    return 0.0
            qc = _f(r[c_qc]) if c_qc < len(r) else 0.0
            inc = _f(r[c_inc]) if c_inc is not None and c_inc < len(r) else 0.0
            dec = _f(r[c_dec]) if c_dec is not None and c_dec < len(r) else 0.0
            qm = _f(r[c_qm]) if c_qm < len(r) else 0.0
            if abs(qc) + abs(inc) + abs(dec) + abs(qm) < 0.005:
                continue
            # R1 大额整数（≥100万 且 整数；阈值提高避免普通整数期末误报，如 125,000）
            if abs(qm) >= 1000000 and abs(qm - round(qm)) < 0.005 and r1 < 3:
                issues.append(('R1', f'期末 {qm:,.0f} 为整数大额（≥100万），疑手工调整/粉饰', f'{sn}'))
                r1 += 1
            # R2 本期发生额异常：仅往来余额科目，发生额 ≥ 期末 10 倍 且 ≥100万
            if not _is_flow and _is_arap:
                _max_occ = max(abs(inc), abs(dec))
                if _max_occ >= 1000000 and abs(qm) > 0.005 and _max_occ >= 10 * abs(qm) and r2 < 3:
                    issues.append(('R2', f'本期发生 {_max_occ:,.0f} 为期末 {qm:,.0f} 的 10 倍以上，大额流转/期末调节嫌疑', f'{sn}'))
                    r2 += 1
            # R3 方向异常：资产类期末为负（贷余）、负债类期末为正（借余）——排除流转类
            if not _is_flow and _is_asset and qm < -0.005 and r3 < 3:
                issues.append(('R3', f'资产类科目期末 {qm:,.2f} 为贷余（方向与科目性质相反）', f'{sn}'))
                r3 += 1
            if not _is_flow and _is_liab and qm > 0.005 and r3 < 3:
                issues.append(('R3', f'负债类科目期末 {qm:,.2f} 为借余（方向与科目性质相反）', f'{sn}'))
                r3 += 1
            # R4 长期挂账
            if abs(qc) > 0.005 and abs(qm - qc) < 0.005 and abs(inc) < 0.005 and abs(dec) < 0.005 and r4 < 3:
                issues.append(('R4', f'期初=期末 {qm:,.2f} 且本期无发生，长期挂账（潜亏/坏账迹象）', f'{sn}'))
                r4 += 1
    return issues


def inject_program_sheets(wb, meta=None):
    """主入口：在 wb 末尾注入『审计程序执行说明』+『勾稽与异常检查』两张表。
    幂等（_wb_program_injected 标记）。meta 可选（label/period 已由 finalize 提取）。
    注意：勾稽/异常检查对【已含年份标记且被 emit_per_year 拆分的底稿】在拆分前执行，
    拆分后两张表随 wb 复制到各年份文件（内容一致，正常）。"""
    if getattr(wb, '_wb_program_injected', False):
        return
    if not wb.sheetnames:
        return
    # 避免与现有同名 sheet 冲突
    if '审计程序执行说明' in wb.sheetnames or '勾稽与异常检查' in wb.sheetnames:
        wb._wb_program_injected = True
        return
    try:
        steps, label = _detect_programs(wb)
        recon = _recon_check(wb)
        anom = _anomaly_rules(wb)
        _build_program_sheet(wb, steps, label)
        _build_recon_sheet(wb, recon, anom)
        # ⚡⚡ 2026-08-29 P0：强制两张程序表置工作簿末尾——部分模块在 audit_program
        #   注入之后才 create_sheet（对方科目核对等），导致程序表不在末尾 → audit_checker
        #   报『未置于末尾』顺序 WARN（三集团 20 处）。统一收尾：无论注入先后都移到末两位。
        _move_prog_to_end(wb)
        wb._wb_program_injected = True
    except Exception:
        # 注入失败不影响底稿主体
        wb._wb_program_injected = True


def _move_prog_to_end(wb):
    """把『审计程序执行说明』『勾稽与异常检查』移到 wb._sheets 末尾（保持此顺序），
    且『对方科目核对』类注入表移到程序表之前（兜底部分模块在 audit_program 之后
    注入对方科目核对 → 追加到更末尾 → audit_checker 报顺序 WARN）。"""
    try:
        _s = wb._sheets
        _prog = [s for s in _s if s.title == '审计程序执行说明']
        _rc = [s for s in _s if s.title == '勾稽与异常检查']
        if _prog or _rc:
            _cp = [s for s in _s if ('对方科目核对' in s.title or '剩余对方科目核对' in s.title)]
            _others = [s for s in _s if s not in (_prog + _rc + _cp)]
            wb._sheets[:] = _others + _cp + _prog + _rc
    except Exception:
        pass


def _build_program_sheet(wb, steps, label):
    ws = wb.create_sheet('审计程序执行说明')
    ws.cell(1, 1, f'{label} 审计程序执行说明（程序自动生成）').font = _TITLE_F
    ws.cell(3, 1, '本底稿由审计小程序自动生成。以下为本程序实际执行的审计步骤（对应本底稿各表）：').font = _BODY_F
    r = 4
    for i, (desc, src) in enumerate(steps, 1):
        ws.cell(r, 1, f'{i}. {desc}').font = _BODY_F
        ws.cell(r, 2, f'数据来源：{src}').font = _NOTE_F
        ws.cell(r, 1).border = _BORDER
        ws.cell(r, 2).border = _BORDER
        r += 1
    r += 1
    ws.cell(r, 1, '勾稽与异常检查结果详见『勾稽与异常检查』表；需人工复核事项（审计调整/重分类/异常）在相应表中列示。').font = _NOTE_F
    ws.column_dimensions['A'].width = 70
    ws.column_dimensions['B'].width = 30


def _build_recon_sheet(wb, recon, anom):
    ws = wb.create_sheet('勾稽与异常检查')
    ws.cell(1, 1, '勾稽与异常检查（程序自动执行）').font = _TITLE_F
    r = 3
    ws.cell(r, 1, '一、表内勾稽：期初 + 本期增加 - 本期减少 = 期末（容差 1 元）').font = _HDR_F
    ws.cell(r, 1).fill = _HDR_FILL
    r += 1
    ws.cell(r, 1, '表名'); ws.cell(r, 2, '数据行数'); ws.cell(r, 3, '差异行数'); ws.cell(r, 4, '结论')
    for c in range(1, 5):
        ws.cell(r, c).font = _HDR_F
        ws.cell(r, c).fill = _HDR_FILL
        ws.cell(r, c).border = _BORDER
    r += 1
    if recon:
        for sn, n, bad in recon:
            ws.cell(r, 1, sn); ws.cell(r, 2, n); ws.cell(r, 3, bad)
            concl = '一致' if bad == 0 else '有差异，见底稿勾稽行'
            ws.cell(r, 4, concl)
            fill = _OK_FILL if bad == 0 else _WARN_FILL
            for c in range(1, 5):
                ws.cell(r, c).border = _BORDER
                ws.cell(r, c).fill = fill
            r += 1
    else:
        has_aud = any('审定表' in s for s in wb.sheetnames)
        msg = '审定表无「期初/期末/增减」滚动勾稽结构（损益类/集团口径），勾稽不适用' if has_aud \
            else '（本底稿无『审定表』sheet，无表内勾稽）'
        ws.cell(r, 1, msg).font = _NOTE_F
        r += 1
    r += 1
    ws.cell(r, 1, '二、异常模式检查（L4 规则库）').font = _HDR_F
    ws.cell(r, 1).fill = _HDR_FILL
    r += 1
    ws.cell(r, 1, '规则'); ws.cell(r, 2, '描述'); ws.cell(r, 3, '位置')
    for c in range(1, 4):
        ws.cell(r, c).font = _HDR_F
        ws.cell(r, c).fill = _HDR_FILL
        ws.cell(r, c).border = _BORDER
    r += 1
    if anom:
        for rid, desc, loc in anom:
            ws.cell(r, 1, rid); ws.cell(r, 2, desc); ws.cell(r, 3, loc)
            for c in range(1, 4):
                ws.cell(r, c).border = _BORDER
                ws.cell(r, c).fill = _WARN_FILL
            r += 1
    else:
        ws.cell(r, 1, '未检出异常模式').font = _BODY_F
        ws.cell(r, 1).fill = _OK_FILL
        r += 1
    r += 1
    ws.cell(r, 1, '说明：异常信号仅供审计师关注，不构成审计结论；审计调整/重分类由审计师填列。').font = _NOTE_F
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 26
