# -*- coding: utf-8 -*-
"""底稿「生成后自动检查器」（2026-07-31 用户要求：嵌套在所有小程序中）。

目的：把「打开底稿才发现的问题」前移到「生成时自动报告」——任何一个小程序单跑
（或 regen 全量）完成终审后，自动扫描该文件夹全部底稿，报告问题清单，避免
「只跑单个小程序仍有错误、要靠人工打开逐张发现」。

检查项（按严重级）：
  E1 ERROR  合计/小计行与明细行不符（公式单元格除外——审定数 =C+D 等联动公式跳过）
  E2 WARN   应有钱的「明细表/主体表」整表无数据（仅表头，而同簿存在带数据的分月/抽凭表）
  W1 WARN   按年文件（*_2025_/*_2026_）内 sheet 混年（如 2025 文件含 2026 行）；
            「期间对比」类双年表为刻意设计，列入白名单跳过
  W2 WARN   抽凭/检查表 摘要列或对方科目列空缺比例 > 15%

性能：read_only 流式 iter_rows 一次性取行，全部检查在内存 list 上完成（不做随机 cell 访问）。

输出：控制台报告（按文件分组 + 汇总 ERROR/WARN 计数）。不修改任何文件。

用法：python audit_checker.py <文件夹> [<文件夹>...]
"""
import paths as P
import os
import re
import sys
import glob

import openpyxl

# 2026-08-04：存货审定表『主体|年度|科目』每主体多年度行属设计（按年度分行核对），
# 加入白名单防混年误报（重跑规范化后出现）。
MIXED_YEAR_WHITELIST = ('期间对比', '差异对比', '勾稽', '存货审定表')
YEAR_VALUES = ('2024', '2025', '2026', '2027')
YEAR_COL_KW = ('年度', '年份', '会计期间', '会计年度')


def _fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _is_total_row_mem(row):
    if not row:
        return False
    a = str(row[0]).strip() if row[0] is not None else ''
    # 2026-08-06 修复：说明/勾稽行（A 列以『勾稽：』等开头，正文含『合计』字样——
    # 如 gp_other 明细表尾『勾稽：明细表合计期末审定数 … vs TB…』）被误判为合计行 →
    # 明细区无值却校验 → 4 个单元格假 ERROR。说明行排除。
    # 2026-08-06 补：职工薪酬『集团汇总数（内容同上面各主体…）』区块标题行含『合计/小计』
    # 字样同样误判（薪酬分配核对表 17 / 多借多贷核对表 13 假 ERROR）。
    if a.startswith(('勾稽', '注:', '注：', '说明', '备注', '口径', '核对说明', '差异', '集团汇总数')):
        return False
    if a in ('合计', '总计', '小计') or a.endswith(('合计', '小计', '总计')):
        return True
    # 2026-08-06 同步 recalc_totals 17:50 修复：A 列【包含】『合计/小计/总计』也识别
    # （loan 明细表『【01FY本级公司 2025 小计】』以『】』结尾，endswith 不命中 →
    # 小计/合计行被当数据行 → 全集团合计 3 倍误报）。
    if '合计' in a or '小计' in a or '总计' in a:
        return True
    for idx in range(1, min(3, len(row))):
        v = row[idx]
        if v is None:
            continue
        x = str(v).strip()
        if x in ('合计', '小计', '总计') or '小计' in x:
            return True
    return False


def _is_header_row_mem(row):
    if _is_total_row_mem(row):
        return False
    hits = 0
    for v in row[:12]:
        if v is None:
            continue
        s = str(v).strip()
        if any(k in s for k in ('核算主体', '科目', '项目', '月份', '期间', '年度', '年份', '序号', '账户', '对方科目', '摘要', '日期', '凭证', '金额', '余额', '数量', '名称', '主体', '账面', '差额')):
            hits += 1
    return hits >= 3


def _mem_total_fix(rows):
    """纯内存版「合计行需修正数」（与 recalc_totals.recalc_sheet 同规则）。"""
    total_idx = [i for i, r in enumerate(rows) if _is_total_row_mem(r)]
    hdr_idx = [i for i, r in enumerate(rows) if _is_header_row_mem(r)]
    fixed = 0
    maxc = max((len(r) for r in rows), default=0)
    maxc = min(maxc, 40)
    for ti in total_idx:
        hds = [h for h in hdr_idx if h < ti]
        if not hds:
            continue
        # 表级合计信任 builder（与 recalc_totals 同步）：『全集团合计/集团合计』由 builder
        # 按各主体小计精确汇总填入，重算会双计父级+子目 → 跳过
        a0 = str(rows[ti][0]) if rows[ti] and rows[ti][0] else ''
        b0 = str(rows[ti][1]) if len(rows[ti]) > 1 and rows[ti][1] else ''
        if '全集团' in a0 or '集团合计' in a0 or '全集团' in b0:
            continue
        # 块级小计范围 = 上一个合计行之后（铁律44：与 recalc_totals 同步）；无则从最近表头后。
        # 『总计』行 = 表级全范围（从最近表头后起算），不用上一个合计行后（期间对比表总计=末年度）
        prev_total = [t for t in total_idx if t < ti]
        if a0 == '总计' or a0 == '总 计':
            start = hds[-1] + 1
        else:
            start = (prev_total[-1] + 1) if prev_total else (hds[-1] + 1)
        # 嵌套小计/合计行【剔除】出求和（2026-08-01 与 recalc_totals 同步：防双计且能修正被污染合计）
        inner = set(t for t in total_idx if start <= t < ti)
        # ⚡ 2026-08-10 豁免『差异/勾稽』列：这些列是计算结果（如应交税费勾稽差异 -23,848.64，
        #   期初+计提−支付−期末 与 匹配列差异），小计≠Σ明细属设计内，非合计错误。
        _hdr_row = rows[hds[-1]] if hds else []
        for c in range(maxc):
            if c < len(_hdr_row) and _hdr_row[c] is not None and \
                    ('差异' in str(_hdr_row[c]) or '勾稽' in str(_hdr_row[c])):
                continue
            vals = []
            for i in range(start, ti):
                if i in inner:
                    continue
                if c < len(rows[i]):
                    v = rows[i][c]
                    if isinstance(v, (int, float)):
                        vals.append(v)
            if not vals:
                continue
            cv = rows[ti][c] if c < len(rows[ti]) else None
            if isinstance(cv, str) and cv.strip().startswith('='):
                continue  # 公式单元格不动
            total = round(sum(vals), 2)
            # ⚡ 2026-08-10 修复浮点误报：合计行浮点值（756542940.1700001）与
            #   round(sum)=756542940.17 精确比较不等 → 61 个 ERROR 假象（数值实际正确）。
            #   改容差比较（0.005=分位）；仍不等才是真需重算。
            if isinstance(cv, (int, float)) and abs(cv - total) > 0.005:
                fixed += 1
    return fixed


def _mem_year_col(rows):
    for r in rows[:6]:
        for c in range(min(len(r), 14)):
            v = r[c]
            if v is not None and str(v).strip() in YEAR_COL_KW:
                return c
    return None


def _mem_mixed_years(rows, col):
    seen = set()
    for r in rows:
        if col < len(r) and r[col] is not None:
            s = str(r[col]).strip()
            if s in YEAR_VALUES:
                seen.add(s)
    return len(seen) > 1


def _mem_date_col(rows):
    for r in rows[:4]:
        for c in range(min(len(r), 16)):
            v = r[c]
            if v is not None and str(v).strip() in ('日期', '凭证日期'):
                return c
    return None


def _mem_mixed_dates(rows, col):
    seen = set()
    for r in rows[2:]:
        if col < len(r) and r[col] is not None:
            m = re.search(r'(?:19|20)\d{2}', str(r[col]))
            if m and m.group(0) in YEAR_VALUES:
                seen.add(m.group(0))
    return len(seen) > 1


def _mem_data_rows(rows):
    """统计"真实数据行"数：排除表头行(第1行)与 合计/小计/总计/年度合计 等汇总行。
    2026-08-01 11文件夹全量测试改进：原实现把表头('序号')与'2025 年度合计'行计入，
    导致 c 单主体1行数据误报为无数据(≤2 命中)、J/Z/S 研发支出空表(仅合计行)漏报。
    2026-08-05 空壳底稿豁免：只计【金额列（前2个数值列）非零】的数据行——
    研发支出等科目未启用时审定表全 0 行（『无数据（科目未启用）』）不算数据，缺附注不报 ERROR。"""
    n = 0
    for i, r in enumerate(rows):
        if i == 0:  # 表头行
            continue
        if not r:
            continue
        t = str(r[0]).strip() if r[0] is not None else ''
        _rowtxt = ' '.join(str(x) for x in r[:4] if x is not None)
        if not t and not _rowtxt:  # 全空行（r[0]=None 的主体延续行 B/C 有内容 → 保留）
            continue
        # 汇总行（合计/小计/总计）——含主体延续行（r[0] 空但 B/C 列有明细名）不误跳
        if any(k in _rowtxt for k in ('合计', '小计', '总计')):
            continue
        # 仅计金额列非零的行（查前 8 列任一数值>0.005；明细表 A主体/B名称 后金额列在 C 起——
        # 2026-08-05 空壳豁免：研发支出全 0 行不算数据，缺附注不报 ERROR）
        _has_val = False
        for x in r[1:8]:
            if isinstance(x, (int, float)) and abs(float(x)) > 0.005:
                _has_val = True
                break
        if not _has_val:
            continue
        n += 1
    return n


def _mem_vouch_blank(rows):
    """抽凭/检查表：摘要/对方科目空缺比例。返回 (摘要空%, 对方空%) 或 (None, None)。"""
    hdr = {}
    for r in rows[:8]:
        for c in range(min(len(r), 16)):
            v = r[c]
            if v is None:
                continue
            s = str(v).strip()
            if s == '摘要':
                hdr.setdefault('sum', c)
            elif s == '对方科目':
                hdr.setdefault('opp', c)
            elif s == '借方金额':
                hdr.setdefault('dr', c)
            elif s == '贷方金额':
                hdr.setdefault('cr', c)
    if 'sum' not in hdr or ('dr' not in hdr and 'cr' not in hdr):
        return None, None
    total = blank_s = blank_o = 0
    for r in rows[2:]:
        a0 = str(r[0]).strip() if r[0] is not None else ''
        if not a0:
            continue  # A 列空 = 非数据行（如 loan 抽凭表 A 列空的「借贷合计」行）
        if '合计' in a0 or '总计' in a0:
            continue  # 合计/总计行不算数据行（避免「0 笔」空表被误判为空缺）
        dv = _fnum(r[hdr['dr']]) if 'dr' in hdr and hdr['dr'] < len(r) else None
        cv = _fnum(r[hdr['cr']]) if 'cr' in hdr and hdr['cr'] < len(r) else None
        if dv is None and cv is None:
            continue  # 借贷均无金额 → 非数据行（核对区/说明）
        total += 1
        sv = r[hdr['sum']] if hdr['sum'] < len(r) else None
        if sv is None or not str(sv).strip():
            blank_s += 1
        if 'opp' in hdr:
            ov = r[hdr['opp']] if hdr['opp'] < len(r) else None
            if ov is None or not str(ov).strip():
                blank_o += 1
    if total == 0:
        return None, None
    return round(blank_s / total * 100, 1), (round(blank_o / total * 100, 1) if 'opp' in hdr else None)


def check_file(fp):
    issues = []
    base = os.path.basename(fp)
    try:
        wb = openpyxl.load_workbook(fp, data_only=False, read_only=True)
    except Exception as ex:
        return [('ERROR', f'文件无法打开：{ex}')]
    is_year_file = bool(re.search(r'_(\d{4})_', base))
    # 预读各 sheet 行（read_only 流式）
    sheet_rows = {}
    try:
        for sn in wb.sheetnames:
            ws = wb[sn]
            sheet_rows[sn] = [list(r) for r in ws.iter_rows(values_only=True)]
    except Exception as ex:
        wb.close()
        return [('ERROR', f'读取失败：{ex}')]
    for sn, rows in sheet_rows.items():
        # E1 合计/小计与明细不符
        # 2026-08-02：职工薪酬 审定表 为「每主体一块、块内含 SUM 汇总行+小计+合计」的特殊结构，
        # 通用合计校验会把 SUM 汇总行当明细导致误报 → 跳过（其正确性由构建勾稽 GAP 行保证）。
        # ⚡⚡ 2026-08-30 XBJ 命名不同（职工薪酬横展审定表_2025/集团审定表_2025）同结构 → 一并豁免
        if sn.startswith('职工薪酬') and '审定表' in sn:
            pass
        else:
            try:
                n_fix = _mem_total_fix(rows)
                if n_fix:
                    issues.append(('ERROR', f'[{sn}] 合计/小计与明细不符（{n_fix} 个合计单元格需重算）'))
            except Exception:
                pass
        # W1 按年文件混年
        if is_year_file and not any(k in sn for k in MIXED_YEAR_WHITELIST):
            yc = _mem_year_col(rows)
            if yc is not None:
                if _mem_mixed_years(rows, yc):
                    issues.append(('WARN', f'[{sn}] 年度列同时含多个年份（混年）'))
            else:
                dc = _mem_date_col(rows)
                if dc is not None and _mem_mixed_dates(rows, dc):
                    issues.append(('WARN', f'[{sn}] 日期列含其他年份凭证（混年）'))
        # E2 应有钱的明细表/主体表整表无数据
        # _mem_data_rows 已排除表头与汇总行 → 真实数据行==0 才算无数据 (2026-08-01 全量测试校准)
        # 注：分部门明细表等「无部门级辅助核算」注记型空表（r1 为说明文字）属设计内不生成，跳过 E2。
        if ('明细表' in sn or '主体比较' in sn or sn.endswith('_2025') or sn.endswith('_2026')):
            if _mem_data_rows(rows) == 0:
                # ⚡⚡ 2026-08-24 修复（空表明细表 WARN 误报）：空表注记可能不在首行
                #   （current_account 子表『无发生额及余额（TB 中无对应科目）』在表头后）。
                #   改为全表扫描注记词——含任一注记 = 正常空表（生成器已标注），跳过。
                _flat0 = ''.join(str(x) if x is not None else '' for row in rows[:40] for x in row[:12])
                if any(k in _flat0 for k in ('未生成', '无「部门」级辅助核算', '无数据',
                                             '无发生额及余额', '未启用')):
                    continue
                other_has = any(sn2 != sn and _mem_data_rows(sheet_rows[sn2]) > 5
                                for sn2 in sheet_rows)
                if other_has:
                    issues.append(('WARN', f'[{sn}] 整表无数据（仅表头），同簿其他表有数据'))
        # W2 抽凭/检查表 摘要/对方科目空缺
        if '凭证抽查' in sn or '检查表' in sn:
            bs, bo = _mem_vouch_blank(rows)
            if bs is not None and bs > 15:
                issues.append(('WARN', f'[{sn}] 摘要空缺 {bs}%'))
            if bo is not None and bo > 15:
                issues.append(('WARN', f'[{sn}] 对方科目空缺 {bo}%'))
    # ============ 规范检查（2026-08-04，从合并工作出发）============
    # 底稿三件套结构必须能被核对程序（tb_recon）自动取数——生成即拦截，防"账套多了漏洞百出"。
    # 检查项：①审定表可取数 ②有附注汇总 sheet ③余额型明细表可取数（发生额型/空壳豁免）。
    # 仅对『*审计底稿*』执行（关联交易核对/核对底稿等辅助文件非科目底稿，豁免）。
    try:
        from tb_recon_detail import _sheet_end_total, _footnote_total
        if '审计底稿' in base:
            _has_data = any(_mem_data_rows(rows) > 5 for rows in sheet_rows.values())
            if _has_data:
                _aud_sn = next((s for s in wb.sheetnames if '审定表' in s), None)
                _det_sn = next((s for s in wb.sheetnames
                                if ('明细' in s or '分类汇总' in s)
                                and not any(k in s for k in ('准备', '计提', '期间对比', '对方科目'))), None)
                _fn_sn = next((s for s in wb.sheetnames if '附注' in s), None)
                if _aud_sn:
                    try:
                        _av, _why = _sheet_end_total(wb[_aud_sn])
                        if _av is None:
                            issues.append(('WARN', f'[{_aud_sn}] 核对程序取不到审定数（{_why}）——表头/合计行结构不规范'))
                    except Exception:
                        pass
                if _fn_sn is None:
                    # 2026-08-04 升级 ERROR：缺『附注汇总』=结构性缺失（合并附注生成依赖），
                    # 不再是可忽略的 WARN——生成即失败提示。
                    issues.append(('ERROR', '缺『附注汇总』sheet（合并附注生成依赖，结构性缺失）'))
                # 余额型明细（分月/变动=发生额型豁免；空壳=无数据豁免已由 _has_data 拦截）
                if _det_sn and not any(k in _det_sn for k in ('分月', '变动', '对比')):
                    try:
                        _dv, _dwhy = _sheet_end_total(wb[_det_sn])
                        if _dv is None:
                            issues.append(('WARN', f'[{_det_sn}] 核对程序取不到明细合计（{_dwhy}）——表头/结构不规范'))
                    except Exception:
                        pass
                # ⚡⚡ 2026-08-25 规范：底稿应含『审计程序执行说明』『勾稽与异常检查』
                #   （audit_program 在 finalize 统一注入；旧底稿未重跑暂无，重跑后补齐）。
                if '审计程序执行说明' not in wb.sheetnames:
                    issues.append(('WARN', '缺『审计程序执行说明』sheet（审计程序记录，audit_program 注入；重跑后补齐）'))
                if '勾稽与异常检查' not in wb.sheetnames:
                    issues.append(('WARN', '缺『勾稽与异常检查』sheet（表内勾稽+L4 异常规则，audit_program 注入）'))
                # 顺序规范：两张程序表应排在工作簿末尾（收尾表）。
                # ⚡⚡ 2026-08-29 放宽：程序表之后只允许『审计调整分录』类收尾表
                #   （bank 等模块在 audit_program 注入后追加调整分录到更末尾，属合理设计，
                #   原要求"程序表必须是倒数 1/2 位"会误报）。
                if '勾稽与异常检查' in wb.sheetnames and '审计程序执行说明' in wb.sheetnames:
                    _ix_prog = wb.sheetnames.index('审计程序执行说明')
                    _ix_rc = wb.sheetnames.index('勾稽与异常检查')
                    _suffix = wb.sheetnames[max(_ix_prog, _ix_rc) + 1:]
                    _tail_ok = {'审计调整分录', '调整分录'}
                    if _ix_prog > _ix_rc or any(s not in _tail_ok for s in _suffix):
                        issues.append(('WARN', '『审计程序执行说明/勾稽与异常检查』未置于末尾（顺序规范）'))
    except Exception:
        pass
    wb.close()
    return issues


def process_folder(folder, quiet=False):
    # 2026-08-04：quiet=True（regen_all 每 builder 后增量检查用）——不打印逐文件明细，
    # 只返回 (ERROR 数, WARN 数)；非 quiet 时照常打印完整报告。
    # ⚡ 2026-08-10 SAP 适配：SAP 账套底稿命名 {科目}审计底稿_{主体}.xlsx（无 _生成 后缀，
    #   主体=1010/300集团 等）且位于 {账套}/底稿/ 子目录（U8 在账套根目录）——
    #   匹配扩展为 *审计底稿*.xlsx 并扫描 根目录+底稿/ 子目录。
    _segs = [folder]
    _sub = os.path.join(folder, '底稿')
    if os.path.isdir(_sub):
        _segs.append(_sub)
    files = []
    for _d in _segs:
        files += glob.glob(os.path.join(_d, '*审计底稿*.xlsx'))
    files = sorted(files)
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    # 跳过"冗余合并稿"：无年份后缀、但存在同名 _20xx_ 按年文件的中间产物
    # （如 current_account 的"应收账款审计底稿_生成.xlsx"——按年文件才是交付物，合并稿不检查）
    yearly = {os.path.basename(f) for f in files if re.search(r'_(?:19|20)\d{2}_生成\.xlsx$', os.path.basename(f))}
    files = [f for f in files
             if re.search(r'_(?:19|20)\d{2}_生成\.xlsx$', os.path.basename(f))
             or not any(os.path.basename(f).replace('_生成.xlsx', '') + '_%s_生成.xlsx' % y == b
                        for y in ('2025', '2026') for b in yearly)]
    if not files:
        print(f'  ⚠️ {folder}：无 审计底稿*.xlsx（U8 *_生成.xlsx 或 SAP 审计底稿_主体.xlsx）')
        return 0, 0
    n_err = n_warn = 0
    if not quiet:
        print(f'--- 自动检查：{folder}（{len(files)} 份底稿）---')
    for fp in files:
        issues = check_file(fp)
        if not issues:
            continue
        base = os.path.basename(fp)
        for lvl, msg in issues:
            if lvl == 'ERROR':
                n_err += 1
            else:
                n_warn += 1
            if not quiet:
                print(f'  {lvl}  {base} {msg}')
    if not quiet:
        print(f'  → 检查完成：ERROR {n_err} / WARN {n_warn}')
    # ⚡ 2026-08-13 跨主体串户检查（#738 配套，用户方法论：串户=同科目多主体审定表
    #   数字完全相同，正常业务不可能——区间文件取错主体的铁证）：
    dup = _cross_dup_check(folder)
    for base, comps in dup:
        n_warn += 1
        if not quiet:
            print(f'  WARN  {base} 串户嫌疑：{comps} 审定表数字完全相同（正常业务不可能）')
    return n_err, n_warn


def _cross_dup_check(folder):
    """跨主体串户检查：同科目族多主体的审定表 sheet 数值签名完全相同 → 串户嫌疑。
    返回 [(文件名, [主体列表])]。签名=审定表全部数值（≥3 个且 >1e-6）的 tuple。"""
    from collections import defaultdict
    _segs = [folder]
    _sub = os.path.join(folder, '底稿')
    if os.path.isdir(_sub):
        _segs.append(_sub)
    groups = defaultdict(list)  # stem -> [(comp, sig)]
    for _d in _segs:
        for fp in glob.glob(os.path.join(_d, '*审计底稿*.xlsx')):
            base = os.path.basename(fp)
            m = re.match(r'^(.*?)审计底稿_(\d{4})\.xlsx$', base)
            if not m:
                continue
            stem, comp = m.group(1), m.group(2)
            sig = _sheet_sig(fp)
            if sig:
                groups[stem].append((comp, sig))
    out = []
    for stem, items in groups.items():
        by_sig = defaultdict(list)
        for comp, sig in items:
            by_sig[sig].append(comp)
        for sig, comps in by_sig.items():
            # 仅「含小数」的签名判串户：整数全同（注册资本 2000 万等）=正常业务；
            # 固定资产 178,043,445.55 这类复杂小数全同=区间文件取错主体的铁证
            if len(comps) >= 2 and any(abs(v - round(v)) > 0.001 for v in sig):
                out.append((f'{stem}审计底稿', sorted(comps)))
    return out


def _sheet_sig(fp):
    """审定表 sheet 数值签名（全部数值 round 2 位），读取失败返回 None。"""
    try:
        wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
        sig = None
        for sn in wb.sheetnames:
            if '审定表' not in sn:
                continue
            vals = []
            for row in wb[sn].iter_rows(values_only=True):
                for v in row:
                    if isinstance(v, (int, float)) and abs(v) > 1e-6:
                        vals.append(round(v, 2))
            if len(vals) >= 3:
                sig = tuple(vals)
            break
        wb.close()
        return sig
    except Exception:
        return None


def main():
    dirs = sys.argv[1:] or [P.G]
    te = tw = 0
    for d in dirs:
        e, w = process_folder(d)
        te += e
        tw += w
    print(f'\n自动检查汇总：共 {te} 个 ERROR、{tw} 个 WARN。')
    return 1 if te else 0


if __name__ == '__main__':
    sys.exit(main())
