# -*- coding: utf-8 -*-
"""bank_reconcile_detail.py —— 银行网银双向核对底稿（#739）

用户需求（2026-08-13）：
  ① 从序时账抽出生成银行日记账（复用 bank_deposit_detail.read_gl_bank，通用 U8/SAP/300）
  ② 网银记录与日记账拼接
  ③ 双向核对：网银→日记账、日记账→网银，不一致直接标注
  ④ 支持「多条网银 = 一条汇总日记账」（用户：企业把相同收款汇总记入银行日记账）

核对规则（企业视角）：
  方向映射：网银『收』↔ 日记账『借』（存入）；网银『支』↔ 日记账『贷』（支出）
  匹配优先级：
    1. 精确：同方向 + 金额相等(±0.01) + 日期差 ≤ window(默认3天)
    2. 汇总：剩余网银按 (日期窗口, 方向) 聚合，多条之和 = 一条日记账
    3. 剩余 = 差异（网银有账无 / 账有网银无），标注原因

用法（拖文件夹或命令行）：
  python bank_reconcile_detail.py <账套目录> [<网银文件夹>]
    账套目录：300/6100（含序时账/科目余额表）；网银文件夹缺省时自动找同级『*网银*』目录
输出：<账套目录>/银行网银核对底稿_<年度>.xlsx
"""
import os
import sys
import re
import json
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from bank_statement_parser import parse_bank_folder

BANK_KWS = ['国开银行', '国开行', '中信', '建设银行', '建行', '交通银行', '交行',
            '农业银行', '农行', '工商银行', '工行', '中国银行', '中行', '农商行', '农村商业银行']
# 归一化：全称 → 简称（防『中国银行』vs『中行』、『国开银行』vs『国开行』分裂）
_BANK_NORM = {'中国银行': '中行', '建设银行': '建行', '农业银行': '农行',
              '工商银行': '工行', '交通银行': '交行', '国开银行': '国开行',
              '农村商业银行': '农商行'}


def _bank_of_name(name):
    """科目名/文件名 → 归一化银行键（'工行'/'中行'/...）。未识别返回 None。"""
    s = str(name or '')
    best, best_len = None, 0
    for kw in BANK_KWS:
        if kw in s and len(kw) > best_len:
            best, best_len = kw, len(kw)
    return _BANK_NORM.get(best, best) if best else None


def _read_u8_gl_rows(adapter, comp, year):
    """U8 综合查询明细表兜底（DQ 形态 2026-08-13）：SAP 适配层只认 SAP 序时账，
    DQ 是 U8 9 列文件（科目名称/日期/字/号/对方科目/摘要/借方金额/贷方金额/辅助核算名称，
    表头行 2、数据行 3+）→ 直接解析返回 read_gl 同构行。"""
    import openpyxl
    try:
        ents = adapter.discover_entities(None) if adapter else {}
    except Exception:
        return []
    yd = (ents.get(comp) or {}).get(str(year))
    glp = (yd or {}).get('gl')
    if not glp or not os.path.exists(str(glp)):
        return []
    wb = openpyxl.load_workbook(str(glp), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    hdr_i = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True)):
        if r and '科目名称' in str(r[0] or '') and '借方金额' in str(r or ''):
            hdr_i = i + 1
            break
    rows = []
    start = hdr_i + 1 if hdr_i else 3
    for r in ws.iter_rows(min_row=start, values_only=True):
        if not r or not r[0]:
            continue
        # ⚡⚡⚡ 2026-08-14 双字段分离（用户需求：对方科目+对方单位都要）：
        #   DQ 综合查询明细表『对方科目』列(col4)与辅助核算(col8)常为空，只有摘要含单位名。
        #   → cp=会计科目（凭证级推断，同凭证其他行科目名，铁律88 思路）；
        #     opp_name=对方单位名（摘要/aux 提取）。两者分开存，银行日记账两列展示。
        _cp_col = str(r[4] or '').strip() if len(r) > 4 else ''   # 原始对方科目列（常空）
        _aux = str(r[8] or '').strip() if len(r) > 8 else ''      # 辅助核算（常空）
        _sm = str(r[5] or '').strip() if len(r) > 5 else ''       # 摘要（含单位名）
        # 对方单位名：aux 优先（核算维度=单位），空则摘要提取
        _opp = _aux
        if not _opp and _sm:
            _m = re.search(
                r'(?:已付|预付|退回|承兑支付|收到|支付|汇给)\s*'
                r'([\u4e00-\u9fa5A-Za-z0-9（）()]{4,40}?(?:公司|厂|站|部|中心|研究院|事务所|有限))', _sm)
            if not _m:
                _m = re.search(
                    r'(?:已付|预付|退回|承兑支付|收到|支付|汇给)\s*'
                    r'([\u4e00-\u9fa5A-Za-z0-9（）()]{4,25}?)(?:[的货款检测费服务费咨询费款项费租金保证金]|%|$)', _sm)
            if _m:
                _opp = _m.group(1).strip()
        # 会计科目：col4 有值用它；否则留空（由凭证级推断填，见函数尾 _infer_u8_opp）
        _cp = _cp_col
        rows.append({
            'name': str(r[0]), 'date': str(r[1] or '')[:10],
            'vtype': str(r[2] or ''), 'vno': str(r[3] or ''),
            'cp': _cp, 'opp_name': _opp, 'sm': _sm,
            'debit': float(r[6] or 0), 'credit': float(r[7] or 0),
            'aux': _aux,
        })
    wb.close()
    # ⚡⚡ 凭证级对方科目推断（铁律88）：银行存款分录的会计科目（应付/预付/应收/收入…）
    #   在同凭证其他行——同 (vtype,vno) 内非本行科目名按金额排序，填入 cp。
    _infer_u8_opp(rows)
    return rows


def _infer_u8_opp(rows):
    """U8 凭证级对方科目推断：同凭证(vtype+vno)内其他科目名→银行存款行 cp。
    实证（DQ）：付货款凭证=银行存款贷 + 应付/预付账款借 → cp='预付账款'；
    收货款=银行存款借 + 应收账款贷 → cp='应收账款'。已填 cp（col4 有值）的跳过。"""
    from collections import defaultdict
    grp = defaultdict(list)
    for i, r in enumerate(rows):
        grp[(str(r.get('vtype') or ''), str(r.get('vno') or ''))].append(i)
    for idxs in grp.values():
        # 各科目金额（取最大笔代表）
        names = {}
        for i in idxs:
            nm = str(rows[i]['name'] or '').strip()
            if not nm:
                continue
            amt = abs(float(rows[i].get('debit') or 0)) + abs(float(rows[i].get('credit') or 0))
            if nm not in names or amt > names[nm]:
                names[nm] = amt
        ordered = sorted(names, key=lambda x: -names[x])
        for i in idxs:
            if rows[i]['cp']:
                continue  # col4 已有值
            self_nm = str(rows[i]['name'] or '').strip()
            opp = ','.join(n for n in ordered if n != self_nm)
            if opp:
                rows[i]['cp'] = opp[:40]


def read_bank_journal(adapter, comp, year='2026'):
    """从 GL 抽银行存款分录 → {bank: [行]}，行含 date/direction/amount/summary/account/vchar。
    方向=企业视角（借=存入、贷=支出），与网银『收/支』对齐。
    ⚡ 2026-08-13：SAP 适配取不到时（U8 形态如 DQ）→ _read_u8_gl_rows 直接解析综合查询明细表。"""
    rows = adapter.read_gl(comp) if adapter else []
    if not rows:
        rows = _read_u8_gl_rows(adapter, comp, year)
    out = {}
    for r in rows:
        nm = str(r.get('name') or '')
        # ⚡ 2026-08-13 修复（DQ 实证 177 笔误抽）：只抽【银行存款】科目——原条件
        #   『'存款' not in nm and '银行' not in nm』太宽，『应收票据-银行承兑汇票』(含"银行")
        #   与『财务费用-存款利息』(含"存款") 被误抽归"其他"（网银无对应 → 一大堆假账有网银无）。
        #   银行存款科目名规范='银行存款-银行名-账户'；结息类科目（含"结息"）兼容。
        # ⚡⚡ 2026-08-14 定期/通知存款排除（用户需求）：网银流水=活期，定期/通知存款
        #   账户与活期分开，抽进来会成假『账有网银无』（DQ2026 实证：嘉兴银行-定期存款
        #   5 行/杭州银行-定期 1 行/交行-定期 1 行）。『定存/定期/通知存款/大额存单/结构性存款』剔除。
        if '银行存款' not in nm and '结息' not in nm:
            continue
        if any(_k in nm for _k in ('定期', '定存', '通知存款', '大额存单', '结构性存款')):
            continue
        bank = _bank_of_name(nm)
        if bank is None:
            continue
        jf = float(r.get('debit') or r.get('dr') or 0.0)
        df = float(r.get('credit') or r.get('cr') or 0.0)
        if jf == 0 and df == 0:
            continue
        _vno = str(r.get('vno') or '')
        _vt = str(r.get('vtype') or '')
        _vchar = ('%s-%s' % (_vt, _vno)).strip('-') or _vno
        item = {
            'date': _norm_date(r.get('date')),
            'summary': str(r.get('sm') or r.get('summary') or ''),
            'account': nm, 'vchar': _vchar,
            'cp': str(r.get('cp') or r.get('opp') or ''),   # ⚡ 凭证级对方科目（铁律88，adapter _infer_opp）
            'opp_name': str(r.get('opp_name') or ''),       # ⚡⚡ 2026-08-14 对方单位名（摘要/aux）
        }
        if jf:
            out.setdefault(bank, []).append(dict(item, direction='借', amount=float(jf)))
        if df:
            out.setdefault(bank, []).append(dict(item, direction='贷', amount=float(df)))
    for b in out:
        out[b] = _net_red_write(out[b])   # ⚡ 红冲净额化（同日期+金额+摘要 借贷抵消）
        out[b].sort(key=lambda x: (x['date'], x['vchar']))
    return out


def _net_red_write(rows):
    """红冲净额化：同 (date, amount, summary) 的 借+贷 对抵消（SAP 红冲重记特征——
    星星充电 01-31 同摘要『收星星充电桩电费款』借 1232.35 与贷 1232.35 各 2 笔，
    净额化后累计匹配才能闭合）。防误伤：须同日期+同金额+同摘要三者一致。"""
    from collections import defaultdict
    groups = defaultdict(list)
    for i, x in enumerate(rows):
        groups[(x['date'], round(x['amount'], 2), str(x.get('summary') or ''))].append(i)
    drop = set()
    for key, idxs in groups.items():
        jf = [i for i in idxs if rows[i]['direction'] == '借']
        df = [i for i in idxs if rows[i]['direction'] == '贷']
        n = min(len(jf), len(df))
        for i in jf[:n] + df[:n]:
            drop.add(i)
    if drop:
        rows = [x for i, x in enumerate(rows) if i not in drop]
    return rows


def _norm_date(v):
    if v is None:
        return ''
    s = str(v)[:10]
    m = re.match(r'^(\d{4})[-/]?(\d{2})[-/]?(\d{2})', s)
    if m:
        return '%s-%s-%s' % (m.group(1), m.group(2), m.group(3))
    return s


def _dday(s):
    try:
        return datetime.date.fromisoformat(str(s)[:10])
    except Exception:
        return None


def _summary_bizdate(s):
    """账侧摘要业务日期解析（DQ/U8 特有 2026-08-13）：月末集中记账的摘要开头含业务日期
    『X.YZ』（如『2.11中国银行转入嘉兴银行1000万』/『12.04中国银行转农业银行1000万』）
    → 返回 (月, 日)；无 → None。⚠️ 排除金额/小数干扰：月 1-12、日 1-31。"""
    m = re.match(r'\s*(\d{1,2})\.(\d{1,2})', str(s or ''))
    if m:
        mo, d = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12 and 1 <= d <= 31:
            return mo, d
    return None


# ⚡⚡⚡ 2026-08-14 业务关键字分组（用户方法论：找【业务关键字】而非【名字】——
#   记账方法各异（网银"数字信用凭据"↔账"工银e信到期"、网银"上海汇付…跨行转入"↔账
#   "上海汇付数据备用金"、网银"45元电子转账"↔账"国望后勤押金退款"），名字千变万化
#   防不胜防，但业务关键词稳定。同义词归一到同一业务组。
# ⚡⚡ 2026-08-15 配置化（跟随账套走，不再刻程序）：bank_business_keywords.json——
#   _default=全局种子；账套键=extend 追加 / list 完全替换。学习器确认的新词一律落配置。
_KEYWORDS_CFG = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             'bank_business_keywords.json')
_KEYWORDS_CACHE = {}   # acct -> [(group, (kws...))]（None 键=全局默认种子）


def _load_keywords_for(acct):
    """账套业务关键词：全局种子 + 账套覆盖（extend 追加 / list 完全替换）。"""
    if acct in _KEYWORDS_CACHE:
        return _KEYWORDS_CACHE[acct]
    try:
        with open(_KEYWORDS_CFG, encoding='utf-8') as f:
            cfg = json.load(f)
    except Exception:
        cfg = {}
    dflt = [(g, tuple(k)) for g, k in (cfg.get('_default') or [])]
    a = cfg.get(acct)
    if isinstance(a, list) and a:
        out = [(g, tuple(k)) for g, k in a]
    elif isinstance(a, dict) and a.get('extend'):
        out = dflt + [(g, tuple(k)) for g, k in a['extend']]
    else:
        out = dflt
    _KEYWORDS_CACHE[acct] = out
    return out


# 当前账套关键词（build 入口 _set_keywords 设置；核对串行执行+大主体子进程隔离，状态安全）
_CUR_KEYWORDS = None
_DEFAULT_KEYWORDS = _load_keywords_for(None)


def _set_keywords(acct):
    global _CUR_KEYWORDS
    _CUR_KEYWORDS = _load_keywords_for(acct or None)


def _biz_key(text):
    """提取业务关键词（返回业务组标识或 None）。先取摘要，counterparty 兜底。
    关键词=当前账套配置（build 入口设置），缺省=全局种子 bank_business_keywords.json。"""
    s = str(text or '')
    kw = _CUR_KEYWORDS if _CUR_KEYWORDS is not None else _DEFAULT_KEYWORDS
    for group, kws in kw:
        for k in kws:
            if k in s:
                return group
    return None


def _shift_month(ym, delta):
    """月份加减（'2026-01' - 1 → '2025-12'），用于月末汇总滞后 ±1 月配对。"""
    try:
        y, m = int(ym[:4]), int(ym[5:7])
    except Exception:
        return None
    m += delta
    while m < 1:
        m += 12; y -= 1
    while m > 12:
        m -= 12; y += 1
    return '%04d-%02d' % (y, m)


def _keyword_sum_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched):
    """⚡⚡⚡ 2026-08-14 关键字分组加总配对（用户方法论核心：『我是找关键字，连续多笔
    可尝试核对，大量重复先加总再在另一边找相应金额；只按名字记账方法各异防不胜防』）。
    对剩余未配行按【业务关键字组】分组（同义词归一），组内两侧【加总】比较：
      ① 同月加总相等（±1 月容差，月末汇总滞后常见）→ 整体配对消除
         （纸管：网银3笔23,650.50 = 账10笔23,650.50；汇付：网银逐笔 = 账月末备用金汇总；
          公积金：网银缴存 = 账汇总）
      ② 同月不配 → 全年加总相等 → 整体配对
      ③ 仍不配 → 小侧加总，从大侧找等额子集（_subset_sum，防组合爆炸限 60 候选）
    防误配：仅同业务组、同方向（网银收↔账借/支↔贷）配对；笔数 ≥3 才加总（大量重复）。
    """
    from collections import defaultdict
    rem_b = [i for i, b in enumerate(bank_rows)
             if i not in used_b and not b.get('_orig_dir')  # ⚡ 红字行跳过（还原语义，留红字对称配对）
             and float(b.get('amount') or 0) > 0]
    rem_j = [i for i, j in enumerate(journal_rows)
             if i not in used_j and not j.get('_orig_dir')
             and float(j.get('amount') or 0) > 0]
    g_b = defaultdict(list)
    g_j = defaultdict(list)
    for i in rem_b:
        b = bank_rows[i]
        k = _biz_key(b.get('summary') or '') or _biz_key(b.get('counterparty') or '')
        if k:
            g_b[(dir_map.get(b.get('direction'), ''), k)].append(i)
    for i in rem_j:
        j = journal_rows[i]
        k = _biz_key(j.get('summary') or '') or _biz_key(j.get('cp') or '')
        if k:
            g_j[(j['direction'], k)].append(i)
    n_pairs = 0
    for (d, k), B in g_b.items():
        J = g_j.get((d, k), [])
        if not J or len(B) + len(J) < 3:
            continue
        # ① 同月加总（±1 月）：网银逐笔 = 账侧月末汇总
        b_by_m = defaultdict(list)
        j_by_m = defaultdict(list)
        for i in B:
            b_by_m[str(bank_rows[i]['date'])[:7]].append(i)
        for i in J:
            j_by_m[str(journal_rows[i]['date'])[:7]].append(i)
        for m in [m for m in b_by_m if any(i not in used_b for i in b_by_m[m])]:
            Bm = [i for i in b_by_m[m] if i not in used_b]
            if not Bm:
                continue
            sb = sum(bank_rows[i]['amount'] for i in Bm)
            for mm in (m, _shift_month(m, -1), _shift_month(m, 1)):
                if not mm:
                    continue
                Jm = [i for i in j_by_m.get(mm, []) if i not in used_j]
                if not Jm:
                    continue
                sj = sum(journal_rows[i]['amount'] for i in Jm)
                if abs(round(sb, 2) - round(sj, 2)) < 0.01 and len(Bm) + len(Jm) >= 3:
                    for i in Bm:
                        used_b.add(i)
                    for i in Jm:
                        used_j.add(i)
                    matched.append(([bank_rows[i] for i in Bm],
                                    [journal_rows[i] for i in Jm],
                                    '关键字汇总(%d网银=%d账,%s)' % (len(Bm), len(Jm), k)))
                    n_pairs += 1
                    break
        # ② 全年加总
        B_rem = [i for i in B if i not in used_b]
        J_rem = [i for i in J if i not in used_j]
        if len(B_rem) + len(J_rem) >= 3 and B_rem and J_rem:
            sb = sum(bank_rows[i]['amount'] for i in B_rem)
            sj = sum(journal_rows[i]['amount'] for i in J_rem)
            if abs(round(sb, 2) - round(sj, 2)) < 0.01:
                for i in B_rem:
                    used_b.add(i)
                for i in J_rem:
                    used_j.add(i)
                matched.append(([bank_rows[i] for i in B_rem],
                                [journal_rows[i] for i in J_rem],
                                '关键字汇总(全年%d网银=%d账,%s)' % (len(B_rem), len(J_rem), k)))
                n_pairs += 1
                continue
            # ③ 小侧加总 = 大侧子集找和（防组合爆炸：大侧金额排序取前 60）
            if len(B_rem) <= len(J_rem):
                small, big = B_rem, J_rem
                small_sum = round(sum(bank_rows[i]['amount'] for i in small), 2)
                big_rows = journal_rows
            else:
                small, big = J_rem, B_rem
                small_sum = round(sum(journal_rows[i]['amount'] for i in small), 2)
                big_rows = bank_rows
            big_sorted = sorted(big, key=lambda i: -big_rows[i]['amount'])[:60]
            hit = _subset_sum(big_sorted, big_rows, small_sum)
            if hit and len(hit) < len(big):
                hit_set = set(hit)
                if big_rows is journal_rows:
                    used_j.update(hit_set)
                    used_b.update(small)
                    matched.append(([bank_rows[i] for i in small],
                                    [journal_rows[i] for i in hit],
                                    '关键字找和(%d网银=%d账,%s)' % (len(small), len(hit), k)))
                else:
                    used_b.update(hit_set)
                    used_j.update(small)
                    matched.append(([bank_rows[i] for i in hit],
                                    [journal_rows[i] for i in small],
                                    '关键字找和(%d网银=%d账,%s)' % (len(hit), len(small), k)))
                n_pairs += 1

    # ④ 账侧关键词 → 网银侧无关键词等额批次（单位倍数配对，用户 21:12 场景）：
    #    账侧"国望后勤押金退款 1,035=45×23" vs 网银 45 元"电子转账"批次（summary 无业务词）——
    #    从网银侧同日(±30天)同方向取 N=账侧金额/K 笔配对（剩余批次保留为其他业务）。
    #    ⚡⚡ 2026-08-14 大额保护：仅账侧金额 ≤10 万元的小额重复业务才做单位倍数配对
    #    （45 元押金/工资/小额手续费）；大额资金池划转（实归/支取 1000 万级）配对风险高
    #    （网银 N 笔×K = 账侧 1 笔可能是不同业务——3100 建行实证：单位倍数配对 6笔×1500万
    #    把资金池支配给错误账行，月度对账支差暴增 5 亿）。
    from collections import defaultdict as _dd2
    b_rem_now = [i for i in rem_b if i not in used_b]
    b_by_amt = _dd2(list)
    for i in b_rem_now:
        b_by_amt[round(bank_rows[i]['amount'], 2)].append(i)
    for ji in list(rem_j):
        if ji in used_j:
            continue
        j = journal_rows[ji]
        k = _biz_key(j.get('summary') or '') or _biz_key(j.get('cp') or '')
        if not k:
            continue
        jd = _dday(j['date'])
        if jd is None:
            continue
        amt = round(j['amount'], 2)
        if amt <= 0 or amt > 100000:  # ⚡ 仅小额重复业务（>10 万不做单位倍数配对）
            continue
        # 从网银侧找等额批次：K×N = amt（N≥2 大量重复），同日窗口内取 N 笔
        for K, pool_all in sorted(b_by_amt.items(), key=lambda x: -len(x[1])):
            if K <= 0 or len(pool_all) < 2 or amt % K != 0:
                continue
            N = int(round(amt / K))
            if N < 2 or N > 100:  # 工资代发 50 人=50 笔批次实证，上限 100 防过度
                continue
            pool = [i for i in pool_all
                    if i not in used_b
                    and dir_map.get(bank_rows[i]['direction'], '') == j['direction']
                    and _dday(bank_rows[i]['date']) is not None
                    and abs((_dday(bank_rows[i]['date']) - jd).days) <= 30]
            if len(pool) >= N:
                pick = pool[:N]
                for i in pick:
                    used_b.add(i)
                used_j.add(ji)
                matched.append(([bank_rows[i] for i in pick], j,
                                '单位倍数配对(%d笔×%.2f,%s)' % (N, K, k)))
                n_pairs += 1
                break
    return n_pairs


# ⚡⚡⚡ 2026-08-14 语义簇自动发现（用户方法论 22:19：『不要只记关键字，每个账套可能都不一样，
#   要学会大量加总后，怎么去找差不多的意思的关键字』）：
#   配置种子（bank_business_keywords.json _default）只是【已知业务快配】，本函数自动学习新账套的业务词簇：
#   ①提取剩余未配行的高频业务词（2-4 字，过滤机构/银行通用后缀）
#   ②网银词 w ↔ 账侧词 v【同方向 + 笔数≥3（大量小额）+ 合计相等】→ 判定同一业务簇，加总配对。
#   字面不同（医疗保障↔生育津贴、汇付↔备用金）但金额总量相等 = 同一业务——靠金额桥接语义。
#   ③迭代 3 轮收敛：配对后剩余变化，再发现新簇（用户：『配对后剩下数量少了，也更容易匹配了，
#   可以再用同样的方法仔细查找』）。
_AUTO_STOPWORDS = set(
    '有限公司 股份有限公司 有限责任公司 电子 转账 汇入 汇出 银行 账户 往来 款项 支付 人民币 '
    '结算 网银 跨行 同城 划转 管理 中心 股份 有限 责任 公司 支行 分行 营业部 汇划 电汇 本行 '
    '跨行转账 收到 付出 转入 转出 收入 支出'.split())


def _auto_biz_words(text):
    """提取业务特征词（2-4 字 n-gram，过滤停用词/机构后缀）。"""
    import re as _re
    s = str(text or '')
    s = _re.sub(r'[^\u4e00-\u9fa5A-Za-z0-9]', '', s)
    out = set()
    for k in (2, 3, 4):
        for i in range(len(s) - k + 1):
            w = s[i:i + k]
            if not _re.search(r'[\u4e00-\u9fa5]', w):
                continue
            if any(st in w for st in _AUTO_STOPWORDS):
                continue
            out.add(w)
    return out


def _auto_cluster_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched):
    """语义簇自动发现加总配对（用户方法论 22:19：大量小额一般不舞弊，加总消除后剩余更少更好配）。
    在 _keyword_sum_match（硬编码种子表）之后跑，兜底未覆盖的新账套业务词。"""
    from collections import defaultdict
    n_pairs = 0
    for _round in range(3):  # 迭代收敛：配对后剩余词频变化，最多 3 轮
        rem_b = [i for i, b in enumerate(bank_rows)
                 if i not in used_b and not b.get('_orig_dir')
                 and float(b.get('amount') or 0) > 0]
        rem_j = [i for i, j in enumerate(journal_rows)
                 if i not in used_j and not j.get('_orig_dir')
                 and float(j.get('amount') or 0) > 0]
        if len(rem_b) < 3 or len(rem_j) < 3:
            break
        b_w = defaultdict(list)
        j_w = defaultdict(list)
        for i in rem_b:
            for w in _auto_biz_words((bank_rows[i].get('counterparty') or '')
                                     + ' ' + (bank_rows[i].get('summary') or '')):
                b_w[w].append(i)
        for i in rem_j:
            for w in _auto_biz_words((journal_rows[i].get('summary') or '')
                                     + ' ' + str(journal_rows[i].get('cp') or '')):
                j_w[w].append(i)
        # 高频词（≥3 笔）才有业务簇意义（大量小额）
        hot_b = {w: set(v) for w, v in b_w.items() if len(set(v)) >= 3}
        hot_j = {w: set(v) for w, v in j_w.items() if len(set(v)) >= 3}
        if not hot_b or not hot_j:
            break
        any_pair = False
        for w, bidx in sorted(hot_b.items(), key=lambda x: -len(x[1])):
            bidx_u = [i for i in bidx if i not in used_b]
            if len(bidx_u) < 3:
                continue
            d0 = dir_map.get(bank_rows[bidx_u[0]]['direction'], '')
            sb = round(sum(bank_rows[i]['amount'] for i in bidx_u), 2)
            for v, jidx in sorted(hot_j.items(), key=lambda x: -len(x[1])):
                jidx_u = [i for i in jidx if i not in used_j]
                if len(jidx_u) < 3:
                    continue
                if journal_rows[jidx_u[0]]['direction'] != d0:
                    continue
                sj = round(sum(journal_rows[i]['amount'] for i in jidx_u), 2)
                if abs(sb - sj) < 0.02:
                    for i in bidx_u:
                        used_b.add(i)
                    for i in jidx_u:
                        used_j.add(i)
                    matched.append(([bank_rows[i] for i in bidx_u],
                                    [journal_rows[i] for i in jidx_u],
                                    '语义簇加总(%d网银=%d账,%s↔%s)' % (len(bidx_u), len(jidx_u), w, v)))
                    n_pairs += 1
                    any_pair = True
                    break
        if not any_pair:
            break
    return n_pairs


def _cumulative_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched):
    """累计匹配：按 (方向映射, 月份) 分组，组内按日期排序，逐笔累计差额，
    累计差归零处闭合一段（该段内 网银↔账 互配，N:M 任意组合）。
    解决「多条网银=一条汇总日记账」「一条=多条拆分」「N:M」——银行对账经典累计法。
    ⚡⚡ 2026-08-14 修复：组内按【摘要关键词】再细分（纸管/公积金/实归支取各算一组）——
    原全月同方向混配，纸管账侧 10 笔被累计匹配配给非纸管网银组合（3100 农行实证），
    导致纸管网银 3 笔落空。同摘要组内累计=不串业务。"""
    from collections import defaultdict
    groups = defaultdict(lambda: {'b': [], 'j': []})

    def _kw(text):
        # ⚡⚡⚡ 2026-08-14 统一业务关键词：改用模块级 _biz_key（含汇付/后勤/医保/E信/公积金/
        #   纸管/社保/工资等 13 组）——原硬编码 8 词表缺汇付等，网银汇付归"_其他"组与全月
        #   其他业务混配（3100 建行汇付 32 笔 vs 账侧 3,477.93 配对失效）。
        return _biz_key(text) or '_其他'

    for i, b in enumerate(bank_rows):
        if i in used_b:
            continue
        m = str(b['date'])[:7]
        # ⚡⚡ 2026-08-14 只取第一个关键词（不拼接 cp）——网银 summary='纸管差价结算_.'+
        #   counterparty='苏州盛虹'→'纸管_其他'，账侧='纸管纸管'，key 不同被分到两组的 bug
        _bk = _kw(b.get('summary') or '') or _kw(b.get('counterparty') or '')
        groups[(dir_map.get(b['direction'], ''), m, _bk)]['b'].append(i)
    for jj, j in enumerate(journal_rows):
        if jj in used_j:
            continue
        m = str(j['date'])[:7]
        _jk = _kw(j.get('summary') or '') or _kw(j.get('cp') or '')
        groups[(j['direction'], m, _jk)]['j'].append(jj)
    for (d, m, _g), g in groups.items():
        B = sorted(g['b'], key=lambda i: bank_rows[i]['date'])
        J = sorted(g['j'], key=lambda jj: journal_rows[jj]['date'])
        sb = sj = 0.0
        seg_b, seg_j = [], []
        i = j = 0
        while i < len(B) or j < len(J):
            if i < len(B) and (sb <= sj or j >= len(J)):
                sb += round(bank_rows[B[i]]['amount'], 2)
                seg_b.append(B[i]); i += 1
            elif j < len(J):
                sj += round(journal_rows[J[j]]['amount'], 2)
                seg_j.append(J[j]); j += 1
            if abs(sb - sj) < 0.01 and (seg_b or seg_j):
                if seg_b and seg_j and abs(sb) > 0.01:
                    for bi in seg_b:
                        used_b.add(bi)
                    for jj in seg_j:
                        used_j.add(jj)
                    matched.append(([bank_rows[x] for x in seg_b],
                                    [journal_rows[x] for x in seg_j],
                                    '累计匹配(%d网银=%d账)' % (len(seg_b), len(seg_j))))
                sb = sj = 0.0
                seg_b, seg_j = [], []


def _tokens(text, n=3):
    """文本 3-gram 词集（去括号/去非中文数字字母），用于名称相似度比对。"""
    s = str(text or '')
    s = re.sub(r'[（(].*?[)）]', '', s)      # 去括号内
    s = re.sub(r'[^\u4e00-\u9fa5A-Za-z0-9]', '', s)
    out = set()
    for k in range(len(s) - n + 1):
        out.add(s[k:k + n])
    return out


def _counterparty_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched):
    """对方关键词关联匹配：网银 counterparty 与账 cp/摘要共享关键词（≥3 字）分组的
    未配行，组内按累计法匹配（如『星星充电』：电商管家平台交易资金 ↔ 收星星充电桩电费款）。
    防误配：仅当组内两边金额合计相等才整体匹配。"""
    from collections import defaultdict
    rem_b = [(i, b) for i, b in enumerate(bank_rows) if i not in used_b]
    rem_j = [(i, j) for i, j in enumerate(journal_rows) if i not in used_j]

    for i, b in rem_b:
        if i in used_b:
            continue
        # ⚡⚡ 2026-08-14 修复：token 须合并 counterparty + summary（原 or 短路只取
        #   counterparty——网银『苏州盛虹纤维 纸管差价结算』漏掉摘要词，与账侧
        #   『纸管差价结算』零共享 → 纸管/公积金 N:M 汇总配不上，3100 农行实证）
        bt = _tokens((b.get('counterparty') or '') + ' ' + (b.get('summary') or ''))
        if not bt:
            continue
        for jj, j in rem_j:
            if jj in used_j:
                continue
            jt = _tokens(j.get('summary') or '') | _tokens(j.get('cp') or '')
            common = bt & jt
            if len(common) < 2:     # 需 ≥2 个 3 字元组共现（防泛化误配）
                continue
            if dir_map.get(b['direction'], '') != j['direction']:
                continue
            # 同关键词组：找所有共享该词的未配行，组内累计匹配
            group_b = [x for x, bb in rem_b if x not in used_b
                       and dir_map.get(bb['direction'], '') == j['direction']
                       and (bt & _tokens((bb.get('counterparty') or '') + ' ' + (bb.get('summary') or '')))]
            group_j = [x for x, jj2 in rem_j if x not in used_j
                       and jj2['direction'] == j['direction']
                       and (jt & _tokens(jj2.get('summary') or '') | _tokens(jj2.get('cp') or ''))]
            if not group_b or not group_j:
                continue
            # ⚡⚡ 2026-08-14 月份约束：网银组与账组须【同月或相邻月】——纸管 5-26 vs 5-31
            #   是月内汇总；防不同月同额误配（3100 农行纸管/公积金实证）
            _mb = [str(bank_rows[x].get('date') or '')[:7] for x in group_b]
            _mj = [str(journal_rows[x].get('date') or '')[:7] for x in group_j]
            _mb = {m for m in _mb if m}
            _mj = {m for m in _mj if m}
            if _mb and _mj and not (_mb & _mj):
                # 相邻月也允许（月末集中记账：网银 5-26 账 5-31 同月；跨月 4-28 账 5-02）
                _ok = False
                for _b_m in _mb:
                    for _j_m in _mj:
                        try:
                            _bm_y, _bm_m = int(_b_m[:4]), int(_b_m[5:7])
                            _jm_y, _jm_m = int(_j_m[:4]), int(_j_m[5:7])
                            if (_bm_y, _bm_m) == (_jm_y, _jm_m) or \
                               abs((_bm_y * 12 + _bm_m) - (_jm_y * 12 + _jm_m)) == 1:
                                _ok = True
                                break
                        except Exception:
                            _ok = True
                            break
                    if _ok:
                        break
                if not _ok:
                    continue
            sb = sum(bank_rows[x]['amount'] for x in group_b)
            sj = sum(journal_rows[x]['amount'] for x in group_j)
            if abs(round(sb, 2) - round(sj, 2)) < 0.02:
                for x in group_b:
                    used_b.add(x)
                for x in group_j:
                    used_j.add(x)
                matched.append(([bank_rows[x] for x in group_b],
                                [journal_rows[x] for x in group_j],
                                '对方匹配(%d网银=%d账)' % (len(group_b), len(group_j))))
            break


def reconcile(bank_rows, journal_rows, window=15):
    """双向核对单银行。返回 (matched, unmatched_bank, unmatched_journal, summary)。
    bank_rows: 网银统一行；journal_rows: 日记账行（direction 借/贷）。
    窗口默认 15 天（月末在途/跨期入账常见，9100 万曾差 9 天）。"""
    dir_map = {'收': '借', '支': '贷'}
    # ⚡⚡ 2026-08-14 红字【对称预配对】：网银『支 -X』（红字冲正/退款）与账侧『贷 -X』
    #   （当日冲正/有误冲正）是同一笔冲正业务——两侧同向同额（负），应配对消除。
    #   实证（DQ 嘉兴银行）：网银『北京朗博 支-6000』(1-16) = 账『1.16有误冲正 贷-6000』，
    #   10 笔红字 9 笔两侧完全一致却全进差异清单（先做对称预配对，防下方归一化拆散）。
    _neg_pairs = 0
    _neg_b = [(i, x) for i, x in enumerate(bank_rows) if float(x.get('amount') or 0) < 0]
    _neg_j = [(i, x) for i, x in enumerate(journal_rows) if float(x.get('amount') or 0) < 0]
    if _neg_b and _neg_j:
        _nj_by = {}
        for _ji, _jx in _neg_j:
            # ⚡ 方向归一化：网银『支』= 账『贷』（dir_map 收→借 / 支→贷）
            _nj_by.setdefault((dir_map.get(_jx['direction'], _jx['direction']),
                               round(float(_jx['amount']), 2)), []).append(_ji)
        _used_neg_b = set()
        _used_neg_j = set()
        _neg_matched = []
        for _bi, _bx in _neg_b:
            _k = (dir_map.get(_bx['direction'], _bx['direction']),
                  round(float(_bx['amount']), 2))
            for _ji in _nj_by.get(_k, []):
                if _ji in _used_neg_j:
                    continue
                _used_neg_b.add(_bi); _used_neg_j.add(_ji)
                _neg_matched.append((_bi, _ji))
                break
        if _neg_matched:
            _neg_pairs = len(_neg_matched)
            used_b = set(_bi for _bi, _ in _neg_matched)
            used_j = set(_ji for _, _ji in _neg_matched)
            matched = [(bank_rows[_bi], journal_rows[_ji], '红字对称冲正')
                       for _bi, _ji in _neg_matched]
        else:
            used_b = set()
            used_j = set()
            matched = []
    else:
        used_b = set()
        used_j = set()
        matched = []
        _used_neg_b = set()
        _used_neg_j = set()

    # ⚡ 2026-08-13 红字负数归一化：网银『收 -X』（红字冲正/退款）=净流出 X=『支 X』；
    #    账侧红字以贷方体现（农商行 7-06 收 -89.78 一码通冲正 ↔ 账 7-22 贷 89.78 配对实证，
    #    7 月收差=支差=-89.78 由此对平）。负数行翻方向+取绝对值参与匹配；原值存
    #    _orig_dir/_orig_amt 供渲染还原（红字在底稿中仍显示为『收 -X』）。
    #    ⚡⚡ 2026-08-14：跳过已对称配对的红字行（保持原样，不翻方向）。
    for _r in list(bank_rows) + list(journal_rows):
        if _r.get('amount') is not None and float(_r['amount']) < 0:
            if id(_r) in {id(bank_rows[i]) for i in _used_neg_b} | {id(journal_rows[i]) for i in _used_neg_j}:
                continue
            _r['_orig_dir'] = _r['direction']
            _r['_orig_amt'] = float(_r['amount'])
            _r['direction'] = '支' if _r['direction'] == '收' else '收'
            _r['amount'] = -float(_r['amount'])

    # ── 1. 精确匹配（同日优先）：先『同日』→ 再『摘要业务日期』（DQ 月末集中记账滞后）→
    #      再『15 天窗口』。⚠️ 2026-08-13 重构：摘要业务日期匹配必须插在 15 天窗口【之前】，
    #      否则同月多笔同额跨行资金被窗口猜测抢配错位（DQ 中行 2-11 网银支 1000 万被配到
    #      2.24 账的实证：账侧『2.11/2.24中行转入嘉兴』贷 4 笔 vs 网银 2-11/2-24 支 4 笔，
    #      窗口先配 → 错位 → 假差异）。 ──
    j_by_key = {}
    for ji, j in enumerate(journal_rows):
        j_by_key.setdefault((j['direction'], round(j['amount'], 2)), []).append(ji)

    # ── 1. 精确匹配（同日优先）：先『同日』→ 再『摘要业务日期』（DQ 月末集中记账滞后）→
    #      再『15 天窗口』。⚠️ 2026-08-13 重构：摘要业务日期匹配必须插在 15 天窗口【之前】，
    #      否则同月多笔同额跨行资金被窗口猜测抢配错位（DQ 中行 2-11 网银支 1000 万被配到
    #      2.24 账的实证：账侧『2.11/2.24中行转入嘉兴』贷 4 笔 vs 网银 2-11/2-24 支 4 笔，
    #      窗口先配 → 错位 → 假差异）。 ──
    # 1a. 同日精确（最高置信：网银日期=账记账日期）
    # ⚡⚡ 2026-08-14 名称一致性优先：同日同额候选 >1 笔时，按网银对方名 ↔ 日记账
    #   摘要/对方科目【共享词】优先配对（DQ 实证：同日多笔同额收款——网银『浙江方圆
    #   电气 IBP』被配给『苏州源丽检测』账行，交叉错配；共享词配对消除假匹配）。
    for bi, b in enumerate(bank_rows):
        if bi in used_b:
            continue
        dkey = (dir_map.get(b['direction'], ''), round(b['amount'], 2))
        bd = _dday(b['date'])
        cands = [ji for ji in j_by_key.get(dkey, [])
                 if ji not in used_j and _dday(journal_rows[ji]['date']) and
                 (bd - _dday(journal_rows[ji]['date'])).days == 0]
        if not cands:
            continue
        # ⚡⚡⚡ 2026-08-14 业务批次保护（用户方法论：找关键字而非名字）：网银行有业务关键词
        #   （汇付/公积金/纸管/后勤/医保/E信…）时，候选账行必须【同业务关键词】才配——
        #   否则网银汇付逐笔（119.46 等）被 1a 错配给非汇付账行，业务批次被拆散，第7步
        #   关键字加总配对失效（3100 建行实证：1 月汇付 32 笔被配走 16 笔 → 残留加总对不上）。
        #   同业务词的正常单笔（公积金缴存 1,120=账 1,120）不受影响，仍走 1a。
        _bk = _biz_key(b.get('summary') or '') or _biz_key(b.get('counterparty') or '')
        if _bk:
            _cands_bk = [ji for ji in cands
                         if _biz_key(journal_rows[ji].get('summary') or '')
                         == _bk]
            if not _cands_bk:
                continue
            cands = _cands_bk
        if len(cands) == 1:
            used_b.add(bi); used_j.add(cands[0])
            matched.append((b, journal_rows[cands[0]], '精确匹配(同日)'))
            continue
        # 多候选：按名称共享词打分（网银对方名 vs 账摘要/对方科目）
        # ⚡⚡ 2026-08-14 修复：best_score=0（候选与网银零共享词）不配对——网银纸管
        #   18513.30 同日候选是『实归支取』（共享词0），配对=错配；留给累计/对方匹配。
        b_toks = _tokens((b.get('counterparty') or '') + ' ' + (b.get('summary') or ''))
        best, best_score = None, -1
        for ji in cands:
            j = journal_rows[ji]
            j_toks = _tokens((j.get('summary') or '') + ' ' + (j.get('cp') or ''))
            score = len(b_toks & j_toks)
            if score > best_score:
                best, best_score = ji, score
        if best is not None and best_score > 0:
            used_b.add(bi); used_j.add(best)
            tag = '精确匹配(同日)' + ('·名符' if best_score > 0 else '·金额')
            matched.append((b, journal_rows[best], tag))
    # 1b. 摘要业务日期匹配（DQ/U8 月末集中记账特有 2026-08-13）：账侧摘要开头『X.YZ』
    #     =业务日期（如『2.11中国银行转入嘉兴银行1000万』，记账日=月末 2-28、业务日=2.11），
    #     **网银日期月日 == 摘要业务月日** → 同笔（滞后记账权威依据）。
    #     仅限：同方向+同金额+网银日期=摘要业务日期；3100 SAP 摘要无 X.Y 格式不受影响。
    rem_b1b = [(i, b) for i, b in enumerate(bank_rows) if i not in used_b]
    rem_j1b = [(i, j) for i, j in enumerate(journal_rows) if i not in used_j]
    for bi, b in rem_b1b:
        if bi in used_b:
            continue
        bd = _dday(b['date'])
        if bd is None:
            continue
        for ji, j in rem_j1b:
            if ji in used_j:
                continue
            if j['direction'] != dir_map.get(b['direction'], '') \
                    or round(j['amount'], 2) != round(b['amount'], 2):
                continue
            biz = _summary_bizdate(j.get('summary'))
            if biz and (bd.month, bd.day) == biz:
                used_b.add(bi)
                used_j.add(ji)
                matched.append((b, j, '摘要业务日期匹配(月末集中记账,%d/%d)' % biz))
                break
    # 1c. 15 天窗口精确（月末在途/跨期入账常见，9100 万曾差 9 天）
    # ⚡⚡ 2026-08-14 名称优先修复（3100 农行纸管实证）：网银『纸管差价结算 18,513.30』
    #   与账侧『实归支取 18,513.30』同日同额，窗口按序误配——真正对应是 5-31 纸管汇总。
    #   多候选时：①同日优先（先于跨日）②共享词打分（防错配）。
    for bi, b in enumerate(bank_rows):
        if bi in used_b:
            continue
        dkey = (dir_map.get(b['direction'], ''), round(b['amount'], 2))
        bd = _dday(b['date'])
        cands = [ji for ji in j_by_key.get(dkey, [])
                 if ji not in used_j and _dday(journal_rows[ji]['date']) and
                 bd and abs((bd - _dday(journal_rows[ji]['date'])).days) <= window]
        if not cands:
            continue
        # ⚡⚡⚡ 2026-08-14 业务批次保护（同 1a）：网银有业务关键词时，窗口候选账行必须同业务词
        _bk = _biz_key(b.get('summary') or '') or _biz_key(b.get('counterparty') or '')
        if _bk:
            _cands_bk = [ji for ji in cands
                         if _biz_key(journal_rows[ji].get('summary') or '')
                         == _bk]
            if not _cands_bk:
                continue
            cands = _cands_bk
        # 同日候选优先
        same_day = [ji for ji in cands if (bd - _dday(journal_rows[ji]['date'])).days == 0]
        pool = same_day if same_day else cands
        if len(pool) == 1:
            ji = pool[0]
            used_b.add(bi); used_j.add(ji)
            jd = _dday(journal_rows[ji]['date'])
            if bd.strftime('%Y-%m') != jd.strftime('%Y-%m'):
                matched.append((b, journal_rows[ji], '跨期匹配(疑似未达账项,差%d天)' % (bd - jd).days))
            else:
                matched.append((b, journal_rows[ji], '精确匹配'))
            continue
        # 多候选：名称共享词打分（网银 cp+sm vs 账 cp+sm），最高分优先
        # ⚡⚡ 2026-08-14 修复：score=0 不配对（防错配，留给累计/对方匹配）
        b_toks = _tokens((b.get('counterparty') or '') + ' ' + (b.get('summary') or ''))
        best, best_score = None, -1
        for ji in pool:
            j = journal_rows[ji]
            j_toks = _tokens((j.get('summary') or '') + ' ' + (j.get('cp') or ''))
            score = len(b_toks & j_toks)
            if score > best_score:
                best, best_score = ji, score
        if best is not None and best_score > 0:
            used_b.add(bi); used_j.add(best)
            jd = _dday(journal_rows[best]['date'])
            tag = '精确匹配(窗口·名符)' if best_score > 0 else '精确匹配(窗口)'
            if bd.strftime('%Y-%m') != jd.strftime('%Y-%m'):
                tag = '跨期匹配(名符,差%d天)' % (bd - jd).days
            matched.append((b, journal_rows[best], tag))

    # ── 2. 累计匹配（核心改进，2026-08-13）：同月同方向按日期排序，累计差归零处
    #        闭合一段（N:M）——解决「多条网银=一条汇总日记账」「一条=多条拆分」「N:M」。
    #        用户实测：总数对得上但精确匹配找不到 → 多为汇总/拆分记账，累计法正解。 ──
    _cumulative_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched)

    # ── 3. 对方关键词关联匹配：剩余网银 counterparty 与剩余账 cp/摘要共享关键词分组，
    #        组内累计匹配（如『星星充电』= 电商管家平台交易资金 ↔ 收星星充电桩电费款） ──
    _counterparty_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched)

    # ── 4. 汇总匹配：多条网银 = 一条日记账（同方向 + 同月聚合） ──
    #    星星充电场景实测：网银逐笔小额（01-05 多笔），账上月末汇总一笔（01-31 1232.35）
    rem_b = [(i, b) for i, b in enumerate(bank_rows) if i not in used_b]
    rem_j = [(i, j) for i, j in enumerate(journal_rows) if i not in used_j]
    # 2a. 同月聚合：按 (方向, YYYY-MM) 汇总【全集】网银/日记账金额，和相等即整体匹配
    # ⚡ 2026-08-13 改用全集而非 rem：前 3 步（精确/累计/对方科目）可能已配掉部分行，rem 里
    #   同月同方向反而不平衡（如 1 月收：精确配掉网银收 X 但账借 Y 未配，X≠Y → bsum≠jsum
    #   → 月汇总不触发 → 整月一码通落空 ub——农商行 1 月收差 0 但 ub 挂 15 笔一码通实证）。
    #   全集判断月平衡（收差=0）→ 配对该月全部未配行，既正确又干净。
    # ⚡⚡ 2026-08-13 红字归一化兼容：负数归一化（收 -X → 支 +X）会破坏月平衡判定（收侧少
    #   44.89、支侧多 44.89 → 两边都 skip → 整月一码通落空）。故：①平衡判定用【还原值】
    #   （_orig_dir/_orig_amt，红字归位）②配对 new_b/new_j 排除红字行（_orig_dir 存在）——
    #   红字只由精确/跨期/代扣匹配找对应贷行（如 7 月 -89.78 ↔ 账贷 89.78），找不到即真实
    #   差异（还原显示『收 -X』）。
    from collections import defaultdict
    b_mon = defaultdict(float)
    for i, b in enumerate(bank_rows):
        m = str(b['date'])[:7]
        if m:
            # ⚡⚡⚡ 2026-08-14 业务批次保护：有业务关键词的行（汇付/公积金/纸管/后勤/医保/
            #   E信/社保/工资/手续费…）不参与全集月汇总——否则月汇总把网银汇付逐笔错配给
            #   非汇付账行（3100 建行实证 152 笔被配走、账侧汇付 5 笔全留 uj）；业务行
            #   完整保留给第 7 步关键字加总配对（同业务组加总相等，语义最准）。
            if _biz_key(b.get('summary') or '') or _biz_key(b.get('counterparty') or ''):
                continue
            d = b.get('_orig_dir', b['direction'])
            a = b.get('_orig_amt', b['amount'])
            b_mon[(dir_map.get(d, ''), m)] += round(a, 2)
    j_mon = defaultdict(float)
    for i, j in enumerate(journal_rows):
        m = str(j['date'])[:7]
        if m:
            if _biz_key(j.get('summary') or '') or _biz_key(j.get('cp') or ''):
                continue
            j_mon[(j['direction'], m)] += round(j['amount'], 2)
    used_mon_b = set()
    used_mon_j = set()
    for (dirv, m), jsum in j_mon.items():
        bsum = b_mon.get((dirv, m), 0.0)
        # ⚡ 2026-08-13 尾差判定：手续费双计已由 _fee_net_journal 净额化消除、红字冲正由
        #   负数归一化配对（未配对的还原为『收 -X』真实差异），月合计差异只应≈0。
        #   原 0.5% 容忍把真实差异当尾差吞掉（农商行 1/2 月红字 +44.89 → 差异清单空但
        #   月度对账显示差，两表打架）；回到 0.02 严格判定，残余小额作为真实尾差显示。
        tol = 0.02
        if abs(bsum - jsum) < tol:
            # ⚡ 2026-08-13 修复：used_mon_b/used_mon_j 累积导致 matched 重复记录——
            #   每次只匹配【本次 (dirv,m)】的行（原代码 append 全部累积集合，之前匹配过的
            #   行会重复进 matched，建行实证 mb 4730 行 vs 去重后 4667 行、重复金额 1,221.11）
            new_b = {i for i, b in enumerate(bank_rows)
                     if i not in used_b
                     and not b.get('_orig_dir')  # ⚡ 红字行不参与月汇总（留给精确/跨期）
                     and not (_biz_key(b.get('summary') or '') or _biz_key(b.get('counterparty') or ''))  # ⚡ 业务批次保护
                     and str(b['date'])[:7] == m
                     and dir_map.get(b['direction'], '') == dirv}
            new_j = {i for i, j in enumerate(journal_rows)
                     if i not in used_j
                     and not (_biz_key(j.get('summary') or '') or _biz_key(j.get('cp') or ''))
                     and str(j['date'])[:7] == m
                     and j['direction'] == dirv}
            if new_b or new_j:
                matched.append(([bank_rows[i] for i in new_b],
                                [journal_rows[i] for i in new_j],
                                '月汇总匹配'))
                used_mon_b |= new_b
                used_mon_j |= new_j
    # 2b. 单笔找和（同方向 + 日期窗口内 DFS 组合）——残余场景
    rem_b2 = [(i, b) for i, b in rem_b if i not in used_mon_b]
    rem_j2 = [(i, j) for i, j in rem_j if i not in used_mon_j]
    for ji, j in rem_j2:
        jd = _dday(j['date'])
        if jd is None:
            continue
        # ⚡ 2026-08-13 修复：pool 必须排除【已 used 网银行】——原代码漏滤导致同一小额网银行
        #   （如广州合利宝/上海汇付备付金）被 _find_sum 重复选中匹配给多笔账行 → used_j 多标记
        #   → uj（账有网银无）低估、差异清单与月度对账打架（3100 建行实证 63 行重复、差 1,221.11）
        pool = [(i, b) for i, b in rem_b2
                if i not in used_b
                and dir_map.get(b['direction'], '') == j['direction']
                and _dday(b['date']) is not None
                and abs((_dday(b['date']) - jd).days) <= window]
        target = round(j['amount'], 2)
        # ⚡ 2026-08-14 多笔网银=1笔账 找和放宽到 25 笔（3300 建行实证：2-24 网银支
        #   12×800万+216万=9816 万 ↔ 账 2-28 贷 9816 万承兑汇票到期，13 笔组合原 max_depth=5
        #   配不上 → 收差 +1030 万假差异）；_find_sum 内部候选裁剪+节点预算防爆炸）
        combo = _find_sum(pool, target, max_depth=25)
        if combo:
            for i, b in combo:
                used_b.add(i)
            used_j.add(ji)
            matched.append(([b for _, b in combo], j, '汇总匹配(%d条网银)' % len(combo)))

    for i, b in rem_b:
        if i in used_mon_b:
            used_b.add(i)
    for i, j in rem_j:
        if i in used_mon_j:
            used_j.add(i)

    # ── 5. 跨期未达账项查找（用户方法论 2026-08-13）：记账时间差导致本月配不上，
    #       自动去相邻月份（跨月，日期差 ≤ 45 天）找同方向同金额记录——找到即确认
    #       「疑似未达账项（跨期）」，从差异清单移出，标注原因。 ──
    rem_b5 = [(i, b) for i, b in enumerate(bank_rows) if i not in used_b]
    rem_j5 = [(i, j) for i, j in enumerate(journal_rows) if i not in used_j]
    j5_by_key = {}
    for ji, j in rem_j5:
        j5_by_key.setdefault((j['direction'], round(j['amount'], 2)), []).append(ji)
    for bi, b in rem_b5:
        if bi in used_b:
            continue
        dkey = (dir_map.get(b['direction'], ''), round(b['amount'], 2))
        bd = _dday(b['date'])
        if bd is None:
            continue
        # ⚡⚡⚡ 2026-08-14 业务批次保护（同 1a/1c）：网银有业务关键词时，跨期候选账行须同业务词
        _bk = _biz_key(b.get('summary') or '') or _biz_key(b.get('counterparty') or '')
        for ji in j5_by_key.get(dkey, []):
            if ji in used_j:
                continue
            if _bk and _biz_key(journal_rows[ji].get('summary') or '') != _bk:
                continue
            jd = _dday(journal_rows[ji]['date'])
            if jd is None:
                continue
            days = (bd - jd).days
            # ⚡ 2026-08-13 DQ 实证：账侧【月末统一记账】（摘要含业务日期 X.YZ，如『2.10已发
            #   职工1月工资』/『8.12中国银行转入交通银行1000』），网银业务日 vs 账记账日**同月**
            #   差 16-30 天（2-10 业务 / 2-28 记账）——原『跨月才找』漏配（DQ 中行支差 500 万
            #   =工资/社保/跨行转账 40 笔滞后记账；交行 3500 万差 16 天、宁波 3000 万差 20 天）。
            #   放宽为 日期差 ≤45 天（同月跨月均可）：跨月=跨期未达，同月=月末集中记账滞后。
            if abs(days) <= 45:
                used_b.add(bi)
                used_j.add(ji)
                matched.append((b, journal_rows[ji],
                                '跨期未达账项(相差%d天,%s)' % (days, '银行先记' if days > 0 else '企业先记')))
                break

    # ── 6. 代扣代缴通道语义匹配（2026-08-13）：网银『待报解/代理国库/收缴』与
    #       账『扣缴/缴/社保/个税/税/公积金』是同一笔代扣（网银显示收缴通道名、
    #       账显示具体税种/社保），摘要不同但金额+方向+日期窗口一致 → 配对
    #       ⚡ 2026-08-14 窗口 15→45：与跨期未达统一——待报解=网银扣款日 vs 账所属期记账
    #       （3100 农商行 7 月待报解 709 万 7-13 扣款 ↔ 账 7-31 社保记账差 18 天，原 15 天
    #       卡住 → 709 万对称挂红区假差异；DQ 月末集中记账普遍差 16-30 天）──
    rem_b6 = [(i, b) for i, b in enumerate(bank_rows) if i not in used_b]
    rem_j6 = [(i, j) for i, j in enumerate(journal_rows) if i not in used_j]
    b6 = [(i, b) for i, b in rem_b6
          if any(k in (b.get('counterparty') or '') + (b.get('summary') or '')
                 for k in ('待报解', '代理国库', '收缴', '财政'))]
    j6 = [(i, j) for i, j in rem_j6
          if any(k in (j.get('summary', '') + ' ' + str(j.get('cp') or ''))
                 for k in ('扣缴', '缴', '社保', '个税', '公积金', '税'))]
    j6_by_key = {}
    for ji, j in j6:
        j6_by_key.setdefault((j['direction'], round(j['amount'], 2)), []).append(ji)
    for bi, b in b6:
        if bi in used_b:
            continue
        dkey = (dir_map.get(b['direction'], ''), round(b['amount'], 2))
        bd = _dday(b['date'])
        if bd is None:
            continue
        for ji in j6_by_key.get(dkey, []):
            if ji in used_j:
                continue
            jd = _dday(journal_rows[ji]['date'])
            if jd is None:
                continue
            if abs((bd - jd).days) <= 45:
                used_b.add(bi)
                used_j.add(ji)
                matched.append((b, journal_rows[ji], '代扣通道匹配(待报解↔扣缴)'))
                break
    # 6b. 代扣通道 1:N 组合匹配：网银 1 笔待报解 = 账侧多笔税费之和
    #     （6900 实测：待报解 16669.63 = 增值税 15673.33+教育附加 235.1+地方教育 156.73
    #       +印花税 55.91+城建税 548.56，账按税种分笔、网银按收缴通道合计）
    #     ⚡ 2026-08-14 窗口 15→45（同 6a/跨期未达）：3100 农商行 7 月待报解 709 万
    #       （7-13 扣款）= 账 7-31 社保 5 笔合计（7,060,266.40+4,294+15,029+2,147+10,735），
    #       差 18 天原 15 天卡住 → 709 万对称挂红区（月度对账 0 但差异清单红区有）
    rem_b6b = [(i, b) for i, b in rem_b6 if i not in used_b]
    rem_j6b = [(i, j) for i, j in rem_j6 if i not in used_j]
    for bi, b in rem_b6b:
        if bi in used_b:
            continue
        bd = _dday(b['date'])
        if bd is None:
            continue
        target = round(b['amount'], 2)
        pool = [(ji, j) for ji, j in rem_j6b
                if ji not in used_j
                and dir_map.get(b['direction'], '') == j['direction']
                and _dday(j['date']) is not None
                and abs((_dday(j['date']) - bd).days) <= 45
                and round(j['amount'], 2) <= target]
        combo = _find_sum([(ji, j) for ji, j in pool], target)
        if combo:
            for ji, _j in combo:
                used_j.add(ji)
            used_b.add(bi)
            matched.append((b, [journal_rows[ji] for ji, _j in combo],
                            '代扣通道汇总匹配(1笔网银=%d笔账税费)' % len(combo)))

    # ── 7. 关键字分组加总配对（用户方法论 2026-08-14：找业务关键字而非名字——
    #        大量重复先加总再在另一边找对应金额）。所有逐笔/累计/跨期/代扣之后兜底，
    #        同业务组（纸管/公积金/汇付/后勤/医保/E信/社保/工资…）两侧加总相等才配。
    #        ⚡ 前置保护：1a/窗口/跨期匹配已避开有业务关键词的网银行（防错配给非本业务账行），
    #        业务批次行完整保留到这里统一加总配对。──
    _keyword_sum_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched)
    # ⚡⚡⚡ 2026-08-14 语义簇自动发现（用户方法论 22:19：不要只记关键字，每个账套可能都不同——
    #   大量加总后自动找"差不多意思"的词，靠金额合计相等桥接语义）。硬编码表配完后，
    #   自动学习剩余未配行的业务词簇（医疗保障↔生育津贴这类字面不同、金额相等=同一业务）。──
    _auto_cluster_match(bank_rows, journal_rows, dir_map, used_b, used_j, matched)

    unmatched_b = [b for i, b in enumerate(bank_rows) if i not in used_b]
    unmatched_j = [j for i, j in enumerate(journal_rows) if i not in used_j]

    # ⚡⚡⚡ 2026-08-14 对称找和配对（用户方法论：差异清单大额+月度对账差小=应配未配）：
    #   剩余未配对中，网银 1 笔 = 账侧 1 笔同额（跨期/名称变体）或 = 账侧 N 笔分拆求和。
    #   实证（3100 建行）：GLOBAL HANTEX 499,952.96 = 账 388,576.22+111,376.74（2 笔拆账）；
    #   GAIN LUCKY 1,214,186.56 = 账 755,909.85+458,276.71。命中→入 matched 并从两侧移除
    #   （真正消除差异，非仅标黄）。用归一化后 amount（红字已翻正）。
    #   ⚡⚡ 2026-08-14 规模保护：单银行未配 >3000 行跳过对称找和（3500 大主体 4 万
    #   未配行组合爆炸 OOM——铁律113/114，大主体核对优先，找和可后补）。
    if len(unmatched_b) + len(unmatched_j) <= 3000:
        try:
            from collections import defaultdict as _dd
            _symb = {}            # unmatched_b 索引 -> [unmatched_j 索引]
            _j_used = set()
            _j_by_amt = _dd(list)
            for _ji, _jx in enumerate(unmatched_j):
                _j_by_amt[round(float(_jx.get('amount') or 0), 2)].append(_ji)
            # 1:1 同额（跨期/名称不同但金额一致）
            for _bi, _bx in enumerate(unmatched_b):
                _a = round(float(_bx.get('amount') or 0), 2)
                for _ji in _j_by_amt.get(_a, []):
                    if _ji not in _j_used:
                        _symb[_bi] = [_ji]
                        _j_used.add(_ji)
                        break
            # 1:N 找和（同方向同月，金额分拆）
            _rem_j = [i for i in range(len(unmatched_j)) if i not in _j_used]
            if _rem_j and len(_rem_j) > 1:
                _j_dir_month = _dd(list)
                for _ji in _rem_j:
                    _jx = unmatched_j[_ji]
                    _j_dir_month[(str(_jx.get('direction') or ''),
                                  str(_jx.get('date') or '')[:7])].append(_ji)
                for _bi, _bx in enumerate(unmatched_b):
                    if _bi in _symb:
                        continue
                    _target = round(float(_bx.get('amount') or 0), 2)
                    if _target <= 0:
                        continue
                    # ⚡ 方向归一化：网银『收/支』↔ 账『借/贷』（dir_map 收→借/支→贷）
                    _bd = dir_map.get(str(_bx.get('direction') or ''),
                                      str(_bx.get('direction') or ''))
                    _bm = str(_bx.get('date') or '')[:7]
                    _cands = _j_dir_month.get((_bd, _bm), [])
                    _cands = [i for i in _cands if float(unmatched_j[i].get('amount') or 0) <= _target][:60]
                    if len(_cands) < 2:
                        continue
                    _hit = _subset_sum(_cands, unmatched_j, _target)
                    if _hit:
                        _symb[_bi] = _hit
                        _j_used.update(_hit)
            if _symb:
                # 组装 matched（原始行对象）
                _sb = set(_symb.keys())
                _sj = set()
                for _jjs in _symb.values():
                    _sj.update(_jjs)
                for _bi, _jjs in _symb.items():
                    _bx = unmatched_b[_bi]
                    _jxs = [unmatched_j[_ji] for _ji in _jjs]
                    matched.append((_bx, _jxs if len(_jxs) > 1 else _jxs[0],
                                    '对称找和配对(%d网银=%d账)' % (1, len(_jxs))))
                unmatched_b = [b for i, b in enumerate(unmatched_b) if i not in _sb]
                unmatched_j = [j for i, j in enumerate(unmatched_j) if i not in _sj]
        except Exception:
            pass

    # ⚡ 2026-08-13 未配对的归一化红字【还原原始值】——负数归一化只用于配对（网银收 -X ↔
    #    账贷 X）；未配对的必须还原『收 -X』，否则月度对账出现收差=支差=X 的对称假差异
    #    （1/2 月 +44.89 实证：网银红字冲正账未记=真实差异，应显示『收 -44.89』不对称红区）。
    for _r in unmatched_b + unmatched_j:
        if _r.get('_orig_dir'):
            _r['direction'] = _r['_orig_dir']
            _r['amount'] = _r['_orig_amt']
            del _r['_orig_dir'], _r['_orig_amt']
    summary = {
        'bank_total': len(bank_rows), 'journal_total': len(journal_rows),
        'matched': len(matched), 'unmatched_b': len(unmatched_b),
        'unmatched_j': len(unmatched_j),
    }
    return matched, unmatched_b, unmatched_j, summary


def _kw_of(s):
    """提取中文 3-gram 滑窗词集合（供网银↔账 摘要关键词共享匹配）。
    用滑窗而非贪婪整串：『收星星充电桩电费款』→ {收星星,星星充,星充电,充电桩,电桩电,桩电费,电费款}，
    与网银『电商管家平台交易资金(星星充电)』共享 {星星充,星充电}。"""
    s = re.sub(r'[^\u4e00-\u9fa5]', '', str(s))
    if len(s) < 3:
        return set()
    return {s[i:i + 3] for i in range(len(s) - 2)}


def _has_change(v):
    """金额是否带角分零头（1232.35→True / 100000.00→False）——零星逐笔流水的特征。"""
    return abs(round(float(v or 0), 2) % 1) > 0.001


def _small_monthly_settle(results, small_amt=None):
    """零星小额核销（用户方法论 2026-08-13）：「网银可能 100 笔，但记账只做了 1 笔（汇总）——
    单笔<阈值，记账方向没错，摘要不清晰没关系」。逐笔账侧小额行，在【同月+同方向】的
    网银小额里找金额组合 = 该账行 → 核销（账 1 笔 = 网银 N 笔之和）。
    找不到组合的保留（真差异/跨期）。返回核销笔数。"""
    if small_amt is None:
        small_amt = _SMALL_AMT
    from collections import defaultdict
    dmap = {'借': '收', '贷': '支'}  # 账方向 → 网银方向
    n = 0
    for b in list(results):
        ub, uj = results[b]['ub'], results[b]['uj']
        if not uj:
            continue
        b_idx = defaultdict(list)
        for i, x in enumerate(ub):
            if x['amount'] < small_amt:
                m = str(x['date'])[:7]
                if m:
                    b_idx[(x['direction'], m)].append((i, x))
        settled_b, settled_j = set(), set()
        settled_notes = []  # ⚡ 2026-08-13 汇总记账核销时记录账行，检查『收入未过往来』
        for ji, jx in enumerate(uj):
            if jx['amount'] >= small_amt:
                continue
            m = str(jx['date'])[:7]
            pool = b_idx.get((dmap.get(jx['direction'], ''), m), [])
            if not pool:
                continue
            # ⚡ 用户方法论：小额=网银单笔<10万 且【大部分有零头】（有角分=逐笔零星流水，
            #    企业才可能汇总记一笔账）；整月都是整数金额 → 非汇总记账场景，不核销
            if sum(1 for _i, x in pool if _has_change(x['amount'])) / len(pool) < 0.5:
                continue
            target = round(jx['amount'], 2)
            # ① 同月同方向网银小额全部合计 = 该账行 → 全核销（网银 100 笔 = 账 1 笔）
            if abs(sum(round(x['amount'], 2) for _i, x in pool) - target) < 0.02:
                for i, _x in pool:
                    settled_b.add(i)
                settled_j.add(ji)
                settled_notes.append(jx)
                continue
            # ①b 按共享关键词缩小 pool（账摘要『收星星充电桩电费款』↔ 网银对方名
            #     『电商管家平台交易资金（星星充电）』共享『星星充电』）→ 全部合计=账行
            kw_pool = [p for p in pool
                       if _kw_of(jx.get('summary', '')) & _kw_of((p[1].get('counterparty') or '')
                                                                 + (p[1].get('summary') or ''))]
            if kw_pool and len(kw_pool) < len(pool) \
                    and abs(sum(round(x['amount'], 2) for _i, x in kw_pool) - target) < 0.02:
                for i, _x in kw_pool:
                    settled_b.add(i)
                settled_j.add(ji)
                settled_notes.append(jx)
                continue
            # ② 部分组合（账 1 笔 = 网银 N 笔子集）——先用关键词分组 pool，失败再全 pool
            #    （2026-08-13：原 `kw_pool or pool` 在 kw_pool 非空但找和失败时不再试全 pool
            #    → 医保津贴 8 笔=账 1 笔(摘要'拨付…护理'vs'护理假津贴'不共享)漏配）
            combo = _find_sum(kw_pool, target, max_depth=20)
            if not combo and len(pool) != len(kw_pool):
                combo = _find_sum(pool, target, max_depth=20)
            if combo:
                for i, _x in combo:
                    settled_b.add(i)
                settled_j.add(ji)
                settled_notes.append(jx)
        # ⚡ 2026-08-13 用户方法论：汇总记账核销时——账侧入【往来】（应收/预收）=核对一致；
        #   直接记【收入】=『收入未过往来』，标记提示但认可核对一致（先标记，不算差异）
        for jx in settled_notes:
            j_cp = str(jx.get('cp') or '')
            if any(k in j_cp for k in _INCOME_KWS):
                results[b].setdefault('issues', []).append({
                    'bank': b, 'date': jx['date'], 'direction': jx['direction'],
                    'amount': jx['amount'], 'level': '关注', 'desc': '收入未过往来',
                    'bank_text': '（网银多笔小额收款=账汇总 1 笔）', 'j_cp': j_cp[:36],
                    'j_sm': str(jx.get('summary') or '')[:30],
                    'reason': f'网银多笔小额收款财务汇总 1 笔，对方科目「{j_cp[:24]}」为收入类——'
                              + '收入未过往来账项（一般应先进应收账款），标记提示；汇总金额已核对一致，认可',
                })
        if settled_b or settled_j:
            # ⚡ 2026-08-13 v2：记录被核销的行（供 _monthly_totals 同步剔除——小额汇总记账
            #   两边对称，剔除后发生额列更真实，差异列不变）
            results[b]['small_b_rows'] = results[b].get('small_b_rows', []) + [ub[i] for i in settled_b]
            results[b]['small_j_rows'] = results[b].get('small_j_rows', []) + [uj[i] for i in settled_j]
            results[b]['ub'] = [x for i, x in enumerate(ub) if i not in settled_b]
            results[b]['uj'] = [x for i, x in enumerate(uj) if i not in settled_j]
            results[b]['small_settled'] = len(settled_b) + len(settled_j)
            n += len(settled_b) + len(settled_j)
    return n


def _contra_settle(results, window=60):
    """账侧双向对销踢虚增（用户方法论 2026-08-13 修正）：同银行内两笔【反方向 + 同对方科目
    + 同金额（±0.01）】的「账有网银无」剩余 → 对销（财务记错后更正/冲销，净额 0，银行账面
    不受影响）→ 问题不大，从差异移除（绿区「✅ 双向对销」）。
    ⚡ 2026-08-13 修正：**只处理账侧**——「网银有账无」即使成对（资金一来一回仅隔 1 分钟）
    也是**实际业务发生但财务未记账 = 重大漏记**，绝不自动核销（由 _bank_contra_issue 标重大）。
    返回核销笔数。"""
    n = 0
    for b in list(results):
        if b == '特殊科目':  # 币种/现流转换已单列非差异，不参与对销统计
            continue
        res = results[b]
        # ── 账侧：借↔贷 + 同 cp ──
        uj = res['uj']
        by_key = {}
        for i, x in enumerate(uj):
            cp = _cp_norm(x.get('cp'))
            if cp:
                by_key.setdefault((x['direction'], cp, round(x['amount'], 2)), []).append(i)
        contra = set()
        for (d, cp, amt), idxs in by_key.items():
            op = '借' if d == '贷' else '贷'
            for i in idxs:
                if i in contra:
                    continue
                bd = _dday(uj[i]['date'])
                for j in by_key.get((op, cp, amt), []):
                    if j in contra or i == j:
                        continue
                    jd = _dday(uj[j]['date'])
                    if bd and jd and abs((bd - jd).days) <= window:
                        contra.add(i)
                        contra.add(j)
                        n += 2
                        break
        if contra:
            res['contra_settled'] = res.get('contra_settled', 0) + len(contra)
            # ⚡ 2026-08-13 v2：记录被对销的账行（日期/方向/金额），供 _monthly_totals 同步剔除
            #   ——否则月度对账显示原始发生额差（如 3100 中行 5 月内转借+贷 3.9 亿）而差异清单
            #   已对销移除，两表口径打架（用户 2026-08-13 指出「不应成为差异」）
            res['contra_rows'] = res.get('contra_rows', []) + [uj[i] for i in contra]
            res['uj'] = [x for i, x in enumerate(uj) if i not in contra]
    return n


def _entrust_pay_settle(results, window=15):
    """受托支付双计剔除（用户方法论 2026-08-13：『借贷差异一致=发生额双计/跨账户』）：
    银行【受托支付】业务——企业申请贷款 X，银行直接把 X 划给供应商（不经企业活期账户），
    账侧记双计：借『银行存款 X』（放贷入账）+ 贷『银行存款 X』（受托支付货款）——同日同额
    **跨凭证**成对 = 发生额虚增（银行实际在该账户无发生），净额 0 非差异。
    3100 农行实证：9 组『借 短期借款…放贷 X ↔ 贷 9999#@!原料款/货款 X』（1-02/5-19/6-15/6-16/
    6-17/7-01/7-02，合计 9.6 亿），网银农行全年无放贷/受托支付流水（全是资金池实归上存/支取成对）
    → 收差=支差=-9.5 亿双计特征。识别规则（保守）：同银行 uj 内【借含 放贷/购汇/结汇】与
    【贷含 9999#@! 或 GW…发货款】、同额（±0.01）、日期差 ≤window 天 → 双计对 → 从差异移除
    （绿区）。⚠️ 只配『放贷类借 ↔ 特殊/发货类贷』，绝不碰网银已匹配的真实付款。
    ⚡⚡ 2026-08-14 扩展（3200/3400 实证）：
    ① 贷侧关键词放宽到任意『#@!』前缀（260320#@!货款/26-PH-GH…#@!…/9999#@!）——
       #@! 是 SAP 受托支付/采购划款的统一特殊标记；
    ② 借侧增加『内转』（3200 中行：借『内转』4 亿/1.4 亿/1.3 亿 ↔ 贷『260320#@!货款』
       同额成对=受托支付双计，网银无对应）；
    ③ 支持 **N:M 合计平衡**（3400 建行：借『放贷』200M+180M=380M ↔ 贷『9999#@!…』7 笔
       合计 380M，同月内 2 借 = 7 贷）——按【同月】借合计==贷合计（±0.02）成组剔除。
    返回剔除笔数。"""
    from collections import defaultdict
    n = 0
    for b in list(results):
        if b == '特殊科目':
            continue
        res = results[b]
        uj = res['uj']
        loan = defaultdict(list)  # amt -> [idx]
        pay = defaultdict(list)
        loan_rows = []
        pay_rows = []
        for i, x in enumerate(uj):
            txt = str(x.get('cp') or '') + ' ' + str(x.get('summary') or '')
            if x['direction'] == '借' and (any(k in txt for k in ('放贷', '购汇', '结汇'))
                                           or '内转' in str(x.get('cp') or '')):
                loan[round(x['amount'], 2)].append(i)
                loan_rows.append((i, x))
            elif x['direction'] == '贷' and ('#@!' in txt or 'GW' in str(x.get('cp') or '')
                                             or '发货款' in txt or '预付款' in txt):
                pay[round(x['amount'], 2)].append(i)
                pay_rows.append((i, x))
        removed = set()
        # ── 1:1 同额+窗口（原逻辑，优先匹配精确同额） ──
        for amt, li in loan.items():
            for i in li:
                if i in removed:
                    continue
                di = _dday(uj[i]['date'])
                if di is None:
                    continue
                for j in pay.get(amt, []):
                    if j in removed:
                        continue
                    dj = _dday(uj[j]['date'])
                    if dj and abs((di - dj).days) <= window:
                        removed.add(i)
                        removed.add(j)
                        n += 2
                        break
        # ── 2 N:M 同月合计平衡（2026-08-14：3400 建行 2 借=7 贷 / 3200 中行 3 借=3 贷） ──
        rem_loan = [(i, x) for i, x in loan_rows if i not in removed]
        rem_pay = [(i, x) for i, x in pay_rows if i not in removed]
        if rem_loan and rem_pay:
            by_m = defaultdict(lambda: {'借': [], '贷': []})
            for i, x in rem_loan:
                m = str(x['date'])[:7]
                if m:
                    by_m[m]['借'].append((i, x))
            for i, x in rem_pay:
                m = str(x['date'])[:7]
                if m:
                    by_m[m]['贷'].append((i, x))
            for m, d in by_m.items():
                if not d['借'] or not d['贷']:
                    continue
                dr = sum(x['amount'] for _i, x in d['借'])
                cr = sum(x['amount'] for _i, x in d['贷'])
                if dr > 0 and abs(dr - cr) < 0.02:
                    for i, _x in d['借'] + d['贷']:
                        removed.add(i)
                    n += len(d['借']) + len(d['贷'])
        if removed:
            res['entrust_rows'] = res.get('entrust_rows', []) + [uj[i] for i in removed]
            res['uj'] = [x for i, x in enumerate(uj) if i not in removed]
    return n


def _pool_wage_settle(results):
    """资金池划转成对对冲（2026-08-14 3400 实证，用户『收差=支差一致』方法论）：
    集团资金池【实归支取/内转/客户收款】↔【工资/社保/公积金/税/实归上存】账侧借贷成对、
    网银无对应（资金从集团主账户直付/归集，不经该银行活期账户）→ 发生额双计净额 0 非差异。
    3400 中行实证（7 月 140 笔 uj 全部 1:1 同额成对）：
      借客户收款-宿迁雪创 10 万 ↔ 贷实归上存 10 万（资金池归集）；
      借实归支取 549.5 万 ↔ 贷9999#@!增值税 549.5 万；借内转 1200 万 ↔ 贷实归上存 1200 万。
    3400 兴业实证：借内转 785 万（9 笔）↔ 贷工资 7 笔+年终奖 10 万（同月合计平衡）。
    实现（两级）：
      ① 1:1 同额配对——同月内【借 amt ↔ 贷 amt】逐笔配对移除（贪心，覆盖客户收款↔实归上存）；
      ② 整月合计平衡——剩余行按【月】借合计==贷合计（±0.02）成组对冲（覆盖 N:M，如兴业 785 万）。
    ⚠️ 只配『资金池借 ↔ 工资税费/上存贷』组合，网银有对应流水已匹配的行不在 uj 不受影响；
       3100 农行实证：账侧实归上存贷=网银实归上存流水已匹配，不会进 uj，天然免疫。
    返回剔除笔数。"""
    from collections import defaultdict
    n = 0
    dr_kw = ('实归支取', '内转', '联动支付', '客户收款', '罚款收入')
    cr_kw = ('工资', '社保', '公积金', '年终奖', '#@!', '税', '实归上存', '托管')
    # ⚡⚡ 2026-08-14 学习闭环扩展：资金池借/贷关键词也支持学习词注入
    try:
        from diff_learner import learned_keywords as _lk
        _dr = _lk('pool_dr')
        if _dr:
            dr_kw = dr_kw + _dr
        _cr = _lk('pool_cr')
        if _cr:
            cr_kw = cr_kw + _cr
    except Exception:
        pass
    for b in list(results):
        if b == '特殊科目':
            continue
        res = results[b]
        uj = res['uj']
        removed = set()
        # ── ① 1:1 同额配对（同月内借 amt ↔ 贷 amt） ──
        by_month_b = defaultdict(lambda: defaultdict(list))  # m -> amt -> [idx]
        by_month_c = defaultdict(lambda: defaultdict(list))
        for i, x in enumerate(uj):
            cp = str(x.get('cp') or '') + ' ' + str(x.get('summary') or '')
            m = str(x['date'])[:7]
            if not m:
                continue
            if x['direction'] == '借' and any(k in cp for k in dr_kw):
                by_month_b[m][round(x['amount'], 2)].append(i)
            elif x['direction'] == '贷' and any(k in cp for k in cr_kw):
                by_month_c[m][round(x['amount'], 2)].append(i)
        for m in by_month_b:
            for amt, bl in by_month_b[m].items():
                cl = by_month_c.get(m, {}).get(amt, [])
                k = min(len(bl), len(cl))
                for i in range(k):
                    removed.add(bl[i])
                    removed.add(cl[i])
                    n += 2
        # ── ② 整月合计平衡（剩余 N:M） ──
        rem = [(i, x) for i, x in enumerate(uj) if i not in removed]
        by_m2 = defaultdict(lambda: {'借': [], '贷': []})
        for i, x in rem:
            cp = str(x.get('cp') or '') + ' ' + str(x.get('summary') or '')
            m = str(x['date'])[:7]
            if not m:
                continue
            if x['direction'] == '借' and any(k in cp for k in dr_kw):
                by_m2[m]['借'].append(i)
            elif x['direction'] == '贷' and any(k in cp for k in cr_kw):
                by_m2[m]['贷'].append(i)
        for m, d in by_m2.items():
            if not d['借'] or not d['贷']:
                continue
            dr = sum(uj[i]['amount'] for i in d['借'])
            cr = sum(uj[i]['amount'] for i in d['贷'])
            if dr > 0 and abs(dr - cr) < 0.02:
                for i in d['借'] + d['贷']:
                    removed.add(i)
                n += len(d['借']) + len(d['贷'])
        if removed:
            res['pool_j_rows'] = res.get('pool_j_rows', []) + [uj[i] for i in removed]
            res['uj'] = [x for i, x in enumerate(uj) if i not in removed]
    return n


def _inner_transfer_net(results, window=15):
    """内转 1:N 对销（用户方法论 2026-08-13：『内转』=本企业账户间内部资金划转，借+贷合计
    对冲净额 0 非差异 → 从差异移除）。原 _contra_settle 只做 1:1 同额同 cp；**一笔内转借 =
    多笔内转贷合计**（拆笔划转）漏配。3100 华夏实证：1-16 借 2000 万（内转）= 1-19 贷
    885+593+522 万（内转，合计 2000 万，差 3 天）→ 假差异 2000 万（月度合计对冲会被 4.5 万
    小内转干扰，故用 _find_sum 精确 1:N 找和）。
    实现：同银行 uj 内『cp 含内转』行，窗口内（≤15 天）每笔借找贷组合和=借金额（及反向
    多笔借=一笔贷）→ 对销移除。⚠️ 仅限『内转』cp（内部划转专用），绝不动货款/税费。"""
    n = 0
    for b in list(results):
        if b == '特殊科目':
            continue
        res = results[b]
        uj = res['uj']
        rows = [(i, x) for i, x in enumerate(uj)
                if '内转' in str(x.get('cp') or '') + str(x.get('summary') or '')]
        if len(rows) < 2:
            continue
        dr_ids = [i for i, x in rows if x['direction'] == '借']
        cr_ids = [i for i, x in rows if x['direction'] == '贷']
        removed = set()
        # ① 每笔借 → 贷组合（1:N）
        for i in dr_ids:
            if i in removed:
                continue
            di = _dday(uj[i]['date'])
            if di is None:
                continue
            target = round(uj[i]['amount'], 2)
            pool = [(j, uj[j]) for j in cr_ids if j not in removed
                    and _dday(uj[j]['date']) is not None
                    and abs((_dday(uj[j]['date']) - di).days) <= window]
            combo = _find_sum(pool, target)
            if combo:
                removed.add(i)
                for j, _x in combo:
                    removed.add(j)
        # ② 每笔贷 → 借组合（N:1，防先剔除方向影响）
        for j in cr_ids:
            if j in removed:
                continue
            dj = _dday(uj[j]['date'])
            if dj is None:
                continue
            target = round(uj[j]['amount'], 2)
            pool = [(i, uj[i]) for i in dr_ids if i not in removed
                    and _dday(uj[i]['date']) is not None
                    and abs((_dday(uj[i]['date']) - dj).days) <= window]
            combo = _find_sum(pool, target)
            if combo:
                removed.add(j)
                for i, _x in combo:
                    removed.add(i)
        if removed:
            res['transfer_rows'] = res.get('transfer_rows', []) + [uj[i] for i in removed]
            res['uj'] = [x for i, x in enumerate(uj) if i not in removed]
            n += len(removed)
    return n


# ⚡ 2026-08-13 期末重要性水平（用户方法论：期末平的但存在银行有财务无→需重要性衡量）
_MATERIALITY_PCT = 0.01    # 期末未达净额占期末余额 1% 为重要性阈值
_MATERIALITY_AMT = 1_000_000  # 或绝对额 ≥100 万


def _materiality_check(results, banks):
    """期末重要性水平：每银行期末未达净额（网银未配净额−账未配净额，资金池/通道类除外）
    vs 年累计流水（收+支合计，最可靠口径；网银 balance 列多银行解析不准不可用）→
    比例超阈值或绝对额大 → 提示。⚡ 特殊科目（币种/现流转换）不参与。返回超阈值家数。"""
    n = 0
    for b in results:
        if b == '特殊科目':  # 过渡科目非真实账户，不参与重要性
            continue
        res = results[b]
        net_b = sum((x['amount'] if x['direction'] == '收' else -x['amount'])
                    for x in res['ub'] if not _is_auto_channel(x))
        net_j = sum((x['amount'] if x['direction'] == '借' else -x['amount'])
                    for x in res['uj'])
        tot_flow = sum(x['amount'] for x in banks.get(b, []))  # 年累计流水（收+支）
        unrec = abs(net_b - net_j)
        ratio = unrec / tot_flow if tot_flow else (1.0 if unrec else 0.0)
        flag = unrec >= _MATERIALITY_AMT or (tot_flow and ratio >= _MATERIALITY_PCT)
        res['materiality'] = {'net_b': net_b, 'net_j': net_j, 'unrec': unrec,
                              'balance': tot_flow, 'ratio': ratio, 'flag': bool(flag)}
        if flag:
            n += 1
    return n
#   ① 资金池归集（银行每日营业末划至归集户/营业初划回，财务不作账正常）
#   ② 代扣代缴通道（待报解-TIPS 扣税/公积金/社保——账上记应交税费/应付职工薪酬）
#   ③ 银行批量自动处理（批量账务/代理付款/集中处理/供应链合约户/结售汇/远期）
_AUTO_CHANNEL_KWS = ('资金归集', '资金池', '归集', '上划', '下拨', '上存', '联动支付', '划回',
                     '划转回', '成员归集', '转存',
                     '待报解', '扣税', '代理国库', '税收收缴', '代扣', '公积金', '社保',
                     '批量', '集中处理', '代理付款', '供应链')
# ⚡ 2026-08-13 修正 v2：结售汇/远期/外汇 从通道类移除——网银购汇/结汇流水（光大远期结售汇、
#   工行即期结售汇）是真实购汇业务，须与账侧真实账户购汇行（借'农行-人民币'收 1370 万等）配对
# 集团关联方特征词：成对（收↔支同对方同金额）的集团关联大额划转 = 集团资金调度/跨境资金池
# （6100 实测：SHENGHONG PETROCHEMICAL ↔ 盛虹炼化 9 亿收+9 亿支成对=集团资金池），单列不计漏记
_GROUP_CORP = ('盛虹', '国望', '港虹', '新视界', '中鲈', '苏震', '东方盛虹', '盛虹科技',
               '盛虹化纤', '苏州盛虹', '盛虹炼化', 'SHENGHONG')


def _is_auto_channel(x):
    """银行自动/通道类判定：摘要/对方含特征词（银行自动归集/代扣代缴/批量处理，财务不逐笔记正常）。"""
    txt = str(x.get('counterparty') or '') + ' ' + str(x.get('summary') or '')
    return any(k in txt for k in _AUTO_CHANNEL_KWS)


def _daikou_monthly_settle(results, banks, journal):
    """待报解月汇总核对（用户方法论 2026-08-13）：网银『待报解-TIPS 扣税』全部按【月】汇总
    支出合计 vs 账侧同月『应交税费/税金及附加』借方（=银行存款贷方支出，含已 1:N 匹配行）
    合计——金额一致 → 该月全部待报解算核对一致（网银按扣款日逐笔、账按所属期汇总计提，
    N:M 跨月属正常口径），从差异移除。⚡ 用全量（banks/journal）而非未配剩余——
    账税费行大多已被 1:N 代扣匹配占用，只看 uj 会漏。返回核销笔数。"""
    from collections import defaultdict
    n = 0
    for b in list(results):
        res = results[b]
        bm = defaultdict(float)
        pool_rows = []
        for x in banks.get(b, []):
            if x['direction'] == '支' and '待报解' in (str(x.get('counterparty') or '') + str(x.get('summary') or '')):
                m = str(x['date'])[:7]
                if m:
                    bm[m] += x['amount']
                    pool_rows.append((x, m))
        if not bm:
            continue
        jm = defaultdict(float)
        for jx in journal.get(b, []):
            if jx['direction'] == '贷':
                cp = str(jx.get('cp') or '')
                # ⚡ 2026-08-13 修正 v3（用户：税费不跨银行、一般一个户）：3100 农商行=税费+社保
                #   （待报解-TIPS 扣税 + 'X月社保缴费'），农行=公积金（'X月公积金缴费'→网银
                #   '苏州市住房公积金管理中心'）——**分户分通道**。筛选保留社保/缴费、排除公积金
                if ('退税' not in cp and '公积金' not in cp
                        and any(k in cp for k in ('税', '附加', '社保', '缴费', '个税', '代扣'))):
                    m = str(jx['date'])[:7]
                    if m:
                        jm[m] += jx['amount']
        settled = []
        settled_months = set()
        diffs = []
        for m, amt in bm.items():
            jtot = jm.get(m, 0.0)
            # ⚡ 2026-08-13 尾差容忍：月合计差异 <5万 或 <网银月合计 0.5% → 算核对一致
            #   （银行手续费/取整/少量跨日小额，3100 农商行 3 月 -0.6万/6 月 -3万）
            if jtot > 0 and abs(amt - jtot) < max(50_000.0, amt * 0.005):
                settled.extend(x for x, mm in pool_rows if mm == m)
                settled_months.add(m)
            elif jtot > 0:
                # ⚡ 2026-08-13 用户方法论「双向核对无重要性标准，有真实差异均列示」：
                #   待报解月合计 ≠ 账税费借方月合计 → 先做【跨月对冲确认】——次月账侧补记
                #   （分录正常）则属跨月记账错位（网银扣款日 vs 账所属期），不算差异，仅关注；
                #   无相邻月对冲的才是真差异（黄区列示）
                diffs.append((m, amt, jtot, amt - jtot))
        if settled:
            sid = set(map(id, settled))
            res['ub'] = [x for x in res['ub'] if id(x) not in sid]
            # ⚡ 2026-08-13 v2 对称移除：核销月同时移除账侧该月 uj 中税费/社保/缴费类贷方行
            #   ——否则红区网银侧待报解移除、账侧税费行保留（不对称），月度对账（两侧都保留）
            #   与红区打架（3100 农商行 2/3/5/6 月待报解核销 5377 万，账侧 6357 万税费贷留红区）
            # ⚡⚡ 2026-08-14 v3：改为从【全量 journal】筛而非 res['uj']——已被 reconcile 代扣
            #   通道（6a/6b）配掉的账税费贷不在 uj，若只筛 uj → 网银侧待报解全剔、账侧只剔未配
            #   部分 → 月度对账支差（3300 农商行实证：待报解 5434 万全核销，账侧 1279 万已被
            #   代扣匹配的行漏剔 → 支差 -1280 万假差异）。全量筛+expl id 去重天然免疫重复剔除。
            j_rm = [jx for jx in journal.get(b, [])
                    if jx['direction'] == '贷' and str(jx['date'])[:7] in settled_months
                    and any(k in str(jx.get('cp') or '') for k in ('税', '附加', '社保', '缴费', '个税', '代扣'))
                    and '退税' not in str(jx.get('cp') or '') and '公积金' not in str(jx.get('cp') or '')]
            if j_rm:
                jrm_ids = set(map(id, j_rm))
                res['uj'] = [x for x in res['uj'] if id(x) not in jrm_ids]
                res['daikou_j_rows'] = res.get('daikou_j_rows', []) + j_rm
            res['daikou_monthly'] = res.get('daikou_monthly', 0) + len(settled)
            # ⚡ 2026-08-14 v3 对称记录：网银侧已核销待报解支也记录（供 _monthly_totals expl
            #   同步剔除）——否则待报解支被 pool/其他环节移走后，月度对账网银侧减、账侧不减
            #   （3300 农商行 42 笔待报解 4154 万被 pool 误判资金池成对移走，账税费贷 4154 万
            #   只从 uj 移除未同步月度对账 → 支差 -4154 万假差异）
            res['daikou_b_rows'] = res.get('daikou_b_rows', []) + settled
            n += len(settled)
        # ⚡ 跨月对冲确认：当月差异 ≈ 相邻月（±1 月）反向差异 → 次月补记，跨月错位正常
        if diffs:
            used = set()
            cross = []
            real = []
            for i, (m1, a1, j1, d1) in enumerate(diffs):
                if i in used:
                    continue
                paired = None
                for j, (m2, a2, j2, d2) in enumerate(diffs):
                    if j == i or j in used:
                        continue
                    mm1, mm2 = m1[:7], m2[:7]
                    try:
                        y1, mo1 = int(mm1[:4]), int(mm1[5:7])
                        y2, mo2 = int(mm2[:4]), int(mm2[5:7])
                        gap = (y2 * 12 + mo2) - (y1 * 12 + mo1)
                    except Exception:
                        gap = 99
                    # ⚡ 2026-08-13 用户方法论『和下月发生额加在一起能否匹配』：相邻月差异
                    #   之和 < 尾差 → 两月滚动匹配（跨月错位，下月补记），降级关注
                    if abs(gap) == 1 and abs(d1 + d2) < max(50_000.0, abs(d1) * 0.02, abs(d2) * 0.02):
                        paired = j
                        break
                if paired is not None:
                    used.add(i)
                    used.add(paired)
                    cross.append((diffs[i], diffs[paired]))
                else:
                    real.append(diffs[i])
            res['daikou_cross_month'] = res.get('daikou_cross_month', []) + cross
            res['daikou_monthly_diff'] = res.get('daikou_monthly_diff', []) + real
    return n


def _classify_bank_unmatched(results, window=7, min_amt=100_000, banks=None, journal=None):
    """网银有账无分类（用户方法论 2026-08-13 修正）：
    ① 资金池自动归集（银行每日营业末划至归集户/次日营业初划回，财务不作账）= 正常 → 单列不计差异
    ② 非资金池【成对】（收+支同对方单位+同金额+日期差≤window）→ 资金实际发生但账未记
       = 重大漏记发生额（哪怕仅隔 1 分钟），标重大 issue
    ③ 非资金池【单边】→ 期末对不平=真差异（≥min_amt 标重大 / 小额留红区）
    每银行存 res['ub_class'] = {'pool': n, 'contra': n, 'single': n, 'single_amt': 金额}。
    返回标注 issue 数。"""
    n_iss = 0
    for b in list(results):
        res = results[b]
        ub = res['ub']
        auto_pool, rest = [], []
        for x in ub:
            # ⚡ 2026-08-14 自动通道（公积金/社保/待报解/批量等）**不参与 pool 移除**——
            #   它们是真实业务（3100 农行公积金缴费 1117 万实证：网银支『苏州市住房公积金
            #   管理中心』账侧有对应贷行，误当资金池剔除 → 网银侧减、账侧不减 → 支差 -1142 万
            #   假差异）；自动通道由 daikou 月核销/正常匹配处理，仅在 ub_class 单列
            (auto_pool if _is_auto_channel(x) else rest).append(x)
        pool = list(auto_pool)  # 仅计数展示用
        by_key = {}
        for i, x in enumerate(rest):
            cp = str(x.get('counterparty') or '').strip()
            if cp:
                by_key.setdefault((x['direction'], cp, round(x['amount'], 2)), []).append(i)
        contra_idx = set()
        for (d, cp, amt), idxs in by_key.items():
            # ⚡ 2026-08-14 集团内小额成对豁免：集团成员（盛虹/国望等）同单位同额收+支成对
            #    = 资金池自动归集/往来挂账（账不作账正常），不受 min_amt 门槛限制——
            #    3000 建行 6-08 收+支 216.63 盛虹化纤实证（原 10 万门槛挡掉 → 216 元假差异挂红区）
            is_group = any(g in cp for g in _GROUP_CORP)
            if amt < min_amt and not is_group:
                continue
            op = '收' if d == '支' else '支'
            for i in idxs:
                if i in contra_idx:
                    continue
                bd = _dday(rest[i]['date'])
                for j in by_key.get((op, cp, amt), []):
                    if j in contra_idx or i == j:
                        continue
                    jd = _dday(rest[j]['date'])
                    if bd and jd and abs((bd - jd).days) <= window:
                        # 集团关联方成对划转（SHENGHONG↔盛虹炼化 9 亿成对）→ 集团资金调度/跨境资金池
                        # （银行自动归集性质，财务未逐笔记正常），单列不计漏记
                        if any(g in cp for g in _GROUP_CORP):
                            pool.append(rest[i])
                            pool.append(rest[j])
                            contra_idx.add(i)
                            contra_idx.add(j)
                            break
                        contra_idx.add(i)
                        contra_idx.add(j)
                        x_i, x_j = rest[i], rest[j]
                        res.setdefault('issues', []).append({
                            'bank': b, 'date': x_i['date'], 'direction': x_i['direction'],
                            'amount': x_i['amount'], 'level': '重大', 'desc': '网银成对漏记',
                            'bank_text': (cp + ' ' + str(x_i.get('summary') or '')).strip()[:50],
                            'j_cp': '（网银有账无）', 'j_sm': '',
                            'reason': f'网银「{cp}」{x_i["direction"]} {x_i["amount"]:,.0f} 元与 '
                                      + f'{x_j["date"]} {x_j["direction"]} {x_j["amount"]:,.0f} 元成对'
                                      + '（非资金池/集团），账面无对应——实际业务发生但财务未记账，'
                                      + '哪怕间隔仅 1 分钟也是漏记发生额（重大），重点核查',
                        })
                        n_iss += 1
                        break
        # ⚡ 2026-08-13 镜像成对：同银行同金额 收+支（counterparty 可不同，如 SHENGHONG 收 9 亿
        #   ↔ 盛虹炼化支 9 亿=同一笔跨境划转两半）——任一侧含集团关联词 → 集团资金调度单列
        # ⚡⚡ 2026-08-14 v2 全量范围：在【rest + auto_pool 全量】中找收+支配对——3300 建行
        #   实证：1-30 四笔 1036 万（东方盛虹资金归集 收+支 2 笔含'归集'进 auto_pool + 虹港
        #   石化退汇收/转账支 2 笔进 rest）= 同一集团资金调度净额 0，若只在 rest 找 → 收 1036
        #   万单边残留假差异（收差=支差=+1036 万）；auto 通道单边（公积金/社保/待报解=真实
        #   业务）无配对应保留，只有【成对】才移除
        all_ub = rest + auto_pool
        amt_key = {}
        for i, x in enumerate(all_ub):
            amt_key.setdefault(round(x['amount'], 2), []).append(i)
        mirror_ids = set()
        for amt, idxs in amt_key.items():
            if amt < min_amt:
                continue
            recv = [i for i in idxs if all_ub[i]['direction'] == '收']
            pay = [i for i in idxs if all_ub[i]['direction'] == '支']
            for ri in recv:
                if ri in mirror_ids:
                    continue
                d1 = _dday(all_ub[ri]['date'])
                for pj in pay:
                    if pj in mirror_ids:
                        continue
                    d2 = _dday(all_ub[pj]['date'])
                    if d1 and d2 and abs((d1 - d2).days) <= window:
                        cp_a = (str(all_ub[ri].get('counterparty') or '') + ' ' +
                                str(all_ub[pj].get('counterparty') or '')).upper()
                        if any(g.upper() in cp_a for g in _GROUP_CORP):
                            mirror_ids.add(ri)
                            mirror_ids.add(pj)
                        break
        # 镜像成对行（含 auto 通道行）加入移除；rest 中的镜像对也标 contra_idx
        mirror_rest_idx = {i for i in mirror_ids if i < len(rest)}
        contra_idx |= mirror_rest_idx
        # auto_pool 中成对部分也移除（pair_rows），单边保留
        auto_mirror = [all_ub[i] for i in mirror_ids if i >= len(rest)]
        single = [x for i, x in enumerate(rest) if i not in contra_idx]
        single_big = [x for x in single if x['amount'] >= min_amt]
        for x in single_big:
            res.setdefault('issues', []).append({
                'bank': b, 'date': x['date'], 'direction': x['direction'],
                'amount': x['amount'], 'level': '重大', 'desc': '网银单边未记',
                'bank_text': (str(x.get('counterparty') or '') + ' ' + str(x.get('summary') or '')).strip()[:50],
                'j_cp': '（网银有账无）', 'j_sm': '',
                'reason': f'网银单边{x["direction"]} {x["amount"]:,.0f} 元（{str(x.get("counterparty") or "")[:18]}），'
                          + '账面无对应且不成对——期末应对不平，实际业务发生未记账（重大），重点核查',
            })
            n_iss += 1
        # 成对（同key成对 contra_idx + 镜像成对 auto_mirror）单独收集——auto_pool 单边保留
        pair_rows = ([x for i, x in enumerate(rest) if i in contra_idx] + auto_mirror)
        res['ub_class'] = {'pool': len(auto_pool) + len(pair_rows), 'contra': len(contra_idx),
                           'single': len(single),
                           'single_big': len(single_big),
                           'single_amt': sum(x['amount'] for x in single)}
        # ⚡ 2026-08-14 资金池成对【应收配平】从差异移除：集团成员同单位同额收+支成对
        #   （资金池自动归集/往来挂账，账不作账正常，净流入 0）= 非真实差异 → 红区不显示、
        #   月度对账同步剔除（3000 建行 6-08 收+支 216.63 盛虹化纤实证）
        # ⚡⚡ 2026-08-14 v2 只移除【成对】，自动通道（公积金/社保/待报解）单边保留——
        #   单边自动通道是真实业务（3100 农行公积金 1117 万），由 daikou/正常匹配处理；
        #   auto 通道中【成对】部分（3300 建行 1-30 资金归集收+支 1036 万）仍应移除
        if pair_rows:
            pool_ids = set(map(id, pair_rows))
            res['ub'] = [x for x in res['ub'] if id(x) not in pool_ids]
            res['pool_rows'] = res.get('pool_rows', []) + pair_rows
    return n_iss


def _cross_account_match(results, banks, journal, window=20):
    """跨账户归因（2026-08-13 用户质疑：6100 建行 2 笔/工行 11 笔大额「网银有账无」
    实为集团资金池划转——网银 A 账户收/支 ↔ 账面记在 B 账户（其他应付款-往来款 国望高科）。
    按【其他银行账户】日记账 同方向+同金额+日期窗口 关联；归因成功的从网银未配移除，
    存 results[b]['cross_acct'] = [(网银行, 关联银行, 账行)]，差异清单黄底显示「已归因」。
    返回归因条数。
    ⚡ 2026-08-13 v2：归因目标必须是该行账侧【未配】行（真『账有网银无』）——jidx 排除已
    matched 账行，否则会把已与网银配平的真实账行错判为跨账户（3100 实证：中信账 1204 万
    已匹配被归因剔除 → 月度对账收差 0 → +1204 万，且差异清单绿区误删）。"""
    from collections import defaultdict
    dir_map = {'收': '借', '支': '贷'}
    used_j = set()
    for b in results:
        for item in results[b]['matched']:
            jside = item[1]
            for j in (jside if isinstance(jside, list) else [jside]):
                used_j.add(id(j))
    jidx = defaultdict(list)
    for jb, rows in journal.items():
        for x in rows:
            if id(x) in used_j:
                continue
            jidx[(x['direction'], round(x['amount'], 2))].append((jb, x))
    n = 0
    removed_j = set()  # ⚡ v3：被归因的目标账行 id（需从 jb 银行 uj 对称移除）
    for b in list(results):
        cross = []
        new_ub = []
        for x in results[b]['ub']:
            bd = _dday(x['date'])
            if bd is None:
                new_ub.append(x)
                continue
            found = None
            for jb, jx in jidx.get((dir_map.get(x['direction'], ''), round(x['amount'], 2)), []):
                if jb == b:  # 仅跨账户；同账户未配已在 reconcile 各阶段处理
                    continue
                jd = _dday(jx['date'])
                if jd and abs((jd - bd).days) <= window:
                    found = (jb, jx)
                    break
            if found:
                cross.append((x, found[0], found[1]))
                removed_j.add(id(found[1]))
                n += 1
            else:
                new_ub.append(x)
        results[b]['ub'] = new_ub
        results[b]['cross_acct'] = cross
    # ⚡ v3：对称移除目标账行——跨账户归因=「网银 A 户 ↔ 账 B 户同一笔资金」，目标账行（jx）
    #   若仍留在 B 银行 uj，则差异清单红区账侧保留、网银侧已移除（不对称），且月度对账同步剔
    #   jx 后与红区打架（3100 实证：华夏账 2000 万/农商行账 5377 万被归因 → 月度对账剔、红区留）
    if removed_j:
        for b in results:
            uj = results[b]['uj']
            if any(id(y) in removed_j for y in uj):
                results[b]['uj'] = [y for y in uj if id(y) not in removed_j]
    return n


def _find_sum(pool, target, max_depth=5):
    """在 (idx,row) 列表找金额和=target 的组合（先单条，再 DFS 组合，最多 max_depth 条）。
    ⚡ 2026-08-13 防爆炸：候选 >80 按接近度裁剪；DFS 节点预算 5 万（3100 农行 6296 笔+
    建行 5185 笔+账 18407 行，无限制会组合爆炸卡死）。"""
    cands = [(i, b) for i, b in pool if round(b['amount'], 2) <= target]
    if len(cands) > 80:
        cands = sorted(cands, key=lambda x: abs(x[1]['amount'] - target))[:80]
    for i, b in cands:
        if abs(round(b['amount'], 2) - target) < 0.01:
            return [(i, b)]
    # DFS 组合（带节点预算）
    budget = [50000]

    def dfs(start, remain, picked):
        budget[0] -= 1
        if budget[0] < 0:
            return None
        if abs(remain) < 0.01:
            return list(picked)
        if len(picked) >= max_depth:
            return None
        for k in range(start, len(cands)):
            i, b = cands[k]
            if round(b['amount'], 2) > remain + 0.01:
                continue
            r = dfs(k + 1, round(remain - b['amount'], 2), picked + [(i, b)])
            if r:
                return r
        return None
    return dfs(0, target, [])


_BANK_FULL2SHORT = {
    '招商银行': '招行', '建设银行': '建行', '农业银行': '农行', '工商银行': '工行',
    '交通银行': '交行', '中国银行': '中行', '国开银行': '国开行', '农村商业银行': '农商行',
    '农商银行': '农商行', '上海银行': '上海银行', '光大银行': '光大', '兴业银行': '兴业',
    '北京银行': '北京银行', '华夏银行': '华夏', '南京银行': '南京银行', '宁波银行': '宁波',
    '平安银行': '平安', '邮储银行': '邮储', '江苏银行': '江苏银行', '浙商银行': '浙商',
    '苏州银行': '苏州银行', '民生银行': '民生', '进出口银行': '进出口行', '中信银行': '中信',
    '嘉兴银行': '嘉兴银行', '浦发银行': '浦发', '北京银行': '北京银行',
}
_BANK_FULLS = tuple(sorted(_BANK_FULL2SHORT.keys(), key=len, reverse=True))
_BANK_SHORTS = ('国开行', '农商行', '招行', '建行', '农行', '工行', '中行', '交行', '光大',
                '兴业', '华夏', '宁波', '平安', '邮储', '浙商', '民生', '进出口', '口行', '中信',
                '嘉兴', '泰隆', '浦发')


def _voucher_net(journal, banks=None):
    """凭证级净额化（用户方法论 2026-08-13：『同凭证已凭证级净额化免疫，跨凭证双计待确认』）：
    同凭证（vchar）内【借 X + 贷 X 同额】= 凭证内内部对冲（一进一出净额 0，如手工凭证
    1700000256 借内转 3219.97 万+贷内转 3219.97 万）→ 从账侧剔除防双计——否则借方/贷方
    只被网银抢配一边、另一边落空成假差异（3100 建行 1-19 实证：1700000256 贷被网银支
    3219.97 万匹配、借落空 → 假账有网银无）。
    ⚡⚡ 2026-08-13 收紧：**只处理『内转』cp 且该银行网银无『同名/内转』流水**——裸『同凭证
    借X+贷X』净额化误杀真实业务（工行外币掉期 52 亿/农行放贷受托 1.22 亿实证）；农行内转
    同凭证借贷对对应网银『同名划转』（真实内部划转）也不能剔（3100 农行反弹 1.22 亿实证）。
    在 reconcile 前对 journal 做（每银行每凭证内借贷同额对销）。返回 {bank: 剔除行数}。"""
    # ⚡ 网银有『同名划转』流水的银行=内转是真实业务（3100 农行实证：网银同名划转=真内转，
    #   剔除即反弹 1.22 亿），不净额化；建行『同名划付款』非内转流水 → 不跳过（1700000256
    #   借内转+贷内转同凭证=凭证内对冲，网银无对应 → 净额化剔除）
    skip = set()
    for b in journal:
        for x in banks.get(b, []):
            txt = str(x.get('counterparty') or '') + ' ' + str(x.get('summary') or '')
            if '同名划转' in txt:
                skip.add(b)
                break
    net = {}
    for b in journal:
        if b in skip:
            continue
        rows = journal[b]
        by_v = {}
        for x in rows:
            by_v.setdefault(str(x.get('vchar') or ''), []).append(x)
        drop = set()
        for v, rs in by_v.items():
            if not v:
                continue
            used = set()
            idx = {}
            for ri, x in enumerate(rs):
                # ⚡ 只处理内转凭证（内部划转），绝不碰货款/放贷/结汇等真实业务
                if '内转' not in str(x.get('cp') or '') + str(x.get('summary') or ''):
                    continue
                k = (x['direction'], round(x['amount'], 2))
                op = '贷' if x['direction'] == '借' else '借'
                mate = idx.get((op, round(x['amount'], 2)), [])
                if mate:
                    ri2 = mate.pop()
                    if ri2 not in used:
                        used.add(ri)
                        used.add(ri2)
                        continue
                idx.setdefault(k, []).append(ri)
            for ri in used:
                drop.add(id(rs[ri]))
        if drop:
            keep = [x for x in rows if id(x) not in drop]
            net[b] = len(rows) - len(keep)
            journal[b] = keep
    return net


def _fee_net_journal(journal, banks=None):
    """账侧凭证级【收款净额化】（农商行二维码收款实证 2026-08-13，用户问『部分收支差异的
    月份是一致的，为何不能对平』）：
    同凭证（vchar）内『借 收款全额 + 贷 手续费/服务费/年费』→ 收款净额 = 借合计 - 费用合计，
    费用贷行从 journal 移除、借方收款行金额扣减——银行扣手续费直接从到账金额中扣（网银
    一码通到账 = 全额 - 手续费），账侧『借全额 + 贷手续费』双计 → 收差 = 支差 = 手续费金额
    的对称假差异（净额平衡，非真实差异）。
    农商行实证：1 月凭证 1800000605『借 26,655.24 二维码收款 + 贷 147.33 农商行扣取手续费』，
    网银一码通 88,649.72 = 账借 88,797.05 - 147.33 → 1 月收差=支差=-147.33；2 月 -204.88 /
    3 月 -469.56 / 4 月 -184.03 / 5 月 -170.82 同模式（各月二维码收款凭证手续费双计）。
    ⚡ 第二级【月聚合净额化】（3 月实证：手续费 469.56 在【独立凭证】1800001598，不在二维码
    凭证内 → 同凭证判定不适用）：当月『账借二维码合计 > 网银一码通收合计』且存在手续费贷行
    → 从二维码借行扣手续费、移除手续费贷行（二维码借 - 网银一码通 = 手续费 + 活动费退回 280
    等，扣手续费后剩余 280 由活动费退回与网银配对，收差归零）。
    判定：借方 cp 含『二维码/一码通』（二维码收款区间汇总，如 20251231-20260129二维码收款）
    + 贷方 cp/摘要含『手续费/服务费/收费/年费/利息』→ 对冲；⚠️ 绝不碰货款/税费/社保、
    不碰『客户收款-XX公司』（客户名称也含'收款'但为真实收款，无手续费对冲）、不动红字冲回
    （贷方含『二维码/收款』的——那是红字冲正，由负数归一化处理）。"""
    net = {}
    from collections import defaultdict
    # ⚡ 2026-08-14 v2：recv 增『饮料机/浴室充值』（3300 农商行收款渠道名，如『饮料机收款』
    #   『浴室充值、水费』——不含『二维码/一码通』关键词，原判定漏收 → recv_amt 太小 →
    #   月聚合平衡判定失败 → 退押金/手续费永不净额化）；fee 增『退押金』（3300 收款冲减名，
    #   与 3100 手续费同构：账借全额+贷退押金=网银一码通到账净额）
    recv_kw = ('二维码', '一码通', '饮料机', '浴室充值')
    fee_kw = ('手续费', '服务费', '收费', '年费', '利息', '退押金')
    # ⚡⚡ 2026-08-14 学习闭环打通：学习器确认的学习词自动注入（diff_learner_memory.
    #   json accepted），下次核对按新关键词净额化——人确认 → 程序自动记住 → 自动
    #   生效，无需手动改本表。缺失/空记忆不影响默认行为。
    try:
        from diff_learner import learned_keywords as _lk
        _lf = _lk('fee')
        if _lf:
            fee_kw = fee_kw + _lf
        _lr = _lk('recv')
        if _lr:
            recv_kw = recv_kw + _lr
    except Exception:
        pass
    for b in list(journal):
        rows = journal[b]
        by_v = defaultdict(list)
        for x in rows:
            by_v[str(x.get('vchar') or '')].append(x)
        removed_ids = set()
        info = []
        # ── 第一级：同凭证净额化 ──
        for v, rs in by_v.items():
            if not v:
                continue
            recv = [x for x in rs if x['direction'] == '借'
                    and any(k in str(x.get('cp') or '') + str(x.get('summary') or '') for k in recv_kw)]
            fees = [x for x in rs if x['direction'] == '贷'
                    and any(k in str(x.get('cp') or '') + str(x.get('summary') or '') for k in fee_kw)]
            if not recv or not fees:
                continue
            fee_amt = round(sum(float(x['amount']) for x in fees), 2)
            recv_amt = round(sum(float(x['amount']) for x in recv), 2)
            if fee_amt <= 0 or fee_amt > recv_amt:
                continue
            # 从借方收款行扣减（优先从金额最大的行扣，不够则依次扣）
            remain = fee_amt
            for x in sorted(recv, key=lambda x: -float(x['amount'])):
                if remain <= 0:
                    break
                amt = float(x['amount'])
                if amt >= remain:
                    x['amount'] = round(amt - remain, 2)
                    remain = 0.0
                else:
                    x['amount'] = 0.0
                    remain = round(remain - amt, 2)
            for x in fees:
                removed_ids.add(id(x))
            info.append((v, round(recv_amt - fee_amt, 2), len(fees)))
        # ── 第二级：月聚合净额化（手续费独立凭证，如 3 月 1800001598） ──
        if banks:
            # ⚡ 网银该月已有同金额支出（年费/收费，网银有独立记录）→ 不是『从收款中扣除』
            #   的手续费，排除——否则误伤（3 月年费 22.50/收费 15.00 网银有支，被净额化移除
            #   → 网银支 22.50/15.00 落空 → 支差 +37.50 假差异，实证）
            bank_exp_amt = defaultdict(set)
            for x in banks.get(b, []):
                if x['direction'] == '支':
                    m = str(x['date'])[:7]
                    if m:
                        bank_exp_amt[m].add(round(float(x['amount']), 2))
            b_one = defaultdict(float)
            for x in banks.get(b, []):
                if x['direction'] == '收' and '一码通' in str(x.get('counterparty') or ''):
                    m = str(x['date'])[:7]
                    if m:
                        b_one[m] += round(float(x['amount']), 2)
            j_recv = defaultdict(lambda: [0.0, []])
            j_fees = defaultdict(list)
            for x in rows:
                # ⚡⚡ 2026-08-14 v2 跳过第一级已移除的行：第二级遍历的 rows 是函数开头
                #   `rows = journal[b]` 的旧列表引用（第一级结束后 journal[b] 重绑新列表但
                #   局部 rows 仍指旧列表）→ 第一级已移除的 fee 贷行在第二级被重复收集 →
                #   同凭证手续费双重扣减（3300 农商行 4 月实证：1800000235 扣 3.80 + 月聚合
                #   又扣 3.80 = 7.60 → 收差 -3.80 变 +3.80 假差异）
                if id(x) in removed_ids:
                    continue
                cp = str(x.get('cp') or '') + str(x.get('summary') or '')
                m = str(x['date'])[:7]
                if not m:
                    continue
                if x['direction'] == '借' and any(k in cp for k in recv_kw):
                    j_recv[m][0] += round(float(x['amount']), 2)
                    j_recv[m][1].append(x)
                elif x['direction'] == '贷' and any(k in cp for k in fee_kw):
                    # ⚡ 网银该月已有同金额支出 → 独立扣费（网银有记录），不净额化
                    if round(float(x['amount']), 2) in bank_exp_amt.get(m, set()):
                        continue
                    j_fees[m].append(x)
            for m in j_recv:
                fees = j_fees.get(m, [])
                if not fees:
                    continue
                fee_amt = round(sum(float(x['amount']) for x in fees), 2)
                if fee_amt <= 0:
                    continue
                recv_amt, recv_rows = j_recv[m][0], j_recv[m][1]
                # ⚡ round 比较防浮点误差：b_one 累加含小数 → 98,048.44999999998 < 98,048.45
                #   → recv<=b_one 误判 False → 第一级已净额化的月份被第二级重复扣（5 月实证
                #   双重扣 170.82 → 收差 +170.82 假差异）
                # ⚡⚡ 2026-08-14 删除 b_one 判定：『网银一码通收合计』与『手续费是否独立支出』
                #   无关（3300 农商行实证：一码通收 14,316 而账借 recv 仅 5,235——账侧分类汇总
                #   收款与网银逐笔跨期错位，recv<b_one 恒成立 → 退押金 270/手续费 23.91 永不
                #   净额化 → 每月收差=支差 -682 对称假差异）。真正判定=bank_exp_amt（网银当月
                #   有无同额支出：有=独立扣费不净额化，无=从收款净额扣）。3100 年费 22.50 网银
                #   有支→bank_exp 排除不受影响；手续费 147.33 网银无支→净额化不变。
                if round(fee_amt, 2) > round(recv_amt, 2):
                    continue
                remain = fee_amt
                for x in sorted(recv_rows, key=lambda x: -float(x['amount'])):
                    if remain <= 0:
                        break
                    amt = float(x['amount'])
                    if amt >= remain:
                        x['amount'] = round(amt - remain, 2)
                        remain = 0.0
                    else:
                        x['amount'] = 0.0
                        remain = round(remain - amt, 2)
                for x in fees:
                    removed_ids.add(id(x))
                info.append(('%s月聚合' % m, round(recv_amt - fee_amt, 2), len(fees)))
        if removed_ids:
            journal[b] = [x for x in rows if id(x) not in removed_ids]
            net[b] = info
    return net


def _disp(x):
    """红字负数归一化还原（reconcile 入口对负数行翻方向+取绝对值参与匹配）→
    返回 (direction, amount) 显示值——底稿中红字仍显示『收 -X』原始语义；
    合计/月度对账用归一化值（reconcile 已就地修改行对象），口径一致。"""
    if not isinstance(x, dict):
        return '', ''
    if x.get('_orig_dir'):
        return x['_orig_dir'], x['_orig_amt']
    return x.get('direction', ''), x.get('amount', '')


def _bank_of_name(name):
    """账侧科目名（银行存款-工商银行-…）→ 银行简称（工行）——2026-08-13 修复：
    原用 BANK_KWS 匹配不归一化 → '工商银行'/'工行' 拆成两组对不上账。"""
    for full in _BANK_FULLS:
        if full in name:
            return _BANK_FULL2SHORT[full]
    for short in _BANK_SHORTS:
        if short in name:
            return '进出口行' if short == '口行' else short
    return '其他'


def read_sap_journal_file(fp, comp):
    """解析 SAP 银行日记账导出（一文件多主体，如『3100-3400银行日记账.XLSX』，公司代码列）→
    {bank: [统一账行]}（仅该主体，2026-08-13 用户需求 3100）。
    列：0公司代码 / 1凭证编号 / 10过帐日期 / 25过账码(40借·50贷) / 27借/贷标识(S/H) /
    30总账科目 / 31科目长文本 / 32金额 / 33文本 / 65银行帐户 / 66帐户标识。"""
    import openpyxl
    wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    out = {}
    n_cx = 0
    cx_amt = 0.0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        if str(row[0]).strip() != comp:
            continue
        date = str(row[10] or '')[:10]
        pcode = str(row[25] or '')
        direction = '借' if pcode.startswith('40') else ('贷' if pcode.startswith('50') else '')
        if not direction:
            s = str(row[27] or '')
            direction = '借' if s == 'S' else ('贷' if s == 'H' else '')
        raw = str(row[32] or '').strip()
        # ⚡ 2026-08-13 修复：SAP 贷方行金额为负（-2535681.71），abs 后方向由过账码决定
        amt = abs(float(raw)) if raw not in ('', '0') else 0.0
        if amt <= 0 or not direction:
            continue
        # ⚡⚡ 2026-08-13 用户点破（Q 列=冲销标识）：col17 冲销标识 1=冲销凭证/2=被冲销凭证
        #   ——成对红冲=**虚增发生额**（原凭证 +X 与红冲 -X 双计，业务净额 0），且跨月/负金额
        #   导致「同月同金额同cp」成对剔除覆盖不到（3100 实证：剔除后全银行收差 -69.7亿→-17.4亿，
        #   上海银行 -4029万→-2元、中信/交行/光大/南京/宁波全平 0）。**匹配前直接剔除**，
        #   只保留空白行（真实发生额）——与用户「银行日记账 Q 列只保留空白、1和2剔除」一致。
        cx = str(row[16] or '') if len(row) > 16 else ''
        if cx in ('1', '2'):
            n_cx += 1
            cx_amt += amt
            continue
        name = str(row[31] or '')
        # ⚡⚡ 2026-08-14 定期/通知存款排除（用户需求，与 U8 路径一致）：网银=活期流水，
        #   定期/通知存款账户剔除防假『账有网银无』
        if any(_k in name for _k in ('定期', '定存', '通知存款', '大额存单', '结构性存款')):
            continue
        # ⚡ 2026-08-13 修正 v2（用户：剔除虚增后发生额应与网银一致）：币种转换/现流转换
        #   = **真实账户购汇/结汇的镜像过渡**（实证：农行结汇 200 万美元=真实账户借'农行-人民币'
        #   1370.7万+贷'农行-美元'200万，镜像走币种转换贷1370.7万+借200万）——**剔除防发生额
        #   双计虚增**，购汇业务由真实账户行与网银购汇流水核对（单列『特殊科目-过渡镜像』）
        if '币种转换' in name or '现流转换' in name:
            bank = '特殊科目'
        else:
            bank = _bank_of_name(name)
        out.setdefault(bank, []).append({
            'bank': bank, 'date': date, 'direction': direction, 'amount': amt,
            'name': name,
            # ⚡ 2026-08-13 修复：col33 是币种（'CNY'）非摘要！真正文本在 col37（对方科目描述）
            #   与 col23（凭证抬头文本）。summary=col23+col37 拼接，保证代扣通道匹配/语义核对可用
            'summary': (str(row[23] or '') + ' ' + str(row[37] or ''))[:60],
            # ⚡ 对方科目：银行日记账 col37 文本=对方科目描述（'短期借款中国农业银行…'）
            'cp': str(row[37] or '').strip()[:40], 'vchar': str(row[1] or ''),
            'account': str(row[66] or '') or str(row[65] or ''),
        })
    wb.close()
    if n_cx:
        print(f'[冲销剔除] {n_cx} 行冲销凭证（Q列标识 1/2，虚增发生额）已剔除，金额 {cx_amt:,.0f} 元')
    return out


def build_bank_reconcile(data_dir, stmt_folder=None, comp=None, year='2026', journal_file=None):
    """主入口：账套目录(+网银文件夹) → 核对底稿 Excel。返回输出路径。
    data_dir 可为 300 根目录（多主体）或单主体子目录（300/6100）：
      ① 子目录 discover 为空 → 自动上溯父目录并取子目录名为 comp
      ② comp 缺省时从网银文件夹名提取（『6100网银…』→ 6100）"""
    import openpyxl
    import sap_adapter as A

    if not os.path.isdir(data_dir):
        print(f'[ERROR] 账套目录不存在：{data_dir}')
        return None
    # 网银文件夹：显式指定，否则自动找同级『*网银*』
    if stmt_folder is None:
        parent = os.path.dirname(os.path.normpath(data_dir))
        cands = [os.path.join(parent, d) for d in os.listdir(parent)
                 if '网银' in d and os.path.isdir(os.path.join(parent, d))]
        if not cands:
            print('[ERROR] 未找到网银明细文件夹（同级需含『网银』的目录，或显式传入）')
            return None
        stmt_folder = cands[0]
    print(f'网银文件夹：{stmt_folder}')

    # ── 定位数据根与主体（子目录自动上溯） ──
    root = data_dir
    A.set_root(root)
    if not A.discover_entities(None):
        up = os.path.dirname(os.path.normpath(data_dir))
        A.set_root(up)
        if A.discover_entities(None):
            root = up
            comp = comp or os.path.basename(os.path.normpath(data_dir))
            print(f'识别为 {up} 下的主体目录，数据根={root}，主体={comp}')
        else:
            print(f'[ERROR] 无法识别账套结构：{data_dir}')
            return None
    # comp：显式 > 网银文件夹名（『6100网银…』→ 6100）> 根目录下唯一主体
    if not comp:
        m = re.search(r'(\d{4})', os.path.basename(os.path.normpath(stmt_folder)))
        if m:
            comp = m.group(1)
        else:
            ents = A.discover_entities(None)
            if len(ents) == 1:
                comp = next(iter(ents))
            else:
                print(f'[ERROR] 多主体（{sorted(ents)[:6]}…）需指定主体（--comp=xxxx）')
                return None
    print(f'主体：{comp}')

    # ⚡⚡ 2026-08-15 账套级业务关键词注入（bank_business_keywords.json 跟随账套走）
    _set_keywords(os.path.basename(os.path.normpath(root)))
    print(f'业务关键词：{len(_CUR_KEYWORDS or _DEFAULT_KEYWORDS)} 组（账套 {os.path.basename(os.path.normpath(root))}）')

    # ── 1. 解析网银 ──
    banks, unknown, total = parse_bank_folder(stmt_folder)
    print(f'网银解析：{total} 笔 / {len(banks)} 家' + (f'；未识别 {unknown}' if unknown else ''))
    # ⚡ 2026-08-13 网银侧银行名与账侧归一化对齐（'进出口'↔'进出口行'/'国开'↔'国开行'，
    #    否则同银行拆两组永远对不上：3100 进出口 14 笔 vs 进出口行 30 行的教训）
    for _k, _v in (('进出口', '进出口行'), ('国开', '国开行'), ('农商银行', '农商行'),
                   ('中国银行', '中行'), ('交通银行', '交行'), ('嘉兴', '嘉兴银行'),
                   # ⚡ 2026-08-14 3200/3300/3400 网银短关键词（'江苏'/'苏州'/'上海'/'泰隆'）
                   #    与长关键词拆组兼容（parse_bank_file 已改返回 parser 标准名，此为兜底）
                   ('江苏', '江苏银行'), ('苏州', '苏州银行'), ('上海', '上海银行'),
                   ('泰隆', '泰隆银行'),
                   # ⚡ 2026-08-14 网银全称 vs 账侧简称归一化（3400 实证：兴业/平安/招行/宁波
                   #    拆两组——网银 parser 返回 '兴业银行'，账侧 _bank_of_name 返回 '兴业'）
                   ('兴业银行', '兴业'), ('平安银行', '平安'), ('招商银行', '招行'),
                   ('宁波银行', '宁波'), ('浦发银行', '浦发'), ('光大银行', '光大'),
                   ('邮储银行', '邮储'), ('民生银行', '民生'),
                   # ⚡⚡ 2026-08-15 漏网：华夏银行（3100 实证：账侧『华夏』24 笔 vs 网银
                   #   『华夏银行』12 笔拆两组 → 配对永远失效，华夏核对全错）——兴业/平安/
                   #   招行/宁波等 2026-08-14 已加，唯华夏漏掉
                   ('华夏银行', '华夏')):
        if _k in banks and _k != _v:
            banks.setdefault(_v, [])
            banks[_v].extend(banks.pop(_k))

    # ── 2. 读银行日记账 ──
    A.set_comp(comp)
    if journal_file:  # ⚡ 2026-08-13 用户提供 SAP 银行日记账文件（3100-3400 多主体）→ 直接解析
        journal = read_sap_journal_file(journal_file, comp)
        print(f'银行日记账文件：{os.path.basename(journal_file)}（{comp} 抽取 {sum(len(v) for v in journal.values())} 笔 / {len(journal)} 家银行）')
    else:
        journal = read_bank_journal(A, comp, year)
        print(f'日记账：{sum(len(v) for v in journal.values())} 笔 / {len(journal)} 家银行')
    # ⚡ 2026-08-13 账侧银行名归一化对齐（'嘉兴'→'嘉兴银行' 等，U8 科目名变体）；
    #    2026-08-14 加 '江苏'/'苏州'/'上海'/'泰隆'（网银短关键词 → 账侧全称键）
    for _k, _v in (('嘉兴', '嘉兴银行'), ('江苏', '江苏银行'), ('苏州', '苏州银行'),
                   ('上海', '上海银行'), ('泰隆', '泰隆银行')):
        if _k in journal and _k != _v:
            journal.setdefault(_v, [])
            journal[_v].extend(journal.pop(_k))
    # ⚡ 2026-08-13 凭证级净额化（同凭证内转借+贷同额=凭证内对冲净额0，防抢配假差异；
    #    仅对网银无『同名/内转』流水的银行——农行网银同名划转=真实内转不剔）
    _vnet = _voucher_net(journal, banks)
    if _vnet:
        print('[凭证净额化] ' + '；'.join(f'{b} {n} 行' for b, n in _vnet.items())
              + '（同凭证内转借+贷同额=内部对冲净额0，剔除防双计）')

    # ⚡ 2026-08-13 收款净额化（农商行二维码收款实证：账侧同凭证/月聚合『借全额+贷手续费』
    #    双计，网银到账=全额-手续费 → 收差=支差一致假差异；净额化后两边配平）
    _fnet = _fee_net_journal(journal, banks)
    if _fnet:
        tot = sum(sum(1 for _v, _n, _nf in v) for v in _fnet.values())
        print(f'[收款净额化] {tot} 张收款凭证 扣减手续费双计（借全额-手续费=网银到账净额，'
              + '；'.join(f'{b} {sum(_nf for _v,_n,_nf in v)} 行' for b, v in _fnet.items())
              + '）')

    # ── 3. 逐银行双向核对 ──
    results = {}
    for b in sorted(set(banks) | set(journal)):
        matched, ub, uj, summary = reconcile(banks.get(b, []), journal.get(b, []))
        # ⚡ 2026-08-13 用户方法论：资金收付单位名称 ↔ 日记账对方科目一致性核对（已匹配对）
        issues = _semantic_check(b, banks.get(b, []), journal.get(b, []), matched)
        # 汇总口径：仅「重大/关注」（方向矛盾=真问题）；「提示」（科目未挂单位名）单独显示
        n_iss = len([i for i in issues if i['level'] in ('重大', '关注')])
        results[b] = {'matched': matched, 'ub': ub, 'uj': uj, 'summary': summary,
                      'issues': issues}
        print(f'  {b}: 网银{summary["bank_total"]} / 账{summary["journal_total"]} / '
              f'匹配{summary["matched"]} / 网银未配{summary["unmatched_b"]} / 账未配{summary["unmatched_j"]}'
              + (f' / ⚠️科目异常{n_iss}' if n_iss else ''))

    # ── 3.5 跨账户归因（资金池划转：网银 A 账户 ↔ 账面 B 账户往来款） ──
    n_cross = _cross_account_match(results, banks, journal)
    if n_cross:
        print(f'[跨账户归因] {n_cross} 笔大额划转已关联到其他银行账户（从网银未配移除，见差异清单「已归因」区）')

    # ── 3.5b 零星小额月合计核销（用户方法论：总数对得上=没问题） ──
    n_small = _small_monthly_settle(results)
    if n_small:
        print(f'[小额核销] {n_small} 笔零星小额（<{_SMALL_AMT:,}元）同月合计两边相等，已核销（从差异移除）')

    # ── 3.5b' 双向对销踢虚增（用户方法论 2026-08-13：同科目反方向两笔=更正/冲销，净额0非异常） ──
    n_contra = _contra_settle(results)
    if n_contra:
        print(f'[账侧对销] {n_contra} 笔同对方科目、反方向、同金额 已对销（财务更正/冲销，非虚增，从差异移除）')
    # ── 3.5b'' 受托支付双计剔除（用户方法论 2026-08-13：放贷+受托支付=一笔业务两笔分录，
    #    银行实际不经该账户 → 账侧『借放贷 X ↔ 贷9999#@!货款 X』同日同额双计净额0，非差异） ──
    n_entrust = _entrust_pay_settle(results)
    if n_entrust:
        print(f'[受托支付双计] {n_entrust} 笔『借放贷/购汇 ↔ 贷9999#@!货款/发货款』同日同额成对 已剔除（银行受托支付双计，净额0非差异）')
    # ── 3.5b''' 内转 1:N 对冲（用户方法论 2026-08-13：内部账户划转借+贷合计对冲净额0非差异） ──
    n_transfer = _inner_transfer_net(results)
    if n_transfer:
        print(f'[内转对冲] {n_transfer} 笔同月『内转』借合计=贷合计 已对冲（内部账户划转拆多笔，净额0非差异）')
    # ⚡ 2026-08-14 资金池划转成对对冲（3400 实证：实归支取借↔工资/社保/税费贷 同月平衡）
    n_pool = _pool_wage_settle(results)
    if n_pool:
        print(f'[资金池对冲] {n_pool} 笔同月『实归支取/内转借=工资/社保/税费贷』已对冲（集团资金池直付，双计净额0非差异）')
    # ⚡ 2026-08-13 修正：网银成对（收↔支同对方同金额）→ 不核销，标重大漏记（资金实际发生账未记）
    n_bank_contra = _classify_bank_unmatched(results, banks=banks, journal=journal)
    if n_bank_contra:
        print(f'[网银未配分类] 标 {n_bank_contra} 笔重大（成对漏记+单边大额；资金池/代扣通道已单列不算差异）')
    # ⚡ 2026-08-13 待报解月汇总核对（用户方法论：报解财务汇总记账、对方应交税费/税金及附加借方、
    #   金额合计一致=核对一致）——网银待报解月合计 vs 账税费贷方月合计
    n_daikou = _daikou_monthly_settle(results, banks, journal)
    if n_daikou:
        print(f'[待报解月汇总] {n_daikou} 笔网银待报解月合计=账税费借方月合计，核对一致（核销）')

    # ── 3.5b''' 期末重要性水平（用户方法论 2026-08-13：期末平的但单边存在→需重要性水平衡量） ──
    n_mat = _materiality_check(results, banks)
    if n_mat:
        print(f'[重要性水平] {n_mat} 家银行期末未达净额占期末余额超 {_MATERIALITY_PCT:.0%}（或≥{_MATERIALITY_AMT:,}元），需评估是否超出审计重要性')

    # ── 3.5b'' 个人大额收付检查（网银有账无侧；分红/借款/工资等除外） ──
    n_personal = 0
    for b in results:
        for u in results[b]['ub']:
            cp_name = str(u.get('counterparty') or '').strip()
            if u['amount'] < _PERSONAL_AMT or not _is_personal_name(cp_name):
                continue
            sm = str(u.get('summary') or '')
            if any(k in sm for k in _PERSONAL_OK):
                continue
            results[b].setdefault('issues', []).append({
                'bank': b, 'date': u['date'], 'direction': u['direction'],
                'amount': u['amount'], 'level': '重大', 'desc': '个人大额收付',
                'bank_text': (cp_name + ' ' + sm).strip()[:50],
                'j_cp': '（网银有账无）', 'j_sm': '',
                'reason': f'个人「{cp_name}」大额{u["direction"]} {u["amount"]:,.0f} 元（非分红/借款/工资类），'
                          + '账上无对应记录，重点核查是否未入账/资金体外循环',
            })
            n_personal += 1
    if n_personal:
        print(f'[个人大额] {n_personal} 笔个人大额收付（非分红/借款/工资类）已标注异常（对方科目核对 sheet）')

    # ── 3.5c 汇总口径更新：网银未匹配/账未匹配改为【净未配】（减跨账户归因+小额核销+对销） ──
    n_total_explained = 0
    for b in results:
        s = results[b]['summary']
        s['unmatched_b'] = len(results[b]['ub'])
        s['unmatched_j'] = len(results[b]['uj'])
        s['explained'] = (len(results[b].get('cross_acct', []))
                          + results[b].get('small_settled', 0)
                          + results[b].get('contra_settled', 0)
                          + len(results[b].get('entrust_rows', []))
                          + len(results[b].get('transfer_rows', []))
                          + len(results[b].get('pool_j_rows', []))
                          + results[b].get('daikou_monthly', 0))
        n_total_explained += s['explained']
    if n_total_explained:
        print(f'[已解释合计] 归因 {n_cross} + 小额核销 {n_small} + 对销 {n_contra} = {n_total_explained} 笔（差异清单绿区，不计差异）')

    # ── 3.6 全年/分月合计对账（用户方法论第 1 层：先对上总额→才有找平可行域） ──
    # ⚡ 2026-08-13 v3：传 results 同步剔除已解释项（对销/核销/跨账户归因）——月度对账差异
    #   与差异清单红区口径一致（3100 中行 5 月内转 3.9 亿对销后月度差异由 -3.9亿 → -4万）
    monthly = _monthly_totals(banks, journal, results)
    for b, d in monthly.items():
        t = d['TOTAL']
        print(f'  [对账] {b}: 全年 网银收{t["b_inc"]:,.2f}/账借{t["j_dr"]:,.2f} '
              f'差{t["diff_inc"]:,.2f} | 网银支{t["b_exp"]:,.2f}/账贷{t["j_cr"]:,.2f} 差{t["diff_exp"]:,.2f}'
              + (' ✅' if t['diff_inc'] == 0 and t['diff_exp'] == 0 else ' ⚠️'))

    # ⚡ 2026-08-14 月度对账先行诊断（用户方法论固化）：先看月度对账——两边对平或收差=支差
    #   （净额平衡）→ 差异清单红区多为【未配平】而非真实差异，应找配平路径（对称项自动计数）
    diag = _symmetry_diag(monthly, results)
    n_bal = sum(1 for d in diag.values() if d['net_balanced'])
    n_sym = sum(1 for d in diag.values() if d['balanced_but_red'])
    print(f'[月度先行诊断] {n_bal} 家净额平衡（收差≈支差→差异多为对称双计/未配平，应找配平路径）；'
          f'{n_sym} 家两边对平但红区仍有未配（对称未配对：N:M 汇总/跨期错位，应继续配对）')

    # ⚡⚡⚡ 2026-08-14 差异智能统一入口（架构收口：把散落的三处体检/知识/学习调用
    #   收敛为单一模块 diff_intelligence.run_intelligence——P1-P12 现象诊断 + 层2
    #   知识兜底 + 学习器候选，统一报告；任何一步失败不影响主流程）。
    # ⚡⚡ 2026-08-14 大主体开关：BANK_NO_INTEL=1 跳过（3500 等 4 万笔未配 → 学习器
    #   S-聚合/词切分全量扫描内存放大，铁律113 大主体优先串行核对，智能体检可后补）。
    if os.environ.get('BANK_NO_INTEL') != '1':
        try:
            from diff_intelligence import run_intelligence as _run_intel
            _run_intel(banks, journal, results, monthly)
        except Exception:
            pass

    # ── 4. 输出 Excel ──
    out_dir = os.path.dirname(os.path.normpath(data_dir)) if os.path.basename(data_dir) == comp else data_dir
    out_path = os.path.join(data_dir, f'银行网银核对底稿_{comp}_{year}.xlsx')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    # ⚡ 2026-08-13 应收回款核对（应收客户贷方回款 vs 网银收款）——暂缓（用户指示先专注双向核对）
    ar_check = {}
    # try:
    #     ar_check = _ar_repayment_check(A, banks, comp)
    # except Exception as _e:
    #     print(f'  ⚠️ 应收回款核对跳过: {_e}')
    _build_sheets(wb, results, banks, journal, year, monthly, ar_check, diag)
    try:
        wb.save(out_path)
    except PermissionError:
        # 原文件被占用（Excel/WPS 打开中）→ 自动另存 _new，不阻塞生成
        alt = out_path.replace('.xlsx', '_new.xlsx')
        wb.save(alt)
        print(f'[提示] 原文件被占用（可能已打开），已另存：{alt}')
        return alt
    except TypeError:
        import openpyxl as _op
        from openpyxl.styles.fills import Fill as _Fill
        bad = [(sh.title, c.coordinate, repr(c.fill)[:80])
               for sh in wb.worksheets for row in sh.iter_rows() for c in row
               if not isinstance(c.fill, _Fill)]
        print('[诊断] 非 Fill 单元格:', bad[:10])
        raise
    print(f'[DONE] 已生成：{out_path}')
    return out_path


# ══════════════════════════════════════════════════════════════════
# 审计语义层（2026-08-13 用户方法论）：网银摘要 ↔ 日记账对方科目一致性核对
# 用户原则：先对全年合计、分月合计（能对上→找平可行域）；再核对对方科目——
#   「网银摘要与日记账对方科目不一致 = 重大问题」，尤其大额款项，绝不能遗漏。
# 启发式预筛：按业务关键词分类，日记账对方科目不在合理科目集内 → 标注待核。
# 误报可能，但作为「审计师人工复核的预筛清单」，宁可多标不可漏。
# ══════════════════════════════════════════════════════════════════
# 业务关键词组 → 合理对方科目关键词集 → 业务说明 → 适用方向（'收'/'支'/'双向'）
# ⚡ 方向过滤是核心：收款只期望收入类科目、付款只期望支出/应付类科目——
#   「收货款」记应收账款=正常、「支电费」记应付账款=正常，方向反了才是异常。
# ⚡ 匹配时【摘要优先、对方名其次】——公司名常含「材料/采购/服务」等干扰词。
_SEMANTIC = [
    # 代扣代缴通道（放最前优先命中）：国库/待报解收缴=社保、个税、税费的代扣通道，
    # 对方科目按实际扣缴内容记（应付职工薪酬/应交税费/其他应收款代垫等）→ 直接放行
    (['代理国库', '待报解', '税收收缴', '代扣代缴', '代扣', '扣缴'],
     ['应付职工薪酬', '应交税费', '其他应收款', '其他应付款', '应付账款', '预付账款'],
     '代扣代缴通道', '双向'),
    (['货款', '销售', '回款', '收款', '租金', '预收', '收入', '充电', '平台', '电商',
      '一码通', '聚合', '扫码', '收款码', '码上'],
     ['应收账款', '预收账款', '预收款项', '主营业务收入', '营业收入', '合同负债', '其他业务收入'],
     '销售/收款', '收'),
    # 收服务费/咨询费 = 服务收入（记应收/其他业务收入正确）——放财务费用组前，
    # 「支服务费」仍由财务费用组拦（方向过滤），「收服务费」不会误报
    (['服务费', '咨询费', '技术服务', '信息服务', '管理服务', '技术开发', '设计费'],
     ['应收账款', '预收账款', '其他业务收入', '营业收入', '合同负债', '其他应付款'],
     '服务收入', '收'),
    (['货款', '采购', '材料', '商品', '进项', '供应商'],
     ['应付账款', '预付账款', '主营业务成本', '存货', '原材料', '库存商品',
      '货款', '退货款', '内部采购'],
     '采购/付款', '支'),
    (['电费', '水费', '物业', '房租', '通行费', '燃气'],
     ['应付账款', '预付账款', '管理费用', '销售费用', '制造费用', '主营业务成本'],
     '费用支付', '支'),
    (['税', '增值税', '附加税', '所得税', '个税', '缴款'],
     ['应交税费', '应交税金', '税金及附加'],
     '税费', '双向'),
    (['工资', '薪酬', '社保', '公积金', '奖金', '津贴', '劳务'],
     ['应付职工薪酬', '应付工资', '管理费用', '销售费用', '研发支出', '生产成本', '制造费用',
      '应付账款', '其他应付款'],
     '薪酬', '支'),
    (['利息', '手续费', '服务费', '账户管理', '年费', '结息'],
     ['财务费用', '利息收入', '利息支出', '利息费用', '管理费用', '销售费用', '研发支出',
      '应付账款', '预付账款', '其他应付款'],
     '财务费用', '双向'),
    (['划转', '转存', '同名', '归集', '下拨', '上划', '调拨',
      '内部转账', '内部划转', '内部往来', '集团内', '资金池'],
     ['银行存款', '其他货币资金', '其他应收款', '其他应付款', '内部往来',
      '资金归集', '资金池', '内转', '归集', '上存', '下存'],
     '内部划转', '双向'),
    (['借款', '贷款', '还款', '放款', '授信', '贴现'],
     ['短期借款', '长期借款', '应付票据', '其他应付款', '其他应收款',
      '委贷', '委托贷款', '内部', '受托', '结汇', '售汇'],
     '借贷', '双向'),
    (['押金', '保证金', '备用金', '往来', '代垫'],
     ['其他应收款', '其他应付款', '归集', '上存', '资金池', '内转'],
     '往来', '双向'),
    (['罚款', '赔偿', '捐赠', '违约金', '滞纳金'],
     ['营业外收入', '营业外支出'],
     '营业外', '双向'),
]
# 大额阈值（元）：≥此金额的科目异常 → 标注「重大」
_MAJOR_AMT = 1_000_000
# 零星小额阈值（元）：单笔 < 此金额 且 同月同方向合计两边相等 → 直接核销（用户方法论：
# 「小额10万以下，只要记账方向没错（摘要不清晰也没关系，汇总记账不写具体），
#   月合计对上=没问题」——月汇总记账场景，无需逐笔配对）
_SMALL_AMT = 100_000
# ⚡ 2026-08-13 收入类对方科目（用户方法论：汇总记账若直接记收入=『收入未过往来』，标记提示）
_INCOME_KWS = ('主营业务收入', '营业收入', '其他业务收入', '收入')
# ⚡ 2026-08-13 个人大额收付阈值（用户方法论：现在很少现金/个人与银行之间大额往来，分红/借款除外）
_PERSONAL_AMT = 50_000
# 个人大额但属正常往来 → 放行（摘要含这些词）
_PERSONAL_OK = ('分红', '借款', '工资', '薪酬', '报销', '社保', '公积金', '个税', '代扣',
                '利息', '劳务', '手续费', '退款', '备用金', '奖金', '津贴', '补助', '补贴',
                '货款', '往来款')
# 企业/机构特征词 → 非个人名
_CORP_SUFFIX = ('公司', '有限', '股份', '集团', '银行', '厂', '中心', '社', '所', '院', '局',
                '委员会', '合作社', '商店', '超市', '支行', '分行', '保险', '证券', '基金',
                '信托', '租赁', '科技', '贸易', '实业', '发展', '建设', '材料', '纤维',
                '能源', '生物', '电子', '纺织', '化工', '高科', '企业', '商务', '供应',
                '大学', '学院',
                '营业部', '政府', '学校', '医院', '事务所', '部队', '税务', '管理', '市场')


def _is_personal_name(s):
    """粗略判断收付单位是否为个人姓名：去噪后纯中文 2-6 字、不含企业/机构特征词。
    排除通道名/机构代收名（暂收款/待报解/国库/海关等中行 csv 关税代缴对手方）。"""
    s = re.sub(r'[\s\d（()）【】\[\]·:：\-—\\/,。]+', '', str(s or ''))
    if not re.fullmatch(r'[\u4e00-\u9fa5]{2,6}', s):
        return False
    if any(k in s for k in _CORP_SUFFIX):
        return False
    if any(k in s for k in _PERSONAL_OK + ('暂收', '国库', '税务', '海关', '待报解',
                                           '代理', '财政', '收缴', '汇兑', '结转', '提现',
                                           '结售汇', '远期', '外汇', '信用卡')):
        return False
    return True


def _cp_norm(s):
    """对方科目归一化（取最长中文段并截 6 字），供双向对销匹配：
    '662245281401#@!预付国望高科2月90%电费' 与 '预付国望高科电费' → 同键。"""
    segs = re.findall(r'[\u4e00-\u9fa5]{2,}', str(s or ''))
    cp = max(segs, key=len) if segs else ''
    return cp[:6]


# 公司后缀：从单位名提取主体时剥离（有限公司/集团/厂/支行…）
_SUFFIXES = ('股份有限公司', '有限责任公司', '有限公司', '责任公司', '集团公司', '有限公司分公司',
             '集团', '分公司', '支行', '分行', '银行', '信用社', '中心', '工厂', '厂', '公司',
             '协会', '事务所', '学校', '医院', '合作社', '经营部', '门市部', '网点', '代理点',
             '结算中心', '托管中心', '账户', '专户', '备付金')


def _unit_core(s):
    """提取单位名主体（去公司后缀/标点，取最长非空段）→ 供名称级核对。"""
    s = str(s or '')
    s = re.sub(r'[（）()【】\[\]·\s:：\-—]+', '', s)
    for suf in _SUFFIXES:
        s = s.replace(suf, '')
    s = s.strip('企业账户活期定期一般基本专用户头账号卡折对公结算往来')
    return s if len(s) >= 2 else ''


# ⚡ 2026-08-13 用户方法论（审计实务核心）：
#   「一般记账，大额都过往来账项（应收/应付/预收/预付/其他应收应付/合同类）；
#    若大额未过往来（直接记收入/采购/费用），只能看两边摘要是否一致 → 标黄人工核；
#    零星小额不过往来正常；审计最怕大额遗漏或大额异常没发现。」
_WANGLAI_KWS = ['应收账款', '预收账款', '应付账款', '预付账款', '其他应收款', '其他应付款',
                '合同资产', '合同负债', '应收票据', '应付票据', '备用金', '内部往来',
                '应收利息', '应付利息', '长期应收款', '长期应付款', '应收股利', '应付股利',
                '内部', '委贷', '委托贷款']
_PL_KWS = ['收入', '采购', '成本', '费用', '原材料', '库存商品', '存货', '主营业务',
           '其他业务', '营业外', '研发', '销售费用', '管理费用', '财务费用', '制造费用',
           '生产成本', '利息']


def _sems_of(text):
    """文本命中的业务组名集合（供摘要一致性判断）。"""
    return {x[2] for x in _SEMANTIC if any(k in str(text) for k in x[0])}


def _bigram_sim(a, b):
    """中文 bigram 相似度（0~1）：两边摘要的公共二字词比例 → 摘要一致性兜底。"""
    def _bg(s):
        s = re.sub(r'[^\u4e00-\u9fa5]', '', str(s))
        return {s[i:i + 2] for i in range(len(s) - 1)} if len(s) >= 2 else set()
    sa, sb = _bg(a), _bg(b)
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / max(1.0, min(len(sa), len(sb)))


def _ar_repayment_check(adapter, banks, comp, threshold=100_000):
    """应收回款核对（用户方法论 2026-08-13）：应收账款明细中【每个客户贷方发生额（回款）】
    vs 网银收到的该客户款项（收款方向）。客户回款（应收贷方）应≈网银收款；大额差异=追查项。
    - 应收贷方 > 网银收款：回款含非银行方式（票据/预收冲抵/坏账核销/未到账）
    - 应收贷方 < 网银收款：收款挂预收/未冲应收/客户名不匹配
    返回 {客户: dict(recv_cr, bank_inc, diff)} 仅差异≥threshold 或任一侧为大额。"""
    from collections import defaultdict
    out = {}
    if adapter is None:
        return out
    try:
        rows = adapter.read_gl_net(comp=comp, pfx='应收账款')
    except Exception:
        try:
            rows = adapter.read_gl_net(comp=comp, pfx='1122')
        except Exception:
            return out
    if not rows:
        return out
    recv_cr = defaultdict(float)
    for r in rows:
        cust = str(r.get('aux') or r.get('cust') or r.get('supp') or '').strip()
        cr = float(r.get('cr') or 0)
        if cust and cr:
            recv_cr[cust] += cr
    bank_inc = defaultdict(float)
    for b, bs in banks.items():
        for x in bs:
            if x['direction'] == '收':
                cp = str(x.get('counterparty') or '').strip()
                if cp:
                    bank_inc[cp] += x['amount']
    # 客户名 ↔ 网银单位 名称级匹配
    bank_by_core = {}
    for cp, amt in bank_inc.items():
        bank_by_core.setdefault(_unit_core(cp) or cp, []).append((cp, amt))
    for cust, cr in recv_cr.items():
        core = _unit_core(cust) or cust
        hits = []
        for bcore, items in bank_by_core.items():
            if bcore and (core in bcore or bcore in core):
                hits.extend(items)
        bank_amt = sum(a for _, a in hits)
        diff = cr - bank_amt
        if abs(diff) >= threshold or cr >= threshold or bank_amt >= threshold:
            out[cust] = {'recv_cr': cr, 'bank_inc': bank_amt, 'diff': diff,
                         'bank_cp': '; '.join(cp for cp, _ in hits[:3])}
    return out


def _semantic_check(b, bank_rows, journal_rows, matched, threshold=_MAJOR_AMT):
    """资金收付单位/摘要 ↔ 对方科目 审计路径核对（用户方法论 2026-08-13）：
    ① 单位名主体出现在对方科目中 → ✅ 一致
    ② 大额（≥阈值）：过往来账项 → ✅ 规范；未过往来（直接损益/成本）→
       两边摘要一致 → 🟡 关注（人工核）；摘要不一致 → 🔴 重大
    ③ 小额：仅「方向与科目性质矛盾」标注（收→支出类/支→收入类）；其余忽略
    返回 [{bank,date,direction,amount,level,desc,bank_text,j_cp,j_sm,reason}]"""
    issues = []
    for m in matched:
        bb = m[0] if isinstance(m[0], list) else [m[0]]
        jr = m[1] if isinstance(m[1], dict) else (m[1][0] if m[1] else {})
        if not isinstance(jr, dict):
            continue
        for b2 in bb:
            sm = str(b2.get('summary') or '').strip()
            cp_name = str(b2.get('counterparty') or '').strip()
            j_cp = str(jr.get('cp') or '').strip()
            j_sm = str(jr.get('summary') or '').strip()
            amt = b2['amount']
            if not j_cp:
                continue
            # ── ① 单位名 → 对方科目 名称级匹配（最高优先） ──
            if cp_name:
                core = _unit_core(cp_name)
                if core and (core in _unit_core(j_cp) or _unit_core(j_cp) in core):
                    continue  # ✅ 收付单位名出现在对方科目中 → 一致
            bank_text = (cp_name + ' ' + sm).strip()
            if not bank_text:
                continue
            # ── ①b 个人大额收付（用户方法论 2026-08-13）：现在很少现金/个人与银行之间
            #     发生大额资金往来；个人大额 = 异常（分红款/借款/工资报销等除外） ──
            if amt >= _PERSONAL_AMT and _is_personal_name(cp_name):
                if any(k in (sm + ' ' + j_sm + ' ' + j_cp) for k in _PERSONAL_OK):
                    continue  # 分红/借款/工资/报销等正常个人往来 → 放行
                issues.append({
                    'bank': b, 'date': b2['date'], 'direction': b2['direction'],
                    'amount': amt, 'level': '重大', 'desc': '个人大额收付',
                    'bank_text': bank_text, 'j_cp': j_cp, 'j_sm': j_sm,
                    'reason': f'个人「{cp_name}」大额{b2["direction"]} {amt:,.0f} 元（非分红/借款/工资类），'
                              + '现实务中个人与银行大额往来极少，重点核查是否公款私存/未入账/资金体外循环',
                })
                continue
            # ── ② 大额路径检查（用户方法论核心：大额必须过往来） ──
            if amt >= threshold:
                if any(k in j_cp for k in _WANGLAI_KWS):
                    continue  # ✅ 大额过往来账项 → 规范（哪怕科目未含单位名）
                # 代扣代缴通道（代理国库/待报解收缴税/社保/个税）：
                # 科目记应交税费/应付职工薪酬/代垫（研发-专家咨询=代扣个税）均正常 → 放行
                if any(k in sm for k in ('代理国库', '待报解', '税收收缴', '代扣', '扣缴')):
                    continue
                if any(k in j_cp for k in _PL_KWS):
                    # 大额未过往来（直接记损益/成本）→ 只能靠摘要一致性判断
                    g_b = _sems_of(sm)
                    g_j = _sems_of(j_sm)
                    core = _unit_core(cp_name) if cp_name else ''
                    unit_hit = bool(core and core in j_sm)
                    same = bool(g_b & g_j) or unit_hit or _bigram_sim(sm, j_sm) >= 0.2
                    # ⚡ 2026-08-13 合同号/订单号记账（'GW2511270042'/'TEMSH202510-02' 等）
                    #     = 按合同/订单维度记往来，非虚增特征 → 摘要不一致也仅「关注」
                    is_contract_no = bool(re.search(r'[A-Za-z]{2,8}\d{4,}', j_cp)) or '#@!' in j_cp
                    level = '关注' if (same or is_contract_no) else '重大'
                    issues.append({
                        'bank': b, 'date': b2['date'], 'direction': b2['direction'],
                        'amount': amt, 'level': level,
                        'desc': '大额未过往来',
                        'bank_text': bank_text, 'j_cp': j_cp, 'j_sm': j_sm,
                        'reason': ('大额未过往来账项（直接记损益/成本），网银与日记账摘要'
                                   + ('一致，需人工复核业务真实性' if same
                                      else '不一致 → 重大异常，重点核查')),
                    })
                    continue
                # 大额但科目性质特殊（银行存款互转/应交税费等）→ 业务组矛盾检查
                hit = next((x for x in _SEMANTIC
                            if x[3] in ('双向', b2['direction'])
                            and any(k in sm for k in x[0])), None)
                if hit and not (any(k in j_cp for k in hit[1])
                or any(k in j_cp for k in hit[0])):  # 关键词回环：对方科目含触发词=记账一致
                    # ⚡ 2026-08-13 定级校准（用户方法论「真正异常的笔数是很少的」）：
                    #   业务组用词不匹配 ≠ 虚增——多为集团资金调拨/受托支付/付息/预付等
                    #   科目性质合理的用词变体 → 降级「关注」（人工核）。
                    #   「重大」只留给：大额未过往来+摘要不一致、个人大额收付。
                    issues.append({
                        'bank': b, 'date': b2['date'], 'direction': b2['direction'],
                        'amount': amt, 'level': '关注', 'desc': hit[2],
                        'bank_text': bank_text, 'j_cp': j_cp, 'j_sm': j_sm,
                        'reason': f'大额{hit[2]}类业务，对方科目「{j_cp}」用词不符（科目性质待核，集团资金调拨/受托支付常见），人工复核',
                    })
                continue
            # ── ③ 小额：仅方向与科目性质矛盾才标注 ──
            hit = next((x for x in _SEMANTIC
                        if x[3] in ('双向', b2['direction'])
                        and any(k in sm for k in x[0])), None)
            if hit and not (any(k in j_cp for k in hit[1])
                or any(k in j_cp for k in hit[0])):  # 关键词回环：对方科目含触发词=记账一致
                issues.append({
                    'bank': b, 'date': b2['date'], 'direction': b2['direction'],
                    'amount': amt, 'level': '关注', 'desc': hit[2],
                    'bank_text': bank_text, 'j_cp': j_cp, 'j_sm': j_sm,
                    'reason': f'{hit[2]}类业务方向与对方科目「{j_cp}」性质矛盾（小额，注意异常）',
                })
    return issues


def _big_amt_rows(results, banks, journal):
    """大金额专项核对（用户方法论 2026-08-14：『大金额的核对是要非常仔细的』）——
    ① 摘要名称是否与对方名称一致（网银对方名 ↔ 账侧摘要/对方科目）
    ② 入账科目是否合乎逻辑（大额应过往来：应收/应付/预收/预付；直接损益=需复核）
    ③ 个人大额进出是否出现（个人名+大额=重点核查，分红/借款/工资除外）
    ④ 入账是否准确（未配大额=重点查）
    提取所有银行大额交易（已匹配对+未匹配），逐笔标注核对要点。返回行列表（按金额降序）。"""
    rows = []
    for b in sorted(results):
        res = results[b]
        for m in res.get('matched', []):
            bb = m[0] if isinstance(m[0], list) else [m[0]]
            jr = m[1] if isinstance(m[1], dict) else (m[1][0] if isinstance(m[1], list) and m[1] else {})
            if not isinstance(jr, dict):
                continue
            j_cp = str(jr.get('cp') or '')
            j_sm = str(jr.get('summary') or '')
            for b2 in bb:
                amt = float(b2.get('amount') or 0)
                if amt < _MAJOR_AMT:
                    continue  # 仅大额（≥100万）
                cp = str(b2.get('counterparty') or '')
                sm = str(b2.get('summary') or '')
                # ① 名称一致性：网银对方名核心 vs 账侧摘要/对方科目
                #   ⚡ 资金池/内转/待报解类业务（网银显示渠道名/空，账侧记"资金归集/实归/内转"）
                #     名称天然不同——标"资金池类"而非"名称不同"（3400 实证 966 笔大额多为此类，
                #     避免审计员疲于核查已知合理差异）
                _pool_kws = ('资金归集', '内转', '实归', '支取', '归集', '划转', '待报解', '代发', '批量')
                if cp and any(k in (sm + j_sm + j_cp) for k in _pool_kws):
                    name_ok = '资金池类'
                else:
                    cb = _unit_core(cp) if cp else ''
                    cj = _unit_core(j_sm) if j_sm else ''
                    cc = _unit_core(j_cp) if j_cp else ''
                    if not cp:
                        name_ok = '—'
                    elif cb and (cb in (cj + cc) or (cj and cb[:4] in cj) or (cc and cb[:4] in cc)):
                        name_ok = '✅ 一致'
                    elif cb and (cj or cc):
                        name_ok = '⚠️ 名称不同'
                    else:
                        name_ok = '—'
                # ② 入账科目逻辑
                if any(k in j_cp for k in _WANGLAI_KWS):
                    acct = '✅ 过往来'
                elif any(k in j_cp for k in _PL_KWS):
                    acct = '⚠️ 直接损益/成本'
                else:
                    acct = '观察'
                # ③ 个人大额
                personal = ''
                if _is_personal_name(cp) and amt >= _PERSONAL_AMT:
                    personal = ('正常' if any(k in (sm + j_sm + j_cp) for k in _PERSONAL_OK)
                                else '🔴 个人大额')
                rows.append([b, str(b2.get('date'))[:10], b2.get('direction'), amt,
                             (cp + ' ' + sm).strip()[:36], j_sm[:28], j_cp[:28],
                             m[2] if len(m) > 2 else '', name_ok, acct, personal])
        # 未匹配大额（入账准确性：大额未配=重点查）
        for u in res.get('ub', []):
            if float(u.get('amount') or 0) >= _MAJOR_AMT:
                rows.append([b, str(u.get('date'))[:10], u.get('direction'), u['amount'],
                             (str(u.get('counterparty') or '') + ' ' + str(u.get('summary') or '')).strip()[:36],
                             '', '', '❌ 网银有账无', '🔴 未配', '大额未配', '🔴 重点查'])
        for u in res.get('uj', []):
            if float(u.get('amount') or 0) >= _MAJOR_AMT:
                rows.append([b, str(u.get('date'))[:10], u.get('direction'), u['amount'],
                             '', str(u.get('summary') or '')[:28], str(u.get('cp') or '')[:28],
                             '❌ 账有网银无', '🔴 未配', '大额未配', '🔴 重点查'])
    rows.sort(key=lambda r: -r[3])
    return rows


def _monthly_totals(banks, journal, results=None):
    """全年/分月合计对账（用户方法论第 1 层，2026-08-13 净额化 v3）：
    逐银行逐月 网银收/支 vs 账借/贷。
    ⚡ 净额化：①排除特殊科目（币种转换/现流转换=过渡镜像，防发生额双计）；
    ②账侧【冲销凭证（Q 列标识 1/2）已在 read_sap_journal_file 层剔除】——SAP 官方冲销标记
    （1=冲销凭证/2=被冲销凭证，红字成对=虚增发生额，3100 实证剔除后全银行收差 -69.7亿→-17.4亿）；
    ③⚡ v3（2026-08-13 用户指出「差异清单与月度对账口径打架」）：同步剔除【已解释项】——
    账侧对销（contra_rows：同cp反方向同金额成对，净额0非差异，如 3100 中行 5 月内转借+贷 3.9 亿
    一般户↔数币户）+ 小额核销（small_*_rows：月合计两边相等）+ 跨账户归因（cross_acct：网银 A 户
    ↔ 账 B 户同笔资金，两边对称剔除）——剔除后月度对账差异列 = 差异清单红区口径，两表相互印证。
    待报解月汇总（daikou）**不剔除**：网银待报解支与账税费贷对称存在于两侧，差额本就为 0。
    ⚡ 不再做「同月同金额同cp」成对剔除——3100 实证误杀工行 52 亿真实业务
    （外币掉期等同额借+贷是真实资金往来，非虚增），虚增剔除以 Q 列为准。
    返回 {bank: {month: {b_inc,b_exp,j_dr,j_cr,diff_inc,diff_exp}, 'TOTAL': {...}}}"""
    from collections import defaultdict
    # 已解释行 id 集合（每银行分 网银侧 b / 账侧 j）——用 id() 定位原始 banks/journal 行
    expl = defaultdict(lambda: {'b': set(), 'j': set()})
    if results:
        for b in results:
            r = results[b]
            # ⚡⚡ 2026-08-14 对称找和配对（reconcile 内新增）：网银 1 笔=账 1 笔同额
            #   / N 笔拆账求和——已配对即非差异，月度对账须同步剔除（否则月度 vs
            #   清单口径差：3400 建行支差 -360 vs 清单 -730 实证）。
            #   ⚡⚡⚡ 2026-08-14 v2：同步剔除【关键字汇总/关键字找和/单位倍数】配对行——
            #   否则月度对账账侧含关键字配对行、清单已移除 → 打架（3900 建行
            #   收国望退货款 5,000 万被配对后：清单 0 笔但月度收差 -5,000 万实证）。
            for _m in r.get('matched', []):
                _tag = str(_m[2] if len(_m) > 2 else '')
                # ⚡ 月汇总匹配也剔除：配对前提=当月两边总额平衡，配对行两边对称剔除
                #   （3900 建行收国望退货款 5,000 万被月汇总配给网银收行→清单 0 但月度
                #   -5,000 万打架实证；剔除后月度=清单口径统一）
                if not ('对称' in _tag or '关键字' in _tag or '单位倍数' in _tag
                        or '月汇总' in _tag):
                    continue
                _bb = _m[0] if isinstance(_m[0], list) else [_m[0]]
                _jj = _m[1] if isinstance(_m[1], list) else [_m[1]]
                for _x in _bb:
                    expl[b]['b'].add(id(_x))
                for _x in _jj:
                    expl[b]['j'].add(id(_x))
            for x, jb, jx in r.get('cross_acct', []):
                expl[b]['b'].add(id(x))
                expl[jb]['j'].add(id(jx))
            for x in r.get('small_b_rows', []):
                expl[b]['b'].add(id(x))
            for x in r.get('small_j_rows', []):
                expl[b]['j'].add(id(x))
            for x in r.get('contra_rows', []):
                expl[b]['j'].add(id(x))
            for x in r.get('entrust_rows', []):
                expl[b]['j'].add(id(x))
            for x in r.get('transfer_rows', []):
                expl[b]['j'].add(id(x))
            for x in r.get('pool_j_rows', []):
                expl[b]['j'].add(id(x))  # ⚡ 2026-08-14 资金池划转双计（账侧实归支取借↔工资税费贷）
            for x in r.get('pool_rows', []):
                expl[b]['b'].add(id(x))  # ⚡ 2026-08-14 资金池成对（网银侧收+支同额同单位）
            for x in r.get('daikou_b_rows', []):
                expl[b]['b'].add(id(x))  # ⚡ 2026-08-14 待报解已核销网银支（对称剔除）
            for x in r.get('daikou_j_rows', []):
                expl[b]['j'].add(id(x))  # ⚡ 2026-08-14 待报解已核销账税费贷（对称剔除）
    out = {}
    for b in sorted(set(banks) | set(journal)):
        if b == '特殊科目':
            continue  # 过渡镜像剔除，不参与月度对账
        bm = defaultdict(lambda: [0.0, 0.0])
        for x in banks.get(b, []):
            if id(x) in expl[b]['b']:
                continue  # 已解释（对销/核销/跨账户归因），不参与月度对账
            m = str(x['date'])[:7]
            if m:
                bm[m][0 if x['direction'] == '收' else 1] += float(x['amount'])
        jm = defaultdict(lambda: [0.0, 0.0])
        for x in journal.get(b, []):
            if id(x) in expl[b]['j']:
                continue  # 已解释（对销/核销/跨账户归因），不参与月度对账
            m = str(x['date'])[:7]
            if m:
                jm[m][0 if x['direction'] == '借' else 1] += float(x['amount'])
        months = sorted(set(bm) | set(jm))
        d = {}
        tb = [0.0, 0.0]; tj = [0.0, 0.0]
        for m in months:
            bi, be = bm[m][0], bm[m][1]
            jd, jc = jm[m][0], jm[m][1]
            d[m] = {'b_inc': bi, 'b_exp': be, 'j_dr': jd, 'j_cr': jc,
                    'diff_inc': round(bi - jd, 2), 'diff_exp': round(be - jc, 2)}
            tb[0] += bi; tb[1] += be; tj[0] += jd; tj[1] += jc
        d['TOTAL'] = {'b_inc': tb[0], 'b_exp': tb[1], 'j_dr': tj[0], 'j_cr': tj[1],
                      'diff_inc': round(tb[0] - tj[0], 2), 'diff_exp': round(tb[1] - tj[1], 2)}
        out[b] = d
    return out


def _symmetry_diag(monthly, results, tol=100.0):
    """月度对账先行诊断（用户方法论 2026-08-14 固化）：
    用户原则：「先看月度对账数据——两边如果对平，或借贷差异相同（收差=支差），那产生
    大量差异列表可能是没配平；当然也可能往来账户名称不一致/摘要有明显区别/涉及个人款项
    应列示」。本函数把该原则固化为自动诊断：
    ① 净额平衡（借贷差异相同）：|收差 - 支差| < 容差 → 两边净流入相等 → 差异多为对称
       双计/未配对（应收配平），差异清单红区应【找配平路径】而非视为真实差异
       （3100 农商行手续费双计 / 中行补助+内转成对 实证）
    ② 两边对平但红区非空：收差≈0 且支差≈0 但差异清单有项 → 对称未配对（N:M 汇总记账/
       跨期错位，如待报解↔社保、二维码↔一码通），应继续配对找平
    ③ 红区对称项计数：ub收X↔uj贷X（跨侧对称，N:M 汇总记账未配对）/ uj 借X+贷X 同额
       （账侧成对双计，净额0）/ ub 收X+支X 同额（网银侧成对，资金池/漏记）——
       对称项越多，说明『未配平』而非『真实差异』的可能性越大
    返回 {bank: {'net_balanced': 收差≈支差, 'balanced_but_red': 对平但红区非空,
                'sym_pairs': 对称项数}}"""
    from collections import Counter
    diag = {}
    for b in monthly:
        if b == '特殊科目':
            continue
        t = monthly[b]['TOTAL']
        r = results.get(b, {})
        ub, uj = r.get('ub', []), r.get('uj', [])
        d = {'net_balanced': False, 'balanced_but_red': False, 'sym_pairs': 0}
        diff_inc, diff_exp = t['diff_inc'], t['diff_exp']
        # 容差：绝对值 100 或 总额的 0.01%（净流入平衡判断，允许汇兑/取整尾差）
        tol2 = max(tol, abs(t['b_inc'] + t['b_exp']) * 0.0001)
        if abs(diff_inc - diff_exp) < tol2:
            d['net_balanced'] = True
        if abs(diff_inc) < tol2 and abs(diff_exp) < tol2 and (ub or uj):
            d['balanced_but_red'] = True
        # ③ 红区对称项
        pairs = 0
        ub_cnt = Counter((x['direction'], round(x['amount'], 2)) for x in ub)
        uj_cnt = Counter((x['direction'], round(x['amount'], 2)) for x in uj)
        for (dd, amt), n in ub_cnt.items():
            mate = '贷' if dd == '收' else '借'
            pairs += min(n, uj_cnt.get((mate, amt), 0))
        uj_drc = Counter()
        for x in uj:
            uj_drc[round(x['amount'], 2)] += 1
            # 账侧成对：同额出现 ≥2 次且有借有贷
        uj_dir = {}
        for x in uj:
            uj_dir.setdefault(round(x['amount'], 2), []).append(x['direction'])
        for amt, dirs in uj_dir.items():
            if '借' in dirs and '贷' in dirs:
                pairs += min(dirs.count('借'), dirs.count('贷'))
        ub_dir = {}
        for x in ub:
            ub_dir.setdefault(round(x['amount'], 2), []).append(x['direction'])
        for amt, dirs in ub_dir.items():
            if '收' in dirs and '支' in dirs:
                pairs += min(dirs.count('收'), dirs.count('支'))
        d['sym_pairs'] = pairs
        diag[b] = d
    return diag


def _bank_unmatch_reason(b):
    """网银未配差异归因（帮审计判断性质，2026-08-13；2026-08-14 增强：
    外汇业务/工资代发/内部划转形态全覆盖——3100 江苏/工行、3400 建行、
    DQ 嘉兴实证的差异形态统一分类，审计一眼看出该找什么流水）。"""
    cp = str(b.get('counterparty') or '')
    sm = str(b.get('summary') or '')
    amt = b['amount']
    _all = cp + sm
    # ⚡⚡ 2026-08-14 外汇业务（结汇/购汇/掉期/外币贷款）——账侧人民币变动，网银人民币流水不覆盖
    if any(k in _all for k in ('结汇', '购汇', '售汇', '掉期', '外币', '欧元', '美元', '港币')):
        return '外汇业务（结汇/购汇/掉期/外币贷款）——账侧人民币账户变动，网银人民币流水不覆盖，需索取外币/掉期账户流水'
    # ⚡⚡ 2026-08-14 工资/社保/公积金代发（网银逐笔 vs 账侧月汇总）
    if any(k in _all for k in ('工资', '奖金', '社保', '公积金', '代发', '生育金', '医保')):
        return '工资/社保/公积金代发（网银逐笔 vs 账侧月汇总），按月核对汇总金额，差额为真实差异'
    # ⚡⚡ 2026-08-14 内部划转（资金归集/内转/同名）
    if any(k in _all for k in ('资金归集', '内转', '划转', '归集', '集中支付')):
        return '内部划转（资金归集/内转/同名划转）——核查对方账户流水/现流科目，网银显示收付单位与银行流水不一致'
    # 已有分类（保留）
    if '星星' in cp or '电商' in cp or '平台' in cp or '充电' in cp:
        return '电商平台小额（星星充电等，账按月汇总记账），建议按月级核对合计'
    if '待报解' in cp or '财政' in cp or '医保' in cp or '社保' in cp or '税款' in cp:
        return '财政/税费/社保类（账记其他应付款等科目），核查科目归属'
    if amt >= 1e6:
        return '大额网银未配，重点核查（未入账/记错科目/跨期/网银未提供账户？）'
    return '网银存在但日记账未找到（未入账/漏记/跨期/记错科目？）'


def _journal_unmatch_reason(j):
    """账未配差异归因（2026-08-14 增强：外汇/工资/内转形态全覆盖）。"""
    sm = str(j.get('summary') or '')
    acct = str(j.get('account') or '')
    _all = sm + acct
    if '数字人民币' in acct:
        return '数字人民币账户（网银未提供该账户明细）'
    if sm.startswith('9999#@!'):
        return '特殊凭证（9999#@! 前缀，多为以前年度/特殊调整），核查'
    # ⚡⚡ 2026-08-14 外汇业务（结汇/购汇/掉期/外币贷款/还本付息）
    if any(k in _all for k in ('结汇', '购汇', '售汇', '掉期', '外币', '欧元', '美元', '港币', '还本金', '还利息')):
        return '外汇业务（结汇/购汇/掉期/外币贷款还本付息）——网银人民币流水不覆盖，需索取外币/贷款账户流水'
    # ⚡⚡ 2026-08-14 工资/社保/公积金代发（账侧月汇总 vs 网银逐笔）
    if any(k in _all for k in ('工资', '奖金', '社保', '公积金', '代发', '生育金', '医保')):
        return '工资/社保/公积金代发（账侧月汇总 vs 网银逐笔），按月核对，差额为真实差异'
    # ⚡⚡ 2026-08-14 内部划转（资金归集/内转/同名）
    if any(k in _all for k in ('资金归集', '内转', '划转', '归集', '集中支付')):
        return '内部划转（资金归集/内转/同名）——核查对方账户流水，网银可能未导该内部账户'
    return '日记账存在但网银未找到（网银未提供该账户/虚记/内部转账/跨期？）'


def _unmatch_cross_ref(u, pool):
    """月度对账印证分类（用户方法论 2026-08-13）：差异行在同银行剩余未配中是否有
    【同月反方向同金额】配对：
    - 'pair'   → 借贷成对（净额 0 但发生额双计）：月度对账余额可能平、发生额有差异 → 需进一步判断
    - 'single' → 单侧差异：月度对账余额两边不一致（余额调节表调节项）"""
    op = {'收': '支', '支': '收', '借': '贷', '贷': '借'}.get(u['direction'])
    m = str(u['date'])[:7]
    amt = round(u['amount'], 2)
    for x in pool:
        if x is u:
            continue
        if x['direction'] == op and str(x['date'])[:7] == m and round(x['amount'], 2) == amt:
            return 'pair'
    return 'single'


def _sym_pair_ub_uj(res, tol=0.01):
    """⚡⚡ 2026-08-14 跨侧金额对称检测：同一银行内『网银有账无』(ub) 与
    『账有网银无』(uj) 若存在同额对应（含 1:N / N:M 金额拆分求和），
    说明两侧金额一致、应配未配（跨期/名称差异/汇总拆分），非真差异。

    用户方法论（3100 建行实证）：网银『无锡市晨阳 电子转账 55,776』= 账侧
    『客户收款-无锡市晨阳 55,776』，金额一致却因跨期 6 个月+名称前缀不同
    双双落入差异清单——月度对账差很小但差异清单大额。本检测把这类对称
    项标记为【应配未配·金额对称】，从红区真差异降级为待核平。

    返回 {ub_idx: [uj_idx,...]}（贪心同额 + 拆分求和配对）。
    """
    from collections import defaultdict
    ub = res.get('ub', [])
    uj = res.get('uj', [])
    if not ub or not uj:
        return {}
    # 1:N 同额：uj 按金额索引（含求和），先做 1:1 精确同额
    j_by_amt = defaultdict(list)
    for ji, j in enumerate(uj):
        j_by_amt[round(float(j.get('amount') or 0), 2)].append(ji)
    pair = {}
    used_j = set()
    for bi, b in enumerate(ub):
        amt = round(float(b.get('amount') or 0), 2)
        for ji in j_by_amt.get(amt, []):
            if ji not in used_j:
                pair[bi] = [ji]
                used_j.add(ji)
                break
    # N:1 / N:M 求和配对：未配的 uj 组合求和 == 单笔 ub（拆账）
    # ⚡⚡ 2026-08-14 增强：候选不再限 ≤8 项，改为【同方向+同月+金额≤target】过滤 +
    #   排序剪枝（3100 建行实证：GLOBAL HANTEX 499,952.96 = 账 388,576.22+111,376.74，
    #   GAIN LUCKY 1,214,186.56 = 账 755,909.85+458,276.71——1 笔网银收=2 笔账拆账）。
    #   防爆炸：候选 ≤40 项且每笔金额 ≤ target（更大金额不可能参与和）。
    rem_j = [ji for ji in range(len(uj)) if ji not in used_j]
    # 建行等大量 uj → 按金额索引分组快速命中
    if rem_j:
        rem_j_by_dir_month = defaultdict(list)
        for ji in rem_j:
            j = uj[ji]
            _d = j.get('direction', '')
            _m = str(j.get('date') or '')[:7]
            rem_j_by_dir_month[(_d, _m)].append(ji)
        for bi in range(len(ub)):
            if bi in pair:
                continue
            target = round(float(ub[bi].get('amount') or 0), 2)
            if target <= 0:
                continue
            b = ub[bi]
            _bd = b.get('direction', '')
            _bm = str(b.get('date') or '')[:7]
            # 同方向同月候选（优先）；无则放宽到同月
            cands = rem_j_by_dir_month.get((_bd, _bm), [])
            if not cands:
                cands = [ji for ji in rem_j if str(uj[ji].get('date') or '')[:7] == _bm]
            # 金额过滤：≤target，且留 60 笔上限
            cands = [ji for ji in cands if float(uj[ji].get('amount') or 0) <= target][:60]
            if len(cands) < 2:
                continue
            _hit = _subset_sum(cands, uj, target, tol)
            if _hit:
                pair[bi] = _hit
                used_j.update(_hit)
    return pair


def _subset_sum(indices, rows, target, tol=0.01):
    """从 indices 中找子集，其金额和 == target（≤5 个元素，排序+剪枝）。
    ⚡⚡ 2026-08-14 增强：金额降序 + 前缀和剪枝 + 超目标提前终止，支持 60 项候选。
    （3100 建行实证：GLOBAL HANTEX = 账 388,576.22+111,376.74 两笔找和）"""
    n = len(indices)
    if n == 0:
        return None
    # 金额降序：大金额优先，快速收敛
    items = sorted(indices, key=lambda i: -abs(float(rows[i].get('amount') or 0)))
    vals = [float(rows[i].get('amount') or 0) for i in items]
    # 前缀和（升序累加用于剪枝：最小 k 项和 > target → 更大 k 无解）
    sorted_asc = sorted(vals)
    pref = [0.0]
    for v in sorted_asc:
        pref.append(pref[-1] + v)
    for size in (2, 3, 4, 5):
        if size > n:
            break
        # 最小 size 项和已超 target → 更大 size 更超
        if pref[size] > target + tol:
            continue
        # DFS：只取 ≤ target 的项
        stack = [(0, [], 0.0)]
        while stack:
            start, chosen, total = stack.pop()
            if len(chosen) == size:
                if abs(total - target) <= tol:
                    return [items[i] for i in chosen]
                continue
            if total > target + tol:
                continue
            for k in range(start, n - (size - len(chosen)) + 1):
                v = vals[k]
                if v > target + tol:
                    continue  # 单笔已超，跳过（后续更小）
                stack.append((k + 1, chosen + [k], total + v))
    return None


def _build_sheets(wb, results, banks, journal, year, monthly=None, ar_check=None, diag=None):
    diag = diag or {}
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HFILL = PatternFill(fill_type='solid', fgColor='DDEBF7')
    RED = PatternFill(fill_type='solid', fgColor='FCE4E4')
    GREEN = PatternFill(fill_type='solid', fgColor='E2EFDA')
    TOT = PatternFill(fill_type='solid', fgColor='FCE4D6')
    YELLOW = PatternFill(fill_type='solid', fgColor='FFF2CC')
    thin = Side(style='thin', color='BFBFBF')
    BD = Border(left=thin, right=thin, top=thin, bottom=thin)
    C = Alignment(horizontal='center', vertical='center', wrap_text=True)
    L = Alignment(horizontal='left', vertical='center', wrap_text=True)
    R = Alignment(horizontal='right', vertical='center')
    NF = '#,##0.00'
    F = Font(name='Times New Roman', size=10)
    B = Font(name='Times New Roman', size=10, bold=True)

    def hdr(ws, row, cols):
        for j, h in enumerate(cols, 1):
            c = ws.cell(row, j, h)
            c.fill = HFILL; c.font = B; c.border = BD; c.alignment = C

    # ── 汇总 sheet ──
    ws = wb.create_sheet('核对汇总')
    ws.cell(1, 1, '银行网银双向核对汇总（%s）' % year).font = Font(name='Times New Roman', size=14, bold=True)
    hdr(ws, 3, ['银行', '网银笔数', '日记账笔数', '匹配笔数', '网银有账无', '账有网银无',
                '已解释(归因/核销)', '核对结论', '月度先行诊断'])
    r = 4
    tot = [0, 0, 0, 0, 0, 0]
    for b in sorted(results):
        s = results[b]['summary']
        ok = '✅ 一致' if (s['unmatched_b'] == 0 and s['unmatched_j'] == 0) else '❌ 差异'
        if b == '特殊科目':
            # ⚡ 2026-08-13 过渡镜像（币种转换/现流转换=真实账户购汇/结汇镜像）——剔除防发生额
            #   双计虚增，非差异；购汇业务由真实账户行与网银购汇流水核对
            ok = '✅ 过渡镜像(剔除)'
        # ⚡ 2026-08-14 月度先行诊断列（用户方法论固化）：净额平衡（收差≈支差）→ 差异多为
        #   对称双计/未配平；对平但红区非空 → 对称未配对；对称项数 → 未配平证据
        dd = diag.get(b, {})
        diag_txt = ''
        if dd.get('net_balanced'):
            diag_txt += '收差≈支差(净额平衡)→差异多为对称双计/未配平，应找配平路径'
        elif dd.get('balanced_but_red'):
            diag_txt += '两边对平但红区有未配→对称未配对(N:M汇总/跨期)，应继续配对'
        if dd.get('sym_pairs'):
            diag_txt += (('；' if diag_txt else '') + f'红区对称项{dd["sym_pairs"]}对(未配平证据)')
        vals = [b, s['bank_total'], s['journal_total'], s['matched'],
                s['unmatched_b'], s['unmatched_j'], s.get('explained', 0), ok, diag_txt]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.border = BD; c.font = F
            c.alignment = C if j in (1, 7, 8) else R
            if j in (2, 3, 4, 5, 6):
                c.number_format = NF
        if '❌' in ok:
            for j in range(1, 9):
                ws.cell(r, j).fill = RED
            if dd.get('net_balanced') or dd.get('balanced_but_red'):
                ws.cell(r, 9).fill = YELLOW  # 净额平衡/对称未配 → 黄标（可能非真实差异）
        tot = [tot[i] + (vals[i + 1] if isinstance(vals[i + 1], int) else 0) for i in range(6)]
        r += 1
    for j, v in enumerate(['合计'] + tot + ['—', ''], 1):
        c = ws.cell(r, j, v)
        c.font = B; c.fill = TOT; c.border = BD
        c.alignment = C if j in (1, 7, 8) else R
        if 2 <= j <= 6:
            c.number_format = NF
    for j, w in enumerate([10, 12, 12, 12, 12, 12, 15, 12, 34], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = 'A4'
    # ⚡ 2026-08-13 期末重要性水平说明（用户方法论：期末平的但存在银行有财务无→需重要性衡量）
    rr = r + 2
    # ⚡ 2026-08-13 期末未达账项（用户方法论修正：双向核对**无重要性标准**，有真实差异均列示，
    #   不设阈值过滤——重要性属审计师判断，程序只负责把差异完整列出来）
    ws.cell(rr, 1, '期末未达账项（网银未配净额−账未配净额；资金池/代扣通道已单列除外——双向核对不设重要性标准，真实差异均列示，重要性由审计师判断）：').font = Font(name='Times New Roman', size=10, bold=True)
    rr += 1
    for b in sorted(results):
        m = results[b].get('materiality')
        if not m:
            continue
        cell = ws.cell(rr, 1, f'{b}: 期末未达净额 {m["unrec"]:,.0f} 元 / 年累计流水 {m["balance"]:,.0f} 元'
                             f'（{m["ratio"]:.2%}）——真实差异，逐项列示待人工核')
        cell.font = Font(name='Times New Roman', size=9)
        rr += 1
    ws.cell(rr, 1, '说明：期末未达净额=网银未配净额−账未配净额（余额调节表的调节项）。不为 0 即存在未达账项/未匹配差异，均已在前述各 sheet 逐笔列示；是否重大由审计师结合重要性水平判断，程序不做过滤。').font = Font(name='Times New Roman', size=9, italic=True)

    # ── 银行日记账（从序时账抽取，用户核心需求①；借贷分两列直观） ──
    ws4 = wb.create_sheet('银行日记账')
    ws4.cell(1, 1, '银行日记账（从序时账抽取，按银行账户分组；收入(借)=存入、支出(贷)=付出；'
                   '对方科目=凭证级推断会计科目，对方单位=摘要/辅助核算提取）').font = Font(name='Times New Roman', size=14, bold=True)
    hdr(ws4, 3, ['银行', '日期', '收入金额(借)', '支出金额(贷)', '摘要', '对方科目', '对方单位', '凭证号', '银行账户'])
    r = 4
    for b in sorted(journal):
        for x in journal[b]:
            _dx, _ax = _disp(x)
            is_dr = _dx == '借'
            vals = [b, x['date'],
                    _ax if is_dr else '',
                    _ax if not is_dr else '',
                    x.get('summary', '')[:40], x.get('cp', '')[:40],
                    x.get('opp_name', '')[:30],
                    x.get('vchar', ''), x.get('account', '')]
            _wr(ws4, r, vals, None, NF)
            r += 1
    for j, w in enumerate([10, 12, 14, 14, 36, 36, 28, 12, 30], 1):
        ws4.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws4.freeze_panes = 'A4'

    # ── 差异清单（净化版 2026-08-13：只列真实未匹配——网银有账无/账有网银无两类；
    #   含通道类（财务按业务分笔记账，月汇总差额即真实差异）；银行级金额合计便于与月度对账对照） ──
    ws2 = wb.create_sheet('差异清单')
    ws2.cell(1, 1, '双向核对差异清单（只列真实未匹配：网银有账无 / 账有网银无）').font = Font(name='Times New Roman', size=14, bold=True)
    ws2.cell(2, 1, '术语：网银有账无=银行流水存在但账面未找到（漏记/未达账项/记错账户）；账有网银无=账面存在但流水未找到（虚记/网银未提供）。收入金额=网银收/账借，支出金额=网银支/账贷。勾稽不上原因列含月度对账印证：单侧=月度对账余额两边不一致；借贷成对=余额可能平但发生额有差异，需进一步判断。底部金额合计=各银行真差异合计（网银有账无收入/支出 + 账有网银无借方/贷方），与月度对账差异一致。').font = Font(name='Times New Roman', size=9, italic=True)
    hdr(ws2, 3, ['银行', '来源', '日期', '收入金额', '支出金额', '对方/摘要', '凭证号', '勾稽不上原因'])
    r = 4
    for b in sorted(results):
        res = results[b]
        # ⚡⚡ 2026-08-14 跨侧金额对称检测：ub↔uj 同额（含拆分求和）→ 应配未配
        #   （3100 建行实证：网银『无锡晨阳 55,776』= 账『客户收款-无锡晨阳 55,776』
        #   跨期+名称差异双落差异清单 → 降级为待核平，非真差异）
        _sym = _sym_pair_ub_uj(res)
        # ① 网银有账无（红）——含通道类（待报解/资金池等银行自动通道：财务按业务分笔记账，
        #   已月汇总核销的部分不在差异，剩余月差额=真实差异）
        for ui, u in enumerate(res['ub']):
            _d, _a = _disp(u)
            inc, exp = (_a, '') if _d == '收' else ('', _a)
            ref = _unmatch_cross_ref(u, res['ub'])
            if ui in _sym:
                reason = ('【应配未配·金额对称】账侧存在同额对应（跨期/名称前缀差异/汇总拆分'
                          + '导致未自动配对）——两侧金额一致，非真差异，待人工核平'
                          + (f'（账侧 {len(_sym[ui])} 笔合计 {sum(float(res["uj"][jj]["amount"]) for jj in _sym[ui]):,.2f}）' if _sym[ui] else ''))
                _wr(ws2, r, [b, '网银有账无', u['date'], inc, exp,
                             (u.get('counterparty') or '') + ' ' + (u.get('summary') or '')[:20], '',
                             reason], YELLOW, NF); r += 1
            elif _is_auto_channel(u):
                reason = ('通道类（待报解-TIPS扣税/资金池/公积金/批量代付等银行自动通道）：'
                          + '财务按业务分笔记账（税费/社保/薪酬），网银通道行与账行 N:M 跨月对应，'
                          + '月汇总核对后的剩余差额=真实差异（网银扣款账未记足/滞后），需追查')
                vals = [b, '网银有账无', u['date'], inc, exp,
                        (u.get('counterparty') or '') + ' ' + (u.get('summary') or '')[:20], '',
                        reason]
                _wr(ws2, r, vals, RED if not _is_auto_channel(u) else YELLOW, NF); r += 1
            elif ref == 'pair':
                reason = f'{_bank_unmatch_reason(u)}；【借贷成对】网银同月收+支同额成对——资金实际发生账未记（漏记发生额，重大），月度对账余额可能平但发生额有差异'
                vals = [b, '网银有账无', u['date'], inc, exp,
                        (u.get('counterparty') or '') + ' ' + (u.get('summary') or '')[:20], '',
                        reason]
                _wr(ws2, r, vals, RED, NF); r += 1
            else:
                reason = f'{_bank_unmatch_reason(u)}；【单侧】月度对账余额两边不一致（余额调节表调节项）'
                vals = [b, '网银有账无', u['date'], inc, exp,
                        (u.get('counterparty') or '') + ' ' + (u.get('summary') or '')[:20], '',
                        reason]
                _wr(ws2, r, vals, RED, NF); r += 1
        # ② 账有网银无（红）——特殊科目（过渡镜像）已单列解释，不在此表
        if b == '特殊科目':
            continue
        for uj_i, u in enumerate(res['uj']):
            _d, _a = _disp(u)
            inc, exp = (_a, '') if _d == '借' else ('', _a)
            ref = _unmatch_cross_ref(u, res['uj'])
            # ⚡⚡ 对称配对：该 uj 被某 ub 匹配 → 标黄色待核（不重复标红）
            _is_sym = any(uj_i in v for v in _sym.values())
            if _is_sym:
                reason = ('【应配未配·金额对称】网银侧存在同额对应（跨期/名称前缀差异/汇总拆分'
                          + '导致未自动配对）——两侧金额一致，非真差异，待人工核平')
                _wr(ws2, r, [b, '账有网银无', u['date'], inc, exp,
                             u.get('summary', '')[:30], u.get('vchar', ''), reason],
                    YELLOW, NF); r += 1
            elif ref == 'pair':
                reason = (f'{_journal_unmatch_reason(u)}；【借贷成对】账侧同月借+贷同额成对——'
                          + '更正/冲销或虚增发生额（净额0但发生额双计），需进一步判断；月度对账余额可能平、发生额有差异')
                _wr(ws2, r, [b, '账有网银无', u['date'], inc, exp,
                             u.get('summary', '')[:30], u.get('vchar', ''), reason],
                    RED, NF); r += 1
            else:
                reason = f'{_journal_unmatch_reason(u)}；【单侧】月度对账余额两边不一致（余额调节表调节项）'
                _wr(ws2, r, [b, '账有网银无', u['date'], inc, exp,
                             u.get('summary', '')[:30], u.get('vchar', ''), reason],
                    RED, NF); r += 1
        # ③ 银行级真差异金额合计（网银有账无收入/支出 + 账有网银无借方/贷方）——
        #   与月度对账差异口径一致（未配净额），便于相互印证
        t_bi = sum(x['amount'] for x in res['ub'] if x['direction'] == '收')
        t_be = sum(x['amount'] for x in res['ub'] if x['direction'] == '支')
        t_jd = sum(x['amount'] for x in res['uj'] if x['direction'] == '借')
        t_jc = sum(x['amount'] for x in res['uj'] if x['direction'] == '贷')
        if t_bi or t_be or t_jd or t_jc:
            _wr(ws2, r, [b, '─ 真差异金额合计 ─', '', t_bi or '', t_be or '',
                         '', '', f'网银有账无收 {t_bi:,.0f} / 支 {t_be:,.0f}；账有网银无借 {t_jd:,.0f} / 贷 {t_jc:,.0f}'],
                None, NF); r += 1
    # 金额列（4/5）格式：_wr 硬编码 4/8，本表金额列在 4/5，统一修正
    for rr in range(4, ws2.max_row + 1):
        for cc in (4, 5):
            cell = ws2.cell(rr, cc)
            if isinstance(cell.value, (int, float)):
                cell.number_format = NF
                cell.alignment = R
    if ws2.max_row < 4:
        ws2.cell(4, 1, '✅ 全部匹配，无真实差异').font = F
    for j, w in enumerate([10, 12, 12, 14, 14, 40, 12, 58], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws2.freeze_panes = 'A4'

    # ── 网银明细 / 日记账明细（拼接对照；借贷分两列） ──
    ws3 = wb.create_sheet('网银与日记账对照')
    ws3.cell(1, 1, '网银记录与银行日记账拼接对照（匹配行同行并列；核对收付单位 ↔ 对方科目/对方单位是否一致）').font = Font(name='Times New Roman', size=14, bold=True)
    hdr(ws3, 3, ['银行', '网银日期', '网银收入', '网银支出', '网银对方/摘要',
                 '日记账日期', '日记账借方', '日记账贷方', '日记账摘要', '对方科目', '对方单位', '凭证号', '匹配状态'])
    r = 4
    for b in sorted(results):
        res = results[b]
        for m in res['matched']:
            bb = m[0] if isinstance(m[0], list) else [m[0]]
            jr_list = m[1] if isinstance(m[1], list) else [m[1]]
            jr = jr_list[0] if jr_list else {}
            tag = m[2]
            n_j = len(jr_list)
            for bi, x in enumerate(bb):
                _dx, _ax = _disp(x)
                binc, bexp = (_ax, '') if _dx == '收' else ('', _ax)
                _dj, _aj = _disp(jr)
                jdr, jcr = (_aj, '') if _dj == '借' else ('', _aj)
                vals = [b, x['date'], binc, bexp,
                        (x.get('counterparty') or '') + ' ' + (x.get('summary') or '')[:14],
                        jr.get('date', '') if bi == 0 else '',
                        jdr if bi == 0 else '', jcr if bi == 0 else '',
                        jr.get('summary', '')[:20] if bi == 0 else '',
                        jr.get('cp', '')[:28] if bi == 0 else '',
                        jr.get('opp_name', '')[:24] if bi == 0 else '',
                        jr.get('vchar', '') if bi == 0 else '',
                        (tag + ('（账侧%d笔汇总）' % n_j)) if bi == 0 and n_j > 1 else tag]
                _wr(ws3, r, vals,
                    GREEN if '精确' in tag else (YELLOW if '未达' in tag else None), NF)
                r += 1
        for u in results[b]['ub']:
            _d, _a = _disp(u)
            binc, bexp = (_a, '') if _d == '收' else ('', _a)
            vals = [b, u['date'], binc, bexp,
                    (u.get('counterparty') or '') + ' ' + (u.get('summary') or '')[:14],
                    '', '', '', '', '', '', '', '❌ 网银有账无']
            _wr(ws3, r, vals, RED, NF); r += 1
        for u in results[b]['uj']:
            _d, _a = _disp(u)
            jdr, jcr = (_a, '') if _d == '借' else ('', _a)
            vals = [b, '', '', '', '', u['date'], jdr, jcr,
                    u.get('summary', '')[:20], u.get('cp', '')[:28], u.get('vchar', ''), '❌ 账有网银无']
            _wr(ws3, r, vals, RED, NF); r += 1
    # 金额列（3/4 网银、7/8 日记账）格式：_wr 硬编码 4/8，本表金额列在 3/4/7/8，统一修正
    for rr in range(4, ws3.max_row + 1):
        for cc in (3, 4, 7, 8):
            cell = ws3.cell(rr, cc)
            if isinstance(cell.value, (int, float)):
                cell.number_format = NF
                cell.alignment = R
    for j, w in enumerate([10, 12, 14, 14, 36, 12, 14, 14, 30, 28, 12, 18], 1):
        ws3.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws3.freeze_panes = 'A4'

    # ── 月度对账（用户方法论第 1 层：全年/分月合计先对上 → 找平可行域） ──
    if monthly:
        ws5 = wb.create_sheet('月度对账')
        ws5.cell(1, 1, '全年/分月合计对账（%s）——先对总额：对得上→明细应可找平；对不上→存在漏记/多记/记错科目' % year).font = Font(name='Times New Roman', size=14, bold=True)
        ws5.cell(2, 1, '（网银「收/支」= 企业「借/贷」视角；账借方/贷方为【剔除冲销凭证后】净额——银行日记账 Q 列冲销标识 1=冲销凭证/2=被冲销凭证（红字成对=虚增发生额，已在取数层剔除），特殊科目过渡镜像不参与。另已剔除【已解释项】（对销/小额核销/跨账户归因——净额 0 非差异，如 3100 中行 5 月内转 3.9 亿一般户↔数币户对销）——本表差异列与「差异清单」红区口径一致。差异≠0 即真实未达/未配，标红）。').font = Font(name='Times New Roman', size=9, italic=True)
        hdr(ws5, 4, ['银行', '期间', '网银收入', '网银支出', '账借方(净)', '账贷方(净)',
                     '收差异(网银-账)', '支差异(网银-账)'])
        r = 5
        for b in sorted(monthly):
            for m in sorted(monthly[b]):
                d = monthly[b][m]
                is_tot = (m == 'TOTAL')
                diff_bad = d['diff_inc'] != 0 or d['diff_exp'] != 0
                vals = [b if not is_tot else (b + ' 全年'), m if not is_tot else '合计',
                        d['b_inc'], d['b_exp'], d['j_dr'], d['j_cr'],
                        d['diff_inc'], d['diff_exp']]
                for j, v in enumerate(vals, 1):
                    c = ws5.cell(r, j, v)
                    c.border = BD
                    c.font = B if is_tot else F
                    c.alignment = C if j in (1, 2) else R
                    if j >= 3:
                        c.number_format = NF
                    # ⚡ 勿赋 None：openpyxl cell.fill 不接受 None（保存崩溃）
                    if is_tot:
                        c.fill = TOT
                    elif diff_bad:
                        c.fill = RED
                r += 1
        for j, w in enumerate([10, 10, 14, 14, 14, 14, 16, 16], 1):
            ws5.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
        ws5.freeze_panes = 'A5'

    # ── 大额交易核对（用户方法论 2026-08-14：『大金额的核对是要非常仔细的』——
    #     ①摘要名称是否与对方名称一致 ②入账科目是否合乎逻辑（大额应过往来）
    #     ③个人大额进出是否出现（分红/借款/工资除外）④入账是否准确（未配=重点查） ──
    _big = _big_amt_rows(results, banks, journal)
    ws6 = wb.create_sheet('大额交易核对')
    ws6.cell(1, 1, '大金额专项核对（≥%d元；个人大额≥%d元）——大金额核对非常仔细：'
                   '①摘要名称↔对方名称一致 ②入账科目逻辑（大额应过往来）'
                   '③个人大额进出 ④入账准确性（未配=重点查）'
             % (_MAJOR_AMT, _PERSONAL_AMT)).font = Font(name='Times New Roman', size=14, bold=True)
    hdr(ws6, 3, ['银行', '日期', '方向', '金额', '网银对方/摘要', '账侧摘要', '账侧对方科目',
                 '匹配状态', '名称一致性', '科目逻辑', '专项提示'])
    r = 4
    for row in _big:
        _t = str(row[7]) + str(row[8]) + str(row[10])
        _fill = RED if ('❌' in _t or '🔴' in _t) else (YELLOW if '⚠️' in str(row[8]) else GREEN)
        _wr(ws6, r, row, _fill, NF)
        r += 1
    for j, w in enumerate([8, 11, 6, 15, 38, 28, 26, 14, 12, 12, 12], 1):
        ws6.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws6.freeze_panes = 'A4'
    n_big_red = sum(1 for row in _big if '❌' in str(row[7]) or '🔴' in str(row[10]))
    print(f'[大额核对] {len(_big)} 笔大额（≥{_MAJOR_AMT:,}元）已列示，其中 {n_big_red} 笔需重点核查（未配/个人大额）')

    # ⚡ 2026-08-13 用户指示：「对方科目核对」与「应收回款核对」暂时不做，先把双向核对
    #   问题解决——差异清单已净化（只列 网银有账无/账有网银无 两类 + 勾稽原因）。


def _wr(ws, r, vals, fill, nf):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    thin = Side(style='thin', color='BFBFBF')
    BD = Border(left=thin, right=thin, top=thin, bottom=thin)
    C = Alignment(horizontal='center', vertical='center', wrap_text=True)
    L = Alignment(horizontal='left', vertical='center', wrap_text=True)
    R = Alignment(horizontal='right', vertical='center')
    F = Font(name='Times New Roman', size=10)
    for j, v in enumerate(vals, 1):
        c = ws.cell(r, j, v)
        c.font = F; c.border = BD
        c.alignment = C if j in (1, 3, 6, 7, 10, 11) else R if j in (4, 8) else L
        if j in (4, 8):
            c.number_format = nf
        if fill:
            c.fill = fill


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print('用法：python bank_reconcile_detail.py <账套目录> [<网银文件夹>] [--journal-file <银行日记账.xlsx>]')
        return 1
    jf = None
    if '--journal-file' in argv:
        i = argv.index('--journal-file')
        jf = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    data_dir = argv[0]
    stmt = argv[1] if len(argv) > 1 else None
    out = build_bank_reconcile(data_dir, stmt, journal_file=jf)
    return 0 if out else 1


if __name__ == '__main__':
    sys.exit(main())
