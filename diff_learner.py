# -*- coding: utf-8 -*-
"""差异模式学习器（2026-08-14 用户方法论：程序是死的、知识是活的——能否不断通过
知识扩展+自我学习，做到举一反三，不用靠人一步步催促？）。

回答：能。关键设计=【程序主动提出候选，人只做判断题（确认/驳回），确认的自动写
入知识库】。把"人出填空题（XX 也是）"变成"程序出判断题（这个候选对吗）"。

闭环：
    新账套数据 → 执行器（核对+体检）→ 未定性差异
        → 学习器（共性挖掘）→ 疑似新模式候选
        → 人确认/驳回（判断题）→ 知识库增长
        → 下次全账套自动命中（举一反三）→ 回到执行器

用户补充（2026-08-14）："举一反三并不要求你一次就成功，可以有试错过程"。
→ 学习器允许试错：候选宁多勿漏（判题成本低），人被驳回的候选记入记忆文件
  （REJECTED），下次自动不再提；确认的候选记入记忆（ACCEPTED）并提示写入知识库
  关键词表。试错本身是学习的一环——被驳回的负样本同样有价值。

学习器三条挖掘通道（不依赖人指出，程序主动提）：
    A. 高频词共现：未配行摘要/对方户名高频词（≥2 笔）→ 疑似新双计关键词
       [如 3400 农商"后勤押金"贷 4 笔 -1,125 —— 学习器应自动挖出"押金"→ 加 fee_kw]
    B. 结构同构：收差≈支差 / 同额收+支 / 相邻月符号翻转 → 复用已有 P1/P8/P9 处理
    C. 跨账套同构：多家账套同一银行同构 → 记账习惯共性（自动推广）

试错记忆（self-learning feedback，JSON 落盘到小程序目录）：
    diff_learner_memory.json  = {"accepted": {"词": "理由"}, "rejected": {"词": "理由"}}
    accept(word, reason)       → 确认候选（写入 accepted，之后不再作为候选提出）
    reject(word, reason)       → 驳回候选（写入 rejected，之后不再提出）
    suggest() 自动过滤 accepted/rejected 中已判过的词。

用法：
    mine_keywords(rows)              → 通道A：高频词共现挖掘
    suggest(rows, bank, monthly)     → 汇总三条通道的学习建议（自动过滤已判词）
    accept(word, reason) / reject(word, reason) → 试错记忆（人确认/驳回）
"""
from collections import Counter, defaultdict
import json
import os
import re

# 试错记忆文件（JSON，落盘到本文件同目录）
_MEMORY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            'diff_learner_memory.json')


def _load_memory():
    """读试错记忆：{"accepted": {词: 理由}, "rejected": {词: 理由}}。"""
    try:
        with open(_MEMORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {'accepted': {}, 'rejected': {}}


def _save_memory(mem):
    try:
        with open(_MEMORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(mem, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def accept(word, reason='', kw_type='fee'):
    """人确认候选词：写入 accepted（之后不再提出），并自动注入下次核对的处理
    关键词（闭环最后一环：确认→下次自动生效，无需手动改 fee_kw/recv_kw）。

    kw_type: 'fee'（费用/冲减类，注入 fee_kw，如 押金/支库）
             'recv'（收款渠道类，注入 recv_kw，如 饮料机/浴室充值）
    """
    mem = _load_memory()
    mem.setdefault('accepted', {})[word] = {
        'reason': reason or '确认有效', 'type': kw_type,
    }
    mem.setdefault('rejected', {}).pop(word, None)
    _save_memory(mem)
    print(f'  [学习器] 已确认 "{word}"（{reason}）→ 已自动注入 {kw_type}_kw，'
          f'下次核对自动生效')


def reject(word, reason=''):
    """人驳回候选词：写入 rejected（负样本，之后不再提出）。"""
    mem = _load_memory()
    mem.setdefault('rejected', {})[word] = reason or '驳回'
    mem.setdefault('accepted', {}).pop(word, None)
    _save_memory(mem)
    print(f'  [学习器] 已驳回 "{word}"（{reason}）→ 下次不再作为候选')


def learned_keywords(kw_type='fee'):
    """已确认学习词（自动注入核对处理关键词用）。返回该类型全部已确认词元组。
    供 bank_reconcile_detail 的 fee_kw/recv_kw 动态扩展——闭环：人确认 → 程序
    自动记住 → 下次核对按新关键词净额化/配对。"""
    mem = _load_memory()
    acc = mem.get('accepted', {})
    return tuple(w for w, info in acc.items()
                 if isinstance(info, dict) and info.get('type') == kw_type)


def _judged_words():
    """已判过的词（accepted + rejected 的并集），学习器自动跳过。"""
    mem = _load_memory()
    return set(mem.get('accepted', {})) | set(mem.get('rejected', {}))


def memory_summary():
    """打印试错记忆概览。"""
    mem = _load_memory()
    acc, rej = mem.get('accepted', {}), mem.get('rejected', {})
    print(f'  [学习器记忆] 已确认 {len(acc)} 词、已驳回 {len(rej)} 词（试错负样本）')
    if acc:
        print(f'    确认：{"、".join(acc.keys())}')
    if rej:
        print(f'    驳回：{"、".join(rej.keys())}')

# 已知关键词（已入知识库，不再建议）——新候选应避开它们
KNOWN_KEYWORDS = (
    '手续费', '服务费', '退押金', '二维码', '一码通', '饮料机', '浴室充值',
    '购汇', '结汇', '售汇', '预付款', '受托', '委托贷款',
    '待报解', '代理国库', '收缴', '实归', '支取', '工资归集', '税贷',
    '社保', '汇总', '理财',
)

# 停用词（高频但无学习价值：银行名/通用动词/金额单位）
STOP_WORDS = {
    '转账', '划转', '支付', '收款', '付款', '收入', '支出', '银行', '账户',
    '网银', '人民币', '公司', '有限', '集团', '交易', '退回', '汇入', '汇出',
    '入账', '出账', '电子', '业务', '扣款', '结算', '划扣', '结转', '跨期',
    '上海', '中国', '建设', '工商', '农业', '交通', '民生', '光大', '兴业',
    '招商', '中信', '浦发', '泰隆', '农商', '华夏', '浙商', '广发', '平安',
    '一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '年', '月', '日',
    # 长串切分出的无信息量碎片/通用词（注意：金库/支库/吴江 等是 P10 待报解真实
    # 关键词，保留作候选，不在此停用）
    '般户', '民币', '人民', '存款', '般存', '户存', '银行存款', '上海银行',
    '一般', '州市', '库苏', '市吴',
    # 账户名/银行名（非摘要业务词，通道S聚合模式应跳过）
    '工行', '中行', '农行', '建行', '商行', '农商行', '农村商业银行',
    '专户', '账户', '账号', '基本户', '专用户',
}


def _norm_text(row):
    """把行内可读文本拼成一段，用于词频分析。"""
    parts = []
    for k in ('summary', 'counterparty', 'cp', 'name', 'subject'):
        v = row.get(k)
        if v is not None:
            parts.append(str(v))
    return ' '.join(parts)


def _grams(text):
    """提取候选词：完整连续汉字串（2-4字）+ 串首/串尾 2 字词根。
    超长串（如'国家金库苏州市吴江区支库'）只取首尾 2 字（'国家'/'支库'），
    避免 '市吴'/'库苏' 这类跨字碎片。"""
    han = re.findall(r'[\u4e00-\u9fff]{2,}', text)
    out = set()
    for seg in han:
        if len(seg) <= 4:
            # 完整短串直接作为候选
            if seg not in STOP_WORDS:
                out.add(seg)
            # 2 字词根（前缀/后缀）
            for g in (seg[:2], seg[-2:]):
                if g not in STOP_WORDS and not any(k in g for k in KNOWN_KEYWORDS):
                    out.add(g)
        else:
            # 超长串：只取首尾 2 字词根（词根有语义，跨字碎片没有）
            for g in (seg[:2], seg[-2:]):
                if g not in STOP_WORDS and not any(k in g for k in KNOWN_KEYWORDS):
                    out.add(g)
    return out


def _dedup_fragments(word_rows):
    """去碎片：短词 X 被更长词 Y 包含且行集相同，且 X 既不是 Y 的前缀也不是后缀
    （如 '勤押'）→ X 是跨字碎片，删 X。保留词根（'后勤'是前缀、'押金'是后缀，
    都有知识价值，保留便于给出可复用的核心关键词）。"""
    words = sorted(word_rows, key=len, reverse=True)
    keep = set(words)
    for short in words:
        for long in words:
            if long == short or len(long) <= len(short):
                continue
            if short in long and not long.startswith(short) and not long.endswith(short):
                short_rows = set(id(r) for r in word_rows[short])
                long_rows = set(id(r) for r in word_rows[long])
                if short_rows <= long_rows:
                    keep.discard(short)
                    break
    return keep


def mine_keywords(rows, min_rows=2, max_cand=8):
    """通道A：高频词共现——统计未配行里出现 ≥min_rows 次的词，按笔数+金额降序。

    返回：[{word, rows, amt, samples}]（rows=命中的行数，amt=绝对金额合计）。
    """
    word_rows = defaultdict(list)  # word -> [row, ...]
    judged = _judged_words()      # 已确认/已驳回的词（试错记忆）自动跳过
    for r in rows:
        for g in _grams(_norm_text(r)):
            if g in judged:
                continue
            word_rows[g].append(r)
    keep = _dedup_fragments(word_rows)
    cands = []
    for w in keep:
        rs = word_rows[w]
        if len(rs) < min_rows:
            continue
        amt = sum(abs(float(r.get('amount') or 0)) for r in rs)
        # 金额太小且笔数少 → 噪声（如 45 元镜像）
        if len(rs) == min_rows and amt < 500:
            continue
        samples = [str(r.get('summary') or r.get('counterparty') or '')[:20] for r in rs[:3]]
        cands.append({'word': w, 'rows': len(rs), 'amt': amt, 'samples': samples})
    # 排序：笔数优先，其次金额；同笔数时 2 字词根（更可复用）排前
    cands.sort(key=lambda c: (-c['rows'], -c['amt'], len(c['word'])))
    return cands[:max_cand]


def _structure_symmetry(rows):
    """通道B：结构同构——收差≈支差（对称双计特征）。"""
    inc = sum(float(r.get('amount') or 0) for r in rows if r.get('direction') == '收')
    exp = sum(float(r.get('amount') or 0) for r in rows if r.get('direction') == '支')
    if abs(inc - exp) < max(100.0, abs(inc) * 0.01) and (abs(inc) > 0.01 or abs(exp) > 0.01):
        return f'收差{inc:,.2f}≈支差{exp:,.2f}（对称双计 → P1 净额化路径）'
    return None


# ════════════════════════════════════════════════════════════════════
# 通道S · 结构信号 → 摘要雷同 → 聚合模式（用户 2026-08-14 03:02 方法论：
#   "当你发现有大量频繁的交易，无法一一匹配，但合计数差异又不大，这个时候
#    自己要考虑他的摘要内容是否雷同与押金，应该用押金模式进行核对"）
#
# 与通道A（关键词共现）的本质区别：
#   通道A = 文字驱动：先统计高频词 → 建议加关键词（还在"匹配文字"）
#   通道S = 结构驱动：先看结构信号（笔数多 + 收差≈支差）→ 自动怀疑"摘要是否
#           雷同" → 建议按摘要聚合配对（押金模式：聚合后净额配对，而非逐笔）
#
# 押金模式（聚合配对）= P11 汇总记账/对称双计的通用解法：网银/账侧同一摘要
#   多笔小额频繁发生（押金/手续费/结息/社保/待报解），逐笔无法匹配，但同摘要
#   聚合后与对方侧能配对 → 按摘要聚合后净额化，而非逐笔死磕。
# ════════════════════════════════════════════════════════════════════


def _dir_norm(d):
    """方向归一化：网银收/支 与 账侧借/贷 统一到收/支（收=借=资金流入）。"""
    return '收' if d in ('收', '借') else '支'


def detect_pool_patterns(rows, min_rows=3, max_net_ratio=0.15):
    """结构信号 → 摘要雷同 → 建议押金模式（聚合配对）。

    触发条件：
      ① 笔数多：未配行 >= min_rows（大量频繁交易，逐笔无法一一匹配）
      ② 合计差异小：收差≈支差，净额占总额比 <= max_net_ratio（合计几乎平衡）
         ——这是"聚合型业务"的指纹：总额大、净额小、单笔对不上
      ③ 存在雷同子集：某摘要词出现 >= min_rows 笔（摘要雷同，如"押金"）
    满足 → 建议按该摘要聚合配对（押金模式核对）。

    注意：方向已归一化——ub（收/支）与 uj（借/贷）可混传，'借'按'收'、'贷'
    按'支'处理（收=借=资金流入）。返回 [dict,...] 或 []：每个雷同词一组
    {word, rows, amt, inc, exp, net, coverage, samples}。覆盖率只作展示
    （结构信号已满足时，雷同子集即线索，不要求覆盖全部未配行）。"""
    if len(rows) < min_rows:
        return []
    inc = sum(abs(float(r.get('amount') or 0)) for r in rows
              if _dir_norm(r.get('direction')) == '收')
    exp = sum(abs(float(r.get('amount') or 0)) for r in rows
              if _dir_norm(r.get('direction')) == '支')
    total = inc + exp
    if total <= 0:
        return []
    net = abs(inc - exp)
    # ② 合计差异小（聚合型指纹）：净额/总额 <= 阈值
    if net / total > max_net_ratio:
        return []
    # ③ 摘要雷同子集：找出现 >= min_rows 笔的词
    word_rows = defaultdict(list)
    for r in rows:
        for g in _grams(_norm_text(r)):
            word_rows[g].append(r)
    groups = []
    for w, rs in word_rows.items():
        if len(rs) < min_rows:
            continue
        r_inc = sum(abs(float(r.get('amount') or 0)) for r in rs
                    if _dir_norm(r.get('direction')) == '收')
        r_exp = sum(abs(float(r.get('amount') or 0)) for r in rs
                    if _dir_norm(r.get('direction')) == '支')
        samples = [str(r.get('summary') or r.get('counterparty') or '')[:20]
                   for r in rs[:3]]
        groups.append({
            'word': w, 'rows': len(rs), 'amt': r_inc + r_exp,
            'inc': r_inc, 'exp': r_exp, 'net': abs(r_inc - r_exp),
            'coverage': len(rs) / len(rows), 'samples': samples,
        })
    # 按（笔数, 金额）降序：覆盖越多/金额越大的雷同词越可能是主线索
    groups.sort(key=lambda g: (-g['rows'], -g['amt']))
    return groups


def suggest(rows, bank='', monthly=None):
    """汇总学习建议。返回 [{'type', 'bank', 'desc', 'suggest'}]。
    优先级：通道S（结构信号→聚合模式）> 通道A（关键词共现）> 通道B（结构同构）
            > 通道C（跨账套同构）。自动跳过试错记忆里已判过的词。
    用户方法论（03:02）：不能只停留在"匹配文字"——大量频繁交易无法一一匹配但
    合计数差异不大时，应自动怀疑摘要雷同，用押金模式（聚合配对）核对。"""
    out = []
    # 通道S：结构信号 → 摘要雷同 → 押金模式（最接近审计思维，最高优先级）
    pools = detect_pool_patterns(rows)
    if pools:
        _signed = sum(float(r.get('amount') or 0) *
                      (1 if _dir_norm(r.get('direction')) == '收' else -1) for r in rows)
        _tot = sum(abs(float(r.get('amount') or 0)) for r in rows)
        for pool in pools[:3]:  # 最多报前 3 个雷同组
            out.append({
                'type': 'S-聚合模式',
                'bank': bank,
                'desc': (f'未配行 {len(rows)} 笔、总额 {_tot:,.2f}，'
                         f'净额 {abs(_signed):,.2f}'
                         f'（合计差异小，净额/总额={abs(_signed)/max(1,_tot)*100:.1f}%）'
                         f'→ 摘要雷同"{pool["word"]}" {pool["rows"]} 笔 '
                         f'{pool["amt"]:,.2f}，覆盖率 {pool["coverage"]*100:.0f}%'
                         f'（例：{"/".join(pool["samples"])}）'),
                'suggest': ('按摘要聚合配对（押金模式：同摘要聚合成一笔净额与对方侧配对），'
                            '而非逐笔死磕——检查是否应加聚合关键词/汇总记账规则'),
            })
    # 通道A：高频词共现
    judged = _judged_words()
    kws = mine_keywords(rows)
    for c in kws:
        tag = ''
        if c['word'] in judged:
            tag = '（已判过，试错记忆跳过）'
        out.append({
            'type': 'A-关键词共现',
            'bank': bank,
            'desc': f'未配行出现"{c["word"]}" {c["rows"]} 笔、合计 {c["amt"]:,.2f}'
                    f'（例：{"/".join(c["samples"])}）{tag}',
            'suggest': f'疑似新双计关键词，确认后加入知识库关键词表（如 fee_kw）',
        })
    # 通道B：结构同构
    sym = _structure_symmetry(rows)
    if sym:
        out.append({
            'type': 'B-结构同构',
            'bank': bank,
            'desc': sym,
            'suggest': '净额平衡 → 差异多为对称双计/未配平，应找配平路径（复用 _fee_net_journal 等）',
        })
    # 通道C：跨账套同构（由调用方传入多家账套的体检结果做并集，这里占位提示）
    if monthly is not None:
        n_bal = sum(1 for d in monthly.values()
                    if abs(d.get('TOTAL', {}).get('diff_inc', 0) -
                           d.get('TOTAL', {}).get('diff_exp', 0)) < 100)
        if n_bal >= 2:
            out.append({
                'type': 'C-跨账套同构',
                'bank': bank,
                'desc': f'{n_bal} 家银行同时净额平衡',
                'suggest': '跨账户共性 → 记账习惯级规则（如购汇/受托/资金池），一次修复多账套受益',
            })
    return out


def print_suggestions(suggestions, title='差异模式学习建议'):
    """打印学习器建议（引擎/核对程序调用）。"""
    print()
    print(f'===== {title} =====')
    if not suggestions:
        print('  ✅ 未发现待学习的新模式（未定性差异无共性）')
        return
    for s in suggestions:
        print(f'  [{s["type"]}] {s["bank"]}: {s["desc"]}')
        print(f'      → {s["suggest"]}')


if __name__ == '__main__':
    # 自检1：结构信号场景（用户方法论核心）——大量频繁押金交易、无法逐笔匹配、
    #   但合计差异小 → 应自动怀疑摘要雷同 → 建议押金模式（聚合配对）
    pool_rows = [
        {'direction': '支', 'amount': 530.0,  'summary': '后勤押金', 'counterparty': '物业'},
        {'direction': '支', 'amount': 280.0,  'summary': '后勤押金', 'counterparty': '食堂'},
        {'direction': '支', 'amount': 225.0,  'summary': '押金-后勤', 'counterparty': '物业'},
        {'direction': '支', 'amount': 90.0,   'summary': '押金', 'counterparty': '食堂'},
        {'direction': '支', 'amount': 150.0,  'summary': '退押金', 'counterparty': '物业'},
        {'direction': '支', 'amount': 320.0,  'summary': '押金退款', 'counterparty': '食堂'},
        {'direction': '收', 'amount': 1480.0, 'summary': '押金退还', 'counterparty': '物业'},
    ]
    print('== 自检1：结构信号 → 押金模式（大量频繁+合计差异小+摘要雷同）==')
    pool = detect_pool_patterns(pool_rows)
    print(f'  detect_pool_patterns → {"命中 " + str(len(pool)) + " 组" if pool else "未命中"}')
    for p in pool[:2]:
        print(f'    摘要"{p["word"]}" {p["rows"]} 笔 {p["amt"]:,.2f}、覆盖率 {p["coverage"]*100:.0f}%')
    assert pool and any('押金' in p['word'] for p in pool), '押金模式检测失败'
    print('  ✅ 押金模式命中')
    print()
    # 自检2：对照——无结构信号（合计差异大）→ 不应建议聚合模式
    nonpool = [
        {'direction': '支', 'amount': 5000000.0, 'summary': '押金', 'counterparty': '物业'},
        {'direction': '收', 'amount': 1000.0,    'summary': '押金', 'counterparty': '物业'},
    ]
    assert detect_pool_patterns(nonpool) == [], '合计差异大不应命中聚合模式'
    print('  对照：合计差异大 → 不误报聚合模式 ✅')
    print()
    # 自检3：完整 suggest 输出（含通道S/A/B）
    print_suggestions(suggest(pool_rows, bank='3400 农商行'))
    print()
    # 自检4：试错闭环——驳回"后勤"后不再提出
    reject('后勤', '碎片词，押金已确认足够')
    print()
    print_suggestions(suggest(pool_rows, bank='3400 农商行'))
    memory_summary()
