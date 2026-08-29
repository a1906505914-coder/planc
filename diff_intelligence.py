# -*- coding: utf-8 -*-
"""差异智能统一入口（2026-08-14 架构收口：把散落在 build_bank_reconcile 里的
三处体检/知识/学习调用收敛成单一模块，单一职责、可复用、可挂进要求层）。

组成（知识库四件套 → 一个入口）：
    ① bank_diff_patterns.diagnose   P1-P12 现象诊断（已知模式自动扫描）
    ② fictitious_knowledge         两层证据知识库（层2 SAP 记账习惯兜底扫描）
    ③ diff_learner.suggest         自我学习（通道A 文字挖掘 + 通道S 结构信号）
    ④ diff_learner.learned_keywords 学习记忆（确认词自动注入处理关键词）

设计原则：
    - run() 是唯一入口，输出统一报告（体检 + 知识 + 学习），任何一步失败
      都不影响主流程（各自 try/except 兜底）
    - 纯诊断/建议，不改 results——核对结果在调用本模块前已定稿
    - 学习词注入（learned_keywords→fee_kw/recv_kw）发生在核对【处理阶段】
      （_fee_net_journal 内），本模块只负责"报告 + 提示确认"。

用法：
    from diff_intelligence import run_intelligence
    run_intelligence(banks, journal, results, monthly, verbose=True)
"""
from collections import defaultdict


def run_intelligence(banks, journal, results, monthly, verbose=True):
    """统一执行差异智能（体检+知识+学习），返回 {bank: [命中的条目]}。

    输入：banks/journal（原始）、results（核对后 ub/uj）、monthly（月度对账）。
    输出：{bank: [(kind, name, evidence, advice)]}，kind ∈ {'P1'..'P12', '知识', '学习'}。
    任何一步失败不影响其他步骤（try/except 各自兜底）。
    """
    report = defaultdict(list)
    # ⚡⚡ 2026-08-14 大数据保护阈值：单银行未配行超此值时跳过 学习器/知识扫描
    #   （词切分/聚合内存放大，3500 实证 4 万笔 → 62.8 亿全量聚合）；P1-P12 体检仍跑。
    MAX_LEARN_ROWS = 8000

    # ── ① 现象诊断：P1-P12 已知模式 ──
    try:
        from bank_diff_patterns import diagnose as _diag
        for b, hits in _diag(banks, journal, results, monthly).items():
            for pid, name, ev, adv in hits:
                report[b].append((pid, name, ev, adv))
    except Exception:
        pass

    # ── ② 两层知识：层2 SAP 记账习惯兜底扫描 ──
    try:
        from fictitious_knowledge import match_habits
        for b in sorted(set(banks) | set(journal)):
            if b == '特殊科目':
                continue
            r = results.get(b, {})
            _rows = r.get('ub', []) + r.get('uj', [])
            if len(_rows) > MAX_LEARN_ROWS:
                continue  # 大主体保护：跳过度量知识扫描（体检仍跑）
            cands = match_habits(_rows)
            if cands:
                # 去重（习惯+日期）
                seen = set()
                for c in cands:
                    k = (c['habit'], c['row'].get('date', ''))
                    if k in seen:
                        continue
                    seen.add(k)
                    report[b].append(('知识', c['habit'], c['why'], c['fix']))
    except Exception:
        pass

    # ── ③ 自我学习：通道A 文字挖掘 + 通道S 结构信号 ──
    try:
        from diff_learner import suggest as _learn
        for b in sorted(set(banks) | set(journal)):
            if b == '特殊科目':
                continue
            r = results.get(b, {})
            _rows = r.get('ub', []) + r.get('uj', [])
            if len(_rows) > MAX_LEARN_ROWS:
                continue
            for s in _learn(_rows, bank=b):
                report[b].append((s['type'], s['desc'], '', s['suggest']))
    except Exception:
        pass

    if verbose:
        print_report(report)
    return dict(report)


def print_report(report, title='差异智能报告（体检 P1-P12 + 知识层2 + 学习候选）'):
    """统一打印差异智能报告。"""
    print()
    print(f'===== {title} =====')
    if not report:
        print('  ✅ 无已知差异模式/知识线索/学习候选（差异均已定性或对平）')
        return
    for b in sorted(report):
        print(f'【{b}】')
        for kind, name, ev, adv in report[b]:
            if kind.startswith('P'):
                print(f'  [{kind} {name}] {ev}')
                print(f'      → {adv}')
            elif kind == '知识':
                print(f'  [知识·层2 {name}] {ev}')
                print(f'      → {adv}')
            else:
                print(f'  [{kind}] {name}')
                print(f'      → {adv}')


if __name__ == '__main__':
    # 自检：用模拟数据跑通四件套
    rows = [
        {'direction': '支', 'amount': 530.0, 'summary': '后勤押金', 'counterparty': '物业'},
        {'direction': '支', 'amount': 280.0, 'summary': '后勤押金', 'counterparty': '食堂'},
        {'direction': '支', 'amount': 225.0, 'summary': '押金-后勤', 'counterparty': '物业'},
        {'direction': '支', 'amount': 90.0,  'summary': '押金', 'counterparty': '食堂'},
        {'direction': '支', 'amount': 150.0, 'summary': '退押金', 'counterparty': '物业'},
        {'direction': '支', 'amount': 320.0, 'summary': '押金退款', 'counterparty': '食堂'},
        {'direction': '收', 'amount': 1480.0, 'summary': '押金退还', 'counterparty': '物业'},
    ]
    fake = {
        '农商行': {
            'ub': rows,
            'uj': [{'direction': '借', 'amount': 100.0, 'summary': '手续费', 'date': '2026-01-05'}],
        }
    }
    monthly = {
        '农商行': {
            'TOTAL': {'diff_inc': 115.0, 'diff_exp': -100.0},
            '2026-01': {'diff_inc': 115.0, 'diff_exp': -100.0},
        }
    }
    run_intelligence(['农商行'], ['农商行'], fake, monthly)
