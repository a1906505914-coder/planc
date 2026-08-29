# -*- coding: utf-8 -*-
"""应收 + 合同资产 联动账龄复核（专项底稿，2026-08-16 用户定）。

业务逻辑（用户 00:51 需求）：
  · 应收账款（1122 等）与合同资产（XBJ 1481 / AJ 1125 / AH 1124）同时存在时，
    不能按纯应收模式分龄——合同资产中【未到期】部分（质保金/审计金/未验收款）
    尚未获得无条件收款权，不应进账龄桶。
  · 处理：合同资产按类别未到期比例（manifest ca_cats）拆分：
      ① 未到期部分 → 单独列示（不参与账龄）
      ② 已到期部分 → 与应收合并，按 GL 交易日期 FIFO 进六档账龄桶
  · 账龄桶：1年以内/1-2/2-3/3-4/4-5/5年以上（与专项账龄复核六档一致）

输出（三件套，对标其他专项）：
  ① 集团底稿：{账套}/底稿/{year}/应收合同资产账龄复核_{year}.xlsx
     Sheet1 账龄汇总(按主体+科目)  Sheet2 合同资产未到期明细
     Sheet3 账龄明细(FIFO)        Sheet4 与TB核对
  ② 单体拆分：按主体行（已在 Sheet1/3 体现）
  ③ 整合：进 应收账款审计底稿（special_reviews 路由）

用法：
  python aging_review.py            # 全部
  python aging_review.py XBJ        # 指定账套
"""
import argparse
import os
import re
import sys
from collections import defaultdict
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import aging_manifest as AM
import audit_common as A

# 样式
F_HEAD = Font(name='微软雅黑', size=10, bold=True)
F = Font(name='微软雅黑', size=10)
THIN = Border(*[Side(style='thin', color='BBBBBB')] * 4)
FILL_HDR = PatternFill('solid', fgColor='D9E2F3')
FILL_WARN = PatternFill('solid', fgColor='FFF2CC')
NUM = '#,##0.00'
PCT = '0.0%'

BUCKETS = ['1年以内', '1-2年', '2-3年', '3-4年', '4-5年', '5年以上']


def _as_float(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0


def _age_bucket_idx(date_str, end_date):
    """按日期计算账龄桶索引。date=交易日期，end=期末（2025-12-31）。"""
    try:
        d = datetime.strptime(str(date_str)[:10], '%Y-%m-%d')
        e = datetime.strptime(end_date, '%Y-%m-%d')
        days = (e - d).days
    except (ValueError, TypeError):
        return 5  # 无法解析归 5年以上（保守）
    if days <= 365:
        return 0
    if days <= 730:
        return 1
    if days <= 1095:
        return 2
    if days <= 1460:
        return 3
    if days <= 1825:
        return 4
    return 5


def _settle_cycle_months(cycle):
    """结算周期 → 月数（合同资产到期时点判断）。月结=1/季结=3/半年=6/年结=12/竣工=None(竣工时点)。"""
    s = str(cycle or '').strip()
    if '月' in s and '季' not in s:
        m = re.search(r'(\d+)', s)
        return int(m.group(1)) if m else 1
    if '季' in s:
        return 3
    if '半年' in s:
        return 6
    if '年' in s and '半年' not in s:
        return 12
    return None   # 竣工结算/其他 → 未知，回退比例法


def _aux_ca_opening(data_dir, ent):
    """读主体辅助核算：合同资产【期初余额】合计（跨年滚入批次，实然判断用）。"""
    import glob as _glob
    pat = os.path.join(data_dir, f'*{ent}*辅助核算余额表*.xlsx')
    files = _glob.glob(pat)
    if not files:
        return 0.0
    try:
        wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception:
        return 0.0
    hi = 0
    for i, r in enumerate(rows[:6]):
        if r and any('科目' in str(x) for x in r):
            hi = i
            break
    if hi >= len(rows):
        return 0.0
    hdr = rows[hi]
    # 定位 期初金额 列（名称含'期初'且含'金额/余额'）
    qc_col = None
    for i, h in enumerate(hdr):
        hs = str(h or '')
        if '期初' in hs and ('金额' in hs or '余额' in hs or '方向' in hs):
            qc_col = i
            if '金额' in hs or '余额' in hs:
                break
    if qc_col is None:
        return 0.0
    tot = 0.0
    for r in rows[hi + 1:]:
        if not r or len(r) <= qc_col:
            continue
        if '合同资产' not in str(r[0] or ''):
            continue
        tot += _as_float(r[qc_col])
    return tot


def _ca_unexpired_by_actual(data_dir, cfg, ent, end_date):
    """⚡ 实然优先（2026-08-16 用户：合同虽规定但执行偏差大——以账面实际执行为准）：
    从该主体 GL 合同资产流水反推【实际结算周期】（FIFO：贷方转出冲最早借方确认，
    持有天数按确认金额加权平均），再判断期末未转出的确认批次：
      确认日距期末 < 实然周期 → 未到期（尚未到实际转出时点，不进账龄）
      确认日距期末 ≥ 实然周期 → 已到期（应转未转 → 逾期进账龄）
    返回 (unexpired, desc) / (None, None) 无配对数据（交给合同台账/比例兜底）。
    desc 示例: '账面反推实然周期5.5月'。
    """
    import glob as _glob
    pat = os.path.join(data_dir, f'*{ent}*综合查询明细表*.xlsx')
    files = _glob.glob(pat)
    if not files:
        return None, None
    try:
        wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception:
        return None, None
    hi = 0
    for i, r in enumerate(rows[:6]):
        if r and any('科目' in str(x) for x in r) and any('金额' in str(x) for x in r):
            hi = i
            break
    ca_in = []    # (日期, 借金额) 确认
    ca_out = []   # (日期, 贷金额) 转出
    for r in rows[hi + 1:]:
        if not r or len(r) < 8:
            continue
        if '合同资产' not in str(r[0] or ''):
            continue
        dt = str(r[1] or '')[:10]
        jf = _as_float(r[6])
        df = _as_float(r[7])
        if jf > 0.005:
            ca_in.append((dt, jf))
        if df > 0.005:
            ca_out.append((dt, df))
    if not ca_in or not ca_out:
        return None, None
    try:
        from datetime import datetime as _dt
        from datetime import timedelta as _td
        _P = lambda s: _dt.strptime(str(s)[:10], '%Y-%m-%d')
        in_q = sorted([(_P(d), a) for d, a in ca_in], key=lambda x: x[0])
        out_l = sorted([(_P(d), a) for d, a in ca_out], key=lambda x: x[0])
        end = _dt.strptime(end_date, '%Y-%m-%d')
    except (ValueError, TypeError):
        return None, None
    # ⚡ 期初滚入批次（跨年合同资产）：近似确认日=年初（持有≈365天），FIFO 最先被冲；
    #    只参与【残留/未到期判断】，不参与周期反推（365 天是近似，会污染实然周期）。
    open_ca = _aux_ca_opening(data_dir, ent)
    open_d = end - _td(days=365)
    # 一次性完整 FIFO：队列=期初批次 + 本期确认（期初在前）；转出按日期顺序冲
    q = [[open_d, open_ca]] if open_ca > 0.005 else []
    q += [[d, a] for d, a in in_q]
    holds = []   # 仅本期精确配对记录 (确认日, 持有天数, 金额)
    for od, oa in out_l:
        rem = oa
        while rem > 0.005 and q:
            id_, ia = q[0]
            take = min(ia, rem)
            if id_ != open_d:      # 本期确认被冲 → 精确配对记录
                hold = (od - id_).days
                if hold >= 0:
                    holds.append((id_, hold, take))
            q[0][1] = ia - take
            rem -= take
            if q[0][1] <= 0.005:
                q.pop(0)
    if not holds:
        return None, None
    total = sum(a for _, _, a in holds)
    total_in_cur = sum(a for _, a in in_q)
    if total_in_cur > 0.005 and total < total_in_cur * 0.3:
        # 本期确认配对率 <30% → 数据不足，诚实回落合同台账/比例
        return None, None
    avg_days = sum(h * a for _, h, a in holds) / total
    if avg_days < 15:
        # 周期 <15 天 = 同日/隔日重分类干扰（红字冲回等），非真实结算行为 → 回落
        return None, None
    unexp = sum(a for d, a in q if a > 0.005 and (end - d).days < avg_days)
    desc = f'账面反推实然周期{avg_days / 30.44:.1f}月'
    return unexp, desc


def _ca_unexpired_by_contracts(data_dir, cfg, ent, ca_end, end_date):
    """⚡⚡ 2026-08-16 用户要求：合同资产未到期判断按【每个合同的具体付款条件】，
    与项目制收入底稿（project_manifest.contracts）同源。

    合同资产=已确认收入尚未到结算条件的部分。未到期金额按业主（辅助核算往来单位）匹配
    合同台账的 结算周期（settle_cycle：月结=1月/季结=3月/竣工=竣工时点）+ 质保期（quality_months）：
      · 合同资产本期新增（辅助核算本期借方）按 结算周期 推 应转应收时点：
          确认日 + 结算周期 <= 期末 → 已到期（进账龄）
          确认日 + 结算周期 >  期末 → 未到期（不进账龄）
      · 质保金（辅助核算科目名含 '质保'）按 quality_months：质保期内未到期
      · 无台账项目/未知周期 → 回退 ca_cats 类别比例（默认 0.8）
    返回 (未到期金额, 匹配方式) — 匹配方式: '账面反推实然周期X.X月' / '合同台账' / '类别比例'。
    """
    # ⚡⚡ 实然优先（用户 2026-08-16：合同规定≠实际执行，偏差大）：
    # 先从账面流水反推实际结算周期——有配对数据即用实然周期（合同台账仅作对照，
    # 不再作为唯一依据）；无流水数据再回落合同台账/类别比例。
    try:
        unexp_act, desc_act = _ca_unexpired_by_actual(data_dir, cfg, ent, end_date)
    except Exception:
        unexp_act, desc_act = None, None
    if unexp_act is not None:
        return min(unexp_act, max(0.0, ca_end)), desc_act
    # 读合同台账（与项目制收入底稿同源）
    contracts = cfg.get('contracts', [])
    try:
        import project_manifest as PM
        if not contracts:
            _pc = PM.jobs().get(cfg.get('acct', ''), {}).get('subs', {}).get('_root', {})
            contracts = _pc.get('contracts', []) or []
    except Exception:
        pass
    # 读主体辅助核算：合同资产 按 业主（往来单位）→ {unit: {'jf': 本期借, 'qm': 期末, 'is_qual': 是否质保金}}
    aux_ca = {}
    import glob as _glob
    pat = os.path.join(data_dir, f'*{ent}*辅助核算余额表*.xlsx')
    files = _glob.glob(pat)
    if files:
        try:
            wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
        except Exception:
            rows = []
        hi = 0
        for i, r in enumerate(rows[:6]):
            if r and any('科目' in str(x) for x in r):
                hi = i
                break
        for r in rows[hi + 1:]:
            if not r or len(r) < 8:
                continue
            km = str(r[0] or '')
            if '合同资产' not in km:
                continue
            unit = str(r[1] or '').strip()
            if not unit:
                continue
            d = aux_ca.setdefault(unit, {'jf': 0.0, 'qm': 0.0, 'is_qual': '质保' in km})
            d['jf'] += _as_float(r[4])          # 本期借方（新增合同资产）
            d['qm'] += _as_float(r[7])          # 期末余额
    if not aux_ca:
        # 无辅助核算 → 回退比例法
        return ca_end * 0.8, '类别比例'
    unexp_total = 0.0
    matched = 0
    # 合同台账按 project 键（业主名）索引
    cmap = {}
    for c in contracts:
        cmap.setdefault(str(c.get('project', '')), c)
        cmap.setdefault(str(c.get('customer', '')), c)
    for unit, d in aux_ca.items():
        c = cmap.get(unit)
        if c:
            cycle = _settle_cycle_months(c.get('settle_cycle', ''))
            qm_months = int(c.get('quality_months') or 0)
            # 本期新增部分按结算周期判断到期
            jf = d['jf']
            if d['is_qual']:
                # 质保金：质保期内未到期（默认按质量期，未配=1年）
                if qm_months > 0:
                    unexp_total += min(d['qm'], jf * (qm_months / 12.0) if jf else d['qm'])
                else:
                    unexp_total += d['qm']      # 无质保期配置 → 全额未到期（保守）
            elif cycle:
                # 工程款：结算周期内新增=未到期（期末留存均为结算周期内新增）
                # 简化：未到期≈期末×min(1, cycle/12)（周期内新增未到期；超出=已到期）
                unexp_total += d['qm'] * min(1.0, cycle / 12.0)
            else:
                unexp_total += d['qm'] * 0.8    # 竣工结算/未知 → 回退
            matched += 1
        else:
            # 无台账合同 → 类别比例（工程款 0.8/质保 1.0）
            unexp_total += d['qm'] * (1.0 if d['is_qual'] else 0.8)
    # 兜底：未到期不超过期末
    return min(unexp_total, max(0.0, ca_end)), ('合同台账' if matched else '类别比例')


# ---------------------------------------------------------------- 读取
def _load_gl_flows(data_dir, cfg):
    """从序时账提取 应收/合同资产 明细流水（含凭证号——互转识别基础）。

    返回 {'ar': [(主体, 日期, 凭证号, 金额, 摘要)],
          'ca': [(主体, 日期, 凭证号, 金额, 摘要, 类别)]}
    金额=借方-贷方（正=增加应收/合同资产）。
    ⚡ 凭证号列=index3（U8 序时账第 4 列），用于互转识别（应收借方=合同资产贷方）。
    """
    codes = cfg['codes']
    # 名称关键词（序时账科目列=名称，非科目码——名称主键，铁律5）
    ar_kw = ['应收账款']
    ca_kw = ['合同资产']
    entities = A.discover_entities(data_dir)
    flows = {'ar': [], 'ca': []}
    for ent in sorted(entities):
        for yy, paths in entities[ent].items():
            gl = paths.get('gl')
            if not gl:
                continue
            try:
                wb = openpyxl.load_workbook(gl, read_only=True, data_only=True)
                ws = wb[wb.sheetnames[0]]
                rows = list(ws.iter_rows(values_only=True))
                wb.close()
            except Exception as ex:
                print(f'  ⚠️ {ent[:20]} GL 读取失败: {ex}')
                continue
            # 表头行
            hi = 0
            for i, r in enumerate(rows[:6]):
                if r and any('科目' in str(x) for x in r) and any('金额' in str(x) for x in r):
                    hi = i
                    break
            for r in rows[hi + 1:]:
                if not r or len(r) < 8:
                    continue
                km = str(r[0] or '')
                dt = str(r[1] or '')[:10]
                vno = str(r[3] or '') if len(r) > 3 else ''
                jf = _as_float(r[6])
                df = _as_float(r[7])
                sm = str(r[5] or '')
                amt = jf - df
                if abs(amt) < 0.005:
                    continue
                # 名称匹配（含辅助核算明细行：科目名称可能含"应收账款\\单位\\合同"）
                is_ar = any(k in km for k in ar_kw) and '合同资产' not in km
                is_ca = any(k in km for k in ca_kw)
                if is_ar:
                    flows['ar'].append((ent, dt, vno, amt, sm))
                elif is_ca:
                    # 合同资产类别（从名称含的子段识别）
                    cat = '其他'
                    for cc in cfg.get('ca_cats', []):
                        if cc['kw'] != '_default' and cc['kw'] in km:
                            cat = cc['kw'].replace('金', '')
                            break
                    flows['ca'].append((ent, dt, vno, amt, sm, cat))
    return flows


def _identify_transfers(ar_flows, ca_flows):
    """凭证号级互转识别：同主体同凭证号内 应收借方 = 合同资产贷方。

    互转（合同资产到期转应收）在合并账龄时应剔除——同一笔钱在应收借方
    与合同资产贷方重复出现，若各自分龄会双计。

    返回 (transfer_keys, net_flows)：
      transfer_keys = {(主体, 凭证号)} 互转凭证集合
      net_flows     = 剔除互转后的纯新增流水（保留原结构）
    """
    # 同主体同凭证号：应收借方总额 vs 合同资产贷方总额
    ar_jf_by = defaultdict(float)   # (ent,vno) -> 应收借方
    ca_df_by = defaultdict(float)   # (ent,vno) -> 合同资产贷方
    for ent, dt, vno, amt, sm in ar_flows:
        if amt > 0 and vno:
            ar_jf_by[(ent, vno)] += amt
    for ent, dt, vno, amt, sm, cat in ca_flows:
        if amt < 0 and vno:
            ca_df_by[(ent, vno)] += abs(amt)
    transfer_keys = set()
    for key in ar_jf_by:
        if key in ca_df_by and abs(ar_jf_by[key] - ca_df_by[key]) < 1.0:
            transfer_keys.add(key)
    return transfer_keys


def _actual_settle_cycle(ca_in, ca_out):
    """⚡ 实然结算周期反推（铁律131f-双轨 2026-08-16）：
    从合同资产流水反推【实际执行】的结算周期——合同规定（应然）vs 账面执行（实然）
    的偏差正是审计发现点（合同说月结、实际半年才转应收）。
    方法：FIFO——贷方转出按日期从早到晚，优先冲最早的借方确认；
    每笔确认被冲时计算持有天数，按确认金额加权平均 → 月数。
    返回 (avg_months, matched_amt, unmatched_in, unmatched_out, n_pairs)：
      avg_months  加权平均持有月数（无配对数据=None）
      matched_amt 参与配对的确认金额合计
      unmatched_in/out 未配对金额（期初滚入/期末未转出）
      n_pairs     配对笔数
    """
    try:
        from datetime import datetime as _dt
        _P = lambda s: _dt.strptime(str(s)[:10], '%Y-%m-%d')
        in_q = sorted([(_P(d), a) for d, a in ca_in if a > 0.005], key=lambda x: x[0])
        in_q0 = list(in_q)
        out_l = sorted([(_P(d), a) for d, a in ca_out if a > 0.005], key=lambda x: x[0])
    except (ValueError, TypeError):
        return None, 0.0, 0.0, 0.0, 0
    wsum = 0.0
    total = 0.0
    n = 0
    qi = 0  # 队列头指针（in_q 已经按日期排序，直接顺序消费）
    for od, oa in out_l:
        rem = oa
        while rem > 0.005 and qi < len(in_q):
            id_, ia = in_q[qi]
            take = min(ia, rem)
            hold = (od - id_).days
            if hold > 0:
                wsum += take * hold
                total += take
                n += 1
            in_q[qi] = (id_, ia - take)
            rem -= take
            if ia - take <= 0.005:
                qi += 1
    matched = total
    total_in = sum(a for _, a in in_q0) if 'in_q0' in dir() else 0.0
    unmatched_in = sum(a for _, a in in_q if a > 0.005)
    consumed = total_in - unmatched_in
    total_out = sum(a for _, a in out_l)
    unmatched_out = max(0.0, total_out - consumed)  # 期初滚入被转出（无确认可配）
    if total < 0.005:
        return None, 0.0, total_in, total_out, 0
    avg_months = wsum / total / 30.44 if total > 0.005 else None
    return avg_months, matched, unmatched_in, unmatched_out, n


def _fifo_age(open_balance, flows, end_date):
    """FIFO 账龄：期末余额按'先进先出'摊入各账龄桶。

    输入：
      open_balance  期初余额（正=借余；可负）
      flows         本年逐笔流水 [(日期, 金额)]，正=增加、负=收回/转出
    原理：
      · 期初余额视为"最先发生"，整笔归 1-2年桶（形成于上年）
      · 本年流水按日期从早到晚，累计净发生；FIFO 抵销：收回先冲期初，
        再冲最早的本年增加
      · 期末余额 = 期初 + Σ流水（与 TB 控制数核对）
      · 分龄：期末余额按"还剩哪些批次"摊入对应桶（期初批次→1-2年，
        本年批次→按其日期对应桶）
    返回 (buckets, end_calc)
    """
    buckets = [0.0] * 6
    # 本年流水按日期排序
    sorted_f = sorted(flows, key=lambda x: x[0])
    # 累计净发生
    net = sum(a for d, a in sorted_f)
    end_calc = open_balance + net
    if abs(end_calc) < 0.005:
        return buckets, 0.0
    # 本年增加批次（正发生，按日期）
    pos_batches = [(d, a) for d, a in sorted_f if a > 0]
    neg_total = sum(a for d, a in sorted_f if a < 0)  # 负数
    # FIFO：收回（neg_total）先冲期初，再冲本年最早增加
    # 期初残留 = max(0, 期初 + 收回)
    open_residual = max(0.0, open_balance + neg_total)
    # 本年批次摊余：pos_batches 按日期累计，减去被冲部分
    if open_balance > 0:
        # 期初有余额，收回优先冲期初 → 本年正发生可能也被部分冲减
        # 期末仍保留期初部分的 = open_residual（归 1-2年）
        pass
    # 计算各批次贡献（含期初残留）
    if end_calc > 0:
        # 正向余额：先期初残留(1-2年)，再按日期早→晚取本年正发生
        remaining = end_calc
        if open_residual > 0:
            take = min(remaining, open_residual)
            buckets[1] += take
            remaining -= take
        # 本年正发生按日期从早到晚（先发生先被期末保留——FIFO 先进先出，
        # 期末保留的是【最近发生】的，所以从晚到早取！）
        for d, a in reversed(pos_batches):
            if remaining <= 0:
                break
            take = min(remaining, a)
            buckets[_age_bucket_idx(d, end_date)] += take
            remaining -= take
        # 尾差（如有）
        if remaining > 1:
            buckets[5] += remaining
    else:
        # 负向余额（贷方，预收类）→ 全部归 1年以内（简化，罕见）
        buckets[0] += end_calc
    return buckets, end_calc


# ---------------------------------------------------------------- 输出
def _write_workbook(out_path, cfg, flows, tb_totals, mode):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    year = cfg['year']
    end_date = f'{year}-12-31'
    data_dir = cfg.get('data_dir', '')

    # 期初余额（TB qc，按主体）——FIFO 分龄起点
    ent_open_ar = defaultdict(float)
    ent_open_ca = defaultdict(float)
    for prefix in cfg['codes']['ar']:
        for row in tb_totals.get('_rows', []):
            e, c, n, y, v = row
            if str(c).startswith(prefix):
                ent_open_ar[e] += _as_float(v.get('qc'))
    for prefix in cfg['codes']['ca']:
        for row in tb_totals.get('_rows', []):
            e, c, n, y, v = row
            if str(c).startswith(prefix):
                ent_open_ca[e] += _as_float(v.get('qc'))

    # ---------- Sheet1 账龄汇总（按主体+科目） ----------
    ws = wb.create_sheet('账龄汇总(按主体)')
    ws.append(['应收 + 合同资产 联动账龄复核（%d 年）' % year])
    ws.append(['口径：先合并剔除互转（凭证号级），再统一 FIFO 分龄；合同资产按类别未到期比例剥离（分科账龄+计提见 Sheet2）。'])
    ws.append([])
    hdr = ['核算主体', '期初应收', '应收账款期末', '合同资产期末', '合同资产未到期', '合同资产已到期',
           '应收+已到期合计'] + BUCKETS + ['账龄合计', '核对(期末-合计)']
    ws.append(hdr)

    # ⚡⚡ 2026-08-16 用户方法论重构：
    #   ① 识别互转（凭证号级：应收借方=合同资产贷方，合同资产到期转应收）
    #   → ② 互转的应收借方【正常进应收账龄】（它本就是合同资产已到期），
    #       合同资产侧只对【未到期部分】剥离（互转部分已到期不再计入未到期）
    #   → ③ 分科账龄（应收全量 / 合同资产已到期）→ ④ 各自计提比例测算坏账准备
    transfer_keys = _identify_transfers(flows['ar'], flows['ca'])

    # 逐主体：期末 = TB 控制数（qm），账龄桶 = FIFO（期初残留1-2年 + 本年批次按日期）
    ent_ar = defaultdict(float)
    ent_ca = defaultdict(float)
    ent_ca_unexp = defaultdict(float)
    ca_unexpired_by = defaultdict(str)            # 未到期判断方式：合同台账/类别比例
    ent_age = defaultdict(lambda: [0.0] * 6)       # 合并债权池账龄
    ent_age_ar = defaultdict(lambda: [0.0] * 6)    # 应收账款账龄
    ent_age_ca = defaultdict(lambda: [0.0] * 6)    # 合同资产账龄（已到期部分）
    ent_prov = defaultdict(float)                  # 计提坏账准备（应收+合同资产）
    all_ents = sorted(set(ent_open_ar) | set(ent_open_ca) |
                      set(e for e, d, v, a, s in flows['ar']) |
                      set(e for e, d, v, a, s, c in flows['ca']))
    for ent in all_ents:
        open_ar = ent_open_ar[ent]
        open_ca = ent_open_ca[ent]
        # ⚡ 互转处理：应收借方（互转）正常进应收账龄（=合同资产到期转应收）；
        #   合同资产贷方（互转）标记已到期——不计入未到期剥离
        ar_flows = [(d, a) for e, d, v, a, s in flows['ar'] if e == ent]
        ca_flows = [(d, a) for e, d, v, a, s, c in flows['ca'] if e == ent]
        ar_all = ar_flows
        ca_all = ca_flows
        # 期末 = TB 控制数（优先）
        ar_end = tb_totals.get('_ent_ar', {}).get(ent, open_ar + sum(a for d, a in ar_all))
        ca_end = tb_totals.get('_ent_ca', {}).get(ent, open_ca + sum(a for d, a in ca_all))
        ent_ar[ent] = ar_end
        ent_ca[ent] = ca_end
        # ① 应收全量 FIFO 分龄（含互转转入——它们就是合同资产已到期）
        ar_buckets, _ = _fifo_age(open_ar, ar_flows, end_date)
        # ② 合同资产分龄：仅【未到期剥离后已到期】进桶；互转贷方（转出）不计未到期
        # ⚡⚡ 2026-08-16 用户要求：未到期按合同台账付款条件（结算周期/质保期）逐项目判断，
        #    替代固定比例——_ca_unexpired_by_contracts 匹配不到时回退 0.8/类别比例
        unexp, unexp_by = _ca_unexpired_by_contracts(data_dir, cfg, ent, ca_end, end_date)
        ca_unexpired_by[ent] = unexp_by
        ca_end_pos = max(0.0, ca_end)
        ca_exp_total = ca_end_pos - unexp
        ca_buckets, ca_end_calc = _fifo_age(open_ca, ca_flows, end_date)
        if ca_end_calc > 0 and sum(ca_buckets) > 0:
            scale = ca_exp_total / sum(ca_buckets)
            ca_buckets = [b * scale for b in ca_buckets]
        else:
            ca_buckets = [0.0] * 6
        # ③ 合并账龄 + 分科账龄（应收全量 / 合同资产已到期）
        for i in range(6):
            ent_age[ent][i] += ar_buckets[i] + ca_buckets[i]
            ent_age_ar[ent][i] += ar_buckets[i]
            ent_age_ca[ent][i] += ca_buckets[i]
        # 未到期 = 合同台账付款条件判断（无台账回退期末×0.8/类别比例）
        unexp = min(unexp, max(0.0, ca_end))
        ent_ca_unexp[ent] = unexp
        exp = ca_end - unexp
        total_ar = ar_end + exp
        # ④ 计提坏账准备（应收比例 × 应收桶 + 合同资产比例 × 合同资产桶）
        prov_ar = cfg['provision']['ar']
        prov_ca = cfg['provision']['ca']
        ent_prov[ent] = sum(ent_age_ar[ent][i] * prov_ar[i] for i in range(6)) + \
                        sum(ent_age_ca[ent][i] * prov_ca[i] for i in range(6))
        age_sum = sum(ent_age[ent])
        row = [ent, round(open_ar, 2), round(ar_end, 2), round(ca_end, 2), round(unexp, 2), round(exp, 2),
               round(total_ar, 2)] + [round(x, 2) for x in ent_age[ent]] + \
              [round(age_sum, 2), round(total_ar - age_sum, 2)]
        ws.append(row)
    # 全集团合计
    tot_ar = sum(ent_ar.values()); tot_ca = sum(ent_ca.values())
    tot_unexp = sum(ent_ca_unexp.values())
    tot_exp = tot_ca - tot_unexp
    tot_age = [sum(ent_age[e][i] for e in ent_age) for i in range(6)]
    grow = ['全集团合计', round(sum(ent_open_ar.values()), 2), round(tot_ar, 2), round(tot_ca, 2),
            round(tot_unexp, 2), round(tot_exp, 2), round(tot_ar + tot_exp, 2)] + \
           [round(x, 2) for x in tot_age] + [round(sum(tot_age), 2),
                                             round(tot_ar + tot_exp - sum(tot_age), 2)]
    ws.append(grow)
    _style(ws, n_col=15)

    # ---------- Sheet2 分科账龄 + 计提测算 ----------
    ws2 = wb.create_sheet('分科账龄与计提')
    ws2.append(['应收账款 / 合同资产 分科账龄 + 坏账准备计提测算（%d 年）' % year])
    ws2.append(['口径：互转（合同资产到期转应收）计入应收账龄；合同资产仅未到期剥离后已到期分龄；两科各自计提比例。'])
    ws2.append([])
    hdr2 = ['核算主体', '应收 1年内', '应收 1-2', '应收 2-3', '应收 3-4', '应收 4-5', '应收 5年以上',
            '合同资产 1年内', '合同资产 1-2', '合同资产 2-3', '合同资产 3-4', '合同资产 4-5', '合同资产 5年以上',
            '应收计提', '合同资产计提', '计提合计']
    ws2.append(hdr2)
    prov_ar = cfg['provision']['ar']; prov_ca = cfg['provision']['ca']
    tot_age_ar = [0.0] * 6; tot_age_ca = [0.0] * 6
    for ent in all_ents:
        a_ar = ent_age_ar[ent]; a_ca = ent_age_ca[ent]
        p_ar = sum(a_ar[i] * prov_ar[i] for i in range(6))
        p_ca = sum(a_ca[i] * prov_ca[i] for i in range(6))
        for i in range(6):
            tot_age_ar[i] += a_ar[i]; tot_age_ca[i] += a_ca[i]
        ws2.append([ent] + [round(x, 2) for x in a_ar] + [round(x, 2) for x in a_ca] +
                   [round(p_ar, 2), round(p_ca, 2), round(p_ar + p_ca, 2)])
    g_p_ar = sum(tot_age_ar[i] * prov_ar[i] for i in range(6))
    g_p_ca = sum(tot_age_ca[i] * prov_ca[i] for i in range(6))
    ws2.append(['全集团合计'] + [round(x, 2) for x in tot_age_ar] + [round(x, 2) for x in tot_age_ca] +
               [round(g_p_ar, 2), round(g_p_ca, 2), round(g_p_ar + g_p_ca, 2)])
    ws2.append([])
    ws2.append(['计提比例（应收）：', '1年内', '1-2', '2-3', '3-4', '4-5', '5年以上'])
    ws2.append([''] + [f'{r*100:.0f}%' for r in prov_ar])
    ws2.append(['计提比例（合同资产）：', '1年内', '1-2', '2-3', '3-4', '4-5', '5年以上'])
    ws2.append([''] + [f'{r*100:.0f}%' for r in prov_ca])
    ws2.append([])
    ws2.append(['⚠ 计提比例为默认占位（应收 5%~100%/合同资产 1%~50%），待您提供公司坏账政策后覆盖（manifest provision）。'])
    _style(ws2, n_col=16)

    # ---------- Sheet3 互转抵消明细 ----------
    ws3 = wb.create_sheet('互转抵消明细')
    ws3.append(['合同资产→应收账款 互转（凭证号级识别）'])
    ws3.append(['同主体同凭证号内 应收借方 = 合同资产贷方 → 合同资产到期转应收，'
                '该笔计入应收账龄、不计入合同资产未到期（避免双计）'])
    ws3.append([])
    ws3.append(['核算主体', '凭证号', '日期', '应收借方(转出合同资产到期)', '合同资产贷方', '互转金额', '摘要'])
    n_trans = 0
    shown = 0
    for ent, dt, vno, amt, sm in sorted(flows['ar'], key=lambda x: x[1]):
        if (ent, vno) in transfer_keys:
            n_trans += 1
            if shown < 200:
                ws3.append([ent, vno, dt, round(amt, 2), '', round(amt, 2), sm[:30]])
                shown += 1
    ws3.append([])
    ws3.append([f'互转凭证数: {n_trans} 个（前 {shown} 个列示；互转金额已计入应收账龄、不计合同资产未到期）'])
    _style(ws3, n_col=7)

    # ---------- Sheet4 账龄明细(FIFO) ----------
    ws4 = wb.create_sheet('账龄明细(FIFO)')
    ws4.append(['应收 + 合同资产已到期 账龄明细（GL 交易日期 FIFO）'])
    ws4.append(['互转（应收借方=合同资产贷方）正常计入应收账龄；合同资产未到期按比例剥离。'])
    ws4.append([])
    ws4.append(['核算主体', '科目类型', '交易日期', '金额', '账龄桶', '摘要'])
    for ent, dt, vno, amt, sm in sorted(flows['ar'], key=lambda x: x[1]):
        idx = _age_bucket_idx(dt, end_date)
        tag = ' [互转]' if (ent, vno) in transfer_keys else ''
        ws4.append([ent, '应收账款' + tag, dt, round(amt, 2), BUCKETS[idx], sm[:30]])
    for ent, dt, vno, amt, sm, cat in sorted(flows['ca'], key=lambda x: x[1]):
        unexp_ratio = 0.8
        for cc in cfg.get('ca_cats', []):
            if cc['kw'] != '_default' and cc['kw'] in cat:
                unexp_ratio = cc['unexpired']
                break
        if amt > 0:
            exp_amt = amt * (1 - unexp_ratio)
            idx = _age_bucket_idx(dt, end_date)
            ws4.append([ent, f'合同资产-{cat}', dt, round(exp_amt, 2), BUCKETS[idx], sm[:30]])
        else:
            idx = _age_bucket_idx(dt, end_date)
            tag = ' [互转]' if (ent, vno) in transfer_keys else ''
            ws4.append([ent, f'合同资产-{cat}{tag}', dt, round(amt, 2), BUCKETS[idx], sm[:30]])
    _style(ws4, n_col=6)

    # ---------- Sheet5 与 TB 核对 ----------
    ws5 = wb.create_sheet('与TB核对')
    ws5.append(['与科目余额表核对（期末）'])
    ws5.append([])
    ws5.append(['项目', '账龄表合计(期末)', 'TB 期末', '差异', '说明'])
    codes = cfg['codes']
    tb = tb_totals or {}
    for label, prefix, key in [('应收账款', codes['ar'][0], 'ar'),
                               ('合同资产', codes['ca'][0], 'ca'),
                               ('坏账准备', codes['baddebt'][0], 'baddebt')]:
        tb_val = tb.get(prefix, 0.0)
        if key == 'ar':
            calc = sum(ent_ar.values())
        elif key == 'ca':
            calc = sum(ent_ca.values())
        else:
            calc = 0.0
        ws5.append([label, round(calc, 2), round(tb_val, 2), round(calc - tb_val, 2),
                    '' if abs(calc - tb_val) < 1 else '差异需查'])
    ws5.append([])
    ws5.append(['注：账龄复核以序时账流水为主（账龄=流水属性）；应收/合同资产期末=期初+本年净发生。',
               '', '', '', '', ''])
    ws5.append(['    TB 控制数为参考——差异>1万元=该主体序时账流水可能不完整（部分往来走辅助核算/未导出），',
               '', '', '', '', ''])
    ws5.append(['    需补全流水或按 TB 差异调整（与应收账款账龄分析表同口径披露）。',
               '', '', '', '', ''])
    _style(ws5, n_col=5)

    # ---------- Sheet6 口径说明 ----------
    ws6 = wb.create_sheet('口径说明')
    notes = [
        '1. 依据：企业会计准则第14号——收入（合同资产=已履约未获无条件收款权）；第22号——金融工具（坏账准备=预期信用损失）。',
        '2. ⚡ 方法论（2026-08-16 用户定）：①【合并剔除互转】同主体同凭证号 应收借方=合同资产贷方 的互转',
        '   （合同资产到期转应收）先抵消，避免双计；②统一债权池 FIFO 分龄（期初残留1-2年+本年按日期）；',
        '   ③【按合同/来源拆回两科】应收全量 + 合同资产已到期部分，各自维护六档账龄；',
        '   ④【各自计提比例】应收账龄桶 × 应收比例 + 合同资产账龄桶 × 合同资产比例 = 坏账准备测算。',
        '3. 账龄桶：1年以内/1-2/2-3/3-4/4-5/5年以上（六档，全集团统一）。',
        '4. ⚡⚡ 合同资产未到期判断（2026-08-16 用户要求，与项目制收入底稿同源）——【实然优先三级】：',
        '   合同资产=已确认收入尚未到结算条件的部分。用户方法论：合同规定≠实际执行、执行偏差大——以账面实际执行为准。',
        '   ① 实然优先：从账面 GL 合同资产流水 FIFO 反推【实际结算周期】（确认→转出持有天数按金额加权），',
        '      确认日距期末<实然周期=未到期；≥实然周期=应转未转逾期进账龄；期初滚入超周期即已到期；',
        '      本期配对率<30%或周期<15天（同日重分类干扰）→ 数据不足回落下一级（诚实标注）；',
        '   ② 合同台账（应然，project_manifest.contracts）：结算周期（月结/季结/半年/年/竣工）+质保期按业主匹配；',
        '   ③ 类别比例兜底（无台账无流水）：质保金/审计金=100%未到期、工程款=80%未到期。',
        '   未到期部分单独列示，已到期并入应收统一分龄；判断方式逐主体在汇总表标注。',
        '5. 计提比例（manifest provision，默认占位）：应收 5%/10%/20%/30%/50%/100%，合同资产 1%/5%/10%/15%/20%/50%',
        '   ——待您提供公司坏账政策后覆盖（应收与合同资产比例可不同）。',
        '6. 后续细化：合同台账到位后，未到期比例/计提比例可逐合同覆盖。',
        f'7. 复核主体范围：{cfg["org"]}；年度：{year}。',
        '8. 差异>1元需查；坏账准备（1231）单列不参与分龄；互转明细见 Sheet3。',
    ]
    for n in notes:
        ws6.append([n])
    ws6.column_dimensions['A'].width = 110

    wb.save(out_path)
    wb.close()


def _style(ws, n_col):
    for r in ws.iter_rows(min_row=1, max_row=min(3, ws.max_row), max_col=n_col):
        for c in r:
            if c.value is not None:
                c.font = F_HEAD
                c.fill = FILL_HDR
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=n_col):
        for c in r:
            if c.value is None:
                continue
            c.border = THIN
            if isinstance(c.value, (int, float)):
                c.number_format = NUM
    for i in range(1, n_col + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16
    if ws.max_row > 3:
        ws.freeze_panes = ws.cell(row=4, column=1).coordinate


# ---------------------------------------------------------------- 主流程
def run_one(acct, sub, out_dir=None):
    acct_cfg, sub_cfg = AM.job_for(acct, sub)
    cfg = dict(sub_cfg)
    # ⚡ 计提比例默认占位（manifest 未配置时用 DEFAULT_PROVISION；用户提供公司政策后覆盖）
    cfg.setdefault('provision', AM.DEFAULT_PROVISION)
    # ⚡⚡ 2026-08-16：注入账套名 + 合同台账（与项目制收入底稿同源，付款条件判断未到期）
    cfg['acct'] = acct
    try:
        import project_manifest as PM2
        _pc = PM2.jobs().get(acct, {}).get('subs', {}).get(sub, {})
        cfg.setdefault('contracts', _pc.get('contracts', []) or [])
    except Exception:
        pass
    year = cfg['year']
    data_dir = os.path.join(AM.DATA_ROOT, cfg['data_dir'].replace('/', os.sep))
    cfg['data_dir'] = data_dir
    print(f'=== {acct}/{sub} 应收+合同资产账龄复核 ===')
    print(f'  数据目录: {data_dir}')

    flows = _load_gl_flows(data_dir, cfg)
    n_ar = len(flows['ar']); n_ca = len(flows['ca'])
    print(f'  序时账流水: 应收 {n_ar} 笔, 合同资产 {n_ca} 笔')
    if n_ar + n_ca == 0:
        print('  ⚠️ 无应收/合同资产流水')
        return 1

    # TB 期末（核对基准）
    entities = A.discover_entities(data_dir)
    tb = A.read_tb_full(data_dir, entities) if entities else {}
    tb_totals = {'_rows': [], '_ent_ar': defaultdict(float), '_ent_ca': defaultdict(float)}
    codes = cfg['codes']
    for prefix in codes['ar'] + codes['ca'] + codes['baddebt']:
        tot = 0.0
        for (e, c, n, y), v in tb.items():
            if str(c).startswith(prefix) and str(y) == str(year):
                tot += _as_float(v.get('qm'))
                tb_totals['_rows'].append((e, c, n, y, v))
                if prefix in codes['ar']:
                    tb_totals['_ent_ar'][e] += _as_float(v.get('qm'))
                if prefix in codes['ca']:
                    tb_totals['_ent_ca'][e] += _as_float(v.get('qm'))
        tb_totals[prefix] = tot
    print(f'  TB 期末: 应收 {tb_totals.get(codes["ar"][0],0):,.2f}, '
          f'合同资产 {tb_totals.get(codes["ca"][0],0):,.2f}')

    out = os.path.join(os.path.dirname(AM.resolve('XBJ')), acct, '底稿', str(year))
    # ⚡ 子集团目录：非 _root 的按 底稿/{year}/{sub}/ 输出（避免同名覆盖）
    if sub != '_root':
        out = os.path.join(out, sub)
    os.makedirs(out, exist_ok=True)
    out_path = os.path.join(out, f'应收合同资产账龄复核_{year}.xlsx')
    _write_workbook(out_path, cfg, flows, tb_totals, '')
    print(f'  ✅ 账龄复核表已生成: {out_path}')
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('acct', nargs='?', default=None)
    ap.add_argument('sub', nargs='?', default=None)
    a = ap.parse_args(argv)
    rc = 0
    for acct, acct_cfg in AM.jobs().items():
        if a.acct and acct != a.acct:
            continue
        for sub in acct_cfg['subs']:
            if a.sub and sub != a.sub:
                continue
            rc |= run_one(acct, sub)
    return rc


if __name__ == '__main__':
    sys.exit(main())
