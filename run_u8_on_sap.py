# -*- coding: utf-8 -*-
"""run_u8_on_sap.py —— 治本架构驱动：用 U8 生成器 + sap_adapter 跑 SAP 账套（逐主体）。

用户原则（2026-08-09）：无论 SAP 还是 U8 只是数据来源不一样，形成底稿用同一套生成器。
SAP 88 家全量 GL 在内存会 OOM → 必须【逐主体】驱动：
    对每个主体构造 entities={comp:{...}} → 调 U8 生成器核心 build（只含该主体行）
    → 输出文件重命名为 {科目}审计底稿_{comp}.xlsx（对齐 SAP 交付命名）。

用法：
    python run_u8_on_sap.py            # 全量 88 家
    python run_u8_on_sap.py 1010 1030  # 指定主体
    python run_u8_on_sap.py --subj equity loan   # 指定科目
    python run_u8_on_sap.py --no-resume         # 强制全量（默认断点续跑，跳过签名未变的已完成科目）
"""
import paths as P
import os
import re
import sys
import time
import shutil
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

# ⚡⚡ 2026-08-27 P0 修复（1010 固定资产/费用/往来底稿全 SKIP 根因之一）：原默认
#   DATA = os.environ.get('AUDIT_DATA_ROOT') or P.YY = D:\底稿测试\AH（账套根，非数据目录）。
#   不传路径参数时 discover 扫 AH 根找不到 科目余额表/序时账（在 AH/数据/2026 下）→
#   全部科目 SKIP 静默失败。SAP 主账套数据目录固定 = P.YY/数据/2026；仍支持首参拖拽覆盖。
DATA = os.environ.get('AH_DATA_ROOT') or os.path.join(P.YY, '数据', '2026')
OUT = os.path.join(DATA, '底稿')


def _infer_year(data):
    """从数据目录路径推断年份，消除 '2026' 硬编码（跨账套/跨年不再改常量）。
    识别形态：`.../数据/2026`、`.../数据/2025/...`、`.../AH/2026`（路径中 4 位数字段）。
    ⚡⚡ 2026-08-30 修复：原 YEAR='2026' 写死 → XBJ(2025)/明年(2027) 必须改代码。
    拖拽传参改 DATA 后须重新推断（main 内同步）。"""
    m = re.search(r'[\\/]数据[\\/]?(\d{4})', data)
    if m:
        return m.group(1)
    m = re.findall(r'(\d{4})', data)
    return m[-1] if m else '2026'


YEAR = _infer_year(DATA)
# ⚡ 2026-08-10 修复：多进程并行（_sap_batch）时各家庭若都向 DATA 根目录写同名中间文件
#   （{科目}审计底稿_2026_生成.xlsx）→ 进程间互相覆盖 → 文件写坏（实测 196 个 BadZipFile）。
#   每家庭改用独立工作目录 _work_{comp}：生成器输出到该目录，再 move 到 OUT；读数据走
#   adapter（忽略 data_dir 参数，用 current_comp），传工作目录不影响取数。
_cur_work = None


def _work():
    """当前主体的独立输出工作目录（生成器输出落地处）。"""
    return _cur_work if _cur_work else DATA

import sap_adapter as A

# ⚡ 2026-08-10 集团模式（--group）：把数据根下全部主体视为同一集团，每科目生成一份
#   多主体列底稿（列=各主体、尾部集团合计/合并抵消/合并报表数 4 行，audit_common 多实体
#   分支自动启用；铁律81 单体裁剪在 len(ents)<=1 才触发，4 主体自然输出集团行）。
GROUP_MODE = False
# ⚡ 2026-08-13 回归集支持：--only 限定集团模式只跑指定主体（回归验证用，不全量 88 家）
GROUP_ONLY = None


def _comp_of(comp):
    """按主体构造 entities（SAP 主体结构 {comp: {2026: {km, gl}}}）。集团模式返回全部主体。"""
    full = A.discover_entities(DATA)
    if GROUP_MODE:
        if GROUP_ONLY:
            return {c: full[c] for c in GROUP_ONLY if c in full}
        return full
    if comp not in full:
        return None
    return {comp: full[comp]}


def _patch_local_readers(mod):
    """把生成器模块的本地读取引用替换为 backend 注入版（形态分发 DAO 单点）。
    2026-08-30 批次B：逻辑内聚到 ledger_backend.install_backend_readers（get_backend
    判定形态；SAP 系注入 sap_adapter 读取器 + _adapter，U8 无需注入），本函数仅转发。
    历史背景：生成器大多 `from audit_common import read_km as _read_km, read_gl as _read_gl`，
    内部聚合基于这两个函数 → patch 模块属性即覆盖（inventory 的 _read_km_cached 等包装自动生效）。
    ⚡ 2026-08-11 P0 修复：补设 mod._adapter = A —— 生成器 SAP 分支普遍判断
    `_adapter is not None and _adapter.is_sap(...)`（如 bank.read_gl_bank / revenue._aggregate /
    expense._read_gl_sap 等），但模块级 _adapter 恒为 None（只有 main() 才 import sap_adapter），
    导致 SAP 场景（单主体+集团）永远走 U8 分支读 list → 银行对方科目核对/销售采购/剩余核对、
    收入分月/分析性程序等全空。统一注入后 SAP 分支才真正生效。"""
    from ledger_backend import install_backend_readers
    return install_backend_readers(mod, DATA)


def _move_out(comp):
    """把生成器输出的 {科目}审计底稿_{year}_生成.xlsx 重命名为 {科目}审计底稿_{comp}.xlsx。
    ⚡ 2026-08-15 铁律130c（AH 集团底稿）：集团模式（GROUP_MODE）默认保留原名
    {科目}审计底稿_{year}_生成.xlsx（集团级每科目一份，与 U8 账套 ga/jj 形态一致）。
    ⚡⚡ 2026-08-24 补正（用户按三集团核对）：--group-name 传 4 位集团码（2468/1010/1357）
        时改名 _生成 → _{group-name}，使三集团输出命名统一可区分（原漏改名 17 个/集团）。"""
    if not os.path.isdir(OUT):
        os.makedirs(OUT, exist_ok=True)
    n = 0
    for fn in os.listdir(_work()):
        if fn.endswith(f'_{YEAR}_生成.xlsx') and '审计底稿' in fn:
            src = os.path.join(_work(), fn)
            if GROUP_MODE and re.fullmatch(r'\d{4}', str(comp)):
                # --group-name 为 4 位集团码 → 改名带集团标识（2468/1010/1357）
                base = fn.replace(f'_{YEAR}_生成.xlsx', '')
                dst = os.path.join(OUT, f'{base}_{comp}.xlsx')
            elif GROUP_MODE:
                dst = os.path.join(OUT, fn)
            else:
                base = fn.replace(f'_{YEAR}_生成.xlsx', '')
                dst = os.path.join(OUT, f'{base}_{comp}.xlsx')
            try:
                from audit_common import move_if_free
                if move_if_free(src, dst):
                    n += 1
            except Exception:
                pass
    return n


def _move_out_gen(comp, pattern):
    """通用改名：{前缀}审计底稿_生成.xlsx 或 {前缀}审计底稿_{y}_生成.xlsx → {前缀}审计底稿_{comp}.xlsx
    ⚡ 2026-08-28 修复：『关联交易和余额核对_生成.xlsx』不含"审计底稿"字样被跳过，
       正式目录关联交易核对因此长期停留在旧版——放宽条件纳入该前缀。"""
    if not os.path.isdir(OUT):
        os.makedirs(OUT, exist_ok=True)
    n = 0
    for fn in os.listdir(_work()):
        if fn.endswith('_生成.xlsx') and ('审计底稿' in fn or fn.startswith('关联交易和余额核对')) and fn.count('_') <= 3:
            src = os.path.join(_work(), fn)
            base = fn.replace('_生成.xlsx', '')
            # base 形如：{科目}审计底稿_{year} → 去掉年份
            if base.endswith(f'_{YEAR}'):
                base = base[:-(len(YEAR) + 1)]
            dst = os.path.join(OUT, f'{base}_{comp}.xlsx')
            try:
                from audit_common import move_if_free
                if move_if_free(src, dst):
                    n += 1
            except Exception:
                pass
    return n


# ============================================================
# 各科目驱动（逐主体调 U8 生成器核心 build，输出重命名）
# ============================================================
def run_equity(comp):
    import equity_detail as EQ
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    tb = A.read_tb_full(DATA, ents1)
    gl = A.read_gl_rows(DATA, ents1)
    n = 0
    for gkey in EQ.GROUPS:
        try:
            EQ._build_year_workbook(_work(), ents1, tb, gl, YEAR, gkey)
            n += 1
        except Exception:
            traceback.print_exc()
            print(f'  ⚠️ equity {gkey} 失败')
    _move_out(comp)
    return n


def run_loan(comp):
    import loan_detail as LN
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    # ⚡ 2026-08-28 修复：run_loan 不走 process_folder（无 set_fin_data_root），
    #   融资台账默认读 ADF → AH 短期借款底稿混入 ADF 借款数据。显式按当前项目设置。
    LN.set_fin_data_root(DATA)
    tb = A.read_tb_full(DATA, ents1)
    gl = A.read_gl_rows(DATA, ents1)
    aux_map = LN._discover_aux(DATA, ents1)
    model = {}
    for e in sorted(ents1):
        for y in sorted(ents1[e]):
            aux = aux_map.get((e, y))
            for sk, subj in LN.SUBJECTS.items():
                rec = LN._extract_loan_detail(aux, tb, e, y, subj)
                if rec:
                    rec['_sk'] = sk
                    model[(e, y, sk)] = rec
    if not model:
        return 0
    LN._build_split_workbooks(_work(), model, [YEAR], gl, tb, ents1, None)
    _move_out(comp)
    return 1


def run_rd_expense(comp):
    import rd_expense_detail as RD
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    tb = A.read_tb_full(DATA, ents1)
    gl = A.read_gl_rows(DATA, ents1)
    detail, control, skipped = RD.build_detail(ents1, tb)
    all_years = sorted({y for (e, y) in control} | {y for (e, y) in skipped})
    if not all_years:
        return 0
    RD.write_paper(_work(), detail, control, skipped, years=all_years, tb=tb,
                   entities=ents1, gl=gl)
    _move_out(comp)
    return 1


def run_pl(comp):
    import pl_detail as PL
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    # pl 走 audit_common read_gl_rows/read_tb_full（patch 覆盖）+ 本地 read_pl_gl 等
    # 本地单文件读取在 SAP 下由 _read_km/_read_gl patch 兜底（若 pl 用了）
    _patch_local_readers(PL)
    PL.build_pl_combined(_work())
    _move_out_gen(comp, None)
    return 1


def run_revenue(comp):
    import revenue_detail as RV
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    _patch_local_readers(RV)
    RV.build_revenue_workbook(_work())
    _move_out_gen(comp, None)
    return 1


def run_inventory(comp):
    import inventory_detail as IV
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    _patch_local_readers(IV)
    IV.process_folder(_work())
    _move_out_gen(comp, None)
    return 1


def run_longterm(comp):
    import longterm_assets_detail as LT
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    _patch_local_readers(LT)
    LT.process_folder(_work())
    _move_out_gen(comp, None)
    return 1


def run_gp_other(comp):
    import gp_other as GO
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    GO.build_all(_work())
    _move_out_gen(comp, None)
    return 1


def run_expense(comp):
    import expense_detail as EX
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    _patch_local_readers(EX)
    EX.build_expense_combined(_work())
    _move_out_gen(comp, None)
    return 1


def run_tax(comp):
    import tax_detail as TX
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    TX.build_tax_combined(_work())
    _move_out_gen(comp, None)
    return 1


def run_payroll(comp):
    import payroll_detail as PY
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    PY.build_payroll_combined(_work())
    _move_out_gen(comp, None)
    return 1


def run_bank(comp):
    import bank_deposit_detail as BK
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    # ⚡ 2026-08-10 并行修复：bank 内部硬编码写『{dir}/银行存款审计底稿_{y}_生成.xlsx』（文件名不含主体），
    # 多进程并行每家写同名 → 互相覆盖写坏。改：每家独立 work 子目录 → os.replace 原子移入 OUT。
    wdir = os.path.join(DATA, f'_work_bank_{comp}')
    os.makedirs(wdir, exist_ok=True)
    BK.build_bank_detail_km_combined(DATA, os.path.join(wdir, 'bank.xlsx'))
    src = os.path.join(wdir, f'银行存款审计底稿_{YEAR}_生成.xlsx')
    dst = os.path.join(OUT, f'银行存款审计底稿_{comp}.xlsx')
    if os.path.exists(src):
        os.makedirs(OUT, exist_ok=True)
        from audit_common import move_if_free
        move_if_free(src, dst)
    return 1


def run_current_account(comp):
    import current_account_detail as CA
    ents1 = _comp_of(comp)
    if not ents1:
        return 0
    # ⚡ 2026-08-11 修复：锁定单主体——前面 DRIVER（bank set_comp 曾硬置 None）可能污染
    #   current_comp → _discover_entities 误走全量 88 家 → 40 分钟 + 集团口径错。显式重设。
    if not GROUP_MODE:
        A.set_comp(comp)
    CA.build_all_in_dir(_work(), out_dir=_work())
    # 往来款输出 {label}审计底稿_2026_生成.xlsx → {label}审计底稿_{comp}.xlsx
    # ⚡⚡ 2026-08-23 修复：移动前确保 OUT 目录存在——os.replace 目标目录缺失会静默失败
    #   （此前文件留在 _work_{comp} 未移入底稿目录）。
    os.makedirs(OUT, exist_ok=True)
    for fn in os.listdir(_work()):
        if fn.endswith(f'_{YEAR}_生成.xlsx') and '审计底稿' in fn:
            base = fn.replace(f'_{YEAR}_生成.xlsx', '')
            src = os.path.join(_work(), fn)
            dst = os.path.join(OUT, f'{base}_{comp}.xlsx')
            try:
                from audit_common import move_if_free
                move_if_free(src, dst)
            except Exception:
                pass
    return 1


# 科目驱动注册表（13 科目全接入，2026-08-09）
DRIVERS = {
    'equity': run_equity,
    'loan': run_loan,
    'rd_expense': run_rd_expense,
    'pl': run_pl,
    'revenue': run_revenue,
    'inventory': run_inventory,
    'longterm': run_longterm,
    'gp_other': run_gp_other,
    'expense': run_expense,
    'tax': run_tax,
    'payroll': run_payroll,
    'bank': run_bank,
    'current_account': run_current_account,
}


def _snapshot_out_files():
    """三态判定的输出快照：_cur_work 工作目录 + 底稿目录 的 xlsx {abspath: (size, mtime)}。
    ⚡ 2026-08-13 修复：原只记文件集（有无）→ 重跑已存在文件（覆盖写）集合不变 → 误判 SKIP；
    改为记录大小+修改时间，跑后有「新增 或 更新」即 OK。"""
    import glob as _g
    base = {}
    for _d in (_cur_work, OUT):
        if os.path.isdir(_d):
            for _fp in _g.glob(os.path.join(_d, '*.xlsx')):
                try:
                    _st = os.stat(_fp)
                    base[os.path.abspath(_fp)] = (_st.st_size, _st.st_mtime)
                except Exception:
                    pass
    return base


# ⚡⚡ 2026-08-29 断点续跑（用户诉求：大任务崩溃续跑增强）：
#   每个 (主体, 科目) 完成后写断点 .resume_{comp}.jsonl，记录状态 + 数据签名 + 代码签名。
#   --resume（默认开）跳过「已完成 OK/SKIP 且签名未变」的科目；FAIL 不跳过（续跑重试，
#   连续 3 次 FAIL 标记 FAILED_PERM 跳过并告警）；--no-resume 强制全量。
#   签名：代码签名=小程序全部 .py 聚合；数据签名=主体 km+gl 文件聚合。任一变化 → 全部失效重跑
#   （改代码/重导数据后必须全量，这是正确行为）。三集团并行各写各的 .resume_{group} 无冲突。
_code_sig_cache = None


def _code_sig():
    """小程序目录全部 .py 的 (size, mtime) 聚合哈希（改代码 → 变化 → 断点失效）。"""
    global _code_sig_cache
    if _code_sig_cache is not None:
        return _code_sig_cache
    import hashlib
    h = hashlib.sha1()
    try:
        for fn in sorted(os.listdir(HERE)):
            if not fn.endswith('.py'):
                continue
            fp = os.path.join(HERE, fn)
            try:
                st = os.stat(fp)
                h.update(f'{fn}|{st.st_size}|{st.st_mtime:.3f}|'.encode('utf-8', 'replace'))
            except Exception:
                pass
    except Exception:
        pass
    _code_sig_cache = h.hexdigest()[:16]
    return _code_sig_cache


def _data_sig(ents):
    """主体数据签名：km + 全部 gl 文件的 (size, mtime) 聚合（重导数据 → 变化 → 断点失效）。"""
    import hashlib
    h = hashlib.sha1()
    seen = set()
    try:
        for _e in (ents or {}).values():
            for _y in (_e or {}).values():
                _km = _y.get('km') if isinstance(_y, dict) else None
                if _km:
                    seen.add(os.path.abspath(_km))
                for _gl in (_y.get('gl') or []) if isinstance(_y, dict) else []:
                    if _gl:
                        seen.add(os.path.abspath(_gl))
    except Exception:
        pass
    for fp in sorted(seen):
        try:
            st = os.stat(fp)
            h.update(f'{os.path.basename(fp)}|{st.st_size}|{st.st_mtime:.3f}|'.encode('utf-8', 'replace'))
        except Exception:
            pass
    return h.hexdigest()[:16]


def _resume_file(comp):
    return os.path.join(DATA, f'.resume_{comp}.jsonl')


def _resume_load(comp):
    """读断点 → {subj: {status, ts, csig, dsig, tries}}。文件损坏则返回空（安全兜底）。"""
    import json as _json
    out = {}
    fp = _resume_file(comp)
    if not os.path.exists(fp):
        return out
    try:
        with open(fp, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = _json.loads(line)
                    out[rec['subj']] = rec
                except Exception:
                    continue
    except Exception:
        return {}
    return out


def _resume_mark(comp, subj, status, csig, dsig):
    """写入/更新单科目断点（整文件重写保原子性，状态小无性能问题）。"""
    import json as _json
    recs = _resume_load(comp)
    prev = recs.get(subj, {})
    tries = prev.get('tries', 0)
    if status == 'FAIL':
        tries += 1
        # 连续 3 次 FAIL → 永久跳过（防无限重试死循环），否则保留 FAIL 待续跑重试
        if tries >= 3:
            status = 'FAILED_PERM'
    recs[subj] = {'subj': subj, 'status': status, 'ts': time.strftime('%Y-%m-%d %H:%M:%S'),
                  'csig': csig, 'dsig': dsig, 'tries': tries,
                  'host': __import__('socket').gethostname()}   # ⚡ 2026-08-30 多用户留痕：记录生成机器（不影响跨机续跑，签名校验兜底）
    fp = _resume_file(comp)
    try:
        tmp = fp + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            for s in sorted(recs):
                f.write(_json.dumps(recs[s], ensure_ascii=False) + '\n')
        os.replace(tmp, fp)
    except Exception:
        pass


def _heartbeat_file(comp):
    return os.path.join(DATA, f'.heartbeat_{comp}')


def _touch_heartbeat(comp, name):
    """每科目开始/完成写心跳文件（看门狗检测停滞超时 → 卡死告警）。"""
    try:
        with open(_heartbeat_file(comp), 'w', encoding='utf-8') as f:
            f.write(f'{time.strftime("%Y-%m-%d %H:%M:%S")} {comp} {name}\n')
    except Exception:
        pass


def _git_head():
    """小程序目录当前 git commit（失败返回 None，用于运行报告归档）。"""
    try:
        import subprocess as _sp
        return _sp.check_output(['git', '-C', HERE, 'rev-parse', '--short', 'HEAD'],
                                stderr=_sp.DEVNULL).decode().strip()[:12]
    except Exception:
        return None


def _write_run_report(argv, run_log, stats, skips, failed_details):
    """运行报告落盘（tmp/run_history/run_{ts}.md）：参数归档 + 明细 + 失败 traceback。
    ⚡ 2026-08-29 用户诉求：每次运行可回溯「哪批代码+哪批数据+结果如何」。"""
    try:
        d = os.path.join(HERE, 'run_history')
        os.makedirs(d, exist_ok=True)
        ts = time.strftime('%Y%m%d_%H%M%S')
        fp = os.path.join(d, f'run_{ts}.md')
        L = []
        L.append('# 运行报告')
        L.append(f'- 时间：{time.strftime("%Y-%m-%d %H:%M:%S")}')
        L.append(f'- 数据根：{DATA}')
        L.append(f'- 输出：{OUT}')
        L.append(f'- 参数：{" ".join(argv)}')
        L.append(f'- git commit：{_git_head() or "（非 git）"}')
        L.append(f'- 代码签名：{_code_sig()}')
        L.append(f'- 结果：OK {stats["OK"]} / SKIP {stats["SKIP"]} / FAIL {stats["FAIL"]}'
                 f'（断点续跑跳过 {sum(1 for x in run_log if x[2] == "RESUME")}）')
        L.append('')
        L.append('## 明细')
        L.append('| 主体 | 科目 | 状态 | 时间 | 耗时s |')
        L.append('|---|---|---|---|---|')
        for comp, name, st, ts_, dur in run_log:
            L.append(f'| {comp} | {name} | {st} | {ts_} | {dur} |')
        if failed_details:
            L.append('')
            L.append('## 失败明细')
            L.append('')
            for comp, name, tb_ in failed_details:
                L.append(f'### {comp} / {name}')
                L.append('```')
                L.append(tb_[-1200:])
                L.append('```')
        if skips:
            L.append('')
            L.append('## SKIP 明细（需人工确认是否真无数据）')
            L.append('')
            for s in skips[:100]:
                L.append(f'- {s}')
        with open(fp, 'w', encoding='utf-8') as f:
            f.write('\n'.join(L))
        print(f'  📋 运行报告已写：{fp}', flush=True)
    except Exception as _ex:
        print(f'  ⚠️ 运行报告写入失败：{_ex}', flush=True)


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    global GROUP_MODE, GROUP_ONLY, _cur_work, DATA, OUT, YEAR
    # ⚡ 2026-08-11 拖拽支持：首参为文件夹路径（不以 '-' 开头）→ 视为数据根目录。
    #   拖入 yy/300 文件夹即可直接生成（与 current_account_detail 的拖入快捷方式对齐）。
    # ⚡ 2026-08-12 修复（落地错目录）：OUT 模块级由 DATA 导入时计算（默认 yy/底稿），
    #   拖拽/传参改 DATA 后若不同步 OUT，_move_out 会把文件移到旧 DATA 的底稿目录
    #   （3200/3500 全部落到 yy/底稿 的错账套事故）。此处强制同步。
    if argv and not argv[0].startswith('-') and os.path.isdir(argv[0]):
        DATA = os.path.abspath(argv[0])
        OUT = os.path.join(DATA, '底稿')
        YEAR = _infer_year(DATA)   # ⚡⚡ 2026-08-30 拖拽改 DATA 后同步年份
        print(f'[拖拽模式] 数据根目录：{DATA}（年份 {YEAR}）', flush=True)
        argv = argv[1:]
    subj_only = None
    if '--subj' in argv:
        i = argv.index('--subj')
        subj_only = argv[i + 1].split(',')
        argv = argv[:i] + argv[i + 2:]
    # ⚡ 2026-08-13 回归集支持：--only 指定主体（集团模式 _comp_of 过滤）
    if '--only' in argv:
        i = argv.index('--only')
        GROUP_ONLY = argv[i + 1].split(',')
        argv = argv[:i] + argv[i + 2:]
        # ⚡⚡ 2026-08-16 集团主体过滤（#744 对平抽查根因）：注入 adapter 全局——
        #   revenue/current_account 等生成器内部 discover 全量 88 家，靠此一处收口
        A._GROUP_COMPS = set(GROUP_ONLY)
        print(f'[--only] 仅限主体：{GROUP_ONLY}', flush=True)
    # ⚡⚡ 2026-08-15 铁律130c（AH 3 集团每科目集团底稿）：--group-name 自定义集团输出名
    #   （输出 {科目}审计底稿_{YEAR}_生成.xlsx 到 --out-dir，与试算表/合并稿同目录、按集团隔离）
    _group_name = None
    if '--group-name' in argv:
        i = argv.index('--group-name')
        _group_name = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    _out_dir = None
    if '--out-dir' in argv:
        i = argv.index('--out-dir')
        _out_dir = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    group_mode = '--group' in argv
    if group_mode:
        argv.remove('--group')
    # ⚡ 2026-08-29 断点续跑：--no-resume 强制全量；默认 --resume（跳过签名未变的已完成科目）
    resume = '--no-resume' not in argv
    if '--no-resume' in argv:
        argv.remove('--no-resume')
    comps = [a for a in argv if not a.startswith('-')] or sorted(A.discover_entities(DATA).keys())
    A.set_root(DATA)
    A.patch_audit_common()
    # ⚡⚡ 2026-08-30 防误用（曾把集团码 1357 当主体码传入 → current_account 静默 SKIP 0.0s）：
    #   --only 指定的代码不在 discover 主体集 → 立即告警并剔除，不让静默空结果交付。
    #   AH 三集团（1010/1357/2468）是集团码非主体码，须走 ah_parallel_run（内部按集团
    #   拆真实主体列表传 --only），此处命中即提示正确用法。
    if GROUP_ONLY:
        try:
            _dkeys = set(A.discover_entities(DATA).keys())
            _miss = [c for c in GROUP_ONLY if c not in _dkeys]
            if _miss:
                _hit = [c for c in GROUP_ONLY if c in _dkeys]
                print(f'⚠️ [--only] {len(_miss)} 个非主体代码（可能误用集团码）: {_miss}', flush=True)
                if not _hit:
                    print(f'⚠️ [--only] 全部无效 → 无主体可跑，请核对参数。AH 集团请用 ah_parallel_run。', flush=True)
                GROUP_ONLY = _hit
                A._GROUP_COMPS = set(GROUP_ONLY)
        except Exception:
            pass
    # ⚡⚡ 2026-08-30 大集团自动 xw 化（88 家合并 current_account openpyxl 卡死教训）：
    #   group 模式主体数超 40 → 强制 CA_XW=1。xw 只影响 current_account 分支，
    #   其他模块无 CA_XW 分支、行为不变 → 自动启用安全。
    if group_mode:
        try:
            _nent = len(A.discover_entities(DATA))
            if _nent > 40 and not os.environ.get('CA_XW'):
                os.environ['CA_XW'] = '1'
                print(f'⚡ [{_nent} 个主体] 大集团自动启用 xw 渲染（current_account），'
                      f'避免 openpyxl 88 家大表卡死', flush=True)
        except Exception:
            pass
    if _out_dir:
        OUT = os.path.abspath(_out_dir)
        print(f'[--out-dir] 输出目录：{OUT}', flush=True)
    # ⚡ 2026-08-13 三态返回（用户方法论：防「驱动成功即 ok」吞掉静默缺失——88/88 假收口根子）：
    #   OK=跑后输出目录有新增文件；SKIP=无产出且无异常（TB 无数据正常跳过）；FAIL=抛异常。
    stats = {'OK': 0, 'SKIP': 0, 'FAIL': 0}
    skips = []
    run_log = []            # ⚡ 2026-08-29 运行报告：每科目 (comp, name, status, ts, dur)
    failed_details = []     # 失败科目 traceback 归档
    if group_mode:
        # ⚡ 2026-08-10 集团模式：不 set_comp（current_comp=None → discover 返回全部主体），
        #   以『{根目录名}集团』伪主体单次循环，输出 {科目}审计底稿_{根目录名}集团.xlsx
        #   （2026-08-10 修复：原写死『300集团』→ yy 的集团底稿也沿用 _300集团 命名误导；
        #   现由数据根目录名派生：300→300集团、yy→yy集团，与账套身份一致）
        GROUP_MODE = True
        comps = [_group_name or (os.path.basename(DATA.rstrip('/\\')) + '集团')]
    for comp in comps:
        if not GROUP_MODE:
            A.set_comp(comp)
        # ⚡ 2026-08-10：每家庭独立输出工作目录（防多进程并行同名中间文件互相覆盖写坏）
        _cur_work = os.path.join(DATA, f'_work_{comp}')
        os.makedirs(_cur_work, exist_ok=True)
        print(f'=== [{comp}] ===', flush=True)
        # ⚡ 2026-08-29 断点续跑：算签名 + 加载断点（本轮签名的锚点，用于跳过判断）
        _csig = _code_sig()
        _ents = _comp_of(comp) or {}
        _dsig = _data_sig(_ents)
        _resume = _resume_load(comp)
        _n_resume = 0
        for name, fn in DRIVERS.items():
            if subj_only and name not in subj_only:
                continue
            # ⚡ 2026-08-29 断点续跑：签名一致且 OK/SKIP/FAILED_PERM 的科目跳过
            # ⚡⚡ 2026-08-30 改进（bank 重跑被旧 SKIP 断点拦截的教训）：--subj 明确指定
            #   某科目时【忽略断点跳过】——用户指定=要重跑（补跑失败/验证修复），
            #   断点残留（如早期 SKIP）不应阻止。
            if resume and not subj_only:
                _rc = _resume.get(name)
                if _rc and _rc.get('csig') == _csig and _rc.get('dsig') == _dsig:
                    if _rc.get('status') in ('OK', 'SKIP'):
                        _n_resume += 1
                        run_log.append((comp, name, 'RESUME', time.strftime('%H:%M:%S'), 0))
                        print(f'  ⏭ [{time.strftime("%H:%M:%S")}] {name}: 跳过（已 {_rc["status"]}，{_rc.get("ts")}）', flush=True)
                        continue
                    if _rc.get('status') == 'FAILED_PERM':
                        run_log.append((comp, name, 'RESUME', time.strftime('%H:%M:%S'), 0))
                        print(f'  ⛔ {name}: 跳过（连续失败 {_rc.get("tries")} 次，需人工排查后 --no-resume 重试）', flush=True)
                        continue
            # ⚡⚡ 2026-08-16 铁律132（看门狗配套）：科目开始打印+时间戳——卡死时能定位到
            #   具体 (主体,科目)，而非仅"完成后"日志（此前卡在科目中间无任何输出）
            print(f'  ⏳ [{time.strftime("%H:%M:%S")}] {name} 开始…', flush=True)
            _touch_heartbeat(comp, name)
            _t0 = time.time()
            _before = _snapshot_out_files()
            try:
                n = fn(comp)
                _after = _snapshot_out_files()
                _n_add = sum(1 for f in _after if f not in _before)
                _n_upd = sum(1 for f in _after if f in _before and _after[f] != _before[f])
                if _n_add or _n_upd:
                    stats['OK'] += 1
                    run_log.append((comp, name, 'OK', time.strftime('%H:%M:%S'), round(time.time() - _t0, 1)))
                    print(f'  ✅ {name}: {n}（新增 {_n_add} / 更新 {_n_upd} 文件，{time.time()-_t0:.0f}s）', flush=True)
                    _resume_mark(comp, name, 'OK', _csig, _dsig)
                else:
                    stats['SKIP'] += 1
                    skips.append(f'{comp} {name}')
                    run_log.append((comp, name, 'SKIP', time.strftime('%H:%M:%S'), round(time.time() - _t0, 1)))
                    print(f'  ⚪ {name}: SKIP 无产出（TB 无该组数据？）', flush=True)
                    _resume_mark(comp, name, 'SKIP', _csig, _dsig)
            except Exception:
                stats['FAIL'] += 1
                tb = traceback.format_exc()
                run_log.append((comp, name, 'FAIL', time.strftime('%H:%M:%S'), round(time.time() - _t0, 1)))
                failed_details.append((comp, name, tb))
                traceback.print_exc()
                print(f'  ❌ {name} 失败', flush=True)
                _resume_mark(comp, name, 'FAIL', _csig, _dsig)
            _touch_heartbeat(comp, name)
        if _n_resume:
            print(f'  ↪ [{comp}] 断点续跑跳过 {_n_resume} 个已完成科目（--no-resume 强制全量）', flush=True)
    print(f'完成：OK {stats["OK"]} / SKIP {stats["SKIP"]} / FAIL {stats["FAIL"]}', flush=True)
    _write_run_report(sys.argv[1:], run_log, stats, skips, failed_details)
    if skips:
        print(f'SKIP 明细（{len(skips)} 项，需人工确认是否真无数据）：', flush=True)
        for s in skips[:80]:
            print(f'  ⚪ {s}', flush=True)
        if len(skips) > 80:
            print(f'  … 其余 {len(skips) - 80} 项', flush=True)
    # ⚡ 2026-08-11 自动质检（用户核心诉求落地：生成后必须自检，不能靠人工翻底稿发现）：
    #   ①空表扫描（该有数据的表空=ERROR，TB 判定）；②audit_checker（结构规范）。
    #   有 ERROR 时打印醒目告警并写 gap_report.txt；--skip-check 可跳过。
    if '--skip-check' not in argv:
        try:
            import sheet_gap_check
            # ⚡⚡ 2026-08-29 P0 修复（用户痛点"空白底稿要逐张点开检查"根因）：
            #   原 scan_folder_with_tb(DATA) 扫 DATA/底稿（旧单体稿目录），集团模式实际
            #   输出在 --out-dir → 集团底稿从未被空表扫描覆盖。改扫实际 OUT。
            _errs, _issues = sheet_gap_check.scan_folder_with_tb(DATA, folder=OUT)
            print(f'\n═════ 自动质检（扫 {OUT}）═════')
            print(f'  空表扫描: {_errs} ERROR / {len(_issues)} 项')
            _n_err = 0
            _rep = []
            for _l, _m in _issues:
                if _l == 'ERROR':
                    _n_err += 1
                    print(f'    ❌ {_m[:110]}')
                    _rep.append(f'[{_l}] {_m}')
            if _n_err == 0:
                print('  ✅ 关键表全部有数据（无取数丢失）')
            try:
                import audit_checker
                _te, _tw = audit_checker.process_folder(OUT if os.path.isdir(OUT) else DATA, quiet=True)
                print(f'  audit_checker: {_te} ERROR / {_tw} WARN')
            except Exception as _ex:
                print(f'  ⚠️ audit_checker 异常: {_ex}')
            if _rep:
                _rp = os.path.join(DATA, '底稿质检报告.txt')
                try:
                    with open(_rp, 'w', encoding='utf-8') as _f:
                        _f.write(f'自动质检 {time.strftime("%Y-%m-%d %H:%M")} — {_n_err} ERROR\n')
                        _f.write('\n'.join(_rep))
                    print(f'  报告已写: {_rp}')
                except Exception:
                    pass
            # ⚡ 2026-08-13 指纹制度化（用户方法论：把「改完一遍遍人工翻查」变成「自动对比」）：
            #   收口重跑后自动对比上次快照（消失/大小变化=回归风险），并滚动本次为下次基准
            try:
                import _snapshot_yy as _S
                _S.DIR = os.path.join(DATA, '底稿')
                _prev = os.path.join(DATA, '_snapshot_prev.json')
                _curr = os.path.join(DATA, '_snapshot_curr.json')
                if os.path.exists(_prev):
                    _S.snap(_curr)
                    _bad = _S.diff(_prev, _curr, [])
                    print(f'  [指纹] 对比完成：{"✅ 无回归" if _bad == 0 else f"❌ {_bad} 处变化需检查"}')
                    try:
                        os.replace(_curr, _prev)
                    except Exception:
                        pass
                else:
                    _S.snap(_prev)
                    print('  [指纹] 首次收口：已拍基准快照 _snapshot_prev.json（下次重跑自动对比）')
            except Exception as _ex:
                print(f'  ⚠️ 指纹对比失败: {_ex}')
            # ⚡⚡ 2026-08-24（用户需求）：集团模式生成后【自动跑阶段③同源核对】
            #   —— 集团审定数 vs 自建试算表，0 差异才算集团底稿取数无 bug（不再人工翻查）。
            if group_mode and _out_dir:
                try:
                    import recon_self
                    _ok, _diffs = recon_self.recon_self(DATA, wp_dir=_out_dir)
                    if _ok and not _diffs:
                        print('  [阶段③集团核对] ✅ 底稿审定数=自建试算表（0 差异）')
                    else:
                        print(f'  [阶段③集团核对] ⚠️ {len(_diffs)} 处差异（需人工确认）')
                except Exception as _ex:
                    print(f'  ⚠️ 阶段③集团核对失败: {_ex}')
                # ⚡⚡ 2026-08-24 架构自检：空白底稿检测（集团模式，自动接入）
                try:
                    import self_check
                    _sok, _slines = self_check.self_check(DATA, group_dir=_out_dir, quiet=True)
                    _sbug = [l for l in _slines if '❌' in l]
                    if _sbug:
                        print(f'  ⛔ [空白底稿自检] {len(_sbug)} 处空白底稿（TB 有数据，需人工处理）：')
                        for l in _sbug[:8]:
                            print('      ' + l.strip())
                    else:
                        print('  ✅ [空白底稿自检] 无空白底稿')
                except Exception as _ex:
                    print(f'  ⚠️ 空白底稿自检失败: {_ex}')
        except Exception as _ex:
            print(f'  ⚠️ 自动质检失败: {_ex}')
    return 0


if __name__ == '__main__':
    main()
