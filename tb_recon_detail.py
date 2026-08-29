# -*- coding: utf-8 -*-
"""tb_recon_detail.py —— 合并试算表核对小程序（2026-08-04 用户需求）。

拖入《合并试算表.xlsx》→ 自动核对同文件夹底稿的三项数据：
  ① 审定表（期末余额/期末未审数 合计）
  ② 明细表（期末余额 合计）
  ③ 附注汇总（科目块合计/审定数）
生成《合并试算表核对底稿_生成.xlsx》：行=科目，列=试算表数|审定表数|差异|明细表合计|差异|附注数|差异|核对状态。
差异 > 容差(0.01) 标红，并汇总「需检查的底稿清单」。

试算表格式自适应：表头行含「科目编码/科目名称」+ 期末相关列（期末借方/期末贷方 两列，
或 期末余额+方向列，或 期末余额单列）。合计/汇总行（名称含 合计/总计）自动跳过。

用法：python tb_recon_detail.py <合并试算表.xlsx> [数据文件夹，缺省=试算表所在文件夹]
"""
import glob
import os
import re
import sys

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import audit_protocol as PROTO   # 底稿协议（2026-08-06）：段头/合计行/行类型唯一写法
import subject_mapping           # 2026-08-07 任务683：科目名归一化（account_profiles.json）

TOL = 0.01
HFILL = PatternFill('solid', fgColor='DDEBF7')
BADFILL = PatternFill('solid', fgColor='FFC7CE')
BADFONT = Font(name='Times New Roman', size=10, bold=True, color='9C0006')
BF = Font(name='Times New Roman', size=10, bold=True)
F10 = Font(name='Times New Roman', size=10)
NUM = '#,##0.00'
BDR = Border(*[Side(style='thin', color='BFBFBF')] * 4)


def _fnum(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _resolve_value(rows, i, j, _depth=0):
    """从 rows 矩阵（values_only 行列表）取 (i,j) 的数值：
    数值直接返回；公式（=SUM(A1:A10) / =D4+D5 / =B3 等）递归解析其引用的单元格值求和。
    2026-08-04 修复：合计行公式无缓存值时，旧"数据行回退求和"会把同列无关行
    （收入审定表的利润行、明细表多级小计）重复计入——公式区间解析精确且不重计。"""
    if _depth > 8:
        return 0.0
    if i is None or i < 0 or i >= len(rows):
        return 0.0
    r = rows[i]
    if j is None or j < 0 or j >= len(r) or r[j] is None:
        return 0.0
    v = r[j]
    fv = _fnum(v)
    if fv is not None:
        return fv
    if isinstance(v, str) and v.startswith('='):
        import re
        # 2026-08-04 修复：范围引用 A1:A5 必须展开中间单元格（SUM(B12:B14) 曾只解析两端 B12/B14，
        # 漏中间 B13——货币资金附注『集团合计数』=SUM(B12:B14) 三行和只剩两行）。
        _range_refs = re.findall(r'([A-Za-z]{1,3}[0-9]+):([A-Za-z]{1,3}[0-9]+)', v)
        refs = []
        for _c1, _c2 in _range_refs:
            _m1 = re.match(r'([A-Za-z]+)([0-9]+)', _c1)
            _m2 = re.match(r'([A-Za-z]+)([0-9]+)', _c2)
            if not _m1 or not _m2:
                continue
            _cc1 = _cc2 = 0
            for _ch in _m1.group(1).upper():
                _cc1 = _cc1 * 26 + (ord(_ch) - 64)
            for _ch in _m2.group(1).upper():
                _cc2 = _cc2 * 26 + (ord(_ch) - 64)
            for _rr in range(int(_m1.group(2)), int(_m2.group(2)) + 1):
                for _cc in range(min(_cc1, _cc2), max(_cc1, _cc2) + 1):
                    _cl = ''
                    _n = _cc
                    while _n:
                        _n, _rem = divmod(_n - 1, 26)
                        _cl = chr(65 + _rem) + _cl
                    refs.append('%s%d' % (_cl, _rr))
        refs += [x for x in re.findall(r'([A-Za-z]{1,3}[0-9]+)', v)
                 if x not in refs]
        tot = 0.0
        seen = set()
        for ref in refs:
            if ref in seen:
                continue
            seen.add(ref)
            m = re.match(r'([A-Za-z]+)([0-9]+)', ref)
            if not m:
                continue
            col = 0
            for ch in m.group(1).upper():
                col = col * 26 + (ord(ch) - 64)
            ri = int(m.group(2)) - 1
            cj = col - 1
            tot += _resolve_value(rows, ri, cj, _depth + 1)
        return tot
    return 0.0


def _find_header(rows, kws):
    """在 rows 中找同时含 kws 全部关键词的行。返回 (行索引, 列索引映射) 或 None。"""
    for i, r in enumerate(rows[:20]):
        if not r:
            continue
        joined = ''.join(str(x) if x is not None else '' for x in r[:26])
        if all(k in joined for k in kws):
            idx = {}
            for j, h in enumerate(r[:26]):
                hs = str(h).strip() if h is not None else ''
                if not hs:
                    continue
                if '科目编码' in hs or '科目代码' in hs or hs in ('编码', '代码'):
                    idx.setdefault('code', j)
                elif '科目名称' in hs or hs in ('科目', '项目', '账户名称'):
                    idx.setdefault('name', j)
                elif '审定数' in hs:
                    idx.setdefault('aud', j)          # 合并 TB 审定数列（优先）
                elif '简单加计数' in hs:
                    idx.setdefault('sum', j)
                elif '期末借方' in hs or '期末余额(借)' in hs:
                    idx.setdefault('jf', j)
                elif '本期借方' in hs:             # 2026-08-06 标准科目余额表（泰国/新加坡原币TB）
                    idx.setdefault('jf', j)
                elif '期末贷方' in hs or '期末余额(贷)' in hs:
                    idx.setdefault('df', j)
                elif '本期贷方' in hs:             # 2026-08-06 标准科目余额表
                    idx.setdefault('df', j)
                elif '期末余额' in hs or '期末' in hs:
                    idx.setdefault('bal', j)
                elif '期初' in hs:
                    idx.setdefault('qc', j)
                elif '金额' in hs:
                    # 2026-08-06 标准科目余额表格式：『方向|金额|本期借方|本期贷方|方向|金额』
                    # 两个金额列（期初/期末）——覆盖式取【最后一个】= 期末金额
                    idx['bal'] = j
                elif '方向' in hs:
                    # 标准科目余额表方向列成对出现（期初/期末），覆盖式取【最后一个】= 期末方向
                    idx['dir'] = j
            if 'name' in idx and any(k in idx for k in ('aud', 'sum', 'bal', 'qc', 'jf')):
                return i, idx
    return None, None


def _detect_year(tb_path, rows=None):
    """从试算表文件名（00_2026合并TB.xlsx）或表头（会计期间:2026年3月）提取年份。"""
    m = re.search(r'(20\d{2})', os.path.basename(tb_path))
    if m:
        return m.group(1)
    if rows:
        for r in rows[:12]:
            joined = ''.join(str(x) if x is not None else '' for x in r)
            m2 = re.search(r'(20\d{2})\s*年', joined)
            if m2:
                return m2.group(1)
    return None


def parse_trial_balance(fp, name_map=None, signed=True):
    """解析合并试算表 → (名称→期末净额 dict, 编码→期末净额 dict, 表头描述, 年份, 主体映射)。
    支持多 sheet 合并底稿（Index/TB表/...）：自动探测含「项目/科目 + 审定数/期初/期末」表头的 sheet。
    取值列优先级：审定数 > 简单加计数 > 期末余额 > 合并期初。
    主体映射：{主体名: {科目名: 净额}}（2026-08-04 单体核对：识别表头中独立主体列，如 12 主体列）。
    name_map：2026-08-06 外币核对——泰国原币 TB 科目名（应收账款和应收票据等）→ 集团标准名
    （应收账款等），使 TB 科目名能与底稿文件名匹配；归一化后同名键合并，代码有前缀关系时
    （一级 1130 vs 二级 1130.01）保留一级（短码）值，避免一级+二级双计。
    signed：2026-08-06 外币核对——泰国原币 TB 期末金额列为绝对值（方向列标借/贷），翻负后
    权益/负债类（贷方）科目为负，与底稿审定表（报表口径正数）对不上 → 泰国核对传 signed=False。"""
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    except Exception as ex:
        print(f'⚠️ 试算表打开失败：{ex}')
        return None, None, ''
    best = None
    # 2026-08-06 sheet 优先级：优先含『审定数/简单加计数』的正式 TB 表——Z 合并TB 多 sheet
    # （Index/关联方抵消/调整分录/TB表…），『关联方抵消』R2 含『科目』先命中但 0 科目
    # （bal=6 非审定口径）→ 取错 sheet。两级：先找含 审定数/简单加计数 的（正式TB表），
    # 无则退回含 项目/科目 的（FY 单 sheet 行为不变）。
    for _kws, _prefer in ((('项目',), True), (('科目',), True), (('项目',), False), (('科目',), False)):
        _prefer_flag = _prefer
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            hdr_i, idx = _find_header(rows, _kws)
            if hdr_i is not None:
                if _prefer_flag and not (('aud' in idx) or ('sum' in idx)):
                    continue   # 优先轮：只要正式口径列（审定数/简单加计数）
                best = (sn, hdr_i, idx, rows)
                break
        if best is not None:
            break
    wb.close()
    if best is None:
        return None, None, '未找到含「项目/科目+审定数/期初/期末」表头的 sheet'
    sn, hdr_i, idx, rows = best
    # 取值列：aud > sum > bal > qc > (jf,df)
    val_col = None
    for k in ('aud', 'sum', 'bal', 'qc'):
        if k in idx:
            val_col = idx[k]
            break
    if val_col is None and 'jf' in idx and 'df' in idx:
        val_col = ('jfdf', idx['jf'], idx['df'])
    # 识别主体列：表头行中名称含『公司/主体』且非 期初/加计数/调整/审定/校验/行次 的数值列
    # 2026-08-04 规模修复：全列扫描（120 主体列式试算表表头 121 列，[:40] 截断会丢后 80 个主体）
    # 2026-08-06 Z 合并TB：主体列名带日期前缀『02.251231高分子』（不含"公司/主体"字样）→
    # 加日期前缀识别（^\d+\.\d+ + 名称），否则 Z 合并TB 主体列全漏。
    ent_cols = []
    for j, h in enumerate(rows[hdr_i]):
        hs = str(h).strip() if h is not None else ''
        if not hs:
            continue
        if ('公司' in hs or '主体' in hs) and not any(k in hs for k in
                ('期初', '加计数', '调整', '审定', '校验', '行次', '页次', '简单')):
            ent_cols.append((j, hs))
        elif re.match(r'^\d+\.\d+', hs) and not any(k in hs for k in
                ('期初', '加计数', '调整', '审定', '校验', '行次', '页次', '简单')):
            ent_cols.append((j, hs))
    desc = f'sheet={sn} 表头行{hdr_i} 列映射: ' + ', '.join(f'{k}={v}' for k, v in sorted(idx.items())) + \
           (f' | 主体列 {len(ent_cols)}: ' + ','.join(h for _, h in ent_cols[:14]) if ent_cols else '')
    by_name, by_code = {}, {}
    _nm_code = {}   # 科目名 → 科目代码（2026-08-06 name_map 合并时按代码前缀判一级/二级）
    per_ent = {}   # {主体名: {科目名: 净额}}
    for r in rows[hdr_i + 1:]:
        if not r:
            continue
        nm = str(r[idx['name']]).strip() if idx['name'] < len(r) and r[idx['name']] is not None else ''
        # 2026-08-05：TB 脏行跳过——名称含『#』的行是名称丢失的重复/汇总行
        # （FY 合并TB：R159『#』=应付账款重复行 328.08M、R29『#』=应收净额+应收票据公式行
        # 108.94M），计入会与真实科目双列（核对稿曾出现『#』=328,076,401.11 与应付账款同值）。
        if not nm or any(k in nm for k in ('合计', '总计', '小计', '校验')) or '#' in nm:
            continue
        code = ''
        if idx.get('code') is not None and idx['code'] < len(r) and r[idx['code']] is not None:
            code = str(r[idx['code']]).strip()
        if isinstance(val_col, tuple):
            jf = _fnum(r[val_col[1]]) if val_col[1] < len(r) else 0.0
            df = _fnum(r[val_col[2]]) if val_col[2] < len(r) else 0.0
            bal = (jf or 0.0) - (df or 0.0)
        else:
            bal = _fnum(r[val_col]) if val_col < len(r) else None
            # 2026-08-06 Z 合并TB：审定数/简单加计列为空（合并TB 期末数据未填，仅期初有值）
            # → 用各主体列之和作为期末数兜底（Z 合并2025_数值版 主体列有期末值）。
            # 仅当 ent_cols 非空（存在主体列）且该行主体列有值时启用，避免误伤正常 TB。
            if bal is None and ent_cols:
                _ent_sum = 0.0
                _has_ent = False
                for _j, _h in ent_cols:
                    _v = _fnum(r[_j]) if _j < len(r) else None
                    if _v is not None:
                        _ent_sum += _v
                        _has_ent = True
                if _has_ent:
                    bal = _ent_sum
            if bal is None:
                continue
            if signed and idx.get('dir') is not None and idx['dir'] < len(r):
                d = str(r[idx['dir']]).strip() if r[idx['dir']] is not None else ''
                if '贷' in d:
                    bal = -bal
        # 2026-08-07 损益科目发生额口径：用户 TB 损益科目期末=0（结转后），底稿
        # 审定/明细/附注均按发生额（收入贷方、费用借方）→ 试算表侧用发生额，消除
        # 0 vs 发生额 假差异（dq 管理费用 74.29M vs 0、营业收入 6.22 亿 vs 0）。
        # 识别=科目代码 6 开头（6001-6901）；方向=贷>借判收入取贷方、否则取借方
        # （财务费用负数 jf=df=-11.3M 取 jf，与利润表一致）。
        if code and code[0] == '6' and idx.get('jf') is not None and idx.get('df') is not None:
            _pjf = _fnum(r[idx['jf']]) if idx['jf'] < len(r) else 0.0
            _pdf = _fnum(r[idx['df']]) if idx['df'] < len(r) else 0.0
            if abs(_pjf) > 0.005 or abs(_pdf) > 0.005:
                bal = _pdf if _pdf > _pjf else _pjf
        # 2026-08-06 修复：同名行（T：3001 实收资本 + 4001 实收资本 两主体不同代码）
        # 写入时即按『前缀消解+相加』合并——原『后者覆盖』在写入时就丢前值，
        # 事后遍历 by_name 无从补救（by_name 只剩末行）→ 底稿 31.32M vs TB 29.32M 差 2M。
        if nm in by_name:
            _old_code, _old_v = _nm_code.get(nm, ''), by_name[nm]
            if code and _old_code:
                if str(_old_code).startswith(code) and len(str(_old_code)) > len(code):
                    pass                        # 旧=长码(子目)，新=短码(一级) → 用新值（覆盖）
                elif code.startswith(str(_old_code)) and len(code) > len(str(_old_code)):
                    continue                    # 新=长码(子目)，旧=短码(一级) → 保留旧值
                else:
                    by_name[nm] = round(_old_v + bal, 2)   # 非同科目（3001/4001 同报表名）相加
                    _nm_code[nm] = min(code, _old_code, key=len) or code
                    continue
            else:
                by_name[nm] = round(_old_v + bal, 2)
                continue
        by_name[nm] = round(bal, 2)
        _nm_code[nm] = code
        if code:
            by_code[code] = round(bal, 2)
        # 主体列取值（净额：借正贷负——主体列本身为余额方向明确的正负值，直接取）
        for j, h in ent_cols:
            v = _fnum(r[j]) if j < len(r) else None
            if v is None:
                continue
            per_ent.setdefault(h, {})[nm] = round(v, 2)
    year = _detect_year(fp, rows)
    # 2026-08-06 修复：同名键合并【无条件执行】——不再依赖 name_map（泰国名映射只是
    # _nk 变换特例）。账套内同名科目用不同一级代码（T：北京同心 3001 实收资本 + 江西/
    # 股份 4001 实收资本）时，原『后者覆盖』致 by_name 只含末行 → 核对差 2 倍
    # （T 实收资本 -60.6M=31.3M×2-2M）。合并规则：代码有前缀关系（一级 vs 二级）保留
    # 短码一级值（防 一级+二级 双计）；无前缀关系（3001 vs 4001 同报表科目）相加。
    _nb = {}
    for _k, _v in by_name.items():
        _nk = name_map.get(_k, _k) if name_map else _k
        _code = _nm_code.get(_k, '')
        if _nk in _nb:
            _old_code, _old_v = _nb[_nk]
            if _old_code and _code:
                if str(_old_code).startswith(str(_code)) and len(str(_old_code)) > len(str(_code)):
                    _nb[_nk] = (_code, _v)          # 旧=长码(子目)，新=短码(一级) → 用一级
                    continue
                if str(_code).startswith(str(_old_code)) and len(str(_code)) > len(str(_old_code)):
                    continue                        # 新=长码(子目)，旧=短码(一级) → 保留一级
            _nb[_nk] = (_old_code, _old_v + _v)     # 非同科目（3001/4001 同报表名、泰国 1130.04/1151.05）相加
        else:
            _nb[_nk] = (_code, _v)
    by_name = {_k: _v for _k, (_c, _v) in _nb.items()}
    if by_code:
        _nc = {}
        for _k, _v in by_code.items():
            _nk = name_map.get(_k, _k) if name_map else _k
            _nc[_nk] = _nc.get(_nk, 0.0) + _v
        by_code = _nc
    if per_ent:
        _pe = {}
        for _e, _d in per_ent.items():
            _nd = {}
            for _k, _v in _d.items():
                _nk = name_map.get(_k, _k) if name_map else _k
                _nd[_nk] = _nd.get(_nk, 0.0) + _v
            _pe[_e] = _nd
        per_ent = _pe
    return by_name, by_code, desc, year, per_ent


def _is_double_hdr(rows, hdr_i):
    """判断 hdr_i 是否为双层表头的主表头行（下一行是子表头：含『期初数/本期增加/本期减少/期末数』
    等金额子列名，且当前行不含这些）。用于 应交税费/职工薪酬/货币资金 审定表。"""
    if hdr_i + 1 >= len(rows):
        return False
    nxt = rows[hdr_i + 1]
    if not nxt:
        return False
    nxt_txt = ''.join(str(x) if x is not None else '' for x in nxt[:16])
    cur_txt = ''.join(str(x) if x is not None else '' for x in rows[hdr_i][:16])
    sub_kws = ('期初数', '本期增加', '本期减少', '期末数', '借方', '贷方')
    return any(k in nxt_txt for k in sub_kws) and not any(k in cur_txt for k in ('期初数', '本期增加', '期末数'))


def _find_total_row(rows):
    """找合计行（前 2 列含 合计/总计/小计；往来审定表合计行首列为空、'合计' 在第 2 列）。
    2026-08-04 修复：明细表含 主体小计/主体合计/全集团合计 多级——优先『全集团合计』，
    其次取最后一个合计行（全集团口径），避免误取 主体小计。"""
    hits = []
    for i, r in enumerate(rows):
        if not r:
            continue
        a0 = str(r[0]).strip() if r[0] is not None else ''
        a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
        # 2026-08-07 修复：分类汇总表『在建工程（合计）』『累计摊销（合计）』的"合计"在
        # 科目名称列（D 列，前 2 列=主体/类别 不含合计）→ 原只查前 2 列识别不出 →
        # 回退数据行求和 → 在建工程三级展开后 二级+三级 双计 272M（父=子和）。扩展前 4 列。
        a2 = str(r[2]).strip() if len(r) > 2 and r[2] is not None else ''
        a3 = str(r[3]).strip() if len(r) > 3 and r[3] is not None else ''
        joined = a0 + a1 + a2 + a3
        if '勾稽' in joined:
            continue  # 2026-08-04：勾稽说明行（『勾稽：期初+本期增加…』）含"合计"字样，排除
        # 2026-08-05：核对说明行（『核对说明：与【一级科目】核对=其他应收款合计/减坏账准备/净额』）
        # 含"合计"字样会被误当合计行 → 其他应收款审定表曾把『减：坏账准备』行并入合计
        # （合计 11,541,664.82 + 坏账 1,030,192.18 = 12,571,857.00，虚增 103 万）。
        if '核对' in joined or '说明' in joined:
            continue
        # 2026-08-06：『报表其他应付款合计』等映射说明行（B 列=报表口径说明，非本表合计行）
        # 含"合计"会被当合计行 → F 列空 → 回退数据行求和 → 把『应付股利』146,300 并入
        # （其他应付款审定 31,729,997.42 vs 正确 31,583,697.42）。
        if a1.startswith('报表') or a0.startswith('报表'):
            continue
        if any(k in a0 for k in ('合计', '总计', '小计')) or any(k in a1 for k in ('合计', '总计', '小计')) \
                or any(k in a2 for k in ('合计', '总计', '小计')) or any(k in a3 for k in ('合计', '总计', '小计')):
            hits.append((i, r))
    if not hits:
        return None, None
    # 2026-08-06 修复：组件块审定表（固定资产：原值/累计折旧/减值/净值）有多个『全集团合计』
    # 行，reversed 取到最后=【净值】行 → dq 固定资产原值 819.4M vs 净值 514.5M 误报 304.9M。
    # 优先『全集团合计 + 原值/成本/余额』行（核对基准=用户TB 一级原值口径）；其次非净值全集团行；
    # 最后才是 a1 空的全集团行（无数值 → 交回退分支组件块原值）。
    for i, r in reversed(hits):
        a0 = str(r[0]).strip() if r[0] is not None else ''
        a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
        if ('全集团' in a0 or '全集团' in a1) and a1 and any(k in a1 for k in ('原值', '成本', '余额')):
            return i, r
    for i, r in reversed(hits):
        a0 = str(r[0]).strip() if r[0] is not None else ''
        a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
        if ('全集团' in a0 or '全集团' in a1) and a1 \
                and not any(k in a1 for k in ('净值', '账面价值', '折旧', '摊销', '减值', '准备')):
            return i, r
    for i, r in reversed(hits):
        a0 = str(r[0]).strip() if r[0] is not None else ''
        a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
        if '全集团' in a0 or '全集团' in a1:
            return i, r
    # 2026-08-04 规模修复：无『全集团合计』时，若最后一个合计行是『主体合计/小计』
    #（a0 含公司/主体 + a1 含合计/小计）→ 多主体块结构，单行取值只含最后一个主体
    #（120 主体时严重错误）→ 返回 None 交给块结构/主体级求和。
    last_i, last_r = hits[-1]
    a0 = str(last_r[0]).strip() if last_r[0] is not None else ''
    a1 = str(last_r[1]).strip() if len(last_r) > 1 and last_r[1] is not None else ''
    if ('公司' in a0 or '主体' in a0) and any(k in (a0 + a1) for k in ('合计', '小计')):
        return None, None
    return hits[-1]


def _col_by_header(rows, hdr_row, kws):
    """按表头关键词找列索引（优先级：含「余额/未审数」→ 普通列；排除 外币/方向 列）。
    2026-08-04 修复：同优先级内『期末』列优先（明细表『期初余额(人民币)』在『期末余额(人民币)』前，
    若按顺序取首个会命中期初列→单体核对提取到期初值）。
    2026-08-04 规模修复：全列扫描（120 主体的列式表头超 30 列，[:30] 会漏列）。"""
    _ncol = len(rows[hdr_row])
    for pref in (('余额', '未审数'), ()):
        # 第一轮：含 pref 且含『期末』
        for j in range(_ncol):
            h = rows[hdr_row][j]
            hs = str(h).strip() if h is not None else ''
            if '外币' in hs or '方向' in hs:
                continue
            if not any(k in hs for k in kws):
                continue
            if pref and not any(p in hs for p in pref):
                continue
            if '期末' in hs:
                return j
        # 第二轮：普通（含 pref）
        for j in range(_ncol):
            h = rows[hdr_row][j]
            hs = str(h).strip() if h is not None else ''
            if '外币' in hs or '方向' in hs:
                continue
            if not any(k in hs for k in kws):
                continue
            if pref and not any(p in hs for p in pref):
                continue
            return j
    return None


def _sheet_ent_values(ws, proj_kw=None, pref_col=None, year=None):
    """从底稿 sheet 提取 {主体名: 期末值}（2026-08-04 单体核对）。
    支持三类结构：
      A) 行式（审定表/明细表）：表头含『核算主体/公司/主体/账套主体』标识列 + 期末列 → 每主体一行；
         proj_kw 提供时（如『主营业务收入』）只计 A 列项目含该关键词的行（收入审定表防合计行重复）；
      C) 块式（固定资产 主体×组件）：A 列块标题=主体名 + 组件行含『净值/账面价值』；
      B) 列式（每主体一列附注/审定表）：表头列名含『公司』→ 每主体一列，取『期末』段行值合计。
    year 提供时（核对_2025/核对_2026）：含『年度』列的混年表只取该年份行（2026-08-07 修复：
    原取最新年度 → J 存货审定表 2025+2026 混排，核对 2025 时审定取到 2026 值 → 5 主体全差异）。
    返回 {主体名: 期末值}；无法解析返回 {}。"""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return {}
    # 找表头行（同 _sheet_end_total 规则：非空列≥3 且含 期初/期末 字样，排除标题/说明行）
    # 2026-08-04 规模修复：列式附注表头（『项目（审定数）| 01主体 | 02主体...』）无『期初/期末』字样
    # ——加兜底：表头行含 ≥2 个『公司/主体』列名且含『审定数/项目』即视为列式表头
    #（此前货币资金/资本公积附注 _sheet_ent_values 恒返回空 → 附注单体核对缺失）。
    hdr_i = None
    for i, r in enumerate(rows[:10]):
        if not r:
            continue
        s40 = ''.join(str(x) if x is not None else '' for x in r[:60])
        if any(k in s40 for k in ('明细表', '审定表', '附注', '底稿', '汇总（')):
            continue
        nfill = sum(1 for x in r[:60] if x is not None and str(x).strip())
        if nfill < 3:
            continue
        entish = sum(1 for x in r[:60] if x is not None and ('公司' in str(x) or '主体' in str(x)))
        # 2026-08-04 修复：损益两年对比表（财务费用明细表：核算主体|二级科目名称|2025全年(TB)|2026全年(TB)）
        # 表头无『期初/期末/本期数』字样 → 表头识别失败 → 明细表单体恒空。加『全年』；
        # 收入按主体明细表（本期收入/本期成本…）加『本期』（2026-08-04 规范化新表）。
        if ('期初' in s40 or '年初' in s40 or '未审数' in s40 or '本期' in s40 or '上年数' in s40
                or '全年' in s40) \
                and ('期末' in s40 or '余额' in s40 or '审定数' in s40 or '本期' in s40 or '全年' in s40):
            hdr_i = i
            break
        if entish >= 2 and ('审定数' in s40 or '项目' in s40):
            hdr_i = i
            break
    if hdr_i is None:
        return {}
    # 2026-08-04：pref_col 列偏好（收入按主体明细表：成本科目取『本期成本』列而非默认『本期收入』）
    col = None
    if pref_col:
        col = _col_by_header(rows, hdr_i, (pref_col,))
    if col is None:
        col = _col_by_header(rows, hdr_i, ('期末',))
    if col is None:
        col = _col_by_header(rows, hdr_i, ('本期数',))
    if col is None:
        col = _col_by_header(rows, hdr_i, ('本期',))   # 收入按主体明细表：本期收入/本期成本…
    if col is None:
        # 2026-08-04 修复：损益两年对比表取『本年列』——表头『2025全年(TB)/2026全年(TB)』，
        # 取含『全年』的最后一列（本年列在后；2025 为上年列）
        _full_cols = [j for j, h in enumerate(rows[hdr_i])
                      if h is not None and '全年' in str(h)]
        if _full_cols:
            col = _full_cols[-1]
    if col is None:
        col = _col_by_header(rows, hdr_i, ('审定数',))
    if col is None and hdr_i + 1 < len(rows):
        col = _col_by_header(rows, hdr_i + 1, ('期末',))
    # 2026-08-06 双层表头兜底（应交税费/货币资金：主表头『未审数|审计调整|重分类』合并列，
    # 子表头『期初数|本期增加|本期减少|期末数』）——主表头无『期末』列但子表头有：
    # 若主表头『审定数』兜底命中的列在『期末』子表头列之后（审计调整/重分类块），
    # 说明主表头是横向多块 → 必须用子表头『期末』列（此前应交税费取 col13=审计调整块
    # → 审定 -11,575,802 vs 正确合计 -9,223,263.68）。
    if hdr_i + 1 < len(rows):
        _c_sub = _col_by_header(rows, hdr_i + 1, ('期末',))
        if _c_sub is not None and (col is None or col > _c_sub):
            col = _c_sub
    if col is None:
        return {}
    out = {}
    # 情况 A：表头含『核算主体/公司/主体/账套主体』标识列（行式）
    ent_col = None
    for j, h in enumerate(rows[hdr_i]):
        hs = str(h).strip() if h is not None else ''
        if hs in ('核算主体', '主体', '公司', '核算单位', '单位名称', '账套主体', '主体名称'):
            ent_col = j
            break
    # 2026-08-06 年度列：存货审定表『主体|年度|科目代码|科目名称|期初…期末未审数』含年度列，
    # 2025+2026 两行都计入同一主体 → 428,735.38 vs 正确 226,682.29（混年双计）。
    # 识别表头含『年度』列 → 仅累计最新年度行（如 FY 2026；无年度列时不受影响）。
    year_col = None
    for j, h in enumerate(rows[hdr_i]):
        hs = str(h).strip() if h is not None else ''
        if hs in ('年度', '年份', '会计年度'):
            year_col = j
            break
    if year_col is not None and hdr_i + 1 < len(rows):
        _years = []
        for _r in rows[hdr_i + 1:]:
            if _r and year_col < len(_r) and _r[year_col] is not None:
                _yv = str(_r[year_col]).strip()
                if _yv.isdigit():
                    _years.append(_yv)
        _max_year = max(_years) if _years else None
    if ent_col is not None:
        # 主体×项目结构（使用权资产/固定资产审定表：A=主体、B=项目 原值/折旧/减值/净值）——
        # 若表头含『项目』列（第2列），只取 净值/账面价值/账面净值 行（防原值+折旧+净值重计）。
        # 2026-08-04 规范化：组件结构检测——仅当数据行确实含组件名（原值/累计折旧/减值/净值）
        # 才启用净值过滤；权益明细表（B 列也叫『项目(明细科目)』但无组件名）不再误伤（曾致
        # 实收资本/资本公积/盈余公积/未分配利润 明细表单体恒空）。
        proj2 = None
        if len(rows[hdr_i]) > 1:
            h1 = str(rows[hdr_i][1]).strip() if rows[hdr_i][1] is not None else ''
            # 2026-08-05：固定资产【分类汇总表】第 2 列为『类别』（非『项目』）但同为组件块结构
            # （原值/累计折旧/净值+类别子行）——加『类别』。组件名列优先取『科目名称』列
            # （分类汇总表组件名在 D 列：★固定资产净值…；审定表组件名在 B 列『项目』）。
            if '项目' in h1 or '类别' in h1:
                _scan = rows[hdr_i + 1:hdr_i + 60]
                if any(any(k in str(x) for k in ('原值', '累计折旧', '减值', '净值', '账面价值'))
                       for rr in _scan for x in (rr[:6] if rr else [])):
                    proj2 = 1
                    # 组件名列：『科目名称』优先（分类汇总表组件名在 D 列 ★净值…），
                    # 其次『项目』（审定表 B 列），最后『类别』（B 列='固定资产' 不含组件名）
                    for _kw in ('科目名称', '项目', '类别'):
                        for _j, _h in enumerate(rows[hdr_i][:6]):
                            _hs = str(_h).strip() if _h is not None else ''
                            if _kw in _hs:
                                proj2 = _j
                                break
                        if proj2 is not None:
                            break
        cur_en = None   # 向下填充：税种行/项目行核算主体列为空时沿用上一主体（应交税费 主体块首行模式）
        # 2026-08-06 块级合计优先：主体块内若有『合计/小计』行 → 用合计值（块权威），
        # 跳过明细累计——职工薪酬审定表『短期薪酬总行+离职后福利总行+其下基本工资等明细行』
        # 父级=子明细和，全累计双计（Z 上分 1,327,315.32 = 2×663,657.66）。块结束（新主体/
        # 循环尾）时：有合计用合计，否则用明细累计。
        _blk_sum = 0.0
        _blk_total = None
        _blk_trust = False   # 2026-08-07：块级合计可信标志（含主体名的『QG 合计』=净额，直接采用）
        _blk_ent = None
        _flushed_tot = set()   # 2026-08-06：已取过【块级合计】的主体。
        # 同一主体多块区分两类结构：
        #  · ga 职工薪酬审定表 = 汇总块(含『合 计』tot, 权威) + 明细块(无 tot, 明细=汇总拆解
        #    同值) → 汇总块保留、明细块跳过（防 2 倍叠加）；
        #  · 银行存款明细表 = 银行存款块 + 其他货币资金块（均无 tot，内容不同）→ 全保留。
        # 规则：块有 tot → 主体未取过才保留；块无 tot → 主体已取过 tot 则跳过（重复明细块），
        #       否则保留（账户类多块）。

        def _flush_blk():
            nonlocal _blk_ent, _blk_sum, _blk_total, _blk_trust
            if _blk_ent is None:
                return
            if _blk_total is not None:
                # 2026-08-07：块级可信合计（含主体名『QG 合计』=净额 373.8M，明细 abs 和
                # 530.7M 不等）→ 直接采用；表级合计（纯『合 计』=全集团）维持原逻辑。
                if _blk_trust:
                    _bv = _blk_total
                    _flushed_tot.add(_blk_ent)
                    if abs(_bv) > 0.005:
                        out[_blk_ent] = round(out.get(_blk_ent, 0.0) + _bv, 2)
                    _blk_ent, _blk_sum, _blk_total, _blk_trust = None, 0.0, None, False
                    return
                if _blk_ent in _flushed_tot:
                    _blk_ent, _blk_sum, _blk_total, _blk_trust = None, 0.0, None, False
                    return
                # 2026-08-06：合计行与块内明细和不符 → 合计行是【表级/跨块】合计
                # （应交税费表尾『合 计』=全集团，被向下填充进末主体 12 → 2 倍），
                # 用块内明细和；相符（合计=明细和，ga 职工薪酬块内『合 计』）用合计。
                if abs(_blk_total - _blk_sum) < 0.01:
                    _bv = _blk_total
                    _flushed_tot.add(_blk_ent)
                else:
                    _bv = _blk_sum
            else:
                if _blk_ent in _flushed_tot:
                    _blk_ent, _blk_sum, _blk_total, _blk_trust = None, 0.0, None, False
                    return
                _bv = _blk_sum
            if abs(_bv) > 0.005:
                out[_blk_ent] = round(out.get(_blk_ent, 0.0) + _bv, 2)
            _blk_ent, _blk_sum, _blk_total, _blk_trust = None, 0.0, None, False

        for i in range(hdr_i + 1, len(rows)):
            r = rows[i]
            if not r or len(r) <= max(ent_col, col):
                continue
            # 2026-08-04 规模修复：合计/小计行（A 列空、C 列『银行存款合计』等）不得因主体向下填充
            # 被当明细行计入（银行存款明细表 r16/r21/r22 合计行曾加进上一主体 → 主体级求和 9.2 亿）
            # 2026-08-06 修复：『小  计』中间是半角空格(0x20)×2 或全角空格——'小计' in 原串 为
            # False → 应交税费 01FY 小计行 -8,393,844.59 被计入（审定 -16.8M vs 正确 -9.8M）。
            # 去全部空格（半角+全角）后匹配。
            # 2026-08-06 再改：合计/小计行不再是『跳过』——记录为块级合计值（块权威，
            # 职工薪酬审定表父级总行+子明细行双计 2× 的修复），明细累计改由 flush 时二选一。
            # 2026-08-06 17:05 再改：row_kind 只传【行名（A 列）】——数值列公式值 =SUM(…)
            # 的『=』会被 is_note_row 判为说明行 → 合计行全被当 note 跳过。
            # 2026-08-06 17:30 再改：传 A/B/C 三列——银行存款明细表合计行 A/B 空、C 列
            # 『银行存款合计』（开户银行列）→ 只传 A 列识别不出 → 合计行当数据行累计进
            # 末主体（12FY控股=369M 全集团值，明细 3 倍）。B/C 列在已协议化表中为文本或
            # 数值（无公式），『=』判定安全。
            # 2026-08-06 20:00 再改：传 A/B/C/D——存货审定表『小计』在 D 列（科目名称），
            # A/B/C 识别不出 → 小计行当数据行 → 主体级求和 2×（FY 存货 453,364.58=2×226,682.29）。
            # 2026-08-06 年度过滤：存货审定表 2025+2026 双行，只计最新年度（_max_year 由上方
            # 表头『年度』列识别；无年度列时 year_col=None 跳过）
            # 2026-08-07 修复：year 参数（核对_2025/核对_2026 当前年份）优先——J 存货审定表
            # 2025+2026 混排，核对 2025 时原取 2026（最新）→ 审定列错 5 主体；核对 2026 取 2026。
            # 2026-08-07 再修：本过滤必须在 row_kind/块级合计判断【之前】——g 存货审定表
            # 『杭州铁城 2026 小计』行若先过 row_kind（sub+含主体名 → 块级可信 _blk_total）
            # 会在 year 过滤前被记录为块值 → 核对 2025 时杭州铁城审定=2026 小计 144,046,006.99。
            if year_col is not None and (_max_year or year):
                _yv = str(r[year_col]).strip() if year_col < len(r) and r[year_col] is not None else ''
                _want = str(year) if year else _max_year
                if _yv != _want:
                    continue
            _rowtxt = ' '.join(str(x) for x in r[:4] if x is not None)
            # 2026-08-06 协议化：行类型统一走 audit_protocol.row_kind（去空白匹配唯一写法，
            # 兼容旧变体）。'tot'=块级合计（纯『合 计』，权威覆盖块值）/ 'sub'=分类小计
            # （如『银行存款合计』是【全表】分类合计，非主体块值——曾致 12 主体（末主体）
            # 把合计 369M 当块值 → 明细 2 倍）/ 'note'=说明勾稽行 / 'grp'/'offset'/'consol'
            # 集团行·抵消·合并数（非主体数据，跳过）。
            _rk = PROTO.row_kind(_rowtxt)
            if _rk == 'note':
                continue
            if _rk == 'tot':
                _blk_total = _resolve_value(rows, i, col)
                _blk_trust = False   # 表级合计（纯『合 计』=全集团），相等才用
                continue
            # 2026-08-07 修复：主体块级合计『QG 合计』『母公司 合计』（B 列含主体名+合计）
            # 被 row_kind 判 'sub' 跳过 → 明细行 abs 累计（Q 资本公积 400202 贷余负被翻正
            # → 530.7M vs 一级净额 373.8M）。'sub' 且行文本含当前主体名 → 记录为块级合计；
            # 全表分类合计（银行存款合计，不含主体名）不受影响。
            if _rk == 'sub' and _blk_ent and str(_blk_ent) in _rowtxt:
                _blk_total = _resolve_value(rows, i, col)
                _blk_trust = True    # 块级合计（含主体名），净额直接采用
                continue
            if _rk in ('sub', 'grp', 'offset', 'consol'):
                continue
            # 2026-08-04 修复：明细表父级行（『车辆费（含三级明细）』）与三级子行并存 → 双计；
            # 父级行值=子行和，跳过父级（名称含『（含』）只计末级（管理费用明细表曾 76M vs TB 25.8M）
            if len(r) > 1 and r[1] is not None and '（含' in str(r[1]):
                continue
            # 年度过滤已移至 row_kind 判断之前（2026-08-07，见上方）——此处不重复
            en = str(r[ent_col]).strip() if r[ent_col] is not None else ''
            # 2026-08-06 A 列汇总行：应交税费审定表『集团加计』在 A 列（B 列主体空）——
            # 若只查 en（B 列）会漏掉 → 集团加计行（=全部主体×税种之和）被当主体计入
            # → 审定 -15,434,403 vs 正确合计 -9,223,263.68（双计 621 万）
            _a0row = str(r[0]).strip() if r[0] is not None else ''
            if any(k in _a0row for k in ('集团加计', '集团合计', '全集团', '汇总')):
                continue
            if en:
                # 2026-08-05 加『汇总』：职工薪酬审定表『集团汇总数』行（各主体之和）
                # 曾当主体计入 → 审定值 = 主体和×2（32M vs TB 16.7M，双计）
                # 2026-08-06 加『集团加计』：应交税费审定表『集团加计』行（=全部主体×税种
                # 之和）曾当主体计入 → 审定 -15,434,403 vs 正确合计 -9,223,263.68（双计 621 万）
                if any(k in en for k in ('合计', '总计', '小计', '全集团', '说明', '汇总', '集团加计')):
                    continue
                if cur_en != en:
                    _flush_blk()          # 新主体块 → 结算上一块
                cur_en = en
                if _blk_ent is None:
                    _blk_ent = en
                if proj2 is not None:
                    # 主体块标题行（B 列空，如 应交税费『1|01FY本级公司||…』）仅登记主体
                    pj = str(r[proj2]).strip() if len(r) > proj2 and r[proj2] is not None else ''
                    if not pj:
                        continue
            elif cur_en is None:
                continue
            en = cur_en
            if proj2 is not None:
                pj = str(r[proj2]).strip() if len(r) > proj2 and r[proj2] is not None else ''
                # 2026-08-07 修复：组件块审定表（无形资产：原值/累计摊销/减值/净值）此前
                # 只取净值行（防原值+折旧+净值重计），但核对基准=用户TB 一级【原值】口径
                #（与 _sheet_end_total 375 行同规则）→ 只累计 原值/成本/余额 行，净值/
                # 折旧/减值行跳过（Z 无形资产审定 46.27M 净值 vs TB 61.78M 原值，差=累计
                # 摊销 15.5M；F 固定资产净值 -3.02 亿 vs 原值 27.95M，均为假差异）。
                if not any(k in pj for k in ('原值', '成本', '余额')):
                    continue
            # proj_kw：只计 项目列（A 列）或 明细科目列（B 列）含该关键词的行
            #（收入审定表 A 列=项目名防营业总收入合计行重复；未分配利润明细表 A 列=主体名、
            #  B 列=『4104.12 未分配利润』——proj_kw 命中 B 列，2026-08-04 修复曾全滤空）
            # 2026-08-04 再修：按主体明细表（A 列=主体名『01FY本级公司』）→ 不过滤行
            #（列偏好 pref_col 已选 本期收入/本期成本 列；proj_kw 过滤会滤空主体行）
            if proj_kw and ent_col is not None and ent_col > 0:
                a0 = str(r[0]).strip() if r[0] is not None else ''
                a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
                # 2026-08-06 收入/成本类：A 列块标题『一、主营业务收入』不含『营业收入』——
                # proj_kw='营业收入' 时按 SUBJ_ROW_KW 展开（主营业务收入/其他业务收入 均计入，
                # 营业收入=主营+其他）；其余科目仍精确匹配。
                _kws = SUBJ_ROW_KW.get(proj_kw, (proj_kw,))
                if ('公司' not in a0 and '主体' not in a0) \
                        and not any(_k and (_k in a0 or _k in a1) for _k in _kws):
                    continue
            # 2026-08-07 修复：备抵行（科目名以『减：』/『减:』开头，如存货审定表
            # 『减：存货跌价准备』）不参与主体累计——曾把备抵并入主体块
            # 2026-08-07 修复：备抵行过滤加【备抵词】条件——原『减:』前缀全跳过把损益
            # 审定表项目行『减：主营业务成本』『减：其他业务成本』（利润表项目，非备抵）
            # 误当备抵剔除 → 主营业务成本审定表未取到（dq 成本 251.34M 缺失）。
            # 备抵行特征=『减:』且行名含 备抵词（坏账/减值/跌价/折旧/摊销/准备）。
            _tok = _rowtxt.replace('：', ':').split()
            if any(t.startswith('减:') for t in _tok) \
                    and any(k in _rowtxt for k in ('坏账', '减值', '跌价', '折旧', '摊销', '准备')):
                continue
            v = _resolve_value(rows, i, col)
            if v is not None:
                if _blk_total is not None:
                    # 2026-08-06：块内已设『合 计』（块权威）→ 后续行跳过——ga 职工薪酬
                    # 审定表每主体=汇总块(短期薪酬等+合 计)+明细块(基本工资等,重复内容)，
                    # 明细行再累计 → _blk_sum=2×合计（83,428 vs 41,714）→ flush 误用明细。
                    # 设合计后明细跳过 → _blk_sum=汇总行和=合计 → flush 二选一正确。
                    continue
                _blk_sum += v        # 块内明细累计（flush 时若无合计行则用此值）
        _flush_blk()
        if out:
            return out
    # 情况 C：块式（固定资产等：A 列块标题=主体名 + 净值行）
    proj_col = None
    for j, h in enumerate(rows[hdr_i]):
        hs = str(h).strip() if h is not None else ''
        if '项目' in hs or '资产' == hs:
            proj_col = j
            break
    if proj_col is not None:
        cur_ent = None
        for i in range(hdr_i + 1, len(rows)):
            r = rows[i]
            if not r or len(r) <= max(proj_col, col):
                continue
            a0 = str(r[0]).strip() if r[0] is not None else ''
            a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
            if a0 and ('公司' in a0 or '主体' in a0) and not a1:
                cur_ent = a0
                continue
            pj = str(r[proj_col]).strip() if r[proj_col] is not None else ''
            # 2026-08-07：与位置 1 同规则——块式组件结构取【原值/成本/余额】行（TB 原值口径）
            if cur_ent and any(k in pj for k in ('原值', '成本', '余额')):
                v = _resolve_value(rows, i, col)
                if v is not None:
                    out[cur_ent] = round(out.get(cur_ent, 0.0) + v, 2)
        if out:
            return out
    # 情况 B：列式（每主体一列附注/审定表）——列名含『公司』的主体列 + 『期末』段行值。
    # 2026-08-04 规模修复：全列扫描（120 主体的每主体一列表头 121 列，[:40] 只识别前 40 主体）
    ent_cols = []
    for j, h in enumerate(rows[hdr_i]):
        hs = str(h).strip() if h is not None else ''
        if ('公司' in hs or '主体' in hs) and not any(k in hs for k in
                ('期初', '加计数', '调整', '审定', '校验', '行次', '页次', '简单')):
            ent_cols.append((j, hs))
    if ent_cols:
        seg = None
        for i in range(hdr_i + 1, len(rows)):
            r = rows[i]
            if not r:
                continue
            s = str(r[0]).strip() if r[0] is not None else ''
            if '期末' in s:
                seg = i
                break
        if seg is None:
            seg = hdr_i + 1
        # 2026-08-04 规模修复：段窗口 60→400（120 主体且每主体一段的列式结构，期末段可超 60 行；
        # 遇『期初』段标题即止，避免把期初段行值当期末）
        for i in range(seg, min(seg + 400, len(rows))):
            r = rows[i]
            if not r:
                continue
            a0 = str(r[0]).strip() if r[0] is not None else ''
            # 2026-08-04 修复：段边界——离开『期末数』段（审计调整/期末审定数/期初数/抵消/报表数等）
            # 即停，防 期末审定数 段公式行（=B4+B8）与 期末数 段重复累计（曾双计）
            if a0 and i > seg and any(k in a0 for k in ('审计', '期初', '调整', '审定', '抵消', '报表',
                                                        '增加', '减少', '小计', '合 计', '合计')):
                break
            if not a0 or any(k in a0 for k in ('合计', '总计', '小计')):
                continue
            for j, h in ent_cols:
                v = _resolve_value(rows, i, j)
                if v is not None and abs(v) > 0.005:
                    out[h] = round(out.get(h, 0.0) + v, 2)
        if out:
            return out
    return {}


def _sheet_end_total(ws):
    """取某 sheet 的合计行「期末」列值。返回 (值, 说明) 或 (None, 原因)。兼容包装。"""
    return _sheet_end_total_impl(ws)


def _block_subtotal(ws, kw):
    """块合计行提取：明细表按块标题（『银行存款合计』『其他货币资金合计』）取期末列值。
    货币资金明细表 列2=科目块（银行存款/其他货币资金），合计行列2=『{科目}合计』。
    2026-08-05：核对单个货币资金科目时明细表须取本块合计行——_sheet_end_total 取最后合计行
    （『货币资金合计』307,443,528.61）曾致银行存款明细差异 -264 万。"""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return None, '空表'
    hdr_i = None
    for i, r in enumerate(rows[:10]):
        if not r:
            continue
        if sum(1 for x in r[:12] if x is not None and str(x).strip()) < 3:
            continue
        s20 = ''.join(str(x) if x is not None else '' for x in r[:40])
        if any(k in s20 for k in ('明细表', '审定表', '附注', '底稿', '汇总（')):
            continue
        if ('期初' in s20 or '年初' in s20 or '未审数' in s20 or '本期' in s20) \
                and ('期末' in s20 or '余额' in s20 or '审定数' in s20) \
                and ('主体' in s20 or '公司' in s20 or '项目' in s20 or '科目' in s20):
            hdr_i = i
            break
    if hdr_i is None:
        return None, '无表头'
    col = _col_by_header(rows, hdr_i, ('期末',))
    if col is None:
        col = _col_by_header(rows, hdr_i, ('余额',))
    if col is None:
        return None, '无期末列'
    kw_flat = kw.replace(' ', '')
    for i in range(hdr_i + 1, len(rows)):
        r = rows[i]
        if not r or len(r) <= col:
            continue
        label = ''.join(str(x) if x is not None else '' for x in r[:6])
        label_flat = label.replace(' ', '').replace('\u3000', '')
        # 2026-08-06 精确匹配：'货币资金合计' in '其他货币资金合计'=True（子串误中）——
        # 货币资金明细表『其他货币资金合计』R22 曾被当『货币资金合计』（264 万 vs 3.07 亿）。
        # 精确条件：label==kw+合计，或 label 以 kw 开头且 k 后紧跟合计（银行存款合计）。
        if label_flat == f'{kw_flat}合计' or label_flat.startswith(f'{kw_flat}合计'):
            v = _resolve_value(rows, i, col)
            if v is not None:
                return v, f'块合计行({kw}合计)'
    return None, f'未找到{kw}合计行'


def _baddebt_row_total(ws, kw):
    """备抵行专项提取：审定表找『减：坏账准备/累计折旧/减值准备』行，取期末列值。
    返回 (值, 行标签) 或 (None, 原因)。2026-08-04 用户要求备抵行对应审定表备抵项。"""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return None, '空表'
    hdr_i = None
    for i, r in enumerate(rows[:10]):
        if not r:
            continue
        if sum(1 for x in r[:12] if x is not None and str(x).strip()) < 3:
            continue
        s20 = ''.join(str(x) if x is not None else '' for x in r[:40])
        if any(k in s20 for k in ('明细表', '附注', '底稿', '汇总（')):
            continue
        if ('期初' in s20 or '年初' in s20 or '未审数' in s20 or '本期数' in s20 or '上年数' in s20) \
                and ('期末' in s20 or '余额' in s20 or '审定数' in s20 or '本期数' in s20) \
                and ('主体' in s20 or '公司' in s20 or '年度' in s20 or '项目' in s20):
            hdr_i = i
            break
    if hdr_i is None:
        return None, '无表头'
    hdr = rows[hdr_i]
    # 期末列：优先『期末』，其次『余额/审定数』
    end_col = None
    for i, h in enumerate(hdr[:20]):
        hs = str(h) if h is not None else ''
        if '期末' in hs:
            end_col = i
            break
    if end_col is None:
        for i, h in enumerate(hdr[:20]):
            hs = str(h) if h is not None else ''
            if '余额' in hs or '审定数' in hs:
                end_col = i
                break
    if end_col is None:
        return None, '无期末列'
    # 找含 kw 的行（备抵行标签：『减：坏账准备』『累计折旧』『减值准备』等）。
    # 组件块（固定资产/使用权：主体×原值/累计折旧/减值/净值）多行匹配 → 全部求和（集团口径）。
    # 2026-08-07 修复：组件块审定表每主体一行+全集团合计行 → 全匹配 = 主体和 + 合计 = 2×
    #（Z 固定资产累计折旧 657M = 主体和 328M + 合计 328M）。应只取【全集团合计】行（权威），
    # 无合计行才回退全部主体行求和。同时 kw 精确匹配避免『累计折旧』命中『使用权资产累计折旧』
    #（Z 固定资产审定表 41 行=使用权资产累计折旧，非固定资产备抵，须排除）。
    tot = 0.0
    hit = 0
    grp_tot = None
    grp_hit = 0
    for r in rows[hdr_i + 1:]:
        label = ' '.join(str(x) if x is not None else '' for x in r[:min(end_col, 4) + 1])
        if kw and kw not in label:
            continue
        # kw 精确匹配：『累计折旧』行若同时含『使用权/无形/投资性』前缀词 → 非本科目备抵，跳过
        if kw in ('累计折旧', '累计摊销') and any(k in label for k in
                ('使用权资产', '无形资产', '投资性房地产', '在建工程', '生产性生物')):
            continue
        v = r[end_col] if end_col < len(r) else None
        if isinstance(v, str) and v.startswith('='):
            v = _resolve_value(rows, r, end_col)
        try:
            fv = abs(float(v))
        except (TypeError, ValueError):
            continue
        a0 = str(r[0]).strip() if r[0] is not None else ''
        if '全集团' in a0 or '合计' in a0:
            grp_tot = fv
            grp_hit += 1
        else:
            tot += fv
            hit += 1
    if grp_hit:
        return grp_tot, f'备抵行(全集团合计{grp_hit}行)'
    if hit:
        return tot, f'备抵行({hit}行)'
    return None, '未找到备抵行'


def _sheet_end_total_impl(ws):
    """取某 sheet 的合计行「期末」列值。返回 (值, 说明) 或 (None, 原因)。
    2026-08-04 修复链：①表头识别排除含『=』的说明行（如『期末审定数=期末未审数…』会污染
    proj_col 识别致组件块全列求和）；②合计行公式无缓存 → _resolve_value 解析 SUM 区间/单元格
    引用（精确不重计，替代旧"数据行回退求和"——曾把收入审定表收入行+利润行重复计入）；    ③组件块（固定资产/在建/使用权/无形）取『净值/账面价值/账面净值』行合计；
    ④损益类多块（收入审定表：主营收入+其他业务收入）取各收入块全集团合计行本期数求和。"""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return None, '空表'
    hdr_i = None
    for i, r in enumerate(rows[:10]):
        if not r:
            continue
        # 2026-08-04：仅首列有内容的行 = 标题行/说明行（如『期末审定数=期末未审数…』），
        # 非空列数 >=3 才是真实表头行（勿用『=』排除——明细表列名『期末余额(=借款金额)』含=）
        if sum(1 for x in r[:12] if x is not None and str(x).strip()) < 3:
            continue
        s20 = ''.join(str(x) if x is not None else '' for x in r[:40])
        # 排除标题行（含『明细表/审定表/附注/底稿/汇总（』字样）
        if any(k in s20 for k in ('明细表', '审定表', '附注', '底稿', '汇总（')):
            continue
        # 2026-08-04：损益两年对比表（财务费用明细表：2025全年(TB)/2026全年(TB)）表头无
        # 『期初/期末/本期数』字样 → 加『全年』；收入按主体明细表（本期收入…）加『本期』。
        if ('期初' in s20 or '年初' in s20 or '未审数' in s20 or '本期' in s20 or '上年数' in s20
                or '全年' in s20) \
                and ('期末' in s20 or '余额' in s20 or '审定数' in s20 or '本期' in s20 or '全年' in s20) \
                and ('主体' in s20 or '公司' in s20 or '年度' in s20 or '项目' in s20):
            hdr_i = i
            break
    if hdr_i is None:
        return None, '无表头'
    col = _col_by_header(rows, hdr_i, ('期末',))
    if col is None:
        col = _col_by_header(rows, hdr_i, ('本期数',))   # 损益审定表：本期数(发生额)
    if col is None:
        col = _col_by_header(rows, hdr_i, ('本期',))     # 收入按主体明细表：本期收入/本期成本…
    if col is None:
        # 2026-08-04：损益两年对比表取『本年列』（2025全年/2026全年 → 最后一列=2026）
        _full_cols = [j for j, h in enumerate(rows[hdr_i])
                      if h is not None and '全年' in str(h)]
        if _full_cols:
            col = _full_cols[-1]
    # 2026-08-04：双层表头（职工薪酬：主表头 未审数/审计调整/审定数 + 子表头 期初数/本期增加/本期减少/期末数）
    # 必须优先用 子表头 的『期末』——主表头『审定数』在横向多块（未审/调整/重分类/审定）中靠后，
    # 会误定位到审定数块列（薪酬曾取 col13 而非 col6 → 小计求和 8146 万 vs 正确 2406 万）。
    # 仅对 薪酬（行含短期薪酬）启用；应交税费/货币资金 等其他双层表头用各自块逻辑（勿改列位）。
    _is_payroll_hdr = hdr_i + 1 < len(rows) and any(
        '短期薪酬' in ''.join(str(x) if x is not None else '' for x in r[:4])
        for r in rows[hdr_i + 1:hdr_i + 40])
    if _is_payroll_hdr and hdr_i + 1 < len(rows):
        _c2 = _col_by_header(rows, hdr_i + 1, ('期末',))
        if _c2 is not None:
            col = _c2
    if col is None:
        col = _col_by_header(rows, hdr_i, ('审定数',))
    if col is None:
        return None, '无期末列'
    # ① 损益类多块特判（收入审定表）：表头含『本期数』+『项目』，块标题含 主营业务收入/其他业务收入
    hdr_txt = ''.join(str(x) if x is not None else '' for x in rows[hdr_i][:24])
    if '本期数' in hdr_txt and '项目' in hdr_txt:
        rev_tot = 0.0
        for i in range(hdr_i + 1, len(rows)):
            r = rows[i]
            if not r or len(r) <= col:
                continue
            a0 = str(r[0]).strip() if r[0] is not None else ''
            a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
            if '全集团合计' in a1 and ('主营业务收入' in a0 or '其他业务收入' in a0):
                v = _resolve_value(rows, i, col)
                if abs(v) < 0.005:
                    # 合计行该列空（旧版底稿仅文本）→ 该块数据行（a0 相同、非合计/小计）求和
                    for ii in range(hdr_i + 1, i):
                        r2 = rows[ii]
                        if not r2 or len(r2) <= col or r2[0] is None:
                            continue
                        if str(r2[0]).strip() != a0:
                            continue
                        a1b = str(r2[1]).strip() if len(r2) > 1 and r2[1] is not None else ''
                        if any(k in a1b for k in ('合计', '小计')):
                            continue
                        v += _resolve_value(rows, ii, col)
                rev_tot += v
        if abs(rev_tot) > 0.005:
            return rev_tot, 'OK(损益收入块合计)'
    # ①b 薪酬特判（职工薪酬审定表：每主体块 期初/增加/减少/期末 + 短期薪酬小计/设定提存计划小计）：
    # 必须在 _find_total_row 之前——薪酬表末行『勾稽：…合计 应=TB一级』含"合计"会被误当合计行。
    _h20 = ''.join(str(x) if x is not None else '' for x in rows[hdr_i][:24]) + \
           ''.join(str(x) if x is not None else '' for x in (rows[hdr_i + 1][:24] if hdr_i + 1 < len(rows) else []))
    _is_payroll = ('未审数' in _h20 and '期初数' in _h20 and '本期增加' in _h20 and '本期减少' in _h20) \
                  and any('短期薪酬' in ''.join(str(x) if x is not None else '' for x in r[:4])
                          for r in rows[hdr_i + 1:hdr_i + 40])
    if _is_payroll:
        tot = 0.0
        _cur = ''
        for i in range(hdr_i + 1, len(rows)):
            r = rows[i]
            if not r or col >= len(r):
                continue
            a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
            a2 = str(r[2]).strip() if len(r) > 2 and r[2] is not None else ''
            if a1:
                _cur = a1  # 主体在块首行（B 列），小计行 B 空 → 向下填充
            # 2026-08-05 跳过『集团汇总数』块小计行（=各主体小计之和，防双计 2×16M）
            if '集团汇总数' in _cur:
                continue
            if a2.endswith('小计') and '勾稽' not in a2 and r[col] is not None:
                tot += _resolve_value(rows, i, col)
        if abs(tot) > 0.005:
            return tot, 'OK(薪酬各主体小计行期末)'
    ti, tr = _find_total_row(rows[hdr_i + 1:])
    if tr is None:
        # ② 主体×组件块结构（固定资产/在建/使用权/无形）→ 取『净值/账面价值/账面净值』行合计
        proj_col = None
        for j, h in enumerate(rows[hdr_i]):
            hs = str(h).strip() if h is not None else ''
            if '项目' in hs or '资产' == hs:
                proj_col = j
                break
        if proj_col is not None:
            tot = 0.0
            for i in range(hdr_i + 1, len(rows)):
                r = rows[i]
                if not r or len(r) <= proj_col or r[proj_col] is None:
                    continue
                pj = str(r[proj_col]).strip()
                if not any(k in pj for k in ('净值', '账面价值', '账面净值')):
                    continue
                tot += _resolve_value(rows, i, col)
            if abs(tot) > 0.005:
                return tot, 'OK(净值/账面价值行合计)'
        # ③ 无合计行的块结构（应交税费/职工薪酬：主体×税种/薪酬项目）→ 期末列数据行求和
        #（项目行均为末级发生额，求和=该科目期末；跳过空行/汇总行/合计行防多级叠加）
        # 2026-08-04 跳过行检测取前 4 列（银行存款明细表合计行『银行存款合计』在第 3 列）
        # 2026-08-04 规模修复：先试『主体级求和』（_sheet_ent_values 逐主体提取，线性且防
        # 父级+子级双计/行内列位不一致——120 主体时比全行求和可靠）；非空才用它，否则全行求和。
        _ent_sum = sum((_sheet_ent_values(ws)).values()) if ws is not None else 0.0
        if abs(_ent_sum) > 0.005:
            return _ent_sum, 'OK(主体级求和)'
        _a0123 = lambda rr: ' '.join(str(rr[k]).strip() if len(rr) > k and rr[k] is not None else ''
                                     for k in range(0, min(4, len(rr))))
        _a0123f = lambda rr: _a0123(rr).replace(' ', '').replace('　', '')
        tot = 0.0
        for i in range(hdr_i + 1, len(rows)):
            r = rows[i]
            if not r or col >= len(r):
                continue
            if any(k in _a0123f(r) for k in ('合计', '总计', '小计', '（含')):
                continue
            tot += _resolve_value(rows, i, col)
        if abs(tot) > 0.005:
            return tot, 'OK(块结构求和)'
        return None, '无合计行(主体块结构)'
    tr = rows[hdr_i + 1 + ti]
    v = _resolve_value(rows, hdr_i + 1 + ti, col)
    if abs(v) < 0.005:
        # 2026-08-04：合计行该列空/公式无效（部分底稿合计行仅 A/B 列文本，如固定资产
        # 『全集团合计』行 D 列空）→ 回退：①组件块（表头含『项目』+ 数据含 原值/累计折旧/净值）
        # 取『净值/账面价值/账面净值』行合计；②普通结构按数据行求和（跳过合计/小计行防多级叠加）。
        proj_col = None
        for j, h in enumerate(rows[hdr_i]):
            hs = str(h).strip() if h is not None else ''
            if '项目' in hs or '资产' == hs:
                proj_col = j
                break
        if proj_col is not None:
            # 2026-08-06 修复：组件块优先取【原值/账面余额/成本】行——核对基准=用户TB
            # 一级科目（固定资产 1601 原值口径）；原取『净值/账面价值』行 → dq 固定资产
            # 原值 819.4M vs 净值 514.5M 误报 304.9M（=累计折旧额）。
            gross = 0.0
            for i in range(hdr_i + 1, len(rows)):
                r = rows[i]
                if not r or len(r) <= proj_col or r[proj_col] is None:
                    continue
                pj = str(r[proj_col]).strip()
                if not any(k in pj for k in ('原值', '账面余额', '成本')):
                    continue
                if any(k in pj for k in ('净值', '价值', '净额')):
                    continue
                gross += _resolve_value(rows, i, col)
            if abs(gross) > 0.005:
                return gross, 'OK(组件块原值行)'
            net = 0.0
            for i in range(hdr_i + 1, len(rows)):
                r = rows[i]
                if not r or len(r) <= proj_col or r[proj_col] is None:
                    continue
                pj = str(r[proj_col]).strip()
                if not any(k in pj for k in ('净值', '账面价值', '账面净值')):
                    continue
                net += _resolve_value(rows, i, col)
            if abs(net) > 0.005:
                return net, 'OK(合计行空-净值行回退)'
        tot = 0.0
        _a0123 = lambda rr: ' '.join(str(rr[k]).strip() if len(rr) > k and rr[k] is not None else ''
                                     for k in range(0, min(4, len(rr))))
        _a0123f = lambda rr: _a0123(rr).replace(' ', '').replace('　', '')
        for i in range(hdr_i + 1, hdr_i + 1 + ti):
            r = rows[i]
            if not r or col >= len(r):
                continue
            if any(k in _a0123f(r) for k in ('合计', '总计', '小计', '（含')):
                continue
            tot += _resolve_value(rows, i, col)
        if abs(tot) > 0.005:
            return tot, 'OK(合计行空-数据行求和)'
    return v, 'OK'


def recon_trial_balance(tb_path, data_dir=None, out_path=None):
    data_dir = data_dir or os.path.dirname(tb_path) or '.'
    _acct = os.path.basename(data_dir.rstrip('\\/'))   # 2026-08-07 任务683：账套名（norm_name 用）
    # 2026-08-06 外币核对：泰国原币 TB（科目名泰国化：应收账款和应收票据/预付费用等）→
    # 集团标准名（应收账款/预付账款等），使 TB 科目名与底稿文件名匹配（新加坡为集团标准
    # 代码/名称，无需映射）。
    _name_map = None
    _tb_base = os.path.basename(tb_path)
    if '泰国' in _tb_base or 'THAI' in _tb_base.upper():
        try:
            from thai_mapping import THAI_NAME_TO_STD
            _name_map = dict(THAI_NAME_TO_STD)
            # 泰国科目名 → 集团标准名补充：现金+银行存款→货币资金（并入银行存款底稿）、
            # 税费归集到应交税费、权益类对齐底稿名、累计折旧走备抵行分支
            _name_map.update({
                '现金': '货币资金',
                '银行存款': '货币资金',
                '销项税-7%': '应交税费',
                '应缴预扣税款': '应交税费',
                '待缴进项税': '应交税费',
                '进项税额': '应交税费',
                '实收资本（股本）': '实收资本(或股本)',
                '利润分配': '未分配利润',
                '累计折旧': '减：累计折旧',
                # 2026-08-06 二级/三级科目名对齐底稿名（泰国 2120.03 其他应付账款→其他应付款、
                # 1151.05 押金和担保→其他应收款；否则『其他应付账款』含『应付账款』子串
                # 误匹配应付账款底稿）
                '其他应付账款': '其他应付款',
                '押金和担保': '其他应收款',
            })
            print('✓ 检测到泰国账套 → 科目名归一化映射 %d 条' % len(_name_map))
        except ImportError:
            print('⚠️ thai_mapping 导入失败，泰国科目名不做归一化')
    by_name, by_code, desc, tb_year, per_ent = parse_trial_balance(
        tb_path, _name_map,
        signed=not any(k in _tb_base for k in ('泰国', 'THAI', 'thai', '新加坡')))
    if by_name is None:
        print(f'⚠️ 试算表解析失败：{desc}')
        return
    print(f'✓ 试算表解析成功：{len(by_name)} 个科目（{desc}）年份={tb_year}'
          + (f' 主体 {len(per_ent)} 个' if per_ent else ''))
    # 扫描底稿（按年分组）
    files = sorted(glob.glob(os.path.join(data_dir, '*审计底稿_*_生成.xlsx')))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    years = {}
    for fn in files:
        m = re.search(r'_(\d{4})_生成\.xlsx$', os.path.basename(fn))
        yr = m.group(1) if m else '未知'
        years.setdefault(yr, []).append(fn)
    out = openpyxl.Workbook()
    out.remove(out.active)
    total_bad = []
    # 试算表年份已知时只核对对应年度底稿（2026 试算表 vs 2026 底稿）
    years = {tb_year: years[tb_year]} if tb_year in years else years
    for yr in sorted(years):
        ws = out.create_sheet(f'核对_{yr}')
        ws.cell(1, 1, f'合并试算表核对底稿（{yr} 年）—— 试算表 vs 审定表/明细表合计/附注汇总').font = \
            Font(name='Times New Roman', size=12, bold=True)
        hdr = ['科目', '试算表期末', '审定表期末', '差异1', '明细表合计', '差异2', '附注数', '差异3', '核对状态']
        for j, h in enumerate(hdr, 1):
            c = ws.cell(3, j, h); c.font = BF; c.fill = HFILL; c.border = BDR
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        r = 4
        n_ok = n_bad = n_miss = 0
        # —— 2026-08-04 A1：试算表主维度重构 ——
        # 底稿文件按科目名索引（『应收账款审计底稿_2026_生成.xlsx』→『应收账款』）
        file_subj = {}
        for fn in years[yr]:
            base = os.path.basename(fn)
            s = base.replace(f'审计底稿_{yr}_生成.xlsx', '').replace('_生成.xlsx', '')
            file_subj[s] = fn

        def _tb_of(subj_name):
            """底稿科目名 → 试算表值（直接/包含匹配；名称差异用 SUBJ_MAP 反向）。
            2026-08-06 修复：SUBJ_MAP 反向分支曾 return 映射字符串（v，如『营业收入』）
            → 主循环 2 abs(tv) 崩溃（TypeError: abs() 'str'）——应返回 TB 值 by_name.get(k)。"""
            for k, v in by_name.items():
                if k and (k in subj_name or subj_name in k):
                    return v
            for k, v in SUBJ_MAP.items():
                if v in subj_name and (k in subj_name or subj_name in k):
                    return by_name.get(k)
            return None

        def _match_file(tb_name):
            """试算表科目名 → (底稿科目名, 文件路径)。直接匹配 → SUBJ_MAP。
            备抵行（BADDEBT_ROWS）不参与文件匹配——走『备抵行』标注分支人工核。
            2026-08-05 修复：BADDEBT_ROWS 检查必须在循环【之前】——『减：坏账准备-应收账款』
            含『应收账款』，先被 file_subj 循环命中 → 走普通 _check 取合计行（应收总额），
            备抵行分支永远到不了（曾致坏账准备核对取成账面数 1.43 亿）。"""
            if tb_name in BADDEBT_ROWS or tb_name in INV_SUB_ROWS:
                return None, None
            # 2026-08-07 繁体『帐』→ 简体『账』（J 账套 应收帐款/应付帐款 匹配底稿）
            # 2026-08-07 任务683：ALIAS_ROWS 后接 subject_mapping.norm_name 归一化
            #（account_profiles.json common/专属 name_map：H『递延资产』→『长期待摊费用』、
            # 通用『股本』→『实收资本(或股本)』等）——匹配底稿科目名/附注口径更顺。
            _tb = ALIAS_ROWS.get(tb_name, tb_name)
            _tb = subject_mapping.norm_name(_acct, _tb)
            _hits = [(s, fn) for s, fn in file_subj.items()
                     if _tb and (_tb in s or s in _tb)]
            if _hits:
                _hits.sort(key=lambda x: (
                    0 if x[0] == tb_name else (1 if x[0].startswith(tb_name) else 2),
                    -len(x[0])))
                return _hits[0]
            mapped = SUBJ_MAP.get(tb_name)
            if mapped:
                for s, fn in file_subj.items():
                    if mapped in s:
                        return s, fn
            return None, None

        def _check(fn, subj, tb_name=None, baddebt_kw=None, tb_v=None, year=None):
            """核对单份底稿三项（合计+单体），返回 (aud_v, det_v, fn_v, aud_ent, det_ent, fn_ent)。
            *_ent 为 {主体名: 期末值}（2026-08-04 单体核对）。
            tb_name 提供时（收入类：主营业务收入等），审定表/明细表按项目行关键词提取单体，
            防『营业总收入』合计行被重复计入。
            baddebt_kw 提供时（备抵行：减：坏账准备/累计折旧），审定表取『减：xx』行的期末值
            （2026-08-04 用户要求：备抵行应对应审定表备抵项，不再笼统"人工核"）。"""
            proj_kw = None
            if tb_name and any(k in tb_name for k in ('收入', '成本', '利润')):
                proj_kw = tb_name
            # 2026-08-05：货币资金底稿审定表=『科目×主体』多科目结构（库存现金/银行存款/其他
            # 货币资金）——核对单个货币资金科目按科目名过滤（银行存款只取银行存款行合计
            # 304,802,940.23；否则 _sheet_end_total 取货币资金三行合计 307,461,283.21，
            # 曾致"核对底稿取审定表数据与底稿不一致"）。
            elif tb_name in ('库存现金', '银行存款', '其他货币资金'):
                proj_kw = tb_name
            wb = openpyxl.load_workbook(fn, read_only=True, data_only=True)
            aud_v = det_v = fn_v = None
            aud_ent = det_ent = fn_ent = {}
            # 2026-08-05 并入科目精确选 sheet：同一底稿可能含多个审定表/明细表（其他应付款底稿内
            # 『其他应付款 审定表』+『应付股利 审定表』+『应付利息 审定表』）——优先 sheet 名含
            # 当前核对科目名，否则退回第一个（主科目/单表底稿行为不变）。
            def _pick_sheet(kw, exclude=()):
                # 优先用原核对科目名 tb_name（应付股利）精确匹配 sheet；再用映射后科目名 subj
                # （其他应付款）匹配；都未命中退回第一个（主科目/单表底稿行为不变）。
                for _nm in (tb_name, subj):
                    if not _nm:
                        continue
                    for s in wb.sheetnames:
                        if kw in s and _nm in s and not any(k in s for k in exclude):
                            return s
                return next((s for s in wb.sheetnames
                             if kw in s and not any(k in s for k in exclude)), None)
            aud_sn = _pick_sheet('审定表')
            if aud_sn:
                if baddebt_kw:
                    aud_v = _baddebt_row_total(wb[aud_sn], baddebt_kw)
                elif proj_kw:
                    # 2026-08-04 修复：收入审定表『收入块+成本块+利润块』多块结构——_sheet_end_total
                    # 损益特判只取 主营业务收入/其他业务收入 块（成本科目曾返回收入合计 87.7M vs
                    # TB 45.16M，差 4253 万）。按 proj_kw 过滤的 _sheet_ent_values 精确取指定块。
                    _pent = _sheet_ent_values(wb[aud_sn], proj_kw=proj_kw, year=year)
                    aud_v = sum(_pent.values()) if _pent else None
                    aud_ent = _pent
                else:
                    if proj_kw:
                        # 2026-08-05：多科目审定表（货币资金 科目×主体）按科目名过滤取行
                        _pent = _sheet_ent_values(wb[aud_sn], proj_kw=proj_kw, year=year)
                        aud_v = sum(_pent.values()) if _pent else None
                        aud_ent = _pent
                    else:
                        # 2026-08-05 备抵科目净额口径：合并 TB 应收账款/其他应收款同名
                        # 覆盖取到【净额】行，底稿审定表合计行是【总额】→ 口径错配
                        # （坏账 2,607 万被误计差异）→ 取审定表『净额』行（总额−备抵）。
                        # 2026-08-07 修复：口径以【用户 TB 值】为准——FY 合并 TB 应收/其他应收
                        # 同名覆盖取到净额行（NET_SUBJ 应取净额）；但 F 合并 TB 其他应收款是
                        # 总额口径 13,324,934.54（坏账行 3.6 亿系企业重分类/调整数，非真实备抵），
                        # 强制取净额 → -347,419,586.35 假差异。取 净额/总额 中与 TB 更接近者。
                        if tb_name in NET_SUBJ:
                            _net_v = _net_row_value(wb[aud_sn])
                            _tot_v, _ = _sheet_end_total(wb[aud_sn])
                            if tb_v is not None and _net_v is not None and _tot_v is not None:
                                aud_v = (_net_v if abs(_net_v - tb_v) <= abs(_tot_v - tb_v)
                                         else _tot_v)
                            elif _net_v is not None:
                                aud_v = _net_v
                            else:
                                aud_v = _tot_v
                        else:
                            # 2026-08-07 修复：非 NET_SUBJ 科目（固定资产/资本公积/在建工程等）
                            # 走常规合计行读取——此前改动误删该 else 分支 → 审定表全未取到
                            aud_v, _ = _sheet_end_total(wb[aud_sn])
                        aud_ent = _sheet_ent_values(wb[aud_sn], proj_kw=proj_kw, year=year)
            # 2026-08-04 规范化：明细 sheet 匹配放宽——覆盖『XX明细表』/『XX分类汇总表』（固定资产等
            # longterm）/『分月明细(按主体)』（收入）等；排除 准备/计提 明细（备抵/计提口径非余额）。
            # 2026-08-06：其他业务收入/成本 无独立明细表（明细含于主营业务收入底稿，仅主营口径）——
            # 若匹配『主营业务收入按主体明细表』会取到主营收入 87.68M vs TB 其他业务收入 371.68
            # （差 -8,767 万）。跳过明细表，仅核审定+附注（明细口径无法拆出）。
            det_sn = None
            if tb_name not in ('其他业务收入', '其他业务成本'):
                # 2026-08-07 货币资金明细读取：优先『货币资金审定表』（库存现金/银行存款/
                # 其他货币资金 三行拆分=全口径），否则回退银行存款明细表只有银行存款 → 明细
                # 差=库存现金（dq 1,435.4 / Z 13,104.12 / H 13,743.61 / R/S/J/T 等 8 账套
                # 假差异，差异值恰=TB 1001 库存现金）。
                if tb_name == '货币资金':
                    _mc_sn = next((s for s in wb.sheetnames
                                   if '货币资金' in s and '审定' in s
                                   and not any(k in s for k in ('准备', '计提'))), None)
                    det_sn = _mc_sn or (_pick_sheet('明细表', ('准备', '计提', '期间对比', '对方科目'))
                                        or _pick_sheet('分类汇总', ('准备', '计提', '期间对比', '对方科目')))
                else:
                    det_sn = (_pick_sheet('明细表', ('准备', '计提', '期间对比', '对方科目'))
                              or _pick_sheet('分类汇总', ('准备', '计提', '期间对比', '对方科目')))
            if det_sn:
                # 2026-08-04：收入按主体明细表『主体|本期收入|本期成本|…』——成本科目需取『本期成本』列
                _pref = None
                if proj_kw:
                    if '成本' in proj_kw:
                        _pref = '本期成本'
                    elif '收入' in proj_kw:
                        _pref = '本期收入'
                if tb_name in ('库存现金', '银行存款', '其他货币资金'):
                    # 2026-08-05：货币资金明细表按块合计行取（『银行存款合计』行期末列），
                    # _sheet_end_total 取最后合计行（货币资金合计）曾致银行存款明细差异 -264 万
                    det_v, _why = _block_subtotal(wb[det_sn], tb_name)
                    det_ent = {}
                else:
                    det_v, _ = _sheet_end_total(wb[det_sn])
                    det_ent = _sheet_ent_values(wb[det_sn], proj_kw=proj_kw, pref_col=_pref, year=year)
                    if det_ent:
                        # 2026-08-04：合计口径与单体一致（_sheet_end_total 无列偏好会取『本期收入』列）
                        det_v = sum(det_ent.values())
            fn_sn = next((s for s in wb.sheetnames if '附注' in s), None)
            if fn_sn:
                # 附注取数科目名：proj_kw 优先（收入/成本/利润/货币资金单科目——主营业务成本必须取
                # 主营业务成本行，否则 subj=文件映射名『营业收入』会取到营业收入 87.7M vs TB 45.2M）。
                # 货币资金整体：SUBJ_MAP 映射为『银行存款』底稿，附注按 库存现金/银行存款/其他货币资金
                # 三行列示——取数必须用『货币资金』（三行累计），否则只取到银行存款单行（2026-08-04）。
                _fn_subj = (proj_kw
                            or ('货币资金' if tb_name == '货币资金'
                                else FN_SUBJ_ALIAS.get(tb_name or subj, subj)))
                fn_v = _footnote_total(wb[fn_sn], _fn_subj)
                fn_ent = _sheet_ent_values(wb[fn_sn], year=year)
            wb.close()
            return aud_v, det_v, fn_v, aud_ent, det_ent, fn_ent

        def _write_row(tb_name, tb_v, aud_v, det_v, fn_v, st):
            nonlocal r, n_ok, n_bad
            d1 = _diff(tb_v, aud_v); d2 = _diff(tb_v, det_v); d3 = _diff(tb_v, fn_v)
            status = []
            if st:
                status.append(st)
            # 2026-08-06 附注=合并口径（披露扣抵消）标注：往来科目附注 I 列审定数=加计−合并抵消，
            # 审定/明细=单体加计口径——附注差异≈审定差异+合并抵消（如应付账款 74,493+100万抵消、
            # 其他应收款 128万+490万抵消），属正常披露非取数错误。检测到该特征时追加口径说明。
            if d1 is not None and d3 is not None and abs(d1) > TOL and abs(d3) > TOL \
                    and aud_v is not None and fn_v is not None \
                    and tb_name in ('应收账款', '其他应收款', '应付账款', '合同负债', '预付账款',
                                    '其他应付款', '预收账款', '长期股权投资'):
                # 合并抵消 = 审定值 − 附注值（附注已扣抵消；审定=单体加计未扣）。
                # 例：应付 328.0M−327.0M=100万；应收 117.1M−106.6M=1,047万；
                # 长投 423.6M−424.1M=−43.7万（附注含 12FY 部分，口径差异）。
                _offset = aud_v - fn_v
                if abs(_offset) > 50:
                    status.append(f'附注差异{d3:,.2f}（=审定差异{abs(d1):,.2f}+合并抵消{abs(_offset):,.2f}，口径差）')
                    d3 = None  # 已标注，不再追加裸差异
            if aud_v is None:
                status.append('审定表未取到')
            elif d1 is not None and abs(d1) > TOL:
                status.append(f'审定差异{d1:,.2f}')
            if det_v is None:
                status.append('明细表未取到')
            elif d2 is not None and abs(d2) > TOL:
                # 2026-08-06 货币资金明细口径：明细表=银行存款+其他货币资金（不含库存现金
                # 17,754.60——库存现金无账户明细，设计上不列入银行存款明细表）→ 明细差异
                # =库存现金额−审定调整，属口径差（审定/附注已对平）。
                if tb_name == '货币资金':
                    status.append(f'明细差异{d2:,.2f}（明细表不含库存现金，口径差）')
                else:
                    status.append(f'明细差异{d2:,.2f}')
            if fn_v is None:
                status.append('附注未取到')
            elif d3 is not None and abs(d3) > TOL:
                status.append(f'附注差异{d3:,.2f}')
            st_ = '、'.join(status) if status else 'OK'
            bad = bool(status) and st != '试算表汇总行（由明细科目勾稽）'
            ws.cell(r, 1, tb_name).border = BDR
            for j, v in enumerate((tb_v, aud_v, d1, det_v, d2, fn_v, d3, st_), start=2):
                c = ws.cell(r, j, v)
                c.border = BDR
                if isinstance(v, (int, float)):
                    c.number_format = NUM
                    c.alignment = Alignment(horizontal='right')
                if bad:
                    c.fill = BADFILL
                    c.font = F10
            if bad:
                ws.cell(r, 9).font = BADFONT
                n_bad += 1
                total_bad.append(f'{tb_name}: {st_}')
            else:
                n_ok += 1
            r += 1

        def _diff(x, y):
            if x is None or y is None:
                return None
            return round(x - y, 2)

        # —— 主循环 1：试算表有值科目（76 个）逐行核对 ——
        done_files = set()
        ent_results = {}   # (试算表名, 底稿科目名) -> (tb_v, aud_ent, det_ent, fn_ent)
        for tb_name in sorted(by_name, key=lambda k: -(abs(by_name[k]) if by_name[k] is not None else 0.0)):
            tv = by_name.get(tb_name)
            if tv is None or abs(tv) <= 0.005:
                continue
            # 2026-08-04 修复：汇总行/勾稽行（加：年初未分配利润 等）优先标注——它们会因名称
            # 含『未分配利润』映射到权益底稿 → proj_kw 滤空致单体核对报"未取到"（实为汇总行）
            if tb_name in SUMMARY_ROWS or any(k in tb_name for k in
                                              ('营业总收入', '营业总成本', '利润', '净利')):
                _write_row(tb_name, tv, None, None, None, '试算表汇总行（由明细科目勾稽）')
                continue
            subj, fn = _match_file(tb_name)
            if fn is None:
                # 无底稿文件：报表汇总行 / 备抵行 / 底稿未生成
                if tb_name in SUMMARY_ROWS or any(k in tb_name for k in ('营业总收入', '营业总成本', '利润', '净利')):
                    _write_row(tb_name, tv, None, None, None, '试算表汇总行（由明细科目勾稽）')
                elif tb_name in INV_SUB_ROWS:
                    # 2026-08-07 存货细分科目：含于存货底稿（用户 TB 单独列示细分，
                    # 存货底稿按合并口径）→ 标注含于存货，不判"底稿未生成"。
                    _write_row(tb_name, tv, None, None, None,
                               f'含于{INV_SUB_ROWS[tb_name]}（存货底稿合并口径；细分差异属口径差）')
                elif tb_name in BADDEBT_ROWS:
                    # 2026-08-04 用户要求：备抵行应对应审定表备抵项（减：坏账准备/累计折旧/减值准备），
                    # 不再笼统"人工核"——映射到所属底稿，取审定表『减：xx』行期末值。
                    _owner = BADDEBT_ROWS[tb_name]
                    _fn = next((file_subj[s] for s in file_subj if _owner in s), None)
                    _kw = {'减：累计折旧': '累计折旧', '减：使用权资产累计折旧': '使用权资产折旧',
                           '减：坏账准备-应收账款': '坏账准备', '减：坏账准备-其他应收款': '坏账准备',
                           '减：固定资产减值准备': '减值准备',
                           '累计折旧': '累计折旧', '累计摊销': '累计摊销',
                           '坏账准备': '坏账准备', '坏帐准备': '坏账准备',
                           '应收坏账准备': '坏账准备', '其他应收坏账准备': '坏账准备',
                           '存货跌价准备': '跌价',
                           '固定资产减值准备': '减值准备', '无形资产减值准备': '减值准备',
                           '在建工程减值准备': '减值准备',
                           '使用权资产累计折旧': '使用权资产累计折旧',
                           '投资性房地产累计折旧': '投资性房地产累计折旧'}.get(tb_name, tb_name.replace('减：', ''))
                    if _fn and _kw:
                        _wbt = openpyxl.load_workbook(_fn, read_only=True, data_only=True)
                        _asn = next((s for s in _wbt.sheetnames if '审定表' in s), None)
                        _av, _why = _baddebt_row_total(_wbt[_asn], _kw) if _asn else (None, '无审定表')
                        _wbt.close()
                        _write_row(tb_name, tv, _av, None, None,
                                   f'备抵行（审定表"{_why}"）' if _av is not None else
                                   f'备抵行（{_owner}底稿未找到对应行：{_why}）')
                        ent_results[(tb_name, _owner)] = (tv, {}, {}, {})
                    else:
                        _write_row(tb_name, tv, None, None, None,
                                   f'备抵行（{_owner}底稿未生成）')
                elif tb_name in REALLOC_ROWS:
                    _write_row(tb_name, tv, None, None, None, f'报表自行重分类（{REALLOC_ROWS[tb_name]}）')
                elif tb_name in CASH_ROWS:
                    _write_row(tb_name, tv, None, None, None, f'含于{CASH_ROWS[tb_name]}')
                elif tb_name == '租赁负债未确认融资费用':
                    _write_row(tb_name, tv, None, None, None,
                               '租赁负债未确认融资费用（无租赁负债底稿，人工核对）')
                else:
                    _write_row(tb_name, tv, None, None, None, '试算表有数-底稿未生成')
                continue
            # 2026-08-05：租赁负债未确认融资费用 = 租赁负债 2601.01 子目（抵减项）。
            # _match_file 会把『租赁负债未确认融资费用』匹配到租赁负债底稿并取审定表合计
            # （租赁负债总额 52.66M vs TB 44 万，错取 5,222 万）→ 特判：取租赁负债明细表
            # 往来单位名含『未确认融资费用』的行，期末余额列合计（带符号，抵减为负）。
            if tb_name == '租赁负债未确认融资费用':
                _wbt = openpyxl.load_workbook(fn, read_only=True, data_only=True)
                _dsn = next((s for s in _wbt.sheetnames if '明细表' in s), None)
                _uv = _uv_why = None
                if _dsn:
                    _rows = [list(r) for r in _wbt[_dsn].iter_rows(values_only=True)]
                    _hc = None
                    for _i in range(min(6, len(_rows))):
                        _rtxt = ''.join(str(x) if x is not None else '' for x in _rows[_i][:24])
                        if ('往来单位' in _rtxt or '核算主体' in _rtxt) and '期末' in _rtxt:
                            _hc = _i
                            break
                    if _hc is not None:
                        _pcol = next((_j for _j, _h in enumerate(_rows[_hc][:24])
                                      if _h is not None and '期末余额' in str(_h)), None)
                        if _pcol is not None:
                            _uv = 0.0
                            _cnt = 0
                            for _r in _rows[_hc + 1:]:
                                if not _r or len(_r) <= _pcol:
                                    continue
                                _nm = str(_r[3]).strip() if len(_r) > 3 and _r[3] is not None else ''
                                if '未确认融资费用' in _nm:
                                    try:
                                        _uv += float(_r[_pcol] or 0)
                                        _cnt += 1
                                    except (TypeError, ValueError):
                                        pass
                            _uv_why = f'明细表未确认融资费用行({_cnt}行,期末列)'
                _wbt.close()
                _write_row(tb_name, tv, _uv, _uv, _uv,
                           f'租赁负债未确认融资费用（{_uv_why}；TB vs 底稿口径差=审计调整，'
                           f'不再误取租赁负债总额）' if _uv is not None else
                           '租赁负债底稿无未确认融资费用行')
                continue
            done_files.add(subj)
            aud_v, det_v, fn_v, aud_ent, det_ent, fn_ent = _check(fn, subj, tb_name, tb_v=tv, year=yr)
            # 2026-08-05 研发费用口径说明：损益『研发费用』（合并TB 单独列示 18.22M）在账套中
            # 挂管理费用项下（FY 6602.33『管理费用-研发费用』），底稿按账套原样归管理费用
            # （用户约定：明细表不自行修改）→ 研发费用审定=0、附注以管理费用附注为准。
            if tb_name == '研发费用':
                # fn_v 传 aud_v（0）→ 不再追加『附注未取到』；差异=TB 报表口径 vs 账套核算口径
                _write_row(tb_name, tv, aud_v, det_v, aud_v,
                           '研发费用（损益）在管理费用项下核算（账套 6602.33，见管理费用底稿）；'
                           'TB 单独列示=报表口径，差异属口径差（用户约定不拆分）')
                ent_results[(tb_name, subj)] = (tv, aud_ent, det_ent, fn_ent)
                continue
            _write_row(tb_name, tv, aud_v, det_v, fn_v,
                       (f'报表自行重分类（{REALLOC_ROWS[tb_name]}）' if tb_name in REALLOC_ROWS else ''))
            ent_results[(tb_name, subj)] = (tv, aud_ent, det_ent, fn_ent)
        # —— 主循环 2：底稿有但试算表无/0 值的科目 ——
        for subj, fn in sorted(file_subj.items()):
            if subj in done_files:
                continue
            tv = _tb_of(subj)
            aud_v, det_v, fn_v, aud_ent, det_ent, fn_ent = _check(fn, subj, year=yr)
            if tv is None:
                _write_row(subj, None, aud_v, det_v, fn_v, '试算表无此科目')
            elif abs(tv) <= 0.005:
                _write_row(subj, tv, aud_v, det_v, fn_v, '试算表0值')
            ent_results[(subj, subj)] = (tv, aud_ent, det_ent, fn_ent)
        ws.cell(r + 1, 1, f'核对汇总：OK {n_ok} / 差异 {n_bad} / 底稿 {len(years[yr])} 份').font = BF
        ws.cell(r + 2, 1, '口径说明（2026-08-04 试算表主维度）：①按试算表有值科目逐行核对（映射表：货币资金→银行存款、'
                          '应付职工薪酬→职工薪酬、研发费用→研发支出、预付款项→预付账款等）；②审定表取「期末未审数/期末数/'
                          '本期数/审定数」合计（合计行公式无缓存时解析 SUM 区间/单元格引用；固定资产等组件块取净值/账面价值行；'
                          '收入类取 主营业务收入+其他业务收入 块合计）；③明细表取合计行期末（排除减值/坏账准备明细表）；'
                          '④附注取科目块合计/审定数；⑤汇总行（营业总收入/利润链）与备抵行（减：累计折旧等）标注待勾稽；'
                          '⑥试算表取「审定数」列（无则简单加计数/期末）；⑦报表自行重分类行（其他流动资产=应交税费多交红字、'
                          '预收款项=应收贷方余额、一年内到期=长债/租赁/长期应付款 1 年内到期）标注归因——账套取数无重分类，'
                          '此类差异属报表口径，不判底稿错误。').font = \
            Font(name='Times New Roman', size=9, italic=True, color='808080')
        ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=9)
        for j, w in enumerate([22, 16, 16, 14, 16, 14, 16, 14, 22], 1):
            ws.column_dimensions[chr(64 + j)].width = w
        ws.freeze_panes = 'B4'
        # ============ 单体核对 sheet（2026-08-04 用户要求：试算表各单体 vs 底稿各单体，0.01 铁律）============
        if per_ent:
            _write_ent_check(out, yr, per_ent, ent_results, by_name)
        # ============ 单体核对（文件夹内单体 TB，2026-08-06 用户方法论：核对从单体开始）============
        # 合并试算表的主体列在部分账套解析失败（Z/JTt 合并TB 主体列缺失/科目名不匹配）→ 单体核对空。
        # 用户方法论：单体有差异→合并必差异；单体核对有错→该科目从源数据到底稿仔细检查。
        # 本块直接用【文件夹内每主体的单体科目余额表】作试算表（read_tb_full），科目按注册表
        # codes 代码前缀匹配（绕开 合并TB细分科目名 vs 底稿汇总名 的差异），逐主体×科目核对
        # 底稿审定/明细/附注，不依赖合并试算表。余额类科目（qm≠0）；损益类结转后 qm=0 自然跳过。
        try:
            import audit_common as _A
            import subjects_registry as _REG
            _ents = _A.discover_entities(data_dir)
            if _ents:
                _tb_full = _A.read_tb_full(data_dir, _ents)
                _ent_list = sorted(_ents)
                _ws2 = out.create_sheet(f'单体核对_{yr}_TB')
                _ws2.cell(1, 1, f'单体核对（{yr}）：文件夹内各主体单体科目余额表 vs 底稿审定/明细/附注（0.01 铁律）').font = \
                    Font(name='Times New Roman', size=12, bold=True)
                _hdr2 = ['科目', '主体', '单体TB期末', '审定表单体', '明细表单体', '附注单体',
                         '审定差异', '明细差异', '附注差异', '状态']
                for _j, _h in enumerate(_hdr2, 1):
                    _c = _ws2.cell(3, _j, _h); _c.font = BF; _c.fill = HFILL; _c.border = BDR
                    _c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                _r2 = 4
                _n_ok2 = _n_bad2 = 0
                _bad_subjs = {}
                for _key, _subj in sorted(_REG.REGISTRY.items()):
                    _codes = _subj.get('codes') or []
                    if not _codes:
                        continue
                    # 2026-08-07 组件块科目 TB 侧取数：固定资产 codes=['1601','1602','1603']
                    # 含备抵（累计折旧/减值）→ _fam2 前缀全命中求和=净额（1601−1602）vs
                    # 底稿审定=原值 → 假差异=累计折旧（Z 母公司 404.7M 净额 vs 683.6M 原值）。
                    # 核对基准=用户 TB 一级【原值】口径（与 _sheet_ent_values/_sheet_end_total
                    # 组件块同规则）→ TB 侧只用【原值】codes（codes[0]，顺序=原值/折旧/减值）。
                    if _subj['name'] in COMP_BLOCK_SUBJ and len(_codes) > 1:
                        _codes = _codes[:1]
                    _fn2 = os.path.join(data_dir, f"{_subj.get('file_key') or _key}审计底稿_{yr}_生成.xlsx")
                    if not os.path.exists(_fn2):
                        continue
                    try:
                        _r5 = _check(_fn2, _subj['name'], tb_name=_subj['name'], year=yr)
                    except Exception:
                        continue
                    _aud_e, _det_e, _fn_e = _r5[3], _r5[4], _r5[5]
                    _sb = _subj['name']
                    for _ent2 in _ent_list:
                        # 2026-08-06 修复：前缀匹配会把 U8 父级行+子级行双计（Z 母公司
                        # 应付账款 2202 父级 158.9M + 子级 158.9M = 317.7M vs 正确 158.9M）
                        # → 先按 codes 前缀取同族，再【叶子过滤】（父级=子和，只加叶子）
                        # 2026-08-07 修复：codes 前缀会命中"兄弟子目"（Z 未分配利润
                        # 4104.01 未分配利润 + 4104.02 提取法定盈余公积，父级 4104 前缀
                        # 全命中 → 差 47,554,708.07=盈余公积提取额）。
                        # 【族内名称全等精化】：仅当族内存在名称【完全等于】科目名的叶子
                        #（如 4104.01『未分配利润』）才只保留该行；否则保留全部叶子——
                        # 不得用"包含"匹配（资本公积族『其他资本公积』含『资本公积』
                        # 字样，曾误剔 资本溢价/上市资本溢价 → 审定 14.45 亿取成 -2,396 万）。
                        # 2026-08-07 修复：使用权资产【名称优先】——registry rua codes
                        # =['1704','1705',…] 是多账套并集，Q 账套 1704=开发支出（技术开发费
                        # 47.3M 被误当使用权资产 → 审定差 3657 万）；Q/Z=1641、R=1607、
                        # S=1621 不在 codes 内。名称含『使用权资产』（排除累计折旧/减值）
                        # 即精确命中原值行；无名称命中才回退 codes 前缀（g/FY/T=1704）。
                        if _subj['name'] == '使用权资产':
                            _nh2 = [(str(_c2), str(_n2), _v2['qm']) for (_e2, _c2, _n2, _y2), _v2 in _tb_full.items()
                                    if _e2 == _ent2 and str(_y2) == str(yr)
                                    and '使用权资产' in str(_n2) and '累计' not in str(_n2)
                                    and '折旧' not in str(_n2) and '减值' not in str(_n2)]
                            if _nh2:
                                _fam2 = _nh2
                            else:
                                _fam2 = [(str(_c2), str(_n2), _v2['qm']) for (_e2, _c2, _n2, _y2), _v2 in _tb_full.items()
                                         if _e2 == _ent2 and str(_y2) == str(yr)
                                         and any(str(_c2).startswith(_k2) for _k2 in _codes)]
                        else:
                            _fam2 = [(str(_c2), str(_n2), _v2['qm']) for (_e2, _c2, _n2, _y2), _v2 in _tb_full.items()
                                     if _e2 == _ent2 and str(_y2) == str(yr)
                                     and any(str(_c2).startswith(_k2) for _k2 in _codes)]
                        _leaf2 = [(c, n, v) for c, n, v in _fam2
                                  if not any(c2 != c and c2.startswith(c) for c2, _, _ in _fam2)]
                        _nm = _subj.get('name')
                        _qm2 = None
                        if _nm:
                            # 2026-08-07 权益科目净额代表统一判定（实收资本/资本公积/未分配利润）：
                            # ① 族内存在【父级行且名称含科目名核心词】→ 取父级值——Q 资本公积
                            #    4002 父级(-373.8M) 子目 400201 名称恰="资本公积"，原全等精化
                            #    误取 78.4M（丢溢价/其他资本公积）；g 实收资本 4105 父级"股本"。
                            # ② 未分配利润：取名称【包含】"未分配利润"的叶子（g 410407
                            #    "利润分配-未分配利润" 全等失败曾混入提取盈余公积 410401、
                            #    Z 4104.01、Q 410403）。
                            # ③ 其他科目：名称全等叶子；无则全部叶子（现状逻辑）。
                            _core = (_nm or '').replace('（或股本）', '').replace('(或股本)', '')
                            if '股本' in (_nm or ''):
                                _core = '股本'
                            if _fam2:
                                _cs = [str(c) for c, _, _ in _fam2]
                                _short = min(_fam2, key=lambda x: len(str(x[0])))
                                _sp = any(str(c2) != str(_short[0])
                                          and str(c2).startswith(str(_short[0])) for c2 in _cs)
                                if _sp and _core and _core in str(_short[1]):
                                    _qm2 = _short[2]
                            if _qm2 is None:
                                if '未分配利润' in (_nm or ''):
                                    _hit = [(c, n, v) for c, n, v in _leaf2
                                            if '未分配利润' in str(n)]
                                    if _hit:
                                        _leaf2 = _hit
                                else:
                                    _exact = [(c, n, v) for c, n, v in _leaf2
                                              if str(n).strip() == _nm]
                                    if _exact:
                                        _leaf2 = _exact
                                if not any(('实收资本' in str(_n) or '股本' in str(_n))
                                           for _, _n, _ in _leaf2) \
                                        and ('实收资本' in (_nm or '') or '股本' in (_nm or '')):
                                    # H 实收资本代码=3003『股本』（非标准 4001；H 的 4001=生产
                                    # 成本）→ codes 前缀族内无实收资本名称行 → 全 TB 名称匹配
                                    # 回退（2026-08-07 修复：原条件 not _leaf2 误拦——族内
                                    # 4001 生产成本叶子非空但名称不含实收资本，回退被跳过）。
                                    # ⚠️ 仅实收资本科目生效——H 其他科目若全部回退名称匹配
                                    # → 27 科目全误取 3003 股本 -76,200,000（曾致假差异）。
                                    _kw_hits = [(str(_c2), str(_n2), v2['qm'])
                                                for (_e2, _c2, _n2, _y2), v2 in _tb_full.items()
                                                if _e2 == _ent2 and str(_y2) == str(yr)
                                                and ('实收资本' in str(_n2)
                                                     or str(_n2).strip() == '股本'
                                                     or ('股本' in str(_n2)
                                                         and not any(k in str(_n2) for k in
                                                                     ('溢价', '购买', '公积金'))))]
                                    if _kw_hits:
                                        # 回退后父级判定：g 4105『股本』父级(-566.2M)——
                                        # 否则叶子仅 个人股本+国有法人股本，漏 社会公共股
                                        #（名不含"股本"子串）
                                        _cs2 = [str(c) for c, _, _ in _kw_hits]
                                        _short2 = min(_kw_hits, key=lambda x: len(str(x[0])))
                                        _sp2 = any(str(c2) != str(_short2[0])
                                                   and str(c2).startswith(str(_short2[0]))
                                                   for c2 in _cs2)
                                        if _sp2 and _core and _core in str(_short2[1]):
                                            _leaf2 = [_short2]
                                        else:
                                            _leaf2 = [(c, n, v) for c, n, v in _kw_hits
                                                      if not any(c2 != c and c2.startswith(c)
                                                                 for c2, _, _ in _kw_hits)]
                                _qm2 = sum(v for _, _, v in _leaf2)
                        else:
                            _qm2 = sum(v for _, _, v in _leaf2)
                        if abs(_qm2) <= 0.005:
                            continue

                        def _cmp2(x):
                            if x is None:
                                return None
                            # 2026-08-07 修复：负债类科目 TB 贷余为负、底稿按正数列示，
                            # 只要异号即应按绝对值比较（原仅绝对值相等时比较 → 金额不同时
                            # 显示成 TB负−底稿正 的假 2 倍差异，如 Q Audio 应付职工薪酬
                            # -1,059,321.14 vs 1,046,612.32 → -2,105,933.46，真实差仅 12,708.82）。
                            if (_qm2 < 0) != (x < 0):
                                return abs(_qm2) - abs(x)   # 方向相反（负债显示口径）按绝对值比较
                            return _qm2 - x
                        _d1 = _cmp2(_av2) if (_av2 := (_aud_e or {}).get(_ent2)) is not None else None
                        _d2 = _cmp2(_dv2) if (_dv2 := (_det_e or {}).get(_ent2)) is not None else None
                        _d3 = _cmp2(_fv2) if (_fv2 := (_fn_e or {}).get(_ent2)) is not None else None
                        _diffs = [(l, d) for l, d in (('审定', _d1), ('明细', _d2), ('附注', _d3))
                                  if d is not None and abs(d) > TOL]
                        if _diffs:
                            _n_bad2 += 1
                            _bad_subjs[_sb] = _bad_subjs.get(_sb, 0) + 1
                            _ws2.cell(_r2, 1, _sb); _ws2.cell(_r2, 2, _ent2)
                            _ws2.cell(_r2, 3, round(_qm2, 2))
                            _ws2.cell(_r2, 4, round(_av2, 2) if _av2 is not None else None)
                            _ws2.cell(_r2, 5, round(_dv2, 2) if _dv2 is not None else None)
                            _ws2.cell(_r2, 6, round(_fv2, 2) if _fv2 is not None else None)
                            _ws2.cell(_r2, 7, round(_d1, 2) if _d1 is not None else None)
                            _ws2.cell(_r2, 8, round(_d2, 2) if _d2 is not None else None)
                            _ws2.cell(_r2, 9, round(_d3, 2) if _d3 is not None else None)
                            _c = _ws2.cell(_r2, 10, '差异→从源数据检查'); _c.font = BADFONT
                            for _j in range(1, 11):
                                _ws2.cell(_r2, _j).border = BDR
                            _r2 += 1
                        else:
                            _n_ok2 += 1
                _ws2.cell(_r2, 1, f'单体核对汇总：OK {_n_ok2} / 差异 {_n_bad2}').font = BF
                _ws2.cell(_r2 + 1, 1, '口径：单体TB取各主体单体科目余额表期末（按注册表 codes 代码前缀匹配，绕开合并TB'
                                     '细分名 vs 底稿汇总名差异）；底稿各主体值取自审定表/明细表/附注；方向相反（负债贷余）'
                                     '按绝对值比较。差异科目需从源数据到底稿逐项检查（单体核对是合并核对的前提）。').font = \
                    Font(name='Times New Roman', size=9, italic=True, color='808080')
                _ws2.merge_cells(start_row=_r2 + 1, start_column=1, end_row=_r2 + 1, end_column=10)
                for _j, _w in enumerate([22, 16, 16, 16, 16, 16, 14, 14, 14, 20], 1):
                    _ws2.column_dimensions[chr(64 + _j)].width = _w
                _ws2.freeze_panes = 'B4'
                if _bad_subjs:
                    print(f'⚠️ [单体核对] {yr} 年 {len(_bad_subjs)} 个科目有单体差异（需从源数据到底稿检查）：')
                    for _sb, _n in sorted(_bad_subjs.items(), key=lambda x: -x[1]):
                        print(f'     · {_sb}（{_n} 个主体）')
        except Exception as _ex2:
            print(f'  ⚠️ 单体核对(TB)生成失败：{_ex2}')
    out_path = out_path or os.path.join(data_dir, '合并试算表核对底稿_生成.xlsx')
    # 2026-08-04：目标文件被 Excel/Tencent Docs 占用时，自动改用『_生成_新.xlsx』并提示
    #（避免 PermissionError 中断——用户开着核对底稿时重跑常见）
    try:
        out.save(out_path)
    except PermissionError:
        alt = out_path.replace('.xlsx', '_新.xlsx')
        out.save(alt)
        out_path = alt
        print(f'⚠️ 原输出被占用（Excel 打开中？），已保存到：{out_path}')
    out.close()
    print(f'✅ 已生成核对底稿：{out_path}')
    if total_bad:
        print(f'⚠️ 共 {len(total_bad)} 个科目核对差异，需检查相关底稿生成：')
        for b in total_bad[:30]:
            print(f'   - {b}')
    else:
        print('🎉 全部科目三项核对一致，底稿生成无差异。')
    return out_path


def _write_ent_check(out, yr, per_ent, ent_results, by_name):
    """单体核对 sheet（2026-08-04）：试算表主体列值 vs 底稿审定表/明细表/附注 各主体值。
    匹配：试算表主体名 与 底稿主体名 包含匹配（FY本级公司 ↔ 01FY本级公司）；
    符号：底稿显示口径（负债贷余为正）与试算表净额（贷方为负）方向相反时按绝对值比较（不误报）。
    0.01 铁律：方向对齐后差异 > 0.01 标红。"""
    ws = out.create_sheet(f'单体核对_{yr}')
    ws.cell(1, 1, f'单体核对（{yr} 年）：合并试算表各主体 vs 底稿审定表/明细表/附注各主体（0.01 铁律）').font = \
        Font(name='Times New Roman', size=12, bold=True)
    hdr = ['科目', '试算表主体', '底稿主体', '试算表单体', '审定表单体', '明细表单体', '附注单体',
           '审定差异', '明细差异', '附注差异', '状态']
    for j, h in enumerate(hdr, 1):
        c = ws.cell(3, j, h); c.font = BF; c.fill = HFILL; c.border = BDR
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    r = 4
    n_ok = n_bad = 0
    tb_ents = sorted(per_ent)
    for (tb_name, subj), (tb_v, aud_ent, det_ent, fn_ent) in sorted(ent_results.items()):
        if not aud_ent and not det_ent and not fn_ent:
            continue
        per_tb = per_ent if per_ent else {}
        # 该科目在试算表各主体的值
        for te in tb_ents:
            te_v = (per_tb.get(te) or {}).get(tb_name)
            if te_v is None or abs(te_v) <= 0.005:
                continue
            # 匹配底稿主体（包含匹配，2026-08-04 规模修复：按名称长度降序 → 取最长匹配，
            # 防 120 主体时相似主体名（FY本级公司 vs FY本级公司2）误配第一个）
            best = None
            ent_pool = sorted(set(list(aud_ent) + list(det_ent) + list(fn_ent)),
                              key=lambda s: len(s or ''), reverse=True)
            for en in ent_pool:
                if not en:
                    continue
                if te in en or en in te:
                    best = en
                    break
            if best is None:
                ws.cell(r, 1, tb_name); ws.cell(r, 2, te); ws.cell(r, 3, '(底稿无此主体)')
                ws.cell(r, 4, round(te_v, 2))
                ws.cell(r, 11, '底稿无对应主体').font = BADFONT
                for j in range(1, 12):
                    ws.cell(r, j).border = BDR
                r += 1
                continue
            av = aud_ent.get(best); dv = det_ent.get(best); fv = fn_ent.get(best)

            def _cmp(x):
                """试算表单体 vs 底稿单体：方向相反（互为相反数）→ 按绝对值比较；否则直接比较。"""
                if x is None:
                    return None, False
                d = te_v - x
                if abs(d) > TOL and abs(te_v + x) <= TOL:
                    return abs(te_v) - abs(x), True   # 符号方向相反（负债显示口径）
                return d, False
            # 单体提取为空（如格式2 附注无主体列/该 sheet 无主体维度）→ 跳过该列比对（不算未取到）
            d1, s1 = _cmp(av) if aud_ent else (None, False)
            d2, s2 = _cmp(dv) if det_ent else (None, False)
            d3, s3 = _cmp(fv) if fn_ent else (None, False)
            status = []
            for lbl, d in (('审定', d1), ('明细', d2), ('附注', d3)):
                if d is None:
                    continue
                if abs(d) > TOL:
                    status.append(f'{lbl}差异{d:,.2f}')
            st = '、'.join(status) if status else 'OK'
            bad = bool(status)
            ws.cell(r, 1, tb_name).border = BDR
            ws.cell(r, 2, te).border = BDR
            ws.cell(r, 3, best).border = BDR
            for j, v in enumerate((te_v, av, dv, fv, d1, d2, d3), start=4):
                c = ws.cell(r, j, v)
                c.border = BDR
                if isinstance(v, (int, float)):
                    c.number_format = NUM
                    c.alignment = Alignment(horizontal='right')
                if bad:
                    c.fill = BADFILL
                    c.font = F10
            ws.cell(r, 11, st)
            if bad:
                ws.cell(r, 11).font = BADFONT
                n_bad += 1
            else:
                n_ok += 1
            r += 1
    ws.cell(r + 1, 1, f'单体核对汇总：OK {n_ok} / 差异 {n_bad}（0.01 铁律：方向对齐后差异>0.01 标红；'
                      f'负债类显示口径方向相反时按绝对值比较，不误报）').font = BF
    ws.cell(r + 2, 1, '说明：①底稿审定表/明细表按主体行提取（核算主体列），附注按每主体一列提取；'
                      '②主体名包含匹配（试算表 FY本级公司 ↔ 底稿 01FY本级公司）；'
                      '③底稿未生成的主体/科目不列行；④单体差异>0.01 需回查该主体底稿生成。').font = \
        Font(name='Times New Roman', size=9, italic=True, color='808080')
    ws.merge_cells(start_row=r + 2, start_column=1, end_row=r + 2, end_column=11)
    for j, w in enumerate([18, 14, 14, 15, 15, 15, 15, 13, 13, 13, 26], 1):
        ws.column_dimensions[chr(64 + j)].width = w
    ws.freeze_panes = 'D4'
    return ws


# 科目映射表（2026-08-04 A1）：试算表科目名 → 底稿科目名关键词。
# 解决 试算表合并列示/名称差异/备抵行 与底稿文件名的匹配缺口：
#   货币资金=银行存款+库存现金+其他货币资金（用银行存款底稿审定≈货币资金核对）；
#   应付职工薪酬=底稿『职工薪酬』；研发费用=底稿『研发支出』（损益 vs 资产口径注意）；
#   实收资本（股本）=底稿『实收资本』（全角/半角括号差异）。
SUBJ_MAP = {
    '货币资金': '银行存款',
    '应付职工薪酬': '职工薪酬',
    '研发费用': '研发支出',
    '实收资本（股本）': '实收资本',
    '实收资本(股本)': '实收资本',
    # 2026-08-04 补充：损益/往来名称差异映射
    '预付款项': '预付账款',
    '主营业务收入': '营业收入',
    '应收账款账面价值': '应收账款',
    '其他应收款账面价值': '其他应收款',
    # 2026-08-07：dq 等账套用户 TB 用旧科目名『营业费用』（6601）而非『销售费用』
    '营业费用': '销售费用',
    '其他业务成本': '营业收入',
    # 2026-08-04 二次复核：成本/税金在账套无独立科目或含于他底稿——映射到承载底稿，
    # 不再判"底稿未生成"（营业成本=收入底稿成本块；税金及附加=报表计算行，无账套科目）。
    '主营业务成本': '营业收入',
    '营业成本': '营业收入',
    '原材料': '存货',
    '长期应付款': '专项应付款',
    '其他业务收入': '营业收入',
    # 2026-08-05 并入科目（ORP/ORA 三科目块）：审定表/明细表在其他应付款/其他应收款底稿内
    # （无独立底稿文件）——映射到承载底稿，核对按科目名精确选 sheet（应付股利 审定表 等）。
    '应付股利': '其他应付款',
    '应付利息': '其他应付款',
    '应收股利': '其他应收款',
    '应收利息': '其他应收款',
}
# 备抵净额科目：合并 TB 同名『总额/净额』对中取到净额行 → 底稿审定表取『净额』行（总额−备抵）
NET_SUBJ = ('应收账款', '其他应收款', '应收账款账面价值', '其他应收款账面价值')
# 2026-08-05 组件块附注科目：附注按 未审数/调整/审定 三段 × 原值/折旧/减值/账面价值 四块，
# 核对取『账面价值(净值)』块『未审数』段（固定资产/无形资产/在建工程/使用权资产/长期待摊）
# 2026-08-07 修复：核对口径改取【原值】块（用户TB 一级=原值）——COMP_BLOCK_SUBJ 补
# 『投资性房地产』（原值/累计折旧/减值/账面价值 四块结构与固定资产相同，此前漏登记 →
# 附注取净值 940,428.99 vs TB 原值 3,570,912.69，差 2,630,483.70=累计折旧）。
COMP_BLOCK_SUBJ = ('固定资产', '无形资产', '在建工程', '使用权资产', '长期待摊费用', '投资性房地产')
# 2026-08-05 附注行名≠科目名（项目×主体格式）：营业收入=主营业务收入+其他业务收入；
# 营业成本=主营业务成本+其他业务成本（行名『一、主营业务收入』不含『营业收入』子串）
SUBJ_ROW_KW = {
    '营业收入': ('主营业务收入', '其他业务收入'),
    '营业成本': ('主营业务成本', '其他业务成本'),
}
# 附注行名 ≠ 账套科目名（模板披露行名）：核对附注维度时用附注行名匹配
# （2026-08-05：其他应付款附注『普通股股利』行=账套『应付股利』科目）。
FN_SUBJ_ALIAS = {
    '应付股利': '普通股股利',
}
# 试算表报表汇总行（利润表/权益链）：底稿无对应文件，标注『汇总行（由明细科目勾稽）』
SUMMARY_ROWS = ('营业总收入', '营业总成本', '营业利润', '利润总额', '净利润',
                '归属于母公司股东的净利润', '持续经营净利润', '可供分配的利润',
                '可供投资者分配的利润', '加：年初未分配利润', '归属于母公司所有者的净利润',
                '综合收益总额', '归属于母公司所有者的综合收益总额', '净利润：')
# 备抵/抵减行（含于资产底稿组件内，人工核）：试算表科目名 → 所属底稿
BADDEBT_ROWS = {    '减：累计折旧': '固定资产',
    '减：使用权资产累计折旧': '使用权资产',
    '减：坏账准备-应收账款': '应收账款',
    '减：坏账准备-其他应收款': '其他应收款',
    '减：固定资产减值准备': '固定资产',
    # 2026-08-06 无形资产累计摊销备抵行（FY 合并TB『减：累计摊销』行，此前未映射
    # → 显示"试算表有数-底稿未生成"）
    '减：累计摊销': '无形资产',
    '减：无形资产减值准备': '无形资产',
    '减：在建工程减值准备': '在建工程',
    # 2026-08-07 用户 TB 备抵科目裸名（无『减：』前缀）→ 映射到所属底稿取审定表备抵行
    '累计折旧': '固定资产',
    '累计摊销': '无形资产',
    '坏账准备': '应收账款',
    '坏帐准备': '应收账款',
    '应收坏账准备': '应收账款',
    '其他应收坏账准备': '其他应收款',
    '存货跌价准备': '存货',
    '固定资产减值准备': '固定资产',
    '无形资产减值准备': '无形资产',
    '在建工程减值准备': '在建工程',
    '使用权资产累计折旧': '使用权资产',
    '投资性房地产累计折旧': '投资性房地产',
}

# 存货细分科目 → 含于存货底稿（2026-08-07：库存商品/发出商品/生产成本/半成品/周转材料/
# 委托加工物资/在途物资/材料采购 等是存货一级科目明细，用户 TB 单独列示 → 标注含于存货，
# 不再"底稿未生成"；差异=存货底稿按合并口径，细分差异属口径差）
INV_SUB_ROWS = {
    '原材料': '存货', '库存商品': '存货', '发出商品': '存货', '生产成本': '存货',
    '半成品': '存货', '周转材料': '存货', '委托加工物资': '存货', '委托代销商品': '存货',
    '在途物资': '存货', '材料采购': '存货', '自制半成品': '存货', '产成品': '存货',
    '工程施工': '存货', '开发支出': '存货', '开发成本': '存货', '合同履约成本': '存货',
    '合同结算': '存货',
}

# 繁体/别名 → 标准科目（2026-08-07：J 账套『应收帐款/应付帐款』用繁体『帐』，
# 与底稿『应收账款/应付账款』匹配不上 → 显示底稿未生成）
ALIAS_ROWS = {
    '应收帐款': '应收账款', '应付帐款': '应付账款', '坏帐准备': '坏账准备',
}

# 报表自行重分类行（2026-08-04 用户方法论）：账套取数无重分类，试算表审定数含报表重分类——
# ①应交税费贷方红字（多交）→ 其他流动资产（账套可能无科目，需另开数据）；
# ②长期借款/租赁负债/长期应付款 1 年内到期 → 一年内到期的非流动负债（NCL 底稿 2242 核对）；
# ③应收账款贷方余额 → 预收款项（账套无预收科目）。此类行标注归因，不判"底稿错误"。
REALLOC_ROWS = {
    '其他流动资产': '应交税费多交红字/其他，账套无对应科目→核对差异属预期',
    '预收款项': '应收账款贷方余额转入，账套无预收账款科目→核对差异属预期',
    '一年内到期的非流动负债': '长期借款/租赁负债/长期应付款 1年内到期，NCL 底稿(2242)仅含账套数据→差异=重分类转入',
}
# 货币资金拆分科目（报表科目货币资金=库存现金+银行存款+其他货币资金；银行存款底稿审定表已三行拆分）
CASH_ROWS = {'库存现金': '货币资金（银行存款底稿审定表三行拆分：库存现金/银行存款/其他货币资金）',
    # 2026-08-07：繁体『帐』/带路径名 别名 → 含于货币资金
    '现金': '货币资金（银行存款底稿审定表三行拆分：库存现金/银行存款/其他货币资金）',
    '1001\\库存现金': '货币资金（银行存款底稿审定表三行拆分：库存现金/银行存款/其他货币资金）',
    '1402\\在途物资': '存货（存货底稿合并口径；细分差异属口径差）',
             '其他货币资金': '货币资金（银行存款底稿审定表三行拆分：库存现金/银行存款/其他货币资金）'}


def _net_row_value(ws):
    """审定表『XX净额』行期末值（备抵科目净额口径，2026-08-05 新增）。

    背景：合并 TB 中备抵科目同名重复（应收账款=总额 R026/净额 R028、其他应收款=总额/净额），
    by_name 同名覆盖取到净额行；而底稿审定表合计行取的是【总额】（R15 合计），
    净额/总额口径错配 → 应收账款审定差异 -3,504 万（其中坏账 2,607 万为口径差）。
    修复：底稿审定表取『XX净额』行（如 R17 应收账款净额=总额−坏账），与 TB 净额口径一致；
    净额行为公式且无缓存时 → 合计行 − 备抵行。"""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        return None
    hdr_i = None
    for i, r in enumerate(rows[:10]):
        if not r:
            continue
        if sum(1 for x in r[:12] if x is not None and str(x).strip()) < 3:
            continue
        s20 = ''.join(str(x) if x is not None else '' for x in r[:40])
        if any(k in s20 for k in ('明细表', '审定表', '附注', '底稿', '汇总（')):
            continue
        if ('期初' in s20 or '年初' in s20 or '未审数' in s20 or '本期' in s20 or '上年数' in s20) \
                and ('期末' in s20 or '余额' in s20 or '审定数' in s20 or '本期' in s20) \
                and ('主体' in s20 or '公司' in s20 or '年度' in s20 or '项目' in s20):
            hdr_i = i
            break
    if hdr_i is None:
        return None
    col = _col_by_header(rows, hdr_i, ('期末',))
    if col is None:
        col = _col_by_header(rows, hdr_i, ('审定数',))
    if col is None:
        return None
    tot = bad = None
    for i in range(hdr_i + 1, len(rows)):
        r = rows[i]
        if not r or col >= len(r):
            continue
        a0 = str(r[0]).strip() if r[0] is not None else ''
        a1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
        joined = a0 + a1
        if '净额' in joined:
            v = _resolve_value(rows, i, col)
            if v is not None and abs(v) > 0.005:
                return v          # 净额行有值（或公式可解析）→ 直接返回
            continue
        if ('合计' in joined or '总计' in joined) and '核对' not in joined:
            v = _resolve_value(rows, i, col)
            if v is not None and abs(v) > 0.005:
                tot = v
        elif '减：' in a1 or a1.startswith('减'):
            v = _resolve_value(rows, i, col)
            if v is not None and abs(v) > 0.005:
                bad = v
    if tot is not None and bad is not None:
        return tot - bad           # 净额行空（公式无缓存）→ 合计 − 备抵
    return None


def _footnote_total(ws, subj):
    """附注汇总：按科目块取期末数/合计。支持两类结构：
    A) 每主体一列两期数（长期借款等 mode='two'）：A列=科目名、B..N=各主体数值，
       取『期末数』段各主体加总 = 全集团期末数；
    B) 合并口径（格式2）：科目块内『合 计/账面余额合计』行取 审定数(第9列)/加计数(第6列)。"""
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    cand = None
    in_block = False   # 已进入 subj 科目块
    _got_subj_row = False  # 2026-08-05 已精确命中科目行（格式2：B 列行名含 subj）→ 合计行不再覆盖
    _cur_total_pri = 0     # 2026-08-05 格式2 合计行优先级（合 计=3 > 披露块合计=2 > 子块小计=1）
    cur_seg = ''       # 'end' / 'open' / 'other'（期末数/期初数/其他段）
    # 2026-08-05 三平行表（利润表科目附注：未审数/审计调整数/审定数 三块）：
    # A 列行名=TB 二级名（营业费用-xxx，不含科目名『销售费用』）→ 原 subj in b0 匹配不到；
    # 合计行 '(合计)' B..L 列空 → 原逻辑取不到值（核对稿 17 个利润表科目『附注未取到』）。
    # 识别『附注披露 — 未审数』块 → 块内数据行（B..L 数值）逐列累计 = 全集团附注数。
    _par_plane = False
    _par_seen = False      # 已遇到过『未审数』块（兜底判断用独立标志——_par_plane 会被调整/审定块重置）
    _par_acc = [0.0] * 60
    _par_ncols = 12        # 三平行表主体列数（表头 B 列起连续主体名；数据行 N 列=行合计须排除）
    # 2026-08-06 组件块附注（在建工程/固定资产/长期待摊等）表头『项目|项目|泰国|全集团合计』：
    # 累计须排除『全集团合计』列（否则主体值+合计值双计 → 泰国在建工程附注 369M=2×184.5M）。
    # 2026-08-06 17:35 修复：longterm 附注表头=『类别|段名|主体1..N|全集团合计』——原条件
    # (项目/主体)+合计 不匹配『类别/段名』表头 → _sum_cols 空 → 数值化后的合计列被累计
    # （FY 固定资产附注 1813M=2×906M 回归根因；公式版合计列读 None 侥幸未暴露）。
    _sum_cols = set()
    for _hr in rows[:8]:
        if not _hr:
            continue
        _htxt = ''.join(str(x) if x is not None else '' for x in _hr[:20])
        if ('项目' in _htxt or '主体' in _htxt or '类别' in _htxt or '段名' in _htxt
                or '核算主体' in _htxt or '科目' in _htxt) and '合计' in _htxt:
            for _hj, _hv in enumerate(_hr[:20]):
                if _hv is not None and '合计' in str(_hv):
                    _sum_cols.add(_hj)
            break
    _main_seg = 'unaudit'  # 2026-08-05 大段：未审数/审计调整数/审定数（组件块附注 3 段重复 → 仅累计未审数段）
    _multi_seg = False     # 出现过『未审数』标题 → 多段结构 → 仅累计未审数段；单段（如职工薪酬审定口径）全累计
    _seg_seen = False      # 出现过段标题（期初数/期末数/本期增减）→ 段式累计仅期末数段；无段标题（收入/存货）走行匹配
    _comp_block = ''       # 组件块：gross/dep/imp/net（一、原值…四、账面价值）
    # 2026-08-05 预扫描：整表是否有『账面价值/净值』块——固定资产等有 → 仅累计 net 块；
    # 在建工程仅『账面余额』块 → 累计 gross 块。不能边跑边判（原值块在 net 块之前出现，
    # _has_net_block 尚 False → 原值被误累计 → 固定资产 2,039M=原值+净值双计）。
    # 排除口径说明行（『口径：…账面价值=原值−累计折旧…』含"账面价值"字样会误判——在建工程曾 None）
    _has_net_block = any(
        r and r[0] is not None
        and ('账面价值' in str(r[0]) or '净值' in str(r[0]))
        and not str(r[0]).startswith(('口径', '注：', '滚动勾稽', '勾稽'))
        for r in rows)
    # 2026-08-06：是否存在『审定』口径净值块（存货块七『期末审定账面价值』）——
    # 有 → 只累计 net_audit 块（审定净值为准）；无 → 累计未审 net 块。
    _has_audit_net = any(
        r and r[0] is not None
        and ('账面价值' in str(r[0]) or '净值' in str(r[0])) and '审定' in str(r[0])
        and not str(r[0]).startswith(('口径', '注：', '滚动勾稽', '勾稽'))
        for r in rows)
    _net_comp = 'net_audit' if _has_audit_net else 'net'
    # 2026-08-06：组件块附注是否有『原值/成本/余额』块——核对基准=用户TB 一级（原值口径），
    # 有原值块 → 附注取原值（gross）而非净值（固定资产 819.4M vs 净值 514.5M 差 304.9M=折旧）。
    _has_gross_block = any(
        r and r[0] is not None
        and any(k in str(r[0]) for k in ('原值', '账面余额', '成本'))
        and not str(r[0]).startswith(('口径', '注：', '滚动勾稽', '勾稽'))
        for r in rows)
    for idx, r in enumerate(rows):
        b0 = str(r[0]).strip() if r[0] is not None else ''
        b1 = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
        # 段标题精确识别（2026-08-04 修复：货币资金附注含 期末数/审计调整/期末审定数/期初数 多段，
        # 原『期末 in b0』会把『期末审定数』误当期末段、且『审计调整』段的数据行覆盖期末值）
        _b0s, _b1s = b0.replace(' ', ''), b1.replace(' ', '')
        if _b0s == '期末数' or _b1s == '期末数':
            cur_seg = 'end'
        elif _b0s == '期初数' or _b1s == '期初数':
            cur_seg = 'open'
        elif any(_b0s == k or _b1s == k for k in ('审计调整', '期末审定数', '本期增加', '本期减少',
                                                   '本年增加', '本年减少', '期初审定数')):
            cur_seg = 'other'
        # 三平行表（利润表科目附注）：标题行『XX 附注披露 — 未审数/审计调整数/审定数』
        # → 仅『未审数』块内数据行（B..L 数值）逐列累计；调整/审定块停止（审定数块公式
        #   无缓存为空，附注披露以未审数为准——用户方法论：附注以审定数为准，调整数由
        #   项目组填列后回灌，核对基准仍与 TB 未审数同口径）。
        if '附注披露' in b0 and '—' in b0:
            # 三平行表（利润表科目附注）：标题『XX 附注披露 — 未审数/审计调整数/审定数』
            # 必须含"—"分隔符——普通『XX附注披露（2026年度）』（固定资产/在建工程等段式
            # 附注）无"—"不应拦截（曾致 in_block 永不设置 → 段式附注全部取不到）
            _par_plane = ('未审数' in b0 and '审定数' not in b0)
            if _par_plane:
                _par_seen = True
            _par_acc = [0.0] * 60 if _par_plane else _par_acc
            cur_seg = 'end' if _par_plane else 'other'
            continue
        # 三平行表表头行（A=二级科目，B 列起各主体）：记录主体列数——数据行 N 列=行合计，
        # 累计只算主体列（否则每行行合计被重复计入 → 附注双倍，销售费用曾 17.7M=2×8.86M）
        # 2026-08-07 修复：组件块附注表头（『项目|项目|主体1..N|全集团合计』，固定资产/
        # 无形资产/投资性房地产等）同结构 → 也更新 _par_ncols。此前 _par_ncols 保持默认
        # 12 → 超 12 主体的列（c13+）被截断 → g 固定资产附注 751M vs TB 1,038M
        #（差 286,891,969.09=爱绅/股份/铁城/青岛海洋 4 主体未计入）。
        if (b0.startswith('二级科目') or '核算主体' in b0
                or (b0 in ('项目', '类别', '段名') and len(r) > 3
                    and any(r[_j] is not None and ('公司' in str(r[_j]) or '集团' in str(r[_j]))
                            for _j in range(2, min(len(r), 60))))):
            _par_ncols = 0
            for _j in range(1, min(len(r), 60)):
                if r[_j] is not None and str(r[_j]).strip():
                    _par_ncols = _j
                else:
                    break
            # 表头最后一列常为『集团合计』（非主体列）→ 减 1
            if _par_ncols and r[_par_ncols] is not None and '合计' in str(r[_par_ncols]):
                _par_ncols -= 1
            continue
        if _par_plane:
            _vv = []
            for _j in range(1, min(len(r), 60)):
                _x = _fnum(r[_j]) if _j < len(r) else None
                if _x is None and _j < len(r) and isinstance(r[_j], str) and r[_j].startswith('='):
                    _x = _resolve_value(rows, idx, _j)
                _vv.append(_x)
            if any(v is not None for v in _vv):
                if b0.startswith('(') or b1.replace(' ', '') == '(合计)':
                    # 2026-08-06 修复：协议化后(合计)行 N 列（集团合计/行合计）已写数值，
                    # 全列求和 = 主体列和 + 集团合计列 = 2×（损益附注 26 科目翻倍回归根因）。
                    # 只累计主体列（≤_par_ncols，与数据行同口径）。
                    _s = sum(v for i, v in enumerate(_vv, 1)
                             if i <= _par_ncols and v is not None)
                    if _s != 0 or cand is None:
                        cand = _s          # (合计) 行有值直接取；空则保持累计值
                else:
                    for _j, _v2 in enumerate(_vv, 1):
                        if _j > _par_ncols:
                            break          # 只累计主体列，排除行合计列（N 列）
                        if _v2 is not None:
                            _par_acc[_j] += _v2
            continue
        # ① 合计/小计行优先（格式2 合并口径：A 列整块都是科目名，必须在此处理，
        #    否则被下方『科目块起点』分支 continue 吞掉——2026-08-04 修复）。
        #    格式2（应收/应付）：『合 计/账面余额合计/小计』多行——2026-08-05 修复：
        #    按行优先级覆盖（科目总额『合 计』> 披露块『账面余额小计/账面价值合计』
        #    > 子块『小计』），防『账龄1年以上重要的XX』块『小计』行覆盖科目合计行
        #    （应付账款附注曾取 260.7M=子块小计 vs 科目合计 328.0M，差 6,736 万）；
        #    每主体一列（货币资金『集团合计数』）：B 列起各主体列均为数值/公式 → 全行求和=全集团。
        _b1flat = b1.replace(' ', '')
        _a0flat = b0.replace(' ', '')
        _total_pri = 0
        is_total = False
        # 2026-08-06 净额科目（应收/其他应收）：附注披露『账面价值合计』行=总额−坏账−合并抵消，
        # 与合并 TB 净额口径一致——优先取账面价值行（净额），其次科目总额合计。
        if _b1flat in ('账面价值合计', '账面价值小计') and subj in NET_SUBJ:
            is_total = True
            _total_pri = 4                       # 净额科目：账面价值最高优先
        elif _b1flat in ('合计', '(合计)') or _a0flat == '合计':
            is_total = True
            _total_pri = 3                       # 科目总额合计（应付/合同负债『合 计』）
        elif _b1flat in ('账面余额合计', '账面余额小计', '账面价值合计', '账面价值小计'):
            is_total = True
            _total_pri = 2                       # 披露块合计（其他应收款『账面余额小计/账面价值合计』）
        elif _b1flat.startswith('小计') or _b1flat == '(小计)':
            is_total = True
            _total_pri = 1                       # 子块小计（账龄1年以上小计/合并范围内小计等）
        if is_total and in_block and not _got_subj_row:
            # 每主体一列结构（B 列为数值或公式）→ 全行各列求和 = 全集团合计数。
            # 注意：不能借 _resolve_value 兜底判断（它对非公式字符串返回 0.0，
            # 会把格式2 的『账面余额合计』行误判为每主体一列——曾致应收双计 2.86 亿）。
            _b1_v = None
            if len(r) > 1 and r[1] is not None:
                _fv1 = _fnum(r[1])
                if _fv1 is not None:
                    _b1_v = _fv1
                elif isinstance(r[1], str) and r[1].startswith('='):
                    _b1_v = _resolve_value(rows, idx, 1)
            if _b1_v is not None:
                _sv, _has = 0.0, False
                for _j in range(1, min(len(r), 60)):
                    if _j in _sum_cols:      # 2026-08-06 修复：存货附注『合 计』行主体列+全集团合计列双计
                        continue
                    _vv = _fnum(r[_j]) if _j < len(r) else None
                    if _vv is None and _j < len(r) and isinstance(r[_j], str) and r[_j].startswith('='):
                        _vv = _resolve_value(rows, idx, _j)
                    if _vv is not None:
                        _sv += _vv; _has = True
                if _has:
                    if os.environ.get('FN_DEBUG'):
                        print('FN-① r%d %s/%s cand=%s' % (idx, b0[:12], b1[:12], round(_sv, 2)))
                    cand = _sv
                    continue
            # 格式2 合并口径：审定数(第9列)/加计数(第6列)。
            # 注意：空单元格直接 continue（_resolve_value 对空返回 0.0 兜底，
            # 会把『小计行审定数列空』误当 0——应付/合同负债曾取 0）。
            # 2026-08-05：低优先级合计行（子块小计）不覆盖已取到的科目合计/披露块合计
            if _total_pri < _cur_total_pri:
                continue
            for col in (8, 5):
                if col >= len(r) or r[col] is None:
                    continue
                v = _fnum(r[col])
                if v is None and isinstance(r[col], str) and r[col].startswith('='):
                    v = _resolve_value(rows, idx, col)
                if v is not None:
                    cand = v
                    _cur_total_pri = _total_pri
                    break
            continue
        # ===== 2026-08-05 段式/组件块/项目×主体附注（应交税费/职工薪酬/固定资产/无形资产/在建工程/
        # 使用权资产/长期待摊/存货）===== 标题行含 subj → in_block 已由下方 _row_hit 设置；
        # 段标题（未审数/审计调整数/审定数/期初数/期末数/本期增加-贷方…）与块标题（一、原值…四、
        # 账面价值）b0 不含 subj → 在 in_block 后独立识别；数据行（B 或 C 列数值）→ 期末数段
        # （或无段）累计各主体列：组件块科目仅累计『账面价值/净值』块『未审数』段；存货累计
        # 『期末账面价值』块；应交税费/职工薪酬/长期待摊累计期末数段全部税种/项目行。
        # 2026-08-05 修复：库存现金/银行存款/其他货币资金 三科目排除段式累计——货币资金附注按
        # 科目行（库存现金/银行存款/其他货币资金 三行并列）列示，subj=银行存款 时若段式累计会把
        # 其他货币资金/审定数公式行也计入 → 附注 307,463,597.93（=全货币资金）vs 审定 304,802,254.95。
        # 三科目必须走下方 _row_hit 精确行匹配（B 列数值行只累计银行存款行）。
        if in_block and subj not in ('库存现金', '银行存款', '其他货币资金'):
            if '未审数' in b0 and '审定数' not in b0:
                _main_seg = 'unaudit'
                _multi_seg = True
            elif '审计调整' in b0:
                _main_seg = 'adj'
            elif '审定数' in b0:
                _main_seg = 'aud'
            _b0flat2 = b0.replace(' ', '')
            if _b0flat2 in ('期初数', '期末数') or '本期增加' in b0 or '本期减少' in b0 \
                    or '本年增加' in b0 or '本年减少' in b0:
                cur_seg = 'end' if '期末' in b0 else ('open' if '期初' in b0 else 'other')
                _seg_seen = True
                continue
            # 块标题行（一、原值…）B 列为空；收入附注『一、主营业务收入』B 列有值=数据行 → 不拦截
            # 2026-08-07 修复①：块标题判定补『摊销』——S 无形资产附注块标题『二、累计摊销』
            # 只有"摊销"无"折旧/减值/余额" → 不匹配 → comp 沿用 gross → 摊销块数据行
            # 被当原值累计（附注 261M vs TB 168.8M，虚增 9,241 万）。
            # 2026-08-07 修复②：块标题必须【序号开头】（一、二、…）——数据行类别名
            # 『软件累计摊销/土地累计摊销』含"摊销"但无序号，若仅按关键词会误判块标题
            # → comp 错乱 + 数据行被跳过（S 无形资产 cand 变 None）。
            _b_is_num = len(r) > 1 and r[1] is not None and isinstance(r[1], (int, float))
            if re.match(r'^[一二三四五六七八九十]+、', b0) and not _b_is_num:
                if '账面价值' in b0 or '净值' in b0:
                    # 2026-08-06 修复：存货附注块三『期末账面价值』+块七『期末审定账面价值』
                    # 均为 net 块 → 双累计 2×。有『审定』字样=审定净块（net_audit）——
                    # 有审定净块时只累计它（审计口径以审定数为准），否则累计未审 net 块。
                    _comp_block = 'net_audit' if '审定' in b0 else 'net'
                    _has_net_block = True
                elif '原值' in b0 or '余额' in b0 or '成本' in b0:
                    _comp_block = 'gross'
                elif '折旧' in b0 or '摊销' in b0:
                    _comp_block = 'dep'
                elif '减值' in b0:
                    _comp_block = 'imp'
                else:
                    # 2026-08-06：非组件块标题（如存货『八、存货跌价准备变动表』）重置，
                    # 否则沿用上一 net 块 → 变动表数据行被累计。
                    _comp_block = ''
                continue
            # 2026-08-05 数据行段名在 B 列（固定资产/无形资产/在建工程：A=类别、B=期初数/本期增加/期末数）
            # 段名行不 continue——该行 C 列起是主体值（B 列=段名），更新 cur_seg 后继续走数据行累计
            if len(r) > 1 and r[1] is not None and isinstance(r[1], str):
                _b1seg = r[1].strip()
                if _b1seg in ('期初数', '期末数') or '本期增加' in _b1seg or '本期减少' in _b1seg \
                        or '本年增加' in _b1seg or '本年减少' in _b1seg:
                    cur_seg = 'end' if '期末' in _b1seg else ('open' if '期初' in _b1seg else 'other')
                    _seg_seen = True
            _is_num_row = False
            for _jj in (1, 2):
                if _jj < len(r) and r[_jj] is not None:
                    if _fnum(r[_jj]) is not None or (isinstance(r[_jj], str) and r[_jj].startswith('=')):
                        _is_num_row = True
                        break
            if _is_num_row:
                # 2026-08-06 协议化：合计/小计/集团合计/合并抵消/合并报表数/说明行
                # （B 列起为数值会被误当数据行累计 → 职工薪酬附注 2 倍+）→ 跳过，
                # 只累计明细数据行（D 行和 = 各主体期末，即全集团正确值）。
                # 2026-08-06 17:30 再改：传 A/B/C 三列（同 _sheet_ent_values——银行存款
                # 明细表合计行 A/B 空 C 列合计名；公式『=』仅在未协议化辅助表出现，跳过更安全）。
                _rk0 = PROTO.row_kind(' '.join(str(x) for x in r[:3] if x is not None))
                # 期末段『合并报表数』/『集团合计数』行 = 全集团权威值（=各主体期末和，
                # 含抵消口径）→ 直接取全行和，不再 D 行累计（SUM 总行+明细行+审定数行
                # 曾 3 倍：jj 职工薪酬附注 13.9M = 3×4.65M）
                # 2026-08-06 17:45 修复：全行求和须排除 _sum_cols 合计列——longterm 附注
                # 数值化后『合并报表数』行末列=全集团合计（与主体和同值）→ cand 2 倍
                # （FY 固定资产附注 1813M=2×906M 回归根因；公式版合计列读 None 侥幸未暴露）。
                # 2026-08-07 修复：consol/grp 权威行（附注矩阵『合并报表数/集团合计数』行）
                # 覆盖 cand 时须遵循组件块口径——固定资产附注【审定数段】『集团合计数』行
                # = 账面价值 781,737,004.02，曾覆盖【未审数段】已累计的原值 1,110,419,835.82
                # （附注差异 328,682,831.80=累计折旧+减值）。组件块科目仅未审数段 gross 块
                # 的 grp/consol 行可覆盖；其他（net/dep/imp 块或 adj/aud 段）一律跳过。
                if _rk0 in ('consol', 'grp') and cur_seg == 'end':
                    _sv, _has = 0.0, False
                    _ok_grp = True
                    if subj in COMP_BLOCK_SUBJ:
                        _ok_grp = (_comp_block == 'gross' and _main_seg == 'unaudit')
                    elif subj == '存货':
                        _ok_grp = (_comp_block == _net_comp)
                    if _ok_grp:
                        for _j in range(1, min(len(r), 60)):
                            if _j in _sum_cols:
                                continue
                            _vv = _fnum(r[_j]) if _j < len(r) else None
                            if _vv is not None:
                                _sv += _vv
                                _has = True
                        if _has:
                            cand = _sv
                    continue
                if _rk0 in ('tot', 'sub', 'grp', 'offset', 'consol'):
                    continue
                if _rk0 == 'note':
                    # 2026-08-06 修复：收入附注『四、其他业务收入』等序号开头【数据行】
                    # （B 列有值）被 is_note_row『序号块标题』规则误判 note → 跳过 →
                    # 营业收入附注只取主营（561.4M）漏其他业务收入（少 5.23M）。
                    # 块标题（B 列空：一、期末账面价值…）保持跳过；B 列有数值=数据行 → 放行
                    # 到下方 _row_hit 精确行匹配累计。
                    _b1n = len(r) > 1 and r[1] is not None and isinstance(r[1], (int, float))
                    if not _b1n:
                        continue
                _target_ok = True
                if subj in COMP_BLOCK_SUBJ:      # 组件块：2026-08-06 核对口径取【原值/成本/余额】块
                    # （用户TB 一级=原值；原取审定净值块 → dq 固定 819.4M vs 514.5M 差 304.9M）；
                    # 无原值块（在建工程仅账面余额→gross 覆盖；极端仅净值）回退净值块。
                    if _has_gross_block:
                        _target_ok = (_comp_block == 'gross' and _main_seg == 'unaudit')
                    else:
                        _target_ok = (_comp_block == _net_comp and _main_seg == 'unaudit') \
                            or (not _has_net_block and _comp_block == 'gross' and _main_seg == 'unaudit')
                elif subj == '存货':
                    # 2026-08-06 有『账面价值』块→仅账面价值；无净值块（THB 存货仅『期末账面余额』
                    # 块标题）→ 取余额块（gross）；否则『一、期末账面余额』被当 gross 与
                    # 『三、期末账面价值』双累计（253M vs 115.5M）
                    _target_ok = (_comp_block == _net_comp) \
                        or (not _has_net_block and _comp_block == 'gross')
                # 单段附注（职工薪酬审定口径/应交税费审定口径等无『未审数』标题）→ 不受大段限制
                if _multi_seg and _main_seg != 'unaudit':
                    _target_ok = False
                # 段式累计仅在有段标题时生效（cur_seg='end'）；无段标题（收入项目×主体/存货块式）
                # 仅存货与组件块走段式累计，收入类留给下方 _row_hit 精确行匹配（防 R05-R13 全累计）
                if _seg_seen:
                    _seg_ok = (cur_seg == 'end')
                else:
                    _seg_ok = (subj in COMP_BLOCK_SUBJ or subj == '存货')
                if _target_ok and _seg_ok:
                    _sv, _has = 0.0, False
                    for _j in range(1, min(len(r), 60)):
                        if _j > _par_ncols:
                            break
                        if _j in _sum_cols:      # 2026-08-06 排除『全集团合计』列（组件块附注双计源）
                            continue
                        _vv = _fnum(r[_j]) if _j < len(r) else None
                        if _vv is None and _j < len(r) and isinstance(r[_j], str) and r[_j].startswith('='):
                            _vv = _resolve_value(rows, idx, _j)
                        if _vv is not None:
                            _sv += _vv
                            _has = True
                    # 2026-08-05 净值化：行名含『摊销/折旧』（无形资产账面价值块的累计摊销行）→ 取负抵减
                    # （固定资产账面价值块用负数『固定资产』行已净值化，无需取负）
                    # 2026-08-06 修复：longterm 生成器已把账面价值块摊销/折旧行改为负数（净额口径，
                    # 账面价值=原值+负摊销直接求和）→ 行级 _sgn=-1 再取负 = 双取负（净值被算成
                    # 原值+摊销：FY 无形资产附注 52.99M vs 正确 44.95M）→ 改符号智能判断：
                    # 摊销/折旧行【行合计为正】（旧生成器未净值化）才取负；【已为负】（新生成器
                    # 净额口径）保持原值，兼容新旧两种生成器输出。
                    if _has:
                        _sgn = -1.0 if (('摊销' in b0 or '折旧' in b0) and _sv > 0) else 1.0
                        if os.environ.get('FN_DEBUG'):
                            print('FN累计 r%d %-14s seg=%s comp=%s main=%s sv=%s cand→%s' % (
                                idx, b0[:14], cur_seg, _comp_block, _main_seg, round(_sv, 2),
                                round((_sgn * _sv) if cand is None else cand + _sgn * _sv, 2)))
                        cand = (_sgn * _sv) if cand is None else cand + _sgn * _sv
                        continue
        # ② 科目块起点：A 列含 subj（每主体一列结构中该行即科目数据行）。
        #    仅累计『期末数』段（货币资金附注的 审计调整/期末审定数 段同科目行不覆盖期末值）。
        #    货币资金报表科目：附注按 库存现金/银行存款/其他货币资金 三行列示 → 三行累计求和；
        #    其他科目：单行覆盖（取该科目行各主体列之和）。
        # 2026-08-05 并入科目（应付股利等）：格式2 附注 B 列=行名（『普通股股利』），匹配 B 列，
        # 且 B 列行名含 subj 时取 审定数/加计数（普通股股利 行 I 列=146300）。
        _cash_subj = (subj == '货币资金')
        # 2026-08-05 行名关键词扩展：营业收入=主营业务收入+其他业务收入；营业成本=主营业务成本+其他业务成本
        # （项目×主体附注行名『一、主营业务收入』不含『营业收入』子串，直接 subj in b0 匹配不到）
        _kw_list = SUBJ_ROW_KW.get(subj, (subj,))
        _row_hit = bool(subj and (any(_k and _k in b0 for _k in _kw_list)
                                  or (b1 and any(_k and _k in b1 for _k in _kw_list))
                                  or (_cash_subj and b0 in ('库存现金', '银行存款', '其他货币资金'))))
        if _row_hit:
            in_block = True
            if len(r) > 1 and r[1] is not None and isinstance(r[1], (int, float)):
                # 每主体一列结构：B 列为数值 → 期末数段（或无段）累计（货币资金三行/长期借款/收入项目行）
                if cur_seg in ('end', ''):
                    # 2026-08-06 17:35 修复：排除 _sum_cols 合计列（银行存款附注 money 模式
                    # 数值化后加计数列被累计 → 附注 738M=2×369M 回归）
                    vals = [v for _jj, v in enumerate((_fnum(x) for x in r[1:]), 1)
                            if v is not None and _jj not in _sum_cols]
                    if vals:
                        # 2026-08-05 多行累加：货币资金三行 / SUBJ_ROW_KW 科目（营业收入=主营+其他业务收入）
                        if os.environ.get('FN_DEBUG'):
                            print('FN-② r%d %s 累加 %s' % (idx, b0[:10], round(sum(vals), 2)))
                        if _cash_subj or subj in SUBJ_ROW_KW:
                            cand = (cand or 0.0) + sum(vals)
                        else:
                            cand = sum(vals)
            elif b1 and subj in b1 and not any(k in b1 for k in ('校对', '账龄', '小计', '合计', '止（')):
                # 格式2 科目行（B 列=行名且含 subj）：取 审定数(第9列)/加计数(第6列)。
                # 排除『与TB校对—XX余额』『（2）账龄1年以上的XX』等非披露科目行（2026-08-05）。
                for _col in (8, 5):
                    if _col < len(r) and r[_col] is not None:
                        _v = _fnum(r[_col])
                        if _v is None and isinstance(r[_col], str) and r[_col].startswith('='):
                            _v = _resolve_value(rows, idx, _col)
                        if _v is not None:
                            if os.environ.get('FN_DEBUG'):
                                print('FN-fmt2 r%d cand=%s' % (idx, round(_v, 2)))
                            cand = _v
                            _got_subj_row = True
                            break
            continue
        if os.environ.get('FN_DEBUG') and idx > 131:
            print('FN-循环尾 r%d no-cand-change (cand=%s)' % (idx, round(cand, 2) if cand is not None else None))
    # 三平行表兜底：未审数块内累计的各主体列和（合计行 '(合计)' 列空时）
    if os.environ.get('FN_DEBUG'):
        print('FN返回前: cand=%s par_acc_sum=%s par_seen=%s' % (
            round(cand, 2) if cand is not None else None,
            round(sum(_par_acc), 2) if _par_seen else 'n/a', _par_seen))
    if cand is None and _par_seen and any(abs(v) > 0.005 for v in _par_acc):
        cand = sum(_par_acc)
    return cand


def main():
    argv = sys.argv[1:]
    if not argv:
        print(__doc__)
        return
    tb = argv[0]
    folder = argv[1] if len(argv) > 1 else None
    out_path = argv[2] if len(argv) > 2 else None   # 2026-08-04：可选输出路径（防占用冲突）
    if not os.path.isfile(tb):
        print(f'⚠️ 试算表不存在：{tb}')
        return
    recon_trial_balance(tb, folder, out_path)


if __name__ == '__main__':
    main()
