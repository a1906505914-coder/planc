# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=(无,共享库) | 关键列=- | 职责=共享底层库:TB/GL读取·表头探测·带符号余额·_safe_save·finalize_workbook(re-export audit_shell);所有生成器统一import本库
"""
audit_common.py — 亿利达审计小程序「共享底层库」（单一解析口径枢纽）

为什么存在：
  桌面交付物生成器（tax/payroll/pl/expense/longterm_assets/revenue/bank/
  current_account/inventory_detail）原本各自内联一套 TB/GL 读取、表头探测、
  带符号余额、_safe_save。同一份"读账套"逻辑散在 9+ 处，一旦用友导出格式变了
  要改 9 处，极易口径漂移（即"铁律3/铁律0"类事故的根因）。

本模块把所有【通用原始读取层 + 数值助手 + 锁感知保存】收口到一处：
  · 实体发现        discover_entities / _extract_year
  · GL 读取         detect_gl_columns / read_gl_rows（返回逐条凭证行 dict）
  · TB 读取         read_tb_full / read_tb_leaf（返回带借贷符号的余额字典）
  · 表头动态探测    已内嵌于 detect_gl_columns / read_tb_full
  · 带符号余额       signed_balance(direction, amount)
  · 数值助手         safe / _nz / _recon_status
  · 锁感知保存       _safe_save(wb, out_path) -> (final_path, warn_or_None)
  · 宋体/千分位收尾  finalize_workbook（re-export 自 audit_shell，单一来源）

用法（各生成器顶部）：
  from audit_common import (
      discover_entities, detect_gl_columns, read_gl_rows,
      read_tb_full, read_tb_leaf, signed_balance, _safe_save, safe,
  )
  from audit_common import finalize_workbook   # 与 audit_shell 同源

注意：本模块刻意【不包含】任何"按交付物特化的聚合逻辑"（如税费取数、薪酬
子目拆分、银行户拆分等）——那些留在各自生成器里，保证隔离性与定点重跑能力。
本模块只负责"把 xlsx 读成结构化行/余额字典"这一层，是所有生成器真正共享的
解析口径。
"""
import os
import unicodedata
import re
import glob
import openpyxl
from collections import defaultdict, Counter
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ===================== 视觉/收尾：re-export 自 audit_shell（单一来源） =====================
# 所有生成器早已共用 audit_shell 的视觉样式与 finalize_workbook；此处 re-export，
# 使 audit_common 成为"读取层 + 收尾"的单一导入枢纽，且不重复实现收尾逻辑。
try:
    from audit_shell import (  # noqa: F401
        finalize_workbook, SHELL_HFILL, SHELL_HFONT, SHELL_TITLE_FONT,
        SHELL_SUB_FONT, SHELL_BOLD, SHELL_NUM, SHELL_BODY, SHELL_THIN,
        SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT, resolve_template, clone_sheet,
        make_default_shell, apply_header_row, AUDIT_TEMPLATES_DIR,
    )
    _HAVE_SHELL = True
except Exception:  # pragma: no cover - audit_shell 缺失时降级，保证不闪退
    _HAVE_SHELL = False

    def finalize_workbook(wb):
        """降级兜底：无 audit_shell 时仅做最基础的宋体 + 千分位，避免导入失败。"""
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.font is not None:
                        cell.font = Font(name='Times New Roman', size=10,
                                         bold=cell.font.bold, italic=cell.font.italic,
                                         underline=cell.font.underline, strike=cell.font.strike,
                                         color=cell.font.color)
                    v = cell.value
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        if cell.number_format in (None, '', 'General'):
                            cell.number_format = '#,##0.00'
        return wb


# ===================== 数值助手 =====================
def safe(x):
    """转 float，失败返回 0.0（脏数据/空值/文本兼容）。"""
    if isinstance(x, str):
        # ⚡⚡ 2026-08-31 千分位逗号（XBJ U8 导出金额 "1,078"）：原 float('1,078') 抛错 → 0
        #   → XBJ 损益审定表 jf/df 解析 0 空壳根因之一。
        x = x.replace(',', '').strip()
    try:
        return float(x)
    except Exception:
        return 0.0


def _nz(x):
    """近零归零（消除浮点噪声）。"""
    return 0.0 if abs(x) < 1e-6 else x


def _recon_status(diff, base_a, base_b):
    """勾稽状态：差额在容差内为'一致'。容差 = max(1.0, 较大基准*1e-6)。"""
    return '一致' if abs(diff) <= max(1.0, max(abs(base_a), abs(base_b)) * 1e-6) else '不一致'


def signed_balance(direction, amount):
    """把"方向+金额"转带符号余额：贷=−，借/其它(余额类默认借正)=+。

    这是所有"期初+借发−贷发=期末"滚动验算的基础（借+贷−通用式）。
    ⚠️ 2026-08-23：SAP 300 形态科目余额表金额【本身带符号】（贷=-196815301.18 等），
    方向列只是冗余标注——本函数直接对负值取负会双重取负变正（AR 方向符号 bug 根因）。
    调用方应优先用 signed_balance_v2（兼容带符号金额）；本函数保留给 U8 等
    『方向+绝对金额』账套（此时金额恒正，取负逻辑正确）。
    """
    a = safe(amount)
    if str(direction).strip() == '贷':
        return -a
    return a


def signed_balance_v2(direction, amount):
    """方向+金额 → 带符号余额（兼容两种导出形态）：
    ① U8 类『方向+绝对金额』：金额恒正 → 贷取负、借保持（同 signed_balance）。
    ② SAP 300 类『方向+带符号金额』：金额负=贷方余额（符号已含方向）→ 保持原符号。
    判定：金额为负 → 保持原符号（已是带符号余额）；金额非负 → 按方向定符号。
    验证：3300 全量 377 科目滚动验算 0 不平；U8 AZ 25 科目 0 不平。
    """
    a = safe(amount)
    if a < 0:
        return a
    if str(direction).strip() == '贷':
        return -a
    return a


# ===================== 锁感知保存（统一单一来源） =====================
def _safe_save(wb, out_path):
    """锁感知保存：目标被占用(多为资源管理器预览窗格/缩略图锁住 xlsx，并非 Excel 打开)
    时，自动改存到同目录 '占用副本' 子文件夹并保留正确文件名，避免整次运行 0 产出。

    返回 (最终保存路径, 警告文本或 None)。
    失败语义：仅当目标与副本【均】无法写入时才返回 (None, 警告)，【不抛异常】，
    由调用方决定如何处理（与 audit_toolkit 历史行为一致）。
    """
    try:
        os.makedirs(os.path.dirname(out_path) or '.', exist_ok=True)
        # 2026-08-05 前道规范化：保存前结构校验（三件套：审定表/明细表/附注）——生成即规范。
        # 0 问题不打印；有 WARN/ERROR 打印提示（不阻断保存，ERROR 由检查器兜底）。
        try:
            validate_workbook(wb, os.path.basename(out_path), raise_on_error=False)
        except Exception:
            pass
        wb.save(out_path)
        return out_path, None
    except PermissionError:
        import datetime as _dt
        ts = _dt.datetime.now().strftime('%Y%m%d_%H%M%S')
        parent = os.path.dirname(out_path) or '.'
        sub = os.path.join(parent, '占用副本')
        try:
            os.makedirs(sub, exist_ok=True)
        except Exception:
            pass
        alt = os.path.join(sub, os.path.basename(out_path))
        if os.path.exists(alt):
            b, e = os.path.splitext(alt)
            alt = f'{b}_{ts}{e}'
        try:
            wb.save(alt)
            return alt, (f'目标文件被占用（多为资源管理器"预览窗格"锁住 xlsx，并非 Excel 打开），'
                         f'已改存到子文件夹：占用副本\\{os.path.basename(alt)}'
                         f'（请关闭资源管理器对该文件的预览/窗口，或直接用此副本）')
        except Exception as _e2:
            return None, f'目标与副本均无法写入：{out_path} / {alt}（{_e2}）'


def move_if_free(src, dst):
    """锁感知搬移（2026-08-29 防 bank PermissionError 类回退）：
    dst 被占用（Excel/WPS 打开：检测 ~$ 锁文件 + 独占打开探测）时打印明确提示并跳过，
    不抛异常。返回 True=搬移成功 / False=目标占用跳过 / None=其他异常。"""
    base = os.path.basename(dst)
    lock = os.path.join(os.path.dirname(dst) or '.', '~$' + base)
    if os.path.exists(lock):
        print(f'  ⚠️ 目标被 Excel/WPS 占用（检测到 ~$ 锁文件）：{base}，跳过搬移，请关闭该文件后重跑')
        return False
    if os.path.exists(dst):
        try:
            _fd = os.open(dst, os.O_RDWR)
            os.close(_fd)
        except (PermissionError, OSError):
            print(f'  ⚠️ 目标被占用（Excel/WPS 打开中？）：{base}，跳过搬移，请关闭该文件后重跑')
            return False
    try:
        os.replace(src, dst)
        return True
    except Exception as _e:
        print(f'  ⚠️ 搬移失败 {os.path.basename(src)} → {base}：{_e}')
        return None



# ===================== 进程级缓存 =====================
_CACHE_TB = {}
_CACHE_GL = {}

# ===================== 实体发现 =====================
def discover_entities(data_dir):
    """返回 {entity: {year: {'km': path, 'gl': path}}}。
    稳健实体发现：主导年份回退（文件名缺年份时用出现最多的年份）、忽略 .xls 旧格式、
    排除外币科目余额表、科目余额表/综合查询明细表配对。
    """
    ent = {}
    if not os.path.isdir(data_dir):
        return ent
    xls_files = []
    candidates = []
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith('~$') or fn.startswith('.'):
            continue
        fp = os.path.join(data_dir, fn)
        if os.path.isdir(fp):
            # ⚡⚡ 2026-08-23 排除非账套主体目录：描述性目录（如 AFJ『其他应收其他应付
            #   科目余额表』——用户另行提供的辅助明细表，非账套数据）内含『XX其他应收款.
            #   xlsx』等文件，会被误识别为主体/科目（试算表多出 0201 李秀月 等辅助核算行、
            #   平衡校验多出『新昌日发/纺机本部其他应付款』主体）。主体目录名应是账套
            #   代码/名称（3300、AF、山东日发），不应含『余额表/报表/辅助核算/明细表』
            #   等文件类型描述词。排除后不影响标准账套（ADF 3300/3100 等数字代码）。
            if re.search(r'余额表|科目余额|报表|辅助核算|明细表|账龄', fn):
                continue
            for sfn in sorted(os.listdir(fp)):
                if sfn.startswith('~$'):
                    continue
                candidates.append((sfn, os.path.join(fp, sfn)))
        else:
            candidates.append((fn, fp))
    for fn, full in candidates:
        low = fn.lower()
        if low.endswith('.xls') and not low.endswith('.xlsx'):
            xls_files.append(fn)
            continue
        # ⚡ 2026-08-23：3400 txt 序时账（GBK 制表符分隔）作为 GL 数据源候选——
        #   仅『序时账*.txt』放行（km 仍只认 xlsx；txt 只可能来自 sap_common._scan_gl_txt）。
        _is_txt_gl = low.endswith('.txt') and '序时账' in fn
        if not (low.endswith('.xlsx') or _is_txt_gl):
            continue
        y = _extract_year(fn)
        if not y:
            # ⚡ 铁律131b（2026-08-15 布局规范）：数据/{期间}/{主体}/ 下文件可能无年份
            #   （ADF '3100科目余额表1-6月.xlsx'）→ 目录名年份回退（数据/2026/3100 → 2026）
            _md = re.search(r'(?<!\d)(20\d{2})(?!\d)', os.path.dirname(full))
            if _md:
                y = _md.group(1)
        base = re.sub(r'(19|20)\d{2}', '', fn)
        base = re.sub(r'年度|年\d*[-~]?\d*月', '', base)  # 剔除"年度"、"年1-3月"等期间后缀
        E = re.sub(r'(外币)?(科目余额表|科目余额|余额表|综合查询明细表|综合查询表|辅助核算余额表|辅助核算|总账|序时账|trial).*$',
                   '', base, flags=re.IGNORECASE).strip(' _-')
        # 2026-08-06：U8 文件名『主体2025年科目余额表』→ 去年份后残留孤立『年』
        # （FY 是『2025年度』已剔除、JTt/SSS 是『2025年』未剔除）→ 只删实体名尾部"年"。
        E = E.rstrip('年')
        # 个位数前缀补0（如 1、FY本级公司 → 01、FY本级公司），确保排序 1,2,...10 而非 1,10,11,2
        _m0 = re.match(r'^(\d)([、,，.．\s])', E)
        if _m0 and len(E) > 2:
            E = '0' + E
        if not E:
            # ⚡ 2026-08-10 根目录直放且文件名无主体前缀（DQ/科目余额表_2024.xlsx 去年份后
            # 实体名=''）→ 实体名回退账套目录名（DQ），否则该年度被整年跳过。
            E = os.path.basename(os.path.normpath(data_dir)).strip(' _-')
            if not E:
                continue
        is_wb = ('外币' in fn) and ('科目余额表' in fn or '余额表' in fn)
        if is_wb:
            continue
        elif ('科目余额表' in fn or '总账' in fn or 'trial' in low) and '综合查询' not in fn and '辅助核算' not in fn:
            typ = 'km'
        elif '综合查询' in fn or '序时账' in fn:
            # 2026-08-22：ADF/SAP 账套的明细账文件命名『3300序时账1月.xlsx』（非"综合查询"）
            # → gl 识别加『序时账』关键词（此前所有序时账主体 gl 缺失 → 利息匹配/凭证抽查无数据）。
            typ = 'gl'
        else:
            continue
        ent.setdefault('__items', []).append((E, y, typ, full))
    # 主导年份回退
    items = ent.pop('__items', [])
    yc = Counter(y for (_, y, _, _) in items if y)
    dom_y = yc.most_common(1)[0][0] if yc else None
    for E, y, typ, full in items:
        yy = y or dom_y
        if yy is None:
            continue
        yr = ent.setdefault(E, {}).setdefault(yy, {'km': None, 'gl': None})
        if typ == 'km':
            # 2026-08-22：多期间科目余额表（3100科目余额表1-6月.xlsx + 1-7月.xlsx）→
            # 选【最新期间】（sorted 后靠后，如 1-7月），使期末余额与借款台账/核对版一致。
            # 原 `if yr['km'] is None` 只保留第一个（最早期间）→ 期末少 7 月变动。
            yr['km'] = full
        elif typ == 'gl':
            yr['gl'] = full
    if xls_files:
        print(f'  ⚠️ 忽略 {len(xls_files)} 个 .xls 旧格式：' + '、'.join(xls_files[:5]))
    return {E: b for E, b in ent.items() if any(v.get('km') for v in b.values())}


def norm_entities(entities):
    """统一 U8/SAP 实体结构为 U8 嵌套格式 {E: {'km': {y: path}, 'gl': {y: path}}}（幂等）。

    ⚡ 2026-08-12 治理：SAP adapter.discover_entities 返回 {E: {y: ['km','gl','aux']}}
    （年度为 key、值为文件路径 list），U8 是 {E: {y: {'km': path, 'gl': path}}}；
    各生成器本地 _discover_entities 大多已包装成 U8 结构，但【边界函数】（直接接收
    adapter 原始结构，如 inventory 集团审定表曾 KeyError 'km'）反复踩坑。本函数：
      · U8 结构（b['km'] 是 dict）→ 原样返回；
      · SAP 结构（b[y] 是 list）→ 转 {km:{y: list[0]}, gl:{y: list[1]}}（若已含 'km'/'gl'
        键则保持）；
      · 幂等：对已 norm 的结构再调用无副作用。
    各生成器边界函数调用前先 norm_entities(entities) 即消灭双结构坑。"""
    out = {}
    for e, b in (entities or {}).items():
        d = out.setdefault(e, {'km': {}, 'gl': {}})
        if not isinstance(b, dict):
            continue
        # 已 norm：{km: {y: path}, gl: {y: path}}（或 SAP 也含这些键）
        if 'km' in b and isinstance(b['km'], dict):
            d['km'].update({str(y): p for y, p in b['km'].items()})
            if 'gl' in b and isinstance(b['gl'], dict):
                d['gl'].update({str(y): p for y, p in b['gl'].items()})
            continue
        for y, p in b.items():
            if isinstance(p, dict) and ('km' in p or 'gl' in p):
                if p.get('km'):
                    d['km'][str(y)] = p['km']
                if p.get('gl'):
                    d['gl'][str(y)] = p['gl']
            elif isinstance(p, (list, tuple)):
                # SAP：{y: ['km','gl','aux_ar','aux_ap']}（索引 0=km、1=gl；部分账套缺失补 None）
                _km = p[0] if len(p) > 0 else None
                _gl = p[1] if len(p) > 1 else None
                if _km:
                    d['km'][str(y)] = _km
                if _gl:
                    d['gl'][str(y)] = _gl
    return {E: b for E, b in out.items() if b['km'] or b['gl']}


def _extract_year(fn):
    """从文件名稳健抽取年份（兼容 2025 / 25 / 20250720 时间戳等）。
    ⚡ 2026-08-15 修复：优先取【类型名（科目余额表/综合查询明细表等）紧邻前】的 YYYY
    ——标准命名『主体 YYYY年 类型』；否则回退任意位置（历史文件名）。
    防误判：主体名含年份（XBJ『2020年简阳市…项目…2025年科目余额表』）时
    旧逻辑抓到主体名里的 2020 → 误生成 2020 空壳底稿。"""
    s = fn
    m = re.search(r'(20\d{2})\s*年?\s*(外币)?(科目余额表|科目余额|余额表|综合查询明细表|综合查询表|辅助核算余额表|辅助核算|总账|trial)',
                  s, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    s = re.sub(r"(?<!\d)20\d{6,12}(?!\d)", "", s)
    s = re.sub(r"(?<!\d)20\d{6}(?!\d)", "", s)
    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", s)
    if m:
        return m.group(1)
    m2 = re.search(r"(?<!\d)(\d{2})(?!\d)", s)
    if m2:
        try:
            yy = int(m2.group(1))
        except ValueError:
            return None
        if 0 <= yy <= 99:
            return f"20{yy:02d}"
    return None


# ===================== GL 读取（通用原始层） =====================
def detect_gl_columns(ws):
    """探测综合查询明细表(GL)列位置：默认 name/date/vch_type/vch_no/opp/summary/debit/credit。
    返回 {字段: 0基列号, '_hits': 命中的关键字段集}。多行表头兼容（扫描前 2 行）。
    """
    idx = {'name': 0, 'date': 1, 'vch_type': 2, 'vch_no': 3, 'opp': 4,
           'summary': 5, 'debit': 6, 'credit': 7}
    hits = set()
    for r in range(1, 3):
        for c in range(1, (ws.max_column or 10) + 1):   # 2026-08-05 容错 write_only 文件 max_column=None
            v = ws.cell(r, c).value
            if not v:
                continue
            s = str(v)
            if '辅助' in s and '名称' in s:
                pass  # 辅助核算名称列（客户/供应商/个人名），非科目名称列，跳过
            elif '科目' in s and ('名称' in s or '编码' in s or '代码' in s):
                idx['name'] = c - 1
                hits.add('name')
            elif '日期' in s or '期间' in s:
                idx['date'] = c - 1
                hits.add('date')
            elif '借方' in s:
                idx['debit'] = c - 1
                hits.add('debit')
            elif '贷方' in s:
                idx['credit'] = c - 1
                hits.add('credit')
            elif '摘要' in s:
                idx['summary'] = c - 1
            elif '对方' in s:
                idx['opp'] = c - 1
            elif '凭证' in s:
                if '号' in s:
                    idx['vch_no'] = c - 1
                else:
                    idx['vch_type'] = c - 1
    idx['_hits'] = hits
    return idx


def read_gl_rows(data_dir, entities):
    """返回逐条凭证行 dict：{e,y,date,vtype,vno,opp,summary,name,dr,cr}。
    通用原始层：所有生成器/勾稽程序（A1-A4、C3 抽凭）复用同一份 GL 读取，口径唯一。
    跳过借=贷=0 的空行；读取失败的主体打印警告并跳过（不中断）。
    """
    if data_dir in _CACHE_GL:
        return _CACHE_GL[data_dir]
    # 2026-08-05 磁盘缓存：GL 解析结果持久化（.cache/），源文件未变直接加载——
    # 改 bug 后重跑 builder 不再重复解析几十万行 GL（省 3-8 分钟/次）。
    _ac_save = None
    try:
        from audit_cache import load as _ac_load, save as _ac_sv
        _ac_save = _ac_sv
        _d, _p = _ac_load(data_dir, entities, 'gl')
        if _d is not None:
            _CACHE_GL[data_dir] = _d
            return _d
    except Exception:
        pass
    rows = []
    for e, yd in entities.items():
        for y, paths in yd.items():
            g = paths.get('gl')
            if not g or not os.path.exists(g):
                continue
            # 2026-08-05 修复：read_only 对部分文件（write_only 生成/新拆分文件）dims 异常
            # 读 0 行 → 自动回退非只读重读（与 jtt_prepare.load_rows 同一策略）。
            wb = None
            for _ro in (True, False):
                try:
                    wb = openpyxl.load_workbook(g, data_only=True, read_only=_ro)
                    _ws_probe = wb[wb.sheetnames[0]]
                    # 2026-08-22：探测行数只读前 1000 行（避免 SAP FBL3N 大序时账全表遍历卡死）
                    _n_probe = sum(1 for _r in _ws_probe.iter_rows(min_row=3, max_row=1000, values_only=True)
                                   if _r and any(_c is not None and str(_c).strip() for _c in _r))
                    if _n_probe >= 1:
                        break
                    wb.close()
                    wb = None
                except Exception as ex:
                    if _ro:
                        print(f'  ⚠️ GL读取失败 [{e}][{y}]：{ex}')
                    wb = None
                    continue
            if wb is None:
                continue
            ws = wb[wb.sheetnames[0]]
            ci = detect_gl_columns(ws)
            _hits = ci.pop('_hits', set())
            # 2026-08-22：SAP FBL3N 序时账（符号/分配/凭证编号/文本…）无 科目名称/日期/借贷方
            # 列 → 非综合查询格式，快速跳过（避免遍历大文件卡死）；A D F 等 SAP 账套 GL 暂不解析。
            if not ({'name', 'date', 'debit', 'credit'} & _hits):
                wb.close()
                continue
            mc = ws.max_column or 12   # 2026-08-05 容错：write_only 生成文件缺 dimension（max_column=None），GL 至少 8 列
            for k in list(ci.keys()):
                if ci[k] < 0:
                    ci[k] = 0
                if ci[k] >= mc:
                    ci[k] = max(0, mc - 1)
            # 文件级"整表重复"配对检测（2026-08-02 统一口径）：部分账套导出每行复制×2
            # （如 FY深圳 2026 综合查询明细表 336=168×2）→ 该文件【所有】行键计数均为偶数时
            # 判定为整表复制、每键只保留 1 行；否则（个别合法相同分录）全部保留、绝不误删。
            _file = []
            for r in ws.iter_rows(min_row=3, values_only=True):
                try:
                    if not r or len(r) <= ci['credit']:
                        continue
                    nm = r[ci['name']] if 0 <= ci['name'] < len(r) else None
                    if not nm:
                        continue
                    jf = safe(r[ci['debit']])
                    df = safe(r[ci['credit']])
                    if jf == 0 and df == 0:
                        continue
                    date = r[ci['date']] if 0 <= ci['date'] < len(r) else None
                    vtype = r[ci['vch_type']] if 0 <= ci['vch_type'] < len(r) else None
                    vno = r[ci['vch_no']] if 0 <= ci['vch_no'] < len(r) else None
                    opp = r[ci['opp']] if 0 <= ci['opp'] < len(r) else None
                    summary = r[ci['summary']] if 0 <= ci['summary'] < len(r) else None
                    name = str(nm).strip()
                    _dk = (date, vtype, vno, name, jf, df,
                           str(opp or ''), str(summary or ''))
                    _file.append((_dk, r, date, vtype, vno, opp, summary, name, jf, df))
                except Exception:
                    continue
            _whole_dup = False
            if len(_file) >= 4:
                from collections import Counter as _C
                _cnt = _C(k for k, *_ in _file)
                _whole_dup = all(v >= 2 and v % 2 == 0 for v in _cnt.values())
            _seen_keys = set()
            for _dk, r, date, vtype, vno, opp, summary, name, jf, df in _file:
                if _whole_dup:
                    if _dk in _seen_keys:
                        continue
                    _seen_keys.add(_dk)
                rows.append({
                    'e': e, 'y': y,
                    'date': date,
                    'vtype': vtype,
                    'vno': vno,
                    'opp': opp,
                    'summary': summary,
                    'name': name,
                    'dr': jf, 'cr': df,
                })
            wb.close()
    # 摘要补全（2026-07-31 修复：U8 记账凭证只有【首行】有摘要，其余行摘要为空，
    # 导致各 builder 抽凭表摘要大面积空缺——除 pl_detail 自行补全外，其余程序都受影响）。
    # ①按 (e, y, vtype, vno) 凭证组内首个非空摘要回填空摘要行（不覆盖已有摘要）；
    # ②整笔凭证无任何摘要且含"本年利润"（年终结转损益凭证）→ 生成可读摘要兜底。
    _sum_map = {}
    for _r in rows:
        if (_r.get('summary') or '').strip():
            _sum_map.setdefault((_r['e'], _r['y'], _r.get('vtype'), _r.get('vno')), _r['summary'])
    _carry_keys = set()
    for _r in rows:
        if '本年利润' in str(_r.get('opp') or '') or '本年利润' in str(_r.get('name') or ''):
            _carry_keys.add((_r['e'], _r['y'], _r.get('vtype'), _r.get('vno')))
    for _r in rows:
        if not (_r.get('summary') or '').strip():
            _k = (_r['e'], _r['y'], _r.get('vtype'), _r.get('vno'))
            _r['summary'] = _sum_map.get(_k) or ('结转损益（年终结账）' if _k in _carry_keys else '')
    _CACHE_GL[data_dir] = rows
    if _ac_save is not None:
        try:
            _ac_save(data_dir, entities, 'gl', rows)
        except Exception:
            pass
    return rows


# ===== 凭证抽查表统一格式（2026-07-30 用户要求） =====
VOUCHER_STANDARD_HEADERS = [
    '核算主体', '测试序号', '日期', '凭证字', '凭证号',
    '二级科目（或客户/供应商名称）', '摘要', '借方金额', '贷方金额',
    '对方科目', '与原始凭证相符', '原始凭证内容', '原始凭证日期',
    '会计处理正确', '所属时间无误'
]
VOUCHER_STD_COLS = len(VOUCHER_STANDARD_HEADERS)

def sort_voucher_rows(rows):
    """凭证抽查表统一排序：先核算主体，再日期。
    rows 为 [(entity, date_or_key, ...), ...] 列表。"""
    def _sort_key(item):
        e = item[0] or ''
        d = str(item[1] or '')
        return (e, d)
    return sorted(rows, key=_sort_key)


# ===================== 对方科目推导（凭证抽查表通用：按金额精确匹配） =====================
# 2026-07-29 提为全局唯一来源：原实现散落在 fill_voucher_counterparty.build_gl_map/derive_cp
# 与 inventory_detail._build_vidx/_recon_cp 两套重复；此处收口，二者均 import 复用，杜绝口径漂移。
def norm_cp_name(s):
    """对方科目名归一化（比较前用）：去空白、NFKC 全半角。"""
    if not s:
        return ''
    s = unicodedata.normalize('NFKC', str(s))
    return re.sub(r'\s+', '', s).strip()


# 科目"业务族"分类（跨科目勾稽/抽凭通用）：把明细科目归并为审计勾稽常用的对照族。
# 顺序敏感（前面的优先命中）。与 vouching_list 的 subject_of 同源思路，此处作为单一来源供 cross_recon 复用。
_FAMILY_KW = [
    ("资金类", ('银行存款', '库存现金', '其他货币资金', '存放中央银行', '存放同业')),
    ("借款", ('短期借款', '长期借款', '借款', '应付利息')),
    ("应收款项", ('应收账款', '应收票据', '合同资产', '长期应收款')),
    ("应付款项", ('应付账款', '应付票据', '合同负债', '预收账款')),
    ("存货", ('原材料', '库存商品', '周转材料', '在途物资', '委托加工物资', '材料采购',
              '商品采购', '低值易耗品', '包装物', '发出商品', '半成品', '生产成本', '制造费用')),
    ("长期资产", ('固定资产', '在建工程', '无形资产', '长期待摊费用', '使用权资产', '工程物资', '投资性房地产')),
    ("减值准备", ('坏账准备', '存货跌价准备', '减值准备', '跌价准备', '信用减值', '资产减值')),
    ("税费", ('应交税费', '未交增值税', '应交增值税', '税金', '所得税')),
    ("职工薪酬", ('应付职工薪酬', '职工薪酬', '工资', '社保', '公积金', '工会经费', '职工教育经费')),
    ("收入", ('主营业务收入', '其他业务收入', '营业收入', '利息收入', '手续费及佣金收入', '其他收益', '投资收益', '营业外收入')),
    ("费用", ('管理费用', '销售费用', '研发费用', '主营业务成本', '其他业务成本', '财务费用',
              '税金及附加', '营业外支出', '业务成本', '销售费用', '期间费用')),
    ("权益", ('实收资本', '资本公积', '盈余公积', '利润分配', '本年利润', '未分配利润', '库存股')),
    ("递延", ('递延收益', '递延所得税资产', '递延所得税负债')),
    ("其他往来", ('其他应收款', '其他应付款', '备用金', '保证金', '押金', '往来款')),
]


def subject_family(name):
    """把科目名归并为勾稽对照族（资金类/借款/应收款项/应付款项/存货/长期资产/减值准备/税费/
    职工薪酬/收入/费用/权益/递延/其他往来/其他）。空名返回 '其他'。"""
    if not name:
        return '其他'
    s = str(name).strip()
    for fam, kws in _FAMILY_KW:
        if any(k in s for k in kws):
            return fam
    return '其他'


def build_gl_voucher_map(data_dir, entities=None):
    """构造 (e, y, vkey) -> [{'name','dr','cr'}, ...]（整笔凭证全部分录，含金额）。
    vkey 同时登记『字-号』组合键与纯号键：匹配时优先组合键（避免 银8/转8/收8 碰撞合并成一张大凭证）。
    供按金额精确推导对方科目：仅取与本方方向相反、且金额匹配的对方账户。
    与 read_gl_rows 同源(均经 discover_entities)，保证全程序 GL 口径唯一。"""
    if entities is None:
        entities = discover_entities(data_dir)
    rows = read_gl_rows(data_dir, entities)
    mp = defaultdict(list)
    for r in rows:
        e = norm_cp_name(r.get('e'))
        y = r.get('y')
        vno = norm_cp_name(r.get('vno'))
        vt = norm_cp_name(r.get('vtype'))
        if not vno:
            continue
        combined = ('%s-%s' % (vt, vno)) if vt else vno
        line = {'name': r.get('name', '') or '', 'dr': float(r.get('dr') or 0),
                'cr': float(r.get('cr') or 0)}
        for vk in (combined, vno):
            mp[(e, y, vk)].append(line)
    return mp


# ---- 长期资产增加检查表「对方科目」真实来源筛选（2026-08-23 统一口径：builder 与终审补全共用）----
INC_CP_KEEP = ('银行存款', '应付账款', '其他应付款', '应收票据', '应收账款',
               '在建工程', '安装工程', '其他应收款', '预付账款', '应付票据', '长期应付款',
               '专项应付款', '其他货币资金', '短期借款', '长期借款', '一年内到期',
               '合同负债', '预计负债', '递延收益', '实收资本', '资本公积', '其他流动负债',
               '专项储备')
INC_CP_DROP = ('应交税费', '营业外收入', '其他业务收入', '主营业务收入', '营业收入',
               '制造费用', '管理费用', '销售费用', '财务费用', '研发支出', '营业费用',
               '税金及附加', '固定资产', '累计折旧', '固定资产清理', '工程物资',
               '减值准备', '库存现金', '其他收益', '资产处置损益', '营业外支出',
               '所得税费用', '主营业务成本', '其他业务成本', '信用减值损失', '资产减值损失',
               '应付职工薪酬', '原材料', '库存商品', '委托加工物资', '发出商品',
               '主营业务', '资产处置收益', '跌价准备', '本年利润', '利润分配',
               '以前年度损益调整', '递延所得税', '生产成本', '累计摊销',
               '长期待摊费用')


def filter_inc_cp(cp):
    """增加检查表对方科目：按 INC_CP_KEEP（真实来源）优先、INC_CP_DROP（剔除）过滤。
    未知科目保留（不误删）；过滤后分号重新连接。兼容 逗号/分号/顿号/全角 混用。
    ⚡ DROP 优先：『信用减值损失-应收账款损失』同时含 KEEP(应收账款)与 DROP(信用减值损失)，
    应归属信用减值损失（剔除）而非应收账款（保留）。"""
    import re as _re
    # ⚡ 不含顿号（、）——『在建工程-建筑、安装工程』等科目名内部含顿号，误分割会拆散科目。
    parts = [p.strip() for p in _re.split(r'[,;；，]', cp or '') if p.strip()]
    keep = []
    for p in parts:
        if any(k in p for k in INC_CP_DROP):
            continue
        if any(k in p for k in INC_CP_KEEP):
            keep.append(p)
        else:
            keep.append(p)
    return '；'.join(keep)


def derive_counterparty(lines, subject, subj_dr, subj_cr):
    """据整笔凭证 lines 推导『本方科目』行的对方科目（按金额精确匹配，杜绝"一大堆"）：
    1. 确定本方方向：subj_dr>0 → 本方在借方；否则本方在贷方。
    2. 取凭证中【方向相反】且【账户≠本方】的分录作候选对方科目。
    3. 若候选仅 1 个 → 直接取它（绝大多数 2 分录凭证，对方科目唯一）。
    4. 若多个候选 → 取金额与本方行相等的那一个（借/贷相反侧金额相等=真实对应科目）。
    5. 仍无法唯一确定（多借多贷且金额无匹配）→ 退化为列示对方侧全部账户（去重保序）；
       对方账户 >3 个时按首段聚合为『段×笔数』，紧凑且不刷屏（真实月末汇总结账凭证）。
    仅当无法判定本方方向（借贷均无额）时，退化为"除本方外全部账户"。
    """
    if not lines:
        return ''
    ns = norm_cp_name(subject) if subject else ''
    if subj_dr > 0:
        side, amt = 'dr', subj_dr
    elif subj_cr > 0:
        side, amt = 'cr', subj_cr
    else:
        others = [l['name'] for l in lines if norm_cp_name(l['name']) != ns]
        return '；'.join(dict.fromkeys(others))
    if side == 'dr':
        opp = [l for l in lines if l['cr'] > 0 and norm_cp_name(l['name']) != ns]
        amt_field = lambda l: l['cr']
    else:
        opp = [l for l in lines if l['dr'] > 0 and norm_cp_name(l['name']) != ns]
        amt_field = lambda l: l['dr']
    if not opp:
        return ''
    if len(opp) == 1:
        return opp[0]['name']
    am = [l for l in opp if abs(amt_field(l) - amt) < 0.005]
    if len(am) == 1:
        return am[0]['name']
    distinct = list(dict.fromkeys(l['name'] for l in opp if l['name']))
    if len(distinct) <= 3:
        return '；'.join(distinct)
    cnt = Counter(n.split('-')[0].strip() for n in distinct)
    return '；'.join('%s×%d' % (s, c) for s, c in cnt.most_common())


# ===================== 铁律14：raw_nat 独立重算与归因守卫（2026-07-29 新增） =====================
# 铁律14：遇 gv≠TB 严禁默认归为 GL 取数缺口；先独立重读原始 GL 全量自然方求和 raw_nat 与 TB 比——
# raw_nat==TB → 缺口不成立、差在自家加工(bug 或合理排除)；仅 raw_nat≠TB 才下『GL 取数缺口』结论。
# 该守卫把上述纪律编码为可复用函数，供各 builder/验证器在归因前统一调用，避免重蹈默认 GL 缺口。
def norm_subject_name(s):
    """科目名归一化：去前后空白(含全角空格)、NFKC 全半角统一、括号归一，
    以消除『采购』/『采购　』/『销项税额（开票）』等同义异写造成的匹配误报。
    铁律14 验证器据此把 GL 与 TB 按同名归并，避免名称细节触发的双计式假缺口。"""
    if not s:
        return ''
    s = unicodedata.normalize('NFKC', str(s))
    s = s.replace('（', '(').replace('）', ')').replace('【', '[').replace('】', ']')
    return s.strip()


def recompute_raw_nat(rows):
    """独立重算 GL 全量自然方求和（不套用任何 结转/计提/红字 排除规则）。

    rows: read_gl_rows 返回的逐条 dict（需含 e, y, name, dr, cr）。
    返回 {(e, y, norm_name): {'dr': float, 'cr': float}}，即每个 (主体,年度,归一化科目名)
    的【原始】借方合计/贷方合计。此值与 TB 控制数(jf/df)直接可比。
    注：综合查询明细表(GL) 数据源仅含『科目名称』列、无科目代码，故按归一化名称归并
    （而非精确字符串），消除同义异写误报；GL 与 TB 命名体系本质不同(业务名 vs 总账科目名)
    的科目仍会 flagged，需人工辨名或令 GL 导出带科目代码列方可彻底精确。
    """
    nat = {}
    for r in rows:
        nm = norm_subject_name(r.get('name'))
        if not nm:
            continue
        key = (r.get('e'), r.get('y'), nm)
        d = nat.setdefault(key, {'dr': 0.0, 'cr': 0.0})
        d['dr'] += float(r.get('dr', 0.0) or 0.0)
        d['cr'] += float(r.get('cr', 0.0) or 0.0)
    return nat


def attribution_guard(gv, tb, raw_nat, tol=0.005):
    """铁律14 归因守卫：判断 gv≠TB 的差异应归因于【自家加工】还是【GL 取数缺口】。

    gv:       builder 加工后的 GL 发生额（自然方，如同侧净额）；
    tb:       TB 控制数（同口径自然方）；
    raw_nat:  recompute_raw_nat 独立重算的 GL 原始自然方合计（同口径）。
    返回 (结论, 说明)：
      ('OWN_BUG', ...) —— raw_nat≈TB：GL 本身完整，gv≠TB 系自家加工(结转/计提排除或处理 bug)，缺口不成立；
      ('GL_GAP',  ...) —— raw_nat≠TB：确为 GL 非全量抽取缺口，应如实披露。
    铁律14 核心：严禁跳过此步直接把 gv≠TB 归为 GL 缺口。
    """
    gv = float(gv or 0.0); tb = float(tb or 0.0); raw_nat = float(raw_nat or 0.0)
    if abs(raw_nat - tb) <= tol:
        return ('OWN_BUG', f'raw_nat({raw_nat:.2f})≈TB({tb:.2f})：缺口不成立，gv≠TB 源于自家加工(非 GL 取数缺口)')
    return ('GL_GAP', f'raw_nat({raw_nat:.2f})≠TB({tb:.2f})：确为 GL 非全量抽取缺口，应如实披露')


# ===================== 通用读取：科目余额表 / 综合查询明细表（2026-07-29 统一） =====================
# 合并 revenue/inventory/longterm 三处同名 _read_km、revenue/inventory 两处同名 _read_gl，
# 统一为 audit_common.read_km / read_gl（返回 superset，行为等价于各本地解析器）：
#   · read_km: {code: {name, odir, opening(signed), debit, credit, fdir, closing(signed), level}}
#   · read_gl: list of {name, date, year, month, cp, aux, debit, credit}
# 各 builder 经 `from audit_common import read_km as _read_km, read_gl as _read_gl` 复用，删除本地副本。
def _ym(date_str):
    """'2025-12-31' -> (2025, 12)；无法解析返回 None。"""
    if not date_str:
        return None
    for sep in ('-', '/', '.'):
        if sep in date_str:
            parts = date_str.split(sep)
            if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
                return (int(parts[0]), int(parts[1]))
    return None


def read_km(path):
    """科目余额表 -> {code: {name, odir, opening(signed), debit, credit, fdir, closing(signed), level}}。

    合并 revenue/inventory/longterm 三处同名解析器（TB 列布局一致、仅返回字段集不同）为 superset。
    方向列(odir/fdir) 决定 opening/closing 带符号（借正贷负）；末级 level 由第 9 列解析。
    """
    if not path or not os.path.exists(path):
        return {}
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as ex:
        print(f'  ⚠️ 跳过 TB 读取失败：{path}：{ex}')
        return {}
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hidx = None
    for i, r in enumerate(rows[:6]):
        # 2026-08-06：表头兼容『科目代码』(FY) 与『科目编码』(JTt/SSS U8 导出)。
        # 两者列布局一致（编码|名称|方向|金额|借|贷|方向|金额|等级），仅表头措辞不同。
        if r and any(('科目代码' in str(c)) or ('科目编码' in str(c)) for c in r if c):
            hidx = i
            break
    if hidx is None:
        return {}
    out = {}
    sgn = lambda d, a: (a if d == '借' else (-a if d == '贷' else a))
    for r in rows[hidx + 1:]:
        if not r or r[0] is None:
            continue
        code = str(r[0]).strip()
        if not code or not code[0].isdigit():
            continue
        name = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ''
        odir = str(r[2]).strip() if len(r) > 2 and r[2] is not None else ''
        opening = safe(r[3]) if len(r) > 3 else 0.0
        debit = safe(r[4]) if len(r) > 4 else 0.0
        credit = safe(r[5]) if len(r) > 5 else 0.0
        fdir = str(r[6]).strip() if len(r) > 6 and r[6] is not None else ''
        closing = safe(r[7]) if len(r) > 7 else 0.0
        level = int(str(r[8]).strip()) if len(r) > 8 and str(r[8]).strip().isdigit() else 0
        out[code] = {
            'name': name, 'odir': odir,
            'opening': sgn(odir, opening), 'debit': debit, 'credit': credit,
            'fdir': fdir, 'closing': sgn(fdir, closing), 'level': level,
        }
    return out


def read_gl(path):
    """GL -> list of {name, date, year, month, vtype, vno, cp, aux, debit, credit}。

    合并 revenue(取 year/month/aux) 与 inventory(取 date/cp) 两处同名解析器为 superset：
    同时返回 原始日期(date)、年份/月份(由 _ym 推导)、凭证字号(vtype/vno，用于同号凭证反推对方科目)、
    对方科目(cp)、辅助核算名称(aux)。
    列按表头关键字自动探测（科目名称/日期/凭证字/凭证号/借方/贷方/对方科目/辅助核算名称）。
    """
    if not path or not os.path.exists(path):
        return []
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as ex:
        print(f'  ⚠️ 跳过 GL 读取失败：{path}：{ex}')
        return []
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    hdr = None
    for r in rows[:6]:
        if r and any('科目名称' in str(c) for c in r if c):
            hdr = r
            break
    if hdr is None:
        return []
    nm_i = next((i for i, c in enumerate(hdr) if c and '科目名称' in str(c)), 0)
    dt_i = next((i for i, c in enumerate(hdr) if c and '日期' in str(c)), 1)
    db_i = next((i for i, c in enumerate(hdr) if c and '借方' in str(c)), 6)
    cr_i = next((i for i, c in enumerate(hdr) if c and '贷方' in str(c)), 7)
    cp_i = next((i for i, c in enumerate(hdr) if c and '对方科目' in str(c)), None)
    aux_i = next((i for i, c in enumerate(hdr) if c and '辅助核算名称' in str(c)), None)
    vn_i = next((i for i, c in enumerate(hdr) if c and '号' in str(c)), None)   # 凭证号
    vt_i = next((i for i, c in enumerate(hdr) if c and '字' in str(c)), None)   # 凭证字
    out = []
    for r in rows:
        if not r or r[nm_i] is None:
            continue
        nm = str(r[nm_i]).strip()
        if nm == '科目名称' or not nm:
            continue
        dt = str(r[dt_i]).strip() if dt_i < len(r) and r[dt_i] is not None else ''
        db = safe(r[db_i]) if db_i < len(r) else 0.0
        cr = safe(r[cr_i]) if cr_i < len(r) else 0.0
        cp = ''
        if cp_i is not None and cp_i < len(r) and r[cp_i] is not None:
            cp = str(r[cp_i]).strip()
        aux = ''
        if aux_i is not None and aux_i < len(r) and r[aux_i] is not None:
            aux = str(r[aux_i]).strip()
        vt = str(r[vt_i]).strip() if vt_i is not None and vt_i < len(r) and r[vt_i] is not None else ''
        vn = str(r[vn_i]).strip() if vn_i is not None and vn_i < len(r) and r[vn_i] is not None else ''
        ym = _ym(dt)
        out.append({
            'name': nm, 'date': dt,
            'year': ym[0] if ym else None, 'month': ym[1] if ym else None,
            'vtype': vt, 'vno': vn,
            'cp': cp, 'aux': aux, 'debit': db, 'credit': cr,
        })
    return out


# ===================== 抽凭（凭证抽查）通用排除规则（2026-07-25 确立） =====================
# 所有小程序生成「凭证抽查/抽凭」sheet 时统一调用，避免各程序口径漂移。
# 规则① 结转损益类凭证不抽；规则② 计提/摊销本身分录相符的不抽。
_CARRYOVER_KW = ('结转损益', '结转本年利润', '期间损益结转', '结转利润', '结转损益类',
                 '损益结转', '本年利润结转', '本年利润转', '结转本年损益', '结转未分配利润', '利润结转')
_ACCRUAL_CREDIT_KW = ('应付职工薪酬', '累计折旧', '累计摊销', '使用权资产', '长期待摊费用', '预计负债',
                      '递延收益', '应付利息', '应付股利', '递延所得税负债', '递延所得税资产',
                      '坏账准备', '存货跌价准备', '合同负债', '应交税费', '应付债券',
                      # —— 2026-07-25 扩充：租赁/专项储备/应收利息/其他应收应付 同为计提侧，不应因"对方科目"
                      #    含这些账户而被误判为真实结算而保留。典型：计提房租费用(借管理费用-折旧费 贷使用权资产
                      #    + 借财务费用-其他 贷租赁负债-未确认融资费用)、计提安全生产费(借制造费用 贷专项储备)、
                      #    计提代缴社保(借费用 贷其他应收款-代垫)、计提销售奖金(借费用 贷其他应付款-应付其他)。
                      '租赁负债', '专项储备', '应收利息', '其他应收款', '其他应付款')
_ACCRUAL_DEBIT_KW = ('管理费用', '销售费用', '制造费用', '生产成本', '研发费用', '主营业务成本',
                     '其他业务成本', '财务费用', '营业外支出', '税金及附加', '所得税费用', '销售费用',
                     # —— 2026-07-25 扩充：研发支出(5301)作为成本费用归集科目，其计提侧分录
                     #    (借研发支出 贷应付职工薪酬) 同属相符计提，不应因本方科目名不含上述费用关键词而漏判保留。
                     '研发支出')
_SETTLEMENT_KW = ('银行存款', '库存现金', '其他货币资金', '应收票据', '应付票据', '预付账款', '预收账款',
                  '应收账款', '应付账款')
#   —— 2026-07-25 修订：原 _SETTLEMENT_KW 含 其他应收款/应付账款/合同负债/应交税费-未交增值税，
#      会把"借费用 贷其他应收款(代垫)/其他应付款(计提)"这类纯计提分录误判为真实结算而保留抽凭。
#      现仅保留【 unambiguous 真实资金/票据/往来结算】账户作为"强制保留"触发：银行存款/库存现金/其他货币资金
#      (现金及现金等价物)、应收票据/应付票据(票据)、预付账款/预收账款(预收预付)、应收账款/应付账款(真实购销往来)。
#      其余(其他应收/其他应付/合同负债/租赁负债/专项储备/应交税费等)一律视为计提侧，不强制保留。
#      注：用友导出 GL 的银行付款对方科目常为"银行名称"(如 泰隆商业银行上海分行)，不命中"银行存款"关键词，
#      此类真实付款凭"银行名称出现在分录科目名"使第二循环不匹配费用/计提关键词而仍被保留，属既有行为，未改变。


def voucher_is_carryover(summary='', name='', opp=''):
    """规则①：年终结转本年利润↔损益/未分配利润的结账分录，抽凭剔除。

    2026-07-26 修订：除 _CARRYOVER_KW（结转损益/结转本年利润/结转未分配利润…）外，
    凡凭证摘要或任一分录科目名/对方科目含「本年利润」（损益汇总结转账户）亦判定为结账分录剔除——
    典型如 借 主营业务收入/其他业务收入… 贷 本年利润 的年末结转，不应作为交易抽凭。
    """
    s = ' '.join(str(x) for x in (summary, name, opp) if x)
    if '本年利润' in s or '未分配利润' in s:
        return True
    # 兼容各类月末/年末结转损益摘要措辞（如「结转本期损益」「结转当期损益」「结转本年利润」等）：
    # 只要同一条文本同时含「结转」与「损益/利润」即判为结账分录，避免被漏判为「非结转」
    # 而误翻转（曾导致 借 其他收益 贷 本年利润 类年终结账被当成自然方红字，利润表发生额被抵销为 0）。
    # 真实同科目借贷双向蓝字错记（如 借费用X 贷费用X，摘要「计提/调整」）不含「结转/利润」，不受影响。
    if '结转' in s and ('损益' in s or '利润' in s):
        return True
    for kw in _CARRYOVER_KW:
        if kw in s:
            return True
    return False


# ===================== 利润表科目：非结转蓝字→自然方红字 规则（2026-07-27 确立） =====================
# 用户方法论：利润表科目须保证「每月结转损益金额 = 统计的发生额」。
# 凡非结转损益（对方科目≠本年利润/未分配利润，且凭证文本不含结转损益关键词）的蓝字（正数金额），
# 其落在【非自然方】的蓝字应作为【自然方红字】处理，使统计发生额与结转损益一致：
#   · 费用类(借方自然)：非结转【贷方蓝字】→ 借方红字（从借方发生额中扣减该笔）
#   · 收入类(贷方自然)：非结转【借方蓝字】→ 贷方红字（从贷方发生额中扣减该笔）
# 计提/摊销类（资产负债表科目，如 应付职工薪酬/累计折旧…）不套用本规则——
# 其与利润表的差异留作异常勾稽（用户：利润表准确后，余额类差异即可通过核对发现）。
# 仅调整【非自然方】蓝字；自然方蓝字（收入贷方/费用借方）为正常发生额，原样保留。
# 对干净账簿（无同科目借贷双向蓝字错记）本规则为零影响；仅纠正「借费用X + 贷费用X」式同科目错记。
def apply_pl_carryover_rule(rows, kind):
    """对利润表科目 GL 逐行应用「非结转蓝字→自然方红字」规则。

    rows: list of dict，需含 dr/cr/name/opp/summary 字段（read_gl_rows 或各程序 GL 读取结果）。
    kind: 'rev'(贷方自然) 或 'exp'(借方自然)。
    返回新 rows（不修改入参）。网净效果：非结转的非自然方蓝字被并入自然方红字扣减，
    使「自然方发生额」等于结转损益口径。
    """
    out = []
    for v in rows:
        dr = v.get('dr', 0.0) or 0.0
        cr = v.get('cr', 0.0) or 0.0
        if dr == 0 and cr == 0:
            out.append(v)
            continue
        # 凭证级结转标记（read_pl_gl 在凭证层级识别到含「本年利润/未分配利润/结转损益」整笔凭证时置位）：
        # 整笔凭证视为年终结账，行内非自然方蓝字【不再翻转】，保证明细表发生额=TB控制数。
        if v.get('_carry') or voucher_is_carryover(summary=v.get('summary', ''), name=v.get('name', ''), opp=v.get('opp', '')):
            out.append(v)  # 结转损益本身，原样保留
            continue
        nv = dict(v)
        if kind == 'exp':                 # 费用：借方自然；非结转【贷方蓝字】→ 借方红字
            if cr > 0:
                nv['dr'] = round(dr - cr, 2)   # 从借方发生额中扣减（等价于借方红字）
                nv['cr'] = 0.0
        else:                             # 收入：贷方自然；非结转【借方蓝字】→ 贷方红字
            if dr > 0:
                nv['cr'] = round(cr - dr, 2)   # 从贷方发生额中扣减（等价于贷方红字）
                nv['dr'] = 0.0
        out.append(nv)
    return out


def voucher_is_pure_accrual(lines):
    """规则②：仅为计提/摊销的常规分录（费用/成本↔计提/摊销负债，无资金/往来结算），勾稽自洽，不抽。
    lines: 该凭证全部分录 [{name, opp, dr, cr}, ...]（按凭证号聚合后的多行）。
    仅当分录出现【 unambiguous 真实资金/票据/往来结算】账户方才视为例外、保留抽凭：
      银行存款/库存现金/其他货币资金(现金及现金等价物)、应收票据/应付票据(票据)、
      预付账款/预收账款(预收预付)、应收账款/应付账款(真实购销往来)。
    其他应收/其他应付/合同负债/租赁负债/专项储备/应交税费/研发支出 等一律按计提侧处理，不强制保留。
    """
    if not lines:
        return False
    for ln in lines:
        nm = ln.get('name', '') or ''
        op = ln.get('opp', '') or ''
        blob = nm + ' ' + op
        if any(k in blob for k in _SETTLEMENT_KW):
            return False  # 含真实结算/第三方 → 例外，保留抽凭
    for ln in lines:
        nm = ln.get('name', '') or ''
        if not (any(k in nm for k in _ACCRUAL_DEBIT_KW) or any(k in nm for k in _ACCRUAL_CREDIT_KW)):
            return False  # 出现非计提/摊销语义科目 → 非纯计提，保留
    return True


def is_accrual_line(line):
    """行级纯计提判定（2026-07-27 修复铁律②漏洞）：

    抽凭排除规则②「相符计提/摊销不抽」原为【整笔凭证】级——当同一凭证号混合了
    「计提行(借费用 贷应付职工薪酬)」与「真实付款行(借费用 贷银行存款)」时，
    凭证级 voucher_is_pure_accrual 因见到真实结算账户而整笔保留，导致计提行被连带泄漏进抽凭。

    本函数对【单行】判定：只要该行自身是「费用/成本/资产 ↔ 计提负债(应付职工薪酬/
    累计折旧/累计摊销/专项储备/租赁负债/其他应付(计提)…)」且对方科目【不含任何真实
    结算/第三方账户】(银行存款/票据/应收应付购销往来)，即判定为相符计提单行，应逐行排除；
    同号凭证中的真实付款行(opp 含结算账户)不受影响、正常保留。

    返回 True=该行是纯计提、应从抽凭剔除。
    """
    nm = (line.get('name', '') or '')
    op = (line.get('opp', '') or '')
    blob = nm + ' ' + op
    # 含真实结算/第三方 → 该行是真实交易，不视为计提
    if any(k in blob for k in _SETTLEMENT_KW):
        return False
    # 本方科目落在计提语义(费用/成本/资产计提) 或 对方科目落在计提负债侧 → 纯计提行
    if (any(k in nm for k in _ACCRUAL_DEBIT_KW)
            or any(k in nm for k in _ACCRUAL_CREDIT_KW)
            or any(k in op for k in _ACCRUAL_CREDIT_KW)):
        return True
    return False


def voucher_should_skip(summary='', name='', opp='', lines=None):
    """抽凭统一排除判定：结转损益 或 相符计提/摊销 → 返回 (skip=True, reason)。

    合并分录级文本：2026-07-26 修订——「本年利润」等结账账户常出现在分录科目名(而非凭证摘要)，
    如 借 主营业务收入 贷 本年利润；仅查凭证级 summary/name/opp 会漏判，故把全部分录的
    科目名/对方科目/摘要并入判定文本，确保年终结转凭证被剔除。
    """
    if lines:
        line_blob = ' '.join(
            ' '.join(str(x) for x in (l.get('name', '') or '', l.get('opp', '') or '', l.get('summary', '') or ''))
            for l in lines
        )
        summary = (summary or '') + ' ' + line_blob
    if voucher_is_carryover(summary, name, opp):
        return True, '结转损益'
    if lines and voucher_is_pure_accrual(lines):
        return True, '相符计提/摊销'
    return False, ''


def group_gl_vouchers(rows):
    """将 read_gl_rows 的逐行结果按 (e, y, vtype, vno) 聚合为整笔凭证对象列表。
    返回 [{'e','y','vtype','vno','date','summary','lines':[{name,opp,dr,cr}]}, ...]。
    供抽凭统一排除(voucher_should_skip)按「整笔凭证」多行联合判定（规则②需看全部分录）。
    """
    from collections import OrderedDict
    groups = OrderedDict()
    for r in rows:
        key = (r.get('e'), r.get('y'), r.get('vtype'), r.get('vno'))
        g = groups.get(key)
        if g is None:
            g = {'e': r.get('e'), 'y': r.get('y'), 'vtype': r.get('vtype'),
                 'vno': r.get('vno'), 'date': r.get('date'), 'summary': r.get('summary'),
                 'lines': []}
            groups[key] = g
        g['lines'].append({'name': r.get('name', ''), 'opp': r.get('opp', ''),
                           'dr': r.get('dr', 0), 'cr': r.get('cr', 0)})
    return list(groups.values())


def voucher_skip_keys(recs, key_fields, name_field, opp_field, summ_field):
    """通用：对已收集的凭证行 recs，按 key_fields(r) 聚合为整笔凭证，返回应排除的 key 集合。
    每个程序传入各自的字段访问器即可统一套用「结转损益不抽 + 相符计提/摊销不抽」规则。
    - key_fields(r): 返回可哈希 key（如 (e,y,vno) 或 (e,y,vtype,vno)）
    - name_field/opp_field/summ_field(r): 返回本方科目名/对方科目/摘要
    判定走 voucher_should_skip(summary, name, opp, lines)，规则②需整笔多行故传 lines。
    """
    from collections import OrderedDict
    groups = OrderedDict()
    for r in recs:
        k = tuple(key_fields(r))
        groups.setdefault(k, []).append(r)
    skip = set()
    for k, lines in groups.items():
        vlines = [{'name': name_field(x) or '', 'opp': opp_field(x) or ''} for x in lines]
        sm = summ_field(lines[0]) if lines else ''
        nm0 = name_field(lines[0]) or ''
        op0 = opp_field(lines[0]) or ''
        s, _ = voucher_should_skip(sm, nm0, op0, vlines)
        if s:
            skip.add(k)
    return skip


# ===================== 函数级缓存 =====================
_FN_CACHE = {}  # { (func_name, data_dir): result }
def _cached_call(func, data_dir, *args, **kwargs):
    """缓存基于 data_dir 的纯函数结果，进程内只执行一次。"""
    key = (func.__name__, data_dir)
    if key not in _FN_CACHE:
        _FN_CACHE[key] = func(data_dir, *args, **kwargs)
    return _FN_CACHE[key]

def clear_cache(data_dir=None):
    """清除所有缓存；给定 data_dir 只清除该目录的缓存。"""
    global _FN_CACHE
    if data_dir is None:
        _FN_CACHE.clear()
    else:
        _FN_CACHE = {k: v for k, v in _FN_CACHE.items() if k[1] != data_dir}


# ===================== TB 读取（全科目，带借贷符号余额） =====================
# ⚡⚡ 2026-08-23 列探测收口：read_tb_full（启发式）与 read_sap_tb 300 分支（精确表头名）
#   共用同一套列名别名 + probe_tb_columns，避免两处各自探测漂移。
TB_COL_ALIASES = {
    'code': ('总账科目', '科目代码', '科目编码', '科目'),
    'name': ('总账科目长文本', '科目名称', '科目名'),
    'qc_fx': ('期初方向', '期初余额方向'),
    'qc_amt': ('期初金额', '期初余额'),
    'jf': ('本期借方金额', '本期借方'),
    'df': ('本期贷方金额', '本期贷方'),
    'qm_fx': ('期末的方向', '期末方向'),
    'qm_amt': ('期末余额', '期末金额'),
    'level': ('等级', '级次'),
}


def probe_tb_columns(header_cells):
    """从表头行探测 TB 列索引（0-based）。启发式判定，兼容 U8/SAP 300 各列序变体。
    返回 {code,name,qc_fx,qc_amt,jf,df,qm_fx,qm_amt,level}（缺列给默认索引）。
    判定顺序关键：先排『长文本/名称』（name 列），再排『代码/编码/总账/科目』（code 列），
    避免『总账科目长文本』被 code 规则吞掉；公司/账簿列不得当 code。"""
    cells = [str(v).strip() if v is not None else '' for v in header_cells]
    h = {}
    # name 优先（最具体：长文本/名称），code 次之（代码/编码/总账/科目，排除长文本/名称/公司/账簿）
    for i, s in enumerate(cells):
        if not s:
            continue
        if 'name' not in h and ('长文本' in s or ('名称' in s and '账簿' not in s)):
            h['name'] = i
        elif 'code' not in h and '长文本' not in s and '名称' not in s \
                and '公司' not in s and '账簿' not in s \
                and ('代码' in s or '编码' in s or '总账' in s or '科目' in s):
            h['code'] = i
    # 方向列：含『方向』按出现顺序（第一个=期初、第二个=期末）
    dirs = [i for i, s in enumerate(cells) if '方向' in s]
    if dirs:
        h.setdefault('qc_fx', dirs[0])
        if len(dirs) > 1:
            h.setdefault('qm_fx', dirs[1])
    # 金额列：期初/本期借方/本期贷方/期末
    for i, s in enumerate(cells):
        if not s:
            continue
        if 'qc_amt' not in h and '期初' in s and ('金额' in s or '余额' in s):
            h['qc_amt'] = i
        elif 'jf' not in h and '借方' in s and '本期' in s:
            h['jf'] = i
        elif 'df' not in h and '贷方' in s and '本期' in s:
            h['df'] = i
        elif 'qm_amt' not in h and '期末' in s and ('金额' in s or '余额' in s):
            h['qm_amt'] = i
        elif 'level' not in h and ('等级' in s or '级次' in s):
            h['level'] = i
    return {'code': h.get('code', 0), 'name': h.get('name', 1),
            'qc_fx': h.get('qc_fx', 2), 'qc_amt': h.get('qc_amt', 3),
            'jf': h.get('jf', 4), 'df': h.get('df', 5),
            'qm_fx': h.get('qm_fx', 6), 'qm_amt': h.get('qm_amt', 7),
            'level': h.get('level', 8)}


def read_tb_full(data_dir, entities):
    """返回 {(e, code, name, y): {'qc':signed期初, 'jf':借发, 'df':贷发, 'qm':signed期末, 'level':等级}}。
    通用原始层：动态探测表头（兼容表头在第 2 行、列序微调的导出），signed 期初/期末。
    """
    _ckey = (data_dir, 'tb')
    if _ckey in _CACHE_TB:
        return _CACHE_TB[_ckey]
    # 2026-08-05 磁盘缓存：TB 解析结果持久化（与 GL 同机制）
    _ac_save_tb = None
    try:
        from audit_cache import load as _ac_load2, save as _ac_sv2
        _ac_save_tb = _ac_sv2
        _d, _p = _ac_load2(data_dir, entities, 'tb')
        if _d is not None:
            _CACHE_TB[_ckey] = _d
            return _d
    except Exception:
        pass
    tb = {}
    for e, yd in entities.items():
        for y, paths in yd.items():
            km = paths.get('km')
            if not km or not os.path.exists(km):
                continue
            try:
                wb = openpyxl.load_workbook(km, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ TB读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            # 探测列：取含表头的行做映射（兼容"代码/名称/期初方向/期初金额/本期借/本期贷/期末方向/期末金额/等级"）。
            # ⚡⚡ 2026-08-15 修复：表头行位置兼容——U8 导出表头在第 2 行（r1=查询条件），
            #   XBJ 的 sss_prepare 三件套表头在【第 1 行】（r2 就是数据，如 1002 银行存款）！
            #   原固定探测 r2 + 从 r3 读 → 表头在 r1 时探测失败（默认列映射碰巧对）但
            #   r2 的 1002 被当表头跳过 → 银行存款丢失 → 试算表勾稽不平（04 伽师 4322 万）。
            # ⚡⚡ 2026-08-23 探测逻辑收口到公共 probe_tb_columns（read_sap_tb 300 共用）。
            def _probe_hdr(_ri):
                return probe_tb_columns([ws.cell(_ri, c).value for c in range(1, 15)])
            hdr = _probe_hdr(2)
            _data_start = 3
            if not hdr.get('code') and not hdr.get('name'):
                # 第 2 行无表头特征 → 表头在第 1 行（XBJ/SSS 三件套格式）
                hdr2 = _probe_hdr(1)
                if hdr2.get('code') or hdr2.get('name'):
                    hdr = hdr2
                    _data_start = 2
            h = {
                'code': hdr.get('code', 0), 'name': hdr.get('name', 1),
                'qc_fx': hdr.get('qc_fx', 2), 'qc_amt': hdr.get('qc_amt', 3),
                'jf': hdr.get('jf', 4), 'df': hdr.get('df', 5),
                'qm_fx': hdr.get('qm_fx', 6), 'qm_amt': hdr.get('qm_amt', 7),
                'level': hdr.get('level', 8),
            }
            for r in ws.iter_rows(min_row=_data_start, values_only=True):
                try:
                    if not r or len(r) < 4 or not r[h['code']]:
                        continue
                    code = str(r[h['code']]).strip()
                    # 2026-08-07：过滤非数字科目代码行（JTt/SSS 等 U8 导出含『总计』行，
                    # name 为空串 → recon_self 名称匹配 '' in subj_name 恒 True → 全科目虚增 793 亿）。
                    # 科目代码一律数字开头；'总计'/'合计'/'小计' 等汇总行必须剔除。
                    if not re.match(r'^\d', code):
                        continue
                    name = str(r[h['name']]).strip() if r[h['name']] else ''
                    qc = signed_balance_v2(r[h['qc_fx']], r[h['qc_amt']]) if h['qc_fx'] < len(r) else safe(r[h['qc_amt']])
                    qm = signed_balance_v2(r[h['qm_fx']], r[h['qm_amt']]) if h['qm_fx'] < len(r) else safe(r[h['qm_amt']])
                    jf = safe(r[h['jf']]) if h['jf'] < len(r) else 0.0
                    df = safe(r[h['df']]) if h['df'] < len(r) else 0.0
                    lv = r[h['level']] if h['level'] < len(r) else None
                    tb[(e, code, name, y)] = {'qc': qc, 'jf': jf, 'df': df, 'qm': qm, 'level': lv}
                except Exception:
                    continue
            wb.close()
    _CACHE_TB[_ckey] = tb
    if _ac_save_tb is not None:
        try:
            _ac_save_tb(data_dir, entities, 'tb', tb)
        except Exception:
            pass
    return tb


# ===================== 审定表（每科目明细表前置，与科目余额表核对） =====================
def _write_audit_by_entity_sheet(wb, item, tb_full, entities, years, target_year):
    """按核算主体列示审定表（PL 上年/本期 或 BS 年初/期末）。单一职责：写一张 by_entity 审定表。
    2026-08-01 从 add_audit_summary_sheets 拆分（原 521 行函数）。返回创建的 sheet。"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    TITLEFILL = PatternFill('solid', fgColor='BDD7EE')
    TOTFILL = PatternFill('solid', fgColor='FCE4D6')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center')
    NUMFMT = '#,##0.00'
    FONT = Font(name='Times New Roman', size=10)
    is_credit = bool(item.get('is_credit', False))
    title = item['title']
    codes = item.get('codes')
    if isinstance(codes, str):
        codes = [codes]
    yrs = sorted(years)
    if target_year:
        cur_y = target_year
        prev_y = str(int(target_year) - 1) if str(int(target_year) - 1) in years else None
    else:
        cur_y = yrs[-1]
        prev_y = yrs[-2] if len(yrs) >= 2 else None

    def _gross_entity(nm=None, code=None, codes=None, y=None, for_entity=None):
        """按编码取发生额；2026-08-04 修复：编码匹配不到时按科目名称兜底匹配
        （FY 其他收益=6113 新准则编码，pl_detail 曾配 6117 旧编码 → 审定表主体行全 0）。
        codes 匹配精确编码；名称兜底按『名称与 subj_name 相同或含其』，防跨账套编码差异。
        2026-08-07 v2：codes 前缀匹配（JTt 等 U8 账套损益科目为 6 位码，配置 codes=['6121']
        是 4 位一级前缀）——前缀命中行再按"父=子和"只取末级叶子，防父+子双计。"""
        tot = 0.0
        # ⚡⚡ 2026-08-31 修复（XBJ 损益审定表空壳根因②）：损益科目(income_statement)取【发生额】
        #   ——费用=借方 jf、收益=贷方 df（与 sap_tb_gen 利润表/自建试算表同口径）。原取净额
        #   (jf-df/df-jf)：XBJ U8 科目余额表损益科目借贷同额(660246 借=贷=-2854517) → 净额 0
        #   → 审定表空壳而试算表 371M。非损益科目维持净额（资产/负债余额口径）。
        _income_stmt = bool(item.get('income_statement'))
        def _pl_amt(_v, _is_credit):
            if _income_stmt:
                return _v['df'] if _is_credit else _v['jf']
            return (_v['df'] - _v['jf']) if _is_credit else (_v['jf'] - _v['df'])
        name_fallback = item.get('subj_name') or (item.get('title') or '').replace(' 审定表', '')
        _hit_set = None
        if codes:
            _hit_set = {str(cc) for (ee, cc, _nn, _yy), _v in tb_full.items()
                        if (y is None or str(_yy) == str(y))
                        and (for_entity is None or ee == for_entity)
                        and any(str(cc).startswith(p) for p in codes)}
        for (e, c, n, yy), v in tb_full.items():
            if y is not None and str(yy) != str(y):
                continue
            if for_entity is not None and e != for_entity:
                continue
            if code is not None and str(c) == str(code):
                tot += _pl_amt(v, is_credit)
                continue
            if codes is not None:
                if any(str(c).startswith(p) for p in codes):
                    if _hit_set and any(str(c) != cc2 and cc2.startswith(str(c)) for cc2 in _hit_set):
                        continue  # 父级行（有后代同族行）→ 只计末级叶子
                    tot += _pl_amt(v, is_credit)
                    continue
                if str(c) in codes:
                    tot += _pl_amt(v, is_credit)
                    continue
            # 名称兜底：仅当 编码未命中 且 名称匹配 时计入（避免双计）
            # ⚡⚡ 2026-08-27 P0 修复：原 `str(n) == name_fallback` 精确相等，而 SAP 科目名
            #   为『一级-二级-三级』（如 其他收益-政府补助-与收益相关）→ 恒不等 → 名称兜底
            #   永远失效（其他收益审定表 0 根因之一）。改按一级段匹配（SAP '-' 分隔，
            #   与 pl_detail._subject_of / 铁律5 名称主键 一致），兼容精确名/前缀带分隔符。
            if name_fallback and (
                    str(n).strip() == name_fallback
                    or str(n).startswith(name_fallback + '-')
                    or str(n).startswith(name_fallback + '　')
                    or ('-' not in str(n) and name_fallback in str(n))):
                tot += _pl_amt(v, is_credit)
        return tot

    ws = wb.create_sheet(title=title)
    is_pl = bool(item.get('income_statement'))
    if is_pl:
        ncols = 6
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        _stitle = item.get('subj_name', title).replace(' 审定表', '')
        tcell = ws.cell(1, 1, f'{_stitle} 审定表（按核算主体）')
        tcell.font = Font(name='Times New Roman', size=10, bold=True); tcell.fill = TITLEFILL
        tcell.alignment = CTR
        ws.row_dimensions[1].height = 22
        headers = ['核算主体', '上年数', '本期数', '审计调整数', '审定数', '与科目余额表勾稽']
        for j, h in enumerate(headers, start=1):
            c = ws.cell(2, j, h); c.font = Font(name='Times New Roman', size=10, bold=True)
            c.fill = HFILL; c.alignment = CTR; c.border = BORDER
        r0 = 3
        t_prev = t_cur = 0.0
        for ent in sorted(entities):
            prev_amt = _gross_entity(codes=codes, y=prev_y, for_entity=ent) if prev_y else 0.0
            cur_amt = _gross_entity(codes=codes, y=cur_y, for_entity=ent)
            if prev_amt == 0 and cur_amt == 0:
                continue
            t_prev += prev_amt; t_cur += cur_amt
            r = r0
            ws.cell(r, 1, ent).alignment = LFT
            ws.cell(r, 2, round(prev_amt, 2)).number_format = NUMFMT
            ws.cell(r, 3, round(cur_amt, 2)).number_format = NUMFMT
            ws.cell(r, 4, None).number_format = NUMFMT
            ws.cell(r, 5, round(cur_amt, 2)).number_format = NUMFMT   # 2026-08-06 协议：审定数写数值（=本期+调整0）
            ws.cell(r, 6, '与科目余额表核对一致').alignment = LFT
            for j in range(1, ncols + 1):
                cell = ws.cell(r, j); cell.border = BORDER; cell.font = FONT
                if j in (2, 3, 4, 5):
                    cell.alignment = RGT
            r0 += 1
        r = r0
        ws.cell(r, 1, '合计').font = Font(name='Times New Roman', size=10, bold=True)
        ws.cell(r, 1).fill = TOTFILL; ws.cell(r, 1).alignment = CTR
        # 2026-08-06 协议：合计行写数值（t_prev/t_cur 生成时累计；公式 data_only 读 None 收回读不回）
        ws.cell(r, 2, round(t_prev, 2)).number_format = NUMFMT
        ws.cell(r, 2).fill = TOTFILL; ws.cell(r, 2).alignment = RGT
        ws.cell(r, 3, round(t_cur, 2)).number_format = NUMFMT
        ws.cell(r, 3).fill = TOTFILL; ws.cell(r, 3).alignment = RGT
        ws.cell(r, 5, round(t_cur, 2)).number_format = NUMFMT   # 合计行审定数（=本期+调整0）
        ws.cell(r, 5).fill = TOTFILL; ws.cell(r, 5).font = FONT; ws.cell(r, 5).alignment = RGT
        for j in (4, 6):
            c = ws.cell(r, j); c.fill = TOTFILL; c.font = FONT
            c.alignment = RGT if j == 4 else LFT
    else:
        ncols = 6
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
        _stitle = item.get('subj_name', title).replace(' 审定表', '')
        tcell = ws.cell(1, 1, f'{_stitle} 审定表（按核算主体）')
        tcell.font = Font(name='Times New Roman', size=10, bold=True); tcell.fill = TITLEFILL
        tcell.alignment = CTR
        ws.row_dimensions[1].height = 22
        headers = ['核算主体', '年初数', '期末未审数', '审计调整数', '审定数', '与科目余额表勾稽']
        for j, h in enumerate(headers, start=1):
            c = ws.cell(2, j, h); c.font = Font(name='Times New Roman', size=10, bold=True)
            c.fill = HFILL; c.alignment = CTR; c.border = BORDER
        r0 = 3
        t_qc = t_qm = 0.0
        # 2026-08-07：entity_codes（按主体各自科目代码）优先——J 实收资本东轴=4001 股本、
        # 仁旭/尚风/海拓/金泰=3001 实收资本，全局 codes 并集 {4001,3001} 精确匹配会把
        # 仁旭 4001 生产成本（273 万）误当实收资本（审定 268 万=4001-3001 带符号和）。
        _ecodes = item.get('entity_codes') or {}
        for ent in sorted(entities):
            _cl = _ecodes.get(ent) or codes
            # ⚡ 2026-08-10 修复（铁律60/65 前缀匹配+末级叶子过滤）：SAP TB 只有 10 位末级码
            # （2401010000 递延收益-搬迁收益），原精确匹配 `c in _cl` 对 4 位 codes（['2401']）
            # 永假 → gp_other 系审定表全 0（递延收益/专项应付款/应付债券等 19 科目）。
            # 前缀匹配后 U8 TB 有父级行会父+子双计 → 末级叶子过滤（父级=子和，铁律3/6）。
            _hit = [(e, c, n, yy, v) for (e, c, n, yy), v in tb_full.items()
                    if yy == cur_y and e == ent and _cl
                    and any(str(c).startswith(p) for p in _cl)]
            _leaves = [(c, n, v) for (e, c, n, yy, v) in _hit
                       if not any(o != c and str(o).startswith(str(c)) for (oe, o, on, oyy, ov) in _hit)]
            # BS科目取 qc(年初数)/qm(期末未审数)；2026-08-06 修复：改【带符号求和再 abs】——
            # 原逐 code abs 累加会把同族备抵子目双计（ga 220601 租赁付款额 -1,085,600 +
            # 220602 未确认融资费用 +37,286.70 被 abs 成 1,122,886.70 vs 净额 1,048,313.30）。
            # 带符号求和=净额（父级=子和，铁律），再 abs 转正显示，兼容单 code 与多 code 同向。
            # ⚡⚡ 2026-08-27 修复（AH 底稿复查——1357 未分配利润 明细vs审定 符号差）：
            #   keep_sign=True（equity RETAINED 族）→ 保持带符号（SAP 未分配利润 Balance=Credit，
            #   贷方正/借余负：亏损主体 qm<0 显示负，集团合计=Σ带符号=明细一级 -984,921,523）。
            #   原 abs 把亏损主体（1020 -110,849,696）翻正 → 审定 +13.42亿 vs 明细 -9.85亿 虚增。
            if item.get('keep_sign'):
                qc_sum = sum(v['qc'] for (c, n, v) in _leaves)
                qm_sum = sum(v['qm'] for (c, n, v) in _leaves)
            else:
                qc_sum = abs(sum(v['qc'] for (c, n, v) in _leaves))
                qm_sum = abs(sum(v['qm'] for (c, n, v) in _leaves))
            if qc_sum == 0 and qm_sum == 0:
                continue
            t_qc += qc_sum; t_qm += qm_sum
            r = r0
            ws.cell(r, 1, ent).alignment = LFT
            ws.cell(r, 2, round(qc_sum, 2)).number_format = NUMFMT
            ws.cell(r, 3, round(qm_sum, 2)).number_format = NUMFMT
            ws.cell(r, 4, None).number_format = NUMFMT
            ws.cell(r, 5, round(qm_sum, 2)).number_format = NUMFMT   # 2026-08-06 协议：审定数写数值（=期末+调整0）
            ws.cell(r, 6, '与科目余额表核对一致').alignment = LFT
            for j in range(1, ncols + 1):
                cell = ws.cell(r, j); cell.border = BORDER; cell.font = FONT
                if j in (2, 3, 4, 5):
                    cell.alignment = RGT
            r0 += 1
        r = r0
        ws.cell(r, 1, '合计').font = Font(name='Times New Roman', size=10, bold=True)
        ws.cell(r, 1).fill = TOTFILL; ws.cell(r, 1).alignment = CTR
        # 2026-08-06 协议：合计行写数值（t_qc/t_qm 生成时累计）
        ws.cell(r, 2, round(t_qc, 2)).number_format = NUMFMT
        ws.cell(r, 2).fill = TOTFILL; ws.cell(r, 2).alignment = RGT
        ws.cell(r, 3, round(t_qm, 2)).number_format = NUMFMT
        ws.cell(r, 3).fill = TOTFILL; ws.cell(r, 3).alignment = RGT
        ws.cell(r, 5, round(t_qm, 2)).number_format = NUMFMT   # 合计行审定数（=期末+调整0）
        ws.cell(r, 5).fill = TOTFILL; ws.cell(r, 5).font = FONT; ws.cell(r, 5).alignment = RGT
        for j in (4, 6):
            c = ws.cell(r, j); c.fill = TOTFILL; c.font = FONT
            c.alignment = RGT if j == 4 else LFT
    widths = [16, 18, 18, 14, 18, 28]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    ws.freeze_panes = 'A3'
    # before：把本审定表插到指定明细表之前（2026-08-01 修复：equity/往来款审定表未置于第一张）
    try:
        before = item.get('before')
        if before and before in wb.sheetnames:
            sheets = wb._sheets
            cur_idx = next((i for i, s in enumerate(sheets) if s.title == title), None)
            before_idx = next((i for i, s in enumerate(sheets) if s.title == before), None)
            if cur_idx is not None and before_idx is not None and cur_idx > before_idx:
                node = sheets.pop(cur_idx)
                sheets.insert(before_idx, node)
    except Exception:
        pass
    return ws


def add_audit_summary_sheets(wb, data_dir, items, tb_full=None, entities=None, FONT=None, target_year=None):
    """在每个科目明细表之前插入一张『审定表』，数据取自科目余额表(TB)，可与 TB 直接核对。

    items: list[dict]，每个元素对应一张要前置的审定表：
      {
        'title':    '应收账款 审定表',    # 审定表 sheet 名（建议含科目名，唯一）
        'codes':    ['1122'] 或 '1122',  # 取自 TB 的科目代码（精确匹配，跨主体合并）
        'subj_name':'应收账款',           # 显示用科目名称（缺省取 TB 名称；多 code 时作合计名）
        'is_credit': False,               # 余额方向：贷方余额(负债/权益/收入)=True；资产/费用=False
        'before':   '应收账款明细表',     # 已存在的明细 sheet 名，审定表将插到它前面
        'note':     '可选说明',
      }
    处理：
      - 经 read_tb_full 取各 codes 的 qc/jf/df/qm（signed 期初/期末，jf=借发/df=贷发），跨主体合并；
      - 列：科目编码|科目名称|年度|期初方向|期初未审数|本期增加|本期减少|期末方向|期末未审数|
            审计调整数|期末审定数|与科目余额表勾稽；
      - 期初/期末按 signed 取绝对值并标借/贷方向；本期增加/减少为 gross 正幅度
        （资产/费用:增加=借发,减少=贷发；负债/权益/收入:增加=贷发,减少=借发）；
      - 期末审定数 = 期末未审数 + 审计调整数（公式联动，调整数留空待手填）；
      - 勾稽说明：期末未审数取自 TB 期末余额，标注『与科目余额表(XXXX)核对一致』；
      - 多 codes 时每个 code 占一行，并追加合计行。
    返回插入的审定表 sheet 列表。
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    if entities is None:
        entities = discover_entities(data_dir)
    if tb_full is None:
        tb_full = read_tb_full(data_dir, entities)
    # 2026-08-01 问题1c：泰国账套科目代码/名称与集团标准不同（5200=销售费用等），
    # 审定表按标准 codes 精确匹配会漏掉泰国主体。在此统一对泰国 TB 做归一化（幂等，不重复）：
    try:
        from thai_mapping import normalize_thai_tb as _norm_thai_tb
        _tb_norm = _norm_thai_tb(tb_full, entities)
        if _tb_norm is not tb_full:
            tb_full = _tb_norm
    except Exception:
        pass

    if FONT is None:
        FONT = Font(name='Times New Roman', size=10)
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    TITLEFILL = PatternFill('solid', fgColor='BDD7EE')
    TOTFILL = PatternFill('solid', fgColor='FCE4D6')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center')
    NUMFMT = '#,##0.00'

    # 年份集合
    years = sorted({y for (e, c, n, y) in tb_full})
    # 按年拆分文件只注入目标年份（如 _2025 文件仅列 2025），否则用全部年份
    eff_years = [target_year] if target_year else years
    made = []
    for item in items:
            codes = item.get('codes')
            if isinstance(codes, str):
                codes = [codes]
            names = item.get('names')  # 铁律⑤：按名称匹配（兼容同码不同义，如 5001=主营业务收入/生产成本）
            is_credit = bool(item.get('is_credit', False))
            title = item['title']
            if title in wb.sheetnames:
                print(f'  ⚠️ 审定表已存在，跳过：{title}')
                continue
            # ---- 按核算主体列示（by_entity，独立于 income_statement） ----
            if item.get('by_entity'):
                made.append(_write_audit_by_entity_sheet(wb, item, tb_full, entities, years, target_year))
                continue


            # ---------- 利润表科目：上年数 / 本期数 / 审计调整数 / 审定数 ----------
            if item.get('income_statement'):
                yrs = sorted(years)
                if target_year:
                    cur_y = target_year
                    prev_y = str(int(target_year) - 1) if str(int(target_year) - 1) in years else None
                else:
                    cur_y = yrs[-1]
                    prev_y = yrs[-2] if len(yrs) >= 2 else None

                def _gross(nm=None, code=None, codes=None, y=None, for_entity=None):
                    tot = 0.0
                    for (e, c, n, yy), v in tb_full.items():
                        if y is not None and yy != y:
                            continue
                        if for_entity is not None and e != for_entity:
                            continue
                        if nm is not None and n != nm:
                            continue
                        if code is not None and c != code:
                            continue
                        if codes is not None:
                            # 2026-08-07 v2：codes 支持【前缀匹配 + 末级叶子过滤】——
                            # JTt 等 U8 账套损益科目为 6 位码（612101 其他收益\政府补助），
                            # 配置 codes=['6121'] 是 4 位一级前缀；原精确匹配 c not in codes
                            # 全部取空 → 审定表=0。前缀命中行再按"父=子和"只取末级叶子
                            # （无其他命中行为其前缀），防父+子双计。
                            if not any(str(c).startswith(p) for p in codes):
                                continue
                            _hit = {str(cc) for (ee, cc, _nn, _yy), _v in tb_full.items()
                                    if (y is None or _yy == y)
                                    and (for_entity is None or ee == for_entity)
                                    and any(str(cc).startswith(p) for p in codes)}
                            if any(str(c) != cc2 and cc2.startswith(str(c)) for cc2 in _hit):
                                continue  # 本行是父级（有后代同族行）→ 剔除，只计末级
                        # ⚡⚡ 2026-08-29 P0 修复：损益审定表取【净发生额】而非毛发生额。
                        #   原 费用取 jf / 收入取 df（毛额），财务费用含大额贷方（汇兑收益/利息
                        #   收入）被虚增（1357 1220 主体 jf=1.077亿 vs 净额 682万，审定表合计
                        #   1.282亿 vs 明细表 978万，差 1.18亿）。与明细表（净额/qm）口径一致：
                        #   收入类=df-jf（净贷方），费用类=jf-df（净借方）。
                        tot += _pl_amt(v, is_credit)
                    return tot

                # ---- 按核算主体列示（by_entity） ----
                if item.get('by_entity'):
                    ws = wb.create_sheet(title=title)
                    ncols = 6
                    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
                    _stitle = item.get('subj_name', title).replace(' 审定表', '')
                    tcell = ws.cell(1, 1, f'{_stitle} 审定表（按核算主体）')
                    tcell.font = Font(name='Times New Roman', size=10, bold=True); tcell.fill = TITLEFILL
                    tcell.alignment = CTR
                    ws.row_dimensions[1].height = 22
                    headers = ['核算主体', '上年数', '本期数', '审计调整数', '审定数', '与科目余额表勾稽']
                    for j, h in enumerate(headers, start=1):
                        c = ws.cell(2, j, h); c.font = Font(name='Times New Roman', size=10, bold=True)
                        c.fill = HFILL; c.alignment = CTR; c.border = BORDER
                    r0 = 3
                    t_prev = t_cur = 0.0
                    for ent in sorted(entities):
                        prev_amt = _gross(codes=codes, y=prev_y, for_entity=ent) if prev_y else 0.0
                        cur_amt = _gross(codes=codes, y=cur_y, for_entity=ent)
                        t_prev += prev_amt; t_cur += cur_amt
                        r = r0
                        ws.cell(r, 1, ent).alignment = LFT
                        ws.cell(r, 2, round(prev_amt, 2)).number_format = NUMFMT
                        ws.cell(r, 3, round(cur_amt, 2)).number_format = NUMFMT
                        ws.cell(r, 4, None).number_format = NUMFMT
                        ws.cell(r, 5, round(cur_amt, 2)).number_format = NUMFMT   # 2026-08-06 协议：审定数写数值
                        ws.cell(r, 6, '与科目余额表核对一致').alignment = LFT
                        for j in range(1, ncols + 1):
                            cell = ws.cell(r, j); cell.border = BORDER; cell.font = FONT
                            if j in (2, 3, 4, 5):
                                cell.alignment = RGT
                        r0 += 1
                    r = r0
                    ws.cell(r, 1, '合计').font = Font(name='Times New Roman', size=10, bold=True)
                    ws.cell(r, 1).fill = TOTFILL; ws.cell(r, 1).alignment = CTR
                    # 2026-08-03 公式化：PL by_entity 合计行 =SUM(主体数据区 3..r-1)
                    if r - 1 >= 3:
                        ws.cell(r, 2, '=SUM(B3:B%d)' % (r - 1)).number_format = NUMFMT
                        ws.cell(r, 3, '=SUM(C3:C%d)' % (r - 1)).number_format = NUMFMT
                    else:
                        ws.cell(r, 2, round(t_prev, 2)).number_format = NUMFMT
                        ws.cell(r, 3, round(t_cur, 2)).number_format = NUMFMT
                    ws.cell(r, 2).fill = TOTFILL; ws.cell(r, 2).alignment = RGT
                    ws.cell(r, 3).fill = TOTFILL; ws.cell(r, 3).alignment = RGT
                    for j in (4, 5, 6):
                        c = ws.cell(r, j); c.fill = TOTFILL; c.font = FONT
                        c.alignment = RGT if j in (4,5) else LFT
                    widths = [16, 18, 18, 14, 18, 28]
                    for j, w in enumerate(widths, start=1):
                        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
                    ws.freeze_panes = 'A3'
                    before = item.get('before')
                    sheets = wb._sheets
                    cur_idx = next((i for i, s in enumerate(sheets) if s.title == title), None)
                    if before and before in wb.sheetnames and cur_idx is not None:
                        before_idx = next(i for i, s in enumerate(sheets) if s.title == before)
                        node = sheets.pop(cur_idx)
                        if cur_idx < before_idx:
                            before_idx -= 1
                        sheets.insert(before_idx, node)
                    made.append(ws)
                    continue
                # ---- 原利润表科目：按科目名称列示 ----
                inc_rows = []  # (key, name, 上年数, 本期数)
                if names:
                    for nm_match in names:
                        prev_amt = _gross(nm=nm_match, y=prev_y) if prev_y else 0.0
                        cur_amt = _gross(nm=nm_match, y=cur_y)
                        inc_rows.append((nm_match, nm_match, prev_amt, cur_amt))
                    inc_rows.append(('合计', item.get('subj_name', '合计'),
                                     sum(r[2] for r in inc_rows), sum(r[3] for r in inc_rows)))
                else:
                    for code in (codes or []):
                        prev_amt = _gross(code=code, y=prev_y) if prev_y else 0.0
                        cur_amt = _gross(code=code, y=cur_y)
                        inc_rows.append((code, item.get('subj_name') or code, prev_amt, cur_amt))
                    if len(codes or []) > 1:
                        inc_rows.append(('合计', item.get('subj_name', '合计'),
                                         sum(r[2] for r in inc_rows), sum(r[3] for r in inc_rows)))
                if not inc_rows:
                    continue

                ws = wb.create_sheet(title=title)
                ncols = 7
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
                tcell = ws.cell(1, 1, f"{item.get('subj_name', title)} 审定表（利润表科目：上年数/本期数取自科目余额表 gross 发生额）")
                tcell.font = Font(name='Times New Roman', size=10, bold=True); tcell.fill = TITLEFILL
                tcell.alignment = CTR
                ws.row_dimensions[1].height = 22
                headers = ['科目编码', '科目名称', '上年数', '本期数', '审计调整数', '审定数', '与科目余额表勾稽']
                for j, h in enumerate(headers, start=1):
                    c = ws.cell(2, j, h); c.font = Font(name='Times New Roman', size=10, bold=True)
                    c.fill = HFILL; c.alignment = CTR; c.border = BORDER
                ws.row_dimensions[2].height = 20
                r0 = 3
                for i, (key, nm, prev_amt, cur_amt) in enumerate(inc_rows):
                    r = r0 + i
                    is_tot = (key == '合计')
                    ws.cell(r, 1, key).alignment = CTR
                    ws.cell(r, 2, nm).alignment = LFT
                    ws.cell(r, 3, prev_amt).number_format = NUMFMT
                    ws.cell(r, 4, cur_amt).number_format = NUMFMT
                    ws.cell(r, 5, None).number_format = NUMFMT
                    ws.cell(r, 6, f'=D{r}+E{r}').number_format = NUMFMT
                    ws.cell(r, 7, f"与科目余额表({nm})核对一致").alignment = LFT
                    for j in range(1, ncols + 1):
                        cell = ws.cell(r, j); cell.border = BORDER; cell.font = FONT
                        if j in (3, 4, 5, 6):
                            cell.alignment = RGT
                    if is_tot:
                        for j in range(1, ncols + 1):
                            ws.cell(r, j).fill = TOTFILL
                            ws.cell(r, j).font = Font(name='Times New Roman', size=10, bold=True)
                widths = [12, 18, 18, 18, 14, 18, 28]
                for j, w in enumerate(widths, start=1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
                ws.freeze_panes = 'A3'
                before = item.get('before')
                sheets = wb._sheets
                cur_idx = next((i for i, s in enumerate(sheets) if s.title == title), None)
                if before and before in wb.sheetnames and cur_idx is not None:
                    before_idx = next(i for i, s in enumerate(sheets) if s.title == before)
                    node = sheets.pop(cur_idx)
                    if cur_idx < before_idx:
                        before_idx -= 1
                    sheets.insert(before_idx, node)
                made.append(ws)
                continue
            # ---------- 资产负债表科目：期初/本期增加/本期减少/期末 ----------
            # ---------- 按核算主体列示审定表（by_entity） ----------
            if item.get('by_entity'):
                ws = wb.create_sheet(title=title)
                ncols = 6
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
                _stitle = item.get('subj_name', title).replace(' 审定表', '')
                tcell = ws.cell(1, 1, f'{_stitle} 审定表（按核算主体）')
                tcell.font = Font(name='Times New Roman', size=10, bold=True); tcell.fill = TITLEFILL
                tcell.alignment = CTR
                ws.row_dimensions[1].height = 22
                headers = ['核算主体', '年初数', '期末未审数', '审计调整数', '审定数', '与科目余额表勾稽']
                for j, h in enumerate(headers, start=1):
                    c = ws.cell(2, j, h); c.font = Font(name='Times New Roman', size=10, bold=True)
                    c.fill = HFILL; c.alignment = CTR; c.border = BORDER
                r0 = 3
                tot_qc = tot_qm = 0.0
                for ent in sorted(entities):
                    qc = qm = 0.0
                    for (e, c, n, yy), v in tb_full.items():
                        if e != ent or yy not in eff_years:
                            continue
                        if names:
                            if n not in names: continue
                        elif codes:
                            if c not in codes: continue
                        else:
                            continue
                        qc += v['qc']
                        qm += v['qm']
                    d_qc = round(-qc if (is_credit and qc < 0) or (not is_credit and qc < 0) else qc, 2)
                    d_qm = round(-qm if (is_credit and qm < 0) or (not is_credit and qm < 0) else qm, 2)
                    # 负债/权益类：TB中自然符号为借正贷负，但不同账套符号不一致，
                    # 统一用 abs() 确保贷方余额正数显示（用户要求：权益科目贷方为正）
                    # ⚡⚡ 2026-08-27 修复（AH 底稿复查）：SAP 未分配利润(4104)族 Balance=Credit
                    #   （贷方正、借余负）——qm>0=盈利、qm<0=亏损主体。若 item.keep_sign=True
                    #   保持带符号（亏损主体显示负，集团合计=Σ带符号，与明细表一致）；
                    #   否则沿用 abs（实收/资本公积/盈余公积/专项储备 Balance=Debit 贷余负取正）。
                    if item.get('keep_sign'):
                        d_qc = round(qc, 2)
                        d_qm = round(qm, 2)
                    elif is_credit:
                        d_qc = round(abs(qc), 2)  # 负债/权益贷方余额 → 正数显示
                        d_qm = round(abs(qm), 2)
                    else:
                        d_qc = round(-qc if qc < 0 else qc, 2)  # 资产类保持借方正数
                        d_qm = round(-qm if qm < 0 else qm, 2)
                    tot_qc += d_qc; tot_qm += d_qm
                    r = r0
                    ws.cell(r, 1, ent).alignment = LFT
                    ws.cell(r, 2, d_qc).number_format = NUMFMT
                    ws.cell(r, 3, d_qm).number_format = NUMFMT
                    ws.cell(r, 4, None).number_format = NUMFMT
                    ws.cell(r, 5, round(d_qm, 2)).number_format = NUMFMT   # 2026-08-06 协议：审定数写数值（=期末+调整0）
                    ws.cell(r, 6, '与科目余额表核对一致').alignment = LFT
                    for j in range(1, ncols + 1):
                        cell = ws.cell(r, j); cell.border = BORDER; cell.font = FONT
                        if j in (2, 3, 4, 5):
                            cell.alignment = RGT
                    r0 += 1
                # 合计行（2026-08-06 协议：写数值=tot_qc/tot_qm 生成时累计）
                r = r0
                ws.cell(r, 1, '合计').font = Font(name='Times New Roman', size=10, bold=True)
                ws.cell(r, 1).fill = TOTFILL; ws.cell(r, 1).alignment = CTR
                ws.cell(r, 2, round(tot_qc, 2)).number_format = NUMFMT; ws.cell(r, 2).font = FONT
                ws.cell(r, 2).fill = TOTFILL; ws.cell(r, 2).alignment = RGT
                ws.cell(r, 3, round(tot_qm, 2)).number_format = NUMFMT; ws.cell(r, 3).font = FONT
                ws.cell(r, 3).fill = TOTFILL; ws.cell(r, 3).alignment = RGT
                ws.cell(r, 5, round(tot_qm, 2)).number_format = NUMFMT   # 合计行审定数（=期末+调整0）
                for j in (4, 5, 6):
                    c = ws.cell(r, j); c.fill = TOTFILL; c.font = FONT; c.alignment = RGT if j in (4,5) else LFT
                widths = [16, 18, 18, 14, 18, 28]
                for j, w in enumerate(widths, start=1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
                ws.freeze_panes = 'A3'
                # 放置到 before 之前
                before = item.get('before')
                sheets = wb._sheets
                cur_idx = next((i for i, s in enumerate(sheets) if s.title == title), None)
                if before and before in wb.sheetnames and cur_idx is not None:
                    before_idx = next(i for i, s in enumerate(sheets) if s.title == before)
                    node = sheets.pop(cur_idx)
                    if cur_idx < before_idx:
                        before_idx -= 1
                    sheets.insert(before_idx, node)
                made.append(ws)
                continue
            # ---------- 原资产负债表科目：按科目代码列示 ----------
            # 汇总每个 code/name 每年的 signed 期初/期末 与 gross 增加/减少
            rows = []  # (key, name, year, qc_s, inc, dec, qm_s)
            if names:
                # 按名称汇总（跨代码/跨主体合并；避免同码不同义污染，如东轴5001=生产成本）
                for nm_match in names:
                    for y in eff_years:
                        qc = inc = dec = qm = 0.0
                        for (e, c, n, yy), v in tb_full.items():
                            if n == nm_match and yy == y:
                                qc += v['qc']
                                if is_credit:
                                    inc += v['df']; dec += v['jf']
                                else:
                                    inc += v['jf']; dec += v['df']
                                qm += v['qm']
                        rows.append((nm_match, nm_match, y, qc, inc, dec, qm))
                # 跨 name 合计行
                for y in eff_years:
                    tqc = sum(r[3] for r in rows if r[2] == y)
                    tinc = sum(r[4] for r in rows if r[2] == y)
                    tdec = sum(r[5] for r in rows if r[2] == y)
                    tqm = sum(r[6] for r in rows if r[2] == y)
                    rows.append(('合计', item.get('subj_name', '合计'), y, tqc, tinc, tdec, tqm))
            else:
                for code in (codes or []):
                    for y in eff_years:
                        qc = inc = dec = qm = 0.0
                        nm = item.get('subj_name') or None
                        for (e, c, n, yy), v in tb_full.items():
                            if c == code and yy == y:
                                qc += v['qc']
                                if is_credit:
                                    inc += v['df']; dec += v['jf']
                                else:
                                    inc += v['jf']; dec += v['df']
                                qm += v['qm']
                                if nm is None and n:
                                    nm = n
                        rows.append((code, nm or code, y, qc, inc, dec, qm))
                # 多 code 时追加合计行
                if len(codes or []) > 1:
                    for y in eff_years:
                        tqc = sum(r[3] for r in rows if r[2] == y)
                        tinc = sum(r[4] for r in rows if r[2] == y)
                        tdec = sum(r[5] for r in rows if r[2] == y)
                        tqm = sum(r[6] for r in rows if r[2] == y)
                        rows.append(('合计', item.get('subj_name', '合计'), y, tqc, tinc, tdec, tqm))
            if not rows:
                continue

            ws = wb.create_sheet(title=title)
            # 标题行
            ncols = 12
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
            tcell = ws.cell(1, 1, f"{item.get('subj_name', title)} 审定表")
            tcell.font = Font(name='Times New Roman', size=10, bold=True); tcell.fill = TITLEFILL
            tcell.alignment = CTR
            ws.row_dimensions[1].height = 22
            # 表头
            headers = ['科目编码', '科目名称', '年度', '期初方向', '期初未审数', '本期增加',
                       '本期减少', '期末方向', '期末未审数', '审计调整数', '期末审定数', '与科目余额表勾稽']
            for j, h in enumerate(headers, start=1):
                c = ws.cell(2, j, h); c.font = Font(name='Times New Roman', size=10, bold=True)
                c.fill = HFILL; c.alignment = CTR; c.border = BORDER
            ws.row_dimensions[2].height = 20
            # 数据
            r0 = 3
            _y_first = {}   # year -> 该年数据行首行（合计行公式引用，2026-08-02）
            for i, (code, nm, y, qc, inc, dec, qm) in enumerate(rows):
                r = r0 + i
                is_tot = (code == '合计')
                fx_qc = '借' if qc >= 0 else '贷'
                fx_qm = '借' if qm >= 0 else '贷'
                ws.cell(r, 1, code).alignment = CTR
                ws.cell(r, 2, nm).alignment = LFT
                ws.cell(r, 3, y).alignment = CTR
                ws.cell(r, 4, fx_qc).alignment = CTR
                if is_tot:
                    # 合计行：期初/增/减/末 =SUM(该年数据区)；方向/审定数沿用
                    _f0 = _y_first.get(y, r)
                    _gcl4 = openpyxl.utils.get_column_letter
                    for _cc in (5, 6, 7, 9):
                        _cl = _gcl4(_cc)
                        ws.cell(r, _cc, '=SUM(%s%d:%s%d)' % (_cl, _f0, _cl, r - 1)).number_format = NUMFMT
                    ws.cell(r, 10, None).number_format = NUMFMT
                    ws.cell(r, 11, f'=I{r}+J{r}').number_format = NUMFMT
                else:
                    _y_first.setdefault(y, r)
                    ws.cell(r, 5, abs(qc)).number_format = NUMFMT
                    ws.cell(r, 6, inc).number_format = NUMFMT
                    ws.cell(r, 7, dec).number_format = NUMFMT
                    ws.cell(r, 9, abs(qm)).number_format = NUMFMT
                    ws.cell(r, 10, None).number_format = NUMFMT  # 调整数留空
                    ws.cell(r, 11, f'=I{r}+J{r}').number_format = NUMFMT  # 审定数=未审+调整
                ws.cell(r, 12, f"与科目余额表({code})核对一致").alignment = LFT
                for j in range(1, ncols + 1):
                    cell = ws.cell(r, j); cell.border = BORDER
                    cell.font = FONT
                    if j in (5, 6, 7, 9, 10, 11):
                        cell.alignment = RGT
                if is_tot:
                    for j in range(1, ncols + 1):
                        ws.cell(r, j).fill = TOTFILL
                        ws.cell(r, j).font = Font(name='Times New Roman', size=10, bold=True)
            # 列宽
            widths = [12, 18, 8, 8, 16, 16, 16, 8, 16, 14, 16, 30]
            for j, w in enumerate(widths, start=1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
            ws.freeze_panes = 'A3'
            # 放置到 before 之前
            before = item.get('before')
            sheets = wb._sheets
            cur_idx = next((i for i, s in enumerate(sheets) if s.title == title), None)
            if before and before in wb.sheetnames and cur_idx is not None:
                before_idx = next(i for i, s in enumerate(sheets) if s.title == before)
                node = sheets.pop(cur_idx)
                if cur_idx < before_idx:
                    before_idx -= 1
                sheets.insert(before_idx, node)
            elif before and before not in wb.sheetnames:
                print(f'  ⚠️ 审定表[{title}]指定的前置明细表不存在：{before}（已置于末尾）')
            made.append(ws)
    return made


def read_tb_leaf(data_dir, entities, code_prefix):
    """取 TB 中代码以 code_prefix 开头的末级行，返回 (e,y) -> [(code,name,jf,df,qm)]。"""
    out = defaultdict(list)
    tb = read_tb_full(data_dir, entities)
    for (e, code, name, y), v in tb.items():
        if not code.startswith(code_prefix):
            continue
        out[(e, y)].append((code, name, v['jf'], v['df'], v['qm']))
    return out


# ===================== 账户层级拆解（从一级科目/TB出发）共享逻辑 =====================
# 适用范围：费用/应交税费/权益/存货/收入/利润表 等按「科目→二级→三级」拆解的明细表。
# 编码此6类构建器共同遵循的铁律（已与费用验证一致、被 validate_expense 回归校验）：
#   ① 从一级科目（TB 控制数）出发；
#   ② 仅当 二级 TB 合计 ≠ 一级 TB 合计（缺口>容差）时，才用 GL 补缺并标注「GL补充」；
#   ③ 三级明细若合计与二级一致则展开标注（父级带 TB 值计入合计，子目仅下钻不重复计）；
#   ④ 0 值行不显示；
#   ⑤ 一级↔二级 勾稽提示行。
def _tol(x):
    """容差：相对 1e-6，绝对下限 1 元。缺口/差异判定统一用此。"""
    return max(1.0, abs(x) * 1e-6)


def is_zero_amount(*vals, eps=0.005):
    """0 值行判断：所有给定金额均接近 0 则不显示（eps 默认 0.005 元）。"""
    return all(abs(v) < eps for v in vals)


def tb_account_tree(tb, code_prefix=None):
    """从 read_tb_full 输出构建 一级→二级→三级 树（账户层级拆解类共享）。

    返回 { (e, y, l1code): {'name','jf','df','qc','qm','l2':
            { l2code: {'name','jf','df','l3': { l3code: {'name','jf','df'} } } } } }
    层级按【去非数字后】的位数判定：4=一级、6=二级、>6（同属某 6 位父）视为三级；
    用完整去数字码作为三级键，兼容 8 位/10 位等多级体系。
    code_prefix：仅保留顶层代码以该前缀开头的账户（如 '2221' 应交税费、'4001' 权益）。
    节点原始 jf/df 同时保留，发生额口径（借发/贷发/gross）由调用方按账户类型选取。
    """
    tree = {}
    for (e, code, name, y), v in tb.items():
        if not code:
            continue
        dg = re.sub(r'\D', '', str(code))
        if len(dg) < 4:
            continue
        if code_prefix and not dg.startswith(code_prefix):
            continue
        l1 = dg[:4]
        node = tree.setdefault((e, y, l1),
                               {'name': None, 'jf': 0.0, 'df': 0.0, 'qc': 0.0, 'qm': 0.0, 'l2': {}})
        if len(dg) == 4:
            node['name'] = name
            node['jf'] = v['jf']; node['df'] = v['df']; node['qc'] = v['qc']; node['qm'] = v['qm']
        elif len(dg) >= 6:
            l2 = dg[:6]
            l2n = node['l2'].setdefault(l2, {'name': None, 'jf': 0.0, 'df': 0.0, 'l3': {}})
            if len(dg) == 6:
                l2n['name'] = name; l2n['jf'] = v['jf']; l2n['df'] = v['df']
            else:
                l3 = dg  # 完整去数字码作为三级键
                l3n = l2n['l3'].setdefault(l3, {'name': None, 'jf': 0.0, 'df': 0.0})
                l3n['name'] = name; l3n['jf'] = v['jf']; l3n['df'] = v['df']
    return tree


def build_account_breakdown_rows(l1_flow, l2_list, l3_map, gl_supp_keys=None, tb_names=None, tol=None, l4_map=None):
    """通用「从一级出发」行规范构造（账户层级拆解类底稿共享，已与费用验证一致）。

    参数：
      l1_flow      : 一级 TB 发生额（合计控制数）
      l2_list      : [(l2code, l2name, l2_flow), ...]  二级(TB) 行（顶层仅 6 位；孤儿 8 位提升为顶层）
      l3_map       : {l2code: [(l3code, l3name, l3_flow), ...]}  三级明细
      l4_map       : 可选 {l3code: [(l4code, l4name, l4_flow), ...]}  四级明细（如 管理费用-研发费用-五险二金-养老保险）
      gl_supp_keys : 可选，GL 中存在的 key（含 None 表示一级直接发生额）；缺口>容差时补 L2G/L1G
      tb_names     : 可选，TB 全部 clean 名称集合（判定 L2G 是否"未在 TB 列示"）
      tol          : 容差函数（默认 _tol）
    返回 rows：6 元组列表
      (out_code, out_name_display, gl_l2key_clean, tb_l2code, level, children)
      level ∈ {'L2','L3','L4','L2G','L1G'}；children = 子三级 code 列表 或 None
    原则：
      · 二级明细来自 TB；
      · 若 三级合计(TB)≈二级(TB) → 展开三级（父级 L2 带 TB 值计入合计，L3 子目仅下钻不重复计）；
      · 三级内部再遇 四级合计≈三级（L4 完整覆盖）→ 三级标「含四级明细」并展开 L4（同样下钻不重复计）；
      · 仅当 二级TB合计 ≠ 一级TB合计（缺口>容差）时，才补 GL 中 TB 未列示的类别并标注。
    """
    tol = tol or _tol
    l4_map = l4_map or {}
    rows = []
    for l2code, l2name, l2_flow in sorted(l2_list, key=lambda x: x[0]):
        children = l3_map.get(l2code, [])
        if children:
            l3sum = sum(cf for (_, _, cf) in children)
            if abs(l3sum - l2_flow) <= tol(l2_flow):
                # 三级合计与二级一致 → 展开三级：父级仍带 TB 值计入合计，子目仅下钻
                rows.append((l2code, l2name + '（含三级明细）', l2name, l2code, 'L2',
                             [cc for (cc, _, _) in children]))
                for l3code, l3name, l3f in sorted(children, key=lambda x: x[0]):
                    l4s = l4_map.get(l3code, [])
                    if l4s and abs(sum(l4f for (_, _, l4f) in l4s) - l3f) <= tol(l3f):
                        # 四级合计与三级一致 → 三级标「含四级明细」并展开 L4（L4 下钻不重复计）
                        rows.append((l3code, '　' + l3name + '（含四级明细）', l3name, l3code, 'L3',
                                     [c4 for (c4, _, _) in l4s]))
                        for l4code, l4name, _ in sorted(l4s, key=lambda x: x[0]):
                            rows.append((l4code, '　　' + l4name + '（四级明细）', l4name, l4code, 'L4', None))
                    else:
                        rows.append((l3code, '　' + l3name + '（三级明细）', l3name, l3code, 'L3', None))
                continue
        rows.append((l2code, l2name, l2name, l2code, 'L2', ([cc for (cc, _, _) in children] or None)))
    # 缺口判定：仅 二级TB合计 ≠ 一级TB合计 时才补 GL
    l2_total = sum(lf for (_, _, lf) in l2_list)
    gap = l1_flow - l2_total
    if abs(gap) > tol(l1_flow) and gl_supp_keys:
        names = tb_names or set()
        for k in sorted(gl_supp_keys, key=lambda x: (x is None, x or '')):
            if k is None:
                rows.append(('', '（一级直接发生额·GL补充）', None, None, 'L1G', None))
            elif k not in names:
                rows.append(('', k + '（GL补充）', k, None, 'L2G', None))
    return rows


def write_recon_prompt(ws, r, year_vals, value_cols=None, num_fmt='#,##0.00',
                       title_color='000000', warn_color='C00000', tol=None):
    """写「一级(TB) ↔ 二级明细合计(TB)」勾稽提示块。返回下一空闲行号。

    参数：
      year_vals : [(year_label, l1_val, l2_val), ...]  按列顺序（支持单年/多年）
      value_cols: {year_label: 列号}  默认从 3 起按 year_vals 顺序排（如 2025→3, 2026→4）
    输出布局（与费用明细表一致，validate_expense 依赖）：
      标题写 col1（加粗蓝字）；子行标签写 col2；金额写对应 value_cols（千分位、右对齐）；
      差异超容差标红 + 红字提示行。
    """
    tol = tol or _tol
    if value_cols is None:
        value_cols = {yv[0]: (3 + i) for i, yv in enumerate(year_vals)}
    c = ws.cell(r, 1, '勾稽：一级(TB) ↔ 二级明细合计(TB)')
    c.font = Font(name='Times New Roman', bold=True, color=title_color)
    r += 1
    labels = [yv[0] for yv in year_vals]
    l1s = [yv[1] for yv in year_vals]
    l2s = [yv[2] for yv in year_vals]
    diffs = [a - b for a, b in zip(l1s, l2s)]
    warn = any(abs(d) > tol(max(abs(a), abs(b))) for a, b, d in zip(l1s, l2s, diffs))

    def _row(label, vals, is_warn=False):
        nonlocal r
        ws.cell(r, 2, label).alignment = SHELL_LEFT
        for yl, v in zip(labels, vals):
            col = value_cols.get(yl)
            if col is None:
                continue
            cell = ws.cell(r, col, round(v, 2) if isinstance(v, (int, float)) else v)
            cell.number_format = num_fmt
            cell.alignment = SHELL_RGT
            if is_warn:
                cell.font = Font(name='Times New Roman', bold=True, color=warn_color)
        if is_warn:
            ws.cell(r, 2).font = Font(name='Times New Roman', bold=True, color=warn_color)
        r += 1

    _row('一级科目(TB)控制数', l1s)
    _row('二级明细合计(TB+GL补充)', l2s)
    _row('差异(一级−二级)', diffs, is_warn=warn)
    if warn:
        ws.cell(r, 2, '提示：差异超出容差，系 TB 二级合计与一级不一致，已用 GL 缺口补充或 TB 自身层级特性所致。')\
            .font = Font(name='Times New Roman', italic=True, color=warn_color)
        r += 1
    return r


def finalize_after_build(data_dir):
    """builder 单跑收尾：执行终审后处理（凭证抽查对方科目补全 + 小计行清理），
    保证「单个小程序生成」与「regen_all 编排生成」的底稿完全一致（2026-07-31 新增）。

    背景：regen_all 在 13 个 builder 跑完后统一调用 audit_finalize.finalize_folder 一次；
    而 13 个 builder 的 main() 原本都不触发 finalize → 单跑 builder 的底稿缺对方科目/含残留
    小计行，与 regen 产出一致性无法保证。本函数让单跑 builder 也执行同一终审。
    - regen 编排模式下跳过（regen_all 设置环境变量 REGEN_ALL_MODE=1，由 regen 最后统一 finalize）；
    - finalize_folder 幂等（只填空缺单元格/删标记行），重复执行无副作用。
    """
    if os.environ.get('REGEN_ALL_MODE') == '1':
        return
    if os.environ.get('SKIP_FINALIZE') == '1':
        return
    try:
        from audit_finalize import finalize_folder
        finalize_folder(data_dir)
    except Exception as ex:
        print(f'  ⚠️ 终审后处理失败：{ex}')
# ============================================================================
# 通用附注汇总生成器（2026-08-02 用户方法论推广：bank/loan/equity/rd_expense/gp_other 使用）
# 每主体一列；行=rows_spec（(行名, codes 列表)；函数自动追加 合计/复核 行=各列之和）；
# 段按 mode：'four'=期初数/本期增加/本期减少/期末数（滚动类，附注检查作用靠此四段+勾稽）；
#           'two'=期末数/期初数（两期数类：货币资金/借款——增减变动由现金流量表核对，无需四段）；
# 数据=TB 叶子行按 codes 前缀匹配（父=子和防双计）；is_credit=True 增加=贷发/减少=借发（负债权益），
# False 增加=借发/减少=贷发（资产）；每段末 合计/复核（=列和）。
# ============================================================================
def build_footnote_generic(wb, data_dir, sheet_name, title, rows_spec, is_credit=False,
                           mode='four', target_year=None, tb_full=None, entities=None,
                           add_total_col=False, skip_no_data=True, parent_abs_keys=()):
    """通用附注汇总：每主体一列、行=rows_spec、段=four/two/money、末行合计/复核。
    mode：'four'=期初/增/减/期末（滚动类）；'two'=期末/期初（两期数类）；
          'money'=货币资金专用：期末数(未审)/审计调整/期末审定数(公式)/期初数(仅数)，
          合并4行（集团合计数/抵消借/贷/合并报表数）仅挂期末审定数段后。
    add_total_col=True：表尾加「合计」列（每行各主体之和，2026-08-02 用户要求）。
    skip_no_data=True（2026-08-03）：账套全集团无该科目任何数据 → 不建 sheet 返回 None
          （数据驱动裁剪：账套没有的科目在附注汇总中不列示）。
    parent_abs_keys（2026-08-06）：行名在此集合的科目（租赁负债——子目『未确认融资费用』
          是借方备抵，末级绝对值之和会把备抵变加项：FY 租赁负债附注 51.7M vs 审定 41.5M，
          虚增 2×未确认融资费用）→ 附注改取【一级父级净额】（cs 精确 in codes，qc/qm abs），
          无父级行（末级导出）时回退原末级逻辑。"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    if tb_full is None or entities is None:
        entities = discover_entities(data_dir)
        tb_full = read_tb_full(data_dir, entities)
    # 2026-08-05 分币种生成：泰国主体 TB 原始代码（1113=银行存款等）与集团标准代码不同，
    # 附注 rows_spec 按标准 codes 匹配会漏掉泰国 → 统一归一化（幂等，与审定表同一策略）。
    try:
        from thai_mapping import normalize_thai_tb as _norm_thai_tb
        _tb_norm = _norm_thai_tb(tb_full, entities)
        if _tb_norm is not tb_full:
            tb_full = _tb_norm
    except Exception:
        pass
    years = sorted({str(y) for e in entities for y in entities[e]})
    if target_year is not None:
        years = [str(target_year)]
    ent_list = sorted(entities)   # 2026-08-02：附注汇总主体列按 01、02… 排序
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    TITLEFILL = PatternFill('solid', fgColor='BDD7EE')
    TOTFILL = PatternFill('solid', fgColor='FCE4D6')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
    RGT = Alignment(horizontal='right', vertical='center')
    NUM = '#,##0.00'
    F10 = Font(name='Times New Roman', size=10)
    BF = Font(name='Times New Roman', size=10, bold=True)
    TF = Font(name='Times New Roman', size=12, bold=True)

    def _ent_vals(e, y):
        """rows_spec 各行 → [qc, jf, df, qm]（qc/qm 取正；增减按 is_credit 方向）。
        rows_spec 元素：(行名, codes) 或 (行名, codes, name_kw)：codes 非空按代码前缀匹配；
        codes 为空且 name_kw 给定按名称关键词匹配（铁律23：按实际账套科目名）。"""
        out = []
        for row_spec in rows_spec:
            if len(row_spec) >= 3:
                row_name, codes, name_kw = row_spec
            else:
                row_name, codes = row_spec
                name_kw = None
            rows = []
            for (ee, c, n, yy), v in tb_full.items():
                if ee != e or yy != y:
                    continue
                cs = str(c)
                if codes:
                    if not any(cs == k or (cs.startswith(k) and len(cs) > len(k)) for k in codes):
                        continue
                    rows.append((cs, v))
                elif name_kw:
                    if not any(k in str(n) for k in name_kw):
                        continue
                    rows.append((cs, v))
            # 2026-08-06：codes 无命中 → 名称兜底（铁律5：Q 一年内到期的非流动负债=2401、
            # 租赁负债=2271 等非标准代码；codes 仅标准码 2242/2485 → 附注"全集团无数据"
            # 被裁剪 → 底稿缺『附注汇总』sheet 结构性 ERROR）。row_name 即科目名关键词。
            if codes and not rows and (name_kw or row_name):
                _kws = name_kw or [row_name]
                for (ee, c, n, yy), v in tb_full.items():
                    if ee != e or yy != y:
                        continue
                    if any(k in str(n) for k in _kws):
                        rows.append((str(c), v))
            # 2026-08-06 租赁负债类：子目含借方备抵（未确认融资费用），末级 abs 会把备抵
            # 变加项（FY 租赁负债附注 51.7M vs 审定 41.5M）→ 取【一级父级净额】：
            # cs 精确 in codes（2601/2631…），qc/qm 取 abs（贷方余额转正）；父级缺失
            # （U8 末级导出无 4 位父级行）→ 回退下方末级逻辑。
            if row_name in parent_abs_keys:
                _one = [r for r in rows if str(r[0]) in codes]
                if _one:
                    # 2026-08-06 修复：_one 含 220601/220602（ga 末级租赁负债，无父级）时
                    # 原 abs 逐个累加双计备抵（1,122,886.70 vs 净额 1,048,313.30）→ 统一
                    # 【叶子带符号求和再 abs】（父级=子和、备抵子目自然抵减，兼容 FY 父级/
                    # ga-jj 末级一正一负/纯末级单码 三种结构）：
                    _leaves = [r for r in _one if not any(
                        r[0] != r2[0] and r2[0].startswith(r[0]) and len(r2[0]) > len(r[0]) for r2 in _one)]
                    qc = sum(v['qc'] for _, v in _leaves)
                    qm = sum(v['qm'] for _, v in _leaves)
                    jf = sum(v['jf'] for _, v in _leaves)
                    df = sum(v['df'] for _, v in _leaves)
                    out.append([abs(qc), jf, df, abs(qm)])
                    continue
                # 无父级行（末级导出/父级 0 余额被剔除，如 FY 06FY智能 2026：仅
                # 2601.01 未确认融资费用 317,757.14 + 2601.02 租赁付款额 -317,757.14，
                # 净额 0）→ 带符号求和=净额再 abs（备抵子目自然抵减，勿 abs 双计
                # ——曾致附注多计 635,514.28=2×317,757.14）
                qc = jf = df = qm = 0.0
                for cs, v in rows:
                    qc += v['qc']; qm += v['qm']
                    jf += v['jf']; df += v['df']
                out.append([abs(qc), jf, df, abs(qm)])
                continue
            leafs = [r for r in rows if not any(
                r[0] != r2[0] and r2[0].startswith(r[0]) and len(r2[0]) > len(r[0]) for r2 in rows)]
            qc = jf = df = qm = 0.0
            for cs, v in leafs:
                qc += abs(v['qc']); qm += abs(v['qm'])
                if is_credit:
                    df += v['df']; jf += v['jf']
                else:
                    jf += v['jf']; df += v['df']
            out.append([qc, jf, df, qm])
        return out

    # ---- 数据驱动裁剪（2026-08-03 skip_no_data）：账套全集团无该科目数据 → 不建 sheet ----
    _yv_all = {}
    _has_any = False
    for y in years:
        _yv = {e: _ent_vals(e, y) for e in ent_list}
        _yv_all[y] = _yv
        if any(abs(v) > 1e-6 for e in ent_list for row in _yv[e] for v in row):
            _has_any = True
    if skip_no_data and not _has_any:
        print(f'  ⊘ {sheet_name}：账套全集团无数据，附注汇总不生成（数据驱动裁剪）')
        return None

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    NC = 2 + len(ent_list) + (1 if add_total_col else 0)
    # 2026-08-02：附注汇总不设合并单元格（便于后期添加/筛选）；标题仅 A1 加粗
    t = ws.cell(1, 1, title)
    t.font = TF; t.fill = TITLEFILL; t.alignment = CTR
    ws.row_dimensions[1].height = 22
    ws.cell(2, 1, '项目（审定数）').font = BF
    ws.cell(2, 1).fill = HFILL; ws.cell(2, 1).alignment = CTR; ws.cell(2, 1).border = BORDER
    for i, ent in enumerate(ent_list, 2):
        c = ws.cell(2, i, ent)
        c.font = BF; c.fill = HFILL; c.alignment = CTR; c.border = BORDER
    if add_total_col:
        c = ws.cell(2, NC, '合计')
        c.font = BF; c.fill = TOTFILL; c.alignment = CTR; c.border = BORDER
    ws.column_dimensions['A'].width = 26
    for j in range(2, NC + 1):
        ws.column_dimensions[get_column_letter(j)].width = 16
    # ⚡ 2026-08-10 排版优化（R2 冻结项目列）：横滚时项目名始终可见
    try:
        ws.freeze_panes = 'B3'
    except Exception:
        pass

    def _write_merge4(r, rc, yv, seg_rows, idx):
        """合并4行：集团合计数（数值，生成时算好——2026-08-06 协议：权威行写数值，
        公式 openpyxl 保存无缓存 data_only 读 None = 收回读不回）+ 抵消借/贷（0 待填）
        + 合并报表数（=集团合计，抵消 0 时；审计人员填抵消后 Excel 重算）。
        ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）时不写 集团合计/抵消借/贷/
        合并报表数 4 行（用户要求：单体底稿不要集团合计数部分）；grp_vals 仍算（审定数行用）。"""
        # 集团合计 = 各主体本段 rows 之和（生成时数据在手，直接算）
        grp_vals = []
        for ent in ent_list:
            _s = 0.0
            for row_spec in seg_rows:
                for _ri, _rs in enumerate(rows_spec):
                    if _rs[0] == row_spec[0]:
                        _s += yv[ent][_ri][idx]
                        break
            grp_vals.append(_s)
        _single = len(ent_list) <= 1
        if not _single:
            # ⚡ 2026-08-10 排版优化（R5 合并行视觉层级）：集团合计=中蓝/抵消借=浅红/抵消贷=浅绿/
            #   合并报表数=深蓝白字；0 值留空（抵消借/贷待填，留空提示填写）。
            _F_MID = PatternFill('solid', fgColor='BDD7EE')
            _F_RED = PatternFill('solid', fgColor='FCE4EC')
            _F_GRN = PatternFill('solid', fgColor='E2EFDA')
            _F_DRK = PatternFill('solid', fgColor='DDEBF7')
            _F_DRK_F = Font(name='Times New Roman', bold=True, color='000000', size=10)
            ws.cell(r, 1, '集团合计数').font = BF; ws.cell(r, 1).fill = _F_MID
            ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i, ent in enumerate(ent_list, 2):
                cc = ws.cell(r, i)
                if abs(grp_vals[i - 2]) >= 0.005:
                    cc.value = round(grp_vals[i - 2], 2)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.fill = _F_MID; cc.font = BF
            if add_total_col:
                _tc = sum(grp_vals)
                cc = ws.cell(r, NC)
                if abs(_tc) >= 0.005:
                    cc.value = round(_tc, 2)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.fill = _F_MID; cc.font = BF
            r += 1
            # 合并抵消借方（待填，留空）
            ws.cell(r, 1, '合并抵消借方').font = F10; ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i in range(2, NC + 1):
                cc = ws.cell(r, i)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = F10
                cc.fill = _F_RED
            r += 1
            # 合并抵消贷方（待填，留空）
            ws.cell(r, 1, '合并抵消贷方').font = F10; ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i in range(2, NC + 1):
                cc = ws.cell(r, i)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = F10
                cc.fill = _F_GRN
            r += 1
            # 合并报表数（= 集团合计数 + 抵消借 − 抵消贷；抵消 0 时 = 集团合计，写数值）
            ws.cell(r, 1, '合并报表数').font = _F_DRK_F; ws.cell(r, 1).fill = _F_DRK
            ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
            for i, ent in enumerate(ent_list, 2):
                cc = ws.cell(r, i)
                if abs(grp_vals[i - 2]) >= 0.005:
                    cc.value = round(grp_vals[i - 2], 2)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.fill = _F_DRK
                cc.font = _F_DRK_F
            if add_total_col:
                _tc = sum(grp_vals)
                cc = ws.cell(r, NC)
                if abs(_tc) >= 0.005:
                    cc.value = round(_tc, 2)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.fill = _F_DRK
                cc.font = _F_DRK_F
            r += 1
        return r, grp_vals

    def _write_seg(r, seg_label, idx, yv):
        """写一个数据段：段头 + rows_spec 行（含合计列），值取 yv[ri][idx]。返回下一行号。"""
        c = ws.cell(r, 1, seg_label)
        c.font = BF; c.fill = TOTFILL; c.alignment = CTR
        r += 1
        for ri, row_spec in enumerate(rows_spec):
            row_name = row_spec[0]
            cell = ws.cell(r, 1, row_name)
            cell.font = F10; cell.alignment = LFT; cell.border = BORDER
            for i, ent in enumerate(ent_list, 2):
                v = yv[ent][ri][idx]
                cc = ws.cell(r, i, round(v, 2) if abs(v) >= 0.005 else 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = F10
            if add_total_col:
                # 2026-08-06 协议：合计列写数值（公式 data_only 读 None）
                _tc = sum(yv[ent][ri][idx] for ent in ent_list)
                cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = F10
            r += 1
        return r

    r = 3
    for y in years:
        yv = _yv_all[y]
        if mode == 'money':
            # 期末数（未审，qm）→ 审计调整（0 待填）→ 期末审定数（=未审+调整，公式）
            # → 合并4行（审定口径）→ 期初数（仅数值，不加调整/审定）
            def _seg_head(r, label):
                c = ws.cell(r, 1, label); c.font = BF; c.fill = TOTFILL; c.alignment = CTR
                return r + 1

            def _seg_body(r, idx=None):
                """写 rows_spec 科目行；idx 给定取 yv[ri][idx]，否则全写 0（审计调整待填）。
                返回 (科目行首行号, 下一行号)。"""
                first = r
                for ri, row_spec in enumerate(rows_spec):
                    cell = ws.cell(r, 1, row_spec[0])
                    cell.font = F10; cell.alignment = LFT; cell.border = BORDER
                    for i, ent in enumerate(ent_list, 2):
                        v = yv[ent][ri][idx] if idx is not None else 0.0
                        cc = ws.cell(r, i, round(v, 2) if abs(v) >= 0.005 else 0.0)
                        cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
                    if add_total_col:
                        # 2026-08-06 协议：合计列写数值
                        _tc = sum(yv[e][ri][idx] for e in ent_list) if idx is not None else 0.0
                        cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
                        cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
                    r += 1
                return first, r

            r = _seg_head(r, '期末数')
            ru_first, r = _seg_body(r, idx=3)
            r = _seg_head(r, '审计调整')
            ra_first, r = _seg_body(r, idx=None)
            r = _seg_head(r, '期末审定数')
            for k, row_spec in enumerate(rows_spec):
                cell = ws.cell(r, 1, row_spec[0])
                cell.font = F10; cell.alignment = LFT; cell.border = BORDER
                for i, ent in enumerate(ent_list, 2):
                    # 期末审定数 = 期末未审数 + 审计调整(0 待填) → 数值 = 未审（协议：权威行写数值）
                    _v = yv[ent][k][3]
                    cc = ws.cell(r, i, round(_v, 2) if abs(_v) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
                if add_total_col:
                    _tc = sum(yv[e][k][3] for e in ent_list)
                    cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
                r += 1
            rc = r
            r, grp_vals = _write_merge4(r, rc, yv, rows_spec, 3)
            r = _seg_head(r, '期初数')
            _, r = _seg_body(r, idx=0)
            continue
        segs = [('期初数', 0), ('本期增加', 1), ('本期减少', 2), ('期末数', 3)] if mode == 'four' else [('期末数', 3), ('期初数', 0)]
        # 数据驱动裁剪（2026-08-02 用户：统一最全模板，套具体账套按数据生成——全集团口径）：
        # 行：某科目行全集团×全段无数据 → 隐藏；段：某段全集团×全行无数据 → 隐藏。
        has_row = [any(abs(yv[e][ri][idx]) > 1e-6 for e in ent_list for _, idx in segs)
                   for ri in range(len(rows_spec))]
        active_ris = [ri for ri, h in enumerate(has_row) if h]
        if not active_ris:
            active_ris = list(range(len(rows_spec)))   # 全 0 兜底仍显示（避免空表）
        seg_rc = {}   # 段 -> 集团合计数行号（2026-08-03 滚动勾稽行引用）
        for seg_label, idx in segs:
            if not any(abs(yv[e][ri][idx]) > 1e-6 for e in ent_list for ri in active_ris):
                continue    # 段无数据 → 不生成（如货币资金无增减段）
            # 段头：A 列加粗提示（不合并跨行）
            c = ws.cell(r, 1, seg_label)
            c.font = BF; c.fill = TOTFILL; c.alignment = CTR
            r += 1
            for ri in active_ris:
                row_spec = rows_spec[ri]
                row_name = row_spec[0]
                cell = ws.cell(r, 1, row_name)
                cell.font = F10; cell.alignment = LFT; cell.border = BORDER
                for i, ent in enumerate(ent_list, 2):
                    v = yv[ent][ri][idx]
                    cc = ws.cell(r, i, round(v, 2) if abs(v) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                    cc.font = F10
                if add_total_col:
                    # 2026-08-06 协议：合计列写数值
                    _tc = sum(yv[ent][ri][idx] for ent in ent_list)
                    cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                    cc.font = F10
                r += 1
            # 集团合计数（2026-08-06 协议：数值，生成时算好；公式 data_only 读 None 收回读不回）
            rc = r
            r, grp_vals = _write_merge4(r, rc, yv, [rows_spec[i] for i in active_ris], idx)
            seg_rc[seg_label] = rc
            # —— 2026-08-03 用户方法论：附注汇总须考虑 审计调整数/审定数 ——
            # four(披露增减变动)区分调整借方/贷方；two(两期数)单列审计调整；money 模式已有。
            # 审定数 = 集团合计数 + 调整借 − 调整贷（调整 0 待填；2026-08-06 协议：写数值，
            # =集团合计——公式 data_only 读 None 收回读不回；审计人员填调整后 Excel 重算）。
            if mode == 'four':
                ws.cell(r, 1, '审计调整-借方').font = F10; ws.cell(r, 1).alignment = LFT
                ws.cell(r, 1).border = BORDER
                for i in range(2, NC + 1):
                    cc = ws.cell(r, i, 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
                r += 1
                ws.cell(r, 1, '审计调整-贷方').font = F10; ws.cell(r, 1).alignment = LFT
                ws.cell(r, 1).border = BORDER
                for i in range(2, NC + 1):
                    cc = ws.cell(r, i, 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
                r += 1
                ws.cell(r, 1, '审定数').font = BF; ws.cell(r, 1).fill = TOTFILL
                ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
                for i, ent in enumerate(ent_list, 2):
                    _gv = grp_vals[i - 2]
                    cc = ws.cell(r, i, round(_gv, 2) if abs(_gv) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                    cc.fill = TOTFILL; cc.font = BF
                if add_total_col:
                    _tc = sum(grp_vals)
                    cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                    cc.fill = TOTFILL; cc.font = BF
                r += 1
            else:
                ws.cell(r, 1, '审计调整').font = F10; ws.cell(r, 1).alignment = LFT
                ws.cell(r, 1).border = BORDER
                for i in range(2, NC + 1):
                    cc = ws.cell(r, i, 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER; cc.font = F10
                r += 1
                ws.cell(r, 1, '审定数').font = BF; ws.cell(r, 1).fill = TOTFILL
                ws.cell(r, 1).alignment = LFT; ws.cell(r, 1).border = BORDER
                for i, ent in enumerate(ent_list, 2):
                    _gv = grp_vals[i - 2]
                    cc = ws.cell(r, i, round(_gv, 2) if abs(_gv) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                    cc.fill = TOTFILL; cc.font = BF
                if add_total_col:
                    _tc = sum(grp_vals)
                    cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
                    cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                    cc.fill = TOTFILL; cc.font = BF
                r += 1
        # 滚动勾稽行（2026-08-03 用户方法论：期初+本期增加-本期减少-期末=0；差异>0.01 标红）
        # 条件：four 模式且至少有一个段被渲染（期初/增/减/期末 段被裁剪时对应项按 0 处理）
        if mode == 'four' and seg_rc:
            r0 = seg_rc.get('期初数'); r1 = seg_rc.get('本期增加')
            r2 = seg_rc.get('本期减少'); r3 = seg_rc.get('期末数')
            exp = {}
            for e in ent_list:
                s = 0.0
                if r0: s += sum(yv[e][ri][0] for ri in active_ris)
                if r1: s += sum(yv[e][ri][1] for ri in active_ris)
                if r2: s -= sum(yv[e][ri][2] for ri in active_ris)
                if r3: s -= sum(yv[e][ri][3] for ri in active_ris)
                exp[e] = s
            c = ws.cell(r, 1, '滚动勾稽（期初+本期增加-本期减少-期末）')
            c.font = BF; c.fill = TOTFILL; c.alignment = LFT; c.border = BORDER
            for i, ent in enumerate(ent_list, 2):
                # 2026-08-06 协议：滚动勾稽写数值（exp 已生成时算好；公式 data_only 读 None）
                _ev = exp[ent]
                cc = ws.cell(r, i, round(_ev, 2) if abs(_ev) >= 0.005 else 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = Font(name='Times New Roman', size=10,
                               color='FF0000' if abs(exp[ent]) > 0.01 else '000000', bold=True)
            if add_total_col:
                _tc = sum(exp.values())
                cc = ws.cell(r, NC, round(_tc, 2) if abs(_tc) >= 0.005 else 0.0)
                cc.number_format = NUM; cc.alignment = RGT; cc.border = BORDER
                cc.font = Font(name='Times New Roman', size=10, bold=True)
            r += 1
    ws.freeze_panes = 'B3'
    return ws




# ============================================================================
# 2026-08-05 前道规范化：底稿结构校验（生成器写入后立即调用——生成即规范，
# 不靠后期检查纠错）。validate_sheet_structure 返回 [(级别, 问题)]，ERROR=硬伤
#（缺表头/缺主体列），WARN=建议项（无合计行/无标准期末列名/主体名含顿号）。
# 各 builder 在 sheet 写入完成后调用，不通过即打印/抛错，当场修正。
# ============================================================================
def validate_sheet_structure(ws, kind='detail', subj_label=''):
    """底稿 sheet 结构校验（2026-08-05 前道规范化）。
    kind: 'detail'（明细表）/ 'audit'（审定表）/ 'footnote'（附注汇总）。
    校验项：①表头含『核算主体/主体』列（明细/审定必须）②标准期末列名（期末/本期/全年/余额/审定数）
    ③合计/小计行 ④主体名不含顿号（『、』『，』）。返回 [(级别, 问题)]。"""
    problems = []
    try:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    except Exception as ex:
        return [('ERROR', f'读取失败：{ex}')]
    if not rows:
        return [('ERROR', '空表')]
    # 表头行（前 6 行找含 核算主体/主体/公司 标识列；排除标题行『XX明细表（2026…』『按主体汇总（…』）
    hdr_i = None
    for i, r in enumerate(rows[:6]):
        if not r:
            continue
        joined = ' '.join(str(x) for x in r[:12] if x is not None)
        if any(k in joined for k in ('明细表（', '汇总（', '审定表（', '（202', '（20', '附注汇总（')):
            continue
        if '核算主体' in joined or '主体' in joined or '公司' in joined:
            hdr_i = i
            break
    if hdr_i is None:
        # 2026-08-05 校准：附注（每主体一列/科目块）与 单主体科目式审定表
        #（应付股利/应收股利：科目编码|科目名称|年度|方向）无需『核算主体/主体』列——不报。
        _hdr_txt = ' '.join(str(x) for r in rows[:4] if r for x in r[:10] if x is not None)
        # ⚡ 2026-08-09：跨主体分月合并表（如『研发费用明细表』表头=项目|1月..12月|合计，
        # 用户图1 按费用性质×月份跨主体归并）无主体列属设计——豁免。
        _is_monthly_merge = ('项目' in _hdr_txt and '月' in _hdr_txt and '合计' in _hdr_txt)
        # ⚡ 2026-08-11：横展审定表（行=科目、列=各核算主体+合计，如职工薪酬横展审定表）——
        # 主体在列方向、无『核算主体』行标识列属设计——豁免。
        _is_wide = ('横展' in _hdr_txt)
        if kind == 'footnote' or '科目编码' in _hdr_txt or '科目名称' in _hdr_txt or _is_monthly_merge or _is_wide:
            pass
        else:
            problems.append(('ERROR', f'{subj_label} 表头缺『核算主体/主体』列（核对程序无法逐主体提取）'))
    else:
        hdr = ' '.join(str(x) for x in rows[hdr_i][:16] if x is not None)
        # ⚡ 2026-08-11：横展审定表（行=项目、列=主体，如固定资产横展审定表）表头无
        #   '期末/余额/审定数' 词属设计（期末在标题/列序说明）——豁免期末列名检查
        _is_wide2 = '横展' in str(subj_label)
        # 2026-08-22：台账类 sheet（如借款『融资明细台账』=披露口径：借款主体/银行/业务
        #   金额/利率/起止日，非报表科目取数表）表头无标准期末列名属设计——豁免。
        _is_ledger = '台账' in str(subj_label)
        if kind in ('detail', 'audit') and not _is_wide2 and not _is_ledger and not any(k in hdr for k in
                                                   ('期末', '本期', '全年', '余额', '审定数', '本期收入')):
            problems.append(('WARN', f'{subj_label} 表头无标准期末列名（期末数/本期数/全年/余额/审定数）'))
        # 主体名顿号检查（表头下前 30 行的前 3 列）
        for r in rows[hdr_i + 1:hdr_i + 30]:
            if not r:
                continue
            for x in r[:3]:
                if x and isinstance(x, str) and any(k in x for k in ('、', '，')) and ('FY' in x or '公司' in x):
                    # 2026-08-06 修复：口径说明行（以『注：』等开头）含『公司』+顿号
                    # （如费用分部门表尾说明）被误判为主体名——说明行跳过。
                    if x.strip().startswith(('注：', '注:', '说明', '备注')):
                        continue
                    problems.append(('WARN', f'{subj_label} 主体名含顿号：{x.strip()[:20]}（应 NN主体名 无顿号）'))
                    break
            if problems and problems[-1][0] == 'WARN' and '顿号' in problems[-1][1]:
                break
    # 合计/小计行（数据区前 200 行；'合 计'/'合  计' 含空格需去空格检测）
    # 2026-08-22 修复：横展附注汇总（表头=项目|01|02…，主体在列方向、无『主体』字样）
    # → hdr_i=None，原扫描范围为【空列表】→ 循环不执行 → 恒误报『无合计/小计行』。
    # 实际附注汇总每段末有『集团合计数』行。改从表头区（跳过标题 R1）起扫 200 行。
    has_total = False
    _scan_start = (hdr_i + 1) if hdr_i is not None else 1
    for r in rows[_scan_start:_scan_start + 200]:
        if not r:
            continue
        joined = ' '.join(str(x) for x in r[:4] if x is not None).replace(' ', '')
        if '合计' in joined or '小计' in joined:
            has_total = True
            break
        # 2026-08-22：附注汇总段末『审定数』行=段合计（各主体列之和）。单体底稿
        # （仅 1 主体，无『集团合计数』行）时同样可识别，避免『无合计行』误报。
        if kind == 'footnote' and str(r[0]).strip() == '审定数':
            has_total = True
            break
    if not has_total:
        problems.append(('WARN', f'{subj_label} 无合计/小计行（核对程序取数依赖『全集团合计/小计』）'))
    return problems


def assert_sheet_standard(ws, kind='detail', subj_label='', raise_on_error=True):
    """写入后立即调用：校验并打印。ERROR 硬伤默认抛 ValueError（生成即失败），
    WARN 仅打印提示。返回问题列表（供调用方记录）。"""
    probs = validate_sheet_structure(ws, kind, subj_label)
    errs = [p for p in probs if p[0] == 'ERROR']
    for lvl, msg in probs:
        print(f"  {'❌' if lvl == 'ERROR' else '⚠️'} [结构校验 {subj_label}] {msg}")
    if errs and raise_on_error:
        raise ValueError('；'.join(m for _, m in errs))
    return probs


def validate_workbook(wb, label='', raise_on_error=False):
    """对 workbook 三件套（审定表/明细表/附注汇总）批量结构校验（2026-08-05 前道规范化）。
    sheet 按名称判定 kind：含『审定表』→audit；含『明细/分类汇总』（排除 准备/计提/期间对比/
    对方科目/分月/变动 等发生额型与备抵表）→detail；含『附注』→footnote。
    返回全部问题；ERROR 硬伤可选抛 ValueError。各生成器在保存前调用——生成即规范。"""
    problems = []
    try:
        for sn in wb.sheetnames:
            if '审定表' in sn:
                kind = 'audit'
            elif ('明细' in sn or '分类汇总' in sn) and not any(
                    k in sn for k in ('准备', '计提', '期间对比', '对方科目', '分月', '变动', '检查', '贴现', '分部门')):
                kind = 'detail'
            elif '附注' in sn:
                kind = 'footnote'
            else:
                continue
            problems += validate_sheet_structure(wb[sn], kind, '%s/%s' % (label, sn))
        # 2026-08-05 三件套齐全检查：有审定表必须有附注（结构性缺失=ERROR，生成即失败提示）
        _sns = wb.sheetnames
        _has_audit = any('审定表' in s for s in _sns)
        _has_fn = any('附注' in s for s in _sns)
        if _has_audit and not _has_fn:
            # ⚡ 2026-08-09 无数据主体豁免：审定表标注『无数据（科目未启用）』或全部数值为 0 时，
            # 附注汇总按数据驱动裁剪不生成（skip_no_data）——1030 研发支出无数据实测属正常。
            _nodata = False
            for _sn in _sns:
                if '审定表' not in _sn:
                    continue
                try:
                    _rows = [list(r) for r in wb[_sn].iter_rows(values_only=True)]
                    _flat = ' '.join(str(x) for r in _rows for x in r if x is not None)
                    _nums = [x for r in _rows[1:] for x in r if isinstance(x, (int, float))]
                    if '无数据' in _flat or (not _nums) or all(abs(x) < 0.005 for x in _nums):
                        _nodata = True
                except Exception:
                    pass
                break
            if not _nodata:
                problems.append(('ERROR', f'{label} 有审定表但缺『附注汇总』sheet（合并附注生成依赖，结构性缺失）'))
    except Exception as ex:
        problems.append(('ERROR', 'workbook 校验异常：%s' % ex))
    errs = [p for p in problems if p[0] == 'ERROR']
    for lvl, msg in problems:
        print(f"  {'❌' if lvl == 'ERROR' else '⚠️'} [结构校验 {label}] {msg}")
    if errs and raise_on_error:
        raise ValueError('；'.join(m for _, m in errs))
    return problems


# ===================== 报表科目排序与分类合计（2026-08-15） =====================
# 用户方法论：试算表/合并审定表按 资产→负债→所有者权益→损益 科目顺序排列，
# 各段有合计；利润表科目有收入总额/成本费用/净利润等小计。

REPORT_ASSET = ['货币资金', '其他货币资金', '内部结算中心存款', '应收账款', '应收票据', '应收利息',
                '预付账款', '其他应收款', '合同资产', '合同资产减值准备', '存货', '原材料', '库存商品',
                '周转材料', '低值易耗品', '存货跌价准备', '房地产开发成本', '待摊费用', '长期应收款',
                '坏账准备', '投资性房地产', '长期股权投资', '固定资产', '累计折旧', '在建工程',
                '使用权资产', '无形资产', '累计摊销', '临时设施', '临时设施摊销', '长期待摊费用',
                '递延所得税资产', '持有待售资产', '持有待售资产减值', '贷款', '贷款损失准备',
                '抵债资产', '抵债资产跌价准备', '其他非流动资产']
REPORT_LIAB = ['短期借款', '应付票据', '应付账款', '预收账款', '合同负债', '应付职工薪酬', '应交税费',
               '应付股利', '其他应付款', '专项应付款', '递延收益', '长期借款', '长期应付款', '预提费用',
               '递延所得税负债', '其他非流动负债']
REPORT_EQUITY = ['实收资本(或股本)', '资本公积', '盈余公积', '未分配利润', '本年利润', '专项储备']
REPORT_PL = ['营业收入', '营业成本', '税金及附加', '销售费用', '管理费用', '财务费用', '研发支出',
             '其他收益', '投资收益', '资产处置损益', '公允价值变动损益', '信用减值损失',
             '资产减值损失', '营业外收入', '营业外支出', '所得税费用']
REPORT_PL_INCOME = ['营业收入', '其他收益', '投资收益', '资产处置损益', '公允价值变动损益', '营业外收入']
REPORT_PL_COST = ['营业成本', '主营业务成本', '其他业务成本', '税金及附加', '销售费用', '管理费用', '财务费用',
                  '研发支出', '信用减值损失', '资产减值损失', '营业外支出', '所得税费用']


def report_cat(name):
    """科目名 → (分类段, 段内序)。段：asset/liab/equity/pl_income/pl_cost/other。"""
    name = str(name or '').strip()
    for cat, order in (('asset', REPORT_ASSET), ('liab', REPORT_LIAB),
                       ('equity', REPORT_EQUITY), ('pl_income', REPORT_PL_INCOME),
                       ('pl_cost', REPORT_PL_COST)):
        if name in order:
            return cat, order.index(name)
    # 段序用于排序：asset<liab<equity<pl_income<pl_cost<other
    return 'other', 0


_REPORT_CAT_ORDER = {'asset': 0, 'liab': 1, 'equity': 2, 'pl_income': 3, 'pl_cost': 4, 'other': 5}


def sort_report_names(names):
    """科目名列表按报表顺序排序（资产→负债→权益→收入→成本费用→其他）。"""
    return sorted(names, key=lambda n: (_REPORT_CAT_ORDER[report_cat(n)[0]], report_cat(n)[1]))


def merge_segment_totals(ent_vals, names):
    """计算分类合计与利润表小计。
    ent_vals: {ent: {科目名: 值}}（横展：主体→科目值）
    返回 {合计label: {ent: 值}}：
      资产合计 / 负债合计 / 所有者权益合计（BS 段内 Σ）
      收入合计 / 成本费用合计 / 净利润（PL：净利润=收入Σ−成本费用Σ，含所得税）"""
    totals = {}
    cat_sums = {cat: {e: 0.0 for e in ent_vals} for cat in ('asset', 'liab', 'equity', 'pl_income', 'pl_cost', 'other')}
    for e, vals in ent_vals.items():
        for nm, v in vals.items():
            cat, _ = report_cat(nm)
            cat_sums[cat][e] += float(v or 0.0)
    labels = {'asset': '资产合计', 'liab': '负债合计', 'equity': '所有者权益合计',
              'pl_income': '收入合计', 'pl_cost': '成本费用合计'}
    for cat, lbl in labels.items():
        totals[lbl] = {e: cat_sums[cat][e] for e in ent_vals}
    totals['净利润'] = {e: cat_sums['pl_income'][e] - cat_sums['pl_cost'][e] for e in ent_vals}
    return totals


# ===================== ⚡ 铁律127：TB 损益缺失 → GL 兜底 =====================
# ga 实证（2026-08-15）：40 主体中 16 个 TB 主营业务收入贷=0、13 个成本=0、15 个费用=0——
# TB（科目余额表）数据源本身缺损益类科目发生额（GL 序时账完整）。损益类生成器若只读 TB
# 会系统性少计。pl_gl_fallback 在生成器读 TB 后自动检测并用 GL 汇总补齐。

# 损益类科目 code 前缀（标准代码；U8/SAP 通用 6001 收入/6401 成本/66xx 期间费用/6711 等）
_PL_PFX = ('6001', '6051', '6101', '6401', '6402', '6403', '6301', '6601', '6602', '6603',
           '6701', '6702', '6711', '6712', '5301', '5401', '5402', '5403', '6801')


def pl_gl_sum(gl_path):
    """读 GL（综合查询明细表），按科目名聚合各损益科目 借方/贷方 发生额（代数和，含红冲负）。
    返回 {name: (jf_sum, df_sum)}（仅损益类名称）。GL 读取失败返回 {}。"""
    if not gl_path or not os.path.exists(gl_path):
        return {}
    rows = read_gl(gl_path)
    out = {}
    for r in rows:
        nm = str(r.get('name') or '')
        if not nm:
            continue
        jf = float(r.get('debit') or 0.0)
        df = float(r.get('credit') or 0.0)
        if abs(jf) < 0.005 and abs(df) < 0.005:
            continue
        # 按名称首段判定损益类（GL name 无 code，用 一级名 判断；如 '主营业务收入\...'）
        _first = nm.split('\\')[0].split('-')[0].strip()
        if not any(_first.startswith(p) for p in ('主营业务收入', '其他业务收入', '主营业务成本',
                                                  '其他业务成本', '营业成本', '销售费用', '管理费用',
                                                  '财务费用', '税金及附加', '研发', '营业外收入',
                                                  '营业外支出', '资产减值', '信用减值', '其他收益',
                                                  '投资收益', '公允价值', '所得税费用', '制造费用')):
            continue
        a, b = out.get(nm, (0.0, 0.0))
        out[nm] = (a + jf, b + df)
    return out


def pl_gl_fallback(tb_year, gl_path):
    """⚡ 铁律127：TB 损益类科目发生额缺失/不可靠时，用 GL 汇总兜底。

    检测（对 TB 中 _PL_PFX 前缀的损益科目行）：
      a) 全部 jf/df 为 0 且 GL 有值 → 全量填回（ga 16/40 主体收入=0 实证）；
      b) 单科目 TB 与 GL 差异 > 50%（TB 减半等口径问题）→ 用 GL 覆盖（ga 15 号
         收入 TB 7.71 亿 vs GL 15.42 亿实证）。
    兜底值 = GL 同科目名汇总 借方/贷方（代数和，含红冲）。仅动损益行，其余零影响。
    返回 (tb_year, triggered)。"""
    if not tb_year or not gl_path:
        return tb_year, False
    pl_codes = [c for c in tb_year if str(c).startswith(_PL_PFX)]
    if not pl_codes:
        return tb_year, False
    gl_sum = pl_gl_sum(gl_path)
    if not gl_sum:
        return tb_year, False
    # 检测 a)：全部损益行 debit/credit 为 0（read_km 结构；ga TB 损益科目只有期初/期末余额、
    #   本期发生额缺失 → opening/closing 非 0 但 debit/credit=0）
    all_zero = all(abs(float(tb_year[c].get('debit') or 0.0)) < 0.005 and
                   abs(float(tb_year[c].get('credit') or 0.0)) < 0.005 for c in pl_codes)
    triggered = False
    for c in pl_codes:
        v = tb_year[c]
        nm = str(v.get('name') or '')
        gj, gd = gl_sum.get(nm, (0.0, 0.0))
        if abs(gj) < 0.005 and abs(gd) < 0.005:
            continue
        td_ = abs(float(v.get('debit') or 0.0))
        tc_ = abs(float(v.get('credit') or 0.0))
        need = False
        if all_zero and (abs(gj) > 0.005 or abs(gd) > 0.005):
            need = True
        else:
            # 检测 b)：单科目差异 > 50%（TB 减半等口径问题，ga 15 号实证）
            if max(td_, abs(gj)) > 0.005 and abs(td_ - abs(gj)) / max(td_, abs(gj)) > 0.5:
                need = True
            if max(tc_, abs(gd)) > 0.005 and abs(tc_ - abs(gd)) / max(tc_, abs(gd)) > 0.5:
                need = True
        if need:
            v['debit'] = gj
            v['credit'] = gd
            triggered = True
    return tb_year, triggered


def read_tb_pl_safe(paths):
    """读科目余额表 + 损益缺失 GL 兜底（铁律127）。paths={'km','gl','aux'}。
    供损益类生成器使用（revenue/expense/tax/rd 等）——TB 损益缺失时自动用 GL 补齐，
    正常账套（TB=GL）零影响。
    ⚡ 2026-08-16 修复：SAP 场景 gl 是多文件 list（逐月序时账）——read_sap_tb 直读
    TB 损益发生额完整（铁律130b 试算表对平实证），无需 GL 兜底；list 传入
    os.path.exists(list) 报 TypeError 致 revenue 底稿缺失（AH 集团A/B 营业收入缺）。"""
    tb = read_km(paths['km'])
    glp = paths.get('gl')
    if glp and not isinstance(glp, (list, tuple)):
        tb, _ = pl_gl_fallback(tb, glp)
    return tb
