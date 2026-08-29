# -*- coding: utf-8 -*-
"""SAP 账套读取层（sap_reader.py，2026-08-07 新建）。

SAP 导出与用友 U8 结构完全不同，本模块把 SAP 数据归一化为程序内部
标准结构（TB/GL/aux），供现有 13 个 builder + 自建试算表使用。

支持的两种 SAP 导出布局：
  A) ah-sap 分目录式：<账套>/{科目余额表,序时账,应收,应付,客户行项目,供应商行项目,费用,利润表,资产负债表}/
  B) 3300 单文件夹式：<账套>/{<N>科目余额表<年>.XLSX, <M>月.XLSX}

核心设计（2026-08-07 v2 架构：会计语言优先，代码方言隔离）：
  - 科目识别以【名称】为主键（SAP 科目名=层级名，如 银行存款-工行-人民币-一般户）；
  - 科目余额表公司代码列=「公司名 代码」→ 拆分实体名+代码；
  - 期间=月度行（剩余 2026=期初结转，一月~十二月 2026=月度行）；
  - 利润中心拆行镜像抵消（未分配的/事业部 金额互为相反数）→ 按公司聚合防双计；
  - 虚增虚减识别 classify_voucher()：文本关键词 + PK 过账码 + 「（GL）」对科目排重。

数据归一化后的内部结构（与 audit_common 对齐）：
  TB:  {(实体, 科目代码, 科目名, 年度): {'qc','jf','df','qm','level'}}
  GL:  [{'y','e','date','vno','vtype','name','dr','cr','summary','opp','cust','supp'}]
  aux: read_opening_balances 兼容（应收/应付往来明细）
"""
import paths as P
import os
import re
import glob
import json

_KB = ('结转', '重分类', '冲销', '清账', '抵消', '调整', '期末', '汇兑', '红冲',
       '分摊', '结算', '汇总', '内部', '计提', '暂估', '尾差')

# ---- 2026-08-23 布局显式声明（account_profiles.json accounts.<账套>.layout）----
# 声明名 → detect 返回值：'u8'→None（通用 U8 路径）、'300'、'ah'→'ah-sap'（AH SAP FAGL）、'3300'
_LAYOUT_DECL_MAP = {'u8': None, '300': '300', 'ah': 'ah-sap', '3300': '3300'}
_PROFILE_CACHE = {}


def _acct_from_data_dir(data_dir):
    """从数据目录路径推断账套名（.../账套/数据/年度）——仅当路径含『数据』段。"""
    d = os.path.normpath(str(data_dir))
    parts = d.split(os.sep)
    for i, p in enumerate(parts):
        if p == '数据' and i > 0:
            return parts[i - 1]
    return None


def _load_declared_layout(acct):
    """读 account_profiles.json 的 accounts.<账套>.layout（进程内缓存）。"""
    global _PROFILE_CACHE
    if not _PROFILE_CACHE:
        try:
            fp = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'account_profiles.json')
            with open(fp, encoding='utf-8') as f:
                prof = json.load(f)
            _PROFILE_CACHE = prof.get('accounts', {})
        except Exception:
            _PROFILE_CACHE = {}
    cfg = _PROFILE_CACHE.get(acct)
    if isinstance(cfg, dict):
        return cfg.get('layout')
    return None


def _apply_declared_layout(detected, data_dir):
    """检测结果与显式声明冲突时，以声明为准并告警（根治文件名启发式误判）。"""
    acct = _acct_from_data_dir(data_dir)
    if not acct:
        return detected
    declared = _load_declared_layout(acct)
    if not declared:
        return detected
    expected = _LAYOUT_DECL_MAP.get(declared, object())
    if expected == detected:
        return detected
    det_name = detected or 'U8'
    print(f"⚠️ layout 声明冲突：{acct} 声明={declared}，检测={det_name}——以声明为准（{declared}）")
    return expected

# 统计性/凭证流过账码（SAP PK）——虚增虚减识别
_STAT_PK = ('99',)                    # 统计性过账
_FLOW_PK = ('81', '83', '86', '89', '91', '93', '96')  # 凭证流（清账/冲销/未清项）

# 对科目排除词：仅排除明确非真实记账科目（SAP 同名校对科目）
# 注：『（GL）』科目是真实总账记账（1010 制造收入（GL）净额 6.79亿=报表收入差异），
#     不能全局排除；虚增虚减由序时账 classify_voucher 处理。
_GL_DUP_KW = ('虚拟', '不再使用', '统计科目')


# ============================================================
# 一、布局发现
# ============================================================
def _looks_sap300(fp):
    """表头内容探测：SAP 300 形态科目余额表含『总账科目长文本/总账』列（ADF 等）；
    U8 含『科目名称/科目编码/科目代码』列（AZ 币种分目录、AJ 分主体目录等）。"""
    import openpyxl
    wb = None
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        for r in ws.iter_rows(values_only=True):
            joined = ' '.join(str(x) for x in r if x is not None)
            if '科目' not in joined:
                continue
            if '总账科目长文本' in joined or '总账科目' in joined or '总账' in joined:
                return True
            if '科目名称' in joined or '科目编码' in joined or '科目代码' in joined:
                return False
        return False
    except Exception:
        return False
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def _looks_sap_ah(fp):
    """AH 账套 SAP FAGL 形态内容探测：科目余额表表头含『公司代码』+『科目号/利润中心/货币类型』
    （AH 数据/2026/科目余额表/1010.xlsx：货币类型|货币|科目号|功能范围|业务范围|段|利润中心|公司代码|期间/年度）。"""
    import openpyxl
    wb = None
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        for r in ws.iter_rows(values_only=True):
            joined = ' '.join(str(x) for x in r if x is not None)
            if '公司代码' in joined and ('科目号' in joined or '利润中心' in joined or '货币类型' in joined):
                return True
            if '科目名称' in joined or '科目编码' in joined or '科目代码' in joined:
                return False
        return False
    except Exception:
        return False
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass


def detect_sap_layout(data_dir):
    """识别账套数据形态，返回 SAP 专属形态标识；None = U8（通用路径）。
    '300'  = SAP 300：每主体一子文件夹，科目余额表含『总账科目长文本/总账』列（ADF）
    'ah-sap' = AH 账套 SAP FAGL：顶层『科目余额表/序时账』目录，科目余额表含『公司代码/科目号』列
               ⚠️ 历史标识名 ah-sap 并非"用友2026"——AH 实为 SAP 导出（2026-08-23 用户确认）。
    '3300' = SAP 3300 单文件夹（保留历史形态）。
    其余（AZ 币种分目录/AJ 分主体目录/GJX 单文件夹 U8）一律返回 None(U8)。
    2026-08-23 起：检测后与 account_profiles.json 的 accounts.<账套>.layout 显式声明比对，
    冲突时以声明为准并告警（根治文件名启发式误判）。"""
    detected = _detect_sap_layout_core(data_dir)
    return _apply_declared_layout(detected, data_dir)


def _detect_sap_layout_core(data_dir):
    if not os.path.isdir(data_dir):
        return None
    subs = set(os.listdir(data_dir))
    # ⚡⚡ 2026-08-23 判定全部改为【表头内容探测】，不再凭文件名猜形态。
    #   AH（SAP FAGL）：顶层『科目余额表』目录内文件为 SAP 表头（公司代码/科目号/利润中心）。
    if '科目余额表' in subs and '序时账' in subs:
        kd = os.path.join(data_dir, '科目余额表')
        try:
            for f in sorted(os.listdir(kd)):
                if f.lower().endswith('.xlsx'):
                    if _looks_sap_ah(os.path.join(kd, f)):
                        return 'ah-sap'
                    break
        except OSError:
            pass
    if any(re.search(r'科目余额表\d*\.XLSX$', f, re.I) for f in subs) and \
       any(re.match(r'^\d+月\.XLSX$', f, re.I) for f in subs):
        # ⚡⚡ 2026-08-23 3300 形态同样需表头内容验证：U8 账套顶层『XX科目余额表.XLSX』
        #   也命中文件名前半规则（ACB/ADZ/AFJ/AJ/AQ/AS/ATT/AYL 等），若恰好有按月文件即误判。
        #   表头含『总账科目长文本/总账』（SAP）才判 3300。
        for f in subs:
            m = re.search(r'科目余额表.*\.XLSX$', f, re.I)
            if m and _looks_sap300(os.path.join(data_dir, f)):
                return '3300'
    # ⚡⚡ 2026-08-23 删除 dq 文件名规则：GJX『DQ2025科目余额表.xlsx』实为 U8 表头
    #   （科目代码|科目名称|方向|金额|...）——GJX 走通用 U8 路径（_discover_dq 已清理，2026-08-24）。
    # ⚡ 2026-08-10 300 形态：每账套一个子文件夹，内含『科目余额表』文件名
    # （如 300/3700/3700科目余额表20260731.XLSX + 序时账 + 客户/供应商余额表）
    # ⚡⚡ 2026-08-23 文件名启发式不可靠（AZ 币种分目录『数链2025科目余额表.xlsx』、
    #   AJ 分主体目录『01...2025年科目余额表.xlsx』都是 U8 导出也含『科目余额表』）
    #   → 改为【表头内容探测】：SAP 300 含『总账科目长文本/总账』列，U8 含『科目名称/科目编码』。
    for d in subs:
        # ⚡⚡ 2026-08-23 排除描述性目录（AFJ『其他应收其他应付科目余额表』等非账套
        #   主体目录）：其内含『XX科目余额表.xlsx』会误判为 SAP 300 形态（→ U8 账套
        #   被误走 read_sap_tb → 0 主体）。主体目录名应为账套代码/名称，不含文件
        #   类型描述词。
        if re.search(r'余额表|科目余额|报表|辅助核算|明细表|账龄', d):
            continue
        dp = os.path.join(data_dir, d)
        if os.path.isdir(dp):
            try:
                inner = os.listdir(dp)
            except OSError:
                continue
            for f in inner:
                if '科目余额表' in f and f.lower().endswith('.xlsx'):
                    if _looks_sap300(os.path.join(dp, f)):
                        return '300'
    return None


def discover_sap_entities(data_dir):
    """发现 SAP 实体：返回 {实体: {年度: {'km': path, 'gl': path}}}。
    ah-sap：公司代码列（名称+代码）→ 实体=代码；合并文件内多公司拆分。
    3300：单文件夹，实体=公司代码前缀（3300科目余额表2025 → 3300/2025）。
    300（2026-08-10）：每账套一个子文件夹（3700/3900/...），实体=文件夹名。"""
    layout = detect_sap_layout(data_dir)
    ent = {}
    if layout == 'ah-sap':
        return _discover_ah_sap(data_dir)
    if layout == '3300':
        return _discover_3300(data_dir)
    if layout == '300':
        return _discover_300(data_dir)
    return ent


def _discover_300(data_dir):
    """300 形态：每账套一个子文件夹，文件在文件夹内：
       <前缀>科目余额表<日期>.XLSX（取最新日期=731 版）、<前缀>序时账*.XLSX（可多文件合并）、
       客户余额表/供应商余额表（往来 aux，取最新日期）。
       aux 通过 discover 返回的 'aux' 键传给 read_aux_balance（300 分支按此路径读）。"""
    ent = {}
    for d in sorted(os.listdir(data_dir)):
        dp = os.path.join(data_dir, d)
        if not os.path.isdir(dp):
            continue
        try:
            files = os.listdir(dp)
        except OSError:
            continue
        tb = None
        tb_date = ''
        gls = []
        gls_xse = []   # ⚡ 2026-08-12：应收/应付借方贷方发生额（GL 筛选形态，无对方科目）
        unknown = []   # ⚡ 2026-08-12：文件名规则未命中的 xlsx（交表头内容兜底判定）
        aux_ar = None
        aux_ap = None
        aux_ar_date = aux_ap_date = ''
        for f in files:
            if not f.lower().endswith(('.xlsx', '.xls', '.txt')):
                continue
            fp = os.path.join(dp, f)
            if '科目余额表' in f:
                m = re.search(r'科目余额表(\d{8})?', f)
                date = (m.group(1) if m and m.group(1) else '')
                if date >= tb_date:      # 字符串比较 20260731 > 20260630 ✓（取 731 版）
                    tb, tb_date = fp, date
            elif '序时账' in f:
                gls.append(fp)
            elif '发生额' in f:   # ⚡ 2026-08-12：3200/3300/3400 变体（GL 筛选形态，含客户/供应商列）
                gls_xse.append(fp)
            elif '客户' in f and ('余额' in f or '账龄' in f):   # 3700客户余额表/3900客户余额/6900客户账龄
                m = re.search(r'(?:客户余额|客户账龄)(\d{8})?', f)
                date = (m.group(1) if m and m.group(1) else '')
                if date >= aux_ar_date:
                    aux_ar, aux_ar_date = fp, date
            elif '供应商' in f and ('余额' in f or '账龄' in f):   # 3700供应商余额表/3900供应商{date}余额/6900供应商账龄
                # ⚡ 2026-08-12 修复：供应商余额表日期后缀兼容 8 位（20260731）与
                #   MMDD（0630/0731，3500 改名后 `供应商余额表0630.xlsx`）——统一转
                #   MMDD4 比较（20260731→'0731'、0731→'0731'、630→'0630'），
                #   否则 8 位/4 位混用字符串比较失序（'0731' vs '20260731'）。
                def _mmdd4(d):
                    s = str(d or '')
                    if len(s) >= 8:
                        return s[-4:]
                    if len(s) == 4:
                        return s
                    if len(s) == 3:
                        return '0' + s
                    return ''
                m = re.search(r'(?:供应商.*?)(\d{8}|\d{3,4})?', f)
                date = (m.group(1) if m and m.group(1) else '')
                _d4 = _mmdd4(date)
                if _d4 >= aux_ap_date:
                    aux_ap, aux_ap_date = fp, _d4
            # ⚡ 2026-08-12 3500 变体：`3500供应商630.xlsx`/`3500供应商731.xlsx`
            #   （MMDD 日期后缀，无"余额/账龄"字样；供应商余额表 630=6月30日/731=7月31日）。
            #   与标准模式同一日期竞争逻辑（取最新日期；不能限定 aux_ap is None——
            #   否则 630 先匹配后 731 不再进此分支 → 误取旧版）。
            elif '供应商' in f and re.search(r'供应商(\d{3,4})\.', f):
                m = re.search(r'供应商(\d{3,4})\.', f)
                _d = m.group(1) if m else ''
                # 月份日期规范化：630→0630、731→0731（字符串比较 0731>0630 ✓）
                _d4 = _d if len(_d) == 4 else ('0' + _d if len(_d) == 3 else '')
                if _d4 >= aux_ap_date:
                    aux_ap, aux_ap_date = fp, _d4
            else:
                unknown.append(fp)   # ⚡ 2026-08-12：未命中任何文件名规则 → 表头内容兜底
        # ⚡ 2026-08-12 修复：主体有完整序时账时忽略"发生额"文件（3500 案例——用户
        #   11:53 新导 3500应收借方发生额.xlsx（44.4万行凭证小计）与 7 个月序时账重复，
        #   双计拉低对方科目核对覆盖率）。发生额仅在没有序时账时作为 GL 兜底。
        if not gls:
            gls = gls_xse
        # ⚡ 2026-08-12 变体自动识别兜底：文件名规则未命中的 xlsx（新主体命名变体），
        #   按【表头内容】判定 GL/客户余额/供应商余额——以后任何新命名变体零改代码接入
        #   （今日 3500/3200/3300/3400 逐个改文件名规则的教训）。
        for _f in unknown:
            _k = _classify_300_by_header(_f)
            if _k == 'gl':
                gls_xse.append(_f)
            elif _k == 'aux_ap' and not aux_ap:
                aux_ap = _f
            elif _k == 'aux_ar' and not aux_ar:
                aux_ar = _f
        if not gls:
            gls = gls_xse
        if tb:
            ent[d] = {'2026': {'km': tb, 'gl': gls, 'aux_ar': aux_ar, 'aux_ap': aux_ap}}
    return ent


def _classify_300_by_header(fp):
    """按表头内容判定 300 变体文件类型（文件名规则兜底）：
    'gl'=序时账/发生额形态（凭证编号+总账科目+金额列）；'aux_ap'=供应商余额表（供应商+余额）；
    'aux_ar'=客户余额表（客户+余额）；None=无法判定。用 _safx 流式只读前 12 行。"""
    import _safx
    try:
        hdrs = set()
        for r in _safx.iter_rows(fp, keep_cols=None):
            for v in r.values():
                if isinstance(v, str):
                    hdrs.add(str(v).strip())
            if len(hdrs) > 25 or _has(hdrs, ('凭证编号', '总账科目', '本币金额', '金额')):
                break
            if len(hdrs) > 60:
                break
    except Exception:
        return None
    # GL 序时账形态：凭证编号 + 总账科目 + 金额列（本币金额/金额/求和项:金额/本位币金额）
    if '凭证编号' in hdrs and '总账科目' in hdrs and any(k in hdrs for k in ('本币金额', '金额', '求和项:金额', '本位币金额')):
        return 'gl'
    if '供应商' in hdrs and any(k in hdrs for k in ('余额', '借方', '贷方', '期初', '期末')):
        return 'aux_ap'
    if '客户' in hdrs and any(k in hdrs for k in ('余额', '借方', '贷方', '期初', '期末')):
        return 'aux_ar'
    return None


def _has(hdrs, keys):
    return any(k in hdrs for k in keys)


def _discover_ah_sap(data_dir):
    """ah-sap 分目录：科目余额表/{公司代码}.xlsx 或 合并导出{A~B}.xlsx。
    ⚡ 2026-08-09 修复：①TB 区间合并文件（2250~2300.xlsx 等）只登记起点公司 → 漏 27 家，
    改为【TB 按文件内容扫描】公司代码（col7 尾部 4 位，_safx 流式）；
    ②序时账为单公司按月文件（{公司}-{月}月.xlsx）→ 文件名前缀直接映射，无需扫内容（防 OOM）。"""
    tb_dir = os.path.join(data_dir, '科目余额表')
    gl_dir = os.path.join(data_dir, '序时账')
    # 科目余额表文件 → 内含公司代码（合并文件多公司，按内容扫描）
    tb_comps = {}   # {公司代码: tb_path}
    for f in sorted(os.listdir(tb_dir)):
        if not f.lower().endswith('.xlsx'):
            continue
        fp = os.path.join(tb_dir, f)
        for code in _scan_tb_companies(fp):
            tb_comps.setdefault(code, fp)
    # 序时账文件 → 单公司按月文件：文件名 4 位前缀=公司代码（区间文件按内容补扫）
    gl_files = {}
    for f in sorted(os.listdir(gl_dir)):
        if not f.lower().endswith(('.xlsx', '.xls')):
            continue
        fp = os.path.join(gl_dir, f)
        base = os.path.splitext(f)[0]
        mpre = re.match(r'^(\d{4})', base)
        if not mpre:
            continue
        msuf = re.search(r'[-~](\d{4})$', base)
        if msuf and int(mpre.group(1)) < int(msuf.group(1)):
            # 区间合并文件：公司代码 = 文件名区间 ∩ TB 扫描公司全集
            # ⚡ 2026-08-12 修复：原按内容全扫 col6——大区间文件（2040-2080 等）
            #   每文件 6-23s，86 个序时账 discover 达 3-5 分钟（被误判卡死/超时）。
            #   TB 扫描（38 文件 9s）已给出公司全集，序时账公司必然 ⊆ TB 公司集
            #   → 区间 ∩ 全集 秒级等价，无需扫文件内容。
            for code in tb_comps:
                if int(mpre.group(1)) <= int(code) <= int(msuf.group(1)):
                    gl_files.setdefault(code, []).append(fp)
        else:
            gl_files.setdefault(mpre.group(1), []).append(fp)
    # 合并实体：公司代码全集 = TB 内容代码 ∪ GL 内代码
    ent = {}
    for code in sorted(set(tb_comps) | set(gl_files)):
        ent[code] = {'2026': {'km': tb_comps.get(code), 'gl': gl_files.get(code)}}
    return ent


def _scan_tb_companies(fp):
    """扫描科目余额表文件全部行提取公司代码（col7 尾部 4 位）。
    用 _safx 流式（zipfile+ET）只取 col7，防大区间文件 OOM。"""
    out = set()
    try:
        import _safx
        for r in _safx.iter_rows(fp, keep_cols={7}):
            v = r.get(7)
            if v is None:
                continue
            m = re.search(r'(\d{4})\s*$', str(v).strip())
            if m:
                out.add(m.group(1))
    except Exception:
        pass
    return out


def _scan_gl_companies(fp, limit=None):
    """扫描序时账文件行提取公司代码（col6，纯 4 位）。
    ⚡ 2026-08-09：limit=200 会漏大区间文件（2610~5040 等）后段公司 → 默认全扫，
    且改用 _safx 流式只取 col6，防 openpyxl 全表迭代 OOM。"""
    out = set()
    try:
        import _safx
        for i, r in enumerate(_safx.iter_rows(fp, keep_cols={6})):
            if limit is not None and i > limit:
                break
            c = str(r.get(6) or '').strip()
            if re.match(r'^\d{4}$', c):
                out.add(c)
    except Exception:
        pass
    return out


def _discover_3300(data_dir):
    """3300 单文件夹：<N>科目余额表<年>.XLSX + <M>月.XLSX。"""
    ent = {}
    tb = gl = None
    year = None
    for f in sorted(os.listdir(data_dir)):
        m = re.search(r'(\d{4})科目余额表(\d{4})\.XLSX$', f, re.I)
        if m:
            tb = os.path.join(data_dir, f)
            year = m.group(2)
            code = m.group(1)
        elif re.match(r'^\d+月\.XLSX$', f, re.I):
            gl = gl or os.path.join(data_dir, f)
    if tb and code:
        ent[code] = {year: {'km': tb, 'gl': gl}}
    return ent


# ============================================================
# 二、科目余额表 → TB（内部结构）
# ============================================================
def parse_sap_tb_name(raw):
    """科目号列原始值 → (名称, 内部码)。如
    '库存现金-人民币-杭州   HY01/1001010200' → ('库存现金-人民币-杭州', '1001010200')
    SAP 名称与内部码由 2+ 空格分隔；内部码格式 HY01/xxxx 或 xxxxxxxx。"""
    if raw is None:
        return '', ''
    s = str(raw).strip()
    m = re.match(r'^(.*?)\s{2,}([A-Z]+\d*/\d+|\d+)$', s)
    if m:
        return m.group(1).strip(), m.group(2)
    return s, ''


def _signed_tb(amt, direction):
    """余额方向列 → 带符号余额：借=+、贷=−、平/空=0（300 形态 TB 用）。"""
    d = str(direction or '').strip()
    if '贷' in d:
        return -abs(amt)
    if '借' in d:
        return abs(amt)
    return 0.0


def read_sap_tb(data_dir, entities, year='2026'):
    """读取 SAP 科目余额表 → {(实体, 科目名, 名称, 年度): {qc,jf,df,qm,level}}。
    ah-sap/3300：见下方通用分支；300 形态（2026-08-10）：11 列汇总式一行一科目。"""
    import openpyxl
    tb = {}
    layout = detect_sap_layout(data_dir)
    # ⚡ 2026-08-10 300 形态：11 列汇总式科目余额表（一行一科目）
    # col0公司代码/col1总账科目(10位)/col2科目长文本/col3期初金额/col4期初方向/
    # col5本期借(=1-7月累计)/col6本期贷/col7本年累计借/col8本年累计贷/col9期末余额/col10期末方向
    if layout == '300':
        # ⚡ 2026-08-23 列探测收口：与 read_tb_full 共用 probe_tb_columns（audit_common），
        #   避免两处各自探测漂移。
        from audit_common import probe_tb_columns
        for d in sorted(os.listdir(data_dir)):
            dp = os.path.join(data_dir, d)
            if not os.path.isdir(dp):
                continue
            for f in os.listdir(dp):
                if '科目余额表' not in f or not f.lower().endswith(('.xlsx', '.xls')):
                    continue
                fp = os.path.join(dp, f)
                try:
                    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
                except Exception:
                    continue
                ws = wb[wb.sheetnames[0]]
                # ⚡ 2026-08-10 表头动态映射：3700/3900/6900 与 6100 列序不同
                # （6100 期初方向在 col3、期初金额在 col5；其余账套期初金额 col3）。
                # read_only 流式：首行=表头，后续行=数据。
                probed = False
                i_code = i_name = i_qc = i_qc_dir = i_jf = i_df = i_qm = i_qm_dir = 0
                for r in ws.iter_rows(values_only=True):
                    if not probed:
                        probed = True
                        if r:
                            h = probe_tb_columns(r)
                            i_code = h['code']; i_name = h['name']
                            i_qc = h['qc_amt']; i_qc_dir = h['qc_fx']
                            i_jf = h['jf']; i_df = h['df']
                            i_qm = h['qm_amt']; i_qm_dir = h['qm_fx']
                        continue
                    if not r or len(r) <= max(i_code, i_name, i_qc, i_qc_dir,
                                              i_jf, i_df, i_qm, i_qm_dir):
                        continue
                    try:
                        code = str(r[i_code] or '').strip()
                        name = str(r[i_name] or '').strip()
                    except Exception:
                        continue
                    if not code or not name:
                        continue
                    # ⚡ 2026-08-10 300 形态「净额列示」镜像行排除：SAP 递延所得税资产/负债
                    #   在同一 4 位科目下额外生成『净额列示』行（290107/181111 等），其
                    #   余额/发生额与真实明细行相反（镜像）→ 带符号求和=0 → 审定表/附注
                    #   全 0（3900 DTL 固定资产加速折旧 24.2 万被净额列示行抵销）。
                    #   净额列示=报表展示辅助行，非真实业务明细 → 排除。
                    if '净额列示' in name:
                        continue
                    comp = str(r[0] or '').strip()   # 公司列固定 col0
                    try:
                        qc = _signed_tb(float(r[i_qc] or 0), str(r[i_qc_dir] or ''))
                        qm = _signed_tb(float(r[i_qm] or 0), str(r[i_qm_dir] or ''))
                        jf = float(r[i_jf] or 0)
                        df = float(r[i_df] or 0)
                    except (ValueError, TypeError):
                        continue   # 格式异常行（如说明/合计行）跳过
                    key = (comp, code, name, str(year))
                    tb[key] = {'qc': qc, 'jf': jf, 'df': df,
                               'qm': qm, 'level': len(code)}
                wb.close()
        return tb
    tb_dir = os.path.join(data_dir, '科目余额表') if layout == 'ah-sap' else data_dir
    files = []
    if layout == 'ah-sap':
        for f in sorted(os.listdir(tb_dir)):
            if f.lower().endswith('.xlsx'):
                files.append(os.path.join(tb_dir, f))
    else:
        for f in sorted(os.listdir(tb_dir)):
            if re.search(r'科目余额表.*\.XLSX$', f, re.I):
                files.append(os.path.join(tb_dir, f))
    for fp in files:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        # 聚合键：科目名+期间 → 跨利润中心求和（镜像抵消）
        agg = {}
        for r in ws.iter_rows(min_row=3, values_only=True):
            if not r or len(r) < 17 or r[2] is None:
                continue
            name, code_raw = parse_sap_tb_name(r[2])
            if not name or any(k in name for k in _GL_DUP_KW):
                continue
            comp_raw = str(r[7]).strip() if r[7] else ''
            m = re.search(r'(\d{4})\s*$', comp_raw)
            if not m:
                continue
            comp = m.group(1)
            period = str(r[8]).strip() if r[8] else ''
            code = code_raw.split('/')[-1] if code_raw else ''
            # ⚡⚡ 2026-08-27 修复（任务885 工资 2000 万差异根因）：聚合键加 code——
            #   原 key=(comp,name,period) 把【同名不同 code】的子目错误合并（1010 实发工资
            #   2211101011/2211101012 两个独立科目名称相同 → 合并后 code 被覆盖为 1012、
            #   金额=两者之和），payroll 用 code 前缀匹配读到合并值 vs GL 分开的 1011 计提
            #   → 控制差异 2,135 万（GL 计提 3.16亿 vs TB 2.94亿）。按 code 分开后各自独立。
            key = (comp, code, name, period)
            a = agg.setdefault(key, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0, 'code': ''})
            a['code'] = code
            if '剩余' in period:
                a['qc'] += float(r[9] or 0)
            else:
                # ⚡⚡ 2026-08-28 修复（专项应付款明细表红字）：SAP 发生额列带符号——
                #   正=该方向发生，负=红字（反向发生）。贷方余额科目（专项应付款等）的
                #   部分贷方发生显示为"借方列负数"（红字借方）→ 原直接累加把负值带进
                #   jf → 明细表"本期减少"红字、"本期增加"少算。
                #   修复：按"正数 + 对方列红字"还原毛发生额（jf−df 净额不变）：
                #     借方发生 = Σ max(col12,0) + Σ max(-col13,0)
                #     贷方发生 = Σ max(col13,0) + Σ max(-col12,0)
                _j12 = float(r[12] or 0)
                _c13 = float(r[13] or 0)
                a['jf'] += max(_j12, 0.0) + max(-_c13, 0.0)
                a['df'] += max(_c13, 0.0) + max(-_j12, 0.0)
            if '七月' in period:
                a['qm'] += float(r[16] or 0)
        wb.close()
        # 转内部结构：期末=七月累计余额；借发/贷发=各月累计和
        # ⚡ 2026-08-09（建议2 完整版）：key 第二项改【真实内部码】（原=名称，导致下游依赖
        #    内部码的逻辑静默失效）。调用方核对：load_tb_raw/rd_expense/recon/tb_gen 均用
        #    _code 占位忽略，仅 gp_other/coverage_scan 使用 code（已同步适配）。
        for (comp, _code, name, period), a in agg.items():
            key = (comp, a.get('code', ''), name, str(year))
            v = tb.setdefault(key, {'qc': 0.0, 'jf': 0.0, 'df': 0.0, 'qm': 0.0, 'level': 4})
            v['qc'] += a['qc']
            v['jf'] += a['jf']
            v['df'] += a['df']
            v['qm'] += a['qm']
            if not v.get('code'):
                v['code'] = a.get('code', '')
    return tb


# ============================================================
# 三、序时账 → GL（内部结构，读取即聚合防 OOM）
# ============================================================
def classify_voucher(pk, text):
    """虚增虚减凭证分类：返回 'normal' | 'stat' | 'flow' | 'text'。
    - stat: 统计性过账（PK=99）
    - flow: 凭证流（PK 清账/冲销等）
    - text: 文本含结转/重分类/冲销/抵消等关键词
    - normal: 真实业务
    """
    pk_s = str(pk or '').strip()
    if pk_s in _STAT_PK:
        return 'stat'
    if pk_s in _FLOW_PK:
        return 'flow'
    txt = str(text or '')
    if any(k in txt for k in _KB):
        return 'text'
    return 'normal'


def _read_sap_gl_colnamed(fp):
    """⚡⚡ 按【列名定位】读 SAP 序时账（2026-08-16 3400/3300 适配）：
    3300(xlsx sharedStrings) 与 3400(txt GBK 制表符) 列序不同——必须按表头列名定位，
    不能按固定列位置（ah-sap 39 列注释仅适用于 AH）。
    列名映射：凭证编号→vno、凭证日期→date、本币金额→金额、科目→name(代码)、
              文本→summary、客户→cust、供应商→supp、PK→pk、公司→e、财年→y。
    金额正=借 负=贷；日期兼容 MM/DD/YYYY、YYYY-MM-DD、Excel 序列号。
    返回标准行 dict 列表。流式读取（txt 逐行 / xlsx iterparse），不整表缓存。
    """
    import zipfile
    import xml.etree.ElementTree as ET
    from datetime import date, timedelta
    ext = os.path.splitext(fp)[1].lower()
    out = []

    def _parse_date(v):
        v = str(v or '').strip()
        if not v:
            return ''
        m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', v)
        if m:
            return f'{m.group(3)}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
        m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', v)
        if m:
            return f'{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}'
        if v.isdigit() and int(v) > 40000:   # Excel 序列号
            try:
                return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
            except Exception:
                pass
        return ''

    def _amt(v):
        try:
            return float(str(v or 0).replace(',', '').replace('，', '').strip())
        except ValueError:
            return 0.0

    def _emit(rowd):
        vno = rowd.get('凭证编号', '')
        amt = _amt(rowd.get('本币金额', 0))
        if not vno and amt == 0:
            return
        pk = rowd.get('PK', '')
        name = rowd.get('科目', '') or rowd.get('总账科目', '')
        summary = rowd.get('文本', '') or rowd.get('记', '')
        vclass = classify_voucher(pk, str(summary or '') + str(name or ''))
        dr, cr = (amt, 0.0) if amt >= 0 else (0.0, -amt)
        out.append({
            'y': str(rowd.get('财年', '2026') or ''),
            'e': str(rowd.get('公司', '') or '').strip(),
            'date': _parse_date(rowd.get('凭证日期', '')),
            'vno': str(vno or ''),
            'vtype': str(rowd.get('类型', '') or ''),
            'name': str(name or ''),
            'dr': dr, 'cr': cr,
            'summary': str(summary or ''),
            'opp': '',
            'cust': str(rowd.get('客户', '') or ''),
            'supp': str(rowd.get('供应商', '') or ''),
            'pk': str(pk or ''),
            'vclass': vclass,
        })

    _COL_NAMES = ('凭证编号', '凭证日期', '本币金额', '科目', '文本', '客户',
                  '供应商', 'PK', '公司', '财年', '类型', '总账科目')

    if ext == '.txt':
        # 3400 txt：GBK 制表符分隔；表头=含"凭证编号"的行
        colmap = {}
        with open(fp, encoding='gbk', errors='replace') as fh:
            for line in fh:
                line = line.rstrip('\n').rstrip('\r')
                if not line.strip():
                    continue
                cells = line.split('\t')
                if not colmap:
                    if '凭证编号' not in line:
                        continue
                    for i, c in enumerate(cells):
                        c = c.strip()
                        if c in _COL_NAMES:
                            colmap[c] = i
                    if '凭证编号' not in colmap:
                        colmap = {}
                    continue
                rowd = {c: (cells[i] if i < len(cells) else '') for c, i in colmap.items()}
                _emit(rowd)
    else:
        # 3300 xlsx：zipfile 流式 + sharedStrings
        try:
            z = zipfile.ZipFile(fp)
        except Exception:
            return out
        try:
            ss_xml = z.read('xl/sharedStrings.xml').decode('utf-8', errors='replace')
        except KeyError:
            ss_xml = ''
        items = re.findall(r'<si>(.*?)</si>', ss_xml)
        ss = [''.join(re.findall(r'<t[^>]*>(.*?)</t>', si)) for si in items]
        ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        # ⚡ 列字母覆盖 A-AN（40 列）——3300 的凭证日期在 AG/本币金额在 AH，超出单字母 A-Z
        _letters_all = []
        for i in range(1, 41):
            s = ''
            n = i
            while n:
                n, r = divmod(n - 1, 26)
                s = chr(65 + r) + s
            _letters_all.append(s)
        letters = _letters_all
        colmap = {}
        try:
            with z.open('xl/worksheets/sheet1.xml') as fh:
                for ev, el in ET.iterparse(fh, events=('end',)):
                    if el.tag != ns + 'row':
                        continue
                    cells = {}
                    for c in el:
                        ref = c.get('r')
                        if not ref:
                            continue
                        col = re.match(r'([A-Z]+)', ref).group(1)
                        t = c.get('t')
                        v = c.find(ns + 'v')
                        val = v.text if v is not None else ''
                        if t == 's' and val:
                            try:
                                val = ss[int(val)]
                            except (ValueError, IndexError):
                                pass
                        cells[col] = val
                    if not colmap:
                        hdr = [str(cells.get(l, '')) for l in letters]
                        if '凭证编号' not in ' '.join(hdr):
                            el.clear()
                            continue
                        for l in letters:
                            c = str(cells.get(l, '')).strip()
                            if c in _COL_NAMES:
                                colmap[c] = l
                        el.clear()
                        continue
                    rowd = {c: cells.get(l, '') for c, l in colmap.items()}
                    _emit(rowd)
                    el.clear()
        except Exception:
            pass
    return out


def read_sap_gl(data_dir, entities, year='2026', keep_virtual=True):
    """读取 SAP 序时账 → [{'y','e','date','vno','vtype','name','dr','cr','summary','opp','cust','supp','pk','vclass'}]。
    读取即聚合：按文件流式读入，只保留必要列；keep_virtual=False 时剔除虚增虚减行。

    ah-sap 序时账 39 列（实测 2026-08-07）：
      0分配 1凭证编号 2部门 3类型 4PK 5文本 6公司 7项 8财年 9年度/月份 10期间 11科目
      12客户 13参照 14利润中心 15合同号 16合同类型 17WBS 18FA 19功能范围 20成本要素
      21供应商 22WBS 23RCd 24订单 25物料 26项目 27采购凭证 28凭证日期 29本币金额
      30本币 31过账日期 32凭证货币金额 33货币 34实际汇率 35货币 36总帐金额 37付款日期 38支付货币金额
    """
    import openpyxl
    rows = []
    layout = detect_sap_layout(data_dir)
    gl_dir = os.path.join(data_dir, '序时账') if layout == 'ah-sap' else data_dir
    files = []
    if layout == 'ah-sap':
        for f in sorted(os.listdir(gl_dir)):
            if f.lower().endswith(('.xlsx', '.xls')):
                files.append(os.path.join(gl_dir, f))
    elif layout in ('300', '3300'):
        # ⚡⚡ 2026-08-16 3400/3300 适配：<前缀>序时账*.xlsx/txt（按列名定位读取）
        for f in sorted(os.listdir(gl_dir)):
            if '序时账' in f and f.lower().endswith(('.xlsx', '.xls', '.txt')):
                files.append(os.path.join(gl_dir, f))
        for fp in files:
            out = _read_sap_gl_colnamed(fp)
            if not keep_virtual:
                out = [r for r in out if r['vclass'] == 'normal']
            rows.extend(out)
        return rows
    else:
        for f in sorted(os.listdir(gl_dir)):
            if re.match(r'^\d+月\.XLSX$', f, re.I):
                files.append(os.path.join(gl_dir, f))
    for fp in files:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 30 or r[11] is None:
                continue
            comp = str(r[6]).strip() if r[6] else ''
            if not re.match(r'^\d{4}$', comp):
                continue
            pk = r[4]
            vclass = classify_voucher(pk, str(r[5] or '') + str(r[11] or ''))
            if not keep_virtual and vclass != 'normal':
                continue
            amt = float(r[29] or 0)  # 本币金额（正=借 负=贷）
            dr, cr = (amt, 0.0) if amt >= 0 else (0.0, -amt)
            rows.append({
                'y': str(r[8] or year),
                'e': comp,
                'date': str(r[28] or '')[:10],
                'vno': str(r[1] or ''),
                'vtype': str(r[3] or ''),
                'name': str(r[11] or ''),
                'dr': dr, 'cr': cr,
                'summary': str(r[5] or ''),
                'opp': '',
                'cust': str(r[12] or ''), 'supp': str(r[21] or ''),
                'pk': str(pk or ''), 'vclass': vclass,
            })
        wb.close()
    return rows


# ============================================================
# 四、应收/应付 → aux（往来明细）
# ============================================================
def read_sap_aux(data_dir, kind='应收'):
    """读取 SAP 应收/应付目录 → [(客户/供应商名, 科目名, 期初, 借发, 贷发, 期末)]。
    表头：编号/客户编码/客户名称/科目编号/科目名称/合同号/期初余额/本期借方/本期贷方/期末余额
    """
    out = []
    layout = detect_sap_layout(data_dir)
    if layout != 'ah-sap':
        return out
    d = os.path.join(data_dir, kind)
    if not os.path.isdir(d):
        return out
    import openpyxl
    for f in sorted(os.listdir(d)):
        if not f.lower().endswith(('.xlsx', '.xls')):
            continue
        fp = os.path.join(d, f)
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 10 or r[2] is None:
                continue
            out.append((str(r[2]).strip(), str(r[4] or '').strip(),
                        float(r[6] or 0), float(r[7] or 0), float(r[8] or 0), float(r[9] or 0)))
        wb.close()
    return out


# ============================================================
# 五、企业报表核对基准（利润表/资产负债表 .xls）
# ============================================================
def read_sap_report(data_dir, kind='利润表', comp='1010'):
    """读取企业报表（.xls 会企格式）→ {项目: {本月/累计/上年同期 或 期末/年初}}。
    kind='利润表'：项目/行次/本月金额/累计金额/上年同期累计
    kind='资产负债表'：资产项目/行次/期末余额/年初余额
    """
    layout = detect_sap_layout(data_dir)
    if layout != 'ah-sap':
        return {}
    d = os.path.join(data_dir, kind)
    fp = os.path.join(d, f'{comp}.xls')
    if not os.path.exists(fp):
        return {}
    try:
        import xlrd
    except Exception:
        return {}
    try:
        wb = xlrd.open_workbook(fp)
        ws = wb.sheet_by_index(0)
    except Exception:
        return {}
    out = {}
    if kind == '利润表':
        for i in range(4, ws.nrows):
            proj = str(ws.cell_value(i, 0)).strip()
            if not proj or proj == '项目':
                continue
            out[proj] = {'month': ws.cell_value(i, 2), 'cum': ws.cell_value(i, 3),
                         'prev': ws.cell_value(i, 4)}
    else:
        # 资产负债表：左右分栏（资产在左 0-3 列，负债权益在右 4-7 列）
        for i in range(4, ws.nrows):
            for col0, col1, col2 in ((0, 2, 3), (4, 6, 7)):
                proj = str(ws.cell_value(i, col0)).strip()
                if not proj or proj == '资产' or proj == '负债和所有者权益':
                    continue
                if proj.endswith('：') or proj.endswith(':'):
                    continue  # 区块标题（流动资产：/流动负债：/合计）
                out[proj] = {'end': ws.cell_value(i, col1), 'begin': ws.cell_value(i, col2)}
    return out


# ============================================================
# 六、兼容入口（供现有程序调用）
# ============================================================
def is_sap_dir(data_dir):
    return detect_sap_layout(data_dir) is not None


if __name__ == '__main__':
    import sys
    d = sys.argv[1] if len(sys.argv) > 1 else P.YY
    print('布局:', detect_sap_layout(d))
    ents = discover_sap_entities(d)
    print('实体数:', len(ents))
    for code in sorted(ents)[:8]:
        info = ents[code]
        for yy, paths in info.items():
            km = os.path.basename(paths.get('km') or '-')
            gl = len(paths.get('gl') or [])
            print('  %s %s: km=%s gl=%d文件' % (code, yy, km, gl))
