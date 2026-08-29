# -*- coding: utf-8 -*-
"""sap_adapter.py —— SAP 数据源适配层（治本架构：一套生成器 + 数据适配层，2026-08-09）。

用户原则：无论 SAP 还是 U8 只是数据来源不一样，形成底稿用同一套生成器。
本层把 SAP（ah-sap/3300）原始数据统一成 U8 audit_common 中间结构，使 12 个 *_detail.py
生成器可直接跑 SAP 账套。SAP 特殊逻辑（区间文件/内部码→名称映射/期初反推）封装在本层，
生成器对数据源透明。

接口对齐（与 audit_common 同名同签名）：
    is_sap(data_dir)                    → bool
    discover_entities(data_dir)         → {E: {y: {'km': path, 'gl': path}}}
    read_tb_full(data_dir, entities)    → {(e, code, name, y): {qc,jf,df,qm,level}}
    read_gl_rows(data_dir, entities)    → [{e,y,date,vtype,vno,opp,summary,name,dr,cr}]

⚡ 内存策略：SAP 88 家全量 GL 在内存会 OOM（1010 一家 67 万行）。SAP 场景必须【逐主体】
   调用（entities 只含 1 个主体），read_gl_rows 内部按单主体读落盘缓存 + 映射名称。
   这由 _sap_batch 改造后按主体循环驱动 U8 生成器实现。

⚡ GL 名称映射（核心）：SAP 序时账无科目名称列（col11=内部码），从 TB 建立
   (comp, code)→名称 映射：精确匹配 → 前8/6/4位前缀回退（实测 1010 覆盖率 100%）。
"""
import os
import sys
import re
import collections

import sap_reader as SR
import sap_common as C

# 进程级缓存
_CACHE_ENT = {}
_CACHE_TB = {}
_CACHE_GL = {}
_CACHE_AUX = {}   # (comp, kind) -> read_aux_balance 结果（往来 8 科目循环复用）
_CACHE_GL_NET = {}  # (comp, pfx, want_asset) -> read_gl_net 净额化结果（7 往来科目循环复用）
_GL_INDEX_BY_CODE4 = None   # 2026-08-25 集团模式 code4 前缀索引（_root, {code4: [rows]}）
_CACHE_CODE2NAME = {}
_CACHE_CUST_NAME = {}


def _load_cust_name_map(comp):
    """应收/{comp}.xlsx → {客户编码: 客户名称}（收入↔应收↔客户配对用，按 comp 缓存）。
    ⚡ 2026-08-10：SAP 序时账收入行客户是编码（col12），名称仅在应收表 col3。"""
    if not comp:
        return {}
    key = ('cust_name', comp)
    if key in _CACHE_CUST_NAME:
        return _CACHE_CUST_NAME[key]
    mp = {}
    p = os.path.join(_DATA_ROOT or '', '应收', f'{comp}.xlsx')
    try:
        import openpyxl
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if r and len(r) >= 3 and r[1] is not None:
                mp.setdefault(str(r[1]).strip(), str(r[2] or '').strip())
        wb.close()
    except Exception:
        pass
    _CACHE_CUST_NAME[key] = mp
    return mp


def is_sap(data_dir):
    """支持任意路径：目录或文件路径都行（生成器传的是单文件路径）。
    ⚡ 2026-08-10 同 discover_entities：忽略 data_dir 参数、统一用 _DATA_ROOT——
       run_u8_on_sap 把生成器 data_dir 传独立工作目录 _work_{comp}（非 SAP 结构），
       is_sap(_work) 会误判 False → 生成器走 U8 本地分支 → os.path.exists(list) 报错。
    ⚡⚡ 2026-08-24 形态判定统一委托 ledger_backend（detect_sap_layout 内容探测 +
       account_profiles 声明覆盖单点）——各 builder 直接调本函数即获得统一判定，
       不再各判各的。"""
    _d = _DATA_ROOT or data_dir
    if _d is None:
        return False
    if os.path.isfile(_d):
        _d = os.path.dirname(_d)
    from ledger_backend import get_backend
    try:
        return get_backend(_d).layout != 'u8'
    except Exception:
        return False


# ============================================================
# 当前驱动主体（生成器本地读取函数在 SAP 模式下用此取数据）
# ============================================================
_current_comp = None
# ⚡ 2026-08-16 集团主体过滤集（GROUP_MODE 下由 run_u8_on_sap 注入；discover_entities 一处过滤）
_GROUP_COMPS = None


def set_comp(comp):
    """驱动脚本逐主体循环时设置当前主体；生成器 SAP 分支用它取 TB/GL/aux。"""
    global _current_comp
    _current_comp = comp


def current_comp():
    return _current_comp


def _scope_comp(path_or_comp):
    """解析调用方传入参数 → 主体代码：显式代码 > current_comp(单主体) > 路径提取(集团兜底)。
    ⚡ 2026-08-10 集团模式（--group）：生成器循环多主体时传各主体 km/gl 路径（{comp}/xxx 或
    {comp}科目余额表...），current_comp=None —— 原实现忽略参数硬用 current_comp → 集团下全空
    （revenue/tax 审定表全 0）。
    ⚡ 2026-08-13 修复（#738 根因治本）：SAP 科目余额表为【区间合并文件】（2250~2300.xlsx
    一文件含多主体），生成器传的文件路径按文件名 `re.match(r'^(\\d{4})', basename)` 会提取到
    【区间起点】主体（2250/6010）而非当前主体 → 串户（2260 底稿=2250 数据）/缺失（6010 不在
    TB 公司集时读空→SKIP）。单主体驱动（current_comp 已设置）时【优先 current_comp】——
    路径解析仅作为集团模式（current_comp=None）的兜底。"""
    if isinstance(path_or_comp, str) and re.match(r'^\d{4}$', path_or_comp):
        return path_or_comp
    if _current_comp:
        return _current_comp
    if isinstance(path_or_comp, str):
        # ⚡⚡ 2026-08-24 修复（ADF 集团 longterm 全 SKIP 根因）：路径含年份+主体目录
        #   （d:...\2026\3300\3300科目余额表1-7月.xlsx），原 `[\\/](\d{4})[\\/]` 只匹配
        #   \2026\（年份，主体段 3300 后是汉字非分隔符）→ read_km 读 2026 主体 → 空 →
        #   固定资产/无形资产等全 SKIP。改为：【文件名开头 4 位数字】优先（主体前缀，
        #   3700科目余额表20260731.XLSX → 3700，天然排除尾部日期 0731）；无则取路径
        #   中最后的数字目录段（ADF 2026\3300\ 主体目录在年份后）。
        _base = os.path.basename(path_or_comp)
        m = re.match(r'^(\d{4})(?=\D|$)', _base)
        if m:
            return m.group(1)
        _digs = re.findall(r'(\d{4})(?=\D|$)', path_or_comp.replace('\\', '/'))
        if _digs:
            return _digs[-1]
        return _current_comp


def tb_entries(comp=None):
    """当前/指定主体的 TB 子目级 → {code: {'name','qc','jf','df','qm','level'}}（U8 read_km 结构）。"""
    comp = _scope_comp(comp) or comp or _current_comp
    if not comp or not _DATA_ROOT:
        return {}
    tb = read_tb_full(_DATA_ROOT, None)
    out = {}
    for (e, c, n, y), v in tb.items():
        if e == comp:
            out[c] = dict(v, name=n)
    return out


def gl_rows(comp=None):
    """当前/指定主体的 GL 标准行（U8 read_gl_rows 结构）。
    ⚡⚡ 2026-08-24 修复（AH 职工薪酬勾稽核对缺失根因）：集团模式（current_comp=None 且
    _GROUP_COMPS 已注入）下 payroll 等生成器调无参 _adapter.read_gl() → 原返回空 →
    计提/发放全 0。改为集团模式返回【全量 GL】（_GROUP_COMPS 过滤后主体），与
    read_gl_net 集团分支一致。"""
    comp = _scope_comp(comp) or comp or _current_comp
    if not comp or not _DATA_ROOT:
        # ⚡⚡ 集团模式兜底：current_comp=None + 有 _GROUP_COMPS → 全量 GL
        if not _current_comp and getattr(sys.modules[__name__], '_GROUP_COMPS', None):
            ents = discover_entities(_DATA_ROOT)
            return read_gl_rows(_DATA_ROOT, ents)
        return []
    ents = {comp: discover_entities(_DATA_ROOT).get(comp, {})}
    return read_gl_rows(_DATA_ROOT, ents)


def read_km(path_or_comp=None):
    """SAP 单主体 TB → U8 read_km 结构 {code: {name, odir, opening, debit, credit, fdir, closing, level}}。
    ⚡ 忽略传入参数（生成器传的是 SAP 文件路径/list），数据一律从 current_comp 取。
    ⚡ SAP TB 无父级行（只有末级子目，如 6001010000）→ 需生成【一级聚合行】
       （code[:4]，name=名称首段如 '主营业务收入'），否则 U8 生成器按一级名精确匹配失败。
       SAP 名称分隔符 '-'（'主营业务收入-制造收入' → 一级 '主营业务收入'）。"""
    comp = _scope_comp(path_or_comp)
    entries = tb_entries(comp)
    out = {}
    # 一级聚合（code[:4]）——U8 生成器一级名匹配依赖（_detect_codes 找 level==1 精确名）
    l1 = {}
    for code, v in entries.items():
        c4 = code[:4]
        if c4 not in l1:
            nm = str(v.get('name', '')).split('-')[0].split('－')[0].strip() or c4
            l1[c4] = {'name': nm, 'odir': '借', 'opening': 0.0, 'debit': 0.0,
                      'credit': 0.0, 'fdir': '借', 'closing': 0.0, 'level': 1}
        l1[c4]['opening'] += float(v.get('qc') or 0.0)
        l1[c4]['debit'] += float(v.get('jf') or 0.0)
        l1[c4]['credit'] += float(v.get('df') or 0.0)
        l1[c4]['closing'] += float(v.get('qm') or 0.0)
    for c4, v in l1.items():
        v['odir'] = '借' if v['opening'] >= 0 else '贷'
        v['fdir'] = '借' if v['closing'] >= 0 else '贷'
        # ⚡ 2026-08-22 收敛：统一 qc/jf/df/qm（标准行结构规范；双写向后兼容）
        v['qc'] = v['opening']; v['jf'] = v['debit']; v['df'] = v['credit']; v['qm'] = v['closing']
        out[c4] = v
    # 末级行原样（明细/分月用）
    for code, v in entries.items():
        if code not in out:
            qc = float(v.get('qc') or 0.0)
            qm = float(v.get('qm') or 0.0)
            out[code] = {
                'name': v.get('name', ''),
                'odir': '借' if qc >= 0 else '贷',
                'opening': qc,
                'debit': float(v.get('jf') or 0.0),
                'credit': float(v.get('df') or 0.0),
                'fdir': '借' if qm >= 0 else '贷',
                'closing': qm,
                # ⚡ 2026-08-22 收敛：统一 qc/jf/df/qm
                'qc': qc, 'jf': float(v.get('jf') or 0.0),
                'df': float(v.get('df') or 0.0), 'qm': qm,
                'level': v.get('level', 4),
            }
    return out


def _infer_opp(rows):
    """凭证级对方科目推断（SAP GL 无对方科目列——铁律：凭证级配对）。

    U8 生成器的剔除/分类逻辑依赖 cp/opp 字段：
      - longterm 内部互转剔除（_is_intragroup_cost_transfer 检查 cp 含同组成本科目）
      - equity 结转凭证抽查剔除（opp 含 '结转'/'本年利润' 关键词）
      - inventory 领用/结转分类（cp 判 5001/6402 等对方科目）
      - expense 制造费用勾稽（cp）
    若 SAP 模式不补此字段（cp=''），以上剔除全部静默失效 → 虚增虚减回归。
    规则：同凭证(vno)内【其他科目名】，按金额绝对值降序去重，逗号连接。"""
    grp = {}
    for r in rows:
        grp.setdefault((r.get('y'), r.get('vno')), []).append(r)
    for _key, rs in grp.items():
        names = {}
        for r in rs:
            nm = str(r.get('name') or '').strip()
            if not nm:
                continue
            amt = abs(float(r.get('dr') or 0.0)) + abs(float(r.get('cr') or 0.0))
            if nm not in names or amt > names[nm]:
                names[nm] = amt
        ordered = sorted(names, key=lambda x: -names[x])
        for r in rs:
            self_nm = str(r.get('name') or '').strip()
            opp = ','.join(n for n in ordered if n != self_nm)
            r['opp'] = opp
            r['cp'] = opp
    return rows


def read_gl(path_or_comp=None):
    """SAP 单主体 GL → U8 read_gl 结构 [{name,date,year,month,vtype,vno,cp,aux,debit,credit}]。
    ⚡ 主体解析：_scope_comp（路径提取 > 直接代码 > current_comp；集团模式生成器传单主体 gl 路径）。
    ⚡ cp/opp=凭证级对方科目推断（_infer_opp），U8 剔除/分类逻辑依赖。
    ⚡⚡ 2026-08-25 性能：结果按 comp 缓存（集团模式 comp=None 全量 455 万行——88 主体各调一次，
        原每次全量构造 20 字段 dict 约 5-12s → 缓存后 1 次构造、后续命中）。调用方只读字段。"""
    comp = _scope_comp(path_or_comp)
    _ck = ('gl', comp)
    if _ck in _CACHE_GL:
        return list(_CACHE_GL[_ck])
    rows = gl_rows(comp)
    out = []
    for r in rows:
        vt = r.get('vtype', '')
        vn = r.get('vno', '')
        out.append({
            'name': r.get('name', ''),
            'code': r.get('code', ''),    # 内部码（费用功能范围分类/660006 特例用）
            'fr': r.get('fr', ''),        # 功能范围（col19：6601销售/6602管理/5101制造/5301研发）
            'date': r.get('date', ''),
            'year': r.get('y', ''),
            'month': r.get('month', 0),
            'vtype': vt,
            'vno': vn,
            'voucher': '%s-%s' % (vt, vn) if (vt or vn) else '',
            'cp': r.get('cp', ''),        # 凭证级对方科目推断（防剔除失效）
            'opp': r.get('opp', ''),      # 同上（equity 用 opp 字段名）
            # ⚡ 2026-08-10 修复：SAP 序时账有客户(col12)/供应商(col21)列，此前 aux=''
            # → revenue rev_cust 全空 → 收入客户明细/前十大客户不生成。逐笔行直接映射。
            'aux': str(r.get('cust') or '').strip() or str(r.get('supp') or '').strip(),
            'sm': r.get('summary', ''),   # 摘要（longterm _read_gl_fa 结构需要）
            'summary': r.get('summary', ''),
            # ⚡ 2026-08-10 指令号透传（GL 39 列 col17 WBS/col24 订单/col25 物料/col26 项目/col27 采购凭证）：
            'wbs': r.get('wbs', ''), 'ord': r.get('ord', ''), 'mat': r.get('mat', ''),
            'proj': r.get('proj', ''), 'po': r.get('po', ''),
            # ⚡ 2026-08-22 收敛：统一 dr/cr（标准行结构规范）；debit/credit 双写向后兼容
            'dr': r.get('dr', 0.0),
            'cr': r.get('cr', 0.0),
            'debit': r.get('dr', 0.0),
            'credit': r.get('cr', 0.0),
        })
    _CACHE_GL[_ck] = out
    return list(out)


def aux_entries(comp=None, name_contains=None, kind='all'):
    """当前/指定主体的 aux 行（read_aux_balance 包装）。
    ⚡ 2026-08-10 集团模式：comp=None 且 current_comp=None → 遍历全部主体，行带 entity。"""
    if comp is None:
        comp = _current_comp
    if comp:
        return read_aux_balance(_DATA_ROOT, comp, name_contains=name_contains, kind=kind)
    out = []
    for c in discover_entities(_DATA_ROOT):
        out.extend(read_aux_balance(_DATA_ROOT, c, name_contains=name_contains, kind=kind))
    return out


def sap_inv_level2(comp=None):
    """SAP 存货 2 级科目（按名称）：{科目码(adapter.read_km 中存在的码): 2级名}。
    SAP 科目体系与 U8 不同（1406=库存产品父级含库存商品/产成品/半成品/在制品）——
    不能按 U8 的 1403/1405/1406 code 聚合。按名称第二段取 2 级名（'库存产品-库存商品'→'库存商品'）。
    ⚡ key 必须是 adapter.read_km 返回中存在的码（4位一级 或 10位末级）——inventory 审定表 km.get(code) 精确匹配。"""
    comp = comp or _current_comp
    km = read_km(comp)
    out = {}
    for code, v in km.items():
        if v.get('level') == 1 or not code.startswith(('1401', '1403', '1404', '1405', '1406', '1407', '1408', '1411', '1412')):
            continue
        nm = str(v.get('name') or '')
        parts = nm.split('-')
        l2name = parts[1].strip() if len(parts) > 1 else parts[0].strip()
        if l2name and code not in out:
            out[code] = l2name
    return out


_DATA_ROOT = None


def set_root(data_dir):
    global _DATA_ROOT
    _DATA_ROOT = data_dir


# ============================================================
# 一、实体发现（对齐 audit_common.discover_entities）
# ============================================================
def discover_entities(data_dir):
    """适配 SAP 实体发现 → U8 结构 {E: {y: {'km': path, 'gl': path}}}。
    ⚡ SAP 的 gl 是多文件 list（逐月序时账），U8 期望单 path——适配层返回原 list，
    由本层 read_gl_rows 内部处理（生成器不得直接 os.path.exists(paths['gl'])）。
    ⚡ 2026-08-09 统一：current_comp 已设置（单主体驱动）→ 返回单主体；未设置（全集团）→ 全量。
    生成器本地 discover 一律收敛为 audit_common.discover_entities（patch 后=本函数），
    单主体过滤在此一处生效，删除各生成器 discover 的 is_sap 分支。
    ⚡ 2026-08-10 修复：**忽略 data_dir 参数、统一用 _DATA_ROOT**——run_u8_on_sap 为防多进程
    并行同名中间文件写坏，把生成器 data_dir 传为独立工作目录 _work_{comp}；若本函数用传入
    目录扫数据 → 换目录后 0 主体。生成器对数据源透明（读数据永远走 _DATA_ROOT）。"""
    _d = _DATA_ROOT or data_dir
    if _d in _CACHE_ENT:
        out = _CACHE_ENT[_d]
    else:
        ents = SR.discover_sap_entities(_d)
        if ents:
            out = {}
            for code, yd in ents.items():
                for yy, paths in yd.items():
                    out.setdefault(code, {})[str(yy)] = {'km': paths.get('km'), 'gl': paths.get('gl'),
                                                         'aux_ar': paths.get('aux_ar'),
                                                         'aux_ap': paths.get('aux_ap')}
        else:
            # ⚡⚡ 2026-08-29 P0 修复：U8 多主体账套（XBJ 等）SAP 布局探测返回 None →
            #   SR 空 → discover 恒 0 主体 → run_u8_on_sap 对 XBJ 全量跑不了（OK 0）。
            #   回退 audit_common 的 U8 通用发现（能识别 XBJ 201 个长项目名主体），
            #   结构转适配层 {code: {year: {km, gl, aux}}}。
            out = {}
            try:
                from audit_common import discover_entities as _de_u8
                _u8 = _de_u8(_d)
                for _c, _yd in _u8.items():
                    for _yy, _bun in _yd.items():
                        if not isinstance(_bun, dict):
                            continue
                        _km = _bun.get('km')
                        _gl = _bun.get('gl')
                        out.setdefault(str(_c), {})[str(_yy)] = {
                            'km': _km,
                            # ⚡⚡ 2026-08-29 二次修复：U8 生成器（pl/expense 等）对 gl 做
                            #   os.path.exists(g) 需单路径字符串；原转 list 导致 TypeError。
                            #   SAP 分支才用 list（read_gl_rows 内部展开）。
                            'gl': _gl,
                            'aux_ar': _bun.get('aux_ar'),
                            'aux_ap': _bun.get('aux_ap'),
                        }
            except Exception:
                out = {}
        _CACHE_ENT[_d] = out
    if _current_comp and _current_comp in out:
        return {_current_comp: out[_current_comp]}
    # ⚡⚡ 2026-08-16 集团主体过滤（#744 对平抽查发现根因）：集团模式（GROUP_MODE）下
    #   run_u8_on_sap 注入 _GROUP_COMPS = 本集团主体集；所有生成器（revenue/current_account/
    #   equity 等）内部 discover 全量 88 家——在此一处过滤，集团底稿才真正按集团拆分，
    #   否则集团A/B/C 底稿内容相同且与集团试算表对不平。
    _gc = _GROUP_COMPS
    if _gc:
        out = {c: v for c, v in out.items() if c in _gc}
    return out


# ============================================================
# 二、TB（对齐 audit_common.read_tb_full）
# ============================================================
def read_tb_full(data_dir, entities=None, year='2026'):
    """SAP 科目余额表 → {(e, code, name, y): {qc,jf,df,qm,level}}。
    read_sap_tb 的 key 结构（comp, code, name, year）与 U8 read_tb_full 完全一致，直接复用。
    ⚡ code=10 位内部码（U8 是 4-6 位层级码）——U8 生成器 code 前缀匹配天然兼容（更精确）。
    ⚡ read_sap_tb 忽略 entities 参数全量读 → 必须按 entities 过滤（防单主体驱动时
       审定表混入全集团主体行，而 GL 只有单主体 → 明细/审定不一致）。
    ⚡ 2026-08-10 同 discover_entities：忽略 data_dir 参数、统一用 _DATA_ROOT。"""
    _d = _DATA_ROOT or data_dir
    key = (_d, year)
    if key in _CACHE_TB:
        tb = _CACHE_TB[key]
    else:
        tb = SR.read_sap_tb(_d, entities, year=year)
        _CACHE_TB[key] = tb
        _build_code2name(tb)
    if entities is not None:
        comps = set(entities)
        tb = {k: v for k, v in tb.items() if k[0] in comps}
    # ⚡⚡ 2026-08-26 还原：read_tb_full(None, None) 语义=全量 88 主体（tax_detail._load_tb_sap
    #   依赖它返回全量后内部按 comp_scope 过滤）。原加 _GROUP_COMPS 过滤破坏该语义 →
    #   1357 应交税费审定表全 0（1020 等主体数据被提前滤掉）。payroll 的集团主体过滤
    #   改在 read_tb_payroll 内部处理（见 payroll_detail）。
    return tb


def _build_code2name(tb):
    """从 TB 建 (comp, code)→名称 精确索引 + 前8/6/4位前缀索引。"""
    exact = {}
    pre8 = collections.defaultdict(list)
    pre6 = collections.defaultdict(list)
    pre4 = collections.defaultdict(list)
    for (e, c, n, y), v in (tb or {}).items():
        if not c or not n:
            continue
        exact[(e, c)] = n
        if len(c) >= 8:
            pre8[(e, c[:8])].append(n)
        if len(c) >= 6:
            pre6[(e, c[:6])].append(n)
        if len(c) >= 4:
            pre4[(e, c[:4])].append(n)
    _CACHE_CODE2NAME[('exact',)] = exact
    _CACHE_CODE2NAME[('pre8',)] = pre8
    _CACHE_CODE2NAME[('pre6',)] = pre6
    _CACHE_CODE2NAME[('pre4',)] = pre4


def resolve_subject_name(comp, code):
    """GL 内部码 → 科目名称：精确 → 前8/6/4位前缀回退（前缀下取第一个，同类科目）。
    未命中返回 None（调用方保留内部码兜底）。"""
    exact = _CACHE_CODE2NAME.get(('exact',)) or {}
    nm = exact.get((comp, code))
    if nm:
        return nm
    for idx_key, cut in (('pre8', 8), ('pre6', 6), ('pre4', 4)):
        idx = _CACHE_CODE2NAME.get((idx_key,)) or {}
        hits = idx.get((comp, code[:cut]))
        if hits:
            return hits[0]
    return None


# ============================================================
# 三、GL（对齐 audit_common.read_gl_rows）
# ============================================================
def _excel_date(v):
    """SAP 序时账凭证日期（col28）是 Excel 序列号（46022.0=2026-01-01）→ 'YYYY-MM-DD'。
    ⚡⚡ 2026-08-23 修复：300 形态 _scan_gl_full 读出的日期列值常为【字符串数字】
       （'46031'，_safx 原样保留），原实现仅 int/float 转换、字符串原样返回 → 年份
       推断成 '4603'（假年份）→ opening 键 (name,'4603') 与 export 年份 '2026' 不匹配
       → 往来明细表空（3300 案例）。字符串数字同样转 Excel 序列号日期。"""
    if v is None:
        return ''
    if isinstance(v, (int, float)):
        try:
            from datetime import datetime, timedelta
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime('%Y-%m-%d')
        except Exception:
            return str(v)
    if isinstance(v, str) and v.strip().isdigit():
        try:
            from datetime import datetime, timedelta
            return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime('%Y-%m-%d')
        except Exception:
            pass
    return str(v)[:10]


def read_gl_rows(data_dir, entities=None, year='2026', exclude_co=False):
    """SAP 序时账 → [{e,y,date,vtype,vno,opp,summary,name,dr,cr,month}]。
    ⚡ 必须逐主体调用（entities 只含 1 主体），内部读单公司落盘缓存 + 内部码→名称映射。
    ⚡ exclude_co=True（2026-08-23）：剔除 CO 结转/分摊凭证（is_co=1，月末成本结转/
       物料重估，GL 绝对发生额虚增而 TB 不含）→ 明细/抽凭口径与 TB 更吻合。默认 False
       不改现有行为。
    ⚡ month=col10 期间转 int（SAP GL 日期是 Excel 序列号，不能用日期解析月份）。
    ⚡ 名称映射依赖 TB（_build_code2name）——若未先行 read_tb_full 则此处兜底构建。
    全量 88 家一次性调用会 OOM——由驱动方（_sap_batch 改造后）按主体循环。
    ⚡ 2026-08-10 同 discover_entities：忽略 data_dir 参数、统一用 _DATA_ROOT。
    ⚡ 2026-08-10 进程级缓存 _CACHE_GL（key=(comp, year)）：同一家庭 13 个生成器
       逐科目驱动都调本函数 → 不缓存则每次全量构造 67 万行 dict + _infer_opp +
       strip_padding，1010 全量重跑 17 分钟（用户痛点"每次都死磕 66 万行"）。
       缓存命中后仅浅拷贝列表（行 dict 共享，调用方只读字段）。"""
    _d = _DATA_ROOT or data_dir
    ents = entities or discover_entities(_d)
    # ⚡ 2026-08-23 exclude_co 影响结果 → 缓存 key 必须区分，否则两模式互相污染
    _ck = (bool(exclude_co), tuple(sorted((e, str(y)) for e, yd in ents.items() for y in yd)))
    if _ck in _CACHE_GL:
        return list(_CACHE_GL[_ck])
    if not _CACHE_CODE2NAME.get(('exact',)):
        read_tb_full(_d, None, year=year)   # 构建名称映射（幂等）
    rows = []
    _is300 = (C.detect_sap_layout(_d) == '300') if hasattr(C, 'detect_sap_layout') else False
    for e, yd in ents.items():
        for yy, paths in yd.items():
            if _is300:
                # ⚡ 2026-08-10 300 形态：_scan_gl_full 已按表头动态映射输出逻辑键行
                # {comp,vno,per,date,code,name,amt,hs,sm}（3700/6100 列序不同，勿用固定列）。
                # 方向铁律（3700 实测）：金额【正=借方、负=贷方】（H 标识行负、S 标识行正，
                # col27 与金额符号相反勿用）——正借负贷与 TB 借/贷完全吻合。
                comp_rows = C.load_comp_gl_disk(_d, e, cols=C._GL_CACHE_COLS_300)
                for r in comp_rows:
                    if exclude_co and r.get('is_co'):
                        continue  # ⚡ 2026-08-23 剔除 CO 结转凭证（多行净额0 + 含704/结转文本）
                    code = str(r.get('code') or '')
                    nm = str(r.get('name') or '') or resolve_subject_name(e, code)
                    amt = float(r.get('amt') or 0)
                    dr, cr = (amt, 0.0) if amt >= 0 else (0.0, -amt)
                    per = str(r.get('per') or '')
                    rows.append({
                        'e': e, 'y': str(yy),
                        'date': _excel_date(r.get('date')),   # 过帐日期（Excel 序列号或字符串）
                        'month': int(per) if per.isdigit() else 0,
                        'vtype': '',
                        'vno': str(r.get('vno') or ''),
                        'opp': '',
                        'summary': str(r.get('sm') or ''),
                        'name': nm or code,
                        'code': code,
                        # ⚡ 2026-08-10 客户/供应商列已由 _scan_gl_full 表头动态定位
                        #   （4 账套列序不同，写死 40/41 全空→往来明细本期借/贷无源）
                        'cust': str(r.get('cust') or '').strip(),
                        'supp': str(r.get('supp') or '').strip(),
                        'fr': '',
                        # ⚡ 2026-08-10 指令号透传（300 FBL3N：订单/项目/物料/采购凭证列；
                        #   wbs 无列置 ''。inventory 按指令号增减变动表数据源）
                        'wbs': '',
                        'ord': str(r.get('ord') or '').strip(),
                        'mat': str(r.get('mat') or '').strip(),
                        'proj': str(r.get('proj') or '').strip(),
                        'po': str(r.get('po') or '').strip(),
                        'dr': dr, 'cr': cr,
                    })
                continue
            comp_rows = C.load_comp_gl_disk(_d, e, cols=C._GL_CACHE_COLS | {17, 24, 25, 26, 27})
            for r in comp_rows:
                code = str(r.get(11) or '')
                nm = resolve_subject_name(e, code)
                amt = float(r.get(29) or 0)
                dr, cr = (amt, 0.0) if amt >= 0 else (0.0, -amt)
                per = str(r.get(10) or '')
                # ⚡ 2026-08-10 期间口径（铁律：GL 与 TB「借方1-7」列一致）：SAP 8/5 过账、期间列=07
                #   的凭证在 1010-7月.xlsx 中（GL 文件按记账期间收纳），而 TB 借方1-7 列按过账日期
                #   1-7 月不含 → 若不过滤，税金及附加差 100,324.65（6500004134）、所得税费用差
                #   4,168,801.22（6500004135）。统一按【凭证日期（col28 Excel 序列号）】∈2026-01~07
                #   过滤，与 TB 口径一致（SAP 8 月未结账、8/5 凭证挂 07 期间属天然差异）。
                # ⚡⚡ 2026-08-10 16:2x 修复：期间判定【禁用 col0】——col0=分配指令号（GL 39 列
                #   铁律），同凭证不同行 col0 混杂『F-GF26.1100.10』（分配指令号）与『20260630』
                #   （日期形态），用 col0 过滤会【误删借方行】→ 专项储备 6500003577/4124 计提凭证
                #   借方『安全投入』行被剔 → 计提核对列支=0、longterm 增加表缺失。统一改 col28。
                # ⚡⚡ 2026-08-26 修复：改按【过账日期 col31】过滤（原 col28=凭证日期）——SAP
                #   「凭证日期=2025年末、过账日期=2026」的冲销凭证（如 6500001378 冲审计调整长期
                #   借款重分类，凭证日期 2025-12-31、过账日期 2026-03-21）凭证日期过滤会被误剔，
                #   TB 按过账日期收纳含它 → 分月明细与 TB 逐月对不上（660302 利息支出 3 月差
                #   300.09 万、1 月差 9.01 万=驻马店 6200000266 冲销行同因）。_GL_CACHE_COLS 已加 col31。
                # ⚡⚡ 2026-08-26 第二层：过账日期过滤会引入「凭证日期=8月初、过账日期倒填7月底」
                #   的 8/5 凭证（如 6500004126/4132/4134 计提7月税，凭证日期 8/4-8/5、过账 7/30-31、
                #   期间07），TB 导出时未过账不含 → 税金及附加 6403 7 月差 100,324.65、所得税 6801
                #   差 4,168,801.22。→ 最终口径【过账日期∈1-7 且 凭证日期≤8月1日】。
                #   ⚡ 8/1 上限由 TB 导出时点反推：VERBR CO 结转（凭证日期 8/1、过账 7/31）TB 含，
                #   计提税（凭证日期 8/4-8/5）TB 不含 → TB 导出时点约 8/1-8/3。
                _dt31 = _excel_date(r.get(31))   # 'YYYY-MM-DD' 或 ''（col31 过账日期）
                _dt28 = _excel_date(r.get(28))   # col28 凭证日期（8/5 凭证甄别）
                _ok0 = False
                if len(_dt31) >= 7 and _dt31[:4].isdigit() and _dt31[5:7].isdigit():
                    _mo_ok = (_dt31[:4] == str(yy) and 1 <= int(_dt31[5:7]) <= 7)
                    if _mo_ok:
                        if len(_dt28) >= 7 and _dt28[:4].isdigit() and _dt28[5:7].isdigit():
                            _y28 = int(_dt28[:4]); _yy_i = int(yy)
                            _ok0 = (_y28 < _yy_i) or (_y28 == _yy_i and _dt28[5:10] <= '08-01')
                if not _ok0:
                    continue
                rows.append({
                    'e': e, 'y': str(yy),
                    'date': _excel_date(r.get(31)),
                    'month': int(per) if per.isdigit() else 0,
                    'vtype': str(r.get(3) or ''),
                    'vno': str(r.get(1) or ''),
                    'opp': '',
                    'summary': str(r.get(5) or ''),
                    'name': nm or code,          # 映射失败保留内部码兜底
                    'code': code,                # 内部码（费用功能范围分类/660006 特例用）
                    'cust': str(r.get(12) or '').strip(),   # 客户（col12，往来款净额化用）
                    'supp': str(r.get(21) or '').strip(),   # 供应商（col21）
                    'fr': str(r.get(19) or '').strip(),     # 功能范围（col19：6601销售/6602管理/5101制造/5301研发）
                    # ⚡ 2026-08-10 指令号列（GL 39 列：col17 WBS/col24 订单/col25 物料/col26 项目/col27 采购凭证）：
                    #   供 inventory 按指令号（WBS/订单）分列期初/增/减/期末明细（用户方法论）。
                    'wbs': str(r.get(17) or '').strip(),    # WBS 元素（P-GF25117.1KJ102-1）
                    'ord': str(r.get(24) or '').strip(),    # 订单号（2808639）
                    'mat': str(r.get(25) or '').strip(),    # 物料号（N2Y00001）
                    'proj': str(r.get(26) or '').strip(),   # 项目（00010）
                    'po': str(r.get(27) or '').strip(),     # 采购凭证（4300012124）
                    'dr': dr, 'cr': cr,
                })
    # ⚡ 2026-08-10 收入客户回填（铁律73：收入↔应收↔客户配对）：
    # SAP 收入确认凭证的客户编码在【同凭证应收账款行】col12，收入行本身 col12 为空
    # （2200 实测 3161 收入行 col12 全空、摘要含"销售人民医院公司"）→ revenue rev_cust
    # 全空 → 收入客户明细/前十大客户不生成。回填后 read_gl 的 aux 即带客户名。
    _INC_KW = ('主营业务收入', '其他业务收入', '营业外收入')
    _cust_by_vno = {}
    for r in rows:
        if str(r.get('name') or '').startswith('应收账款'):
            _cs = str(r.get('cust') or '').strip()
            if _cs:
                _cust_by_vno.setdefault((r.get('e'), r.get('y'), r.get('vno')), _cs)
    for r in rows:
        if str(r.get('name') or '').startswith(_INC_KW) and not str(r.get('cust') or '').strip():
            _cv = _cust_by_vno.get((r.get('e'), r.get('y'), r.get('vno')), '')
            if _cv:
                # 客户编码 → 名称（应收表 col3；映射不到保留编码）
                r['cust'] = _load_cust_name_map(r.get('e')).get(_cv, _cv)
    # ⚡ 凭证级对方科目推断（U8 生成器剔除/分类依赖 cp/opp——SAP 无对方科目列必须补）
    rows = _infer_opp(rows)
    # ⚡ 2026-08-10 跨凭证虚增剔除（用户方法论：同科目同日同额借贷互抵=虚增信号）——
    # 强信号（摘要相同）自动剔除；弱信号（摘要不同）记日志不剔（存货流转/往来清账可能是真实业务）。
    rows = strip_padding(rows)
    _CACHE_GL[_ck] = rows
    return list(rows)


def _row_strip_ident(r):
    """行标识（strip 缓存用）：vno+code+日期+金额+方向。"""
    return (str(r.get('vno') or ''), str(r.get('code') or ''),
            str(r.get('date') or '')[:10],
            round(abs(float(r.get('dr') or 0.0)) + abs(float(r.get('cr') or 0.0)), 2),
            'd' if float(r.get('dr') or 0.0) > 0 else 'c')


def _compute_strip(rows, min_amt=1000000.0):
    """计算应剔除行的标识集合（重估虚拟行 + 强信号摘要相同借贷互抵对）。"""
    if not rows:
        return set()
    from collections import defaultdict
    grp = defaultdict(list)
    for i, r in enumerate(rows):
        amt = abs(float(r.get('dr') or 0.0)) + abs(float(r.get('cr') or 0.0))
        if amt < min_amt:
            continue
        # ⚡⚡ 2026-08-15 修复（6100 实证）：key 用【完整末级 code】而非 code[:4]——
        #   内转凭证『借国开行 110万 + 贷工行 110万』同属 1002 银行存款一级，code[:4]
        #   同组+同日同额+摘要同'内转' → 被误判"虚增"整凭证剔除 → 银行核对账侧漏 3 笔
        #   真实划转（网银两侧都有流水）。跨银行划转是真实业务，仅【同末级科目】借贷互抵
        #   才是冲销/虚增信号。改后：同科目冲销仍匹配剔除，跨账户划转保留。
        key = (r.get('e'), str(r.get('code') or ''),
               str(r.get('date') or '')[:10], round(amt, 2))
        grp[key].append(i)
    drop = set()
    for key, idxs in grp.items():
        drs = [i for i in idxs if float(rows[i].get('dr') or 0.0) > 0.005]
        crs = [i for i in idxs if float(rows[i].get('cr') or 0.0) > 0.005]
        n = min(len(drs), len(crs))
        if n == 0:
            continue
        _is_reval = any('评估' in str(rows[i].get('summary') or '')
                        or '重估' in str(rows[i].get('summary') or '') for i in idxs)
        if _is_reval:
            for rd in drs[:n]:
                drop.add(_row_strip_ident(rows[rd]))
            for rc in crs[:n]:
                drop.add(_row_strip_ident(rows[rc]))
            continue
        matched_dr, matched_cr = set(), set()
        for rd in drs[:n]:
            for rc in crs[:n]:
                sm_d = str(rows[rd].get('summary') or '').strip()
                sm_c = str(rows[rc].get('summary') or '').strip()
                if sm_d and sm_d == sm_c and rd not in matched_dr and rc not in matched_cr:
                    drop.add(_row_strip_ident(rows[rd]))
                    drop.add(_row_strip_ident(rows[rc]))
                    matched_dr.add(rd); matched_cr.add(rc)
    return drop


_CACHE_STRIP = {}


def strip_padding(rows, min_amt=1000000.0):
    """跨凭证同日同额借贷互抵剔除（2026-08-10 用户方法论）。
    只剔【强信号】= 借贷行摘要相同（冲销/重分类虚增）+ 重估虚拟行（摘要含 评估/重估）；
    弱信号（同日同额摘要不同，存货流转/往来清账）不剔——可能是真实业务。
    性能：同主体 strip 结果缓存 _CACHE_STRIP（13 生成器多次 read_gl_rows 只算一次）。"""
    if not rows:
        return rows
    ents = {r.get('e') for r in rows}
    drop = set()
    for e in ents:
        k = ('strip', e)
        if k not in _CACHE_STRIP:
            _CACHE_STRIP[k] = _compute_strip([r for r in rows if r.get('e') == e], min_amt)
        drop |= _CACHE_STRIP[k]
    if not drop:
        return rows
    return [r for r in rows if _row_strip_ident(r) not in drop]





def _is_layout_300():
    """当前 _DATA_ROOT 是否为 300 形态（每账套一文件夹）。"""
    try:
        return C.detect_sap_layout(_DATA_ROOT or '') == '300'
    except Exception:
        return False


def _aux_date_suffix(dp, kw):
    """在账套文件夹 dp 内找含 kw（客户余额表/供应商余额表）的文件，返回最新日期后缀
    （如 '20260731'）；无匹配返回 ''（默认 .xlsx）。"""
    best = ''
    try:
        for f in os.listdir(dp):
            if kw in f and f.lower().endswith(('.xlsx', '.xls')):
                m = re.search(rf'{kw}(\d{{8}})?', f)
                date = m.group(1) if m and m.group(1) else ''
                if date >= best:
                    best = date
    except OSError:
        pass
    return best


def read_aux_balance(data_dir, comp, name_contains=None, kind='all'):
    """SAP 应收/应付表（客户/供应商行项目）→ 行列表。
    应收/{comp}.xlsx 列（1-based）：col2客户编码 col3客户名称 col4科目编号 col5科目名称
    col6合同号 col7期初 col8本期借 col9本期贷 col10期末。
    应付表（铁律74）余额列贷正借负与 TB 反号 → 应付侧 qc/qm 取负还原记账符号（jf/df 保持正发生额）。
    kind: 'ar' 应收 / 'ap' 应付 / 'all' 两者。返回 [{'entity','cust','subject','contract',
    'open','debit','credit','close'}]（金额=1 行全量，由调用方按需净额/聚合）。"""
    import openpyxl as _oxl
    out = []
    root = _DATA_ROOT or data_dir
    # ⚡ 2026-08-10 300 形态：往来余额表=各账套文件夹内『客户余额表/供应商余额表』（账龄式）
    #   客户表（应收侧+其他应收等）：col0客商编码/col1客商描述/col18会计科目(10位)/
    #   col19科目描述/col22未到期余额/col23账龄金额（金额带符号：客户贷余=负→应收侧反号？）
    #   供应商表（应付侧）：col17会计科目/col18科目描述/col19币种/col20账龄金额/col21未到期余额
    if _is_layout_300():
        # ⚡ 2026-08-10 300 形态：优先用 discover 登记的 aux 实际路径（3700=客户余额表{date}、
        #   3900=客户余额{date}/供应商{date}余额，命名不同）——不重新拼文件名。
        _ents = discover_entities(_DATA_ROOT or data_dir)
        _bun = _ents.get(comp, {})
        _paths = next(iter(_bun.values()), {}) if _bun else {}
        _dp = os.path.join(root, comp)
        dirs = []
        _aux_ar = _paths.get('aux_ar')
        _aux_ap = _paths.get('aux_ap')
        if kind in ('ar', 'all'):
            _f = _aux_ar or os.path.join(_dp, f'{comp}客户余额{_aux_date_suffix(_dp, "客户余额")}.xlsx')
            if os.path.exists(_f):
                dirs.append((_f, 1.0))
        if kind in ('ap', 'all'):
            _f = _aux_ap or os.path.join(_dp, f'{comp}供应商{_aux_date_suffix(_dp, "供应商")}余额.xlsx')
            if not os.path.exists(_f):
                _f = os.path.join(_dp, f'{comp}供应商余额{_aux_date_suffix(_dp, "供应商余额")}.xlsx')
            if os.path.exists(_f):
                dirs.append((_f, -1.0))
        _ck = (comp, kind, name_contains, '300')
        if _ck in _CACHE_AUX:
            return list(_CACHE_AUX[_ck])

        def _aux_map(fp):
            """读客户/供应商余额表 → {(cust, subject): 账龄金额合计}。
            ⚡ 2026-08-10 列修复：3900 表『未到期余额』列(col26)仅 40 行有值，真实余额在
            『账龄金额』(col27) 全行有值（应收账款-货款 账龄金额净额 51,344,285.55 = TB 完全一致）；
            优先取账龄金额。"""
            m = {}
            if not fp or not os.path.exists(fp):
                return m
            try:
                wb = _oxl.load_workbook(fp, read_only=True, data_only=True)
            except Exception:
                return m
            ws = wb[wb.sheetnames[0]]
            _hdr = None
            for rr in ws.iter_rows(min_row=1, max_row=3, values_only=True):
                if rr and any('会计科目' in str(v or '') or '一级科目' in str(v or '') for v in rr):
                    _hdr = [str(v or '') for v in rr]
                    break
            _ci = {n: i for i, n in enumerate(_hdr)} if _hdr else {}
            i_name = _ci.get('科目描述', _ci.get('一级科目描述', 19))
            i_bal = _ci.get('账龄金额', _ci.get('未到期余额', 22))
            i_cust = _ci.get('客商描述', 1)
            for rr in ws.iter_rows(min_row=2, values_only=True):
                if not rr:
                    continue
                subj = str(rr[i_name] or '') if i_name < len(rr) else ''
                if name_contains and name_contains not in subj:
                    continue
                cust = str(rr[i_cust] or '').strip() if i_cust < len(rr) else ''
                if not cust:
                    continue
                _bal = float(rr[i_bal] or 0.0) if i_bal < len(rr) else 0.0
                k = (cust, subj)
                m[k] = m.get(k, 0.0) + _bal
            wb.close()
            return m

        # 期末版（最新日期，close）＋ 年初版（最早日期，open）——300 客户余额表多日期
        # 版本（20251231 年初/20260630/20260731 期末），期初=年初版账龄金额。
        def _aux_by_date(kw, latest=True):
            best = None
            bd = ''
            try:
                for f in os.listdir(_dp):
                    if kw in f and f.lower().endswith(('.xlsx', '.xls')):
                        m = re.search(rf'{kw}(\d{{8}})?', f)
                        date = m.group(1) if m and m.group(1) else ''
                        if latest and date >= bd:
                            bd, best = date, os.path.join(_dp, f)
                        elif not latest and (not bd or (date and date < bd)):
                            bd, best = date, os.path.join(_dp, f)
            except OSError:
                pass
            return best

        _f_ar_late = (_paths.get('aux_ar')
                      or _aux_by_date('客户余额', latest=True)
                      or _aux_by_date('客户账龄', latest=True))
        _f_ar_early = _aux_by_date('客户余额', latest=False) or _aux_by_date('客户账龄', latest=False)
        _f_ap_late = (_paths.get('aux_ap')
                      or _aux_by_date('供应商余额', latest=True)
                      or _aux_by_date('供应商账龄', latest=True))
        _f_ap_early = _aux_by_date('供应商余额', latest=False) or _aux_by_date('供应商账龄', latest=False)

        # ⚡ 2026-08-10 本期借/贷发生额补充：GL 序时账行（read_gl_rows 300 分支已带 cust）
        # 按 (客户, 科目名) 聚合 —— aux 客户余额表只有期末列，发生额必须从 GL 补
        # （客户余额表无期初/借/贷列；GL 行客户列=表头『客户描述』动态定位，勿写死 40/41）。
        _gl_dc_map = None
        def _gl_dc():
            nonlocal _gl_dc_map
            if _gl_dc_map is None:
                from collections import defaultdict as _dd
                _gl_dc_map = _dd(lambda: [0.0, 0.0])
                try:
                    _rr = read_gl_rows(_DATA_ROOT, {comp: discover_entities(_DATA_ROOT).get(comp, {})})
                except Exception:
                    _rr = []
                for r in _rr:
                    _cu = str(r.get('cust') or '').strip()
                    _nm = str(r.get('name') or '')
                    if _cu and _nm:
                        _gl_dc_map[(_cu, _nm)][0] += float(r.get('dr') or 0)
                        _gl_dc_map[(_cu, _nm)][1] += float(r.get('cr') or 0)
            return _gl_dc_map

        if kind in ('ar', 'all') and _f_ar_late:
            late_m = _aux_map(_f_ar_late)
            early_m = _aux_map(_f_ar_early) if _f_ar_early else {}
            _dc = _gl_dc()
            for (cust, subj), close in late_m.items():
                _d, _c = _dc.get((cust, subj), (0.0, 0.0))
                out.append({'entity': comp, 'cust': cust, 'subject': subj, 'contract': '',
                            'open': early_m.get((cust, subj), 0.0),
                            'debit': _d, 'credit': _c, 'close': close})
        if kind in ('ap', 'all') and _f_ap_late:
            late_m = _aux_map(_f_ap_late)
            early_m = _aux_map(_f_ap_early) if _f_ap_early else {}
            _dc = _gl_dc()
            for (cust, subj), close in late_m.items():
                _d, _c = _dc.get((cust, subj), (0.0, 0.0))
                out.append({'entity': comp, 'cust': cust, 'subject': subj, 'contract': '',
                            'open': early_m.get((cust, subj), 0.0),
                            'debit': _d, 'credit': _c, 'close': close})
        _CACHE_AUX[_ck] = out
        return out
    dirs = []
    if kind in ('ar', 'all'):
        dirs.append((os.path.join(root, '应收', f'{comp}.xlsx'), 1.0))
    if kind in ('ap', 'all'):
        dirs.append((os.path.join(root, '应付', f'{comp}.xlsx'), -1.0))
    # ⚡ 2026-08-10 性能修复：原 read_only+ws.cell() 逐格访问 + ws.max_row（read_only 下
    #   触发整表扫描）→ current_account 8 科目 × 应收/应付 1.8 万行 ×10 格 = 数分钟卡死；
    #   改 iter_rows 流式 O(行数)。加进程级缓存（同主体多科目循环复用，避免重复读文件）。
    _ck = (comp, kind, name_contains)   # ⚡ name_contains 必须入 key（各往来科目过滤不同）
    if _ck in _CACHE_AUX:
        return list(_CACHE_AUX[_ck])
    for p, sign in dirs:
        if not os.path.exists(p):
            continue
        try:
            wb = _oxl.load_workbook(p, read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 10:
                continue
            # col4=科目编号 col5=科目名称 —— name_contains 必须匹配【科目名称】
            # （原用 r[3]=科目编号，'应收账款' in '1122010000' 恒 False → aux 恒 0）
            subj = str(r[4] or '')
            if name_contains and name_contains not in subj:
                continue
            cust = str(r[2] or '').strip()
            if not cust:
                cust = str(r[1] or '').strip()
            con = str(r[5] or '').strip()
            _open = float(r[6] or 0.0)
            _db = float(r[7] or 0.0)
            _cr = float(r[8] or 0.0)
            _close = float(r[9] or 0.0)
            if sign < 0:
                _open, _close = -_open, -_close   # 应付表余额反号（铁律74）
            out.append({'entity': comp, 'cust': cust, 'subject': subj,
                        'contract': con,
                        'open': _open, 'debit': _db, 'credit': _cr, 'close': _close})
        wb.close()
    _CACHE_AUX[_ck] = out
    return out


def _gl_code4_index():
    """集团模式全量 GL 的 code4 前缀索引（一次遍历建，7 往来科目共用）。
    ⚡⚡ 2026-08-25：read_gl_net 集团模式原每科目全量过滤+净额化 O(n×7)；
    数字前缀（code4）过滤改走索引取子集，过滤结果与全量完全一致（任何 code
    都按前 4 位索引，取回后精确 startswith）。缓存按 _DATA_ROOT 失效。"""
    global _GL_INDEX_BY_CODE4
    _root = _DATA_ROOT
    if _GL_INDEX_BY_CODE4 is not None and _GL_INDEX_BY_CODE4[0] == _root:
        return _GL_INDEX_BY_CODE4[1]
    from collections import defaultdict
    rows = read_gl_rows(_DATA_ROOT, discover_entities(_DATA_ROOT))
    idx = defaultdict(list)
    for r in rows:
        idx[str(r.get('code') or '')[:4]].append(r)
    _GL_INDEX_BY_CODE4 = (_root, idx)
    return idx


def read_gl_net(comp=None, pfx=None, want_asset=True):
    """凭证级净额化（SAP 清账/资金池铁律）：同凭证(vno)内该科目前缀全部行借贷合并计净额。
    粒度 (y, vno, code4)——1010 应收 86.77/85.46 亿 0 差异（曾 (vno,name,unit) 过度净额化）。
    仅用于往来款等存在互转的科目；收入/费用等不净额化（口径=GL 借方合计）。
    返回 read_gl_rows 结构行（净额化后，行自带 cp/entity）。
    ⚡⚡ 2026-08-24 性能优化（用户痛点"集团跑得慢"）：集团模式全量 GL 545 万行，
        7 个往来科目各调一次 → 每次全量净额化 O(545万×7) 重复劳动。净额化结果按
        (comp, pfx, want_asset) 缓存——首次全量净额化一次，同科目复用（read_gl_rows
        本身已缓存，读 GL 非瓶颈；瓶颈是净额化循环）。
    ⚡⚡ 2026-08-25 预索引优化：pfx 支持 str 或 list（current_account 传 [名称, code4]
        名称+代码双条件，特例码账套不漏）；集团模式首次全量 GL 构建 code4→行索引，
        数字前缀（code4）过滤走索引取子集，避免 7 科目各全量遍历。"""
    _ck = (comp or _current_comp, tuple(pfx) if isinstance(pfx, (list, tuple)) else pfx, want_asset)
    if _ck in _CACHE_GL_NET:
        return list(_CACHE_GL_NET[_ck])
    comp = comp or _current_comp
    if comp:
        rows = read_gl_rows(_DATA_ROOT, {comp: discover_entities(_DATA_ROOT).get(comp, {})})
    else:
        # ⚡ 2026-08-10 集团模式：全量（跨主体净额化，key 带 entity 防 vno 冲突）
        rows = read_gl_rows(_DATA_ROOT, discover_entities(_DATA_ROOT))
    if pfx:
        _pfxs = [str(p) for p in (pfx if isinstance(pfx, (list, tuple)) else [pfx])]
        # ⚡⚡ 2026-08-25 code4 预索引：数字前缀（1122 等）先从 code4 索引取子集
        #   （一次全量遍历建索引，7 科目共用），再按原条件过滤——与全量过滤结果
        #   完全一致（任何 code 都按前 4 位索引，取回后精确 startswith），无遗漏。
        _digit = [p for p in _pfxs if p.isdigit()]
        if _digit and not comp:
            _idx = _gl_code4_index()
            _cand = []
            _seen = set()
            for _p in _digit:
                for _r in _idx.get(_p[:4], ()):
                    if id(_r) not in _seen:
                        _seen.add(id(_r))
                        _cand.append(_r)
            rows = [r for r in _cand
                    if any(str(r.get('name') or '').startswith(p)
                           or str(r.get('code') or '').startswith(p) for p in _pfxs)]
        else:
            rows = [r for r in rows
                    if any(str(r.get('name') or '').startswith(p)
                           or str(r.get('code') or '').startswith(p) for p in _pfxs)]
    from collections import defaultdict
    grp = defaultdict(lambda: [0.0, 0.0])
    first_idx = {}   # ⚡ 2026-08-10 性能修复：原「每组再线性扫 rows 找首行」O(组×行) 百亿次
                     # 比较 → current_account 卡死 26 分钟；改为分组时顺带记录首行 O(n)。
    for r in rows:
        key = (r.get('e'), r.get('y'), r.get('vno'), str(r.get('code') or '')[:4])
        grp[key][0] += float(r.get('dr') or 0.0)
        grp[key][1] += float(r.get('cr') or 0.0)
        first_idx.setdefault(key, r)
    out = []
    for (e0, y, vno, c4), (d, c) in grp.items():
        if abs(d) < 0.005 and abs(c) < 0.005:
            continue
        first = first_idx[(e0, y, vno, c4)]
        out.append({'entity': first.get('e') or comp, 'y': y, 'name': first.get('name', ''),
                    'code': first.get('code', ''), 'date': first.get('date', ''),
                    'month': first.get('month', 0), 'vtype': first.get('vtype', ''),
                    'vno': vno, 'summary': first.get('summary', ''), 'sm': first.get('sm', ''),
                    'cust': first.get('cust', ''), 'supp': first.get('supp', ''),
                    'cp': first.get('cp', ''), 'opp': first.get('opp', ''),
                    'dr': max(d, 0.0), 'cr': max(c, 0.0)})
    # ⚡ 2026-08-12 降内存：净额化中间结构（rows/grp/first_idx 合计数 GB）显式释放，
    #   提前让 GC 回收（3500 100 万行 GL：rows 1.5GB + grp 索引 1GB 峰值）。
    del rows, grp, first_idx
    _CACHE_GL_NET[_ck] = out
    return out


def read_gl_tb(comp=None, pfx=None):
    """完全互抵剔除口径（对齐 TB 净额口径，铁律14 深挖结论）：
    按 (year, vno) 分组该科目前缀行，借贷【完全互抵】（|净额|<0.005）的凭证整笔剔除
    （内部互转/红字重记虚增）；其余凭证保留【原始借贷全额】（不净额化——
    借贷不等凭证含真实业务，如红字冲回+新确认）。用于分月明细/分析性程序。"""
    comp = comp or _current_comp
    rows = read_gl_rows(_DATA_ROOT, {comp: discover_entities(_DATA_ROOT).get(comp, {})})
    if pfx:
        rows = [r for r in rows
                if str(r.get('name') or '').startswith(pfx)
                or str(r.get('code') or '').startswith(pfx)]
    from collections import defaultdict
    net = defaultdict(float)
    for r in rows:
        net[(r.get('y'), r.get('vno'))] += float(r.get('dr') or 0.0) - float(r.get('cr') or 0.0)
    drop_vnos = {k for k, n in net.items() if abs(n) < 0.005}
    return [r for r in rows if (r.get('y'), r.get('vno')) not in drop_vnos]


_RECON_FULL_CACHE = {}   # 2026-08-25 (root, comp) -> (full_rows, by_vno) 缓存——recon_counterparties_sap
                         # 原每主体每科目调用都 read_gl_rows(149 万行)+重建 by_vno（68×7=476 次 ≈ 13 分钟）


def recon_counterparties_sap(rows):
    """对方科目核对——凭证级精确分摊（2026-08-10 虚增虚减剔除）：
    粗归集 use_cp 把同凭证所有非自身科目计入且金额=目标行整额 → 同凭证多对方科目金额重复
    （2200 应收贷方侧电费/水费/蒸汽费各 3.66 亿=虚增）。本函数读全量 GL 按 vno 组对方科目
    【真实行金额】：目标行借方→凭证内贷方科目、贷方→借方科目。
    返回 {(ent, pk): dict(sale_cp, sale_cnt, collect_cp, collect_cnt, ...)}（结构对齐
    current_account_detail.reconstruct_counterparties_by_entity）。
    ⚡⚡ 2026-08-25 性能优化：full/by_vno 按 (root, comp) 缓存——原 476 次调用各重建
    （current_account 68 主体 × 7 科目，read_gl_rows 149 万行 pkl 加载 + by_vno 构建）。"""
    from collections import defaultdict
    comp = _current_comp
    _rk = (_DATA_ROOT, comp)
    if _rk in _RECON_FULL_CACHE:
        full, by_vno = _RECON_FULL_CACHE[_rk]
    else:
        if not comp or comp not in discover_entities(_DATA_ROOT):
            # 集团模式（current_comp=None）：全量对方科目核对
            full = read_gl_rows(_DATA_ROOT, discover_entities(_DATA_ROOT))
        else:
            full = read_gl_rows(_DATA_ROOT, {comp: discover_entities(_DATA_ROOT).get(comp, {})})
        by_vno = defaultdict(list)
        for r in full:
            by_vno[(r.get('y'), r.get('vno'))].append(r)
        _RECON_FULL_CACHE[_rk] = (full, by_vno)
    out = {}
    for r in rows:
        ent = r.get('entity') or r.get('e') or comp
        # ⚡ 2026-08-11 修复：pk 用【年度】(date[:4]) 而非 YYYY-MM——_write_recon_counterparty_sheet
        #   按 periods(年度粒度如 ['2026']) 匹配 by_ent.get((ent, pk))，旧 [:7] 致 key 不匹配
        #   → 对方科目核对 sheet 空(有 sheet 无数据)。已验证 1010 AR sale_cp 117.5 亿。
        pk = str(r.get('date') or '')[:4]
        key = (ent, pk)
        d = out.setdefault(key, {'sale_cp': {}, 'collect_cp': {}, 'sale_cnt': {}, 'collect_cnt': {},
                                 'ar_debit_net': 0.0, 'ar_credit_net': 0.0,
                                 'sale_ext': {}, 'collect_ext': {}, 'internal_deb': 0.0, 'internal_cred': 0.0})
        # ⚡ 2026-08-11 P0 修复：rows 来自 read_transactions（current_account_detail），
        #   字段是 debit/credit（大写全名）、无 y 键；read_gl_rows 是 dr/cr。兼容两种。
        _dr = float(r.get('dr') if r.get('dr') is not None else r.get('debit') or 0.0)
        _cr = float(r.get('cr') if r.get('cr') is not None else r.get('credit') or 0.0)
        amt = _dr - _cr
        # ⚡ 2026-08-11：rows(read_transactions) 无 y 键 → 从 date(YYYY-MM-DD) 推断年份，
        #   否则 by_vno 键 (y,vno) 匹配失败 → others 空 → sale/collect_cp 全 0 → 对方科目核对 sheet 不生成。
        _y_key = r.get('y') or r.get('year') or ''
        if not _y_key:
            _ds = str(r.get('date') or '')
            _m = re.search(r'(19|20)\d{2}', _ds)
            if _m:
                _y_key = _m.group(0)
        others = [x for x in by_vno.get((_y_key, r.get('vno')), [])
                  if x.get('name') != r.get('name') and str(x.get('name') or '')]
        if amt > 0.005:
            d['ar_debit_net'] += amt
            for x in others:
                nm = str(x.get('name') or '')
                d['sale_cp'][nm] = d['sale_cp'].get(nm, 0.0) + float(x.get('cr') or 0.0)
                # ⚡ 2026-08-11：sale_cnt 对齐 U8 结构=dict(科目→笔数)，非 int
                #   （_write_recon_counterparty_sheet 用 cnt.get(km,0)，int 会 AttributeError→被吞→sheet 空）
                d['sale_cnt'][nm] = d['sale_cnt'].get(nm, 0) + 1
        elif amt < -0.005:
            d['ar_credit_net'] += -amt
            for x in others:
                nm = str(x.get('name') or '')
                d['collect_cp'][nm] = d['collect_cp'].get(nm, 0.0) + float(x.get('dr') or 0.0)
                d['collect_cnt'][nm] = d['collect_cnt'].get(nm, 0) + 1
    return out


def patch_audit_common():
    """把 audit_common 数据接口切换为 SAP 适配版（生成器本地读取自动走 adapter）。
    ⚡ 只替换属性（AU.xxx = ...）；from-import 绑定不受影响（生成器 SAP 分支应直接调
    _adapter 接口或经 audit_common.discover_entities 属性查找）。"""
    import audit_common as AU
    AU.discover_entities = discover_entities
    AU.read_tb_full = read_tb_full
    AU.read_gl_rows = read_gl_rows
    AU.read_km = read_km
    AU.is_sap = is_sap
    print('[adapter] audit_common 数据接口已切换为 SAP 适配版')
