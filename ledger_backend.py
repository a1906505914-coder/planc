# -*- coding: utf-8 -*-
"""ledger_backend.py —— 账套读取统一抽象层（2026-08-24 架构阶段1+2）。

目标：把『账套形态』的差异（U8/SAP300/AH-SAP 的 列语义/符号/期间/发现规则）
内聚到 backend，业务层（builder）只面向统一接口，不再写 _sap_mode/is_sap 分支。

用法：
    from ledger_backend import get_backend, read_tb
    bk = get_backend(data_dir)          # 自动判定形态 → U8/Sap300/AhBackend
    ents = bk.discover()                # {实体: {年度: {'km','gl','aux'}}}
    tb = bk.read_tb(entities=ents)      # {(e,code,name,y): {qc,jf,df,qm,level}}
    tb = read_tb(data_dir)              # 统一入口（自动 discover）

设计原则：
- 新增模块，不破坏现有调用（read_tb_full/read_sap_tb 等继续可用）
- 形态判定单点：get_backend 只调 detect_sap_layout（含 account_profiles 声明覆盖）
- 每个 backend 只关心自己形态的读取细节
"""
import os

# ============================================================
# 统一入口（业务层直接用，无需关心形态）
# ============================================================
_backend_cache = {}


def get_backend(data_dir):
    """按数据目录形态返回 LedgerBackend 实例（进程内缓存）。"""
    if data_dir in _backend_cache:
        return _backend_cache[data_dir]
    from sap_reader import detect_sap_layout
    layout = detect_sap_layout(data_dir) or 'u8'
    if layout == '300':
        bk = Sap300Backend(data_dir)
    elif layout == 'ah-sap':
        bk = AhBackend(data_dir)
    elif layout == '3300':
        bk = Sap300Backend(data_dir)   # 3300 形态走 SAP 读取（历史形态）
    else:
        bk = U8Backend(data_dir)
    _backend_cache[data_dir] = bk
    return bk


def read_tb(data_dir, entities=None):
    """统一 TB 读取入口：按形态分发 read_tb_full（U8）/ read_sap_tb（SAP 系）。"""
    bk = get_backend(data_dir)
    ents = entities if entities is not None else bk.discover()
    return bk.read_tb(entities=ents)


def read_gl(data_dir, entities=None, pfx=None, comp=None, **kw):
    """统一 GL 读取入口：U8→read_gl_rows（逐行）；SAP→read_gl_net（凭证级净额）。"""
    bk = get_backend(data_dir)
    return bk.read_gl(entities=entities, pfx=pfx, comp=comp, **kw)


def read_aux(data_dir, comp=None, subj=None, **kw):
    """统一辅助余额读取入口。"""
    bk = get_backend(data_dir)
    return bk.read_aux(comp=comp, subj=subj, **kw)


# ============================================================
# 抽象基类
# ============================================================
class LedgerBackend:
    """账套读取统一接口。形态差异内聚到各实现。"""
    layout = 'u8'

    def __init__(self, data_dir):
        self.data_dir = data_dir

    def discover(self, data_dir=None):
        raise NotImplementedError

    def read_tb(self, data_dir=None, entities=None):
        raise NotImplementedError

    def read_gl(self, data_dir=None, entities=None, pfx=None, comp=None, **kw):
        raise NotImplementedError

    def read_aux(self, data_dir=None, comp=None, subj=None, **kw):
        raise NotImplementedError


# ============================================================
# U8 形态：通用科目余额表 + 综合查询明细表（科目代码|科目名称|方向|金额|...）
# ============================================================
class U8Backend(LedgerBackend):
    layout = 'u8'

    def discover(self, data_dir=None):
        from audit_common import discover_entities
        return discover_entities(data_dir or self.data_dir)

    def read_tb(self, data_dir=None, entities=None):
        from audit_common import read_tb_full
        return read_tb_full(data_dir or self.data_dir, entities if entities is not None else self.discover())

    def read_gl(self, data_dir=None, entities=None, pfx=None, comp=None, **kw):
        from audit_common import read_gl_rows
        return read_gl_rows(data_dir or self.data_dir, entities if entities is not None else self.discover())

    def read_aux(self, data_dir=None, comp=None, subj=None, **kw):
        # U8 辅助核算余额表（current_account_detail.read_opening_balances 兼容）
        import current_account_detail as CA
        return CA.read_opening_balances(None, subj, comp=comp)


# ============================================================
# SAP 300 / AH-SAP：SAP 导出（公司|总账|总账科目长文本 / 货币类型|...|公司代码）
# ============================================================
class SapBackend(LedgerBackend):
    """SAP 系公共实现（300/ah-sap 共用读取逻辑，差异在 discover 与 layout）。"""
    layout = 'sap'

    def discover(self, data_dir=None):
        from sap_reader import discover_sap_entities
        return discover_sap_entities(data_dir or self.data_dir)

    def read_tb(self, data_dir=None, entities=None):
        from sap_reader import read_sap_tb
        return read_sap_tb(data_dir or self.data_dir,
                           entities if entities is not None else self.discover())

    def read_gl(self, data_dir=None, entities=None, pfx=None, comp=None, **kw):
        # SAP GL：read_gl_net（凭证级净额，依赖 sap_adapter._DATA_ROOT/current_comp）
        import sap_adapter as SA
        d = data_dir or self.data_dir
        SA.set_root(d)
        if comp:
            SA.set_comp(comp)
        return SA.read_gl_net(pfx=pfx, **kw)

    def read_aux(self, data_dir=None, comp=None, subj=None, **kw):
        import sap_adapter as SA
        d = data_dir or self.data_dir
        SA.set_root(d)
        if comp:
            SA.set_comp(comp)
        return SA.read_aux_balance(d, comp or SA.current_comp(), name_contains=subj.get('kw') if subj else None)


class Sap300Backend(SapBackend):
    layout = '300'


class AhBackend(SapBackend):
    layout = 'ah-sap'


# 便捷：已知形态时直接取 backend（跳过 detect，供形态明确场景）
def backend_for(layout, data_dir):
    if layout == '300':
        return Sap300Backend(data_dir)
    if layout == 'ah-sap':
        return AhBackend(data_dir)
    return U8Backend(data_dir)
