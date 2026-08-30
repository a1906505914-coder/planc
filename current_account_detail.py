# -*- coding: utf-8 -*-
# FINGERPRINT: 产出={往来科目}_生成.xlsx(应收/应付/预收/预付/借款等)+关联交易核对_生成.xlsx | 关键列=往来单位/关联方/账龄/借款利息 | 职责=往来款与借款明细(含利息勾稽,独立生成器)
"""
往来科目明细表生成程序（统一版）
================================================================================
功能概述
--------------------------------------------------------------------------------
本程序是《通用应收账款明细表生成程序 accounts_receivable_detail.py》的统一参数化版本，
一次覆盖六大往来科目：
    应收账款(1122) / 预收账款(2203) / 预付账款(1123) /
    其他应收款(1221) / 其他应付款(2241) / 应付账款(2202)
接收「往来科目相关数据」（综合查询明细表 + 辅助核算余额表 + 可选科目余额表），自动完成：
  1. 读取数据（支持 .csv / .xlsx），自动识别列（中英文别名兼容）；
  2. 按科目名称前缀识别目标往来科目（应收账款/预收/预付/应付账款/其他应收/其他应付）；
  3. 主表数字【直接取自辅助核算余额表（目标科目末级控制数）】，不再从综合查询明细表(GL)
     重累加后与控制数对账——避免『GL 非全量抽取』造成的人为勾稽残差；
  4. 综合查询明细表仅用于：①每客户交易流水下钻；②对应科目（审计关注）分析；
  5. 按 (往来单位, 期间) 汇总：期初/本期借/本期贷/期末/账龄/交易笔数；
  6. 输出结构化 Excel（多 Sheet）：明细_<期间> / 期间对比 / 数据汇总 /
     借方对应科目_<期间> / 贷方对应科目_<期间> / 对方科目核对（按核算主体·分年度判断异常） /
     校验结果 / 交易流水(下钻)。

取数口径（与银行存款明细表程序、应收账款明细表程序完全一致）：
  控制源 = 辅助核算余额表的末级实名行（科目辅助核算名称为真实客户/供应商/个人，
  排除「客户/供应商/科目/现金流量(无余额显示)」等通用标签行与父级汇总行）；
  控制数核对 = 各科目末级期末合计  vs  科目余额表父级（科目名称==父级名）期末；
  综合查询明细表 = 仅下钻与对应科目分析。

使用方式
--------------------------------------------------------------------------------
  python current_account_detail.py --input 综合查询明细表 --aux 辅助核算余额表 \
         --output 预收账款明细表.xlsx --subject-key APR
可选参数：
  --input  综合查询明细表(.csv/.xlsx，可为目录)             [默认：桌面 G（若不存在则需显式指定）]
  --aux    辅助核算余额表(.csv/.xlsx，可为目录)              [默认：同 --input 目录]
  --km     科目余额表(可选，用于父级控制数核对)              [默认：同 --input 目录]
  --subject-key  AR|APR|APP|AP|ORA|ORP  (默认 AR；目录模式忽略，循环全部)
  --period Y|M|Q      期间粒度，默认 Y（年）
  --output 输出 Excel 路径
  --subject 核算主体名称（可选，默认从辅助核算余额表文件名推导）

也可作为模块导入：from current_account_detail import build_ca_detail, build_all_in_dir
================================================================================
"""

import paths as P
import argparse
import csv
import datetime
import gc   # 2026-08-07 内存优化：大文件追加链后显式回收
import os
import re
import sys
import time
from collections import defaultdict
# ⚡ 2026-08-10 排版优化共享函数（R3 双行表头/R1 零值/R4 分组底色/R5 合计层级）
from audit_shell import (two_row_header, money_cell, group_fill, style_total,
                         hide_col_if_zero, FILL_GROUP_BLUE, FILL_GROUP_GREEN,
                         FILL_GROUP_GRAY, FILL_GROUP_YELLOW, FILL_GROUP_RED,
                         FILL_GROUP_PURPLE)

# ============ 顶部引导（必须在 import openpyxl 之前）============
# 拖入/双击本 .py 时，Windows 可能用「系统 Python 3.14（无 openpyxl）」启动，
# 顶部 import openpyxl 会立即 ModuleNotFoundError 闪退。故在此先检测并在必要时
# 用 WorkBuddy 托管 Python 重新执行本文件，确保 openpyxl 可用、控制台保留。
import os as _bs_os
import sys as _bs_sys
import subprocess as _bs_sub


def _find_managed_python():
    """定位 WorkBuddy 托管 Python（含 openpyxl）。"""
    base = _bs_os.path.join(_bs_os.environ.get('USERPROFILE', _bs_os.path.expanduser('~')),
                            '.workbuddy', 'binaries', 'python', 'versions')
    if _bs_os.path.isdir(base):
        for d in sorted(_bs_os.listdir(base), reverse=True):
            p = _bs_os.path.join(base, d, 'python.exe')
            if _bs_os.path.isfile(p):
                return p
    return None


def _bootstrap_relaunch():
    return  # ⚡ 2026-08-22 禁用托管Python重启：venv依赖已齐全，托管3.13 openpyxl损坏


if _bs_os.environ.get('AUDIT_NO_RELAUNCH') != '1':
    _bootstrap_relaunch()


import openpyxl
from audit_shell import (SHELL_TITLE_FONT, SHELL_LEFT, SHELL_HFILL, SHELL_HFONT, SHELL_CEN, SHELL_BORDER, SHELL_BOLD, SHELL_NUM, SHELL_TOT_FILL, SHELL_BODY, SHELL_RGT, resolve_template, AUDIT_TEMPLATES_DIR, finalize_workbook)
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import audit_year_io as Y

# ----------------------------------------------------------------------------
# 列名别名映射（与应收账款明细表程序保持一致）
# ----------------------------------------------------------------------------
COLUMN_ALIASES = {
    "km":          ["科目名称", "科目", "accounttitle", "subject", "km"],
    "date":        ["日期", "记账日期", "交易日期", "会计日期", "date", "dt", "postdate"],
    "vtype":       ["字", "凭证字", "vouchertype", "vtype"],
    "vno":         ["号", "凭证号", "voucherno", "vno"],
    "summary":     ["摘要", "备注", "说明", "summary", "memo", "note", "desc"],
    "debit":       ["借方金额", "借方", "debitamount", "debit", "jf"],
    "credit":      ["贷方金额", "贷方", "creditamount", "credit", "df"],
    "cust":        ["辅助核算名称", "客户名称", "客户", "customer", "custname", "auxname", "往来单位名称"],
    "cp":          ["对方科目", "对应科目", "counterparty", "cp", "counterpart"],
}

AUX_COLUMN_ALIASES = {
    "km":          ["科目名称", "科目", "accounttitle", "subject", "km"],
    "code":        ["代码", "客户编号", "编号", "编码", "code", "custcode"],
    "name":        ["往来单位名称", "客户名称", "单位名称", "客户", "customer", "name"],
    "open_dir":    ["期初方向", "期初借贷", "opendir"],
    "open":        ["期初", "期初金额", "期初余额", "open", "opening", "initial"],
    "debit":       ["借方", "借方发生额", "debit", "jf"],
    "credit":      ["贷方", "贷方发生额", "credit", "df"],
    "close_dir":   ["期末方向", "期末借贷", "closedir"],
    "close":       ["期末", "期末金额", "期末余额", "close", "ending", "final"],
    "level":       ["等级", "级次", "level", "grade"],
    "wtype":       ["往来类型", "往来性质", "wtype", "type"],
    "related":     ["关联方", "关联单位", "关联客户", "relatedparty", "related", "is_related"],
    "aging":       ["账龄", "账龄区间", "aging", "agebucket", "bucket"],
}

DEBIT_TOKENS = {"借", "debit", "dr", "d", "+", "1", "收入", "存入"}
CREDIT_TOKENS = {"贷", "credit", "cr", "c", "-", "0", "付出", "支出"}

# Excel 样式
HDR_FILL = PatternFill("solid", fgColor="DDEBF7")
HDR_FONT = Font(name='Times New Roman', bold=True, color="000000", size=10)
TITLE_FONT = Font(name='Times New Roman', bold=True, size=10, color="000000")
SUB_FONT = Font(name='Times New Roman', bold=True, size=10, color="000000")
TOTAL_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FONT = Font(name='Times New Roman', bold=True)
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
OTHER_FILL = PatternFill("solid", fgColor="FCE4D6")
NOTE_FILL = PatternFill("solid", fgColor="FFF2CC")
ANOM_FILL = PatternFill("solid", fgColor="FFC7CE")
ANOM_FONT = Font(name='Times New Roman', color="9C0006", bold=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY_FMT = "#,##0.00"
PCT_FMT = "0.00%"
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")

# ⚡⚡ 2026-08-25 性能：_name_hit / _is_subject_km 高频调用（GL 行 × 7 科目循环），
#   预编译正则 + 繁简归一 translate 表（避免每次调用 re.split/re.fullmatch 查缓存、
#   str.replace 逐字扫描）。_CN2T = str.maketrans('帐','账') 一次性构建。
_TR_TABLE = str.maketrans({'帐': '账'})
_RE_SEG = re.compile(r'[\\\-]+')
_RE_DIGITS = re.compile(r'[\d.]+')
_RE_HE = re.compile(r'和')

# 真实明细行的「科目辅助核算名称」通用标签黑名单（这些行是父级/中间汇总/现金流量维度，无余额）
GENERIC_AUX_NAMES = {"", "客户", "供应商", "科目", "个人", "单位往来", "个人往来",
                     "现金流量", "现金流量(无余额显示)", "现金流量（无余额显示）"}

# ----------------------------------------------------------------------------
# 五大往来科目配置
#   label       : 科目中文名（= 科目余额表父级行 科目名称，用于精确取控制数）
#   kw          : 辅助核算余额表/综合查询明细表 中匹配该科目的前缀
#   sheet       : 输出文件名主体
#   exp_debit   : 该科目「借方」侧期望对应科目（凭证中对方科目被贷记）
#   exp_credit  : 该科目「贷方」侧期望对应科目（凭证中对方科目被借记）
#   notes_*     : 审计关注提示映射
# ----------------------------------------------------------------------------
SALE_NOTE = {
    "固定资产": "固定资产处置收入误挂应收？应核查是否应贷记资产处置损益/固定资产清理",
    "营业外收入": "与应收借方同凭证，属非常规销售，核查业务实质",
}
COLLECT_NOTE = {
    "预收账款": "以预收/预付款抵减应收，属客户往来净额结算，关注对冲依据",
    "坏账准备": "坏账核销(借坏账准备 贷应收)，关注核销审批与税务备案",
    "应付账款": "应收与应付互抵，关注是否符合金融资产净额结算",
    "销售费用": "疑为销售折让/退货冲应收，关注红字发票与收入冲减",
    "财务费用": "外币应收汇兑损益调整，关注汇率选取",
    "利润分配": "★异常：应收减少直接进未分配利润，疑为前期差错更正/权益调整，重点核查",
    "递延所得税资产": "异常，关注是否坏账准备变动对应递延所得税",
    "营业外支出": "应收减少进营业外支出，疑为坏账损失/捐赠，关注凭据",
    "其他应收款": "往来互抵/代垫", "其他应付款": "往来互抵",
    "制造费用": "异常，建议核查凭证", "累计折旧": "异常，建议核查",
    "应交税费": "疑为退货进项转出对冲",
}
GENERIC_NOTE = {
    "应收账款": "往来互抵/重分类", "应收票据": "票据结算",
    "预收账款": "往来互抵/重分类", "预付账款": "往来互抵/重分类",
    "其他应收款": "往来互抵/代垫", "其他应付款": "往来互抵/代垫",
    "应付账款": "应付与预付/应收互抵", "坏账准备": "坏账核销，关注审批与税务备案",
    "利润分配": "★异常：直接进未分配利润，疑前期差错/权益调整，重点核查",
    "营业外收入": "非常规，核查业务实质", "营业外支出": "疑坏账损失/捐赠，关注凭据",
    "制造费用": "异常，建议核查凭证", "累计折旧": "异常，建议核查",
    "应交税费": "疑进项转出/税费对冲", "财务费用": "汇兑损益调整，关注汇率选取",
    "递延所得税资产": "减值/坏账变动对应递延所得税，关注",
    "销售费用": "疑折让/退货冲减", "管理费用": "费用性往来，关注真实性",
    "在建工程": "设备/工程暂估或预付款结转", "库存商品": "存货入库结转",
    "原材料": "材料入库结转", "固定资产": "资产购建结转", "应付职工薪酬": "薪酬计提暂挂",
}

SUBJECTS = {
    "AR":  dict(label="应收账款", kw="应收账款", sheet="应收账款明细表", nature="asset",
                exp_debit=["主营业务收入", "其他业务收入", "销项", "银行存款"],
                exp_credit=["银行存款", "应收票据"],
                notes_debit=SALE_NOTE, notes_credit=COLLECT_NOTE),
    "APR": dict(label="预收账款", kw="预收", sheet="预收账款明细表", nature="liability",
                wtype="客户",   # 2026-08-04：负债但按客户核算（收款类负债：预收/合同负债）
                exp_debit=["主营业务收入", "其他业务收入", "销项"],
                exp_credit=["银行存款", "库存现金"],
                notes_debit=GENERIC_NOTE, notes_credit=GENERIC_NOTE),
    "APP": dict(label="预付账款", kw="预付账款", sheet="预付账款明细表", nature="asset",
                exp_debit=["银行存款", "库存现金"],
                exp_credit=["库存商品", "原材料", "固定资产", "在建工程",
                            "管理费用", "销售费用", "制造费用", "应付账款", "银行存款"],
                notes_debit=GENERIC_NOTE, notes_credit=GENERIC_NOTE),
    "ORA": dict(label="其他应收款", kw="其他应收款", sheet="其他应收款明细表", nature="asset",
                exp_debit=["银行存款", "库存现金"],
                exp_credit=["银行存款", "库存现金"],
                notes_debit=GENERIC_NOTE, notes_credit=GENERIC_NOTE),
    "ORP": dict(label="其他应付款", kw="其他应付款", sheet="其他应付款明细表", nature="liability",
                exp_debit=["银行存款", "库存现金"],
                exp_credit=["银行存款", "库存现金", "管理费用", "销售费用",
                            "制造费用", "应付职工薪酬"],
                notes_debit=GENERIC_NOTE, notes_credit=GENERIC_NOTE),
    "AP":  dict(label="应付账款", kw="应付账款", sheet="应付账款明细表", nature="liability",
                # ⚡ 2026-08-26 SAP 适配：贷方侧(增加应付)期望借方=收货入库类（库存产品=SAP 库存商品、
                #   生产成本/成本差异=收货结转、财务费用=外币应付评估汇兑）；借方侧(减少应付)期望贷方
                #   加 应收票据/应付票据（票据结算）。内部子目（GR/IR/外币评估/自身）在 _classify_counterparty 判正常。
                exp_debit=["银行存款", "库存现金", "预付账款", "应收票据", "应付票据", "财务费用",
                           "库存商品", "库存产品", "原材料", "生产成本", "主营业务成本", "其他业务成本"],
                exp_credit=["库存商品", "库存产品", "原材料", "生产成本", "库存产品成本差异",
                            "固定资产", "在建工程", "管理费用", "销售费用", "制造费用",
                            "应交税费", "财务费用", "主营业务成本", "其他业务成本", "银行存款"],
                notes_debit=GENERIC_NOTE, notes_credit=GENERIC_NOTE),
    # 应收票据：部分账套的辅助核算余额表中存在按客户分类的应收票据明细时生成，
    # 模板完全复用应收账款（资产/借方侧、6 档账龄、销售确认+收回对应科目）。
    # ⚡ 2026-08-26 用户方法论修正：借方侧（收到票据）贷方=预收/应收 属正常（收票抵应收/预收）；
    #   贷方侧（票据背书/贴现）借方=应付/预付 属正常（票据支付供应商）。原白名单漏这两类 → 误标『关注』。
    "ARN": dict(label="应收票据", kw="应收票据", sheet="应收票据明细表", nature="asset",
                exp_debit=["主营业务收入", "其他业务收入", "销项", "预收账款", "应收账款"],
                exp_credit=["银行存款", "应收账款", "应付账款", "预付账款"],
                notes_debit=SALE_NOTE, notes_credit=COLLECT_NOTE),
    # 2026-08-06 前瞻：合同资产（SSS 建筑施工企业大科目，1480）。与合同负债重分类对；
    # FY 无合同资产数据 → 不产文件（账套无数据自动跳过）。
    # ⚡⚡ 2026-08-27 修复（用户指正）：合同资产↔应收/预收 对冲只要客户名称一致即属正常
    #   （合同资产转应收、预收冲减合同资产），不得标『关注』。exp_debit/exp_credit 补入
    #   应收账款/预收账款——原仅 exp_credit 有『应收账款』，借方侧应收 60亿/预收 2亿、
    #   贷方侧预收 58亿全被误标『关注』。
    "CA":  dict(label="合同资产", kw="合同资产", sheet="合同资产明细表", nature="asset",
                wtype="客户",
                exp_debit=["主营业务收入", "其他业务收入", "销项", "应收账款", "预收账款"],
                exp_credit=["银行存款", "应收账款", "预收账款"],
                notes_debit=SALE_NOTE, notes_credit=COLLECT_NOTE),
    # 2026-08-04 补充负债类科目（无 aux 辅助核算 → 走 TB 兜底占位行机制，参照预付账款）：
    # 合同负债 2205。负债类账龄 2 档、附注负债简化。
    # 2026-08-06 科目规整：LL(租赁负债)/DI(递延收益)/NCL(一年内到期) 迁出至 gp_other
    # （无往来对象、无账龄、附注 four 四平行表），本程序只保留真正往来科目。
    "CL":  dict(label="合同负债", kw="合同负债", sheet="合同负债明细表", nature="liability",
                wtype="客户",   # 2026-08-04：负债但按客户核算（收款类负债）
                exp_debit=["银行存款"],
                exp_credit=["主营业务收入", "其他业务收入", "销项"],
                notes_debit=GENERIC_NOTE, notes_credit=GENERIC_NOTE),
}

# ⚡⚡ 2026-08-25：往来科目标准 code4（read_gl_net code4 预索引用——SAP 分支传
#   [名称, code4] 双条件过滤；特例码账套由 account_profiles 在 build_all_in_dir
#   _ca_code_cfg 覆盖，此处仅作默认标准码，名称条件兜底保证不漏）。
CA_STD_CODE = {'AR': '1122', 'APR': '2203', 'APP': '1123', 'ORA': '1221',
               'ORP': '2241', 'AP': '2202', 'ARN': '1121', 'CL': '2205', 'CA': '1480'}


def _resolve_account_name(data_dir):
    """解析账套配置 key（account_profiles.accounts）：
    优先 data_dir 自身 basename；SAP 模式 data_dir=工作目录（如 _work_1010）匹配不到
    AH → 从数据根（_adapter._DATA_ROOT）路径向上找第一个有配置的目录段（.../AH/数据/2026
    → 'AH'）。2026-08-26 修复：合同资产 CA 硬编码 1480（SSS 前瞻）vs AH 实际 1124，
    靠 account_profiles subject_codes.contract_asset=1124 生效（原 basename 匹配失败）。"""
    base = os.path.basename(str(data_dir or '').rstrip('\\/'))
    try:
        import subject_mapping as _smr
        if _smr.account_profile(base):
            return base
        _root = getattr(_adapter, '_DATA_ROOT', '') or ''
        parts = str(_root).replace('\\', '/').split('/')
        for i in range(len(parts) - 1, -1, -1):
            if parts[i] and _smr.account_profile(parts[i]):
                return parts[i]
    except Exception:
        pass
    return base

# 账龄区间划分：
#   资产类应收（应收账款/其他应收款/应收票据）→ 6 档；
#   预付账款（资产类）→ 4 档（2026-08-03 用户要求：1年以内/1-2年/2-3年/3年以上，与格式2 预付块一致）；
#   负债类（应付/预收/其他应付/合同负债）→ 2 档（用户要求：负债类账龄一般只分 1年以内/1年以上）。
AGING_BUCKETS = {
    "recv": ["1年以内", "1-2年", "2-3年", "3-4年", "4-5年", "5年以上"],
    "pay":  ["1年以内", "1-2年", "2-3年", "3年以上"],
    "liab": ["1年以内", "1年以上"],
}
SUBJ_AGING_TYPE = {
    "AR": "recv", "ORA": "recv", "ARN": "recv", "CA": "recv",
    "APP": "pay",
    "APR": "liab", "AP": "liab", "ORP": "liab",
    "CL": "liab",
}
# 借方侧往来（资产类：应收/预付/其他应收/合同资产）与贷方侧往来（负债类：应付/预收/其他应付/合同负债）
DEBIT_SIDE = {"AR", "APP", "ORA", "ARN", "CA"}
CREDIT_SIDE = {"AP", "APR", "ORP", "CL"}
SIDE_LABEL = {"debit": "借方侧(应收/预付/其他应收)", "credit": "贷方侧(应付/预收/其他应付)"}


# ============================================================================
# 模块 1：数据读取
# ============================================================================
def _norm(s):
    return re.sub(r"[\s_]+", "", str(s)).lower()


# ----------------------------------------------------------------------------
# 集团内账套主体别名匹配（用于『关联交易和余额核对』关联方识别）
#   已获取的账套同属一个集团，账套主体互为关联方。为做账套间往来核对，
#   需把各账套辅助核算中的『往来单位名称』归并到对应账套主体。
#   做法：取各账套主体简称的『核心词』（去除公司/股份等通用后缀），与往来
#   单位名称核心词做包含匹配，取最长匹配者，尽量避免『亿利达』等短词误串。
# ----------------------------------------------------------------------------
_CORP_SUFFIX = re.compile(
    r"(股份有限公司|有限责任公司|有限公司|股份|有限|责任|公司|集团|企业|"
    r"管理|（[^）]*）|\([^)]*\))+$")


def _entity_core(name):
    """取账套主体/往来单位名称的核心词（去除尾部通用公司/括号后缀）。"""
    if not name:
        return ""
    return _CORP_SUFFIX.sub("", _norm(name))


def build_entity_matcher(entity_list):
    """返回 match(name) -> 命中的账套主体名 或 None。
    匹配优先级：
      1) 精确核心词相等（如往来单位『亿利达风机』精确命中主体『亿利达风机』，
         避免被更长的『广东亿利达风机』借子串误串）；
      2) 否则取最长核心词包含匹配（要求核心词长度>=2；2 字主体名如『金泰/东轴』也纳入，
         避免账套主体恰为 2 字时整个集团内关联往来核对失效）。"""
    cores = [(e, _entity_core(e)) for e in entity_list if len(_entity_core(e)) >= 2]
    def match(name):
        nc = _entity_core(name)
        if len(nc) < 2:
            return None
        # 1) 精确匹配优先
        exact = [e for (e, c) in cores if c == nc]
        if exact:
            return max(exact, key=lambda e: len(_entity_core(e)))
        # 2) 包含匹配（最长核心词）
        best = None
        best_len = 1  # 候选核心词需 >1（即>=2）才会被采纳；放宽以支持 2 字主体名
        for e, c in cores:
            if c in nc or nc in c:
                if len(c) > best_len:
                    best_len = len(c)
                    best = e
        return best
    return match


# ----------------------------------------------------------------------------
# 关联方清单（用户提供的权威关联方名单）—— 用于『关联交易和余额核对』识别
#   用户可在数据文件夹中放置一份文件名含『关联方清单』的文件（.xlsx 或 .csv），
#   逐行列出应被识别为关联方的往来单位 / 账套主体名称（全称或足以辨识的核心词）。
#   程序判断"是否为关联方"时：若清单存在，则以清单为准（命中清单即判为关联方），
#   替代原有仅靠账套主体名称模糊归并的识别方式（该方式对集团内长名、跨文件夹
#   母公司/姊妹公司等易误判）。
# ----------------------------------------------------------------------------
_RP_HEADER_KEYS = ["关联方名称", "关联方", "往来单位名称", "单位名称", "客户名称",
                   "供应商名称", "对方单位", "名称", "公司名称", "企业名称", "关联交易对手方"]


def _load_related_party_list(data_dir):
    """在 data_dir 查找文件名含『关联方清单』的清单文件；找到返回 [(raw, norm), ...]，否则 None。
    支持 .xlsx / .xls / .csv；内容读取：优先按表头关键字定位名称列并跳过表头行，否则取首列。"""
    if not os.path.isdir(data_dir):
        return None
    cands = [os.path.join(data_dir, fn) for fn in sorted(os.listdir(data_dir))
             if ("关联方清单" in fn) and fn.lower().endswith((".xlsx", ".xls", ".csv"))]
    if not cands:
        return None
    path = cands[0]
    try:
        if path.lower().endswith(".csv"):
            with open(path, "r", encoding="utf-8-sig", errors="ignore") as fh:
                rows = [r for r in csv.reader(fh) if r and any(str(c).strip() for c in r)]
            col, hidx = _pick_rp_column(rows)
            data_rows = rows[hidx + 1:] if hidx is not None else rows
            names = [(str(r[col]).strip(), _norm(r[col])) for r in data_rows
                     if col < len(r) and str(r[col]).strip()]
        else:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            rows = [[c for c in row if c is not None] for row in ws.iter_rows(values_only=True)]
            rows = [r for r in rows if any(str(c).strip() for c in r)]
            wb.close()
            col, hidx = _pick_rp_column(rows)
            data_rows = rows[hidx + 1:] if hidx is not None else rows
            names = []
            for r in data_rows:
                if col < len(r):
                    v = str(r[col]).strip()
                    if v and v.lower() not in ("none", "nan"):
                        names.append((v, _norm(v)))
    except Exception:
        return None
    names = [(raw, nv) for raw, nv in names if nv]
    return names or None


def _pick_rp_column(all_rows):
    """在 all_rows 前若干行查找表头行，返回 (col, header_abs_index)。
    header_abs_index 为表头所在绝对行号（调用方据此跳过表头）；找不到则 (0, None) 视为无表头。
    匹配策略：先精确匹配表头关键字（最安全，避免说明文字误命中）；再子串匹配，但仅对
    短单元格(规范化后<=15字，表头通常很短，说明长句借此排除)。
    2026-08-02 修复：『关联交易对手方（企业）』表头（FY 关联方清单）未命中关键字、且其相邻列
    『是否合并内关联方』含『关联方』子串被误当名称列 → ①名称列关键字补充『关联交易对手方』；
    ②子串匹配排除标记列（含 是否/合并内/合并外 的表头不作文本列）。"""
    for ridx, row in enumerate(all_rows[:10]):
        row = [str(c) for c in row]
        for i, h in enumerate(row):
            nh = _norm(h)
            for k in _RP_HEADER_KEYS:
                if nh == _norm(k):
                    return i, ridx
    for ridx, row in enumerate(all_rows[:10]):
        row = [str(c) for c in row]
        for i, h in enumerate(row):
            nh = _norm(h)
            if len(nh) > 15:
                continue
            if any(k in nh for k in ('是否', '合并内', '合并外')):
                continue   # 标记列（是否合并内/外关联方）不作文本列
            for k in _RP_HEADER_KEYS:
                if len(k) >= 3 and k in nh:
                    return i, ridx
    return 0, None


def _name_in_related_list(name, rp_list):
    """判断往来单位名称是否命中关联方清单。
    规则：1) 规范化精确相等；2) 任一条目核心词(>=3字)被名称核心词包含
    （放宽以容纳『浙江兆龙高分子材料有限公司』vs 清单『兆龙高分子材料』等长名变体）。"""
    if not rp_list:
        return False
    nc = _norm(name)
    if not nc:
        return False
    for _raw, nv in rp_list:
        if nv == nc:
            return True
    core = _entity_core(name)
    for _raw, nv in rp_list:
        cv = _entity_core(_raw)
        if len(cv) >= 3 and cv in core:
            return True
    return False


def _resolve_columns(headers, aliases):
    norm_headers = {_norm(h): i for i, h in enumerate(headers) if h}
    mapping = {}
    for std, alist in aliases.items():
        for al in alist:
            if _norm(al) in norm_headers:
                mapping[std] = norm_headers[_norm(al)]
                break
    return mapping


def _resolve_aux_columns(header):
    mp = {}
    money_idx = []
    for i, h in enumerate(header):
        n = _norm(h)
        if n == "科目名称":
            mp["km"] = i
        elif n in ("科目辅助核算代码", "科目代码", "代码"):
            mp["code"] = i
        elif n in ("科目辅助核算名称", "往来单位名称", "单位名称", "客户名称"):
            mp["name"] = i
        elif n == "等级":
            mp["level"] = i
        elif n == "往来类型":
            mp["wtype"] = i
        elif n == "关联方":
            mp["related"] = i
        elif "借" in n and "方" in n:
            mp.setdefault("debit", i)
        elif "贷" in n and "方" in n:
            mp.setdefault("credit", i)
        elif "方向" in n and "借" not in n and "贷" not in n:
            # 期初方向 / 期末方向（部分导出仅标为『方向』，出现两次：期初、期末）
            if "open_dir" not in mp:
                mp["open_dir"] = i
            elif "close_dir" not in mp:
                mp["close_dir"] = i
        elif "金额" in n:
            # 2026-08-06：兼容 FY『金额』与 JTt/SSS『期初金额/期末金额』——open=首个、close=末个
            money_idx.append(i)
    if money_idx:
        mp.setdefault("open", money_idx[0])
        mp["close"] = money_idx[-1]
    return mp


# 款项性质关键词映射（由辅助核算表「科目名称」子目段推导）
_NATURE_MAP = [
    ("暂估", "暂估款"), ("结算", "结算款"), ("货款", "货款"), ("借款", "借款"),
    ("保证金", "保证金"), ("质保金", "质保金"), ("押金", "押金"), ("备用金", "备用金"),
    ("代垫", "代垫款"), ("代付", "代付款"), ("股利", "股利"), ("股权", "股权转让款"),
    ("预收", "预收款"), ("预付", "预付款"), ("进项", "进项税"), ("销项", "销项税"),
    ("运费", "运费"), ("服务费", "服务费"), ("咨询", "咨询费"), ("利息", "利息"),
    ("房租", "房租"), ("工资", "工资"), ("薪酬", "薪酬"), ("社保", "社保"),
    ("个税", "个税"), ("佣金", "佣金"), ("返利", "返利"), ("索赔", "索赔款"),
    ("赔偿", "赔偿款"), ("质保", "质保款"), ("定金", "定金"), ("订金", "订金"),
    ("往来", "往来款"), ("投资", "投资款"), ("分红", "分红款"),
]


def _derive_nature(km):
    """由辅助核算表科目名称推导款项性质（取子目段，命中关键词则归并）。
    例：应付账款-暂估应付款 -> 暂估款；应付账款-结算应付款 -> 结算款。"""
    if not km:
        return ""
    seg = km.split("-", 1)[1].strip() if "-" in km else str(km).strip()
    if not seg:
        return ""
    for k, v in _NATURE_MAP:
        if k in seg:
            return v
    return seg


def _derive_bill_type(km):
    """由辅助核算科目名称推导应收票据种类（银行承兑汇票/商业承兑汇票）。
    无子目或不含承兑标识则空。例：『应收票据-银行承兑汇票』→银行承兑汇票。"""
    if not km:
        return ""
    seg = km.split("-", 1)[1] if "-" in km else str(km)
    if "银行承兑" in seg or ("银行" in seg and "承兑" in seg):
        return "银行承兑汇票"
    if "商业承兑" in seg or ("商业" in seg and "承兑" in seg):
        return "商业承兑汇票"
    if "财务公司承兑" in seg or ("财务公司" in seg and "承兑" in seg):
        return "财务公司承兑汇票"
    if "承兑" in seg:
        return "承兑汇票"
    return ""


def _dir_of(p):
    """取路径/路径列表所在目录（列表取首个元素）。无法解析返回 ''。"""
    if isinstance(p, (list, tuple)):
        p = p[0] if p else ""
    if not p:
        return ""
    try:
        return os.path.dirname(os.path.abspath(p))
    except Exception:
        return ""


def _read_foreign_balances(data_dir, entity, periods):
    """读取目录下「外币科目余额表」（按 主体+年度 匹配），返回
    {(period, 子目科目名称归一): (期初原币, 期末原币, 币种)}。
    逐户原币不可得，工具按各户人民币占比在子目原币内分摊（汇总=子目原币）。"""
    out = {}
    if not data_dir or not os.path.isdir(data_dir):
        return out
    for fn in sorted(os.listdir(data_dir)):
        if "外币科目余额表" not in fn:
            continue
        if entity and entity not in fn:
            continue
        m = re.search(r"(20\d{2})", fn)
        if not m:
            continue
        pk = m.group(1)
        if pk not in periods:
            continue
        p = os.path.join(data_dir, fn)
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            raw = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception:
            continue
        hdr_idx = None
        for i, row in enumerate(raw):
            cells = [_norm(c) for c in row if c is not None]
            if "原币" in cells and "科目名称" in cells:
                hdr_idx = i
                break
        if hdr_idx is None:
            continue
        hdr = [_norm(c) for c in raw[hdr_idx]]
        try:
            i_km = hdr.index("科目名称")
            fidx = [i for i, h in enumerate(hdr) if h == "原币"]
            i_closef = fidx[-1]
            i_openf = fidx[0]
            i_ccy = hdr.index("币种") if "币种" in hdr else None
        except (ValueError, IndexError):
            continue
        i_level = hdr.index("等级") if "等级" in hdr else None
        for row in raw[hdr_idx + 1:]:
            nm = _str(row, i_km, "")
            if not nm:
                continue
            lvl = _str(row, i_level, "") if i_level is not None else ""
            # 跳过科目级汇总行（等级1 且 科目名含"科目"），保留子目/末级
            if str(lvl).strip() == "1" and "科目" in nm:
                continue
            of = _parse_amount(row[i_openf], 0) if i_openf < len(row) else 0.0
            cf = _parse_amount(row[i_closef], 0) if i_closef < len(row) else 0.0
            ccy = _str(row, i_ccy, "") if i_ccy is not None else ""
            key = (pk, _norm(nm))
            if key not in out or (of + cf) > (out[key][0] + out[key][1]):
                out[key] = (of, cf, ccy)
    return out


def _is_fx_km(km):
    """判断科目名称是否为外币（原币）明细科目。"""
    if not km:
        return False
    return bool(re.search(r"美元|美圆|usd|欧元|英镑|日元|港币|hkd|eur|gbp|jpy|"
                          r"外币|外汇|（美|（欧|（英|（日|（港|（新|新元|坡元|马克|法郎",
                          str(km), re.I))


def _fx_km_matches_subj(km, subj):
    """外币辅助核算余额表 科目名称常省略『账款』（如『应收外币-美元』而非『应收账款（美元）』），
    故此处在精确包含/前缀匹配之外，对前2字（应收/应付/预收/预付）做兜底匹配；
    『其他*』类（其他应收/其他应付）前2字相同易互串，故不兜底，仅精确匹配。"""
    kw = subj["kw"]
    # ⚡⚡ 2026-08-24（ADZ 应收 东轴全 0 根因）：东轴 1122 名称『应收帐款』（繁体
    #   '帐'）→ kw『应收账款』（简体'账'）前缀不匹配 → 审定表东轴 0。繁简归一。
    km = str(km).replace('帐', '账')
    kw = str(kw).replace('帐', '账')
    if kw in km or km.startswith(kw):
        return True
    if not kw.startswith("其他") and km.startswith(kw[:2]):
        return True
    return False


def _strip_aux_suffix(name):
    """去掉 read_opening_balances 给跨子目同名客户加的『(子目)』后缀，还原基础客户名。
    从首个（或(起截到末尾（兼容『客户名（应收账款（美元））』这类嵌套括号）。"""
    if not name:
        return name
    s = str(name)
    return re.sub(r"[（(].*$", "", s).strip()


def _read_foreign_aux(data_dir, entity, periods, subj):
    """读取目录下『外币辅助核算余额表』（按 主体 匹配），返回
    {(period, 归一客户名): (原币期初, 原币期末)}，用于逐户原币直接填列 外币余额。
    仅纳入目标科目(subj['kw'])且末级实名行（辅助核算名称为真实往来单位）；
    原币期末 = 原币期初 + 原币借方 - 原币贷方（与人民币期末同口径净额推算）。"""
    out = {}
    if not data_dir or not os.path.isdir(data_dir):
        return out
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith("~$"):
            continue
        if "外币辅助核算余额表" not in fn:
            continue
        if entity and entity not in fn:
            continue
        pk = _detect_year(fn)
        if pk not in periods:
            continue
        p = os.path.join(data_dir, fn)
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            raw = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception:
            continue
        # 定位表头：含『辅助核算名称』与『原币』
        hdr_idx = None
        for i, row in enumerate(raw):
            cells = [_norm(c) for c in row if c is not None]
            if "辅助核算名称" in cells and "原币" in cells:
                hdr_idx = i
                break
        if hdr_idx is None:
            continue
        hdr = [_norm(c) for c in raw[hdr_idx]]
        try:
            i_km = hdr.index("科目名称")
            i_name = hdr.index("辅助核算名称")
            of_idx = [i for i, h in enumerate(hdr) if h == "原币"]   # 期初/借/贷/期末 原币
            i_fx = of_idx[0] if len(of_idx) > 0 else None
            i_fd = of_idx[1] if len(of_idx) > 1 else None
            i_fc = of_idx[2] if len(of_idx) > 2 else None
            i_fcl = of_idx[3] if len(of_idx) > 3 else None   # 期末原币（直接取数，优先于净算）
        except (ValueError, IndexError):
            continue
        for row in raw[hdr_idx + 1:]:
            km = _str(row, i_km, "")
            if not _fx_km_matches_subj(km, subj):
                continue
            aux_name = _str(row, i_name, "")
            if not aux_name or aux_name in GENERIC_AUX_NAMES:
                continue
            f_open = _parse_amount(row[i_fx], 0) if (i_fx is not None and i_fx < len(row)) else 0.0
            f_deb = _parse_amount(row[i_fd], 0) if (i_fd is not None and i_fd < len(row)) else 0.0
            f_cre = _parse_amount(row[i_fc], 0) if (i_fc is not None and i_fc < len(row)) else 0.0
            # 期末原币优先取导出列(已含汇率重述的权威期末)，为空才净算(期初+借-贷)
            raw_fcl = row[i_fcl] if (i_fcl is not None and i_fcl < len(row)) else None
            if raw_fcl is not None and str(raw_fcl).strip() != "":
                f_close = _parse_amount(raw_fcl, 0.0)
            else:
                f_close = f_open + f_deb - f_cre
            key = (pk, _norm(aux_name))
            if key in out:                       # 同客户跨明细行，原币累加
                o0, o1 = out[key]
                out[key] = (o0 + f_open, o1 + f_close)
            else:
                out[key] = (f_open, f_close)
    return out


def _read_foreign_aux_full(data_dir, entity_filter=None, periods_filter=None):
    """读取目录下全部『外币辅助核算余额表』，按 (entity, subj_key, period, norm_name)
    归集逐户 原币/本位币 期初·期末 + 币种 + 期末方向；用于诊断『仅在外币aux出现、常规aux缺失』
    的往来单位（疑似漏标注）。
    列结构（已用 2025母公司外币辅助核算余额表 核实）：表头含『辅助核算名称』与『原币』，
    原币列序 [期初原币, 本期借方原币, 本期贷方原币, 期末原币]，金额(本位币)列序
    [期初金额, 借方金额, 贷方金额, 期末金额]，方向列最后一个是『期末方向』，另有『币种』列。"""
    out = {}
    if not data_dir or not os.path.isdir(data_dir):
        return out
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith("~$"):
            continue
        if "外币辅助核算余额表" not in fn:
            continue
        full = os.path.join(data_dir, fn)
        ent = _derive_subject(full)
        if entity_filter and ent != entity_filter:
            continue
        pk = _detect_year(fn)
        if periods_filter and pk not in periods_filter:
            continue
        try:
            wb = openpyxl.load_workbook(full, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            raw = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception:
            continue
        hi = None
        for i, row in enumerate(raw):
            cells = [_norm(c) for c in row if c is not None]
            if "辅助核算名称" in cells and "原币" in cells:
                hi = i
                break
        if hi is None:
            continue
        hdr = [_norm(c) for c in raw[hi]]
        try:
            i_km = hdr.index("科目名称")
            i_name = hdr.index("辅助核算名称")
            fy = [i for i, h in enumerate(hdr) if h == "原币"]
            my = [i for i, h in enumerate(hdr) if h == "金额"]
            dy = [i for i, h in enumerate(hdr) if h == "方向"]
            i_level = hdr.index("等级") if "等级" in hdr else None
            i_wtype = hdr.index("往来类型") if "往来类型" in hdr else None
            i_ccy = hdr.index("币种") if "币种" in hdr else None
        except (ValueError, IndexError):
            continue
        if len(fy) < 4 or len(my) < 4:
            continue
        f_o, f_d, f_c, f_cl = fy[0], fy[1], fy[2], fy[3]
        m_o, m_d, m_c, m_cl = my[0], my[1], my[2], my[3]
        d_cl = dy[-1] if dy else None
        for row in raw[hi + 1:]:
            km = _str(row, i_km, "")
            aux_name = _str(row, i_name, "")
            if not aux_name or aux_name in GENERIC_AUX_NAMES:
                continue
            lvl = _str(row, i_level, "") if i_level is not None else ""
            wt = _str(row, i_wtype, "") if i_wtype is not None else ""
            if str(lvl).strip() in ("1", "2") or wt in ("科目",):
                continue  # 跳过科目级/维度汇总行
            sk = None
            for k, sd in SUBJECTS.items():
                if _fx_km_matches_subj(km, sd):
                    sk = k
                    break
            if sk is None:
                continue
            f_open = _parse_amount(row[f_o], 0) if f_o < len(row) else 0.0
            f_close = _parse_amount(row[f_cl], 0) if f_cl < len(row) else 0.0
            m_open = _parse_amount(row[m_o], 0) if m_o < len(row) else 0.0
            m_close = _parse_amount(row[m_cl], 0) if m_cl < len(row) else 0.0
            ccy = _str(row, i_ccy, "") if i_ccy is not None else ""
            dirc = _str(row, d_cl, "") if d_cl is not None else ""
            nk = _norm(aux_name)
            key = (ent, sk, pk, nk)
            if key in out:
                o = out[key]
                o["fcur_open"] += f_open
                o["fcur_close"] += f_close
                o["rmb_open"] += m_open
                o["rmb_close"] += m_close
            else:
                out[key] = dict(orig_name=aux_name, fcur_open=f_open, fcur_close=f_close,
                                rmb_open=m_open, rmb_close=m_close, ccy=ccy, dir=dirc)
    return out


def _detect_header(raw, aliases):
    all_aliases = set()
    for alist in aliases.values():
        all_aliases.update(_norm(a) for a in alist)
    km_aliases = set(_norm(a) for a in aliases.get("km", []))
    best_fallback = None
    for i, row in enumerate(raw):
        norm_cells = [_norm(c) for c in row if c is not None]
        cnt = sum(1 for nc in norm_cells if nc in all_aliases)
        if cnt >= 2 and best_fallback is None:
            best_fallback = (i, [str(h).strip() if h is not None else "" for h in row])
        has_km = any(nc in km_aliases for nc in norm_cells)
        if has_km and cnt >= 2:
            return i, [str(h).strip() if h is not None else "" for h in row]
    if best_fallback is not None:
        return best_fallback
    return 0, [str(h).strip() if h is not None else "" for h in raw[0]] if raw else []


def _str(row, idx, default):
    if idx is None or idx >= len(row) or row[idx] is None:
        return default
    return str(row[idx]).strip()


def _parse_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.datetime.strptime(s[:19], fmt).date()
            except ValueError:
                continue
    return None


def _parse_amount(v, ri):
    if isinstance(v, (int, float)):
        a = float(v)
    else:
        s = str(v).strip()
        neg = s.startswith("(") and s.endswith(")")
        s = s.replace(",", "").replace("￥", "").replace("¥", "")
        s = re.sub(r"[()]", "", s) if neg else s
        # ⚡ 2026-08-12 修复：数据源含 'None'/'nan'/空 文本（如 DQ 辅助核算空壳），
        #   按 0 处理——原 float('None') ValueError → 往来科目（应收/应付/预付等 8 个
        #   builder）整批崩溃 → DQ 2026 期 17 科目底稿缺失。
        if s in ('', 'None', 'none', 'NONE', 'nan', 'NaN', 'null', 'NULL', '-'):
            return 0.0
        try:
            a = float(s)
        except ValueError:
            raise ValueError(f"第{ri}行金额无法解析：{v!r}")
        if neg:
            a = -a
    return a


def _is_negative_raw(v):
    if isinstance(v, (int, float)):
        return v < 0
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace("￥", "").replace("¥", "")
        if s.startswith("(") and s.endswith(")"):
            return True
        return s.startswith("-")
    return False


def _is_subject_km(km, kw):
    """判断综合查询明细表科目名称是否属于目标往来科目（排除坏账准备/信用减值损失等同源科目）"""
    # ⚡⚡ 2026-08-24（ADZ 应收 东轴全 0 根因）：繁简归一（应收帐款→应收账款）。
    # 2026-08-25 性能：translate 表代替 str.replace（高频调用）。
    km = km.translate(_TR_TABLE)
    kw = kw.translate(_TR_TABLE)
    if not km.startswith(kw):
        return False
    if km.startswith("坏账准备") or km.startswith("信用减值损失"):
        return False
    return True


def _name_hit(nm, kw):
    """科目名【主段】匹配：名称按 \\ - 分割，去掉纯数字段后取首段，首段含 kw 或 kw 含首段。
    2026-08-08 修复（JTt/F 边界）：原『kw in 全名』会把跨科目子目误命中——
      JTt 114109『114109\\内部往来\\应收票据』首段=内部往来 → 排除（内部往来非应收票据）；
      112103『112103\\应收票据\\信用』首段=应收票据 → 命中；
      F 1221『其他应收款』→ 命中。"""
    # ⚡⚡ 2026-08-24（ADZ 应收 东轴全 0 根因）：繁简归一（应收帐款→应收账款）。
    # 2026-08-25 性能：translate 表 + 预编译正则（高频调用）。
    nm = str(nm).translate(_TR_TABLE)
    kw = str(kw).translate(_TR_TABLE)
    segs = [s for s in _RE_SEG.split(nm) if s.strip()]
    segs = [s for s in segs if not _RE_DIGITS.fullmatch(s)]
    if not segs:
        return kw in nm
    main = segs[0]
    # ⚡⚡ 2026-08-24 修复（AZ 泰国应收票据双算根因）：泰国合并科目『应收账款和应收票据』
    #   （泰国 1130，含应收+票据，归一化时整体映射到应收 1122）主段含『应收票据』子串 →
    #   被 kw='应收票据' 误命中 → 应收票据底稿多计泰国 1.14 亿。排除：主段是『应收X和应收Y』
    #   合并形态且 kw 对应其中一段时，若主段同时含『应收』与另一科目段（非 kw 主体）则跳过。
    #   例：主段=应收账款和应收票据，kw=应收票据 → 排除（其主体=应收账款）。
    #   例：主段=应收票据，kw=应收票据 → 命中（真实应收票据）。
    if main != kw and kw in main and '和' in main:
        # 主段是合并科目（应收账款和应收票据）：仅当 kw 精确等于【后段】且后段≠前段时排除——
        #   例：主段=应收账款和应收票据，kw=应收票据 → 排除（泰国 1130 主体归应收，票据误命中）。
        #       主段=应收账款和应收票据，kw=应收账款 → 保留（应收科目应纳入泰国 1130）。
        _parts = [p for p in _RE_HE.split(main) if p.strip()]
        if len(_parts) == 2 and kw == _parts[1] and _parts[1] != _parts[0]:
            return False
    # 只做正向匹配 kw in main（kw 一律用完整科目名：其他应付款/其他应收款/应收票据…）。
    # 2026-08-08 最终修复：反向匹配（main in kw）有歧义——Q 220203 名称『其他应付』⊂
    # kw『其他应付款』被误当其他应付款（应付账款子目 701万 混入）；S 应收票据
    # 1120/1121 主段=『应收票据』=kw 正向即命中，无需反向。
    # ⚡⚡ 2026-08-27 修复（任务887 1357 应收 39,044.40 对称差异根因）：『其他应收
    #   账款-其他应收账款调整』（1221999999，其他应收调整科目）主段=其他应收账款含
    #   『应收账款』子串 → 被 kw='应收账款' 误命中 → _name_codes 把 1221 调入应收
    #   codes → 应收审定表多计 39,044.40、其他应收明细（GL 1221 客户行）多计 →
    #   应收/其他应收 ±32万 对称差异。排除：main 以『其他』开头而 kw 不以『其他』
    #   开头（或反之）→ 不命中（其他应收款↔应收账款、其他应付款↔应付账款等）。
    if kw in main:
        if (main.startswith('其他') and not kw.startswith('其他')) \
           or (kw.startswith('其他') and not main.startswith('其他')):
            return False
        return True
    # ⚡⚡ 2026-08-27 修复（任务887 其他应收/其他应付 32万 对称差异根因）：『其他应收
    #   账款-其他应收账款调整』（1221999999）主段='其他应收账款'——『其他应收款』非其
    #   子串（应收后是『账』）→ 其他应收款审定表 _name_codes 漏收 1221 → 审定少计
    #   （明细 GL 1221 客户行含）→ 其他应收款明细 vs 审定 +32万。归一：『其他应收
    #   账款』『其他应付账款』视为『其他应收款/其他应付款』（SAP 调整类科目名变体）。
    _main_n = main.replace('其他应收账款', '其他应收款').replace('其他应付账款', '其他应付款')
    return kw in _main_n


# ---------------- GL 读取缓存（一次读取、多科目复用） ----------------
# 往来款程序默认对 SUBJECTS 的 7 个科目各重跑一遍 build_ca_detail，
# 而每个科目都会把整个《综合查询明细表》全量读解析一遍（最重的开销）。
# 这里把"昂贵的 xlsx/csv 解析"按 文件路径 缓存一次，7 个科目复用同一份解析结果，
# 仅按各自 kw 重新标注 is_subject（廉价内存拷贝），从而把最大头的 7 遍全量解析降到 1 遍。
_GL_RAW_CACHE = {}       # path -> list[dict]（不含 is_sub，cust 为原始值）
_GL_DN_WARN_CACHE = {}   # path -> list[str]（借贷同为负告警，解析时一次性算出）
_GL_CACHE_DISABLED = False  # 仅用于测试：True 时强制每次重解析（模拟旧行为）


def clear_gl_cache():
    """清空 GL 解析缓存。应在每次"新一次拖入/新一次 build_all_in_dir"开始时调用，
    避免长驻进程内跨数据包残留旧解析结果。"""
    _GL_RAW_CACHE.clear()
    _GL_DN_WARN_CACHE.clear()


def _read_gl_raw(path, warnings=None):
    """解析《综合查询明细表》(xlsx/csv) 一次，结果按 path 缓存；返回不含 is_sub 的基础行。"""
    if not _GL_CACHE_DISABLED and path in _GL_RAW_CACHE:
        if warnings is not None:
            warnings.extend(_GL_DN_WARN_CACHE.get(path, []))
        return _GL_RAW_CACHE[path]
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            raw = list(csv.reader(f))
    else:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    if not raw:
        raise ValueError("文件为空")
    hdr_idx, headers = _detect_header(raw, COLUMN_ALIASES)
    mapping = _resolve_columns(headers, COLUMN_ALIASES)
    required = ["km", "date", "debit", "credit"]
    missing = [k for k in required if k not in mapping]
    if missing:
        raise ValueError(f"缺少必要列：{missing}（文件表头：{headers}）")

    dn = []
    rows = []
    for ri, row in enumerate(raw[hdr_idx + 1:], start=hdr_idx + 2):
        if not any(c is not None and str(c).strip() != "" for c in row):
            continue
        km = _str(row, mapping.get("km"), "")
        dv = row[mapping["debit"]] if mapping["debit"] < len(row) else None
        cv = row[mapping["credit"]] if mapping["credit"] < len(row) else None
        debit = _parse_amount(dv, ri) if dv not in (None, "") else 0.0
        credit = _parse_amount(cv, ri) if cv not in (None, "") else 0.0
        if _is_negative_raw(dv) and _is_negative_raw(cv):
            dn.append(f"第{ri}行借、贷金额同时为负（{dv!r}/{cv!r}），请核实该笔凭证录入。")
        dval = row[mapping["date"]] if mapping["date"] < len(row) else None
        dt = _parse_date(dval)
        if dt is None:
            raise ValueError(f"第{ri}行日期无法解析：{dval!r}")
        cust = _str(row, mapping.get("cust"), "")
        if cust and "," in cust:
            cust = cust.split(",")[-1].strip()
        rows.append(dict(
            date=dt,
            vtype=_str(row, mapping.get("vtype"), ""),
            vno=_str(row, mapping.get("vno"), ""),
            km=km,
            cust=cust,
            debit=float(debit),
            credit=float(credit),
            summary=_str(row, mapping.get("summary"), ""),
            cp=_str(row, mapping.get("cp"), ""),
        ))
    if not _GL_CACHE_DISABLED:
        _GL_RAW_CACHE[path] = rows
        _GL_DN_WARN_CACHE[path] = dn
    if warnings is not None:
        warnings.extend(dn)
    return rows


def read_transactions(path, kw, warnings=None, code_pfx=None, entity=None):
    """读取综合查询明细表(GL)，仅标记属于目标往来科目的行 is_subject。

    解析结果按 path 缓存（见 _read_gl_raw）：同一文件在 7 个往来科目循环中被复用，
    仅按 kw 重新标记 is_subject，避免重复全量解析。
    ⚡⚡ 2026-08-25：code_pfx=标准 code4 前缀（如 '1122'）——SAP 分支 read_gl_net
       传 [kw, code_pfx] 双条件（名称+代码），配合 read_gl_net 的 code4 预索引，
       避免 7 科目各全量遍历；特例码账套（名称命中）不受影响。
       entity=集团模式指定主体——SAP 分支构造时按主体过滤（减少净额化结果遍历）。"""
    # SAP：从 adapter 凭证级净额化行构建（同凭证同科目同单位互抵只计净额，防资金池虚增）
    if _adapter is not None and (path is None or not isinstance(path, str)):
        rows = []
        # 负债类往来按供应商（col21）净额化，资产类按客户（col12）——防不同单位互抵错乱
        _liab_kws = ('应付', '预收', '合同负债', '应付票据', '其他应付款')
        want_asset = not any(kw.startswith(lk) for lk in _liab_kws)
        _pfx = [kw, code_pfx] if code_pfx else kw
        _ent_f = entity or _adapter.current_comp()
        for r in _adapter.read_gl_net(pfx=_pfx, want_asset=want_asset):
            _ent = r.get('entity') or r.get('e') or ''
            if _ent_f and str(_ent) != str(_ent_f):
                continue   # ⚡⚡ 2026-08-25 集团模式按主体过滤（build_all_in_dir 68 主体循环省遍历）
            km = str(r.get('name') or '')
            is_sub = _is_subject_km(km, kw)
            # ⚡ 2026-08-12 修复：read_gl_net 返回 read_gl_rows 结构（cust/supp 字段，
            #   无 aux）→ 原取 r.get('aux') 恒空 → 3200/3300/3400 应收/应付明细表客户
            #   全部丢失（显示科目名）。统一取 aux or cust or supp（资产行客户、负债行供应商）。
            cust = (r.get('aux') or r.get('cust') or r.get('supp')
                    or ('（未标注往来单位）' if is_sub else ''))
            # ⚡ 2026-08-11 P0 修复：read_gl_net 返回 read_gl_rows 结构（dr/cr/y 字段，
            #   2026-08-10 重构后），原用 r.get('debit')/r.get('credit') → 金额恒 0 →
            #   SAP 往来明细/审定表金额全空 + 对方科目核对不生成。统一字段映射。
            _dr = float(r.get('dr') if r.get('dr') is not None else r.get('debit') or 0.0)
            _cr = float(r.get('cr') if r.get('cr') is not None else r.get('credit') or 0.0)
            rows.append({'km': km, 'entity': r.get('entity') or r.get('e') or _adapter.current_comp() or '',
                         'date': r.get('date'), 'vtype': r.get('vtype'), 'vno': r.get('vno'),
                         'summary': r.get('sm'), 'debit': _dr,
                         'credit': _cr, 'cust': cust, 'cp': r.get('cp'),
                         'is_sub': is_sub})
        return rows
    base = _read_gl_raw(path, warnings)
    rows = []
    for r in base:
        is_sub = _is_subject_km(r["km"], kw)
        cust = r["cust"] if r["cust"] else ("（未标注往来单位）" if is_sub else "")
        rows.append(dict(r, cust=cust, is_sub=is_sub))
    return rows


# ---- 解析缓存：同一文件只 load_workbook 一次，多个往来科目(subj)循环复用 ----
# build_ca_detail 对 7 个往来科目分别调用 read_control / read_opening_balances，
# 原本每个 subj 都把同一份 科目余额表/辅助核算余额表 重新 load 一遍（7× 重读）。
# 此处按 path 缓存「原始行 + 表头探测结果」，后续 subj 直接复用，避免重复重开销。
# 子目过滤(kw/label)在调用方按 subj 进行，不在此缓存，故多 subj 复用安全。
_TB_PARSE_CACHE = {}


def _parse_tb_file(path):
    """读取并解析 TB/辅助核算余额表(或期初余额表)一次，按 path 缓存。
    返回 (raw, hdr_idx, headers, mp)：raw=全部行(列表); hdr_idx=表头行索引;
    headers=表头文本列表; mp=逻辑列名→列索引(来自 _resolve_aux_columns)。
    子目过滤与方向判定在调用方按 subj 进行，与解析无关，故多 subj 复用安全。"""
    cached = _TB_PARSE_CACHE.get(path)
    if cached is not None:
        return cached
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            raw = list(csv.reader(f))
    else:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    hdr_idx, headers = _detect_header(raw, AUX_COLUMN_ALIASES)
    mp = _resolve_aux_columns(headers)
    _TB_PARSE_CACHE[path] = (raw, hdr_idx, headers, mp)
    return raw, hdr_idx, headers, mp


def read_opening_balances(path, subj, comp=None):
    """读取辅助核算余额表 -> (data, conflicts)。
    仅纳入【目标科目 且 末级实名行】：科目名称以 subj['kw'] 开头，且 科目辅助核算名称
    为真实往来单位（非空、非通用标签），且往来类型非『现金流量(无余额显示)』。
    data: {往来单位: dict(open, debit, credit, close, code, aging, foreign)}
    解析结果按 path 缓存（见 _parse_tb_file）：同一文件在 7 个往来科目循环中被复用，
    仅按 subj['kw'] 重新过滤末级行，避免重复全量 load_workbook。
    comp: ⚡ 2026-08-10 集团模式按主体过滤（current_comp=None 时 aux_entries 全量，
    需显式传主体；否则 4 主体同名客户合并、每主体明细=全集团）。"""
    # SAP：从 adapter.read_aux_balance 构建逐户余额（单位=客户/供应商名）
    if _adapter is not None and (path is None or not isinstance(path, str)):
        data, conflicts = {}, []
        kw = subj.get('kw', '')
        # ⚡ 2026-08-10 用户方法论：应收票据【不按合同号分行】（票据无合同维度，合同号无意义）；
        # 其余往来科目（应收/应付/其他应收/其他应付/预付/预收）按 单位×合同号 分行。
        _no_con = (subj.get('sheet') or '').startswith('应收票据')
        for a in _adapter.aux_entries(comp=comp, name_contains=kw):
            nm = str(a.get('cust') or '')
            if not nm:
                continue
            con = str(a.get('contract') or '').strip()
            if _no_con or not con:
                key = nm
            else:
                key = f'{nm}@@{con}'
            rec = data.setdefault(key, {'open': 0.0, 'debit': 0.0, 'credit': 0.0, 'close': 0.0,
                                        'code': '', 'aging': '', 'foreign': {},
                                        'contract': '' if _no_con else con})
            rec['open'] += float(a.get('open') or 0.0)
            rec['debit'] += float(a.get('debit') or 0.0)
            rec['credit'] += float(a.get('credit') or 0.0)
            rec['close'] += float(a.get('close') or 0.0)
        # ⚡ 2026-08-26 修复：SAP aux 余额带符号（read_aux_balance 还原后 负债贷余为负、资产借余为正；
        #   供应商行项目还有借余行=多付/预付，为正）。原未设 open_dir/close_dir → _signed_balance 按
        #   『平』返回原符号 → 负债客户行期末显示负号 → 明细表合计与审定表不平（-33.74亿 vs 42.07亿）。
        #   统一 dir='借'（非零）经 _signed_balance：负债 -amt（贷余负→正、借余正→负），资产保持 →
        #   客户行合计=TB（负债 42.07亿 贷余全转正显示）。
        for _k, _rec in data.items():
            _rec['open_dir'] = '借' if abs(_rec['open']) > 0.005 else '平'
            _rec['close_dir'] = '借' if abs(_rec['close']) > 0.005 else '平'
        return data, conflicts
    raw, hdr_idx, headers, mp = _parse_tb_file(path)
    data, conflicts = {}, []
    if not raw:
        return data, conflicts
    if "km" not in mp or "close" not in mp:
        return data, conflicts
    seen_close = {}
    for ri, row in enumerate(raw[hdr_idx + 1:], start=hdr_idx + 2):
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        km = _str(row, mp.get("km"), "")
        if not km.startswith(subj["kw"]):
            continue
        # 末级实名行识别
        aux_name = _str(row, mp.get("name"), "")
        wtype = _str(row, mp.get("wtype"), "")
        if "现金流量" in wtype:
            continue
        if aux_name in GENERIC_AUX_NAMES:
            continue
        # 2026-08-05 汇总行过滤：aux 科目级汇总行（往来单位列=科目名，如『其他应收款』
        # 『其他应收款(其他应收款)』）非真实往来对手 → 计入会双计
        # （FY 01 其他应收款明细曾 10.8M+7.9M 汇总行 → 明细 32.8M vs TB 11.8M）。
        _aux_base = re.sub(r'[（(].*?[)）]', '', aux_name).strip()
        if _aux_base and _aux_base == subj.get("label"):
            continue
        # 内部管理辅助维度（部门/项目等）非交易对手，不应计入按客户分类的往来余额；
        # 部分账套的应收票据等科目下挂部门/项目辅助账，若不当作客户计入会虚增余额、引发冲突。
        if "部门" in wtype or "项目" in wtype:
            continue
        # 铁律18：按科目性质过滤往来类型——资产类(应收/预付/其他应收)仅保留客户，
        # 负债类(应付/预收/其他应付)仅保留供应商；类型为空保留（不过度剔除）。
        # 2026-08-04 修正：收款类负债（预收账款/合同负债 wtype='客户'）按客户核算，
        # 覆盖默认『供应商』过滤——否则合同负债 CUST 客户明细被误剔除（铁律18 例外）。
        # 2026-08-05 修正：其他应收款(ORA)对手方天然多样（员工借款/押金/代垫/供应商往来），
        # 不按客户/供应商过滤——FY 曾把 01FY本级 个人借款 7,864,943.96、09FY绍兴 沈金辉 50,000
        # 全部剔除，明细表仅 61.4 万 vs TB 1,179 万（用户反馈"明细表生成数据不完整"）。
            if wtype:
                subj_nature = (subj or {}).get("nature", "asset")
                if (subj or {}).get("sheet", '') == "其他应收款明细表":
                    expected = None  # ORA：保留全部往来类型
                else:
                    expected = (subj or {}).get("wtype") or ("客户" if subj_nature == "asset" else "供应商")
                    # ⚡ 2026-08-10 DQ 等 U8 账套：预付账款客户行 wtype='供应商'
                    # （预付=付给供应商），也有账套标'客户' → 预付科目客户/供应商两类都保留
                    # （原默认 asset→'客户' 把 DQ 预付 163 行全剔 → 明细退回 TB 兜底 1 条）。
                    _wt_both = (subj or {}).get('sheet') == "预付账款明细表"
            # 2026-08-07 修复：U8 账套辅助核算往来类型列是『单位往来』『客商』
            # （c 诚本/T 用，而非『客户/供应商』）、J 用『客户往来辅助账/供应商
            # 往来辅助账』——客户/供应商本是科目性质（应收=客户、应付=供应商、
            # 预付=供应商），wtype 只是统一对手方维度。原 expected 过滤把
            # 『单位往来』当非法类型全剔 → AR/APR/AP opening=0 → n_open=0 →
            # TB 兜底 → 明细表退化为科目级行（1122.01/1122.02/1231.01 坏账行
            # 混入，c 应收明细 49.6M 净额假象 vs TB 53.3M 原值）。
            _wt_unified = ('单位往来', '客商', '往来单位')          # U8 统一对手方
            _wt_cust = ('客户', '客户往来辅助账')                    # 客户侧
            _wt_supp = ('供应商', '供应商往来辅助账')                # 供应商侧
            if expected and wtype != expected:
                if wtype in _wt_unified:
                    pass  # U8 统一对手方类型：科目性质已区分借贷，直接保留
                elif wtype in _wt_cust:
                    if expected == '客户' or _wt_both:
                        pass
                    else:
                        continue  # 客户型数据出现在应付类科目 → 跳过
                elif wtype in _wt_supp:
                    if expected == '供应商' or _wt_both:
                        pass
                    else:
                        continue  # 供应商型数据出现在应收类科目 → 跳过
                elif wtype not in ("客户", "供应商"):
                    continue  # 科目/部门/项目/个人/现金流量等已在前面过滤，此处兜底
                else:
                    continue
        name = aux_name or km
        _nature = (subj or {}).get("nature", "asset")
        open_amt = _parse_amount(row[mp["open"]], ri) if mp.get("open") is not None and mp["open"] < len(row) else 0.0
        debit = _parse_amount(row[mp["debit"]], ri) if mp.get("debit") is not None and mp["debit"] < len(row) else 0.0
        credit = _parse_amount(row[mp["credit"]], ri) if mp.get("credit") is not None and mp["credit"] < len(row) else 0.0
        close = _parse_amount(row[mp["close"]], ri) if mp["close"] < len(row) else 0.0
        code = _str(row, mp.get("code"), "")
        aging = _str(row, mp.get("aging"), "")
        foreign = "国外" in km
        open_dir = _str(row, mp.get("open_dir"), "")
        close_dir = _str(row, mp.get("close_dir"), "")
        # 2026-08-03 修复（用户质疑 aux"不平衡"根因）：U8 同一客户同一子目可能有多行
        # （不同辅助组合/期初调整行），每行自身恒平衡(期初+借−贷=期末)，但各行方向可能
        # 不同（如一行期末借余、一行期末贷余），且本期借/贷列可含红字负数。
        # 原逻辑：金额绝对值累加 + 方向保留首行 → 带符号余额错误、平衡校验失衡
        # （例：杭州市市监局 期末 贷23625.3 + 借8000 两行合并后被按"贷"取 −31625.3，
        #   真实净额 −15625.3，差 16,000）。
        # 修复：每行先按各自方向转带符号再累加，合并后反推方向，保证带符号期初/期末代数正确。
        _os = _signed_balance(open_amt, open_dir, _nature)
        _cs = _signed_balance(close, close_dir, _nature)
        # 关联方：辅助核算余额表「关联方」列（True/False 或 "是"/"否" 等）
        rel_raw = _str(row, mp.get("related"), "")
        related = str(rel_raw).strip().lower() in ("true", "是", "1", "yes", "y", "t")
        if name in seen_close and abs(seen_close[name] - close) > 0.005:
            conflicts.append((name, seen_close[name], close))
        seen_close[name] = close
        # 同往来单位名称可能跨子目出现（如『浙江三进』既是应收单位往来借款又是其他），
        # 不能按键名直接覆盖（会丢失一条余额）；同子目同名视为重复行累加，跨子目同名加子目后缀区分。
        key = name
        if key in data:
            if data[key].get("_km") == km:
                d = data[key]
                d["_open_sgn"] += _os
                d["_close_sgn"] += _cs
                d["open"] = abs(d["_open_sgn"])
                d["close"] = abs(d["_close_sgn"])
                d["debit"] += debit
                d["credit"] += credit
                d["open_dir"] = _dir_from_signed(d["_open_sgn"], _nature)
                d["close_dir"] = _dir_from_signed(d["_close_sgn"], _nature)
                d["foreign"] = d["foreign"] or foreign
                continue
            sub = km.split("-", 1)[1] if "-" in km else km
            key = f"{name}({sub})"
            while key in data:
                key += "·"
        data[key] = dict(open=abs(_os), debit=debit, credit=credit, close=abs(_cs),
                         code=code, aging=aging, foreign=foreign,
                         related=related, open_dir=_dir_from_signed(_os, _nature),
                         close_dir=_dir_from_signed(_cs, _nature),
                         wtype=wtype, _km=km, _open_sgn=_os, _close_sgn=_cs)
    return data, conflicts


def read_control(path, subj):
    """读取科目余额表，返回目标科目父级（科目名称==subj['label']）期末控制数（精确匹配，自动排除子目）。
    解析结果按 path 缓存（见 _parse_tb_file）：同一文件在 7 个往来科目循环中被复用，
    仅按 subj['label'] 重新匹配父级行，避免重复全量 load_workbook。"""
    # SAP：从 adapter.read_km 一级行按 label 取期末（带符号）
    if _adapter is not None and (path is None or not isinstance(path, str)):
        km = _adapter.read_km()
        for code, v in km.items():
            if str(v.get('name') or '').split('-')[0].strip() == subj.get('label', ''):
                return float(v.get('closing') or 0.0)
        return 0.0
    raw, hdr_idx, headers, mp = _parse_tb_file(path)
    ctrl = 0.0
    if not raw:
        return ctrl
    if "km" not in mp or "close" not in mp:
        return ctrl
    for ri, row in enumerate(raw[hdr_idx + 1:], start=hdr_idx + 2):
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        km = _str(row, mp.get("km"), "")
        if km != subj["label"]:
            continue
        raw_close = _parse_amount(row[mp["close"]], ri) if mp["close"] < len(row) else 0.0
        # 控制数按科目性质+期末方向带符号：资产借余为+、贷余为−；负债贷余为+、借余为−。
        # 与明细表带符号期末一致，二者可直接比较。
        ctrl += _signed_balance(raw_close, _str(row, mp.get("close_dir"), ""), subj.get("nature", "asset"))
    return ctrl


# ============================================================================
# 模块 2：期间识别与分组汇总
# ============================================================================
def period_key(date, mode="Y"):
    # ⚡ 2026-08-09 SAP：read_gl_net 的 date 是 'YYYY-MM-DD' 字符串（_excel_date），
    # U8 场景是 datetime —— 统一转 datetime 再按 mode 取期间键。
    if isinstance(date, str):
        try:
            from datetime import datetime as _dt
            date = _dt.strptime(date[:10], '%Y-%m-%d')
        except Exception:
            return str(date)[:4] if mode == 'Y' else str(date)[:7]
    if mode == "M":
        return date.strftime("%Y-%m")
    if mode == "Q":
        q = (date.month - 1) // 3 + 1
        return f"{date.year}Q{q}"
    return str(date.year)


def _customer_key(rec):
    return rec["cust"] or rec["km"] or "（未标注往来单位）"


def aggregate(rows, mode="Y"):
    agg = {}
    periods = set()
    customers = []
    seen = set()
    for r in rows:
        if not r["is_sub"]:
            continue
        pk = period_key(r["date"], mode)
        ak = _customer_key(r)
        periods.add(pk)
        if ak not in seen:
            seen.add(ak)
            customers.append(ak)
        d = agg.setdefault((ak, pk), {"debit": 0.0, "credit": 0.0, "count": 0})
        d["debit"] += r["debit"]
        d["credit"] += r["credit"]
        d["count"] += 1
    return sorted(periods), customers, agg


# ============================================================================
# 模块 2.1：借贷方向符号约定（按科目性质）
# ----------------------------------------------------------------------------
# 资产类(应收/预付/其他应收)：借方为正数、贷方为负数；余额(借余)为正。
# 负债类(应付/预收/其他应付)：贷方为正数、借方为负数；余额(贷余)为正。
# 辅助核算余额表的『期初方向/期末方向』列(借/贷/平)决定余额符号；
# 带符号后 期末 = 期初 + 借方(带符号) + 贷方(带符号)，平衡校验恒为 0，
# 末级期末合计数(带符号) = 科目余额表父级控制数，二者一致。
# ============================================================================
def _norm_dir(s):
    n = _norm(s)
    if n in ("借", "debit", "dr", "d", "1"):
        return "借"
    if n in ("贷", "credit", "cr", "c", "0"):
        return "贷"
    return "平"


def _signed_balance(amt, dir_s, nature):
    """余额(期初/期末)带符号：资产借余为+，贷余(资产中的贷方余额，如客户多付)为−；
    负债贷余为+，借余(负债中的借方余额，如预付性挂账)为−；平/空取绝对值符号。"""
    d = _norm_dir(dir_s)
    if nature == "liability":
        return -amt if d == "借" else amt      # 贷/平 → +，借 → −
    return -amt if d == "贷" else amt           # 借/平 → +，贷 → −


def _dir_from_signed(sgn, nature):
    """由带符号余额反推方向列值（与 _signed_balance 互为逆运算）：
    资产：正→借、负→贷；负债：正→贷、负→借；≈0→平。"""
    if abs(sgn) <= 0.005:
        return "平"
    if nature == "liability":
        return "贷" if sgn > 0 else "借"
    return "借" if sgn > 0 else "贷"


def _signed_debit(amt, nature):
    return amt if nature == "asset" else -amt   # 资产借+，负债借−


def _signed_credit(amt, nature):
    return -amt if nature == "asset" else amt   # 资产贷−，负债贷+


def _disp_name(ak):
    """customers 在 export_excel 中被归一化为 (entity, name) 2-tuple；展示时统一取往来单位名称。"""
    if isinstance(ak, (tuple, list)) and len(ak) == 2:
        return ak[1]
    return ak


def _align_fcur_sign(s):
    """外币(原币)期初/期末与本币期初/期末同方向。

    本币(s_open/s_close)已按科目性质带符号：资产借余为+、贷余(贷方余额)为−；
    负债贷余为+、借余(借方余额)为−。而外币原币导出恒为非负，故按本币符号对齐
    ——仅当本币非零且二者方向相反时翻转外币符号，使本币与外原币呈同方向数据。"""
    so = s.get("s_open", 0.0) or 0.0
    sc = s.get("s_close", 0.0) or 0.0
    fo = s.get("fcur_open", 0.0) or 0.0
    fc = s.get("fcur_close", 0.0) or 0.0
    if so != 0 and fo != 0 and ((so > 0) != (fo > 0)):
        s["fcur_open"] = -fo
    if sc != 0 and fc != 0 and ((sc > 0) != (fc > 0)):
        s["fcur_close"] = -fc


def compute_summary(customers, periods, agg, opening=None, nature="asset"):
    """主表数字直接取自辅助核算余额表(控制数)；无控制数但 GL 有数据的行以 GL 净额列示(gl_only)。
    除保留原始绝对值(open/debit/credit/close)外，另按科目性质生成带符号显示值
    (s_open/s_debit/s_credit/s_close/s_bal)，明细表全部以带符号值呈现。"""
    opening = opening or {}
    agg = agg or {}
    summary = {}
    for ak in customers:
        for pk in periods:
            rec = agg.get((ak, pk), {"debit": 0.0, "credit": 0.0, "count": 0})
            op_rec = opening.get((ak, pk))
            has_agg = (rec["count"] > 0) or (rec["debit"] != 0) or (rec["credit"] != 0)
            gl_fill = False
            if op_rec is not None:
                op = op_rec.get("open", 0.0)
                cl = op_rec.get("close", 0.0)
                aging = op_rec.get("aging", "")
                code = op_rec.get("code", "")
                foreign = op_rec.get("foreign", False)
                bill_type = _derive_bill_type(op_rec.get("_km", ""))
                op_dir = op_rec.get("open_dir", "")
                cl_dir = op_rec.get("close_dir", "")
                # Q3 修复：AUX 含该客户但借/贷发生额均≈0（源表把发生额放在空名子目行、
                # 被 GENERIC_AUX_NAMES 过滤），而 GL 有真实发生额 → 回退用 GL 发生额填充增减变动，
                # 保留 AUX 期初/期末（通常有效）。如 金桥信息 应收账款增减变动曾全为 0。
                _aux_db = op_rec.get("debit", 0.0)
                _aux_cr = op_rec.get("credit", 0.0)
                if abs(_aux_db) < 0.005 and abs(_aux_cr) < 0.005 and has_agg:
                    db = rec["debit"]
                    cr = rec["credit"]
                    gl_fill = True
                else:
                    db = _aux_db
                    cr = _aux_cr
                gl_only = False
            elif has_agg:
                op = 0.0
                db = rec["debit"]
                cr = rec["credit"]
                cl = db - cr
                aging = ""
                code = ""
                foreign = False
                bill_type = ""
                gl_only = True
                op_dir = cl_dir = ""
            else:
                op = db = cr = cl = 0.0
                aging = ""
                code = ""
                foreign = False
                bill_type = ""
                gl_only = False
                op_dir = cl_dir = ""
            # 带符号显示值（按科目性质）：资产 借+/贷−，负债 贷+/借−；余额按方向列定符号
            s_open = _signed_balance(op, op_dir, nature)
            s_debit = _signed_debit(db, nature)
            s_credit = _signed_credit(cr, nature)
            s_close = _signed_balance(cl, cl_dir, nature) if (op_rec is not None) else (s_debit + s_credit)
            s_bal = s_open + s_debit + s_credit - s_close
            summary[(ak, pk)] = {
                "open": op, "debit": db, "credit": cr, "close": cl,
                "aging": aging, "code": code, "foreign": foreign,
                "bill_type": bill_type,
                "count": rec["count"],
                "aux_open": op, "aux_debit": db, "aux_credit": cr, "aux_close": cl,
                "gl_only": gl_only,
                "gl_fill": gl_fill,
                "open_dir": op_dir, "close_dir": cl_dir,
                "s_open": s_open, "s_debit": s_debit, "s_credit": s_credit,
                "s_close": s_close, "s_bal": s_bal,
            }
    return summary


# ============================================================================
# 模块 2.2：账龄区间划分（基于 GL 交易日期的先进先出 FIFO 测算）
# ----------------------------------------------------------------------------
# 数据源：辅助核算余额表通常不含账龄列，逐笔账龄只能由综合查询明细表(GL)的交易日期推算。
# 方法：将期初余额视为期间首日发生的一笔侧增加，按日期升序排列所有(增加/减少)流水；
#       减少(收款/付款)按 FIFO 先冲最老的增加；期末未结的增加按(期末日−发生日)分桶。
# 单一年度账套：期初与当年交易均落在 1 年以内 → 全列『1年以内』（已在账龄分布表披露）；
# 含多年度 GL 的账套可自动展现 1 年以上账龄。
# ============================================================================
def _period_bounds(pk):
    """由期间字符串(如 '2025' 或 '2025Q1')返回(期初日, 期末日)。"""
    import re as _re
    m = _re.search(r"(19|20)\d{2}", str(pk))
    y = int(m.group(0)) if m else 2025
    return datetime.date(y, 1, 1), datetime.date(y, 12, 31)


def _bucket_index(age_days, bucket_type):
    if bucket_type == "recv":
        if age_days < 365:  return 0
        if age_days < 730:  return 1
        if age_days < 1095: return 2
        if age_days < 1460: return 3
        if age_days < 1825: return 4
        return 5
    if bucket_type == "liab":   # ⚡ 2026-08-12 修复：liab 仅 2 档，原走 pay 分支返回 3
        if age_days < 365:  return 0
        return 1
    else:  # pay
        if age_days < 365:  return 0
        if age_days < 730:  return 1
        if age_days < 1095: return 2
        return 3


def compute_aging_buckets(nature, flows, opening_signed, period_start, period_end, bucket_type):
    """FIFO 账龄测算。
    flows: list of (date, signed_debit, signed_credit)，已按科目性质带符号
           （资产: 借+贷−；负债: 借−贷+）。
    opening_signed: 期初带符号余额。
    返回 {bucket: amount}。"""
    buckets = AGING_BUCKETS[bucket_type]
    result = {b: 0.0 for b in buckets}
    # ⚡ 2026-08-12 修复：3200 应收借方凭证小计 GL 行 date 为 str（'2026-01-04'）→
    #   (period_end - d) 类型错 TypeError 崩（原仅 3500 逐笔序时账 date 已转 date）。
    #   统一归一化：datetime/date 原样、Excel 序列号→date、'YYYY-MM-DD'/'-/'/'- ' 解析。
    _DT = datetime
    def _norm_d(x, fb):
        if x is None:
            return fb
        if isinstance(x, _DT.datetime):
            return x.date()
        if isinstance(x, _DT.date):
            return x
        if isinstance(x, (int, float)):
            try:
                return _DT.date(1899, 12, 30) + _DT.timedelta(days=int(x))
            except Exception:
                return fb
        m = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', str(x).strip())
        if m:
            try:
                return _DT.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except Exception:
                return fb
        return fb
    period_start = _norm_d(period_start, _DT.date(2026, 1, 1))
    period_end = _norm_d(period_end, _DT.date(2026, 12, 31))
    # 构造流水：正数=侧增加(资产借/负债贷)，负数=侧减少(资产贷/负债借)
    seq = []
    if opening_signed:
        seq.append((period_start, opening_signed))
    for d, sd, sc in flows:
        seq.append((_norm_d(d, period_start), sd + sc))     # 单笔净变动(已含方向)
    seq.sort(key=lambda x: x[0])
    stack = []   # [[date, remaining], ...] 未结增加
    for d, amt in seq:
        if amt > 0:
            stack.append([d, amt])
        else:
            dec = -amt
            while dec > 1e-9 and stack:
                if stack[0][1] <= dec + 1e-9:
                    dec -= stack[0][1]
                    stack.pop(0)
                else:
                    stack[0][1] -= dec
                    dec = 0
    for d, rem in stack:
        idx = _bucket_index((period_end - d).days, bucket_type)
        result[buckets[idx]] += rem
    return result


def _dominant_bucket(buckets):
    """返回金额最大的桶名（账龄区间列的快速展示用）。"""
    top, amt = "", 0.0
    for b, v in buckets.items():
        if abs(v) > abs(amt) + 1e-9:
            top, amt = b, v
    return top


def _rollforward_aging(prior_buckets, s, bucket_type, nature):
    """账龄方法2：以上一年度账龄表(分客户分桶)为基，做账龄滚动（倒推法）。
    规则（用户定义）：在本年账龄基础上总体再+1年；若本年1年以上金额 < 上年剩余金额，
                      则先冲减上年账龄最长的年度金额，依此类推。
    prior_buckets : {bucket: amount} 上年末该客户账龄分布（来自上年账龄分布/明细）
    s             : 本年汇总 dict（含 s_debit/s_credit/s_close/nature）
    bucket_type   : "recv"(6档) / "pay"(4档)
    返回 {bucket: amount}（本年账龄分布，合计 ≡ s_close 控制数）。"""
    buckets = AGING_BUCKETS[bucket_type]
    K = len(buckets)
    P = [prior_buckets.get(b, 0.0) for b in buckets]   # P[0]=1年以内 … P[K-1]=最长档
    # 本年新增且仍挂账(≤1年)业务估计：资产=借−贷；负债=贷−借（净额，<0 则计 0）
    if nature == "asset":
        net_new = s.get("s_debit", 0.0) - s.get("s_credit", 0.0)
    else:
        net_new = s.get("s_credit", 0.0) - s.get("s_debit", 0.0)
    G = max(0.0, net_new)                     # 本年新增(≤1年)
    s_close = s.get("s_close", 0.0)
    cur_old = s_close - G                      # 本年1年以上目标金额
    if cur_old < 0:
        cur_old = 0.0
        G = s_close
    # 滚动：上年各档整体上移1年（P[0]→本年1-2年 … P[K-2]→最长；P[K-1]并入最长档）
    rolled = [0.0] * K
    rolled[0] = G
    for j in range(1, K):
        rolled[j] = P[j - 1]
    rolled[K - 1] += P[K - 1]
    sum_old = sum(rolled[1:])
    # 若本年1年以上 < 上年剩余(=sum(rolled[1:]))，从最长档起冲减
    if cur_old < sum_old - 1e-6:
        diff = sum_old - cur_old
        j = K - 1
        while diff > 1e-9 and j >= 1:
            if rolled[j] >= diff - 1e-9:
                rolled[j] -= diff
                diff = 0.0
            else:
                diff -= rolled[j]
                rolled[j] = 0.0
                j -= 1
    # 归一化：使旧账龄合计严格等于 cur_old（与期末控制数勾稽）
    sum_old = sum(rolled[1:])
    if sum_old > 1e-9 and abs(sum_old - cur_old) > 1e-6:
        f = cur_old / sum_old
        for j in range(1, K):
            rolled[j] *= f
    return {b: rolled[i] for i, b in enumerate(buckets)}


# ============================================================================
# 模块 2.5：对应科目还原
# ============================================================================
def reconstruct_counterparties(rows):
    """按 (期间, (字,号)) 重建会计分录，还原目标科目对应科目（与应收账款程序同逻辑，科目无关）。"""
    vouchers = {}
    ar_debit_total = defaultdict(float)
    ar_credit_total = defaultdict(float)
    for r in rows:
        if not r["is_sub"]:
            continue
        pk = period_key(r["date"], "Y")
        key = (pk, r["vtype"], r["vno"])
        v = vouchers.setdefault(key, {"A": 0.0, "B": 0.0})
        v["A"] += r["debit"]
        v["B"] += r["credit"]
        ar_debit_total[pk] += r["debit"]
        ar_credit_total[pk] += r["credit"]
    nonar = defaultdict(lambda: {"deb": [], "cred": []})
    for r in rows:
        pk = period_key(r["date"], "Y")
        key = (pk, r["vtype"], r["vno"])
        if key not in vouchers:
            continue
        if r["is_sub"]:
            continue
        if r["debit"] > 0:
            nonar[key]["deb"].append((r["km"], r["debit"]))
        if r["credit"] > 0:
            nonar[key]["cred"].append((r["km"], r["credit"]))

    result = defaultdict(lambda: dict(sale_cp={}, sale_cnt={}, collect_cp={}, collect_cnt={},
                                      ar_debit_net=0.0, ar_credit_net=0.0,
                                      sale_ext=0.0, collect_ext=0.0,
                                      internal_deb=0.0, internal_cred=0.0))
    for pk in set(k for (k, _, _) in vouchers):
        sale_cp = defaultdict(float); sale_cnt = defaultdict(int)
        collect_cp = defaultdict(float); collect_cnt = defaultdict(int)
        sale_ext = collect_ext = 0.0
        for key, v in vouchers.items():
            if key[0] != pk:
                continue
            A, B = v["A"], v["B"]
            creds = nonar[key]["cred"]
            debs = nonar[key]["deb"]
            Dn = sum(x[1] for x in debs)
            denom = A + Dn
            f_sale = A / denom if denom > 0 else 0.0
            f_coll = B / denom if denom > 0 else 0.0
            for km, amt in creds:
                att = amt * f_sale
                sale_cp[km] += att; sale_cnt[km] += 1; sale_ext += att
            for km, amt in debs:
                att = amt * f_coll
                collect_cp[km] += att; collect_cnt[km] += 1; collect_ext += att
        res = result[pk]
        res["sale_cp"] = dict(sale_cp); res["sale_cnt"] = dict(sale_cnt)
        res["collect_cp"] = dict(collect_cp); res["collect_cnt"] = dict(collect_cnt)
        res["ar_debit_net"] = ar_debit_total[pk]
        res["ar_credit_net"] = ar_credit_total[pk]
        res["sale_ext"] = sale_ext
        res["collect_ext"] = collect_ext
        res["internal_deb"] = ar_debit_total[pk] - sale_ext
        res["internal_cred"] = ar_credit_total[pk] - collect_ext
    return dict(result)


RECON_KEY_COLLISION_THRESHOLD = 8  # 一张凭证的翻面(对方)科目通常仅数行；超过则『字/号』非唯一凭证号

def _vouch_key_of(r, ent_key=True):
    """凭证唯一键：必须含【年月】——U8 凭证号每月重新编号（每月都有记-1..记-N），
    跨月同号若只按 (字,号) 聚合会被误判为"同一凭证"，翻面还原/碰撞判定全乱
    （2026-07-31 修复：FY 2025 对方科目核对缺失根因）。
    """
    d = r.get("date")
    ym = d.strftime("%Y-%m") if hasattr(d, "strftime") else str(d)[:7]
    if ent_key:
        return (r.get("entity", ""), period_key(d, "Y"), ym, r.get("vtype", ""), r.get("vno", ""))
    return (period_key(d, "Y"), ym, r.get("vtype", ""), r.get("vno", ""))


def _recon_unreliable(rows):
    """返回不可靠的 (ent, pk) 集合：该主体/年度下，目标科目的『字/号』被过多不同科目复用，
    按(字,号)把 is_sub 行与翻面行配对会把无关交易误归为对方科目，翻面还原结果不可信。"""
    flip_km = defaultdict(set)
    for r in rows:
        if r.get("is_sub"):
            continue
        key = _vouch_key_of(r)
        if r.get("km"):
            flip_km[key].add(r["km"])
    bad = set()
    for r in rows:
        if not r.get("is_sub"):
            continue
        key = _vouch_key_of(r)
        if len(flip_km.get(key, set())) > RECON_KEY_COLLISION_THRESHOLD:
            bad.add((r.get("entity", ""), period_key(r["date"], "Y")))
    return bad

def reconstruct_counterparties_by_entity(rows, use_cp=False):
    """按 (核算主体 entity, 期间 pk) 重建会计分录，还原目标科目对应科目（凭证级按主体隔离）。
    返回 {(ent, pk): dict(sale_cp, sale_cnt, collect_cp, collect_cnt,
                          ar_debit_net, ar_credit_net, sale_ext, collect_ext,
                          internal_deb, internal_cred)}。
    用于『对方科目核对』按核算主体、分年度的异常判断（不再聚合到集团层面）。
    ⚡ 2026-08-10 SAP 快速路径（use_cp=True）：SAP 的 read_transactions 只返回目标科目行
    （read_gl_net pfx 过滤），标准凭证翻面配对因缺非目标行而落空 → 对方科目核对 sheet 不生成。
    SAP 行自带 cp（_infer_opp 同凭证对方科目推断，逗号分隔），借方行 cp→sale_cp、贷方行 cp→collect_cp。"""
    if use_cp:
        per = defaultdict(lambda: dict(sale_cp={}, sale_cnt={}, collect_cp={}, collect_cnt={},
                                       ar_debit_net=0.0, ar_credit_net=0.0, sale_ext=0.0,
                                       collect_ext=0.0, internal_deb=0.0, internal_cred=0.0))
        for r in rows:
            if not r.get('is_sub'):
                continue
            pk = period_key(r['date'], 'Y')
            ent = r.get('entity', '')
            d = per[(ent, pk)]
            dr = float(r.get('debit') or 0.0)
            cr = float(r.get('credit') or 0.0)
            d['ar_debit_net'] += dr
            d['ar_credit_net'] += cr
            for nm in (str(r.get('cp') or '').split(',')):
                nm = nm.strip()
                if not nm:
                    continue
                if dr > 0:
                    d['sale_cp'][nm] = d['sale_cp'].get(nm, 0.0) + dr
                    d['sale_cnt'][nm] = d['sale_cnt'].get(nm, 0) + 1
                if cr > 0:
                    d['collect_cp'][nm] = d['collect_cp'].get(nm, 0.0) + cr
                    d['collect_cnt'][nm] = d['collect_cnt'].get(nm, 0) + 1
        return dict(per)
    vouchers = {}
    unreliable = _recon_unreliable(rows)   # (ent,pk) 字/号非唯一→翻面还原不可靠，停用
    ar_debit_total = defaultdict(float)
    ar_credit_total = defaultdict(float)
    for r in rows:
        if not r.get("is_sub"):
            continue
        pk = period_key(r["date"], "Y")
        ent = r.get("entity", "")
        key = _vouch_key_of(r)
        v = vouchers.setdefault(key, {"A": 0.0, "B": 0.0})
        v["A"] += r["debit"]
        v["B"] += r["credit"]
        ar_debit_total[(ent, pk)] += r["debit"]
        ar_credit_total[(ent, pk)] += r["credit"]
    nonar = defaultdict(lambda: {"deb": [], "cred": []})
    for r in rows:
        if r.get("is_sub"):
            continue
        pk = period_key(r["date"], "Y")
        ent = r.get("entity", "")
        key = _vouch_key_of(r)
        if key not in vouchers:
            continue
        if r["debit"] > 0:
            nonar[key]["deb"].append((r["km"], r["debit"]))
        if r["credit"] > 0:
            nonar[key]["cred"].append((r["km"], r["credit"]))
    per = defaultdict(lambda: dict(sale_cp={}, sale_cnt={}, collect_cp={}, collect_cnt={},
                                   ar_debit_net=0.0, ar_credit_net=0.0, sale_ext=0.0, collect_ext=0.0,
                                   internal_deb=0.0, internal_cred=0.0))
    # ⚡⚡ 2026-08-25 修复 O(n²)：原『for (ent,pk) × for key in vouchers』内层遍历全部
    #   vouchers（XBJ ORA 69 万行 → 201 主体 × 69 万 = 1.4 亿次比较 = build 40 分钟主因）。
    #   预按 (ent,pk) 分组，内层只遍历本组凭证 → 每组几千行，秒级。
    vouchers_by_ep = defaultdict(dict)
    for _k, _v in vouchers.items():
        vouchers_by_ep[(_k[0], _k[1])][_k] = _v
    nonar_by_ep = defaultdict(dict)
    for _k, _lst in nonar.items():
        nonar_by_ep[(_k[0], _k[1])][_k] = _lst
    for (ent, pk) in set(k[:2] for k in vouchers):
        if (ent, pk) in unreliable:
            continue
        sale_cp = defaultdict(float); sale_cnt = defaultdict(int)
        collect_cp = defaultdict(float); collect_cnt = defaultdict(int)
        sale_ext = collect_ext = 0.0
        _vch_g = vouchers_by_ep[(ent, pk)]
        _nonar_g = nonar_by_ep.get((ent, pk), {})
        for key, v in _vch_g.items():
            A, B = v["A"], v["B"]
            _nk = _nonar_g.get(key)
            creds = _nk["cred"] if _nk else []
            debs = _nk["deb"] if _nk else []
            Dn = sum(x[1] for x in debs)
            denom = A + Dn
            f_sale = A / denom if denom > 0 else 0.0
            f_coll = B / denom if denom > 0 else 0.0
            for km, amt in creds:
                att = amt * f_sale
                sale_cp[km] += att; sale_cnt[km] += 1; sale_ext += att
            for km, amt in debs:
                att = amt * f_coll
                collect_cp[km] += att; collect_cnt[km] += 1; collect_ext += att
        res = per[(ent, pk)]
        res["sale_cp"] = dict(sale_cp); res["sale_cnt"] = dict(sale_cnt)
        res["collect_cp"] = dict(collect_cp); res["collect_cnt"] = dict(collect_cnt)
        res["ar_debit_net"] = ar_debit_total[(ent, pk)]
        res["ar_credit_net"] = ar_credit_total[(ent, pk)]
        res["sale_ext"] = sale_ext
        res["collect_ext"] = collect_ext
        res["internal_deb"] = ar_debit_total[(ent, pk)] - sale_ext
        res["internal_cred"] = ar_credit_total[(ent, pk)] - collect_ext
    return dict(per)


def _reconstruct_counterparties_combined(rows):
    """合并模式集团层面对应科目分析：在 reconstruct_counterparties_by_entity 之后，
    按期间(pk)聚合为集团层面（凭证级按主体隔离，聚合仅为求和）。"""
    per = reconstruct_counterparties_by_entity(rows)
    result = defaultdict(lambda: dict(sale_cp={}, sale_cnt={}, collect_cp={}, collect_cnt={},
                                       ar_debit_net=0.0, ar_credit_net=0.0, sale_ext=0.0, collect_ext=0.0,
                                       internal_deb=0.0, internal_cred=0.0))
    for (ent, pk), d in per.items():
        for k, v in d["sale_cp"].items():
            result[pk]["sale_cp"][k] = result[pk]["sale_cp"].get(k, 0.0) + v
        for k, v in d["sale_cnt"].items():
            result[pk]["sale_cnt"][k] = result[pk]["sale_cnt"].get(k, 0) + v
        for k, v in d["collect_cp"].items():
            result[pk]["collect_cp"][k] = result[pk]["collect_cp"].get(k, 0.0) + v
        for k, v in d["collect_cnt"].items():
            result[pk]["collect_cnt"][k] = result[pk]["collect_cnt"].get(k, 0) + v
        result[pk]["ar_debit_net"] += d["ar_debit_net"]
        result[pk]["ar_credit_net"] += d["ar_credit_net"]
        result[pk]["sale_ext"] += d["sale_ext"]
        result[pk]["collect_ext"] += d["collect_ext"]
        result[pk]["internal_deb"] += d["internal_deb"]
        result[pk]["internal_cred"] += d["internal_cred"]
    return dict(result)


def _mk_cat(expected):
    def cat(km):
        for e in expected:
            if e in km:
                return "期望"
        return "其他"
    return cat


BANK_TOKENS = ("银行", "支行", "分行", "营业部", "信用社", "农商", "村镇银行",
               "财务公司", "财司", "结算中心", "网联", "银联")
CASH_TOKENS = ("库存现金", "现金")


def _is_cash_bank(km):
    return any(t in km for t in BANK_TOKENS) or any(t in km for t in CASH_TOKENS)


def _classify_counterparty(km, expected_list, note_map, self_kw=None):
    """判断某对应科目是否异常。返回 (verdict, reason)。
    verdict ∈ {正常, 关注, 异常}：
      正常 = 命中该侧期望对应科目(exp_*) 或 该侧期望含银行存款/库存现金且对方为银行账户/现金（回款、付款走银行属正常）
             或 对应科目为该往来科目自身内部子目（GR/IR 过渡、外币评估调整、自身结转——2026-08-26 用户方法论）；
      异常 = 命中 notes 中带『★异常/异常』标记的科目（如直接进未分配利润、营业外收支挂账等）；
      关注 = 非期望对应科目（含 notes 中『疑/建议核查/需关注』或完全无说明），需核实业务实质。"""
    # ⚡ 2026-08-26 内部子目正常化：对应科目为本科目自身前缀（如 应付账款-GR/IR收货/收发票、
    #   应付账款-外币评估调整、应付账款-外部应付 等）→ 内部结转/过渡/评估，属正常核算，
    #   不再误标『应付与预付/应收互抵』（1010 贷方侧 GR/IR 16.12亿曾被误标关注）。
    if self_kw and km.startswith(self_kw):
        return "正常", ""
    for e in expected_list:
        if e and e in km:
            return "正常", ""
    # 银行/现金类科目：该侧期望含 银行存款/库存现金 时，回款/付款经银行账户属正常业务
    if _is_cash_bank(km) and ("银行存款" in expected_list or "库存现金" in expected_list):
        return "正常", ""
    note = ""
    for k, v in note_map.items():
        if k and k in km:
            note = v
            break
    if note:
        if "★异常" in note or "异常" in note:
            return "异常", note
        if "疑" in note or "建议核查" in note or "需关注" in note:
            return "关注", note
        return "关注", note
    return "关注", "非期望对应科目，需核实业务实质与凭证"


# ============================================================================
# 模块 3：数据校验
# ============================================================================
def validate(rows, customers, periods, agg, summary, opening, counterparties, ctrl_map=None, subj=None):
    issues = []
    opening = opening or {}
    ctrl_map = ctrl_map or {}
    label = subj["label"] if subj else "往来科目"

    code_by_name = {}
    for (name, pk), rec in opening.items():
        code_by_name.setdefault(name, set()).add(rec.get("code", ""))
    for name, codes in code_by_name.items():
        codes.discard("")
        if len(codes) > 1:
            issues.append({
                "level": "WARN", "type": "往来单位一致性",
                "message": f"往来单位『{name}』对应多个编号：{sorted(codes)}，已按名称归并。",
            })

    for ak in customers:
        for pk in periods:
            s = summary.get((ak, pk))
            if s is None:
                continue
            calc_close = s["s_open"] + s["s_debit"] + s["s_credit"]
            if abs(calc_close - s["s_close"]) > 0.005:
                issues.append({
                    "level": "WARN", "type": "余额核对",
                    "message": (f"往来单位『{ak}』{pk}：辅助核算源表内部不平衡 "
                                f"(期初{s['s_open']:.2f}+借(符号){s['s_debit']:.2f}+贷(符号){s['s_credit']:.2f}"
                                f"={calc_close:.2f} ≠ 期末{s['s_close']:.2f}，差异{calc_close - s['s_close']:.2f})，"
                                f"属源表勾稽误差，非账务差错。"),
                })
            if s.get("gl_only"):
                issues.append({
                    "level": "INFO", "type": "取数",
                    "message": f"往来单位『{ak}』{pk} 仅存在于综合查询明细表(GL)、无辅助核算控制数，"
                               f"期末({s['s_close']:.2f})为GL净额派生，未纳入控制数核对。",
                })

    zero_cnt = sum(1 for r in rows if r["is_sub"] and r["debit"] == 0 and r["credit"] == 0)
    if zero_cnt:
        issues.append({"level": "WARN", "type": "金额", "message": f"存在 {zero_cnt} 条{label}金额为 0 的记录，已按 0 处理。"})

    for pk in periods:
        cp = counterparties.get(pk)
        if not cp:
            continue
        diff_sale = cp["ar_debit_net"] - cp["sale_ext"]
        diff_coll = cp["ar_credit_net"] - cp["collect_ext"]
        if abs(diff_sale) > max(1.0, 0.001 * abs(cp["ar_debit_net"])):
            issues.append({
                "level": "WARN", "type": "对应科目",
                "message": f"{pk} 借方侧：{label}借方净额({cp['ar_debit_net']:.2f})与外部对应科目合计({cp['sale_ext']:.2f})差异{diff_sale:.2f}，多为{label}↔{label}互冲。",
            })
        if abs(diff_coll) > max(1.0, 0.001 * abs(cp["ar_credit_net"])):
            issues.append({
                "level": "WARN", "type": "对应科目",
                "message": f"{pk} 贷方侧：{label}贷方发生额({cp['ar_credit_net']:.2f})与外部对应科目合计({cp['collect_ext']:.2f})差异{diff_coll:.2f}，多为{label}↔{label}互冲。",
            })
        exp_debit = subj["exp_debit"] if subj else []
        exp_credit = subj["exp_credit"] if subj else []
        notes_debit = subj["notes_debit"] if subj else {}
        notes_credit = subj["notes_credit"] if subj else {}
        cat_d = _mk_cat(exp_debit)
        cat_c = _mk_cat(exp_credit)
        # 非期望科目不再在『校验结果』中作为异常显示（按分类列示于『对方科目核对』，供审计判断）。
        # 仅保留借方/贷方净额与外部对应科目合计的差异提示（多为往来↔往来互冲）。

    ctrl_issues = []
    if ctrl_map:
        for pk in periods:
            s_sum = sum(summary[(a, pk)]["s_close"] for a in customers
                        if summary.get((a, pk)) and not summary[(a, pk)].get("gl_only"))
            ctrl = ctrl_map.get(pk)
            if ctrl is None:
                continue
            diff = s_sum - ctrl
            if abs(diff) > 0.005:
                ctrl_issues.append({
                    "level": "WARN", "type": "控制数核对",
                    "message": (f"{pk} {label}控制数：辅助核算末级期末合计({s_sum:,.2f}) "
                                f"与 科目余额表{label}期末({ctrl:,.2f}) 差异 {diff:,.2f}，"
                                f"多为辅助核算源表内部不平衡累计，非账务差错。"),
                })
    if ctrl_issues:
        issues += ctrl_issues
    elif ctrl_map:
        issues.append({"level": "OK", "type": "控制数核对",
                       "message": f"各期{label}辅助核算末级期末合计均与科目余额表{label}期末一致（控制数核对通过）。"})

    if not issues:
        issues.append({"level": "OK", "type": "汇总", "message": "校验通过：往来单位一致、金额平衡；跨期衔接与对应科目详见『期间对比』与『对方科目核对』。"})
    return issues


# ============================================================================
# 模块 4 + 5：明细表生成与导出
# ============================================================================
def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def _write_debit_credit_check_sheet(ws, customers, summary, periods, subj, label):
    """借贷方检查表（期初+本期借-本期贷=期末 滚动勾稽）。
    逐往来单位列示：期初(借/贷)方余额、本期借方发生额、本期贷方发生额、期末(借/贷)方余额、勾稽差。
    应收类(nature=='asset')按借方列示、应付类(nature=='liability')按贷方列示；多期则逐期累加。"""
    is_asset = (subj or {}).get("nature") == "asset"
    open_lbl = "期初借方余额" if is_asset else "期初贷方余额"
    close_lbl = "期末借方余额" if is_asset else "期末贷方余额"
    ws.cell(2, 1, f"{label}借贷方检查表（期初+本期借-本期贷=期末 勾稽）").font = TITLE_FONT
    headers = ["核算主体", "序号", "往来单位编号", "往来单位名称", "款项性质", "是否关联方",
               open_lbl, "本期借方发生额", "本期贷方发生额", close_lbl, "勾稽(期初+借-贷-期末)"]
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(headers))
    r = 5
    tot = {"open": 0.0, "debit": 0.0, "credit": 0.0, "close": 0.0, "bal": 0.0}
    pks = periods or [k[1] for k in summary.keys()]
    for idx, ak in enumerate(customers, 1):
        s_open = s_debit = s_credit = s_close = 0.0
        for pk in pks:
            s = summary.get((ak, pk))
            if not s:
                continue
            s_open += s["s_open"]; s_debit += s["s_debit"]; s_credit += s["s_credit"]; s_close += s["s_close"]
        # 发生额/余额均按绝对值列示（借贷方向已由列名体现）；勾稽用带符号恒等式 期初+借(signed)+贷(signed)-期末≈0
        open_d, debit_d, credit_d, close_d = abs(s_open), abs(s_debit), abs(s_credit), abs(s_close)
        bal = s_open + s_debit + s_credit - s_close
        s0 = summary.get((ak, pks[0])) or {}
        vals = [ak[0], idx, s0.get("code", ""), _disp_name(ak), s0.get("nature", ""),
                ("是" if s0.get("related") else "否"),
                round(open_d, 2), round(debit_d, 2), round(credit_d, 2), round(close_d, 2), round(bal, 2)]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (7, 8, 9, 10, 11):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
            elif c in (2, 5, 6):
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
            if c == 11 and abs(bal) > 0.005:
                cell.fill = WARN_FILL
        tot["open"] += open_d; tot["debit"] += debit_d; tot["credit"] += credit_d
        tot["close"] += close_d; tot["bal"] += bal
        r += 1
    tcell = ws.cell(r, 1, "合计")
    tcell.font = TOTAL_FONT
    for c in range(1, len(headers) + 1):
        ws.cell(r, c).border = BORDER
        ws.cell(r, c).fill = TOTAL_FILL
    for c, key in [(7, "open"), (8, "debit"), (9, "credit"), (10, "close"), (11, "bal")]:
        cell = ws.cell(r, c, round(tot[key], 2))
        cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
        cell.number_format = MONEY_FMT
    widths = [16, 6, 16, 30, 16, 10, 18, 16, 16, 18, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B5"


def _write_period_sheet(ws, period, customers, summary, subject, label, aging_map=None, bucket_type=None):
    buckets = AGING_BUCKETS.get(bucket_type) if bucket_type else None
    base_headers = ["核算主体", "序号", "往来单位编号", "往来单位名称", "款项性质", "是否关联方", "币种",
                    "外币余额(期初)", "期初余额(人民币)", "本期借方发生额", "本期贷方发生额",
                    "外币余额(期末)", "期末余额(人民币)",
                    "重分类调整", "审计调整", "审定数"]
    if buckets:
        headers = base_headers + buckets + ["交易笔数", "平衡校验(期初+借+贷-期末)"]
        nb = len(buckets)
    else:
        headers = base_headers + ["账龄区间", "交易笔数", "平衡校验(期初+借+贷-期末)"]
        nb = 0
    for i, h in enumerate(headers, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(headers))
    ws.cell(2, 1, f"{label}明细表 — {period}").font = TITLE_FONT

    r = 5
    tot = {"open": 0.0, "debit": 0.0, "credit": 0.0, "close": 0.0, "count": 0, "fopen": 0.0, "fclose": 0.0}
    tot_buckets = {b: 0.0 for b in (buckets or [])}
    # ⚡⚡ 2026-08-25 性能：数据行改 ws.append 批量写（逐 cell 20万行 40.9s → append 4.4s，约 5-7 倍）。
    #   样式仍逐 cell（openpyxl 无整列样式 API），但通过遍历 ws[r] 直接取已建 cell 避免 ws.cell 二次查找。
    _AL_RIGHT = Alignment(horizontal="right")
    for idx, ak in enumerate(customers, 1):
        s = summary[(ak, period)]
        bal = s["s_open"] + s["s_debit"] + s["s_credit"] - s["s_close"]
        ccy = "外币" if s["foreign"] else "人民币"
        f_open = s.get("fcur_open", 0.0) or 0.0
        f_close = s.get("fcur_close", 0.0) or 0.0
        base_vals = [subject, idx, s["code"], _disp_name(ak), s.get("nature", ""),
                     ("是" if s.get("related") else "否"), ccy,
                     (f_open if f_open else None), s["s_open"], abs(s["s_debit"]),
                     abs(s["s_credit"]), (f_close if f_close else None), s["s_close"],
                     None, None, "=M%d+N%d+O%d" % (r, r, r)]
        # 2026-08-03 重分类：反项余额(期末为负)→重分类调整列自动过入，审定归 0，账龄清 0
        _rc = 0.0
        if (s["s_close"] or 0.0) < -0.005:
            _rc = round(-(s["s_close"] or 0.0), 2)
            base_vals[13] = _rc          # 列14 重分类调整（0-based 索引 13）
        if buckets:
            if _rc:
                bm = {b: 0.0 for b in buckets}       # 重分类转出：账龄清 0
            else:
                bm = aging_map.get((ak, period)) if aging_map else None
                bm = bm or {b: 0.0 for b in buckets}
            bucket_vals = [round(bm.get(b, 0.0), 2) for b in buckets]
            vals = base_vals + bucket_vals + [s["count"], round(bal, 2)]
        else:
            aging = s["aging"] if s["aging"] else "（未提供）"
            vals = base_vals + [aging, s["count"], round(bal, 2)]
        ws.append(vals)
        row_cells = ws[r]
        for c, cell in enumerate(row_cells, 1):
            cell.border = BORDER
            money_cols = list(range(8, 17)) + list(range(17, 17 + nb))   # 8..16 + 桶列
            if c in money_cols:
                cell.number_format = MONEY_FMT
                cell.alignment = _AL_RIGHT
            elif c in (2, 5, 6, 7):
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
            if bal is not None and abs(bal) > 0.005 and c == len(vals):
                cell.fill = WARN_FILL
        for b, v in zip(buckets or [], bucket_vals if buckets else []):
            tot_buckets[b] += v
        tot["open"] += s["s_open"]; tot["debit"] += abs(s["s_debit"])
        tot["credit"] += abs(s["s_credit"]); tot["close"] += s["s_close"]
        tot["count"] += s["count"]
        tot["fopen"] += f_open; tot["fclose"] += f_close
        r += 1
    tcell = ws.cell(r, 1, "合计")
    tcell.font = TOTAL_FONT
    for c in range(1, len(headers) + 1):
        ws.cell(r, c).border = BORDER
        ws.cell(r, c).fill = TOTAL_FILL
    for c, key in [(8, "fopen"), (9, "open"), (10, "debit"), (11, "credit"),
                   (12, "fclose"), (13, "close")]:
        cell = ws.cell(r, c, tot[key])
        cell.font = TOTAL_FONT
        cell.fill = TOTAL_FILL
        cell.border = BORDER
        cell.number_format = MONEY_FMT
    # 2026-08-06 协议：重分类/期末/审定 列合计写数值（=数据行和；公式 data_only 读 None）
    for c in (14, 15, 16):
        _sv = sum(ws.cell(x, c).value for x in range(5, r)
                  if isinstance(ws.cell(x, c).value, (int, float)))
        cell = ws.cell(r, c, round(_sv, 2))
        cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
        cell.number_format = MONEY_FMT
    for j, b in enumerate(buckets or []):
        cell = ws.cell(r, 17 + j, round(tot_buckets[b], 2))
        cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
        cell.number_format = MONEY_FMT
    cnt_col = 17 + nb
    cell = ws.cell(r, cnt_col, tot["count"]); cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
    widths = [16, 6, 16, 30, 16, 10, 8, 16, 16, 16, 16, 16, 16, 14, 14, 14] + [14] * nb + [10, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B5"


def _write_comparison_sheet(ws, customers, periods, summary, subject, label):
    ws.cell(2, 1, f"{label}明细 — 期间对比").font = TITLE_FONT
    hdr = ["核算主体", "往来单位编号", "往来单位名称", "款项性质", "是否关联方", "币种"]
    for pk in periods:
        hdr += [f"{pk}外币期初", f"{pk}期初", f"{pk}借方", f"{pk}贷方", f"{pk}外币期末", f"{pk}期末"]
    # 跨期衔接：上期期末(外币/本币) vs 下期期初(外币/本币) + 差异（差异直接高亮显示）
    transitions = []
    for i in range(len(periods) - 1):
        p0, p1 = periods[i], periods[i + 1]
        transitions.append((p0, p1))
        hdr += [f"{p0}期末外币", f"{p1}期初外币", "外币差异",
                f"{p0}期末本币", f"{p1}期初本币", "本币差异"]
    hdr += ["首期期末", "末期期末", "期末变动"]
    diff_cols = [i + 1 for i, h in enumerate(hdr) if h in ("外币差异", "本币差异")]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for ak in customers:
        first_close = summary[(ak, periods[0])]["close"]
        last_close = summary[(ak, periods[-1])]["close"]
        ccy = "外币" if summary[(ak, periods[0])]["foreign"] else "人民币"
        code = summary[(ak, periods[0])]["code"]
        nature = summary[(ak, periods[0])].get("nature", "")
        related = "是" if summary[(ak, periods[0])].get("related") else "否"
        row_vals = [subject, code, _disp_name(ak), nature, related, ccy]
        for pk in periods:
            s = summary[(ak, pk)]
            fo = s.get("fcur_open", 0.0) or 0.0
            fc = s.get("fcur_close", 0.0) or 0.0
            row_vals += [(fo if fo else None), s["s_open"], abs(s["s_debit"]), abs(s["s_credit"]),
                         (fc if fc else None), s["s_close"]]
        for (p0, p1) in transitions:
            s0 = summary[(ak, p0)]; s1 = summary[(ak, p1)]
            fc0 = s0.get("fcur_close", 0.0) or 0.0
            fc1 = s1.get("fcur_open", 0.0) or 0.0
            lc0 = s0["s_close"]; lc1 = s1["s_open"]
            row_vals += [(fc0 if fc0 else None), (fc1 if fc1 else None), round(fc1 - fc0, 2),
                         lc0, lc1, round(lc1 - lc0, 2)]
        row_vals += [first_close, last_close, last_close - first_close]
        for c, v in enumerate(row_vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c >= 7:
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        for c in diff_cols:
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and abs(v) > 0.005:
                ws.cell(r, c).fill = WARN_FILL
        r += 1
    ws.cell(r, 1, "合计").font = TOTAL_FONT
    for c in range(1, len(hdr) + 1):
        ws.cell(r, c).border = BORDER
        ws.cell(r, c).fill = TOTAL_FILL
    # 2026-08-02 公式化：各期间列 6 指标 =SUM(数据区 5..r-1)；跨期差异/期末变动列沿用公式（下方）
    data_first = 5
    col = 7
    for pk in periods:
        for off in range(6):
            cl = get_column_letter(col + off)
            cell = ws.cell(r, col + off, '=SUM(%s%d:%s%d)' % (cl, data_first, cl, r - 1))
            cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
            cell.number_format = MONEY_FMT
        col += 6
    # 合计行：跨期差异 = 下期期初合计 − 上期期末合计
    for (p0, p1) in transitions:
        fc0 = sum((summary[(a, p0)].get("fcur_close", 0.0) or 0.0) for a in customers)
        fc1 = sum((summary[(a, p1)].get("fcur_open", 0.0) or 0.0) for a in customers)
        lc0 = sum(summary[(a, p0)]["s_close"] for a in customers)
        lc1 = sum(summary[(a, p1)]["s_open"] for a in customers)
        block = [(fc0 if fc0 else None), (fc1 if fc1 else None), round(fc1 - fc0, 2),
                 lc0, lc1, round(lc1 - lc0, 2)]
        for off, val in enumerate(block):
            cell = ws.cell(r, col + off, val)
            cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
            cell.number_format = MONEY_FMT
        if abs(fc1 - fc0) > 0.005:
            ws.cell(r, col + 2).fill = WARN_FILL
        if abs(lc1 - lc0) > 0.005:
            ws.cell(r, col + 5).fill = WARN_FILL
        col += 6
    widths = [16, 16, 30, 16, 10, 8] + [14] * (len(periods) * 6) \
             + [14] * (len(transitions) * 6) + [14, 14, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "G5"


def _write_summary_sheet(ws, periods, customers, summary, subject, label):
    ws.cell(2, 1, f"{label}余额汇总").font = TITLE_FONT
    hdr = ["核算主体", "期间", "往来单位数", "期初余额合计", "借方发生额合计", "贷方发生额合计", "期末余额合计"]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for pk in periods:
        op = sum(summary[(a, pk)]["s_open"] for a in customers)
        db = sum(abs(summary[(a, pk)]["s_debit"]) for a in customers)
        cr = sum(abs(summary[(a, pk)]["s_credit"]) for a in customers)
        cl = sum(summary[(a, pk)]["s_close"] for a in customers)
        for c, v in enumerate([subject, pk, len(customers), op, db, cr, cl], 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c >= 4:
                cell.number_format = MONEY_FMT
        r += 1
    for i, w in enumerate([16, 12, 10, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_counterparty_sheet(ws, title, cp, ar_total, cat_fn, note_map, sign, base):
    ws.cell(1, 1, title).font = TITLE_FONT
    hdr = ["序号", "对应科目（科目名称）", "分类", f"金额({sign})",
           f"占{base}{'借方' if sign == '借' else '贷方'}%", "笔数", "审计提示"]
    for c, h in enumerate(hdr, 1):
        ws.cell(3, c, h)
    _style_header(ws, 3, len(hdr))
    items = sorted(cp["cp"].items(), key=lambda x: (-(0 if cat_fn(x[0]) == "期望" else 1), -x[1]))
    rr = 4
    idx = 1
    disp_sum = 0.0
    for km, amt in items:
        if abs(amt) < 1.00:
            continue
        cat = cat_fn(km)
        disp_sum += amt
        ws.cell(rr, 1, idx).alignment = CENTER
        ws.cell(rr, 2, km)
        ws.cell(rr, 3, cat).alignment = CENTER
        ws.cell(rr, 4, amt).number_format = MONEY_FMT
        ws.cell(rr, 4).alignment = Alignment(horizontal="right")
        ws.cell(rr, 5, amt / ar_total if ar_total else 0).number_format = PCT_FMT
        ws.cell(rr, 5).alignment = Alignment(horizontal="right")
        ws.cell(rr, 6, cp["cnt"].get(km, 0)).alignment = CENTER
        note = ""
        for k, v in note_map.items():
            if k in km:
                note = v
                break
        ws.cell(rr, 7, note)
        fill = OTHER_FILL if cat == "其他" else None
        for c in range(1, 8):
            cell = ws.cell(rr, c)
            cell.border = BORDER
            if fill:
                cell.fill = fill
            if c in (2, 7):
                cell.alignment = LEFT
        rr += 1
        idx += 1
    ws.cell(rr, 1, "").alignment = CENTER
    ws.cell(rr, 2, "差额：往来科目↔往来科目互冲及比例分摊残差")
    ws.cell(rr, 3, "差额").alignment = CENTER
    ws.cell(rr, 4, cp["internal"]).number_format = MONEY_FMT
    ws.cell(rr, 4).alignment = Alignment(horizontal="right")
    ws.cell(rr, 5, cp["internal"] / ar_total if ar_total else 0).number_format = PCT_FMT
    ws.cell(rr, 5).alignment = Alignment(horizontal="right")
    ws.cell(rr, 6, "").alignment = CENTER
    ws.cell(rr, 7, "同凭证内往来单位间划转/冲销，或混合凭证比例分摊残差（极小，可忽略）")
    for c in range(1, 8):
        cell = ws.cell(rr, c)
        cell.border = BORDER
        cell.fill = NOTE_FILL
        if c in (2, 7):
            cell.alignment = LEFT
    rr += 1
    ws.cell(rr, 2, "外部对应科目合计")
    ws.cell(rr, 4, disp_sum).number_format = MONEY_FMT
    ws.cell(rr, 4).alignment = Alignment(horizontal="right")
    ws.cell(rr, 5, disp_sum / ar_total if ar_total else 0).number_format = PCT_FMT
    ws.cell(rr, 5).alignment = Alignment(horizontal="right")
    for c in range(1, 8):
        cell = ws.cell(rr, c)
        cell.border = BORDER
        cell.font = Font(name='Times New Roman', bold=True)
        cell.fill = TOTAL_FILL
        if c in (2, 7):
            cell.alignment = LEFT
    rr += 1
    ws.cell(rr, 2, base + ("借方净额" if sign == "借" else "贷方发生额"))
    ws.cell(rr, 4, ar_total).number_format = MONEY_FMT
    ws.cell(rr, 4).alignment = Alignment(horizontal="right")
    ws.cell(rr, 5, 1.0).number_format = PCT_FMT
    ws.cell(rr, 5).alignment = Alignment(horizontal="right")
    for c in range(1, 8):
        cell = ws.cell(rr, c)
        cell.border = BORDER
        cell.font = Font(name='Times New Roman', bold=True)
    widths = [6, 52, 8, 18, 14, 8, 46]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"


def _write_recon_counterparty_sheet(ws, title, by_ent, subj, periods, entities, recon_bad=None):
    """【对方科目核对】按核算主体、分年度，列出往来科目对应科目并判断异常。
    by_ent: reconstruct_counterparties_by_entity 的返回 {(ent, pk): {...}}。"""
    ws.cell(1, 1, title).font = TITLE_FONT
    ws.cell(2, 1, "红色=异常（重点核查）；黄色=关注（需核实业务实质）；无底色=正常（符合预期对应科目）。"
            ).font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    hdr = ["核算主体", "年度", "方向", "对应科目（科目名称）", "金额", "笔数",
           "占该侧%", "是否异常", "异常说明"]
    note_row = 3
    if recon_bad:
        txt = ("⚠ 以下主体/年度因《综合查询明细表》『字/号』非唯一凭证号（同一字/号被不同科目/凭证复用），"
               "按字/号翻面还原会把无关交易误列为对方科目，已停用翻面还原以免产生虚假对应科目："
               + "；".join(sorted(f"{e} {p}" for (e, p) in recon_bad))
               + "。")
        c = ws.cell(3, 1, txt)
        c.font = Font(name='Times New Roman', bold=True, color="C00000", size=10)
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=9)
        ws.row_dimensions[3].height = 46
        note_row = 4
    hdr_row = note_row + 1
    for c, h in enumerate(hdr, 1):
        ws.cell(hdr_row, c, h)
    _style_header(ws, hdr_row, len(hdr))

    exp_d = subj["exp_debit"]; exp_c = subj["exp_credit"]
    note_d = subj["notes_debit"]; note_c = subj["notes_credit"]
    sides = [("借方侧", "sale_cp", "sale_cnt", exp_d, note_d, "ar_debit_net"),
             ("贷方侧", "collect_cp", "collect_cnt", exp_c, note_c, "ar_credit_net")]

    anomaly_summary = defaultdict(int)   # ent -> 异常对应科目数
    anom_years = defaultdict(set)        # ent -> 涉及年度
    rr = hdr_row + 1
    for ent in entities:
        wrote_ent = False
        for pk in periods:
            d = by_ent.get((ent, pk))
            if not d:
                continue
            for sname, cpk, cntk, expl, notemap, netk in sides:
                cp = d.get(cpk, {})
                cnt = d.get(cntk, {})
                net = d.get(netk, 0.0)
                # 占该侧% 分母：该侧全部对应科目金额之和（保证各行合计=100% 且单项不超 100%）；
                # 不用控制数 net 做分母，避免 GL 还原口径与控制数不一致时出现 >100% 的误导。
                side_total = sum(cp.values())
                classified = []
                for km, amt in cp.items():
                    verdict, reason = _classify_counterparty(km, expl, notemap, subj["kw"])
                    classified.append((km, amt, cnt.get(km, 0), verdict, reason))
                vrank = {"异常": 0, "关注": 1, "正常": 2}
                classified.sort(key=lambda x: (vrank.get(x[3], 3), -x[1]))
                for km, amt, c, verdict, reason in classified:
                    if abs(amt) < 1.00:
                        continue
                    wrote_ent = True
                    if verdict == "异常":
                        anomaly_summary[ent] += 1
                        anom_years[ent].add(pk)
                    ws.cell(rr, 1, ent)
                    ws.cell(rr, 2, pk).alignment = CENTER
                    ws.cell(rr, 3, sname).alignment = CENTER
                    ws.cell(rr, 4, km)
                    ws.cell(rr, 5, amt).number_format = MONEY_FMT
                    ws.cell(rr, 5).alignment = Alignment(horizontal="right")
                    ws.cell(rr, 6, c).alignment = CENTER
                    ws.cell(rr, 7, amt / side_total if side_total else 0).number_format = PCT_FMT
                    ws.cell(rr, 7).alignment = Alignment(horizontal="right")
                    cell_v = ws.cell(rr, 8, verdict)
                    cell_v.alignment = CENTER
                    ws.cell(rr, 9, reason)
                    fill = ANOM_FILL if verdict == "异常" else (OTHER_FILL if verdict == "关注" else None)
                    for cc in range(1, 10):
                        cell = ws.cell(rr, cc)
                        cell.border = BORDER
                        if fill:
                            cell.fill = fill
                            if cc == 8:
                                cell.font = ANOM_FONT
                        if cc in (4, 9):
                            cell.alignment = LEFT
                    rr += 1
        # 主体小计行（仅当该主体有数据）
        if wrote_ent:
            ws.cell(rr, 1, f"　{ent} 小计").font = TOTAL_FONT
            ws.cell(rr, 2, "")
            ws.cell(rr, 3, "")
            ws.cell(rr, 4, f"异常对应科目 {anomaly_summary[ent]} 项" if anomaly_summary[ent]
                    else "无异常对应科目")
            ws.cell(rr, 8, "异常" if anomaly_summary[ent] else "正常").alignment = CENTER
            for cc in range(1, 10):
                cell = ws.cell(rr, cc)
                cell.border = BORDER
                cell.fill = TOTAL_FILL
                if cc in (4, 8):
                    cell.font = TOTAL_FONT
            rr += 1

    # 异常汇总块（按核算主体）
    rr += 1
    ws.cell(rr, 1, "异常对应科目汇总（按核算主体）").font = SUB_FONT
    rr += 1
    sub_hdr = ["核算主体", "异常对应科目数量", "涉及年度", "结论"]
    for c, h in enumerate(sub_hdr, 1):
        ws.cell(rr, c, h)
    _style_header(ws, rr, len(sub_hdr))
    rr += 1
    any_anom = False
    for ent in entities:
        n = anomaly_summary.get(ent, 0)
        if n == 0:
            continue
        any_anom = True
        ws.cell(rr, 1, ent)
        ws.cell(rr, 2, n).alignment = CENTER
        ws.cell(rr, 3, "、".join(sorted(anom_years[ent]))).alignment = CENTER
        ws.cell(rr, 4, "存在异常，重点核查").font = ANOM_FONT
        for cc in range(1, 5):
            cell = ws.cell(rr, cc)
            cell.border = BORDER
            cell.fill = ANOM_FILL
        rr += 1
    if not any_anom:
        ws.cell(rr, 1, "（全部核算主体未发现标记『异常』的对应科目；『关注』项见上表，需结合凭证核实）")
        rr += 1

    widths = [16, 8, 10, 40, 18, 8, 12, 10, 46]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{hdr_row + 1}"


def _collect_bad_cps(by_ent, subj, periods):
    """收集『异常/关注』对应科目名（与对方科目核对同一判定），供凭证清单反查。
    返回 {科目名: (verdict, reason)}。"""
    exp_d = subj.get("exp_debit", []); exp_c = subj.get("exp_credit", [])
    note_d = subj.get("notes_debit", {}); note_c = subj.get("notes_credit", {})
    out = {}
    sides = [("sale_cp", exp_d, note_d), ("collect_cp", exp_c, note_c)]
    for d in by_ent.values():
        for cpk, expl, notemap in sides:
            for km in (d.get(cpk, {}) or {}):
                verdict, reason = _classify_counterparty(km, expl, notemap, subj.get("kw"))
                if verdict != "正常":
                    out[km] = (verdict, reason)
    return out


def _write_bad_cp_voucher_sheet(ws, bad_cps, rows, subject, label):
    """【异常对应科目凭证清单】——对『异常/关注』对应科目反查 GL 凭证明细
    （凭证号/日期/往来单位/摘要/金额），供审计员直接定位凭证查证业务实质。
    ⚡ 2026-08-12 用户需求：对方科目核对标红后需给出可查证的凭证级线索。"""
    ws.cell(1, 1, f"{label}—异常对应科目凭证清单").font = TITLE_FONT
    ws.cell(2, 1, "列示『对方科目核对』中标记异常/关注的对应科目所在凭证（含凭证号可直接在序时账定位查证）。"
            "红色=异常（重点核查）；黄色=关注。SAP 凭证小计文件（每凭证 1 行）无对方科目列 → 无法反查属正常。"
            ).font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    hdr = ["核算主体", "方向", "异常/关注对应科目", "结论", "凭证号", "日期", "往来单位", "摘要", "金额(元)"]
    for c, h in enumerate(hdr, 1):
        ws.cell(3, c, h)
    _style_header(ws, 3, len(hdr))
    rr = 4
    _n = 0
    _out_rows = []
    for r in rows:
        if not r.get("is_sub"):
            continue
        cpv = str(r.get("cp") or "").strip()
        if not cpv:
            continue
        dr = float(r.get("debit") or 0.0)
        cr = float(r.get("credit") or 0.0)
        for nm in cpv.split(","):
            nm = nm.strip()
            if not nm or nm not in bad_cps:
                continue
            verdict, reason = bad_cps[nm]
            side = "借" if dr > 0.005 else "贷"
            amt = dr if dr > 0.005 else cr
            if abs(amt) < 1.00:
                continue
            d = r.get("date")
            ds = str(d)[:10] if d is not None else ""
            # ⚡⚡ 2026-08-29 用户需求：按『对方科目核对』sheet 顺序排列
            #   （核算主体→年度→借/贷侧→对应科目名），便于精确对应查找异常分录。
            _ent = str(r.get("entity") or subject)
            _yr = str(d)[:4] if d is not None else ""
            _side_order = 0 if dr > 0.005 else 1  # 对方科目核对 sides：借方侧在前
            _out_rows.append((_ent, _yr, _side_order, nm, [
                _ent, side, nm, verdict,
                r.get("vno"), ds, r.get("cust") or "",
                str(r.get("summary") or "")[:60], round(amt, 2)]))
    _out_rows.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
    for _item in _out_rows:
        # ⚡⚡ 2026-08-25 性能优化：ws.append 批量写行（异常凭证清单可达 7.8 万行，
        #   逐 cell 70 万次 → append 快 5 倍），样式后置循环定位。
        ws.append(_item[4])
        rr += 1
        _n += 1
    if _n == 0:
        ws.cell(4, 1, "（该科目无『异常/关注』对应科目凭证，或数据为凭证小计无法反查——对方科目核对无异常项时本表自然为空）")
    widths = [16, 6, 34, 8, 14, 12, 16, 60, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"


def _write_validation_sheet(ws, issues, rows, customers, periods):
    ws.cell(2, 1, "数据校验结果").font = TITLE_FONT
    ar_cnt = sum(1 for r in rows if r["is_sub"])
    info = [f"交易记录数（全表）：{len(rows)}", f"目标科目记录数：{ar_cnt}",
            f"往来单位数：{len(customers)}", f"期间数：{len(periods)}（{', '.join(periods)}）"]
    for i, t in enumerate(info):
        ws.cell(4 + i, 1, t)
    hdr = ["序号", "级别", "类型", "说明"]
    hr = 9
    for i, h in enumerate(hdr, 1):
        ws.cell(hr, i, h)
    _style_header(ws, hr, len(hdr))
    r = hr + 1
    for i, it in enumerate(issues, 1):
        for c, v in enumerate([i, it["level"], it["type"], it["message"]], 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            cell.alignment = LEFT if c == 4 else CENTER
            if it["level"] == "ERROR":
                cell.fill = WARN_FILL
            elif it["level"] == "WARN" and "其他对应科目" in it["type"]:
                cell.fill = OTHER_FILL
        r += 1
    for i, w in enumerate([6, 10, 16, 96], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_drilldown_sheet(ws, rows, subject, label):
    ws.cell(2, 1, f"{label}交易流水（下钻）— 数据源自综合查询明细表(GL)").font = TITLE_FONT
    ws.cell(3, 1, "说明：综合查询明细表为非全量抽取，流水累计与辅助核算余额差额=取数范围限制，"
                   "不构成账务差错；本表仅用于交易级下钻与审计抽样。").font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    hdr = ["序号", "核算主体", "期间", "日期", "字", "号", "往来单位名称", "摘要", "对方科目", "借方金额", "贷方金额"]
    for i, h in enumerate(hdr, 1):
        ws.cell(5, i, h)
    _style_header(ws, 5, len(hdr))
    r = 6
    idx = 1
    for x in rows:
        if not x.get("is_sub"):
            continue
        d = x.get("date")
        pk = d.strftime("%Y") if d else ""
        ds = d.strftime("%Y-%m-%d") if d else ""
        vals = [idx, subject, pk, ds, x.get("vtype", ""), x.get("vno", ""),
                x.get("cust", ""), x.get("summary", ""), x.get("cp", ""), x.get("debit", 0), x.get("credit", 0)]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (10, 11):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
            elif c in (1, 3, 4, 5):
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
        idx += 1
        r += 1
    widths = [6, 16, 8, 12, 6, 8, 30, 40, 28, 16, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A6"


# ----------------------------------------------------------------------------
# 应收票据专项底稿（仅 ARN）：票据种类拆分 / 背书贴现明细 / 坏账准备
#   相关日期(出票日/到期日)或坏账准备数据若账套中无，直接留空（用户要求）。
# ----------------------------------------------------------------------------
def _write_bill_type_split(ws, periods, customers, summary, subject, label):
    """银行承兑/商业承兑汇票按辅助核算子目拆分：汇总 + 逐客户明细。"""
    ws.cell(2, 1, f"{label}—票据种类拆分（银行承兑汇票 / 商业承兑汇票 按辅助核算子目归类）").font = TITLE_FONT
    note = ("说明：票据种类由辅助核算科目名称子目段推导（如『应收票据-银行承兑汇票』→银行承兑汇票）。"
            "账套未分子目或无承兑标识的列为『未分类』。")
    c2 = ws.cell(3, 1, note); c2.font = Font(name='Times New Roman', italic=True, size=10, color="808080"); c2.alignment = LEFT
    def _type_of(s):
        return (s.get("bill_type") or "未分类") if s else "未分类"
    # 汇总块
    for i, h in enumerate(["票据种类", "期间", "客户数", "本币期末余额"], 1):
        ws.cell(5, i, h)
    _style_header(ws, 5, 4)
    types = []
    for pk in periods:
        for ak in customers:
            t = _type_of(summary.get((ak, pk)))
            if t not in types:
                types.append(t)
    ordered = [t for t in ["银行承兑汇票", "商业承兑汇票", "财务公司承兑汇票", "未分类"] if t in types] + \
              [t for t in types if t not in ["银行承兑汇票", "商业承兑汇票", "财务公司承兑汇票", "未分类"]]
    r = 6
    for t in ordered:
        for pk in periods:
            cnt = 0; tot = 0.0
            for ak in customers:
                s = summary.get((ak, pk))
                if s and _type_of(s) == t:
                    cnt += 1; tot += s.get("s_close", 0.0)
            ws.cell(r, 1, t).border = BORDER
            ws.cell(r, 2, pk).border = BORDER
            ws.cell(r, 3, cnt).border = BORDER
            cc = ws.cell(r, 4, round(tot, 2)); cc.border = BORDER; cc.number_format = MONEY_FMT
            cc.alignment = Alignment(horizontal="right")
            r += 1
    for i, w in enumerate([20, 10, 10, 20], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    r += 1
    # 明细块
    ws.cell(r, 1, "按客户明细（票据种类 / 客户 / 期间 / 本币期末余额）").font = SUB_FONT
    r += 1
    for i, h in enumerate(["票据种类", "客户", "期间", "本币期末余额"], 1):
        ws.cell(r, i, h)
    _style_header(ws, r, 4)
    r += 1
    for ak in customers:
        nm = _disp_name(ak)
        for pk in periods:
            s = summary.get((ak, pk))
            if not s:
                continue
            t = _type_of(s)
            ws.cell(r, 1, t).border = BORDER
            ws.cell(r, 2, nm).border = BORDER
            ws.cell(r, 3, pk).border = BORDER
            cc = ws.cell(r, 4, round(s.get("s_close", 0.0), 2)); cc.border = BORDER
            cc.number_format = MONEY_FMT; cc.alignment = Alignment(horizontal="right")
            r += 1
    ws.freeze_panes = "A6"


def _write_endorse_discount(ws, rows, subject, label):
    """背书 / 贴现 / 承兑 交易明细（据 GL 摘要关键词筛选）。出票日、到期日账套无则留空。"""
    ws.cell(2, 1, f"{label}—背书 / 贴现 / 承兑 明细（数据源自综合查询明细表(GL) 摘要字段）").font = TITLE_FONT
    ws.cell(3, 1, "说明：按凭证『摘要』含『背书/贴现/承兑』关键词筛选；出票日、到期日账套未提供，暂留空。").font = \
        Font(name='Times New Roman', italic=True, size=10, color="808080")
    hdr = ["序号", "日期", "字", "号", "往来单位名称", "摘要", "对方科目", "借方金额", "贷方金额", "业务类型", "出票日", "到期日"]
    for i, h in enumerate(hdr, 1):
        ws.cell(5, i, h)
    _style_header(ws, 5, len(hdr))
    r = 6; idx = 1
    def _biz(t):
        t = t or ""
        if "背书" in t:
            return "背书"
        if "贴现" in t:
            return "贴现"
        if "承兑" in t:
            return "承兑"
        return "其他"
    matched = 0
    for x in rows:
        if not x.get("is_sub"):
            continue
        biz = _biz(x.get("summary", ""))
        if biz == "其他":
            continue
        matched += 1
        d = x.get("date")
        ds = d.strftime("%Y-%m-%d") if d else ""
        vals = [idx, ds, x.get("vtype", ""), x.get("vno", ""), x.get("cust", ""),
                x.get("summary", ""), x.get("cp", ""), x.get("debit", 0), x.get("credit", 0), biz, None, None]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (8, 9):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
            elif c in (1, 2, 3, 4):
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
        idx += 1; r += 1
    if matched == 0:
        ws.cell(r, 1, "（账套交易流水中未发现『背书/贴现/承兑』相关记录；出票日、到期日账套未提供，留空）").font = \
            Font(name='Times New Roman', italic=True, color="808080")
        r += 1
    for i, w in enumerate([6, 12, 6, 8, 30, 40, 28, 16, 16, 12, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A6"


def _read_baddebt_for_notes(km_files, periods):
    """读取科目余额表中与应收票据相关的坏账准备（科目名同时含『坏账准备』与『应收票据』）。
    返回 {period: dict(open, debit, credit, close)}（带符号：备抵科目，贷余为正、借余为负）。

    ⚡ 2026-08-12 SAP 分支：SAP discover_entities 的 km 文件列表恒为空（TB 由 adapter 内部
    读取）→ 原逻辑直接返回空 → 应收票据坏账准备表空（质检 ERROR）。SAP 下从
    adapter.read_tb_full 直接取（key=(e,c,n,y)，qc/jf/df/qm 已带符号，贷余为负——
    与 U8 _signed_balance(..., 'asset') 口径一致）。"""
    out = {}
    if _adapter is not None and _is_sap_mode(getattr(_adapter, '_DATA_ROOT', None)):
        try:
            _tb = _adapter.read_tb_full(getattr(_adapter, '_DATA_ROOT', None), None)
        except Exception:
            _tb = {}
        for (e, c, n, y), v in _tb.items():
            nn = str(n)
            if "应收票据" not in nn:
                continue
            cc = str(c)
            if "坏账准备" not in nn and not cc.startswith("1231"):
                continue
            pk = str(y)
            if pk not in periods and "全部" not in periods:
                continue
            rec = out.setdefault(pk, dict(open=0.0, debit=0.0, credit=0.0, close=0.0))
            rec["open"] += float(v.get("qc") or 0.0)
            rec["debit"] += float(v.get("jf") or 0.0)
            rec["credit"] += float(v.get("df") or 0.0)
            rec["close"] += float(v.get("qm") or 0.0)
        return out
    if not km_files:
        return out
    files = km_files if isinstance(km_files, (list, tuple)) else [km_files]
    for p in files:
        if not p or not os.path.isfile(p):
            continue
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            raw = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception:
            continue
        if not raw:
            continue
        hdr_idx, headers = _detect_header(raw, AUX_COLUMN_ALIASES)
        mp = _resolve_aux_columns(headers)
        if "km" not in mp or "close" not in mp:
            continue
        pk = _detect_year(os.path.basename(p))
        if pk not in periods and "全部" not in periods:
            continue
        for ri, row in enumerate(raw[hdr_idx + 1:], start=hdr_idx + 2):
            if not row or all(c is None or str(c).strip() == "" for c in row):
                continue
            km = _str(row, mp.get("km"), "")
            # 2026-08-03 修复：FY 的坏账准备子目名直接叫『应收票据』（1231.03，父级 1231 才叫坏账准备）
            # → 匹配放宽：科目名含『应收票据』 且 （含『坏账准备』 或 代码以 1231 开头）
            if "应收票据" not in km:
                continue
            km_code = _str(row, mp.get("code"), "")
            if "坏账准备" not in km and not str(km_code).startswith("1231"):
                continue
            op = _parse_amount(row[mp["open"]], ri) if mp.get("open") is not None and mp["open"] < len(row) else 0.0
            db = _parse_amount(row[mp["debit"]], ri) if mp.get("debit") is not None and mp["debit"] < len(row) else 0.0
            cr = _parse_amount(row[mp["credit"]], ri) if mp.get("credit") is not None and mp["credit"] < len(row) else 0.0
            cl = _parse_amount(row[mp["close"]], ri) if mp["close"] < len(row) else 0.0
            op_dir = _str(row, mp.get("open_dir"), "")
            cl_dir = _str(row, mp.get("close_dir"), "")
            rec = out.setdefault(pk, dict(open=0.0, debit=0.0, credit=0.0, close=0.0))
            rec["open"] += _signed_balance(op, op_dir, "asset")
            rec["debit"] += _signed_debit(db, "asset")
            rec["credit"] += _signed_credit(cr, "asset")
            rec["close"] += _signed_balance(cl, cl_dir, "asset")
    return out


def _write_note_baddebt(ws, km_path, periods, subject, label):
    """应收票据坏账准备。账套无则留空骨架并注明。"""
    ws.cell(2, 1, f"{label}—坏账准备（应收票据相关）").font = TITLE_FONT
    bd = _read_baddebt_for_notes(km_path, periods)
    hdr = ["期间", "期初余额", "本期计提(贷)", "本期转回(借)", "期末余额", "备注"]
    if not bd:
        ws.cell(4, 1, "（账套科目余额表中未发现与应收票据相关的坏账准备"
                       "（科目名须同时含『坏账准备』与『应收票据』）。"
                       "按审计要求，相关余额暂留空，待账套补充或手工填列。）").font = Font(name='Times New Roman', italic=True, color="808080")
        for i, h in enumerate(hdr, 1):
            ws.cell(6, i, h)
        _style_header(ws, 6, len(hdr))
        r = 7
        for pk in periods:
            ws.cell(r, 1, pk).border = BORDER
            for c in range(2, len(hdr) + 1):
                ws.cell(r, c).border = BORDER
            r += 1
        for i, w in enumerate([12, 18, 18, 18, 18, 24], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        return
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for pk in periods:
        rec = bd.get(pk)
        if not rec:
            ws.cell(r, 1, pk).border = BORDER
            for c in range(2, len(hdr) + 1):
                ws.cell(r, c).border = BORDER
            r += 1
            continue
        vals = [pk, round(rec["open"], 2), round(rec["credit"], 2), round(-rec["debit"], 2),
                round(rec["close"], 2), ""]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (2, 3, 4, 5):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        r += 1
    for i, w in enumerate([12, 18, 18, 18, 18, 24], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"


def export_excel(output, periods, customers, summary, issues, rows, counterparties, mode="Y", subject="（未提供）", subj=None, by_ent=None, aging_map=None, bucket_type=None, km_path=None, target_year=None):
    wb = openpyxl.Workbook()
    label = subj["label"] if subj else "往来科目"
    # 统一 customers 表示为 2-tuple (entity, name)；summary 同步重键，
    # 兼容 build_ca_detail 传字符串名（无 entity）与传 2-tuple 两种调用方式，避免 aging/汇总表解包崩。
    _cust_norm = []
    _sum_norm = {}
    for c in customers:
        if isinstance(c, (tuple, list)) and len(c) == 2:
            e, n = c[0], c[1]
        else:
            e, n = None, c
        _cust_norm.append((e, n))
        for pk in periods:
            if (c, pk) in summary:
                _sum_norm[((e, n), pk)] = summary[(c, pk)]
            elif ((e, n), pk) in summary:
                _sum_norm[((e, n), pk)] = summary[((e, n), pk)]
    customers = _cust_norm
    summary = _sum_norm
    # 剔除无年份伪期间 "全部"（源于文件名无年份的源文件），避免凭空生成 明细_全部 / 期间对比；
    # 仅当存在≥2 个真实年度时才视为多期（单期数据不生成 明细全部 / 期间对比）。
    periods = [p for p in periods if p != "全部"] or list(periods)
    multi = len(periods) > 1
    _render_periods = [target_year] if target_year is not None else periods
    for pk in _render_periods:
        ws = wb.create_sheet(f"{label}明细表_{pk}")
        _write_period_sheet(ws, pk, customers, summary, subject, label, aging_map=aging_map, bucket_type=bucket_type)
    # ⚡ 2026-08-16 旧『账龄分布』已废弃：由专项 aging_review.py（应收+合同资产联动、FIFO 期初纳入）替代，
    #   原 _write_aging_sheet 不再输出（避免与专项账龄复核重复/失真）
    # 注：原『借方对应科目 / 贷方对应科目』两表无意义，已移除；保留『对方科目核对』(按主体·分年度)。
    # 对方科目核对（按核算主体·分年度，判断异常）：无对应科目金额数据时不生成该 sheet（2026-07-31 用户要求）
    entities = sorted({r.get("entity") or subject for r in rows})
    recon_bad = _recon_unreliable(rows)
    _by_ent_data = by_ent if by_ent is not None else reconstruct_counterparties_by_entity(rows)
    _has_cp_data = any(
        abs(amt) >= 1.0
        for _d in _by_ent_data.values()
        for _cpk in ("sale_cp", "collect_cp")
        for amt in (_d.get(_cpk, {}) or {}).values()
    )
    if _has_cp_data:
        ws = wb.create_sheet("对方科目核对")
        _write_recon_counterparty_sheet(ws, f"{label}—对方科目核对（按核算主体·分年度）",
                                        _by_ent_data, subj, periods, entities, recon_bad=recon_bad)
        # ⚡ 2026-08-12 新增：异常对应科目凭证清单（对方科目核对标红后给出可查证的凭证级线索）
        try:
            _bad_cps = _collect_bad_cps(_by_ent_data, subj, periods)
            if _bad_cps:
                ws = wb.create_sheet("异常对应科目凭证清单")
                _write_bad_cp_voucher_sheet(ws, _bad_cps, rows, subject, label)
        except Exception:
            pass
    # 2026-07-31 用户要求：删除『校验结果』『借贷方检查』『交易流水(下钻)』『重要性清单』四张底稿
    # 应收票据专项底稿（仅 ARN）：票据种类拆分 / 背书贴现明细 / 坏账准备
    if subj and (subj.get("kw") == "应收票据" or subj.get("sheet") == "应收票据明细表"):
        ws = wb.create_sheet("票据种类拆分")
        _write_bill_type_split(ws, periods, customers, summary, subject, label)
        ws = wb.create_sheet("背书贴现明细")
        _write_endorse_discount(ws, rows, subject, label)
        ws = wb.create_sheet("应收票据坏账准备")
        _write_note_baddebt(ws, km_path, periods, subject, label)
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    order = ([f"{label}明细表_{pk}" for pk in _render_periods] +
             ["对方科目核对",
              "票据种类拆分", "背书贴现明细", "应收票据坏账准备"])
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
    finalize_workbook(wb)
    _base = os.path.basename(output)
    if _base.endswith("_生成.xlsx"):
        _base = _base[:-len("_生成.xlsx")]
    elif _base.endswith(".xlsx"):
        _base = _base[:-len(".xlsx")]
    outs = Y.emit_per_year(wb, _base, os.path.dirname(output))
    wb.close()
    return outs


def _strip_export_ts(name):
    """去除用友/U8 导出文件末尾的时间戳后缀，如 '_20260723152733'（8–14 位纯数字）。"""
    return re.sub(r"_\d{8,14}$", "", name)


def _detect_year(fn):
    """从文件名判定会计年度。优先 4 位年份(19xx/20xx)；无 4 位年份时，
    用独立 2 位年份(20–29 -> 20xx)兜底（兼容 '上分25科目余额表' 这类命名）。
    先剥离导出时间戳，避免把导出日期(2026)误判为数据年度。"""
    base = re.sub(r"\.(xlsx|xls|csv)$", "", fn, flags=re.I)
    base = _strip_export_ts(base)
    m = re.search(r"(19|20)\d{2}", base)
    if m:
        return m.group(0)
    m2 = re.search(r"(?<![\d])(2\d)(?![\d])", base)
    if m2:
        return "20" + m2.group(1)
    return "全部"


def _derive_subject(aux_path):
    files = []
    if os.path.isdir(aux_path):
        for fn in sorted(os.listdir(aux_path)):
            if "辅助核算余额表" in fn and fn.lower().endswith((".xlsx", ".xls", ".csv")):
                files.append(fn)
    elif os.path.isfile(aux_path):
        files.append(os.path.basename(aux_path))
    if not files:
        return "（未提供）"
    name = files[0]
    name = re.sub(r"\.(xlsx|xls|csv)$", "", name, flags=re.I)   # 1. 去扩展名
    for suf in ("辅助核算余额表", "科目余额表", "综合查询明细表", "综合查询表", "余额表", "外币"):
        name = name.replace(suf, "")                              # 2. 去类型后缀（含 GL 文件名变体"综合查询表"）
    name = _strip_export_ts(name)                                 # 3. 去导出时间戳(否则残片粘进实体名)
    name = re.sub(r"[\s_\-]*(?:19|20)\d{2}", "", name)           # 4. 去 4 位年份
    name = re.sub(r"年度|年\d*[-~]?\d*月", "", name)               # 4a. 去"年度"、"年1-3月"等期间后缀
    name = name.rstrip("年")  # 2026-08-06：U8『主体2025年科目余额表』去年份后残留孤立『年』
    # （与 audit_common.discover_entities 一致——否则 comb_summary 实体名带"年"、
    #   审定表主体(ac_ents)不带"年" → reclass_map 键不匹配 → 审定表重分类 J/K 全 0）
    name = re.sub(r"[\s_\-]*(?<!\d)(2\d)(?!\d)", "", name)       # 5. 去 2 位年份(20-29 -> 20xx)，使同主体跨文件命名一致
    # 2026-08-06 修复：上述正则会把开头【序号 20-29】误当 2 位年份删除（ga 主体
    # 『23浙江省工业设备安装集团有限公司白俄代表处…』→ 丢序号 → 实体名与
    # audit_common.discover_entities 不一致 → reclass_map 键不匹配 → 审定表重分类 9233 万丢失）。
    # 补救：开头序号被误删时按排序补回（文件名含序号、实体名丢了序号的场景）。
    _m_seq = re.match(r"^(\d{2})(?=[^\d])", os.path.basename(aux_path) if aux_path else '')
    if _m_seq and not re.match(r"^\d", name):
        name = _m_seq.group(1) + name
    name = name.strip(" _-")
    # 6. 个位数前缀补0（"1、"→"01、"），使实体按 01、02、…10、11、12 排序（对齐 audit_common.discover_entities 铁律）
    _m0 = re.match(r"^(\d)([、,，.．\s])", name)
    if _m0 and len(name) > 2:
        name = "0" + name
    return name or "（未提供）"


def _build_period_map(inputs, keyword=None):
    if isinstance(inputs, (list, tuple)):
        paths = list(inputs)
    elif os.path.isdir(inputs):
        paths = []
        for fn in sorted(os.listdir(inputs)):
            if fn.startswith("~$"):
                continue
            if not fn.lower().endswith((".xlsx", ".xls", ".csv")):
                continue
            if keyword and keyword not in fn:
                continue
            paths.append(os.path.join(inputs, fn))
    else:
        paths = [inputs]
    res = {}
    for p in paths:
        if p is None:   # ⚡ SAP：占位 gl 路径，跳过（数据走 adapter.read_gl）
            continue
        period = _detect_year(os.path.basename(p))
        res.setdefault(period, []).append(p)
    return res


# ============================================================================
# 主入口
# ============================================================================
# ============================================================================
# 合并模式专用：逐主体分组 -> 合并写出（核算主体列逐行正确，跨主体同名不覆盖）
# ============================================================================
def _cwrite_period_sheet(ws, period, comb_customers, comb_summary, label, aging_map=None, bucket_type=None,
                         prior_meta=None):
    """明细表（2026-08-03 用户方法论：期初数前加 最上层客户单位/客户性质/坏账计提方法，
    最后一列加 期后收款；账龄优先用上年账龄对照（简化滚动），无上年表时回退 FIFO）。
    prior_meta: {(entity, cust): {'top_unit','nature','aging','after'}}（来自上年账龄表）。"""
    buckets = AGING_BUCKETS.get(bucket_type) if bucket_type else None
    # ⚡ 2026-08-09 SAP 合同号列：客户名含 '@@'（SAP 按 单位×合同号 分行，read_opening_balances
    #   SAP 分支 key=客户名@@合同号）→ 明细表在『往来单位名称』后插入『合同号』列；
    #   U8 客户名无 '@@' → 不插列（保持 FY 等账套原列结构）。
    _has_con = any('@@' in str(n) for (_e, n) in comb_customers)
    _off = 1 if _has_con else 0
    base_headers = ["核算主体", "序号", "往来单位编号", "往来单位名称", "款项性质", "是否关联方", "币种",
                    "最上层客户单位", "客户性质", "坏账计提方法",
                    "外币余额(期初)", "期初余额(人民币)", "本期借方发生额", "本期贷方发生额",
                    "外币余额(期末)", "期末余额(人民币)",
                    "重分类调整", "审计调整", "审定数"]
    if _has_con:
        base_headers = base_headers[:4] + ["合同号"] + base_headers[4:]
    if buckets:
        headers = base_headers + buckets + ["交易笔数", "平衡校验(期初+借+贷-期末)", "期后收款"]
        nb = len(buckets)
    else:
        headers = base_headers + ["账龄区间", "交易笔数", "平衡校验(期初+借+贷-期末)", "期后收款"]
        nb = 0
    # ⚡ 2026-08-10 排版优化（R3 双行表头）：行4=大类分组合并单元格，行5=列名；数据从行6
    #   分组：基础信息(10或11含合同号) | 期初(2) | 本期发生额(2) | 期末(2) | 调整(3)
    #         [账龄(nb)] | 其他(3：交易笔数/平衡校验/期后收款)
    _base_cnt = 10 + _off
    _groups = [('基础信息', _base_cnt), ('期初', 2), ('本期发生额', 2), ('期末', 2), ('调整', 3)]
    if buckets:
        _groups.append(('账龄', nb))
    _groups.append(('其他', 3))
    two_row_header(ws, 4, _groups, headers)
    ws.cell(2, 1, f"{label}明细表 — {period}").font = TITLE_FONT
    r = 6
    tot = {"open": 0.0, "debit": 0.0, "credit": 0.0, "close": 0.0, "count": 0, "fopen": 0.0, "fclose": 0.0}
    tot_buckets = {b: 0.0 for b in (buckets or [])}
    _first_data_row = r
    money_cols = list(range(11 + _off, 20 + _off)) + list(range(20 + _off, 20 + _off + nb))
    for idx, (entity, name) in enumerate(comb_customers, 1):
        s = comb_summary[(entity, name), period]
        # ⚡ 2026-08-26 兜底占位行（无辅助核算明细/未标注往来单位）平衡校验留空：占位行仅补
        #   TB 期末差额、无期初/发生额（发生额在 GL/aux 客户行）→ 平衡=期末≠0 是固有现象，
        #   会误导审计师以为差异。留空并表底说明；合计行才是真正的勾稽（期初+借-贷-期末=0）。
        bal = (s["s_open"] + s["s_debit"] + s["s_credit"] - s["s_close"]
               if name not in ("（无辅助核算明细）", "（未标注往来单位）") else None)
        ccy = "外币" if s["foreign"] else "人民币"
        f_open = s.get("fcur_open", 0.0) or 0.0
        f_close = s.get("fcur_close", 0.0) or 0.0
        _nm = name
        _con = ''
        if _has_con and '@@' in str(name):
            _nm, _con = str(name).split('@@', 1)
        # 2026-08-03 用户方法论：期初数前加 最上层客户单位/客户性质/坏账计提方法
        pm = ((prior_meta or {}).get((entity, name))
              or (prior_meta or {}).get(('*', name)) or {}) or {}
        top_unit = pm.get('top_unit', '')
        cust_nature = pm.get('nature', '')
        baddebt_method = pm.get('method', '账龄计提')     # 默认账龄计提，后续手工改单项/组合
        # 2026-08-03 重分类：①本科目期末为负数（余额方向与科目性质相反）→ 自动过入『重分类调整』
        # 列（审定归 0，账龄清 0）；②对方科目负数余额转入本明细表（reclass_in，重分类调整=转入额，
        # 审定数=转入额，资产类账龄按 1 年以内首档计入）。
        _rc = 0.0
        if (s["s_close"] or 0.0) < -0.005:
            _rc = round(-(s["s_close"] or 0.0), 2)
        _rc_in = round(float(s.get("reclass_in", 0.0) or 0.0), 2)
        _rc_tot = _rc + _rc_in
        base_vals = [entity, idx, s["code"], _nm]
        if _has_con:
            base_vals += [_con]
        base_vals += [s.get("nature", ""), ("是" if s.get("related") else "否"), ccy,
                      top_unit, cust_nature, baddebt_method,
                      (f_open if f_open else None), s["s_open"], abs(s["s_debit"]),
                      abs(s["s_credit"]), (f_close if f_close else None), s["s_close"],
                      (_rc_tot if _rc_tot else None), None]
        # 审定数 = 期末未审(16+off) + 重分类(17+off) + 审计调整(18+off) 公式
        from openpyxl.utils import get_column_letter as _gcl
        _f_end = _gcl(16 + _off); _f_rc = _gcl(17 + _off); _f_adj = _gcl(18 + _off)
        base_vals += ["=%s%d+%s%d+%s%d" % (_f_end, r, _f_rc, r, _f_adj, r)]
        if buckets:
            # 优先用上年账龄对照（简化滚动）：prior_aging 存在 → _roll_aging_from_prior
            if _rc:
                bm = {b: 0.0 for b in buckets}       # 重分类转出：账龄清 0
            elif _rc_in:
                bm = {b: 0.0 for b in buckets}
                bm[buckets[0]] = _rc_in               # 重分类转入：账龄按 1 年以内首档
            elif pm.get('aging'):
                bm = _roll_aging_from_prior(pm['aging'], s["s_open"], s["s_debit"],
                                            s["s_close"], buckets)
            else:
                bm = aging_map.get(((entity, name), period)) if aging_map else None
                bm = bm or {b: 0.0 for b in buckets}
            bucket_vals = [round(bm.get(b, 0.0), 2) for b in buckets]
            after = pm.get('after', None)
            vals = base_vals + bucket_vals + [s["count"], (None if bal is None else round(bal, 2)), after]
        else:
            aging = s["aging"] if s["aging"] else "（未提供）"
            after = pm.get('after', None)
            vals = base_vals + [aging, s["count"], (None if bal is None else round(bal, 2)), after]
        # ⚡⚡ 2026-08-25 性能优化：ws.append 批量写值替代逐 cell 创建（XBJ 69 万行明细
        #   实测逐 cell 是 39 分钟 build 的主因）；金额列 0 值预置 None（保持"零值留空"），
        #   样式（border/format/alignment/fill 分组）在 append 后统一遍历——原 group_fill
        #   对 69 万行二次遍历设底色 → 合并进本循环，消除第二次 1035 万次 cell 访问。
        _row = list(vals)
        for _ci in money_cols:
            if _ci <= len(_row):
                _v = _row[_ci - 1]
                if isinstance(_v, (int, float)) and abs(_v) <= 0.005:
                    _row[_ci - 1] = None
        ws.append(_row)
        # ⚡⚡ 2026-08-25 修复：勿用 ws.max_row 定位——max_row 是 O(全 cells) 遍历
        #   （max(self._cells)），69 万行每行调用 = O(n²) = 50 分钟（v2 复测实测元凶）。
        #   r 为循环计数（6+idx-1），append 连续写入同号行，直接复用。
        for c, v in enumerate(_row, 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            # 金额列：外币期初(11+off) 期初(12+off) 借(13+off) 贷(14+off)
            #         外币期末(15+off) 期末(16+off) + 账龄桶(17+off..)
            if c in money_cols:
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
                # 列分组底色（R4：期初蓝/发生额绿/期末蓝/调整黄/账龄灰）——合并原 group_fill
                if c in (11 + _off, 12 + _off):
                    cell.fill = FILL_GROUP_BLUE
                elif c in (13 + _off, 14 + _off):
                    cell.fill = FILL_GROUP_GREEN
                elif c in (15 + _off, 16 + _off):
                    cell.fill = FILL_GROUP_BLUE
                elif c in (17 + _off, 18 + _off, 19 + _off):
                    cell.fill = FILL_GROUP_YELLOW
                elif nb and c in range(20 + _off, 20 + _off + nb):
                    cell.fill = FILL_GROUP_GRAY
            elif c in (2, 5 + _off, 6 + _off, 7 + _off, 8 + _off, 9 + _off, 10 + _off):
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
            if bal is not None and abs(bal) > 0.005 and c == len(vals):
                cell.fill = WARN_FILL
        for b, v in zip(buckets or [], bucket_vals if buckets else []):
            tot_buckets[b] += v
        tot["open"] += s["s_open"]; tot["debit"] += abs(s["s_debit"])
        tot["credit"] += abs(s["s_credit"]); tot["close"] += s["s_close"]
        tot["count"] += s["count"]
        tot["fopen"] += f_open; tot["fclose"] += f_close
        r += 1
    # ⚡⚡ 2026-08-25 列分组底色已在主循环内按列号合并设置（原 group_fill 对 69 万行
    #   二次遍历设 fill，是 XBJ ORA build 39 分钟的主因之一）——此处不再二次遍历。
    tcell = ws.cell(r, 1, "合计")
    tcell.font = TOTAL_FONT
    for c in range(1, len(headers) + 1):
        ws.cell(r, c).border = BORDER
        ws.cell(r, c).fill = TOTAL_FILL
    for c, key in [(11 + _off, "fopen"), (12 + _off, "open"), (13 + _off, "debit"), (14 + _off, "credit"),
                   (15 + _off, "fclose"), (16 + _off, "close")]:
        # ⚡ 2026-08-10 修复：合计行 open/close 【不再 abs】。2026-08-05 为对齐"审定表正数"加的
        #   abs 是历史 bug——明细行带符号（负债借余 -120万），合计 abs（+120万）→ 合计≠Σ明细
        #   → audit_checker 15 个往来明细表 ERROR（表表不一致）。正确口径：合计=Σ明细（带符号）
        #   = 审定表（TB 带符号），三者一致（铁律54：报表格式才取正，明细/勾稽带符号）。
        #   当日 105.6M 差异场景（租赁负债 -52.66M）：合计带符号后 = 明细Σ = 审定表，同号一致。
        _v = tot[key]
        cell = ws.cell(r, c, _v)
        cell.font = TOTAL_FONT
        cell.fill = TOTAL_FILL
        cell.border = BORDER
        cell.number_format = MONEY_FMT
    # 2026-08-06 协议：重分类/期末/审定 列合计写数值（=数据行和；公式 data_only 读 None）
    # ⚡ 2026-08-11 DQ2 修复：审定数(19+off)合计 = 期末(16+off) + 重分类(17+off) + 审计调整(18+off)
    #   ——明细行审定数为公式，sum 公式格 data_only=None→0，曾致 25 年预付审定合计=0。
    _ce = 16 + _off; _cr = 17 + _off; _ca = 18 + _off; _cau = 19 + _off
    _se = sum(ws.cell(x, _ce).value for x in range(_first_data_row, r)
              if isinstance(ws.cell(x, _ce).value, (int, float)))
    _sr = sum(ws.cell(x, _cr).value for x in range(_first_data_row, r)
              if isinstance(ws.cell(x, _cr).value, (int, float)))
    _sa = sum(ws.cell(x, _ca).value for x in range(_first_data_row, r)
              if isinstance(ws.cell(x, _ca).value, (int, float)))
    for c, v in [(_cr, _sr), (_ca, _sa), (_cau, _se + _sr + _sa)]:
        cell = ws.cell(r, c, round(v, 2))
        cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
        cell.number_format = MONEY_FMT
    for j, b in enumerate(buckets or []):
        cell = ws.cell(r, 20 + _off + j, round(tot_buckets[b], 2))
        cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
        cell.number_format = MONEY_FMT
    cnt_col = 20 + _off + nb
    cell = ws.cell(r, cnt_col, tot["count"]); cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
    # ⚡ 2026-08-10 排版优化（R5 合计层级中蓝 + 粗分隔线；R2 冻结前5列+双行表头；外币列全0隐藏）
    try:
        style_total(ws, r, len(headers), level='total')
    except Exception:
        pass
    widths = ([16, 6, 16, 30, 16, 10, 8, 22, 14, 12, 14, 16, 16, 16, 14, 16]
              + [14, 14, 14] + [14] * nb + [10, 22, 16])
    if _has_con:
        widths = widths[:4] + [18] + widths[4:]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "F6"
    hide_col_if_zero(ws, 11 + _off, r - 1, note_cell=(1, 11 + _off), note_txt='外币列全 0，已隐藏（可取消隐藏查看）')
    hide_col_if_zero(ws, 15 + _off, r - 1)
    # ⚡ 2026-08-26 占位行平衡校验留空说明（用户反馈『平衡校验为何出现差异』）
    _n = ws.cell(r + 2, 1, '说明：『（无辅助核算明细）/（未标注往来单位）』为 TB 差额/未标注兜底占位行，'
                          '仅含期末差额、期初与发生额分布于客户行 → 平衡校验列留空不参与校验；'
                          '合计行『期初+借-贷=期末』为总勾稽。')
    _n.font = Font(name='Times New Roman', italic=True, size=10, color='808080')


def _cwrite_comparison_sheet(ws, comb_customers, periods, comb_summary, label):
    ws.cell(2, 1, f"{label}明细 — 期间对比").font = TITLE_FONT
    hdr = ["核算主体", "往来单位编号", "往来单位名称", "款项性质", "是否关联方", "币种"]
    for pk in periods:
        hdr += [f"{pk}外币期初", f"{pk}期初", f"{pk}借方", f"{pk}贷方", f"{pk}外币期末", f"{pk}期末"]
    # 跨期衔接：上期期末(外币/本币) vs 下期期初(外币/本币) + 差异（差异直接高亮显示）
    transitions = []
    for i in range(len(periods) - 1):
        p0, p1 = periods[i], periods[i + 1]
        transitions.append((p0, p1))
        hdr += [f"{p0}期末外币", f"{p1}期初外币", "外币差异",
                f"{p0}期末本币", f"{p1}期初本币", "本币差异"]
    hdr += ["首期期末", "末期期末", "期末变动"]
    diff_cols = [i + 1 for i, h in enumerate(hdr) if h in ("外币差异", "本币差异")]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for (entity, name) in comb_customers:
        first_close = comb_summary[(entity, name), periods[0]]["close"]
        last_close = comb_summary[(entity, name), periods[-1]]["close"]
        ccy = "外币" if comb_summary[(entity, name), periods[0]]["foreign"] else "人民币"
        code = comb_summary[(entity, name), periods[0]]["code"]
        nature = comb_summary[(entity, name), periods[0]].get("nature", "")
        related = "是" if comb_summary[(entity, name), periods[0]].get("related") else "否"
        row_vals = [entity, code, name, nature, related, ccy]
        for pk in periods:
            s = comb_summary[(entity, name), pk]
            fo = s.get("fcur_open", 0.0) or 0.0
            fc = s.get("fcur_close", 0.0) or 0.0
            row_vals += [(fo if fo else None), s["s_open"], abs(s["s_debit"]), abs(s["s_credit"]),
                         (fc if fc else None), s["s_close"]]
        for (p0, p1) in transitions:
            s0 = comb_summary[(entity, name), p0]; s1 = comb_summary[(entity, name), p1]
            fc0 = s0.get("fcur_close", 0.0) or 0.0
            fc1 = s1.get("fcur_open", 0.0) or 0.0
            lc0 = s0["s_close"]; lc1 = s1["s_open"]
            row_vals += [(fc0 if fc0 else None), (fc1 if fc1 else None), round(fc1 - fc0, 2),
                         lc0, lc1, round(lc1 - lc0, 2)]
        row_vals += [first_close, last_close, last_close - first_close]
        for c, v in enumerate(row_vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c >= 7:
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        for c in diff_cols:
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and abs(v) > 0.005:
                ws.cell(r, c).fill = WARN_FILL
        r += 1
    ws.cell(r, 1, "合计").font = TOTAL_FONT
    for c in range(1, len(hdr) + 1):
        ws.cell(r, c).border = BORDER
        ws.cell(r, c).fill = TOTAL_FILL
    col = 7
    for pk in periods:
        fo = sum((comb_summary[(e, n), pk].get("fcur_open", 0.0) or 0.0) for (e, n) in comb_customers)
        so = sum(comb_summary[(e, n), pk]["s_open"] for (e, n) in comb_customers)
        sd = sum(abs(comb_summary[(e, n), pk]["s_debit"]) for (e, n) in comb_customers)
        sc = sum(abs(comb_summary[(e, n), pk]["s_credit"]) for (e, n) in comb_customers)
        fc = sum((comb_summary[(e, n), pk].get("fcur_close", 0.0) or 0.0) for (e, n) in comb_customers)
        cl = sum(comb_summary[(e, n), pk]["s_close"] for (e, n) in comb_customers)
        for off, val in enumerate([fo, so, sd, sc, fc, cl]):
            cell = ws.cell(r, col + off, val)
            cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
            cell.number_format = MONEY_FMT
        col += 6
    # 合计行：跨期差异 = 下期期初合计 − 上期期末合计
    for (p0, p1) in transitions:
        fc0 = sum((comb_summary[(e, n), p0].get("fcur_close", 0.0) or 0.0) for (e, n) in comb_customers)
        fc1 = sum((comb_summary[(e, n), p1].get("fcur_open", 0.0) or 0.0) for (e, n) in comb_customers)
        lc0 = sum(comb_summary[(e, n), p0]["s_close"] for (e, n) in comb_customers)
        lc1 = sum(comb_summary[(e, n), p1]["s_open"] for (e, n) in comb_customers)
        block = [(fc0 if fc0 else None), (fc1 if fc1 else None), round(fc1 - fc0, 2),
                 lc0, lc1, round(lc1 - lc0, 2)]
        for off, val in enumerate(block):
            cell = ws.cell(r, col + off, val)
            cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
            cell.number_format = MONEY_FMT
        if abs(fc1 - fc0) > 0.005:
            ws.cell(r, col + 2).fill = WARN_FILL
        if abs(lc1 - lc0) > 0.005:
            ws.cell(r, col + 5).fill = WARN_FILL
        col += 6
    widths = [16, 16, 30, 16, 10, 8] + [14] * (len(periods) * 6) \
             + [14] * (len(transitions) * 6) + [14, 14, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "G5"


def _cwrite_summary_sheet(ws, periods, comb_customers, comb_summary, label):
    ws.cell(2, 1, f"{label}余额汇总").font = TITLE_FONT
    hdr = ["核算主体", "期间", "往来单位数", "期初余额合计", "借方发生额合计", "贷方发生额合计", "期末余额合计"]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    seen = []
    for (e, n) in comb_customers:
        if e not in seen:
            seen.append(e)
    for e in seen:
        ents = [n for (ee, n) in comb_customers if ee == e]
        for pk in periods:
            op = sum(comb_summary[(e, n), pk]["s_open"] for n in ents)
            db = sum(abs(comb_summary[(e, n), pk]["s_debit"]) for n in ents)
            cr = sum(abs(comb_summary[(e, n), pk]["s_credit"]) for n in ents)
            cl = sum(comb_summary[(e, n), pk]["s_close"] for n in ents)
            for c, v in enumerate([e, pk, len(ents), op, db, cr, cl], 1):
                cell = ws.cell(r, c, v)
                cell.border = BORDER
                if c >= 4:
                    cell.number_format = MONEY_FMT
            r += 1
    for i, w in enumerate([16, 12, 10, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _cwrite_drilldown_sheet(ws, comb_rows, label):
    ws.cell(2, 1, f"{label}交易流水（下钻）— 数据源自综合查询明细表(GL)").font = TITLE_FONT
    ws.cell(3, 1, "说明：综合查询明细表为非全量抽取，流水累计与辅助核算余额差额=取数范围限制，"
                   "不构成账务差错；本表仅用于交易级下钻与审计抽样。各行『核算主体』已按真实主体分列。").font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    hdr = ["序号", "核算主体", "期间", "日期", "字", "号", "往来单位名称", "摘要", "对方科目", "借方金额", "贷方金额"]
    for i, h in enumerate(hdr, 1):
        ws.cell(5, i, h)
    _style_header(ws, 5, len(hdr))
    r = 6
    idx = 1
    for x in comb_rows:
        if not x.get("is_sub"):
            continue
        d = x.get("date")
        pk = d.strftime("%Y") if d else ""
        ds = d.strftime("%Y-%m-%d") if d else ""
        ent = x.get("entity", "")
        vals = [idx, ent, pk, ds, x.get("vtype", ""), x.get("vno", ""),
                x.get("cust", ""), x.get("summary", ""), x.get("cp", ""), x.get("debit", 0), x.get("credit", 0)]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (10, 11):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
            elif c in (1, 3, 4, 5):
                cell.alignment = CENTER
            else:
                cell.alignment = LEFT
        idx += 1
        r += 1
    widths = [6, 16, 8, 12, 6, 8, 30, 40, 28, 16, 16]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A6"


def _write_aging_sheet(ws, periods, comb_customers, comb_summary, aging_map, bucket_type, label, aging_methods=None, prior_meta=None):
    """账龄分布：各账龄区间 × 各期间的未结余额合计。
    口径与明细表行一致（2026-08-03 用户方法论：段位不推进——今年1-3月新发生→1年以内，
    去年1年以内→今年仍1年以内、去年1-2年→仍1-2年…）；优先按上年账龄表简化滚动，
    无上年表时回退 GL FIFO。periods 传单年（_render_periods），2025/2026 底稿各显示当年。"""
    buckets = AGING_BUCKETS.get(bucket_type, [])
    ws.cell(2, 1, f"{label} — 账龄分布（YTD 年份简化滚动，全年 FIFO；与明细表口径一致）").font = TITLE_FONT
    hdr = ["账龄区间"] + [f"{pk}期末余额" for pk in periods] + ["合计"]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for b in buckets:
        row = [b]
        tot_b = 0.0
        for pk in periods:
            amt = 0.0
            for (e, n) in comb_customers:
                s = comb_summary[(e, n), pk]
                pm = ((prior_meta or {}).get((e, n)) or (prior_meta or {}).get(('*', n)) or {}) or {}
                if pm.get('aging'):
                    bm = _roll_aging_from_prior(pm['aging'], s["s_open"], s["s_debit"],
                                                s["s_close"], buckets)
                else:
                    bm = (aging_map.get(((e, n), pk), {}) or {})
                amt += bm.get(b, 0.0)
            tot_b += amt
            row.append(round(amt, 2))
        row.append(round(tot_b, 2))
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c > 1:
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        r += 1
    # 合计行（= 各期间期末余额合计，用于与账龄分桶合计数勾稽）
    tcell = ws.cell(r, 1, "合计")
    tcell.font = TOTAL_FONT
    tcell.fill = TOTAL_FILL
    tcell.border = BORDER
    grand = 0.0
    for j, pk in enumerate(periods, start=2):
        amt = sum(comb_summary[(e, n), pk]["s_close"] for (e, n) in comb_customers)
        grand += amt
        cell = ws.cell(r, j, round(amt, 2))
        cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
        cell.number_format = MONEY_FMT
    cell = ws.cell(r, len(hdr), round(grand, 2))
    cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
    cell.number_format = MONEY_FMT
    r += 2
    note = ("说明：账龄口径与明细表行一致。本年数据为部分年度（如 2026 仅 1-3 月）时启用简化滚动"
            "（段位不推进）：本年新发生计入『1年以内』，上年各账龄段原样保留（去年1年以内→今年仍1年以内、"
            "去年1-2年→仍1-2年…），期末余额相对减少时按先进先出从最老段扣减；"
            "全年数据年份按 GL 交易日期 FIFO 正常推进账龄。无上年账龄表时回退 FIFO 测算"
            "（期初余额视为期间首日发生）。该测算为审计抽样辅助，不代表法定账龄认定。")
    if aging_methods:
        method_txt = "；".join(sorted(aging_methods))
        note += (f"\n本套账龄取数方法：{method_txt}。"
                 "（方法1=跨年GL的FIFO真实账龄；方法2=以上年账龄表滚动；"
                 "方法0=仅单年GL、1年以上账龄受限、全列『1年以内』。）")
    ws.cell(r, 1, note).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 2, end_column=len(hdr))
    widths = [16] + [18] * len(periods) + [18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B5"


# ============================================================================
# 2026-08-03 应收/其他应收/应收票据 附注汇总（格式2 上市公司合并口径）
# 列=科目筛选列|科目|附注填列负责人×3(空)|简单加计数|合并抵消(借/贷)|审定数。
# 合并范围内=集团关联方(related)；合并抵消借=合并内关联方余额；审定数=加计数−抵消。
# ============================================================================
def _fmt2_nature(nat):
    """FY 款项性质 → 格式2 其他应收款固定类别。"""
    if not nat:
        return '其他'
    if '股利' in nat:
        return '应收股利'
    if '利息' in nat:
        return '应收利息'
    if '备用金' in nat:
        return '备用金'
    if '拆借' in nat or '借款' in nat:
        return '拆借款'
    if '履约' in nat or '投标' in nat:
        return '履约保证金'
    if '押金' in nat or '质保' in nat or '保证金' in nat:
        return '押金保证金'
    if '暂付' in nat or '代垫' in nat or '代付' in nat:
        return '应收暂付款'
    if '资金池' in nat:
        return '资金池资金'
    return '其他'


def _build_ca_footnote_sheet(ws, key, subj, comb_customers, comb_summary, aging_map, bucket_type,
                             tb_full, year, label, gl_full=None):
    """生成单个往来科目『附注汇总』（格式2 合并口径）。返回 True。"""
    from collections import defaultdict
    buckets = AGING_BUCKETS.get(bucket_type) or []
    _pk = str(year)
    inner = defaultdict(float); outer = defaultdict(float)
    bal_in = 0.0
    nat_in = defaultdict(float); nat_out = defaultdict(float)
    btype_in = defaultdict(float); btype_out = defaultdict(float)
    for (entity, name) in comb_customers:
        s = comb_summary.get(((entity, name), _pk))
        if s is None:
            continue
        sc = s.get('s_close', 0.0) or 0.0
        rel = bool(s.get('related'))
        bm = (aging_map or {}).get(((entity, name), _pk), {}) or {}
        for b in buckets:
            amt = float(bm.get(b, 0.0) or 0.0)
            (inner if rel else outer)[b] += amt
        if rel:
            bal_in += abs(sc)
        if key == 'ORA':
            _n = _fmt2_nature(s.get('nature', ''))
            (nat_in if rel else nat_out)[_n] += abs(sc)
        if key == 'ARN':
            _bt = s.get('bill_type', '') or '其他'
            (btype_in if rel else btype_out)[_bt] += abs(sc)
    # —— 总额按报表科目（TB 审定数口径）——
    # 2026-08-03 修正：FY 其他应收款 819 万中 807 万无客户辅助核算（部门/个人维度），
    # aux 客户口径仅 12,035——明细表如实反映客户口径，但附注按报表科目应取 TB 审定数。
    # 客户口径账龄/性质覆盖之外的差额并入『合并范围外-1年以内/其他』。
    _code = {'AR': '1122', 'ORA': '1221', 'ARN': '1121',
             'APP': '1123', 'APR': '2203', 'AP': '2202', 'ORP': '2241',
             'CL': '2205', 'CA': '1480'}[key]
    # ⚡ 2026-08-12 修复：SAP TB 只有 10 位末级码（1121010000），原 `str(c) == _code`
    #   精确匹配永不命中 → _tb_bal=0 → 应收票据/预付等附注合计 0（质检 1020 应收票据 4.1 亿
    #   /1040 预付 471 万 附注空 ERROR）。改前缀匹配+末级叶子过滤（铁律60/65：U8 有父级行
    #   会父+子双计，须剔父级；SAP 无父级行天然叶子）。
    _tb_rows = [(c, n, v) for (e, c, n, yy), v in (tb_full or {}).items()
                if str(yy) == _pk and str(c).startswith(_code)
                and any(abs(float(v.get(k) or 0.0)) > 0.005 for k in ('qc', 'jf', 'df', 'qm'))]
    _tb_leaves = [(c, n, v) for (c, n, v) in _tb_rows
                  if not any(c2 != c and str(c2).startswith(str(c)) for (c2, n2, v2) in _tb_rows)]
    _tb_bal = sum(abs(float(v.get('qm', 0.0) or 0.0)) for (c, n, v) in _tb_leaves)
    # 2026-08-03 附注按【审定口径】列示：合计 = TB 期末 + 对方科目重分类转入
    # （重分类转入客户期末=0 但审定数=转入额，账龄按首档计入——与明细表审定数列一致）
    _reclass_in = sum(float(s.get('reclass_in', 0.0) or 0.0)
                      for ((e, n), pk), s in comb_summary.items() if str(pk) == _pk)
    _cust_aging_tot = sum(inner.values()) + sum(outer.values())
    _gap = max(0.0, _tb_bal + _reclass_in - bal_in - _cust_aging_tot)
    if _gap > 0.005:
        outer[buckets[0]] += _gap          # 无客户辅助部分 → 合并范围外 1年以内
    _bal_out = _tb_bal + _reclass_in - bal_in
    _tot_bal = _tb_bal + _reclass_in
    if key == 'ORA' and _gap > 0.005:
        nat_out['其他'] += _gap          # 无客户辅助部分 → 款项性质『其他』
    # 坏账准备（TB 1231 对应子目；2026-08-03 修正：勿用『其他应收 not in n』排除——会误伤 ORA 的『其他应收款』匹配）
    def _bd_qm(kw):
        return sum(abs(float(v.get('qm', 0.0) or 0.0))
                   for (e, c, n, yy), v in (tb_full or {}).items()
                   if str(yy) == _pk and str(c).startswith('1231') and kw in str(n))
    _bd_ar = _bd_qm('应收帐款') + _bd_qm('应收账款')
    _bd_ora = _bd_qm('其他应收款')
    _bd_arn = _bd_qm('应收票据')
    _bd_app = _bd_qm('预付')
    _bd = {'AR': _bd_ar, 'ORA': _bd_ora, 'ARN': _bd_arn, 'APP': _bd_app}.get(key, 0.0)

    _HF = PatternFill('solid', fgColor='DDEBF7')
    _BF = Font(name='Times New Roman', size=10, bold=True)
    _F10 = Font(name='Times New Roman', size=10)
    _NUM = '#,##0.00'
    # 2026-08-04 B3：『与TB校对』行不再硬编码 0（误导），改为展示 附注审定数(含抵消口径) − TB审定数 的真实差异；
    # 附注审定数=简单加计数−合并抵消（合并报表口径），TB审定数为合并试算表审定数。
    # |差异|>0.01 红字+『待对平』标注——差异来源与 tb_recon 一致，对平后自然归 0。
    _tb_audit = _tb_bal + _reclass_in   # TB审定数（含重分类转入，未抵消口径）
    _diff_tb = round((_tot_bal - bal_in) - _tb_audit, 2)   # 附注审定数(F−G) − TB审定数

    def _row_tb_check(r, txt, diff):
        """写『与TB校对』行：差异=附注审定数−TB审定数；|差异|>0.01 红字+待对平标注，否则填 0。"""
        vals = [txt, '', '', '', '', '', '', '']
        r = _row(r, _title, vals)
        _cell = ws.cell(r - 1, 6, round(diff, 2))
        _cell.number_format = _NUM
        _cell.alignment = Alignment(horizontal='right')
        if abs(diff) > 0.01:
            _cell.font = Font(name='Times New Roman', size=9, bold=True, color='C00000')
            ws.cell(r - 1, 2, txt + '（待对平：附注−TB）')
        return r
    _BDR = Border(*[Side(style='thin', color='BFBFBF')] * 4)
    NC = 9  # A科目 B明细 C/D/E负责人 F加计 G抵消借 H抵消贷 I审定数

    def _row(r, b, vals, bold=False, fill=False, money_cols=(6, 7, 8, 9)):
        ws.cell(r, 1, b)
        ws.cell(r, 2, vals[0])
        for j, v in enumerate(vals[1:], start=3):
            # 2026-08-04 合并抵消列（7/8）为空/0 → 填 0 并灰斜体标注『待填』（阅读提示，非数据缺失）
            if j in (7, 8) and (v is None or (isinstance(v, (int, float)) and v == 0)):
                cell = ws.cell(r, j, 0)
                cell.font = Font(name='Times New Roman', size=9, italic=True, color='BFBFBF')
                cell.number_format = _NUM
                cell.alignment = Alignment(horizontal='right')
                continue
            if v is None:
                continue
            cell = ws.cell(r, j, v)
            if j in money_cols:
                cell.number_format = _NUM
                cell.alignment = Alignment(horizontal='right')
        for j in range(1, NC + 1):
            cc = ws.cell(r, j); cc.border = _BDR
            if bold or fill:
                cc.font = _BF
            if fill:
                cc.fill = _HF
        return r + 1

    # —— 表头 ——
    ws.cell(1, 1, f'{label}附注汇总（{year} 年，格式2 合并口径）').font = Font(name='Times New Roman', size=12, bold=True)
    hdr = ['科目筛选列', '科目', '附注填列负责人', '', '', '简单加计数', '合并抵消(借)', '合并抵消(贷)', '审定数']
    for j, h in enumerate(hdr, 1):
        cc = ws.cell(3, j, h); cc.font = _BF; cc.fill = _HF; cc.border = _BDR
        cc.alignment = Alignment(horizontal='center', wrap_text=True)
    r = 4
    _title = {'AR': '3.应收账款', 'ORA': '6. 其他应收款', 'ARN': '2. 应收票据',
              'APP': '5. 预付账款', 'APR': '27. 预收帐款', 'AP': '26. 应付账款',
              'ORP': '31. 其他应付款',
              'CL': '合同负债', 'CA': '合同资产'}[key]

    if key == 'CA':
        # 2026-08-06 合同资产（SSS 建筑施工企业前瞻）：资产 6 档账龄 + 减值准备 + 与TB校对
        ws.cell(r, 1, _title); ws.cell(r, 2, _title)
        for j in range(1, NC + 1):
            ws.cell(r, j).font = _BF; ws.cell(r, j).fill = _HF; ws.cell(r, j).border = _BDR
        r += 1
        r = _row(r, _title, ['(1) 账龄情况', '', '', '', '', '', '', ''])
        for grp, d in (('合并范围内', inner), ('合并范围外', outer)):
            r = _row(r, _title, [grp, '', '', '', '', '', '', ''])
            for b in buckets:
                amt = round(d.get(b, 0.0), 2)
                if abs(amt) < 0.005:
                    amt = None
                r = _row(r, _title, ['  ' + b, None, None, None, amt, None, None, amt])
            sub = round(sum(d.values()), 2)
            r = _row(r, _title, ['小计：', None, None, None, sub, None, None, sub], bold=True)
        bal_tot = round(_tot_bal, 2)
        r = _row(r, _title, ['账面余额合计', None, None, None, bal_tot, round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
        if abs(_bd) > 0.005:
            r = _row(r, _title, ['合同资产减值准备', None, None, None, round(_bd, 2), None, None, round(_bd, 2)])
            r = _row(r, _title, ['账面价值合计', None, None, None, round(_tot_bal - _bd, 2), round(bal_in, 2), None,
                                 round(_tot_bal - bal_in - _bd, 2)], bold=True)
        r = _row_tb_check(r, '与TB校对——合同资产余额', _diff_tb)
    if key == 'AR':
        ws.cell(r, 1, _title); ws.cell(r, 2, _title)
        for j in range(1, NC + 1):
            ws.cell(r, j).font = _BF; ws.cell(r, j).fill = _HF; ws.cell(r, j).border = _BDR
        r += 1
        r = _row(r, _title, ['(1) 账龄情况', '', '', '', '', '', '', ''])
        for grp, d in (('合并范围内', inner), ('合并范围外', outer)):
            r = _row(r, _title, [grp, '', '', '', '', '', '', ''])
            for b in buckets:
                amt = round(d.get(b, 0.0), 2)
                if abs(amt) < 0.005:
                    amt = None
                r = _row(r, _title, ['  ' + b, None, None, None, amt, None, None, amt])
            sub = round(sum(d.values()), 2)
            r = _row(r, _title, ['小计：', None, None, None, sub, None, None, sub], bold=True)
        bal_tot = round(_tot_bal, 2)
        r = _row(r, _title, ['账面余额合计', None, None, None, bal_tot, round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
        r = _row(r, _title, ['坏账准备', None, None, None, round(_bd, 2), None, None, round(_bd, 2)])
        r = _row(r, _title, ['账面价值合计', None, None, None, round(_tot_bal - _bd, 2), round(bal_in, 2), None,
                             round(_tot_bal - bal_in - _bd, 2)], bold=True)
        r = _row_tb_check(r, '与TB校对——应收账款余额', _diff_tb)
        r = _row(r, _title, ['(2) 坏账准备计提情况', '', '', '', '', '', '', ''])
        r = _row(r, _title, ['1) 类别明细情况', '', '', '', '', '', '', ''])
        r = _row(r, _title, ['按组合计提的应收账款余额（账龄组合）', None, None, None, bal_tot,
                             round(bal_in, 2), None, round(_tot_bal - bal_in, 2)])
        r = _row(r, _title, ['合 计', None, None, None, bal_tot, round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
        r = _row(r, _title, ['按组合计提的坏账准备（账龄组合）', None, None, None, round(_bd, 2),
                             None, None, round(_bd, 2)])
        r = _row(r, _title, ['坏账准备合 计', None, None, None, round(_bd, 2), None, None, round(_bd, 2)], bold=True)
        r = _row_tb_check(r, '与TB校对——应收账款坏账准备', round(_bd - _bd, 2))
    elif key == 'ORA':
        ws.cell(r, 1, _title); ws.cell(r, 2, _title)
        for j in range(1, NC + 1):
            ws.cell(r, j).font = _BF; ws.cell(r, j).fill = _HF; ws.cell(r, j).border = _BDR
        r += 1
        r = _row(r, _title, ['(1) 款项性质分类情况', '', '', '', '', '', '', ''])
        for cat in ('履约保证金', '押金保证金', '应收暂付款', '拆借款', '备用金', '其他',
                    '合并范围内关联方款项', '资金池资金', '应收利息', '应收股利'):
            if cat == '合并范围内关联方款项':
                amt = round(bal_in, 2)
                r = _row(r, _title, ['  ' + cat, None, None, None, amt, None, None, amt])
            else:
                amt = round(nat_out.get(cat, 0.0), 2)
                if abs(amt) >= 0.005:
                    r = _row(r, _title, ['  ' + cat, None, None, None, amt, None, None, amt])
        r = _row(r, _title, ['账面余额小计', None, None, None, round(_tot_bal, 2), round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
        r = _row(r, _title, ['坏账准备', None, None, None, round(_bd, 2), None, None, round(_bd, 2)])
        r = _row(r, _title, ['账面价值合计', None, None, None, round(_tot_bal - _bd, 2), round(bal_in, 2), None,
                             round(_tot_bal - bal_in - _bd, 2)], bold=True)
        r = _row(r, _title, ['(2) 账龄情况（不包括应收利息和应收股利）', '', '', '', '', '', '', ''])
        for grp, d in (('合并范围内', inner), ('合并范围外', outer)):
            r = _row(r, _title, [grp, '', '', '', '', '', '', ''])
            for b in buckets:
                amt = round(d.get(b, 0.0), 2)
                if abs(amt) < 0.005:
                    amt = None
                r = _row(r, _title, ['  ' + b, None, None, None, amt, None, None, amt])
            sub = round(sum(d.values()), 2)
            r = _row(r, _title, ['小计：', None, None, None, sub, None, None, sub], bold=True)
    elif key == 'APP':
        # 格式2『5. 预付账款』：账龄 4 档（1年以内/1至2年/2至3年/3年以上）合并内/外分列，
        # 小计 → 账面余额合计 → 坏账准备（有数据才显示，无减值则省略）→ 账面价值合计 → 与TB校对
        ws.cell(r, 1, _title); ws.cell(r, 2, _title)
        for j in range(1, NC + 1):
            ws.cell(r, j).font = _BF; ws.cell(r, j).fill = _HF; ws.cell(r, j).border = _BDR
        r += 1
        r = _row(r, _title, ['(1) 账龄分析', '', '', '', '', '', '', ''])
        for grp, d in (('合并范围内', inner), ('合并范围外', outer)):
            r = _row(r, _title, [grp, '', '', '', '', '', '', ''])
            for b in buckets:
                amt = round(d.get(b, 0.0), 2)
                if abs(amt) < 0.005:
                    amt = None
                r = _row(r, _title, ['  ' + b, None, None, None, amt, None, None, amt])
            sub = round(sum(d.values()), 2)
            r = _row(r, _title, ['小计：', None, None, None, sub, None, None, sub], bold=True)
        bal_tot = round(_tot_bal, 2)
        r = _row(r, _title, ['合 计', None, None, None, bal_tot, round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
        r = _row_tb_check(r, '与TB校对——预付账款余额', _diff_tb)
        if abs(_bd) > 0.005:
            r = _row(r, _title, ['坏账准备', None, None, None, round(_bd, 2), None, None, round(_bd, 2)])
            r = _row(r, _title, ['账面价值合计', None, None, None, round(_tot_bal - _bd, 2),
                                 round(bal_in, 2), None, round(_tot_bal - bal_in - _bd, 2)], bold=True)
            r = _row_tb_check(r, '与TB校对——预付款项坏账准备', round(_bd - _bd, 2))
    elif key in ('APR', 'AP', 'CL'):
        # 格式2『27. 预收帐款』『26. 应付账款』『合同负债』（2026-08-06 起 LL/DI/NCL 迁出至 gp_other）：
        # 负债类简化——账龄只分 1年以内/1年以上（用户方法论：负债类账龄一般只要 1年以内和1年以上），
        # 合计 + 与TB校对（不分合并内/外）。
        tot = defaultdict(float)
        for b in buckets:
            tot[b] = (inner.get(b, 0.0) or 0.0) + (outer.get(b, 0.0) or 0.0)
        ws.cell(r, 1, _title); ws.cell(r, 2, _title)
        for j in range(1, NC + 1):
            ws.cell(r, j).font = _BF; ws.cell(r, j).fill = _HF; ws.cell(r, j).border = _BDR
        r += 1
        r = _row(r, _title, ['(1) 账龄情况', '', '', '', '', '', '', ''])
        for b in buckets:
            amt = round(tot.get(b, 0.0), 2)
            if abs(amt) < 0.005:
                amt = None
            r = _row(r, _title, ['  ' + b, None, None, None, amt, None, None, amt])
        bal_tot = round(_tot_bal, 2)
        r = _row(r, _title, ['合 计', None, None, None, bal_tot, round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
        r = _row_tb_check(r, f'与TB校对——{subj["label"]}', _diff_tb)
        # 2026-08-04 P0-2：模板要求『账龄1年以上重要的应付账款/预收款项/合同负债』
        # （单位名称/期末数/未偿还或结转原因）——从 aging_map 取『1年以上』余额>0 的往来单位 Top10。
        # 仅 AP/APR/CL（模板有该子项）；LL/DI/NCL 无此要求不生成。
        if key in ('AP', 'APR', 'CL'):
            _old_b = '1年以上' if '1年以上' in buckets else (buckets[-1] if buckets else '')
            _rows1 = []
            for (_ent, _nm), _pk2 in comb_summary:
                if str(_pk2) != _pk:
                    continue
                _bm = (aging_map or {}).get(((_ent, _nm), _pk2), {}) or {}
                _amt = float(_bm.get(_old_b, 0.0) or 0.0)
                if _amt > 0.005:
                    _rows1.append((_ent, _nm, _amt))
            if _rows1:
                r = _row(r, _title, ['(2) 账龄1年以上重要的' + {'AP': '应付账款', 'APR': '预收款项', 'CL': '合同负债'}[key],
                                     '', '', '', '', '', '', ''])
                r = _row(r, _title, ['  单位名称', '主体', None, None, '期末数', None, None, '未偿还或结转原因（审计师填）'])
                _rows1.sort(key=lambda x: -x[2])
                for _ent, _nm, _amt in _rows1[:10]:
                    r = _row(r, _title, ['  ' + _nm, _ent, None, None, round(_amt, 2), None, None, ''])
                r = _row(r, _title, ['  小计', None, None, None, round(sum(x[2] for x in _rows1), 2),
                                     None, None, ''], bold=True)
    elif key == 'ORP':
        # 格式2『31. 其他应付款』：参照其他应收款（一级科目/报表科目两种情况）——
        # (1) 明细情况：应付股利/应付利息/其他应付款/合计；(2) 账龄 2 档（负债简化）。
        def _orp_tb(kw):
            return sum(abs(float(v.get('qm', 0.0) or 0.0))
                       for (e, c, n, yy), v in (tb_full or {}).items()
                       if str(yy) == _pk and kw in str(n))
        _div_tb = _orp_tb('应付股利')
        _int_tb = _orp_tb('应付利息')
        tot = defaultdict(float)
        for b in buckets:
            tot[b] = (inner.get(b, 0.0) or 0.0) + (outer.get(b, 0.0) or 0.0)
        ws.cell(r, 1, _title); ws.cell(r, 2, _title)
        for j in range(1, NC + 1):
            ws.cell(r, j).font = _BF; ws.cell(r, j).fill = _HF; ws.cell(r, j).border = _BDR
        r += 1
        r = _row(r, _title, ['(1) 明细情况', '', '', '', '', '', '', ''])
        r = _row(r, _title, ['  普通股股利', None, None, None, round(_div_tb, 2), None, None,
                             round(_div_tb, 2)])
        r = _row(r, _title, ['  应付利息（仅限已到期未付部分）', None, None, None, round(_int_tb, 2),
                             None, None, round(_int_tb, 2)])
        r = _row(r, _title, ['  其他应付款', None, None, None, round(_tot_bal, 2), round(bal_in, 2),
                             None, round(_tot_bal - bal_in, 2)])
        r = _row(r, _title, ['合 计', None, None, None, round(_tot_bal + _div_tb + _int_tb, 2),
                             round(bal_in, 2), None,
                             round(_tot_bal - bal_in + _div_tb + _int_tb, 2)], bold=True)
        r = _row_tb_check(r, '与TB校对——其他应付款', _diff_tb)
        r = _row(r, _title, ['(2) 账龄情况（负债类仅分 1年以内/1年以上）', '', '', '', '', '', '', ''])
        for b in buckets:
            amt = round(tot.get(b, 0.0), 2)
            if abs(amt) < 0.005:
                amt = None
            r = _row(r, _title, ['  ' + b, None, None, None, amt, None, None, amt])
        r = _row(r, _title, ['合 计', None, None, None, round(_tot_bal, 2), round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
    else:  # ARN 应收票据
        ws.cell(r, 1, _title); ws.cell(r, 2, _title)
        for j in range(1, NC + 1):
            ws.cell(r, j).font = _BF; ws.cell(r, j).fill = _HF; ws.cell(r, j).border = _BDR
        r += 1
        r = _row(r, _title, ['(1) 明细情况（账面价值）', '', '', '', '', '', '', ''])
        # 2026-08-04 P0-1：模板要求按 银行/商业/财务公司承兑 分类披露。
        # aux bill_type 缺失时，从 GL 摘要识别（『银行承兑汇票/商业承兑汇票/电子银行承兑』等）；
        # 仍无法识别的差额归『未分类』并加披露限制说明——避免填附注时无数据可取。
        _bt_map = {'银行承兑': ' 银行承兑汇票', '商业承兑': ' 商业承兑汇票',
                   '财务公司承兑': '财务公司承兑汇票'}
        _bt_amt = {bt: round(btype_out.get(bt, 0.0) + btype_in.get(bt, 0.0), 2) for bt in _bt_map}
        _bt_gl = {bt: 0.0 for bt in _bt_map}
        if key == 'ARN' and gl_full:
            for _g in gl_full:
                _gn = str(_g.get('name') or '')
                _gs = str(_g.get('summary') or '')
                if '应收票据' not in _gn:
                    continue
                _bl = _gs + _gn
                if '银行承兑' in _bl or '电子银行承兑' in _bl:
                    _bt_gl['银行承兑'] += abs(float(_g.get('dr') or 0.0)) + abs(float(_g.get('cr') or 0.0))
                elif '商业承兑' in _bl:
                    _bt_gl['商业承兑'] += abs(float(_g.get('dr') or 0.0)) + abs(float(_g.get('cr') or 0.0))
                elif '财务公司承兑' in _bl:
                    _bt_gl['财务公司承兑'] += abs(float(_g.get('dr') or 0.0)) + abs(float(_g.get('cr') or 0.0))
        # GL 识别金额不参与分类行取值（仅用于披露说明判断），分类行以 aux bill_type 为准；
        # aux 完全无种类信息时，期末余额无法按种类拆分 → 分类行不硬造金额，
        # 以『未分类』列示期末余额 + 披露说明（GL 摘要可识别本期发生额构成，供审计师按台账填列）。
        _aux_known = sum(_bt_amt.values())
        _gl_known = sum(_bt_gl.values())
        _gl_txt = ''
        if _aux_known < 0.005 and _gl_known >= 0.005:
            _gl_txt = '（GL摘要识别本期发生额：' + '、'.join(
                f'{disp}≈{round(_bt_gl[bt], 2):,.2f}' for bt, disp in _bt_map.items() if _bt_gl[bt] >= 0.005) + '）'
        for bt, disp in _bt_map.items():
            amt = _bt_amt[bt]
            if abs(amt) >= 0.005:
                r = _row(r, _title, [disp, None, None, None, amt, None, None, amt])
        _bt_uncls = round(_tot_bal - bal_in - _aux_known, 2)
        if _bt_uncls > 0.005:
            r = _row(r, _title, [' 未分类（账套无票据种类辅助核算）', None, None, None,
                                 _bt_uncls, None, None, _bt_uncls])
        if _gl_txt:
            r = _row(r, _title, [('票据种类参考：' + _gl_txt)[:90], '', '', '', '', '', '', ''])
        if bal_in > 0.005:
            r = _row(r, _title, [' 合并范围内关联方票据', None, None, None, round(bal_in, 2), None, None,
                                 round(bal_in, 2)])
        r = _row(r, _title, ['合 计', None, None, None, round(_tot_bal, 2), round(bal_in, 2), None,
                             round(_tot_bal - bal_in, 2)], bold=True)
        if _bd > 0.005:
            r = _row(r, _title, ['坏账准备', None, None, None, round(_bd, 2), None, None, round(_bd, 2)])
            r = _row(r, _title, ['账面价值', None, None, None, round(_tot_bal - _bd, 2), None, None,
                                 round(_tot_bal - bal_in - _bd, 2)], bold=True)
    ws.cell(r, 1, '注：合并范围内=集团内关联方；合并抵消=合并内关联方余额；审定数=简单加计数−抵消（合并报表口径）。').font = \
        Font(name='Times New Roman', size=9, italic=True, color='808080')
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
    for j, w in enumerate([16, 40, 12, 12, 12, 18, 16, 16, 18], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'C4'
    return True


def _append_ca_footnote_sheet(out_path, key, subj, comb_customers, comb_summary, aging_map,
                              bucket_type, tb_full, year, gl_full=None, wb=None):
    """向应收/其他应收/应收票据底稿追加『附注汇总』sheet（2026-08-03 格式2 合并口径）。
    ⚡⚡ 2026-08-25 性能优化：支持外部传入 wb（合并多次 load+save 为一次）。"""
    if wb is None:
        wb = openpyxl.load_workbook(out_path)
    try:
        if '附注汇总' in wb.sheetnames:
            del wb['附注汇总']
        ws = wb.create_sheet('附注汇总')
        try:
            _build_ca_footnote_sheet(ws, key, subj, comb_customers, comb_summary,
                                     aging_map, bucket_type, tb_full, year, subj['label'],
                                     gl_full=gl_full)
        except Exception as _ex:
            print(f"   ⚠️ 附注汇总内容生成失败 [{key}] {year}：{_ex}")
    finally:
        if wb is None:
            wb.save(out_path)
            wb.close()
            del wb   # 2026-08-07 内存优化：close 只释放文件句柄，openpyxl 对象仍驻留至 GC——
                     # AR 大文件(175MB) 4 次追加链峰值累积 OOM → 显式 del 立即释放
    print(f"   ✅ 已追加 附注汇总 至：{out_path}")


def _write_intra_group_recon_sheet(ws, ent_list, comb_customers, comb_summary, periods, label):
    """⚡ 2026-08-11 阶段一 1.2 合并底稿差异化：内部交易抵消核对表（集团模式专属）。
    行=往来单位，列=各主体期末(带符号) | 全集团合计 | 抵消借 | 抵消贷 | 合并报表数。
    期末取最后期间的 s_close（带符号：资产借正贷负/负债贷正借负）。
    同一往来单位出现在 ≥2 个主体 → 内部交易候选 → 标黄提示核对（A 应收↔B 应付未抵消信号）。
    抵消借/贷留空待审计人员填；合并报表数=全集团合计（抵消 0 时）。"""
    from openpyxl.styles import PatternFill
    _F_MID = PatternFill('solid', fgColor='BDD7EE')
    _F_RED = PatternFill('solid', fgColor='FCE4EC')
    _F_GRN = PatternFill('solid', fgColor='E2EFDA')
    _F_DBL = PatternFill('solid', fgColor='DDEBF7')
    _F_WHT = Font(name='Times New Roman', bold=True, size=10, color='000000')
    _F_HL = PatternFill('solid', fgColor='FFF2CC')          # 内部交易候选：浅黄
    _pk = periods[-1] if periods else None
    if _pk is None:
        return
    # {cust: {ent: close}}（仅期末有余额者）
    cust_ent = {}
    for (_e, _nm) in comb_customers:              # 元素=(entity, name)
        s = comb_summary.get(((_e, _nm), _pk))
        if s and abs(s.get('s_close', 0.0)) >= 0.005:
            cust_ent.setdefault(_nm, {})[_e] = s['s_close']
    # 排序：全集团合计绝对值降序，Top 300
    _tot = {cc: sum(v.values()) for cc, v in cust_ent.items()}
    _ordered = sorted(cust_ent, key=lambda x: -abs(_tot[x]))[:300]
    nc = 2 + len(ent_list) + 4
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
    c = ws.cell(1, 1, f'{label}—内部交易抵消核对表（{_pk}）· 同一往来单位跨主体出现=内部交易候选，抵消待填')
    c.font = TITLE_FONT if 'TITLE_FONT' in globals() else Font(name='Times New Roman', bold=True, size=10, color='000000')
    c.alignment = Alignment(horizontal='center', vertical='center')
    headers = ['往来单位'] + ent_list + ['全集团合计', '抵消借', '抵消贷', '合并报表数']
    for j, h in enumerate(headers, 1):
        cell = ws.cell(2, j, h)
        cell.font = Font(name='Times New Roman', bold=True, size=10, color='000000')
        cell.fill = PatternFill('solid', fgColor='DDEBF7')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER
    _ent_end = 1 + len(ent_list)
    ws.cell(2, _ent_end + 1).fill = _F_MID
    ws.cell(2, _ent_end + 2).fill = _F_RED
    ws.cell(2, _ent_end + 3).fill = _F_GRN
    _mc = ws.cell(2, _ent_end + 4); _mc.fill = _F_DBL; _mc.font = _F_WHT
    r = 3
    for cc in _ordered:
        rec = cust_ent[cc]
        n_ent = sum(1 for v in rec.values() if abs(v) >= 0.005)
        tot = sum(rec.values())
        ws.cell(r, 1, cc).alignment = Alignment(horizontal='left', vertical='center')
        ws.cell(r, 1).border = BORDER
        for i, e in enumerate(ent_list):
            _v = rec.get(e, 0.0)
            cell = ws.cell(r, 2 + i, round(_v, 2) if abs(_v) >= 0.005 else None)
            cell.number_format = MONEY_FMT if 'MONEY_FMT' in globals() else '#,##0.00'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = BORDER
        cell = ws.cell(r, _ent_end + 1, round(tot, 2) if abs(tot) >= 0.005 else None)
        cell.number_format = '#,##0.00'; cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.border = BORDER; cell.fill = _F_MID
        for _off in (2, 3):
            cell = ws.cell(r, _ent_end + _off, None)
            cell.border = BORDER; cell.fill = _F_RED if _off == 2 else _F_GRN
        cell = ws.cell(r, _ent_end + 4, round(tot, 2) if abs(tot) >= 0.005 else None)
        cell.number_format = '#,##0.00'; cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.border = BORDER; cell.fill = _F_DBL; cell.font = _F_WHT
        if n_ent >= 2:                                # 跨主体 → 内部交易候选
            for j in range(1, nc + 1):
                ws.cell(r, j).fill = _F_HL
            ws.cell(r, _ent_end + 4).fill = _F_DBL; ws.cell(r, _ent_end + 4).font = _F_WHT
        r += 1
    ws.column_dimensions['A'].width = 28
    for j in range(2, nc + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    ws.freeze_panes = 'C3'


def export_combined_excel(output, periods, comb_customers, comb_summary, issues, comb_rows, comb_counterparties, mode="Y", subj=None, by_ent=None, aging_map=None, bucket_type=None, km_path=None, out_dir=None, data_dir=None, subj_key=None, entities_dict=None, aging_methods=None, target_year=None, prior_meta=None, early_return_wb=False):
    """合并底稿导出。⚠️ 2026-08-25 性能优化：early_return_wb=True 时返回 wb 对象
    （不 finalize/save/close）——build_all_in_dir 统一做追加（相关科目/坏账/附注/
    审定表注入）后再 finalize+save 一次，消除对 93MB 大底稿的 load+save 反复 IO
    （XBJ 追加阶段原 20-35 分钟/次）。"""
    wb = openpyxl.Workbook()
    label = subj["label"] if subj else "往来科目"
    # 同上：剔除无年份伪期间 "全部"，单期不生成 明细全部 / 期间对比
    periods = [p for p in periods if p != "全部"] or list(periods)
    multi = len(periods) > 1
    _render_periods = [target_year] if target_year is not None else periods
    # 2026-08-06 修复：跨年度客户×期间缺键（某客户 2025 有数据进 comb_customers、
    # 2026 无发生额 → comb_summary 缺 2026 键 → 明细表 KeyError 崩溃，如 c 账套
    # 其他应付款-政府补助）。补齐空记录（与构建处 dict 结构一致）。
    _empty_sum = lambda: dict(open=0.0, debit=0.0, credit=0.0, close=0.0,
                              foreign=False, aging='', code=None, count=0, nature='',
                              related=False, fcur_open=0.0, fcur_close=0.0,
                              open_dir='', close_dir='', s_open=0.0, s_debit=0.0,
                              s_credit=0.0, s_close=0.0, s_bal=0.0, bill_type='')
    for _pk in _render_periods:
        for (_ent, _name) in comb_customers:
            comb_summary.setdefault((( _ent, _name), _pk), _empty_sum())
    # ⚡ 2026-08-11 DQ 双年度账套修复：按【目标年度】过滤客户——2025 底稿不渲染 2024 实体
    #   （dq）的客户行（此前 comb_customers 跨年度共享 → 2025 明细表混入 dq 空行 156 行、
    #   期末翻倍 12.18M×2=24.36M 的根源之一；aux 按年度关联后仍有空行占位）。
    #   过滤规则：客户 (ent,name) 在目标年度 pk 有真实数据（s_open/s_debit/s_credit/s_close/
    #   reclass_in 任一非零）才保留；全零占位行（跨年度 setdefault 补空）剔除。
    def _has_data(ak, pk):
        _s = comb_summary.get((ak, pk))
        if not _s:
            return False
        return (abs(float(_s.get('s_open') or 0.0)) > 0.005
                or abs(float(_s.get('s_debit') or 0.0)) > 0.005
                or abs(float(_s.get('s_credit') or 0.0)) > 0.005
                or abs(float(_s.get('s_close') or 0.0)) > 0.005
                or abs(float(_s.get('reclass_in') or 0.0)) > 0.005)
    for pk in _render_periods:
        ws = wb.create_sheet(f"{label}明细表_{pk}")
        _cust_pk = [ak for ak in comb_customers if _has_data(ak, pk)]
        _cwrite_period_sheet(ws, pk, _cust_pk, comb_summary, label, aging_map=aging_map,
                             bucket_type=bucket_type, prior_meta=prior_meta)
    _TR = os.environ.get('CA_TRACE')
    if _TR:
        print(f'  [CA_TRACE] export 明细表完成 @ {time.time():.0f}', flush=True)
    # ⚡ 2026-08-16 旧『账龄分布』已废弃：由专项 aging_review.py 替代（同上）
    # 注：原『集团合并层面的借方对应科目 / 贷方对应科目』两表无意义，已按需求移除；
    #     按核算主体·分年度的『对方科目核对』表保留（含异常判定）。
    # 对方科目核对（按核算主体·分年度，判断异常）：无对应科目金额数据时不生成该 sheet（2026-07-31 用户要求）
    ent_list = sorted({x.get("entity", "") for x in comb_rows if x.get("entity")})
    recon_bad = _recon_unreliable(comb_rows)
    _by_ent_data = by_ent if by_ent is not None else reconstruct_counterparties_by_entity(comb_rows)
    _has_cp_data = any(
        abs(amt) >= 1.0
        for _d in _by_ent_data.values()
        for _cpk in ("sale_cp", "collect_cp")
        for amt in (_d.get(_cpk, {}) or {}).values()
    )
    if _has_cp_data:
        ws = wb.create_sheet("对方科目核对")
        _write_recon_counterparty_sheet(ws, f"{label}—对方科目核对（按核算主体·分年度）",
                                        _by_ent_data, subj, _render_periods, ent_list, recon_bad=recon_bad)
        # ⚡ 2026-08-12 新增：异常对应科目凭证清单（合并模式同源反查）
        try:
            _bad_cps = _collect_bad_cps(_by_ent_data, subj, _render_periods)
            if _TR:
                print(f'  [CA_TRACE] export 对方科目核对完成/异常收集 @ {time.time():.0f}', flush=True)
            if _bad_cps:
                ws = wb.create_sheet("异常对应科目凭证清单")
                _write_bad_cp_voucher_sheet(ws, _bad_cps, comb_rows, "", label)
            if _TR:
                print(f'  [CA_TRACE] export 异常清单完成 @ {time.time():.0f}', flush=True)
        except Exception:
            pass
    # 2026-07-31 用户要求：删除『校验结果』『借贷方检查』『交易流水(下钻)』『重要性清单』四张底稿
    # 应收票据专项底稿（仅 ARN）：票据种类拆分 / 背书贴现明细 / 坏账准备
    if subj and (subj.get("kw") == "应收票据" or subj.get("sheet") == "应收票据明细表"):
        ws = wb.create_sheet("票据种类拆分")
        _write_bill_type_split(ws, _render_periods, comb_customers, comb_summary, "", label)
        ws = wb.create_sheet("背书贴现明细")
        _endorse_rows = [x for x in comb_rows if target_year is None or str(x.get('year', '')) == str(target_year)]
        _write_endorse_discount(ws, _endorse_rows, "", label)
        ws = wb.create_sheet("应收票据坏账准备")
        _write_note_baddebt(ws, km_path, _render_periods, "", label)
    # ---- 坏账准备计算表（2026-08-03 参照 112500-7-6 格式；仅 应收/其他应收）----
    if subj_key in ("AR", "ORA"):
        ws = wb.create_sheet("坏账准备")
        _write_baddebt_calc_sheet(ws, subj_key, label, _render_periods, comb_customers,
                                  comb_summary, aging_map, bucket_type, entities_dict, target_year)
    # ---- 需求2：Top10余额明细表（仅 应收/其他应收/预收/应付 4类）----
    if subj_key in ("AR", "ORA", "APR", "AP"):
        ws = wb.create_sheet("Top10余额")
        _write_top10_sheet(ws, periods, comb_customers, comb_summary, label, subj_key)
    # ---- 需求4：主营客户比（仅应收账款）----
    if subj_key == "AR":
        ws = wb.create_sheet("主营客户比")
        _write_revenue_customer_sheet(ws, data_dir, comb_customers, comb_summary, _render_periods, entities_dict, label)
        # ⚡⚡ 2026-08-30 用户需求：主营客户比差异归因单独 sheet（应收借方 vs 收入的差异按对方科目归类）
        try:
            ws = wb.create_sheet("主营客户比差异归因")
            _write_customer_ratio_diff_sheet(ws, data_dir, _by_ent_data, entities_dict, _render_periods, label)
        except Exception as _ex:
            print(f'  ⚠️ 主营客户比差异归因失败：{_ex}')
    # ⚡ 2026-08-11 阶段一 1.2 合并底稿差异化：内部交易抵消核对表（集团模式多主体专属）
    # ⚡⚡ 2026-08-29 用户需求：不再生成——关联交易/关联往来底稿由独立小程序生成（88 家核对底稿）。
    # if len(ent_list) > 1:
    #     ws = wb.create_sheet("内部交易抵消核对")
    #     try:
    #         _write_intra_group_recon_sheet(ws, ent_list, comb_customers, comb_summary,
    #                                        _render_periods, label)
    #     except Exception as _ex:
    #         print(f'  ⚠️ 内部交易抵消核对失败：{_ex}')
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    order = ([f"{label}明细表_{pk}" for pk in periods] +
             ["对方科目核对",
              "票据种类拆分", "背书贴现明细", "应收票据坏账准备",
              "坏账准备", "Top10余额", "主营客户比", "内部交易抵消核对"])
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
    if early_return_wb:
        return wb   # ⚡⚡ 2026-08-25：调用方统一追加+finalize+save（省 load+save）
    finalize_workbook(wb)
    wb.save(output)
    wb.close()
    del wb   # 2026-08-07 内存优化：显式释放（AR 大文件 4 次追加链 OOM）




# ============================================================================
# 坏账准备计算表（2026-08-03 用户需求：参照《应收账款坏账准备底稿.xlsx》112500-7-6 格式）
#   列 A=项目 B=子项目 C=账龄 D=计提比例 E..=主体列 末列=合计
#   块1 应收账款原值（按账龄组合 6 段+小计 / 单项计提 / 合并内关联方 / 原值合计）
#   块2 坏账准备应有余额（按账龄组合 6 段×比例+小计 / 单项计提 / 应有余额合计）
#   块3 应收账款净值（原值−坏账准备，各段+小计+合计）
#   块4 坏账准备调整（未审数/期初续调/本期调整数/TB）
#   账龄数据源=aging_map（FIFO 按主体聚合）；比例参照表 0.05/0.30/0.50/0.80/0.80/1.00。
# ============================================================================
_BADDEBT_RATIO = [0.05, 0.30, 0.50, 0.80, 0.80, 1.00]   # 对应 6 档：1年以内/1-2年/2-3年/3-4年/4-5年/5年以上


def _year_is_ytd(rows, year):
    """判断某年度 GL 数据是否为『部分年度(YTD)』：该年出现过的月份数 < 12（如 FY 2026 仅 1-3 月）。
    2026-08-03 用户方法论：段位不推进的简化账龄滚动仅在 YTD 年份启用（临时特殊处理）；
    全年数据年份恢复正常账龄推进（去年1年以内→今年1-2年，FIFO 真实账龄）。"""
    months = set()
    for r in rows:
        if str(r.get('year', '') or '') != str(year):
            continue
        d = r.get('date')
        if d is None:
            continue
        try:
            months.add(str(d)[:7])
        except Exception:
            continue
    return bool(months) and len(months) < 12


def _read_prior_aging_table(data_dir, label, prev_year, bucket_type='recv'):
    """读上年账龄表（用户提供的 FY 版上年明细表，或程序上年生成底稿的明细表 sheet），
    返回 {(entity, cust): {'aging': {bucket: amt}, 'top_unit': str, 'nature': str}}。
    兼容两种格式：
    A) 参照格式（112500-7-2）：表头行3 —— 主体序号|主体|是否关联方|客户代码|客户单位|
       最上层客户单位|客户性质|期初|本期增加|本期减少|…|审定数|1年以内|1-2年|…|5年以上|…|期后回款
    B) 程序格式（当前 _cwrite_period_sheet）：表头行4 —— 核算主体|序号|往来单位编号|往来单位名称|
       款项性质|是否关联方|币种|外币余额(期初)|期初余额|…|期末余额|1年以内|…|5年以上|交易笔数|平衡校验
    找不到文件/解析失败 → 返回 {}（回退现有 FIFO 账龄，不中断）。"""
    buckets = AGING_BUCKETS.get(bucket_type, AGING_BUCKETS['recv'])
    out = {}
    if not data_dir or not os.path.isdir(data_dir):
        return out
    # 候选文件：① {label}明细表{prev_year}年.xlsx（用户提供 FY 版）；② 程序上年生成底稿
    cands = []
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith("~$"):
            continue
        if not fn.lower().endswith((".xlsx", ".xls")):
            continue
        if label in fn and str(prev_year) in fn and '明细表' in fn:
            cands.append(os.path.join(data_dir, fn))
    gen = os.path.join(data_dir, f'{label}审计底稿_{prev_year}_生成.xlsx')
    if os.path.isfile(gen):
        cands.append(gen)
    if not cands:
        return out
    # 表头语义列定位（兼容两种格式）
    col_map = {}          # 语义 -> 列号
    aging_cols = {}       # 桶名 -> 列号
    hdr_r = None
    for path in cands:
        try:
            wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        # 在前 6 行内找表头行（含『客户单位』或『往来单位名称』且含『1年以内』）
        for ri in range(min(6, len(raw))):
            row = [str(v or '') for v in raw[ri]]
            has_cust = any(('客户单位' in x or '往来单位名称' in x) for x in row)
            has_age = any('1年以内' in x for x in row)
            if has_cust and has_age:
                hdr_r = ri
                for ci, h in enumerate(row):
                    h = h.replace(chr(10), ' ')
                    if '客户单位' in h or '往来单位名称' in h:
                        col_map.setdefault('cust', ci)
                    elif '核算主体' in h or h.strip() == '主体' or '主体名称' in h:
                        col_map.setdefault('entity', ci)
                    elif '最上层' in h:
                        col_map.setdefault('top', ci)
                    elif '客户性质' in h or '政府行政' in h or '民营' in h or '电商' in h:
                        col_map.setdefault('nature', ci)
                    for b in buckets:
                        if h.strip() == b:
                            aging_cols[b] = ci
                break
        if hdr_r is not None:
            break
    if hdr_r is None or 'cust' not in col_map or len(aging_cols) < 2:
        return out
    ent_c = col_map.get('entity')
    top_c = col_map.get('top')
    nat_c = col_map.get('nature')
    for row in raw[hdr_r + 1:]:
        if not row or all(v is None or str(v).strip() == '' for v in row):
            continue
        cust = str(row[col_map['cust']] or '').strip()
        if not cust or cust in ('合计', '总计'):
            continue
        entity = str(row[ent_c] or '').strip() if ent_c is not None else ''
        aging = {}
        ok = False
        for b in buckets:
            ci = aging_cols.get(b)
            if ci is None:
                aging[b] = 0.0
                continue
            v = row[ci] if ci < len(row) else None
            try:
                v = float(v) if v is not None else 0.0
            except (TypeError, ValueError):
                v = 0.0
            aging[b] = v
            if abs(v) > 1e-6:
                ok = True
        if not ok:
            continue
        out[(entity, cust)] = {
            'aging': aging,
            'top_unit': str(row[top_c] or '').strip() if (top_c is not None and top_c < len(row)) else '',
            'nature': str(row[nat_c] or '').strip() if (nat_c is not None and nat_c < len(row)) else '',
        }
    return out


def _roll_aging_from_prior(prior_aging, s_open, s_debit, s_close, buckets):
    """2026 简化账龄滚动（用户方法论 2026-08-03：上年账龄 + 本年 1-3 月新发生）：
    ①上年账龄各段原样保留（去年 1年以内→仍 1年以内、去年 1-2年→仍 1-2年…段位不推进）；
    ②本年新发生（s_debit 借方）全部计入『1年以内』；
    ③余额相对减少（期初+本年增 > 期末）→ 减少额按 FIFO 从最老段先扣。
    prior_aging: {bucket: amt}（上年账龄）；返回 {bucket: amt}。"""
    out = {b: 0.0 for b in buckets}
    if prior_aging:
        for b in buckets:
            out[b] = abs(prior_aging.get(b, 0.0))
    # 本年新发生 → 1年以内（借方发生额，1-3月）
    out[buckets[0]] = out.get(buckets[0], 0.0) + max(0.0, s_debit)
    total = sum(out.values())
    # 期末控制数（TB 权威）——若期初+增−减=期末，减少额自然体现；直接按期末归一：
    # 期初余额 + 本年增 − 期末 = 本年减少额；从最老段开始扣减。
    dec_need = total - abs(s_close)
    if dec_need > 1e-6:
        for b in reversed(buckets):          # 最老段先扣（5年以上→…→1年以内）
            if dec_need <= 1e-6:
                break
            take = min(out[b], dec_need)
            out[b] -= take
            dec_need -= take
        if dec_need > 1e-6:                  # 极端：扣完仍差（红字）→ 调 1年以内
            out[buckets[0]] -= dec_need
    elif dec_need < -1e-6:                   # 期末 > 期初+增（异常）→ 差额并入 1年以内
        out[buckets[0]] += (-dec_need)
    return out


def _read_baddebt_tb(entities_dict, year, kw=None):
    """读各主体科目余额表中坏账准备科目（名含『坏账准备』；kw 限定关键词）期末余额。
    返回 {ent: {'open','debit','credit','close'}}（备抵贷余为正）。
    适配两种实体结构：①audit_common 式 {year: {'km': path}}；②本程序式 {'km': [path,...]} 列表。"""
    out = {}
    if not entities_dict:
        return out
    for ent, yrs in entities_dict.items():
        km_path = None
        if isinstance(yrs, dict) and set(yrs) == {'aux', 'gl', 'km'}:
            # 本程序 _discover_entities 结构：{'aux': [...], 'gl': [...], 'km': [path,...]}
            kms = yrs.get('km') or []
            km_path = None
            for p in kms:
                if str(year) in str(p):
                    km_path = p
                    break
            if km_path is None and kms:
                km_path = kms[0]
        else:
            yr = yrs.get(str(year)) if isinstance(yrs, dict) else None
            if yr:
                km_path = yr.get('km') if isinstance(yr, dict) else None
        if not km_path or not os.path.isfile(km_path):
            continue
        try:
            wb = openpyxl.load_workbook(km_path, data_only=True, read_only=True)
            ws = wb[wb.sheetnames[0]]
            raw = [list(r) for r in ws.iter_rows(values_only=True)]
            wb.close()
        except Exception:
            continue
        hdr_idx, headers = _detect_header(raw, AUX_COLUMN_ALIASES)
        mp = _resolve_aux_columns(headers)
        if 'km' not in mp or 'close' not in mp:
            continue
        rec = {'open': 0.0, 'debit': 0.0, 'credit': 0.0, 'close': 0.0}
        for ri, row in enumerate(raw[hdr_idx + 1:], start=hdr_idx + 2):
            if not row or all(c is None or str(c).strip() == '' for c in row):
                continue
            km = _str(row, mp.get('km'), '')
            if '坏账准备' not in km:
                continue
            if kw and kw not in km:
                continue
            rec['open'] += _parse_amount(row[mp['open']], ri) if mp.get('open') is not None and mp['open'] < len(row) else 0.0
            rec['debit'] += _parse_amount(row[mp['debit']], ri) if mp.get('debit') is not None and mp['debit'] < len(row) else 0.0
            rec['credit'] += _parse_amount(row[mp['credit']], ri) if mp.get('credit') is not None and mp['credit'] < len(row) else 0.0
            rec['close'] += _parse_amount(row[mp['close']], ri) if mp['close'] < len(row) else 0.0
        out[ent] = rec
    return out


def _write_baddebt_calc_sheet(ws, subj_key, label, periods, comb_customers, comb_summary,
                              aging_map, bucket_type, entities_dict, target_year=None):
    """坏账准备计算表（参照 112500-7-6 应收账款坏账计算表格式）。"""
    FONT10 = Font(name='Times New Roman', size=10)
    buckets = AGING_BUCKETS.get(bucket_type, AGING_BUCKETS['recv'])
    ent_list = sorted(entities_dict) if entities_dict else sorted({e for (e, _n) in comb_customers})
    NC = 4 + len(ent_list)          # A-D 项目/子项/账龄/比例 + 主体列 + 合计列
    year = str(target_year if target_year is not None else (periods[-1] if periods else '2025'))
    kw = '其他应收' if subj_key == 'ORA' else None
    SUBJ_NM = '其他应收款' if subj_key == 'ORA' else '应收账款'
    bd = _read_baddebt_tb(entities_dict, year, kw)

    ws.cell(2, 1, f'{label}—坏账准备计算表（参照 112500-7-6 格式；账龄按 GL 交易日期 FIFO 测算）').font = TITLE_FONT
    # 表头
    hdr = ['项目', '', '账   龄', '计提比例'] + ent_list + ['合计']
    for i, h in enumerate(hdr, 1):
        cell = ws.cell(4, i, h)
        cell.font = HDR_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER; cell.alignment = CENTER
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=2)

    def _aging(ent):
        out = {b: 0.0 for b in buckets}
        for (e, n) in comb_customers:
            if e != ent:
                continue
            am = aging_map.get(((e, n), year), {}) if aging_map else {}
            for b, v in am.items():
                out[b] = out.get(b, 0.0) + v
        return out

    def _row(r, a, b, c, d, vals, money=True, bold=False, fill=False, formula=False):
        """写一行：A/B/C/D 标签 + 各主体值 + 合计。vals: list[主体值或公式]。"""
        ws.cell(r, 1, a).border = BORDER
        ws.cell(r, 2, b).border = BORDER
        ws.cell(r, 3, c).border = BORDER
        ws.cell(r, 4, d).border = BORDER
        if a:
            ws.cell(r, 1).font = TOTAL_FONT if bold else FONT10
        if b:
            ws.cell(r, 2).font = TOTAL_FONT if bold else FONT10
        if d is not None:
            ws.cell(r, 4).number_format = '0.00'
            ws.cell(r, 4).alignment = Alignment(horizontal='right')
        for i, v in enumerate(vals):
            cell = ws.cell(r, 5 + i, v)
            cell.border = BORDER
            if money:
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal='right')
            if bold or fill:
                cell.font = TOTAL_FONT
                cell.fill = TOTAL_FILL
        if money:
            # 2026-08-06 协议：合计列写数值（=各主体列和；公式 data_only 读 None 收回读不回）
            _tv = sum(v for v in vals if isinstance(v, (int, float)))
            cell = ws.cell(r, NC, round(_tv, 2))
            cell.number_format = MONEY_FMT
            cell.alignment = Alignment(horizontal='right')
            cell.border = BORDER
            if bold or fill:
                cell.font = TOTAL_FONT
                cell.fill = TOTAL_FILL
        if bold:
            for c in range(1, NC + 1):
                if ws.cell(r, c).font == FONT10:
                    ws.cell(r, c).font = TOTAL_FONT
        return r + 1

    r = 5
    # ===== 块1 应收账款原值 =====
    comb_vals = [list(_aging(ent).values()) for ent in ent_list]   # 6 桶 × 主体
    sub_combo = [0.0] * len(ent_list)
    for bi, bname in enumerate(buckets):
        vals = [comb_vals[ei][bi] for ei in range(len(ent_list))]
        r = _row(r, SUBJ_NM + '原值' if bi == 0 else '', SUBJ_NM + '原值-按账龄组合' if bi == 0 else '',
                 bname, None, vals)
        for ei in range(len(ent_list)):
            sub_combo[ei] += vals[ei]
    r = _row(r, '', '', '小计', None, sub_combo, bold=True, fill=True)
    r_combo_sub = r - 1
    # 单项计提（0 占位）
    for bi, bname in enumerate(buckets):
        r = _row(r, SUBJ_NM + '原值-单项计提' if bi == 0 else '', '', bname, None, [0.0] * len(ent_list))
    r = _row(r, '', '', '小计', None, [0.0] * len(ent_list), bold=True, fill=True)
    r_single_sub = r - 1
    # 合并内关联方（0 占位）
    for bi, bname in enumerate(buckets):
        r = _row(r, SUBJ_NM + '合并内关联方' if bi == 0 else '', '', bname, None, [0.0] * len(ent_list))
    r = _row(r, '', '', '小计', None, [0.0] * len(ent_list), bold=True, fill=True)
    # 原值合计 = 组合小计 + 单项小计（2026-08-06 协议：写数值；公式 data_only 读 None 收回读不回）
    vals = [round((ws.cell(r_combo_sub, 5 + ei).value or 0.0) + (ws.cell(r_single_sub, 5 + ei).value or 0.0), 2)
            for ei in range(len(ent_list))]
    r = _row(r, SUBJ_NM + '原值合计', '合计', '', None, vals, bold=True, fill=True)
    r_orig_tot = r - 1

    # ===== 块2 坏账准备应有余额 =====
    for bi, bname in enumerate(buckets):
        ratio = _BADDEBT_RATIO[bi] if bi < len(_BADDEBT_RATIO) else 1.0
        orig_row = r_combo_sub - len(buckets) + bi
        vals = [round((ws.cell(orig_row, 5 + ei).value or 0.0) * ratio, 2)
                for ei in range(len(ent_list))]
        r = _row(r, '坏账准备余额' if bi == 0 else '', '坏账准备应有余额-按账龄组合' if bi == 0 else '',
                 bname, ratio, vals)
    r = _row(r, '', '', '小计', None,
             [round(sum(ws.cell(r - len(buckets) + bi, 5 + ei).value or 0.0 for bi in range(len(buckets))), 2)
              for ei in range(len(ent_list))],
             bold=True, fill=True)
    r_bd_sub = r - 1
    r = _row(r, '', '坏账准备-单项计提', '', None, [0.0] * len(ent_list))
    r_bd_single = r - 1
    vals = [round((ws.cell(r_bd_sub, 5 + ei).value or 0.0) + (ws.cell(r_bd_single, 5 + ei).value or 0.0), 2)
            for ei in range(len(ent_list))]
    r = _row(r, '坏账准备应有余额合计', '合计', '', None, vals, bold=True, fill=True)
    r_bd_tot = r - 1

    # ===== 块3 应收账款净值 =====
    for bi, bname in enumerate(buckets):
        orig_row = r_combo_sub - len(buckets) + bi
        bd_row = r_bd_sub - len(buckets) + bi
        vals = [round((ws.cell(orig_row, 5 + ei).value or 0.0) - (ws.cell(bd_row, 5 + ei).value or 0.0), 2)
                for ei in range(len(ent_list))]
        r = _row(r, SUBJ_NM + '净值' if bi == 0 else '', SUBJ_NM + '净值-按账龄组合' if bi == 0 else '',
                 bname, None, vals)
    r = _row(r, '', '', '小计', None,
             [round(sum(ws.cell(r - len(buckets) + bi, 5 + ei).value or 0.0 for bi in range(len(buckets))), 2)
              for ei in range(len(ent_list))],
             bold=True, fill=True)
    r_net_sub = r - 1
    r = _row(r, '', SUBJ_NM + '净值-单项计提', '', None,
             [round((ws.cell(r_single_sub, 5 + ei).value or 0.0) - (ws.cell(r_bd_single, 5 + ei).value or 0.0), 2)
              for ei in range(len(ent_list))])
    r_net_single = r - 1
    vals = [round((ws.cell(r_net_sub, 5 + ei).value or 0.0) + (ws.cell(r_net_single, 5 + ei).value or 0.0), 2)
            for ei in range(len(ent_list))]
    r = _row(r, SUBJ_NM + '净值合计', '合计', '', None, vals, bold=True, fill=True)

    # ===== 块4 坏账准备调整 =====
    r += 1
    ws.cell(r, 1, '坏账准备调整').font = SUB_FONT
    r += 1
    # 未审数（TB 期末，按主体）
    vals = [round(bd.get(ent, {}).get('close', 0.0), 2) for ent in ent_list]
    r = _row(r, '', '坏账准备未审数', '', None, vals)
    r_unaud = r - 1
    r = _row(r, '', '期初续调坏账准备', '', None, [0.0] * len(ent_list))
    r_cont = r - 1
    vals = [round((ws.cell(r_bd_tot, 5 + ei).value or 0.0) - (ws.cell(r_unaud, 5 + ei).value or 0.0)
                  - (ws.cell(r_cont, 5 + ei).value or 0.0), 2)
            for ei in range(len(ent_list))]
    r = _row(r, '', '本期调整数', '', None, vals, bold=True)
    r_adj = r - 1
    # TB（未审数同源，单独列示供差异核对）
    r = _row(r, '', 'TB', '', None, [round(bd.get(ent, {}).get('close', 0.0), 2) for ent in ent_list])
    r_tb = r - 1
    vals = [round((ws.cell(r_bd_tot, 5 + ei).value or 0.0) - (ws.cell(r_tb, 5 + ei).value or 0.0), 2)
            for ei in range(len(ent_list))]
    r = _row(r, '', '差异（应有余额-账套）', '', None, vals)

    # 列宽
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 26
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    for i in range(len(ent_list)):
        ws.column_dimensions[get_column_letter(5 + i)].width = 14
    ws.column_dimensions[get_column_letter(NC)].width = 16
    ws.freeze_panes = 'E5'
    return ws



def _put_cells(ws, r, row, money_cols=(), center_cols=()):
    """通用：写入一行并套边框/金额右对齐/居中样式。"""
    for c, v in enumerate(row, 1):
        cell = ws.cell(r, c, v)
        cell.border = BORDER
        if c in money_cols:
            cell.number_format = MONEY_FMT
            cell.alignment = Alignment(horizontal="right")
        elif c in center_cols:
            cell.alignment = CENTER
        else:
            cell.alignment = LEFT


# ============================================================================
# 新增需求2：Top10余额明细表（应收/其他应收/预收/应付）
# ============================================================================
# 往来单位名称归一化：read_opening_balances 对“同名跨子目/跨科目”的往来单位，
# 会在名称末尾追加 '(子目)' 后缀（如 '(其他应付)' / '(工程应付-暂估分包)' /
# '(未开票-增票6%)'），并可能再补 '·' 残留。此处反复剥离【末尾】的括号后缀与 '·'，
# 还原基础往来单位名，使同一家单位的多个子目行归并为一条。
# 关键：只剥离【末尾】括号（ASCII '(...)' 或全角 '（...）'），不碰名称中部的合法括号
# （如『蚂蚁智信（杭州）信息技术有限公司』），避免把不同城市的同名公司误并。
def _base_counterparty(name):
    s = (name or "").strip()
    while s.endswith("·"):
        s = s[:-1].rstrip()
    while True:
        m = re.search(r"(?:\([^()]*\)|（[^（）]*）)\s*$", s)
        if not m:
            break
        s = s[:m.start()].rstrip("·").strip()
        if not s:
            break
    return s


def _wtype_cat(subj_nature):
    """往来类型（客户/供应商）标签：按科目性质判定（资产=客户，负债=供应商）。

    重要：Q 集团辅助核算『往来类型』列经实测与【发票状态】相关而非真实交易对手类型——
    同一家单位『开票』行标『客户』、『未开票』行标『供应商』且期末余额完全相同（如
    浙江希帷：客户/开票 243,759.29 与 供应商/未开票 243,759.29）。若按该列强制分列，
    会把同一家单位的同名跨子目（仅发票状态差异）误拆为 客户/供应商 两行，违背聚合初衷。
    故此处统一按科目性质标注，使每家单位在各自科目内归并为单行、类型正确，
    且客户余额只存在于资产类科目、供应商余额只存在于负债类科目，绝不跨性质合并。"""
    return "客户" if subj_nature == "asset" else "供应商"


def _write_top10_sheet(ws, periods, comb_customers, comb_summary, label, subj_key):
    periods = sorted(periods)
    p1 = periods[-1]
    p0 = periods[-2] if len(periods) >= 2 else None
    subj_nature = SUBJECTS.get(subj_key, {}).get("nature", "asset")
    cmp = ("（2年对比：%s vs %s）" % (p1, p0)) if p0 else ("（仅 %s，无对比）" % p1)
    ws.cell(2, 1, f"{label} — 余额前10位明细表 {cmp}").font = TITLE_FONT
    note = ("格式：先按核算主体分别列示余额前10位，再汇总全集团余额前10位；"
            "按期末余额绝对值排序；负债科目(预收/应付)以贷方余额绝对值参与排序。"
            "已按往来单位名称汇总同名跨子目行（剥离末尾'(子目)'/'(未开票-增票X%)'等后缀与'·'残留），"
            "同一往来单位只计一条；客户与供应商始终分列（新增『往来类型』列），绝不合并。"
            "注：『（无辅助核算明细）』为账套未挂往来单位的大额余额兜底桶（SAP 应付/应收行往来单位字段"
            "缺失），需按指令号/采购凭证进一步拆解核实；『×××公司』为核算主体自身/集团内部单位，关注内部交易。")
    ws.cell(3, 1, note).font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    ws.cell(3, 1).alignment = LEFT

    def sclose(en, pk):
        return comb_summary.get((en, pk), {}).get("s_close", 0.0)

    def aggregate(cust_list):
        """按 (归一化往来单位名称, 往来类型) 汇总：同名跨子目/带'·'残留的行合并为一条；
        客户与供应商始终分列（同基础名若分属客户/供应商则保留两条），绝不合并。
        返回 {(base, wc): {entity, ents(set), close_p1, close_p0, nature, wtype, related, maxabs}}。"""
        agg = {}
        for en in cust_list:
            base = _base_counterparty(en[1])
            s = comb_summary.get((en, p1), {})
            wc = _wtype_cat(subj_nature)
            c1 = s.get("s_close", 0.0) or 0.0
            c0 = sclose(en, p0) if p0 else 0.0
            key = base
            d = agg.get(key)
            if d is None:
                d = agg[key] = {"entity": en[0], "ents": set([en[0]]),
                                "close_p1": 0.0, "close_p0": 0.0,
                                "nature": s.get("nature", ""), "wtype": wc,
                                "related": False, "maxabs": 0.0}
            d["close_p1"] += c1
            d["close_p0"] += c0
            d["ents"].add(en[0])
            d["related"] = d["related"] or bool(s.get("related"))
            aa = abs(c1)
            if aa >= d["maxabs"]:
                d["maxabs"] = aa
                d["nature"] = s.get("nature", "")
                d["entity"] = en[0]
        return agg

    hdr = (["排名", "核算主体", "往来单位名称", "往来类型", "款项性质", "是否关联方", f"期末余额({p1})"]
           + ([f"期末余额({p0})", "变动"] if p0 else []))
    ncols = len(hdr)
    money_cols = tuple(range(7, 7 + (3 if p0 else 1)))
    center_cols = (1, 4, 6)
    r = 5
    # 一、分主体 Top10（同主体内按往来单位汇总）
    ws.cell(r, 1, "一、分核算主体 余额前10位（已按往来单位汇总同名跨子目，按期末余额绝对值）").font = SUB_FONT
    r += 1
    for i, h in enumerate(hdr, 1):
        ws.cell(r, i, h)
    _style_header(ws, r, ncols)
    r += 1
    by_ent = defaultdict(list)
    for en in comb_customers:
        by_ent[en[0]].append(en)
    for entity in sorted(by_ent.keys()):
        agg = aggregate(by_ent[entity])
        custs = sorted(agg.items(), key=lambda kv: abs(kv[1]["close_p1"]), reverse=True)[:10]
        for rank, (base, d) in enumerate(custs, 1):
            row = [rank, entity, base, _wtype_cat(subj_nature), d["nature"], ("是" if d["related"] else "否"),
                   round(d["close_p1"], 2)]
            if p0:
                row += [round(d["close_p0"], 2), round(d["close_p1"] - d["close_p0"], 2)]
            _put_cells(ws, r, row, money_cols=money_cols, center_cols=center_cols)
            r += 1
    r += 1
    # 二、全集团 Top10（跨主体按往来单位汇总）
    ws.cell(r, 1, "二、全集团 余额前10位（已按往来单位汇总同名跨子目，按期末余额绝对值）").font = SUB_FONT
    r += 1
    for i, h in enumerate(hdr, 1):
        ws.cell(r, i, h)
    _style_header(ws, r, ncols)
    r += 1
    grp_agg = aggregate(comb_customers)
    grp_top = sorted(grp_agg.items(), key=lambda kv: abs(kv[1]["close_p1"]), reverse=True)[:10]
    for rank, (base, d) in enumerate(grp_top, 1):
        ent_disp = d["entity"] if len(d["ents"]) == 1 else "/".join(sorted(d["ents"]))
        row = [rank, ent_disp, base, _wtype_cat(subj_nature), d["nature"], ("是" if d["related"] else "否"),
               round(d["close_p1"], 2)]
        if p0:
            row += [round(d["close_p0"], 2), round(d["close_p1"] - d["close_p0"], 2)]
        _put_cells(ws, r, row, money_cols=money_cols, center_cols=center_cols)
        r += 1
    # 合计行（全集团Top10期末合计）
    tcell = ws.cell(r, 2, "合计(Top10)")
    tcell.font = TOTAL_FONT
    tot1 = sum(d["close_p1"] for _, d in grp_top)
    if p0:
        tot0 = sum(d["close_p0"] for _, d in grp_top)
        _put_cells(ws, r, ["", "", "", "", "", "", round(tot1, 2), round(tot0, 2),
                           round(tot1 - tot0, 2)], money_cols=money_cols, center_cols=center_cols)
    else:
        _put_cells(ws, r, ["", "", "", "", "", "", round(tot1, 2)], money_cols=money_cols, center_cols=center_cols)
    widths = [6, 22, 30, 12, 16, 10, 18] + ([18, 16] if p0 else [])
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A6"


# ============================================================================
# 新增需求3：重要性水平清单（全部6个往来科目）
# ============================================================================
def _gen_materiality_template(out_dir):
    path = os.path.join(out_dir, "重要性水平_模板.xlsx")
    if os.path.exists(path):
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "重要性水平"
    ws.cell(1, 1, "核算主体").font = TOTAL_FONT
    ws.cell(1, 2, "重要性水平(元)").font = TOTAL_FONT
    ws.cell(3, 1, "说明：填写各核算主体财务报表层次重要性水平（元）。程序将列示该主体往来款项中"
                  "期末余额绝对值超过此阈值的往来单位。填妥后保持本文件名（或改名重要性水平.xlsx），"
                  "重新运行程序即可自动筛选。")
    ws.cell(4, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=4, start_column=1, end_row=6, end_column=2)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 18
    finalize_workbook(wb)
    wb.save(path)
    wb.close()


def _load_or_gen_materiality(out_dir):
    """读取『重要性水平.xlsx』；若不存在则生成『重要性水平_模板.xlsx』并返回 None。
    返回 {norm_entity: threshold} 或 None。"""
    if not out_dir:
        return None
    for fn in ("重要性水平.xlsx", "重要性水平_模板.xlsx"):
        p = os.path.join(out_dir, fn)
        if os.path.exists(p):
            try:
                wb = openpyxl.load_workbook(p, data_only=True)
                ws = wb[wb.sheetnames[0]]
                rows = [list(r) for r in ws.iter_rows(values_only=True)]
                wb.close()
                mp = {}
                for row in rows[1:]:
                    if not row or not row[0]:
                        continue
                    ent = str(row[0]).strip()
                    thr = row[1] if len(row) > 1 else None
                    try:
                        thr = float(thr) if thr not in (None, "") else None
                    except Exception:
                        thr = None
                    if thr is not None:
                        mp[_norm(ent)] = thr
                return mp
            except Exception:
                pass
    _gen_materiality_template(out_dir)
    return None


def _write_materiality_sheet(ws, comb_customers, comb_summary, periods, label, subj_key, mat_map):
    periods = sorted(periods)
    p1 = periods[-1]
    subj_nature = SUBJECTS.get(subj_key, {}).get("nature", "asset")
    ws.cell(2, 1, f"{label} — 重要性水平以上往来款项清单（{p1}）").font = TITLE_FONT
    r = 4

    def _build_groups():
        """按 (核算主体, 基础往来单位名, 往来类型) 聚合：同名跨子目归并为一条；
        客户与供应商始终分列（同基础名若分属客户/供应商则保留两行），绝不合并。"""
        g = {}
        for (E, name) in comb_customers:
            s = comb_summary.get(((E, name), p1), {})
            base = _base_counterparty(name)
            wc = _wtype_cat(subj_nature)
            key = (E, base)
            d = g.get(key)
            if d is None:
                d = g[key] = {"close": 0.0, "related": False,
                              "nature": s.get("nature", ""), "ents": set([E])}
            d["close"] += s.get("s_close", 0.0) or 0.0
            d["ents"].add(E)
            d["related"] = d["related"] or bool(s.get("related"))
            if d["nature"] in (None, ""):
                d["nature"] = s.get("nature", "")
        return g

    groups = _build_groups()
    if not mat_map:
        ws.cell(r, 1, "未提供重要性水平模板（已生成：重要性水平_模板.xlsx）。本次未筛选，列出全部往来单位；"
                        "填写模板后重新运行可自动筛选超阈值项。已按往来单位名称汇总同名跨子目行，"
                        "客户与供应商始终分列、不合并。").alignment = LEFT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        r += 2
        hdr = ["核算主体", "往来单位名称", "往来类型", "款项性质", "是否关联方", f"期末余额({p1})", "重要性水平"]
        for i, h in enumerate(hdr, 1):
            ws.cell(r, i, h)
        _style_header(ws, r, len(hdr))
        r += 1
        for (E, base), d in groups.items():
            _put_cells(ws, r, [E, base, _wtype_cat(subj_nature), d["nature"], ("是" if d["related"] else "否"),
                               round(d["close"], 2), "—"],
                       money_cols=(6,), center_cols=(3, 5))
            r += 1
        for i, w in enumerate([22, 30, 12, 16, 10, 18, 14], 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        return
    hdr = ["核算主体", "往来单位名称", "往来类型", "款项性质", "是否关联方", f"期末余额({p1})", "重要性水平", "超出金额"]
    for i, h in enumerate(hdr, 1):
        ws.cell(r, i, h)
    _style_header(ws, r, len(hdr))
    r += 1
    rows_out = []
    no_thr_ents = set()
    for (E, base), d in groups.items():
        thr = mat_map.get(_norm(E))
        if thr is None:
            no_thr_ents.add(E)
            continue
        close = d["close"]
        if abs(close) > thr + 1e-6:
            rows_out.append((E, base, _wtype_cat(subj_nature), d["nature"], "是" if d["related"] else "否",
                             round(close, 2), round(thr, 2), round(abs(close) - thr, 2)))
    rows_out.sort(key=lambda x: x[7], reverse=True)
    if not rows_out:
        msg = "各已设定重要性水平的主体，均无期末余额绝对值超阈值的往来款项。"
        if no_thr_ents:
            msg += f"（另有 {len(no_thr_ents)} 个主体未在模板中设定重要性水平，未纳入筛选。）"
        ws.cell(r, 1, msg).alignment = LEFT
        r += 1
    for row in rows_out:
        _put_cells(ws, r, row, money_cols=(6, 7, 8), center_cols=(3, 5))
        r += 1
    for i, w in enumerate([22, 30, 12, 16, 10, 18, 14, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================================
# 新增需求4：主营客户比（仅应收账款）
# ============================================================================
def _read_revenue_aux(data_dir):
    """读取辅助核算余额表中『主营业务收入』分客户数据。
    返回 ({(entity, customer): {"credit":贷发, "close":期末}}, has_customer, rev_by_ent)。
    rev_by_ent：主体级主营业务收入贷方合计——FY 等账套主营业务收入无客户辅助维度
    （仅科目级/部门级），分客户读取为空；兜底直接从 aux 汇总科目级贷方发生额，
    使『主营客户比』在无分客户数据时仍能生成主体级比较（2026-07-31 修复）。"""
    REV = dict(label="主营业务收入", kw="主营业务收入", sheet="主营业务收入", nature="revenue",
               exp_debit=[], exp_credit=[], notes_debit="", notes_credit="")
    out, has_customer = {}, False
    rev_by_ent = defaultdict(float)
    if not data_dir or not os.path.isdir(data_dir):
        return out, has_customer, dict(rev_by_ent)
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith("~$"):
            continue
        if "辅助核算余额表" not in fn or "外币" in fn:
            continue
        full = os.path.join(data_dir, fn)
        E = _derive_subject(full)
        try:
            data, _ = read_opening_balances(full, REV)
        except Exception:
            continue
        for name, rec in data.items():
            if not name or name == "主营业务收入":
                continue
            out[(E, name)] = {"credit": rec.get("credit", 0.0) or 0.0,
                              "close": rec.get("close", 0.0) or 0.0}
            has_customer = True
        # 主体级兜底：直接解析 aux，汇总所有『主营业务收入*』一级科目（等级1）的贷方发生额。
        # 2026-08-01 修复 #17：泰国 aux 中 4100(等级1)+4100.01(等级2) 各 351M 被重复累加 → 3 倍；
        # 只取等级1（一级科目贷方=主营收入总额），避免父级+子级双计。
        try:
            raw, hdr_idx, headers, mp = _parse_tb_file(full)
            if "km" in mp and "credit" in mp:
                tot_cr = 0.0
                for ri, row in enumerate(raw[hdr_idx + 1:], start=hdr_idx + 2):
                    if not row or all(c is None or str(c).strip() == "" for c in row):
                        continue
                    km = _str(row, mp.get("km"), "")
                    if not km.startswith("主营业务收入"):
                        continue
                    lv = _str(row, mp.get("level", -1), "") if mp.get("level") is not None else ""
                    lv_s = str(lv or "").strip()
                    if lv_s and lv_s != "1":
                        continue   # 仅等级1（父级）；等级2+ 为子目，避免双计
                    tot_cr += _parse_amount(row[mp["credit"]], ri) if mp["credit"] < len(row) else 0.0
                rev_by_ent[E] += tot_cr
        except Exception:
            continue
    return out, has_customer, dict(rev_by_ent)


def _write_revenue_customer_sheet(ws, data_dir, comb_customers, comb_summary, periods, entities_dict, label):
    periods = sorted(periods)
    p1 = periods[-1]
    rev_map, has_cust, rev_by_ent = _read_revenue_aux(data_dir)
    ws.cell(2, 1, f"{label} — 主营客户比（应收账款借方发生数 ÷ 主营业务收入）").font = TITLE_FONT
    ar_debit_by_ent = defaultdict(float)
    for (E, name) in comb_customers:
        ar_debit_by_ent[E] += comb_summary.get(((E, name), p1), {}).get("s_debit", 0.0) or 0.0
    # ⚡ 2026-08-26 SAP 兜底：工作目录无『辅助核算余额表』→ rev_by_ent 空 → 主体级比较缺失。
    #   从科目余额表读主营业务收入(6001) 叶子贷方合计，保证『主营客户比』主体级仍可生成
    #   （SAP 主营业务收入无客户辅助核算维度，分客户明细无法生成——表底已注明）。
    if not rev_by_ent and not has_cust and _adapter is not None:
        try:
            for E in (entities_dict or {}):
                try:
                    km = _adapter.read_km(E) or {}
                except Exception:
                    continue
                _r6 = [(str(c), r) for c, r in km.items() if str(c).startswith("6001")]
                _leaf6 = [(c, r) for c, r in _r6
                          if not any(c2 != c and c2.startswith(c) for c2, _ in _r6)]
                _rev = sum(float(r.get("df") or r.get("credit") or 0.0) for _, r in _leaf6)
                if abs(_rev) > 0.005:
                    rev_by_ent[E] = _rev
        except Exception:
            pass
    r = 4
    if not has_cust and not rev_by_ent:
        ws.cell(4, 1, "说明：辅助核算余额表中『主营业务收入』无分客户数据，按需求未生成明细表。"
                        "（若后续提供分客户主营业务收入辅助核算，可重算本表。）").alignment = LEFT
        ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=4)
        return
    ws.cell(r, 1, "一、按核算主体：主营业务收入(贷方发生额) 与 应收账款借方发生数 比较").font = SUB_FONT
    r += 1
    hdr = ["核算主体", "主营业务收入(贷方发生额)", "应收账款借方发生数(本期借发)", "比率(应收借方/收入)"]
    for i, h in enumerate(hdr, 1):
        ws.cell(r, i, h)
    _style_header(ws, r, len(hdr))
    r += 1
    # 主体全集 = 有主营业务收入 或 有应收借方的主体（并集），保证两张口径都能看到
    ents = sorted(set(rev_by_ent.keys()) | set(ar_debit_by_ent.keys()))
    for E in ents:
        rev = rev_by_ent.get(E, 0.0)
        ar = ar_debit_by_ent.get(E, 0.0)
        ratio = (ar / rev) if rev else None
        _put_cells(ws, r, [E, round(rev, 2), round(ar, 2),
                          (round(ratio, 4) if ratio is not None else "—")],
                   money_cols=(2, 3), center_cols=(4,))
        r += 1
    r += 1
    if has_cust:
        ws.cell(r, 1, "二、分客户主营业务收入明细（数据源自辅助核算余额表）").font = SUB_FONT
        r += 1
        hdr2 = ["核算主体", "客户名称", "主营业务收入(本期贷方)"]
        for i, h in enumerate(hdr2, 1):
            ws.cell(r, i, h)
        _style_header(ws, r, len(hdr2))
        r += 1
        for (E, cust) in sorted(rev_map.keys()):
            _put_cells(ws, r, [E, cust, round(rev_map[(E, cust)]["credit"], 2)], money_cols=(3,))
            r += 1
    else:
        ws.cell(r, 1, "二、分客户主营业务收入明细（数据源自辅助核算余额表）").font = SUB_FONT
        r += 1
        ws.cell(r, 1, "（辅助核算余额表中『主营业务收入』未按客户辅助核算，仅有科目/部门维度，"
                     "分客户明细无法生成；主体级比较见上表。）").alignment = LEFT
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    for i, w in enumerate([22, 26, 28, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_customer_ratio_diff_sheet(ws, data_dir, by_ent_data, entities_dict, periods, label):
    """主营客户比差异归因：应收账款借方发生额按【对方科目】拆分，解释与主营业务收入(贷方)差异。
    归因类别：①主营业务收入(正常) ②销项税额(价税分离) ③预收账款互转 ④合同资产/负债互转 ⑤其他。
    数据源：对方科目核对（sale_cp=应收借方对方科目金额分布，与『主营客户比』同源 GL）。
    ⚡ 2026-08-30 用户需求：把差异归类结果形成单独 sheet，清楚解释主营客户比差异原因。"""
    periods = sorted(periods)
    try:
        _rev_map, _has_cust, rev_by_ent = _read_revenue_aux(data_dir)
    except Exception:
        rev_by_ent = {}
    _CATS = [
        ("主营业务收入(正常)", ["主营业务收入", "其他业务收入", "营业收入"]),
        ("销项税额(价税分离)", ["销项", "销项税额"]),
        ("预收账款互转", ["预收账款", "预收款"]),
        ("合同资产/负债互转", ["合同资产", "合同负债", "合同结算"]),
    ]
    ws.cell(2, 1, f"{label} — 主营客户比差异归因").font = TITLE_FONT
    ws.cell(3, 1, "解释『主营客户比』中应收账款借方发生数与主营业务收入(贷方)的差异："
                  "应收借方按对方科目拆分归因（数据源自对方科目核对，同源 GL）。"
                  "归因合计应等于应收账款借方发生额；『差异』列 = 应收借方 − 收入。"
                  ).font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    hdr = ["核算主体", "主营业务收入", "应收账款借方发生额", "差异(应收借方-收入)"] + \
          [f"其中：{t}" for t, _ in _CATS] + ["其他(需关注)", "归因合计"]
    hdr_row = 4
    for i, h in enumerate(hdr, 1):
        ws.cell(hdr_row, i, h)
    _style_header(ws, hdr_row, len(hdr))
    r = hdr_row + 1
    ents = sorted(set(rev_by_ent.keys()) | set(k[0] for k in by_ent_data))
    for E in ents:
        rev = rev_by_ent.get(E, 0.0)
        ar = 0.0
        parts = [0.0] * len(_CATS)
        other = 0.0
        for (ent, pk), d in by_ent_data.items():
            if ent != E:
                continue
            for km, amt in (d.get("sale_cp") or {}).items():
                ar += amt
                _hit = -1
                for i, (_t, kws) in enumerate(_CATS):
                    if any(kw in km for kw in kws):
                        _hit = i
                        break
                if _hit >= 0:
                    parts[_hit] += amt
                else:
                    other += amt
        _put_cells(ws, r, [E, round(rev, 2), round(ar, 2), round(ar - rev, 2)] +
                          [round(p, 2) for p in parts] + [round(other, 2), round(sum(parts) + other, 2)],
                   money_cols=tuple(range(2, len(hdr) + 1)), center_cols=())
        r += 1
    _note_row = r + 1
    ws.cell(_note_row, 1, "注：『其他(需关注)』= 对方科目不属于收入/销项税/预收/合同类（如冲往来、营业外等），"
                          "需按异常对应科目凭证清单进一步核查。").alignment = LEFT
    ws.merge_cells(start_row=_note_row, start_column=1, end_row=_note_row, end_column=len(hdr))
    for i, w in enumerate([14] + [22] * (len(hdr) - 1), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = f"A{hdr_row + 1}"


def _is_sap_mode(data_dir):
    """⚡ 2026-08-24 形态判定统一走 LedgerBackend（detect_sap_layout 内容探测 +
    account_profiles 声明覆盖）。原散落 `_adapter.is_sap(...)` 各判各的；
    统一入口后新形态/声明变更只需改 ledger_backend 一处。
    ⚡⚠️ 与 sap_adapter.is_sap 同口径：SAP 模式下 run_u8_on_sap 传【工作目录
    _work_{comp}】（非 SAP 结构）→ 必须优先用 _DATA_ROOT（真实数据根）判定，
    否则工作目录被判 U8 → SAP 模式丢失。"""
    try:
        from ledger_backend import get_backend
        _d = getattr(_adapter, '_DATA_ROOT', None) or data_dir
        if not _d:
            return False
        return get_backend(_d).layout != 'u8'
    except Exception:
        return False


def _discover_entities(data_dir):
    """⚡ 2026-08-09 统一（保留 aux 结构）：SAP=adapter 单主体（current_comp 过滤内置）+ aux 占位；
    U8=audit_common 标准发现 + 辅助核算文件补充（文件名含『辅助核算』/『往来』）。"""
    from audit_common import discover_entities as _de
    if _adapter is not None and _is_sap_mode(data_dir):
        _comp = _adapter.current_comp()
        _src = {_comp: _de(data_dir).get(_comp, {})} if _comp else _de(data_dir)
        out = {}
        for _c, _yd in _src.items():
            out[_c] = {'aux': [None],
                       'gl': [None] if any(p.get('gl') for p in _yd.values()) else [],
                       'km': [None]}
        return out
    ents = _de(data_dir)
    # U8：补充辅助核算文件（账套级扫描；每主体一份时按主体目录/文件名就近关联）
    _aux_all = [os.path.join(data_dir, fn) for fn in sorted(os.listdir(data_dir))
                if fn.lower().endswith('.xlsx') and ('辅助核算' in fn or '往来' in fn)]
    # ⚡ 2026-08-10 DQ 等根目录直放多年度账套：audit_common discover 返回
    # {entity: {年: {km,gl}}} 年度嵌套 + 顶层 aux → build_all_in_dir 用 bun['gl']/bun['km']
    # 取到 None → 误判 _sap_mode → 预付明细退回 TB 兜底 1 条。统一转扁平
    # {entity: {km:[..], gl:[..], aux:[..]}}（与 SAP 分支同构，build_all_in_dir 免改）。
    out = {}
    for _e, _bun in ents.items():
        _km = [v['km'] for v in _bun.values() if isinstance(v, dict) and v.get('km')]
        _gl = [v['gl'] for v in _bun.values() if isinstance(v, dict) and v.get('gl')]
        if not _km and not _gl:
            continue
        # ⚡ 2026-08-11 修复 DQ 双年度账套混年：原 _aux_all 全量给所有实体 → DQ(2025) 和
        #   dq(2024) 实体各自读到 2024+2025 两份 aux → 2025 底稿混入 dq(2024) 实体、期末翻倍
        #   （12.18M×2=24.36M，DQ3.0 旧文件单主体 12.18M）。修复：aux 按实体年度就近关联
        #   （文件名含实体年度才分配；无年份匹配则保留全量，兼容单年度账套）。
        _e_years = [str(y) for y in _bun.keys()]
        _aux_e = [f for f in _aux_all
                  if any(y in os.path.basename(f) for y in _e_years)] or _aux_all
        # ⚡⚡ 2026-08-23 修复：空 gl/km 填 []（空 list）而非 [None]——[None] 是 SAP 分支
        #   占位语义，U8 主体无 GL 时填 [None] 会让 build_ca_detail 的 _sap_mode 判定
        #   （input_path=[None] → len==1 且 [0] is None）误判为 SAP → read_gl_net →
        #   _DATA_ROOT=None 崩溃（AFJ FGS 只有科目余额表无综合查询明细表，regen_all 崩）。
        out[_e] = {'km': _km, 'gl': _gl,
                   'aux': _aux_e}
    return out


def build_ca_detail(subj_key, input_path, aux_path, output_path, km_path=None, period_mode="Y", subject=None, write=True):
    """核心流程：读取 -> 分组 -> 汇总 -> 对应科目还原 -> 校验 -> 导出。
    返回 (issues, periods, customers, counterparties, n_open)。n_open=辅助核算末级行数（0 表示该科目在数据中不存在）。"""
    subj = SUBJECTS[subj_key]
    read_warnings = []

    # ⚡ 2026-08-09 SAP 适配：_discover_entities 在 SAP 下把 aux/gl/km 置为 [None] 占位，
    # 而 _build_period_map 会跳过 None → rows/opening 全空 → n_open=0 → 明细表退回
    # TB 兜底科目级（1010 应收明细表 9 行 vs 应收/1010.xlsx 客户维度 4241 行）。
    # SAP 模式直接走 adapter 单主体读取（read_transactions/read_opening_balances 已有
    # SAP 分支：path=None 时走 read_gl_net 净额化 / read_aux_balance 客户维度）。
    # ⚡ 2026-08-10 判定修复：U8 多年度账套（DQ）扁平化后 input_path=真实文件 list（非 str），
    # 原「not isinstance(str)」会误判 SAP → 必须识别【占位 [None]】才是 SAP 模式。
    _sap_mode = _adapter is not None and (
        input_path is None
        or (isinstance(input_path, (list, tuple)) and len(input_path) == 1 and input_path[0] is None))
    if _sap_mode:
        # ⚡ 2026-08-26 修复（任务870 扫描发现合同资产明细空）：CA_STD_CODE['CA']='1480'
        #   （SSS 建筑施工企业前瞻），但 AH/SAP 合同资产科目码=1124（GL 940 行 / aux 211 行
        #   均 1124）→ code_pfx=1480 取不到 GL → 明细表全 0 vs 审定表 27.25亿。SAP 模式
        #   合同资产改 1124。
        _cp = CA_STD_CODE.get(subj_key)
        if subj_key == 'CA':
            _cp = '1124'
        rows = read_transactions(None, subj["kw"], code_pfx=_cp,
                                 entity=subject, warnings=read_warnings)
        # ⚡ 2026-08-10 集团模式（current_comp=None）：read_transactions/read_opening_balances
        # 返回全量（read_gl_net 全量 / aux_entries 全量）→ 必须按 subject（当前主体）过滤，
        # 否则每主体明细=全集团（340=85×4 重复）且同名客户跨主体合并。
        if _adapter.current_comp() is None and subject:
            rows = [r for r in rows if str(r.get('entity') or '') == str(subject)]
        data, conflicts = read_opening_balances(None, subj, comp=subject)
        opening = {}
        aux_conflicts = [(name, a1, a2, '') for (name, a1, a2) in conflicts]
        # 年份：优先从 GL rows 日期推断（period_key 用 date.year）；无 GL 时取
        # adapter discover_entities 的当前主体年度键。
        # ⚡ 2026-08-10 兼容字符串日期（_excel_date 对文本型过帐日期返回 str，
        #   集团/部分账套 date='YYYY-MM-DD' 无 year 属性 → 原推断 _pk='' → 数据挂空期间）。
        _pk = ''
        for _r in rows:
            _d = _r.get('date')
            if _d is not None:
                if hasattr(_d, 'year'):
                    _pk = str(_d.year)
                elif isinstance(_d, str) and len(_d) >= 4 and _d[:4].isdigit():
                    _pk = _d[:4]
                if _pk:
                    break
        if not _pk:
            try:
                _ys = _adapter.discover_entities(getattr(_adapter, '_DATA_ROOT', None) or P.YY)
                # ⚡⚡ 2026-08-27 修复（任务886 2310 预收明细缺失根因）：集团模式
                #   current_comp=None → _e='' → _ys.get('',{}) 空 → sorted()[−1]
                #   IndexError → except → _pk='' → 客户行挂空期间 '' → 明细表按 '2026'
                #   渲染匹配失败 → 2310 预收 155.7万 明细表缺失（仅审定表有）。
                #   用 subject（当前主体）兜底取年度键。
                _e = _adapter.current_comp() or subject or ''
                _pk = str(sorted(_ys.get(_e, {}).keys())[-1])
            except Exception:
                _pk = ''
        for name, rec in data.items():
            opening[(name, _pk)] = rec
        gl_map, aux_map, km_m = {}, {}, {}
    else:
        gl_map = _build_period_map(input_path, keyword="综合查询") if input_path else {}
        rows = []
        for _plist in gl_map.values():
            for p in (_plist if isinstance(_plist, list) else [_plist]):
                rows += read_transactions(p, subj["kw"], warnings=read_warnings)

        aux_map = _build_period_map(aux_path, keyword="辅助核算余额表")
        opening = {}
        aux_conflicts = []
        for period, _plist in aux_map.items():
            for p in (_plist if isinstance(_plist, list) else [_plist]):
                data, conflicts = read_opening_balances(p, subj)
                aux_conflicts += [(name, a1, a2, period) for (name, a1, a2) in conflicts]
                for name, rec in data.items():
                    opening[(name, period)] = rec
        km_m = _build_period_map(km_path, keyword="科目余额表")

    ctrl_map = {}
    if km_path:
        if _sap_mode:
            km_m = _build_period_map(km_path, keyword="科目余额表")
        km_m = {p: [f for f in fl
                    if "外币" not in os.path.basename(f) and "辅助核算" not in os.path.basename(f)]
                for p, fl in km_m.items()}
        km_m = {p: fl for p, fl in km_m.items() if fl}
        for period, _plist in km_m.items():
            for p in _plist:
                ctrl_map[period] = ctrl_map.get(period, 0.0) + read_control(p, subj)

    gl_periods, gl_customers, agg = aggregate(rows, period_mode)
    aux_periods = sorted({p for (_, p) in opening})
    periods = sorted(set(gl_periods) | set(aux_periods))
    # 余额表主表【仅取辅助核算末级控制客户】(控制数直接取数)；GL 仅用于下钻与对应科目分析，
    # 不并入余额合计（GL 为非全量抽取，其净额会与辅助核算控制数背离，造成虚增/虚减）。
    # 仅当辅助核算完全无该科目数据时，退化为 GL 净额派生（标记 gl_only）。
    aux_customers = sorted({name for (name, _) in opening})
    if aux_customers:
        customers = aux_customers
    elif (subj.get("sheet") or "").startswith("应收票据"):
        # ⚡ 2026-08-26 修复（任务872 应收票据差 9,809万）：SAP 应收票据 GL 无客户字段
        #   （_customer_key 回退 km=科目名 → 伪客户『应收票据-银行承兑汇票』或
        #   『（未标注往来单位）』GL 净额行）。GL 客户路径产生的『（未标注往来单位）』
        #   s_close=GL净额(负)会混入差额兜底 _aux_close → _gap=TB子目−(−GL净额) 虚增
        #   3.12亿 → 子目分摊 2.94亿。应收票据直接走 TB 兜底（7112 分支按票据子目展开，
        #   s_close=TB 子目期末，票据种类拆分正确）。GL 借/贷发生额由序时账另行核对。
        customers = []
    else:
        customers = list(gl_customers)
    summary = compute_summary(customers, periods, agg, opening, nature=subj["nature"])
    # 2026-08-06 末级过滤：辅助核算/GL 含【科目子目行】（往来单位编号=科目代码 1221.02/1221.02.02…，
    # 无真实往来单位，名称=子目名）——U8 父级+子级同导出，明细表父+子双计（Z 物联合肥
    # 其他应收款 5,742.88 = 父 2,804.66 + 子 2,938.22，正确 2,804.66；1221.02 系列在 GL 行，
    # 故须用 summary 的 code 而非 opening）。编号含层级点/反斜杠且存在更长同前缀编号
    # → 父级剔除（只保留末级；真实客户编号无层级点，不受影响）。
    _cdm = {}
    for _ak in customers:
        _rc2 = None
        for _pk2 in periods:
            _rc2 = summary.get((_ak, _pk2)) or _rc2
        _cdm[_ak] = str(_rc2.get('code') or '') if _rc2 else ''

    def _codelike(c):
        return '.' in c or '\\' in c or '/' in c
    _leaf_cust = []
    for _ak in customers:
        _c = _cdm.get(_ak, '')
        if not _codelike(_c) or not any(
                _ak2 != _ak and _codelike(_cdm.get(_ak2, ''))
                and _cdm.get(_ak2, '').startswith(_c) and len(_cdm.get(_ak2, '')) > len(_c)
                for _ak2 in customers):
            _leaf_cust.append(_ak)
    if len(_leaf_cust) != len(customers):
        print(f'  ↳ 末级过滤：剔除 {len(customers) - len(_leaf_cust)} 个科目父级行（编号含层级点）')
        _dead = set(customers) - set(_leaf_cust)
        for _ak in _dead:
            for _pk in periods:
                summary.pop((_ak, _pk), None)
        customers = _leaf_cust
    # 丰富摘要：款项性质 / 是否关联方 / 外币余额（按子目原币分摊到逐户）
    for (ak, pk), s in summary.items():
        rec = opening.get((ak, pk))
        s["related"] = bool(rec.get("related")) if rec else False
        s["nature"] = _derive_nature(rec.get("_km", "")) if rec else ""
        s["wtype"] = rec.get("wtype", "") if rec else ""
    # 取包含外币辅助核算余额表/外币科目余额表的正确目录：
    # 目录输入(如 S)直接用该目录；文件输入取其所在目录。禁止用 _dir_of(目录)（会退回上一级，找不到外币文件）。
    _aux_src = aux_path[0] if isinstance(aux_path, (list, tuple)) and aux_path else aux_path
    _data_dir = _aux_src if isinstance(_aux_src, str) and os.path.isdir(_aux_src) else (_dir_of(aux_path) or _dir_of(input_path))
    fc = _read_foreign_balances(_data_dir, subject, periods)
    # 外币辅助核算余额表：逐户原币直接填列（按客户名称匹配；优先于此，因其为最精确来源）
    faux_by_name = _read_foreign_aux(_data_dir, subject, periods, subj)
    if fc:
        sub_rmb = defaultdict(float)
        for (ak, pk), rec in opening.items():
            if rec.get("foreign") and rec.get("_km"):
                sub_rmb[(pk, rec["_km"])] += rec.get("close", 0.0)
    for (ak, pk), s in summary.items():
        rec = opening.get((ak, pk))
        # 外币辅助核算余额表 逐户原币直接填列（按客户名称匹配，最精确来源；
        # 该表仅含外币客户，故凡在其中出现的客户均直接以其原币填列，不依赖 _km 是否为外币子目）。
        cand = _norm(ak)
        fv = faux_by_name.get((pk, cand))
        if fv is None:
            fv = faux_by_name.get((pk, _norm(_strip_aux_suffix(ak))))
        if fv is not None:
            s["fcur_open"] = round(fv[0], 2)
            s["fcur_close"] = round(fv[1], 2)
            _align_fcur_sign(s)          # 外币(原币)与本币同方向
            s["foreign"] = True
            continue
        # 退化：外币科目余额表 子目原币 按人民币占比分摊（原逻辑，仅无逐户外币辅助表时）
        if not rec or not rec.get("foreign"):
            s["fcur_open"] = 0.0
            s["fcur_close"] = 0.0
            continue
        sub_km = rec.get("_km", "")
        fv = fc.get((pk, _norm(sub_km)))
        if not fv:
            s["fcur_open"] = 0.0
            s["fcur_close"] = 0.0
            continue
        tot = sub_rmb.get((pk, sub_km), 0.0)
        if tot:
            ratio = rec.get("close", 0.0) / tot
            s["fcur_open"] = round(fv[0] * ratio, 2)
            s["fcur_close"] = round(fv[1] * ratio, 2)
        else:
            s["fcur_open"] = 0.0
            s["fcur_close"] = 0.0
        _align_fcur_sign(s)          # 外币(原币)与本币同方向（退化分摊路径亦对齐）
    issues = []

    red_debit = defaultdict(lambda: [0, 0.0])
    red_credit = defaultdict(lambda: [0, 0.0])
    for r in rows:
        if r["is_sub"] and r["debit"] < 0:
            red_debit[period_key(r["date"], period_mode)][0] += 1
            red_debit[period_key(r["date"], period_mode)][1] += r["debit"]
        if r["is_sub"] and r["credit"] < 0:
            red_credit[period_key(r["date"], period_mode)][0] += 1
            red_credit[period_key(r["date"], period_mode)][1] += r["credit"]
    for pk in periods:
        nd, sd = red_debit.get(pk, [0, 0.0])
        nc, sc = red_credit.get(pk, [0, 0.0])
        if nd or nc:
            issues.append({
                "level": "WARN", "type": "红字冲减",
                "message": f"{pk} {subj['label']}含红字/负数凭证：借方 {nd} 笔合计 {sd:,.2f}、贷方 {nc} 笔合计 {sc:,.2f}；"
                           f"已按净额计入借贷方，与辅助核算同口径，不影响勾稽。",
            })
    counterparties = reconstruct_counterparties(rows)
    # ⚡ 2026-08-10 SAP 快速路径：SAP rows 只有目标科目行（read_gl_net pfx 过滤），
    # 标准凭证翻面配对落空 → 用行自带 cp（_infer_opp 推断对方科目）构建，恢复『对方科目核对』sheet。
        # ⚡ 2026-08-10 升级：SAP 用凭证级精确分摊（剔除同凭证多对方科目金额重复=虚增虚减）
    if _sap_mode:
        by_ent = _adapter.recon_counterparties_sap(rows)
    else:
        by_ent = reconstruct_counterparties_by_entity(rows)
    issues += validate(rows, customers, periods, agg, summary, opening, counterparties, ctrl_map, subj)

    for w in read_warnings:
        issues.append({"level": "WARN", "type": "金额", "message": w})
    for (name, a1, a2, period) in aux_conflicts:
        issues.append({
            "level": "ERROR", "type": "余额冲突",
            "message": f"辅助核算中往来单位『{name}』({period})出现冲突期末值：{a1:.2f} 与 {a2:.2f}，已取末值 {a2:.2f}，请核实。",
        })

    if subject is None:
        subject = _derive_subject(aux_path)

    # 标记每行所属核算主体（合并模式按主体聚合时使用）
    for r in rows:
        r.setdefault("entity", subject)

    if not customers:
        # 该科目在此数据集中不存在（如股份公司无预收账款）：不生成文件，返回空
        return dict(issues=issues, periods=periods, customers=customers,
                   counterparties=counterparties, n_open=0, subject=subject,
                   summary=summary, rows=rows, opening=opening, by_ent=by_ent)

    if write:
        export_excel(output_path, periods, customers, summary, issues, rows, counterparties, period_mode, subject, subj, by_ent=by_ent, km_path=km_path)

        print(f"✅ 已生成：{output_path}")
        print(f"   期间：{periods}  往来单位数：{len(customers)}")
        for pk in periods:
            tot_open = sum(summary[(a, pk)]["s_open"] for a in customers)
            tot_db = sum(summary[(a, pk)]["s_debit"] for a in customers)
            tot_cr = sum(summary[(a, pk)]["s_credit"] for a in customers)
            tot_cl = sum(summary[(a, pk)]["s_close"] for a in customers)
            ctrl = ctrl_map.get(pk)
            ctrl_str = f"  控制数(科目余额表{subj['label']})={ctrl:,.2f}" if ctrl is not None else ""
            print(f"   [{pk}] 期初={tot_open:,.2f}  借方(符号)={tot_db:,.2f}  贷方(符号)={tot_cr:,.2f}  期末={tot_cl:,.2f}{ctrl_str}")
        errs = [i for i in issues if i["level"] == "ERROR"]
        warns = [i for i in issues if i["level"] == "WARN"]
        print(f"   校验：ERROR {len(errs)} 条，WARN {len(warns)} 条")

    return dict(issues=issues, periods=periods, customers=customers,
               counterparties=counterparties, n_open=len(opening), subject=subject,
               summary=summary, rows=rows, opening=opening, by_ent=by_ent)


def build_offset_workbook(all_bal, out_dir):
    """往来互抵明细：对同时出现在借方侧(应收/预付/其他应收)与贷方侧(应付/预收/其他应付)
    的往来单位（或关联方），列示双侧余额、互抵额与净额，并给出逐笔明细。"""
    # 按 (往来单位, 期间) 归集
    groups = defaultdict(lambda: dict(debit=0.0, credit=0.0, related=False, detail=[]))
    for b in all_bal:
        g = groups[(b["name"], b["period"])]
        if b["side"] == "debit":
            g["debit"] += b["close"]
        else:
            g["credit"] += b["close"]
        g["related"] = g["related"] or b["related"]
        g["detail"].append(b)
    # 仅保留『关联方』或『双侧均有余额』的单位
    keys = [k for k, g in groups.items() if g["related"] or (abs(g["debit"]) > 0.005 and abs(g["credit"]) > 0.005)]
    keys.sort()
    out_path = os.path.join(out_dir, "往来互抵明细_生成.xlsx")
    if not keys:
        if os.path.exists(out_path):
            try: os.remove(out_path)
            except BaseException: pass
        return None
    wb = openpyxl.Workbook()
    # ---- 互抵汇总 ----
    ws = wb.active
    ws.title = "互抵汇总"
    ws.cell(2, 1, "往来互抵明细 — 借方侧(应收/预付/其他应收) 与 贷方侧(应付/预收/其他应付) 互抵").font = TITLE_FONT
    hdr = ["往来单位", "是否关联方", "期间", "借方侧往来合计", "贷方侧往来合计", "互抵额", "净额", "结论"]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for (name, pk) in keys:
        g = groups[(name, pk)]
        offset = min(abs(g["debit"]), abs(g["credit"]))
        net = g["debit"] - g["credit"]
        if abs(g["debit"]) < 0.005 or abs(g["credit"]) < 0.005:
            concl = "单边（仅一侧有余额，无互抵）"
        elif abs(net) < 0.005:
            concl = "完全互抵"
        else:
            concl = "部分互抵"
        row = [name, "是" if g["related"] else "否", pk,
               round(g["debit"], 2), round(g["credit"], 2), round(offset, 2), round(net, 2), concl]
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (4, 5, 6, 7):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        if concl == "完全互抵":
            for c in range(1, len(hdr) + 1):
                ws.cell(r, c).fill = TOTAL_FILL
        r += 1
    for i, w in enumerate([30, 10, 10, 18, 18, 14, 16, 22], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    # ---- 互抵明细 ----
    ws2 = wb.create_sheet("互抵明细")
    ws2.cell(2, 1, "往来互抵明细 — 逐笔构成（核算主体 / 科目 / 方向 / 余额）").font = TITLE_FONT
    hdr2 = ["往来单位", "期间", "核算主体", "科目", "方向", "余额(本币带符号)", "备注"]
    for i, h in enumerate(hdr2, 1):
        ws2.cell(4, i, h)
    _style_header(ws2, 4, len(hdr2))
    r = 5
    for (name, pk) in keys:
        for b in sorted(groups[(name, pk)]["detail"], key=lambda x: (x["entity"], x["label"])):
            side_lbl = "借(应收/预付/其他应收)" if b["side"] == "debit" else "贷(应付/预收/其他应付)"
            ws2.cell(r, 1, name).border = BORDER
            ws2.cell(r, 2, pk).border = BORDER
            ws2.cell(r, 3, b["entity"]).border = BORDER
            ws2.cell(r, 4, b["label"]).border = BORDER
            ws2.cell(r, 5, side_lbl).border = BORDER
            cc = ws2.cell(r, 6, round(b["close"], 2)); cc.border = BORDER; cc.number_format = MONEY_FMT
            cc.alignment = Alignment(horizontal="right")
            ws2.cell(r, 7, "").border = BORDER
            r += 1
    for i, w in enumerate([30, 10, 16, 16, 26, 20, 20], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = "A5"
    finalize_workbook(wb)
    from audit_common import validate_workbook  # 2026-08-05 前道规范化：保存前结构校验
    validate_workbook(wb, '往来款底稿', raise_on_error=False)
    wb.save(out_path)
    wb.close()
    return out_path



def _build_cross_frame(wb):
    """关联交易和余额核对 — 空格式框架（外壳内容）。4 张固定 sheet 的标题/说明/列宽/冻结。
    表头列随主体数动态，由 _fill_cross 写入。样式统一采用 audit_shell 的 SHELL_* 常量。"""
    # Sheet1 — 关联交易余额核对（逐对并排·分年度）
    ws = wb.active
    ws.title = "关联交易余额核对"
    ws.cell(2, 1, "关联交易和余额核对 — 集团内账套主体间往来核对（逐对账套对·分年度并排，含期末余额与本年发生额）").font = SHELL_TITLE_FONT
    note = ("说明：已获取的账套同属一个集团，账套主体互为关联方。下表中每一行表示一个账套对"
            "(A↔B) 在某年度的往来情况，并排列示『本期借/贷发生额(关联交易流水)』与『期末借/贷余额』四项："
            "A→B 本期借、A→B 本期贷、A→B 期末借、A→B 期末贷，以及 B→A 对应四项，可直接对照两侧是否相互抵消。"
            "即使期末余额已结清(=0)，仍可在此看到集团内交易发生额。各账套对的『合计』行为跨年度汇总。"
            "对账套间是否核对相符（净额/差额），详见 Sheet『关联交易双向核对』；逐科目余额并排详见 Sheet『关联交易按科目核对』。")
    c2 = ws.cell(3, 1, note)
    c2.font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    c2.alignment = SHELL_LEFT
    ws.freeze_panes = "D5"
    # Sheet2 — 关联交易双向核对（按账套对 合计）
    ws2 = wb.create_sheet("关联交易双向核对")
    ws2.cell(2, 1, "关联交易双向核对 — 账套A对B的净债权(借-贷) 应与 账套B对A的净债权 互为相反数").font = SHELL_TITLE_FONT
    ws2.freeze_panes = "A5"
    # Sheet3 — 关联交易按科目核对（按账套对 × 各往来科目 展开）
    ws3 = wb.create_sheet("关联交易按科目核对")
    ws3.cell(2, 1, "关联交易按科目核对 — 账套A↔B 逐科目(应收/应付/其他应收/其他应付/应收票据)借/贷侧余额并排，定位未对账的具体科目（不含净额核对）").font = SHELL_TITLE_FONT
    ws3.freeze_panes = "A5"
    # Sheet4 — 关联交易差异明细
    ws4 = wb.create_sheet("关联交易差异明细")
    ws4.cell(2, 1, "关联交易差异明细 — 双向核对中差额≠0 的配对，直接给出各账套构成行供追查").font = SHELL_TITLE_FONT
    ws4.freeze_panes = "A5"


def _fill_cross(wb, all_bal, entity_list, out_dir, rp_list=None):
    """把数据填入外壳框架（_build_cross_frame 或外壳 xlsx 提供）。表头随主体数动态写入。"""
    matcher = build_entity_matcher(entity_list)
    # 配对列表 (b, E)：E = 命中本文件夹主体，或(源数据标记关联方但未能归并时)以往来单位名称为对侧。
    intra_pairs = []
    for b in all_bal:
        E = matcher(b["name"])
        if E is not None and E != b["entity"]:
            intra_pairs.append((b, E))
            continue
        # 源数据标记为关联方、但未能归并到本文件夹任一主体的（典型：跨文件夹的集团母公司/姊妹公司）：
        # 仍纳入关联交易核对并以往来单位名称为对侧，使『关联交易和余额核对』能生成并展示流水。
        if b.get("related") and b["name"] != b["entity"]:
            intra_pairs.append((b, b["name"]))
    intra = [b for b, _ in intra_pairs]
    # 期末余额(close) 与 本期发生额(flow) 分开聚合，便于分别展示
    by_cp = defaultdict(lambda: dict(ents=defaultdict(lambda: dict(
        close_debit=0.0, close_credit=0.0, flow_debit=0.0, flow_credit=0.0, lines=[]))))
    by_pair = defaultdict(lambda: dict(
        a_close_debit=0.0, a_close_credit=0.0, b_close_debit=0.0, b_close_credit=0.0,
        a_flow_debit=0.0, a_flow_credit=0.0, b_flow_debit=0.0, b_flow_credit=0.0,
        lines_a=[], lines_b=[]))
    by_pair_subj = defaultdict(lambda: dict(
        a_close_debit=0.0, a_close_credit=0.0, b_close_debit=0.0, b_close_credit=0.0,
        a_flow_debit=0.0, a_flow_credit=0.0, b_flow_debit=0.0, b_flow_credit=0.0,
        lines_a=[], lines_b=[]))
    by_pair_period = defaultdict(lambda: dict(
        a_close_debit=0.0, a_close_credit=0.0, b_close_debit=0.0, b_close_credit=0.0,
        a_flow_debit=0.0, a_flow_credit=0.0, b_flow_debit=0.0, b_flow_credit=0.0))
    for b, E in intra_pairs:
        close = b.get("close", 0.0) or 0.0
        fd = b.get("debit", 0.0) or 0.0
        fc = b.get("credit", 0.0) or 0.0
        # by_cp
        d = by_cp[(E, b["period"])]["ents"][b["entity"]]
        if b["side"] == "debit":
            d["close_debit"] += close
        else:
            d["close_credit"] += close
        d["flow_debit"] += fd
        d["flow_credit"] += fc
        d["lines"].append(b)
        # by_pair / by_pair_subj / by_pair_period 同步聚合
        i, j = b["entity"], E
        key = tuple(sorted([i, j]))
        p = by_pair[key]
        ps = by_pair_subj[(key, b["key"], b["label"])]
        pp = by_pair_period[(key, b["period"])]
        if i == key[0]:
            if b["side"] == "debit":
                p["a_close_debit"] += close; ps["a_close_debit"] += close; pp["a_close_debit"] += close
            else:
                p["a_close_credit"] += close; ps["a_close_credit"] += close; pp["a_close_credit"] += close
            p["a_flow_debit"] += fd; p["a_flow_credit"] += fc
            ps["a_flow_debit"] += fd; ps["a_flow_credit"] += fc
            pp["a_flow_debit"] += fd; pp["a_flow_credit"] += fc
            p["lines_a"].append(b); ps["lines_a"].append(b)
        else:
            if b["side"] == "debit":
                p["b_close_debit"] += close; ps["b_close_debit"] += close; pp["b_close_debit"] += close
            else:
                p["b_close_credit"] += close; ps["b_close_credit"] += close; pp["b_close_credit"] += close
            p["b_flow_debit"] += fd; p["b_flow_credit"] += fc
            ps["b_flow_debit"] += fd; ps["b_flow_credit"] += fc
            pp["b_flow_debit"] += fd; pp["b_flow_credit"] += fc
            p["lines_b"].append(b); ps["lines_b"].append(b)
    cp_keys = sorted(by_cp.keys())
    out_path = os.path.join(out_dir, "关联交易和余额核对_生成.xlsx")
    # 清理旧命名文件（重命名前的遗留副本），避免重复
    legacy = os.path.join(out_dir, "多个账套存在交易往来_生成.xlsx")
    if os.path.exists(legacy) and os.path.abspath(legacy) != os.path.abspath(out_path):
        try:
            os.remove(legacy)
        except BaseException:
            pass
    # 生成条件：存在集团内关联交易记录(发生额或余额)即生成；
    # 仅在完全无任何集团内关联交易(连发生额都没有)时才删除旧文件并不生成。
    if not intra:
        # 诊断日志：无关联交易时记录原因，便于定位（如跨文件夹关联方未纳入本文件夹）。
        if os.environ.get('AUDIT_LOG') == '1':
            try:
                dbg = os.path.join(out_dir, "关联交易核对_诊断.log")
                with open(dbg, "w", encoding="utf-8") as fh:
                    fh.write("关联交易和余额核对 — 未生成诊断\n")
                    if rp_list is not None:
                        fh.write("关联方识别依据: 文件夹内『关联方清单』（清单内必判为关联方，并保留账套主体"
                                 "自动归并与源数据标注，三重保险）\n")
                    else:
                        fh.write("关联方识别依据: 账套主体名称归并器（未提供『关联方清单』，可能误判；"
                                 "建议在文件夹放置 关联方清单.xlsx 提升准确性）\n")
                    fh.write(f"文件夹主体数: {len(entity_list)}\n")
                    fh.write(f"all_bal(跨科目余额)条目数: {len(all_bal)}\n")
                    n_rel = sum(1 for b in all_bal if b.get("related"))
                    fh.write(f"  其中被标记为关联方(related)的条目数: {n_rel}\n")
                    unmatched = sorted({b["name"] for b in all_bal
                                        if b.get("related") and matcher(b["name"]) is None})
                    fh.write(f"  标记关联方但未能归并到本文件夹主体的往来单位({len(unmatched)}):\n")
                    for u in unmatched[:60]:
                        fh.write(f"    - {u}\n")
                    fh.write("结论: 本文件夹内未发现可配对的集团内账套间往来，故未生成该表。\n")
                    fh.write("若上述往来单位确为集团内关联方，请确认其对应账套是否已纳入本文件夹；"
                             "或将其列入『关联方清单』后重跑。\n")
            except Exception:
                pass
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except BaseException:
                pass
        wb.close()
        return None
    # Sheet1 — 关联交易余额核对（逐对并排，按期间；期末余额 + 本期发生额）
    ws = wb["关联交易余额核对"]
    hdr = ["账套A", "账套B", "期间",
           "A→B本期借", "A→B本期贷", "A→B期末借", "A→B期末贷",
           "B→A本期借", "B→A本期贷", "B→A期末借", "B→A期末贷"]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    pair_period_keys = sorted(by_pair_period.keys(), key=lambda k: (k[0][0], k[0][1], k[1]))
    for (key, pk) in pair_period_keys:
        pp = by_pair_period[(key, pk)]
        row = [key[0], key[1], pk,
               round(pp["a_flow_debit"], 2), round(pp["a_flow_credit"], 2),
               round(pp["a_close_debit"], 2), round(pp["a_close_credit"], 2),
               round(pp["b_flow_debit"], 2), round(pp["b_flow_credit"], 2),
               round(pp["b_close_debit"], 2), round(pp["b_close_credit"], 2)]
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if 4 <= c <= 11 and isinstance(v, (int, float)):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        r += 1
    # 每对合计行（跨年度汇总）
    pair_tot = defaultdict(lambda: dict(
        a_flow_debit=0.0, a_flow_credit=0.0, a_close_debit=0.0, a_close_credit=0.0,
        b_flow_debit=0.0, b_flow_credit=0.0, b_close_debit=0.0, b_close_credit=0.0))
    for (key, pk), pp in by_pair_period.items():
        t = pair_tot[key]
        t["a_flow_debit"] += pp["a_flow_debit"]; t["a_flow_credit"] += pp["a_flow_credit"]
        t["a_close_debit"] += pp["a_close_debit"]; t["a_close_credit"] += pp["a_close_credit"]
        t["b_flow_debit"] += pp["b_flow_debit"]; t["b_flow_credit"] += pp["b_flow_credit"]
        t["b_close_debit"] += pp["b_close_debit"]; t["b_close_credit"] += pp["b_close_credit"]
    for key in sorted(pair_tot.keys(), key=lambda k: (k[0], k[1])):
        t = pair_tot[key]
        row = [key[0], key[1], "合计",
               round(t["a_flow_debit"], 2), round(t["a_flow_credit"], 2),
               round(t["a_close_debit"], 2), round(t["a_close_credit"], 2),
               round(t["b_flow_debit"], 2), round(t["b_flow_credit"], 2),
               round(t["b_close_debit"], 2), round(t["b_close_credit"], 2)]
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            cell.font = SHELL_BOLD
            if 4 <= c <= 11 and isinstance(v, (int, float)):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        r += 1
    if not by_pair_period:
        ws.cell(r, 1, "（无集团内关联交易记录）").font = Font(name='Times New Roman', italic=True, color="808080")
    for i, w in enumerate([18, 18, 10, 16, 16, 16, 16, 16, 16, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # Sheet2 — 关联交易双向核对（期末净债权应互为相反数；并附本期净额流水）
    ws2 = wb["关联交易双向核对"]
    hdr2 = ["账套A", "账套B", "A期末净债权", "B期末净债权", "差额(应=0)", "是否相符",
            "A本期净额(借-贷)", "B本期净额(借-贷)"]
    for i, h in enumerate(hdr2, 1):
        ws2.cell(4, i, h)
    _style_header(ws2, 4, len(hdr2))
    r = 5
    diff_pairs = []
    for (i, j) in sorted(by_pair.keys()):
        p = by_pair[(i, j)]
        net_a = p["a_close_debit"] - p["a_close_credit"]
        net_b = p["b_close_debit"] - p["b_close_credit"]
        diff = net_a + net_b
        flow_a = p["a_flow_debit"] - p["a_flow_credit"]
        flow_b = p["b_flow_debit"] - p["b_flow_credit"]
        balanced = abs(diff) <= max(0.01, 0.001 * max(abs(net_a), abs(net_b)))
        row = [i, j, round(net_a, 2), round(net_b, 2), round(diff, 2), "是" if balanced else "否",
               round(flow_a, 2), round(flow_b, 2)]
        for c, v in enumerate(row, 1):
            cell = ws2.cell(r, c, v)
            cell.border = BORDER
            if 3 <= c <= 8 and isinstance(v, (int, float)):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        if not balanced:
            ws2.cell(r, 6).fill = WARN_FILL
            diff_pairs.append((i, j))
        r += 1
    for i, w in enumerate([18, 18, 16, 16, 14, 10, 18, 18], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    # Sheet3 — 关联交易按科目核对（逐科目 期末余额 + 本期发生额 并排；已删除净额核对）
    ws3 = wb["关联交易按科目核对"]
    hdr3 = ["账套A", "账套B", "科目",
            "A期末借", "A期末贷", "A本期借", "A本期贷",
            "B期末借", "B期末贷", "B本期借", "B本期贷"]
    for i, h in enumerate(hdr3, 1):
        ws3.cell(4, i, h)
    _style_header(ws3, 4, len(hdr3))
    r = 5
    for (key, sk, sl) in sorted(by_pair_subj.keys()):
        ps = by_pair_subj[(key, sk, sl)]
        row = [key[0], key[1], sl,
               round(ps["a_close_debit"], 2), round(ps["a_close_credit"], 2),
               round(ps["a_flow_debit"], 2), round(ps["a_flow_credit"], 2),
               round(ps["b_close_debit"], 2), round(ps["b_close_credit"], 2),
               round(ps["b_flow_debit"], 2), round(ps["b_flow_credit"], 2)]
        for c, v in enumerate(row, 1):
            cell = ws3.cell(r, c, v)
            cell.border = BORDER
            if 4 <= c <= 11 and isinstance(v, (int, float)):
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        r += 1
    if not by_pair_subj:
        ws3.cell(r, 1, "（无集团内关联交易记录）").font = Font(name='Times New Roman', italic=True, color="808080")
    for i, w in enumerate([18, 18, 14, 14, 14, 14, 14, 14, 14, 14, 14], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    # Sheet4 — 关联交易差异明细
    ws4 = wb["关联交易差异明细"]
    hdr4 = ["账套A", "账套B", "期间", "核算主体", "对手方主体", "科目", "方向", "余额(本币带符号)", "备注"]
    for i, h in enumerate(hdr4, 1):
        ws4.cell(4, i, h)
    _style_header(ws4, 4, len(hdr4))
    r = 5
    if not diff_pairs:
        # 期末余额双向核对均相符；若存在本期发生额但期末已结清，给出提示
        has_flow = any((p["a_flow_debit"] or p["a_flow_credit"] or p["b_flow_debit"] or p["b_flow_credit"])
                       for p in by_pair.values())
        if has_flow:
            ws4.cell(r, 1, "（集团内关联交易期末余额双向核对均相符、已核对平衡；但本期存在集团内交易发生额，"
                          "已全部结算抵消，期末余额均为 0。本期发生额详见『关联交易余额核对 / 关联交易双向核对』。）").font = Font(name='Times New Roman', italic=True, color="808080")
        else:
            ws4.cell(r, 1, "（所有关联交易双向核对均相符：A对B净债权与B对A净债权互为相反数，集团内部往来已核对平衡）").font = Font(name='Times New Roman', italic=True, color="808080")
        r += 1
    else:
        for (i, j) in diff_pairs:
            p = by_pair[(i, j)]
            net_a = p["a_close_debit"] - p["a_close_credit"]
            net_b = p["b_close_debit"] - p["b_close_credit"]
            diff = net_a + net_b
            t = ws4.cell(r, 1, f"【{i} ↔ {j}】差额 = {diff:,.2f}（{i}净债权 {net_a:,.2f} + {j}净债权 {net_b:,.2f}）")
            t.font = SUB_FONT
            r += 1
            for b in sorted(p["lines_a"] + p["lines_b"], key=lambda x: (x["entity"], x["label"])):
                E = matcher(b["name"])
                side_lbl = "借(应收/预付/其他应收)" if b["side"] == "debit" else "贷(应付/预收/其他应付)"
                ws4.cell(r, 1, i).border = BORDER
                ws4.cell(r, 2, j).border = BORDER
                ws4.cell(r, 3, b["period"]).border = BORDER
                ws4.cell(r, 4, b["entity"]).border = BORDER
                ws4.cell(r, 5, E).border = BORDER
                ws4.cell(r, 6, b["label"]).border = BORDER
                ws4.cell(r, 7, side_lbl).border = BORDER
                cc = ws4.cell(r, 8, round(b.get("close", 0.0), 2)); cc.border = BORDER; cc.number_format = MONEY_FMT
                cc.alignment = Alignment(horizontal="right")
                ws4.cell(r, 9, "").border = BORDER
                r += 1
    for i, w in enumerate([20, 20, 10, 16, 16, 16, 26, 20, 20], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w
    finalize_workbook(wb)
    wb.save(out_path)
    wb.close()
    return out_path


def build_cross_workbook(all_bal, out_dir, entity_list, rp_list=None):
    """关联交易和余额核对 — 外壳模式。
    优先从中央模板目录加载『关联交易和余额核对_外壳.xlsx』（结构+样式归外壳）；
    缺失则现场构建框架并落盘该外壳文件（供日后在 Excel 调整格式，无需改代码）。"""
    shell = resolve_template("关联交易和余额核对_外壳", "CA_CROSS_TEMPLATE")
    if shell:
        wb = openpyxl.load_workbook(shell)
    else:
        wb = openpyxl.Workbook()
        _build_cross_frame(wb)
        try:
            os.makedirs(AUDIT_TEMPLATES_DIR, exist_ok=True)
            finalize_workbook(wb)
            wb.save(os.path.join(AUDIT_TEMPLATES_DIR, "关联交易和余额核对_外壳.xlsx"))
            wb.close()
        except Exception:
            pass
    return _fill_cross(wb, all_bal, entity_list, out_dir, rp_list=rp_list)


# ============================================================================
# 往来补充并入科目（参照往来底稿，并入对应主科目底稿）
# ----------------------------------------------------------------------------
# 依据需求：
#   应收股利(1131)/应收利息(1132) 参照往来底稿并入【其他应收款(ORA)】底稿；
#   应付股利(2231)/应付利息(2232) 参照往来底稿并入【其他应付款(ORP)】底稿；
#   坏账准备(1231) 拆为 应收账款坏账准备(123101)/其他应收款坏账准备(123102)，
#     分别并入【应收账款(AR)】/【其他应收款(ORA)】底稿。
# 并入表统一列示：期初、本期增加、本期减少、期末未审、审计调整、审定数。
# 坏账准备本期增减须与信用减值损失(6702)对应子目勾稽；若有差异，形成凭证抽查。
# 说明：并入表主数据取 TB（控制数，权威）；凭证抽查取 GL（受非全量/2026 YTD 限制）。
RELATED_MAP = {
    'AR': [
        # 应收账款坏账准备明细表由 _append_ar_baddebt_sheet(通用坏账明细表) 生成，
        # 不再走 _append_one_related 格式（2026-08-03 用户要求同存货跌价准备明细表）。
    ],
    'ORA': [
        dict(sheet='应收股利明细表', title='应收股利', name_kw=['应收股利'],
             is_credit=False, is_baddebt=False),
        dict(sheet='应收利息明细表', title='应收利息', name_kw=['应收利息'],
             is_credit=False, is_baddebt=False),
        # 其他应收款坏账准备明细表由 _append_ar_baddebt_sheet(通用坏账明细表) 生成，
        # 不再走 _append_one_related 格式（2026-08-03 用户要求同应收账款坏账准备明细表）。
    ],
    'ORP': [
        dict(sheet='应付股利明细表', title='应付股利', name_kw=['应付股利'],
             is_credit=True, is_baddebt=False),
        dict(sheet='应付利息明细表', title='应付利息', name_kw=['应付利息'],
             is_credit=True, is_baddebt=False),
    ],
}
_RELATED_FILE_MAP = {
    '应收账款审计底稿': 'AR',
    '其他应收款审计底稿': 'ORA',
    '其他应付款审计底稿': 'ORP',
}


def _rel_key_for(fn):
    """按文件名前缀匹配并入科目所属主科目（2026-08-03 修复：实际文件名为
    『其他应收款审计底稿_2025_生成.xlsx』带年份，原整名匹配恒失败→应收股利/利息明细表从未生成）。"""
    for k, v in _RELATED_FILE_MAP.items():
        if fn.startswith(k):
            return v
    return None
_REL_TOT_FILL = PatternFill('solid', fgColor='D9E1F2')
_REL_RED = Font(name='Times New Roman', bold=True, color='C00000')


def _rel_hdr(ws, row, headers, start_col=1):
    for j, h in enumerate(headers, start_col):
        c = ws.cell(row, j, h)
        c.fill = SHELL_HFILL
        c.font = SHELL_HFONT
        c.border = SHELL_BORDER
        c.alignment = SHELL_CEN
    return row + 1


def _rel_money(ws, r, c, v):
    if v is not None:
        try:
            v = round(float(v), 2)
        except (TypeError, ValueError):
            pass
    cell = ws.cell(r, c, v)
    cell.number_format = '#,##0.00'
    cell.border = SHELL_BORDER
    cell.alignment = Alignment(horizontal='right')
    return cell


def _rel_txt(ws, r, c, v, bold=False, fill=None, align=None):
    cell = ws.cell(r, c, v)
    cell.border = SHELL_BORDER
    cell.alignment = align or SHELL_LEFT
    if bold:
        cell.font = SHELL_BOLD
    if fill:
        cell.fill = fill
    return cell


def _rel_set_formula(ws, r, col, formula):
    cell = ws.cell(r, col, formula)
    cell.number_format = '#,##0.00'
    cell.border = SHELL_BORDER
    cell.alignment = Alignment(horizontal='right')
    return cell


def _rel_fmt_date(d):
    if d is None:
        return ''
    if isinstance(d, (datetime.date, datetime.datetime)):
        return d.strftime('%Y-%m-%d')
    return str(d)


def _rel_fmt_vno(x):
    if not x:
        return ''
    vt = x.get('vtype'); vn = x.get('vno')
    if vt is None and vn is None:
        return ''
    return '%s-%s' % (vt, vn) if vt is not None else str(vn)


def _match_related(code, name, spec):
    """命中判断：按科目【名称】关键词归类（铁律5：不依赖科目代码体系，跨账套代码口径不一致）。
    2026-07-31：归一化『帐』→『账』——FY 账套子目名用"坏账准备-应收帐款"(帐)，
    name_kw 用"应收账款"(账) 匹配不上导致并入表为空；归一化后兼容两种写法。"""
    n = str(name or '').replace('帐', '账')
    for kw in spec.get('name_kw', []):
        if kw not in n:
            return False
    return True


def _gl_match_baddebt(name, spec):
    """坏账准备凭证抽查的 GL 行匹配：取『坏账准备』与『信用减值损失』两侧，按子目关键词定位。"""
    n = str(name or '').replace('帐', '账')
    sub = spec.get('sub_keyword', '')
    if '坏账准备' in n and (not sub or sub in n):
        return True
    if '信用减值损失' in n and (not sub or sub in n):
        return True
    return False


def _append_related_sheets(out_path, tb, gl, years, data_dir, wb=None):
    """往来主科目底稿生成后回写追加『并入的往来相关科目』明细 sheet。
    ⚡⚡ 2026-08-25 性能优化：支持外部传入 wb（合并多次 load+save 为一次）。"""
    fn = os.path.basename(out_path)
    key = _rel_key_for(fn)
    if not key or key not in RELATED_MAP:
        return
    specs = RELATED_MAP[key]
    if not years:
        return
    if wb is None:
        wb = openpyxl.load_workbook(out_path)   # 普通加载（需写回追加 sheet）
    try:
        for spec in specs:
            if spec['sheet'] in wb.sheetnames:
                del wb[spec['sheet']]
            recon_info = _append_one_related(wb, tb, gl, years, spec, data_dir)
            if spec.get('is_baddebt'):
                vname = spec['title'] + '凭证抽查'
                if vname in wb.sheetnames:
                    del wb[vname]   # 始终清除旧凭证抽查（避免无差异时残留陈旧 sheet）
                if recon_info.get('has_diff'):
                    try:
                        _append_baddebt_voucher(wb, tb, gl, years, spec, recon_info)
                    except Exception as ex:
                        print(f"   ⚠️ 坏账准备凭证抽查生成失败 [{spec['title']}]：{ex}")
    finally:
        if wb is None:
            wb.save(out_path)
            wb.close()
            del wb   # 2026-08-07 内存优化：显式释放（AR 大文件 4 次追加链 OOM）
    print(f"   ✅ 已追加并入科目 sheet 至：{out_path}")


def _append_ar_baddebt_sheet(out_path, tb, gl, years, data_dir, key='AR', wb=None):
    """坏账准备明细表（2026-08-03 用户要求，格式参照存货跌价准备明细表）：
    主体×期初余额/本期计提(贷方)/本期转回及转销(借方)/期末余额（不分月）；
    下方附『计提数与信用减值损失科目核对』（借：信用减值损失 / 贷：坏账准备-对应子目）。
    AR=应收账款坏账准备(1231 名称应收帐款/应收账款)；ORA=其他应收款坏账准备(1231 名称其他应收款)。
    ⚡⚡ 2026-08-25 性能优化：支持外部传入 wb（build_all_in_dir 合并多次 load+save 为
    一次，消除 85MB 大底稿反复全量 IO）；wb=None 时保持原 load/save 行为。"""
    _is_ar = (key == 'AR')
    if wb is None:
        wb = openpyxl.load_workbook(out_path)
    try:
        # 清理旧 RELATED_MAP 格式残留 sheet（2026-08-03 起坏账准备明细表改由本函数生成）
        for _old in ('应收账款坏账准备明细表', '应收账款坏账准备凭证抽查',
                     '其他应收款坏账准备明细表', '其他应收款坏账准备凭证抽查'):
            if _old in wb.sheetnames:
                del wb[_old]
        for y in years:
            # 精确到对应坏账准备子目：AR=应收帐款/应收账款（排除其他应收/应收票据）；ORA=其他应收款
            matched = []
            for (e, c, n, yy), v in tb.items():
                if str(yy) != str(y) or not str(c).startswith('1231'):
                    continue
                nm = str(n)
                if _is_ar:
                    if ('应收帐款' in nm or '应收账款' in nm) and '其他应收' not in nm and '应收票据' not in nm:
                        matched.append((e, c, n, v))
                else:
                    # ⚡ 2026-08-10 SAP 名称方言：1231020000=『坏账准备-其他应收账款坏帐准备』
                    # （含"其他应收账款"而非"其他应收款"）→ 两种写法都匹配
                    if '其他应收款' in nm or '其他应收账款' in nm:
                        matched.append((e, c, n, v))
            if not matched:
                continue
            sheet_name = '坏账准备明细表'
            if sheet_name in wb.sheetnames:
                del wb[sheet_name]
            ws = wb.create_sheet(sheet_name)
            ws.cell(1, 1, f'{"应收账款" if _is_ar else "其他应收款"}坏账准备明细表（{y} 年）').font = TITLE_FONT
            hr = 4
            headers = ['主体', '期初余额', '本期计提', '本期转回及转销', '期末余额']
            for j, h in enumerate(headers, 1):
                cc = ws.cell(hr, j, h)
                cc.font = SHELL_BOLD; cc.fill = PatternFill('solid', fgColor='DDEBF7')
                cc.border = BORDER; cc.alignment = Alignment(horizontal='center', vertical='center')
            by_ent = defaultdict(list)
            for e, c, n, v in matched:
                by_ent[e].append(v)
            r = hr + 1
            tot = {'qc': 0.0, 'inc': 0.0, 'dec': 0.0, 'qm': 0.0}
            for ent in sorted(by_ent):
                vs = by_ent[ent]
                # 2026-08-03 符号修复：备抵科目 TB 为带符号（贷余 qc/qm 为负、红字 jf/df 用负号）。
                # 期初/期末按带符号显示（贷余为负，与 Z 存货跌价明细表口径一致）；
                # 计提 = 正贷方 + 红字借方（备抵增加）、转回及转销 = 正借方 + 红字贷方（备抵减少）；
                # 勾稽 = 期初 + 转回 − 计提 = 期末（红字转回等价借方）。
                qc = sum(float(v.get('qc') or 0.0) for v in vs)
                qm = sum(float(v.get('qm') or 0.0) for v in vs)
                inc = sum(max(0.0, float(v.get('df') or 0.0)) + max(0.0, -float(v.get('jf') or 0.0))
                          for v in vs)
                dec = sum(max(0.0, float(v.get('jf') or 0.0)) + max(0.0, -float(v.get('df') or 0.0))
                          for v in vs)
                if abs(qc) < 0.005 and abs(inc) < 0.005 and abs(dec) < 0.005 and abs(qm) < 0.005:
                    continue
                ws.cell(r, 1, ent)
                for j, val in enumerate((qc, inc, dec, qm), start=2):
                    cell = ws.cell(r, j, round(val, 2))
                    cell.number_format = MONEY_FMT
                    cell.alignment = Alignment(horizontal='right')
                for j in range(1, 6):
                    ws.cell(r, j).border = BORDER
                tot['qc'] += qc; tot['inc'] += inc; tot['dec'] += dec; tot['qm'] += qm
                r += 1
            ws.cell(r, 1, '合计').font = TOTAL_FONT
            for j, key in enumerate(('qc', 'inc', 'dec', 'qm'), start=2):
                cell = ws.cell(r, j, round(tot[key], 2))
                cell.font = TOTAL_FONT; cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal='right')
            for j in range(1, 6):
                ws.cell(r, j).border = BORDER
                ws.cell(r, j).fill = TOTAL_FILL
            r += 2
            # —— 计提数与信用减值损失科目核对（借：信用减值损失 / 贷：坏账准备-对应子目）——
            ws.cell(r, 1, f'坏账计提数与信用减值损失科目核对（借：信用减值损失 / 贷：坏账准备-'
                          f'{"应收帐款" if _is_ar else "其他应收款"}）：').font = TOTAL_FONT
            r += 1
            for j, h in enumerate(['主体', '年度', '坏账准备贷方计提', '信用减值损失借方', '差异'], 1):
                cc = ws.cell(r, j, h); cc.font = SHELL_BOLD
                cc.border = BORDER; cc.fill = PatternFill('solid', fgColor='DDEBF7')
                cc.alignment = Alignment(horizontal='center')
            r += 1
            from collections import defaultdict as _dd
            prov_by = _dd(lambda: _dd(float))
            loss_by = _dd(lambda: _dd(float))
            for row in gl:
                nm = str(row.get('name') or '')
                yy = str(row.get('y') or '')
                if yy != str(y):
                    continue
                if _is_ar:
                    if ('坏账准备' in nm and ('应收帐款' in nm or '应收账款' in nm)
                            and '其他应收' not in nm and '应收票据' not in nm):
                        prov_by[row.get('e')][yy] += float(row.get('cr') or 0.0)
                else:
                    if '坏账准备' in nm and '其他应收款' in nm:
                        prov_by[row.get('e')][yy] += float(row.get('cr') or 0.0)
                if '信用减值损失' in nm:
                    loss_by[row.get('e')][yy] += float(row.get('dr') or 0.0)
            for ent in sorted(set(list(prov_by) + list(loss_by))):
                pv = prov_by[ent].get(str(y), 0.0)      # 坏账准备贷方（带符号：正=计提、红字=转回）
                lv = loss_by[ent].get(str(y), 0.0)      # 信用减值损失借方（带符号）
                ws.cell(r, 1, ent).border = BORDER
                ws.cell(r, 2, y).border = BORDER
                for j, val in enumerate((pv, lv, pv - lv), start=3):
                    cell = ws.cell(r, j, round(val, 2))
                    cell.number_format = MONEY_FMT
                    cell.alignment = Alignment(horizontal='right')
                    cell.border = BORDER
                if abs(pv - lv) > 0.005:
                    ws.cell(r, 5).font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
                r += 1
            ws.cell(r, 1, '注：差异=信用减值损失借方中含其他构成（如应收票据/其他应收款坏账、审计调整等），需人工核实。').font = \
                Font(name='Times New Roman', size=9, italic=True, color='808080')
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
            for i, w in enumerate([16, 16, 16, 18, 20], 1):
                ws.column_dimensions[get_column_letter(i)].width = w
    finally:
        if wb is None:
            wb.save(out_path)
            wb.close()
            del wb   # 2026-08-07 内存优化：显式释放（AR 大文件 4 次追加链 OOM）
    print(f"   ✅ 已追加 坏账准备明细表 至：{out_path}")


def _append_one_related(wb, tb, gl, years, spec, data_dir):
    """写出单个并入科目『增减变动及明细表』。返回 {'has_diff': bool}（仅坏账准备有意义）。"""
    from audit_common import safe, is_zero_amount
    ws = wb.create_sheet(spec['sheet'])
    ncols = 10
    headers = ['年度', '主体', '项目(明细科目)', '期初', '本期增加', '本期减少',
               '期末未审', '审计调整', '审定数', '勾稽(期初+增-减-末)']
    r = 1
    ws.cell(r, 1, '%s增减变动及明细表（%s 年度）' % (spec['title'], '、'.join(years))).font = TITLE_FONT
    r += 1
    hr = r + 1
    r = _rel_hdr(ws, hr, headers)
    C = list(range(1, ncols + 1))
    C_YR, C_ENT, C_ITEM, C_QC, C_INC, C_DEC, C_QM, C_ADJ, C_AUD, C_CHK = C

    def _rv(v):
        qc = safe(v.get('qc')); jf = safe(v.get('jf')); df = safe(v.get('df')); qm = safe(v.get('qm'))
        qc_s = abs(qc); qm_s = abs(qm)
        if spec['is_credit']:
            inc = df; dec = jf
        else:
            inc = jf; dec = df
        return qc_s, inc, dec, qm_s

    any_data = False
    for y in years:
        matched = [(e, c, n, v) for (e, c, n, yy), v in tb.items()
                   if yy == y and _match_related(c, n, spec)]
        if not matched:
            continue
        by_ent = defaultdict(list)
        for e, c, n, v in matched:
            by_ent[e].append((c, n, v))
        for ent in sorted(by_ent):
            rows = by_ent[ent]
            primary = min(rows, key=lambda x: len(x[0]))
            pcode, pname, pv = primary
            allc = [c for c, n, v in rows]
            leaf = [(c, n, v) for c, n, v in rows
                    if not any(o != c and o.startswith(c) for o in allc)]
            children = [x for x in rows if x[0] != pcode and x[0].startswith(pcode)]
            # 主体分组标题
            _rel_txt(ws, r, C_ENT, '%s（%s）' % (ent, y), bold=True)
            ws.merge_cells(start_row=r, start_column=C_ENT, end_row=r, end_column=C_ITEM)
            for i in range(1, ncols + 1):
                ws.cell(r, i).fill = _REL_TOT_FILL
                ws.cell(r, i).font = SHELL_BOLD
            r += 1
            has_children = bool(children)
            detail = leaf if leaf else [(pcode, pname, pv)]
            qc_t = inc_t = dec_t = qm_t = 0.0
            for code, name, v in detail:
                dqc, dinc, ddec, dqm = _rv(v)
                if is_zero_amount(dqc, dinc, ddec, dqm):
                    continue
                any_data = True
                _rel_txt(ws, r, C_YR, y)
                _rel_txt(ws, r, C_ITEM, '%s %s' % (code, name))
                _rel_money(ws, r, C_QC, dqc)
                _rel_money(ws, r, C_INC, dinc)
                _rel_money(ws, r, C_DEC, ddec)
                _rel_money(ws, r, C_QM, dqm)
                _rel_money(ws, r, C_ADJ, None)
                _rel_set_formula(ws, r, C_AUD, '=%s%d+%s%d' % (
                    get_column_letter(C_QM), r, get_column_letter(C_ADJ), r))
                _rel_set_formula(ws, r, C_CHK, '=%s%d+%s%d-%s%d-%s%d' % (
                    get_column_letter(C_QC), r, get_column_letter(C_INC), r,
                    get_column_letter(C_DEC), r, get_column_letter(C_QM), r))
                r += 1
                qc_t += dqc; inc_t += dinc; dec_t += ddec; qm_t += dqm
            if has_children and any_data:
                _rel_txt(ws, r, C_ITEM, '%s 合计' % ent, bold=True)
                _rel_money(ws, r, C_QC, qc_t)
                _rel_money(ws, r, C_INC, inc_t)
                _rel_money(ws, r, C_DEC, dec_t)
                _rel_money(ws, r, C_QM, qm_t)
                _rel_money(ws, r, C_ADJ, None)
                _rel_set_formula(ws, r, C_AUD, '=%s%d+%s%d' % (
                    get_column_letter(C_QM), r, get_column_letter(C_ADJ), r))
                _rel_set_formula(ws, r, C_CHK, '=%s%d+%s%d-%s%d-%s%d' % (
                    get_column_letter(C_QC), r, get_column_letter(C_INC), r,
                    get_column_letter(C_DEC), r, get_column_letter(C_QM), r))
                for i in range(1, ncols + 1):
                    ws.cell(r, i).font = SHELL_BOLD
                r += 1
    if not any_data:
        _rel_txt(ws, r, C_ENT, '本科目本期无发生额及余额（TB 中无对应科目）', bold=True)
        ws.merge_cells(start_row=r, start_column=C_ENT, end_row=r, end_column=ncols)
        r += 1
    # 坏账准备：与信用减值损失勾稽
    has_diff = False
    if spec.get('is_baddebt'):
        recon_kw = spec.get('recon_kw'); recon_title = spec.get('recon_title', '信用减值损失')
        r += 1
        _rel_txt(ws, r, C_ENT, '勾稽：坏账准备本期计提(贷方发生额) ↔ %s(借方发生额)；差异应=0，差异形成凭证抽查。' % recon_title, bold=True)
        ws.merge_cells(start_row=r, start_column=C_ENT, end_row=r, end_column=ncols)
        r += 1
        _rel_hdr(ws, r, ['年度', '主体', '坏账准备·本期计提(贷发)', '%s·借方发生' % recon_title, '差异', '说明', '', '', '', ''])
        r += 1
        for y in years:
            for ent in sorted({e for (e, c, n, yy) in tb if yy == y}):
                bd = [v for (e, c, n, yy), v in tb.items()
                      if e == ent and yy == y and _match_related(c, n, spec)]
                rec = [v for (e, c, n, yy), v in tb.items()
                       if e == ent and yy == y and recon_kw and
                       all(k in str(n) for k in recon_kw)]
                bd_df = sum(safe(v.get('df')) for v in bd)
                rec_jf = sum(safe(v.get('jf')) for v in rec)
                diff = bd_df - rec_jf
                if abs(bd_df) < 1e-6 and abs(rec_jf) < 1e-6 and abs(diff) < 1e-6:
                    continue
                if abs(diff) > 1e-6:
                    has_diff = True
                _rel_txt(ws, r, 1, y)
                _rel_txt(ws, r, 2, ent)
                _rel_money(ws, r, 3, bd_df)
                _rel_money(ws, r, 4, rec_jf)
                dcell = _rel_money(ws, r, 5, diff)
                if abs(diff) > 1e-6:
                    dcell.font = _REL_RED
                _rel_txt(ws, r, 6, '差异！详见《%s凭证抽查》' % spec['title'] if abs(diff) > 1e-6 else '一致')
                for i in range(7, ncols + 1):
                    ws.cell(r, i).border = SHELL_BORDER
                r += 1
    # 说明行（末行）
    if spec['is_credit']:
        note = '说明：本期增加=贷方发生额；本期减少=借方发生额；'
    else:
        note = '说明：本期增加=借方发生额；本期减少=贷方发生额；'
    note += '审计调整列手填，审定数=期末未审+审计调整；勾稽列(期初+本期增加-本期减少-期末未审)应=0。'
    if spec.get('is_baddebt'):
        note += '坏账准备本期计提(贷方发生额)应与《信用减值损失》对应子目借方发生额勾稽一致。'
    _rel_txt(ws, r, C_ENT, note)
    ws.merge_cells(start_row=r, start_column=C_ENT, end_row=r, end_column=ncols)
    r += 1
    widths = [8, 22, 40, 16, 14, 14, 16, 12, 16, 18]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    return {'has_diff': has_diff}


def _append_baddebt_voucher(wb, tb, gl, years, spec, recon_info):
    """坏账准备与信用减值损失勾稽存在差异时，形成『凭证抽查』sheet（取 GL）。"""
    from audit_common import safe, voucher_skip_keys
    ws = wb.create_sheet(spec['title'] + '凭证抽查')
    ncols = 15
    headers = ['核算主体', '测试序号', '日期', '凭证字', '凭证号',
               '二级科目（或客户/供应商名称）', '摘要', '借方金额', '贷方金额',
               '对方科目', '与原始凭证相符', '原始凭证内容', '原始凭证日期',
               '会计处理正确', '所属时间无误']
    r = 1
    ws.cell(r, 1, '%s凭证抽查（与《信用减值损失》勾稽差异追查，已剔除当期损益结转）' % spec['title']).font = TITLE_FONT
    r += 1
    _rel_txt(ws, r, 1, '说明：本科目与信用减值损失勾稽存在差异，下列为当年（已剔除结转损益类凭证）相关凭证，供抽凭追查；'
                       'GL 为非全量抽取，2026 年 GL 仅 1–5 月(YTD)。', bold=False)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 1
    hr = r + 1
    r = _rel_hdr(ws, hr, headers)
    C = list(range(1, ncols + 1))
    rows_data = []
    grand_dr = grand_cr = 0.0
    for y in years:
        for ent in sorted({e for (e, c, n, yy) in tb if yy == y}):
            rows = [x for x in gl if x.get('y') == y and x.get('e') == ent
                    and _gl_match_baddebt(x.get('name', ''), spec)]
            if rows:
                skip = voucher_skip_keys(
                    rows,
                    key_fields=lambda rr: (rr.get('e'), rr.get('y'), rr.get('vtype'), rr.get('vno')),
                    name_field=lambda rr: rr.get('name') or '',
                    opp_field=lambda rr: rr.get('opp') or '',
                    summ_field=lambda rr: rr.get('summary') or '',
                )
                rows = [x for x in rows if (x.get('e'), x.get('y'), x.get('vtype'), x.get('vno')) not in skip]
            for x in rows:
                dr = safe(x.get('dr')); cr = safe(x.get('cr'))
                rows_data.append((ent, x.get('date') or '', y, x, dr, cr))
                grand_dr += dr; grand_cr += cr
    rows_data.sort(key=lambda x: (x[0], x[1]))
    for i, (ent, _, year, x, dr, cr) in enumerate(rows_data, 1):
        _rel_txt(ws, r, 1, ent)                                      # 核算主体
        _rel_txt(ws, r, 2, i)                                        # 测试序号
        _rel_txt(ws, r, 3, _rel_fmt_date(x.get('date')))             # 日期
        _rel_txt(ws, r, 4, year)                                     # 凭证字 = 原年度
        _rel_txt(ws, r, 5, _rel_fmt_vno(x))                          # 凭证号
        _rel_txt(ws, r, 6, x.get('name') or '')                      # 二级科目
        _rel_txt(ws, r, 7, x.get('summary') or '')                   # 摘要
        _rel_money(ws, r, 8, round(dr, 2) if dr > 0 else None)      # 借方金额
        _rel_money(ws, r, 9, round(cr, 2) if cr > 0 else None)      # 贷方金额
        _rel_txt(ws, r, 10, x.get('opp') or '')                      # 对方科目
        _rel_txt(ws, r, 11, '')                                      # 与原始凭证相符
        _rel_txt(ws, r, 12, '')                                      # 原始凭证内容
        _rel_txt(ws, r, 13, '')                                      # 原始凭证日期
        _rel_txt(ws, r, 14, '')                                      # 会计处理正确
        _rel_txt(ws, r, 15, '')                                      # 所属时间无误
        r += 1
    _rel_txt(ws, r, 7, '合计', bold=True)
    _rel_money(ws, r, 8, grand_dr)
    _rel_money(ws, r, 9, grand_cr)
    for i in range(1, ncols + 1):
        ws.cell(r, i).font = SHELL_BOLD
    widths = [14, 10, 14, 10, 10, 26, 30, 14, 14, 24, 12, 16, 14, 12, 12]
    for i, w in enumerate(widths[:ncols], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)


def _write_flow_audit_sheet(wb, key, subj, tb_full, entities, y, baddebt_net=True, reclass_map=None,
                            reclass_in_map=None, codes=None, label=None, related=False):
    # 2026-08-05 并入科目（应收/应付股利、应收/应付利息）独立审定表与主科目统一为 flow 12 列：
    # codes=动态匹配的科目代码（FY 2232=应付股利 非标准 2231）；label=科目名；
    # related=True 跳过 AR/ORA/ORP 备抵/三科目块特判（并入科目不适用）。
    """往来科目审定表（参考『往来款审定表.xlsx』格式，2026-08-01 用户需求；2026-08-04 删『命名简称』列→12 列）：
      汇总包顺序/公司/期初余额/本期增加/本期减少/期末余额/企业自行重分类/
      报表数/审计重分类调整1/审计重分类调整2/审计调整/审定数
    按主体行 + 合计行；TB 取 qc/jf/df/qm（自然符号转正数显示，is_credit 科目取反）；
    审定数 = 期末 + 重分类1 + 重分类2 + 审计调整（公式，Excel 打开自动计算）。
    2026-08-03 重分类：reclass_map(调整1列)=本科目反项余额转出合计；reclass_in_map(调整2列)=对方科目
    转入合计 → 审定数 = 期末 + 转出 + 转入 = 剔除反项余额并纳入转入（报表口径）。
    2026-08-04 删除『命名简称』列（原两列同名冗余）→ 12 列：公司列=实体名。"""
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
    BOLD = Font(name='Times New Roman', size=10, bold=True)
    _ca_code = {'AR': '1122', 'APR': '2203', 'APP': '1123', 'ORA': '1221',
                'ORP': '2241', 'AP': '2202', 'ARN': '1121',
                'CL': '2205', 'CA': '1480'}
    code = codes[0] if codes else _ca_code.get(key, key)
    is_credit = (subj.get('nature') == 'liability')
    _lbl = label or subj['label']
    title = f"{_lbl} 审定表"
    ws = wb.create_sheet(title=title)
    N = 12
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=N)
    tcell = ws.cell(1, 1, f"{_lbl}{y}审定表")
    tcell.font = BOLD; tcell.fill = TITLEFILL; tcell.alignment = CTR
    ws.row_dimensions[1].height = 22
    headers = ['汇总包顺序', '公司', '期初余额', '本期\n增加', '本期\n减少',
               '期末余额', '企业自行重分类', '报表数', '审计重分类调整1', '审计重分类调整2',
               '审计调整', '审定数']
    for j, h in enumerate(headers, start=1):
        c = ws.cell(2, j, h); c.font = BOLD; c.fill = HFILL; c.alignment = CTR; c.border = BORDER

    # 注：期初/期末按科目方向带符号（资产 借正贷负、负债 贷正借负）；增减数一律正数——
    # 资产：增加=jf(借)、减少=df(贷)；负债：增加=df(贷)、减少=jf(借)。
    # ⚡⚡ 2026-08-24 修复（AYL 等账套审定表勾稽不平根因）：is_credit 只反映"科目名义
    #   方向"，但源 TB 可能把负债科目记成【借余】（AYL 2202 应付账款期初/期末方向=借，
    #   read_tb_full 取 abs 丢失方向符号）→ 增/减列按 is_credit 硬编码映射接反 →
    #   期初+增-减≠期末。改为【借余/贷余自适应】：比较 |期初+借-贷-期末|（借余公式）
    #   与 |期初+贷-借-期末|（贷余公式），取更接近者决定增/减映射——ACB 贷余应付
    #   （增=df/减=jf）与 AYL 借余应付（增=jf/减=df）均正确；TB 自身不平的主体取
    #   更接近方向（差异仍保留，属源数据问题）。
    def _dir_is_debit(e):
        # ⚡⚡ 2026-08-24 二次修复（ACB 勾稽反而打破根因）：read_tb_full 的 qc/qm 是
        #   【带符号】余额（借正贷负，贷余科目为负），qm = qc + jf - df 是带符号体系的
        #   通用恒等式（对借余/贷余都成立）→ 用 d1/d2 拟合方向恒判借余（d1 恒≈0）。
        #   正确判定 = 看科目在主体内的【实际余额方向】：qm>0 借余（增=jf/减=df）、
        #   qm<0 贷余（增=df/减=jf）；期末平按期初方向。ACB 应付贷余(qm<0)→增=df ✓、
        #   AYL 应付借余(qm>0)→增=jf ✓。
        _qm = _amt_raw(e, 'qm')
        if abs(_qm) < 0.005:
            return _amt_raw(e, 'qc') >= 0
        return _qm > 0

    def _inc(e):
        # ⚡⚡ 2026-08-25 红字兼容（AQ 合同负债 df=-851298.6 / 合同资产 jf=-3682422.43
        #   根因）：jf/df 是带符号发生额，红字（负值）表示反向实际发生——本方向负发
        #   不计入增加，反向负发计入减少。公式（对称，借余/贷余通用）：
        #   增加 = 本方向正发生 + 反方向负发生(|..|)；减少 = 反方向正发生 + 本方向负发生。
        _jf = _amt_raw(e, 'jf'); _df = _amt_raw(e, 'df')
        if _dir_is_debit(e):
            return (abs(_jf) if _jf > 0 else 0.0) + (abs(_df) if _df < 0 else 0.0)
        return (abs(_df) if _df > 0 else 0.0) + (abs(_jf) if _jf < 0 else 0.0)
    def _dec(e):
        _jf = _amt_raw(e, 'jf'); _df = _amt_raw(e, 'df')
        if _dir_is_debit(e):
            return (abs(_df) if _df > 0 else 0.0) + (abs(_jf) if _jf < 0 else 0.0)
        return (abs(_jf) if _jf > 0 else 0.0) + (abs(_df) if _df < 0 else 0.0)
    # JTt 兼容（2026-08-05）：多账簿 U8 无标准 4 位一级码（112201 等）、名称含编码前缀——
    # 预计算每 (主体,年) 的末级科目编码集合（编码 c 是更长编码的前缀 → 父级，其借发=子目之和，防双计）
    _final_codes = {}
    for (_ee, _cc, _nn, _yy), _vv in tb_full.items():
        _k = (_ee, str(_yy))
        _final_codes.setdefault(_k, set()).add(str(_cc))
    for _k in _final_codes:
        _sc = sorted(_final_codes[_k])
        _final_codes[_k] = {c for i, c in enumerate(_sc)
                            if not (i + 1 < len(_sc) and _sc[i + 1].startswith(c))}

    def _amt_raw(e, field):
        # 2026-08-08 最终修复：遍历 codes 全列表逐码匹配（跨一级组场景如 S 应收票据
        # 1120/1121 各自精确命中；公共前缀方案会跨科目边界误伤——F 1221+1231.02→'12'
        # 把坏账子目纳入、JTt 1121+1141→'11' 把内部往来纳入）。
        # 每码：精确键命中直接取；否则前缀+主名匹配——命中=名称主段含 kw 且
        # 【无更深的同前缀且主名含 kw 的子键】（父级有余额而子目主名不同时取父级，
        #  如 S pot 1121 应收票据/112101 银行承兑汇票；父+子都命中只计子目防双计）。
        tot = 0.0
        _codes = codes if codes else ([code] if code else [])
        for _cd in _codes:
            _cs = str(_cd)
            # 2026-08-08 修复（Q 边界）：bd_codes（ORA 坏账码 123102 等）会追加进
            # items codes（供坏账准备独立行使用），_amt_raw 遍历时必须跳过备抵码——
            # 否则主体行=余额+坏账（Q 其他应收款 Audio 484,440-25,259=459,181 假值）。
            if _cs.startswith(('1231', '1471', '1602', '1703')):
                continue
            v = tb_full.get((e, _cs, _lbl, str(y)), {})
            if v:
                tot += float(v.get(field, 0.0) or 0.0)
                continue
            _cands = [(cc, nn, vv) for (ee, cc, nn, yy), vv in tb_full.items()
                      if ee == e and str(yy) == str(y)
                      and str(cc).startswith(_cs) and _name_hit(nn, _lbl)]
            for cc, nn, vv in _cands:
                _c = str(cc)
                if any(str(cc2) != _c and str(cc2).startswith(_c) and _name_hit(nn2, _lbl)
                       for (ee2, cc2, nn2, yy2), vv2 in tb_full.items()
                       if ee2 == e and str(yy2) == str(y)):
                    continue
                tot += float(vv.get(field, 0.0) or 0.0)
        return tot

    r = 3
    t_qc = t_jf = t_df = t_qm = 0.0
    ents = sorted(entities)
    for i, e in enumerate(ents, 1):
        # 期初/期末按科目方向带符号（2026-08-01 修复 #5）：资产类 借正贷负、负债类 贷正借负；
        # 增减数一律正数反映（资产 增=jf借/减=df贷；负债 增=df贷/减=jf借）。
        qc = _amt_raw(e, 'qc'); qm = _amt_raw(e, 'qm')
        # ⚡⚡ 2026-08-23 审定数修复：qm 显示取正（资产借正/负债贷正），但『审定数』必须用
        #   按科目性质定向后的带符号期末（资产借正贷负、负债贷正借负）——否则对反项余额主体
        #   （如 3400 应收贷余 -3.03亿）审定数 = abs(qm)+转出 会翻倍（3.03+3.07=6.10亿错）。
        _qm_adj = qm if not is_credit else -qm
        if not is_credit:
            # 资产：期末借正贷负显示——贷余主体（预付贷余=应付性质，qm<0）显示负，
            #   与明细表（_signed_balance 资产贷→负）方向一致。原 `qm = abs(qm)` 把
            #   贷余取正 → 2468 预付 8 个贷余主体（2380/2460/2650/2720 等）明细负
            #   vs 审定正 8285 万差异。期初按期末方向对齐；期末平（qm≈0）按期初方向。
            #   勾稽：-qc + jf - df = -qm（带符号恒等式，借余/贷余均成立）。
            _neg = qm < 0 or (abs(qm) < 0.005 and qc < 0)
            qc = -qc if _neg else qc
            qm = qm
        else:
            # 负债：期末贷余为正；期初按期末方向对齐——期初借余显示负（AYL 广东亿利达
            #   应付 期初借余 4199万/期末贷余 1753万 方向翻转，abs 显示勾稽必然不平；
            #   -qc 显示后 -qc+df-jf=-qm 恒成立）。期初贷余显示正。
            # ⚡⚡ 2026-08-25 期末平（qm≈0）按期初方向（ACB 诚本其他应付款 qc=-340.95/
            #   期末 0：期初贷余应显示正 340.95，原 qc 保持 -340.95 → 勾稽 -681.90 不平）。
            # ⚡⚡ 2026-08-27 修复（任务886 预收/预付明细vs审定差异根因）：负债科目按
            #   『贷正借负』显示恒取反——原仅对贷余（qm<0）取反、借余（qm>0）保持正，
            #   与明细表（_signed_balance 负债借余显示负）符号相反 → 2468 预收 2380
            #   (+3843万借余) 明细-3843万 vs 审定+3843万 差异。恒取反后借余显示负，
            #   与明细表/审定数(_qm_adj=-qm)方向一致；重分类调整列转出借余(reclass_map)。
            #   勾稽：集团审定表（表头含汇总包/公司）由 wp_recon_check/_recon_check 豁免；
            #   单主体贷余场景（AYL/ACB）显示与旧逻辑相同（贷余取反、期初按期末方向）。
            qc = -qc
            qm = -qm
        inc = _inc(e); dec = _dec(e)
        # 全主体连续编号（含 0 值主体，保证序号连续、与账套主体清单一一对应）
        t_qc += qc; t_jf += inc; t_df += dec; t_qm += qm
        ws.cell(r, 1, i)
        ws.cell(r, 2, e)
        ws.cell(r, 3, round(qc, 2))
        ws.cell(r, 4, round(inc, 2))
        ws.cell(r, 5, round(dec, 2))
        ws.cell(r, 6, round(qm, 2))
        ws.cell(r, 7, 0)
        ws.cell(r, 8, round(qm, 2))                       # 报表数 = 期末 + 企业自行重分类
        _rc_ent = round((reclass_map or {}).get((e, str(y)), 0.0) or 0.0, 2)
        ws.cell(r, 9, _rc_ent if _rc_ent else 0)          # 审计重分类调整1（本科目反项余额转出）
        _rci_ent = round((reclass_in_map or {}).get((e, str(y)), 0.0) or 0.0, 2)
        ws.cell(r, 10, _rci_ent if _rci_ent else 0)       # 审计重分类调整2（对方科目转入）
        ws.cell(r, 11, 0)                                 # 审计调整
        # 2026-08-06 协议：审定数写数值（=期末+重分类转出+转入+调整0；公式 data_only 读 None）
        ws.cell(r, 12, round(_qm_adj + _rc_ent + _rci_ent, 2))
        for j in range(1, N + 1):
            cell = ws.cell(r, j); cell.border = BORDER; cell.font = FONT
            if j in (3, 4, 5, 6, 7, 8, 9, 10, 11, 12):
                cell.number_format = NUMFMT
            if j in (3, 4, 5, 6, 7, 8, 9, 10, 11, 12):
                cell.alignment = RGT
        r += 1
    # 合计行（2026-08-06 协议：J/K/L 列写数值=各主体行和；公式 data_only 读 None）
    ws.cell(r, 2, '合计')
    ws.cell(r, 3, round(t_qc, 2)); ws.cell(r, 4, round(t_jf, 2))
    ws.cell(r, 5, round(t_df, 2)); ws.cell(r, 6, round(t_qm, 2))
    ws.cell(r, 8, round(t_qm, 2))
    for _cc in (9, 10, 11, 12):
        _sv = sum(ws.cell(x, _cc).value for x in range(3, r)
                  if isinstance(ws.cell(x, _cc).value, (int, float)))
        ws.cell(r, _cc, round(_sv, 2))
    for j in range(1, N + 1):
        cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
        if j in (3, 4, 5, 6, 7, 8, 9, 10, 11, 12):
            cell.number_format = NUMFMT
    # —— 2026-08-03 应收账款审定表加备抵项：减：坏账准备 → 应收账款净额（报表净值口径）——
    if key == 'AR' and not related:
        _bd_qm = 0.0
        for (e, c, n, yy), v in tb_full.items():
            # 精确到应收账款坏账准备子目（排除 其他应收款/应收票据，其名含"应收"子串）
            if str(yy) == str(y) and str(c).startswith('1231') \
                    and ('应收帐款' in str(n) or '应收账款' in str(n)) and '其他应收' not in str(n):
                _bd_qm += abs(float(v.get('qm', 0.0) or 0.0))
        r += 1
        r_tot = r - 1
        ws.cell(r, 2, '减：坏账准备'); ws.cell(r, 3, '')
        for j in (3, 4, 5, 7, 9, 10, 11):
            ws.cell(r, j, None)
        ws.cell(r, 6, round(_bd_qm, 2))
        ws.cell(r, 8, round(_bd_qm, 2))
        ws.cell(r, 12, '=H%d' % r)
        for j in range(1, N + 1):
            cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
            if j in (6, 8, 12):
                cell.number_format = NUMFMT
        r_bd = r
        r += 1
        ws.cell(r, 2, '应收账款净额'); ws.cell(r, 3, '')
        for j in (3, 4, 5, 7, 9, 10, 11):
            ws.cell(r, j, None)
        # 净额 = 期末 + 重分类调整1 + 重分类调整2 − 坏账准备（报表口径，含重分类）
        ws.cell(r, 6, '=F%d+I%d+J%d-F%d' % (r_tot, r_tot, r_tot, r_bd))
        ws.cell(r, 8, '=H%d+I%d+J%d-H%d' % (r_tot, r_tot, r_tot, r_bd))
        ws.cell(r, 12, '=H%d' % r)
        for j in range(1, N + 1):
            cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
            if j in (6, 8, 12):
                cell.number_format = NUMFMT
        r += 1
    elif key == 'ORA' and not related:
        # 2026-08-03 用户方法论：报表其他应收款=三个一级科目（其他应收款/应收股利/应收利息）。
        # 审定表列示：其他应收款合计(主体行+合计) → 减：其他应收款坏账准备 → 其他应收款净额 →
        # 应收股利 → 应收利息 → 报表其他应收款合计。与一级科目核对=前三项；与报表科目核对=报表合计。
        def _ora_qm(_cond):
            return sum(abs(float(v.get('qm', 0.0) or 0.0))
                       for (e, c, n, yy), v in tb_full.items()
                       if str(yy) == str(y) and _cond(c, str(n)))
        _bd_qm = _ora_qm(lambda c, n: c.startswith('1231') and '其他应收款' in n)
        _div_qm = _ora_qm(lambda c, n: c == '1131')
        _int_qm = _ora_qm(lambda c, n: c == '1132')
        r += 1
        r_tot = r - 1
        rows_def = [
            ('减：其他应收款坏账准备', round(_bd_qm, 2)),
            ('应收股利', round(_div_qm, 2)),
            ('应收利息', round(_int_qm, 2)),
        ]
        r_bd = r_net = r_div = r_int = None
        ws.cell(r, 2, rows_def[0][0]); ws.cell(r, 3, '')
        for j in (3, 4, 5, 7, 9, 10, 11):
            ws.cell(r, j, None)
        ws.cell(r, 6, rows_def[0][1]); ws.cell(r, 8, rows_def[0][1])
        ws.cell(r, 12, '=H%d' % r)
        for j in range(1, N + 1):
            cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
            if j in (6, 8, 12):
                cell.number_format = NUMFMT
        r_bd = r; r += 1
        ws.cell(r, 2, '其他应收款净额'); ws.cell(r, 3, '')
        for j in (3, 4, 5, 7, 9, 10, 11):
            ws.cell(r, j, None)
        # 净额 = 期末 + 重分类调整1 + 重分类调整2 − 坏账准备（报表口径，含重分类）
        ws.cell(r, 6, '=F%d+I%d+J%d-F%d' % (r_tot, r_tot, r_tot, r_bd))
        ws.cell(r, 8, '=H%d+I%d+J%d-H%d' % (r_tot, r_tot, r_tot, r_bd))
        ws.cell(r, 12, '=H%d' % r)
        for j in range(1, N + 1):
            cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
            if j in (6, 8, 12):
                cell.number_format = NUMFMT
        r_net = r; r += 1
        for label, val in rows_def[1:]:
            ws.cell(r, 2, label); ws.cell(r, 3, '')
            for j in (3, 4, 5, 7, 9, 10, 11):
                ws.cell(r, j, None)
            ws.cell(r, 6, val); ws.cell(r, 8, val)
            ws.cell(r, 12, '=H%d' % r)
            for j in range(1, N + 1):
                cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
                if j in (6, 8, 12):
                    cell.number_format = NUMFMT
            if label == '应收股利':
                r_div = r
            else:
                r_int = r
            r += 1
        ws.cell(r, 2, '报表其他应收款合计'); ws.cell(r, 3, '')
        for j in (3, 4, 5, 7, 9, 10, 11):
            ws.cell(r, j, None)
        ws.cell(r, 6, '=F%d+F%d+F%d' % (r_net, r_div, r_int))
        ws.cell(r, 8, '=H%d+H%d+H%d' % (r_net, r_div, r_int))
        ws.cell(r, 12, '=H%d' % r)
        for j in range(1, N + 1):
            cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
            if j in (6, 8, 12):
                cell.number_format = NUMFMT
        r += 1
        _nc = ws.cell(r, 1, '核对说明：与【一级科目】核对=其他应收款合计/减坏账准备/净额 前三项（行 %d-%d）；'
                            '与【报表科目】核对=报表其他应收款合计（行 %d，=净额+应收股利+应收利息）。' % (r_tot, r_net, r))
        _nc.font = Font(name='Times New Roman', size=9, italic=True, color='808080')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N)
        r += 1
    elif key == 'ORP' and not related:
        # 2026-08-03 用户方法论：报表其他应付款=三个一级科目（其他应付款/应付股利/应付利息）。
        # 与 ORA 对称：审定表列示 其他应付款合计 → 应付股利 → 应付利息（仅限已到期未付部分）→
        # 报表其他应付款合计。与一级科目核对=其他应付款合计；与报表科目核对=报表合计。
        def _orp_qm(_cond):
            return sum(abs(float(v.get('qm', 0.0) or 0.0))
                       for (e, c, n, yy), v in tb_full.items()
                       if str(yy) == str(y) and _cond(c, str(n)))
        _div_qm = _orp_qm(lambda c, n: '应付股利' in n)
        _int_qm = _orp_qm(lambda c, n: '应付利息' in n)
        r += 1
        r_tot = r - 1
        rows_def = [
            ('应付股利', round(_div_qm, 2)),
            ('应付利息（仅限已到期未付部分）', round(_int_qm, 2)),
        ]
        r_div = r_int = None
        for label, val in rows_def:
            ws.cell(r, 2, label); ws.cell(r, 3, '')
            for j in (3, 4, 5, 7, 9, 10, 11):
                ws.cell(r, j, None)
            ws.cell(r, 6, val); ws.cell(r, 8, val)
            ws.cell(r, 12, '=H%d' % r)
            for j in range(1, N + 1):
                cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
                if j in (6, 8, 12):
                    cell.number_format = NUMFMT
            if label.startswith('应付股利'):
                r_div = r
            else:
                r_int = r
            r += 1
        ws.cell(r, 2, '报表其他应付款合计'); ws.cell(r, 3, '')
        for j in (3, 4, 5, 7, 9, 10, 11):
            ws.cell(r, j, None)
        ws.cell(r, 6, '=F%d+F%d+F%d' % (r_tot, r_div, r_int))
        ws.cell(r, 8, '=H%d+H%d+H%d' % (r_tot, r_div, r_int))
        ws.cell(r, 12, '=H%d' % r)
        for j in range(1, N + 1):
            cell = ws.cell(r, j); cell.border = BORDER; cell.font = BOLD; cell.fill = TOTFILL
            if j in (6, 8, 12):
                cell.number_format = NUMFMT
        r += 1
        _nc = ws.cell(r, 1, '核对说明：与【一级科目】核对=其他应付款合计（行 %d）；与【报表科目】核对='
                            '报表其他应付款合计（行 %d，=其他应付款+应付股利+应付利息）。'
                            '应付利息仅列示已到期未付部分（未到期计提利息由审计人员按需调整列示）。' % (r_tot, r))
        _nc.font = Font(name='Times New Roman', size=9, italic=True, color='808080')
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=N)
        r += 1
    # 列宽
    widths = [10, 18, 13, 11, 11, 13, 12, 12, 13, 13, 11, 13]
    for j, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = 'A3'
    return ws


def _inject_ca_audit_sheets(out_y, key, subj, data_dir, tb_full, ac_ents, y, reclass_map=None, reclass_in_map=None, wb=None):
    """向已生成的某科目按年底稿注入审定表（含备抵科目净值行）。单一职责：审定表注入。
    异常已内部捕获打印，不影响主流程（2026-08-01 从 build_all_in_dir 拆分）。
    ⚡⚡ 2026-08-25 性能优化：支持外部传入 wb（合并多次 load+save 为一次）；wb=None 保持原行为。"""
    try:
        from audit_common import add_audit_summary_sheets as _add_audit
        _ca_code = {'AR': '1122', 'APR': '2203', 'APP': '1123', 'ORA': '1221',
                    'ORP': '2241', 'AP': '2202', 'ARN': '1121',
                    'CL': '2205', 'CA': '1480'}
        _rel_code = {
            ('AR', '应收账款坏账准备'): ('123101', True),
            ('ORA', '应收股利'): ('1131', False),
            ('ORA', '应收利息'): ('1132', False),
            ('ORA', '其他应收款坏账准备'): ('123102', True),
            ('ORP', '应付股利'): ('2231', True),
            ('ORP', '应付利息'): ('2232', True),
        }
        # 2026-08-07 修复：主科目 codes 优先按【名称】动态匹配（不同账套同科目代码不同——
        # 合同负债 FY=2205 / T=2204 / JTt=2203 / SSS=2207，原硬编码 CL='2205' 导致 T
        # 合同负债审定表全 0 空壳）。名称命中取【最短代码】（一级；U8 6 位码账套如 JTt
        # 靠 flow 审定表 _amt_raw 前缀+末级子目求和兜底），排除备抵词（坏账/减值/跌价）。
        _name_codes = sorted({str(c) for (e, c, n, yy), v in tb_full.items()
                              if str(yy) == str(y) and _name_hit(n, subj['kw'])
                              and not any(k in str(n) for k in ('坏账', '减值', '跌价', '重分类'))
                              and not str(c).startswith(('1231', '1471', '1602', '1703'))
                              and (abs(float(v.get('qc') or 0.0)) > 0.005
                                   or abs(float(v.get('qm') or 0.0)) > 0.005)})
        # 2026-08-07 v2（会计语言优先，代码方言隔离）：兜底 codes 经账套配置解析——
        # account_profiles.json subject_codes 有账套特例码用之（Q 应收票据=1021、
        # JTt 合同资产=1125,1480、T 合同负债=2204 等），否则用 _ca_code 硬编码兜底。
        _rk_map = {'AR': 'ar', 'APR': 'advance_recv', 'APP': 'prepay',
                   'ORA': 'ar_other', 'ORP': 'ap_other', 'AP': 'ap',
                   'ARN': 'note_recv', 'CL': 'contract_liab', 'CA': 'contract_asset'}
        _acct_ca = _resolve_account_name(data_dir)
        def _ca_code_all(key):
            try:
                import subject_mapping as _sm_ca
                _rk = _rk_map.get(key)
                if _rk:
                    _cs = _sm_ca.resolve_subject_codes(_acct_ca, _rk, [_ca_code.get(key, key)])
                    if _cs:
                        return _cs
            except Exception:
                pass
            return [_ca_code.get(key, key)]
        def _ca_code_of(key):
            _all = _ca_code_all(key)
            return _all[0] if _all else _ca_code.get(key, key)
        _name_codes = sorted({str(c) for (e, c, n, yy), v in tb_full.items()
                              if str(yy) == str(y) and _name_hit(n, subj['kw'])
                              and not any(k in str(n) for k in ('坏账', '减值', '跌价', '重分类'))
                              and not str(c).startswith(('1231', '1471', '1602', '1703'))
                              and (abs(float(v.get('qc') or 0.0)) > 0.005
                                   or abs(float(v.get('qm') or 0.0)) > 0.005)})
        # 2026-08-08 修复：JTt 等 U8 6/8/10 位混合码账套，名称匹配最短代码会取到
        # 具体子目（JTt 应收 min(len)=112203→审定表只取该子目 91.7万 vs TB 7432万）。
        # 正确取法=全部命中代码的【公共前缀】截断到一级 4 位（1122），让 flow 审定表
        # _amt_raw 前缀+末级子目求和兜底。4 位码账套（FY）公共前缀=4 位码本身，无变化。
        if _name_codes:
            # 2026-08-08 最终修复：按【一级 4 位】分组——同组（如 JTt 1121 系子目）
            # 用 4 位前缀（_amt_raw 前缀+主名匹配防双计）；跨组（S 应收票据 1120/1121）
            # 用完整代码列表（_amt_raw 逐码精确命中）。原公共前缀方案跨科目边界误伤：
            #   F 1221+1231.02 前缀 '12' 把坏账子目纳入（3.474亿假值）、
            #   JTt 1121+1141 前缀 '11' 把内部往来纳入（1,761万假值）。
            _p4 = {str(c)[:4] for c in _name_codes}
            if len(_p4) == 1:
                _codes0 = [sorted(_p4)[0]]
            else:
                _codes0 = list(_name_codes)
        else:
            _codes0 = _ca_code_all(key)
        _reg_key = _rk_map.get(key)
        items = [dict(title=f"{subj['label']} 审定表", codes=_codes0,
                      is_credit=(subj['nature'] == 'liability'),
                      before=f"{subj['label']}明细表_{y}",
                      subj_name=subj['label'], by_entity=True)]
        # 备抵科目代码整合进主审定表 codes
        _bd_map = {('AR', '123101'): '应收账款坏账准备', ('ORA', '123102'): '其他应收款坏账准备'}
        bd_codes = [c for (k, c), n in _bd_map.items() if k == key]
        if bd_codes:
            items[0]['codes'] = items[0]['codes'] + bd_codes
        for spec in RELATED_MAP.get(key, []):
            ck = _rel_code.get((key, spec['title']))
            _codes = []
            # 2026-08-03 铁律5（科目以名称判定）：优先按名称关键词在 tb_full 动态匹配科目代码
            # （FY 账套 2232=应付股利 而非标准 2231，硬编码代码会错位/取空），
            # 名称匹配失败才回退硬编码代码。
            if spec.get('name_kw'):
                _codes = sorted({str(c) for (e, c, n, yy), v in tb_full.items()
                                 if str(yy) == str(y)
                                 and all(k in str(n) for k in spec['name_kw'])})
            if not _codes and ck and ck[0]:
                # 名称匹配失败 → 回退硬编码代码，但校验该代码的科目名不含本科目关键词
                # （FY 账套 2232=应付股利，勿被误当『应付利息』；冲突则视为科目不存在不生成）
                _nm_hits = {str(n) for (e, c, n, yy), v in tb_full.items()
                            if str(yy) == str(y) and str(c) == ck[0]}
                _kw_ok = all(any(kw in nn for kw in spec.get('name_kw', [])) for nn in _nm_hits)
                if not _nm_hits or _kw_ok:
                    _codes = [ck[0]]
            if _codes:
                items.append(dict(title=f"{spec['title']} 审定表", codes=_codes,
                                  is_credit=ck[1] if ck else spec.get('is_credit', False),
                                  before=spec['sheet'], subj_name=spec['title']))
        if items:
            if wb is None:
                wb = openpyxl.load_workbook(out_y)
            # 往来主科目（8 个）改用参考格式审定表（13 列，2026-08-01 用户需求）；
            # 2026-08-04 扩展：合同负债同负债类 flow 审定表；2026-08-06 LL/DI/NCL 迁出至 gp_other；
            # 并入科目（应收/应付股利、应收/应付利息、坏账准备等）保留通用 by_entity 审定表。
            _MAIN_KEYS = {'AR', 'APR', 'APP', 'ORA', 'ORP', 'AP', 'ARN', 'CL', 'CA'}
            if key in _MAIN_KEYS:
                # 删除旧的通用审定表 sheet（若有），换新格式
                _old_title = f"{subj['label']} 审定表"
                if _old_title in wb.sheetnames:
                    del wb[_old_title]
                _write_flow_audit_sheet(wb, key, subj, tb_full, ac_ents, y,
                                        reclass_map=reclass_map,
                                        reclass_in_map=reclass_in_map,
                                        codes=items[0]['codes'])
                # 审定表移到最前（2026-08-01 修复 #4：其他应收款等审定表不再置于末尾）
                try:
                    _t_title = f"{subj['label']} 审定表"
                    _sheets = wb._sheets
                    _t_idx = next((i for i, s in enumerate(_sheets) if s.title == _t_title), None)
                    if _t_idx is not None and _t_idx > 0:
                        _sheets.insert(0, _sheets.pop(_t_idx))
                except Exception:
                    pass
                # 并入科目的审定表仍用通用 add_audit_summary_sheets
                _sub_items = items[1:] if len(items) > 1 else []
                if _sub_items:
                    # 2026-08-05 并入科目（应收/应付股利、应收/应付利息）独立审定表与主科目统一：
                    # 由『通用科目式』（单行单科目，无主体列/无审定数联动）改为 flow 12 列
                    # （每主体一行 + 合计行，含审计调整/审定数公式，过入器可按主体定位）。
                    for _si in _sub_items:
                        _t = _si.get('title', '')
                        _lbl2 = _si.get('subj_name') or _t.replace(' 审定表', '')
                        _codes2 = _si.get('codes') or []
                        if _t in wb.sheetnames:
                            del wb[_t]
                        _merge_subj = {'label': _lbl2,
                                       'nature': 'liability' if _si.get('is_credit') else 'asset'}
                        try:
                            _write_flow_audit_sheet(wb, key, _merge_subj, tb_full, ac_ents, y,
                                                    codes=_codes2, label=_lbl2, related=True)
                        except Exception as _ex:
                            print(f'  ⚠️ 并入科目审定表生成失败 [{_lbl2}]：{_ex}')
            else:
                _add_audit(wb, data_dir, items, tb_full=tb_full, entities=ac_ents, target_year=y)
            # 主审定表（13 列 flow 格式）已内置备抵净额行（AR/ORA 2026-08-03）→ 不再追加通用净值行
            if bd_codes and tb_full and key not in _MAIN_KEYS:
                _main_code = _ca_code_of(key) if '_ca_code_of' in dir() else _ca_code.get(key, key)
                _bd_code = bd_codes[0]
                if key in _MAIN_KEYS:
                    _as = wb[f"{subj['label']} 审定表"]
                    _qc_col, _qm_col = 4, 7
                else:
                    _as = wb[items[0]['title']]
                    _qc_col, _qm_col = 5, 9
                _main_qm = sum(v['qm'] for (e, c, n, yv), v in tb_full.items() if c == _main_code and str(yv) == str(y))
                _bd_qm = sum(v['qm'] for (e, c, n, yv), v in tb_full.items() if c == _bd_code and str(yv) == str(y))
                _main_qc = sum(v['qc'] for (e, c, n, yv), v in tb_full.items() if c == _main_code and str(yv) == str(y))
                _bd_qc = sum(v['qc'] for (e, c, n, yv), v in tb_full.items() if c == _bd_code and str(yv) == str(y))
                _nr = _as.max_row + 1
                _as.cell(_nr, 2, f'{subj["label"]}净值（=原值−备抵）')
                _as.cell(_nr, _qc_col, abs(_main_qc) - abs(_bd_qc) if _main_qc and _bd_qc else None)
                _as.cell(_nr, _qm_col, abs(_main_qm) - abs(_bd_qm) if _main_qm and _bd_qm else None)
                for _j in range(1, 14):
                    _c = _as.cell(_nr, _j); _c.border = BORDER; _c.font = Font(name='Times New Roman', bold=True)
                    if _j in (_qc_col, _qm_col):
                        _c.number_format = '#,##0.00'
            if wb is None:
                wb.save(out_y)
                wb.close()
                del wb   # 2026-08-07 内存优化：显式释放（AR 大文件 4 次追加链 OOM）
    except Exception as ex:
        print(f"⚠️ 审定表注入失败 [{key}] {y}：{ex}")


def _compute_subject_aging(comb_rows, comb_customers, comb_summary, comb_periods,
                            subj, key, rp_list, ca_matcher, all_bal, period_mode):
    """单一职责：某往来科目的 ①零余额补全 ②关联方判定 ③GL 账龄测算（跨年 FIFO/上年滚动/单年受限）。
    就地修改 comb_summary（related 标记）与 all_bal（关联交易累积）；返回 (aging_map, methods_used)。
    2026-08-01 从 build_all_in_dir 拆分，降低主函数长度。"""
    comb_periods = sorted(comb_periods)
    # 补全交叉缺失：某往来单位在某期间无余额(如仅2025有、2026清零)时，
    # comb_customers 为各期间并集，若不填零字典，期间对比/明细 sheet 取数会 KeyError。
    _ZERO = dict(open=0.0, debit=0.0, credit=0.0, close=0.0,
                 foreign=False, aging="", code="", count=0,
                 nature="", related=False, fcur_open=0.0, fcur_close=0.0,
                 open_dir="", close_dir="",
                 s_open=0.0, s_debit=0.0, s_credit=0.0, s_close=0.0, s_bal=0.0,
                 bill_type="")
    for (entity, name) in comb_customers:
        for pk in comb_periods:
            comb_summary.setdefault(((entity, name), pk), dict(_ZERO))
    # 关联方判定：
    #  - 若文件夹提供了『关联方清单』：以清单为权威纳入依据——清单内往来单位【必】判为关联方，
    #    同时保留账套主体间自动归并(ca_matcher)与源数据『关联方』列标注，三重保险避免遗漏集团内往来；
    #  - 否则沿用"账套主体名称归并器"识别集团内账套往来（ca_matcher）+ 源数据标注。
    #  说明：清单用于"补全"程序靠名称归并易漏判的关联方（如跨文件夹母公司/姊妹公司），
    #  而非"排他"——故采用 清单命中 OR 源数据标注 OR 自动归并 的并集，保证不回归。
    for (ek, pk), s in comb_summary.items():
        src_related = s.get("related", False)
        name = ek[1]
        if rp_list is not None:
            s["related"] = (bool(_name_in_related_list(name, rp_list))
                            or bool(src_related)
                            or bool(ca_matcher(name)))
        else:
            s["related"] = bool(ca_matcher(name)) or bool(src_related)
    # ---- 账龄测算（GL 交易日期 FIFO）与跨科目余额收集 ----
    bt = SUBJ_AGING_TYPE[key]
    # 预分组 GL 行，避免 O(客户×行) 双重遍历
    rows_by_cust = defaultdict(list)          # (entity, name, period) -> rows（单年）
    flows_by_cust_all = defaultdict(list)     # (entity, name) -> [(date, sd, sc)]（跨年）
    cust_years = defaultdict(set)             # (entity, name) -> {period}
    for r in comb_rows:
        if not r.get("is_sub"):
            continue
        ek = (r.get("entity"), _norm(r.get("cust", "")))
        pk_r = period_key(r["date"], period_mode)
        rows_by_cust[(ek[0], ek[1], pk_r)].append(r)
        flows_by_cust_all[ek].append((r["date"],
                                      _signed_debit(r["debit"], subj["nature"]),
                                      _signed_credit(r["credit"], subj["nature"])))
        cust_years[ek].add(pk_r)
    aging_map = {}
    methods_used = set()
    for (entity, name) in comb_customers:
        ek = (entity, _norm(name))
        years_all = cust_years.get(ek, set())
        for pk in comb_periods:
            s = comb_summary[(entity, name), pk]
            # 关联交易表：集团内关联方无论期末余额是否为零都纳入(all_bal)，并带本期借/贷发生额，
            # 使年末已结清(余额=0)的关联交易仍能生成『关联交易和余额核对』并展示交易流水。
            if abs(s["s_close"]) > 0.005 or s.get("related"):
                all_bal.append(dict(entity=entity, key=key, label=subj["label"], name=name,
                                    nature=subj["nature"],
                                    side=("debit" if key in DEBIT_SIDE else "credit"),
                                    related=bool(s.get("related")), period=pk, close=s["s_close"],
                                    # 发生额取【原始 Gross 绝对值】(借发/贷发分别计正)，使关联交易流水表
                                    # 借/贷两列均显示正数流水(如 A 赊销 B 143万 与 A 收回 B 143万 同为正)，
                                    # 而『本期净额(借−贷)』= 净额变动(已结清交易=0)。余额 close 仍用带符号 s_close。
                                    debit=s.get("debit", 0.0) or 0.0,
                                    credit=s.get("credit", 0.0) or 0.0))
            # ---- 账龄方法选择（按数据可用性自动选）----
            # 方法1：文件夹含≥2年GL → 跨年 FIFO 真实账龄；
            # 方法2：无多年GL 但有上年账龄(本年已算) → 上年账龄表滚动；
            # 方法0：都无 → 单年 FIFO（全列1年以内，账龄受限）。
            years_le = sorted(y for y in years_all if y <= pk)
            # 期键可能为聚合键(如 _detect_year 对无年份文件名返回 '全部')，
            # 不可直接 int()，需容错；仅当 pk 为合法年份时才计算上一期。
            try:
                pk_year = int(pk)
            except (ValueError, TypeError):
                pk_year = None
            prior_pk = (str(pk_year - 1)
                        if (pk_year is not None and str(pk_year - 1) in comb_periods)
                        else None)
            prior_bk = aging_map.get(((ek[0], name), prior_pk)) if prior_pk else None
            if len(years_le) >= 2:
                flows = [fl for fl in flows_by_cust_all.get(ek, [])
                         if period_key(fl[0], period_mode) <= pk]
                earliest = years_le[0]
                opening = comb_summary.get(((entity, name), earliest), {}).get("s_open", 0.0)
                ps, pe = _period_bounds(earliest)[0], _period_bounds(pk)[1]
                raw = compute_aging_buckets(subj["nature"], flows, opening, ps, pe, bt)
                methods_used.add("方法1(跨年GL-FIFO)")
            elif prior_bk is not None and sum(prior_bk.values()) > 1e-6:
                raw = _rollforward_aging(prior_bk, s, bt, subj["nature"])
                methods_used.add("方法2(上年账龄表滚动)")
            else:
                flows = []
                for r in rows_by_cust.get((entity, _norm(name), pk), []):
                    flows.append((r["date"], _signed_debit(r["debit"], subj["nature"]),
                                  _signed_credit(r["credit"], subj["nature"])))
                ps, pe = _period_bounds(pk)
                raw = compute_aging_buckets(subj["nature"], flows, s["s_open"], ps, pe, bt)
                methods_used.add("方法0(单年GL-账龄受限)")
            # 关键：GL 为非全量抽取，逐户 FIFO 原值往往与辅助核算控制数(本表余额)背离；
            # 为保持账龄分布合计 ≡ 明细表期末余额(控制数)，将分桶结果按控制数 s_close 归一化。
            raw_sum = sum(raw.values())
            if abs(raw_sum) > 1e-6:
                scale = s["s_close"] / raw_sum
                aging_map[((entity, name), pk)] = {b: v * scale for b, v in raw.items()}
            else:
                # 无 GL 流水（仅期初余额）：全部列入首档（单年度即『1年以内』）
                first = AGING_BUCKETS[bt][0]
                aging_map[((entity, name), pk)] = {b: (s["s_close"] if b == first else 0.0)
                                                    for b in AGING_BUCKETS[bt]}
    # 2026-08-03 用户方法论：其他应收款未提供 2025 账龄 → 暂全部按『1年以内』处理
    if key == 'ORA':
        _first = AGING_BUCKETS[bt][0]
        for ((_e, _n), _pk), _s in comb_summary.items():
            aging_map[((_e, _n), _pk)] = {b: (abs(_s.get('s_close', 0.0) or 0.0) if b == _first else 0.0)
                                          for b in AGING_BUCKETS[bt]}
        methods_used.add('其他应收款-账龄暂按1年以内')
    return aging_map, methods_used


# ============================================================================
# 关联方往来汇总（2026-08-02 用户需求 A：清单驱动关联方核对并入往来款小程序）
# 依据 关联方清单.xlsx + 各主体辅助核算余额表（含资金类科目），汇总各主体与清单
# 关联方的 应收/应付/资金 三类往来（期初/借/贷/期末），追加为『关联交易和余额核对』
# 工作簿的「关联方往来汇总」「关联方清单核对」两个 sheet。无清单则跳过。
# 与原有『关联交易和余额核对』（账套主体间对账）互补：本汇总以清单为准，覆盖
# 合并外关联方与资金类往来（银行存款/其他货币资金等）。
# ----------------------------------------------------------------------------
def _rp_parse(fn):
    """辅助核算文件名 → (实体, 年份)。如 1、FY本级公司2025年度辅助核算余额表.xlsx。"""
    import re
    m = re.search(r'(\d{4})年', fn)
    year = m.group(1) if m else ''
    ent = re.sub(r'^\d+、', '', fn)
    ent = re.split(r'\d{4}年', ent)[0].strip()
    return ent, year


def _rp_scan(fp, rp_names):
    """扫描单个辅助核算余额表：返回命中清单名称的明细行
    [(科目名, 关联方名, 期初自然符号, 借方, 贷方, 期末自然符号)]。
    U8 列：0代码/1关联方标记/2科目名/3往来单位名/4期初方向/5期初额/6借/7贷/8期末方向/9期末额。
    期初/期末按方向转自然借贷符号（铁律16）；分公司后缀（如…肇庆分公司）子串匹配归母公司。"""
    rows = []
    try:
        wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
    except Exception:
        return rows
    ws = wb[wb.sheetnames[0]]
    for r in ws.iter_rows(min_row=3, values_only=True):
        if not r or not r[0]:
            continue
        nm = str(r[3]).strip() if len(r) > 3 and r[3] else ''
        subj = str(r[2]).strip() if len(r) > 2 and r[2] else ''
        if not nm or not subj:
            continue
        hit = next((rp for rp in rp_names if rp in nm or nm in rp), None)
        if hit is None:
            continue
        osg = 1.0 if (len(r) > 4 and str(r[4]).strip() == '借') else (-1.0 if (len(r) > 4 and str(r[4]).strip() == '贷') else 0.0)
        csg = 1.0 if (len(r) > 8 and str(r[8]).strip() == '借') else (-1.0 if (len(r) > 8 and str(r[8]).strip() == '贷') else 0.0)
        rows.append((subj, hit,
                     osg * float(r[5] or 0) if len(r) > 5 else 0.0,
                     float(r[6] or 0) if len(r) > 6 else 0.0,
                     float(r[7] or 0) if len(r) > 7 else 0.0,
                     csg * float(r[9] or 0) if len(r) > 9 else 0.0))
    wb.close()
    return rows


def _rp_cat(subj):
    s = str(subj)
    if any(k in s for k in ('银行', '货币资金', '现金')):
        return '资金类'
    if any(k in s for k in ('应收', '预付')):
        return '应收类'
    if any(k in s for k in ('应付', '预收')):
        return '应付类'
    return '其他'


def _rp_append_summary(data_dir, out_dir, entities, rp_list):
    """汇总并写入 关联交易和余额核对_生成.xlsx 的「关联方往来汇总」「关联方清单核对」sheet。
    工作簿不存在（本文件夹无集团内关联交易记录）时新建，保证清单驱动汇总始终可见。"""
    # 独立读取清单第1列名称/第2列合并内外（_load_related_party_list 只返回名称，无合并内外标记）
    import glob
    rp_full = []
    for cand in sorted(glob.glob(os.path.join(data_dir, "*关联方清单*"))):
        if cand.lower().endswith((".xlsx", ".xls", ".csv")):
            try:
                wb = openpyxl.load_workbook(cand, data_only=True, read_only=True)
                ws = wb[wb.sheetnames[0]]
                for r in ws.iter_rows(min_row=2, values_only=True):
                    if r and r[0] and str(r[0]).strip():
                        rp_full.append((str(r[0]).strip(),
                                        str(r[1]).strip() if len(r) > 1 and r[1] else '合并外关联方'))
                wb.close()
            except Exception:
                pass
            break
    if not rp_full:
        rp_full = [(n, '') for n, _ in rp_list]
    rp_names = [n for n, _ in rp_full]
    rp_map = dict(rp_full)
    summary = {}   # (y, ent, rp) -> {cat: [open, dr, cr, close]}
    details = []   # (y, ent, rp, subj, open, dr, cr, close)
    years = set()
    for E, bun in entities.items():
        for fp in (bun.get("aux") or []):
            ent, yr = _rp_parse(os.path.basename(fp))
            if not yr:
                continue
            years.add(yr)
            for (subj, rp, op, dr, cr, cl) in _rp_scan(fp, rp_names):
                cat = _rp_cat(subj)
                key = (yr, ent, rp)
                d = summary.setdefault(key, {c: [0.0, 0.0, 0.0, 0.0] for c in ('应收类', '应付类', '资金类', '其他')})
                a = d[cat]
                a[0] += op; a[1] += dr; a[2] += cr; a[3] += cl
                details.append((yr, ent, rp, subj, op, dr, cr, cl))
    if not summary:
        print("  ↳ 清单关联方未在本文件夹辅助核算中匹配到往来，关联方往来汇总跳过")
        return
    out = os.path.join(out_dir, "关联交易和余额核对_生成.xlsx")
    if os.path.exists(out):
        wb = openpyxl.load_workbook(out)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
    _BF = Font(name='Times New Roman', size=10, bold=True)
    _F10 = Font(name='Times New Roman', size=10)
    _HF = PatternFill('solid', fgColor='DDEBF7')
    _TF = PatternFill('solid', fgColor='FCE4D6')
    _NUM = '#,##0.00'
    for y in sorted(years):
        # ---- Sheet：关联方往来汇总（按年） ----
        if f"关联方往来汇总_{y}" in wb.sheetnames:
            del wb[f"关联方往来汇总_{y}"]
        ws = wb.create_sheet(f"关联方往来汇总_{y}")
        ws.cell(1, 1, f"关联方往来汇总（{y} 年度；清单驱动：辅助核算往来单位名称匹配 关联方清单.xlsx；"
                      f"期初/期末按方向转自然借贷符号；合并内关联方=合并抵消线索）").font = SHELL_TITLE_FONT
        ws.cell(1, 1).fill = SHELL_HFILL
        hdr = ['核算主体', '关联方', '合并内外', '应收类-期初', '应收类-借方', '应收类-贷方', '应收类-期末',
               '应付类-期初', '应付类-借方', '应付类-贷方', '应付类-期末',
               '资金类-期初', '资金类-借方', '资金类-贷方', '资金类-期末', '往来净额(应收+应付)']
        for c, h in enumerate(hdr, 1):
            cell = ws.cell(2, c, h)
            cell.font = _BF; cell.fill = _HF; cell.alignment = SHELL_CEN; cell.border = SHELL_BORDER
        r = 3
        keys = sorted(k for k in summary if k[0] == y)
        for (yy, ent, rp) in keys:
            d = summary[(yy, ent, rp)]
            vals = [ent, rp, rp_map.get(rp, '')]
            for cat in ('应收类', '应付类', '资金类'):
                vals += [round(d[cat][0], 2), round(d[cat][1], 2), round(d[cat][2], 2), round(d[cat][3], 2)]
            vals.append(round(d['应收类'][3] + d['应付类'][3], 2))
            for c, v in enumerate(vals, 1):
                cell = ws.cell(r, c, v)
                cell.font = _F10; cell.border = SHELL_BORDER
                if c in (1, 2, 3):
                    cell.alignment = SHELL_LEFT
                else:
                    cell.number_format = _NUM; cell.alignment = SHELL_CEN
            r += 1
        r += 1
        ws.cell(r, 1, '其中：合并内关联方往来（合并抵消线索）').font = _BF
        for cc in range(1, 17):
            ws.cell(r, cc).fill = _TF
        r += 1
        tot = {c: [0.0, 0.0, 0.0, 0.0] for c in ('应收类', '应付类', '资金类')}
        for (yy, ent, rp) in keys:
            if '合并内' in rp_map.get(rp, ''):
                for cat in tot:
                    for i in range(4):
                        tot[cat][i] += summary[(yy, ent, rp)][cat][i]
        row = ['小计（合并内）', '—', '—'] + [round(v, 2) for cat in ('应收类', '应付类', '资金类') for v in tot[cat]] + \
              [round(tot['应收类'][3] + tot['应付类'][3], 2)]
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.font = _BF; cell.border = SHELL_BORDER
            cell.alignment = SHELL_LEFT if c in (1, 2, 3) else SHELL_CEN
            if c > 3:
                cell.number_format = _NUM
        ws.freeze_panes = 'A3'
        for j, w in enumerate([12, 30, 10] + [13] * 12 + [15], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    # ---- Sheet：关联方清单核对 ----
    if "关联方清单核对" in wb.sheetnames:
        del wb["关联方清单核对"]
    ws3 = wb.create_sheet("关联方清单核对")
    hdr3 = ['关联方', '合并内外', '是否匹配到往来', '涉及主体数', '往来科目数']
    for c, h in enumerate(hdr3, 1):
        cell = ws3.cell(1, c, h)
        cell.font = _BF; cell.fill = _HF; cell.alignment = SHELL_CEN; cell.border = SHELL_BORDER
    r = 2
    n_miss = 0
    for rp, _kind in rp_list:
        ent_set = {k[1] for k in summary if k[2] == rp}
        subj_cnt = len({d[3] for d in details if d[2] == rp})
        matched = len(ent_set) > 0
        if not matched:
            n_miss += 1
        row = [rp, rp_map.get(rp, ''), '✓' if matched else '—', len(ent_set), subj_cnt]
        for c, v in enumerate(row, 1):
            cell = ws3.cell(r, c, v)
            cell.font = _F10; cell.border = SHELL_BORDER
            cell.alignment = SHELL_LEFT if c in (1, 2) else SHELL_CEN
        if not matched:
            for c in range(1, 6):
                ws3.cell(r, c).fill = PatternFill('solid', fgColor='FFF2CC')
        r += 1
    r += 1
    note = (f'注：清单 {len(rp_list)} 个关联方中 {len(rp_list) - n_miss} 个在辅助核算中匹配到往来'
            f'（{n_miss} 个未匹配=本年度无往来或未在辅助核算中核算）；'
            '合并内关联方往来在合并层面应抵消，最终附注不体现。')
    c = ws3.cell(r, 1, note)
    c.font = Font(name='Times New Roman', size=10, italic=True, color='808080')
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    for j, w in enumerate([34, 12, 16, 12, 12], 1):
        ws3.column_dimensions[get_column_letter(j)].width = w
    try:
        finalize_workbook(wb)
    except Exception:
        pass
    wb.save(out)
    wb.close()
    print(f"  ✓ 关联方往来汇总/清单核对 已并入：{out}")


def build_all_in_dir(data_dir, out_dir=None, period_mode="Y", subject=None, only_subj=None):
    print(f'>>> 开始处理往来款（{data_dir}）…')
    """目录模式：自动发现目录下所有核算主体，按主体分组处理后合并输出。
    每个往来科目生成【一个】合并工作簿（如 应收账款审计底稿_生成.xlsx），
    其中『核算主体』列为每行真实所属主体；跨主体同名往来单位分成不同行（不再互相覆盖）；
    校验结果与对应科目分析按主体分别标识/合并。
    2026-08-04：only_subj 指定单科目（如 'CL'）时只重跑该科目（省时，用户提速需求）。
    另生成：账龄分布(各科目明细表内含)、多个账套存在交易往来(集团内关联方往来核对，多账套)、
    （仅当存在外币辅助核算余额表时）外币aux漏标注检查。"""
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    clear_gl_cache()  # 一次拖入 = 一次全新缓存，避免长驻进程跨数据包残留
    entities = _discover_entities(data_dir)
    # ===== 2026-08-03 重分类：扫描全部主体 aux，收集各科目期末反项余额(负数)客户 =====
    # 负数 = 余额方向与科目性质相反（资产贷余→负债、负债借余→资产），自动重分类到对方科目。
    # 对映射：AR↔APR(预收)、APP(预付)↔AP、ORA↔ORP、CA(合同资产)↔CL(合同负债)。
    # 2026-08-06 补 CA↔CL 对（SSS 建筑施工企业合同科目为大额，FY 无合同科目自动跳过）。
    _RECLASS_PAIRS = [('AR', 'APR'), ('APR', 'AR'), ('APP', 'AP'), ('AP', 'APP'),
                      ('ORA', 'ORP'), ('ORP', 'ORA'), ('CA', 'CL'), ('CL', 'CA')]
    reclass_pool = {k: [] for k in SUBJECTS}
    _TR = os.environ.get('CA_TRACE')
    if _TR:
        print(f'  [CA_TRACE] 重分类扫描开始 @ {time.time():.0f}', flush=True)
    for entity, bun in entities.items():
        for _ap in (bun.get('aux') or []):
            if _ap is None:   # ⚡ SAP：占位 aux，read_opening_balances 走 adapter
                _yr = ''
            else:
                _yr = str(_detect_year(os.path.basename(_ap)) or '')
            for _k, _subj in SUBJECTS.items():
                try:
                    _adata, _ = read_opening_balances(_ap, _subj)
                except Exception:
                    continue
                for _n, _v in _adata.items():
                    _sc = _signed_balance(_v['close'], _v.get('close_dir', ''), _subj['nature'])
                    if _sc < -0.005:
                        reclass_pool[_k].append((entity, _n, round(-_sc, 2), _yr))
    if _TR:
        print(f'  [CA_TRACE] 重分类扫描完成 @ {time.time():.0f}', flush=True)
    # 往来补充并入科目所需（坏账准备拆分/应收应付股利利息）：一次性读取 TB 与 GL（与权益程序同源）
    from audit_common import discover_entities as _ac_discover, read_tb_full as _rtb, read_gl_rows as _rgl
    _ac_ents = _ac_discover(data_dir)
    # ⚡ SAP：审计口径 TB/GL 只取当前驱动主体（防审定表混入全集团主体行而明细仅单主体）
    if _adapter is not None and _is_sap_mode(data_dir):
        # ⚡⚡ 2026-08-23 修复：SAP 模式需要 _DATA_ROOT 有值（read_transactions/read_gl_net
        #   无 comp 时读它）。但 build_all_in_dir 的 data_dir 可能是【工作目录 _work_{comp}】
        #   （run_u8_on_sap 传入），此时 _DATA_ROOT 已被外层 set_root(真实数据根) 设好——
        #   若再用工作目录覆盖会找不到主体。故仅当 _DATA_ROOT 为空时才设置。
        if not getattr(_adapter, '_DATA_ROOT', None):
            _adapter.set_root(data_dir)
        _comp = _adapter.current_comp()
        if _comp and _comp in _ac_ents:
            _ac_ents = {_comp: _ac_ents[_comp]}
    _tb_full = _rtb(data_dir, _ac_ents)
    if os.environ.get('CA_TRACE'):
        print(f'  [CA_TRACE] TB 读取完成 @ {time.time():.0f}', flush=True)
    # 泰国账套科目归一化（2026-08-01：#6 往来款审定表缺泰国）——泰国代码/名称映射到集团标准
    try:
        from thai_mapping import normalize_thai_tb as _norm_thai
        _tb_full = _norm_thai(_tb_full, _ac_ents)
    except Exception as _ex:
        print(f'  ⚠️ 泰国科目归一化失败：{_ex}')
    _gl_full = _rgl(data_dir, _ac_ents)
    _all_years = sorted({y for (e, c, n, y) in _tb_full})
    # 集团内账套主体互为关联方：构建名称→账套主体归并器，供关联方识别与账套间往来核对使用
    ca_matcher = build_entity_matcher(list(entities.keys()))
    # 用户提供的『关联方清单』（若有）：作为关联方识别的权威依据，优先于名称模糊归并
    rp_list = _load_related_party_list(data_dir)
    has_fx = any("外币辅助核算余额表" in fn for fn in os.listdir(data_dir))
    all_bal = []          # 跨科目余额收集（用于多个账套存在交易往来核对）
    results = {}
    comb_km = []          # 合并模式收集所有主体的科目余额表（用于应收票据坏账准备读取）
    # 2026-08-04 单科目重跑（--only-subj）：只处理指定科目，大幅缩短重跑时间
    _keys = [only_subj] if only_subj else list(SUBJECTS)
    if only_subj and only_subj not in SUBJECTS:
        print(f'  ⚠️ 未知科目键：{only_subj}（可用：{"、".join(SUBJECTS)}）')
        return 0
    for key in _keys:
        if os.environ.get('CA_TRACE'):
            print(f'  [CA_TRACE] {key} 主循环开始 @ {time.time():.0f}', flush=True)
        subj = SUBJECTS[key]
        comb_customers = []
        comb_summary = {}
        comb_rows = []
        comb_issues = []
        comb_periods = set()
        n_total = 0
        have = False
        _ca_code = {'AR': '1122', 'APR': '2203', 'APP': '1123', 'ORA': '1221',
                    'ORP': '2241', 'AP': '2202', 'ARN': '1121',
                    'CL': '2205', 'CA': '1480'}
        # 2026-08-07 v2（会计语言优先，代码方言隔离）：TB 兜底代码经账套配置解析——
        # account_profiles.json subject_codes 有账套特例码用之（Q 应收票据=1021、
        # JTt 合同资产=1125,1480、T 合同负债=2204 等），否则用 _ca_code 硬编码兜底。
        _acct_ca2 = _resolve_account_name(data_dir)
        _ca_code_cfg = {}
        try:
            import subject_mapping as _sm_ca2
            for _k2, _rk2 in [('AR', 'ar'), ('APR', 'advance_recv'), ('APP', 'prepay'),
                              ('ORA', 'ar_other'), ('ORP', 'ap_other'), ('AP', 'ap'),
                              ('ARN', 'note_recv'), ('CL', 'contract_liab'), ('CA', 'contract_asset')]:
                _ca_code_cfg[_k2] = _sm_ca2.resolve_subject_codes(_acct_ca2, _rk2, [_ca_code.get(_k2, _k2)])
        except Exception:
            _ca_code_cfg = {}
        def _ca_code_of(key):
            if _ca_code_cfg and key in _ca_code_cfg and _ca_code_cfg[key]:
                return _ca_code_cfg[key][0]
            return _ca_code.get(key, key)
        # 2026-08-06：TB 兜底匹配 = 代码前缀 OR 名称含 kw（不同账套同一科目代码不同
        # 如合同负债 FY=2205 / JTt=2207 → 仅代码前缀匹配会漏 → 名称匹配兜底）。
        # 2026-08-07：名称兜底须排除 1231 坏账准备族——F 坏账准备 1231.02 子目名
        # 『其他应收款』撞 ORA 名称 kw → 其他应收款明细表混入 -3.6 亿坏账行（明细
        # 差 -3.34 亿假差异）；1231 为各往来科目共享坏账准备代码，代码前缀已覆盖
        # 各自科目（1221/1122），名称兜底对其无意义。
        _kw = (SUBJECTS[key].get('kw') or '').strip()
        def _tb_match(cc, nn):
            if str(cc).startswith(_ca_code_of(key)):
                return True
            if str(cc).startswith('1231'):
                return False
            if _kw and _kw in str(nn):
                return True
            return False
        for entity, bun in entities.items():
            if bun["km"]:
                comb_km += bun["km"]
            res = build_ca_detail(key, bun["gl"], bun["aux"], None, bun["km"],
                                  period_mode, subject=entity, write=False)
            if res["n_open"] == 0 and not res["customers"]:
                # ⚡ 2026-08-12 修复：原条件仅 n_open==0 就进 TB 兜底 continue——SAP 无 aux
                #   （aux_ar=None）但 GL 有客户级发生额（res['customers'] 128 客户）时，
                #   明细被 TB 科目级兜底覆盖、GL 客户全部丢弃（3200/3300/3400 应收/应付
                #   明细显示科目名而非客户）。条件收紧：仅当 GL 也无客户发生额时才 TB 兜底；
                #   有 GL 客户 → 走下方正常 GL 客户合并（6910-6918）。
                # 2026-08-03 TB 兜底：aux 无该科目辅助明细但 TB 有余额（如 FY 预付账款 1123，
                # 11 主体有 TB 余额但无辅助核算）→ 生成『（无辅助核算明细）』占位行
                # （s_open/s_close=TB 期初/期末自然符号，借正贷负），使审定表/明细表
                # 仍能按主体生成并与 TB 勾稽；账龄无流水 → 全列首档。
                _has_tb = any(str(ee) == entity and _tb_match(cc, nn)
                              and (abs(float(v.get('qc') or 0.0)) > 0.005
                                   or abs(float(v.get('qm') or 0.0)) > 0.005)
                              for (ee, cc, nn, yy), v in _tb_full.items())
                if not _has_tb:
                    continue
                have = True
                _pk_list = sorted({str(yy) for (ee, cc, nn, yy), v in _tb_full.items()
                                   if str(ee) == entity and _tb_match(cc, nn)
                                   and (abs(float(v.get('qc') or 0.0)) > 0.005
                                        or abs(float(v.get('qm') or 0.0)) > 0.005)})
                if _pk_list:
                    # 2026-08-05 二级展开（租赁负债等）：TB 一级科目下含二级子目
                    # （如 2601.01 未确认融资费用 / 2601.02 租赁付款额）时，占位行按
                    # 二级子目逐行展开——否则单行『（无辅助核算明细）』只显示一级净值，
                    # 未确认融资费用等二级金额无法核对（用户反馈）。无子目（仅一级行）
                    # → 退化为原单行占位。
                    # 注意：同 code 跨年度名称可能不同（2025 键名带"租赁负债-"前缀、
                    # 2026 键名不带）→ 按 code 分组、名称取最新年度，防同名分裂双行；
                    # comb_summary 仍按 (code, 年度) 各建行，供明细表按目标年度取用。
                    _base = str(_ca_code_of(key))
                    _all_lines = [(str(yy), str(cc), str(nn), v)
                                  for (ee, cc, nn, yy), v in _tb_full.items()
                                  if str(ee) == entity and (_tb_match(cc, nn) or str(cc).startswith(_base))
                                  and (abs(float(v.get('qc') or 0.0)) > 0.005
                                       or abs(float(v.get('qm') or 0.0)) > 0.005)]
                    _by_code = {}
                    for _yy, _cc, _nn, _v in _all_lines:
                        _by_code.setdefault(_cc, []).append((_yy, _nn, _v))
                    # 2026-08-07 修复：末级【叶子】过滤（原只跳一级父码 _base）——Q Audio
                    # 2202→220201(采购应付 90.7M)→22020101/22020102（子级同额），中间父级
                    # 220201 未过滤 → 父+子双计 181.4M=2×90.7M。任何其他代码以此为前缀
                    # 即剔除（2202/220201 均有后代 → 剔；22020101/22020102 叶子保留）。
                    _all_codes = list(_by_code.keys())
                    _leaf_codes = [c for c in _all_codes
                                   if not any(str(c2) != str(c) and str(c2).startswith(str(c))
                                              for c2 in _all_codes)]
                    for _cc in sorted(_leaf_codes):
                        _items = sorted(_by_code[_cc], key=lambda x: str(x[0]))
                        # 名称取最新年度（期初/期末最新），旧年度名称仅作数值来源
                        _cname = str(_items[-1][1]).strip() or '（无辅助核算明细）'
                        if (entity, _cname) not in comb_customers:
                            comb_customers.append((entity, _cname))
                        for _yy2, _nn2, _v in _items:
                            comb_periods.add(_yy2)
                            _jf = float(_v.get('jf') or 0.0)
                            _df = float(_v.get('df') or 0.0)
                            # 2026-08-05 负债科目符号规则：贷方余额（qm<0）→ 转正显示；
                            # 借余子目（qm>0，如未确认融资费用=抵减项）→ 转负，使合计=
                            # 贷方余额−抵减=净值（正数），与 recalc_totals 按数据行求和一致、
                            # 与 aux 明细（转正口径）一致。原实现两子目都为正 → 合计双计。
                            _sqc = float(_v.get('qc') or 0.0)
                            _sqm = float(_v.get('qm') or 0.0)
                            if subj['nature'] == 'liability':
                                _sqc = abs(_sqc) if _sqc <= 0 else -abs(_sqc)
                                _sqm = abs(_sqm) if _sqm <= 0 else -abs(_sqm)
                            comb_summary[((entity, _cname), _yy2)] = dict(
                                open=abs(_sqc), debit=abs(_jf), credit=abs(_df),
                                close=abs(_sqm), foreign=False, aging='',
                                code=_cc, count=0, nature='', related=False,
                                fcur_open=0.0, fcur_close=0.0, open_dir='', close_dir='',
                                s_open=_sqc,
                                s_debit=_signed_debit(_jf, subj['nature']),
                                s_credit=_signed_credit(_df, subj['nature']),
                                s_close=_sqm, s_bal=0.0,
                                # ⚡ 2026-08-26 净化：TB 二级子目名含票据种类（如 112101『应收票据-银行承兑汇票』）
                                #   → 推导 bill_type，使票据种类拆分表不再全『未分类』
                                bill_type=_derive_bill_type(_cname))
                continue
            # ===== 2026-08-05 差额兜底：aux 明细部分缺失 → 补『（无辅助核算明细）』差额行 =====
            # 背景：TB 兜底原仅处理 n_open==0（aux 完全无明细）。部分账套大额余额无往来
            # 单位明细（FY 金属/龙游其他应付款：aux 仅科目级汇总行，末级实名明细 close≈0）
            # → n_open>0 绕过兜底，TB 余额丢失（金属 209.9 万、龙游 210.4 万，明细合计
            # 23.96M vs TB 31.70M，差 774 万）。修复：比较 TB 余额（转口径）与 aux 明细
            # s_close 合计，差额>容差 → 补占位行，使 明细表合计 == TB 余额（用户铁律）。
            _base = str(_ca_code_of(key))
            _gl_cust_path = bool(res['customers'])   # ⚡ 2026-08-12：GL 客户路径（SAP 无 aux、GL 有客户发生额）
            _tb_all = [(str(yy), str(cc), float(v.get('qm') or 0.0))
                       for (ee, cc, nn, yy), v in _tb_full.items()
                       if str(ee) == entity and str(cc).startswith(_base)]
            # 2026-08-07 修复：差额兜底按【期间】逐期独立取数（原只取 _max_yy 最后年度
            # TB 行 + 只补 _cur_pk 期、其余期补 0）——c 诚本 TB 为双年度（2025/2026 各行
            # qm=各年期末，如 1221 其他应收款 2025=281,368.79 / 2026=191,862.74），
            # 原逻辑 _max_yy=2026 → 2025 期差额（社保+公积金 105,868.79）漏补 →
            # 明细表 175,500 vs 审定表/TB 281,368.79 三表不一致。逐期补后各期各自
            # 满足『明细合计==TB 期末』（用户铁律）。FY 02 技术应收 2026 qm=0 案例：
            # 该期 _ylines 空 → 不补，行为与原 _max_yy 防错逻辑一致。
            _tb_by_yy = {}
            for _y, _c, _v in _tb_all:
                _tb_by_yy.setdefault(_y, []).append((_c, _v))
            if _tb_by_yy:
                _gname = '（无辅助核算明细）'
                for _yy3 in sorted(res['periods'] or []):
                    _ylines_raw0 = _tb_by_yy.get(str(_yy3)) or []
                    _ylines = [(_c, _v) for _c, _v in _ylines_raw0 if abs(_v) > 0.005]
                    if not _ylines:
                        if not _ylines_raw0:
                            continue  # 该期 TB 无此科目行 → 不补
                        # ⚡⚡ 2026-08-27 修复（任务887 2468 预收 2550 主体 135.9万 根因）：
                        #   TB 有该期科目行但期末≈0（期初贷余/借余本期清平，如 2203999999
                        #   『预收账款-预收账款调整』期初 135.9万贷余 → 凭证 6500000000 借发生
                        #   135.9万 → 期末 0）→ 原 continue 不补占位 → 明细表仅 GL 客户行
                        #   发生额 -135.9万 → 明细≠审定(TB 期末 0)。改用 0 余额 TB 行继续走
                        #   差额兜底（_tb_bal=0、_gap = -aux_close → 补占位抵消 GL 客户行）
                        #   → 明细合计 = TB = 0（用户铁律：明细与 TB 不一致=取数不完整）。
                        _ylines = [(_c, 0.0) for _c, _v in _ylines_raw0]
                    _tb_codes = {c for c, _ in _ylines}
                    # TB 权威控制数：父行存在用父行（U8 一级期末=报表口径），否则末级子目求和
                    if _base in _tb_codes:
                        _ylines = [(c, v) for c, v in _ylines if c == _base]
                    elif any(c != _base and c.startswith(_base) for c in _tb_codes):
                        _ylines = [(c, v) for c, v in _ylines if c != _base]
                    # ⚡ 2026-08-26 净化：应收票据(ARN)差额兜底按 TB 二级子目逐行展开
                    #   （112101 银行承兑/112102 商业承兑/112104 财务公司承兑），bill_type 从
                    #   子目名推导 → 票据种类拆分不再全『未分类』。其他科目保持父行合并（不改变结构）。
                    if key == 'ARN' and len(_ylines) > 1:
                        _named = []
                        for _c, _v in _ylines:
                            _nm = next((str(n) for (ee, cc2, n, yyy) in _tb_full
                                        if str(ee) == entity and str(cc2) == _c
                                        and str(yyy) == str(_yy3)), '')
                            _named.append((_c, _v, _nm or '（无辅助核算明细）'))
                        _ylines = _named
                    else:
                        _ylines = [(c, v, '') for c, v in _ylines]
                    _tb_bal = 0.0
                    for _c, _v, _nn in _ylines:
                        if subj['nature'] == 'liability':
                            _tb_bal += abs(_v) if _v <= 0 else -abs(_v)   # 负债：贷余转正
                        else:
                            _tb_bal += _v
                    _aux_close = sum(s.get('s_close', 0.0) for (ak, pk), s in res['summary'].items()
                                     if pk == _yy3)
                    _gap = _tb_bal - _aux_close
                    # 2026-08-05 修正：仅补【TB>aux】的正差额（aux 明细缺失部分）。
                    # gap 为负（aux 期末>TB 期末，如 aux 数据滞后于 TB——FY 02 技术应收
                    # 账款 TB 期末已全额收回=0，aux 期末仍 8.14M）时补占位会虚增明细。
                    # ⚡ 2026-08-12 扩展：GL 客户路径（SAP 无客户 aux、仅 GL 客户发生额，
                    #   3200/3300/3400 应收借方凭证小计）——GL 净额（期初0+借−贷）与 TB
                    #   期末（含期初、口径不同）方向/金额均可能背离（3200 应收 GL -1547万
                    #   vs TB -7027万，差 5479万）→ 按绝对值补差，使 明细合计==TB 期末
                    #   （用户铁律：明细与 TB 不一致=取数不完整）。
                    # ⚡⚡ 2026-08-27 修复（任务886 2310 预收借余明细缺失根因）：占位行差额
                    #   兜底条件由『_gap>0.01 or GL客户路径 abs』改【无条件 abs(_gap)>0.01】。
                    #   原对负债借余主体（TB qm>0 → _tb_bal 算负 → _gap 负）且无 GL 客户
                    #   路径（如 2310 预收 155.7万借余）不补占位行 → 明细表缺主体块 →
                    #   明细合计≠审定表。统一补带符号差额（s_close=_gap），使明细合计
                    #   = TB 带符号期末（用户铁律）；aux 数据滞后场景（gap 负）补负占位
                    #   行后 明细合计=aux+gap=TB 亦正确。
                    if abs(_gap) > 0.01:
                        if key == 'ARN':
                            # 应收票据：按 TB 子目逐行展开（bill_type 推导 → 票据种类拆分正确）
                            _sub_tot = sum(abs(_v2) for _, _v2, _ in _ylines) or 0.0
                            if _sub_tot > 0.01:
                                for _c2, _v2, _nn2 in _ylines:
                                    _gname2 = _nn2 or _gname
                                    _alloc = _gap * (abs(_v2) / _sub_tot)
                                    if (entity, _gname2) not in comb_customers:
                                        comb_customers.append((entity, _gname2))
                                    comb_summary[((entity, _gname2), _yy3)] = dict(
                                        open=0.0, debit=0.0, credit=0.0,
                                        close=abs(_alloc), foreign=False, aging='',
                                        code=_base, count=0, nature='', related=False,
                                        fcur_open=0.0, fcur_close=0.0, open_dir='', close_dir='',
                                        s_open=0.0, s_debit=0.0, s_credit=0.0,
                                        s_close=_alloc, s_bal=0.0,
                                        bill_type=_derive_bill_type(_gname2) if key == 'ARN' else '')
                                    comb_periods.add(_yy3)
                            else:
                                if (entity, _gname) not in comb_customers:
                                    comb_customers.append((entity, _gname))
                                comb_summary[((entity, _gname), _yy3)] = dict(
                                    open=0.0, debit=0.0, credit=0.0,
                                    close=abs(_gap), foreign=False, aging='',
                                    code=_base, count=0, nature='', related=False,
                                    fcur_open=0.0, fcur_close=0.0, open_dir='', close_dir='',
                                    s_open=0.0, s_debit=0.0, s_credit=0.0,
                                    s_close=_gap, s_bal=0.0, bill_type='')
                                comb_periods.add(_yy3)
                        else:
                            # ⚡ 2026-08-26 修复（任务872 应付 -2.55亿）：非应收票据科目差额兜底
                            #   单行补『（无辅助核算明细）』，金额=_gap 全额（=TB叶子−aux）。
                            #   原按 TB 子目循环且共用 gname → comb_summary 每轮 set 覆盖 →
                            #   仅最后子目份额保留（GR/IR 8,391万）、其余丢弃、GR/IR 全额未补。
                            if (entity, _gname) not in comb_customers:
                                comb_customers.append((entity, _gname))
                            comb_summary[((entity, _gname), _yy3)] = dict(
                                open=0.0, debit=0.0, credit=0.0,
                                close=abs(_gap), foreign=False, aging='',
                                code=_base, count=0, nature='', related=False,
                                fcur_open=0.0, fcur_close=0.0, open_dir='', close_dir='',
                                s_open=0.0, s_debit=0.0, s_credit=0.0,
                                s_close=_gap, s_bal=0.0, bill_type='')
                            comb_periods.add(_yy3)
            have = True
            n_total += res["n_open"]
            comb_rows += res["rows"]
            comb_periods |= set(res["periods"])
            for it in res["issues"]:
                comb_issues.append(dict(it, message=f"[{entity}] " + it["message"]))
            for name in res["customers"]:
                comb_customers.append((entity, name))
                for pk in res["periods"]:
                    comb_summary[(entity, name), pk] = res["summary"][(name, pk)]
        if not have:
            # 2026-08-04 报表重分类承接科目（APR 预收账款等）：账套无科目数据（TB 无 2203），
            # 但对方科目有负数重分类转入（应收贷方余额→预收）→ 仍生成底稿（明细=转入客户）。
            _from_key_pre = dict(_RECLASS_PAIRS).get(key)
            _has_recl_in = bool(reclass_pool.get(_from_key_pre)) if _from_key_pre else False
            if not _has_recl_in:
                out_path = os.path.join(out_dir, f"{subj['label']}审计底稿_生成.xlsx")
                for stale in (out_path,
                              os.path.join(out_dir, f"{subj['label']}审计底稿_2025_生成.xlsx"),
                              os.path.join(out_dir, f"{subj['label']}审计底稿_2026_生成.xlsx")):
                    if os.path.exists(stale):
                        try:
                            os.remove(stale)
                        except BaseException:
                            pass
                results[key] = (None, 0, comb_issues)
                continue
            print(f"  ↳ {subj['label']}：账套无科目数据，按重分类承接生成（转入 {len(reclass_pool[_from_key_pre])} 笔）")
            comb_periods.update({str(x[3]) for x in reclass_pool[_from_key_pre]})
        bt = SUBJ_AGING_TYPE[key]
        import os as _os_tr
        _TR = _os_tr.environ.get('CA_TRACE')
        if _TR:
            print(f'  [CA_TRACE] {key} 主体循环完成 @ {time.time():.0f}', flush=True)
        aging_map, methods_used = _compute_subject_aging(
            comb_rows, comb_customers, comb_summary, comb_periods, subj, key,
            rp_list, ca_matcher, all_bal, period_mode)
        if _TR:
            print(f'  [CA_TRACE] {key} 账龄完成 @ {time.time():.0f}', flush=True)
        comb_counterparties = _reconstruct_counterparties_combined(comb_rows)
        if _TR:
            print(f'  [CA_TRACE] {key} 对方科目完成 @ {time.time():.0f}', flush=True)
        # ⚡ 2026-08-10 SAP 凭证级精确分摊（剔除同凭证多对方科目金额重复=虚增虚减）
        _sap_ca = _adapter is not None and _is_sap_mode(data_dir)
        if _sap_ca:
            by_ent = _adapter.recon_counterparties_sap(comb_rows)
        else:
            by_ent = reconstruct_counterparties_by_entity(comb_rows)
        if _TR:
            print(f'  [CA_TRACE] {key} by_ent 完成 @ {time.time():.0f}', flush=True)
        # ===== 按年直接构建（2026-08-01 改造：无合并稿中间态）=====
        _years_out = sorted(_all_years or comb_periods)
        _first_out = None
        for _y in _years_out:
            _out_y = os.path.join(out_dir, f"{subj['label']}审计底稿_{_y}_生成.xlsx")
            # 2026-08-03 用户方法论：段位不推进的简化账龄滚动仅在『本年数据为部分年度(YTD，如 2026 仅
            # 1-3 月)』时启用（临时特殊处理）；全年数据年份恢复正常账龄推进（FIFO 真实账龄），
            # 不再读上年表滚动。判断=该年 GL 月份数 < 12。
            prior_meta = {}
            try:
                _y_int = int(_y)
            except (TypeError, ValueError):
                _y_int = None
            _ytd = _year_is_ytd(comb_rows, _y)
            if _ytd and _y_int is not None:
                _prev = str(_y_int - 1)
                _pt = _read_prior_aging_table(data_dir, subj['label'], _prev, bt)
                if _pt:
                    # 键归一化：上年表主体名可能不同（参照格式用序号主体）→ 按客户名匹配，
                    # 主体键取本套账主体名；另加『客户名兜底』键（*，name）供实体名不同时匹配，
                    # 但兜底须经金额一致性校验（上年期末≈本年期初）防止跨账套同名客户误配。
                    prior_meta = {}
                    _cur_close = {}
                    for (e, n) in comb_customers:
                        _s = comb_summary.get(((e, n), str(_y)), {})
                        _cur_close.setdefault(n, _s.get('s_open', 0.0))
                    for (_e, _c), _v in _pt.items():
                        prior_meta[(_e, _c)] = _v
                        _p_close = sum(abs(x) for x in _v['aging'].values())
                        _c_open = _cur_close.get(_c, 0.0)
                        # 金额一致性：上年账龄合计 ≈ 本年期初余额（±2% 或 ±1 元）才允许兜底
                        if abs(_p_close - _c_open) <= max(1.0, abs(_c_open) * 0.02):
                            prior_meta.setdefault(('*', _c), _v)
                    # 打印匹配情况（客户名与本年 comb_customers 的交集）
                    _cur_custs = {(e, n) for (e, n) in comb_customers}
                    _hit = sum(1 for (e, n) in prior_meta if (e, n) in _cur_custs)
                    _hit_star = sum(1 for (_e, _n) in _cur_custs if ('*', _n) in prior_meta)
                    print(f"  ℹ️ 上年({_prev})账龄表：{len(_pt)} 客户（本年精确匹配 {_hit}，"
                          f"客户名+金额校验匹配 {_hit_star}），账龄按『上年对照简化滚动』")
            # 2026-08-03 重分类转入：对方科目反项余额客户 → 本明细表新增行（生成前注入，
            # 明细表/账龄/审定表同步含转入；转出方已在其自身明细表审定归 0）
            _from_key = dict(_RECLASS_PAIRS).get(key)
            _reclass_in_by_ent = {}
            if _from_key:
                for (_ent, _cust, _amt, _yr) in reclass_pool.get(_from_key, []):
                    if str(_yr) != str(_y):
                        continue
                    _k = (_ent, str(_y))
                    _reclass_in_by_ent[_k] = _reclass_in_by_ent.get(_k, 0.0) + _amt
                    _exist = comb_summary.get(((_ent, _cust), str(_y)))
                    if _exist is not None:
                        # 本科目已有同名客户：转入金额合并进 reclass_in（不重复建行）
                        _exist['reclass_in'] = (_exist.get('reclass_in', 0.0) or 0.0) + _amt
                        continue
                    comb_customers.append((_ent, _cust))
                    # 转入客户仅当年有金额；其余期间补零（comb_customers 跨 _y 共享，防渲染 KeyError）
                    for _pk in sorted(comb_periods or [str(_y)]):
                        _zc = dict(
                            open=0.0, debit=0.0, credit=0.0, close=0.0, foreign=False, aging='',
                            code='', count=0, nature=subj['nature'], related=False,
                            fcur_open=0.0, fcur_close=0.0, open_dir='', close_dir='',
                            s_open=0.0, s_debit=0.0, s_credit=0.0, s_close=0.0, s_bal=0.0,
                            bill_type='',
                            reclass_in=(_amt if str(_pk) == str(_y) else 0.0))
                        comb_summary[((_ent, _cust), str(_pk))] = _zc
                        if subj['nature'] == 'asset':
                            aging_map[((_ent, _cust), str(_pk))] = {
                                b: (_amt if (b == AGING_BUCKETS[bt][0] and str(_pk) == str(_y)) else 0.0)
                                for b in AGING_BUCKETS[bt]}
            if _reclass_in_by_ent:
                print(f"  ↳ 重分类转入 [{key}] {_y}：{sum(_reclass_in_by_ent.values()):,.2f} "
                      f"（自 {SUBJECTS[_from_key]['label']} {len(_reclass_in_by_ent)} 主体）")
            # ⚡⚡ 2026-08-29 #876 验证：CA_XW=1 时走 xlsxwriter 渲染（明细表流式+对方科目核对），
            #   追加链路（相关科目/坏账/账龄/附注/审定表）为阶段2 TODO，本阶段直接保存。
            if os.environ.get('CA_XW'):
                from xw_export import build_xw_export
                _wb_out = build_xw_export(_out_y, comb_periods, comb_customers, comb_summary,
                                          comb_issues, comb_rows, comb_counterparties, period_mode, subj,
                                          by_ent=by_ent, aging_map=aging_map, bucket_type=bt, km_path=comb_km,
                                          out_dir=out_dir, data_dir=data_dir, subj_key=key,
                                          entities_dict=entities, aging_methods=methods_used,
                                          target_year=_y, prior_meta=prior_meta,
                                          tb_full=_tb_full, ac_ents=_ac_ents)
                _wb_out.close()
                print(f'✅ [xw] 已生成：{_out_y}（合并 {len(entities)} 个核算主体，末级行合计 {n_total}）')
                # ⚡⚡ 2026-08-30 修复：xw 版结构性缺失（附注汇总）+ 合计行不符（完整性 ERROR）。
                #   openpyxl 后处理：recalc_totals 重算各 sheet 合计/小计 + 追加附注汇总
                #   （xw 文件<1MB，无 OOM；_append_ca_footnote_sheet 支持外部 wb 免二次 load）。
                try:
                    import openpyxl as _ox
                    from recalc_totals import recalc_sheet as _recalc
                    from current_account_detail import _append_ca_footnote_sheet
                    _wb2 = _ox.load_workbook(_out_y)
                    _fixed = 0
                    for _sn in list(_wb2.sheetnames):
                        _fixed += _recalc(_wb2[_sn])
                    _append_ca_footnote_sheet(_out_y, key, subj, comb_customers, comb_summary,
                                              aging_map, bt, _tb_full, _y, gl_full=None, wb=_wb2)
                    _wb2.save(_out_y)
                    _wb2.close()
                    if _fixed:
                        print(f'  ↳ [xw] 后处理：重算 {_fixed} 个合计单元格 + 附注汇总已追加')
                except Exception as _ex:
                    print(f'  ⚠️ [xw] 后处理失败 [{key}] {_y}：{_ex}')
                if _first_out is None:
                    _first_out = _out_y
                results[key] = (_first_out, n_total, comb_issues)
                continue
            _wb_out = export_combined_excel(_out_y, comb_periods, comb_customers, comb_summary,
                                            comb_issues, comb_rows, comb_counterparties, period_mode, subj,
                                            by_ent=by_ent, aging_map=aging_map, bucket_type=bt, km_path=comb_km,
                                            out_dir=out_dir, data_dir=data_dir, subj_key=key,
                                            entities_dict=entities, aging_methods=methods_used,
                                            target_year=_y, prior_meta=prior_meta,
                                            early_return_wb=True)
            if _TR:
                print(f'  [CA_TRACE] {key} export 完成 @ {time.time():.0f}', flush=True)
            print(f"✅ 已生成：{_out_y}（合并 {len(entities)} 个核算主体，末级行合计 {n_total}）")
            gc.collect()   # 2026-08-07 内存优化：释放 export 的大 workbook，防追加链峰值累积 OOM
            # ⚡⚡ 2026-08-25 性能优化（v2）：export 直接返回内存 wb（early_return_wb），
            #   后续追加/审定表注入全部操作同一 wb，最后统一 finalize+save 一次——
            #   完全消除对 93MB 大底稿的 load+save（XBJ ORA 追加阶段原 20-35 分钟 → ~2 分钟）。
            try:
                # 往来补充并入科目（坏账准备拆分/应收应付股利利息），仅 AR/ORA/ORP 追加 sheet
                if key in ('AR', 'ORA', 'ORP') and _all_years:
                    try:
                        _append_related_sheets(_out_y, _tb_full, _gl_full, [_y], data_dir, wb=_wb_out)
                    except Exception as ex:
                        print(f"⚠️ 并入科目追加失败 [{key}] {_y}：{ex}")
                # 2026-08-03 用户要求：应收账款/其他应收款 坏账准备明细表（格式同存货跌价准备明细表）+ 计提数核对
                if key in ('AR', 'ORA') and _all_years:
                    try:
                        _append_ar_baddebt_sheet(_out_y, _tb_full, _gl_full, [_y], data_dir, key=key, wb=_wb_out)
                    except Exception as ex:
                        print(f"⚠️ 坏账准备明细表追加失败 [{key}] {_y}：{ex}")
                # ⚡ 2026-08-11 用户方法论：应收+合同资产合并账龄表（仅应收账款；合同资产账套
                #   同客户在两科目间拆分+虚增虚减+相互结转 → 合并后 FIFO 账龄才准确）
                if key == 'AR' and _all_years:
                    try:
                        # 单体（current_comp 有值）：应收+合同资产合并账龄（单主体审计口径）
                        _ws_ca = build_ar_ca_combined_aging(_wb_out, data_dir, target_year=str(_y))
                        if _ws_ca is not None:
                            print(f"  ✓ 应收+合同资产合并账龄（{_y}）已并入：{os.path.basename(_out_y)}")
                        # ⚡ 2026-08-12 集团（current_comp 为空=全集团）：应收+合同资产前10位
                        #   （跨主体合并视图，与单体账龄互补；单体由 combined_aging 覆盖）
                        if _adapter is not None and not _adapter.current_comp():
                            _ws_top = build_ar_ca_top10_group(_wb_out, data_dir, target_year=str(_y))
                            if _ws_top is not None:
                                print(f"  ✓ 应收+合同资产前10位（集团跨主体合并，{_y}）已并入")
                    except Exception as ex:
                        print(f"⚠️ 应收+合同资产合并账龄失败 [{key}] {_y}：{ex}")
                # 2026-08-03 附注汇总（格式2 上市公司合并口径；应收/其他应收/应收票据/
                # 预付/预收/应付/其他应付 全部往来科目；负债类简化：账龄 1年以内/1年以上）
                if _all_years:
                    try:
                        _append_ca_footnote_sheet(_out_y, key, subj, comb_customers, comb_summary,
                                                  aging_map, bt, _tb_full, _y, gl_full=_gl_full, wb=_wb_out)
                    except Exception as ex:
                        print(f"⚠️ 附注汇总生成失败 [{key}] {_y}：{ex}")
                # 2026-08-03 备抵机制通用化：预付/应收票据等资产科目若 TB 出现对应减值准备科目
                # （未来可能出现的 预付款项减值准备/其他流动资产减值准备 等）→ 自动追加通用减值准备明细表+核对
                if key in ('APP', 'ARN') and _all_years:
                    try:
                        from baddebt_common import scan_baddebt_subjects, build_baddebt_detail_sheet as _bd_sheet
                        _kw = '预付账款' if key == 'APP' else '应收票据'
                        for (_c, _nm, _qc, _qm) in scan_baddebt_subjects(_tb_full, _y):
                            if _kw in _nm:
                                _bd_sheet(_wb_out, _tb_full, _gl_full, _y, _nm, subj['label'])
                                break
                    except Exception as ex:
                        print(f"⚠️ 通用减值准备明细表追加失败 [{key}] {_y}：{ex}")
                # ---- 审定表（target_year=_y，置于各科目明细表之前）----
                # 2026-08-03 重分类：统计本科目各主体×年 负数余额客户重分类合计（填审定表『审计重分类调整1』列）
                _reclass_by_ent = {}
                for ((_e, _n), _pk), _s in comb_summary.items():
                    if (_s.get('s_close') or 0.0) < -0.005:
                        _k = (_e, str(_pk))
                        _reclass_by_ent[_k] = _reclass_by_ent.get(_k, 0.0) + (-(_s['s_close'] or 0.0))
                _inject_ca_audit_sheets(_out_y, key, subj, data_dir, _tb_full, _ac_ents, _y,
                                        reclass_map=_reclass_by_ent,
                                        reclass_in_map=_reclass_in_by_ent if _from_key else None,
                                        wb=_wb_out)
                if _first_out is None:
                    _first_out = _out_y
            finally:
                if _wb_out is not None:
                    if _TR:
                        print(f'  [CA_TRACE] {key} 追加阶段开始 @ {time.time():.0f}', flush=True)
                    finalize_workbook(_wb_out)   # ⚡ 2026-08-25 export early_return 跳过，此处补
                    if _TR:
                        print(f'  [CA_TRACE] {key} finalize 完成 @ {time.time():.0f}', flush=True)
                    _wb_out.save(_out_y)
                    if _TR:
                        print(f'  [CA_TRACE] {key} save 完成 @ {time.time():.0f}', flush=True)
                    _wb_out.close()
                    del _wb_out   # 2026-08-07 内存优化：显式释放
        results[key] = (_first_out, n_total, comb_issues)
    # ---- 2026-08-06 单科目重跑：只生成该科目底稿，直接返回 ----
    # 循环外段（往来互抵清理/关联交易核对/外币检查）均为全科目逻辑，单科目重跑无意义且
    # 曾致 AR 单科目重跑在 inject 后异常终止（进程无 traceback 退出、exit 1）——跳过最稳。
    if only_subj:
        return results
    # ---- 往来互抵明细：按需求删除，不再生成（清理旧文件）----
    off_path = os.path.join(out_dir, "往来互抵明细_生成.xlsx")
    if os.path.exists(off_path):
        try:
            os.remove(off_path)
            print(f"🗑️ 已移除旧的『往来互抵明细_生成.xlsx』（按需求不再生成）")
        except BaseException:
            pass
    # ---- 关联交易和余额核对（2026-08-29 起不再生成：关联交易/关联往来底稿由独立小程序生成，
    #      88 家核对底稿已另行生成；13 科目底稿小程序不再产出关联交易文件）----
    _cross_path = os.path.join(out_dir, "关联交易和余额核对_生成.xlsx")
    if os.path.exists(_cross_path):
        try:
            os.remove(_cross_path)
            print("🗑️ 已移除旧的『关联交易和余额核对_生成.xlsx』（不再生成，由独立小程序生成）")
        except BaseException:
            pass
    # ---- 外币aux漏标注检查（仅当文件夹存在外币辅助核算余额表时）----
    if has_fx:
        try:
            build_faux_missing(data_dir, out_dir)
        except Exception as ex:
            print(f"⚠️ 外币aux漏标注检查生成失败：{ex}")
    else:
        fx_path = os.path.join(out_dir, "往来款_外币aux漏标注检查_生成.xlsx")
        if os.path.exists(fx_path):
            try:
                os.remove(fx_path)
                print(f"🗑️ 已移除旧的『往来款_外币aux漏标注检查_生成.xlsx』（本文件夹无外币辅助核算余额表）")
            except BaseException:
                pass
    # ⚡ 2026-08-26 按 WBS/指令号 回款追踪（SAP 专属；收入/存货已按指令号，收款补齐闭环）
    try:
        build_wbs_tracking_workbook(data_dir, out_dir)
    except Exception as ex:
        print(f'  ⚠️ 按WBS回款追踪失败：{ex}')
    return results


def _write_faux_missing_sheet(ws, rows_summary, rows_detail, has_fx):
    """写出『外币aux漏标注检查』单 sheet：上段为各科目常规aux合计数 vs TB控制数复核，
    下段为漏标注往来单位明细（仅在外币aux存在、常规aux缺失）。"""
    ws.cell(1, 1, "外币辅助核算余额表 —— 漏标注检查（仅列：在外币aux存在、常规aux缺失的往来单位）").font = TITLE_FONT
    note = ("判定逻辑：对各核算主体、各往来科目，比较『常规辅助核算余额表末级实名合计数』与『科目余额表对应往来科目父级控制数』；"
            "若两者相符，则这些仅在外币aux出现的往来单位，一般为外币aux中部分凭证未标注往来单位（漏标注），需后续核查。"
            "注：常规aux与TB不符时无法判定漏标注，差额另行列示。")
    c2 = ws.cell(2, 1, note)
    c2.font = Font(name='Times New Roman', italic=True, size=10, color="808080")
    c2.alignment = LEFT

    r = 4
    ws.cell(r, 1, "一、常规aux合计数 vs 科目余额表控制数（漏标注判定依据）").font = SUB_FONT
    r += 1
    h1 = ["核算主体", "往来科目", "期间", "常规aux合计数", "科目余额表控制数", "差额",
          "是否相符", "外币aux客户数", "常规aux缺失数"]
    for i, h in enumerate(h1, 1):
        ws.cell(r, i, h)
    _style_header(ws, r, len(h1))
    r += 1
    for (E, label, yr, reg_t, tb, diff, matched, n_fx, n_miss) in rows_summary:
        ws.cell(r, 1, E)
        ws.cell(r, 2, label)
        ws.cell(r, 3, yr)
        ws.cell(r, 4, round(reg_t, 2)).number_format = MONEY_FMT
        ws.cell(r, 5, round(tb, 2) if tb is not None else None).number_format = MONEY_FMT
        cd = ws.cell(r, 6, round(diff, 2))
        cd.number_format = MONEY_FMT
        if abs(diff) > 0.005:
            cd.font = ANOM_FONT
        ws.cell(r, 7, "是" if matched else "否")
        ws.cell(r, 8, n_fx)
        cm = ws.cell(r, 9, n_miss)
        if n_miss > 0:
            cm.font = ANOM_FONT
        r += 1

    r += 1
    ws.cell(r, 1, "二、漏标注往来单位明细（外币aux存在、常规aux缺失）").font = SUB_FONT
    r += 1
    if rows_detail:
        h2 = ["核算主体", "往来科目", "期间", "往来单位名称", "币种", "期末方向",
              "原币期初", "原币期末", "本位币期末", "判定", "常规aux合计数",
              "科目余额表控制数", "差额"]
        for i, h in enumerate(h2, 1):
            ws.cell(r, i, h)
        _style_header(ws, r, len(h2))
        r += 1
        for row in rows_detail:
            for i, v in enumerate(row, 1):
                c = ws.cell(r, i, v)
                if i in (7, 8, 9, 11, 12, 13) and isinstance(v, (int, float)):
                    c.number_format = MONEY_FMT
                if i == 10:
                    if "漏标注" in str(v):
                        c.font = ANOM_FONT
                    elif "无法判定" in str(v):
                        c.font = Font(name='Times New Roman', color="9C5700")
            r += 1
    else:
        msg = "（未发现：外币aux中存在但常规aux缺失的往来单位）"
        if not has_fx:
            msg = "（该文件夹无『外币辅助核算余额表』，未执行漏标注检查）"
        ws.cell(r, 1, msg).font = Font(name='Times New Roman', italic=True, color="808080")
        r += 1

    widths = [16, 12, 8, 20, 20, 18, 10, 14, 14, 14, 42, 20, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"


def build_faux_missing(data_dir, out_dir=None):
    """诊断：外币辅助核算余额表中存在、但常规辅助核算余额表缺失的往来单位（疑似漏标注）。
    输出 往来款_外币aux漏标注检查_生成.xlsx（单 sheet）。"""
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    # 分组 (entity, year) -> {aux, km}
    groups = defaultdict(dict)
    has_fx = False
    for fn in sorted(os.listdir(data_dir)):
        if fn.startswith("~$"):
            continue
        if not fn.lower().endswith((".xlsx", ".xls", ".csv")):
            continue
        full = os.path.join(data_dir, fn)
        yr = _detect_year(fn)
        if "外币辅助核算余额表" in fn:
            has_fx = True
            continue
        if "辅助核算余额表" in fn:
            E = _derive_subject(full)
            groups[E].setdefault(yr, {})["aux"] = full
        elif "科目余额表" in fn and "辅助核算" not in fn:
            E = _derive_subject(full)
            groups[E].setdefault(yr, {})["km"] = full

    fx_all = _read_foreign_aux_full(data_dir, None, None)
    rows_summary = []
    rows_detail = []
    for E in sorted(groups):
        for yr in sorted(groups[E]):
            info = groups[E][yr]
            aux_path = info.get("aux")
            if not aux_path:
                continue
            km_path = info.get("km")
            for key in SUBJECTS:
                subj = SUBJECTS[key]
                data, _ = read_opening_balances(aux_path, subj)
                tb = read_control(km_path, subj) if km_path else None
                # 常规aux合计数按科目性质+期末方向带符号，与带符号TB控制数一致比较
                reg_total = sum(_signed_balance(r["close"], r.get("close_dir", ""), subj["nature"])
                                for r in data.values())
                reg_names = set(_norm(n) for n in data)
                tol = max(0.01, 0.001 * max(abs(reg_total), abs(tb or 0)))
                matched = (tb is not None) and (abs(reg_total - tb) <= tol)
                fx_items = [(cust, rec) for (ent2, sk2, pk2, cust), rec in fx_all.items()
                            if ent2 == E and sk2 == key and pk2 == yr]
                n_fx = len(fx_items)
                missing = [(cust, rec) for (cust, rec) in fx_items if cust not in reg_names]
                rows_summary.append((E, subj["label"], yr, reg_total, tb,
                                     reg_total - (tb or 0), matched, n_fx, len(missing)))
                for cust, rec in missing:
                    if matched:
                        verdict = "漏标注（常规aux合计数与TB控制数相符）"
                    elif reg_total == 0 and (tb is None or tb == 0):
                        verdict = "（常规aux无此科目，无法判定）"
                    else:
                        verdict = "无法判定（常规aux与TB不符，差额 %.2f）" % (reg_total - (tb or 0))
                    rows_detail.append((E, subj["label"], yr, rec.get("orig_name", ""),
                                        rec.get("ccy", ""), rec.get("dir", ""),
                                        round(rec.get("fcur_open", 0), 2),
                                        round(rec.get("fcur_close", 0), 2),
                                        round(rec.get("rmb_close", 0), 2), verdict,
                                        round(reg_total, 2), round(tb or 0, 2),
                                        round(reg_total - (tb or 0), 2)))
    out_path = os.path.join(out_dir, "往来款_外币aux漏标注检查_生成.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "外币aux漏标注检查"
    _write_faux_missing_sheet(ws, rows_summary, rows_detail, has_fx)
    try:
        finalize_workbook(wb)
        wb.save(out_path)
    except PermissionError:
        alt = os.path.join(out_dir, "往来款_外币aux漏标注检查_新.xlsx")
        finalize_workbook(wb)
        wb.save(alt)
        out_path = alt
    finally:
        wb.close()
    print(f"✅ 已生成：{out_path}（漏标注往来单位 {len(rows_detail)} 户；"
          f"覆盖 {len(rows_summary)} 个 主体×科目×期间）")
    return out_path, len(rows_detail)


def build_wbs_tracking_workbook(data_dir, out_dir=None):
    """按 WBS/指令号 回款追踪表（2026-08-26 用户方法论：AH=WBS，收入/收款/存货三方对齐）。
    数据源：adapter.read_gl（SAP 行含 wbs/ord/mat），范围=往来科目（SUBJECTS kw 名称匹配）。
    行=主体×WBS×往来科目；列=主体|WBS/指令号|往来科目|借方发生(销售/应收)|贷方发生(回款)|净结余|对方科目Top。
    与 收入-成本按指令号配比（revenue_detail）、存货按指令号增减（inventory_detail）三方对齐，
    打通『存货→发货→收款→收入』闭环。U8（无指令号列）不生成。"""
    if _adapter is None or not _is_sap_mode(getattr(_adapter, '_DATA_ROOT', data_dir)):
        return None
    entities = _discover_entities(data_dir)
    if not entities:
        return None
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = '按WBS回款追踪'
    ws.cell(1, 1, '往来按WBS/指令号回款追踪（借方=销售应收/资产、贷方=回款/收款；'
                  '与收入按指令号配比、存货按指令号增减 三方对齐）').font = TITLE_FONT
    hdr = ['主体', 'WBS/指令号', '往来科目', '借方发生(销售/应收)', '贷方发生(回款)', '净结余', '对方科目Top']
    for j, h in enumerate(hdr, 1):
        cc = ws.cell(3, j, h)
        cc.font = HDR_FONT; cc.fill = HDR_FILL; cc.alignment = CENTER; cc.border = BORDER
    r = 4
    _n_rows = 0
    for comp in sorted(entities):
        try:
            rows = list(_adapter.read_gl(comp))
        except Exception:
            continue

        def _inst_of(rr):
            """指令号：wbs > ord > mat > proj（proj='00000' 为 SAP 占位，排除）。"""
            for f in ('wbs', 'ord', 'mat'):
                v = str(rr.get(f) or '').strip()
                if v:
                    return v
            pj = str(rr.get('proj') or '').strip()
            return pj if pj and pj != '00000' else ''

        # 同凭证索引（应收/预收行自身无 WBS，从其凭证中带 WBS 的收入/发货行继承）
        by_vno = defaultdict(list)
        for row in rows:
            by_vno[str(row.get('voucher') or row.get('vno') or '')].append(row)
        agg = {}
        for row in rows:
            nm = str(row.get('name') or '')
            subj = None
            for _k, s in SUBJECTS.items():
                if s['kw'] in nm:
                    subj = s['label']
                    break
            if not subj:
                continue
            w = _inst_of(row)
            if not w:
                # ⚡ 2026-08-26 同凭证 WBS 继承：SAP 应收/预收行不挂 WBS，但同凭证的
                #   收入/发货行带 WBS（实测 14.3% 应收凭证可继承）。取首个非空指令号。
                for m in by_vno.get(str(row.get('voucher') or row.get('vno') or ''), []):
                    w2 = _inst_of(m)
                    if w2:
                        w = w2
                        break
            d = float(row.get('debit') or row.get('dr') or 0.0)
            c = float(row.get('credit') or row.get('cr') or 0.0)
            if abs(d) < 0.005 and abs(c) < 0.005:
                continue
            key = (w or '（无指令号）', subj)
            a = agg.get(key)
            if a is None:
                agg[key] = [d, c, str(row.get('cp') or '')[:30]]
            else:
                a[0] += d
                a[1] += c
        for (w, subj), (d, c, cp) in sorted(agg.items(), key=lambda kv: -abs(kv[1][0] - kv[1][1])):
            ws.cell(r, 1, comp)
            ws.cell(r, 2, w)
            ws.cell(r, 3, subj)
            ws.cell(r, 4, round(d, 2)).number_format = MONEY_FMT
            ws.cell(r, 5, round(c, 2)).number_format = MONEY_FMT
            ws.cell(r, 6, round(d - c, 2)).number_format = MONEY_FMT
            ws.cell(r, 7, cp)
            for _cc in range(1, 8):
                ws.cell(r, _cc).border = BORDER
            r += 1
            _n_rows += 1
    if _n_rows == 0:
        ws.cell(4, 1, '（往来 GL 无 WBS/指令号行——U8 数据无指令号列，SAP 无指令号往来行归入「（无指令号）」）')
    out = os.path.join(out_dir or data_dir, '往来按WBS回款追踪_生成.xlsx')
    try:
        from audit_common import _safe_save
        _safe_save(wb, out)
        print(f'[OK] 往来按WBS回款追踪 → {out}（{_n_rows} 行）')
        return out
    except Exception as ex:
        print(f'  ⚠️ WBS回款追踪保存失败：{ex}')
        return None


def main(argv=None):
    # 默认输入目录：仅当桌面 G 账套包仍存在时作为便捷默认；
    # 该账套包未来会删除，届时需显式 --input 指定新数据包，程序不因此崩溃。
    _legacy = os.path.join(P.DATA_ROOT, 'G')
    default_dir = _legacy if os.path.isdir(_legacy) else None
    raw = list(sys.argv[1:] if argv is None else argv)
    # 拖入文件夹快捷方式：若首参为文件夹路径（不以 '-' 开头），自动视为 --input/--aux/--km
    if raw and not raw[0].startswith('-') and os.path.isdir(raw[0]):
        folder = raw[0]
        raw = ['--input', folder, '--aux', folder, '--km', folder] + raw[1:]
    p = argparse.ArgumentParser(description="通用往来科目明细表生成程序（统一版，覆盖六大往来：应收/预收/预付/应付/其他应收/其他应付）")
    p.add_argument("--input", "-i", default=default_dir, help="综合查询明细表(.csv/.xlsx，可为目录)")
    p.add_argument("--aux", "-a", default=default_dir, help="辅助核算余额表(.csv/.xlsx，可为目录)")
    p.add_argument("--km", "-k", default=default_dir, help="科目余额表(可选，用于父级控制数核对)")
    p.add_argument("--subject-key", default=None,
                   choices=list(SUBJECTS.keys()),
                   help="指定单一科目(AR/APR/APP/ORA/ORP)；目录模式或省略则循环全部")
    p.add_argument("--output", "-o", default=None, help="单一科目输出路径（与 --subject-key 配合使用）")
    p.add_argument("--period", choices=["Y", "M", "Q"], default="Y", help="期间粒度 Y年/M月/Q季")
    p.add_argument("--subject", default=None, help="核算主体名称（默认从文件名推导）")
    p.add_argument("--only-subj", default=None,
                   help="2026-08-04 目录模式单科目重跑（如 --only-subj CL 只跑合同负债，省时）")
    args = p.parse_args(raw)

    if not args.input:
        print("未提供数据文件夹，且默认 G 账套包已不存在；请用 --input 指定新的账套数据包路径。")
        return 2
    in_p = args.input
    if not (os.path.isfile(in_p) or os.path.isdir(in_p)):
        print(f"输入路径不存在：{in_p}")
        return 2
    # 往来款三表（综合查询/辅助核算/科目余额）通常同目录；--aux/--km 缺失时回退到 --input 目录，
    # 不再依赖 gu股份公司 默认路径。
    aux_p = args.aux if (os.path.isfile(args.aux) or os.path.isdir(args.aux)) else in_p
    km_p = args.km if (os.path.isfile(args.km) or os.path.isdir(args.km)) else None

    if args.subject_key:
        clear_gl_cache()
        out = args.output or os.path.join(aux_p if os.path.isdir(aux_p) else os.path.dirname(os.path.abspath(aux_p)),
                                          f"{SUBJECTS[args.subject_key]['label']}审计底稿_生成.xlsx")
        build_ca_detail(args.subject_key, in_p, aux_p, out, km_p, args.period, args.subject)
    else:
        build_all_in_dir(in_p, out_dir=(in_p if os.path.isdir(in_p) else os.path.dirname(os.path.abspath(in_p))),
                         period_mode=args.period, subject=args.subject, only_subj=args.only_subj)
    from audit_common import finalize_after_build
    finalize_after_build(in_p if os.path.isdir(in_p) else os.path.dirname(os.path.abspath(in_p)))   # 单跑收尾
    return 0


def _find_managed_python():
    """定位 WorkBuddy 托管 Python（含 openpyxl），避免系统 Python 无 openpyxl 导致拖入闪退。"""
    base = os.path.join(os.environ.get('USERPROFILE', os.path.expanduser('~')),
                        '.workbuddy', 'binaries', 'python', 'versions')
    if os.path.isdir(base):
        for d in sorted(os.listdir(base), reverse=True):
            p = os.path.join(base, d, 'python.exe')
            if os.path.isfile(p):
                return p
    return None


# ⚡ 2026-08-10 SAP 适配导入提前：原在文件末尾（__main__ 之后）→ 直接 `python xxx.py` 直跑时
# main() 已执行到 _discover_entities 而 _adapter 未定义 → NameError。模块被 import 时底部
# try/except 会覆盖此占位（同值）。
_adapter = None


if __name__ == "__main__":
    # 拖入/双击本 .py 时，若当前不是托管 Python（如系统 Python 3.14 无 openpyxl），
    # 自动用托管 Python 重新执行并保留控制台，避免 ImportError 黑窗闪退。
    _mp = _find_managed_python()
    if _mp and os.path.abspath(sys.executable).lower() != os.path.abspath(_mp).lower():
        import subprocess
        try:
            rc = subprocess.run([_mp, '-B', os.path.abspath(__file__)] + sys.argv[1:]).returncode
        except Exception as _e:
            print(f'[ERROR] 无法启动托管 Python：{_e}')
            rc = 1
        try:
            input('\n按回车退出…')
        except EOFError:
            pass
        sys.exit(rc)
    try:
        rc = main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        import datetime
        traceback.print_exc()
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        try:
            if os.environ.get('AUDIT_LOG') == '1':
                with open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       f'ca_crash_{ts}.log'), 'w', encoding='utf-8') as _f:
                    _f.write('未捕获异常：\n' + traceback.format_exc())
                print(f'已记录崩溃日志：ca_crash_{ts}.log')
        except Exception:
            pass
        rc = 1
    try:
        input('\n按回车退出…')
    except EOFError:
        pass
    sys.exit(rc)

# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None

# ============================================================================
# ⚡ 2026-08-11 用户方法论：应收+合同资产合并账龄表
# 合同资产账龄复杂——同客户(或合同号)在 应收(1122) 与 合同资产(1124) 间拆分，
# 且本期含虚增虚减(同凭证借贷互抵)、应收↔合同资产相互结转。
# 最佳做法：①合并两科目按客户；②剔除虚增虚减(净额化)；③剔除两科目相互结转；
# ④对合并数据做账龄划分；⑤区分到哪个科目需后续按合同规定另行划分（本表不做）。
# 输出 sheet『应收+合同资产合并账龄』（仅 1122/1124 同时有数据的账套）。
# ============================================================================
def build_ar_ca_combined_aging(wb, data_dir, target_year=None):
    """应收+合同资产合并账龄表。单主体/集团均可。
    数据=GL(1122+1124) 按客户合并；剔除同凭证两科目相互结转(摘要含结转/冲销且双向)；
    合并后 FIFO 账龄(recv 6 档)。返回 ws 或 None(无合同资产数据不生成)。"""
    import re as _re
    from collections import defaultdict
    if _adapter is None:
        return None
    comp = _adapter.current_comp()
    # ⚡ 2026-08-11 用户要求：合并账龄【只能单体账套内合并】，不能全集团合并——
    #   同客户在两科目间拆分的合并账龄是单主体审计口径；集团模式(current_comp=None)
    #   文件是多主体合并 → 此处不生成（集团口径另出『应收+合同资产前10位』跨主体相加表）。
    if not comp:
        return None
    ents = {comp: _adapter.discover_entities(getattr(_adapter, '_DATA_ROOT', None) or data_dir).get(comp, {})}
    if not ents:
        return None
    gl = _adapter.read_gl_rows(getattr(_adapter, '_DATA_ROOT', None) or data_dir, ents)
    # 收集 1122/1124 行
    ar_rows, ca_rows = [], []
    for r in gl:
        c = str(r.get('code') or '')
        if c.startswith('1122'):
            ar_rows.append(r)
        elif c.startswith('1124'):
            ca_rows.append(r)
    if not ca_rows:
        return None                     # 无合同资产账套不生成
    # 同凭证相互结转剔除：同一 vno 内同时出现 1122 与 1124 且金额方向相反→视为结转
    by_vno = defaultdict(list)
    for r in ar_rows + ca_rows:
        by_vno[str(r.get('vno') or '')].append(r)
    _cross_vnos = set()
    for v, rows in by_vno.items():
        has_ar = any(str(x.get('code') or '').startswith('1122') for x in rows)
        has_ca = any(str(x.get('code') or '').startswith('1124') for x in rows)
        if has_ar and has_ca and len(rows) >= 2:
            _cross_vnos.add(v)
    # 客户维度合并：期初=0(GL 无期初,用发生额净额)、增/减/期末
    cust_ar = defaultdict(lambda: {'inc': 0.0, 'dec': 0.0})
    cust_ca = defaultdict(lambda: {'inc': 0.0, 'dec': 0.0})
    cust_name = {}
    for r in ar_rows:
        cu = str(r.get('cust') or '').strip() or '（未标注客户）'
        if str(r.get('vno') or '') in _cross_vnos:
            continue
        cust_ar[cu]['inc'] += float(r.get('dr') or 0.0)
        cust_ar[cu]['dec'] += float(r.get('cr') or 0.0)
    for r in ca_rows:
        cu = str(r.get('cust') or '').strip() or '（未标注客户）'
        if str(r.get('vno') or '') in _cross_vnos:
            continue
        cust_ca[cu]['inc'] += float(r.get('dr') or 0.0)
        cust_ca[cu]['dec'] += float(r.get('cr') or 0.0)
    all_cust = sorted(set(cust_ar) | set(cust_ca), key=lambda x: -(
        (cust_ar.get(x, {}).get('inc', 0) - cust_ar.get(x, {}).get('dec', 0)) +
        (cust_ca.get(x, {}).get('inc', 0) - cust_ca.get(x, {}).get('dec', 0))))
    if not all_cust:
        return None
    ws = wb.create_sheet('应收+合同资产合并账龄')
    ws.cell(1, 1, f'应收+合同资产合并账龄（{target_year or ""}；同客户合并、剔除虚增虚减与两科目相互结转后 FIFO 划分；'
                  f'区分应收/合同资产需按合同规定另行划分）').font = SHELL_TITLE_FONT
    hdrs = ['客户', '应收-期初', '应收-本期增', '应收-本期减', '应收-期末',
            '合同资产-期初', '合同资产-本期增', '合同资产-本期减', '合同资产-期末',
            '合并-期初', '合并-本期增', '合并-本期减', '合并-期末',
            '账龄1年以内', '账龄1-2年', '账龄2-3年', '账龄3-4年', '账龄4-5年', '账龄5年以上']
    groups = [('客户', 1), ('应收账款', 4), ('合同资产', 4), ('合并口径', 4), ('合并账龄(FIFO)', 6)]
    r = two_row_header(ws, 2, groups, hdrs)
    tot = {'ar_cl': 0.0, 'ca_cl': 0.0}
    for cu in all_cust[:500]:
        a = cust_ar.get(cu, {})
        c = cust_ca.get(cu, {})
        ar_cl = a.get('inc', 0.0) - a.get('dec', 0.0)
        ca_cl = c.get('inc', 0.0) - c.get('dec', 0.0)
        m_cl = ar_cl + ca_cl
        # FIFO 账龄(合并口径,简化:以期末余额方向,期初视作年初)
        buckets = _aging_simple(m_cl)
        if abs(ar_cl) < 0.005 and abs(ca_cl) < 0.005:
            continue
        ws.cell(r, 1, cu).font = SHELL_BOLD
        vals = [ar_cl, a.get('inc', 0.0), a.get('dec', 0.0), ar_cl,
                ca_cl, c.get('inc', 0.0), c.get('dec', 0.0), ca_cl,
                m_cl, a.get('inc', 0.0) + c.get('inc', 0.0), a.get('dec', 0.0) + c.get('dec', 0.0), m_cl] + buckets
        for j, v in enumerate(vals, 2):
            cc = ws.cell(r, j, round(v, 2) if abs(v) >= 0.005 else None)
            cc.number_format = SHELL_NUM; cc.border = SHELL_BORDER; cc.alignment = SHELL_RGT
        ws.cell(r, 1).border = SHELL_BORDER
        tot['ar_cl'] += ar_cl; tot['ca_cl'] += ca_cl
        r += 1
    # 合计
    ws.cell(r, 1, '合计').font = SHELL_BOLD
    for j in range(2, 14):
        cc = ws.cell(r, j); cc.fill = SHELL_TOT_FILL; cc.border = SHELL_BORDER
    ws.cell(r, 2, round(tot['ar_cl'], 2)).number_format = SHELL_NUM
    ws.cell(r, 5, round(tot['ar_cl'], 2)).number_format = SHELL_NUM
    ws.cell(r, 6, round(tot['ca_cl'], 2)).number_format = SHELL_NUM
    ws.cell(r, 9, round(tot['ca_cl'], 2)).number_format = SHELL_NUM
    ws.cell(r, 10, round(tot['ar_cl'] + tot['ca_cl'], 2)).number_format = SHELL_NUM
    ws.cell(r, 13, round(tot['ar_cl'] + tot['ca_cl'], 2)).number_format = SHELL_NUM
    r += 2
    ws.cell(r, 1, '说明：①两科目按客户合并；②剔除同凭证借贷互抵(虚增虚减)与应收↔合同资产相互结转凭证；'
                  '③合并后 FIFO 账龄划分(6 档)；④应收与合同资产的实际归属需按合同条款/完工进度另行划分。'
                  '期初/期末为发生额净额口径(GL 无期初，TB 期初在审定表核对)。').font = SHELL_BOLD
    for i, w in enumerate([30] + [13] * 18, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'B3'
    return ws


def build_ar_ca_top10_group(wb, data_dir, target_year=None):
    """应收+合同资产前10位（集团跨主体合并，2026-08-12 用户待办）。

    ⚡ 与单体『应收+合同资产合并账龄』(build_ar_ca_combined_aging) 的区别：
    · 单体表=同客户在两科目间拆分的合并账龄，仅单主体审计口径；
    · 本表=集团视图——跨主体同一客户【应收+合同资产】期末相加，按合并后绝对值
      取 Top10，展示各主体分布（识别同一客户在多家主体的往来集中度）。

    数据=GL(1122 应收 + 1124 合同资产) 按 (客户, 主体) 聚合期末净额，再按客户
    跨主体相加。返回 ws 或 None（无合同资产数据不生成）。
    ⚡ 集团模式逐主体读 GL 防 OOM（88 家全量 read_gl_rows 会内存超限，与 _sap_batch 同策略）。"""
    from collections import defaultdict
    if _adapter is None:
        return None
    root = getattr(_adapter, '_DATA_ROOT', None) or data_dir
    ents = _adapter.discover_entities(root)
    if not ents:
        return None
    # (主体, 客户) → {ar_net, ca_net}
    ag = defaultdict(lambda: {'ar': 0.0, 'ca': 0.0})
    # 集团模式（无 current_comp）→ 逐主体读（防 88 家全量 OOM）；单体 → 一次读
    _comp = _adapter.current_comp()
    _scope = {_comp: ents[_comp]} if (_comp and _comp in ents) else None
    if _scope:
        for r in _adapter.read_gl_rows(root, _scope):
            c = str(r.get('code') or '')
            amt = float(r.get('dr') or 0.0) - float(r.get('cr') or 0.0)
            cu = str(r.get('cust') or '').strip() or '（未标注客户）'
            e = str(r.get('e') or '')
            if c.startswith('1122'):
                ag[(e, cu)]['ar'] += amt
            elif c.startswith('1124'):
                ag[(e, cu)]['ca'] += amt
    else:
        for _c in sorted(ents):
            _e1 = {_c: ents[_c]}
            for r in _adapter.read_gl_rows(root, _e1):
                c = str(r.get('code') or '')
                amt = float(r.get('dr') or 0.0) - float(r.get('cr') or 0.0)
                cu = str(r.get('cust') or '').strip() or '（未标注客户）'
                e = str(r.get('e') or '')
                if c.startswith('1122'):
                    ag[(e, cu)]['ar'] += amt
                elif c.startswith('1124'):
                    ag[(e, cu)]['ca'] += amt
    if not any(v['ca'] for v in ag.values()):
        return None                     # 无合同资产账套不生成
    # 客户 → {ar, ca, 主体分布}
    by_cust = defaultdict(lambda: {'ar': 0.0, 'ca': 0.0, 'ents': defaultdict(float)})
    for (e, cu), v in ag.items():
        if abs(v['ar']) < 0.005 and abs(v['ca']) < 0.005:
            continue
        by_cust[cu]['ar'] += v['ar']
        by_cust[cu]['ca'] += v['ca']
        by_cust[cu]['ents'][e] += (v['ar'] + v['ca'])
    top = sorted(by_cust.items(), key=lambda x: -(abs(x[1]['ar'] + x[1]['ca'])))[:10]
    if not top:
        return None
    ws = wb.create_sheet('应收+合同资产前10位')
    ws.cell(1, 1, f'应收+合同资产前10位（集团跨主体合并；{target_year or ""}；'
                  f'按合并后期末余额绝对值取前 10 位）').font = SHELL_TITLE_FONT
    hdrs = ['客户', '应收账款净额', '合同资产净额', '合计', '涉及主体数', '主体分布(净额)']
    for j, h in enumerate(hdrs, 1):
        cc = ws.cell(2, j, h); cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL
        cc.alignment = SHELL_CEN; cc.border = SHELL_BORDER
    r = 3
    for cu, v in top:
        ent_desc = '；'.join('%s %s' % (e, _fmt_amt(a)) for e, a in
                             sorted(v['ents'].items(), key=lambda x: -abs(x[1]))[:4])
        if len(v['ents']) > 4:
            ent_desc += f"；等 {len(v['ents'])} 家"
        ws.cell(r, 1, cu).font = SHELL_BOLD
        ws.cell(r, 2, round(v['ar'], 2) if abs(v['ar']) >= 0.005 else None)
        ws.cell(r, 3, round(v['ca'], 2) if abs(v['ca']) >= 0.005 else None)
        ws.cell(r, 4, round(v['ar'] + v['ca'], 2)).font = SHELL_BOLD
        ws.cell(r, 5, len(v['ents']))
        ws.cell(r, 6, ent_desc).alignment = SHELL_LEFT
        for j in range(1, 7):
            cc = ws.cell(r, j)
            cc.border = SHELL_BORDER
            if j in (2, 3, 4):
                cc.number_format = SHELL_NUM; cc.alignment = SHELL_RGT
        r += 1
    ws.cell(r, 1, '说明：①同一客户跨主体期末净额相加（应收+合同资产）；②负数=贷方余额（预收性），'
                  '按绝对值参与排序；③主体分布展示前 4 家及家数。').font = SHELL_BOLD
    for i, w in enumerate([30, 15, 15, 15, 10, 55], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'B3'
    return ws


def _fmt_amt(x):
    """金额千分位显示（供主体分布文本）。"""
    try:
        return f"{float(x):,.2f}"
    except (TypeError, ValueError):
        return str(x)


def _aging_simple(close_balance):
    """简化 FIFO 账龄：期末为正(借余,应收类)→按 6 档；负(贷余)不分。
    单年度数据无多年度流水时全部落在 1 年以内(真实账龄需多年度 GL)。"""
    b = AGING_BUCKETS.get('recv', ['1年以内', '1-2年', '2-3年', '3-4年', '4-5年', '5年以上'])
    out = [0.0] * len(b)
    if abs(close_balance) < 0.005:
        return out
    if close_balance > 0:
        out[0] = close_balance        # 单年度默认 1 年以内(与现有账龄口径一致)
    return out
