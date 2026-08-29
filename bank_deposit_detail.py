# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=银行存款审计底稿_生成.xlsx | 关键列=银行户期末余额/对方科目/异常 | 职责=银行存款明细与未达账·异常分析(独立生成器)
"""
通用银行存款明细表生成程序
================================================================================
功能概述
--------------------------------------------------------------------------------
本程序接收「银行存款交易数据」（交易级明细），自动完成：
  1. 读取数据（支持 .csv / .xlsx），自动识别列（中英文别名兼容）；
  2. 自动识别每条记录所属「会计期间」并按期间分组（年 / 月 / 季 可配）；
  3. 按 (账号, 币种, 期间) 汇总：期初余额、借方发生额、贷方发生额、期末余额；
  4. 同时适配「多期数据」与「单期数据」：
       - 单期：直接生成该期明细表；
       - 多期：按期间分别生成明细，并额外生成「期间对比」表；
  5. 数据校验：账户一致性（同账号名称唯一）、金额平衡（期初+借-贷=期末、
     跨期结转连续）、字段合法性；
  6. 输出结构化 Excel（多 Sheet）。

银行存款科目为资产类，约定：借 = 增加（存入），贷 = 减少（付出）。
  期末余额 = 期初余额 + 借方发生额 - 贷方发生额

使用方式
--------------------------------------------------------------------------------
  python bank_deposit_detail.py --input 交易数据.xlsx --output 银行存款明细表.xlsx
  # --input 也可直接传【文件夹】：自动识别其中的交易数据文件与期初余额文件
可选参数：
  --period  Y|M|Q      期间粒度，默认 Y（年）；M=月(2025-01)，Q=季(2025Q1)
  --opening 期初余额.csv  可选，提供各账户各币种首期期初余额，格式：
                          账号,币种,期初余额  （币种缺省视为人民币）
  --currency 列名       指定币种列名（数据含多币种时）

也可作为模块导入：from bank_deposit_detail import build_bank_detail
================================================================================
"""

import argparse
import csv
import datetime
import os
import re
import sys

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
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
# 统一外壳加载逻辑（六程序共享）：中央 audit_templates/ 定位与落盘
from audit_shell import resolve_template, AUDIT_TEMPLATES_DIR, finalize_workbook
from audit_common import _safe_save  # 共享库：统一锁感知保存（单一来源）

# ----------------------------------------------------------------------------
# 列名别名映射：标准字段 -> 可接受的中/英文别名（不区分大小写、忽略空格与下划线）
# ----------------------------------------------------------------------------
COLUMN_ALIASES = {
    "date":        ["日期", "记账日期", "交易日期", "会计日期", "date", "dt", "postdate"],
    "account":     ["账户名称", "户名", "银行账户名称", "account", "accountname", "acctname"],
    "account_no":  ["账号", "银行账号", "账户号", "accountno", "acctno", "accountnumber"],
    "direction":   ["借贷方向", "方向", "借贷", "direction", "dc", "drcr"],
    "amount":      ["金额", "发生额", "交易金额", "amount", "amt", "value"],
    "summary":     ["摘要", "备注", "说明", "summary", "memo", "note", "desc"],
    "currency":    ["币种", "货币", "currency", "ccy"],
}

# 借贷方向识别
DEBIT_TOKENS = {"借", "debit", "dr", "d", "+", "1", "收入", "存入"}
CREDIT_TOKENS = {"贷", "credit", "cr", "c", "-", "0", "付出", "支出"}

# Excel 样式（⚡ 2026-08-11 阶段二 2.1：统一引用 audit_shell 共享常量，消除重复定义）
from audit_shell import (SHELL_HFILL as HDR_FILL, SHELL_HFONT as HDR_FONT,
                         SHELL_TITLE_FONT as TITLE_FONT, SHELL_SUB_FONT as SUB_FONT,
                         SHELL_TOT_FILL as TOTAL_FILL, SHELL_BOLD as TOTAL_FONT,
                         SHELL_BORDER as BORDER, SHELL_NUM as MONEY_FMT,
                         SHELL_CEN as CENTER, SHELL_LEFT as LEFT)
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")


# ============================================================================
# 模块 1：数据读取
# ============================================================================
def _norm(s):
    """规范化列名：去空格、转小写、去下划线"""
    return re.sub(r"[\s_]+", "", str(s)).lower()


def _resolve_columns(headers):
    """将文件表头映射到标准字段，返回 {标准字段: 列索引}"""
    norm_headers = {_norm(h): i for i, h in enumerate(headers) if h}
    mapping = {}
    for std, aliases in COLUMN_ALIASES.items():
        for al in aliases:
            if _norm(al) in norm_headers:
                mapping[std] = norm_headers[_norm(al)]
                break
    return mapping


def read_transactions(path, warnings=None):
    """
    读取交易数据，返回 list[dict]，统一字段：
    date(datetime.date), account, account_no, direction('借'/'贷'),
    amount(float>0), summary, currency
    warnings: 可选 list，用于收集读取期的警告（如负金额）。
    """
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            raw = list(reader)
    else:  # xlsx
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    if not raw:
        raise ValueError("文件为空")
    headers = [str(h).strip() if h is not None else "" for h in raw[0]]
    mapping = _resolve_columns(headers)
    required = ["date", "account", "direction", "amount"]
    missing = [k for k in required if k not in mapping]
    if missing:
        raise ValueError(f"缺少必要列：{missing}（文件表头：{headers}）")

    for ri, row in enumerate(raw[1:], start=2):
        if not any(c is not None and str(c).strip() != "" for c in row):
            continue  # 跳过空行
        rec = {}

        # 日期
        dval = row[mapping["date"]] if mapping["date"] < len(row) else None
        rec["date"] = _parse_date(dval)
        if rec["date"] is None:
            raise ValueError(f"第{ri}行日期无法解析：{dval!r}")

        # 账户 / 账号
        rec["account"] = _str(row, mapping.get("account"), "")
        rec["account_no"] = _str(row, mapping.get("account_no"), "")
        if not rec["account_no"] and not rec["account"]:
            rec["account_no"] = f"未知账户{ri}"
        if not rec["account"]:
            rec["account"] = rec["account_no"]

        # 借贷方向
        dval = row[mapping["direction"]] if mapping["direction"] < len(row) else None
        rec["direction"] = _parse_direction(dval, ri)

        # 金额
        aval = row[mapping["amount"]] if mapping["amount"] < len(row) else None
        rec["amount"] = _parse_amount(aval, ri)
        # 负金额提示：方向已单列，原值负通常意味着数据录入/方向错误
        if _is_negative_raw(aval):
            msg = (f"第{ri}行金额为负（{aval!r}），已按绝对值计入『{rec['direction']}』"
                   f"方向；请核实该笔借贷方向是否正确。")
            if warnings is not None:
                warnings.append(msg)

        # 摘要 / 币种
        rec["summary"] = _str(row, mapping.get("summary"), "")
        rec["currency"] = _str(row, mapping.get("currency"), "人民币") or "人民币"

        rows.append(rec)
    return rows


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


def _parse_direction(v, ri):
    s = _str([v], 0, "").lower()
    if s in DEBIT_TOKENS:
        return "借"
    if s in CREDIT_TOKENS:
        return "贷"
    raise ValueError(f"第{ri}行借贷方向无法识别：{v!r}")


def _parse_amount(v, ri):
    if isinstance(v, (int, float)):
        a = float(v)
    else:
        s = str(v).strip()
        neg = s.startswith("(") and s.endswith(")")
        s = s.replace(",", "").replace("￥", "").replace("¥", "")
        s = re.sub(r"[()]", "", s) if neg else s
        try:
            a = float(s)
        except ValueError:
            raise ValueError(f"第{ri}行金额无法解析：{v!r}")
        if neg:
            a = -a
    if a == 0:
        return 0.0
    # 方向已单独给出，金额取绝对值；借贷符号由 direction 决定增减
    return abs(a)


def read_opening_balances(path):
    """读取期初余额表 -> ({(账号, 币种): 金额}, conflicts)
    conflicts: 同(账号,币种)出现不一致金额时的冲突列表 [(no,cur,a1,a2), ...]
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            raw = list(csv.reader(f))
    else:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    res = {}
    conflicts = []
    if not raw:
        return res, conflicts
    headers = [str(h).strip() if h is not None else "" for h in raw[0]]
    def idx(name):
        for i, h in enumerate(headers):
            if _norm(h) == _norm(name):
                return i
        return None
    ci, ki, vi = idx("账号"), idx("币种"), idx("期初余额")
    if ci is None or vi is None:
        return res, conflicts
    seen = {}
    for ri, row in enumerate(raw[1:], start=2):
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        no = _str(row, ci, "")
        cur = _str(row, ki, "人民币") or "人民币"
        amt = _parse_amount(row[vi] if vi < len(row) else None, ri)
        if (no, cur) in seen and abs(seen[(no, cur)] - amt) > 0.005:
            conflicts.append((no, cur, seen[(no, cur)], amt))
        seen[(no, cur)] = amt
        res[(no, cur)] = amt
    return res, conflicts


def _is_negative_raw(v):
    """判断原始金额值是否为负（用于负金额提示）"""
    if isinstance(v, (int, float)):
        return v < 0
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace("￥", "").replace("¥", "")
        if s.startswith("(") and s.endswith(")"):
            return True
        return s.startswith("-")
    return False


# 银行存款(资产类)科目识别：覆盖全称与常见简称。
# 仅当『科目名称』以银行机构名开头才判定为银行存款账户，避免误收
# 『财务费用-利息收入-银行存款活期利息』『银行承兑汇票』等。
# ⚡⚡ 2026-08-15 配置化：银行存款科目码（标准 1002 + 本地码 1113 泰国），
#   account_profiles.json subject_codes.bank_deposit 可覆盖（如 XBJ=1002+1004 内部结算中心存款）。
_BANK_DEPOSIT_CODES = ('1002', '1113')


def _set_bank_codes(acct):
    """按账套注入银行存款科目码（默认 1002+1113；XBJ 等加 1004 本地码）。"""
    global _BANK_DEPOSIT_CODES
    try:
        import subject_mapping as _sm
        codes = _sm.resolve_subject_codes(acct, 'bank_deposit', list(_BANK_DEPOSIT_CODES))
        _BANK_DEPOSIT_CODES = tuple(codes) or ('1002', '1113')
    except Exception:
        pass


def _acct_of(data_dir):
    """从数据目录向上找账套名（XBJ/prepared → XBJ；账套根 → 自身）。"""
    import paths as _P
    p = os.path.abspath(data_dir)
    while True:
        b = os.path.basename(p)
        if b in _P.ACCT_NAMES:
            return b
        parent = os.path.dirname(p)
        if parent == p:
            return b
        p = parent
BANK_SUBJECT_KEYWORDS = (
    # 通用科目名（非具体银行名，如集团级 GL 只列"银行存款"父级）
    "银行存款",
    # 全称
    "杭州银行", "农村合作银行", "浙商银行", "华夏银行", "上海浦东发展银行", "浦发银行",
    "平安银行", "中国光大银行", "光大银行", "民生银行", "温州银行", "稠州银行",
    "兴业银行", "中信银行", "台州银行", "宁波银行", "建设银行", "农业银行",
    "中国银行", "工商银行", "交通银行", "招商银行", "邮储银行", "邮政储蓄银行",
    "北京银行", "上海银行", "南京银行", "江苏银行", "渤海银行", "广发银行",
    "恒丰银行", "农村商业银行", "农村信用社", "微商银行", "厦门银行", "天津银行",
    "贵阳银行", "郑州银行", "长沙银行", "成都银行", "重庆银行", "青岛银行",
    "齐鲁银行", "徽商银行", "东莞银行", "汉口银行", "贵阳银行",
    # 简称（多不含『银行』二字）
    "建行", "农行", "中行", "工行", "交行", "招行", "浦发", "光大", "邮储",
    "农商行", "农商", "农信", "宁波分行", "宁波银行",
)


def _is_bank_subject(name):
    """判断 GL 科目名称是否为银行存款(1002)资产账户。

    采用『关键词出现在名称任意位置』匹配（而非仅开头），以兼容
    『路桥工行美元户』这类带前缀的银行网点名；同时排除损益类利息、
    票据、其他货币资金等非银行存款科目。
    """
    if name is None:
        return False
    s = str(name).strip()
    if not s:
        return False
    # 损益类利息科目不算资金账户
    if s.startswith("财务费用") or "利息收入" in s:
        return False
    # 其他货币资金(支付宝/微信/票据保证金等)不算银行存款
    if "其他货币资金" in s or "票据" in s or "承兑汇票" in s:
        return False
    for kw in BANK_SUBJECT_KEYWORDS:
        if kw in s:
            return True
    return False


def _norm_key(name):
    """归一化账户名：去空白、转小写，便于科目余额表↔GL 名称匹配。"""
    return re.sub(r"\s+", "", str(name).strip()).lower()


def _norm_key_strip_digits(name):
    """归一化并去除尾部账号数字，作为名称匹配的兜底键。"""
    return re.sub(r"\d+$", "", _norm_key(name))


def read_gl_bank(path, warnings=None):
    """读取『综合查询明细表(GL)』，抽取银行存款(1002)分录 -> 内部统一行格式。

    返回 list[dict]，字段同 read_transactions：
    date, account, account_no, direction, amount, summary, currency。
    银行存款在 GL 中以具体银行网点名(如『杭州银行台州分行503437』)作为科目名称，
    本函数按 _is_bank_subject 过滤，并将双栏(借方金额/贷方金额)映射为 借贷方向+金额。
    """
    # SAP：从 adapter 干净行聚合银行分录（name=银行账户名，映射借贷方向）
    # ⚡ 2026-08-11 P0 修复：无参 _adapter.read_gl() 依赖 current_comp，集团模式=None →
    #   空。从 path(list[0] 含 {comp}) 推断主体显式传 comp（_scope_comp 仅认 str）。
    if _adapter is not None and (path is None or not isinstance(path, str)):
        _c_comp = None
        if _adapter.current_comp():
            _c_comp = _adapter.current_comp()
        else:
            p0 = path[0] if isinstance(path, list) and path else path
            _m = re.search(r'[\\/](\d{4})[\\/]', str(p0))
            if _m:
                _c_comp = _m.group(1)
            else:
                _m2 = re.match(r'^(\d{4})', os.path.basename(str(p0)))
                if _m2:
                    _c_comp = _m2.group(1)
        out = []
        for r in _adapter.read_gl(_c_comp):
            nm = str(r.get('name') or '')
            if not _is_bank_subject(nm):
                continue
            jf = float(r.get('debit') or 0.0)
            df = float(r.get('credit') or 0.0)
            # ⚡ 2026-08-12 修复：透传 vno/vtype/year——对方科目推导段(2173-2219)
            #   依赖 _r['vchar']/_r['vno']/_r['entity']/_r['year'] 匹配 build_gl_voucher_map，
            #   缺字段 → 推导 0 条 → 对方科目核对/销售精确汇总/采购精确汇总/剩余 4 表空
            #   （质检 6 主体 ERROR）。vchar 兼容 U8 '字-号' 形态。
            _vno = str(r.get('vno') or '')
            _vt = str(r.get('vtype') or '')
            _vchar = ('%s-%s' % (_vt, _vno)).strip('-') or _vno
            _y = str(r.get('year') or r.get('y') or '')
            if jf:
                out.append({'date': r.get('date'), 'account': nm, 'account_no': '',
                            'direction': '借', 'amount': jf, 'summary': r.get('sm'),
                            'currency': '', 'vno': _vno, 'vtype': _vt, 'vchar': _vchar,
                            'year': _y})
            if df:
                out.append({'date': r.get('date'), 'account': nm, 'account_no': '',
                            'direction': '贷', 'amount': df, 'summary': r.get('sm'),
                            'currency': '', 'vno': _vno, 'vtype': _vt, 'vchar': _vchar,
                            'year': _y})
        return out
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
        return []
    # 表头探测：含『科目名称』且含(日期 or 借方金额/贷方金额)
    hdr_idx, hdr = 0, [str(h).strip() if h is not None else "" for h in raw[0]]
    for i, row in enumerate(raw):
        cells = [str(c) for c in row if c is not None]
        joined = " ".join(cells)
        if "科目名称" in joined and ("日期" in joined or "借方金额" in joined or "贷方金额" in joined):
            hdr_idx, hdr = i, [str(h).strip() if h is not None else "" for h in row]
            break

    def cidx(*names):
        for n in names:
            for j, h in enumerate(hdr):
                if _norm(h) == _norm(n):
                    return j
        return None

    ci = cidx("科目名称", "科目")
    di = cidx("日期")
    wi = cidx("字")
    ni = cidx("号")
    pi = cidx("对方科目", "对方科目名称")
    si = cidx("摘要")
    ji = cidx("借方金额", "借方", "借方发生额")
    fi = cidx("贷方金额", "贷方", "贷方发生额")
    ai = cidx("辅助核算名称")
    if ci is None or ji is None or fi is None or di is None:
        return []

    rows = []
    red_debit = red_credit = 0.0
    for ri, row in enumerate(raw[hdr_idx + 1:], start=hdr_idx + 2):
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        name = _str(row, ci, "")
        if not _is_bank_subject(name):
            continue
        dval = row[di] if di < len(row) else None
        dt = _parse_date(dval)
        if dt is None:
            continue
        debit = _signed_amt(row[ji]) if ji < len(row) else 0.0
        credit = _signed_amt(row[fi]) if fi < len(row) else 0.0
        # 红字(负数)净额：借/贷列可能为负，属正常冲减
        if debit < 0:
            red_debit += debit
        if credit < 0:
            red_credit += credit
        # 携带有符号借/贷列值，便于与科目余额表(1002)『本期借/贷』(红字作负贡献)口径一致地勾稽；
        # 显示用 direction/amount 同样保留红字负号（贷-红字 = 红字冲减），避免把红字贷方误转成借方
        # 而虚增『双借双贷』(Δ借=Δ贷) 的伪差异。期末余额(期初+借-贷)两种表示等价，不受影响。
        sdebit, scredit = debit, credit
        if debit != 0:
            direction, amount = "借", debit
        elif credit != 0:
            direction, amount = "贷", credit
        else:
            continue
        # 凭证号
        vchar = _str(row, wi, "")
        vno = _str(row, ni, "")
        voucher = f"{vchar}-{vno}" if (vchar or vno) else ""
        # 摘要 + 对方科目
        summ = _str(row, si, "")
        opp = _str(row, pi, "")
        if opp and opp not in summ:
            summ = (summ + "｜对方:" + opp).strip("｜")
        # 币种：账户名含外币关键字
        currency = "人民币"
        if "美元" in name:
            currency = "美元"
        elif "欧元" in name:
            currency = "欧元"
        rows.append({
            "date": dt,
            "account": name,
            "account_no": name,  # 以账户名作为唯一键，便于与科目余额表匹配
            "direction": direction,
            "amount": amount,
            "sdebit": sdebit,    # 有符号借方列值(GL 红字借方为负)
            "scredit": scredit,  # 有符号贷方列值(GL 红字贷方为负)
            "summary": summ,
            "currency": currency,
            "opp": opp,          # 对方科目（独立字段，供对方科目核对 sheet 使用）
            "vchar": vchar,      # 凭证字（供异常 sheet 凭证抽取）
            "vno": vno,          # 凭证号（供异常 sheet 凭证抽取）
        })
    if red_debit < -0.005 or red_credit < -0.005:
        if warnings is not None:
            warnings.append(
                f"GL 红字冲减：借方合计 {red_debit:.2f}，贷方合计 {red_credit:.2f}（属正常冲减，已按净额计入）。")
    return rows


def _safe_amt(v):
    try:
        return _parse_amount(v, 0)
    except (ValueError, TypeError):
        return 0.0


def _signed_amt(v):
    """保留正负号的金额解析（用于 GL 双栏：红字冲减以负数表示）。"""
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace(",", "").replace("￥", "").replace("¥", "")
        neg = s.startswith("(") and s.endswith(")")
        s = re.sub(r"[()]", "", s) if neg else s
        try:
            a = float(s)
        except ValueError:
            return 0.0
        return -a if neg else a
    return 0.0


def read_km_bank_openings(path):
    """读取科目余额表，提取银行存款(1002)各子目期初余额。

    返回 dict：key = 归一化账户名(_norm_key)，value = (期初金额(带借贷符号), 币种, 显示名)。
    同时写入『去尾部数字』兜底键，缓解 GL 与科目余额表账号位数不一致导致的匹配失败。
    """
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
        return {}
    code_aliases = set(_norm(a) for a in KM_COLUMN_ALIASES["code"])
    amt_aliases = set(_norm(a) for a in KM_COLUMN_ALIASES["open_amt"]
                      + KM_COLUMN_ALIASES["close_amt"])
    hdr_idx, hdr = 0, [str(h).strip() if h is not None else "" for h in raw[0]]
    for i, row in enumerate(raw):
        norm_cells = [_norm(c) for c in row if c is not None]
        has_amt = any(nc in amt_aliases for nc in norm_cells)
        has_dc = any(_norm(h) in ("本期借方", "本期贷方", "借方", "贷方") for h in row if h is not None)
        has_dir = any(_norm(h) == "方向" for h in row if h is not None)
        if any(nc in code_aliases for nc in norm_cells) and (has_amt or has_dc or has_dir):
            hdr_idx, hdr = i, [str(h).strip() if h is not None else "" for h in row]
            break

    def cidx(std):
        targets = set(_norm(a) for a in KM_COLUMN_ALIASES[std])
        for i, h in enumerate(hdr):
            if _norm(h) in targets:
                return i
        return None

    ci = cidx("code")
    ni = cidx("name")
    # 期初余额列：优先别名；否则取『方向』列后紧邻的『金额』列(兼容表头仅写"金额"的布局)
    dirs = [i for i, h in enumerate(hdr) if _norm(h) == "方向"]
    oai = cidx("open_amt")
    if oai is None and dirs:
        oai = dirs[0] + 1 if dirs[0] + 1 < len(hdr) else None
    odi = cidx("open_dir") or (dirs[0] if dirs else None)
    if ci is None or ni is None or oai is None:
        return {}
    res = {}
    for row in raw[hdr_idx + 1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        if ci >= len(row) or row[ci] is None:
            continue
        code = _norm(row[ci]).replace(".0", "")
        # 仅取 1002 子目（1002*），排除『1002』父行（父行=全部子目汇总数，非真实账户，
        # 若计入期初会导致控制数期初被整体重复计列）。
        if code == "1002" or not code.startswith("1002"):
            continue
        name = _str(row, ni, "")
        if not name:
            continue
        open_amt = _safe_amt(row[oai] if oai < len(row) else 0)
        od = _str(row, odi, "借")
        signed = open_amt if (_norm(od) in ("借", "debit", "dr", "j") or _norm(od) == "") else -open_amt
        currency = "人民币"
        if "美元" in name:
            currency = "美元"
        elif "欧元" in name:
            currency = "欧元"
        k1 = _norm_key(name)
        res[k1] = (signed, currency, name)
        res[_norm_key_strip_digits(name)] = (signed, currency, name)
    return res


def _build_opening_from_km(rows, km_path):
    """根据 GL 抽取出的银行存款账户，从科目余额表(1002)推导各账户期初余额。

    返回 (opening_dict, warns)。
      opening_dict: {(账户名, 币种): 期初金额}
      warns: 未匹配到账套科目余额表的账户提示列表
    """
    opening_dict, warns = {}, []
    if not km_path:
        return opening_dict, warns
    km = read_km_bank_openings(km_path)
    if not km:
        return opening_dict, warns
    seen = set()
    for r in rows:
        ak = _account_key(r)
        if ak in seen:
            continue
        seen.add(ak)
        no, cur = ak
        k1 = _norm_key(no)
        k2 = _norm_key_strip_digits(no)
        hit = km.get(k1) or km.get(k2)
        if hit is not None:
            amt, ccur, _ = hit
            if ccur != cur:
                warns.append(f"账户『{no}』币种不一致：GL 为{cur}，科目余额表为{ccur}，按期初{cur}计。")
            opening_dict[ak] = amt
        else:
            opening_dict[ak] = 0.0
            warns.append(f"账户『{no}』({cur})未在科目余额表1002中找到期初，按 0 计，请核对。")
    # 补充：科目余额表1002全部子目（含 GL 无交易的休眠账户），保证期初与控制数(1002)口径一致，
    # 避免休眠账户期初被漏计导致控制台账期初/期末差异（如 农行路桥横街支行03061-1）。
    covered = {_norm_key(no) for (no, _cur) in opening_dict}
    for (amt, ccur, name) in km.values():
        kn = _norm_key(name)
        if kn in covered:
            continue
        opening_dict[(name, ccur)] = amt
        covered.add(kn)
    return opening_dict, warns


def _build_km_year_map(folder, exclude_name=None):
    """扫描文件夹，返回 {年份: 主科目余额表路径}，用于按年核对 银行存款(1002) 控制数。

    仅取『主科目余额表/trial/总账』，排除『外币科目余额表』『辅助核算』『明细表』，
    避免把外币专用余额表(仅外币子目)误当 1002 控制数，或与主表重复计列。
    """
    res = {}
    if not os.path.isdir(folder):
        return res
    for fn in sorted(os.listdir(folder)):
        if not fn.lower().endswith((".xlsx", ".csv")):
            continue
        if exclude_name and fn == exclude_name:
            continue
        if "外币" in fn or "辅助核算" in fn or "明细" in fn:
            continue
        if not any(k in fn for k in ("科目余额表", "trial", "总账", "科目余额")):
            continue
        yrs = re.findall(r"20\d{2}", fn)
        if not yrs:
            continue
        yr = yrs[-1]
        if yr not in res:
            res[yr] = os.path.join(folder, fn)
    return res


def _build_opening_from_km_multi(rows, km_by_year):
    """按各账户最早出现年份，从对应年份主科目余额表(1002)取期初；并补充全部 1002 子目(含休眠账户)。

    返回 (opening_dict, warns)。相比 _build_opening_from_km(单表)，本函数支持多年 GL：
    首期(最早年)期初取自该年主科目余额表，后续期由余额表自身跨期结转；并对各年 1002 全部子目
    （含 GL 无交易的休眠账户，如 农行路桥横街支行03061-1）补列期初，保证期初/期末与控制数口径一致。
    """
    opening_dict, warns = {}, []
    if not km_by_year:
        return opening_dict, warns
    # 每个账户最早出现年份
    earliest = {}
    for r in rows:
        ak = _account_key(r)
        yr = str(r["date"].year)
        if ak not in earliest or yr < earliest[ak]:
            earliest[ak] = yr
    km_cache = {}
    for yr, p in km_by_year.items():
        km_cache[yr] = read_km_bank_openings(p)
    for ak, yr in earliest.items():
        no, cur = ak
        km = km_cache.get(yr)
        if not km:
            opening_dict[ak] = 0.0
            warns.append(f"账户『{no}』({cur}) 年份{yr}未找到对应科目余额表，期初按 0 计。")
            continue
        hit = km.get(_norm_key(no)) or km.get(_norm_key_strip_digits(no))
        if hit is not None:
            amt, ccur, _ = hit
            if ccur != cur:
                warns.append(f"账户『{no}』币种不一致：GL 为{cur}，科目余额表为{ccur}，按期初{cur}计。")
            opening_dict[ak] = amt
        else:
            opening_dict[ak] = 0.0
            warns.append(f"账户『{no}』({cur})年份{yr}未在科目余额表1002中找到期初，按 0 计，请核对。")
    # 补充：各年 KM 全部 1002 子目（含 GL 无交易的休眠账户），保证期初与控制数(1002)口径一致
    covered = {_norm_key(no) for (no, _c) in opening_dict}
    for yr, km in km_cache.items():
        for (amt, ccur, name) in km.values():
            kn = _norm_key(name)
            if kn in covered:
                continue
            opening_dict[(name, ccur)] = amt
            covered.add(kn)
    return opening_dict, warns



# ============================================================================
# 模块 2：期间识别与分组汇总
# ============================================================================
def period_key(date, mode="Y"):
    """从日期得到期间键"""
    if mode == "M":
        return date.strftime("%Y-%m")
    if mode == "Q":
        q = (date.month - 1) // 3 + 1
        return f"{date.year}Q{q}"
    return str(date.year)


def _account_key(rec):
    """账户唯一键：(账号, 币种)；账号为空则用账户名"""
    no = rec["account_no"] or rec["account"]
    return (no, rec["currency"])


def aggregate(rows, mode="Y"):
    """
    按 (账号, 币种, 期间) 聚合。
    返回 periods(set), accounts(有序列表 of (no,cur)),
    agg = {(no,cur,period): dict(debit,credit,count)}
    """
    agg = {}
    periods = set()
    accounts = []
    seen = set()
    for r in rows:
        pk = period_key(r["date"], mode)
        ak = _account_key(r)
        periods.add(pk)
        if ak not in seen:
            seen.add(ak)
            accounts.append(ak)
        d = agg.setdefault((ak[0], ak[1], pk), {"debit": 0.0, "credit": 0.0, "count": 0})
        if "sdebit" in r and "scredit" in r:
            # GL 模式：按有符号借/贷列值汇总，与科目余额表(1002)『本期借/贷』(红字作负贡献)口径一致
            d["debit"] += r["sdebit"]
            d["credit"] += r["scredit"]
        elif r["direction"] == "借":
            d["debit"] += r["amount"]
        else:
            d["credit"] += r["amount"]
        d["count"] += 1
    return sorted(periods), accounts, agg


def compute_summary(accounts, periods, agg, opening=None):
    """
    计算每账户每期 期初/借/贷/期末。
    期初规则：首期 = 期初余额表提供值（缺省 0）；后续期 = 上期期末（结转）。
    返回 summary[(no,cur,period)] = dict(open,debit,credit,close,count)
    """
    opening = opening or {}
    summary = {}
    for ak in accounts:
        no, cur = ak
        prev_close = None
        for pk in periods:
            rec = agg.get((no, cur, pk), {"debit": 0.0, "credit": 0.0, "count": 0})
            if prev_close is None:
                op = opening.get((no, cur), 0.0)
            else:
                op = prev_close
            close = op + rec["debit"] - rec["credit"]
            summary[(no, cur, pk)] = {
                "open": op, "debit": rec["debit"],
                "credit": rec["credit"], "close": close, "count": rec["count"],
            }
            prev_close = close
    return summary


# ============================================================================
# 模块 3：数据校验
# ============================================================================
def validate(rows, accounts, periods, agg, summary, opening=None):
    """
    返回 issues：list of dict(level, type, message)
    level: 'ERROR' / 'WARN'
    """
    issues = []
    opening = opening or {}

    # (1) 账户一致性：同账号 -> 账户名称集合
    name_by_no = {}
    for r in rows:
        no = r["account_no"] or r["account"]
        nm = r["account"]
        name_by_no.setdefault(no, set()).add(nm)
    for no, names in name_by_no.items():
        if len(names) > 1:
            issues.append({
                "level": "WARN", "type": "账户一致性",
                "message": f"账号『{no}』对应多个账户名称：{sorted(names)}，已按行各自名称处理。",
            })

    # (2) 金额平衡（期末 = 期初+借-贷）与跨期结转连续性
    for ak in accounts:
        no, cur = ak
        prev_close = None
        for pk in periods:
            s = summary[(no, cur, pk)]
            calc_close = s["open"] + s["debit"] - s["credit"]
            if abs(calc_close - s["close"]) > 0.005:
                issues.append({
                    "level": "ERROR", "type": "金额平衡",
                    "message": f"账户『{no}』({cur}) {pk}：期初+借-贷({calc_close:.2f})≠期末({s['close']:.2f})",
                })
            if prev_close is not None and abs(prev_close - s["open"]) > 0.005:
                issues.append({
                    "level": "WARN", "type": "跨期结转",
                    "message": f"账户『{no}』({cur}) {pk} 期初({s['open']:.2f})≠上期期末({prev_close:.2f})，结转不连续。",
                })
            prev_close = s["close"]

    # (3) 期初余额表与首期结转冲突（若提供期初且账户在多期出现，首期期初应等于提供值）
    for ak in accounts:
        no, cur = ak
        first_pk = periods[0]
        s = summary[(no, cur, first_pk)]
        if (no, cur) in opening and abs(opening[(no, cur)] - s["open"]) > 0.005:
            issues.append({
                "level": "WARN", "type": "期初余额",
                "message": f"账户『{no}』({cur}) 首期{first_pk}期初({s['open']:.2f})与期初余额表({opening[(no,cur)]:.2f})不一致。",
            })

    # (4) 字段合法性：负金额已在读取时取绝对值；此处检查极端（金额=0 提示）
    zero_cnt = sum(1 for r in rows if r["amount"] == 0)
    if zero_cnt:
        issues.append({"level": "WARN", "type": "金额", "message": f"存在 {zero_cnt} 条金额为 0 的记录，已按 0 处理。"})

    if not issues:
        issues.append({"level": "OK", "type": "汇总", "message": "校验通过：账户一致、金额平衡、跨期结转连续。"})
    return issues


# ============================================================================
# 模块 4：明细表生成 + 模块 5：导出
# ============================================================================
def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def _write_period_sheet(ws, period, accounts, summary, cur_label_map):
    """写入单个期间明细表（⚡ 2026-08-11 阶段二 2.1：计算产结构化行 → render_into 渲染）。
    布局规范化：R1 标题 / R2 表头 / R3 起数据（原 R2/R4/R5，值不变）。"""
    from audit_render import render_into, row as _arow
    headers = ["序号", "账户名称", "账号", "币种", "期初余额", "借方发生额",
               "贷方发生额", "期末余额", "交易笔数", "平衡校验(期初+借-贷-期末)"]
    MONEY = {5, 6, 7, 8, 10}
    name_map = cur_label_map
    _rows = []
    tot = {"open": 0.0, "debit": 0.0, "credit": 0.0, "close": 0.0, "count": 0}
    for idx, ak in enumerate(accounts, 1):
        no, cur = ak
        s = summary[(no, cur, period)]
        bal = s["open"] + s["debit"] - s["credit"] - s["close"]
        acct_name = name_map.get(ak, no)
        _rows.append(_arow([idx, acct_name, no, cur, s["open"], s["debit"],
                            s["credit"], s["close"], s["count"], round(bal, 2)],
                           num=MONEY,
                           align='c' if False else None,
                           cell_fill=({10: WARN_FILL} if abs(bal) > 0.005 else None)))
        tot["open"] += s["open"]; tot["debit"] += s["debit"]
        tot["credit"] += s["credit"]; tot["close"] += s["close"]
        tot["count"] += s["count"]
    # 合计行
    _rows.append(_arow(["合计", None, None, None, tot["open"], tot["debit"],
                        tot["credit"], tot["close"], tot["count"], None],
                       b=True, fill='sub', num=MONEY))
    render_into(ws, headers, _rows, MONEY, title=f"银行存款明细表 — {period}")
    # 列宽
    widths = [6, 28, 22, 10, 16, 16, 16, 16, 10, 22]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_comparison_sheet(ws, accounts, periods, summary, name_map):
    """期间对比：每账户一行，各期 期初/借/贷/期末 + 期末变动"""
    ws.cell(2, 1, "银行存款明细 — 期间对比").font = TITLE_FONT
    # 表头
    hdr = ["账户名称", "账号", "币种"]
    for pk in periods:
        hdr += [f"{pk}期初", f"{pk}借方", f"{pk}贷方", f"{pk}期末"]
    hdr += ["首期期末", "末期期末", "期末变动"]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for ak in accounts:
        no, cur = ak
        first_close = summary[(no, cur, periods[0])]["close"]
        last_close = summary[(no, cur, periods[-1])]["close"]
        row_vals = [name_map.get(ak, no), no, cur]
        for pk in periods:
            s = summary[(no, cur, pk)]
            row_vals += [s["open"], s["debit"], s["credit"], s["close"]]
        row_vals += [first_close, last_close, last_close - first_close]
        for c, v in enumerate(row_vals, 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c >= 4:
                cell.number_format = MONEY_FMT
                cell.alignment = Alignment(horizontal="right")
        r += 1
    # 合计行
    ws.cell(r, 1, "合计").font = TOTAL_FONT
    for c in range(1, len(hdr) + 1):
        ws.cell(r, c).border = BORDER
        ws.cell(r, c).fill = TOTAL_FILL
    col = 4
    for pk in periods:
        so = sum(summary[(a[0], a[1], pk)]["open"] for a in accounts)
        sd = sum(summary[(a[0], a[1], pk)]["debit"] for a in accounts)
        sc = sum(summary[(a[0], a[1], pk)]["credit"] for a in accounts)
        cl = sum(summary[(a[0], a[1], pk)]["close"] for a in accounts)
        for off, val in enumerate([so, sd, sc, cl]):
            cell = ws.cell(r, col + off, val)
            cell.font = TOTAL_FONT; cell.fill = TOTAL_FILL; cell.border = BORDER
            cell.number_format = MONEY_FMT
        col += 4
    widths = [28, 22, 10] + [14] * (len(periods) * 4) + [14, 14, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_summary_sheet(ws, periods, accounts, summary):
    """数据汇总：各期总体 期初/借/贷/期末"""
    ws.cell(2, 1, "银行存款余额汇总").font = TITLE_FONT
    hdr = ["期间", "账户数", "期初余额合计", "借方发生额合计", "贷方发生额合计", "期末余额合计"]
    for i, h in enumerate(hdr, 1):
        ws.cell(4, i, h)
    _style_header(ws, 4, len(hdr))
    r = 5
    for pk in periods:
        op = sum(summary[(a[0], a[1], pk)]["open"] for a in accounts)
        db = sum(summary[(a[0], a[1], pk)]["debit"] for a in accounts)
        cr = sum(summary[(a[0], a[1], pk)]["credit"] for a in accounts)
        cl = sum(summary[(a[0], a[1], pk)]["close"] for a in accounts)
        for c, v in enumerate([pk, len(accounts), op, db, cr, cl], 1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c >= 3:
                cell.number_format = MONEY_FMT
        r += 1
    for i, w in enumerate([12, 10, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_validation_sheet(ws, issues, rows, accounts, periods):
    ws.cell(2, 1, "数据校验结果").font = TITLE_FONT
    info = [f"交易记录数：{len(rows)}", f"账户数：{len(accounts)}", f"期间数：{len(periods)}（{', '.join(periods)}）"]
    for i, t in enumerate(info):
        ws.cell(4 + i, 1, t)
    hdr = ["序号", "级别", "类型", "说明"]
    hr = 8
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
        r += 1
    for i, w in enumerate([6, 10, 14, 90], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def export_excel(output, periods, accounts, summary, issues, rows, name_map, mode="Y"):
    """导出多 Sheet Excel"""
    wb = openpyxl.Workbook()
    multi = len(periods) > 1
    # 期间明细 sheets
    for pk in periods:
        title = f"银行存款明细表_{pk}"
        ws = wb.create_sheet(title)
        _write_period_sheet(ws, pk, accounts, summary, name_map)
    if multi:
        ws = wb.create_sheet("期间对比")
        _write_comparison_sheet(ws, accounts, periods, summary, name_map)
    ws = wb.create_sheet("数据汇总")
    _write_summary_sheet(ws, periods, accounts, summary)
    ws = wb.create_sheet("校验结果")
    _write_validation_sheet(ws, issues, rows, accounts, periods)
    # 删除默认空 sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    # 调整 sheet 顺序：汇总在前更易读
    order = ([f"明细_{pk}" for pk in periods] + (["期间对比"] if multi else []) +
             ["数据汇总", "校验结果"])
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)
    finalize_workbook(wb)
    wb.save(output)


def _build_name_map(rows):
    """(账号,币种) -> 账户名称（取首次出现）"""
    m = {}
    for r in rows:
        ak = _account_key(r)
        if ak not in m:
            m[ak] = r["account"]
    return m


# ============================================================================
# 主入口
# ============================================================================
def _discover_inputs(folder, exclude_name=None):
    """扫描输入文件夹，自动识别『交易数据文件』『期初余额文件』『科目余额表(控制数)』『综合查询明细表(GL)』。

    识别规则（按文件名关键字，不区分大小写）：
      - 期初余额文件：文件名含『期初』或『opening』；
      - 科目余额表(控制数)：文件名含『科目余额表/trial/总账/科目余额』且不含『明细』；
      - 综合查询明细表(GL)：文件名含『综合查询』；
      - 以下文件一律视为『非交易文件』，绝不当作交易数据：
        科目余额表 / 辅助核算 / 综合查询 / 余额表 / 总账 / trial / 控制数 / 明细表
      - 交易数据文件：非上述文件，且文件名含『银行/存款/流水/交易/日记账』之一；
      - 若未命中关键字，则退化为：文件夹内除期初/科目余额表/非交易/输出文件外的任意 .xlsx/.csv。
    返回 (txn_files: list, opening_file: str|None, km_file: str|None, gl_files: list)。
    """
    EXCLUDE = ("科目余额表", "辅助核算", "综合查询", "余额表", "总账",
               "trial", "控制数", "明细表")
    TXN_KW = ("银行", "存款", "流水", "交易", "日记账")
    GL_KW = ("综合查询",)
    txn, openf, kmf, gl = [], None, None, []
    files = sorted(os.listdir(folder))
    for fn in files:
        if not fn.lower().endswith((".xlsx", ".csv")):
            continue
        if exclude_name and fn == exclude_name:
            continue
        if "期初" in fn or "opening" in fn.lower():
            if openf is None:
                openf = os.path.join(folder, fn)
            continue
        if any(k in fn for k in ("科目余额表", "trial", "总账", "科目余额")) and "明细" not in fn:
            if kmf is None:
                kmf = os.path.join(folder, fn)
            continue
        if any(k in fn for k in GL_KW):
            gl.append(os.path.join(folder, fn))
            continue
        # 明确排除 GL/余额类文件，绝不当作交易文件
        if any(k in fn for k in EXCLUDE):
            continue
        if any(k in fn for k in TXN_KW):
            txn.append(os.path.join(folder, fn))
    if not txn:
        for fn in files:
            if not fn.lower().endswith((".xlsx", ".csv")):
                continue
            if exclude_name and fn == exclude_name:
                continue
            if openf and os.path.join(folder, fn) == openf:
                continue
            if kmf and os.path.join(folder, fn) == kmf:
                continue
            if any(k in fn for k in EXCLUDE):
                continue
            if any(k in fn for k in GL_KW):
                continue
            txn.append(os.path.join(folder, fn))
    return txn, openf, kmf, gl


# 科目余额表列别名（用于提取银行存款 1002 控制数）
KM_COLUMN_ALIASES = {
    "code": ["科目代码", "代码", "科目编码", "科目代号", "科目编号"],
    "name": ["科目名称", "名称", "科目"],
    "open_dir": ["期初方向", "期初借贷方向", "期初借贷"],
    "open_amt": ["期初余额", "期初", "期初金额"],
    "debit": ["本期借方", "借方", "借方发生额", "本期借方发生额"],
    "credit": ["本期贷方", "贷方", "贷方发生额", "本期贷方发生额"],
    "close_dir": ["期末方向", "期末借贷方向", "期末借贷"],
    "close_amt": ["期末余额", "期末", "期末金额"],
}


def _read_km_1002(path):
    """读取科目余额表，提取银行存款(1002)控制数（含子目）。

    返回 dict {open, debit, credit, close}（均带借贷符号，资产借为+）；找不到返回 None。
    规则：优先取代码精确 '1002' 的父行；若不存在则汇总所有 '1002' 开头的子目，避免双重计列。
    """
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
        return None

    def _safe_amt(v):
        if v is None:
            return 0.0
        try:
            return _parse_amount(v, 0)
        except (ValueError, TypeError):
            return 0.0

    def _signed(v, direction):
        a = _safe_amt(v)
        d = _norm(direction) if direction is not None else ""
        return a if d in ("借", "debit", "dr", "j") or d == "" else -a

    # 定位表头：含 code 别名，且含余额类列名或『方向/本期借(贷)方』(兼容表头仅写"金额"的布局)
    code_aliases = set(_norm(a) for a in KM_COLUMN_ALIASES["code"])
    amt_aliases = set(_norm(a) for a in KM_COLUMN_ALIASES["open_amt"]
                      + KM_COLUMN_ALIASES["close_amt"])
    hdr_idx, hdr = 0, [str(h).strip() if h is not None else "" for h in raw[0]]
    for i, row in enumerate(raw):
        norm_cells = [_norm(c) for c in row if c is not None]
        has_amt = any(nc in amt_aliases for nc in norm_cells)
        has_dc = any(_norm(h) in ("本期借方", "本期贷方", "借方", "贷方") for h in row if h is not None)
        has_dir = any(_norm(h) == "方向" for h in row if h is not None)
        if any(nc in code_aliases for nc in norm_cells) and (has_amt or has_dc or has_dir):
            hdr_idx, hdr = i, [str(h).strip() if h is not None else "" for h in row]
            break

    def cidx(std):
        targets = set(_norm(a) for a in KM_COLUMN_ALIASES[std])
        for i, h in enumerate(hdr):
            if _norm(h) in targets:
                return i
        return None

    ci = cidx("code")
    di = cidx("debit")
    cri = cidx("credit")
    # 期初/期末余额列：优先别名；否则取『方向』列后紧邻的『金额』列(兼容表头仅写"金额"的布局)
    dirs = [i for i, h in enumerate(hdr) if _norm(h) == "方向"]
    oai = cidx("open_amt")
    if oai is None and dirs:
        oai = dirs[0] + 1 if dirs[0] + 1 < len(hdr) else None
    cai = cidx("close_amt")
    if cai is None and len(dirs) >= 2:
        cai = dirs[1] + 1 if dirs[1] + 1 < len(hdr) else None
    odi = cidx("open_dir") or (dirs[0] if dirs else None)
    cdi = cidx("close_dir") or (dirs[1] if len(dirs) >= 2 else None)
    if ci is None or (oai is None and cai is None):
        return None

    def _row_dict(row):
        return {
            "open": _signed(row[oai] if oai is not None and oai < len(row) else 0,
                           row[odi] if odi is not None and odi < len(row) else "借"),
            "debit": _safe_amt(row[di] if di is not None and di < len(row) else 0),
            "credit": _safe_amt(row[cri] if cri is not None and cri < len(row) else 0),
            "close": _signed(row[cai] if cai is not None and cai < len(row) else 0,
                            row[cdi] if cdi is not None and cdi < len(row) else "借"),
        }

    parent = None
    children = []
    for row in raw[hdr_idx + 1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        if ci >= len(row) or row[ci] is None:
            continue
        code = _norm(row[ci]).replace(".0", "")
        if code == "1002":
            parent = _row_dict(row)
        elif code.startswith("1002"):
            children.append(_row_dict(row))
    if parent is not None:
        return parent
    if not children:
        return None
    return {
        "open": sum(c["open"] for c in children),
        "debit": sum(c["debit"] for c in children),
        "credit": sum(c["credit"] for c in children),
        "close": sum(c["close"] for c in children),
    }


def build_bank_detail(input_path=None, output_path=None, period_mode="Y", opening_path=None, rows=None, read_warnings=None, control_path=None, opening_dict=None, control_map=None):
    """核心流程：读取 -> 分组 -> 汇总 -> 校验 -> 导出。返回 (issues, periods, accounts)
    说明：input_path 与 rows 二选一。提供 rows 时跳过文件读取（用于目录多文件合并等场景）。
    control_path：可选科目余额表路径，用于银行存款(1002)控制数勾稽(单期/缺省期)。
    control_map：可选 dict{期间: 科目余额表路径}，多期(GL含多年)时按期间分别核对控制数。
    opening_dict：可选内存期初余额 {(账号,币种): 金额}，优先于 opening_path（GL 模式由科目余额表推导）。"""
    if rows is None:
        if input_path is None:
            raise ValueError("必须提供 input_path 或 rows 之一")
        read_warnings = []
        rows = read_transactions(input_path, warnings=read_warnings)
    else:
        read_warnings = read_warnings or []
    if opening_dict is not None:
        opening, open_conflicts = opening_dict, []
    elif opening_path:
        opening, open_conflicts = read_opening_balances(opening_path)
    else:
        opening, open_conflicts = {}, []
    periods, accounts, agg = aggregate(rows, period_mode)
    # 将『仅存在于科目余额表(期初)而无 GL 分录』的休眠账户并入账户清单，
    # 使其期初/期末进入控制数汇总（否则休眠账户期初被漏计，导致控制数期初/期末差异）。
    if opening_dict:
        _acct_set = set(accounts)
        for _ak in opening_dict:
            if _ak not in _acct_set:
                accounts.append(_ak)
                _acct_set.add(_ak)
    summary = compute_summary(accounts, periods, agg, opening)
    issues = validate(rows, accounts, periods, agg, summary, opening)
    # 合并读取期警告（负金额等）
    for w in read_warnings:
        issues.append({"level": "WARN", "type": "金额", "message": w})
    # 合并期初余额冲突（重复账号不一致）
    for (no, cur, a1, a2) in open_conflicts:
        issues.append({
            "level": "ERROR", "type": "期初余额",
            "message": f"期初余额表中账号『{no}』({cur})出现冲突值：{a1:.2f} 与 {a2:.2f}，已取末值 {a2:.2f}，请核实。",
        })
    # ===== 科目余额表 银行存款(1002) 控制数勾稽（按期间分别核对）=====
    ctrl_sources = control_map or {}
    for pk in periods:
        km_path = ctrl_sources.get(pk) or control_path
        if not km_path:
            continue
        km = _read_km_1002(km_path)
        if km is None:
            issues.append({"level": "WARN", "type": "控制数核对",
                           "message": f"科目余额表『{os.path.basename(km_path)}』未找到银行存款(1002)科目，跳过{pk}控制数核对。"})
            continue
        calc_debit = sum(summary[(a[0], a[1], pk)]["debit"] for a in accounts)
        calc_credit = sum(summary[(a[0], a[1], pk)]["credit"] for a in accounts)
        calc_open = sum(summary[(a[0], a[1], pk)]["open"] for a in accounts)
        calc_close = sum(summary[(a[0], a[1], pk)]["close"] for a in accounts)
        ctrl = {"期初": km["open"], "借方": km["debit"], "贷方": km["credit"], "期末": km["close"]}
        calc = {"期初": calc_open, "借方": calc_debit, "贷方": calc_credit, "期末": calc_close}
        tied = True
        for k in ("期初", "借方", "贷方", "期末"):
            d = calc[k] - ctrl[k]
            if abs(d) > 0.005:
                tied = False
                issues.append({"level": "WARN", "type": "控制数核对",
                               "message": f"银行存款控制数({pk})：计算{k}({calc[k]:.2f})与科目余额表1002{k}({ctrl[k]:.2f})不一致，差异{d:.2f}。"})
        if tied:
            issues.append({"level": "OK", "type": "控制数核对",
                           "message": f"银行存款(1002)与科目余额表控制数对平（{pk}期末{calc_close:.2f}）。"})
    name_map = _build_name_map(rows)
    export_excel(output_path, periods, accounts, summary, issues, rows, name_map, period_mode)
    return issues, periods, accounts


# ============================================================================
# 科目余额表模式（主表数字直接取自科目余额表 1002 子目 + 外币科目余额表原币）
# 说明：本项目方法论中《综合查询明细表》为非全量抽取，其累加的借/贷 ≠ 科目余额表
#       完整借/贷。故银行存款明细表主表数字须直接取自科目余额表(控制数)，而非从
#       GL 反算；GL 仅作交易流水下钻，并标注『非全量抽取』。
# ============================================================================

def _km_find_header(raw):
    """定位科目余额表表头，返回 (hdr_idx, hdr, ci, di, cri, oai, cai, odi, cdi)。"""
    code_aliases = set(_norm(a) for a in KM_COLUMN_ALIASES["code"])
    amt_aliases = set(_norm(a) for a in KM_COLUMN_ALIASES["open_amt"]
                      + KM_COLUMN_ALIASES["close_amt"])
    hdr_idx, hdr = None, None
    for i, row in enumerate(raw):
        nc = [_norm(c) for c in row if c is not None]
        has_amt = any(nc1 in amt_aliases for nc1 in nc)
        has_dc = any(_norm(h) in ("本期借方", "本期贷方", "借方", "贷方") for h in row if h is not None)
        has_dir = any(_norm(h) == "方向" for h in row if h is not None)
        if any(nc1 in code_aliases for nc1 in nc) and (has_amt or has_dc or has_dir):
            hdr_idx, hdr = i, [str(h).strip() if h is not None else "" for h in row]
            break
    if hdr is None:
        return (None,) * 9

    def cidx(std):
        targets = set(_norm(a) for a in KM_COLUMN_ALIASES[std])
        for i, h in enumerate(hdr):
            if _norm(h) in targets:
                return i
        return None

    ci = cidx("code"); di = cidx("debit"); cri = cidx("credit")
    dirs = [i for i, h in enumerate(hdr) if _norm(h) == "方向"]
    oai = cidx("open_amt")
    if oai is None and dirs:
        oai = dirs[0] + 1 if dirs[0] + 1 < len(hdr) else None
    cai = cidx("close_amt")
    if cai is None and len(dirs) >= 2:
        cai = dirs[1] + 1 if dirs[1] + 1 < len(hdr) else None
    odi = cidx("open_dir") or (dirs[0] if dirs else None)
    cdi = cidx("close_dir") or (dirs[1] if len(dirs) >= 2 else None)
    return hdr_idx, hdr, ci, di, cri, oai, cai, odi, cdi


def read_km_bank_accounts(km_path):
    """读取科目余额表，返回 {'1002':(parent,children), '1012':(parent,children)}。
    parent/child 均含本位币 open/debit/credit/close；children 还带 code/name。
    数字直接作为银行存款明细表主表(=控制数)，不反算。"""
    # SAP：从 adapter.read_km 聚合 1001/1002/1012 账户
    # ⚡ 2026-08-10 修复：①out 补 1001 库存现金（原只有 1002/1012 → 1001 永远 (None,[])）；
    # ②父级 p_par 补 code/name 键（原只有金额键 → 1012 有父级无子目时 rows_src=[parent]，
    #   _km_subject_rows 取 ch["code"] 抛 KeyError，2200 银行存款底稿缺失）。
    if _adapter is not None and (km_path is None or not isinstance(km_path, str) or _adapter.is_sap(km_path)):
        out = {'1001': (None, []), '1002': (None, []), '1012': (None, [])}
        km = _adapter.read_km()
        _SAP_NAMES = {'1001': '库存现金', '1002': '银行存款', '1012': '其他货币资金'}
        for pref in ('1001', '1002', '1012'):
            subs = []
            p_par = {'code': pref, 'name': _SAP_NAMES[pref],
                     'open': 0.0, 'debit': 0.0, 'credit': 0.0, 'close': 0.0}
            for code, v in sorted(km.items()):
                if not code.startswith(pref):
                    continue
                if v.get('level') == 1:
                    p_par = {'code': pref, 'name': v.get('name', _SAP_NAMES[pref]),
                             'open': float(v.get('opening') or 0.0), 'debit': float(v.get('debit') or 0.0),
                             'credit': float(v.get('credit') or 0.0), 'close': float(v.get('closing') or 0.0)}
                    continue
                subs.append({'code': code, 'name': v.get('name', ''),
                             'open': float(v.get('opening') or 0.0), 'debit': float(v.get('debit') or 0.0),
                             'credit': float(v.get('credit') or 0.0), 'close': float(v.get('closing') or 0.0)})
            out[pref] = (p_par, subs)
        return out
    try:
        wb = openpyxl.load_workbook(km_path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    except Exception:
        # 解析失败：返回空字典（而非元组），避免下游 accounts.get(...) 崩溃；
        # 该主体银行存款将留空并由调用方在 issues 中提示（数据源非标/损坏）。
        return {"1002": (None, []), "1012": (None, [])}
    hdr_idx, hdr, ci, di, cri, oai, cai, odi, cdi = _km_find_header(raw)
    if hdr is None or ci is None or (oai is None and cai is None):
        return {"1002": (None, []), "1012": (None, [])}

    def _safe(v):
        try:
            if v is None:
                return 0.0
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v).strip().replace(",", "").replace("￥", "").replace("¥", "")
            return float(re.sub(r"[()]", "", s)) if s else 0.0
        except Exception:
            return 0.0

    def _sgn(v, d):
        a = _safe(v)
        dd = _norm(d) if d is not None else ""
        return a if dd in ("借", "debit", "dr", "j") or dd == "" else -a

    def _rd(row, name):
        return {
            "name": name,
            "open": _sgn(row[oai] if oai is not None and oai < len(row) else 0,
                        row[odi] if odi is not None and odi < len(row) else "借"),
            "debit": _safe(row[di] if di is not None and di < len(row) else 0),
            "credit": _safe(row[cri] if cri is not None and cri < len(row) else 0),
            "close": _sgn(row[cai] if cai is not None and cai < len(row) else 0,
                         row[cdi] if cdi is not None and cdi < len(row) else "借"),
        }

    ni = None
    for i, h in enumerate(hdr):
        if _norm(h) in set(_norm(a) for a in KM_COLUMN_ALIASES["name"]):
            ni = i
            break
    # 银行存款/其他货币资金 取数兼容：标准科目表 1002/1012，亦兼容本地科目表
    # （如泰国 1113 银行存款、1112 其他货币资金）。按『代码或科目名』归类，
    # 归并到 "1002"/"1012" 键（下游勾稽逻辑无需改动）。
    BANK_CODES = {"1002", "1113"}     # 银行存款（含本地码）
    OTHER_CODES = {"1012", "1112"}    # 其他货币资金（含本地码）

    def _cat(code, name):
        nc = _norm(name)
        # 先按精确码 / 科目名归类（标准科目表 + 账套配置本地码）
        if code in _BANK_DEPOSIT_CODES or nc == "银行存款":
            return "bank"
        if code in OTHER_CODES or nc == "其他货币资金":
            return "other"
        # 2026-08-05：库存现金(1001)纳入明细表（用户反馈：06FY智能 货币资金明细表
        # 没取到现金科目）——1001 无 aux 明细，父级即唯一账户行。
        if code in ("1001",) or nc in ("库存现金", "现金"):
            return "cash"
        # 再按『代码前缀』归类：本项目账套银行存款/其他货币资金多采用无点子码
        # （如 10021101=中国银行XX支行、101201=银行汇票），必须纳入末级账户，
        # 否则明细表只会剩 1002/1012 父级两行。
        if any(code.startswith(c) for c in _BANK_DEPOSIT_CODES):
            return "bank"
        if code.startswith("1012") or code.startswith("1112"):
            return "other"
        if code.startswith("1001"):
            return "cash"
        return None

    def _bucket(rows):
        """rows: list of (code, name, row)。返回 (parent_dict, children_list)。
        parent = 该类别根码行（最短且为其他同码共同前缀，如 1002/1113）；
        children = 该类别下全部『末级账户』行（不是任何其他同码前缀的码）。
        兼容科目代码『带点』(1002.01.01) 与『无点』(10021101) 两种层级写法：
        本项目账套多为无点写法，故用『前缀包含』而非『前缀+点』判定层级。
        无末级子目时（如 1113 无 1113.xx）children 为空，由调用方将 parent 作为唯一账户行补列。"""
        if not rows:
            return None, []
        codes = [c for c, _, _ in rows]
        # 根码 = 是所有其他码共同前缀的最短码（如 1002 / 1012 / 1113）
        root = None
        for c in sorted(codes, key=len):
            if all(c2 == c or c2.startswith(c) for c2 in codes):
                root = c
                break
        if root is None:
            root = min(codes, key=len)
        pr = next((r for (c, n, r) in rows if c == root), None)
        pname = next((n for (c, n, r) in rows if c == root), root)
        p = _rd(pr, pname); p["code"] = root
        # 末级 = 不是任何其他同码前缀的码（无点/有点通用）
        kids = []
        _by_code = {c: (n, r) for (c, n, r) in rows}
        _BANK_KW = ('银行', '支行', '分行', '营业部', '信用社', '农信', '农商', '村镇银行',
                    '储蓄', '专户', '账户')
        for (c, n, r) in rows:
            if c == root or any(c2 != c and c2.startswith(c) for c2 in codes):
                continue
            kid = {"code": c, "name": n, **_rd(r, n)}
            # 2026-08-07 修复（dq 银行存款未取到真正的下级明细）：末级账户名是
            # 通用性质词（活期存款/定期存款等）时，真正的账户（开户银行）在直接
            # 父级（如 100201 中行嘉兴秀城支行）。向上找直接父级：父级名含银行
            # 特征词（银行/支行/分行/营业部等）→ 账户显示名 = 父级银行名 + 末级
            # 性质词（同银行活期/定期两行时自动区分），并保留 parent_name 供
            # 开户银行列单独展示。
            _pc = c[:-2]
            while _pc and _pc != root and _pc in _by_code:
                _pn, _pr = _by_code[_pc]
                if any(k in str(_pn) for k in _BANK_KW) and _pn != n:
                    kid['parent_name'] = _pn
                    kid['sub_name'] = n
                    kid['name'] = f"{_pn}-{n}" if _pn else n
                    break
                _pc = _pc[:-2]
            kids.append(kid)
        return p, kids

    bank_rows, other_rows, cash_rows = [], [], []
    for row in raw[hdr_idx + 1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        if ci >= len(row) or row[ci] is None:
            continue
        code = re.sub(r"\.0$", "", _norm(row[ci]))
        name = str(row[ni]) if ni is not None and ni < len(row) and row[ni] is not None else code
        cat = _cat(code, name)
        if cat == "bank":
            bank_rows.append((code, name, row))
        elif cat == "other":
            other_rows.append((code, name, row))
        elif cat == "cash":
            cash_rows.append((code, name, row))
    parent, children = _bucket(bank_rows)
    parent_1012, children_1012 = _bucket(other_rows)
    parent_1001, children_1001 = _bucket(cash_rows)
    return {"1001": (parent_1001, children_1001),
            "1002": (parent, children), "1012": (parent_1012, children_1012)}


def read_fc_bank_accounts(fc_path):
    """读取外币科目余额表，返回 FC 银行账户列表(原币口径)。
    元素：{code,name,ccy,open_fc,debit_fc,credit_fc,close_fc}。"""
    out = []
    if not fc_path:
        return out
    try:
        wb = openpyxl.load_workbook(fc_path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    except Exception:
        return out
    hdr, hi = None, None
    for i, row in enumerate(raw):
        nc = [str(c).strip() for c in row if c is not None]
        if "科目代码" in nc and "原币" in nc:
            hdr, hi = row, i
            break
    if hdr is None:
        return out

    def col(name):
        for j, h in enumerate(hdr):
            if str(h).strip() == name:
                return j
        return None

    ci = col("科目代码"); ni = col("科目名称"); ccyc = col("币种")
    yuan = [j for j, h in enumerate(hdr) if str(h).strip() == "原币"]
    if len(yuan) >= 4:
        oi, yi, ci2, cl = yuan[0], yuan[1], yuan[2], yuan[3]
    elif len(yuan) == 1:
        oi = yi = ci2 = cl = yuan[0]
    else:
        return out

    def _fc(v):
        return _safe_amt(v)

    for row in raw[hi + 1:]:
        if not row or ci is None or ci >= len(row) or row[ci] is None:
            continue
        code = re.sub(r"\.0$", "", str(row[ci]).strip())
        if code == "1002" or code == "1012":
            continue
        if not code.startswith("1002") and not code.startswith("1012"):
            continue
        name = str(row[ni]).strip() if ni is not None and ni < len(row) and row[ni] is not None else ""
        ccy = str(row[ccyc]).strip() if ccyc is not None and ccyc < len(row) and row[ccyc] is not None else ""
        out.append({
            "code": code, "name": name, "ccy": ccy,
            "open_fc": _fc(row[oi] if oi < len(row) else 0),
            "debit_fc": _fc(row[yi] if yi < len(row) else 0),
            "credit_fc": _fc(row[ci2] if ci2 < len(row) else 0),
            "close_fc": _fc(row[cl] if cl < len(row) else 0),
        })
    return out


def _parse_account_bank(name):
    """从账户名解析 开户银行 / 账号（账号=尾部数字, 可带 -N；支持半角/全角括号包裹如『…（514467384883）』）。"""
    s = str(name).strip()
    m = re.search(r"([0-9]+(?:-[0-9]+)?)\s*[)）]?\s*$", s)
    if m:
        bank = s[:m.start()].strip().rstrip("（([")
        return bank, m.group(1)
    return s, ""


def _detect_fc(name):
    """按账户名关键字识别外币币种（即使无外币科目余额表也能标注）。"""
    s = str(name)
    for kw, ccy in (("美元", "USD"), ("欧元", "EUR"), ("日元", "JPY"), ("港币", "HKD")):
        if kw in s:
            return ccy
    return None


def _extract_year(fn):
    """从文件名稳健识别会计年度。

    常见陷阱：25 年账套导出文件带 2026 时间戳（如 『…_20260723152733.xlsx』），
    若用 re.search(r'(20\\d{2})', fn) 会先命中时间戳 2026 而非会计年度 2025。
    本函数先剥离导出时间戳（6~14 位 20xx 数字串），再匹配 4 位会计年度(20xx)，
    否则匹配独立 2 位年度(如 25 → 2025)。
    ⚡ 2026-08-15 修复（同步 audit_common）：优先取【类型名（科目余额表/综合查询
    明细表等）紧邻前】的 YYYY——标准命名『主体 YYYY年 类型』；防主体名含年份
    （XBJ『2020年简阳市…』『2022零星维修工程』）误判为年度 → 虚 2020/2022 期间。"""
    s = fn
    # 0) 类型名前年份优先（标准命名『主体 YYYY年 类型』）
    m = re.search(r'(20\d{2})\s*年?\s*(外币)?(科目余额表|科目余额|余额表|综合查询明细表|综合查询表|辅助核算余额表|辅助核算|总账|trial)',
                  s, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    # 1) 剥离 20 + 6~12 位数字的导出时间戳（=8~14 位日期时间串，如 20260723152733）
    s = re.sub(r"(?<!\d)20\d{6,12}(?!\d)", "", s)
    # 2) 剥离独立 8 位 YYYYMMDD 时间戳
    s = re.sub(r"(?<!\d)20\d{6}(?!\d)", "", s)
    # 3) 4 位会计年度（如 2025）
    m = re.search(r"(?<!\d)(20\d{2})(?!\d)", s)
    if m:
        return m.group(1)
    # 4) 独立 2 位年度（如文件名含 25）→ 20xx
    m2 = re.search(r"(?<!\d)(\d{2})(?!\d)", s)
    if m2:
        try:
            yy = int(m2.group(1))
        except ValueError:
            return None
        if 0 <= yy <= 99:
            return f"20{yy:02d}"
    return None


def _discover_aux_balance(folder):
    """发现目录内的『辅助核算余额表』（用于需求4：外币是否可从其取数）。"""
    if not os.path.isdir(folder):
        return None
    for fn in sorted(os.listdir(folder)):
        if "辅助核算" in fn and fn.lower().endswith((".xlsx", ".csv")):
            return os.path.join(folder, fn)
    return None


def _scan_aux_bank(aux_path):
    """扫描辅助核算余额表，返回『银行存款/其他货币资金』相关行数（按 科目代码 1002/1012 或
    科目名称含『银行存款/其他货币资金』判定，避免把『应收票据-银行承兑』误计入）。"""
    try:
        wb = openpyxl.load_workbook(aux_path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    except Exception:
        return 0
    hdr_idx, hdr, ci, di, cri, oai, cai, odi, cdi = _km_find_header(raw)
    if hdr_idx is None:
        # 辅助核算余额表列名与科目余额表不同，退化为全单元格关键字扫描
        n = 0
        for row in raw:
            s = " ".join(str(c) for c in row if c is not None)
            if ("1002" in s or "1012" in s) and ("银行存款" in s or "其他货币资金" in s):
                n += 1
        return n
    ni = None
    for i, h in enumerate(hdr):
        if _norm(h) in set(_norm(a) for a in KM_COLUMN_ALIASES["name"]):
            ni = i
            break
    n = 0
    for row in raw[hdr_idx + 1:]:
        if not row or all(c is None or str(c).strip() == "" for c in row):
            continue
        code = _norm(row[ci]).replace(".0", "") if ci is not None and ci < len(row) and row[ci] is not None else ""
        name = str(row[ni]).strip() if ni is not None and ni < len(row) and row[ni] is not None else ""
        if code.startswith("1002") or code.startswith("1012") or "银行存款" in name or "其他货币资金" in name:
            n += 1
    return n


def _discover_inputs_km(folder, exclude_name=None):
    """按年发现 科目余额表 / 外币科目余额表 / 综合查询明细表。"""
    km_by_year, fc_by_year, gl_by_year = {}, {}, {}
    for fn in sorted(os.listdir(folder)):
        if not fn.lower().endswith((".xlsx", ".csv")):
            continue
        if exclude_name and fn == exclude_name:
            continue
        low = fn.lower()
        y = _extract_year(fn)
        if "外币" in fn and ("科目余额表" in fn or "余额表" in fn):
            if y:
                fc_by_year[y] = os.path.join(folder, fn)
        elif ("科目余额表" in fn or "trial" in low or "总账" in fn) and "明细" not in fn \
                and "综合查询" not in fn and "辅助核算" not in fn:
            if y:
                km_by_year[y] = os.path.join(folder, fn)
        elif "综合查询" in fn:
            if y:
                gl_by_year[y] = os.path.join(folder, fn)
    return km_by_year, fc_by_year, gl_by_year


def _derive_subject(km_by_year):
    for p in km_by_year.values():
        fn = os.path.basename(p)
        m = re.match(r"^(.*?)科目余额表", fn)
        if m and m.group(1).strip():
            return m.group(1).strip()
    return "（核算主体）"


def _km_subject_rows(accounts, fc_map, entity, y, issues):
    """把 read_km_bank_accounts 返回的 {'1002':(parent,children),'1012':(...)}
    转为明细行列表，并做子目合计 vs 父级（控制数）勾稽、期初+借-贷=期末（平衡）校验。
    返回明细行列表（含 subject/subject_name；entity 由调用方决定是否存在）。"""
    rows = []
    tag = f"[{entity}][{y}] " if entity else f"{y} "
    SUBJ_NAMES = {"1001": "库存现金", "1002": "银行存款", "1012": "其他货币资金"}
    for subj in ("1001", "1002", "1012"):
        parent, children = accounts.get(subj, (None, []))
        subj_name = SUBJ_NAMES.get(subj, subj)
        # 2026-08-05：库存现金(1001)无辅助明细但须纳入明细表（用户反馈：06FY智能
        # 货币资金明细表没取到现金科目）——1001 父级即唯一账户行。
        if parent is None and subj == "1002":
            issues.append({"level": "WARN", "type": "控制数",
                           "message": f"{tag}科目余额表未找到 1002 父级（该主体可能采用本地科目表，如 1113 银行存款），"
                                      f"其银行存款未纳入明细表，请确认是否需扩展取数范围。"})
        if parent is not None:
            # 仅当有末级子目时，做『子目合计 vs 父级（控制数）』勾稽；
            # 若无末级子目（如泰国 1113 无 1113.xx），父级本身就是唯一账户行，跳过该勾稽。
            if children:
                for label, key in (("期初", "open"), ("借方", "debit"), ("贷方", "credit"), ("期末", "close")):
                    sval = sum(c[key] for c in children)
                    d = sval - parent[key]
                    if abs(d) > 0.005:
                        issues.append({"level": "WARN", "type": "控制数",
                                       "message": f"{tag}{subj_name}（{subj}）子目合计{label}({sval:.2f})与{subj}父级({parent[key]:.2f})差异{d:.2f}。"})
            # 账户行来源：有末级子目取子目，否则以父级作为唯一账户行。
            rows_src = children if children else [parent]
            for c in rows_src:
                bal = c["open"] + c["debit"] - c["credit"] - c["close"]
                if abs(bal) > 0.005:
                    issues.append({"level": "WARN", "type": "平衡",
                                   "message": f"{tag}{subj_name}账户『{c['name']}』期初+借-贷≠期末，差额{bal:.2f}（来源科目余额表，请核对导出）。"})
        else:
            rows_src = []
        for ch in rows_src:
            code = ch["code"]; name = ch["name"]
            fc = fc_map.get(code)
            # 2026-08-07：末级账户名=性质词（活期存款/定期存款）时，开户银行=父级
            # 银行名（dq 1002→100201 中行嘉兴秀城支行→10020101 活期存款 三层结构），
            # 否则从账户名尾部解析银行名+账号。
            if ch.get('parent_name'):
                bank = ch['parent_name']
                acc = _parse_account_bank(ch.get('sub_name') or name)[1]
            else:
                bank, acc = _parse_account_bank(name)
            ccy_by_name = _detect_fc(name)
            is_fc = fc is not None or ccy_by_name is not None
            ccy = fc["ccy"] if fc else (ccy_by_name or "人民币")
            性质 = "外币账户" if is_fc else "人民币账户"
            # 2026-08-07：账户性质列附加末级性质词（活期存款/定期存款等）——
            # 仅当 name 已含性质词（银行名-活期存款 格式）时展示，避免与开户银行重复。
            if ch.get('sub_name'):
                性质 = f"{性质}（{ch['sub_name']}）"
            row = {"year": y, "subject": subj, "subject_name": subj_name,
                   "code": code, "name": name, "bank": bank, "acc": acc,
                   "性质": 性质, "ccy": ccy,
                   "open": ch["open"], "debit": ch["debit"], "credit": ch["credit"], "close": ch["close"],
                   "open_fc": fc["open_fc"] if fc else None,
                   "debit_fc": fc["debit_fc"] if fc else None,
                   "credit_fc": fc["credit_fc"] if fc else None,
                   "close_fc": fc["close_fc"] if fc else None}
            if entity is not None:
                row["entity"] = entity
            rows.append(row)
    return rows


def build_bank_detail_km(km_by_year, fc_by_year, gl_by_year, output_path,
                         subject_name="", period_mode="Y"):
    """科目余额表模式核心：主表数字直接取自科目余额表1002 + 外币科目余额表原币。
    返回 (issues, years, detail_rows)。"""
    issues = []
    years = sorted(km_by_year.keys())
    detail_rows = []
    parent_totals = {}
    for y in years:
        accounts = read_km_bank_accounts(km_by_year[y])
        fc_list = read_fc_bank_accounts(fc_by_year.get(y)) if fc_by_year.get(y) else []
        fc_map = {f["code"]: f for f in fc_list}
        parent_totals[y] = accounts
        detail_rows.extend(_km_subject_rows(accounts, fc_map, None, y, issues))
    # 交易流水下钻(GL 非全量抽取)
    gl_rows = []
    gl_warn = []
    for y in years:
        g = gl_by_year.get(y)
        if g:
            rr = read_gl_bank(g, gl_warn)
            for x in rr:
                x["period"] = y
            gl_rows += rr
    for w in gl_warn:
        issues.append({"level": "WARN", "type": "取数范围", "message": w})
    if not fc_by_year:
        issues.append({"level": "INFO", "type": "外币原币",
                       "message": "未提供外币科目余额表，外币账户原币列留空（仅列示本位币）。"})
    else:
        for y in years:
            if y not in fc_by_year:
                issues.append({"level": "INFO", "type": "外币原币",
                               "message": f"{y} 无外币科目余额表，该年外币原币列留空。"})
    _export_km_excel(output_path, detail_rows, parent_totals, years, subject_name, gl_rows, issues)
    return issues, years, detail_rows


def _export_km_excel(output_path, detail_rows, parent_totals, years, subject_name, gl_rows, issues):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook()
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="DDEBF7")
    hdr_font = Font(name='Times New Roman', bold=True, color="000000", size=10)
    FONT10 = Font(name='Times New Roman', size=10)
    total_fill = PatternFill("solid", fgColor="DDEBF7")
    grand_fill = PatternFill("solid", fgColor="BDD7EE")
    money = "#,##0.00"

    def style_header(ws, ncol, row=1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def total_row(ws, r, ncol, vals, fill=total_fill, bold=True):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if bold:
                cell.font = Font(name='Times New Roman', bold=True, size=10)
            cell.fill = fill
        for c, v in vals.items():
            ws.cell(row=r, column=c, value=v)
            if c in (8, 9, 10, 11, 12, 13, 14, 15, 17) and isinstance(v, (int, float)):
                ws.cell(row=r, column=c).number_format = money

    # ---- Sheet1 银行存款明细表（与合并模式一致的列方案） ----
    ws = wb.active
    ws.title = "银行存款明细表"
    cols = ["会计期间", "序号", "核算主体", "开户银行", "账号", "账户性质", "币种",
            "期初原币", "期初余额(本币)", "借方原币", "本期借方(本币)",
            "贷方原币", "本期贷方(本币)", "期末原币", "期末余额(本币)",
            "平衡校验",
            "原币(汇兑测算)", "期末汇率", "应计本位币", "汇兑差异",
            "对账单金额", "对账单索引", "对账单差异", "余额调节表索引", "函证确认金额",
            "科目代码", "科目名称"]
    MONEY_COLS = {8, 9, 10, 11, 12, 13, 14, 15, 17, 19, 20, 21, 23, 25}
    ws.append(cols)
    style_header(ws, len(cols))

    def _emit1(ws, r, rows, seq_start, y):
        for i, d in enumerate(rows, seq_start):
            bal = d["open"] + d["debit"] - d["credit"] - d["close"]
            ws.append([y, i, subject_name, d["bank"], d["acc"], d["性质"], d["ccy"],
                       d["open_fc"], d["open"], d["debit_fc"], d["debit"],
                       d["credit_fc"], d["credit"], d["close_fc"], d["close"],
                       round(bal, 2),
                       d["close_fc"], "", None, None,
                       None, "", None, "", None,
                       d["code"], d["subject_name"]])
            ws.cell(row=r, column=19, value='=IF(R{r}="","",Q{r}*R{r})'.format(r=r))
            ws.cell(row=r, column=20, value='=IF(S{r}="","",S{r}-N{r})'.format(r=r))
            ws.cell(row=r, column=23, value='=IF(U{r}="","",U{r}-N{r})'.format(r=r))
            for c in range(1, len(cols) + 1):
                cell = ws.cell(row=r, column=c); cell.border = border
                cell.font = FONT10
                if c in MONEY_COLS:
                    cell.number_format = money
                if abs(bal) > 0.005 and c == 16:
                    cell.fill = warn_fill
            r += 1
        return r

    def _blk1(ws, r, label, rows, subj, fill, yr=""):
        so = sum(d["open"] for d in rows); sd = sum(d["debit"] for d in rows)
        sc = sum(d["credit"] for d in rows); sl = sum(d["close"] for d in rows)
        sf_o = sum(d["open_fc"] or 0 for d in rows); sf_d = sum(d["debit_fc"] or 0 for d in rows)
        sf_c = sum(d["credit_fc"] or 0 for d in rows); sf_l = sum(d["close_fc"] or 0 for d in rows)
        kv = {"1001": "库存现金", "1002": "银行存款", "1012": "其他货币资金", "": "货币资金"}
        total_row(ws, r, len(cols), {1: yr, 2: "", 3: label, 8: sf_o, 9: so, 10: sf_d, 11: sd,
                                     12: sf_c, 13: sc, 14: sf_l, 15: sl, 17: sf_l,
                                     26: subj, 27: kv.get(subj, "")}, fill=fill)
        return r + 1

    r = 2
    for y in years:
        yrows = [d for d in detail_rows if d["year"] == y]
        cash_rows = [d for d in yrows if d["subject"] == "1001"]
        bank_rows = [d for d in yrows if d["subject"] == "1002"]
        other_rows = [d for d in yrows if d["subject"] == "1012"]
        # 2026-08-05：明细表须含库存现金块（1001 无 aux 明细，父级即唯一行；用户反馈
        # 06FY智能 货币资金明细表没取到现金科目）。货币资金合计=现金+银行+其他。
        if cash_rows:
            r = _emit1(ws, r, cash_rows, 1, y)
            r = _blk1(ws, r, "库存现金合计", cash_rows, "1001", total_fill, y)
        r = _emit1(ws, r, bank_rows, len(cash_rows) + 1, y)
        r = _blk1(ws, r, "银行存款合计", bank_rows, "1002", total_fill, y)
        if other_rows:
            r = _emit1(ws, r, other_rows, len(cash_rows) + len(bank_rows) + 1, y)
            r = _blk1(ws, r, "其他货币资金合计", other_rows, "1012", total_fill, y)
            r = _blk1(ws, r, "货币资金合计", yrows, "", grand_fill, y)
        cash_comp = sum(d["close"] for d in cash_rows)
        bank_comp = sum(d["close"] for d in bank_rows); other_comp = sum(d["close"] for d in other_rows)
        acc = parent_totals.get(y, {})
        p1 = acc.get("1002"); p2 = acc.get("1012"); p0 = acc.get("1001")
        bank_ctrl = p1[0]["close"] if (p1 and p1[0] is not None) else 0.0
        other_ctrl = p2[0]["close"] if (p2 and p2[0] is not None) else 0.0
        cash_ctrl = p0[0]["close"] if (p0 and p0[0] is not None) else 0.0
        if cash_rows and abs(cash_comp - cash_ctrl) > 0.005:
            issues.append({"level": "WARN", "type": "勾稽",
                           "message": f"{y} 库存现金明细表期末合计({cash_comp:,.2f})与科目余额表库存现金父级差异{cash_comp - cash_ctrl:,.2f}。"})
        if abs(bank_comp - bank_ctrl) > 0.005:
            issues.append({"level": "WARN", "type": "勾稽",
                           "message": f"{y} 银行存款明细表期末合计({bank_comp:,.2f})与科目余额表银行存款父级差异{bank_comp - bank_ctrl:,.2f}。"})
        if other_rows and abs(other_comp - other_ctrl) > 0.005:
            issues.append({"level": "WARN", "type": "勾稽",
                           "message": f"{y} 其他货币资金明细表期末合计({other_comp:,.2f})与科目余额表其他货币资金父级差异{other_comp - other_ctrl:,.2f}。"})
    to = sum(d["open"] for d in detail_rows); td = sum(d["debit"] for d in detail_rows)
    tc = sum(d["credit"] for d in detail_rows); tl = sum(d["close"] for d in detail_rows)
    tf_o = sum(d["open_fc"] or 0 for d in detail_rows); tf_d = sum(d["debit_fc"] or 0 for d in detail_rows)
    tf_c = sum(d["credit_fc"] or 0 for d in detail_rows); tf_l = sum(d["close_fc"] or 0 for d in detail_rows)
    total_row(ws, r, len(cols), {1: "", 2: "", 3: "总计", 8: tf_o, 9: to, 10: tf_d, 11: td,
                                 12: tf_c, 13: tc, 14: tf_l, 15: tl, 17: tf_l}, fill=grand_fill)
    r += 1
    widths = [10, 6, 14, 22, 18, 12, 8, 14, 16, 14, 16, 14, 16, 14, 16, 12,
              14, 10, 16, 14, 16, 12, 14, 16, 16, 12, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    # ---- Sheet2 期间对比 ----
    ws2 = wb.create_sheet("期间对比")
    cols2 = ["会计期间", "期初余额(本币)", "本期借方(本币)", "本期贷方(本币)", "期末余额(本币)", "外币原币期末合计(折算)"]
    ws2.append(cols2)
    style_header(ws2, len(cols2))
    r = 2
    for y in years:
        yrows = [d for d in detail_rows if d["year"] == y]
        so = sum(d["open"] for d in yrows); sd = sum(d["debit"] for d in yrows)
        sc = sum(d["credit"] for d in yrows); sl = sum(d["close"] for d in yrows)
        ws2.append([y, so, sd, sc, sl, ""])
        for c in range(1, len(cols2) + 1):
            cell = ws2.cell(row=r, column=c)
            cell.border = border
            if c in (2, 3, 4, 5):
                cell.number_format = money
        r += 1
    ws2.append(["总计", to, td, tc, tl, ""])
    for c in range(1, len(cols2) + 1):
        cell = ws2.cell(row=r, column=c)
        cell.border = border
        cell.fill = total_fill
        cell.font = Font(name='Times New Roman', bold=True)
        if c in (2, 3, 4, 5):
            cell.number_format = money
    for i, w in enumerate([12, 18, 18, 18, 18, 22], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ---- Sheet3 汇总 ----
    ws3 = wb.create_sheet("汇总")
    ws3.append(["项目", "金额(本币)"])
    style_header(ws3, 2)
    for i, (k, v) in enumerate([("期初余额合计", to), ("本期借方合计", td),
                                ("本期贷方合计", tc), ("期末余额合计", tl)], 2):
        ws3.append([k, v])
        ws3.cell(row=i, column=2).number_format = money
        for c in (1, 2):
            ws3.cell(row=i, column=c).border = border
    ws3.column_dimensions["A"].width = 20
    ws3.column_dimensions["B"].width = 20

    # ---- Sheet4 校验 ----
    ws4 = wb.create_sheet("校验")
    ws4.append(["类型", "说明", "状态"])
    style_header(ws4, 3)
    r = 2
    ok_fill = PatternFill("solid", fgColor="C6EFCE")
    warn_fill = PatternFill("solid", fgColor="FFEB9C")
    for y in years:
        acc = parent_totals.get(y)
        p1002 = acc.get("1002")[0] if (isinstance(acc, dict) and acc.get("1002") and acc.get("1002")[0] is not None) else None
        if p1002 is not None:
            sl = sum(d["close"] for d in detail_rows if d["year"] == y and d["subject"] == "1002")
            d_ = sl - p1002["close"]
            ws4.append([f"{y} 控制数对平", "银行存款主表期末合计 vs 科目余额表1002期末(构造相等)",
                        "OK" if abs(d_) <= 0.005 else f"差异{d_:.2f}"])
            ws4.cell(row=r, column=3).fill = ok_fill if abs(d_) <= 0.005 else warn_fill
            r += 1
    for it in issues:
        if it["type"] in ("平衡", "控制数") and it["level"] == "WARN":
            ws4.append([it["type"], it["message"], "关注"])
            ws4.cell(row=r, column=3).fill = warn_fill
            r += 1
    for i, w in enumerate([16, 72, 10], 1):
        ws4.column_dimensions[get_column_letter(i)].width = w

    # ---- Sheet5 交易流水(下钻) ----
    ws5 = wb.create_sheet("交易流水(下钻)")
    cols5 = ["日期", "账户", "摘要", "方向", "金额", "币种", "来源期间"]
    ws5.append(cols5)
    style_header(ws5, len(cols5))
    r = 2
    for row in gl_rows:
        ws5.append([row.get("date"), row.get("account"), row.get("summary"),
                    row.get("direction"), row.get("amount"), row.get("currency"), row.get("period", "")])
        cell = ws5.cell(row=r, column=5)
        if isinstance(cell.value, (int, float)):
            cell.number_format = money
        for c in range(1, len(cols5) + 1):
            ws5.cell(row=r, column=c).border = border
        r += 1
    for i, w in enumerate([12, 30, 42, 6, 16, 8, 12], 1):
        ws5.column_dimensions[get_column_letter(i)].width = w
    ws5.freeze_panes = "A2"

    finalize_workbook(wb)
    output_path, _save_warn = _safe_save(wb, output_path)
    if _save_warn:
        print(f'  [WARN] {_save_warn}')
    if output_path is None:
        raise PermissionError(_save_warn or '保存失败：目标与副本均无法写入')
    return output_path


# ============================================================================
# 模块 8：目录模式（多账套 × 多年份）合并输出 + 对方科目核对 + 异常判断
# ============================================================================
_FUND_KW = ("库存现金", "其他货币资金", "备用金", "1001", "1009", "1012",
            "结构性存款", "定期存款", "通知存款", "保证金", "存单", "协定存款")


def _is_fund_counterparty(opp):
    """判断对方科目是否为资金类（疑似内部资金划转/自身账户调拨）。
    注意：利息收入(财务费用-利息收入-银行存款活期利息)等仅是含『银行存款』字样的损益科目，
    并非资金划转，故不再用『银行存款』做宽泛子串匹配，改为：
      · 现金/其他货币资金/备用金/存款类专用科目或产品；
      · 对方科目本身为账户名(银行/支行/分行+账号数字 或 纯账号)。"""
    if not opp:
        return False
    s = opp.replace(" ", "")
    for kw in _FUND_KW:
        if kw in s:
            return True
    if re.search(r"(银行|支行|分行)\D*\d{3,}", s):
        return True
    if re.search(r"^\d{4,}$", s):
        return True
    return False


def _discover_bank_entities(data_dir):
    """按文件名前缀把账套文件归并到各核算主体，并按年分组。
    实体 = 文件名去掉『(外币)?科目余额表/综合查询明细表』与年份后缀。
    返回 {entity: {"km": {year:path}, "fc": {year:path}, "gl": {year:path}}}，
    仅保留至少含一份科目余额表的主体（无控制数则跳过）。"""
    # SAP：转本程序结构（km/fc/gl 均为 {y: path/list}；SAP 无外币表 fc 留空）
    # ⚡ 2026-08-09 统一：adapter.discover_entities 已内置 current_comp 单主体过滤，
    # 此处仅做 U8→本程序嵌套结构转换（km/fc/gl 均为 {y: path/list}；SAP 无外币表 fc 留空）
    # ⚡ 2026-08-12 治理：SAP adapter 结构 {y: ['km','gl',...]}（list）非 {y:{'km':..}}，
    #   原 `p.get('km')` 对 list 报 AttributeError → 统一用 norm_entities 转 U8 嵌套。
    if _adapter is not None and _adapter.is_sap(data_dir):
        from audit_common import norm_entities as _norm_bank
        ent = {}
        for _c, _bun in _norm_bank(_adapter.discover_entities(data_dir)).items():
            ent[_c] = {'km': dict(_bun.get('km') or {}),
                       'fc': {}, 'gl': dict(_bun.get('gl') or {})}
        return ent
    ent = {}
    if not os.path.isdir(data_dir):
        return ent
    for fn in sorted(os.listdir(data_dir)):
        if not fn.lower().endswith((".xlsx", ".xls", ".csv")):
            continue
        full = os.path.join(data_dir, fn)
        y = _extract_year(fn)
        if y is None:
            continue
        # 实体名清理：先剥离导出时间戳(长数字串)与年份，避免 2026 时间戳/25 残留在主体名
        clean = re.sub(r"(?<!\d)\d{6,14}(?!\d)", "", fn)   # 6~14 位导出时间戳
        clean = re.sub(r"(19|20)\d{2}", "", clean)          # 4 位年
        clean = re.sub(r"年度|年\d*[-~]?\d*月", "", clean)    # 剔除"年度"、"年1-3月"等期间后缀
        clean = re.sub(r"(?<!\d)\d{2}(?!\d)年", "", clean)   # 独立 2 位年（如"25年度"→"年度"）；2026-07-31 修复：
                                                              # 旧规则直接删任意独立 2 位数字，误删主体序号前缀"10、11、12"
                                                              # （"10、FY金属公司"→"、FY金属公司"），须后随"年"才删。
        E = re.sub(r"(外币)?(科目余额表|科目余额|余额表|综合查询明细表|综合查询表|总账|trial).*$", "",
                   clean, flags=re.IGNORECASE).strip(" _-")
        # 个位数前缀补0（1、XX → 01、XX），确保排序 1,2,...10 而非 1,10,11,2
        _m0 = re.match(r'^(\d)([、,，.．\s])', E)
        if _m0 and len(E) > 2:
            E = '0' + E
        if not E:
            # ⚡ 2026-08-12 修复（对齐 audit_common.discover_entities 兜底）：根目录直放且
            #   文件名无主体前缀（DQ/科目余额表_2024.xlsx 去年份后实体名=''）→ 回退账套
            #   目录名（DQ），否则该年度被整年跳过（DQ 2024 银行存款 130 亿缺底稿根因）。
            E = os.path.basename(os.path.normpath(data_dir)).strip(' _-')
            if not E:
                continue
        bucket = ent.setdefault(E, {"km": {}, "fc": {}, "gl": {}})
        if "外币" in fn and ("科目余额表" in fn or "余额表" in fn):
            bucket["fc"][y] = full
        elif ("科目余额表" in fn or "总账" in fn or "trial" in fn.lower()) \
                and "综合查询" not in fn and "辅助核算" not in fn:
            bucket["km"][y] = full
        elif "综合查询" in fn:
            bucket["gl"][y] = full
    return {E: b for E, b in ent.items() if b["km"]}


def _build_bank_frame(wb, years):
    """创建银行存款工作簿的静态 sheet 结构：仅按固定命名建空 sheet。
    『货币资金审定表』依赖数据文件夹内的模板（动态），由 _fill_capital_audit_sheet 按需建，
    不在此预建；其余固定 sheet 全部在此生成，作为可外部编辑的「外壳」持久化到 audit_templates/。"""
    for y in years:
        name = f"银行存款明细表_2025" if str(y) == "2025" else f"银行存款明细表_{y}"
        wb.create_sheet(name)
    wb.create_sheet("期间对比")
    wb.create_sheet("对方科目核对")


def build_bank_detail_km_combined(data_dir, output_path, period_mode="Y"):
    """目录模式（多账套×多年份）：扫描目录下所有核算主体的 科目余额表/外币科余表/综合查询明细表，
    合并输出一个工作簿：
      · 每个年度一张『明细_<年>』表（含『核算主体』列）；
      · 『期间对比』（核算主体×年度）；
      · 『对方科目核对』（GL 抽取分录按 核算主体/年度/账户/对方科目 归集，含异常标记）；
      · 『异常』（控制数对平、平衡、对方科目、取数范围等全部关注项）。
    主表数字取自各主体科目余额表 1002 + 外币原币；综合查询明细表(GL)仅下钻与对方科目分析。"""
    # ⚡⚡ 2026-08-15 账套级银行存款科目码注入（XBJ=1002+1004 内部结算中心存款）
    _set_bank_codes(_acct_of(data_dir))
    entities = _discover_bank_entities(data_dir)
    if not entities:
        print(f"❌ 目录内未找到含『科目余额表』的账套文件：{data_dir}")
        return None
    all_years = set()
    detail_rows = []   # 每个银行账户一行（含 entity / year）
    gl_rows = []       # read_gl_bank 行 + entity / year / opp
    issues = []
    parent_totals = {}  # (entity, year) -> parent dict(控制数)

    for entity, bun in entities.items():
        km_years = bun["km"]
        fc_by_year = bun["fc"]
        gl_by_year = bun["gl"]
        for y, km_path in km_years.items():
            all_years.add(y)
            # ⚡ 2026-08-11 P0：集团模式 current_comp=None → adapter.read_km/read_gl 返回空
            # → 期间对比/明细全 0（300 集团银行存款底稿）。逐主体 set_comp 后读取。
            if _adapter is not None and _adapter.is_sap(data_dir):
                if 'bank_saved_comp' not in dir():
                    bank_saved_comp = _adapter.current_comp()   # ⚡ 2026-08-11 修复：保存调用前状态
                _adapter.set_comp(entity)
            accounts = read_km_bank_accounts(km_path)
            _p1002, _ = accounts.get("1002", (None, []))
            _p1012, _ = accounts.get("1012", (None, []))
            if _p1002 is None and _p1012 is None:
                issues.append({"level": "WARN", "type": "控制数",
                               "message": f"[{entity}][{y}] 银行存款/其他货币资金科目余额表解析失败"
                                          f"（表头缺失或非标准格式，如仅有『币别:人民币』前导行），"
                                          f"该主体银行账户未纳入明细表，请核对并重导源文件。"})
            fc_list = read_fc_bank_accounts(fc_by_year.get(y)) if fc_by_year.get(y) else []
            fc_map = {f["code"]: f for f in fc_list}
            parent_totals[(entity, y)] = accounts
            detail_rows.extend(_km_subject_rows(accounts, fc_map, entity, y, issues))
            g = gl_by_year.get(y)
            if g:
                rr = read_gl_bank(g, [])
                for x in rr:
                    x["entity"] = entity
                    x["year"] = y
                gl_rows += rr
            if not fc_by_year.get(y):
                issues.append({"level": "INFO", "type": "外币原币",
                               "message": f"[{entity}][{y}] 无外币科目余额表，该主体外币账户原币列留空。"})
        # ⚡ 2026-08-11 修复：恢复【调用前】current_comp（非硬置 None——曾导致后续
        #   current_account 驱动 current_comp=None → 误处理全集团 88 家 → 40 分钟+口径错）
        if _adapter is not None and _adapter.is_sap(data_dir):
            try:
                _adapter.set_comp(bank_saved_comp)
            except Exception:
                pass

    # ---- 凭证级对方科目推导（GL 对方科目列可能为空，通过同凭证号借贷匹配推导） ----
    sale_rows = []      # 销售精确汇总：银行借方(收款)凭证内 贷方=6销售科目 金额（凭证级行金额）
    purch_rows = []     # 采购精确汇总：银行贷方(付款)凭证内 借方=存货/预付/应付 金额（凭证级行金额）
    remain_rows = []    # 剩余对方科目：原对方科目核对 剔除 销售/采购 后剩余（凭证级对方科目行归集）
    if gl_rows:
        from audit_common import discover_entities as _de_bank_cp, build_gl_voucher_map, derive_counterparty
        if _adapter is not None and _adapter.is_sap(data_dir):
            # ⚡ 2026-08-12 修复：audit_common.discover_entities 不识别 SAP 布局 →
            #   _all_ents_cp 空 → _vmap_cp 空 → 对方科目推导跳过 → 对方科目核对/
            #   销售精确汇总/采购精确汇总/剩余对方科目核对 4 表空（质检 6 主体 ERROR）。
            #   SAP 必须用 adapter.discover_entities（读 科目余额表/ 子目录）。
            _c_cp = _adapter.current_comp()
            _ents_cp = _adapter.discover_entities(data_dir)
            _all_ents_cp = ({_c_cp: _ents_cp.get(_c_cp, {})} if _c_cp and _c_cp in _ents_cp else _ents_cp)
        else:
            _all_ents_cp = _de_bank_cp(data_dir)
        if _all_ents_cp:
            _vmap_cp = build_gl_voucher_map(data_dir, _all_ents_cp)
            if _vmap_cp:
                _n_derived = 0
                for _r in gl_rows:
                    _vk = f"{_r.get('vchar','')}-{_r.get('vno','')}".strip("-")
                    if not _vk:
                        continue
                    _lines = _vmap_cp.get((_r["entity"], _r["year"], _vk))
                    if not _lines:
                        _lines = _vmap_cp.get((_r["entity"], _r["year"], _r.get("vno","")))
                    if not _lines:
                        continue
                    _dr = float(_r["amount"]) if _r["direction"] == "借" else 0.0
                    _cr = float(_r["amount"]) if _r["direction"] == "贷" else 0.0
                    _cp = derive_counterparty(_lines, _r["account"], _dr, _cr)
                    if _cp:
                        _r["opp"] = _cp
                        _n_derived += 1
                if _n_derived > 0:
                    issues.append({"level": "INFO", "type": "对方科目",
                                   "message": f"凭证级对方科目推导：为 {_n_derived}/{len(gl_rows)} 条银行存款 GL 分录推导出对方科目（替代空白列）。"})
                # ---- 销售/采购/剩余 凭证级精确拆分（2026-08-01 新增）----
                # 销售6科目：应收票据/应收账款/预收款项/合同负债/主营业务收入/其他业务收入（按名称含关键词，铁律23）
                # 采购类：存货(材料采购/原材料/库存商品/周转材料/生产成本/半成品等) + 预付账款 + 应付账款/应付票据
                _SALE_KW = ('应收票据', '应收账款', '合同负债', '主营业务收入', '其他业务收入', '预收')
                _PURCH_INV_KW = ('材料采购', '原材料', '在途', '库存商品', '周转材料', '包装物', '低值易耗',
                                 '生产成本', '半成品', '产成品', '发出商品', '委托加工', '存货')
                _PURCH_PRE_KW = ('预付',)
                _PURCH_PAY_KW = ('应付账款', '应付票据')
                _sale = {}    # (e,y,acct,cls) -> amt
                _purch = {}   # (e,y,acct,cls) -> amt
                _rem = {}     # (e,y,acct,opp) -> [n_d,s_d,n_c,s_c]
                # 口径与原『对方科目核对』同源：只处理 gl_rows(银行分录) 实际出现的凭证集合，
                # 银行行识别用 gl_rows 的 account 字段（而非全量 read_gl_rows 名称匹配）。
                # ⚡ 2026-08-12 修复：from-import 绑定在 patch 前固定为 U8 原版 read_gl_rows
                #   → SAP 下 _gl_rows_all 恒空 → 销售/采购/剩余 3 表空。SAP 用 adapter 版。
                if _adapter is not None and _adapter.is_sap(data_dir):
                    _gl_rows_all = _adapter.read_gl_rows(data_dir, _all_ents_cp)
                else:
                    from audit_common import read_gl_rows as _rgr_bank
                    _gl_rows_all = _rgr_bank(data_dir, _all_ents_cp)
                _vouch_grp = {}
                for _gr in _gl_rows_all:
                    _gvk = f"{_gr.get('vtype','')}-{_gr.get('vno','')}".strip('-')
                    if not _gvk:
                        continue
                    _k = (_gr.get('e'), _gr.get('y'), _gvk)
                    _vouch_grp.setdefault(_k, []).append(
                        {'name': _gr.get('name', '') or '', 'dr': float(_gr.get('dr') or 0),
                         'cr': float(_gr.get('cr') or 0)})
                # 原表同源凭证集合：gl_rows 的 (entity, year, vchar-vno)
                # ⚡ 2026-08-12 修复：键必须与 _vouch_grp 完全一致——_vouch_grp 用
                #   (_gr['e'], _gr['y'], vtype-vno)，而 gl_rows 是 {entity, year, vno}。
                #   原 f"{vchar}-{vno}".strip('-') SAP 下拼接错位 + e/y 键名不一致
                #   → 恒不匹配 → 销售/采购/剩余 3 表空。
                _bank_vouch_set = set()
                for _r in gl_rows:
                    _gk = f"{_r.get('vtype','')}-{_r.get('vno','')}".strip('-')
                    if not _gk:
                        continue
                    _bank_vouch_set.add((_r["entity"], str(_r["year"]), _gk))
                # 银行账户名集合：gl_rows account 归一化（与原表 _bank_norm 同源，含 acc 兜底）
                _bank_norm_set = set()
                for _d in detail_rows:
                    if _d.get("name"):
                        _bank_norm_set.add(re.sub(r"[\s\-－_/]", "", str(_d["name"])))
                    if _d.get("acc"):
                        _bank_norm_set.add(re.sub(r"[\s\-－_/]", "", str(_d["acc"])))

                def _is_bank_line(nm):
                    o = re.sub(r"[\s\-－_/]", "", str(nm or ""))
                    if o in _bank_norm_set:
                        return True
                    # 跨账套账户名前缀差异（J 账套 GL 行名为『银行存款-银行存款-工行…』双重前缀）：
                    # 归一化后仍可通过 结尾包含/开头包含 双向匹配（精确账户号唯一，不会误伤损益科目）
                    for a in _bank_norm_set:
                        if len(a) >= 6 and (o.endswith(a) or a.endswith(o)):
                            return True
                    return False

                def _cls_sale(nm):
                    s = str(nm or "")
                    for kw in _SALE_KW:
                        if kw in s:
                            return kw
                    return None

                def _cls_purch(nm):
                    s = str(nm or "")
                    for kw in _PURCH_INV_KW:
                        if kw in s:
                            return '存货'
                    for kw in _PURCH_PRE_KW:
                        if kw in s:
                            return '预付'
                    for kw in _PURCH_PAY_KW:
                        if kw in s:
                            return '应付'
                    return None

                for (_e, _y, _vk), _lines in _vouch_grp.items():
                    if not _lines:
                        continue
                    if (_e, _y, _vk) not in _bank_vouch_set:
                        continue
                    # 凭证内银行行（借方=收款 / 贷方=付款）与对方科目行
                    _bank_lines = [l for l in _lines if _is_bank_line(l.get('name'))]
                    if not _bank_lines:
                        continue
                    _acct = _bank_lines[0].get('name', '')
                    if len({re.sub(r"[\s\-－_/]", "", str(l.get('name', ''))) for l in _bank_lines}) > 1:
                        _acct = '（多银行账户）'
                    # 收款侧：银行行在借方 → 贷方对方科目行 归 销售/剩余
                    if any(l.get('dr', 0) > 0 for l in _bank_lines):
                        for l in _lines:
                            if _is_bank_line(l.get('name')):
                                continue
                            cr = l.get('cr', 0) or 0.0
                            if cr <= 0:
                                continue
                            _cl = _cls_sale(l.get('name'))
                            if _cl:
                                _k = (_e, _y, _acct, _cl)
                                _sale[_k] = _sale.get(_k, 0.0) + cr
                            else:
                                _ko = (_e, _y, _acct, l.get('name', '（空）'))
                                _d = _rem.setdefault(_ko, [0, 0.0, 0, 0.0])
                                _d[0] += 1; _d[1] += cr
                    # 付款侧：银行行在贷方 → 借方对方科目行 归 采购/剩余
                    if any(l.get('cr', 0) > 0 for l in _bank_lines):
                        for l in _lines:
                            if _is_bank_line(l.get('name')):
                                continue
                            dr = l.get('dr', 0) or 0.0
                            if dr <= 0:
                                continue
                            _cp_ = _cls_purch(l.get('name'))
                            if _cp_:
                                _k = (_e, _y, _acct, _cp_)
                                _purch[_k] = _purch.get(_k, 0.0) + dr
                            else:
                                _ko = (_e, _y, _acct, l.get('name', '（空）'))
                                _d = _rem.setdefault(_ko, [0, 0.0, 0, 0.0])
                                _d[2] += 1; _d[3] += dr
                sale_rows = [{"entity": k[0], "year": k[1], "account": k[2], "cls": k[3], "amt": v}
                             for k, v in sorted(_sale.items())]
                purch_rows = [{"entity": k[0], "year": k[1], "account": k[2], "cls": k[3], "amt": v}
                              for k, v in sorted(_purch.items())]
                remain_rows = [{"entity": k[0], "year": k[1], "account": k[2], "opp": k[3],
                                "n_d": v[0], "s_d": v[1], "n_c": v[2], "s_c": v[3]}
                               for k, v in sorted(_rem.items())]
                if sale_rows or purch_rows:
                    issues.append({"level": "INFO", "type": "对方科目",
                                   "message": f"凭证级精确拆分：销售收款科目 {len(sale_rows)} 组、采购付款科目 {len(purch_rows)} 组、"
                                              f"剩余对方科目 {len(remain_rows)} 组（按银行行所在凭证的对方科目行金额归集，非银行行全额）。"})

    years = sorted(all_years)

    # 对方科目归集 + 异常识别（基于 GL 已抽取分录）
    cp = {}
    acct_tot = {}
    for r in gl_rows:
        opp = (r.get("opp") or "").strip()
        key = (r["entity"], r["year"], r["account"], opp)
        d = cp.setdefault(key, {"cnt": [0, 0.0, 0, 0.0], "rows": []})
        if r["direction"] == "借":
            d["cnt"][0] += 1; d["cnt"][1] += r["amount"]
        else:
            d["cnt"][2] += 1; d["cnt"][3] += r["amount"]
        d["rows"].append(r)
        at = acct_tot.setdefault((r["entity"], r["year"], r["account"]), [0.0, 0.0])
        if r["direction"] == "借":
            at[0] += r["amount"]
        else:
            at[1] += r["amount"]

    cp_rows = []
    for (entity, year, account, opp), d in cp.items():
        n_d, s_d, n_c, s_c = d["cnt"]
        td, tc = acct_tot.get((entity, year, account), [0.0, 0.0])
        pct_d = (abs(s_d) / abs(td) * 100) if td else 0.0
        pct_c = (abs(s_c) / abs(tc) * 100) if tc else 0.0
        flags = []
        if not opp:
            flags.append("对方科目缺失")
        elif _is_fund_counterparty(opp):
            flags.append("资金类对方科目(疑似内部划转)")
        if max(pct_d, pct_c) > 80:
            flags.append("单边流量集中度过高(>80%)")
        note = "；".join(flags)
        # 对方科目缺失不单独告警（已在对方科目核对 sheet 排除，并纳入异常凭证抽取）
        if any(f != "对方科目缺失" for f in flags):
            issues.append({"level": "WARN", "type": "对方科目",
                           "message": f"[{entity}][{year}] 账户『{account}』对方科目『{opp or '（空）'}』："
                                      f"借方笔数{n_d}/金额{s_d:,.2f}、贷方笔数{n_c}/金额{s_c:,.2f}，{note}。"})
        cp_rows.append({
            "entity": entity, "year": year, "account": account, "opp": opp or "（空）",
            "n_d": n_d, "s_d": s_d, "n_c": n_c, "s_c": s_c,
            "net": s_d - s_c, "pct_d": pct_d, "pct_c": pct_c, "note": note,
            "flags": flags, "rows": d["rows"],
        })

    # 取数范围披露
    if not gl_rows:
        issues.append({"level": "INFO", "type": "取数范围",
                       "message": "未提供综合查询明细表，无法生成『对方科目核对』，主表仍以科目余额表控制数为准。"})

    # 需求4：外币是否需从『辅助核算余额表』取数？
    aux_path = _discover_aux_balance(data_dir)
    if aux_path:
        n_aux = _scan_aux_bank(aux_path)
        if n_aux == 0:
            issues.append({"level": "INFO", "type": "外币原币",
                           "message": f"已扫描『{os.path.basename(aux_path)}』，其中未含任何银行存款/其他货币资金账户行"
                                      f"（且通常无原币列），外币账户原币仍取自『外币科目余额表』；需求4(从辅助核算余额表取外币)对本数据集不适用。"})
        else:
            issues.append({"level": "INFO", "type": "外币原币",
                           "message": f"『{os.path.basename(aux_path)}』检出 {n_aux} 行银行存款/其他货币资金账户，"
                                      f"如需从辅助核算余额表取外币原币可据此补充（当前版本优先采用『外币科目余额表』）。"})
    else:
        issues.append({"level": "INFO", "type": "外币原币",
                       "message": "未提供『辅助核算余额表』，外币账户原币取自『外币科目余额表』。"})

    # 结构外壳：优先加载中央 audit_templates/ 的「银行存款_外壳.xlsx」（可外部编辑 sheet 结构）；
    # 缺失则按 _build_bank_frame 现建并落盘，便于今后改排版只动外壳、不动代码。
    shell = resolve_template('银行存款_外壳', 'BANK_TEMPLATE')
    if shell:
        wb = openpyxl.load_workbook(shell)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        _build_bank_frame(wb, years)
        try:
            os.makedirs(AUDIT_TEMPLATES_DIR, exist_ok=True)
            finalize_workbook(wb)
            wb.save(os.path.join(AUDIT_TEMPLATES_DIR, '银行存款_外壳.xlsx'))
        except Exception as ex:
            print(f'  ⚠️ 外壳落盘失败（不影响本次生成）：{ex}')
    _out_dir = os.path.dirname(output_path) or data_dir
    _first = None
    for _y in sorted(years):
        if shell:
            _wb_y = openpyxl.load_workbook(shell)
        else:
            _wb_y = openpyxl.Workbook()
            _wb_y.remove(_wb_y.active)
        _out_y = os.path.join(_out_dir, f'银行存款审计底稿_{_y}_生成.xlsx')
        _r = _export_km_combined_excel(_out_y, years, detail_rows, cp_rows, issues,
                                       parent_totals, entities, data_dir, wb=_wb_y, target_year=_y,
                                       sale_rows=sale_rows, purch_rows=purch_rows, remain_rows=remain_rows)
        if _first is None:
            _first = _r
    output_path = _first
    return issues, years, detail_rows, output_path


def _read_km_capital(km_path):
    """读取科目余额表 1001/1002/1012 期末(带借贷符号)及按币种拆分。
    返回 dict: '1001'->[parent_close, {ccy: child_close_sum}], ...（集团/单主体均可）。
    parent_close 为父级(或子目合计)期末带符号数；child_close_sum 按币种(人民币/USD/EUR/JPY/HKD)归集。"""
    try:
        wb = openpyxl.load_workbook(km_path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    except Exception:
        return {}
    hdr_idx, hdr, ci, di, cri, oai, cai, odi, cdi = _km_find_header(raw)
    if hdr is None or ci is None:
        return {}
    def _parse_num(v):
        try:
            if v is None:
                return 0.0
            if isinstance(v, (int, float)):
                return float(v)
            s = str(v).strip().replace(",", "").replace("￥", "").replace("¥", "")
            return float(re.sub(r"[()]", "", s)) if s else 0.0
        except Exception:
            return 0.0
    def _sg(v, d):
        a = _parse_num(v)
        dd = _norm(d) if d is not None else ""
        return a if dd in ("借", "debit", "dr", "j") or dd == "" else -a
    ni = None
    for i, h in enumerate(hdr):
        if _norm(h) in set(_norm(a) for a in KM_COLUMN_ALIASES["name"]):
            ni = i
            break
    res = {c: [0.0, {}] for c in ("1001", "1002", "1012")}
    for row in raw[hdr_idx + 1:]:
        if not row or ci >= len(row) or row[ci] is None:
            continue
        code = _norm(row[ci]).replace(".0", "")
        if code not in ("1001", "1002", "1012") and not (
                code.startswith("1001") or code.startswith("1002") or code.startswith("1012")):
            continue
        nm = str(row[ni]).strip() if ni is not None and ni < len(row) and row[ni] is not None else code
        close = _sg(row[cai] if cai is not None and cai < len(row) else 0,
                   row[cdi] if cdi is not None and cdi < len(row) else "借")
        top = code[:4]
        if code in ("1001", "1002", "1012"):
            res[top][0] += close
        else:
            ccy = _detect_fc(nm) or "人民币"
            res[top][1][ccy] = res[top][1].get(ccy, 0.0) + close
    for c in res:
        pc, cm = res[c]
        if abs(pc) < 0.005 and cm:
            pc = sum(cm.values())
        res[c] = [pc, cm]
    return res


def _find_capital_template(data_dir):
    """在数据文件夹内查找『货币资金审定表*.xlsx』模板（跳过 Excel 锁文件）。"""
    if not data_dir or not os.path.isdir(data_dir):
        return None
    for fn in sorted(os.listdir(data_dir)):
        if fn.lower().endswith(".xlsx") and "货币资金审定表" in fn and not fn.startswith("~$"):
            return os.path.join(data_dir, fn)
    return None


def _build_capital_audit_sheet(wb, data_dir, entities, years, FONT10, detail_rows=None, parent_totals=None):
    '''货币资金审定表：按核算主体列示 库存现金+银行存款+其他货币资金。'''
    from audit_common import add_audit_summary_sheets as _add_bank_audit
    codes = ['1001', '1002', '1012']
    _add_bank_audit(wb, data_dir, [dict(title='货币资金审定表', codes=codes, is_credit=False,
                                         subj_name='货币资金', by_entity=True)])


def _build_capital_audit_3subj(wb, data_dir, target_year):
    """货币资金审定表（2026-08-02 用户要求）：
    先按三个科目（库存现金/银行存款/其他货币资金）列示——每科目下按核算主体行
    （年初数=TB期初 qc、期末未审数=TB期末 qm、审计调整待填、审定数=未审+调整公式）；
    再每科目小计；最后集团各科目合计（库存现金合计/银行存款合计/其他货币资金合计/货币资金总计）。
    年初/期末均取该科目 TB 叶子行（父=子和，防双计）；代码按 1001/1002/1012 前缀+名称识别（铁律23）。"""
    from audit_common import discover_entities as _de3, read_tb_full as _rt3
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    # SAP：单主体驱动时按 current_comp 过滤（防审定表混入 88 家主体行）
    if _adapter is not None and _adapter.is_sap(data_dir):
        _c3 = _adapter.current_comp()
        ents = {_c3: _de3(data_dir).get(_c3, {})} if _c3 else _de3(data_dir)
    else:
        ents = _de3(data_dir)
    tb_full = _rt3(data_dir, ents)
    ty = str(target_year)
    HFILL = PatternFill('solid', fgColor='DDEBF7')
    HFONT = Font(name='Times New Roman', bold=True, color='000000', size=10)
    TFILL = PatternFill('solid', fgColor='D9E1F2')
    TFONT = Font(name='Times New Roman', bold=True, size=10)
    GFILL = PatternFill('solid', fgColor='FFF2CC')
    thin = Side(style='thin', color='BFBFBF')
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    CEN = Alignment(horizontal='center', vertical='center', wrap_text=True)
    LFT = Alignment(horizontal='left', vertical='center')
    RGT = Alignment(horizontal='right', vertical='center')
    NUM = '#,##0.00'
    F10 = Font(name='Times New Roman', size=10)

    def cat_of(c, n):
        s = str(c); nm = str(n)
        if s.startswith('1001') or '库存现金' in nm or nm == '现金':
            return '库存现金'
        if s.startswith('1002') or '银行存款' in nm:
            return '银行存款'
        if s.startswith('1012') or '其他货币资金' in nm:
            return '其他货币资金'
        return None

    rows_by = {}   # (e, cat) -> [(code, v)]
    for (e, c, n, y), v in tb_full.items():
        if str(y) != ty:
            continue
        cat = cat_of(c, n)
        if cat:
            rows_by.setdefault((e, cat), []).append((str(c), v))

    def leaf_sum(items):
        qc = qm = 0.0
        leaves = [v for c, v in items
                  if not any(c != c2 and c2.startswith(c) and len(c2) > len(c) for (c2, _) in items)]
        # 2026-08-07 修复：原逐行 abs() 求和会把负余额账户翻正——S pot 银行存款
        # 外币美元/欧元户负余额（华夏 -6,987 万/南京 -1.06 亿等合计 -4.5 亿）→ 审定
        # 11.31 亿 vs TB 净额 6.81 亿（虚增 4.5 亿）。改取【带符号净额】再 abs：
        # 父=子和（铁律），负余额账户如实抵减（外币汇差/透支是真实状态）。
        qc = abs(sum(float(v['qc'] or 0.0) for v in leaves))
        qm = abs(sum(float(v['qm'] or 0.0) for v in leaves))
        return qc, qm

    cats = ('库存现金', '银行存款', '其他货币资金')
    ent_list = sorted({e for (e, _) in rows_by})
    if '货币资金审定表' in wb.sheetnames:
        del wb['货币资金审定表']
    ws = wb.create_sheet('货币资金审定表')
    ws.cell(1, 1, f'货币资金 审定表（按核算主体 · {ty} 年）').font = Font(name='Times New Roman', bold=True, size=12)
    for c, h in enumerate(['科目', '核算主体', '年初数', '期末未审数', '审计调整数', '审定数'], 1):
        cell = ws.cell(2, c, h); cell.font = HFONT; cell.fill = HFILL
        cell.alignment = CEN; cell.border = BORDER
    r = 3
    cat_total = {c: [0.0, 0.0] for c in cats}
    cat_span = {}   # cat -> (主体区首行, 主体区末行)  用于小计/合计行公式（2026-08-02 合计保留公式）
    for cat in cats:
        ws.cell(r, 1, cat).font = TFONT; ws.cell(r, 1).fill = TFILL
        for c in range(1, 7):
            ws.cell(r, c).border = BORDER
        r += 1
        body_first = r
        sub = [0.0, 0.0]
        for e in ent_list:
            items = rows_by.get((e, cat))
            if not items:
                continue
            qc, qm = leaf_sum(items)
            if abs(qc) < 0.005 and abs(qm) < 0.005:
                continue
            ws.cell(r, 1, cat); ws.cell(r, 2, e)
            ws.cell(r, 3, round(qc, 2)); ws.cell(r, 4, round(qm, 2))
            ws.cell(r, 5, None)
            # 2026-08-06 协议：审定数写数值（=期末+调整0；公式 data_only 读 None）
            ws.cell(r, 6, round(qm, 2))
            for c in range(1, 7):
                ws.cell(r, c).border = BORDER; ws.cell(r, c).font = F10
            for c in (3, 4, 5, 6):
                ws.cell(r, c).number_format = NUM; ws.cell(r, c).alignment = RGT
            ws.cell(r, 1).alignment = LFT; ws.cell(r, 2).alignment = LFT
            sub[0] += qc; sub[1] += qm
            r += 1
        # 每科目小计（2026-08-06 协议：写数值=sub 累计）
        sub_r = r
        ws.cell(r, 1, f'{cat} 小计'); ws.cell(r, 2, '')
        ws.cell(r, 3, round(sub[0], 2)); ws.cell(r, 4, round(sub[1], 2))
        ws.cell(r, 5, None); ws.cell(r, 6, round(sub[1], 2))
        for c in range(1, 7):
            ws.cell(r, c).border = BORDER; ws.cell(r, c).font = TFONT; ws.cell(r, c).fill = TFILL
        for c in (3, 4, 5, 6):
            ws.cell(r, c).number_format = NUM; ws.cell(r, c).alignment = RGT
        ws.cell(r, 1).alignment = LFT
        cat_total[cat] = sub
        cat_span[cat] = (body_first, r - 1)
        r += 1
    # 集团各科目合计（2026-08-06 协议：写数值=cat_total）
    ws.cell(r, 1, '集团各科目合计').font = TFONT; ws.cell(r, 1).fill = GFILL
    for c in range(1, 7):
        ws.cell(r, c).border = BORDER
    r += 1
    grand = [0.0, 0.0]
    for cat in cats:
        bf, bl = cat_span.get(cat, (r, r))
        ws.cell(r, 1, cat + ' 合计'); ws.cell(r, 2, '')
        ws.cell(r, 3, round(cat_total[cat][0], 2)); ws.cell(r, 4, round(cat_total[cat][1], 2))
        ws.cell(r, 5, None); ws.cell(r, 6, round(cat_total[cat][1], 2))
        for c in range(1, 7):
            ws.cell(r, c).border = BORDER; ws.cell(r, c).font = F10
        for c in (3, 4, 5, 6):
            ws.cell(r, c).number_format = NUM; ws.cell(r, c).alignment = RGT
        grand[0] += cat_total[cat][0]; grand[1] += cat_total[cat][1]
        r += 1
    # 货币资金 总计（2026-08-06 协议：写数值=grand）
    first_cat_r = r - 3
    ws.cell(r, 1, '货币资金 总计'); ws.cell(r, 2, '')
    ws.cell(r, 3, round(grand[0], 2)); ws.cell(r, 4, round(grand[1], 2))
    ws.cell(r, 5, None); ws.cell(r, 6, round(grand[1], 2))
    for c in range(1, 7):
        ws.cell(r, c).border = BORDER; ws.cell(r, c).font = TFONT; ws.cell(r, c).fill = GFILL
    for c in (3, 4, 5, 6):
        ws.cell(r, c).number_format = NUM; ws.cell(r, c).alignment = RGT
    ws.cell(r, 1).alignment = LFT
    for c, w in enumerate([16, 18, 16, 16, 13, 13], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A3'
    return ws


def _fill_capital_audit_sheet(wb, tmpl_path, entities, years, FONT10):
    """复制『货币资金审定表』模板结构，期末未审数(B列)由科目余额表(1001/1002/1012)及子目
    按币种拆分(集团合计)取数；其他列(审计调整/审定数)留空。目标年默认 2025。"""
    try:
        # 注意：read_only=True 下 ReadOnlyWorksheet 无 merged_cells/column_dimensions 属性，
        # 会触发 AttributeError 被下方 except 吞掉导致整表跳过；模板仅 24 行，直接普通加载。
        twb = openpyxl.load_workbook(tmpl_path, data_only=True)
        tws = twb[twb.sheetnames[0]]
        trows = [list(r) for r in tws.iter_rows(values_only=True)]
        merged = list(tws.merged_cells.ranges)
        col_dims = {k: v.width for k, v in tws.column_dimensions.items() if v.width}
        twb.close()
    except Exception:
        return
    target_year = "2025" if "2025" in years else (years[0] if years else None)
    agg = {c: [0.0, {}] for c in ("1001", "1002", "1012")}
    if target_year:
        for ent, bun in entities.items():
            kp = bun.get("km", {}).get(target_year)
            if not kp:
                continue
            rc = _read_km_capital(kp)
            for code in ("1001", "1002", "1012"):
                pc, cm = rc.get(code, (0.0, {}))
                agg[code][0] += pc
                for ccy, v in cm.items():
                    agg[code][1][ccy] = agg[code][1].get(ccy, 0.0) + v
        for code in agg:
            pc, cm = agg[code]
            if abs(pc) < 0.005 and cm:
                pc = sum(cm.values())
            agg[code] = [pc, cm]
    cash_total, cash_cm = agg["1001"]
    bank_total, bank_cm = agg["1002"]
    other_total, other_cm = agg["1012"]
    grand = cash_total + bank_total + other_total

    def rmb_portion(total, cm):
        fc = sum(v for k, v in cm.items() if k != "人民币")
        return total - fc
    def cur(cm, ccy):
        return cm.get(ccy, 0.0)
    sec_vals = {
        "库存现金": (cash_total, rmb_portion(cash_total, cash_cm), cur(cash_cm, "USD"), cur(cash_cm, "EUR")),
        "银行存款": (bank_total, rmb_portion(bank_total, bank_cm), cur(bank_cm, "USD"),
                     cur(bank_cm, "EUR"), cur(bank_cm, "JPY")),
        "其他货币资金": (other_total, rmb_portion(other_total, other_cm), cur(other_cm, "USD"),
                      cur(other_cm, "EUR")),
    }

    ws = wb["货币资金审定表"] if "货币资金审定表" in wb.sheetnames else wb.create_sheet("货币资金审定表")
    for ri, row in enumerate(trows, start=1):
        for ci2, val in enumerate(row, start=1):
            ws.cell(row=ri, column=ci2, value=val)
    sec = None
    for ri in range(1, len(trows) + 1):
        name = ws.cell(row=ri, column=1).value
        name_s = str(name) if name is not None else ""
        if "库存现金" in name_s and "本位币" in name_s:
            sec = "库存现金"; ws.cell(row=ri, column=2, value=sec_vals["库存现金"][0])
            ws.cell(row=ri, column=2).number_format = "#,##0.00"
        elif "银行存款" in name_s and "本位币" in name_s:
            sec = "银行存款"; ws.cell(row=ri, column=2, value=sec_vals["银行存款"][0])
            ws.cell(row=ri, column=2).number_format = "#,##0.00"
        elif "其他货币资金" in name_s and "本位币" in name_s:
            sec = "其他货币资金"; ws.cell(row=ri, column=2, value=sec_vals["其他货币资金"][0])
            ws.cell(row=ri, column=2).number_format = "#,##0.00"
        elif "合计" in name_s and "本位币" in name_s:
            ws.cell(row=ri, column=2, value=grand); ws.cell(row=ri, column=2).number_format = "#,##0.00"
        elif sec:
            if "人民币" in name_s:
                ws.cell(row=ri, column=2, value=sec_vals[sec][1]).number_format = "#,##0.00"
            elif "美元" in name_s:
                ws.cell(row=ri, column=2, value=sec_vals[sec][2]).number_format = "#,##0.00"
            elif "欧元" in name_s:
                ws.cell(row=ri, column=2, value=sec_vals[sec][3]).number_format = "#,##0.00"
            elif "日元" in name_s:
                jv = sec_vals[sec][4] if len(sec_vals[sec]) > 4 else 0.0
                ws.cell(row=ri, column=2, value=jv).number_format = "#,##0.00"
        # 保留样表全部列结构（审计调整借/贷、期末审定数、期初审定数及其表头/标签），
        # 仅「期末未审数」(B列)由科目余额表取数覆盖；其余列按用户要求暂不用取数，
        # 沿用样表原有占位内容（多为 0），不再清空以免整列塌陷丢失格式。
    for m in merged:
        try:
            ws.merge_cells(str(m))
        except Exception:
            pass
    for k, w in col_dims.items():
        try:
            ws.column_dimensions[k].width = w
        except Exception:
            pass
    for row in ws.iter_rows():
        for cell in row:
            if cell.font is None or cell.font.size != 10:
                cell.font = FONT10
    # 置于工作簿首位
    order = ["货币资金审定表"] + [s.title for s in wb.worksheets if s.title != "货币资金审定表"]
    wb._sheets.sort(key=lambda s: order.index(s.title))


def _export_km_combined_excel(output_path, years, detail_rows, cp_rows, issues, parent_totals, entities, data_dir, wb=None, target_year=None,
                              sale_rows=None, purch_rows=None, remain_rows=None):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    wb = wb or openpyxl.Workbook()
    thin = Side(style="thin", color="BBBBBB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="DDEBF7")
    hdr_font = Font(name='Times New Roman', bold=True, color="000000", size=10)
    FONT10 = Font(name='Times New Roman', size=10)
    BOLD10 = Font(name='Times New Roman', bold=True, size=10)
    total_fill = PatternFill("solid", fgColor="DDEBF7")
    grand_fill = PatternFill("solid", fgColor="BDD7EE")
    warn_fill = PatternFill("solid", fgColor="FFEB9C")
    err_fill = PatternFill("solid", fgColor="FFC7CE")
    money = "#,##0.00"

    def get_sheet(name):
        """外壳已含该 sheet 则返回，否则新建（实现「外壳优先、动态 sheet 兜底」）。
        注意：外壳模板的空 sheet 在 openpyxl 重载后会出现一行 phantom 空行(max_row=1 但无值)，
        导致后续 ws.append 写到第 2 行、表头错位。若第 1 行全空则删除以使表头回到第 1 行。"""
        ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        if ws.max_row >= 1:
            mc = max(ws.max_column, 1)
            if all(ws.cell(row=1, column=c).value is None for c in range(1, mc + 1)):
                ws.delete_rows(1, 1)
        return ws

    def style_header(ws, ncol, row=1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = hdr_fill; cell.font = hdr_font; cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def total_row(ws, r, ncol, vals, fill=total_fill):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c); cell.border = border
            cell.font = BOLD10; cell.fill = fill
        for c, v in vals.items():
            ws.cell(row=r, column=c, value=v)
            if isinstance(v, (int, float)) and c in MONEY_COLS:
                ws.cell(row=r, column=c).number_format = money

    # ---- Sheet1..N 明细_<年>（按年模式只建 target_year）----
    for y in ([target_year] if target_year is not None else years):
        sheet_name = "银行存款明细表2025" if str(y) == "2025" else f"银行存款明细表{y}"
        ws = get_sheet(sheet_name)
        # 序号在核算主体之前；本币各列前插入对应原币列；表尾追加「期末余额汇兑测算(4)」与「对账单/函证(5)」；
        # 末尾追加『科目代码/科目名称』列以区分 1002 银行存款 与 1012 其他货币资金。
        cols = ["序号", "核算主体", "开户银行", "账号", "账户性质", "币种",
                "期初原币", "期初余额(本币)", "借方原币", "本期借方(本币)",
                "贷方原币", "本期贷方(本币)", "期末原币", "期末余额(本币)",
                "平衡校验", "备注",
                "原币(汇兑测算)", "期末汇率", "应计本位币", "汇兑差异",
                "对账单金额", "对账单索引", "对账单差异", "余额调节表索引", "函证确认金额",
                "科目代码", "科目名称"]
        MONEY_COLS = {7, 8, 9, 10, 11, 12, 13, 14, 17, 19, 20, 21, 23, 25}
        ws.append(cols)
        style_header(ws, len(cols))

        def _emit(ws, r, rows, seq_start):
            for i, d in enumerate(rows, seq_start):
                bal = d["open"] + d["debit"] - d["credit"] - d["close"]
                ws.append([i, d.get("entity", ""), d["bank"], d["acc"], d["性质"], d["ccy"],
                           d["open_fc"], d["open"], d["debit_fc"], d["debit"],
                           d["credit_fc"], d["credit"], d["close_fc"], d["close"],
                           round(bal, 2), "",
                           d["close_fc"], "", None, None,
                           None, "", None, "", None,
                           d["code"], d["subject_name"]])
                # 测算公式：应计本位币=原币×期末汇率；汇兑差异=应计本位币−期末余额(本币)；对账单差异=对账单金额−期末余额(本币)
                ws.cell(row=r, column=19, value='=IF(R{r}="","",Q{r}*R{r})'.format(r=r))
                ws.cell(row=r, column=20, value='=IF(S{r}="","",S{r}-N{r})'.format(r=r))
                ws.cell(row=r, column=23, value='=IF(U{r}="","",U{r}-N{r})'.format(r=r))
                for c in range(1, len(cols) + 1):
                    cell = ws.cell(row=r, column=c); cell.border = border
                    cell.font = FONT10
                    if c in MONEY_COLS:
                        cell.number_format = money
                    if abs(bal) > 0.005 and c == 15:
                        cell.fill = warn_fill
                r += 1
            return r

        def _blk_total(ws, r, label, rows, subj, fill):
            so = sum(d["open"] for d in rows); sd = sum(d["debit"] for d in rows)
            sc = sum(d["credit"] for d in rows); sl = sum(d["close"] for d in rows)
            sf_o = sum(d["open_fc"] or 0 for d in rows); sf_d = sum(d["debit_fc"] or 0 for d in rows)
            sf_c = sum(d["credit_fc"] or 0 for d in rows); sf_l = sum(d["close_fc"] or 0 for d in rows)
            kv = {"1001": "库存现金", "1002": "银行存款", "1012": "其他货币资金", "": "货币资金"}
            total_row(ws, r, len(cols), {1: "", 2: "", 3: label, 7: sf_o, 8: so, 9: sf_d, 10: sd,
                                         11: sf_c, 12: sc, 13: sf_l, 14: sl, 17: sf_l,
                                         26: subj, 27: kv.get(subj, "")}, fill=fill)
            return r + 1

        r = 2
        yrows = [d for d in detail_rows if d["year"] == y]
        bank_rows = [d for d in yrows if d["subject"] == "1002"]
        other_rows = [d for d in yrows if d["subject"] == "1012"]
        # 一、银行存款（1002）
        r = _emit(ws, r, bank_rows, 1)
        r = _blk_total(ws, r, "银行存款合计", bank_rows, "1002", total_fill)
        # 二、其他货币资金（1012）
        if other_rows:
            r = _emit(ws, r, other_rows, len(bank_rows) + 1)
            r = _blk_total(ws, r, "其他货币资金合计", other_rows, "1012", total_fill)
            # 三、货币资金合计（1002 + 1012）
            r = _blk_total(ws, r, "货币资金合计", yrows, "", grand_fill)

        # ---- 年度级勾稽：明细表期末合计 vs 科目余额表父级期末 ----
        bank_comp = sum(d["close"] for d in bank_rows)
        other_comp = sum(d["close"] for d in other_rows)
        bank_ctrl = 0.0; other_ctrl = 0.0
        for (e, yy), acc in parent_totals.items():
            if yy == y:
                p1 = acc.get("1002"); p2 = acc.get("1012")
                if p1 and p1[0] is not None:
                    bank_ctrl += p1[0]["close"]
                if p2 and p2[0] is not None:
                    other_ctrl += p2[0]["close"]
        if abs(bank_comp - bank_ctrl) > 0.005:
            issues.append({"level": "WARN", "type": "勾稽",
                           "message": f"{y} 银行存款明细表期末合计({bank_comp:,.2f})与科目余额表银行存款父级期末合计({bank_ctrl:,.2f})差异{bank_comp - bank_ctrl:,.2f}。"})
        else:
            issues.append({"level": "INFO", "type": "勾稽",
                           "message": f"{y} 银行存款明细表期末合计({bank_comp:,.2f})与科目余额表银行存款父级勾稽相符。"})
        if other_rows:
            if abs(other_comp - other_ctrl) > 0.005:
                issues.append({"level": "WARN", "type": "勾稽",
                               "message": f"{y} 其他货币资金明细表期末合计({other_comp:,.2f})与科目余额表其他货币资金父级期末合计({other_ctrl:,.2f})差异{other_comp - other_ctrl:,.2f}。"})
            else:
                issues.append({"level": "INFO", "type": "勾稽",
                               "message": f"{y} 其他货币资金明细表期末合计({other_comp:,.2f})与科目余额表其他货币资金父级勾稽相符。"})

        widths = [6, 14, 22, 18, 12, 8, 14, 16, 14, 16, 14, 16, 14, 16, 12, 16,
                  14, 10, 16, 14, 16, 12, 14, 16, 16, 12, 14]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"

    # ---- 期间对比 ----
    ws2 = get_sheet("期间对比")
    cols2 = ["核算主体", "会计期间", "期初余额(本币)", "本期借方(本币)", "本期贷方(本币)",
             "期末余额(本币)", "外币账户数"]
    MONEY_COLS = {3, 4, 5, 6}
    ws2.append(cols2)
    style_header(ws2, len(cols2))
    r = 2
    for y in years:
        for entity in sorted(entities.keys()):
            erows = [d for d in detail_rows if d["year"] == y and d["entity"] == entity]
            if not erows:
                continue
            so = sum(d["open"] for d in erows); sd = sum(d["debit"] for d in erows)
            sc = sum(d["credit"] for d in erows); sl = sum(d["close"] for d in erows)
            nfc = sum(1 for d in erows if d["性质"] == "外币账户")
            ws2.append([entity, y, so, sd, sc, sl, nfc])
            for c in range(1, len(cols2) + 1):
                cell = ws2.cell(row=r, column=c); cell.border = border
                cell.font = FONT10
                if c in MONEY_COLS and isinstance(cell.value, (int, float)):
                    cell.number_format = money
            r += 1
        # 该年合计（跨主体）
        yr_rows = [d for d in detail_rows if d["year"] == y]
        so = sum(d["open"] for d in yr_rows); sd = sum(d["debit"] for d in yr_rows)
        sc = sum(d["credit"] for d in yr_rows); sl = sum(d["close"] for d in yr_rows)
        nfc = sum(1 for d in yr_rows if d["性质"] == "外币账户")
        total_row(ws2, r, len(cols2), {1: "合计", 2: y, 3: so, 4: sd, 5: sc, 6: sl, 7: nfc})
        r += 1
    # 总计
    to = sum(d["open"] for d in detail_rows); td = sum(d["debit"] for d in detail_rows)
    tc = sum(d["credit"] for d in detail_rows); tl = sum(d["close"] for d in detail_rows)
    total_row(ws2, r, len(cols2), {1: "总计", 2: "", 3: to, 4: td, 5: tc, 6: tl}, fill=grand_fill)
    for i, w in enumerate([14, 12, 18, 18, 18, 18, 12], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # （2026-08-02 用户要求：『期间差异对比』表已删除——期间对比按年分块已足够，跨年差异列无审计意义）

    # ---- 对方科目核对（剔除借贷双方均为银行存款的内部划转，并删除「净额(借-贷)」列） ----
    # 构建银行存款账户名称/账号集合，用于识别内部划转（对方科目命中银行存款账户）
    _bank_norm = set()
    for d in detail_rows:
        if d.get("name"):
            _bank_norm.add(re.sub(r"[\s\-－_/]", "", str(d["name"])))
        if d.get("acc"):
            _bank_norm.add(re.sub(r"[\s\-－_/]", "", str(d["acc"])))

    def _is_internal_transfer(opp):
        o = (opp or "").strip()
        if not o or o == "（空）":
            return False
        if "银行存款" in o or "1002" in o:
            return True
        return re.sub(r"[\s\-－_/]", "", o) in _bank_norm

    cp_rows_filtered = [r for r in cp_rows
                         if (not _is_internal_transfer(r["opp"])) and r["opp"] not in ("", "（空）")]
    n_excluded = len(cp_rows) - len(cp_rows_filtered)
    if n_excluded:
        issues.append({"level": "INFO", "type": "对方科目",
                       "message": f"『对方科目核对』已剔除 {n_excluded} 条（对方科目缺失或借贷双方均为银行存款的内部划转），不纳入对方科目勾稽。"})

    ws3 = get_sheet("对方科目核对")
    cols3 = ["核算主体", "期间", "银行账户", "对方科目", "借方笔数", "借方金额(本币)",
             "贷方笔数", "贷方金额(本币)", "占账户借方%", "占账户贷方%", "异常标记"]
    MONEY_COLS = {6, 8}
    ws3.append(cols3)
    style_header(ws3, len(cols3))
    r = 2
    for row in sorted(cp_rows_filtered, key=lambda x: (x["entity"], x["year"], x["account"], x["opp"])):
        if target_year is not None and str(row["year"]) != str(target_year):
            continue
        ws3.append([row["entity"], row["year"], row["account"], row["opp"],
                    row["n_d"], row["s_d"], row["n_c"], row["s_c"],
                    round(row["pct_d"], 2), round(row["pct_c"], 2), row["note"]])
        for c in range(1, len(cols3) + 1):
            cell = ws3.cell(row=r, column=c); cell.border = border
            cell.font = FONT10
            if c in MONEY_COLS and isinstance(cell.value, (int, float)):
                cell.number_format = money
            if c == 11 and row["note"]:
                cell.fill = warn_fill
                cell.font = Font(name='Times New Roman', bold=True, color='C00000')   # ⚡ 2026-08-11 N7：异常标记红字强化
        r += 1
    if not cp_rows_filtered:
        ws3.cell(row=r, column=1,
                 value=("注：综合查询明细表(GL)银行行的『对方科目』列全部为空（或均为银行存款之间的内部转账），"
                        "无法列示对方科目明细；主表数字仍以科目余额表(1002)控制数为准。"
                        "若需对方科目明细，请改用『对方科目』列已填的 GL 导出（标准 GL 导出已正常列示对方科目明细），"
                        "或让源系统补填『对方科目』列。")).font = Font(name='Times New Roman', italic=True, color="888888")
    for i, w in enumerate([14, 8, 30, 26, 9, 16, 9, 16, 12, 12, 30], 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = "A2"

    # ---- 销售精确汇总：银行借方(收款)凭证内 贷方=6销售科目 金额（凭证级行金额，剔除其他科目干扰） ----
    ws_sale = get_sheet("销售精确汇总")
    cols_sale = ["核算主体", "期间", "银行账户", "应收票据", "应收账款", "预收款项", "合同负债",
                 "主营业务收入", "其他业务收入", "销售收款合计"]
    ws_sale.append(cols_sale)
    style_header(ws_sale, len(cols_sale))
    _sale_agg = {}   # (e,y,acct) -> {cls: amt}
    for sr in (sale_rows or []):
        if target_year is not None and str(sr["year"]) != str(target_year):
            continue
        _k = (sr["entity"], sr["year"], sr["account"])
        _m = _sale_agg.setdefault(_k, {})
        _m[sr["cls"]] = _m.get(sr["cls"], 0.0) + sr["amt"]
    _rs = 2
    _sale_tot = {}
    for _k in sorted(_sale_agg):
        _e, _y, _acct = _k
        _m = _sale_agg[_k]
        _vals = [_m.get(c, 0.0) for c in ('应收票据', '应收账款', '预收款项', '合同负债', '主营业务收入', '其他业务收入')]
        _sum = sum(_vals)
        ws_sale.append([_e, _y, _acct] + [round(v, 2) for v in _vals] + [round(_sum, 2)])
        for _c in range(1, len(cols_sale) + 1):
            _cell = ws_sale.cell(row=_rs, column=_c); _cell.border = border; _cell.font = FONT10
            if _c >= 4:
                _cell.number_format = money
        for _i, _cl in enumerate(('应收票据', '应收账款', '预收款项', '合同负债', '主营业务收入', '其他业务收入')):
            _sale_tot[_cl] = _sale_tot.get(_cl, 0.0) + _vals[_i]
        _sale_tot['合计'] = _sale_tot.get('合计', 0.0) + _sum
        _rs += 1
    if _sale_agg:
        total_row(ws_sale, _rs, len(cols_sale),
                  dict([(1, "合计")] + [(i + 4, round(_sale_tot.get(c, 0.0), 2)) for i, c in
                                        enumerate(('应收票据', '应收账款', '预收款项', '合同负债', '主营业务收入', '其他业务收入'))] +
                       [(len(cols_sale), round(_sale_tot.get('合计', 0.0), 2))]))
    else:
        ws_sale.cell(row=_rs, column=1, value="注：本期银行借方(收款)凭证中未匹配到销售类对方科目（应收票据/应收账款/预收款项/合同负债/主营业务收入/其他业务收入）。").font = \
            Font(name='Times New Roman', italic=True, color="888888")
    for i, w in enumerate([14, 8, 30, 11, 11, 11, 11, 13, 13, 15], 1):
        ws_sale.column_dimensions[get_column_letter(i)].width = w
    ws_sale.freeze_panes = "A2"

    # ---- 采购精确汇总：银行贷方(付款)凭证内 借方=存货/预付/应付 金额（凭证级行金额） ----
    ws_purch = get_sheet("采购精确汇总")
    cols_purch = ["核算主体", "期间", "银行账户", "存货类", "预付类", "应付类", "采购付款合计"]
    ws_purch.append(cols_purch)
    style_header(ws_purch, len(cols_purch))
    _purch_agg = {}
    for pr in (purch_rows or []):
        if target_year is not None and str(pr["year"]) != str(target_year):
            continue
        _k = (pr["entity"], pr["year"], pr["account"])
        _m = _purch_agg.setdefault(_k, {})
        _m[pr["cls"]] = _m.get(pr["cls"], 0.0) + pr["amt"]
    _rp = 2
    _purch_tot = {}
    for _k in sorted(_purch_agg):
        _e, _y, _acct = _k
        _m = _purch_agg[_k]
        _vals = [_m.get(c, 0.0) for c in ('存货', '预付', '应付')]
        _sum = sum(_vals)
        ws_purch.append([_e, _y, _acct] + [round(v, 2) for v in _vals] + [round(_sum, 2)])
        for _c in range(1, len(cols_purch) + 1):
            _cell = ws_purch.cell(row=_rp, column=_c); _cell.border = border; _cell.font = FONT10
            if _c >= 4:
                _cell.number_format = money
        for _i, _cl in enumerate(('存货', '预付', '应付')):
            _purch_tot[_cl] = _purch_tot.get(_cl, 0.0) + _vals[_i]
        _purch_tot['合计'] = _purch_tot.get('合计', 0.0) + _sum
        _rp += 1
    if _purch_agg:
        total_row(ws_purch, _rp, len(cols_purch),
                  dict([(1, "合计")] + [(i + 4, round(_purch_tot.get(c, 0.0), 2)) for i, c in enumerate(('存货', '预付', '应付'))] +
                       [(len(cols_purch), round(_purch_tot.get('合计', 0.0), 2))]))
    else:
        ws_purch.cell(row=_rp, column=1, value="注：本期银行贷方(付款)凭证中未匹配到采购类对方科目（存货/预付账款/应付账款·应付票据）。").font = \
            Font(name='Times New Roman', italic=True, color="888888")
    for i, w in enumerate([14, 8, 30, 11, 11, 11, 15], 1):
        ws_purch.column_dimensions[get_column_letter(i)].width = w
    ws_purch.freeze_panes = "A2"

    # ---- 剩余对方科目核对：原『对方科目核对』剔除 销售/采购 精确部分后的剩余（凭证级对方科目行归集） ----
    ws_rem = get_sheet("剩余对方科目核对")
    cols_rem = ["核算主体", "期间", "银行账户", "对方科目", "借方笔数", "借方金额(本币)",
                "贷方笔数", "贷方金额(本币)"]
    ws_rem.append(cols_rem)
    style_header(ws_rem, len(cols_rem))
    _rr = 2
    for row in (remain_rows or []):
        if target_year is not None and str(row["year"]) != str(target_year):
            continue
        ws_rem.append([row["entity"], row["year"], row["account"], row["opp"],
                       row["n_d"], round(row["s_d"], 2), row["n_c"], round(row["s_c"], 2)])
        for _c in range(1, len(cols_rem) + 1):
            _cell = ws_rem.cell(row=_rr, column=_c); _cell.border = border; _cell.font = FONT10
            if _c in (6, 8):
                _cell.number_format = money
        _rr += 1
    if not (remain_rows or []):
        ws_rem.cell(row=_rr, column=1, value="注：剔除销售/采购精确部分后无剩余对方科目。").font = \
            Font(name='Times New Roman', italic=True, color="888888")
    else:
        ws_rem.cell(row=_rr, column=1, value=(
            "注：本表为『销售精确汇总』『采购精确汇总』之外的其他对方科目（凭证级对方科目行金额，非银行行全额）；"
            "原『对方科目核对』借方/贷方合计 与 本表+销售+采购 的差额 = 红字冲回/多银行账户凭证的银行行全额（不计入销售/采购）。")).font = \
            Font(name='Times New Roman', italic=True, color="888888")
    for i, w in enumerate([14, 8, 30, 30, 9, 16, 9, 16], 1):
        ws_rem.column_dimensions[get_column_letter(i)].width = w
    ws_rem.freeze_panes = "A2"

    finalize_workbook(wb)
    if target_year is not None:
        # 按年模式：重建审定表（当年）→ 直接保存（期间对比/差异对比保留双年段）
        from audit_common import add_audit_summary_sheets as _rebuild_bank
        from audit_common import discover_entities as _de_bank, read_tb_full as _rtb_bank
        if _adapter is not None and _adapter.is_sap(data_dir):
            _c_b = _adapter.current_comp()
            _ebank = ({_c_b: _de_bank(data_dir).get(_c_b, {})} if _c_b else _de_bank(data_dir))
        else:
            _ebank = _de_bank(data_dir)
        _tbank = _rtb_bank(data_dir, _ebank)
        # 删除其他年份的明细表占位 sheet（外壳模板可能自带）
        for _sn in list(wb.sheetnames):
            if _sn.startswith('银行存款明细表') and _sn != f'银行存款明细表{target_year}':
                del wb[_sn]
        if '货币资金审定表' in wb.sheetnames:
            del wb['货币资金审定表']
        # 2026-08-02 用户要求：货币资金审定表按三科目（库存现金/银行存款/其他货币资金）列示
        # + 每科目小计 + 集团各科目合计（不再用按主体一行的 add_audit_summary_sheets）
        try:
            _build_capital_audit_3subj(wb, data_dir, target_year)
        except Exception as _ex:
            print(f'  ⚠️ 货币资金审定表（三科目）生成失败，回退原版：{_ex}')
            _rebuild_bank(wb, data_dir,
                [dict(title='货币资金审定表', codes=['1001','1002','1012'],
                      is_credit=False, subj_name='货币资金', by_entity=True)],
                tb_full=_tbank, entities=_ebank, target_year=target_year)
        _sh = wb._sheets
        _ix = next((i for i, s in enumerate(_sh) if s.title == '货币资金审定表'), None)
        if _ix is not None and _ix > 0:
            _sh.insert(0, _sh.pop(_ix))
        _dm = f'银行存款明细表{target_year}'
        _ixd = next((i for i, s in enumerate(_sh) if s.title == _dm), None)
        if _ixd is not None and _ixd != 1:
            _sh.insert(1, _sh.pop(_ixd))
        # 附注汇总（货币资金，2026-08-02 用户方法论+当日确认）：
        # 结构=期末数(未审)→审计调整→期末审定数(公式)→期初数(仅数)；合并4行仅挂审定段后；
        # 主体列按 01、02 排序；表尾加「合计」列
        try:
            from audit_common import build_footnote_generic as _bfn_bank
            _bfn_bank(wb, data_dir, '附注汇总',
                      '货币资金附注汇总（审定口径；每主体一列；行=库存现金/银行存款/其他货币资金；'
                      '段=期末数(未审)/审计调整/期末审定数/期初数；增减变动由现金流量表核对）',
                      [('库存现金', ['1001']), ('银行存款', ['1002']), ('其他货币资金', ['1012'])],
                      is_credit=False, mode='money', target_year=target_year,
                      tb_full=_tbank, entities=_ebank, add_total_col=True)
        except Exception as _ex:
            print(f'  ⚠️ 附注汇总生成失败：{_ex}')
        from audit_shell import finalize_workbook as _fw_bank
        _fw_bank(wb)
        from audit_common import validate_workbook  # 2026-08-05 前道规范化：保存前结构校验
        validate_workbook(wb, '银行存款底稿', raise_on_error=False)
        wb.save(output_path)
        wb.close()
        return output_path
    from audit_common import validate_workbook
    validate_workbook(wb, '银行存款底稿', raise_on_error=False)
    wb.save(output_path)
    wb.close()
    return output_path




def main(argv=None):
    raw = list(sys.argv[1:] if argv is None else argv)
    # 拖入文件夹快捷方式：若首参为文件夹路径（不以 '-' 开头），自动视为 --input 并推导 --output
    if raw and not raw[0].startswith('-') and os.path.isdir(raw[0]):
        raw = ['--input', raw[0],
               '--output', os.path.join(raw[0], '银行存款审计底稿_生成.xlsx')] + raw[1:]
    p = argparse.ArgumentParser(description="通用银行存款明细表生成程序")
    p.add_argument("--input", "-i", required=False, help="交易数据文件(.csv/.xlsx) 或 含数据文件的文件夹")
    p.add_argument("--output", "-o", required=False, help="输出 Excel 路径")
    p.add_argument("--period", choices=["Y", "M", "Q"], default="Y", help="期间粒度 Y年/M月/Q季")
    p.add_argument("--opening", default=None, help="可选期初余额表(账号,币种,期初余额)；文件夹/单文件模式可自动识别")
    p.add_argument("--control", "-k", default=None, help="可选科目余额表路径(提取1002控制数)；文件夹/单文件模式可自动识别")
    args = p.parse_args(raw)
    if not args.input:
        print("❌ 用法：拖入【账套文件夹】到本程序，或指定 --input <文件夹> [--output <xlsx>]")
        return 1
    if not args.output:
        args.output = os.path.abspath(os.path.join(args.input, '银行存款审计底稿_生成.xlsx'))

    in_abs = os.path.abspath(args.input)
    out_abs = os.path.abspath(args.output)
    exclude = os.path.basename(out_abs) if os.path.dirname(out_abs) == in_abs else None

    if os.path.isdir(args.input):
        txn_files, auto_open, auto_km, gl_files = _discover_inputs(args.input, exclude_name=exclude)
        km_by_year, fc_by_year, gl_by_year = _discover_inputs_km(args.input, exclude_name=exclude)
        control_path = args.control or auto_km
        if km_by_year:
            # 科目余额表模式（多账套×多年份合并；主表数字直接取自科目余额表1002 + 外币科目余额表原币）
            entities = _discover_bank_entities(args.input)
            issues, years, detail, args.output = build_bank_detail_km_combined(
                args.input, args.output, args.period)
            n_acct = len(detail)
            n_ent = len(entities)
            print(f"📁 科目余额表模式(合并)：识别核算主体 {n_ent} 个、年度 {years}"
                  + (f"，账户数 {n_acct}" if n_acct else "，未识别到银行存款账户")
                  + f"，输出：{args.output}")
        elif gl_files and not txn_files:
            # GL 模式：从综合查询明细表抽取银行存款分录（无需另导银行日记账）
            rows, rw = [], []
            for g in gl_files:
                rows += read_gl_bank(g, rw)
            km_by_year2 = _build_km_year_map(args.input, exclude_name=exclude)
            control_map = km_by_year2
            opening_dict, owarn = _build_opening_from_km_multi(rows, km_by_year2)
            rw.extend(owarn)
            ctrl_fallback = next(iter(km_by_year2.values()), None) or control_path
            issues, periods, accounts = build_bank_detail(
                None, args.output, args.period, None, rows, rw, ctrl_fallback,
                opening_dict=opening_dict, control_map=control_map)
            print(f"📁 目录模式(GL)：识别综合查询明细表 {len(gl_files)} 个，抽取银行存款分录 {len(rows)} 笔"
                  + (f"，按年控制数(1002)：{ {k: os.path.basename(v) for k, v in km_by_year2.items()} }" if km_by_year2 else "（未提供科目余额表）"))
        elif txn_files:
            opening_path = args.opening or auto_open
            rows, rw = [], []
            for f in txn_files:
                rows += read_transactions(f, warnings=rw)
            issues, periods, accounts = build_bank_detail(
                None, args.output, args.period, opening_path, rows, rw, control_path)
            print(f"📁 目录模式：识别交易文件 {len(txn_files)} 个"
                  + (f"，期初文件：{os.path.basename(opening_path)}" if opening_path else "（未提供期初）")
                  + (f"，控制数文件(1002)：{os.path.basename(control_path)}" if control_path else "（未提供科目余额表）"))
        else:
            print(f"❌ 文件夹内未找到交易数据文件或综合查询明细表：{args.input}")
            return 1
    else:
        # 单文件：未显式给控制数时，尝试在同级目录自动发现科目余额表
        control_path = args.control
        if control_path is None:
            d = os.path.dirname(os.path.abspath(args.input))
            _, _, ak, _ = _discover_inputs(d)
            control_path = ak
        issues, periods, accounts = build_bank_detail(
            args.input, args.output, args.period, args.opening, None, None, control_path)
        if control_path:
            print(f"📄 控制数文件(1002)：{os.path.basename(control_path)}")

    print(f"✅ 已生成：{args.output}")
    _periods = periods if "periods" in dir() else (years if "years" in dir() else [])
    _nacct = len(accounts) if "accounts" in dir() else (len(detail) if "detail" in dir() else 0)
    print(f"   期间：{_periods}")
    print(f"   账户数：{_nacct}  记录数：见校验结果表")
    errs = [i for i in issues if i["level"] == "ERROR"]
    warns = [i for i in issues if i["level"] == "WARN"]
    print(f"   校验：ERROR {len(errs)} 条，WARN {len(warns)} 条")
    for it in issues:
        print(f"     [{it['level']}] {it['type']}: {it['message']}")
    from audit_common import finalize_after_build
    if os.path.isdir(args.input):
        finalize_after_build(args.input)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）
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
                                       f'bank_crash_{ts}.log'), 'w', encoding='utf-8') as _f:
                    _f.write('未捕获异常：\n' + traceback.format_exc())
                print(f'已记录崩溃日志：bank_crash_{ts}.log')
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
