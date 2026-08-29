# -*- coding: utf-8 -*-
"""就地注入『审定表』到现有科目明细表工作簿（不重跑 builder、不改位置）。

扫描指定文件夹（默认 J/c/x）中的 *_生成.xlsx 科目明细表，在明细 sheet 之前插入一张
『审定表』，数据取自该文件夹自身的科目余额表(TB)，可与 TB 直接核对。

- 费用/收入类：代码由文件名直接判定（5101/6601/6001+6051 ...）。
- 权益类：从工作簿明细 sheet 的『项目(明细科目)』列提取末级科目编码，反推父级
  权益代码（4001 实收资本 / 4104 利润分配 / 4101 盈余公积 / 4002 资本公积 ...）。
- 已含审定表的工作簿自动跳过（如 g 的权益类、各文件夹的应交税费/存货）。
"""
import paths as P
import os, re, sys
import openpyxl
import audit_common as A
from audit_shell import finalize_workbook

DESKTOP = P.DATA_ROOT

# 文件名(去 _生成.xlsx) -> (codes, is_credit, 显示名)
FEE_MAP = {
    "制造费用审计底稿":      (["5101"], False, "制造费用"),
    "营业外收入审计底稿":    (["6301"], True,  "营业外收入"),
    "税金及附加审计底稿":    (["6403"], False, "税金及附加"),
    "销售费用审计底稿":      (["6601"], False, "销售费用"),
    "管理费用审计底稿":      (["6602"], False, "管理费用"),
    "财务费用审计底稿":      (["6603"], False, "财务费用"),
    "研发费用审计底稿":      (["6604"], False, "研发费用"),
    "资产减值损失审计底稿":  (["6701"], False, "资产减值损失"),
    "信用减值损失审计底稿":  (["6702"], False, "信用减值损失"),
    "营业外支出审计底稿":    (["6711"], False, "营业外支出"),
    "所得税费用审计底稿":    (["6801"], False, "所得税费用"),
    "投资收益审计底稿":      (["6111"], True,  "投资收益"),
    "资产处置损益审计底稿":  (["6115"], True,  "资产处置损益"),
    "其他收益审计底稿":      (["6117"], True,  "其他收益"),
}
# ⚡ 2026-08-26 SAP 科目码映射（AH 是 SAP 账套，新准则码；U8/旧准则码不同）：
#   信用减值损失 6702(旧) → 6741(SAP)；资产处置损益 6115(旧) → 6720(SAP)；其他收益 6117(旧) → 6730(SAP)。
#   原 FEE_MAP 用 U8 码从 TB 取数 → SAP 下取不到 → 审定表近空（1010 实证：其他收益期末 2,086 万全丢）。
SAP_CODE_MAP = {"6702": "6741", "6115": "6720", "6117": "6730"}
# 费用/损益类均为利润表科目：审定表用「上年数/本期数/审计调整/审定数」格式（非资产负债表期初/期末）
FEE_MAP = {k: (v[0], v[1], v[2], True) for k, v in FEE_MAP.items()}

REVENUE_STEM = "营业收入审计底稿"
# 铁律⑤：营业收入按【名称】匹配（不写死代码 6001/6051）。
# J 账套科目表不一致：主营业务收入=6001(东轴)/5001(其余4家)，且东轴 5001=生产成本（同码不同义）。
# 故必须按名称汇总，跨代码/跨主体合并，避免漏数或把生产成本并进收入。
REVENUE_ITEM = dict(names=["主营业务收入", "其他业务收入"], is_credit=True, subj_name="营业收入", income_statement=True)

EQUITY_MAP = {
    "实收资本(或股本)审计底稿": (True, "实收资本(或股本)"),
    "未分配利润审计底稿":       (True, "未分配利润"),
    "盈余公积审计底稿":         (True, "盈余公积"),
    "资本公积审计底稿":         (True, "资本公积"),
}

# 已是独立审定表/非科目明细表/分析类，跳过
SKIP_PREFIX = ("应交税费审计底稿", "存货审计底稿", "库存股", "银行存款审计底稿",
               "货币资金", "专项储备", "审计完善工具包", "函证", "截止性",
               "抽凭", "存货监盘", "科目底稿", "重要性", "报表")


def stem_of(fn):
    if fn.endswith("_生成.xlsx"):
        return fn[:-len("_生成.xlsx")]
    return None


def equity_parent(wb, tb_codes):
    """从权益明细 sheet 的『项目(明细科目)』列提取末级科目编码，反推父级权益代码。"""
    ds = next((s for s in wb.sheetnames if ("明细" in s) and ("审定" not in s)), None)
    if not ds:
        return None
    ws = wb[ds]
    leaves = []
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if isinstance(cell, str):
                m = re.match(r"^(\d[\d.]+)", cell.strip())
                if m:
                    code = m.group(1)
                    if code[:1] == "4" and len(code) >= 4:
                        leaves.append(code)
    leaves = list(dict.fromkeys(leaves))
    if not leaves:
        return None
    for cand in sorted({l[:4] for l in leaves}):
        if cand in tb_codes and all(l.startswith(cand) for l in leaves):
            return [cand]
    return leaves  # 兜底：直接用末级（多行+合计）


def before_sheet(sheets):
    idx = next((i for i, s in enumerate(sheets) if "目录" in s), None)
    if idx is not None and idx + 1 < len(sheets):
        return sheets[idx + 1]
    return sheets[0]


def classify(stem, fp, tb_codes, is_sap_flag=False):
    # 容忍按年拆分导致的文件名年份后缀（管理费用审计底稿_2025 / _2026）
    base = re.sub(r"_(20\d{2})$", "", stem)
    if stem in FEE_MAP or base in FEE_MAP:
        codes, credit, label, income = FEE_MAP[stem if stem in FEE_MAP else base]
        if is_sap_flag:
            codes = [SAP_CODE_MAP.get(c, c) for c in codes]
        # 收入/成本底稿（营业收入底稿 / 制造费用/销售费用/管理费用/财务费用/研发费用）保持原样
        if stem == REVENUE_STEM or base == REVENUE_STEM or any(c in ('5101','6601','6602','6603','6604') for c in codes):
            return dict(codes=codes, is_credit=credit, subj_name=label, income_statement=income)
        return dict(codes=codes, is_credit=credit, subj_name=label, income_statement=income, by_entity=True)
    if stem == REVENUE_STEM or base == REVENUE_STEM:
        return dict(REVENUE_ITEM)
    ematch = next((k for k in EQUITY_MAP if stem.startswith(k) or base.startswith(k)), None)
    if ematch:
        credit, label = EQUITY_MAP[ematch]
        wb = openpyxl.load_workbook(fp, read_only=True)
        pc = equity_parent(wb, tb_codes)
        wb.close()
        if not pc:
            return None
        return dict(codes=pc, is_credit=credit, subj_name=label, by_entity=True)
    return None


def process_folder(folder):
    print(f"\n##### FOLDER {folder}")
    ents = A.discover_entities(folder)
    if not ents:
        print("  无实体，跳过")
        return
    tb = A.read_tb_full(folder, ents)
    tb_codes = {c for (e, c, n, y) in tb}
    for fn in sorted(os.listdir(folder)):
        if not fn.endswith("_生成.xlsx") or "_bak" in fn:
            continue
        stem = stem_of(fn)
        if stem is None or any(stem.startswith(p) for p in SKIP_PREFIX):
            continue
        fp = os.path.join(folder, fn)
        # 检测文件名年份（管理费用审计底稿_2025 → '2025'；注意 tb_full 年份键为字符串）
        ym = re.search(r"_(20\d{2})$", stem)
        target_year = ym.group(1) if ym else None
        try:
            is_sap_flag = bool(A.is_sap(folder)) if hasattr(A, 'is_sap') else False
            item = classify(stem, fp, tb_codes, is_sap_flag)
        except Exception as ex:
            print(f"  ⚠️ 分类失败 {fn}: {ex}")
            continue
        if item is None:
            continue
        wb = openpyxl.load_workbook(fp)
        try:
            title = f"{item['subj_name']} 审定表"
            if title in wb.sheetnames:
                if target_year:
                    # 按年文件：旧审定表可能为错误年度，移除后按目标年份重注入
                    wb.remove(wb[title])
                else:
                    print(f"  ↷ 已有审定表，跳过 {fn}")
                    continue
            item["title"] = title
            item["before"] = before_sheet(wb.sheetnames)
            A.add_audit_summary_sheets(wb, folder, [item], tb_full=tb, entities=ents,
                                       target_year=target_year)
            finalize_workbook(wb)  # 统一字体(Times New Roman/10号) + 千分位（含注入的审定表）
            wb.save(fp)
            print(f"  ✓ 注入 [{title}] -> {fn}  (before='{item['before']}')")
        except Exception as ex:
            print(f"  ✗ 失败 {fn}: {ex}")
        finally:
            wb.close()


if __name__ == "__main__":
    folders = sys.argv[1:] or [os.path.join(DESKTOP, d) for d in ["J", "c", "x"]]
    for f in folders:
        if os.path.isdir(f):
            process_folder(f)
    print("\n完成。")
