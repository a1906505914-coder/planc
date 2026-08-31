# -*- coding: utf-8 -*-
"""sap_report.py — 账套企业报表（利润表/资产负债表 .xls）读取器（2026-08-31 新建）。

背景：AH 类 SAP 账套在数据目录下导出权威企业报表：
  利润表/{comp}.xls      会企 02 表（项目/行次/本月/累计/上年同期）
  资产负债表/{comp}.xls  会企 01 表（资产|行次|期末|年初|负债和所有者权益|行次）
自建试算表（sap_tb_gen）的 Sheet2 资产负债表 / Sheet3 利润表以企业报表为核对基准
（用户铁律：报表↔试算表 核对相符）→ 直接读企业报表，保证口径完全一致。

费用目录 费用/{comp}/{类别}N.xlsx（N=月份序号，取最大=最新累计）：
  按功能范围拆分的期间费用（管理费用/销售费用/研发费用/制造费用），
  每文件『合计：』行 = 企业利润表对应科目（已验证 1010 精确一致）。
  明细行含父级小计（工资附加费合计：）+ 缩进子行（\u3000 开头）→ 取数须排除子行与合计行。
"""
import os
import re
import glob

import xlrd
import openpyxl

# 企业利润表科目行名 → 试算表标准行名（sap_tb_gen.PL_ROWS 的 rn）
_PL_KEYWORDS = {
    '营业收入': ['一、营业收入', '营业收入'],
    '主营业务收入': ['其中：主营业务收入', '主营业务收入'],
    '其他业务收入': ['其他业务收入'],
    '营业成本': ['减：营业成本', '营业成本'],
    '主营业务成本': ['其中：主营业务成本', '主营业务成本'],
    '其他业务成本': ['其他业务成本'],
    '税金及附加': ['税金及附加'],
    '销售费用': ['销售费用'],
    '管理费用': ['管理费用'],
    '研发费用': ['研发费用'],
    '财务费用': ['财务费用'],
    '其他收益': ['加：其他收益', '其他收益'],
    '投资收益': ['投资收益'],
    '信用减值损失': ['信用减值损失'],
    '资产减值损失': ['资产减值损失'],
    '营业外收入': ['加：营业外收入', '营业外收入'],
    '营业外支出': ['减：营业外支出', '营业外支出'],
    '所得税费用': ['减：所得税', '所得税'],
}
# 资产负债表科目行名（资产段 col0 / 负债权益段 col4）→ 标准行名
_BS_KEYWORDS = {
    '货币资金': ['货币资金'], '交易性金融资产': ['交易性金融资产'],
    '应收票据': ['应收票据'], '应收账款': ['应收账款'], '应收款项融资': ['应收款项融资'],
    '预付款项': ['预付款项'], '应收利息': ['应收利息'], '应收股利': ['应收股利'],
    '其他应收款': ['其他应收款'], '存货': ['存货'], '合同资产': ['合同资产'],
    '持有待售资产': ['持有待售资产'], '一年内到期的非流动资产': ['一年内到期的非流动资产'],
    '其他流动资产': ['其他流动资产'], '其他权益工具投资': ['其他权益工具投资'],
    '其他债券投资': ['其他债券投资'], '长期应收款': ['长期应收款'],
    '长期股权投资': ['长期股权投资'], '投资性房地产': ['投资性房地产'],
    '固定资产': ['固定资产'], '在建工程': ['在建工程'], '工程物资': ['工程物资'],
    '固定资产清理': ['固定资产清理'], '使用权资产': ['使用权资产'],
    '无形资产': ['无形资产'], '开发支出': ['开发支出'], '商誉': ['商誉'],
    '长期待摊费用': ['长期待摊费用'], '递延所得税资产': ['递延所得税资产'],
    '其他非流动资产': ['其他非流动资产'],
    '短期借款': ['短期借款'], '应付票据': ['应付票据'], '应付账款': ['应付账款'],
    '预收款项': ['预收款项'], '合同负债': ['合同负债'], '应付职工薪酬': ['应付职工薪酬'],
    '应交税费': ['应交税费'], '应付利息': ['应付利息'], '应付股利': ['应付股利'],
    '其他应付款': ['其他应付款'], '持有待售负债': ['持有待售负债'],
    '一年内到期的非流动负债': ['一年内到期的非流动负债'], '其他流动负债': ['其他流动负债'],
    '长期借款': ['长期借款'], '应付债券': ['应付债券'], '租赁负债': ['租赁负债'],
    '长期应付款': ['长期应付款'], '专项应付款': ['专项应付款'], '预计负债': ['预计负债'],
    '递延收益': ['递延收益'], '递延所得税负债': ['递延所得税负债'],
    '其他非流动负债': ['其他非流动负债'],
    '实收资本': ['实收资本'], '其他权益工具': ['其他权益工具'], '资本公积': ['资本公积'],
    '库存股': ['库存股'], '其他综合收益': ['其他综合收益'], '盈余公积': ['盈余公积'],
    '专项储备': ['专项储备'], '一般风险储备': ['一般风险储备'], '未分配利润': ['未分配利润'],
}


def _resolve_root(data_dir):
    """解析原始数据目录：run_u8_on_sap 传给生成器的 data_dir 是 _work_{comp} 中间目录
    （费用目录/企业报表在原始数据根下）→ 若 data_dir 以 _work 开头或无 费用/利润表 目录，
    回退到父目录。"""
    if os.path.isdir(data_dir):
        for sub in ('费用', '利润表', '资产负债表'):
            if os.path.isdir(os.path.join(data_dir, sub)):
                return data_dir
        base = os.path.basename(os.path.normpath(data_dir))
        if base.startswith('_work'):
            parent = os.path.dirname(data_dir)
            if os.path.isdir(os.path.join(parent, '费用')):
                return parent
    return data_dir


def has_enterprise_reports(data_dir):
    """账套是否导出企业报表目录（利润表/资产负债表）。"""
    data_dir = _resolve_root(data_dir)
    return (os.path.isdir(os.path.join(data_dir, '利润表'))
            and os.path.isdir(os.path.join(data_dir, '资产负债表')))


def _find_row(rows, keywords):
    """在企业报表行列表里找关键词匹配的行（行文本前缀/包含关键词，跳过表头行）。"""
    for i, r0 in enumerate(rows):
        if not r0:
            continue
        for kw in keywords:
            if r0 == kw or r0.startswith(kw) or kw in r0:
                return r0
    return None


def read_ent_profit(data_dir, comp):
    """读企业利润表 {comp}.xls → {标准行名: 累计金额}。无文件/无匹配返回 {}。"""
    data_dir = _resolve_root(data_dir)
    fp = os.path.join(data_dir, '利润表', f'{comp}.xls')
    if not os.path.exists(fp):
        return {}
    sh = xlrd.open_workbook(fp).sheet_by_index(0)
    rows = []
    for i in range(sh.nrows):
        rows.append((str(sh.cell_value(i, 0)).strip(), sh.cell_value(i, 3)))  # 累计金额 col3
    out = {}
    for rn, kws in _PL_KEYWORDS.items():
        for r0, v in rows:
            if r0 and isinstance(v, (int, float)) and any(kw in r0 for kw in kws):
                out[rn] = v
                break
    return out


def read_ent_balance(data_dir, comp):
    """读企业资产负债表 {comp}.xls → {标准行名: 期末余额}。资产取 col2、负债权益取 col7。"""
    data_dir = _resolve_root(data_dir)
    fp = os.path.join(data_dir, '资产负债表', f'{comp}.xls')
    if not os.path.exists(fp):
        return {}
    sh = xlrd.open_workbook(fp).sheet_by_index(0)
    out = {}
    for i in range(4, sh.nrows):
        a = str(sh.cell_value(i, 0)).strip()
        l = str(sh.cell_value(i, 4)).strip()
        av = sh.cell_value(i, 2)
        lv = sh.cell_value(i, 7) if sh.ncols > 7 else None
        for rn, kws in _BS_KEYWORDS.items():
            if a and isinstance(av, (int, float)) and any(kw in a for kw in kws):
                out.setdefault(rn, 0.0)
                out[rn] += av
            if l and isinstance(lv, (int, float)) and any(kw in l for kw in kws):
                out.setdefault(rn, 0.0)
                out[rn] += lv
    return out


def _to_f(x):
    try:
        return float(str(x).replace(',', '').strip())
    except (ValueError, TypeError):
        return 0.0


def read_fee_scope(data_dir, comp):
    """读费用目录（功能范围拆分期间费用）：
    格式 A（子目录）：费用/{comp}/{类别}N.xlsx（N=月份序号，取最大=最新累计，本期累计 col3）
    格式 B（根目录单文件）：费用/{comp}.xlsx（sheets=类别，列=1月..N月，累计=Σ月份）
    返回 {类别: [(科目名, 金额), ...]}，科目名=非缩进非合计的一级项。
    无费用目录/无文件 → {}。"""
    data_dir = _resolve_root(data_dir)
    base = os.path.join(data_dir, '费用', str(comp))
    if not os.path.isdir(base):
        fp = os.path.join(data_dir, '费用', f'{comp}.xlsx')
        if not os.path.exists(fp):
            return {}
        return _read_fee_flat(fp)
    out = {}
    for cat in ('管理费用', '销售费用', '研发费用', '制造费用'):
        files = glob.glob(os.path.join(base, f'{cat}*.xlsx'))
        if not files:
            continue
        # 最大序号 = 最新月份累计
        def _seq(fp):
            m = re.search(r'(\d+)\.xlsx$', fp)
            return int(m.group(1)) if m else 0
        best = max(files, key=_seq)
        wb = openpyxl.load_workbook(best, read_only=True, data_only=True)
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
        # ⚡⚡ 2026-08-31 格式探测：表头含『月』列（如 1月/2月..N月）→ 逐月列格式
        #   （1030 研发费用.xlsx 混入子目录；科目名 col1、合计=Σ月份）；否则标准格式 A
        #   （表头 项目/本月/本期累计，科目名 col0、取本期累计 col3）
        _is_monthly = bool(rows) and any(
            c is not None and str(c).strip().endswith('月') for c in rows[0])
        items = []
        for r in rows[1:]:
            if not r or not r[0]:
                continue
            if _is_monthly:
                nm = str(r[1]) if len(r) > 1 and r[1] else ''
                s = nm.strip()
                if s.startswith('合计') or nm.startswith('\u3000'):
                    continue
                v = sum(_to_f(x) for x in r[2:])
            else:
                nm = str(r[0])
                s = nm.strip()
                if s.startswith('合计') or nm.startswith('\u3000'):
                    continue
                v = _to_f(r[2])  # 本期累计数
            items.append((s, v))
        wb.close()
        out[cat] = items
    return out


def _read_fee_flat(fp):
    """格式 B：费用/{comp}.xlsx（sheets=类别，列=1月..N月，累计=Σ月份列）。"""
    out = {}
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    for cat in ('管理费用', '销售费用', '研发费用', '制造费用'):
        if cat not in wb.sheetnames:
            continue
        ws = wb[cat]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 2:
            continue
        items = []
        for r in rows[1:]:
            if not r or len(r) < 3 or not r[1]:
                continue
            nm = str(r[1])
            if nm.startswith('\u3000'):
                continue   # 缩进子行（父级小计的拆分）
            s = nm.strip()
            if s.startswith('合计'):
                continue
            v = sum(_to_f(x) for x in r[2:])   # 1月..N月累计
            items.append((s, v))
        if items:
            out[cat] = items
    wb.close()
    return out
