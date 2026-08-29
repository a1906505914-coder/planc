# -*- coding: utf-8 -*-
"""sheet_gap_check.py —— 底稿空表扫描（2026-08-11 用户核心诉求落地）。
检测"该有数据的表却是空的"——audit_checker 只查结构(表头/合计)，不查数据是否取到。
本模块：对每个底稿文件，扫描所有 sheet：
  · 关键表（审定表/明细表/附注汇总/对方科目核对等）为空（无任何非零数值行）→ 记录；
  · 结合该科目 TB 控制数：TB 有余额/发生额但审定表空 → ERROR（取数丢失）；
  · TB 也无数据 → INFO（科目无业务，空表属正常）。
返回 [(level, msg), ...] 供驱动方汇总展示。
"""
import os
import glob
import functools

# 判定为"关键表"的关键词（空则告警；含这些词的 sheet 空=异常）
# ⚡ 2026-08-29 P0：移除'勾稽'——『勾稽与异常检查』是 audit_checker 收尾追加的检查表，
#   无异常时为空属正常设计（此前集团模式空表扫描把它误报 ERROR，掩盖真实问题）。
_KEY_SHEET_KW = ('审定表', '明细表', '附注', '账龄', '对方科目', '核对',
                 '分月', '汇总', '抽查', '分析', '配比', '全链条', '凭证', '余额',
                 'Top10', '前十大', '坏账', '折旧', '摊销', '计提')
# 已知"允许为空"的模板表（空模板正常）
_OPTIONAL_SHEET_KW = ('审计调整分录', '调整分录', '底稿说明', '内部交易抵消核对',
                      '多借多贷核对', '差异凭证清单', '无研发支出科目披露', '勾稽与异常检查',
                      # ⚡ 2026-08-12：分部门明细表在 SAP 下为注记型空表（SAP 无『部门』级
                      # 辅助核算，表内 E2 已注明"无法生成分部门明细表"）——非取数丢失
                      '分部门明细表', '分部门')


def _has_data(ws, max_rows=500, max_cols=30):
    """sheet 是否存在非零数值（前 max_rows 行内）。
    ⚡ 2026-08-12：集团宽表 88 主体时数据列可能超 30（如附注 8040 主体在第 57 列）→
    原 max_cols=30 把有数据的集团表误判为空（交易性金融资产附注 241.9 万误报 ERROR）。
    max_cols=None 时扫全部列（read_only 大表可用 max_column 上限）。"""
    try:
        from openpyxl.utils import get_column_letter
        _mc = ws.max_column or 0
        _limit = _mc if max_cols is None else min(_mc, max_cols)
        for r in range(1, min(ws.max_row, max_rows) + 1):
            for c in range(1, _limit + 1):
                v = ws.cell(r, c).value
                if isinstance(v, (int, float)) and abs(v) > 0.005:
                    return True
    except Exception:
        return True  # 读取失败不误报
    return False


def scan_file(fp, tb_codes=None, tb_has_data=False, quiet=False):
    """扫描单个底稿文件。tb_codes=该科目相关科目码；tb_has_data=TB 是否有余额/发生额。
    返回 [(level, msg)]。"""
    from openpyxl import load_workbook
    issues = []
    try:
        wb = load_workbook(fp, read_only=True, data_only=True)
    except Exception as e:
        return [('ERROR', f'无法打开 {os.path.basename(fp)}: {e}')]
    empty_serious = []
    empty_optional = []
    # ⚡ 2026-08-12：集团文件（_yy集团 后缀，88 主体宽表）→ 全列扫描（max_cols=None），
    #   防 8040 等靠后主体数据列被 max_cols=30 截断误判空。
    _wide = ('集团' in os.path.basename(fp))
    for sn in wb.sheetnames:
        ws = wb[sn]
        if _has_data(ws, max_cols=None if _wide else 30):
            continue
        if any(k in sn for k in _OPTIONAL_SHEET_KW):
            empty_optional.append(sn)
            continue
        if any(k in sn for k in _KEY_SHEET_KW):
            # ⚡ 2026-08-12 注记型空表识别：表空但含『无发生额/无该科目/无法生成/未发现』
            # 等说明文字 → 属"真无数据"的设计空表（如 SAP 无部门辅助核算的分部门明细表、
            # 无应收利息科目的应收利息明细表、无在建工程减值的减值准备明细表），
            # 非取数丢失 → 归 INFO 不报 ERROR（避免 TB 有数据但子表真无数据误报）。
            _note = ''
            for row in ws.iter_rows(max_row=8, max_col=4):
                for c in row:
                    if isinstance(c.value, str) and any(
                            k in c.value for k in ('无发生额', '无该科目', '无法生成',
                                                   '未发现', '无需要抽查', '无数据',
                                                   '不产生', '未生成')):
                        _note = c.value.strip()[:40]
                        break
                if _note:
                    break
            if _note:
                empty_optional.append(f'{sn}（注记：{_note}…）')
            else:
                # ⚡ 2026-08-12 全 0 合计表识别：表空但存在『合计』行且全部数值为 0
                #   （含字符串 '0'，如 baddebt_common 无减值数据生成的减值准备明细表
                #   合计 0）→ 属"真无数据"的设计空壳，非取数丢失 → INFO。
                _zero_tot = False
                try:
                    for _row in ws.iter_rows(max_row=30):
                        if any(str(c.value or '').strip() in ('合计', '总计') for c in _row):
                            _nums = [c.value for c in _row if isinstance(c.value, (int, float))]
                            _str0 = [c.value for c in _row if isinstance(c.value, str) and c.value.strip() == '0']
                            if _nums and all(abs(v) < 0.005 for v in _nums) or (_str0 and not _nums):
                                _zero_tot = True
                            break
                except Exception:
                    pass
                if _zero_tot:
                    empty_optional.append(f'{sn}（合计全 0：无数据）')
                else:
                    empty_serious.append(sn)
    wb.close()
    if empty_serious:
        lvl = 'ERROR' if tb_has_data else 'INFO'
        issues.append((lvl, f'{os.path.basename(fp)} 关键表为空: {empty_serious}'
                            + ('' if tb_has_data else '（TB 该科目也无数据，属正常）')))
    elif empty_optional:
        issues.append(('INFO', f'{os.path.basename(fp)} 可选表空(正常): {empty_optional}'))
    return issues


def scan_folder(folder, tb_has_by_file=None, quiet=False):
    """扫描目录下全部底稿。tb_has_by_file: {文件名: bool} 可选——该科目 TB 是否有数据。
    返回 (ERROR 数, [(level, msg)])."""
    from collections import Counter
    all_issues = []
    files = sorted(glob.glob(os.path.join(folder, '*审计底稿*.xlsx')))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]
    for fp in files:
        name = os.path.basename(fp)
        # 从文件名推断科目（去掉 审计底稿_xxx 后缀）
        base = name.split('审计底稿')[0]
        tb_had = False
        if tb_has_by_file:
            tb_had = tb_has_by_file.get(name, False)
        issues = scan_file(fp, tb_has_data=tb_had, quiet=quiet)
        all_issues += issues
    errs = sum(1 for l, _ in all_issues if l == 'ERROR')
    return errs, all_issues


if __name__ == '__main__':
    import sys
    folder = sys.argv[1] if len(sys.argv) > 1 else '.'
    errs, issues = scan_folder(folder)
    print(f'空表扫描: {errs} ERROR / {len(issues)} 项')
    for l, m in issues:
        print(f'  [{l}] {m}')


# ===================== TB 控制数判定 =====================
# 按底稿文件名推断科目 → 从 TB 查该科目有无余额/发生额。
# 有数据但关键表空=ERROR（取数丢失）；无数据=INFO（科目无业务）。
_SUBJ_KEYWORDS = {
    '应收账款': ('1122',), '应付账款': ('2202',), '其他应收款': ('1221',), '其他应付款': ('2241',),
    '预付账款': ('1123',), '预收账款': ('2203',), '应收票据': ('1121',), '合同资产': ('1124',),
    '合同负债': ('2205',), '存货': ('1403', '1405', '1406', '1407', '1408', '1409', '1411'),
    '固定资产': ('1601', '1602', '1603'), '无形资产': ('1701', '1702', '1703'),
    '在建工程': ('1604',), '投资性房地产': ('1521', '1522', '1523'), '使用权资产': ('1608', '1609'),
    '长期待摊费用': ('1801',), '递延收益': ('2401',), '研发支出': ('5301',),
    '应交税费': ('2221',), '应付职工薪酬': ('2211',), '职工薪酬': ('2211',),
    '营业收入': ('6001',), '营业成本': ('6401',), '研发费用': ('6602',),
    '管理费用': ('6602',), '销售费用': ('6601',), '财务费用': ('6603',),
    '制造费用': ('5101',), '所得税费用': ('6801',), '投资收益': ('6111',),
    '其他收益': ('6117',), '营业外收入': ('6301',), '营业外支出': ('6711',),
    '税金及附加': ('6403',), '资产减值损失': ('6701',), '信用减值损失': ('6741',),
    '资产处置损益': ('6715',), '实收资本': ('4001',), '资本公积': ('4002',),
    '盈余公积': ('4101',), '未分配利润': ('4104',), '库存股': ('4201',),
    '短期借款': ('2001',), '长期借款': ('2501',), '应付债券': ('2502',),
    '应付票据': ('2201',), '长期股权投资': ('1511',), '长期应付款': ('2701',),
    '递延所得税资产': ('1811',), '递延所得税负债': ('2901',), '专项储备': ('4102',),
    '专项应付款': ('2711',), '预计负债': ('2215',), '交易性金融资产': ('1101',),
    '一年内到期': ('2101',), '租赁负债': ('2260',),
}


# ⚡⚡ 2026-08-24 缓存：self_check 全量扫描时每空表科目调一次 TB 判定（每次重读 TB 超慢+刷警告）
@functools.lru_cache(maxsize=16)
def _load_tb_cached(data_dir):
    """按账套缓存 TB 全量（self_check 多空表判定共用，避免每次重读+刷 .xls 警告）。"""
    import sap_adapter
    if sap_adapter.is_sap(data_dir):
        sap_adapter.set_root(data_dir)
        return sap_adapter.read_tb_full(data_dir, None)
    import audit_common
    entities = audit_common.discover_entities(data_dir)
    return audit_common.read_tb_full(data_dir, entities)


def tb_has_subject_data(data_dir, subj_name, comp=None):
    """查某科目在 TB 是否有余额/发生额（comp 给定时限该主体；否则任一主体任一年度）。
    返回 bool。读取失败保守返回 True（不误报 INFO）。"""
    try:
        tb = _load_tb_cached(data_dir)
        codes = _SUBJ_KEYWORDS.get(subj_name, ())
        # ⚡⚡ 2026-08-24 名称主键（铁律58）：命中代码行须名称含 subj_name 才判有数据
        #   （5301 在 ADZ=营业外收入 ≠ 研发支出 → 不误判研发支出有数据；防止代码跨账套异义）。
        #   代码命中后名称校验放宽：subj_name 或其前 2 字在名称中（应付职工薪酬/职工薪酬
        #   互为包含；实收资本(或股本) 前2字=实收资本 命中 TB『实收资本』）。
        if not codes:
            # 无代码映射 → 名称兜底（如研发支出走名称判定，不靠 5301 代码）
            for (e, c, n, y), v in tb.items():
                if comp and str(e) != str(comp):
                    continue
                if subj_name and (subj_name in str(n) or subj_name[:2] in str(n)):
                    if any(abs(v.get(k) or 0) > 0.005 for k in ('qc', 'jf', 'df', 'qm')):
                        return True
            return False
        for (e, c, n, y), v in tb.items():
            if comp and str(e) != str(comp):
                continue
            if str(c).startswith(codes):
                # ⚡ 2026-08-12 铁律110（用户明确）：资产负债表科目只要有发生数（jf/df 任一
                #   非零），即使期末=0，也要生成底稿。故所有科目统一看 余额+发生额 任一非零
                #   （原余额类只查 qc/qm 会把"期末0但有发生额"主体误判无数据→漏生成）。
                if subj_name and not (subj_name in str(n) or subj_name[:2] in str(n) or str(n) in subj_name):
                    continue  # ⚡⚡ 名称主键：代码命中但名称不符（5301=营业外收入）→ 跳过
                if any(abs(v.get(k) or 0) > 0.005 for k in ('qc', 'jf', 'df', 'qm')):
                    return True
        return False
    except Exception:
        return True


def build_tb_map(data_dir, folder=None):
    """构建 {底稿文件名: 该科目 TB 是否有数据}。文件名如 应收账款审计底稿_1010.xlsx。
    底稿在 folder（默认 data_dir/底稿）；TB 判定用 data_dir。
    ⚡ 2026-08-12：解析文件名主体后缀（_1010）→ 按该主体判定；集团文件（_yy集团）
    用全集团判定。原全集团判定使无数据主体的空表误报 ERROR。"""
    import re as _re
    folder = folder or os.path.join(data_dir, '底稿')
    if not os.path.isdir(folder):
        folder = data_dir
    out = {}
    for f in os.listdir(folder):
        if not f.endswith('.xlsx') or '审计底稿' not in f or f.startswith('~$'):
            continue
        base = f.split('审计底稿')[0]
        hit = None
        for name in _SUBJ_KEYWORDS:
            if name in base:
                hit = name
                break
        if not hit:
            out[f] = True
            continue
        # 主体后缀：_1010.xlsx → 1010；_yy集团.xlsx → 集团（None=全集团）
        _m = _re.search(r'_(\d{4})\.xlsx$', f)
        _comp = _m.group(1) if _m else None
        out[f] = tb_has_subject_data(data_dir, hit, comp=_comp)
    return out


def scan_folder_with_tb(data_dir, folder=None, quiet=False):
    """带 TB 判定的目录扫描：有数据但关键表空=ERROR。返回 (ERROR 数, issues)。
    data_dir=账套根目录（TB 判定用）；folder=底稿目录（默认 data_dir/底稿）。"""
    folder = folder or os.path.join(data_dir, '底稿')
    if not os.path.isdir(folder):
        folder = data_dir
    tb_map = build_tb_map(data_dir, folder)
    return scan_folder(folder, tb_has_by_file=tb_map, quiet=quiet)
