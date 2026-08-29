# -*- coding: utf-8 -*-
"""audit_sampling.py —— 审计抽样规模自动计算（方案五b，2026-08-11 用户提出）

痛点：复核需逐行统计总体 + 手工重算抽样规模，耗时最长；易少抽/漏抽大额异常样本。

功能：
  1) compute_min_sample_size(loop_type, om, pm, sad, risk, pop_amount)
     · attribute 属性抽样（控制测试）：n = Z²·p(1-p)/e²
       Z：95%=1.96 / 90%=1.645 / 80%=1.282；p=预计偏差率(默认2%)；e=可容忍偏差率(默认5%)
     · variable 变量抽样（余额/交易测试）：n ≈ (总体金额×Z²)/(可容忍−预计错报)
       可容忍错报=PM、预计错报=SAD（金额单位：元）
     · pps 货币单位抽样：n = (OM×可靠性系数)/(可容忍−预计错报)
       可靠性系数：95%=3.0 / 90%=2.31 / 80%=1.61
  2) inject_sampling_header(ws, subj, cfg)：抽凭 sheet 顶部固定 3-5 行
     （科目参数 + 总体 + 最低样本量 n + 必抽数 + 已抽数 + 大额样本标'必抽'）
  3) 接入点：audit_shell.finalize_workbook 统一注入（检测抽凭 sheet，改 1 处全生效）

配置：sampling_config.json —— default 全局默认 + accounts.<账套>.<科目> 覆盖。
  om=整体重要性(元) / pm=实际执行重要性 / sad=明显微小阈值 /
  risk=固有×控制风险等级(high=95% / medium=90% / low=80%) / loop=循环类型(attribute|variable|pps)
"""
from __future__ import annotations
import math
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

# 置信度 → Z 值 / 可靠性系数
_Z = {'high': 1.96, 'medium': 1.645, 'low': 1.282}
_RF = {'high': 3.0, 'medium': 2.31, 'low': 1.61}

_CFG_CACHE = {}


def load_config(data_dir=None):
    """读取抽样配置（sampling_config.json）。合并：default 全局默认 + accounts.<账套> 覆盖。"""
    key = data_dir or ''
    if key in _CFG_CACHE:
        return _CFG_CACHE[key]
    cfg = {'default': {'om': 10000000, 'pm': 5000000, 'sad': 500000,
                       'risk': 'medium', 'loop': 'variable',
                       'tol_rate': 0.05, 'exp_rate': 0.02},
           'accounts': {}}
    fp = os.path.join(HERE, 'sampling_config.json')
    if os.path.exists(fp):
        try:
            import json
            with open(fp, encoding='utf-8') as f:
                _d = json.load(f)
            cfg['default'].update(_d.get('default', {}))
            cfg['accounts'].update(_d.get('accounts', {}))
        except Exception:
            pass
    _CFG_CACHE[key] = cfg
    return cfg


def account_config(data_dir, subj):
    """某账套×科目的配置（default 合并账套覆盖）。subj=科目中文名。"""
    cfg = load_config(data_dir)
    d = dict(cfg['default'])
    acct = os.path.basename((data_dir or '').rstrip('\\/'))
    acct_cfg = (cfg.get('accounts') or {}).get(acct) or {}
    # 科目级覆盖（精确名 > 含关键词）
    if subj in acct_cfg:
        d.update(acct_cfg[subj])
    else:
        for k, v in acct_cfg.items():
            if k and k in subj:
                d.update(v)
                break
    return d


def compute_min_sample_size(loop_type, om=None, pm=None, sad=None, risk='medium',
                            pop_amount=None, tol_rate=0.05, exp_rate=0.02):
    """按循环类型计算最低样本量 n。返回 int。"""
    om = om or 0.0
    pm = pm or 0.0
    sad = sad or 0.0
    risk = risk if risk in _Z else 'medium'
    z = _Z[risk]
    if loop_type == 'attribute':
        # 属性抽样：n = Z²·p(1-p)/e²（p=预计偏差率，e=可容忍偏差率）
        p = exp_rate or 0.02
        e = tol_rate or 0.05
        n = z * z * p * (1 - p) / (e * e)
        return max(int(math.ceil(n)), 5)
    if loop_type == 'pps':
        # 货币单位抽样：n = OM×可靠性系数/(可容忍−预计错报)
        rf = _RF[risk]
        te = pm - sad
        if te <= 0:
            return min(pop_amount or 60, 60)
        n = om * rf / te
        return max(min(int(math.ceil(n)), 200), 5)
    # variable 变量抽样（默认）：n ≈ (总体金额×Z²)/(可容忍−预计错报)
    te = pm - sad
    pop = pop_amount or 0.0
    if te <= 0 or pop <= 0:
        return 30
    n = pop * z * z / te
    # 上限：总体量的 10%（避免样本量超总体）
    cap = max(int(pop / 10), 60) if pop > 0 else 60
    return max(min(int(math.ceil(n)), cap), 5)


def _find_data_cols(ws, header_row):
    """从表头行识别金额列/凭证号列。返回 (amt_cols, pop_col)。"""
    hdr = [ws.cell(header_row, c).value for c in range(1, ws.max_column + 1)]
    amt_cols = [i for i, h in enumerate(hdr, 1)
                if h and isinstance(h, str) and ('金额' in h or '借方' in h or '贷方' in h)]
    pop_col = next((i for i, h in enumerate(hdr, 1)
                    if h and isinstance(h, str) and ('凭证号' in h or '测试序号' in h)), None)
    return amt_cols, pop_col


def inject_sampling_header(ws, subj, data_dir=None):
    """抽凭 sheet 顶部注入抽样信息（固定 3-5 行）。
    统计总体（数据行数 + 金额合计）→ 计算最低样本量 n → 标必抽 → 写顶部。"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from audit_shell import SHELL_TITLE_FONT, SHELL_BORDER, SHELL_NUM, SHELL_LEFT, SHELL_CEN

    if ws.max_row < 2 or ws.max_column < 2:
        return
    # 表头行 = 第一个含『凭证号/测试序号』且含『金额』的行
    hr = None
    for r in range(1, min(ws.max_row, 6) + 1):
        row_vals = [str(ws.cell(r, c).value or '') for c in range(1, min(ws.max_column, 15) + 1)]
        if any('凭证号' in v or '测试序号' in v for v in row_vals) and any('金额' in v for v in row_vals):
            hr = r
            break
    if hr is None:
        return
    # 统计总体：数据行金额
    amt_cols, pop_col = _find_data_cols(ws, hr)
    pop = 0
    pop_amt = 0.0
    for r in range(hr + 1, ws.max_row + 1):
        if not ws.cell(r, 1).value and not ws.cell(r, 2).value and not ws.cell(r, 3).value:
            continue
        pop += 1
        for c in amt_cols:
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)):
                pop_amt += abs(v)
    if pop == 0:
        return
    cfg = account_config(data_dir, subj)
    n = compute_min_sample_size(cfg.get('loop', 'variable'),
                                om=cfg.get('om'), pm=cfg.get('pm'), sad=cfg.get('sad'),
                                risk=cfg.get('risk', 'medium'), pop_amount=pop_amt,
                                tol_rate=cfg.get('tol_rate', 0.05), exp_rate=cfg.get('exp_rate', 0.02))
    # 必抽：单笔 > PM 或 > SAD（大额/异常强制纳入）
    pm = cfg.get('pm', 0.0)
    sad = cfg.get('sad', 0.0)
    must_rows = []
    for r in range(hr + 1, ws.max_row + 1):
        for c in amt_cols:
            v = ws.cell(r, c).value
            if isinstance(v, (int, float)) and (abs(v) > pm or abs(v) > sad):
                must_rows.append(r)
                break
    must = len(set(must_rows))
    # ---- 顶部注入（insert 到表头前，原表头下移）----
    n_rows = 5
    ws.insert_rows(1, n_rows)
    _F_WARN = PatternFill('solid', fgColor='FFF2CC')
    _F_HDR = PatternFill('solid', fgColor='DDEBF7')
    _F_WHITE = Font(name='Times New Roman', size=10, bold=True, color='000000')
    _F_BOLD = Font(name='Times New Roman', size=10, bold=True)
    _F_NOTE = Font(name='Times New Roman', size=9, italic=True, color='808080')
    lines = [
        f'审计抽样规模（{subj}）—— 最低样本量 n={n}；总体 {pop} 笔 / {pop_amt:,.2f} 元；'
        f'实际已抽 {pop} 笔；必抽(>PM {pm:,.0f} 或 >SAD {sad:,.0f}) {must} 笔',
        f'参数：OM={cfg.get("om", 0):,.0f} 元；PM={pm:,.0f} 元；SAD={sad:,.0f} 元；'
        f'固有×控制风险={cfg.get("risk", "medium")}（置信 '
        f'{"95%" if cfg.get("risk") == "high" else "90%" if cfg.get("risk") == "medium" else "80%"}）；'
        f'循环类型={cfg.get("loop", "variable")}',
        f'属性抽样 n=Z²p(1-p)/e²；变量抽样 n≈(总体×Z²)/(PM−SAD)；PPS n=OM×RF/(PM−SAD)',
        f'大额/异常样本（>PM 或 >SAD）{must} 笔已强制纳入必抽；随机补足至 {n} 笔。',
        f'抽样配置：小程序目录 sampling_config.json（全局默认 + 本账套覆盖）。',
    ]
    for i, txt in enumerate(lines, 1):
        c = ws.cell(i, 1, txt)
        c.font = _F_BOLD if i == 1 else _F_NOTE
        c.alignment = SHELL_LEFT
        if i == 1:
            c.fill = _F_WARN
            c.font = _F_BOLD
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=min(ws.max_column, 12))
    # 表头行加粗保持
    for c in range(1, ws.max_column + 1):
        ws.cell(n_rows + 1, c).border = SHELL_BORDER
    # 必抽行标记（表头已下移 n_rows 行）
    for r in must_rows:
        rr = r + n_rows
        for c in range(1, min(ws.max_column, 12) + 1):
            ws.cell(rr, c).fill = _F_WARN
    # 冻结：表头行下方
    try:
        ws.freeze_panes = f'A{n_rows + 2}'
    except Exception:
        pass


# 抽凭 sheet 识别（finalize_workbook 统一注入用）
SAMPLING_SHEET_KW = ('凭证抽查', '抽查凭证', '费用抽查', '凭证抽查表')


def inject_all_sampling_sheets(wb, data_dir=None):
    """对工作簿内全部抽凭 sheet 注入抽样头（finalize_workbook 调用）。
    data_dir 优先取显式参数；否则取 wb._data_dir（builder 创建 wb 时设置，
    2026-08-11：让抽样注入可用账套级 sampling_config 覆盖）。"""
    if data_dir is None:
        data_dir = getattr(wb, '_data_dir', None)
    from openpyxl.utils import get_column_letter  # noqa
    label = ''
    for s in wb.sheetnames:
        if '审定表' in s:
            label = s.replace('审定表', '').strip()
            break
    if not label:
        for s in wb.sheetnames:
            m = re.search(r'^(.*?)[-_ ]?(20\d{2}|明细|汇总|审计)', s)
            if m:
                label = m.group(1)
                break
    for sn in wb.sheetnames:
        if any(k in sn for k in SAMPLING_SHEET_KW):
            try:
                inject_sampling_header(wb[sn], label or sn, data_dir)
            except Exception:
                pass


def main():
    """命令行：python audit_sampling.py <账套目录> [科目名]  # 试算并打印 n"""
    args = [a for a in sys.argv[1:] if not a.startswith('-')]
    if not args:
        print('用法：python audit_sampling.py <账套目录> [科目名]')
        return 1
    data_dir = args[0]
    subj = args[1] if len(args) > 1 else '测试科目'
    cfg = account_config(data_dir, subj)
    print(f'{subj} 配置：{cfg}')
    for loop in ('attribute', 'variable', 'pps'):
        for risk in ('high', 'medium', 'low'):
            n = compute_min_sample_size(loop, om=cfg.get('om'), pm=cfg.get('pm'),
                                        sad=cfg.get('sad'), risk=risk, pop_amount=1e8)
            if risk == cfg.get('risk', 'medium'):
                print(f'  {loop:10s} {risk}: n={n}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
