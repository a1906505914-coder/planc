# -*- coding: utf-8 -*-
"""sap_current_account_detail.py —— SAP 往来款底稿生成器（2026-08-08 v1）。

方法论（2026-08-08 固化，AR/AP 117 家均已 100% <0.5% 对平）：
  - 科目圈定：应收=1122、其他应收=1221（排除 122104 资金结算）、预付=1123、
    应付=2202、其他应付=2241、预收=2203（票据 1121/2201、资金池 122104 独立不计）。
  - 凭证级净额化：同凭证借贷同科目互抵（SAP 清账）→ 净额 0 排除；净额按方向分类。
  - 增减构成（审计关注）：
      资产类（应收/其他应收/预付）：增加=净借（贷方对方：银行=回款、收入=确认）、
      减少=净贷（借方对方：银行/票据=收款、应付=互抵、预收/预付=转）；
      负债类（应付/其他应付/预收）：增加=净贷（借方对方：存货/成本/费用/资产=采购、
      银行=收款）、减少=净借（贷方对方：银行/票据=付款、应收=互抵）。
  - 期末=自建试算表（名称级）；期初=期末−借+贷（净额反推）。
输出：{data_dir}/底稿/往来款审计底稿_{科目名}.xlsx
  Sheet1 审定表（公司×期初/本期增加/本期减少/期末/净变动）
  Sheet2 增减构成（增加/减少按对方科目分类 × 公司）
用法：python sap_current_account_detail.py <data_dir> [<comp|all>] [--out <dir>]
"""
import paths as P
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import _safx
import sap_common as C

# ---- 样式 ----
FONT = Font(name='Times New Roman', size=10)
BOLD = Font(name='Times New Roman', size=10, bold=True)
HEAD_FONT = Font(name='Times New Roman', size=10, bold=True)
HEAD_FILL = PatternFill('solid', fgColor='DDEBF7')
TOT_FILL = PatternFill('solid', fgColor='FCE4D6')
NUMFMT = '#,##0.00'
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
RGT = Alignment(horizontal='right', vertical='center')
LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
_thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

MONTHS = 7
YEAR = '2026'

# 科目配置：key → (前缀, 名称, 性质, aux源)  性质: asset=借方科目 / liab=贷方科目
# aux源: 'cust'=应收文件(客户维度) / 'supp'=应付文件(供应商维度)
# ⚡ 2026-08-09 实测：SAP 客户主数据(应收文件)含 1121/1122/1124/1221/2203(预收!);
#   供应商主数据(应付文件)含 1123(预付)/2202/2241。预收账款是负债但客户维度。
SUBJECTS = [
    ('ar',  ('1122',), '应收账款', 'asset', 'cust'),
    ('arn', ('1121',), '应收票据', 'asset', 'cust'),
    ('ca',  ('1124',), '合同资产', 'asset', 'cust'),
    ('ora', ('1221',), '其他应收款', 'asset', 'cust'),
    # ⚡ 2026-08-09：1222 备用金（SAP 实测独立科目，9 家余额 39 万，原遗漏；应收文件无数据→占位）
    ('petty', ('1222',), '备用金', 'asset', 'cust'),
    ('prepay', ('1123',), '预付账款', 'asset', 'supp'),
    ('ap',  ('2202',), '应付账款', 'liab', 'supp'),
    ('apn', ('2201',), '应付票据', 'liab', 'supp'),
    ('orp', ('2241',), '其他应付款', 'liab', 'supp'),
    ('advrecv', ('2203',), '预收账款', 'liab', 'cust'),
    # 合同负债：审计/报表科目体系预留（2026-08-08 用户方法论）。
    # 科目体系完整=所有科目都能生成底稿；但当前仍按账套科目取数——
    # SAP 账套无 2205 合同负债科目 → 自动跳过不生成；未来有审计调整
    # （借预收账款 贷合同负债）时再生成对应底稿承接核对。
    ('cl', ('2205',), '合同负债', 'liab', 'supp'),
]
FUND_EXCL = '122104'   # 资金结算排除
BANK = ('1001', '1002')     # 银行（票据 1121/2201 已是独立往来科目，不再是对方归类）
INV = ('1401', '1402', '1403', '1404', '1405', '1406', '6401', '6402', '6403',
       '6600', '5101', '5301', '5001', '1601', '1602', '1603', '1604', '1701')
REV = ('6001', '6051')                       # 收入
AR_OPP = ('1122', '1221')                    # 应收（对方互抵）
AP_OPP = ('2202', '2241')                    # 应付（对方互抵）
ADV = ('1123', '2203')                       # 预付/预收


def _gl_files(data_dir, comp):
    gl_dir = os.path.join(data_dir, '序时账')
    out = []
    for fn in sorted(os.listdir(gl_dir)):
        if not fn.lower().endswith('.xlsx'):
            continue
        base = os.path.splitext(fn)[0]
        m = re.match(r'^(\d{4})', base)
        if not m:
            continue
        m2 = re.match(r'^(\d{4})[~-](\d{4})$', base)
        if m.group(1) == comp:
            out.append(os.path.join(gl_dir, fn))
        elif m2 and m2.group(1) <= comp <= m2.group(2):
            out.append(os.path.join(gl_dir, fn))
    return out


def analyze_subject(data_dir, comp, pfx, is_liab):
    """单主体单科目：凭证级净额化 → 增减构成。返回 dict。"""
    v = {}
    for row in C.iter_comp_gl(data_dir, comp, keep_cols={1, 5, 6, 10, 11, 12, 21, 29}):
        if str(row.get(6, '')).strip() != comp:
            continue
        per = str(row.get(10, '')).strip()
        if not (per[:2].isdigit() and 1 <= int(per[:2]) <= 7):
            continue
        subj = str(row.get(11) or '')
        try:
            amt = float(row.get(29) or 0)
        except (TypeError, ValueError):
            continue
        vno = str(row.get(1) or '').strip()
        if not vno or not subj or amt == 0:
            continue
        d = v.setdefault(vno, defaultdict(float))
        d[subj] += amt
        if '_txt' not in d:
            d['_txt'] = str(row.get(5) or '').strip()
        if '_cust' not in d and row.get(12):
            d['_cust'] = str(row.get(12)).strip()
        if '_supp' not in d and row.get(21):
            d['_supp'] = str(row.get(21)).strip()

    stat = defaultdict(float)   # 分类 → 金额
    net_jf = net_df = 0.0       # 净借合计 / 净贷合计
    for vno, d in v.items():
        net = sum(a for s, a in d.items() if s not in ('_txt', '_cust', '_supp')
                  and s.startswith(pfx) and not s.startswith(FUND_EXCL))
        if abs(net) < 0.005:
            continue
        if net > 0:
            net_jf += net
        else:
            net_df += -net
        deb = [(s, a) for s, a in d.items() if s not in ('_txt', '_cust', '_supp') and a > 0.01
               and not (s.startswith(pfx) and not s.startswith(FUND_EXCL))]
        cred = [(s, -a) for s, a in d.items() if s not in ('_txt', '_cust', '_supp') and a < -0.01
                and not (s.startswith(pfx) and not s.startswith(FUND_EXCL))]
        if net > 0:   # 净借
            bank = sum(a for s, a in cred if s.startswith(BANK))
            rev = sum(a for s, a in cred if s.startswith(REV))
            advc = sum(a for s, a in cred if s.startswith(('1123', '2203')))
            apc = sum(a for s, a in cred if s.startswith(AP_OPP))
            other = sum(a for s, a in cred) - bank - rev - advc - apc
            if is_liab:      # 负债类净借=减少
                if bank > 0:
                    stat['付款(贷银行/票据)'] += net
                elif rev > 0:
                    stat['转收入/冲收入(贷收入)'] += net
                elif advc > 0:
                    stat['预付/预收转出(贷预付预收)'] += net
                else:
                    stat['其他减少(贷其他)'] += net
            else:            # 资产类净借=增加
                if rev > 0:
                    stat['收入确认(贷收入)'] += net
                elif bank > 0:
                    stat['回款/代垫(贷银行)'] += net
                elif advc > 0:
                    stat['预收/预付转应收(贷预收预付)'] += net
                elif apc > 0:
                    stat['应付转入(贷应付)'] += net
                else:
                    stat['其他增加(贷其他)'] += net
        else:                # 净贷
            bank = sum(a for s, a in deb if s.startswith(BANK))
            inv = sum(a for s, a in deb if s.startswith(INV))
            arc = sum(a for s, a in deb if s.startswith(AR_OPP))
            revc = sum(a for s, a in deb if s.startswith(REV))
            advc2 = sum(a for s, a in deb if s.startswith(('1123', '2203')))
            other = sum(a for s, a in deb) - bank - inv - arc - revc - advc2
            if is_liab:      # 负债类净贷=增加
                if inv > 0:
                    stat['采购挂账(借存货/成本/资产)'] += -net
                elif bank > 0:
                    stat['收款(借银行)'] += -net
                elif arc > 0:
                    stat['应收互抵转入(借应收)'] += -net
                elif revc > 0:
                    stat['预收确认收入(借收入)'] += -net
                else:
                    stat['其他增加(借其他)'] += -net
            else:            # 资产类净贷=减少
                if bank > 0:
                    stat['收款/结转(借银行/票据)'] += -net
                elif arc > 0:
                    stat['互抵转应付(借应付)'] += -net
                elif advc2 > 0:
                    stat['转预收/预付(借预收预付)'] += -net
                elif revc > 0:
                    stat['收入冲回(借收入)'] += -net
                elif inv > 0:
                    stat['冲预付/成本转(借成本)'] += -net
                else:
                    stat['其他减少(借其他)'] += -net
    # 增加/减少按方向：资产类增加=净借、减少=净贷；负债类反之
    inc = net_jf if not is_liab else net_df
    dec = net_df if not is_liab else net_jf
    # 异常凭证识别（2026-08-08 用户方法论：双向验证+活跃度；写入底稿）
    exc_rows = []
    MIN_EXC = 500000.0
    for vno, d in v.items():
        net = sum(a for s, a in d.items() if s not in ('_txt', '_cust', '_supp')
                  and s.startswith(pfx) and not s.startswith(FUND_EXCL))
        if abs(net) < MIN_EXC:
            continue
        txt = str(d.get('_txt', ''))
        if '1122999998' in d or '上评估' in txt or '重置' in txt:
            continue
        # 排除正常业务：收入确认（贷 6001/6051）、银行收款/付款（对方=银行票据）
        has_rev = any(s.startswith(REV) and a < 0 for s, a in d.items()
                      if s not in ('_txt', '_cust', '_supp'))
        has_bank = any(s.startswith(BANK) for s, a in d.items()
                       if s not in ('_txt', '_cust', '_supp'))
        if has_rev or has_bank:
            continue
        ap_amt = sum(abs(a) for s, a in d.items() if s not in ('_txt', '_cust', '_supp')
                    and s.startswith(AP_OPP))
        supp = str(d.get('_supp', '')).strip()
        act = 0
        if supp:
            act = len({cc for (cc, vv) in v.items()
                      if str(vv.get('_supp', '')).strip() == supp and cc != vno})
        # 业务性质
        if ap_amt > 0 and any(k in txt for k in ('水电', '电费', '水费', '水电气')):
            nature = '正常-水电费结算(连续日常)'
        elif ap_amt > 0 and '暂估' in txt:
            nature = '正常-暂估互抵(应付侧双向确认)'
        elif ap_amt > 0:
            nature = ('正常-持续往来互抵' if act >= 3
                      else ('⚠️ 大额冲抵-%s本年无持续业务(%d笔),需核查' % (supp or '无供应商', act)))
        elif '发票' in txt and any(k in txt for k in ('冲销', '冲回', '红字', '红冲')):
            nature = '发票冲销-需核查红冲依据'
        else:
            nature = '大额调整-需核查事由'
        exc_rows.append((vno, round(net, 2), txt[:40], supp, nature))
    return dict(stat), inc, dec, exc_rows


# ⚡ 2026-08-09：SAP 辅助核算明细表支持（用户要求与 U8 底稿一致——U8 版 9 sheet 含明细表/账龄/Top10）
# 数据源：{data_dir}/应收/{comp}.xlsx（客户维度余额表）+ {data_dir}/应付/{comp}.xlsx（供应商维度余额表）
#         {data_dir}/客户行项目/{comp}.xlsx + 供应商行项目/{comp}.xlsx（流水 → 账龄/交易笔数）
_AUX_QC = 6   # 应收/应付表 期初余额 列
_AUX_JF = 7   # 本期借方
_AUX_DF = 8   # 本期贷方
_AUX_QM = 9   # 期末余额
_AUX_CODE = 3 # 科目编号
_AUX_NAME = 4 # 科目名称
_AUX_CC = 1   # 客户/供应商编码
_AUX_CN = 2   # 客户/供应商名称
_AUX_CON = 5  # 合同号（分项目列示维度，2026-08-09 用户要求）


def _aux_rows(data_dir, comp, want_asset):
    """读取 {应收|应付}/{comp}.xlsx → [(科目码, 科目名, 单位编码, 单位名称, 合同号, 期初, 借, 贷, 期末)]。
    want_asset=True 用应收（客户），False 用应付（供应商）。
    ⚡ 铁律73（2026-08-09 实测）：SAP 应付表余额列【贷正借负】（与 TB/应收表借正贷负反号）——
    1123 预付 aux=-403,479 而 TB=+403,479、2241 其他应付 aux=+423,229 而 TB=-423,229。
    故应付侧 qc/qm 取负还原为记账符号（借正贷负）；jf/df 保持正数发生额。"""
    sub = '应收' if want_asset else '应付'
    sgn = 1.0 if want_asset else -1.0
    fp = os.path.join(data_dir, sub, f'{comp}.xlsx')
    if not os.path.exists(fp):
        return []
    out = []
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or len(r) < 10 or r[_AUX_CODE] is None:
                continue
            code = str(r[_AUX_CODE]).strip()
            nm = str(r[_AUX_NAME] or '').strip()
            cc = str(r[_AUX_CC] or '').strip()
            cn = str(r[_AUX_CN] or '').strip()
            con = str(r[_AUX_CON] or '').strip() if len(r) > _AUX_CON else ''
            try:
                qc = float(r[_AUX_QC] or 0); jf = float(r[_AUX_JF] or 0)
                df = float(r[_AUX_DF] or 0); qm = float(r[_AUX_QM] or 0)
            except (TypeError, ValueError):
                continue
            if abs(qc) + abs(jf) + abs(df) + abs(qm) < 0.005:
                continue
            out.append((code, nm, cc, cn, con, qc * sgn, jf, df, qm * sgn))
        wb.close()
    except Exception:
        pass
    return out


def _supp_line_ledger(data_dir, comp, pfx='2202'):
    """供应商行项目 → 逐笔流水 [(凭证日期, 凭证编号, 文本, 科目码, 供应商编码, 金额)]。
    ⚡ 2026-08-09 实测：GL 2202 应付仅 14% 行带 col21 供应商、凭证号 join 不可靠（69 凭证同凭证多供应商）
    → 改用供应商行项目（col10=供应商编码 293/358 匹配应付表；col1=总帐帐目科目；col11=过账日期；
    col12=本币金额；col3=文本）。行项目含 2024/2025 历史行 → 按过账日期年份 2026 过滤
    （过滤后与应付表 2202 净额完全一致 ✓）。凭证级净额化同 GL 口径。"""
    import datetime as _dt
    from collections import defaultdict
    fp = os.path.join(data_dir, '供应商行项目', f'{comp}.xlsx')
    if not os.path.exists(fp):
        return []
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = ws.iter_rows(values_only=True)
        hdr = list(next(rows))
        col = {}
        for i, h in enumerate(hdr):
            hs = str(h or '').strip()
            if hs in ('总帐帐目', '科目'):
                col.setdefault('code', i)
            elif hs == '凭证编号':
                col.setdefault('vno', i)
            elif hs in ('过账日期', '凭证日期'):
                col.setdefault('date', i)
            elif hs in ('本币金额',):
                col.setdefault('amt', i)
            elif hs in ('文本',):
                col.setdefault('txt', i)
        agg = defaultdict(lambda: [0.0, '', ''])
        for r in rows:
            if not r or 'code' not in col or 'vno' not in col or 'amt' not in col:
                continue
            code = str(r[col['code']] or '').strip()
            if not code.startswith(pfx):
                continue
            supp = str(r[col.get('supp', 10)] or '').strip() if 'supp' in col else \
                (str(r[10] or '').strip() if len(r) > 10 else '')
            if not supp:
                continue
            try:
                amt = float(r[col['amt']] or 0)
            except (TypeError, ValueError):
                continue
            if abs(amt) < 0.005:
                continue
            dt = str(r[col['date']] or '').strip()[:10] if 'date' in col else ''
            if dt[:4] != '2026':   # ⚡ 行项目含历史年，只取本年发生（与应付表本期口径一致）
                continue
            vno = str(r[col['vno']] or '').strip()
            a = agg[(vno, code, supp)]
            a[0] += amt
            if not a[1]:
                a[1] = dt
            if not a[2]:
                a[2] = str(r[col['txt']] or '').strip() if 'txt' in col else ''
        wb.close()
    except Exception:
        return []
    out = [(v[1], vno, v[2], code, supp, round(v[0], 2))
           for (vno, code, supp), v in agg.items() if abs(v[0]) >= 0.005]
    out.sort(key=lambda x: (x[0], x[1]))
    return out


def _gl_ledger_rows(data_dir, comp, pfx, want_asset):
    """GL 序时账中该公司往来科目（pfx）的逐笔流水 → [(凭证日期, 凭证编号, 文本, 科目码, 单位编码, 金额)]。
    ⚡ 铁律73（2026-08-09 实测）：GL 往来科目行客户类 100% 带 col12 客户编码
    （1121/1122/1124/1221/2203）、1123 预付/2241 100% 带 col21 供应商；仅 2202 应付 14%
    → 用凭证号 join GL 内供应商映射补全。凭证日期=col28（Excel 序列号，表头行自动过滤）。
    ⚡ 凭证级净额化（铁律）：同凭证内同科目借贷互抵只计净额（SAP 常同凭证塞红字冲销/重分类），
    与应收/应付表口径一致——否则借方/贷方虚增（1030 实测差异 42 亿）。
    返回按 (日期, 凭证号) 排序的流水，供「往来明细账」sheet 使用。"""
    import datetime as _dt
    from collections import defaultdict
    # 凭证号 → 供应商编码 映射（补全 2202 等供应商缺失行）
    vno_supp = {}
    for r in C.iter_comp_gl(data_dir, comp, keep_cols={1, 11, 21}):
        vno = str(r.get(1) or '').strip()
        s21 = str(r.get(21) or '').strip()
        if vno and s21:
            vno_supp.setdefault(vno, s21)
    # 凭证级聚合：key=(凭证号, 科目码, 单位) → [净额, 日期, 文本]
    agg = defaultdict(lambda: [0.0, '', ''])
    for r in C.iter_comp_gl(data_dir, comp, keep_cols={1, 5, 9, 11, 12, 21, 28, 29}):
        code = str(r.get(11) or '').strip()
        if not code.startswith(pfx):
            continue
        try:
            amt = float(r.get(29) or 0)
        except (TypeError, ValueError):
            continue
        if abs(amt) < 0.005:
            continue
        vno = str(r.get(1) or '').strip()
        if want_asset:
            unit = str(r.get(12) or '').strip()
        else:
            unit = str(r.get(21) or '').strip() or vno_supp.get(vno, '')
        if not unit:
            continue
        a = agg[(vno, code, unit)]
        a[0] += amt
        # 日期：col28 Excel 序列号 → YYYY-MM-DD；异常时退化为期间
        if not a[1]:
            c28 = r.get(28)
            if c28 is not None:
                try:
                    a[1] = str(_dt.date(1899, 12, 30) + _dt.timedelta(days=float(c28)))
                except (TypeError, ValueError):
                    a[1] = str(r.get(9) or '')
            else:
                a[1] = str(r.get(9) or '')
        if not a[2]:
            a[2] = str(r.get(5) or '').strip()
    out = [(v[1], vno, v[2], code, unit, round(v[0], 2))
           for (vno, code, unit), v in agg.items() if abs(v[0]) >= 0.005]
    out.sort(key=lambda x: (x[0], x[1]))
    return out


def _aux_ledger_sheet(wb, name, comp, aux_rows, ledger, is_liab=False, tb_qm=0.0):
    """U8 风格「往来明细账」：账簿式逐笔流水（单位×科目分组，组内按日期/凭证号排序）。
    列：凭证日期/凭证编号/摘要/科目/借方/贷方/余额（期初+借-贷滚动）。
    ⚡ 铁律73（2026-08-09）：负债类科目（应付/预收/其他应付）贷方增借方减 → 余额=期初+贷-借；
      资产类（应收/预付/其他应收）借方增贷方减 → 余额=期初+借-贷。金额符号：GL=借正贷负，
      供应商行项目=负贷正借（SAP 行项目符号与 GL 相反）——借方金额列恒显正数、贷方列恒显负数绝对值。
    期初余额取应收/应付表（单位+科目），期末与表内滚动累计核对。"""
    ws = wb.create_sheet(f'{name}往来明细账')
    hdrs = ['核算主体', '往来单位', '科目', '凭证日期', '凭证编号', '摘要', '借方金额', '贷方金额', '期末余额']
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(1, j, h)
        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
    # 单位编码+科目 → (单位名, 期初)（应收/应付表 qc；同单位同科目多合同号 → 期初按合同号拆分行，
    # 无合同号的按科目合计作为兜底）
    unit_begin = {}
    for code, nm, cc, cn, con, qc, jf, df, qm in aux_rows:
        key = (cc, code)
        a = unit_begin.setdefault(key, [cn, 0.0, 0.0])
        a[1] += qc
        a[2] += qm
    # 分组：先按 单位×科目 分组（组内按日期排序），防同组被跨组行隔开致期初重复
    from collections import OrderedDict
    groups = OrderedDict()
    for dt, vno, txt, code, unit, amt in ledger:
        groups.setdefault((unit, code), []).append((dt, vno, txt, code, unit, amt))
    r = 2
    t_j = t_l = 0.0
    rows = []
    for (unit, code), lines in groups.items():
        lines.sort(key=lambda x: (x[0], x[1]))
        cn = unit_begin.get((unit, code), ['', 0.0, 0.0])[0]
        qc0 = unit_begin.get((unit, code), ['', 0.0, 0.0])[1]
        bal = qc0
        subj_j = subj_l = 0.0
        rows.append(('BEGIN', comp, cn or unit, code, '', '', '期初余额', 0.0, 0.0, round(bal, 2)))
        for dt, vno, txt, code2, unit2, amt in lines:
            # ⚡ 铁律73：金额符号统一 借正贷负（GL 与供应商行项目一致：贷方=负）。
            #   期初/期末经 _aux_rows 已统一为记账符号（借正贷负，资产正/负债负）
            #   → 滚动统一 期末=期初+借-贷（资产/负债同式，is_liab 不再影响方向）。
            j_amt = amt if amt > 0 else 0.0
            l_amt = -amt if amt < 0 else 0.0
            bal = bal + j_amt - l_amt
            subj_j += j_amt; subj_l += l_amt
            t_j += j_amt; t_l += l_amt
            rows.append(('', comp, cn or unit, code2, dt, vno, txt, j_amt, l_amt, round(bal, 2)))
        rows.append(('SUB', comp, cn or unit, '', '', '', '小计', subj_j, subj_l, round(bal, 2)))
    # ⚡ 铁律73：应收/应付表有余额但 GL 无发生（期初=期末）的组补一行——明细账合计须与表全口径对平
    for (unit, code), (cn, qc0, qm0) in unit_begin.items():
        if (unit, code) not in groups:
            rows.append(('BEGIN', comp, cn or unit, code, '', '', '期初余额', 0.0, 0.0, round(qc0, 2)))
            rows.append(('SUB', comp, cn or unit, '', '', '', '小计（无发生，期初=期末）', 0.0, 0.0, round(qc0, 2)))
    # 写表
    for tag, comp_v, unit_v, code_v, dt_v, vno_v, txt_v, j_v, l_v, bal_v in rows:
        if tag == 'SUB':
            for j, v in enumerate([comp_v, unit_v, '', '', '', txt_v, j_v, l_v, bal_v], 1):
                c = ws.cell(r, j, v)
                if isinstance(v, (int, float)):
                    c.number_format = NUMFMT
                c.border = BORDER
                if j == 6:
                    c.font = BOLD
            r += 1
            continue
        for j, v in enumerate([comp_v, unit_v, code_v, dt_v, vno_v, txt_v, j_v, l_v, bal_v], 1):
            c = ws.cell(r, j, v)
            if isinstance(v, (int, float)):
                c.number_format = NUMFMT
            c.border = BORDER
            c.alignment = RGT if j in (7, 8, 9) else LFT
            if tag == 'BEGIN' and j == 6:
                c.font = BOLD
        r += 1
    # 合计行
    sub_tot = sum(x[9] for x in rows if x[0] == 'SUB')
    d_tb = tb_qm - sub_tot
    ws.cell(r, 1, '合计').font = BOLD
    ws.cell(r, 2, f'借方={t_j:,.2f} 贷方={t_l:,.2f}').font = BOLD
    r += 1
    # ⚡ 铁律73：TB 兜底占位行（无辅助核算部分；与明细表口径一致）
    if abs(d_tb) > 0.005:
        ws.cell(r, 1, '（无辅助核算明细）').font = BOLD
        ws.cell(r, 2, f'TB 差额：无单位辅助核算明细，期末=TB-明细账小计 {d_tb:,.2f}').font = BOLD
        ws.cell(r, 9, round(d_tb, 2)).number_format = NUMFMT
        ws.cell(r, 9).border = BORDER
    for j, w2 in enumerate([10, 26, 14, 12, 14, 34, 14, 14, 16], 1):
        ws.column_dimensions[chr(64 + j)].width = w2
    ws.freeze_panes = 'A2'
    return ws


def _line_items(data_dir, comp, want_asset):
    """读取 {客户|供应商}行项目/{comp}.xlsx → [(科目码, 文本, 凭证日期, 清帐日期, 金额)]。
    账龄依据：未清项（清帐日期为空/未清）按凭证日期分桶。"""
    sub = '客户行项目' if want_asset else '供应商行项目'
    fp = os.path.join(data_dir, sub, f'{comp}.xlsx')
    if not os.path.exists(fp):
        return []
    out = []
    try:
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = ws.iter_rows(values_only=True)
        hdr = list(next(rows))
        col = {}
        for i, h in enumerate(hdr):
            hs = str(h or '').strip()
            if hs == '总帐帐目':
                col.setdefault('code', i)
            elif hs in ('凭证日期', '过账日期'):
                col.setdefault('date', i)
            elif hs in ('清帐日期',):
                col.setdefault('clr', i)
            elif hs in ('本币金额',):
                col.setdefault('amt', i)
            elif hs in ('文本',):
                col.setdefault('txt', i)
        for r in rows:
            if not r or 'code' not in col or 'amt' not in col:
                continue
            code = str(r[col['code']] or '').strip()
            try:
                amt = float(r[col['amt']] or 0)
            except (TypeError, ValueError):
                continue
            if not code or abs(amt) < 0.005:
                continue
            dt = str(r[col['date']] or '').strip()[:10] if 'date' in col else ''
            clr = str(r[col['clr']] or '').strip() if 'clr' in col else ''
            txt = str(r[col['txt']] or '').strip() if 'txt' in col else ''
            out.append((code, txt, dt, clr, amt))
        wb.close()
    except Exception:
        pass
    return out


def _aging_bucket(dt, base='2026-07-31'):
    """账龄分桶（2026-08-09 用户明确口径：按凭证日期年份，2026年发生=1年以内、2025年=1-2年…）。
    返回桶名或 None（无日期）。"""
    if not dt or not re.match(r'^\d{4}-\d{2}-\d{2}', dt):
        return None
    y = int(dt[:4])
    by = int(base[:4])
    yrs = by - y
    if yrs < 0:
        return '1年以内'
    if yrs < 1:
        return '1年以内'
    if yrs < 2:
        return '1-2年'
    if yrs < 3:
        return '2-3年'
    if yrs < 4:
        return '3-4年'
    if yrs < 5:
        return '4-5年'
    return '5年以上'


def _aux_detail_sheet(wb, name, comp, aux_rows, lines, tb_qc=0.0, tb_qm=0.0):
    """生成 U8 风格「明细表」sheet：单位×科目维度 期初/借/贷/期末/平衡校验。
    ⚡ SAP 行项目无客户维度关联（分配列仅 6% 有值）→ 账龄/交易笔数仅科目级，放合计行+账龄分布 sheet。
    ⚡ 铁律73：TB 兜底占位行——aux 合计≠TB 时（部分单位无辅助核算），补「（无辅助核算明细）」行=s_open/s_close 差额。"""
    ws = wb.create_sheet(f'{name}明细表')
    hdrs = ['核算主体', '序号', '往来单位编号', '往来单位名称', '款项性质', '期初余额',
            '本期借方发生额', '本期贷方发生额', '期末余额', '平衡校验(期初+借-贷-期末)']
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(1, j, h)
        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
    # 行项目按科目码聚合：账龄桶 + 笔数（科目级；账龄只统计未清项——清帐日期为空=构成期末余额）
    from collections import defaultdict
    li = defaultdict(lambda: defaultdict(float))
    li_cnt = defaultdict(int)
    for code, _t, dt, clr, amt in lines:
        if not clr:   # 未清项才算账龄
            b = _aging_bucket(dt)
            if b:
                li[code][b] += abs(amt)
        li_cnt[code] += 1
    r = 2
    seq = 0
    # 行项目账龄：按凭证年份分桶（用户口径 2026=1年以内…）；未清项=清帐日期空
    # ⚡ SAP 行项目清帐凭证/日期全空 → 全部视为未清，按凭证日期年份分桶
    bucket_amt = defaultdict(float)
    for code, _t, dt, clr, amt in lines:
        b = _aging_bucket(dt)
        if b:
            bucket_amt[b] += abs(amt)
    for code, nm, cc, cn, con, qc, jf, df, qm in aux_rows:
        seq += 1
        bal = round(qc + jf - df - qm, 2)
        vals = [comp, seq, cc, cn, con or '—', nm, qc, jf, df, qm, bal]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            if isinstance(v, (int, float)):
                c.number_format = NUMFMT
            c.border = BORDER
            c.alignment = RGT if j >= 7 else LFT
        r += 1
    # 合计行：期初/借/贷/期末 + 平衡校验
    if seq:
        t = [0.0] * 5
        for code, nm, cc, cn, con, qc, jf, df, qm in aux_rows:
            t[0] += qc; t[1] += jf; t[2] += df; t[3] += qm
        t[4] = round(t[0] + t[1] - t[2] - t[3], 2)
        ws.cell(r, 1, '合计').font = BOLD
        for j, v in zip((7, 8, 9, 10, 11), t):
            c = ws.cell(r, j, round(v, 2))
            c.number_format = NUMFMT; c.border = BORDER
        r += 1
        # ⚡ 铁律73：TB 兜底占位行（aux 合计≠TB 时差额=无辅助核算部分）
        d_qc = tb_qc - t[0]
        d_qm = tb_qm - t[3]
        if abs(d_qc) > 0.005 or abs(d_qm) > 0.005:
            ws.cell(r, 3, '（无辅助核算明细）').font = BOLD
            ws.cell(r, 5, 'TB 差额：该部分无单位辅助核算明细，期末=TB-aux合计').font = BOLD
            for j, v in zip((7, 8, 9, 10, 11), (round(d_qc, 2), 0.0, 0.0, round(d_qm, 2), 0.0)):
                c = ws.cell(r, j, v)
                c.number_format = NUMFMT; c.border = BORDER
            r += 1
        # 账龄分布（按凭证年份分桶，未清发生额）
        ws.cell(r, 5, '账龄分布（按凭证日期年份，未清发生额）').font = BOLD
        for i, k in enumerate(('1年以内', '1-2年', '2-3年', '3-4年', '4-5年', '5年以上')):
            c = ws.cell(r, 7 + i, round(bucket_amt.get(k, 0), 2))
            c.number_format = NUMFMT; c.border = BORDER
    for j, w2 in enumerate([10, 6, 14, 26, 16, 18, 14, 14, 14, 14, 22], 1):
        ws.column_dimensions[chr(64 + j)].width = w2
    ws.freeze_panes = 'A2'
    return ws


def _aux_aging_sheet(wb, name, lines):
    """U8 风格「账龄分布」sheet：按凭证年份分桶汇总（用户口径 2026=1年以内…）。"""
    ws = wb.create_sheet('账龄分布')
    hdrs = ['账龄区间', '期末余额', '合计']
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(1, j, h)
        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
    bucket_amt = defaultdict(float)
    for code, _t, dt, clr, amt in lines:
        b = _aging_bucket(dt)
        if b:
            bucket_amt[b] += abs(amt)
    r = 2
    for k in ('1年以内', '1-2年', '2-3年', '3-4年', '4-5年', '5年以上'):
        v = round(bucket_amt.get(k, 0), 2)
        ws.cell(r, 1, k).border = BORDER
        c = ws.cell(r, 2, v)
        c.number_format = NUMFMT; c.border = BORDER
        c2 = ws.cell(r, 3, v)
        c2.number_format = NUMFMT; c2.border = BORDER
        r += 1
    tot = round(sum(bucket_amt.values()), 2)
    ws.cell(r, 1, '合计').font = BOLD
    c = ws.cell(r, 2, tot); c.number_format = NUMFMT; c.border = BORDER
    c.font = BOLD
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16
    return ws


# 进程级缓存：科目码→名称映射（read_sap_tb 全量读 88 家 ~8s，缓存后每进程一次）
_TB_NAME_CACHE = {}


def _load_name_map(data_dir, comp):
    """读该公司 4位科目码→名称 映射（进程级缓存）。"""
    key = (data_dir, comp)
    if key not in _TB_NAME_CACHE:
        names = {}
        try:
            import sap_reader as SR
            tb = SR.read_sap_tb(data_dir, None, year='2026')
            for (c, _code, name, _yy), vv in tb.items():
                if c == comp and vv.get('code'):
                    p4 = str(vv['code'])[:4]
                    names.setdefault(p4, name)
        except Exception:
            pass
        _TB_NAME_CACHE[key] = names
    return _TB_NAME_CACHE[key]


def _opp_pairs(data_dir, comp, pfx):
    """凭证级对方科目核对：该公司该科目(pfx)全部凭证中，同凭证内其他科目（对方科目）的净额分布。
    返回 [(对方科目4位码, 名称, 借方金额, 贷方金额, 笔数)]。
    ⚡ 铁律73（2026-08-09）：GL 无对方科目列 → 凭证级配对（同凭证内本科目净额 vs 其他科目行）。
    本科目借方净额 → 对方科目记贷方；本科目贷方净额 → 对方科目记借方。凭证级净额化防互抵虚增。"""
    from collections import defaultdict
    v = defaultdict(lambda: defaultdict(float))
    for r in C.iter_comp_gl(data_dir, comp, keep_cols={1, 11, 29}):
        code = str(r.get(11) or '').strip()
        try:
            amt = float(r.get(29) or 0)
        except (TypeError, ValueError):
            continue
        vno = str(r.get(1) or '').strip()
        if not vno or abs(amt) < 0.005:
            continue
        v[vno][code] += amt
    opp = defaultdict(lambda: [0.0, 0.0, 0])   # 4位主码 → [借方金额, 贷方金额, 笔数]
    for vno, lines in v.items():
        self_net = sum(a for c, a in lines.items() if c.startswith(pfx))
        if abs(self_net) < 0.005:
            continue
        for c, a in lines.items():
            if c.startswith(pfx) or abs(a) < 0.005:
                continue
            p4 = c[:4]
            if self_net > 0 and a < 0:      # 本科目借净 → 对方科目贷方
                opp[p4][1] += -a
                opp[p4][2] += 1
            elif self_net < 0 and a > 0:    # 本科目贷净 → 对方科目借方
                opp[p4][0] += a
                opp[p4][2] += 1
    # 科目码 → 名称（进程级缓存，防每份底稿全量重读 TB）
    names = _load_name_map(data_dir, comp)
    out = [(c, names.get(c, ''), v[0], v[1], v[2])
           for c, v in opp.items() if v[0] + v[1] > 0.005]
    out.sort(key=lambda x: -(x[2] + x[3]))
    return out


def _aux_opp_sheet(wb, name, comp, lines, opp=None):
    """U8 风格「对方科目核对」sheet：凭证级对方科目配对（净额化）。
    ⚡ 2026-08-09 增强：原为行项目文本归集（无核对意义）→ 改为 GL 凭证级配对——
    对方科目码+名称+借方/贷方净额+笔数，可核对业务性质（销售→收入、收款→银行等）。"""
    ws = wb.create_sheet('对方科目核对')
    hdrs = ['核算主体', '对方科目编码', '对方科目名称', '借方金额', '贷方金额', '笔数', '说明']
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(1, j, h)
        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
    r = 2
    if opp:
        for code, nm, j_amt, l_amt, cnt in opp:
            ws.cell(r, 1, comp).border = BORDER
            ws.cell(r, 2, code).border = BORDER
            ws.cell(r, 3, nm).border = BORDER
            for j, v in zip((4, 5), (round(j_amt, 2), round(l_amt, 2))):
                c = ws.cell(r, j, v)
                c.number_format = NUMFMT; c.border = BORDER; c.alignment = RGT
            ws.cell(r, 6, cnt).border = BORDER
            ws.cell(r, 7, '凭证级配对（同凭证本科目净额 vs 对方科目）').border = BORDER
            r += 1
    else:
        ws.cell(r, 1, comp).border = BORDER
        ws.cell(r, 7, '无对方科目数据（GL 无该科目凭证）').border = BORDER
    for j, w2 in enumerate([10, 14, 26, 14, 14, 8, 30], 1):
        ws.column_dimensions[chr(64 + j)].width = w2
    return ws


def _aux_analysis_sheet(wb, name, comp, aux_rows, ledger, is_liab):
    """⚡ 铁律73（2026-08-09 用户要求）：往来款【分析性程序】sheet——多角度审计风险提示。
    模块：①期初余额分布（期初导入/继承识别）②期初→期末变动 Top（变动率+大额）③大额单笔发生
    ④余额性质异常（资产贷余/负债借余=重分类提示）⑤集中度（Top5 占比）⑥本年新增/退出/休眠单位。
    数据源：应收/应付表（期初/借/贷/期末）+ GL 凭证级净额流水（已剔虚增虚减）。"""
    ws = wb.create_sheet('分析性程序')
    r = 1
    BOLD = Font(bold=True)
    HEAD2 = Font(bold=True, color='1F4E79')

    def _sec(txt):
        nonlocal r
        c = ws.cell(r, 1, txt)
        c.font = HEAD2
        r += 1

    def _hdr(cols):
        nonlocal r
        for j, h in enumerate(cols, 1):
            c = ws.cell(r, j, h)
            c.font = BOLD; c.border = BORDER
        r += 1

    def _row(vals):
        nonlocal r
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.border = BORDER
            if isinstance(v, (int, float)):
                c.number_format = NUMFMT
        r += 1

    # 单位级聚合（期初/借/贷/期末/合同号）
    from collections import defaultdict
    units = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0, ''])   # 编码 → [期初,借,贷,期末,名称]
    for code, nm, cc, cn, con, qc, jf, df, qm in aux_rows:
        a = units[cc]
        a[0] += qc; a[1] += jf; a[2] += df; a[3] += qm
        a[4] = cn
    # GL 单笔流水（净额化后）
    big_rows = [(dt, vno, txt, code, unit, amt) for dt, vno, txt, code, unit, amt in ledger
                if abs(amt) >= 1000000]

    # ① 期初余额分布（期初导入/继承识别）
    _sec('① 期初余额分布（期初=上年结转/导入余额）')
    _hdr(['往来单位', '期初余额', '期末余额', '说明'])
    n = 0
    for cc, (qc, jf, df, qm, cn) in sorted(units.items(), key=lambda x: -abs(x[1][0])):
        if abs(qc) >= 1000000:
            _row([cn or cc, round(qc, 2), round(qm, 2), '期初大额：上年结转余额，关注账龄与可回收性'])
            n += 1
            if n >= 10:
                break
    if n == 0:
        _row(['无期初 ≥100万 的单位', '', '', ''])
    r += 1

    # ② 期初→期末变动 Top（变动率+大额）
    _sec('② 期初→期末变动分析（大额变动=关注增减原因）')
    _hdr(['往来单位', '期初余额', '期末余额', '变动额', '变动率%', '提示'])
    n = 0
    for cc, (qc, jf, df, qm, cn) in sorted(units.items(), key=lambda x: -abs(x[1][3] - x[1][0])):
        dv = qm - qc
        if abs(dv) >= 1000000:
            rate = (dv / qc * 100) if abs(qc) > 0.005 else 999.0
            tip = '余额由贷转借' if (qc < 0 and qm > 0) else ('余额由借转贷' if (qc > 0 and qm < 0) else '大额增减')
            _row([cn or cc, round(qc, 2), round(qm, 2), round(dv, 2), round(rate, 1), tip])
            n += 1
            if n >= 10:
                break
    if n == 0:
        _row(['无变动 ≥100万 的单位', '', '', '', '', ''])
    r += 1

    # ③ 大额单笔发生
    _sec(f'③ 大额单笔发生（单笔 ≥100万，共 {len(big_rows)} 笔）')
    _hdr(['凭证日期', '凭证编号', '摘要', '金额', '单位'])
    for dt, vno, txt, code, unit, amt in sorted(big_rows, key=lambda x: -abs(x[5]))[:15]:
        _row([dt, vno, txt[:30], round(amt, 2), unit])
    if not big_rows:
        _row(['无单笔 ≥100万 发生', '', '', '', ''])
    r += 1

    # ④ 余额性质异常（重分类提示）
    _sec('④ 余额性质异常（资产科目贷余/负债科目借余 → 重分类提示）')
    _hdr(['往来单位', '期末余额', '性质', '建议'])
    n = 0
    for cc, (qc, jf, df, qm, cn) in sorted(units.items(), key=lambda x: -abs(x[1][3])):
        if (not is_liab and qm < -0.005) or (is_liab and qm > 0.005):
            sug = '重分类至预收账款' if not is_liab else '重分类至预付账款'
            _row([cn or cc, round(qm, 2), '贷方余额' if not is_liab else '借方余额', sug])
            n += 1
    if n == 0:
        _row(['无余额性质异常', '', '', ''])
    r += 1

    # ⑤ 集中度（Top5 占比）
    _sec('⑤ 集中度（Top5 单位期末余额占比）')
    ranked = sorted(units.items(), key=lambda x: -abs(x[1][3]))
    tot_qm = sum(abs(a[3]) for a in units.values())
    if tot_qm > 0.005:
        _hdr(['排名', '往来单位', '期末余额', '占比%', '累计占比%'])
        cum = 0.0
        for i, (cc, (qc, jf, df, qm, cn)) in enumerate(ranked[:5], 1):
            pct = abs(qm) / tot_qm * 100
            cum += pct
            _row([i, cn or cc, round(qm, 2), round(pct, 1), round(cum, 1)])
    else:
        _row(['期末余额合计为 0', '', '', '', ''])
    r += 1

    # ⑥ 本年新增/退出/休眠
    _sec('⑥ 本年新增/退出/休眠单位（无发生=关注真实性）')
    _hdr(['类型', '单位数', '期末余额合计', '说明'])
    new = [(cc, a) for cc, a in units.items() if abs(a[0]) < 0.005 and abs(a[3]) >= 0.005]
    out_ = [(cc, a) for cc, a in units.items() if abs(a[0]) >= 0.005 and abs(a[3]) < 0.005]
    sleep = [(cc, a) for cc, a in units.items() if abs(a[1] + a[2]) < 0.005 and abs(a[0]) >= 0.005]
    _row(['本年新增（期初0期末有余额）', len(new), round(sum(a[3] for _, a in new), 2), '关注新客户信用风险'])
    _row(['本年结清退出（期初有期末0）', len(out_), round(sum(a[3] for _, a in out_), 2), '关注结清方式与回款'])
    _row(['休眠无发生（期初=期末）', len(sleep), round(sum(a[3] for _, a in sleep), 2), '关注长期挂账与坏账风险'])
    for j, w2 in enumerate([26, 16, 16, 16, 14, 18], 1):
        ws.column_dimensions[chr(64 + j)].width = w2
    ws.freeze_panes = 'A2'
    return ws


def _aux_top10_sheet(wb, name, aux_rows):
    """U8 风格 Top10 余额（按期末余额绝对值降序，单位×合同号）。"""
    ws = wb.create_sheet('Top10余额')
    hdrs = ['序号', '往来单位编号', '往来单位名称', '合同号', '款项性质', '期初余额', '本期借方', '本期贷方', '期末余额']
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(1, j, h)
        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
    rows = sorted(aux_rows, key=lambda x: -abs(x[8]))
    for i, (code, nm, cc, cn, con, qc, jf, df, qm) in enumerate(rows[:10], 1):
        for j, v in enumerate([i, cc, cn, con or '—', nm, qc, jf, df, qm], 1):
            c = ws.cell(i + 1, j, v)
            if isinstance(v, (int, float)):
                c.number_format = NUMFMT
            c.border = BORDER
    for j, w2 in enumerate([6, 14, 26, 16, 18, 14, 14, 14, 14], 1):
        ws.column_dimensions[chr(64 + j)].width = w2
    return ws


def load_tb(data_dir):
    tb = {}
    fp = os.path.join(data_dir, '自建试算表_2026.xlsx')
    if not os.path.exists(fp):
        return tb
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    if '试算表' not in wb.sheetnames:
        wb.close()
        return tb
    ws = wb['试算表']
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    comps = [str(h) for h in hdr if str(h).isdigit()]
    for r in rows:
        name = str(r[0] or '').strip()
        if not name:
            continue
        tb[name] = {}
        for j, c in enumerate(comps, 1):
            v = r[j]
            if isinstance(v, (int, float)):
                tb[name][c] = float(v)
    wb.close()
    return tb


def build_workbook(data_dir, comp, out_dir=None):
    tb = load_tb(data_dir)
    out_dir = out_dir or os.path.join(data_dir, '底稿')
    os.makedirs(out_dir, exist_ok=True)
    outs = []
    for key, pfx, name, nature, aux_src in SUBJECTS:
        is_liab = (nature == 'liab')
        stat, inc, dec, exc_rows = analyze_subject(data_dir, comp, pfx, is_liab)
        tb_end = tb.get(name, {}).get(comp, 0.0)
        if abs(tb_end) < 0.5 and not stat:
            continue
        # 期初=TB qc（铁律24：期初=TB qc 权威，不反推——GL 凭证级净额化≠TB 发生额）
        qc_tb, _qm_tb, qc_exist = C.tb_begin_end(data_dir, comp, [name])
        begin = qc_tb if qc_exist else (tb_end - inc + dec)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        # Sheet1 审定表
        ws = wb.create_sheet('审定表')
        hdrs = ['公司', '期初余额', '本期增加', '本期减少', '期末余额']
        for j, h in enumerate(hdrs, 1):
            c = ws.cell(1, j, h)
            c.font = HEAD_FONT
            c.fill = HEAD_FILL
            c.alignment = CTR
            c.border = BORDER
        _w = [comp, begin, inc, dec, tb_end]
        for j, v in enumerate(_w, 1):
            cc = ws.cell(2, j, round(v, 2) if isinstance(v, (int, float)) else v)
            if isinstance(v, (int, float)):
                cc.number_format = NUMFMT
            cc.border = BORDER
            cc.alignment = RGT if j > 1 else LFT
        # ⚡ 2026-08-09 用户原则：原值科目与备抵类放一起，净值可与报表核对（U8 标准）
        # 坏账准备按子目归属本科目：123101应收/123102其他应收/123103合同资产/123104应收票据
        _BD_SUB = {'应收账款': '123101', '其他应收款': '123102',
                   '合同资产': '123103', '应收票据': '123104'}
        bd_pfx = _BD_SUB.get(name)
        if bd_pfx:
            try:
                # ⚡ 2026-08-10 修复断裂引用：原 from sap_longterm_assets_detail import _load_tb_sub_all
                # （该旧生成器已归档 _old_sap_gen_20260809）→ 改用 sap_common.load_tb_sub（仍在）
                _tbs = C.load_tb_sub(data_dir)
                bd_qc = sum(v['qc'] for (c, code), v in _tbs.items()
                            if c == comp and code.startswith(bd_pfx))
                bd_qm = sum(v['qm'] for (c, code), v in _tbs.items()
                            if c == comp and code.startswith(bd_pfx))
                if abs(bd_qm) > 0.005 or abs(bd_qc) > 0.005:
                    # 减：坏账准备（贷余科目显示正数，与 U8 一致）
                    _r = 3
                    c0 = ws.cell(_r, 1, '减：坏账准备')
                    c0.font = BOLD; c0.border = BORDER
                    for jj, vv in zip((2, 5), (-bd_qc, -bd_qm)):
                        cc = ws.cell(_r, jj, round(vv, 2))
                        cc.number_format = NUMFMT; cc.border = BORDER; cc.alignment = RGT
                    ws.cell(_r, 3, 0.0).border = BORDER
                    ws.cell(_r, 4, 0.0).border = BORDER
                    # 净额（原值 + 备抵，备抵为负 → 自动得净值）
                    _r2 = 4
                    c0 = ws.cell(_r2, 1, f'{name}净额')
                    c0.font = BOLD; c0.border = BORDER
                    net_qc = begin + bd_qc
                    net_qm = tb_end + bd_qm
                    for jj, vv in zip((2, 3, 4, 5), (net_qc, inc, dec, net_qm)):
                        cc = ws.cell(_r2, jj, round(vv, 2))
                        cc.number_format = NUMFMT; cc.border = BORDER; cc.alignment = RGT
                        cc.font = BOLD
                    # ⚡ 2026-08-09 U8 对齐：坏账准备明细表（期初/计提/转回/期末，随本科目底稿生成）
                    _BD_LABEL = {'123101': '应收账款坏账准备', '123102': '其他应收款坏账准备',
                                 '123103': '合同资产坏账准备', '123104': '应收票据坏账准备'}
                    _wbd = wb.create_sheet('坏账准备明细表')
                    for j, h in enumerate(['科目', '期初余额', '本期计提(贷)', '本期转回(借)', '期末余额'], 1):
                        c = _wbd.cell(1, j, h)
                        c.font = HEAD_FONT; c.fill = HEAD_FILL; c.alignment = CTR; c.border = BORDER
                    _rb = 2
                    _bd_tot = [0.0, 0.0, 0.0, 0.0]
                    for _cc, _nm in sorted(_BD_LABEL.items()):
                        _qc = sum(v['qc'] for (c, code), v in _tbs.items()
                                  if c == comp and code.startswith(_cc))
                        _qm = sum(v['qm'] for (c, code), v in _tbs.items()
                                  if c == comp and code.startswith(_cc))
                        if abs(_qc) < 0.005 and abs(_qm) < 0.005:
                            continue
                        # 计提=贷方(负)、转回=借方(正)：从 GL 按子目码聚合
                        _jf = _df = 0.0
                        for _r in C.iter_comp_gl(data_dir, comp, keep_cols={6, 10, 11, 29}):
                            if str(_r.get(11) or '').startswith(_cc):
                                _amt = float(_r.get(29) or 0)
                                if _amt > 0:
                                    _jf += _amt          # 借方=转回
                                else:
                                    _df += -_amt         # 贷方=计提
                        _wbd.cell(_rb, 1, _nm).border = BORDER
                        for _jj, _vv in zip((2, 3, 4, 5), (_qc, _df, _jf, _qm)):
                            _cc2 = _wbd.cell(_rb, _jj, round(_vv, 2))
                            _cc2.number_format = NUMFMT; _cc2.border = BORDER; _cc2.alignment = RGT
                        _bd_tot[0] += _qc; _bd_tot[1] += _df; _bd_tot[2] += _jf; _bd_tot[3] += _qm
                        _rb += 1
                    _wbd.cell(_rb, 1, '合计').font = BOLD
                    _wbd.cell(_rb, 1).fill = HEAD_FILL
                    for _jj, _vv in enumerate(_bd_tot, 2):
                        _cc2 = _wbd.cell(_rb, _jj, round(_vv, 2))
                        _cc2.number_format = NUMFMT; _cc2.font = BOLD; _cc2.fill = HEAD_FILL
                        _cc2.alignment = RGT
                    for _jj in range(1, 6):
                        _wbd.cell(_rb, _jj).border = BORDER
                    for _jj, _ww in enumerate([24, 16, 16, 16, 16], 1):
                        _wbd.column_dimensions[chr(64 + _jj)].width = _ww
            except Exception:
                pass
        for j, w2 in enumerate([12, 16, 16, 16, 16], 1):
            ws.column_dimensions[chr(64 + j)].width = w2
        # Sheet2 增减构成
        ws2 = wb.create_sheet('增减构成')
        _hdrs2 = ['增减类别', '金额']
        for j, h in enumerate(_hdrs2, 1):
            c = ws2.cell(1, j, h)
            c.font = HEAD_FONT
            c.fill = HEAD_FILL
            c.alignment = CTR
            c.border = BORDER
        r = 2
        for k, v in sorted(stat.items(), key=lambda x: -abs(x[1])):
            ws2.cell(r, 1, k).border = BORDER
            cc = ws2.cell(r, 2, round(v, 2))
            cc.number_format = NUMFMT
            cc.border = BORDER
            cc.alignment = RGT
            r += 1
        ws2.column_dimensions['A'].width = 30
        ws2.column_dimensions['B'].width = 15
        # Sheet3 异常凭证（审计重点关注，双向验证标记）
        ws3 = wb.create_sheet('异常凭证')
        hdrs3 = ['凭证编号', '科目净额影响', '凭证文本', '供应商', '业务性质']
        for j, h in enumerate(hdrs3, 1):
            c = ws3.cell(1, j, h)
            c.font = HEAD_FONT
            c.fill = HEAD_FILL
            c.alignment = CTR
            c.border = BORDER
        rr = 2
        for vno, netv, txtv, suppv, naturev in exc_rows:
            ws3.cell(rr, 1, vno).border = BORDER
            cc = ws3.cell(rr, 2, netv)
            cc.number_format = NUMFMT
            cc.border = BORDER
            cc.alignment = RGT
            ws3.cell(rr, 3, txtv).border = BORDER
            ws3.cell(rr, 4, suppv).border = BORDER
            c5 = ws3.cell(rr, 5, naturev)
            c5.border = BORDER
            c5.alignment = LFT
            c5.fill = PatternFill('solid', fgColor='E2EFDA' if naturev.startswith('正常')
                                  else 'FFC7CE')
            rr += 1
        for j, w2 in enumerate([14, 16, 40, 14, 34], 1):
            ws3.column_dimensions[chr(64 + j)].width = w2

        # ⚡ 2026-08-09：Sheet4 明细表 + Sheet5 账龄分布 + Sheet6 对方科目核对 + Sheet7 Top10
        #   （对齐 U8 底稿结构；数据源=应收/应付辅助核算；明细按 单位×合同号 分列）
        want_asset = (aux_src == 'cust')
        aux_rows = [a for a in _aux_rows(data_dir, comp, want_asset)
                    if a[0].startswith(pfx)]
        if aux_rows:
            lines = [l for l in _line_items(data_dir, comp, want_asset)
                     if l[0].startswith(pfx)]
            _aux_detail_sheet(wb, name, comp, aux_rows, lines, tb_qc=begin, tb_qm=tb_end)
            # ⚡ 铁律73（2026-08-09）：往来明细账——GL 逐笔流水（客户/供应商 100% 覆盖），账簿式列示
            #   2202 应付 GL 仅 14% 带供应商 → 改用供应商行项目数据源（自带供应商编码，2026 过滤后与应付表净额一致）
            if pfx == ('2202',) and os.path.exists(os.path.join(data_dir, '供应商行项目', f'{comp}.xlsx')):
                ledger = _supp_line_ledger(data_dir, comp, '2202')
            else:
                ledger = _gl_ledger_rows(data_dir, comp, pfx, want_asset)
            if ledger:
                _aux_ledger_sheet(wb, name, comp, aux_rows, ledger, is_liab=is_liab, tb_qm=tb_end)
                _aux_analysis_sheet(wb, name, comp, aux_rows, ledger, is_liab=is_liab)
            # ⚡ 2026-08-16 旧『账龄分布』已废弃：由专项 aging_review.py（应收+合同资产联动）替代，
            #   SAP（AH）迁移后接入专项；此处不再输出账龄分布 sheet
            _aux_opp_sheet(wb, name, comp, lines, opp=_opp_pairs(data_dir, comp, pfx))
            _aux_top10_sheet(wb, name, aux_rows)
        else:
            # 无辅助核算文件（如应付票据 2201/合同负债 2205）→ 空明细表占位
            ws4 = wb.create_sheet(f'{name}明细表')
            c = ws4.cell(1, 1, f'{name}：无辅助核算明细（SAP 无对应行项目文件）')
            c.font = HEAD_FONT

        fp = os.path.join(out_dir, f'往来款审计底稿_{name}_{comp}.xlsx')
        wb._meta = dict(comp=comp, label=name, period='2026年度1-7月')
        # ⚡ 2026-08-09：附注汇总（往来科目×期末，用户要求所有底稿都生成附注汇总→Word附注；
        #   优先用明细表往来单位，无辅助核算时科目单行占位）
        try:
            from sap_u8_sheets import u8_footnote_sheet
            foot = {}
            aux_hit = [a for a in _aux_rows(data_dir, comp, aux_src == 'cust')
                       if a[0].startswith(pfx)]
            if aux_hit:
                # 按往来单位×期末（idx3=单位名称、idx8=期末；无单位名用科目名/编码）
                for a in aux_hit:
                    nm = str(a[3] or '') if len(a) > 3 else ''
                    amt = float(a[8] or 0) if len(a) > 8 else 0.0
                    if not nm:
                        nm = str(a[1] or '')
                    if abs(amt) > 0.005:
                        foot[nm] = foot.get(nm, 0.0) + amt
                u8_footnote_sheet(wb, name, foot, comp)
            elif abs(tb_end) > 0.005:
                u8_footnote_sheet(wb, name, {name: tb_end}, comp)
        except Exception:
            pass
        try:
            from audit_shell import finalize_workbook as _fw
            _fw(wb)
        except Exception:
            pass
        wb.save(fp)
        wb.close()
        outs.append(fp)
        print(f'  ✅ {name}[{comp}]: 期末={tb_end:,.2f} 增={inc:,.2f} 减={dec:,.2f}')
    return outs


def main(argv=None):
    data_dir = sys.argv[1] if len(sys.argv) > 1 else P.YY
    comp = sys.argv[2] if len(sys.argv) > 2 else '3010'
    out_dir = None
    if '--out' in sys.argv:
        out_dir = sys.argv[sys.argv.index('--out') + 1]
    if comp == 'all':
        # 逐家子进程隔离（防内存累积 OOM：117 家单进程在 1050 附近被杀）
        import json as _json
        import subprocess
        tb = load_tb(data_dir)
        comps = sorted({c for m in tb.values() for c in m})
        outs = []
        for c in comps:
            r = subprocess.run([sys.executable, os.path.abspath(__file__),
                                data_dir, c] + ([out_dir] if out_dir else []),
                               capture_output=True, text=True, encoding='utf-8',
                               errors='replace')
            if r.returncode != 0:
                print(f'  ⚠️ [{c}] 子进程失败: {r.stderr.strip()[:80]}', flush=True)
                continue
            outs.append(c)
            print(f'  {c} 完成', flush=True)
        print(f'完成 {len(outs)} 家 / {len(comps)}')
    else:
        outs = build_workbook(data_dir, comp, out_dir)
        print(f'完成 {len(outs)} 份')


if __name__ == '__main__':
    main()
