# -*- coding: utf-8 -*-
"""sap_common.py —— SAP 底稿生成器共享骨架（2026-08-08）。

提供：样式常量、GL 文件定位、自建试算表读取、底稿写入辅助、勾稽核对公共逻辑。
新生成器（sap_bank_detail / sap_inventory_detail / sap_loan_detail /
sap_equity_detail）复用本模块，保证全 SAP 底稿风格统一 + 勾稽核对标准化。
"""
import os
import re
import sys
import pickle
import zipfile
import datetime
from collections import defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)
import _safx
from sap_reader import detect_sap_layout  # ⚡ 2026-08-10 布局检测（300 形态 GL 读取用；sap_reader 不反向依赖本模块）

# ---- 样式 ----
FONT = Font(name='Times New Roman', size=10)
BOLD = Font(name='Times New Roman', size=10, bold=True)
HEAD_FONT = Font(name='Times New Roman', size=10, bold=True)
HEAD_FILL = PatternFill('solid', fgColor='DDEBF7')
TOT_FILL = PatternFill('solid', fgColor='FCE4D6')
OK_FILL = PatternFill('solid', fgColor='E2EFDA')     # 勾稽通过（绿）
DIFF_FILL = PatternFill('solid', fgColor='FFC7CE')   # 勾稽差异（红）
NUMFMT = '#,##0.00'
CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
RGT = Alignment(horizontal='right', vertical='center')
LFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
_thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

MONTHS = 7
YEAR = '2026'

# ---- CO 结转凭证识别（2026-08-23 修订）----
# ⚡ 月末 CO 成本对象分摊/物料价格重估凭证：GL 序时账计绝对发生额、TB 部分科目按
#   净额/不含结转记录 → 两者差异主因。**识别仅凭『含 7042/7043 成本对象科目』**：
#   SAP 7042/7043 是 CO 内部成本对象科目，正常业务不会直接记，出现即内部结转。
# ⚡⚡ 2026-08-23 修订（勿用文本关键词）：此前用『多行净额0 + 文本含计提/摊销』误伤
#   真实计提凭证——3300/6100/6900 的『19』前缀 vno=1900000000『计提XXXX年XX月折旧和
#   摊销』（借费用6600 贷累计折旧1602，借贷成对）被标记 is_co，导致 GL 累计折旧贷方
#   从 0.50亿 掉到 0.05亿、费用明细失准。折旧/摊销/工资计提是真实业务（TB 含发生额），
#   **绝不能排除**；CO 结转只认 7042/7043 科目。GRIR 暂估/采购入库等真实业务同样保留。
#   落地：_scan_gl_full 输出加 is_co 标记（0/1），消费方按需过滤，不改默认行为。
_CO_MIN_ROWS = 8   # 保留（备用），CO 判定不再依赖行数

# ---- SAP 名称归一化（2026-08-08 修复：一级名变体拆分导致底稿漏算）----
# SAP 名称分层分隔符混用：- / —(U+2014) / －(U+FF0D) / \(U+005C)
_L1_SEP = re.compile(r'[-—－\\/]')
# 一级名变体 → 标准名（名称归一化后首段再映射）
_L1_MAP = {
    '银行存款人民币': '银行存款',
    '银行存款—人民币': '银行存款',
    '实收资本(或股本)': '实收资本',
}


def norm_l1(name):
    """SAP 一级名归一化：按多分隔符拆首段 + 变体映射。
    供自建试算表生成（sap_tb_gen._l1）与各生成器 TB 取数共用，保证同源。"""
    if not name:
        return name
    nm = str(name)
    # ⚡⚡ 2026-08-30 修复：投资性房地产-累计折旧/累计摊销/减值准备 应独立一级
    #   （与固定资产 1602『累计折旧』独立一致），而非并入『投资性房地产』——
    #   否则 Sheet1 投资性房地产=净值(原值-折旧 131.9M) vs 底稿原值(185.6M) 差异 54M。
    if '投资性房地产' in nm and ('累计折旧' in nm or '累计摊销' in nm):
        return '投资性房地产累计折旧'
    if '投资性房地产' in nm and '减值准备' in nm:
        return '投资性房地产减值准备'
    l1 = _L1_SEP.split(nm)[0].strip()
    return _L1_MAP.get(l1, l1)


def gl_dir(data_dir):
    return os.path.join(data_dir, '序时账')


def _gl_files(data_dir, comp):
    """返回该公司可能涉及的 GL 文件（前缀匹配 + 区间文件内容判断）。
    ⚡ 2026-08-10 300 形态：GL 在各账套子文件夹内（300/3700/3700序时账1-7月.XLSX）——
    直接扫 {data_dir}/{comp}/ 下的『序时账』文件（可多文件：1-6月+7月合并）。"""
    if detect_sap_layout(data_dir) == '300':
        dp = os.path.join(data_dir, comp)
        if not os.path.isdir(dp):
            return []
        # ⚡ 2026-08-12 3200/3300/3400 变体：无独立『序时账』文件，只有 GL 筛选形态的
        #   『应收借方/贷方发生额』『应付借方/贷方发生额』（含客户/供应商/科目/金额列，
        #   是序时账的科目筛选子集）→ 一并纳入 GL 数据源（往来明细表数据源）。
        # ⚡ 2026-08-22 3400 变体：txt 序时账（GBK 制表符分隔，含 3400序时账1月.txt 等，
        #   41/42/43 列混合——41列=普通序时账行、42列=含总账科目补充列、43列=表头/CO
        #   成本对象分摊行）→ 一并纳入（_scan_gl_txt 专用解析）。
        # ⚡ 2026-08-22 3400 去重：txt 序时账【完整】涵盖发生额 xlsx 视图（发生额是科目
        #   筛选子集），两者同扫会重复计入 → 有『序时账*.txt』时优先只用 txt。
        # ⚡ 2026-08-23 推广到全部主体：凡有『序时账』文件（xlsx/txt）即只用序时账，
        #   排除应收/应付借/贷方发生额（都是序时账子集，同扫重复计入 → 3300/3500
        #   GL 曾虚增 63亿/178亿）。仅当该主体【无】序时账时才回退用发生额文件。
        _all = sorted(os.listdir(dp))
        _gl_all = [f for f in _all
                   if (f.lower().endswith(('.xlsx', '.xls', '.txt')) or f.lower().endswith('.XLSX'))
                   and ('序时账' in f or '发生额' in f or '借方' in f or '贷方' in f)]
        _voucher = [f for f in _all if ('序时账' in f)
                    and (f.lower().endswith(('.xlsx', '.xls', '.txt')) or f.lower().endswith('.XLSX'))]
        if _voucher:
            return [os.path.join(dp, f) for f in _voucher]
        return [os.path.join(dp, f) for f in _all
                if f.lower().endswith(('.xlsx', '.xls'))
                and ('发生额' in f or '借方' in f or '贷方' in f)]
    d = gl_dir(data_dir)
    if not os.path.isdir(d):
        return []
    out = []
    for fn in sorted(os.listdir(d)):
        if not fn.lower().endswith('.xlsx'):
            continue
        base = os.path.splitext(fn)[0]
        m = re.match(r'^(\d{4})', base)
        if not m:
            continue
        pre = m.group(1)
        msuf = re.search(r'-(\d{4})$', base)
        if pre == comp:
            out.append(os.path.join(d, fn))
        elif msuf and int(pre) <= int(comp) <= int(msuf.group(1)):
            out.append(os.path.join(d, fn))
    return out


# ⚡ 2026-08-09（性能-跨进程落盘缓存）：GL 解析结果 pickle 落盘缓存。
# 目的：①88 家全量跑时避免每家重复解析 856MB xlsx；②改生成器逻辑重跑时
#      数据层零重解析（只重算不重读）。失效：源文件 mtime+size 变化自动重解析。
# 缓存目录：{data_dir}/.gl_cache/
_GL_DISK_ENABLED = True


def _gl_disk_key(data_dir, comp, cols):
    # ⚡ 2026-08-11 P0 缓存瘦身：.pkl → .pklz（gzip 流式压缩，文本列压缩率高、IO 加载快）。
    #   后缀变化使旧 .pkl 缓存自然失效，无需手动清理。
    # ⚡⚡ 2026-08-23 修复：原用 hash(tuple(sorted(cols)))——Python hash() 对 tuple 带
    #   随机种子（PYTHONHASHSEED），【不同进程算出的 key 不同】→ 缓存永远命中不了 →
    #   每次触发全量重扫（3300 210 万行 xlsx → OOM 静默 kill → 往来明细表空）。
    #   改用 hashlib.md5 稳定哈希，跨进程/跨机器一致。
    # ⚡⚡ 2026-08-24 架构阶段3：key 生成收口到 cache_manager.gl_cache_key（含 _GL_CACHE_VER 版本前缀）。
    from cache_manager import gl_cache_key
    return gl_cache_key(data_dir, comp, cols)


def _gl_source_sig(data_dir, comp):
    """源文件签名（mtime+size）→ 缓存失效判断。
    ⚡⚡ 2026-08-23 修复：fp 必须 os.path.normpath 统一分隔符——调用方 data_dir 有的传
       'd:/…'（正斜杠）有的传 'd:\\…'（反斜杠），os.path.join 结果混合分隔符 → sig
       字符串不同 → 缓存永远命中不了 → 每次全量重扫（3300 210 万行 OOM 静默 kill）。
    ⚡⚡ 2026-08-24 架构阶段3：签名逻辑收口到 cache_manager.gl_source_sig。"""
    from cache_manager import gl_source_sig
    return gl_source_sig(data_dir, comp)


def _period_from_date(v):
    """从凭证日期值推期间（月份字符串 '1'-'12'）。兼容 _safx 原始值：
    Excel 序列数字（46029）、字符串 'YYYY-MM-DD'/'MM/DD/YYYY'、datetime 对象。
    无法识别返回 None。⚡ 2026-08-23 3500 模板无『记帐期间/期间』列（i_per 落回
    类型列 YS/SA），期间必须从凭证日期推导。"""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return str(v.month)
    if isinstance(v, datetime.date):
        return str(v.month)
    if isinstance(v, (int, float)):
        # Excel 序列日期（基准 1899-12-30，SAP 导出无 1900 闰年虚拟日）
        try:
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(v)))
            return str(d.month)
        except (ValueError, OverflowError):
            return None
    s = str(v).strip()
    if len(s) >= 10 and s[5:7].isdigit() and s[4:5] == '-':
        return str(int(s[5:7]))
    if len(s) >= 7 and s[:2].isdigit() and s[2:3] in ('/', '-'):
        return str(int(s[:2]))
    # 数字字符串（'46029'）
    if s.isdigit() and len(s) >= 5:
        try:
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(s))
            return str(d.month)
        except (ValueError, OverflowError):
            return None
    return None


def _scan_gl_full(data_dir, comp, cols=None):
    """直接全扫该公司 GL（不经缓存，供 load_comp_gl_disk 使用）。
    ⚡ 2026-08-10 300 形态（FBL3N 多模板）：3700/6100/3900/6900 导出列序【不一致】
    （3700：科目码 col30/金额 col32；6100：科目码 col24/金额更靠后）→ 不能写死列号，
    每文件【表头探测】动态定列，输出逻辑键行 {comp,vno,per,date,code,name,amt,hs,sm}。
    期间=记帐期间列 ∈1-7 月；公司=公司代码列。
    ⚡ 2026-08-10 指令号扩展：cols 参数控制保留列集（默认 _GL_CACHE_COLS）；
    read_gl_rows 传扩展集（+17 WBS/24 订单/25 物料/26 项目/27 采购凭证）供
    inventory 按指令号（WBS/订单）分列期初/增/减/期末。"""
    _kc = set(cols or _GL_CACHE_COLS)
    from sap_reader import detect_sap_layout as _dsl
    if _dsl(data_dir) == '300':
        out = []
        for fp in _gl_files(data_dir, comp):
            if fp.lower().endswith('.txt'):
                # ⚡ 2026-08-22 3400 txt 序时账（GBK 制表符分隔）专用解析
                out.extend(_scan_gl_txt(fp, comp))
                continue
            cm, hdr_r, sheet_path = _gl_colmap300(fp)
            if not cm:
                continue
            # ⚡ 2026-08-10 列名变体（6900-1-6月 表头『求和项:金额』『公司码』『总账科目长文本』）
            i_comp = cm.get('公司代码', cm.get('公司码', cm.get('公司', 0)))
            i_per = cm.get('记帐期间', cm.get('期间', 4))
            i_vno = cm.get('凭证编号', 1)
            i_code = cm.get('总账科目', 30)
            i_name = cm.get('总账科目长文本', cm.get('长文本', cm.get('科目文本', None)))  # ⚡ 2026-08-12：3500 无名称列 → None（上层 resolve_subject_name 用 TB 映射补）
            i_amt = cm.get('求和项:金额', cm.get('金额', cm.get('本位币金额', cm.get('本币金额', 32))))
            i_date = cm.get('过帐日期', cm.get('过账日期', cm.get('凭证日期', 10)))
            i_hs = cm.get('借/贷标识', cm.get('借/贷', cm.get('记', 27)))
            i_sm = cm.get('凭证抬头文本', cm.get('抬头文本', cm.get('文本', 24)))
            # ⚡ 2026-08-10 往来客户/供应商列（表头动态定位，勿写死 40/41——4 账套列序不同：
            #   客户描述 3900=41/6100=34/6900=34，客户 3900=91/6100=77/6900=77）→ 明细表
            #   本期借/贷发生额按客户聚合的数据源。排除『贸易伙伴』列（含义不同）。
            i_cust = None
            for _k in ('客户描述', '客户'):
                if _k in cm:
                    i_cust = cm[_k]
                    break
            i_supp = None
            for _k in ('供应商描述', '供应商'):
                if _k in cm:
                    i_supp = cm[_k]
                    break
            # ⚡ 2026-08-10 指令号列（300 FBL3N 表头：订单/销售订单/项目/物料/采购凭证）：
            #   供 inventory 按指令号（订单/项目）增减变动表（与 yy 的 col17 WBS 体系不同，
            #   300 无 WBS 列、用订单列）。表头动态定位，缺列置 None。
            i_ord = cm.get('订单', cm.get('销售订单', None))
            i_proj = cm.get('项目', None)
            i_mat = cm.get('物料', None)
            i_po = cm.get('采购凭证', None)
            cols = {i_comp, i_per, i_vno, i_code, i_name, i_amt, i_date, i_hs, i_sm}
            if i_cust is not None:
                cols.add(i_cust)
            if i_supp is not None:
                cols.add(i_supp)
            for _ic in (i_ord, i_proj, i_mat, i_po):
                if _ic is not None:
                    cols.add(_ic)
            try:
                # ⚡ 2026-08-10 6900-1-6月：SAP『凭证小计报表』（非逐笔序时账）——数据行
                #   每凭证仅 1 行（单侧金额、借贷不配平），其余 57 万行是各级『汇总』行。
                #   判定：数据行数 ≈ 唯一凭证号数（>90%）→ 跳过该文件（明细源不可用，
                #   发生额以 TB 为准）。
                _rows_f = []
                _vnos_f = set()
                _n_data = 0
                cur_comp = None
                for r in _safx.iter_rows(fp, keep_cols=cols,
                                         # ⚡⚡ 2026-08-23 修复：_safx.min_row 是 0-based『跳过前 N 行』，
                                         #   hdr_r 是 1-based 表头行号 → 应传 hdr_r（跳表头），传 hdr_r+1
                                         #   会【多跳 1 行=表头后第一个数据行】——3500 各月文件第 2 行
                                         #   的『已过账』行被丢（vno=1700000004/1700000873 等），
                                         #   凭证只剩未清单边 → 借贷不平衡，GL 借贷差 +1.77 亿。
                                         min_row=(hdr_r if hdr_r else None),
                                         sheet=sheet_path):
                    # ⚡ 2026-08-10 6900-1-6月：公司码列合并单元格——仅每组首行有值
                    #   （1043 条数据中 1042 条 col0 为空）→ 留空时沿用上一非空值。
                    _cv = str(r.get(i_comp) or '').strip()
                    if _cv:
                        cur_comp = _cv
                    if cur_comp != comp:
                        continue
                    per = str(r.get(i_per) or '').strip()
                    if not (per.isdigit() and 1 <= int(per) <= MONTHS):
                        # ⚡ 2026-08-23 3500 模板无『记帐期间/期间』列（i_per 落回类型列
                        #   YS/SA/YF，值非数字）→ 从凭证日期推期间（_period_from_date
                        #   兼容 Excel 序列/字符串/datetime）。仍非法则跳过。
                        per = _period_from_date(r.get(i_date)) or ''
                    if not (per.isdigit() and 1 <= int(per) <= MONTHS):
                        continue
                    # ⚡⚡ 2026-08-23 归一化：期间列可能是『01』（前导零）而日期兜底返回
                    #   『1』（无前导零）→ 同文件内两种格式混存，分月明细会拆成两月。
                    #   统一 str(int()) 去前导零（'01'→'1'），保证期间一致。
                    per = str(int(per))
                    code = str(r.get(i_code) or '').strip()
                    if not code:
                        continue   # 汇总/说明行（科目空）跳过
                    _n_data += 1
                    _vnos_f.add(str(r.get(i_vno) or ''))
                    r[i_per] = per   # ⚡ 2026-08-23 期间兜底后写回（3500 无期间列，per 由日期推导）
                    _rows_f.append(r)
                if _n_data and len(_vnos_f) / _n_data > 0.9:
                    # ⚡ 2026-08-12 修复：凭证小计报表跳过增加例外——文件含客户/供应商列
                    #   且非空率高（≥50%，如 3200/3300/3400 应收借方发生额 65546 行每凭证
                    #   1 行但每行有客户编码）→ 是「凭证级客户发生额」，可作往来明细源
                    #   （借/贷文件互补覆盖双方发生额），不跳过。6900 混乱凭证小计（无客户
                    #   列或非空率低）仍跳过（明细源不可用，发生额以 TB 为准）。
                    _n_unit = 0
                    if i_cust is not None or i_supp is not None:
                        _n_unit = sum(1 for r in _rows_f
                                      if (str(r.get(i_cust) or '').strip()
                                          or str(r.get(i_supp) or '').strip()))
                    if _n_unit / _n_data < 0.5:
                        # 凭证小计报表（无客户/供应商维度）：每凭证 1 行，借贷不配平，不可作明细源
                        print(f'  ⚠️ [GL] {os.path.basename(fp)}：凭证小计报表（{_n_data} 行 '
                              f'{len(_vnos_f)} 张凭证，每凭证仅 1 行，客户/供应商非空率 '
                              f'{_n_unit / _n_data:.0%} < 50%）→ 跳过，1-7月发生额以 TB 为准')
                        _rows_f = []
                # ⚡ 2026-08-23 CO 结转凭证识别（仅含 7042/7043 成本对象科目）
                # ⚡⚡ 勿用文本关键词——『计提折旧和摊销』等真实计提会被误伤（见模块头）。
                _co_keys = set()
                if _rows_f:
                    _vno_rows = defaultdict(list)
                    for _r in _rows_f:
                        _vno_rows[(str(_r.get(i_per) or ''), str(_r.get(i_vno) or ''))].append(_r)
                    for _k, _rs in _vno_rows.items():
                        if any(str(_r.get(i_code) or '').startswith(('7042', '7043')) for _r in _rs):
                            _co_keys.add(_k)
                for r in _rows_f:
                    out.append({'comp': r.get(i_comp), 'vno': r.get(i_vno), 'per': r.get(i_per),
                                'date': r.get(i_date), 'code': str(r.get(i_code) or '').strip(),
                                'name': r.get(i_name), 'amt': r.get(i_amt),
                                'hs': r.get(i_hs), 'sm': r.get(i_sm),
                                'is_co': 1 if (str(r.get(i_per) or ''), str(r.get(i_vno) or '')) in _co_keys else 0,
                                'cust': str(r.get(i_cust) or '').strip() if i_cust is not None else '',
                                'supp': str(r.get(i_supp) or '').strip() if i_supp is not None else '',
                                'ord': str(r.get(i_ord) or '').strip() if i_ord is not None else '',
                                'proj': str(r.get(i_proj) or '').strip() if i_proj is not None else '',
                                'mat': str(r.get(i_mat) or '').strip() if i_mat is not None else '',
                                'po': str(r.get(i_po) or '').strip() if i_po is not None else ''})
            except Exception:
                continue
        return out
    rows = []
    for fp in _gl_files(data_dir, comp):
        try:
            for r in _safx.iter_rows(fp, keep_cols=_kc):
                if str(r.get(6) or '').strip() != comp:
                    continue
                per = str(r.get(10) or '').strip()
                if not (per[:2].isdigit() and 1 <= int(per[:2]) <= MONTHS):
                    continue
                rows.append(r)
        except Exception:
            continue
    return rows


# ⚡ 2026-08-22 3400 txt 序时账适配器
# 背景：3400 无 xlsx 序时账，只有 7 个月『3400序时账X月.txt』（GBK 编码、制表符分隔）。
# 结构（经与『发生额』文件交叉验证破解）：
#   - 行无表头（文件头部是报表标题区），数据行 40-43 列混合；
#   - 41 列行 = 普通序时账行（银行/费用等科目）；42 列行 = 往来行（应付/应收，
#     多出 col41=总账科目）；43 列行 = 表头行（col5=『凭证编号』文本）或 CO 成本
#     对象分摊行（7043 作业/分摊科目，S/H 成对、净额 0）——两者都按 col5 非数字跳过；
#   - 关键列固定（已确认）：col5=凭证号、col8=凭证日期(MM/DD/YYYY)、col10=金额(带符号)、
#     col17=借/贷标识(S=借/正,H=贷/负)、col27=公司、col29=科目码、col37=总账科目码
#     （同 col29，42列行与 col41 一致）、col38=期间、col40=过账日期(MM/DD/YYYY)。
# 方向约定：col17 'S'=借(金额正)、'H'=贷(金额负)——与发生额文件/科目余额表核对一致。
def _scan_gl_txt(fp, comp):
    """解析 3400 形态 txt 序时账（GBK），输出与 _scan_gl_full 300 分支相同的逻辑键行
    {comp,vno,per,date,code,name,amt,hs,sm}。逐行流式解析，7 个月约 1200 万行可跑。"""
    out = []
    try:
        fh = open(fp, encoding='gbk', errors='replace')
    except OSError:
        return out
    with fh:
        for ln in fh:
            if not ln or not ln.strip():
                continue
            cols = ln.rstrip('\n').rstrip('\r').split('\t')
            L = len(cols)
            if L < 40:
                continue
            v5 = cols[5].strip()
            if not (v5.isdigit() and len(v5) >= 6):
                continue  # 表头/标题行（col5 非纯数字）跳过；43 列表头与 CO 分摊行均落此
            # ⚡ 2026-08-22 CO 成本对象分摊行（col7='CO'，43列行，7042/7043 作业/分摊科目，
            #   借贷成对、净额 0，全量占 txt 89% 行量）→ 排除：属成本会计内部结转，
            #   非对外财务凭证；且 1230 万行全量会导致 GL 缓存 OOM（实测）。CO 科目
            #   发生额如需核对以科目余额表为准（read_tb_full 不含 GL 依赖）。
            if cols[7].strip() == 'CO':
                continue
            if cols[27].strip() != comp:
                continue
            per = cols[38].strip()
            if not (per.isdigit() and 1 <= int(per) <= MONTHS):
                continue
            code = cols[37].strip() or cols[29].strip()
            if not code:
                continue
            try:
                amt = float(cols[10].strip().replace(',', ''))
            except (ValueError, IndexError):
                continue
            hs = cols[17].strip()
            out.append({
                'comp': comp,
                'vno': v5,
                'per': per,
                'date': _txt_date(cols[40].strip()),   # 过账日期 MM/DD/YYYY → YYYY-MM-DD
                'code': code,
                'name': '',                            # 无名称列 → 上层用 TB 映射补
                'amt': amt,
                'hs': hs,
                'sm': cols[16].strip()[:80],           # 文本/抬头
            })
    return out


def _txt_date(s):
    """'MM/DD/YYYY' → 'YYYY-MM-DD'；无法解析原样返回。"""
    s = (s or '').strip()
    if len(s) >= 10 and s[2] == '/' and s[5] == '/':
        m, d, y = s[:2], s[3:5], s[6:10]
        if m.isdigit() and d.isdigit() and y.isdigit():
            return f'{y}-{m}-{d}'
    return s


def _gl_colmap300(fp):
    """300 形态：遍历【全部 sheet】找【表头行】→ (cm, hdr_r, sheet_path)：cm={表头文本:列号}，
    hdr_r=Excel 1-based 表头行号，sheet_path=该 sheet 的 xml 路径。
    ⚡ 多 sheet 选表：6900-1-6月 sheet 顺序=['Sheet2'(凭证小计报表), 'Sheet1'(标准逐笔序时账)]，
    两个 sheet 都有表头（Sheet2 表头在筛选条件后的行 9，Sheet1 表头行 1）→ 无法靠表头名区分，
    按【表头行号最小】选（标准导出表头在第 1 行；带筛选条件的报表表头靠后）。
    表头判定：含『总账科目』且含 金额/求和项:金额/本位币金额/凭证编号 之一。"""
    try:
        best = None   # (hdr_r, cm, sp)
        with zipfile.ZipFile(fp) as z:
            for sp in _safx._all_sheet_paths(z):
                hdr_r = 0
                for r in _safx.iter_rows(fp, keep_cols=set(range(100)), sheet=sp):
                    hdr_r += 1
                    cm = {str(v).strip(): i for i, v in r.items() if v is not None}
                    if '总账科目' in cm and ('金额' in cm or '求和项:金额' in cm
                                             or '本位币金额' in cm or '本币金额' in cm
                                             or '凭证编号' in cm):
                        if best is None or hdr_r < best[0]:
                            best = (hdr_r, cm, sp)
                        break   # 该 sheet 已定位表头
                    if hdr_r >= 12:
                        break
        if best:
            hdr_r, cm, sp = best
            return cm, hdr_r, sp
        return {}, None, None
    except Exception:
        return {}, None, None


def load_comp_gl_disk(data_dir, comp, cols=None):
    """读该公司 GL 全部行（列表），带落盘缓存。
    返回 [{col: value}, ...]（仅含 cols ∪ {6,10} 列）。
    ⚡ 注意：本函数返回共享缓存列表，调用方不得修改元素。"""
    cols = set(cols or _GL_CACHE_COLS) | {6, 10}
    if not _GL_DISK_ENABLED:
        return _scan_gl_full(data_dir, comp, cols)
    cache_dir = os.path.join(data_dir, '.gl_cache')
    try:
        os.makedirs(cache_dir, exist_ok=True)
    except OSError:
        return _scan_gl_full(data_dir, comp, cols)
    key = _gl_disk_key(data_dir, comp, cols)
    sig = _gl_source_sig(data_dir, comp)
    sig_fp = key + '.sig'
    # 命中检查
    if os.path.exists(key) and os.path.exists(sig_fp):
        try:
            with open(sig_fp, 'r', encoding='utf-8') as f:
                if f.read() == sig:
                    # ⚡ 2026-08-11：gzip 流式解压（内存峰值=最终列表，不翻倍）
                    import gzip
                    with gzip.open(key, 'rb') as f2:
                        return pickle.load(f2)
        except Exception:
            pass
    # 未命中 → 全扫 + 落盘
    rows = _scan_gl_full(data_dir, comp, cols)
    try:
        import gzip
        with gzip.open(key, 'wb', compresslevel=4) as f:
            pickle.dump(rows, f, protocol=pickle.HIGHEST_PROTOCOL)
        with open(sig_fp, 'w', encoding='utf-8') as f:
            f.write(sig)
    except Exception:
        pass
    return rows


def load_tb(data_dir, year=None):
    """读取自建试算表 → {科目名: {公司: 期末}}。year 支持 2025/2026 覆盖。"""
    tb = {}
    for yy in (['2026', '2025'] if year is None else [year]):
        fp = os.path.join(data_dir, f'自建试算表_{yy}.xlsx')
        if not os.path.exists(fp):
            continue
        wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        if '试算表' not in wb.sheetnames:
            wb.close()
            continue
        ws = wb['试算表']
        rows = ws.iter_rows(values_only=True)
        hdr = list(next(rows))
        comps = [str(h) for h in hdr if str(h).isdigit()]
        for r in rows:
            name = str(r[0] or '').strip()
            if not name:
                continue
            d = tb.setdefault(name, {})
            for j, c in enumerate(comps, 1):
                v = r[j]
                if isinstance(v, (int, float)):
                    d[c] = float(v)
        wb.close()
        if tb:
            break
    return tb


def load_tb_all(data_dir):
    """读取自建试算表 → {(公司, 科目名): 期末}。"""
    tb = load_tb(data_dir)
    out = {}
    for name, m in tb.items():
        for c, v in m.items():
            out[(c, name)] = v
    return out


# 进程级缓存：原始 TB（含 qc/qm/jf/df）
_TB_RAW_CACHE = {}


def load_tb_raw(data_dir, year='2026'):
    """读取原始 SAP TB → {(公司, 一级名归一化): {qc,qm,jf,df}}。
    期初权威=TB qc（铁律24），期末=TB qm。进程级缓存防重复解析。"""
    global _TB_RAW_CACHE
    key = (data_dir, year)
    if key in _TB_RAW_CACHE:
        return _TB_RAW_CACHE[key]
    try:
        import sap_reader as SR
        raw = SR.read_sap_tb(data_dir, None, year=year)
    except Exception:
        _TB_RAW_CACHE[key] = {}
        return _TB_RAW_CACHE[key]
    agg = {}
    for (comp, _code, name, _yy), v in raw.items():
        if comp not in agg:
            agg[comp] = {}
        l1 = norm_l1(name)
        a = agg[comp].setdefault(l1, {'qc': 0.0, 'qm': 0.0, 'jf': 0.0, 'df': 0.0})
        a['qc'] += v['qc']
        a['qm'] += v['qm']
        a['jf'] += v['jf']
        a['df'] += v['df']
    _TB_RAW_CACHE[key] = agg
    return agg


def tb_begin_end(data_dir, comp, names, year='2026'):
    """按一级名集合取该公司 TB 期初/期末（同源权威）。
    names: 归一化一级名（或前缀）列表 → (期初合计, 期末合计)。
    返回 (qc, qm, exists)：exists=该科目名是否真实存在于 TB（qc=0 也是权威）。"""
    raw = load_tb_raw(data_dir, year)
    ent = raw.get(comp, {})
    qc = qm = 0.0
    hit = False
    for l1, a in ent.items():
        if any(l1 == n or l1.startswith(n) for n in names):
            qc += a['qc']
            qm += a['qm']
            hit = True
    return qc, qm, hit


# 进程级缓存：原始 TB 子目级 {(comp, 科目代码): {qc,qm}}
_TB_SUB_CACHE = {}


def load_tb_sub(data_dir, year='2026'):
    """读取原始 SAP TB 子目级（直接从 xlsx 解析，保留内部码）
    → {(公司, 科目代码): {qc,qm,jf,df}}。用于存货/银行等子目级审定。
    注意：read_sap_tb 已剥离内部码，须直接读文件取 col2 的 'HY01/1406030000' 尾码。"""
    global _TB_SUB_CACHE
    key = (data_dir, year)
    if key in _TB_SUB_CACHE:
        return _TB_SUB_CACHE[key]
    import sap_reader as SR
    layout = SR.detect_sap_layout(data_dir)
    tb_dir = os.path.join(data_dir, '科目余额表') if layout == 'ah-sap' else data_dir
    out = {}
    files = []
    for f in sorted(os.listdir(tb_dir)):
        if f.lower().endswith('.xlsx'):
            files.append(os.path.join(tb_dir, f))
    for fp in files:
        try:
            wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
        except Exception:
            continue
        ws = wb[wb.sheetnames[0]]
        agg = {}   # (comp, code) -> [qc,jf,df,qm]
        for r in ws.iter_rows(min_row=3, values_only=True):
            if not r or len(r) < 17 or r[2] is None:
                continue
            try:
                name, code = SR.parse_sap_tb_name(r[2])
            except Exception:
                continue
            if not name or not code or any(k in name for k in SR._GL_DUP_KW):
                continue
            comp_raw = str(r[7]).strip() if r[7] else ''
            m = re.search(r'(\d{4})\s*$', comp_raw)
            if not m:
                continue
            comp = m.group(1)
            code = code.split('/')[-1] if '/' in code else code
            period = str(r[8]).strip() if r[8] else ''
            a = agg.setdefault((comp, code), [0.0, 0.0, 0.0, 0.0])
            if '剩余' in period:
                a[0] += float(r[9] or 0)
            else:
                a[1] += float(r[12] or 0)
                a[2] += float(r[13] or 0)
            if '七月' in period:
                a[3] += float(r[16] or 0)
        wb.close()
        for (comp, code), a in agg.items():
            out[(comp, code)] = {'qc': a[0], 'jf': a[1], 'df': a[2], 'qm': a[3]}
    _TB_SUB_CACHE[key] = out
    return out


_GL_ROWS_CACHE = {}          # {comp: [row_dict...]} 进程级 GL 行缓存（13 生成器共享一次扫描）
# ⚡ 2026-08-11 P0 修复：恢复指令号列全集（17/24/25/26/27）——此前精简为 11 列与磁盘
#   缓存 key（16 列集）不匹配 → 缓存永远 miss → 1010 每次全量重扫 89.6s → current_account 31min。
_GL_CACHE_COLS = {0, 1, 5, 6, 10, 11, 12, 17, 19, 21, 24, 25, 26, 27, 28, 29, 31}   # 常用列全集（分配/凭证/文本/公司/期间/科目/客户/供应商/WBS/功能范围/订单/物料/项目/采购凭证/凭证日期/金额/过账日期）
# ⚡ 2026-08-26 加 col31（过账日期）：read_gl_rows 期间口径过滤原用 col28（凭证日期），
#   把"凭证日期=2025年末、过账日期=2026"的冲销凭证（审计调整/冲回）误过滤 → 分月明细
#   与 TB 逐月对不上（如 660302 利息支出 3 月差 300 万=冲审计调整长期借款重分类）→ 改按 col31。
_GL_CACHE_VER = 'v1'   # ⚡ 2026-08-23 缓存格式版本：列集解析格式变化时 bump，旧缓存自动失效
# ⚡ 2026-08-10 300 形态（FBL3N 导出）列全集：0公司代码/1凭证编号/3财年/4记帐期间/
# 10过帐日期/24凭证抬头文本/27借/贷标识(H借/S贷)/30总账科目/31科目长文本/32金额(带符号)/
# 34本位币金额/37行项目文本/38利润中心/40客户/41客户描述
_GL_CACHE_COLS_300 = {0, 1, 3, 4, 10, 24, 27, 30, 31, 32, 34, 37, 38, 40, 41}


def iter_comp_gl(data_dir, comp, keep_cols=None):
    """遍历该公司全部 GL 行（生成器）。
    ⚡ 2026-08-09（铁律70）：内部过滤依赖 col6 公司/col10 期间，强制并入 keep_cols，
    防调用方漏传致静默 0 行（gp_other 曾中招）。
    ⚡ 2026-08-09（性能双层缓存）：进程级 _GL_ROWS_CACHE（同进程 13 生成器共享一次扫描）
    + 落盘 pickle 缓存（跨进程/跨次运行复用，源文件 mtime+size 变化自动重解析）。
    1010 大公司全量从 60-80min → ~2.5min；二次重跑数据层零重解析。"""
    keep_cols = set(keep_cols or ()) | {6, 10}
    rows = _GL_ROWS_CACHE.get(comp)
    if rows is None:
        # 落盘统一用全集（防同公司多份 pkl），进程缓存也存全集
        rows = load_comp_gl_disk(data_dir, comp, cols=_GL_CACHE_COLS)
        _GL_ROWS_CACHE[comp] = rows
    if keep_cols == _GL_CACHE_COLS:
        for r in rows:
            yield r
        return
    for r in rows:
        if keep_cols is not None and keep_cols != _GL_CACHE_COLS:
            yield {k: v for k, v in r.items() if k in keep_cols}
        else:
            yield r


def sheet_hdr(ws, hdrs, row=1, fill=None):
    """写表头行。"""
    for j, h in enumerate(hdrs, 1):
        c = ws.cell(row, j, h)
        c.font = HEAD_FONT
        c.fill = fill or HEAD_FILL
        c.alignment = CTR
        c.border = BORDER


def txt(ws, r, c, v, bold=False, fill=None):
    cell = ws.cell(r, c, v)
    cell.font = BOLD if bold else FONT
    if fill:
        cell.fill = fill
    cell.border = BORDER
    cell.alignment = LFT
    return cell


def money(ws, r, c, v, fill=None):
    cell = ws.cell(r, c, round(v, 2) if isinstance(v, (int, float)) else v)
    cell.number_format = NUMFMT
    cell.font = FONT
    cell.alignment = RGT
    cell.border = BORDER
    if fill:
        cell.fill = fill
    return cell


def fmt_cols(ws, widths, max_col=None):
    """设置列宽（支持超 26 列：AA/AB…；原实现 >26 列全写 A 列）。"""
    for j, w in enumerate(widths, 1):
        if max_col and j > max_col:
            break
        n = j - 1
        letters = ''
        while True:
            letters = chr(65 + n % 26) + letters
            n = n // 26 - 1
            if n < 0:
                break
        ws.column_dimensions[letters].width = w


def write_recon_sheet(wb, title, items, note=None):
    """写勾稽核对 sheet。

    items: [(项目, 左金额, 右金额, 差异口径描述)]，差异>0.01 红色标。
    note: 底部说明文本（可选）。
    """
    ws = wb.create_sheet(title)
    sheet_hdr(ws, ['勾稽项目', '金额A', '金额B', '差异(A-B)', '说明'])
    r = 2
    for name, a, b, desc in items:
        txt(ws, r, 1, name)
        money(ws, r, 2, a if a is not None else 0.0)
        money(ws, r, 3, b if b is not None else 0.0)
        diff = (a or 0.0) - (b or 0.0)
        cc = money(ws, r, 4, diff, OK_FILL if abs(diff) <= 0.01 else DIFF_FILL)
        txt(ws, r, 5, desc)
        r += 1
    if note:
        r += 1
        txt(ws, r, 1, note, bold=True)
    fmt_cols(ws, [36, 16, 16, 14, 52])
    return ws
