# -*- coding: utf-8 -*-
# FINGERPRINT: 产出=制造费用审计底稿/销售费用审计底稿/管理费用审计底稿/财务费用审计底稿/研发费用审计底稿_生成.xlsx | 关键列=二级科目全年/分月/与TB控制数 | 职责=费用二级科目发生额(GL汇总)与TB勾稽(独立生成器)

from audit_common import _safe_save, build_account_breakdown_rows, write_recon_prompt, voucher_skip_keys, _extract_year  # 共享库：锁感知保存 + 从一级出发行规范 + 勾稽提示 + 抽凭排除规则 + 稳健年份抽取
from fill_voucher_counterparty import derive_cp  # 凭证对方科目按金额精确推导（杜绝"一大堆"对方科目）
"""
expense_detail.py — 费用明细表编制小程序（5101/6601/6602/6603/6604）
=================================================================================
用法：把【账套导出文件夹】拖到同目录的 run_expense_detail.bat 上即可。
  · 自动发现文件夹内各核算主体的「科目余额表」(控制数) + 「综合查询明细表」(GL 凭证级)
  · 按 5 个费用一级科目（5101制造费用 / 6601销售费用 / 6602管理费用 / 6603财务费用 / 6604研发费用）
    各自生成一份底稿：<科目名称>审计底稿_生成.xlsx
  · 每份底稿沿用对应模板（中央 audit_templates/ 文件夹内的 5101~6604 外壳；
    可用环境变量 EXPENSE_TEMPLATE_DIR 覆盖）的表结构：
      说明 / <费用>明细表 / <年份>分月明细 / <年份>主体比较
    其中「<费用>明细表」按二级科目汇总（全年/同比/与科目余额表勾稽），
    「分月明细」为 核算主体×二级科目×月，「主体比较」为 二级科目×核算主体。
  · 模板缺失时自动改用内置默认空壳（格式降级，仍保证有输出），不再静默 0 文件。
  · 模板仅作“格式外壳”，数据全部按被拖文件夹的账套重新计算（可适用于任意集团文件夹）。
依赖：openpyxl（WorkBuddy 托管 Python 自带）
"""
import os
import re
import sys
import shutil
from collections import defaultdict, OrderedDict

# ---- 启动诊断日志（绝对路径，纯内置库，任何崩溃前必写，用于定位"闪退无痕迹"）----
def _write_boot_log(tag):
    import os as _bl_os
    if _bl_os.environ.get('AUDIT_LOG') != '1':
        return
    try:
        import datetime
        _p = os.path.join(os.path.expanduser('~'), 'Desktop', 'expense_boot.log')
        with open(_p, 'a', encoding='utf-8') as _f:
            _f.write('[%s] %s sys=%s argv=%s\n' % (
                datetime.datetime.now().isoformat(timespec='seconds'), tag, sys.executable, sys.argv))
    except Exception:
        pass

_write_boot_log('LOAD')
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


from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import sap_report as RPT   # ⚡⚡ 2026-08-31 费用目录（功能范围拆分）读取
# 统一外壳样式与加载逻辑（六程序共享）：蓝头 1F4E78 / 千分位 / 边框 / 中央 audit_templates/
from audit_shell import (
    SHELL_HFILL, SHELL_HFONT, SHELL_SUB_FONT, SHELL_BOLD, SHELL_NUM, SHELL_THIN,
    SHELL_BORDER, SHELL_CEN, SHELL_LEFT, SHELL_RGT, SHELL_TOT_FILL, SHELL_TITLE_FONT,
    resolve_template, AUDIT_TEMPLATES_DIR,
    finalize_workbook,
)

# 5 个费用类别：名称 → 一级标准代码（输出底稿按此生成，外壳文件名沿用中央 audit_templates 命名）
EXPENSE_CATS = OrderedDict([
    ('制造费用', 5101),
    ('销售费用', 6601),
    ('管理费用', 6602),
    ('财务费用', 6603),
    ('研发费用', 6604),
])

# 泰国账套二级科目名 → 集团标准二级科目名（2026-08-01 问题1a：主体比较表按二级内容对齐）。
# 实测泰国 5401 制造费用二级：员工支出/电费/水费/包装物/物料耗用/折旧费/环保支出/运费/认证费/
# 低值易耗品/维修费/检测费/厂房租金/安全生产费/叉车租赁费/其他/办公费/废水处理费；
# 母公司 5101 二级：职工薪酬/电费/水费/包装物/物料耗用/折旧费/安全生产费用/环保支出/运费/认证费/
# 差旅费/印刷费/辐照费/办公费/租赁费/低值易耗品/加工费/3D打印服务费/开模费/其他/审核费/邮寄费/
# 维修费/检测费/证书维护费/标签购买费/劳务外包费/水电费-上缴的各项税费。
# 仅映射"名称不同但内容同类"的二级（同名者天然对齐，无需映射）。
THAI_L2_NAME_MAP = {
    '员工支出': '职工薪酬',              # 泰国员工支出 = 母公司职工薪酬
    '厂房租金': '租赁费',                # 泰国厂房租金 ≈ 母公司租赁费
    '叉车租赁费': '租赁费',              # 泰国叉车租赁费 ≈ 母公司租赁费
    '安全生产费': '安全生产费用',        # 泰国安全生产费 = 母公司安全生产费用
    '其他-如果金额大请新建科目': '其他',  # 泰国兜底其他 = 母公司其他
    '废水处理费': '环保支出',            # 泰国废水处理费 ≈ 母公司环保支出
    '员工激励': '职工福利费',            # 泰国员工激励 ≈ 母公司职工福利费（三级层，挂靠映射）
    '研讨会和培训': '职工教育经费',      # 泰国研讨会和培训 ≈ 母公司职工教育经费
}
SUBJECTS = [(code, name) for name, code in EXPENSE_CATS.items()]
TARGET_CODES = set(EXPENSE_CATS.values())

# —— 类别判定一律以【一级科目名称】为准，绝不按科目代码归并 ——
# 本集团各账套代码体系不一致：例如 J 账套 4101=制造费用、g/c/x 账套 4101=盈余公积；
# J 账套 4001=生产成本(其他账套 5001)、5601/5602/5603=销售/管理/财务费用(其他账套 6601/6602/6603)。
# 若按代码前缀(5101=制造费用)硬归类，J 账套制造费用(4101)会被整体漏掉。
# 故：一级科目的类别 = _name_to_cat(该一级的科目名称)；二级科目的类别 = 其所属一级的名称所决定的类别。
CAT_NAME_KEYWORDS = {
    '制造费用': ['制造费用'],
    '销售费用': ['销售费用', '营业费用'],
    '管理费用': ['管理费用'],
    '财务费用': ['财务费用'],
    '研发费用': ['研发费用', '研发费', '研发支出'],
}


def _name_to_cat(name):
    if not name:
        return None
    s = str(name)
    # 研发支出 是资产负债表过渡/资本化科目，其借发(发生的研究开发费)已在期末结转计入研发费用，
    # 若一并归入费用类会与其自身双重计列；故 研发支出 不归入任何费用类。
    if '研发支出' in s:
        return None
    s0 = s.split('-')[0].strip()
    # 2026-08-06：兼容 U8 反斜杠层级（JTt/SSS 科目名『660214\管理费用\固定资产』，
    # '-' 首段是『660214』无汉字 → 换 '\' 首段『管理费用』）。
    if not any('\u4e00' <= ch <= '\u9fff' for ch in s0):
        s0 = s.split('\\')[0].strip()
    # 2026-08-02 修复：先按【一级名（首段）】判定——『管理费用-研发费用-基本工资』一级是管理费用，
    # 必须归管理费用（6602），与 TB 一级归类一致（FY 6602.33 管理费用-研发费用）。
    # 2026-08-05 用户明确：企业把研发费用放在管理费用项下核算时，明细表不自行修改（剔除/移动），
    # 按账套原样归管理费用。独立『研发费用』科目（首段即研发费用）仍归研发费用。
    for cat, kws in CAT_NAME_KEYWORDS.items():
        for kw in kws:
            if s0 == kw or s0.startswith(kw):
                return cat
    if '研发费用' in s or '研发费' in s or '研发' in s:
        return '研发费用'
    return None


# ⚡ 2026-08-26 配置化收敛：SAP 费用特判码收口 sap_subject_map（防漂移唯一依据）；
#   兜底默认值保证 import 失败/新模块未接入时行为不回归。
try:
    from sap_subject_map import fee_code as _sap_fee_code, fr_category as _sap_fr_cat
except Exception:
    _sap_fee_code = lambda k: {'rd': '660006', 'finance': '6603', 'pool': '6600',
                               'fx_loss': '660303', 'fx_gain': '660304'}.get(k)
    _sap_fr_cat = lambda fr: {'6601': '销售费用', '6602': '管理费用', '5101': '制造费用',
                              '5301': '研发费用'}.get(str(fr or '').strip())
_RD_CODE = _sap_fee_code('rd') or '660006'
_FIN_CODE = _sap_fee_code('finance') or '6603'
_POOL_CODE = _sap_fee_code('pool') or '6600'
_FX_LOSS_CODE = _sap_fee_code('fx_loss') or '660303'
_FX_GAIN_CODE = _sap_fee_code('fx_gain') or '660304'


def _sap_fee_cat(r):
    """SAP 费用行分类（2026-08-10 用户提醒"看功能列"——对齐旧 SAP 生成器 FR_CAT 语义）：
    ①科目 660006（研究开发费）不论功能范围一律研发费用；
    ②6603 系（财务费用）编码已定性；
    ③其余 6600 系列按【功能范围列】（col19：6601销售/6602管理/5101制造/5301研发）归类。
    SAP TB 编码 6600 是期间费用总池（含 管理/销售/制造/研发 全部明细科目），
    无功能范围列信息无法区分 → 只能从 GL 功能范围列拆分（TB 仅总额控制）。
    返回标准类别名（销售费用/管理费用/制造费用/研发费用/财务费用）或 None。"""
    code = str(r.get('code') or '')
    nm = str(r.get('name') or '')
    # ⚡ 2026-08-10 限定费用科目：仅 6600/6603 系列参与 FR 分类——SAP 成本/收入科目
    # （6402 主营业务成本等）也可能带功能范围（如 6602），误归会把成本行混入费用底稿
    if not (code.startswith(_POOL_CODE) or code.startswith(_FIN_CODE)):
        return None
    if code.startswith(_RD_CODE) or '研究开发费' in nm:
        return '研发费用'
    if code.startswith(_FIN_CODE) or nm.startswith('财务费用'):
        return '财务费用'
    fr = str(r.get('fr') or '').strip()
    _cat = _sap_fr_cat(fr)
    if _cat:
        return _cat
    # 功能范围缺失兜底：6600 系列默认管理费用（常见情形），6600 外的费用编码按名称
    if code.startswith(_POOL_CODE):
        return '管理费用'
    return _name_to_cat(nm)


# 模板所在目录（格式外壳来源）：统一走中央 audit_templates/（与数据文件夹分离，避免误删）。
# 可用环境变量 EXPENSE_TEMPLATE_DIR 覆盖外壳目录。gu股份公司 仅为输入数据账套包，不作为模板来源。
def _find_template(code, name):
    """定位模板（外壳）完整路径；查找顺序：环境变量目录 → 中央 audit_templates/。"""
    return resolve_template(f'{code}{name}', 'EXPENSE_TEMPLATE_DIR')


def _make_default_expense_wb(name, years):
    """模板缺失时的兜底空壳：含 明细表/各年分月明细/各年主体比较/说明，保证总有可用输出。"""
    from openpyxl.utils import get_column_letter as _gcl
    hf = SHELL_HFONT
    hfill = SHELL_HFILL
    bd = SHELL_BORDER
    cen = SHELL_CEN
    yrs = sorted(years)
    y_cur = yrs[-1]
    y_prev = yrs[-2] if len(yrs) >= 2 else None
    if y_prev:
        h_sum = ['二级科目名称', f'{y_prev}全年(TB)', f'{y_cur}全年(TB)', '增减额', '变动率%']
    else:
        h_sum = ['二级科目名称', f'{y_cur}全年(TB)']
    h_mon = ['核算主体', '二级科目名称'] + [f'{m}月' for m in range(1, 13)] + \
            ['合计(净发生额)', '借方发生额(备查)', 'TB控制数', '差异']
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(f'{name}明细表')
    for c, h in enumerate(h_sum, 1):
        cell = ws.cell(1, c, h); cell.font = hf; cell.fill = hfill; cell.alignment = cen; cell.border = bd
    ws.freeze_panes = 'A2'
    for y in sorted(years):
        ws2 = wb.create_sheet(f'{y}分月明细')
        for c, h in enumerate(h_mon, 1):
            cell = ws2.cell(1, c, h); cell.font = hf; cell.fill = hfill; cell.alignment = cen; cell.border = bd
        ws2.freeze_panes = 'D2'
        ws3 = wb.create_sheet(f'{y}主体比较')
        for c, h in enumerate(['二级科目名称', '核算主体', '全集团合计'], 1):
            cell = ws3.cell(1, c, h); cell.font = hf; cell.fill = hfill; cell.alignment = cen; cell.border = bd
    return wb

# ---------------- 样式（统一为 audit_shell 调色板） ----------------
HDR_FONT = SHELL_HFONT
HDR_FILL = SHELL_HFILL
SUB_FONT = SHELL_SUB_FONT
TOT_FONT = SHELL_BOLD
TOT_FILL = SHELL_TOT_FILL
NUM_FMT = SHELL_NUM
THIN = SHELL_THIN
BORDER = SHELL_BORDER
CENTER = SHELL_CEN
LEFT = SHELL_LEFT
RIGHT = SHELL_RGT


def safe(x):
    try:
        return float(x)
    except Exception:
        return 0.0


def parse_month(dt):
    if dt is None:
        return None
    if isinstance(dt, int):
        return None
    s = str(dt)
    m = re.search(r'-(\d{1,2})-', s)
    if m:
        return int(m.group(1))
    m2 = re.search(r'(\d{4})[-/](\d{1,2})', s)
    if m2:
        return int(m2.group(2))
    return None


# ---------------- 发现账套实体 ----------------
def _entity_name(fn):
    """从文件名稳健抽取实体名：先剥年份+期间后缀再剥科余表等关键字。个位数前缀补0。"""
    base = re.sub(r'(19|20)\d{2}', '', fn)
    base = re.sub(r'年度|年\d*[-~]?\d*月', '', base)  # 剔除"年度"、"年1-3月"等期间后缀
    E = re.sub(r'(外币)?(科目余额表|科目余额|余额表|综合查询明细表|综合查询表|辅助核算余额表|辅助核算|总账|trial).*$',
                  '', base, flags=re.IGNORECASE).strip(' _-')
    E = E.rstrip('年')  # 2026-08-06：U8『主体2025年科目余额表』去年份后残留孤立『年』
    # 个位数前缀补0（如 1、FY本级公司 → 01、FY本级公司）
    _m0 = re.match(r'^(\d)([、,，.．\s])', E)
    if _m0 and len(E) > 2:
        E = '0' + E
    return E


def _discover_entities(data_dir):
    # ⚡ 2026-08-09 治本统一：本地 discover 收敛为 audit_common.discover_entities——
    # U8 模式读原文件全量；SAP 模式 patch 后=adapter（current_comp 单主体过滤内置）。
    from audit_common import discover_entities as _de
    return _de(data_dir)


# ---------------- 读科目余额表：一级名→码、二级树、二级借发控制数 ----------------
def read_all_tb(data_dir, entities, years):
    # SAP：从 adapter.read_km 一级聚合行构建（名称首段=一级名；二级名从末级行 code[:6] 去重）
    if _adapter is not None and _adapter.is_sap(data_dir):
        name2code = {}
        code2name = {}
        tb_l2_names = defaultdict(dict)
        tb_l3_names = defaultdict(dict)
        tb_l2_control = {}
        tb_l1_control = {}
        for _cat, _code in EXPENSE_CATS.items():
            name2code[_cat] = _code
            code2name[_code] = _cat
        # ⚡⚡ 2026-08-31 AH 类账套功能范围拆分：6600 期间费用总池（工资及奖金等）无
        #   管理/销售/研发/制造独立科目 → 读账套费用目录 费用/{e}/{类别}N.xlsx
        #   （N=最新月份累计；每文件『合计：』行=企业利润表科目，已验证精确一致）。
        _fee_scope = {}
        # ⚡⚡ 2026-08-31 财务费用（6603）口径：账套企业利润表财务费用=净额（利息收入
        #   冲减），TB 借发双计（利息收入贷方 8M + 支出借方 72M）→ 有企业报表的主体
        #   覆盖为『财务费用』企业报表值（与试算表 Sheet3 同源一致）。
        _fee_fin = {}
        _has_rep = RPT.has_enterprise_reports(data_dir)
        for e in entities:
            _fs = RPT.read_fee_scope(data_dir, e)
            if _fs:
                _fee_scope[e] = _fs
            if _has_rep:
                _pl = RPT.read_ent_profit(data_dir, e)
                if abs(float(_pl.get('财务费用', 0.0))) > 0.005:
                    _fee_fin[e] = float(_pl.get('财务费用', 0.0))
                # ⚡⚡ 2026-08-31 逐类别补全：费用目录缺某类别（如 1030 有管理/销售/
                #   制造文件但无研发费用文件）但企业利润表有值 → 补企业报表值（与
                #   试算表 Sheet3 同源一致）；完全无费用目录的主体同样回退企业报表。
                for _cat in ('管理费用', '销售费用', '研发费用'):
                    if (_cat not in _fs
                            and abs(float(_pl.get(_cat, 0.0))) > 0.005):
                        _fs.setdefault(_cat, [(f'{_cat}（企业报表）', float(_pl.get(_cat, 0.0)))])
        for e, yd in entities.items():
            km = _adapter.read_km(e)
            l2map = {}
            _use_fee = e in _fee_scope
            _use_fin = e in _fee_fin
            for code, v in km.items():
                lv = v.get('level')
                nm = str(v.get('name') or '')
                if lv == 1:
                    s0 = nm.split('-')[0].strip()
                    if s0 and s0 not in name2code:
                        name2code[s0] = int(code) if code.isdigit() else 0
                    if code.isdigit():
                        code2name.setdefault(int(code), s0)
                    for y in years:
                        _k1 = (e, int(code) if code.isdigit() else 0, str(y))
                        tb_l1_control[_k1] = tb_l1_control.get(_k1, 0.0) + float(v.get('debit') or 0.0)
                elif len(code) >= 6:
                    c4 = code[:4]
                    # ⚡⚡ 2026-08-31 有企业利润表的主体：财务费用 6603 TB 行跳过
                    #   （借发口径双计 → 用企业报表净额覆盖，见下方 _fee_fin）
                    if _use_fin and c4 == '6603':
                        continue
                    # ⚡⚡ 2026-08-31 有费用目录的主体：6600 总池行跳过（功能范围拆分
                    #   到费用目录，TB 一级名聚合无法拆分管理/销售/研发/制造）
                    if _use_fee and c4 == '6600':
                        continue
                    # ⚡⚡ 2026-08-31 修复（AH 研发费用底稿 0 vs 试算表 119M）：AH SAP 无
                    #   独立 6604 研发费用，研发费用在 660006 研究开发费（c4='6600' 期间费用
                    #   总池码，不在 EXPENSE_CATS 值集）→ 原 c4 检查挡住 → tb_l2_control 无
                    #   研发费用 → 底稿全 0。放行 _RD_CODE(660006) 前缀，映射到标准码 6604。
                    if c4 in {str(v) for v in EXPENSE_CATS.values()} or code.startswith(_RD_CODE):
                        # ⚡ 2026-08-26 研发费用二级细分：660006 前缀的 8/10 位子目（直接投入/人员人工等）
                        #   c6 截断全是 '660006' → 二级合并成一行、Q-TB 填全部。改用完整科目码作二级 key。
                        c6 = code if code.startswith(_RD_CODE) else code[:6]
                        nm6 = '-'.join(nm.split('-')[:2]) if '-' in nm else nm
                        # ⚡ 2026-08-10 key 类型统一：U8 分支用 int 码（EXPENSE_CATS 值），
                        # SAP 分支原用 str(c4) → build_subject_rows.get(6603) 落空 → 明细表空壳
                        # ⚡⚡ 2026-08-31 660006 研究开发费 → 归研发费用标准码 6604（与
                        #   _write_sap_fee_audit 的 code=6604 匹配）
                        _ci = EXPENSE_CATS['研发费用'] if code.startswith(_RD_CODE) else (int(c4) if c4.isdigit() else 0)
                        tb_l2_names[_ci].setdefault(c6, nm6)
                        # ⚡ 2026-08-10 修复：二级借发控制数（明细表 TB 列依赖）——
                        # 原 SAP 分支只建 tb_l2_names、漏 tb_l2_control → 明细表(change_compare)
                        # TB 列全空 → 财务费用明细表空壳（1010 有审定 8536 万但明细表 0 行）。
                        # ⚡⚡ 2026-08-29 用户铁律：TB=审定表=明细表=附注，禁止"口径不一致"。
                        #   财务费用(6603)明细表控制数改回【TB 净额】(debit−credit)，与审定表
                        #   (TB jf−df) 同源一致；read_all_gl 不再覆盖（原 2026-08-28 用 GL 净额
                        #   导致 1010 审定 7688万 vs 明细 8261万 差异）。
                        #   注意：这里只遍历 len(code)>=6（末级），不含一级父行，天然无双计。
                        for y in years:
                            _k2 = (e, _ci, c6, str(y))
                            # ⚡⚡ 2026-08-31 研发费用（660006）取【借发 debit】（与 sap_tb_gen
                            #   利润表研发费用=借发口径一致）：原统一净额（debit-credit）在 660006
                            #   有贷方冲减（AH 1010 df=546K）时底稿 102.5M vs 试算表 103.0M 差 0.55M。
                            #   制造费用等保留净额（内部分摊冲减设计，其试算表同用净额/无冲减）。
                            _amt = (float(v.get('debit') or 0.0) if code.startswith(_RD_CODE)
                                    else float(v.get('debit') or 0.0) - float(v.get('credit') or 0.0))
                            tb_l2_control[_k2] = tb_l2_control.get(_k2, 0.0) + _amt
                        if len(code) >= 8:
                            tb_l3_names[_ci].setdefault(code, nm)
            # ⚡⚡ 2026-08-31 费用目录（功能范围拆分）→ tb_l2_control：科目名作二级
            #   key（无科目码），金额=本期累计。管理/销售/研发/制造费用底稿审定表
            #   (_write_sap_fee_audit 按 code 汇总) 与明细表(按二级名) 同源。
            if _use_fee:
                for cat, items in _fee_scope[e].items():
                    _ci = EXPENSE_CATS.get(cat)
                    if _ci is None:
                        continue
                    for nm, amt in items:
                        c6 = f'F-{nm}'
                        tb_l2_names[_ci].setdefault(c6, nm)
                        for y in years:
                            _k2 = (e, _ci, c6, str(y))
                            tb_l2_control[_k2] = tb_l2_control.get(_k2, 0.0) + amt
            # ⚡⚡ 2026-08-31 财务费用：企业利润表净额覆盖（6603 借发→净额口径）
            if _use_fin:
                _ci = EXPENSE_CATS.get('财务费用')
                if _ci is not None:
                    tb_l2_names[_ci].setdefault('F-财务费用', '财务费用（企业报表净额）')
                    for y in years:
                        tb_l2_control[(e, _ci, 'F-财务费用', str(y))] = _fee_fin[e]
        return name2code, code2name, tb_l2_names, tb_l2_control, tb_l1_control, tb_l3_names
    name2code = {}                 # 一级名称 -> 代码(int)
    code2name = {}                 # 代码(int) -> 一级名称
    tb_l2_names = defaultdict(dict)  # code -> {l2code(str): l2name(str)}  (6位=二级)
    tb_l3_names = defaultdict(dict)  # code -> {l3code(str): l3name(str)}  (8位=三级)
    tb_l2_control = {}             # (entity, code, l2code, year) -> 借发
    tb_l1_control = {}             # (entity, code, year) -> 一级借发(合计数控制数)
    # 预置 5 个类别名称→标准代码，使"仅 GL 有数据、TB 无对应一级"的类别也能在 GL 阶段正确归并。
    for _cat, _code in EXPENSE_CATS.items():
        name2code[_cat] = _code
        code2name[_code] = _cat
    for e, yd in entities.items():
        for y, paths in yd.items():
            km = paths.get('km')
            if not km or not os.path.exists(km):
                continue
            try:
                wb = load_workbook(km, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ 跳过 TB 读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            # —— 第一遍：建立 一级代码→名称 映射，并按【一级名称】判定每个一级的费用类别 ——
            #   （关键：类别只看名称，不看代码。J 账套 4101=制造费用、g 账套 4101=盈余公积，
            #    仅看代码会错把盈余公积当制造费用或漏掉 J 的制造费用。）
            l1_name = {}   # 一级4位码 -> 名称
            l1_cat = {}    # 一级4位码 -> 类别名(or None)
            all_names = {}   # cn -> 本级名（name.split('-') 末段，兼容"本级名/完整路径"两种导出格式）
            rows = []
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or len(r) < 6:
                    continue
                code = str(r[0]).strip() if r[0] is not None else ''
                name = str(r[1]).strip() if r[1] is not None else ''
                if not code:
                    continue
                cn = re.sub(r'\D', '', code)
                jf = safe(r[4]) if len(r) > 4 else 0.0
                rows.append((cn, name, jf))
                all_names[cn] = name.split('-')[-1].strip() if '-' in name else name
                if len(cn) == 4:
                    l1_name[cn] = name
                elif '\\' in name:
                    # 2026-08-06 修复：JTt U8『末级科目导出』无 4 位父级行（6602 只有 660201…），
                    # 且科目名=反斜杠路径（'660201\\管理费用\\工资'）→ 从路径段提取一级名
                    # （跳过数字段，取首个中文段=父级名），否则 l1_name 空 → 费用类别全判定失败。
                    _segs = [s for s in name.split('\\') if s and not s.strip().isdigit()]
                    if _segs and cn[:4] not in l1_name:
                        l1_name[cn[:4]] = _segs[0]
            # ⚡⚡ 2026-08-15 铁律127：TB 费用科目发生额缺失 → GL 兜底（ga 实证：16/40 主体
            #   TB 损益科目本期发生额全 0，费用控制数=0 → 审定表空壳）。GL 同科目名有借方
            #   发生额且 TB 该行 jf=0 → 用 GL 借方汇总填回（pl_gl_sum 按一级名过滤损益类）。
            _gpath = paths.get('gl')
            if _gpath and os.path.exists(_gpath) and rows:
                try:
                    from audit_common import pl_gl_sum as _pgs
                    _gs = _pgs(_gpath)
                    if _gs:
                        rows = [(c, nm, (_gs.get(nm, (0.0, 0.0))[0] or jf)
                                 if abs(jf) < 0.005 and nm in _gs else jf)
                                for c, nm, jf in rows]
                except Exception:
                    pass
            for _cn, _nm in l1_name.items():
                l1_cat[_cn] = _name_to_cat(_nm)
            # 构建完整路径名（按代码长度逐级拼接父级名）。关键：保证跨级/跨父级重名子目唯一——
            #   6602.09.02 效益奖金 → '管理费用-工资-效益奖金'、6602.33.02 效益奖金 → '管理费用-研发费用-效益奖金'；
            #   否则 GL 归并 key 重名 → 两个父行重复取数 → 分月小计双计（2026-08-02 修复）。
            full_path = {}
            for _cn in sorted(all_names, key=lambda x: (len(x), x)):
                _base = all_names[_cn]
                if len(_cn) == 4:
                    full_path[_cn] = _base
                else:
                    _par = _cn[:len(_cn) - 2]
                    _pf = full_path.get(_par)
                    # 仅当 _base 与父级【末级段】相同才视为重复丢弃；不可用子串包含判断——
                    # '电费' in '管理费用-项目支出-水电费' 为 True（水电费含电费子串），
                    # 会把 6602.43.12.02 电费 误并入父级名 → 10位/8位码重名 → GL 归并双计（2026-08-02）。
                    if _pf and _base != _pf.split('-')[-1]:
                        full_path[_cn] = _pf + '-' + _base
                    else:
                        full_path[_cn] = _pf or _base
            # —— 第二遍：一级按自身名称归类；二级按【所属一级的名称】归类（不看自身代码）——
            for cn, name, jf in rows:
                if len(cn) == 4:
                    cat = l1_cat.get(cn)
                    if not cat:
                        continue
                    ci = EXPENSE_CATS[cat]
                    if name:
                        name2code[name] = ci
                        code2name[ci] = name
                    tb_l1_control[(e, ci, y)] = jf          # 一级合计数控制数（如制造费用）
                elif len(cn) >= 6:
                    parent = cn[:4]                         # 结构取父级（仅定位父，不用代码判定类别）
                    pcat = l1_cat.get(parent)
                    # 2026-08-06：U8 末级导出（JTt/SSS）无 4 位父级行 → l1_cat 无 parent →
                    # 改按科目名【首段】判定一级（『660214\管理费用\固定资产』→ 管理费用）。
                    # 注：此处不再 continue 排除非费用父级（如 5001.04/4001.xx），由 _name_to_cat 天然过滤。
                    if not pcat:
                        pcat = _name_to_cat(name)
                    if not pcat:
                        continue
                    ci = EXPENSE_CATS[pcat]
                    # 2026-08-06：末级导出无父级行 → 一级控制数由子目累加（父级借发=子目之和）。
                    tb_l1_control[(e, ci, y)] = tb_l1_control.get((e, ci, y), 0.0) + jf
                    _fn = full_path.get(cn, name)
                    tb_l2_names[ci][cn] = _fn
                    tb_l2_control[(e, ci, cn, y)] = jf
                    if len(cn) >= 8:                      # 三级及以上明细（8位=三级/10位=四级/12位=五级）
                        tb_l3_names[ci][cn] = _fn
            wb.close()
    return name2code, code2name, tb_l2_names, tb_l2_control, tb_l1_control, tb_l3_names


# ---------------- 二级归属：GL 科目名 → TB 二级全名 ----------------
def _match_l2(nm, code, cands):
    """将 GL 科目名归并到 TB 二级全名（科目余额表）。
    用友综合查询明细表的 科目名 常为完整路径且父级重复，例如
      '管理费用-管理费用-职工薪酬-管理费用-职工薪酬-年终奖励'
    其中 TB 二级全名为 '管理费用-职工薪酬'。若能在 TB 二级名中按【词边界】子串命中，
    则取最长者（最具体）；否则退回末级段（研发类退回含'研发'段）。
    2026-08-02：词边界把『（』也视为分隔符——GL 名带『（三级明细）/（含三级明细）/（四级明细）』
    后缀时（如 '管理费用-研发费用-基本工资（三级明细）'），TB 三级名 '管理费用-研发费用-基本工资'
    命中点后的『（』应视为边界，否则三级永远匹配不上、GL 全被归并到二级父行（研发费用 3/4 级缺失）。"""
    if cands:
        # —— 第一级：尾匹配 —— cands 名必须是 GL 名【末尾】的词（GL 记到该科目本身或更细末级）。
        #   修复 2026-08-02 双计：GL 名 '管理费用-研发费用-效益奖金' 匹配 TB 8 位子目名，
        #   但旧逻辑按【最长词边界】会把 '管理费用-研发费用'（6位父名，8字符）误判为最长命中，
        #   导致子目行全部被父级抢走、L2 行 keys 再展开子目 → 小计双计。
        best = None
        for l2name in cands.values():
            if not l2name:
                continue
            if not nm.endswith(l2name):
                continue
            nxt = len(nm) - len(l2name)
            if nxt > 0 and nm[nxt - 1] not in '-（(：: ':
                continue
            if best is None or len(l2name) > len(best):
                best = l2name
        if best:
            return best
        # —— 第二级：任意词边界匹配（原逻辑）—— GL 名中间含 cands 名：
        #   父级记账行（GL 科目名=父级本身）、或 TB 无更细子目时归到最近父级。
        best = None
        for l2name in cands.values():
            if not l2name:
                continue
            idx = nm.find(l2name)
            while idx != -1:
                prev_ok = (idx == 0) or (nm[idx - 1] in '-（(：: ')
                nxt = idx + len(l2name)
                nxt_ok = (nxt == len(nm)) or (nm[nxt] in '-（(：: ')
                if prev_ok and nxt_ok:
                    if best is None or len(l2name) > len(best):
                        best = l2name
                    break
                idx = nm.find(l2name, idx + 1)
        if best:
            return best
    # —— 兜底：逐级去掉末段，找 TB 中的父级路径（GL 记了 TB 未列示的更细科目，
    #    如 '管理费用-研发费用-会议费' 且 TB 无会议费子目 → 归到 '管理费用-研发费用'）——
    seg = nm.split('-')
    while len(seg) > 1:
        seg = seg[:-1]
        parent = '-'.join(seg)
        if cands and parent in cands.values():
            return parent
    seg = nm.split('-')
    if '研发' in nm:
        idx = next((i for i, p in enumerate(seg) if '研发' in p), 0)
        return seg[idx]
    return seg[-1] if seg else nm


# ---------------- 读 GL：实体×一级×二级×月 借/贷 ----------------
def read_all_gl(data_dir, entities, years, name2code, tb_l2_names=None, tb_l2_control=None):
    # SAP：从 adapter 干净行聚合（类别判定 _sap_fee_cat 按功能范围列拆分；分月=month）
    # ⚡ 2026-08-10 用户提醒"看功能列"：SAP TB 编码 6600=期间费用总池（管理/销售/制造/研发
    # 混合），无 FR 列无法细分 → 费用类别一律从 GL 功能范围列(col19)拆分；
    # 同时补填 tb_l2_names/tb_l2_control（明细表二级名+控制数，SAP 无 TB 分 FR 数据）。
    if _adapter is not None and _adapter.is_sap(data_dir):
        gl = {y: defaultdict(lambda: [0.0] * 24) for y in years}
        for e, yd in entities.items():
            for r in _adapter.read_gl(e):
                cat = _sap_fee_cat(r)
                if not cat:
                    continue
                c = EXPENSE_CATS[cat]
                nm = str(r.get('name') or '')
                seg = nm.split('-')
                code = str(r.get('code') or '')
                # ⚡ 2026-08-26 研发费用二级细分（与 read_all_tb 对齐）：660006 前缀子目用完整码，
                #   否则 c6 截断全 '660006' → 研发二级合并、TB 控制数 Q 列错填全部借发。
                c6 = code if code.startswith(_RD_CODE) else (code[:6] if len(code) >= 6 else code)
                nm6 = '-'.join(seg[:2]) if len(seg) > 1 else seg[0]
                mo = int(r.get('month') or 0)
                # ⚡ 2026-08-10 lk 用名称（U8 语义一致）：GL 补充判定靠名称集合匹配 tb_names，
                # 用编码会被当"GL补充"显示（'660001（GL补充）'）；tb_l2_control 仍用 c6 编码 key
                lk = nm6 or c6
                y = str(r.get('year') or years[0] if years else '')
                arr = gl[y][(e, c, lk)]
                if 1 <= mo <= 12:
                    # ⚡ 2026-08-09 净额口径：SAP GL 借贷 gross 虚增（1010 财务费用借 17.6 亿 vs
                    # TB 借发 8536 万）——期间费用统一存净额(借-贷)（冲回/内部水电费分摊剔除）。
                    # ⚡⚡ 2026-08-28 制造费用(5101)从借发改净额：1010 FR=5101 贷方 1,645万
                    #   全为『生产部向各单位收取电费/水费』『结转生产部电费』内部分摊冲减，
                    #   用户权威《费用》目录即净额口径（制造费用 7,621万 = 借 9,266−贷 1,645），
                    #   read_all_gl 用借发与费用目录/审定表不一致（曾致 9,266万 vs 7,621万）。
                    #   四类费用统一净额后，GL 自足即可还原费用目录，无需读费用目录文件。
                    _db = float(r.get('debit') or 0.0)
                    _cr = float(r.get('credit') or 0.0)
                    arr[mo - 1] += _db - _cr
                    arr[12 + mo - 1] += _cr
                # ⚡ 2026-08-10 补填二级名（6600 系列 GL FR 拆分）。
                # ⚡⚡ 2026-08-29 用户铁律：TB=审定表=明细表=附注一致。原 2026-08-28 在此用
                #   GL 净额覆盖 tb_l2_control，导致明细表(TB列)与审定表(TB权威)不一致
                #   （1010 审定 7688万 vs 明细 8261万）。控制数改由 read_all_tb 填 TB 净额，
                #   此处【不再覆盖】tb_l2_control（GL 数据仍用于分月/附注的月度拆分）。
                if (code.startswith(_POOL_CODE) or code.startswith(_FIN_CODE)) and tb_l2_names is not None and c6 and nm6:
                    tb_l2_names[c].setdefault(c6, nm6)
        return gl
    gl = {y: defaultdict(lambda: [0.0] * 24) for y in years}  # 前12=借, 后12=贷
    for e, yd in entities.items():
        for y, paths in yd.items():
            g = paths.get('gl')
            if not g or not os.path.exists(g):
                continue
            try:
                wb = load_workbook(g, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ 跳过 GL 读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            # 文件级"整表重复"配对检测（2026-08-02）：FY深圳 2026 导出每行×2（336=168×2）→
            # 全部行键计数为偶数才判定整表复制、每键留 1 行；个别合法相同分录不误删。
            _file = []
            for r in ws.iter_rows(min_row=3, values_only=True):
                if not r or len(r) < 8:
                    continue
                nm = str(r[0]) if r[0] is not None else ''
                jf = safe(r[6])
                df = safe(r[7])
                if jf == 0 and df == 0:
                    continue
                _dk = (r[1], r[2], r[3], nm, jf, df,
                       str(r[4] or '') if len(r) > 4 else '',
                       str(r[5] or '') if len(r) > 5 else '')
                _file.append((_dk, r, nm, jf, df))
            _whole_dup = False
            if len(_file) >= 4:
                from collections import Counter as _C
                _cnt = _C(k for k, *_ in _file)
                _whole_dup = all(v >= 2 and v % 2 == 0 for v in _cnt.values())
            _seen_keys = set()
            for _dk, r, nm, jf, df in _file:
                if _whole_dup:
                    if _dk in _seen_keys:
                        continue
                    _seen_keys.add(_dk)
                seg = nm.split('-')
                s0 = seg[0]
                # 类别判定：优先按【名称含"研发"】归类（处理 小企业准则 研发费用计入管理费用子目），
                # 否则退回 首段名称→标准代码 映射。
                cat = _name_to_cat(nm)
                if cat:
                    c = EXPENSE_CATS[cat]
                else:
                    c = name2code.get(s0)
                    if c is None or c not in TARGET_CODES:
                        continue
                if c not in TARGET_CODES:
                    continue
                mo = parse_month(r[1])
                if mo is None or mo < 1 or mo > 12:
                    continue
                # 二级归属：优先按 TB 二级全名（科目余额表）词边界子串匹配，
                # 以对齐"GL查证取数"与"二级科目"；无匹配退回末级段（研发类退回含'研发'段）。
                l2key = _match_l2(nm, c, (tb_l2_names or {}).get(c))
                arr = gl[y][(e, c, l2key)]
                arr[mo - 1] += jf
                arr[12 + mo - 1] += df
            wb.close()
    return gl


# ---------------- 费用分部门明细表（参照『费用分部门明细表.xlsx』格式，2026-08-01 用户需求） ----------------
# 数据源：辅助核算余额表「科目辅助核算名称」（往来类型='部门'）。
# 部门名单动态识别（铁律：不要按参考表固定名单）：取 往来类型='部门' 且 辅助名≠'部门' 的末级部门行。
# 注：辅助核算「等级」=科目深度+2（如 管理费用-劳务费 2级→部门行4级；管理费用-工资-补贴-交通补贴 4级→部门行6级），
#     同一部门名可出现在多个等级（挂在不同明细科目下），故按「科目名称→部门」逐行归集、跨等级全部计入，
#     不按固定等级过滤（否则会漏掉挂在浅层科目下的部门金额）。
_FEE_CAT_PREFIXES = None  # 惰性初始化：一级科目名 → 前缀元组（含别名），见 _read_aux_dept


def _fee_cat_prefixes():
    global _FEE_CAT_PREFIXES
    if _FEE_CAT_PREFIXES is None:
        _FEE_CAT_PREFIXES = {}
        for _cat, _pre in (('制造费用', ('制造费用',)),
                           ('销售费用', ('销售费用', '营业费用')),
                           ('管理费用', ('管理费用',)),
                           ('财务费用', ('财务费用',)),
                           ('研发费用', ('研发费用', '研发费', '研发支出'))):
            _FEE_CAT_PREFIXES[_cat] = _pre
    return _FEE_CAT_PREFIXES


def _read_aux_dept(aux_path, fee_cat):
    """读辅助核算余额表：按 科目辅助核算名称(部门) 归集 某费用一级科目 的本期借方发生额。
    返回 {dept: 借方金额}。
    行结构（辅助核算余额表）：[科目辅助核算代码, 关联方, 科目名称, 科目辅助核算名称, 方向, 金额,
                                 本期借方, 本期贷方, 方向, 期末金额, 等级, 往来类型]
    · 仅取 往来类型='部门' 且 辅助名≠'部门' 的末级部门行（'部门'为汇总节点行，剔除）；
    · 费用发生额取 本期借方（r[6]）；一级科目按名称前缀匹配（别名：销售费用=营业费用）。"""
    prefixes = _fee_cat_prefixes()[fee_cat]
    try:
        wb = load_workbook(aux_path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
    except Exception as ex:
        print(f'  ⚠️ 辅助核算余额表读取失败 {aux_path}：{ex}')
        return {}
    res = {}
    for r in ws.iter_rows(min_row=3, values_only=True):
        if len(r) < 12 or r[2] is None:
            continue
        nm = str(r[2])
        if not nm.startswith(prefixes):
            continue
        aux = str(r[3]) if r[3] is not None else ''
        wtype = str(r[11]) if len(r) > 11 and r[11] is not None else ''
        if wtype != '部门' or not aux or aux == '部门':
            continue
        try:
            dr = float(r[6] or 0.0) if len(r) > 6 else 0.0
        except Exception:
            continue
        if abs(dr) < 0.005:
            continue
        res[aux] = res.get(aux, 0.0) + dr
    try:
        wb.close()
    except Exception:
        pass
    return res


def _aux_dept_all(data_dir, entities, years, fee_cat):
    """读全集团辅助核算：{e: {y: {dept: 借方金额}}}。无 aux 文件的实体返回 {}。"""
    out = {}
    for e, yd in entities.items():
        out[e] = {}
        for y, paths in yd.items():
            ap = paths.get('aux')
            if ap and os.path.exists(ap):
                out[e][y] = _read_aux_dept(ap, fee_cat)
    return out


def write_dept_detail_sheet(ws, name, code, year, entities, aux_by_ent, tb_control):
    """费用分部门明细表（参照『费用分部门明细表.xlsx』）：同一 sheet 纵向 3 块。
      块一 未审数：行=部门、列=部门名称+各主体借方发生额+合并抵消前总计/合并抵消/合并抵消后合计，
                   块底 合计/TB/差异 核对行；
      块二 审计调整：同结构，初始全 0；
      块三 审定数：同结构，审定 = 未审 + 调整（公式联动），块底 合计/TB/差异。
    tb_control: (entity, code, year) -> 一级借发控制数（TB 行来源）。"""
    ents = sorted(entities)
    # 全集团部门并集（本年）
    depts = set()
    for e in ents:
        d = aux_by_ent.get(e, {}).get(year, {})
        depts |= set(d.keys())
    depts = sorted(depts)
    if not depts:
        ws.cell(1, 1, f'（{name} 在 {year} 无「部门」级辅助核算数据，未生成分部门明细表）').font = \
            Font(name='Times New Roman', italic=True, color="888888")
        return
    n_ent = len(ents)
    ncols = 1 + n_ent + 3                    # 部门 + 各主体 + 合并抵消前总计/合并抵消/合并抵消后合计
    HDR_F = Font(name='Times New Roman', bold=True, size=10, color="000000")
    HDR_FL = PatternFill('solid', fgColor='DDEBF7')
    TOT_FL = PatternFill('solid', fgColor='FFF2CC')
    FN10 = Font(name='Times New Roman', size=10)
    BOLD10 = Font(name='Times New Roman', bold=True, size=10)

    def _tb_val(e):
        return tb_control.get((e, code, str(year)), 0.0) if tb_control else 0.0

    def _header_row(r):
        ws.cell(r, 1, '部门名称')
        for i, e in enumerate(ents, 2):
            ws.cell(r, i, e)
        ws.cell(r, ncols - 2, '合并抵消前总计')
        ws.cell(r, ncols - 1, '合并抵消')
        ws.cell(r, ncols, '合并抵消后合计')
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c); cell.font = HDR_F; cell.fill = HDR_FL
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = BORDER
        return r + 1

    def _body_rows(r, getter):
        """getter(e, dept) -> 数值。返回 (r_next, grand_by_ent)。"""
        grand = [0.0] * n_ent
        for dept in depts:
            ws.cell(r, 1, dept)
            row_sum = 0.0
            for i, e in enumerate(ents, 2):
                v = getter(e, dept) or 0.0
                row_sum += v; grand[i - 2] += v
                cell = ws.cell(r, i, round(v, 2))
                cell.border = BORDER; cell.font = FN10; cell.number_format = NUM_FMT
            c = ws.cell(r, ncols - 2, round(row_sum, 2))   # 合并抵消前总计
            c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
            ws.cell(r, ncols - 1, 0).number_format = NUM_FMT      # 合并抵消
            # 2026-08-06 协议：末列写数值（=抵消前+抵消0；公式 data_only 读 None）
            c2 = ws.cell(r, ncols, round(row_sum, 2))
            c2.border = BORDER; c2.font = FN10; c2.number_format = NUM_FMT
            cell0 = ws.cell(r, 1); cell0.border = BORDER; cell0.font = FN10
            r += 1
        # 合计行
        ws.cell(r, 1, '合计')
        for i, e in enumerate(ents, 2):
            cell = ws.cell(r, i, round(grand[i - 2], 2))
            cell.border = BORDER; cell.font = BOLD10; cell.fill = TOT_FL; cell.number_format = NUM_FMT
        gsum = sum(grand)
        c = ws.cell(r, ncols - 2, round(gsum, 2))
        c.border = BORDER; c.font = BOLD10; c.fill = TOT_FL; c.number_format = NUM_FMT
        ws.cell(r, ncols - 1, 0).number_format = NUM_FMT
        # 2026-08-06 协议：末列写数值
        c2 = ws.cell(r, ncols, round(gsum, 2))
        c2.border = BORDER; c2.font = BOLD10; c2.fill = TOT_FL; c2.number_format = NUM_FMT
        c0 = ws.cell(r, 1); c0.border = BORDER; c0.font = BOLD10; c0.fill = TOT_FL
        return r + 1, grand

    def _tb_diff_rows(r, grand):
        """TB 控制数行 + 差异行。TB 取该科目一级本期借发控制数（各主体）。"""
        tb_vals = [_tb_val(e) for e in ents]
        ws.cell(r, 1, 'TB')
        for i, e in enumerate(ents, 2):
            cell = ws.cell(r, i, round(tb_vals[i - 2], 2))
            cell.border = BORDER; cell.font = FN10; cell.number_format = NUM_FMT
        gtb = sum(tb_vals)
        c = ws.cell(r, ncols - 2, round(gtb, 2))
        c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
        ws.cell(r, ncols - 1, 0).number_format = NUM_FMT
        ws.cell(r, ncols, round(gtb, 2)).number_format = NUM_FMT
        r += 1
        ws.cell(r, 1, '差异')
        for i, e in enumerate(ents, 2):
            cell = ws.cell(r, i, round(tb_vals[i - 2] - grand[i - 2], 2))
            cell.border = BORDER; cell.font = FN10; cell.number_format = NUM_FMT
        c = ws.cell(r, ncols - 2, round(gtb - sum(grand), 2))
        c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
        ws.cell(r, ncols - 1, 0).number_format = NUM_FMT
        ws.cell(r, ncols, round(gtb - sum(grand), 2)).number_format = NUM_FMT
        return r + 1

    r = 1
    ws.cell(r, 1, f'{name}分部门明细表').font = Font(name='Times New Roman', bold=True, size=12)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 2
    # 块一：未审数
    ws.cell(r, 1, f'{name}分部门明细 — 未审数（{year} 年，各核算主体借方发生额）').font = BOLD10
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 1
    r = _header_row(r)
    row1_data = r
    r, grand1 = _body_rows(r, lambda e, d: aux_by_ent.get(e, {}).get(year, {}).get(d, 0.0))
    r = _tb_diff_rows(r, grand1)
    r += 1
    # 块二：审计调整
    ws.cell(r, 1, f'{name}分部门明细 — 审计调整数（{year} 年，待审计调整）').font = BOLD10
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 1
    r = _header_row(r)
    row2_data = r
    r, grand2 = _body_rows(r, lambda e, d: 0.0)
    r += 1
    # 块三：审定数 = 未审 + 调整（公式引用块一/块二）
    ws.cell(r, 1, f'{name}分部门明细 — 审定数（{year} 年，未审数 + 审计调整数）').font = BOLD10
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
    r += 1
    r = _header_row(r)
    row3_data = r
    n_l2 = len(depts)
    # 块一数据起始行 = 块三数据起始行 - 块二数据(n_l2) - 块二合计(1) - 空行(1)
    row1_data = row3_data - n_l2 - 2
    for i, dept in enumerate(depts):
        rr3 = row3_data + i
        ws.cell(rr3, 1, dept)
        for j in range(2, ncols - 2 + 1):
            # 2026-08-06 协议：审定行写数值（=块一未审值，调整 0；公式 data_only 读 None）
            _v = ws.cell(row1_data + i, j).value if isinstance(ws.cell(row1_data + i, j).value, (int, float)) else 0.0
            c = ws.cell(rr3, j, round(_v, 2))
            c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
        c = ws.cell(rr3, ncols - 2)
        col = _col_letter(ncols - 2)
        _v = ws.cell(row1_data + i, ncols - 2).value if isinstance(ws.cell(row1_data + i, ncols - 2).value, (int, float)) else 0.0
        c.value = round(_v, 2)
        c.border = BORDER; c.font = FN10; c.number_format = NUM_FMT
        ws.cell(rr3, ncols - 1, 0).number_format = NUM_FMT
        c2 = ws.cell(rr3, ncols)
        _v2 = ws.cell(row1_data + i, ncols).value if isinstance(ws.cell(row1_data + i, ncols).value, (int, float)) else 0.0
        c2.value = round(_v2, 2)
        c2.border = BORDER; c2.font = FN10; c2.number_format = NUM_FMT
    # 块三合计行（= 块一合计；2026-08-06 协议：写数值）
    rr3t = row3_data + n_l2
    ws.cell(rr3t, 1, '合计')
    for j in range(2, ncols + 1):
        _v = ws.cell(row1_data + n_l2, j).value if isinstance(ws.cell(row1_data + n_l2, j).value, (int, float)) else 0.0
        c = ws.cell(rr3t, j, round(_v, 2))
        c.border = BORDER; c.font = BOLD10; c.fill = TOT_FL; c.number_format = NUM_FMT
    c0 = ws.cell(rr3t, 1); c0.border = BORDER; c0.font = BOLD10; c0.fill = TOT_FL
    r = rr3t + 1
    # 块三 TB/差异（审定口径：TB 仍为一级借发控制数，审定数≈未审（调整 0））
    r = _tb_diff_rows(r, grand1)
    # 列宽
    ws.column_dimensions['A'].width = 24
    for j in range(2, ncols + 1):
        ws.column_dimensions[_col_letter(j)].width = 13
    ws.freeze_panes = 'B3'
    # 表尾口径说明（差异归因：未挂「部门」辅助核算的明细科目发生额不在部门行内）
    note = ('注：未审数合计 = 各主体辅助核算余额表「科目辅助核算名称（部门）」归集的借方发生额；'
            'TB 行 = 该费用科目一级本期借发控制数（科目余额表）。两者差异 = 未挂「部门」级辅助核算的'
            '明细科目发生额（如折旧摊销、税费类等），属取数口径差异而非账务差错；'
            '若需与 TB 完全一致，可将差额并入「公司」行或按未挂部门科目单独列示。')
    ws.cell(r + 1, 1, note).font = Font(name='Times New Roman', italic=True, size=9, color='888888')
    ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=ncols)


def _col_letter(i):
    """列号→字母（支持 >26，用 openpyxl.utils.get_column_letter）。"""
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)


# ---------------- 费用凭证级抽查（按核算主体，统一套用共享抽凭排除规则） ----------------
# 行级计提/摊销识别（2026-07-28 增补，修正「整笔凭证级排除漏网」铁律②）：
#   U8 常把「计提行(借费用 贷 应付职工薪酬/使用权资产累计折旧 等)」与「真实付款行(借费用 贷银行存款)」
#   塞进同一笔凭证。整笔级 voucher_skip_keys 因见到真实结算账户而整笔保留，导致计提行泄漏进抽凭。
#   故在此对【单行】再判定：只要该行是费用侧且对方科目为计提/摊销负债（或摘要/二级科目含计提/摊销/工资/社保/
#   公积金/使用权资产折旧 等），即判为相符计提/摊销，逐行剔除；真实付款(opp 含银行存款等结算户)不受影响。
_EXP_SETTLE = ('银行存款', '库存现金', '其他货币资金', '应收账款', '应付账款',
               '应收票据', '应付票据', '预收账款', '预付账款')
_EXP_ACCRUAL_LIAB = ('应付职工薪酬', '使用权资产', '累计折旧', '累计摊销', '专项储备',
                     '租赁负债', '预提', '暂估')
_EXP_ACCRUAL_KW = ('计提', '补提', '补计', '摊销', '分摊', '分配', '公积金', '社保',
                   '住房', '使用权资产折旧', '累计折旧', '累计摊销')
# 费用类一级前缀（用于「对方科目」展示补全时排除同凭证内其它费用行，避免噪声）
_EXPENSE_PREFIXES = ('销售费用', '管理费用', '财务费用', '制造费用', '研发费用')


def _expense_line_is_accrual(cp, sm, l2):
    """行级判定：该费用分录是否系相符计提/摊销（非真实交易），应剔除出抽凭。"""
    opp = cp or ''
    sm = sm or ''
    l2 = l2 or ''
    if any(k in opp for k in _EXP_SETTLE):
        return False  # 对方为真实结算账户 → 真实付款，保留
    if any(k in opp for k in _EXP_ACCRUAL_LIAB):
        return True   # 对方为计提/摊销负债 → 相符计提
    blob = sm + ' ' + l2
    if any(k in blob for k in _EXP_ACCRUAL_KW):
        return True   # 摘要/二级科目明示计提/摊销/工资/社保/公积金/使用权资产折旧 → 相符计提
    return False


def read_vouchers(data_dir, entities, years, name2code, tb_l2_names=None):
    """抽取费用类凭证级分录（含 凭证号/日期/摘要/对方科目/二级科目/借发生额）。
    仅取 费用一级科目(5101/6601/6602/6603/6604) 上的分录，返回记录列表。
    排除判定统一调用 audit_common.voucher_skip_keys（2026-07-25 确立的抽凭通用规则）：
      · 规则① 结转损益类凭证不抽；
      · 规则② 相符计提/摊销（借费用/成本 贷 应付职工薪酬/累计折旧/累计摊销/长期待摊费用等，
        无资金/往来结算）不抽；
      · 含银行/现金/应收应付等真实结算的凭证（如直接付款 借费用 贷银行存款）保留抽凭。
    规则②需看「整笔凭证」全部分录，故先按 (entity, year, 凭证字, 凭证号) 聚合 GL 全部分录，
    套用共享规则得到应排除的凭证集合，再对费用类分录逐行打 skip 标记。"""
    # SAP：从 adapter 干净行聚合费用科目分录（分类按功能范围列 _sap_fee_cat，
    # 无对方科目列，跳过计提/结转排除的简化处理）
    if _adapter is not None and _adapter.is_sap(data_dir):
        recs = []
        for e, yd in entities.items():
            for r in _adapter.read_gl(e):
                nm = str(r.get('name') or '')
                seg = nm.split('-')
                cat = _sap_fee_cat(r)
                if not cat:
                    continue
                c = EXPENSE_CATS[cat]
                if c not in TARGET_CODES:
                    continue
                jf = float(r.get('debit') or 0.0)
                df = float(r.get('credit') or 0.0)
                if jf == 0 and df == 0:
                    continue
                recs.append(dict(entity=e, year=str(r.get('year') or ''), code=c, cat=cat, nm=nm,
                                 # ⚡ 2026-08-10 抽凭 l2key：SAP 科目名（如『工资及奖金』）
                                 # 无 '-' 分段 → seg[1] 空 → 抽查凭证『二级科目』列空白；
                                 # 直接用完整名称（SAP 科目名本身就是末级/二级名）
                                 l2key=seg[1] if len(seg) > 1 else nm, vtype=r.get('vtype'),
                                 vno=r.get('vno'), date=r.get('date'), sm=r.get('sm'),
                                 cp=r.get('cp'), jf=jf, df=df, skip=False,
                                 # ⚡⚡ 2026-08-27 P0 修复：原 skip=False/accrual=False 硬编码，
                                 #   SAP 场景计提/摊销凭证（借费用 贷 应付职工薪酬/累计折旧等）
                                 #   泄漏进费用抽查（1010 管理费用抽到"计提工资"2000万）。
                                 #   SAP cp 列实为对方科目（计提时为应付职工薪酬等），
                                 #   走行级计提/摊销判定剔除；真实付款（银行存款等结算户）保留。
                                 accrual=_expense_line_is_accrual(
                                     str(r.get('cp') or ''), str(r.get('sm') or ''),
                                     seg[1] if len(seg) > 1 else nm)))
        return recs
    recs = []
    for e, yd in entities.items():
        for y, paths in yd.items():
            g = paths.get('gl')
            if not g or not os.path.exists(g):
                continue
            try:
                wb = load_workbook(g, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ 跳过 GL 读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            # 第一遍：收集该 GL 全部有效分录行，按 (凭证字, 凭证号) 聚合为整笔凭证
            raw = []
            for r in ws.iter_rows(min_row=3, values_only=True):
                if not r or len(r) < 9:
                    continue
                nm = str(r[0]) if r[0] is not None else ''
                jf = safe(r[6]); df = safe(r[7])
                if jf == 0 and df == 0:
                    continue
                vtype = str(r[2]) if r[2] is not None else ''
                vno = str(r[3]) if r[3] is not None else (str(r[2]) if r[2] is not None else '')
                date = r[1]
                sm = str(r[5]) if r[5] is not None else ''
                cp = str(r[4]) if r[4] is not None else ''
                raw.append(dict(vtype=vtype, vno=vno, date=date, sm=sm, cp=cp,
                                nm=nm, jf=jf, df=df))
            # 套用共享规则，得到应排除的 (凭证字, 凭证号) 集合
            skip_set = voucher_skip_keys(
                raw,
                key_fields=lambda rr: (rr['vtype'], rr['vno']),
                name_field=lambda rr: rr['nm'],
                opp_field=lambda rr: rr['cp'],
                summ_field=lambda rr: rr['sm'],
            )
            # 同一凭证全部分录（用于「对方科目」填空时取对方账户名）
            vlines = defaultdict(list)
            for rr in raw:
                vlines[(rr['vtype'], rr['vno'])].append(rr)
            # 第二遍：仅取费用类分录，打上整笔凭证 skip 标记 + 行级计提/摊销标记
            for rr in raw:
                nm = rr['nm']
                cat = _name_to_cat(nm)
                if cat:
                    code = EXPENSE_CATS[cat]
                else:
                    s0 = nm.split('-')[0].strip()
                    code = name2code.get(s0)
                    if code is None or code not in TARGET_CODES:
                        continue
                if code not in TARGET_CODES:
                    continue
                l2key = _match_l2(nm, code, (tb_l2_names or {}).get(code))
                orig_cp = rr['cp']
                # 展示用「对方科目」：用同号凭证全部分录按金额精确推导（derive_cp），
                # 仅取与本方方向相反、且金额匹配的真实对应账户，杜绝"一大堆"对方科目。
                # 惯常 2 分录凭证（借费用 贷银行存款）直接得唯一对方科目；多借多贷按金额匹配取唯一；
                # 仅极复杂且金额无匹配才退化为去重聚合（仍远少于原"同凭证全部非费用行"拼接）。
                v_lines = [{'name': o['nm'], 'dr': o['jf'], 'cr': o['df']}
                           for o in vlines[(rr['vtype'], rr['vno'])]]
                cp = derive_cp(v_lines, nm, rr['jf'], rr['df'])
                # 计提/摊销判定一律用【原始】对方科目 + 摘要/二级科目关键词，避免同凭证其它行污染
                accrual = _expense_line_is_accrual(orig_cp, rr['sm'], l2key)
                recs.append(dict(entity=e, year=y, code=code, cat=cat, nm=nm,
                                 l2key=l2key, vtype=rr['vtype'], vno=rr['vno'],
                                 date=rr['date'], sm=rr['sm'],
                                 cp=cp, jf=rr['jf'], df=rr['df'],
                                 skip=((rr['vtype'], rr['vno']) in skip_set),
                                 accrual=accrual))
            wb.close()
    return recs


# ---------------- 制造费用结转勾稽：5101 贷方(转出) ↔ 5001.04 借方(转入) ----------------
def read_transfer_gl(data_dir, entities, years):
    # SAP：从 adapter 干净行聚合 5101 贷/生产成本-制造费用 借 分月
    # ⚡ 2026-08-10 制造费用判定改功能范围列：SAP 制造费用科目名是明细名（"电费"等）无"制造费用"前缀
    if _adapter is not None and _adapter.is_sap(data_dir):
        transfer = {}
        for e, yd in entities.items():
            mfg_cr = [0.0] * 12
            cost_db = [0.0] * 12
            for r in _adapter.read_gl(e):
                nm = str(r.get('name') or '')
                mo = int(r.get('month') or 0)
                fr = str(r.get('fr') or '').strip()
                if 1 <= mo <= 12:
                    # ⚡ 2026-08-10 制造费用贷方限定 6600 系列费用科目：640202 其他业务成本等
                    # 非费用科目也可能带 FR=5101（SAP 实测），误计会虚增制造费用贷方
                    if fr == '5101' and str(r.get('code') or '').startswith(_POOL_CODE):
                        mfg_cr[mo - 1] += float(r.get('credit') or 0.0)
                    elif nm.startswith('生产成本') and '制造费用' in nm:
                        cost_db[mo - 1] += float(r.get('debit') or 0.0)
            transfer[(e, str(years[0]) if years else '')] = {'mfg_cr': mfg_cr, 'cost_db': cost_db}
        return transfer
    """从 GL 抽取制造费用结转的两边分月数据（用于 5101 贷方 ↔ 5001.04 借方 勾稽）。
    借：生产成本-制造费用(5001.04)  ←  贷：制造费用(5101)
      · 5001.04 转入(借)：GL 中 科目名以「生产成本」开头且含「制造费用」的分录借方；
      · 5101 转出(贷)：GL 中 科目名以「制造费用」开头（且非「生产成本-制造费用」）的分录贷方。
    返回 transfer[(entity, year)] = {'mfg_cr': [12月], 'cost_db': [12月]}。
    注：部分账套科目余额表无独立「贷发」列，故结转勾稽以 GL 凭证级分月为准，更可靠。

    2026-08-01 问题1b（泰国拆分结转）：泰国账套将制造费用按二级【拆分结转】到生产成本
    多个二级科目——员工支出→生产成本-工资、电水费→生产成本-水电、其余→生产成本-制造费用
    （5401 贷方 133.13M = 5400.02 工资 40.28M + 5400.03 水电 7.84M + 5400.04 制造费用 85.01M，
     分毫不差）。若仍只认「生产成本-制造费用」借方，泰国会误报 48M 结转差异。
    → 对泰国主体，成本承接方取 生产成本 二级名含【工资/水电/制造费用】的借方；
      「直接原料/半成品」等生产领料与在制品结转不属于制造费用承接，不纳入。"""
    try:
        from thai_mapping import is_thai_entity
    except Exception:
        is_thai_entity = lambda x: False
    _THAI_COST_L2_KW = ('工资', '水电', '制造费用')   # 泰国生产成本承接制造费用的二级
    transfer = {}
    for e, yd in entities.items():
        for y, paths in yd.items():
            g = paths.get('gl')
            if not g or not os.path.exists(g):
                continue
            try:
                wb = load_workbook(g, data_only=True, read_only=True)
            except Exception as ex:
                print(f'  ⚠️ 跳过结转GL读取失败 [{e}][{y}]：{ex}')
                continue
            ws = wb[wb.sheetnames[0]]
            rec = transfer.setdefault((e, y), {'mfg_cr': [0.0] * 12, 'cost_db': [0.0] * 12})
            thai = is_thai_entity(e)
            for r in ws.iter_rows(min_row=3, values_only=True):
                if not r or len(r) < 8:
                    continue
                nm = str(r[0]) if r[0] is not None else ''
                jf = safe(r[6]); df = safe(r[7])
                if jf == 0 and df == 0:
                    continue
                mo = parse_month(r[1])
                if mo is None or mo < 1 or mo > 12:
                    continue
                if nm.startswith('制造费用'):
                    # 制造费用转出（贷方）；排除「生产成本-制造费用」
                    if not nm.startswith('生产成本'):
                        rec['mfg_cr'][mo - 1] += df
                elif nm.startswith('生产成本'):
                    if thai:
                        # 泰国：生产成本承接制造费用的二级（工资/水电/制造费用）借方
                        l2nm = nm.split('-')[-1].strip() if '-' in nm else nm
                        if any(k in l2nm for k in _THAI_COST_L2_KW):
                            rec['cost_db'][mo - 1] += jf
                    elif '制造费用' in nm:
                        # 其他账套：生产成本-制造费用转入（借方）
                        rec['cost_db'][mo - 1] += jf
            wb.close()
    return transfer


def write_interest_voucher_sheet(ws, code, name, recs, entities, years):
    """财务费用『存款利息收入』凭证抽查（2026-08-07 用户要求：dq 财务费用存在大额存款
    利息收入 -11,329,260.54（660302 借方红字），原费用抽查只取借方正数（手续费等）漏了
    利息收入）。取 财务费用 借方红字（jf<0）且科目名含『利息』的凭证，按绝对值取前 10 笔；
    15 列标准格式（借方金额列=利息收入绝对值，摘要含红字说明）。"""
    hdrs = ['核算主体', '测试序号', '日期', '凭证字', '凭证号', '二级科目（或客户/供应商名称）',
            '摘要', '借方金额', '贷方金额', '对方科目', '与原始凭证相符', '原始凭证内容',
            '原始凭证日期', '会计处理正确', '所属时间无误']
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(1, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    all_rows = []
    for e in entities:
        for y in sorted(years):
            # ⚡ 2026-08-10 修复：dq 利息收入记【借方红字】(jf<0)，但 1010/SAP 记【贷方】(df>0)，
            # 原只筛 jf<0 → 1010 利息收入 803 万全部漏抽。扩展：借方红字 OR 贷方发生，
            # 排除月末结转凭证（摘要含 结转/本年利润）。
            pool = [x for x in recs if x['entity'] == e and x['year'] == y
                    and x['code'] == code
                    and (x['jf'] < -0.005 or (x['df'] > 0.005
                         and '结转' not in str(x.get('sm') or '')
                         and '本年利润' not in str(x.get('sm') or '')))
                    and '利息' in str(x.get('nm') or '')]
            pool.sort(key=lambda x: -abs(x['jf']))
            top = pool[:10]
            if not top:
                continue
            for x in top:
                all_rows.append((e, y, x))
    all_rows.sort(key=lambda t: (t[0], t[2]['date']))
    r = 2
    seq = 1
    for e, y, x in all_rows:
        ws.cell(r, 1, e); ws.cell(r, 2, seq); ws.cell(r, 3, x['date'])
        ws.cell(r, 4, ''); ws.cell(r, 5, x['vno']); ws.cell(r, 6, x['l2key'])
        ws.cell(r, 7, x['sm'] + '（红字冲减财务费用）' if x['jf'] < 0 else x['sm'])
        # ⚡ 2026-08-10：借方红字取 |jf|，贷方发生（SAP 利息收入）取 df
        _amt = abs(x['jf']) if x['jf'] < 0 else float(x['df'] or 0.0)
        amt = ws.cell(r, 8, round(_amt, 2)); amt.number_format = NUM_FMT
        ws.cell(r, 9, None); ws.cell(r, 10, x['cp'])
        for c in range(11, 16):
            ws.cell(r, c, None)
        left_cols = {6, 7, 10}
        for c in range(1, 16):
            cc = ws.cell(r, c); cc.border = BORDER
            cc.alignment = LEFT if c in left_cols else CENTER
        r += 1
        seq += 1
    if seq == 1:
        ws.cell(2, 1, '（本年度无存款利息收入红字凭证）')
    widths = {'A': 14, 'B': 10, 'C': 12, 'D': 10, 'E': 16, 'F': 24,
              'G': 56, 'H': 14, 'I': 14, 'J': 18, 'K': 14, 'L': 16,
              'M': 14, 'N': 14, 'O': 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'


def write_voucher_sheet(ws, code, name, recs, entities, years):
    """写「费用抽查凭证」表：按 核算主体×年度 取借方发生额前 10 笔（已剔除自动化/结转）。
    15 列标准格式，数据按(核算主体, 日期)排序后写入。"""
    hdrs = ['核算主体', '测试序号', '日期', '凭证字', '凭证号', '二级科目（或客户/供应商名称）',
            '摘要', '借方金额', '贷方金额', '对方科目', '与原始凭证相符', '原始凭证内容',
            '原始凭证日期', '会计处理正确', '所属时间无误']
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(1, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    # 收集所有数据行
    all_rows = []
    for e in entities:
        for y in sorted(years):
            pool = [x for x in recs if x['entity'] == e and x['year'] == y
                    and x['code'] == code and x['jf'] > 0
                    and not x.get('skip') and not x.get('accrual')]
            pool.sort(key=lambda x: -x['jf'])
            top = pool[:10]
            if not top:
                continue
            for x in top:
                all_rows.append((e, y, x))
    # 按(核算主体, 日期)排序
    all_rows.sort(key=lambda t: (t[0], t[2]['date']))
    # 统一写入
    r = 2
    seq = 1
    for e, y, x in all_rows:
        ws.cell(r, 1, e); ws.cell(r, 2, seq); ws.cell(r, 3, x['date'])
        ws.cell(r, 4, ''); ws.cell(r, 5, x['vno']); ws.cell(r, 6, x['l2key'])
        ws.cell(r, 7, x['sm'])
        amt = ws.cell(r, 8, round(x['jf'], 2)); amt.number_format = NUM_FMT
        ws.cell(r, 9, None); ws.cell(r, 10, x['cp'])
        for c in range(11, 16):
            ws.cell(r, c, None)
        left_cols = {6, 7, 10}
        for c in range(1, 16):
            cc = ws.cell(r, c); cc.border = BORDER
            cc.alignment = LEFT if c in left_cols else CENTER
        r += 1
        seq += 1
    # 抽凭排除规则说明已按用户要求删除（2026-07-31）
    # 列宽
    widths = {'A': 14, 'B': 10, 'C': 12, 'D': 10, 'E': 16, 'F': 24,
              'G': 52, 'H': 14, 'I': 14, 'J': 18, 'K': 14, 'L': 16,
              'M': 14, 'N': 14, 'O': 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'


# ---------------- 计算某一级科目的输出行清单（TB 优先，三级展开，缺口才补 GL）----------------
def _tol(x):
    return max(1.0, abs(x) * 1e-6)


def _row_gl_keys(row, code, tb_l3_names):
    """返回该输出行在 GL 中对应的所有二级 key（用于分月/对账取数）。"""
    _, _, gl_l2key, tb_l2code, level, children = row
    if level in ('L2G', 'L1G'):
        return [gl_l2key] if gl_l2key else []
    if level == 'L2H':                      # 三级展开时的二级标题行，不单独取数
        return []
    if level in ('L3', 'L4'):               # 三级/四级行：gl_l2key 即本层名
        return [gl_l2key] if gl_l2key else []
    # 二级行：本身名 + 其三级子目名（若 GL 把金额记到三级名下）+ 四级子目名
    # （2026-08-02：GL 四级名如『管理费用-研发费用-五险二金-养老保险』只归并在 4 级，
    #  若不加，L2 行 GL 缺 4 级金额 → 分月主体小计与 TB 差 5.24M（FY本级 研发费用））
    keys = [gl_l2key] if gl_l2key else []
    if children:
        l3d = tb_l3_names.get(code, {})
        for cc in children:
            nm = l3d.get(cc)
            if nm:
                keys.append(nm)
            for cc4, nm4 in l3d.items():
                if len(str(cc4)) == 10 and cc4.startswith(str(cc)):
                    keys.append(nm4)
    return keys


# ⚡ 2026-08-28 研发费用 GL 补充归并规则：GL 细类特征词 → TB 二级短名（去掉一级前缀）
# ⚡ 2026-08-28 研发费用归并：SAP 6600 池按功能范围拆分后，研发费用类别下出现大量碎细二级
#   （660001 工资及奖金 / 660002 工资附加费-医疗保险费 / 660036 保密补贴 / 660053 防暑降温费 /
#    660012 机票 / 660007 折旧费 / 660009 机物料消耗…），以及 GL 按功能范围拆出的研发细类。
#   按研发费用归集口径映射到标准二级（人员人工/直接投入/折旧摊销/其他费用…），使明细表干净。
_RD_SUPP_RULES = (
    (('工资附加', '工资及奖金', '工资', '医疗', '养老', '失业', '工伤', '生育', '住房公积金',
      '企业年金', '公积金', '年金', '保密补贴', '防暑降温', '降温', '高温', '取暖', '福利'), '人员人工'),
    (('折旧',), '折旧费用与长期费用摊销'),
    (('设计', '试验'), '设计费'),
    (('机物料', '物料消耗'), '直接投入'),
    (('无形资产', '摊销'), '无形资产摊销'),
    (('差旅', '机票', '交通', '住宿'), '其他费用'),
)

# 研发费用归并映射（build_subject_rows 构建，write_change_compare_sheet 写表时据此把并入金额
# 加到目标二级行）：{target_l2code: {'tb': [被并入的碎片二级码...], 'gl': [被并入的 GL 细类 key...]}}
_RD_MERGE_MAP = {}


def _merge_rd_fragments(gl, years, code, l2_list, gl_supp_keys):
    """研发费用细碎二级 + GL 补充归并：可归属的碎片/细类并入对应标准二级，
    返回 (剩余无法归并的 gl_supp_keys, 归并后的 l2_list)。金额由写表函数按 _RD_MERGE_MAP 并入。"""
    global _RD_MERGE_MAP
    # 反向索引：TB 二级短名（'-'后末段）→ l2code
    # ⚡⚡ 2026-08-28 全集团短名冲突修复：某主体『职工福利费-其他费用』(660003) 与标准研发
    #   『研究开发费-其他费用』(6600069900) 短名都是"其他费用"，setdefault 先取 660003 →
    #   _RD_MERGE_MAP 目标错乱、归并失效（2700 差旅费 5,821.43 丢失）。660006 标准研发二级优先。
    short2code = {}
    for cn, nm, _ in l2_list:
        s = str(nm).split('-')[-1] if '-' in str(nm) else str(nm)
        if s not in short2code or str(cn).startswith('660006'):
            short2code[s] = cn

    def _target_short(nm_or_lk, cn):
        """返回目标二级短名（人员人工/直接投入/…）或 None。"""
        if cn:
            if cn.startswith('660001') or cn.startswith('660002'):
                return '人员人工'
            if cn.startswith('660007'):
                return '折旧费用与长期费用摊销'
            if cn.startswith('660009'):
                return '直接投入'
            if cn.startswith('660012'):
                return '其他费用'
            if cn.startswith('660036') or cn.startswith('660053'):
                return '人员人工'
        for kws, short in _RD_SUPP_RULES:
            if any(k in (nm_or_lk or '') for k in kws):
                return short
        return None

    _RD_MERGE_MAP = {}
    merged_gl = set()
    l2_list = list(l2_list)
    # ① 碎片 TB 二级（非 660006 标准前缀）→ 从展示行移除，但【不记 md['tb']】：
    #   TB 碎片为 6 位合并行（660002 含医保/养老/教育等全部子目），与 GL 细类 lk 粒度不同，
    #   两边都并入会双计（实测『工资附加费-职工教育经费』TB 660002 与 GL lk 同金额被加 2 次）。
    #   金额统一由 ② GL 细类归并（md['gl']）承担，保证与 tb_l2_control 审定数口径一致。
    keep = []
    for cn, nm, flow in l2_list:
        if str(cn).startswith('660006'):
            keep.append((cn, nm, flow))
            continue
        tgt = _target_short(nm, str(cn))
        if tgt is None:
            keep.append((cn, nm, flow))
            continue
        tc = short2code.get(tgt)
        if tc is None:
            keep.append((cn, nm, flow))
            continue
        _RD_MERGE_MAP.setdefault(tc, {'tb': [], 'gl': []})
    # ② GL 细类 → 归并到目标二级（金额由写表 _tb / 分月 month_amt 从 gl[y] 并入）。
    #   跳过标准研发二级（'研究开发费-' 前缀，自身已在 tb_l2_control，归并即双计）。
    for lk in sorted(gl_supp_keys):
        if not lk or str(lk).startswith('研究开发费-'):
            continue
        tgt = _target_short(lk, None)
        if tgt is None:
            continue
        tc = short2code.get(tgt)
        if tc is None:
            continue
        _RD_MERGE_MAP.setdefault(tc, {'tb': [], 'gl': []})['gl'].append(lk)
        merged_gl.add(lk)
    return gl_supp_keys - merged_gl, keep


def build_subject_rows(code, years, gl, tb_l2_names, tb_l2_control, tb_l3_names, tb_l1_control, entities):
    """返回有序输出行清单：每行 (out_code, out_name, gl_l2key, tb_l2code, level, children)。
    构建原则（TB 优先，从一级出发）由 audit_common.build_account_breakdown_rows 统一编码：
      · 二级明细来自 科目余额表(TB) 二级借发；
      · 若 三级合计(TB) 与 二级(TB) 一致 → 展开三级（标注「三级明细」），二级作为分组标题；
      · 仅当 二级 TB 合计 ≠ 一级 TB 合计（存在缺口）时，才补 GL 中 TB 未列示的类别，并标注「GL补充」；
      · 0 值行的隐藏在写表时处理。
    本函数仅把费用特有的【按标准代码跨实体聚合 TB 控制数】映射成共享函数所需的入参
    （l2_list / l3_map / gl_supp_keys / tb_names），输出结构与旧版逐格一致。
    """
    l2_all = tb_l2_names.get(code, {})
    l3 = tb_l3_names.get(code, {})
    # 顶层仅列 二级(6位)；8位及以上一律作为子级挂载，避免深层科目(10/12位)被当独立二级行导致双计
    # （2026-08-01 修复 #15：Z 母公司 6604 研发费用 有 6→8→10→12 四级，旧逻辑把 10/12位 当二级，
    #   且 6位父级与 8位子级同列 → TB 合计虚增 3 倍）。
    # ⚡ 2026-08-26 研发费用(660006)二级 key 用完整 10 位码 → 此处放行 660006 前缀（否则研发二级全被滤掉）
    l2 = {cn: nm for cn, nm in l2_all.items()
          if len(str(cn)) <= 6 or str(cn).startswith(_RD_CODE)}
    # 孤儿深层(无6位父级的8位+)提升为顶层（父级缺失场景，如某些账套直接从8位开始）
    for cn, nm in l3.items():
        if cn[:6] not in l2 and len(str(cn)) > 6:
            l2[cn] = nm

    def tb_sum(cn):
        return sum(tb_l2_control.get((e, code, cn, y), 0.0) for e in entities for y in years)

    l1 = sum(tb_l1_control.get((e, code, y), 0.0) for e in entities for y in years)
    # 构造共享函数入参：二级列表 + 三级/四级映射（按 6 位父码归集）。
    # 2026-08-02：放开原「有更深层则三级不展开」的限制——L4(10位)单独归集为 l4_map 挂 L3 之下，
    # 12位(L5) 忽略不展开（金额含于其 L4 父行）；L2/L3 行 TB 值本就含全部子孙，L3/L4 仅下钻展示
    # 不再重复计入合计（FY 管理费用-研发费用 6602.33 有 L3/L4，如 五险二金-养老保险）。
    l2_list = [(cn, l2[cn], tb_sum(cn)) for cn in sorted(l2)]
    l3_map = {}
    l4_map = {}
    for cn in l2:
        kids8 = [(cc, l3[cc], tb_sum(cc)) for cc in l3 if len(str(cc)) == 8 and cc[:6] == cn]
        if not kids8:
            continue
        l4_map[cn] = {}
        for cc8, _, _ in kids8:
            # 4级=去点后 10 位码，直接父级=8 位三级码（无点码前缀匹配，无 '.'）
            kids10 = [(cc, l3[cc], tb_sum(cc)) for cc in l3
                      if len(str(cc)) == 10 and cc.startswith(cc8)]
            if kids10:
                l4_map[cn][cc8] = kids10
        l3_map[cn] = kids8
    tb_names = set(l2.values()) | set(l3.values())
    # GL 补充 key（含 None 表示一级直接发生额）；gl 为 None（如明细表无 GL 场景）则跳过
    gl_supp_keys = set()
    if gl:
        for y in years:
            for (e, c, lk) in gl[y]:
                if c == code:
                    gl_supp_keys.add(lk)
    # ⚡ 2026-08-28 研发费用(6604) 归并：SAP 6600 池按功能范围拆分产生的碎细二级/GL 细类
    #   （医疗保险费/保密补贴/防暑降温费/机票/工资及奖金…）按归集口径并入标准二级，
    #   不再独立成行或『GL补充』行（金额由写表函数按 _RD_MERGE_MAP 并入目标二级）。
    if gl_supp_keys and str(code) == str(EXPENSE_CATS['研发费用']):
        gl_supp_keys, l2_list = _merge_rd_fragments(gl, years, code, l2_list, gl_supp_keys)
    # 委托共享函数（输出 6 元组与旧版完全一致，validate_expense 回归校验）
    return build_account_breakdown_rows(l1, l2_list, l3_map, gl_supp_keys=gl_supp_keys,
                                        tb_names=tb_names, l4_map={cc: v for m in l4_map.values()
                                                                  for cc, v in m.items()})


def gl_year_sum(gl, year_entity_l2, years, which='debit'):
    """求和某 (entity, code, l2key) 跨年/跨月的借或贷。
    year_entity_l2: 指定 (e, code, l2key)；years: 参与的年列表；which: 'debit'/'credit'。
    """
    total = 0.0
    base = 0 if which == 'debit' else 12
    for y in years:
        arr = gl[y].get(year_entity_l2)
        if arr:
            total += sum(arr[base:base + 12])
    return total


# ---------------- 写「<费用>明细表」（TB 优先：年度取 TB，分月/对账取 GL）----------------
def write_summary_sheet(ws, code, name, rows_spec, gl, tb_l2_control, entities, years, tb_l2_names, tb_l3_names, tb_l1_control=None):
    # 重写表头（自包含，不依赖模板静态表头）
    # 说明：2025全年(GL)/2026(GL)/差异25(TB-GL)/差异26(TB-GL) 四列已移除——
    # TB 为权威控制数，GL 为非全量抽取（2026 仅 1–5 月 YTD），TB−GL 差异纯属取数范围缺口、
    # 非账务差错，列示易误导，故仅保留 TB 年度对比 + 2025 年 1–5 月(GL, YTD 可比) + 占全年%。
    # 表头随"是否有 2026 YTD 对比期"动态变化（2026-07-28 修正 #398）：
    #  - 仅 2025（无多一期数据）：只列 2025全年(TB)，不自动产生 1-5月 列；
    #  - 含 2026 YTD：列 2025全年 / 2026年1-5月(YTD)(TB) / 2025年1-5月(GL)
    #    / 2026年1-5月(YTD)(GL) / 1-5月可比变动率% / 2025年1-5月占全年%。
    has_2026 = '2026' in years
    if has_2026:
        HEAD = ['编码', '二级/三级科目名称', '2025全年(TB)', '2026年1-5月(YTD)(TB)',
                '2025年1-5月(GL)', '2026年1-5月(YTD)(GL)', '1-5月可比变动率%', '2025年1-5月占全年%']
    else:
        HEAD = ['编码', '二级/三级科目名称', '2025全年(TB)']
    for c, h in enumerate(HEAD, 1):
        cell = ws.cell(1, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    # 清除模板可能多出的表头列（如旧模板 12 列），避免残留"TB-GL 差异"等旧表头
    for c in range(len(HEAD) + 1, (ws.max_column or len(HEAD)) + 1):
        ws.cell(1, c).value = None
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None
    yrs = sorted(years)
    y2025 = '2025' if '2025' in years else yrs[0]
    y2026 = '2026' if '2026' in years else None

    def tb_amt(entity, tb_code, y):
        return tb_l2_control.get((entity, code, tb_code, y), 0.0) if tb_l2_code_ok(tb_code) else 0.0

    def gl_amt(entity, y, m, row):
        tot = 0.0
        for k in _row_gl_keys(row, code, tb_l3_names):
            arr = gl[y].get((entity, code, k))
            if arr:
                tot += sum(arr[0:12]) if m is None else arr[m - 1]
        return tot

    def _row_out(row):
        out_code, out_name, gl_l2key, tb_l2code, level, children = row
        c_val = (sum(tb_amt(e, tb_l2code, y2025) for e in entities)
                 if tb_l2_code_ok(tb_l2code) else sum(gl_amt(e, y2025, None, row) for e in entities))
        if has_2026:
            d_val = (sum(tb_amt(e, tb_l2code, y2026) for e in entities)
                     if tb_l2_code_ok(tb_l2code) else sum(gl_amt(e, y2026, None, row) for e in entities))
            g1 = sum(gl_amt(e, y2025, m, row) for e in entities for m in range(1, 6))
            g2 = sum(gl_amt(e, y2026, m, row) for e in entities for m in range(1, 6))
            h_comp = round((g2 - g1) / g1 * 100, 2) if abs(g1) > 0.005 else '—'
            pct_year = round(g1 / c_val * 100, 2) if abs(c_val) > 0.005 else '—'
            return out_code, out_name, level, c_val, d_val, g1, g2, h_comp, pct_year
        return out_code, out_name, level, c_val

    r = 2
    cT = dT = g1T = g2T = 0.0
    for row in rows_spec:
        if has_2026:
            out_code, out_name, level, c_val, d_val, g1, g2, h_comp, pct_year = _row_out(row)
        else:
            out_code, out_name, level, c_val = _row_out(row)
        # 0 值行不显示（仅看 TB 年度与 1-5月 GL 是否有数）
        _chk = [c_val] + ([d_val, g1, g2] if has_2026 else [])
        if all(abs(v) < 0.005 for v in _chk):
            continue
        ws.cell(r, 1, out_code); ws.cell(r, 2, out_name)
        ws.cell(r, 3, round(c_val, 2))
        if has_2026:
            ws.cell(r, 4, round(d_val, 2) if d_val != 0 else None)
            ws.cell(r, 5, round(g1, 2) if g1 != 0 else None)
            ws.cell(r, 6, round(g2, 2) if g2 != 0 else None)
            ws.cell(r, 7, h_comp); ws.cell(r, 8, pct_year)
        for c in range(3, len(HEAD) + 1):
            ws.cell(r, c).number_format = NUM_FMT
            ws.cell(r, c).alignment = RIGHT
        ws.cell(r, 2).alignment = LEFT
        if level == 'L3':
            ws.cell(r, 2).font = Font(name='Times New Roman', italic=True, color='595959')
        for c in range(1, len(HEAD) + 1):
            ws.cell(r, c).border = BORDER
        # 仅 6 位二级父级(其 TB 借发已含全部三级/四级子孙)与 GL 补充行计入合计；
        # 8 位三级、10 位四级仅作下钻展示，不再重复计入（修复 2/3/4 级嵌套双计）。
        if (len(str(out_code)) == 6) or (level in ('L2G', 'L1G')):
            cT += c_val
            if has_2026:
                dT += d_val; g1T += g1; g2T += g2
        r += 1
    # 全集团合计（仅可见数据行；⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出）
    if len(entities) > 1:
        ws.cell(r, 1, ''); ws.cell(r, 2, '(全集团合计)')
        ws.cell(r, 3, round(cT, 2))
        if has_2026:
            hT = round((g2T - g1T) / g1T * 100, 2) if abs(g1T) > 0.005 else '—'
            pT = round(g1T / cT * 100, 2) if abs(cT) > 0.005 else '—'
            ws.cell(r, 4, round(dT, 2)); ws.cell(r, 5, round(g1T, 2))
            ws.cell(r, 6, round(g2T, 2)); ws.cell(r, 7, hT); ws.cell(r, 8, pT)
        for c in range(3, len(HEAD) + 1):
            ws.cell(r, c).number_format = NUM_FMT
            ws.cell(r, c).alignment = RIGHT
        for c in range(1, len(HEAD) + 1):
            ws.cell(r, c).font = TOT_FONT
            ws.cell(r, c).fill = TOT_FILL
            ws.cell(r, c).border = BORDER
    # 注：原"一级(TB) ↔ 二级明细合计(TB)"勾稽提示块与"TB−GL 差异"列已按需求移除——
    # 二级明细合计本就直接取自 TB（与一级同一权威来源），无差异需提示；TB−GL 差异为取数范围缺口、非差错。
    return dict(c=cT, d=dT)


def tb_l2_code_ok(x):
    return x is not None and x != ''


def write_fx_carryforward_check(ws, data_dir, entities, year):
    """汇兑损益结转勾稽校验（2026-08-26 用户方法论）：
    外币评估借贷两边记账 → 每月 660303 汇兑损失 / 660304 汇兑收益 借贷净额(借-贷)
    应 = 月末结转损益凭证中对应金额。
    - 识别结转损益凭证：摘要含 结转损益/结转本年/本年利润，或同凭证含 4103/4104 本年利润科目
    - 找到结转凭证 → 逐月对比净额 vs 结转金额，超容差(5000)标红
    - 未找到 → 说明 SAP 未按月结转损益（实测 1010 序时账无 4103 行），提示以 TB 控制数核对（分月明细 R 列已暴露差异）"""
    from openpyxl.utils import get_column_letter as _gcl
    ws.cell(1, 1, f'汇兑损益结转勾稽校验（{year} 年度；外币评估借贷镜像 → 月净额应=月末结转损益金额）').font = SHELL_TITLE_FONT
    _hdr = ['月份', '汇兑损失净额', '汇兑收益净额', '结转凭证-损失', '结转凭证-收益', '差异', '状态']
    for j, h in enumerate(_hdr, 1):
        cc = ws.cell(3, j, h)
        cc.font = HDR_FONT; cc.fill = HDR_FILL; cc.alignment = CENTER
    net_d = defaultdict(float); net_y = defaultdict(float)
    by_vno = defaultdict(list)
    if _adapter is not None and _adapter.is_sap(data_dir):
        for e in entities:
            for r in _adapter.read_gl(e):
                c = str(r.get('code') or '')
                if not (c.startswith(_FX_LOSS_CODE) or c.startswith(_FX_GAIN_CODE)):
                    continue
                d = str(r.get('date') or '')
                m = int(d[5:7]) if len(d) >= 7 and d[5:7].isdigit() else 0
                net = float(r.get('debit') or 0.0) - float(r.get('credit') or 0.0)
                if c.startswith(_FX_LOSS_CODE):
                    net_d[m] += net
                else:
                    net_y[m] += net
                by_vno[str(r.get('vno') or '')].append(r)
    cog_d = defaultdict(float); cog_y = defaultdict(float)
    cog_found = False
    for vno, vrows in by_vno.items():
        is_cog = any('4103' in str(r.get('code') or '') or '4104' in str(r.get('code') or '')
                     for r in vrows) or any(
            k in str(r.get('sm') or '') for r in vrows for k in ('结转损益', '结转本年', '本年利润'))
        if not is_cog:
            continue
        for r in vrows:
            c = str(r.get('code') or '')
            if not (c.startswith(_FX_LOSS_CODE) or c.startswith(_FX_GAIN_CODE)):
                continue
            cog_found = True
            d = str(r.get('date') or '')
            m = int(d[5:7]) if len(d) >= 7 and d[5:7].isdigit() else 0
            net = float(r.get('debit') or 0.0) - float(r.get('credit') or 0.0)
            if c.startswith(_FX_LOSS_CODE):
                cog_d[m] += net
            else:
                cog_y[m] += net
    r = 4
    tot_nd = tot_ny = 0.0
    for m in range(1, 13):
        nd = net_d.get(m, 0.0); ny = net_y.get(m, 0.0)
        cd = cog_d.get(m, 0.0); cy = cog_y.get(m, 0.0)
        if abs(nd) < 0.005 and abs(ny) < 0.005 and abs(cd) < 0.005 and abs(cy) < 0.005:
            continue
        tot_nd += nd; tot_ny += ny
        diff = (nd - cd) + (ny - cy)
        if cog_found:
            status = '结转核对 ✓' if abs(diff) <= 5000 else '结转差异!'
        else:
            status = '无结转凭证（SAP未按月结转）'
        ws.cell(r, 1, f'{m}月')
        ws.cell(r, 2, round(nd, 2)).number_format = NUM_FMT
        ws.cell(r, 3, round(ny, 2)).number_format = NUM_FMT
        ws.cell(r, 4, round(cd, 2) if cog_found else None).number_format = NUM_FMT
        ws.cell(r, 5, round(cy, 2) if cog_found else None).number_format = NUM_FMT
        # 无结转凭证时差异列留空（避免把汇兑净额误当结转差异）
        cc = ws.cell(r, 6, round(diff, 2) if cog_found else None)
        if cog_found:
            cc.number_format = NUM_FMT
        ws.cell(r, 7, status)
        if '差异' in status:
            ws.cell(r, 7).font = Font(bold=True, color='C00000')
        for j in range(1, 8):
            ws.cell(r, j).border = BORDER
        r += 1
    # 汇兑净损益合计行（无结转凭证时仍可供审计参考）
    ws.cell(r, 1, '汇兑净损益合计').font = Font(bold=True)
    ws.cell(r, 2, round(tot_nd, 2)).number_format = NUM_FMT
    ws.cell(r, 3, round(tot_ny, 2)).number_format = NUM_FMT
    ws.cell(r, 6, round(tot_nd + tot_ny, 2)).number_format = NUM_FMT
    ws.cell(r, 7, '净汇兑损失(正)/收益(负)')
    for j in range(1, 8):
        ws.cell(r, j).border = BORDER
        ws.cell(r, j).fill = TOT_FILL
    r += 1
    r += 1
    if cog_found:
        ws.cell(r, 1, '说明：识别到结转损益凭证，汇兑净额与结转金额逐月核对（容差 5000）；差异=汇兑净额-结转金额。'
                     ).font = Font(name='Times New Roman', italic=True, color='595959')
    else:
        ws.cell(r, 1, '说明：GL 序时账未发现结转损益凭证（摘要含 结转损益/结转本年/本年利润 或同凭证 4103 本年利润）——'
                      'SAP 未按月结转损益（可能年终结转），结转核对列留空；汇兑净额与 TB 控制数的差异请核对分月明细 R 列，'
                      '若 SAP 后续配置月末结转损益，本表将自动启用逐月核对。'
                     ).font = Font(name='Times New Roman', italic=True, color='595959')
    for i, w in enumerate([8, 16, 16, 16, 16, 14, 28], 1):
        ws.column_dimensions[_gcl(i)].width = w
    return ws


# ---------------- 写「<年份>分月明细」 ----------------
def _load_fee_master(data_dir, ent, fee_cat):
    """读用户提供的权威费用明细表 `{data_dir}/费用/{ent}/{fee_cat}{月}.xlsx`（目录式，
    表头：项目/本月数/本期累计数/上年同期累计数）。返回 {月: {项目名: 本月数}}；
    无该数据源（目录不存在/无文件）返回空 dict → 分月明细回退 GL。
    ⚡ 2026-08-28：1010 制造费用 GL 5101 借发 1091万 vs 费用目录 762万（虚增 329万），
    分月明细必须以用户权威表为准。剔除『XX合计：』汇总行；项目名去全角缩进空格。
    ⚡ data_dir 实为输出目录（run_expense 传 _work()）时，回退 adapter 真实数据根。"""
    root = data_dir
    if not root or not os.path.isdir(os.path.join(root, '费用')):
        try:
            root = _adapter._DATA_ROOT
        except Exception:
            root = data_dir
    if not root or not ent:
        return {}
    base = os.path.join(root, '费用', str(ent))
    if not os.path.isdir(base):
        return {}
    out = {}
    for m in range(1, 13):
        fp = os.path.join(base, f'{fee_cat}{m}.xlsx')
        if not os.path.exists(fp):
            continue
        try:
            wb = load_workbook(fp, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            for row in ws.iter_rows(values_only=True):
                nm = str(row[0] or '').strip()
                if not nm or '合计' in nm:
                    continue
                try:
                    v = float(row[1] or 0)
                except (TypeError, ValueError):
                    continue
                nm2 = nm.replace('\u3000', '').strip()
                if not nm2:
                    continue
                _d = out.setdefault(m, {})
                _d[nm2] = _d.get(nm2, 0.0) + v
            wb.close()
        except Exception:
            continue
    return out


def write_monthly_sheet(ws, code, name, rows_spec, gl, tb_l2_control, entities, year, tb_l3_names, sap_fee_cat=False, data_dir=None):
    """写「<年份>分月明细」。
    ⚡ 2026-08-26 sap_fee_cat=True（SAP 期间费用 管理/销售/制造/研发）：TB 6600 无功能范围拆分，
    类别级 TB 控制数不存在（tb_l2_control 中该类别的 key 是 read_all_gl 补填的 GL fr 净额）。
    旧逻辑仍查 tb_l2_control → Q-TB 列读到 GL 自己 → Q=P、差异恒 0【假勾平】。
    修复：SAP 期间费用下 Q/R 列留空、小计 Q/R 留空、表底加说明；一级勾稽看「6600期间费用类别拆分核对」。
    ⚡ 2026-08-28 优先用用户权威《费用》目录（data_dir/费用/{主体}/{费用名}{月}.xlsx）：
    GL 5101 借发含结转/冲回虚增，费用目录才是实际归集数（分月合计应=审定表）。"""
    from openpyxl.utils import get_column_letter as _gcl
    # ⚡⚡ 2026-08-28 统一 GL 主口径：read_all_gl 四类费用已统一净额（含制造费用 5101），
    #   GL 自足可还原用户《费用》目录（1010 四类合计差 2.53 元），不再读取费用目录文件——
    #   否则审定表(GL) 与分月/明细(费用目录) 会在管理费用/研发费用出现 3.5万 口径差。
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None
    # ⚡ 2026-08-28 重写表头：O=合计(净发生额) P=借方发生额(备查) Q=TB控制数 R=差异（模板旧表头
    #   『合计(借发)/净发生额』与本版本列名不符，模板克隆后残留误导）
    _H = ['核算主体', '二级科目名称'] + [f'{m}月' for m in range(1, 13)] + \
         ['合计(净发生额)', '借方发生额(备查)', 'TB控制数', '差异']
    for _c, _t in enumerate(_H, 1):
        _cell = ws.cell(1, _c, _t)
        _cell.font = HDR_FONT; _cell.fill = HDR_FILL; _cell.alignment = CENTER; _cell.border = BORDER

    def month_amt(entity, m, row):
        """返回 (净额, 贷方发生额)：read_all_gl 前12=借-贷净额（四类统一净额口径），
        后12=贷方发生额（备查）。研发费用目标二级行并入 _RD_MERGE_MAP['gl'] 细类
        （碎片 660002 等已在展示行移除，金额由归并承担，保证分月合计=审定）。"""
        tot = 0.0
        tot_cr = 0.0
        for k in _row_gl_keys(row, code, tb_l3_names):
            arr = gl[year].get((entity, code, k))
            if arr:
                tot += arr[m - 1]
                tot_cr += arr[12 + m - 1]
        if str(code) == str(EXPENSE_CATS['研发费用']):
            md = _RD_MERGE_MAP.get(row[3])
            if md:
                for lk in md.get('gl', []):
                    arr = gl[year].get((entity, code, lk))
                    if arr:
                        tot += arr[m - 1]
                        tot_cr += arr[12 + m - 1]
        return tot, tot_cr

    def tb_ctl(entity, tb_l2code):
        # ⚡ 2026-08-26 假勾平修复：SAP 期间费用类别级无真实 TB 控制数，返回 None（Q/R 留空）
        if sap_fee_cat:
            return None
        return tb_l2_control.get((entity, code, tb_l2code, year), 0.0) if tb_l2_code_ok(tb_l2code) else 0.0

    r = 2
    for e in sorted(entities):
        ent_debit = [0.0] * 12
        ent_tb = 0.0
        l2_rows = []            # 该主体 L2/L2G/L1G 行号（小计公式仅对它们求和，L3/L4 下钻行不重复计）
        for row in rows_spec:
            out_code, out_name, gl_l2key, tb_l2code, level, children = row
            if level == 'L2H':                      # 三级标题行不列分月
                continue
            months = []
            months_cr = []
            for m in range(1, 13):
                net, cr = month_amt(e, m, row)
                months.append(net)
                months_cr.append(cr)
            # ⚡⚡ 2026-08-28 用户再反转（前次 08-27 曾要求借方发生额）：财务费用等损益科目
            #   在 GL 中借贷毛发生额含外币重估镜像/跨期冲回等相互抵销（1010 汇兑损失 1 月
            #   借发 1.15 亿 vs TB 全年仅 1191 万），分月明细必须按【每月净额（结转口径）】
            #   展示，汇兑收益/损失自然相抵，每月合计与 TB 控制数核对（差异=GL净额−TB）。
            months_net = list(months)
            db_total = round(sum(n + c for n, c in zip(months, months_cr)), 2)   # 借方发生额(备查)
            months = [round(n, 2) for n in months_net]              # 每月=净发生额(借-贷)
            p = sum(months)                                         # 合计(净发生额)
            p_net = p
            db_sum = p
            tb = tb_ctl(e, tb_l2code)
            s = (db_sum - tb) if tb is not None else None   # 差异 = 净发生额 - TB 控制数
            if abs(p) < 0.005 and abs(p_net) < 0.005:       # 0 值行不显示
                continue
            ws.cell(r, 1, e); ws.cell(r, 2, out_name)
            for m in range(12):
                ws.cell(r, 3 + m, round(months[m], 2) if months[m] != 0 else None)
                ws.cell(r, 3 + m).number_format = NUM_FMT
                ws.cell(r, 3 + m).alignment = RIGHT
            # ⚡ 2026-08-26 列对齐表头：O=合计(净发生额) P=借方发生额(备查) Q=TB控制数 R=差异
            ws.cell(r, 15, round(db_sum, 2)); ws.cell(r, 16, db_total)
            if tb is not None:
                ws.cell(r, 17, round(tb, 2)); ws.cell(r, 18, round(s, 2))
            for c in (15, 16, 17, 18):
                ws.cell(r, c).number_format = NUM_FMT; ws.cell(r, c).alignment = RIGHT
            # 差异超出容差标红（供审计师查因：GL 净额 vs TB 控制数不一致 = 序时账/取数问题）
            if tb is not None and abs(s) > 5000:
                ws.cell(r, 18).font = Font(name='Times New Roman', size=10, bold=True, color='C00000')
            for c in range(1, 19):
                ws.cell(r, c).border = BORDER
            ws.cell(r, 1).alignment = LEFT; ws.cell(r, 2).alignment = LEFT
            if level == 'L3':
                ws.cell(r, 3).font = Font(name='Times New Roman', italic=True, color='595959')
            if level == 'L4':
                ws.cell(r, 3).font = Font(name='Times New Roman', italic=True, color='A6A6A6')
            # 仅 6 位二级父级(其 TB 借发已含全部三级/四级子孙)与 GL 补充行计入主体小计；
            # L3/L4 仅作下钻展示，其值已含于父级 L2，不再重复计入（2026-08-02 修复分月小计双计）。
            if level in ('L2', 'L2G', 'L1G'):
                for m in range(12):
                    ent_debit[m] += months[m]
                ent_tb += (tb or 0.0)
                l2_rows.append(r)
            r += 1
        # 主体小计（2026-08-02 用户要求合计保留公式：分月/全年/TB =SUM(该主体 L2 行，逗号分隔，
        # 排除 L3/L4 下钻行避免双计)；差异=全年−TB 公式）
        ws.cell(r, 1, e + ' 小计'); ws.cell(r, 2, '(主体小计)')
        if l2_rows:
            def _cells(cl):
                return ','.join('%s%d' % (cl, rr) for rr in l2_rows)
            for m in range(12):
                cl = _gcl(3 + m)
                ws.cell(r, 3 + m, '=SUM(%s)' % _cells(cl))
                ws.cell(r, 3 + m).number_format = NUM_FMT; ws.cell(r, 3 + m).alignment = RIGHT
            ws.cell(r, 15, '=SUM(%s)' % _cells('O')); ws.cell(r, 16, '=SUM(%s)' % _cells('P'))
            # ⚡ 2026-08-26 假勾平修复：SAP 期间费用无 TB 类别控制数 → 小计 Q/R 留空（不写公式误导）
            if sap_fee_cat:
                ws.cell(r, 17, None); ws.cell(r, 18, None)
            else:
                ws.cell(r, 17, '=SUM(%s)' % _cells('Q'))
                # ⚡⚡ 2026-08-27 修复：原 '=P-Q'（净额-TB）使跨期差异显示 0，用户无法定位；
                #   差异=借发-TB=O-Q，暴露 GL 借发与 TB 借发的差额（跨期凭证/外币评估镜像）。
                ws.cell(r, 18, '=O%d-Q%d' % (r, r))
        else:
            for m in range(12):
                ws.cell(r, 3 + m, None).number_format = NUM_FMT
            ws.cell(r, 15, 0.0); ws.cell(r, 16, 0.0)
            ws.cell(r, 17, None if sap_fee_cat else 0.0); ws.cell(r, 18, None if sap_fee_cat else 0.0)
        for c in (15, 16, 17, 18):
            ws.cell(r, c).number_format = NUM_FMT; ws.cell(r, c).alignment = RIGHT
        for c in range(1, 19):
            ws.cell(r, c).font = SUB_FONT; ws.cell(r, c).fill = TOT_FILL; ws.cell(r, c).border = BORDER
        ws.cell(r, 1).alignment = LEFT; ws.cell(r, 2).alignment = LEFT
        r += 1
    # ⚡ 2026-08-26 假勾平修复：SAP 期间费用表底说明（类别为 GL 功能范围拆分，TB 无对应控制数）
    if sap_fee_cat:
        r += 1
        _c = ws.cell(r, 1, '说明：SAP 科目余额表 6600=期间费用总池，无功能范围拆分，类别级无 TB 控制数'
                          '（Q/R 列留空）。本表管理/销售/制造/研发类别为 GL 功能范围列拆分净额。'
                          '一级总池勾稽见「6600期间费用类别拆分核对」表。')
        _c.font = Font(name='Times New Roman', italic=True, color='595959')
    else:
        r += 1
        _c = ws.cell(r, 1, '口径说明（2026-08-28）：每月=GL 净发生额（借-贷，四类费用统一净额口径，'
                          '含制造费用内部分摊冲减；即损益结转口径，汇兑收益/损失自然相抵），'
                          '合计(净发生额)=全年净额，借方发生额(备查)=GL 借方毛发生额（含外币'
                          '评估镜像/冲回，仅备查不参与核对），TB控制数=科目余额表借发，差异=净发生额−TB 控制数。'
                          '差异≠0 的成因：①外币期末评估镜像（汇兑损益 GL 借贷大额对冲，TB 已净额化）；'
                          '②贷方冲减科目（利息收入/汇兑收益为贷方科目，TB 借发=0，其净额体现为负数抵减）；'
                          '③跨期凭证（GL 按凭证日期、SAP TB 按记账期间归集）。标红=差异超 5,000 元，'
                              '请按上述成因甄别。')
        _c.font = Font(name='Times New Roman', italic=True, color='595959')


# ---------------- 6600 期间费用类别拆分与 TB 总池核对（SAP 专属，2026-08-26） ----------------
def build_fee_cat_recon_sheet(wb, gl, tb_l1_control, entities, year, data_dir=None):
    """SAP 期间费用类别拆分与 TB 总池核对（用户单一口径要求：费用底稿最终要与 TB 勾平）。
    SAP TB 6600=期间费用总池（无功能范围拆分）；管理/销售/制造/研发类别只能从 GL 功能范围
    列(col19)拆分 → 类别级无 TB 控制数（分月明细 Q/R 列留空，见 write_monthly_sheet）。
    本表做【一级总池勾稽】——同口径【净额 vs 净额】：
      GL 4 类拆分净额合计 vs TB 6600 一级净额（借发-贷发，SAP 贷发=结转/冲回）。
    实测差异 = TB 导出时点缺失（6500004127 计提工资 2,000万 + 安全投入 649,432.50），
    1-6 月逐月 GL=TB 完全一致，证明 GL 完整。"""
    ws = wb.create_sheet('6600期间费用类别拆分核对')
    ws.cell(1, 1, f'6600 期间费用类别拆分与科目余额表总池核对（{year}；单位：元）').font = HDR_FONT
    _hdr = ['类别', 'GL 净额(借-贷)', '占期间费用%', 'TB 6600一级净额', '差异(GL净额-TB净额)']
    for c, h in enumerate(_hdr, 1):
        cell = ws.cell(2, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    # GL 各类别净额（read_all_gl 四类统一净额口径，arr 前12 即借-贷净额——制造费用 5101
    # 原存借方 gross 需再减贷方，2026-08-28 已统一净额，此处不再二次扣减）
    def _cat_net(ccode):
        v = 0.0
        for (e, cc, lk), arr in gl[str(year)].items():
            if cc == ccode:
                v += sum(arr[:12])
        return v
    cats = [('管理费用', 6602), ('销售费用', 6601), ('制造费用', 5101), ('研发费用', 6604)]
    r = 3
    gl_tot = 0.0
    for cname, ccode in cats:
        v = _cat_net(ccode)
        gl_tot += v
        ws.cell(r, 1, cname)
        ws.cell(r, 2, round(v, 2)).number_format = NUM_FMT
        for c in range(1, 6):
            ws.cell(r, c).border = BORDER
        r += 1
    # 期间费用 GL 小计
    ws.cell(r, 1, '期间费用小计（GL 4 类合计）').font = SUB_FONT
    ws.cell(r, 2, round(gl_tot, 2)).number_format = NUM_FMT
    for c in range(1, 6):
        ws.cell(r, c).font = SUB_FONT; ws.cell(r, c).fill = TOT_FILL; ws.cell(r, c).border = BORDER
    r += 1
    # TB 6600 一级净额（借发-贷发；优先 adapter 精确读，退 tb_l1_control 借发口径）
    tb_net = 0.0
    _adapter_tb_ok = False
    if data_dir and _adapter is not None and _adapter.is_sap(data_dir):
        try:
            for e in entities:
                _km = _adapter.read_km(e)
                for _c, _r in _km.items():
                    if str(_c) == '6600':
                        tb_net += float(_r.get('debit') or 0.0) - float(_r.get('credit') or 0.0)
            _adapter_tb_ok = True
        except Exception:
            _adapter_tb_ok = False
    if not _adapter_tb_ok:
        tb_net = sum(tb_l1_control.get((e, 6600, str(year)), 0.0) for e in entities)
    ws.cell(r, 1, 'TB 6600 期间费用总池（一级净额，借发-贷发）')
    ws.cell(r, 2, round(tb_net, 2)).number_format = NUM_FMT
    for c in range(1, 6):
        ws.cell(r, c).border = BORDER
    r += 1
    diff = gl_tot - tb_net
    ws.cell(r, 1, '差异（GL净额 - TB净额）').font = Font(bold=True)
    ws.cell(r, 2, round(diff, 2)).number_format = NUM_FMT
    ws.cell(r, 5, round(diff, 2)).number_format = NUM_FMT
    if abs(diff) > 5000:
        ws.cell(r, 5).font = Font(bold=True, color='C00000')
    for c in range(1, 6):
        ws.cell(r, c).border = BORDER
        ws.cell(r, c).fill = TOT_FILL
    r += 2
    ws.cell(r, 1, '说明：①SAP TB 无功能范围拆分，四类为 GL 功能范围列拆分；TB 6600 一级净额=全部期间费用。'
                  '②GL 与 TB 同为净额口径。'
                  '③差异主要为 TB 导出时点（约8月初）缺失 7 月底凭证：'
                  '如 6500004127 计提工资 2,000万、66001800 安全投入 649,432.50 等——GL 完整（1-6月逐月与 TB 一致），'
                  '需向财务核实 TB 补导。').font = Font(name='Times New Roman', italic=True, color='595959')
    ws.column_dimensions['A'].width = 34
    for col in ('B', 'C', 'D', 'E'):
        ws.column_dimensions[col].width = 18
    return ws


# ---------------- 写「<年份>主体比较」（动态实体列，TB 优先） ----------------
def write_compare_sheet(ws, code, name, rows_spec, gl, entities, year, tb_l2_control, tb_l3_names):
    from openpyxl.utils import get_column_letter as _gcl
    """写「<年份>主体比较」表：行=二级科目，列=主体。
    2026-08-01 问题1a：泰国账套二级科目名与集团标准不同（如 制造费用-员工支出 = 母公司-职工薪酬），
    按内容对齐比较——泰国二级名先经 THAI_L2_NAME_MAP 映射为集团标准名，再与同名主体行聚合为一行
    （泰国列与母公司列同现一行，便于横向比较；三级明细行挂在其映射后父级之下）。"""
    n = len(entities)
    ent_list = sorted(entities)
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体（SAP 逐主体底稿）不输出『全集团合计』列
    _single_cmp = n <= 1
    # 表头（去掉了二级科目编码）
    hdr = ['二级科目名称'] + list(ent_list) + ([] if _single_cmp else ['全集团合计'])
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(1, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    if _single_cmp:
        # ⚡ 2026-08-10 单体裁剪：模板表头残留的『全集团合计』列清空（模板 3 列固定）
        _clr = ws.cell(1, 2 + n)
        _clr.value = None; _clr.font = Font(name='Times New Roman', size=10)
        _clr.fill = PatternFill(); _clr.border = Border()

    def row_val(e, row):
        """该主体该行的年度数：有 TB 控制数取 TB，否则取 GL（缺口补充）。"""
        _, _, gl_l2key, tb_l2code, level, children = row
        if level == 'L2H':
            return 0.0
        if tb_l2_code_ok(tb_l2code):
            return tb_l2_control.get((e, code, tb_l2code, year), 0.0)
        tot = 0.0
        for k in _row_gl_keys(row, code, tb_l3_names):
            arr = gl[year].get((e, code, k))
            if arr:
                tot += sum(arr[0:12])
        return tot

    def _norm_key(out_name):
        """合并键：泰国二级名 → 集团标准基础名（剥「含三级明细/三级明细」后缀）。
        泰国「员工支出（含三级明细）」与母公司「职工薪酬」归并为同一键「职工薪酬」。"""
        if not out_name:
            return out_name
        base = re.sub(r'（含三级明细）|（三级明细）', '', str(out_name)).strip()
        return THAI_L2_NAME_MAP.get(base, base)

    def _is_l3(out_name):
        return '三级明细' in str(out_name or '')

    # —— 阶段1：收集并聚合（泰国映射名与同名主体行合并）——
    merged = {}        # 归一化键 -> [各主体值]
    has_l3 = {}        # 归一化键 -> 是否含三级明细
    order = []         # 归一化键保持首现顺序
    l3_buckets = {}    # 归一化键 -> [三级行(out_name, vals)]
    cur_l2 = None
    for row in rows_spec:
        out_code, out_name, gl_l2key, tb_l2code, level, children = row
        if level == 'L2H':
            continue
        # 2026-08-03 修复：数据列序必须与表头 ent_list(sorted) 一致——entities 原始顺序
        # （文件名发现序，FY 为 10,11,12,01..09）曾导致整列错位 3 位（FY01 3800万跑到 FY04 列）
        vals = [row_val(e, row) for e in ent_list]
        if abs(sum(vals)) < 0.005:
            continue
        if level in ('L3', 'L4'):
            # L3/L4 均为下钻展示（值含于父级 L2），挂当前父级 bucket，不计入主表
            if cur_l2 is not None:
                l3_buckets.setdefault(cur_l2, []).append((out_name, vals))
            continue
        key = _norm_key(out_name)
        cur_l2 = key
        if key not in merged:
            merged[key] = [0.0] * n
            has_l3[key] = _is_l3(out_name)
            order.append(key)
        else:
            has_l3[key] = has_l3[key] or _is_l3(out_name)
        for ci in range(n):
            merged[key][ci] += vals[ci]

    # —— 阶段2：统一输出（先父行，再其三级明细行）——
    r = 2
    col_tot = [0.0] * n
    col_parent_rows = [[] for _ in range(n)]   # 各主体列 父级行号（全集团合计公式用，排除三级下钻）

    def _write_row(label, vals, italic=False, count_total=True):
        nonlocal r
        ws.cell(r, 1, label)
        row_tot = 0.0
        for ci, e in enumerate(ent_list):
            v = vals[ci]
            ws.cell(r, 2 + ci, round(v, 2) if abs(v) >= 0.005 else None)
            ws.cell(r, 2 + ci).number_format = NUM_FMT; ws.cell(r, 2 + ci).alignment = RIGHT
            ws.cell(r, 2 + ci).border = BORDER
            if count_total:
                col_tot[ci] += v
                col_parent_rows[ci].append(r)
            row_tot += v
        if not _single_cmp:
            ws.cell(r, 2 + n, round(row_tot, 2)); ws.cell(r, 2 + n).number_format = NUM_FMT
            ws.cell(r, 2 + n).alignment = RIGHT; ws.cell(r, 2 + n).border = BORDER
        ws.cell(r, 1).alignment = LEFT
        if italic:
            ws.cell(r, 1).font = Font(name='Times New Roman', italic=True, color='595959')
        ws.cell(r, 1).border = BORDER
        r += 1

    for key in order:
        label = key + ('（含三级明细）' if has_l3.get(key) else '')
        _write_row(label, merged[key])
        for (l3name, l3vals) in l3_buckets.get(key, []):
            # 三级明细为下钻展示，其值已含于父级 TB，不计入全集团合计（防双计）
            _write_row(l3name, l3vals, italic=True, count_total=False)
    # 全集团合计行（2026-08-06 协议：写数值=col_tot 累计；公式 data_only 读 None 收回读不回）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出
    if len(entities) > 1:
        ws.cell(r, 1, '全集团合计')
        for ci in range(n):
            ws.cell(r, 2 + ci, round(col_tot[ci], 2) if abs(col_tot[ci]) >= 0.005 else 0.0)
            ws.cell(r, 2 + ci).number_format = NUM_FMT; ws.cell(r, 2 + ci).alignment = RIGHT; ws.cell(r, 2 + ci).border = BORDER
        _gtot = sum(col_tot)
        ws.cell(r, 2 + n, round(_gtot, 2) if abs(_gtot) >= 0.005 else 0.0)
        ws.cell(r, 2 + n).number_format = NUM_FMT; ws.cell(r, 2 + n).alignment = RIGHT; ws.cell(r, 2 + n).border = BORDER
        for c in range(1, 2 + n + 1):
            ws.cell(r, c).font = TOT_FONT; ws.cell(r, c).fill = TOT_FILL; ws.cell(r, c).border = BORDER
        ws.cell(r, 1).alignment = LEFT
    ws.freeze_panes = 'B2'
    # 列宽
    ws.column_dimensions['A'].width = 34
    # 2026-08-06 修复：chr(ord('B')+ci) 在主体数>24 时超 Z 变 '['（非法列名）——
    # ga 40 主体触发 openpyxl 保存崩溃，费用底稿全部失败。改用 get_column_letter 支持多字母列。
    for ci in range(n):
        ws.column_dimensions[_gcl(2 + ci)].width = 13
    if not _single_cmp:
        ws.column_dimensions[_gcl(2 + n)].width = 14




# ---------------- 制造费用专属勾稽：一级↔二级 + 结转 5001.04 ----------------
def _recon_ok(a, b):
    """差异是否可忽略（相对容差 1e-6，绝对下限 1 元）。"""
    diff = a - b
    return abs(diff) <= max(1.0, max(abs(a), abs(b)) * 1e-6)


def write_mfg_recon_sheet(ws, code, name, entities, years, tb_l1_control, tb_l2_names, tb_l2_control, gl, transfer):
    """制造费用(5101) 两张勾稽：
    A. 一级合计数(TB借发) ↔ 二级借发合计(TB)，按主体×年，差异应≈0；
    B. 制造费用结转：5101 贷方(GL月) ↔ 5001.04 借方(GL月)，按主体×年分月+全年，
       若存在差异则标注「不一致」并附差额（即结转未完全计入生产成本-制造费用）。"""
    yrs = sorted(years)
    r = 1
    ws.cell(r, 1, f'{code} {name} — 勾稽核对表（一级↔二级 / 结转生产成本-制造费用）').font = Font(name='Times New Roman', bold=True, size=10, color='000000')
    r += 2

    # ===== Section A：一级合计数 ↔ 二级借发合计 =====
    ws.cell(r, 1, '一、制造费用一级合计数 ↔ 二级明细合计（科目余额表，本期借发）').font = Font(name='Times New Roman', bold=True, size=10, color='000000')
    r += 1
    a_hdr = ['核算主体', '年度', '制造费用一级借发(TB)', '二级借发合计(TB)', '差异(一级-二级)', '二级借发合计(GL)', '差异(一级-GL)', '备注']
    for c, h in enumerate(a_hdr, 1):
        cell = ws.cell(r, c, h); cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    r += 1
    a_ok = True
    for e in entities:
        for y in yrs:
            l1 = tb_l1_control.get((e, code, y), 0.0)
            l2_codes = list(tb_l2_names.get(code, {}).keys())
            # 末级(叶子)码：不被其他码以自身为前缀包含，避免「父+子」重复计列
            # （如 5101.01 职工薪酬 与 5101.01.01 工资奖金 同时存在时只取叶子）。
            leaf_codes = [lc for lc in l2_codes
                          if not any(o != lc and o.startswith(lc) for o in l2_codes)]
            l2_tb = sum(tb_l2_control.get((e, code, lc, y), 0.0) for lc in leaf_codes)
            # GL 二级借发合计（跨所有二级 key，GL 仅末级）
            l2_gl = 0.0
            for (ee, cc, lk) in gl[y]:
                if ee == e and cc == code:
                    l2_gl += sum(gl[y][(ee, cc, lk)][0:12])
            diff_tb = l1 - l2_tb
            diff_gl = l1 - l2_gl
            note = '一致' if (_recon_ok(l1, l2_tb) and _recon_ok(l1, l2_gl)) else '差异需核实'
            if note != '一致':
                a_ok = False
            row_vals = [e, y, round(l1, 2), round(l2_tb, 2), round(diff_tb, 2), round(l2_gl, 2), round(diff_gl, 2), note]
            for c, v in enumerate(row_vals, 1):
                cell = ws.cell(r, c, v); cell.border = BORDER
                if c >= 3:
                    cell.number_format = NUM_FMT; cell.alignment = RIGHT
                if c == 8 and note != '一致':
                    cell.font = Font(name='Times New Roman', bold=True, color='C00000')
            r += 1
    r += 1

    # ===== Section B：制造费用结转 制造费用贷方 ↔ 生产成本-制造费用借方（均按名称取数，代码因账套而异）=====
    ws.cell(r, 1, '二、制造费用结转勾稽：制造费用（一级，代码因账套而异，如 5101/4101/5401）贷方（转出）'
                  '↔ 生产成本-制造费用（借方转入，代码如 5001.04/4001.xx；'
                  '泰国账套按二级拆分结转至生产成本-工资/水电/制造费用，一并纳入承接方）').font = Font(name='Times New Roman', bold=True, size=10, color='000000')
    r += 1
    b1_hdr = ['核算主体', '年度', '制造费用贷方(全年,GL)', '生产成本承接-制造费用借方(全年,GL)', '差额', '是否一致', '备注']
    for c, h in enumerate(b1_hdr, 1):
        cell = ws.cell(r, c, h); cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    r += 1
    b_ok = True
    try:
        from thai_mapping import is_thai_entity as _is_thai
    except Exception:
        _is_thai = lambda x: False
    for e in entities:
        for y in yrs:
            tr = transfer.get((e, y), {'mfg_cr': [0.0] * 12, 'cost_db': [0.0] * 12})
            mfg_cr = sum(tr['mfg_cr']); cost_db = sum(tr['cost_db'])
            diff = mfg_cr - cost_db
            ok = _recon_ok(mfg_cr, cost_db)
            if not ok:
                b_ok = False
            note = '一致（结转全部计入生产成本-制造费用）' if ok else '不一致：结转未完全对应，需核实'
            if _is_thai(e):
                note = ('一致（泰国拆分结转：员工支出→生产成本-工资、电水费→生产成本-水电、'
                        '其余→生产成本-制造费用）' if ok
                        else '不一致：泰国制造费用结转未完全对应生产成本承接二级，需核实')
            row_vals = [e, y, round(mfg_cr, 2), round(cost_db, 2), round(diff, 2), '是' if ok else '否', note]
            for c, v in enumerate(row_vals, 1):
                cell = ws.cell(r, c, v); cell.border = BORDER
                if c in (3, 4, 5):
                    cell.number_format = NUM_FMT; cell.alignment = RIGHT
                if c == 6 and not ok:
                    cell.font = Font(name='Times New Roman', bold=True, color='C00000')
                if c == 7 and not ok:
                    cell.font = Font(name='Times New Roman', color='C00000')
            r += 1
    r += 1

    # B2：分月明细
    ws.cell(r, 1, '   分月明细（5101贷方 vs 5001.04借方）').font = Font(name='Times New Roman', bold=True, size=10, color='000000')
    r += 1
    b2_hdr = ['核算主体', '年度', '月份', '制造费用贷方', '生产成本-制造费用借方', '差额']
    for c, h in enumerate(b2_hdr, 1):
        cell = ws.cell(r, c, h); cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    r += 1
    for e in entities:
        for y in yrs:
            tr = transfer.get((e, y), {'mfg_cr': [0.0] * 12, 'cost_db': [0.0] * 12})
            for m in range(1, 13):
                mc = tr['mfg_cr'][m - 1]; cd = tr['cost_db'][m - 1]; df = mc - cd
                row_vals = [e, y, f'{m}月', round(mc, 2), round(cd, 2), round(df, 2)]
                for c, v in enumerate(row_vals, 1):
                    cell = ws.cell(r, c, v); cell.border = BORDER
                    if c in (4, 5, 6):
                        cell.number_format = NUM_FMT; cell.alignment = RIGHT
                    if c == 6 and not _recon_ok(mc, cd):
                        cell.font = Font(name='Times New Roman', color='C00000')
                r += 1
    r += 1
    ws.cell(r, 1, f'勾稽结论：一级↔二级 {"一致" if a_ok else "存在差异（见上）"}；结转 制造费用↔生产成本-制造费用 '
                  f'{"一致" if b_ok else "存在差异（见上，差额即未计入生产成本-制造费用的结转额）"}。').font = Font(name='Times New Roman', italic=True, size=10, color='808080')
    # 列宽
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 8
    for col in ('C', 'D', 'E', 'F', 'G', 'H'):
        ws.column_dimensions[col].width = 18
    ws.column_dimensions['G'].width = 18


# ---------------- 同主体不同年度 明细费用增减对比表（明细表之后、分月之前） ----------------
def write_change_compare_sheet(ws, code, name, entities, years, tb_l2_names, tb_l2_control, tb_l3_names,
                               gl=None, tb_l1_control=None, data_dir=None):
    """写「{name}明细表」：行=核算主体×二级明细，列=上年/当年(TB)/增减额/变动率%。
    2026-08-02：二级下挂 三级/四级 下钻行（rows_spec 同 build_subject_rows，含 L3/L4）；
    L2 行计入主体小计，L3/L4 仅下钻展示不重复计（其值已含于父级 L2）。始终显示上年数列（无上年填0）。
    ⚡ 2026-08-11 阶段二 2.1：计算产结构化行，渲染收敛到 audit_render.render_sheet。
    ⚡⚡ 2026-08-28 统一 GL 主口径：当年列=GL 功能范围拆分净额（read_all_gl 四类统一净额），
    与审定表/分月明细/附注同源一致（不再用费用目录文件覆盖）。"""
    yrs = sorted(years)
    y_cur = yrs[-1]  # 最近一年为"当年"
    y_prev = yrs[-2] if len(yrs) >= 2 else None  # 有上年则用其数据，否则上年数=0
    hdr = ['核算主体', '二级科目名称',
           f'{y_prev}全年(TB)' if y_prev else '上年数',
           f'{y_cur}全年(TB)',
           '增减额', '变动率%']
    nh = len(hdr)
    MONEY = {3, 4, 5}
    from audit_render import render_sheet, row as _arow
    from openpyxl.styles import Font as _Fx

    # 行规范 = build_subject_rows（L2 + L3/L4 下钻 + GL 补充；gl 为 None 时无 GL 补充）
    rows_spec = build_subject_rows(code, yrs, gl, tb_l2_names, tb_l2_control, tb_l3_names or {},
                                   tb_l1_control or {}, entities)

    def _tb(e, tb_l2code, y):
        base = tb_l2_control.get((e, code, tb_l2code, y), 0.0) if (y and tb_l2_code_ok(tb_l2code)) else 0.0
        # ⚡ 2026-08-28 研发费用归并：目标二级金额 = TB 目标码 + 被并入碎片二级(TB) + GL 细类(借发)
        if y and str(code) == str(EXPENSE_CATS['研发费用']):
            md = _RD_MERGE_MAP.get(tb_l2code)
            if md:
                for src in md.get('tb', []):
                    base += tb_l2_control.get((e, code, src, y), 0.0) if tb_l2_code_ok(src) else 0.0
                for lk in md.get('gl', []):
                    arr = gl[y].get((e, code, lk)) if gl else None
                    if arr:
                        base += sum(arr[0:12])
        return base

    _rows = []
    g_cur = g_prev = 0.0
    ent_sub_rows = []   # 主体小计行号（全集团合计公式引用，2026-08-03 公式化）
    for e in sorted(entities):
        ent_cur = ent_prev = 0.0
        ent_l2_rows = []   # 本主体 L2 数据行号（主体小计公式引用，排除 L3/L4 下钻）
        started = False
        for row in rows_spec:
            out_code, out_name, gl_l2key, tb_l2code, level, children = row
            # ⚡⚡ 2026-08-28 统一 GL 主口径：当年列=GL 功能范围拆分净额（read_all_gl 净额，
            #   制造费用已含内部分摊冲减），不再用费用目录文件覆盖（避免与审定表口径差）。
            v_cur = _tb(e, tb_l2code, y_cur)
            v_prev = _tb(e, tb_l2code, y_prev)
            if abs(v_cur) < 0.005 and abs(v_prev) < 0.005:
                continue
            started = True
            if y_prev:
                chg = v_cur - v_prev
                rate = round(chg / v_prev * 100, 2) if abs(v_prev) > 0.005 else '—'
                vals = [e, out_name, round(v_prev, 2), round(v_cur, 2), round(chg, 2), rate]
            else:
                vals = [e, out_name, 0, round(v_cur, 2), round(v_cur, 2), '—']
            # 下钻行：斜体灰字，不计入主体小计（值已含于父级 L2）
            if level in ('L3', 'L4'):
                _f_ov = {
                    1: _Fx(name='Times New Roman', size=10, color='808080'),
                    2: _Fx(name='Times New Roman', italic=True,
                           color='595959' if level == 'L3' else 'A6A6A6'),
                }
                _rows.append(_arow(vals, num=MONEY, align='l', font=_f_ov))
            else:
                _rows.append(_arow(vals, num=MONEY, align='l'))
            if level in ('L2', 'L2G', 'L1G'):
                ent_cur += v_cur; ent_prev += v_prev
                ent_l2_rows.append(len(_rows))
        if started:
            # 主体小计（2026-08-06 协议：写数值=ent_prev/ent_cur 累计；公式 data_only 读 None）
            if y_prev:
                erate = round((ent_cur - ent_prev) / ent_prev * 100, 2) if abs(ent_prev) > 0.005 else '—'
                vals = [e + ' 小计', '(主体小计)', round(ent_prev, 2), round(ent_cur, 2),
                        round(ent_cur - ent_prev, 2), erate]
            else:
                vals = [e + ' 小计', '(主体小计)', 0, round(ent_cur, 2), round(ent_cur, 2), '—']
            _rows.append(_arow(vals, b=True, fill='sub', num=MONEY, align='l'))
            ent_sub_rows.append(len(_rows))
            g_cur += ent_cur; g_prev += ent_prev
    # 全集团合计（2026-08-06 协议：写数值=g_prev/g_cur 累计；公式 data_only 读 None）
    # ⚡ 2026-08-10 单体裁剪：仅 1 个主体 SAP 逐主体底稿不输出
    if (g_cur != 0 or g_prev != 0) and len(entities) > 1:
        if y_prev:
            gchg = g_cur - g_prev
            grate = round(gchg / g_prev * 100, 2) if abs(g_prev) > 0.005 else '—'
            vals = ['全集团合计', '', round(g_prev, 2), round(g_cur, 2), round(gchg, 2), grate]
        else:
            vals = ['全集团合计', '', 0, round(g_cur, 2), round(g_cur, 2), '—']
        _rows.append(_arow(vals, b=True, fill='total', num=MONEY, align='l'))

    # 清空模板残留的旧列（8列+）和旧数据行（原逻辑保留：模板克隆后需清残留）
    for c in range(nh + 1, (ws.max_column or nh) + 1):
        ws.cell(1, c).value = None
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).value = None
    _render_into(ws, hdr, _rows, MONEY)
    return ws


def _render_into(ws, headers, rows, money_cols, groups=None, freeze='A2'):
    """把结构化 rows 渲染进【已存在的 sheet】（不新建），用于模板克隆场景。
    ⚡ 2026-08-11：render_sheet 面向新建 sheet；模板驱动的 write_* 直接写已有 ws，
    复用渲染器的行渲染逻辑（值/金额格式/边框/对齐/行样式/字体覆盖）。
    groups: 双行表头分组（[(名, 跨列数),...]），提供时 R1=分组行、R2=列名、数据从 R3。"""
    from audit_render import SHELL_BORDER as _B, SHELL_NUM as _NF, SHELL_RGT as _RG, SHELL_LEFT as _LF, \
        SHELL_BOLD as _BL, SHELL_FONT as _FN, SHELL_TOT_FILL as _TF, FILL_TOTAL_MID as _FT, \
        SHELL_HFILL as _HF, SHELL_HFONT as _HFONT, SHELL_CEN as _CEN
    if groups:
        # 双行表头：R1 分组（深蓝白字合并），R2 列名，数据从 R3
        c0 = 1
        for gname, gn in groups:
            for _c in range(c0, c0 + gn):
                ws.cell(1, _c).border = _B
                ws.cell(1, _c).fill = _HF
            if gn <= 1:
                cell = ws.cell(1, c0, gname)
                cell.font = _HFONT
                cell.alignment = _CEN
            else:
                ws.merge_cells(start_row=1, start_column=c0, end_row=1, end_column=c0 + gn - 1)
                cell = ws.cell(1, c0, gname)
                cell.font = _HFONT
                cell.alignment = _CEN
            c0 += gn
        for j, h in enumerate(headers, 1):
            cell = ws.cell(2, j, h)
            cell.fill = _HF
            cell.font = _HFONT
            cell.border = _B
            cell.alignment = _CEN
        _r0 = 3
    else:
        # 单行表头（R1，与 render_sheet 单行表头同款：深蓝白字居中）
        for j, h in enumerate(headers, 1):
            cell = ws.cell(1, j, h)
            cell.fill = _HF
            cell.font = _HFONT
            cell.border = _B
            cell.alignment = _CEN
        _r0 = 2
    r = _r0
    for rs in rows:
        vals = rs.get('v') or []
        b = rs.get('b', False)
        fill = rs.get('fill')
        align_o = rs.get('align')
        num = rs.get('num')
        font_ov = rs.get('font') or {}
        for j, v in enumerate(vals, 1):
            cell = ws.cell(r, j)
            if v is None:
                pass
            elif isinstance(v, str) and v.startswith('='):
                cell.value = v
            elif isinstance(v, (int, float)):
                cell.value = round(float(v), 2)
            else:
                cell.value = v
            cell.border = _B
            if j in font_ov:
                cell.font = font_ov[j]
            elif fill == 'sub':
                cell.fill = _TF; cell.font = _BL
            elif fill == 'total':
                cell.fill = _FT; cell.font = _BL
            elif b:
                cell.font = _BL
            else:
                cell.font = _FN
            if num == 'all' or (num and j in num) or (num is None and j in money_cols):
                cell.number_format = _NF
            cell.alignment = _RG if (num == 'all' or (num and j in num) or (num is None and j in money_cols)) else _LF
        r += 1
    # 列宽 + 冻结
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 34
    for col in ('C', 'D', 'E', 'F'):
        ws.column_dimensions[col].width = 16
    ws.freeze_panes = freeze
    return ws


def _rebuild_detail_per_year(outs, code, name, entities, all_years,
                              tb_l2_names, tb_l2_control, tb_l3_names):
    """对 emit_per_year 拆分后的每份费用文件，按年重新生成明细表。
    2025 底稿：无上年数→仅 2025；2026 底稿：有 2025 数据→显示对比。"""
    import re as _re, openpyxl
    for fp in outs:
        base = os.path.basename(fp)
        ym = _re.search(r'_(\d{4})_', base)
        if not ym:
            continue
        y_target = ym.group(1)
        # 确定该文件适用的年份集：有上年数据时传两年做对比，否则只传当年
        has_prev = any(y < y_target for y in all_years)
        file_years = sorted(all_years) if has_prev else [y_target]
        try:
            wb = openpyxl.load_workbook(fp)
            # 删除旧明细表
            sn = f'{name}明细表'
            if sn in wb.sheetnames:
                del wb[sn]
            ws = wb.create_sheet(sn)
            write_change_compare_sheet(ws, code, name, entities, file_years,
                                       tb_l2_names, tb_l2_control, tb_l3_names or {})
            # 移到审定表之后（第2位）
            sheets = wb._sheets
            idx = next((i for i, s in enumerate(sheets) if s.title == sn), None)
            if idx is not None and idx > 1:
                sheets.insert(1, sheets.pop(idx))
            from audit_shell import finalize_workbook as _fw
            _fw(wb); wb.save(fp); wb.close()
        except Exception as ex:
            print(f'  ⚠️ 追加费用明细表失败 {base}：{ex}')


def _write_sap_fee_audit(wb, code, name, entities, y, tb_l2_control, data_dir=None):
    """SAP 费用审定表自建（2026-08-10 用户"看功能列"修复链）：
    SAP TB 编码 6600=期间费用总池（管理/销售/制造/研发混合），无管理费用/销售费用/
    制造费用/研发费用独立科目 → add_audit_summary_sheets 按 codes=[6602] 从 TB 精确取数必为 0。
    本期数 = GL 功能范围列拆分净额（tb_l2_control 按 (e, code, *, y) 汇总，四类统一净额，
    含制造费用内部分摊冲减；与明细表/分月/附注同源）。
    财务费用（6603）TB 有独立科目，仍走 add_audit（不动）。"""
    sn = f'{name} 审定表'
    if sn in wb.sheetnames:
        del wb[sn]
    ws = wb.create_sheet(sn)
    # 标题 + 表头（对齐 add_audit 输出格式）
    c0 = ws.cell(1, 1, f'{name} 审定表（按核算主体）')
    c0.font = SHELL_HFONT
    hdrs = ['核算主体', '上年数', '本期数', '审计调整数', '审定数', '与科目余额表勾稽']
    for c, h in enumerate(hdrs, 1):
        cell = ws.cell(2, c, h)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
    r = 3
    tot = 0.0
    for e in entities:
        v = sum(x for (ee, cc, c6, yy), x in tb_l2_control.items()
                if ee == e and cc == code and str(yy) == str(y))
        note = 'GL 功能范围拆分（SAP TB 无独立科目）'
        tot += v
        ws.cell(r, 1, e)
        for c in (2, 3, 5):
            cell = ws.cell(r, c, round(v, 2) if c in (3, 5) else 0.0)
            cell.number_format = NUM_FMT; cell.alignment = RIGHT; cell.border = BORDER
        ws.cell(r, 4, None).border = BORDER
        ws.cell(r, 6, note).border = BORDER
        r += 1
    ws.cell(r, 1, '合计')
    for c in (2, 3, 5):
        cell = ws.cell(r, c, round(tot, 2) if c in (3, 5) else 0.0)
        cell.number_format = NUM_FMT; cell.alignment = RIGHT; cell.border = BORDER
        cell.font = TOT_FONT; cell.fill = TOT_FILL
    ws.cell(r, 4, None).border = BORDER
    for c in range(1, 7):
        ws.cell(r, c).font = TOT_FONT; ws.cell(r, c).fill = TOT_FILL
        ws.cell(r, c).border = BORDER
    ws.column_dimensions['A'].width = 14
    for col, w in (('B', 12), ('C', 16), ('D', 12), ('E', 16), ('F', 34)):
        ws.column_dimensions[col].width = w
    return ws


def build_subject_workbook(data_dir, template_path, code, name, entities, years, gl, tb_l2_names, tb_l2_control, vouchers=None, tb_l1_control=None, transfer=None, tb_l3_names=None, aux_dept_all=None):
    """直接按年构建（2026-08-01 改造：消除合并稿中间态，替代 emit_per_year 分拆+按年重建）。
    每年独立 wb：审定表(audit_common 按年) + 明细表(change_compare 按年) + 分月明细/主体比较(当年) + 抽凭(当年)
    + 分部门明细表(辅助核算，2026-08-01) + 附注明细表(同利润表科目格式，2026-08-01)。"""
    tpl = template_path
    if not (tpl and os.path.exists(tpl)):
        tpl = _find_template(code, name)
    rows_spec = build_subject_rows(code, years, gl, tb_l2_names, tb_l2_control, tb_l3_names or {}, tb_l1_control or {}, entities)
    if not rows_spec:
        print(f'  ⚠️ [{code} {name}] 无二级科目数据，跳过。')
        return None
    from audit_shell import finalize_workbook as _fw
    from audit_common import add_audit_summary_sheets as _audit
    outs = []
    for y in sorted(years):
        ys = str(y)
        # ---- 加载模板（每年克隆）----
        if tpl and os.path.exists(tpl):
            try:
                wb = load_workbook(tpl)
            except Exception as ex:
                print(f'  ❌ 读取模板失败（可能被 Excel 占用）：{tpl} -> {ex}')
                return None
            if '说明' in wb.sheetnames:
                del wb['说明']
        else:
            print(f'  ⚠️ 模板缺失（{code}{name}.xlsx），改用内置默认空壳生成（格式降级）。请恢复 audit_templates 中的外壳。')
            wb = _make_default_expense_wb(name, [y])
        # 删除模板中非当年年份的专属 sheet（{other}分月明细/{other}主体比较）
        for sn in list(wb.sheetnames):
            m = re.match(r'^(20\d{2})(分月明细|主体比较)$', sn)
            if m and m.group(1) != ys:
                del wb[sn]
        # ---- 审定表（audit_common 按年；SAP 下管理费用/销售费用/制造费用/研发费用
        #      无 TB 独立科目 → 自建 _write_sap_fee_audit；财务费用 6603 有独立科目，
        #      但 AH 企业报表口径=净额（利息收入冲减）→ 有企业报表的账套也走
        #      _write_sap_fee_audit（tb_l2_control 已存企业报表净额））----
        _has_ent_rep = bool(data_dir) and RPT.has_enterprise_reports(data_dir)
        _sap_fee_audit = _adapter is not None and _adapter.is_sap(data_dir) and (
            code != 6603 or _has_ent_rep)
        if _sap_fee_audit:
            _write_sap_fee_audit(wb, code, name, entities, ys, tb_l2_control, data_dir=data_dir)
        else:
            _audit(wb, data_dir,
                   [dict(title=f'{name} 审定表', codes=[str(code)],
                         is_credit=False, subj_name=name,
                         income_statement=True, by_entity=True)],
                   target_year=ys)
        # ---- 费用明细表（change_compare：有上年数据则对比；与现状 rebuild 后一致）----
        has_prev = any(yy < y for yy in years)
        file_years = sorted(years) if has_prev else [y]
        sn_det = f'{name}明细表'
        if sn_det in wb.sheetnames:
            del wb[sn_det]
        ws_det = wb.create_sheet(sn_det)
        write_change_compare_sheet(ws_det, code, name, entities, file_years,
                                   tb_l2_names, tb_l2_control, tb_l3_names or {},
                                   gl=gl, tb_l1_control=tb_l1_control, data_dir=data_dir)
        # ---- 分月明细（当年）----
        sn_m = f'{ys}分月明细'
        if sn_m in wb.sheetnames:
            write_monthly_sheet(wb[sn_m], code, name, rows_spec, gl, tb_l2_control, entities, y, tb_l3_names or {},
                                sap_fee_cat=_sap_fee_audit, data_dir=data_dir)
        # ---- 6600 期间费用类别拆分核对（SAP 期间费用专属：分月明细 Q/R 留空，此处做一级总池勾稽）----
        if _sap_fee_audit:
            build_fee_cat_recon_sheet(wb, gl, tb_l1_control or {}, entities, y, data_dir=data_dir)
        # ---- 主体比较（当年）----
        sn_c = f'{ys}主体比较'
        if sn_c in wb.sheetnames:
            write_compare_sheet(wb[sn_c], code, name, rows_spec, gl, entities, y, tb_l2_control, tb_l3_names or {})
        # ---- 费用抽查凭证（当年）----
        vsn = '费用抽查凭证'
        if vsn in wb.sheetnames:
            del wb[vsn]
        ws_v = wb.create_sheet(vsn)
        write_voucher_sheet(ws_v, code, name, vouchers or [], entities, [y])
        # ---- 财务费用存款利息收入凭证抽查（2026-08-07 用户要求：财务费用大额存款
        # 利息收入（借方红字）原费用抽查只取借方正数漏掉；利息收入须单独抽凭）----
        if code == 6603:
            ivsn = '利息收入凭证抽查'
            if ivsn in wb.sheetnames:
                del wb[ivsn]
            ws_iv = wb.create_sheet(ivsn)
            write_interest_voucher_sheet(ws_iv, code, name, vouchers or [], entities, [y])
            # ---- 汇兑损益结转勾稽校验（2026-08-26 用户方法论：外币评估借贷镜像 →
            # 汇兑净额 应 = 月末结转损益金额）----
            fxsn = '汇兑损益结转勾稽校验'
            if fxsn in wb.sheetnames:
                del wb[fxsn]
            ws_fx = wb.create_sheet(fxsn)
            write_fx_carryforward_check(ws_fx, data_dir, entities, y)
        # ---- 分部门明细表（辅助核算，参照『费用分部门明细表.xlsx』；无部门数据则注明不生成）----
        if aux_dept_all:
            sn_d = f'{name}分部门明细表'
            if sn_d in wb.sheetnames:
                del wb[sn_d]
            ws_d = wb.create_sheet(sn_d)
            write_dept_detail_sheet(ws_d, name, code, y, entities, aux_dept_all, tb_l1_control)
        # ---- 附注明细表（同利润表科目格式：未审/调整/审定 三张平行表）----
        sn_n = '附注汇总'
        if sn_n in wb.sheetnames:
            del wb[sn_n]
        ws_n = wb.create_sheet(sn_n)
        try:
            from pl_detail import write_footnote_sheet as _wfn, nat_amt as _nat
            # 构造 gl_agg 同源数据：{(subj, e, l2, y): [dr, cr]}
            # ⚡⚡ 2026-08-29 用户铁律：TB=审定表=明细表=附注。附注改由【TB 净额】
            #   （tb_l2_control，与明细表/审定表同源）构造，不再用 GL（read_all_gl）
            #   ——原 1010 财务费用附注=GL 口径（汇兑损失 2571万含期初）与明细表
            #   TB 净额（1192万）不一致。
            gl_agg_note = {}
            for _e in entities:
                for _l2, _nm in (tb_l2_names.get(code) or {}).items():
                    _net = tb_l2_control.get((_e, code, _l2, str(y)), 0.0)
                    if abs(_net) > 0.005:
                        # ⚡⚡ _l2 用【二级名称】作 key（write_footnote_sheet 的 l2_keep
                        #   = rows_spec row[2] 即名称，原用 6 位码 '660301' 匹配失败 → 附注全 0）。
                        #   net 带符号放借方列（nat_amt('exp') 取 debit）：负数科目
                        #   （利息收入/汇兑收益）原放贷方列被忽略 → 显示 0 → 附注合计失真。
                        gl_agg_note[(name, _e, _nm, str(y))] = [_net, 0.0]
            # ⚡⚡ 2026-08-28 统一 GL 主口径：附注同源 GL 净额（read_all_gl 四类统一净额），
            #   不再用费用目录覆盖（避免审定/明细 GL 与附注费用目录口径差）。
            #   ⚡ 研发费用归并碎片（_RD_MERGE_MAP['gl']）金额并入目标二级，
            #   否则附注缺碎片金额（工资附加费-职工教育经费 3,168.32），与审定/明细/分月差。
            if str(code) == str(EXPENSE_CATS['研发费用']) and _RD_MERGE_MAP:
                _tc2glkey = {row[3]: row[2] for row in rows_spec if row[3]}
                for _tc, _md in _RD_MERGE_MAP.items():
                    _glkey = _tc2glkey.get(_tc)
                    if not _glkey:
                        continue
                    for _lk in _md.get('gl', []):
                        for _e in entities:
                            _arr = gl.get(str(y), {}).get((_e, code, _lk))
                            if _arr:
                                _base = gl_agg_note.get((name, _e, _glkey, str(y)), [0.0, 0.0])
                                gl_agg_note[(name, _e, _glkey, str(y))] = [
                                    _base[0] + sum(_arr[:12]), _base[1] + sum(_arr[12:])]
            # 2026-08-02 用户要求：附注汇总只要 2 级明细（3级/4级仅明细表下钻）——传 l2_keep 过滤。
            # 2026-08-03 用户方法论：附注只到二级（同名匹配，二级值含三级/四级合计）；
            # 四级父行（工资-补贴 等 L3 行）不单独进附注，否则与二级行值双计。
            l2_keep_note = {row[2] for row in rows_spec if row[4] in ('L2', 'L2G', 'L1G') and row[2]}
            _wfn(ws_n, name, 'exp', y, entities, gl_agg_note, l2_keep=l2_keep_note)
        except Exception as _ex:
            print(f'  ⚠️ 附注明细表生成失败 {name} {y}：{_ex}')
            ws_n.cell(1, 1, f'（附注明细表生成失败：{_ex}）').font = Font(name='Times New Roman', italic=True, color="888888")
        # ---- 制造费用(5101) 勾稽（当年）----
        if code == 5101:
            rsn = '制造费用勾稽核对'
            if rsn in wb.sheetnames:
                del wb[rsn]
            ws_r = wb.create_sheet(rsn)
            write_mfg_recon_sheet(ws_r, code, name, entities, [y],
                                  tb_l1_control or {}, tb_l2_names, tb_l2_control, gl, transfer or {})
        # ---- 分析性程序（2026-08-09 对齐 SAP 3 模块，用户要求 SAP/U8 统一）----
        ap_sn = '分析性程序'
        if ap_sn in wb.sheetnames:
            del wb[ap_sn]
        ws_ap = wb.create_sheet(ap_sn)
        try:
            ws_ap.cell(1, 1, f'{name} 分析性程序（{y} 年度）').font = SHELL_TITLE_FONT
            r_ap = 3
            ws_ap.cell(r_ap, 1, '① 分月合计与环比（环比波动>50% 标⚠️）').font = SHELL_SUB_FONT
            r_ap += 1
            for j, h in enumerate(['月份', '合计', '环比%', '提示'], 1):
                cc = ws_ap.cell(r_ap, j, h)
                cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
            r_ap += 1
            mon_tot = [0.0] * 12
            for e in entities:
                for k, arr in (gl.get(str(y)) or {}).items():
                    if k[1] != code:
                        continue
                    for m in range(12):
                        mon_tot[m] += arr[m]
            prev = None
            for m in range(12):
                chg = (mon_tot[m] - prev) / prev * 100 if prev else 0
                tip = '⚠️ 环比波动大' if abs(chg) > 50 and prev else ''
                ws_ap.cell(r_ap, 1, f'{m + 1}月').border = SHELL_BORDER
                cc2 = ws_ap.cell(r_ap, 2, round(mon_tot[m], 2) if mon_tot[m] else 0.0)
                cc2.number_format = SHELL_NUM; cc2.alignment = SHELL_RGT; cc2.border = SHELL_BORDER
                cc3 = ws_ap.cell(r_ap, 3, round(chg, 1) if prev else None)
                cc3.number_format = '0.0'; cc3.alignment = SHELL_RGT; cc3.border = SHELL_BORDER
                ws_ap.cell(r_ap, 4, tip).border = SHELL_BORDER
                prev = mon_tot[m]
                r_ap += 1
            ws_ap.cell(r_ap, 1, '全年合计').font = SHELL_BOLD
            cc_t = ws_ap.cell(r_ap, 2, round(sum(mon_tot), 2))
            cc_t.number_format = SHELL_NUM; cc_t.alignment = SHELL_RGT; cc_t.font = SHELL_BOLD
            for c in range(1, 5):
                ws_ap.cell(r_ap, c).fill = SHELL_TOT_FILL
                ws_ap.cell(r_ap, c).border = SHELL_BORDER
            r_ap += 2
            ws_ap.cell(r_ap, 1, '② 项目结构 Top10（占比）——关注集中与异常科目').font = SHELL_SUB_FONT
            r_ap += 1
            for j, h in enumerate(['项目', '金额', '占比%'], 1):
                cc = ws_ap.cell(r_ap, j, h)
                cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
            r_ap += 1
            proj_agg = {}
            for e in entities:
                for k, arr in (gl.get(str(y)) or {}).items():
                    if k[1] != code:
                        continue
                    proj_agg[k[2]] = proj_agg.get(k[2], 0.0) + sum(arr)
            tot_p = sum(proj_agg.values())
            for pname, pv in sorted(proj_agg.items(), key=lambda x: -x[1])[:10]:
                ws_ap.cell(r_ap, 1, pname).border = SHELL_BORDER
                ccp = ws_ap.cell(r_ap, 2, round(pv, 2))
                ccp.number_format = SHELL_NUM; ccp.alignment = SHELL_RGT; ccp.border = SHELL_BORDER
                ccp2 = ws_ap.cell(r_ap, 3, round(pv / tot_p * 100, 1) if tot_p else 0)
                ccp2.number_format = '0.0'; ccp2.alignment = SHELL_RGT; ccp2.border = SHELL_BORDER
                r_ap += 1
            if not proj_agg:
                ws_ap.cell(r_ap, 1, '（无发生额）').border = SHELL_BORDER
                r_ap += 1
            r_ap += 1
            ws_ap.cell(r_ap, 1, '③ 大额项目（年合计 ≥100万）——审计重点').font = SHELL_SUB_FONT
            r_ap += 1
            for j, h in enumerate(['项目', '金额'], 1):
                cc = ws_ap.cell(r_ap, j, h)
                cc.font = SHELL_HFONT; cc.fill = SHELL_HFILL; cc.border = SHELL_BORDER; cc.alignment = SHELL_CEN
            r_ap += 1
            n_big = 0
            for pname, pv in sorted(proj_agg.items(), key=lambda x: -x[1]):
                if pv >= 1000000:
                    ws_ap.cell(r_ap, 1, pname).border = SHELL_BORDER
                    ccb = ws_ap.cell(r_ap, 2, round(pv, 2))
                    ccb.number_format = SHELL_NUM; ccb.alignment = SHELL_RGT; ccb.border = SHELL_BORDER
                    r_ap += 1
                    n_big += 1
            if n_big == 0:
                ws_ap.cell(r_ap, 1, '无年合计 ≥100万 项目').border = SHELL_BORDER
                r_ap += 1
            ws_ap.column_dimensions['A'].width = 36
            ws_ap.column_dimensions['B'].width = 16
            ws_ap.column_dimensions['C'].width = 12
            ws_ap.column_dimensions['D'].width = 16
            ws_ap.freeze_panes = 'A3'
        except Exception as _ex:
            print(f'  ⚠️ 分析性程序生成失败 {name} {y}：{_ex}')
        # ---- sheet 顺序：审定表第1、明细表第2 ----
        sheets = wb._sheets
        t_idx = next((i for i, s in enumerate(sheets) if '审定表' in s.title), None)
        if t_idx is not None and t_idx > 0:
            sheets.insert(0, sheets.pop(t_idx))
        d_idx = next((i for i, s in enumerate(sheets) if s.title == sn_det), None)
        if d_idx is not None and d_idx > 1:
            sheets.insert(1, sheets.pop(d_idx))
        _fw(wb)
        # ⚡ 2026-08-28 #875：对方科目核对（模块内集成；expense 的 gl 为聚合结构 → auto 读全量行）
        try:
            from counterparty_recon import inject_into_wb_auto
            inject_into_wb_auto(wb, data_dir, ys, name, ents_set=set(entities))
        except Exception as _ex:
            print(f'  ⚠️ {name}对方科目核对注入失败：{_ex}')
        out = os.path.join(data_dir, f'{name}审计底稿_{ys}_生成.xlsx')
        from audit_common import validate_workbook
        validate_workbook(wb, '费用底稿', raise_on_error=False)
        wb.save(out)
        wb.close()
        outs.append(out)
        print(f'  ✓ {name} {ys} 已生成')
    return outs[0] if outs else None



# ---------------- 合并入口 ----------------
def build_expense_combined(data_dir, out_dir=None):
    entities = _discover_entities(data_dir)
    if not entities:
        print(f'❌ 目录内未找到含「科目余额表」的账套文件：{data_dir}')
        return []
    out_dir = out_dir or data_dir
    os.makedirs(out_dir, exist_ok=True)
    years = sorted({y for b in entities.values() for y in b})
    print(f'发现核算主体 {len(entities)} 个，年度：{", ".join(years)}')
    # 读 TB + GL
    name2code, code2name, tb_l2_names, tb_l2_control, tb_l1_control, tb_l3_names = read_all_tb(data_dir, entities, years)
    if not name2code:
        print('❌ 未能从科目余额表解析到任何一级科目名称，请检查文件格式。')
        return []
    gl = read_all_gl(data_dir, entities, years, name2code, tb_l2_names, tb_l2_control)
    vouchers = read_vouchers(data_dir, entities, years, name2code, tb_l2_names)
    transfer = read_transfer_gl(data_dir, entities, years)   # 制造费用结转：5101贷 ↔ 5001.04借

    produced = []
    ent_list = list(entities.keys())
    for code, name in SUBJECTS:
        # 该类别在 TB 有二级 或 在 GL 中出现任一实体/年，才算"有数据"。
        has_tb = bool(tb_l2_names.get(code))
        has_gl = any(c == code for y in years for (e, c, lk) in gl[y])
        if not has_tb and not has_gl:
            print(f'  ⊘ [{code} {name}] 本文件夹无相关数据，跳过。')
            continue
        # 分部门明细表数据源：全集团辅助核算余额表（仅读取有 aux 文件的主体；无则分部门表注明不生成）
        aux_dept_all = None
        try:
            aux_dept_all = _aux_dept_all(data_dir, entities, years, name)
        except Exception as ex:
            print(f'  ⚠️ 辅助核算部门归集失败 {name}：{ex}')
        tpl = _find_template(code, name)
        out = build_subject_workbook(data_dir, tpl, code, name, ent_list, years, gl, tb_l2_names, tb_l2_control, vouchers, tb_l1_control, transfer, tb_l3_names, aux_dept_all)
        if out:
            produced.append((code, name, out))
            print(f'  ✅ 已生成：{out}')
    return produced


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


def main(argv=None):
    raw = list(sys.argv[1:] if argv is None else argv)
    if raw and not raw[0].startswith('-') and os.path.isdir(raw[0]):
        raw = [raw[0]]
    else:
        if not raw:
            try:
                inp = input('请拖入或输入 账套导出文件夹 路径：').strip().strip('"').strip("'")
            except EOFError:
                inp = ''
            if inp:
                raw = [inp]
    if not raw:
        print('❌ 未提供文件夹路径。')
        return 1
    data_dir = raw[0].strip().strip('"').strip("'")
    if not os.path.isdir(data_dir):
        print(f'❌ 路径不存在或不是文件夹：{data_dir}')
        return 1
    produced = build_expense_combined(data_dir)
    if not produced:
        print('⚠️ 未生成任何费用底稿（可能该文件夹不含 5101/6601/6602/6603/6604 数据）。')
        return 1
    print(f'\n✅ 共生成 {len(produced)} 份费用明细底稿（外壳来源：{AUDIT_TEMPLATES_DIR}）：')
    for code, name, out in produced:
        print(f'   · {code} {name} -> {out}')
    from audit_common import finalize_after_build
    finalize_after_build(data_dir)   # 单跑收尾：对方科目补全+小计清理（与 regen 产出一致）
    return 0



# ⚡ 2026-08-09 治本：SAP 数据源适配（sap_adapter；is_sap 目录走适配分支，数据接口与 U8 同构）
try:
    import sap_adapter as _adapter
except Exception:
    _adapter = None

if __name__ == '__main__':
    # 拖入/双击本 .py 时，若当前不是托管 Python（如系统 Python 3.14 无 openpyxl），
    # 自动用托管 Python 重新执行本文件并保留控制台，避免 ImportError 黑窗闪退。
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
    # 已在托管 Python 下：执行并兜底捕获异常，写崩溃日志 + 暂停，杜绝闪退。
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
                                       f'expense_crash_{ts}.log'), 'w', encoding='utf-8') as _f:
                    _f.write('费用小程序未捕获异常：\n' + traceback.format_exc())
                print(f'已记录崩溃日志：expense_crash_{ts}.log')
        except Exception:
            pass
        rc = 1
    try:
        input('\n按回车退出…')
    except EOFError:
        pass
    sys.exit(rc)