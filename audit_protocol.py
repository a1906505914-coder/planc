# -*- coding: utf-8 -*-
"""底稿协议（2026-08-06 立规范）

背景：底稿给审计人员编制后收回，能否被程序准确读出整合数据，取决于
底稿结构是否规范。此前各 builder 的 合计/小计/段头 写法存在大量变体
（『小 计』半角空格 / 『合计：』冒号 / 『集团合计数』等），读取器只能
靠文字模糊猜测，收回整合不可靠。

本模块定义【唯一写法】，生成器输出与读取器解析共用同一套常量：
  · 段头（审定表/附注矩阵）：期初数 / 本期增加 / 本期减少 / 期末数
    （附注矩阵段头可带口径后缀，如 本期增加-贷方，前缀必须一致）
  · 合计行：『合 计』『小 计』（全角空格唯一写法）
  · 集团/抵消行：『集团合计』『合并抵消借方』『合并抵消贷方』『合并报表数』
  · 说明行：以『说明』/『勾稽』开头 → 读取器跳过（非数据行）

兼容策略：读取器解析时【去全部空白（半角+全角）后匹配规范写法】，
旧版变体（带空格/冒号/『总计』）自动兼容；生成器逐步统一到唯一写法。
"""
import re

# ---------------- 段头（唯一写法） ----------------
SEG_OPENING = '期初数'          # 审定表/附注矩阵 段头
SEG_INCREASE = '本期增加'
SEG_DECREASE = '本期减少'
SEG_CLOSING = '期末数'
SEG_LIST = (SEG_OPENING, SEG_INCREASE, SEG_DECREASE, SEG_CLOSING)

# ---------------- 行类型（唯一写法） ----------------
TOT_ROW = '合 计'               # 合计行（全角空格）
SUB_ROW = '小 计'               # 小计行
TOTAL_ROW = '总 计'             # 总计行
GRP_TOT = '集团合计'            # 附注矩阵：集团合计（各主体之和）
OFF_DEBIT = '合并抵消借方'      # 附注矩阵：合并抵消（借方）
OFF_CREDIT = '合并抵消贷方'     # 附注矩阵：合并抵消（贷方）
CONSOL_ROW = '合并报表数'       # 附注矩阵：合并报表数 = 集团合计 + 抵消借 - 抵消贷
GROUP_ROW_KEYS = ('集团合计', '集团加计', '集团合计数', '集团汇总数')
OFFSET_ROW_KEYS = ('合并抵消', '抵消借', '抵消贷')
CONSOL_ROW_KEYS = ('合并报表数', '合并数')

# 校验/说明行前缀 → 读取器跳过（非数据行）
NOTE_PREFIXES = ('说明', '勾稽', '注：', '注:', '提示', '其中：')
# 注意：『其中：』是数据行（如 其中：存放在境外的款项总额）——不能用 NOTE 排除！
# 说明类行特征：含『合计应=』『应等于』『=期初』『复核』等勾稽说明文字
CHECK_NOTE_KWS = ('合计应', '应等于', '勾稽', '复核', '应=')
# 单独判定说明行（避免误伤数据行）：
def is_note_row(s):
    """说明/勾稽行（非数据行）→ True。数据行（如『其中：存放在境外』）返回 False。

    2026-08-06 扩展（扫描 298 处"合计行读 None"假阳性根因）：
      · 前缀『注:』『备注』『口径』『核对说明』『项目(』→ 说明/表头行；
      · 含『附注披露』→ 标题行（如 信用减值损失 附注披露 — 审定数（…））；
      · 含『=』→ 说明性标题（如 期末审定数=期末未审数+审计调整、合并报表数=集团合计+抵消）。
    注意：输入只传【行名（A 列）】，不传数值列（公式值 =SUM(…) 的『=』会误判为说明行）。
    """
    s = flat(s)
    if not s:
        return False
    if any(k in s for k in CHECK_NOTE_KWS):
        return True
    if s.startswith(('说明', '提示', '注:', '备注', '口径', '核对说明', '项目(')):
        return True
    if '附注披露' in s or '=' in s:
        return True
    if re.match(r'^[一二三四五六七八九十]+[、.]', s):
        # 序号块标题判定（2026-08-06：识别『六、期末审定数（账面余额）』等附注段标题）。
        # 2026-08-07 修复：误伤损益审定表项目行——『一、主营业务收入』『二、其他业务收入』
        # 是数据行（有主体+数值），被当 note → _sheet_ent_values 整表滤空 → 核对
        # 『审定表未取到』（dq 营业收入审定表曾返回 {}）。区分规则：
        #  · 序号后为【段/组件词】→ 块标题（期初/本期/期末/审定/账面/余额/原值/折旧/
        #    减值/净值/合计）；
        #  · 序号后为【损益项目词】→ 数据行（收入/成本/利润/税金/费用）；
        #  · 其余序号行维持原判定（块标题）。
        _after = re.sub(r'^[一二三四五六七八九十]+[、.]', '', s)
        if any(k in _after for k in ('期初', '本期', '期末', '审定', '账面', '余额',
                                     '原值', '折旧', '减值', '净值', '合计')):
            return True
        if any(k in _after for k in ('收入', '成本', '利润', '税金', '费用')):
            return False
        return True
    return False


# ---------------- 归一化辅助 ----------------
def flat(s):
    """去全部空白（半角空格/全角空格/制表符）+ 全角转半角冒号。"""
    if s is None:
        return ''
    s = str(s)
    s = s.replace('：', ':').replace('（', '(').replace('）', ')')
    return re.sub(r'[\s\u3000\u00a0]', '', s)


def seg_of(s):
    """返回文本所属段（期初数/本期增加/本期减少/期末数），非段头 → None。
    兼容带口径后缀的段头（如 本期增加-贷方 → 本期增加）。"""
    f = flat(s)
    if not f:
        return None
    for seg in SEG_LIST:
        if f == flat(seg) or f.startswith(flat(seg)):
            return seg
    # 兼容旧写法：期初/年初
    if f in ('期初', '年初数', '年初'):
        return SEG_OPENING
    if f in ('期末', '年末数', '年末'):
        return SEG_CLOSING
    return None


def row_kind(s):
    """行类型：'tot' 块级合计（权威）/ 'sub' 分类小计（低优先）/ 'grp' 集团合计 /
    'offset' 抵消 / 'consol' 合并报表数 / 'data' 数据行 / 'note' 说明行 / None 空。"""
    f = flat(s)
    if not f:
        return None
    if is_note_row(f):
        return 'note'
    # 权威值行：审定数/期末审定数（=集团合计+调整，非明细数据行——附注矩阵段末
    # 『审定数』行曾当 data 累计 → 职工薪酬附注 2 倍）
    if '审定数' in f:
        return 'tot'
    # 块级合计（权威）：行名以『合计/总计』开头（可带数值后缀——_rowtxt 含数值列，
    # 如『合 计 52608.05 830780.15』flat 后 startswith('合计')）。职工薪酬审定表
    # 块内『合 计』=块权威值（父级总行+明细行双计的修复关键）。
    if f.startswith(('合计', '总计')) or f == '合计:':
        return 'tot'
    if any(k in f for k in GROUP_ROW_KEYS):
        return 'grp'                     # 集团合计（先于 endswith 判定，避免『集团合计』被当 sub）
    if any(k in f for k in OFFSET_ROW_KEYS):
        return 'offset'
    if any(k in f for k in CONSOL_ROW_KEYS):
        return 'consol'
    # 分类小计/合计（低优先，不设块值）：『银行存款合计』（全表分类合计，非主体块值）、
    # 『短期薪酬小计』（块内分类小计）、『在建工程（合计）』（父级=子级之和，行名
    # 中间含'合计'——协议化前用"包含"匹配命中跳过，曾父+子双计 2 倍）
    # 2026-08-07 修复：_rowtxt 拼接 A/B/C/D 列时行名尾部带金额数字（如
    # 『货币资金 总计 292468827.29 1503738950.82』flat 后尾随数字），endswith 判定
    # 失败 → 总计行被当数据行并入当前主体 → 货币资金审定 2 倍（Z 母公司 14.45 亿
    # +总计 15.04 亿=29.49 亿）。剥离尾部金额数字后再判定。
    _fc = re.sub(r'[\d,.\-+]+$', '', f).rstrip()
    if _fc.endswith(('合计', '总计')) or '小计' in f or '合计' in f:
        return 'sub'
    return 'data'


def is_tot_or_sub(s):
    k = row_kind(s)
    return k in ('tot', 'sub', 'grp', 'offset', 'consol')


def norm_num(v):
    """数值归一化（容忍千分位/负号/全角数字）。"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(',', '').replace('，', '').strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


# ---------------- 审定表列序（唯一约定） ----------------
# 审定表（flow 型）：列序固定（2026-08-06 协议）
COL_OPENING = 0    # 期初数
COL_INCREASE = 1   # 本期增加
COL_DECREASE = 2   # 本期减少
COL_CLOSING = 3    # 期末数
COL_ADJ_DEBIT = 4  # 审计调整借方
COL_ADJ_CREDIT = 5  # 审计调整贷方
COL_RECLASS_DEBIT = 6  # 重分类调整借方
COL_RECLASS_CREDIT = 7  # 重分类调整贷方
