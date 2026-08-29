# -*- coding: utf-8 -*-
"""虚增识别知识库（2026-08-14 用户方法论定稿：踢虚增 = 两层证据 + 格式边界）。

用户原话（方法论）：
    "如何踢虚增也是个知识点，有 2 个层面：
     1 是仔细阅读各列信息，通过列识别是否可以直接提取到数据；
     2 是进一步学习 sap 记账习惯，哪些情况是虚增，这样如果漏了列信息，也能找出来。
     另外如果其他格式，我们没有使用过，实际是无法协助的。"

三层结构（本文件就是该知识本体的可执行化）：
    层1 · 列信息识别（直接证据）：读列 → 直接提取
        数据形态告诉我们"是什么"——同凭证净额化 / 同额镜像 / 科目·日期·合同定位。
        依赖列信息完整；列不全/解析漏列 → 只能给形状，不能定性（交给层2兜底）。
    层2 · SAP 记账习惯（推理证据）：懂习惯 → 推断（虚增 + 科目定位）
        业务语义告诉我们"为什么双计"——购汇/手续费/受托/资金池/待报解/汇总记账。
        即使列信息被漏掉，凭"这类业务必然双计"的规律也能反查锁定。
        用户补充（2026-08-14）：SAP 特有方式不止虚增——**找具体费用科目**（6600
        总池按 col19 功能范围列拆分）、内部码方言、GL 无科目名称列等，同样是 SAP
        记账习惯知识，一并纳入本层。
    边界 · 格式覆盖（诚实声明）：已学格式可自主协助；未用格式无法协助。
        先"学习"再纳入：manifest 声明 + 列探测 + 上期 vs 本期接入核对。

用法：
    check_coverage(fmt)          → 格式边界检查（诚实声明，不硬猜）
    explain_layers()             → 打印两层知识总表
    match_habits(rows, ...)      → 用层2习惯规则在未配行中找虚增候选（供核对程序调用）
    match_subject_customs(rows)  → 用层2科目定位习惯识别 SAP 特有费用科目拆分
"""
from collections import defaultdict

# ════════════════════════════════════════════════════════════════════
# 边界 · 格式覆盖（诚实声明：没学过的格式无法协助，绝不硬猜）
# ════════════════════════════════════════════════════════════════════
FORMAT_COVERAGE = {
    # fmt -> {name, probe(判别方式), parser}
    'u8_tb_gl':      {'name': 'U8 科目余额表/综合查询明细表',
                      'probe': '文件名含主体+年度，一主体一文件',
                      'parser': 'audit_common.read_tb / iter_comp_gl'},
    'sap_39col':     {'name': 'SAP 39 列综合查询明细表（区间文件，一文件多主体）',
                      'probe': '按内容扫 col7 主体列（铁律72）',
                      'parser': 'sap_adapter.read_gl_rows'},
    'icbc_csv':      {'name': '工行 CSV 三格式（3100两列/3000单列/3400转入转出）',
                      'probe': '表头探测',
                      'parser': 'bank_statement_parser.parse_icbc_csv'},
    'tailong':       {'name': '泰隆银行流水',
                      'probe': '表头探测',
                      'parser': 'bank_statement_parser.parse_tailong'},
    'pufa':          {'name': '浦发银行流水',
                      'probe': '表头探测',
                      'parser': 'bank_statement_parser.parse_pufa'},
    'probed_generic': {'name': '通用表头探测（parse_probed 兜底）',
                       'probe': '列名特征探测',
                       'parser': 'bank_statement_parser.parse_probed'},
}


def check_coverage(fmt, strict=False):
    """格式边界检查。strict=True 时未知格式直接拒绝协助（返回 False + 原因）。"""
    if fmt in FORMAT_COVERAGE:
        return True, FORMAT_COVERAGE[fmt]['name']
    if strict:
        return False, f'格式 [{fmt}] 未学习过，无法协助——先做 manifest 声明 + 列探测 + 接入核对'
    return False, f'格式 [{fmt}] 未在学习清单内，识别结果不可靠'


# ════════════════════════════════════════════════════════════════════
# 层1 · 列信息识别（直接证据）：数据形态 → 直接提取
# ════════════════════════════════════════════════════════════════════
COLUMN_SIGNALS = {
    'vchar':    '凭证号：同凭证借+贷 → 净额化（剔除完全互抵）',
    'subject':  '科目：科目判定（名称主键/内部码前缀）',
    'customer': '客户：往来明细维度（按 单位×合同号 分行）',
    'vendor':   '供应商：往来明细维度',
    'vdate':    '凭证日期：期间口径（col28，禁 col0 分配指令）',
    'amount':   '金额：净额化 / 配平（大额差异≥100万呈倍数/滚动/跨主体=bug）',
    'contract': '合同：往来双维度（SAP 按 单位@@合同号）',
    'order':    '订单：业务定位（col25 订单/物料）',
    'func_area': '功能范围：费用分类（6601销售/6602管理/5101制造/5301研发）',
    'wbs':      'WBS：项目定位（col17）',
    'cp':       '对方科目：凭证内借贷互证（铁律62）',
    'purch_doc': '采购凭证：指令号列（col27）',
}

# ════════════════════════════════════════════════════════════════════
# 层2 · SAP 记账习惯（推理证据）：业务语义 → 推断双计（漏列也能锁定）
# ════════════════════════════════════════════════════════════════════
SAP_HABITS = [
    {
        'name': '购汇/结汇双计',
        'sign': '贷侧出现"预付款"；同一外汇业务购汇分录+售汇分录',
        'why':  '银行售汇与企业购汇各记一笔，本质同一笔外汇业务',
        'evidence': '对称双计（收差≈支差）',
        'fix':  '_entrust_pay_settle 贷侧"预付款"关键词',
        'case': '3300 建行 5.2 万购汇',
    },
    {
        'name': '手续费/服务费双计',
        'sign': '收款关键词（二维码/一码通/饮料机/浴室充值）+ 扣费关键词（手续费/服务费/退押金）',
        'why':  '网银记净额、账侧记毛额，收+支同额成对',
        'evidence': '对称双计（P1），月内收差=支差',
        'fix':  '_fee_net_journal 净额化（第二级月聚合跳过 removed_ids）',
        'case': '3100/3300 农商 682 元',
    },
    {
        'name': '受托支付 N:M',
        'sign': '受托放款与受托还款同月多笔，贷侧"预付款"',
        'why':  '受托 N 笔放款对应 M 笔还款，同月合计平衡',
        'evidence': '同月合计平衡（N:M 配对）',
        'fix':  '_entrust_pay_settle N:M 同月合计平衡',
        'case': '农行 9.5 亿受托',
    },
    {
        'name': '资金池对冲',
        'sign': '实归支取借 ↔ 工资/税贷，1:1 同额成对',
        'why':  '集团资金池归集，收支同额互冲净额 0',
        'evidence': '镜像成对（P9）',
        'fix':  '_pool_wage_settle 1:1 同额配对 + 整月平衡兜底',
        'case': '盛虹/国望资金池',
    },
    {
        'name': '待报解通道',
        'sign': '网银"待报解/代理国库/收缴" ↔ 账侧税费',
        'why':  '代扣代缴两种形态：网银待报解过渡户 vs 账记税费',
        'evidence': '待报解残留（P10）',
        'fix':  'daikou 剔除（j_rm 全量 journal 筛）',
        'case': '农商 709 万、3300 农商 5434 万',
    },
    {
        'name': '汇总记账',
        'sign': '网银逐笔 vs 账月末汇总/区间（一码通/社保）',
        'why':  '账侧月末汇总入账，网银逐笔',
        'evidence': '月汇总匹配/找和（P11）',
        'fix':  '月汇总匹配 + 找和算法',
        'case': '一码通/社保',
    },
    {
        'name': '红字冲正/退款',
        'sign': '网银负金额行（退款冲正）',
        'why':  '退款冲正账侧未记或形态不同（收-X vs 贷X）',
        'evidence': '红字未配（P7），负数归一化后应配对',
        'fix':  '红字归一化（_orig_amt/_orig_dir）后配对',
        'case': '一码通退款 -45 元',
    },
    {
        'name': '跨期错位',
        'sign': '某月+X 相邻月-X',
        'why':  '跨期未达/月末滞后记账',
        'evidence': '逐月收差符号相邻相反（P8）',
        'fix':  '45 天窗口/摘要业务日期匹配',
        'case': 'DQ 中行 500 万、6/7 月 2.585 亿',
    },
    {
        'name': '镜像成对（划转）',
        'sign': 'ub 内同额收+支',
        'why':  '资金池/跨境划转账未记，净额 0',
        'evidence': '同额收+支成对（P9）',
        'fix':  '_classify_bank_unmatched 全量镜像配对（auto 单边保留）',
        'case': '3300 建行 45 元镜像×3',
    },
    {
        'name': '定期/通知存款账户',
        'sign': '科目名含"定期/通知存款/大额存单/结构性存款"',
        'why':  '定期/通知存款账户与活期分开，网银流水=活期——定期存取款在网银无对应，抽入成假"账有网银无"',
        'evidence': '账有网银无 + 科目名含定期/通知存款（DQ2025 7科目25行/DQ2026 3科目实证）',
        'fix':  'read_bank_journal/read_sap_journal_file 抽取时排除（科目名关键词过滤）——活期才与网银核对，定期/通知/外币另核',
        'case': 'DQ2025 嘉兴/宁波/招行/杭州/民生定期 25 行；DQ2026 嘉兴/杭州/交行定期',
    },
    {
        'name': 'E信/供应链融资到期清分',
        'sign': '网银『数字信用凭据融资清分往来款项』 ↔ 账侧『工银e信到期』',
        'why':  '工银E信（供应链数字信用凭据）到期清分，网银按凭据逐笔、账侧按到期日汇总，同业务不同表述',
        'evidence': '网银有账无+账有网银无同现，月度对账平（两边总额一致，3200 工行全年差 -1.91 实证）',
        'fix':  '业务关键词配对：数字信用凭据↔工银e信，N:M 找和+净额验证（注意账侧 E信到期可能大于网银数字信用——部分走其他摘要）',
        'case': '3200 工行 7月 网银77笔 2,092万 vs 账侧工银e信到期',
    },
]

_HABIT_INDEX = {h['name']: h for h in SAP_HABITS}

# ════════════════════════════════════════════════════════════════════
# 层2 · SAP 特有科目定位习惯（2026-08-14 用户补充："包括前面找具体费用科目，
#   也是 sap 的特有方式"——SAP 记账习惯不止虚增，还有"怎么从总池拆出明细科目"）
# ════════════════════════════════════════════════════════════════════
SAP_SUBJECT_CUSTOMS = [
    {
        'name': '费用按功能范围拆分',
        'sign': 'GL 有功能范围列（col19）；6600=期间费用总池',
        'why':  'SAP 费用统一记 6600 总池，具体科目（销售/管理/研发/制造）不在科目码上，'
                '而在功能范围列——分类依据=col19（6601销售/6602管理/5101制造/5301研发）；'
                '660006 一律研发；6603=财务',
        'evidence': '6600 总池与利润表明细对不上 → 必须按功能范围列拆',
        'fix':  '_sap_fee_cat（限定 6600/6603 系列按 col19 拆分）',
        'case': '铁律89：利润表费用分类',
    },
    {
        'name': '科目内部码方言',
        'sign': 'SAP 科目代码 ≠ 标准科目码（1811递延所得税资产/2260租赁负债/2401递延收益/'
                '2701长期应付款/2502应付债券/2221应交税费等）',
        'why':  'SAP 记账习惯：内部码自编，与准则科目号不同——科目判定只能靠名称匹配或'
                '内部码前缀，不能按标准码硬套',
        'evidence': '标准科目码在 SAP GL 中查不到 → 查 SAP 内部码方言表',
        'fix':  '名称匹配为主 + SAP 科目代码方言表（GL 侧用内部码前缀）',
        'case': '2026-08-09 实测 30+ 个 SAP 方言码',
    },
    {
        'name': 'GL 无科目名称列',
        'sign': 'SAP GL 只有科目代码列，无名称列',
        'why':  '记账习惯：名称只存在于 TB 侧，GL 侧科目判定只能用内部码前缀',
        'evidence': 'GL 取不到科目名 → 判定走内部码前缀',
        'fix':  'read_sap_gl 科目判定用内部码前缀；名称仅 TB 侧解析',
        'case': '全部 SAP 账套',
    },
]

_SUBJECT_CUSTOMS_INDEX = {c['name']: c for c in SAP_SUBJECT_CUSTOMS}


def match_subject_customs(rows, func_col=None):
    """用层2科目定位习惯识别 SAP 特有费用科目拆分候选。

    rows: 费用总池行（6600/6603 系列）。
    func_col: 功能范围列取值（col19），如 '6602 管理费用'——用于展示拆分依据。
    """
    hits = []
    for r in rows:
        text = ' '.join(str(r.get(k) or '') for k in
                        ('summary', 'counterparty', 'cp', 'name', 'subject'))
        if '6600' in text or '6603' in text:
            hits.append({
                'custom': '费用按功能范围拆分',
                'why': SAP_SUBJECT_CUSTOMS[0]['why'],
                'fix': '_sap_fee_cat 按 col19 拆分（6601销售/6602管理/5101制造/5301研发）',
                'row': r,
                'func_col': func_col,
            })
    return hits


def match_habits(rows, evidence_fn=None):
    """用层2习惯规则扫描未配行，返回命中的习惯候选列表。

    rows: 未配行（ub 或 uj），每行含 direction/amount/summary/counterparty/date。
    evidence_fn: 可选，接收行返回额外证据文本。
    """
    hits = []
    keywords = {
        '购汇/结汇双计': ('购汇', '结汇', '售汇', '预付款'),
        '手续费/服务费双计': ('手续费', '服务费', '退押金', '二维码', '一码通', '饮料机', '浴室充值'),
        '受托支付 N:M': ('受托', '委托贷款'),
        '资金池对冲': ('实归', '支取', '工资归集', '税贷'),
        '待报解通道': ('待报解', '代理国库', '收缴'),
        '汇总记账': ('一码通', '社保', '汇总'),
        # ⚡⚡ 2026-08-14 定期/通知存款（账户级特征：科目名，非摘要关键词——DQ 实证）
        '定期/通知存款账户': ('定期存款', '通知存款', '大额存单', '结构性存款'),
        # ⚡⚡ 2026-08-14 E信/供应链融资（网银数字信用凭据 ↔ 账侧工银e信到期，3200 实证）
        'E信/供应链融资到期清分': ('数字信用凭据', '工银e信', '工银E信', 'E信到期'),
    }
    for r in rows:
        text = ' '.join(str(r.get(k) or '') for k in
                        ('summary', 'counterparty', 'cp', 'name'))
        for habit, kws in keywords.items():
            if any(k in text for k in kws):
                h = _HABIT_INDEX.get(habit)
                if h:
                    ev = evidence_fn(r) if evidence_fn else ''
                    hits.append({'habit': habit, 'why': h['why'],
                                 'fix': h['fix'], 'row': r, 'extra': ev})
    return hits


def explain_layers():
    """打印两层知识总表（自查/培训用）。"""
    print('===== 虚增识别知识库 =====')
    print('── 层1 · 列信息识别（直接证据）──')
    for k, v in COLUMN_SIGNALS.items():
        print(f'  {k:12s}: {v}')
    print('── 层2 · SAP 记账习惯（推理证据）──')
    for h in SAP_HABITS:
        print(f'  【{h["name"]}】{h["sign"]}')
        print(f'      为什么：{h["why"]}')
        print(f'      证据：{h["evidence"]} → 处理：{h["fix"]} [{h["case"]}]')
    print('── 层2 · SAP 特有科目定位习惯（用户补充：找费用科目也是 SAP 特有方式）──')
    for c in SAP_SUBJECT_CUSTOMS:
        print(f'  【{c["name"]}】{c["sign"]}')
        print(f'      为什么：{c["why"]}')
        print(f'      证据：{c["evidence"]} → 处理：{c["fix"]} [{c["case"]}]')
    print('── 边界 · 格式覆盖（诚实声明）──')
    for f, m in FORMAT_COVERAGE.items():
        print(f'  {f:14s}: {m["name"]}（{m["probe"]}）')


if __name__ == '__main__':
    explain_layers()
    print()
    ok, name = check_coverage('sap_39col')
    print(f'check_coverage(sap_39col): {ok} → {name}')
    ok, name = check_coverage('某新银行 xlsx', strict=True)
    print(f'check_coverage(未知, strict): {ok} → {name}')
