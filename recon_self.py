# -*- coding: utf-8 -*-
"""阶段③ recon_self —— 底稿 vs 自建试算表核对（2026-08-06）

背景：用户方法论——拿到账套先自建试算表（摸结构），再生成底稿，然后【底稿 vs 自建试算表】
同源核对：两边都来自单体科目余额表，应完全一致，差异 = 取数 bug（一抓一个准，零噪音）。
（外部试算表核对是最后一步：自建 vs 外部 = 审计调整/合并抵消/重分类口径差。）

机制：
  1) self_tb_gen.build_entity_tb(data_dir) —— 叶子过滤后各主体 TB（权威控制数）
  2) subjects_registry.REGISTRY —— 每科目的 codes（代码前缀）
  3) 底稿侧：snapshot_regress._extract（审定数，data_only 读数值——协议化后权威行全是数值）
  4) 逐科目 × 逐主体：|自建TB该科期末| vs |底稿审定数|，差>0.01 报差异（0.01 铁律）
  5) 损益类（期末余额=0）自动跳过（发生额科目不在试算表期末列）

输出：差异清单（科目/主体/自建TB/底稿审定/差）+ 汇总。返回 (ok, diffs)。
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def recon_self(data_dir, years=None, quiet=False, wp_dir=None):
    """阶段③：底稿审定数 vs 自建试算表。返回 (ok, diffs)。
    diffs = [(科目名, file_key, 主体, 年度, 自建TB期末, 底稿审定, 差)]
    ⚡ 2026-08-24：wp_dir 可选指定底稿目录（集团模式底稿在独立 out-dir；缺省=data_dir）。"""
    import audit_common
    import self_tb_gen as S
    import subjects_registry as REG
    from snapshot_regress import _extract
    if wp_dir is None:
        wp_dir = data_dir

    if not os.path.isdir(data_dir):
        return True, []
    # 1) 自建试算表（各主体叶子 TB）
    # 2026-08-07 修复⑬：不再用 self_tb_gen.build_entity_tb 的【一级父级名替换】产物
    # （4104 叶子名被替换为『利润分配』→ 名称匹配『未分配利润』失败 → 代码聚合
    # 4104 全子目含应付股利 25M → c 未分配利润 -87.4M vs 底稿 112.4M 误报）。
    # 改为直接读 read_tb_full 原始 TB + 叶子过滤（保留子目名『利润分配-未分配利润』）。
    entities = audit_common.discover_entities(data_dir)
    tb_full = audit_common.read_tb_full(data_dir, entities)
    # ⚡⚡ 2026-08-24 修复（AZ 泰国账套同源核对大面积差异根因）：泰国主体科目代码/名称
    #   与集团标准不同（应收 1130≠1122、固定资产非 1601 等）——底稿审定表经
    #   thai_mapping.normalize_thai_tb 归一化已含泰国，而自建 TB 侧未归一化 → 按标准码
    #   取数漏掉泰国主体 → 自建TB(不含泰国) vs 底稿审定(含泰国) 系统性大差异。
    #   此处与 builder 同口径归一化后再核对。
    try:
        from thai_mapping import normalize_thai_tb as _norm_thai
        tb_full = _norm_thai(tb_full, entities)
    except Exception:
        pass
    # ⚡⚡ 2026-08-24 父级一致性修正：源 TB 父级≠子目合计（父级陈旧）时，以子目合计为准
    #   （与底稿叶子口径一致）→ 自建试算表不再系统性大于/小于底稿（AS 固定资产 8.3亿、
    #   GJX 在建 2.76亿、AJJ 在建 1451万 等差异根因）。
    try:
        from tb_parent_fix import leaf_total_tb as _leaf_fix
        tb_full = _leaf_fix(tb_full)
    except Exception:
        pass
    if not tb_full:
        print('  ⚠️ [自建试算表] 无主体 TB，跳过阶段③')
        return True, []
    t_years = sorted({str(yy) for (_e, _c, _n, yy) in tb_full}) or ['2025']
    if years is None:
        years = t_years
    # 逐年度：{ent: {code: 原始TB行}}（叶子过滤，保留原始子目名 + 一级名）
    ent_tb_all = {}
    for ent in sorted(entities):
        ent_tb_all[ent] = {}
        for yy in years:
            items = [(c, v, n) for (e, c, n, y), v in tb_full.items()
                     if e == ent and str(y) == yy]
            leafs = S.leaf_codes([c for c, _v, _n in items])
            # 2026-08-07 修复⑮：read_tb_full 的 v 字典无 'name' 字段（name 在键里）
            # → 必须把 name 塞进 v，否则 merged 后名称匹配全失效（c 固定资产 89.4M
            # 原值 vs 51.6M 净值误报根因）。
            # 2026-08-07 修复⑰：附一级名 l1_name（4 位代码最短行名称）——H 实收资本
            # 3003 股本 叶子=300301~300333 股东名（不含『股本』字样）→ 名称匹配失效
            # → 回退代码 4001 误取生产成本 2.2M（自建TB=2,235,812.23 vs 底稿 76.2M）。
            # 权益类（实收/资本公积/盈余公积）叶子名=股东/子目名，必须用一级名匹配。
            l1_names = {}
            # ⚡⚡ 2026-08-24 修复（ADZ 实收资本 3001 子目漏取根因）：l1_names 原取
            #   『首个该前缀行』名称，若子目（300101 蔡晓锋）先于父级（3001 实收资本）
            #   出现在迭代序 → l1=股东名 → 权益类叶子（300101~300103 股东名）名称匹配
            #   全失效（ADZ 仁旭/海拓/金泰 实收资本 6351万 漏取）。优先用【4 位父级行】
            #   名称（修复⑱仍保留：无父级行时取首个子目非数字段，防 JTt 410401 误配）。
            _parent_segs = {}
            for c, _v, n in items:
                l1 = re.sub(r'\D', '', str(c))[:4]
                if not l1:
                    continue
                segs = [s for s in re.split(r'[\\/\-]', str(n))
                        if s.strip() and not re.match(r'^\d+$', s.strip())]
                nm0 = segs[0].strip() if segs else str(n)
                if str(c) == l1:
                    _parent_segs[l1] = nm0   # 父级行优先（后遇到也覆盖）
                else:
                    _parent_segs.setdefault(l1, nm0)
            l1_names = dict(_parent_segs)
            ent_tb_all[ent][yy] = {c: dict(v, name=n,
                                           l1_name=l1_names.get(re.sub(r'\D', '', str(c))[:4], ''))
                                   for c, v, n in items if c in leafs}
    # 2) 底稿审定数（逐科目文件）
    diffs = []
    _acct = os.path.basename(data_dir.rstrip('\\/'))
    for key, subj in REG.REGISTRY.items():
        # 2026-08-07 v2（会计语言优先，代码方言隔离）：codes 经账套配置解析——
        # 账套特例码（dq 递延收益=2601、H 实收资本=3003、Q 租赁负债=2271 等）登记在
        # account_profiles.json subject_codes，【不进全局注册表】→ 消除跨账套撞车
        # （FY 递延收益 2601=租赁负债误并入、FY 一年内到期 2401=递延收益误并入）。
        try:
            import subject_mapping as _sm
            codes = _sm.resolve_subject_codes(_acct, key, subj.get('codes') or [])
        except Exception:
            codes = subj.get('codes') or []
        if not codes:
            continue
        # 2026-08-07 修复④：损益类科目跳过——自建TB 期末余额=未结转损益（结转后=0），
        # 底稿审定=发生额（gross，收入贷/费用借），两者口径不可比（FY 营业收入
        # 自建TB -7.7M 未结转 vs 底稿 566M 发生额误报）。损益核对由 GL 铁律3/27 负责。
        if any(str(c)[:1] in '678' for c in codes):
            continue
        fk = subj['file_key']
        for yy in years:
            # ⚡ 2026-08-24 兼容集团模式命名（run_u8_on_sap --out-dir 集团稿）：
            #   {科目}审计底稿_{yy}集团.xlsx（集团产物）或 {科目}审计底稿_{yy}_生成.xlsx（单体/常规）。
            #   ⚡ 集团目录可能混有旧 _生成 残留（单体稿）→ 优先匹配集团命名，防取错旧稿。
            fp = os.path.join(wp_dir, f'{fk}审计底稿_{yy}集团.xlsx')
            if not os.path.exists(fp):
                fp = os.path.join(wp_dir, f'{fk}审计底稿_{yy}_生成.xlsx')
            if not os.path.exists(fp):
                continue
            try:
                aud, _det, _fn, _s, _y = _extract(fp, year=yy)
            except Exception:
                continue
            if aud is None:
                continue
            # 自建 TB 侧：该科目期末聚合（abs 对齐底稿正数显示）。
            # 2026-08-06 修复①：按【名称】匹配（铁律5：代码不作语义判定）——
            # H 账套 4001=生产成本（非实收资本）、实收资本=3003『股本』；原 codes
            # startswith(4001) 匹配到生产成本 2.2M 误报。名称匹配=subj 名与实际名
            # 互相包含（'股本'⊂'实收资本(或股本)'）。
            # 2026-08-07 修复⑤：固定资产族双口径自适应——FY 固定资产底稿审定=原值
            # （合计行 1,140M），H 固定资产底稿=净值（原值-累计折旧 337M-144M=193M）。
            # 各账套底稿口径不一 → 自建TB 侧同时算 原值(排除备抵) 与 净值(备抵带符号)，
            # 取与底稿审定数更接近者比对（FY 原值 1,140M ✓ / H 净值 193M ✓）。
            _excl_kw = ('清理', '待处理', '待处置', '受托', '代管',
                        '待认证')  # ⚡⚡ 2026-08-25 待认证进项税额(222126)：借方留抵过渡科目，
                                    #   tax_detail 生成器明确剔除（VAT_KW 含『待认证』），
                                    #   recon_self 同步排除（AYL 应交税费 3390万 误报根因）。
            # 2026-08-07 修复⑯：备抵科目代码前缀排除——F 账套 1231.01 名称=『应收账款』
            # （实为 坏账准备-应收账款 子目，1231.02=其他应收款 -360M），名称匹配误把
            # 备抵子目计入主科目（F 应收账款 差 6.7K、其他应收款 差 -491M 根因）。
            # 名称无法区分（子目名=主科目名），只能按备抵族代码前缀排除。
            _contra_code_prefix = ('1231', '1471', '1472', '1602', '1603',
                                   '1702', '1703', '1522', '1523', '1632')
            # 2026-08-07 修复⑥：备抵词加宽——FY 折旧科目名『使用权资产折旧』（非
            # 『使用权资产累计折旧』），原 _contra_kw 漏判 → 折旧计入原值口径误报
            # （FY 使用权资产 51.5M 净值 vs 底稿 78.6M 原值，差 27M=折旧额）。
            _contra_kw = ('累计折旧', '累计摊销', '减值准备', '跌价准备', '坏账准备',
                          '未确认融资费用', '使用权资产折旧', '折旧', '摊销')
            subj_name = subj.get('name', '')
            # 2026-08-07 修复㉔：registry codes 是多账套并集必然撞车（di 的 2601 在
            # FY=租赁负债、ncl 的 2401 在 FY=递延收益）。判定撞车：【代码命中行】的
            # 名称包含【另一个注册表科目名】（如『租赁负债-未确认融资费用』含『租赁
            # 负债』）→ 该行属其他科目，剔除。不误伤：银行存款/原材料/库存商品 非
            # 注册表科目名，不受影响（FY 货币资金 1002 银行存款 保留）。
            _reg_other_names = [n2 for n2 in REG.REGISTRY.values()
                                for n2 in (n2.get('name', ''),)
                                if n2 and n2 != subj_name and len(n2) >= 3]
            # ⚡⚡ 2026-08-24（AQ 1704 开发支出 / ADZ 4001 生产成本 误配根因）：
            #   registry 多账套并集 codes 含 1704（rua=FY 使用权资产）但 AQ 1704=
            #   开发支出；4001（paid_in）在 ADZ 东轴=股本、仁旭等=生产成本。这两
            #   个他科名不在 REGISTRY（开发支出未收录、生产成本非报表科目）→ 显式
            #   加入撞车词（2026-08-24 二次评估：_name_relate 已改为纯撞车检查，无
            #   强相关性校验，撞车词对存货/在建等子科目名无副作用）。
            _reg_other_names = _reg_other_names + ['开发支出', '生产成本']

            def _name_hit(n, l1n, subj_name):
                """名称分段匹配（铁律5 名称优先）。
                2026-08-07 修复⑭：名称按 \\ / - 分段后段级匹配，兼容：
                · 子目名『利润分配-未分配利润』→ 段『未分配利润』命中（修复⑬ 保留
                  原始子目名后必须分段，否则 - 不分割会整体失配）；
                · U8 反斜杠层级『22030101\\合同负债\\预收工程款\\已开票』→ 段命中；
                · 排除子串误配：『长期待摊费用摊销』（段≠subj 名，subj 名长度差>4
                  判为不同科目）vs『长期待摊费用』；『折旧费用与长期待摊费用』同理。
                2026-08-07 修复⑰：正向优先 + 一级名兜底——权益类叶子名=股东/子目
                名（H 300301~333 股东名不含『股本』），名称匹配须同时查 l1_name
                （一级名『股本』）；且反向包含（TB 名⊂subj 名）仅正向无命中时启用
                （『待摊费用』⊂『长期待摊费用』不误配）。
                """
                def _hit(s, subj_name):
                    if not s or not subj_name:
                        return False
                    # 2026-08-07 修复⑳：提取/分配类子目排除——『提取法定盈余公积』
                    # 『提取任意盈余公积』『应付现金股利』是利润分配(4104)借方发生额
                    # 子目，名称含『盈余公积』/『股利』字样但非余额科目（F 盈余公积
                    # 9.5M、JTt 24.3M 误配根因）。
                    if s.startswith('提取') or '应付现金股利' in s or '应付股利' in s:
                        return False
                    if s == subj_name:
                        return True
                    if subj_name in s and (s.startswith(subj_name) or s.endswith(subj_name)) \
                            and len(s) - len(subj_name) <= 4:
                        return True
                    return False

                def _rev_hit(s, subj_name):
                    if not s or not subj_name:
                        return False
                    # 修复⑳：提取/分配类子目排除（与 _hit 同规则）
                    if s.startswith('提取') or '应付现金股利' in s or '应付股利' in s:
                        return False
                    # 2026-08-07 修复㉓：反向匹配排除通用短词——F 2241.99/4002.02.03/
                    # 6301.99/6602.99 名『其他』，'其他' in '其他应收款' 恒 True → 全部
                    # 误配进其他应收款（-104.4M vs 13.3M）。反向匹配仅当 s 是【完整
                    # 标准科目名】：排除『其他/本年/以前/待处理』等通用词；但『股本』
                    # 『资本』等 2 字标准名必须保留（H 实收资本=3003 股本 76.2M）。
                    if s in ('其他', '本年', '以前', '待处理', '上期', '本期', '期初', '期末'):
                        return False
                    return s in subj_name

                if not subj_name:
                    return False
                # 2026-08-07 修复㉕：全半角括号归一化——FY TB 一级名『实收资本（或
                # 股本）』（全角 0xff08/0xff09）vs 注册表『实收资本(或股本)』（半角）
                # 名称匹配失败 → 权益类叶子名（股东名）l1 兜底失效 → FY 实收资本
                # 自建TB -3.49M vs 底稿 86.4M（差 89.9M 根因）。
                def _norm_br(s):
                    return str(s).replace('（', '(').replace('）', ')').replace('，', ',').replace('、', ',')
                _sn = _norm_br(subj_name)
                cand = [_norm_br(str(n))]
                if l1n:
                    cand.append(_norm_br(str(l1n)))
                # 正向：subj 名 ⊂ TB 名 / 一级名（优先）
                for s in cand:
                    for seg in re.split(r'[\\/\-]', s):
                        if _hit(seg.strip(), _sn):
                            return True
                # 反向：TB 名 / 一级名 ⊂ subj 名（仅正向无命中时启用）
                for s in cand:
                    for seg in re.split(r'[\\/\-]', s):
                        if _rev_hit(seg.strip(), _sn):
                            return True
                return False

            def _name_relate(n, l1n, subj_name):
                """代码命中行是否属本科目（名称主键铁律58）。
                仅做【撞车检查】：名称含其他注册表科目名/已知他科词（开发支出、
                生产成本等）→ 他科剔除。不做强名称相关性——存货子目（原材料/
                库存商品）与『存货』名称无关但属本科目，货币资金子目同理
                （2026-08-24 ADZ 存货 7327万 误剔教训）。
                ⚡⚡ 2026-08-25 修复（AYL 未分配利润 96M 根因，与 _prefix_ok 同步）：
                利润分配(4104)子目『利润分配-提取法定盈余公积』名称含他科名
                『盈余公积』→ 原撞车检查误剔 → 未分配利润只按名称取到 410407，
                漏提取盈余公积净额。利润分配族子目（首段=『利润分配』或 提取*
                分配动作）属本科目，不参与撞车剔除。"""
                if not subj_name:
                    return True
                _n = str(n)
                _first_seg = re.split(r'[\\/\-]', _n)[0].strip() if _n else ''
                if _first_seg == '利润分配' or _n.startswith('提取'):
                    return True
                for _rn in _reg_other_names:
                    if _rn and _rn in _n:
                        return False
                return True

            def _tb_qm(ent_tb):
                """返回 (原值口径合计, 净值口径合计)。
                2026-08-07 修复⑪：名称匹配【正向优先，反向兜底】——反向匹配
                （TB 名 ⊂ subj 名）仅在正向（subj 名 ⊂ TB 名）无任何命中时启用：
                · 长期待摊费用：正向命中 1801 → 反向不启用 → 1472『待摊费用』
                  （'待摊费用'⊂'长期待摊费用'）不再误配；
                · H 实收资本：正向 '实收资本(或股本)'⊂'股本' ✗ → 反向 '股本' ✓ 仍命中。
                修复⑫：_pick 按 codes 数量分流——多科目聚合（货币资金 3 码）名称
                只命中『其他货币资金』漏库存现金/银行存款 → 名称∩代码非空回退代码；
                单科目（未分配利润 4104）代码前缀会带进 4104.03 应付股利等兄弟子目
                → 信任名称（4104.06 未分配利润）。"""
                s_g_n = s_c_n = 0.0
                s_g_v = s_c_v = 0.0
                rows_g_n, rows_c_n = [], []
                rows_g_v, rows_c_v = [], []
                _prefix_cache = {}
                pos_hit = False
                # ① 正向名称匹配（subj 名 ⊂ TB 名）
                for (c, _n0), v in ent_tb.items():
                    n = str(v.get('name', ''))
                    if any(k in n for k in _excl_kw):
                        continue
                    if subj_name and _name_hit(n, str(v.get('l1_name', '')), subj_name):
                        pos_hit = True
                        break
                for (c, _n0), v in ent_tb.items():
                    n = str(v.get('name', ''))
                    if any(k in n for k in _excl_kw):
                        continue
                    # 修复⑯：备抵族代码前缀（1231 坏账准备等）在名称匹配前排除——
                    # 其子目名=主科目名（F 1231.01『应收账款』），名称无法区分。
                    # ⚡ 2026-08-11 修复（DQ 固定资产 305M 误报）：原逻辑把 1602 累计折旧/
                    #   1603 减值等【独立备抵科目】也整行 continue → 净值口径 v2 不含折旧 →
                    #   固定资产/无形资产 双口径都=原值，无法匹配底稿净值（差=折旧额）。
                    #   修复：备抵族行【名称含备抵词】→ 计入净值口径 net（is_contra 不进口径
                    #   gross）；仅『名称=主科目名』的子目（1231.01『应收账款』）才整行排除。
                    if any(str(c).startswith(p) for p in _contra_code_prefix):
                        # ⚡ 2026-08-11 三次修复：备抵行须【属于本科目 codes】才计净值
                        #   （_contra_code_prefix 全局，fa 循环里 1702 无形资产摊销被误扣
                        #   → fa 净值 514.5-20.8=493.7M 串号）。1602/1603∈fa、1702/1703∈ia。
                        if (codes and any(str(c).startswith(k) for k in codes)
                                and any(k in n for k in _contra_kw)):
                            s_g_v += v['qm']; rows_g_v.append(str(c))
                            s_c_v += v['qm']; rows_c_v.append(str(c))
                        continue
                    is_contra = any(k in n for k in _contra_kw)
                    hit = _name_hit(n, str(v.get('l1_name', '')), subj_name)
                    # ⚡ 2026-08-16 修复（ga 租赁负债误配）：名称命中但【代码不在本科目
                    #   codes】且【名称首段≠本科目名】→ 他科子目段（181104『递延所得税
                    #   资产\租赁负债』段'租赁负债'命中，但 1811 非租赁负债 codes、
                    #   首段'递延所得税资产'≠'租赁负债'）→ 排除，防并入他科余额。
                    #   不误伤：220601/220602（2206 ∈ codes）、未分配利润（4104 ∈ codes）。
                    if hit and codes and not any(str(c).startswith(k) for k in codes):
                        _first_seg = re.split(r'[\\/\-]', n)[0].strip()
                        # ⚡⚡ 2026-08-24 修复（ADZ 实收资本 3001 子目漏取根因）：原判定
                        #   『首段 != subj_name → 剔除』误伤 ADZ 3001『实收资本』
                        #   （subj=『实收资本(或股本)』，首段『实收资本』⊂ subj 名，是
                        #   报表同义词/子集）。改为『首段与 subj_name 完全无包含关系
                        #   才剔除』（181104『递延所得税资产\租赁负债』首段与『租赁
                        #   负债』无包含 → 仍剔除，修复⑩语义保留）。
                        # ⚡⚡⚡ 2026-08-24 再修（ADZ 300101 蔡晓锋 等股东名叶子漏取）：
                        #   权益类叶子名=股东名（不含『实收资本』），但 l1_name=『实收
                        #   资本』（修复⑰ 附一级名）。首段无关但 l1 相关 → 保留（H 实收
                        #   资本 300301~333 股东名、ADZ 300101 蔡晓锋 均靠 l1 命中）。
                        _l1n = str(v.get('l1_name', '') or '')
                        # ⚡⚡⚡⚡ 2026-08-24 再修（AQ 22020201『工程』误配在建工程
                        #   根因）：『工程』⊂『在建工程』但『工程』是应付账款子目名
                        #   （泛化短词）。包含判定要求段长≥3（『股本』4001∈codes 不走
                        #   此判定，不受影响；『实收资本』4字⊂『实收资本(或股本)』保留；
                        #   『蔡晓锋』靠 l1『实收资本』保留）。
                        # ⚡⚡⚡⚡⚡ 2026-08-25 再修（AYL 实收资本 4105 只取到 20M 根因）：
                        #   AYL 股本代码=4105（非注册表默认 4001）→ 走此判定，各主体
                        #   叶子 l1_name=『股本』（2字）被 len>=3 排除 → 仅上海朗炫
                        #   （l1=『实收资本』4字）命中，其余 14 主体漏取。l1_name 是
                        #   一级科目名（父级行权威名），不应受段长限制；首段仍要求
                        #   len>=3 防『工程』等泛化短词误配。
                        def _rel(w):
                            return (w == subj_name or
                                    (len(w) >= 3 and (w in subj_name or subj_name in w)))
                        def _rel_l1(w):
                            return bool(w) and (w == subj_name or w in subj_name or subj_name in w)
                        if not _rel(_first_seg) and not _rel_l1(_l1n):
                            hit = False
                    if hit:
                        # ⚡⚡ 2026-08-24 修复（AQ 存货 1.387亿 漏取根因）：损益类
                        #   名称撞词行（640104 存货跌价准备/670104 存货跌价损失，
                        #   qm=0 的费用科目）名称含『存货』被 _name_hit 命中 → rows_n
                        #   非空 → _pick 偏 s_n=0 → 存货代码聚合被弃。损益类（6/7
                        #   开头）科目不并入资产类聚合（recon_self 本就不核对损益）。
                        if str(c)[:1] in '678':
                            continue
                        if not is_contra:
                            s_g_n += v['qm']; rows_g_n.append(str(c))
                        s_g_v += v['qm']; rows_g_v.append(str(c))
                    if codes and any(str(c).startswith(k) for k in codes):
                        # 2026-08-07 修复㉔：codes 撞车剔除——【前缀级】判定：
                        # 该 codes 前缀下【任一】行名含另一个注册表科目名（di 的 2601
                        # 在 FY=『租赁负债-未确认融资费用/租赁付款额』→ 属租赁负债；
                        # ncl 的 2401 在 FY=『递延收益』→ 属递延收益）→ 整前缀剔除。
                        # 不误伤：1002 银行存款/1403 原材料 非注册表科目名 → 保留。
                        _pc = min((k for k in codes if str(c).startswith(k)), key=len)
                        _prefix_ok = True
                        if _pc not in _prefix_cache:
                            _hit_other = False
                            for (_c2, _n2), _v2 in ent_tb.items():
                                if not str(_c2).startswith(_pc):
                                    continue
                                # ⚡⚡ 2026-08-25 修复（AYL 未分配利润 96M 根因）：
                                #   利润分配(4104)子目『利润分配-提取法定盈余公积』首段=
                                #   『利润分配』（分配动作，属本科目），原撞车检查因名称
                                #   含他科名『盈余公积』(surplus_res) 把整个 4104 前缀
                                #   剔除 → 代码聚合只剩 410407 未分配利润，漏提取盈余公积
                                #   净额 → 自建TB vs 底稿差 96M。利润分配族子目不参与撞车。
                                _n2s = str(_v2.get('name', ''))
                                _first2 = re.split(r'[\\/\-]', _n2s)[0].strip() if _n2s else ''
                                if _first2 == '利润分配' or _n2s.startswith('提取'):
                                    continue
                                if any(rn in _n2s or rn in str(_v2.get('l1_name', ''))
                                       for rn in _reg_other_names):
                                    _hit_other = True
                                    break
                            _prefix_cache[_pc] = not _hit_other
                        _prefix_ok = _prefix_cache[_pc]
                        if not _prefix_ok:
                            continue
                        # ⚡⚡ 2026-08-24 修复（AQ 使用权资产 SMART 1704 误配根因）：
                        #   registry codes 是多账套并集（rua 含 1704=FY 使用权资产，
                        #   但 AQ 1704=开发支出），主体无名称命中时代码兜底会把开发
                        #   支出并入（SMART 220.7M→使用权资产）。名称主键（铁律58）：
                        #   代码命中行若【名称与 l1 均与 subj_name 无关】且非备抵 →
                        #   属他科，剔除（开发支出 vs 使用权资产）。不误伤：FY 1704
                        #   名称含『使用权资产』保留；J 存货 1403/1405 l1=『存货』
                        #   保留；H 权益类叶子 l1=『股本』保留。
                        if not is_contra and not _name_relate(n, str(v.get('l1_name', '')), subj_name):
                            continue
                        if not is_contra:
                            s_c_n += v['qm']; rows_c_n.append(str(c))
                        s_c_v += v['qm']; rows_c_v.append(str(c))
                def _pick(s_n, s_c, rows_n, rows_c):
                    # 2026-08-07 修复⑲：名称命中（rows_n 非空）但合计=0（JTt 未分配利润
                    # 410408=0）不得回退代码聚合（4104 全子目含提取盈余公积/应付股利
                    # 24.3M）——"名称无命中"才用代码兜底。
                    # 2026-08-07 修复㉑：名称命中行与代码命中行【无交集】且代码聚合显著
                    # （J 存货叶子名=原材料/库存商品 无『存货』字样，仅 1409 跌价准备
                    # l1 含『存货』被名称命中 → 名称聚合≈0，代码聚合 74.6M 才对）→
                    # 代码优先（名称仅为辅助，代码才是权威控制数）。
                    # ⚡⚡ 2026-08-24 修复（AQ 使用权资产 1704 误并根因）：registry codes
                    #   是多账套并集（rua=['1704','1705','1681','1509','1510','1810','1812']），
                    #   AQ 的 1704=开发支出（技术开发费 47.3M），但 1641 使用权资产不在
                    #   codes 内 → 名称命中 1641、代码命中 1704，原 abs(s_c)>abs(s_n) 取
                    #   s_c 把开发支出并入（差 38.8M）。名称主键（铁律58）：名称命中行
                    #   【不在 codes 内】时名称聚合可信（AQ 1641 精确对应底稿），代码
                    #   聚合只在该行 ⊂ codes 时信任（J 存货 1409⊂codes、货币资金名称
                    #   行⊂codes 均保持原行为）。tb_recon_detail 已同规则处理使用权资产。
                    if not rows_n:
                        return s_c
                    if not rows_c:
                        return s_n
                    if len(codes) > 1:
                        if set(rows_n) <= set(rows_c):
                            return s_c
                        return s_n
                    if set(rows_n) & set(rows_c) or abs(s_c) > abs(s_n):
                        return s_c
                    return s_n
                gross = _pick(s_g_n, s_c_n, rows_g_n, rows_c_n)
                net = _pick(s_g_v, s_c_v, rows_g_v, rows_c_v)
                return gross, net

            # ⚡⚡ 2026-08-24 修复（ADF 集团核对符号口径）：集团稿审定数 = 【各主体 abs】之和
            #   （每主体审定表取正展示，用户确认『资产类负数=该科目含负债性质余额，各自科目反映』），
            #   原全集团带符号合并（净额）在主体方向混合时（部分主体借余/部分贷余）与审定数不等 →
            #   应收/应付等误报。改为逐主体取 abs 后再求和（与集团稿口径一致）。
            #   保留原修复⑧⑩的语义：待摊费用仅在少数主体等仍靠逐主体识别。
            # ⚡⚡⚡ 2026-08-25 再修（AYL 应交税费 3390万 误报根因）：不同账套集团稿
            #   符号口径不一——ADF 按各主体 abs 取正展示；AYL 应交税费集团审定表合计
            #   =【各主体带符号】求和（部分主体负/部分正，合计 7,518,619.51）。Σ abs
            #   在方向混合主体下必然虚高。改为双口径并存：分别算 abs 求和 与 带符号
            #   净额取正，与底稿审定数比更接近者（兼容两种集团稿形态）。
            sum_g_abs = sum_g_sig = sum_v_abs = sum_v_sig = 0.0
            for ent in sorted(ent_tb_all):
                _etb0 = ent_tb_all[ent].get(str(yy), {})
                if not _etb0:
                    continue
                # _tb_qm 期望 key=(code,name) 二元组（原 merged 结构）；单主体转同构
                _etb = {(str(c), str(v.get('name', ''))): v for c, v in _etb0.items()}
                _g, _v = _tb_qm(_etb)
                sum_g_abs += abs(_g); sum_g_sig += _g
                sum_v_abs += abs(_v); sum_v_sig += _v
            g2 = sum_g_abs if abs(sum_g_abs - abs(aud)) <= abs(abs(sum_g_sig) - abs(aud)) else abs(sum_g_sig)
            v2 = sum_v_abs if abs(sum_v_abs - abs(aud)) <= abs(abs(sum_v_sig) - abs(aud)) else abs(sum_v_sig)
            if abs(g2) <= 0.005 and abs(v2) <= 0.005:
                continue
            # 双口径取与底稿审定数更接近者（0.01 铁律）
            d_g = abs(g2 - abs(aud))
            d_v = abs(v2 - abs(aud))
            if min(d_g, d_v) > max(0.01, abs(aud) * 0.0001):
                tot_s = g2 if d_g < d_v else v2
                diffs.append((subj['name'], fk, '全集团', str(yy), round(tot_s, 2), round(aud, 2),
                              round(tot_s - aud, 2)))
    if diffs:
        print('  ⚠️ [阶段③ recon_self] 底稿审定数 vs 自建试算表差异（同源，差异=取数 bug）：')
        for name, fk, ent, yy, tb, aud, diff in diffs[:30]:
            print(f'     · {name}（{fk}）{yy} {ent}: 自建TB={tb:,.2f} 底稿审定={aud:,.2f} 差={diff:,.2f}')
        if len(diffs) > 30:
            print(f'     … 共 {len(diffs)} 项')
        return False, diffs
    print('  ✅ [阶段③ recon_self] 底稿审定数 = 自建试算表（同源 0 差异，取数无 bug）')
    return True, []


def main():
    dirs = sys.argv[1:]
    for d in dirs:
        ok, diffs = recon_self(d)
        print(f'{d}: {"OK" if ok else "差异 %d 项" % len(diffs)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
