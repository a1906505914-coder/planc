# 审计流水线扩展设计（预留框架，2026-08-24）

> 定位：本文档是【设计预留】，不是实现清单。按用户要求，考虑安全性暂不做完整闭环；
> 但把 ④⑤⑥ 阶段与零散外部核对程序的【框架/契约/接口】先定下来，后续实现按此落地，
> 避免继续"长出来"。

## 一、现状盘点（已调研）

### 1.1 已统一的（整合先例，共 4 个）
| 层 | 模块 | 说明 |
|---|---|---|
| 账套读取 | `ledger_backend.py` | LedgerBackend 接口 + U8/Sap300/AhBackend + 统一 read_tb/discover |
| 缓存 | `cache_manager.py` | GL/TB 缓存 key 版本化 + 清理 |
| 脱敏 | `desensitizer.py` | 统一 Desensitizer（confirm/contract 语境隔离）|
| 专项编排 | `special_reviews.py` | REGISTRY 插件式注册 6 专项 + gen_one 统一包装 |

### 1.2 已有的"中间层"预留
- `audit_result_export.py`：审定数 → 结构化 JSON（`adjustments` 参数已预留，`adj_structure` simple/split）
- `audit_protocol.py`：底稿行类型协议（合计/小计/段头识别）

### 1.3 未实现（regen_all 6 阶段中 3 个为空）
```
④ 重分类调整分录   reclass_gen   (False)
⑤ 报告生成          report_gen    (False)
⑥ 外部试算表核对    recon_ext     (False)
```

### 1.4 零散未收编（41 个外部数据/核对程序）
- 独立 `__main__` 运行、无统一总控：`bank_acct_recon` / `confirm_reconcile` / `cross_recon` / `aging_review` / `ah_reconcile` / `confirm_*` / `contract_*` / `currency_*` / `baddebt_*` 等

## 二、设计原则

1. **协议先行，实现后补**：先把"输入/输出契约"用文档+接口定死，实现按契约做。
2. **注册表模式**（复用 special_reviews 先例）：所有专项/核对程序登记到统一 REGISTRY，总控只认注册表。
3. **数据不落地 xlsx 解析**：跨阶段传递用 JSON（audit_result_export 已示范），xlsx 仅作最终交付。
4. **阶段可独立运行**：④⑤⑥ 各自可单独 `--phase` 跑，阶段间通过 JSON 契约衔接。
5. **安全性**：调整分录/报告涉及审计判断 → 默认只输出"建议"，不自动改写底稿；需人工确认。

## 三、目标架构

```
regen_all（6 阶段总控）
  ├─ ① self_tb_gen        ✅ 自建试算表 → 试算表.xlsx + 试算表.json
  ├─ ② 13 builder         ✅ 底稿 xlsx + 审定数沉淀（audit_result_export）
  ├─ ③ recon_self         ✅ 底稿 vs 试算表 同源核对 → 核对结果.json
  ├─ ④ reclass_gen        🔲 重分类调整（设计见 §四）
  ├─ ⑤ report_gen         🔲 报告生成（母分汇总/合并，设计见 §五）
  └─ ⑥ recon_ext          🔲 外部数据核对（设计见 §六）

专项核对层（收编 41 个零散程序）
  └─ special_reviews.REGISTRY 扩展 → 银行/函证/合同/账龄/跨账套/汇率 …
      每专项 = (module, manifest, subject, split, accts)
      输出 = 专项底稿.xlsx + 差异.json（统一差异契约）
```

## 四、④ 重分类调整（reclass_gen）契约

### 4.1 输入
- `试算表.json`（①产物）
- 各科目审定数 JSON（②产物）
- `重分类建议.json`（人工填写，格式见下）

### 4.2 重分类建议 JSON 契约
```json
{
  "version": 1,
  "entity": "3300",
  "year": "2026",
  "adjustments": [
    {
      "id": "REC-001",
      "reason": "预收账款贷方余额重分类至合同负债",
      "status": "draft",           // draft | confirmed | rejected
      "entries": [
        {"code": "2203", "name": "预收账款", "dr": 0, "cr": 123456.78, "currency": "CNY"},
        {"code": "2241", "name": "合同负债", "dr": 123456.78, "cr": 0, "currency": "CNY"}
      ]
    }
  ]
}
```

### 4.3 输出
- `重分类调整_2026.xlsx`（按实体/科目汇总）
- 经确认的调整 → 追加写入 `审计结果/*.json` 的 `adj` 字段（audit_result_export 已预留）

## 五、⑤ 报告生成（report_gen）契约

### 5.1 输入
- 审定数 JSON（②产物，已含 adj）
- `报告配置.json`（报表行→科目映射，母分/合并规则）

### 5.2 输出
- `合并试算表_2026.xlsx`（母+分 逐主体列 + 合并抵消列）
- `试算表核对_2026.xlsx`（审定数 vs 外部试算表）
- 注：合并抵消分录（内部往来/购销）属于审计判断，预留 `offset.json` 人工维护

## 六、⑥ 外部数据核对（recon_ext）契约

### 6.1 统一差异 JSON 契约（所有核对程序输出对齐）
```json
{
  "recon": "bank_acct_recon",       // 来源程序
  "acct": "AYL", "year": "2025",
  "subject": "银行存款",
  "run_at": "2026-08-24T10:00:00",
  "diffs": [
    {
      "entity": "AYL", "code": "1002", "name": "银行存款",
      "gl_amt": 123456.78, "ext_amt": 123000.00, "diff": 456.78,
      "reason": "待查",              // 人工确认
      "evidence": "银行对账单2025年7月",  // 外部证据引用
      "status": "open"               // open | explained | adjusted
    }
  ]
}
```

### 6.2 零散程序收编清单（第一阶段只注册不重写）
| 程序 | 专项 | 目标账套 | 状态 |
|---|---|---|---|
| bank_acct_recon / bank_review | 银行 | ADF/GJX/AYL | ✅ 已收编（开户核对+网银双向对账差异清单）|
| confirm_reconcile / confirm_parse_all | 函证 | AZ | ✅ 已收编（往来函证差异/不符标记）|
| contract_summary / contract_anomaly | 合同 | ADF | 🔲 数据源维度（data_source_registry）|
| aging_review / aging_manifest | 账龄 | 集团 | ✅ 已收编（与TB核对差异）|
| cross_recon | 科目勾稽 | 集团 | ✅ 已收编（勾稽差异列）|
| ah_reconcile / ah_intra_* | AH 关联往来 | AH | 🔲 AH 专用三步专项（暂不并入）|
| currency_split / currency_annotate | 多币种 | AZ/ADF | 🔲 数据源维度（preprocess）|
| baddebt_common / co_carryover_stats | 备抵/CO | ADF | 🔲 专项分析（analyze）|

### 6.3 收编方式（不重写逻辑，已实现）
1. 每个专项在 `recon_registry.py` 登记 diff_sheets 抽取规则（文件 glob + sheet + 差异标记列/数值列）
2. 通用抽取器 `recon_diff_extract.py` 从专项底稿 xlsx 抽差异 → 统一 JSON
3. `recon_ext.py` 总控：按账套过滤适用专项 → 汇总差异 JSON → 输出汇总表
4. `data_source_registry.py`：数据获取/预处理/专项分析类程序（合同OCR/函证OCR/多币种拆分）统一登记
5. `data_pipeline.py`：数据获取→核对→汇总 完整流水线总控

**已收编核对专项（4）**：bank / aging / cross / confirm —— 实测 AYL 银行 124 条、GJX 网银 12 条、XBJ 账龄 1 条、AZ 函证 35 条、ACB/AFJ 勾稽 104/144 条。

**已登记数据源（7）**：currency_split(preprocess) / contract_ocr / contract_summary / confirm_ocr / confirm_parse / bank_stmt_parse(extract) / contract_anomaly(analyze)。


## 七、实施路线（按用户节奏，不紧急）

| 步骤 | 内容 | 风险 | 建议时机 |
|---|---|---|---|
| R1 | 建 `recon_registry.py` 注册表骨架 + `recon_ext.py` 总控骨架（先注册 2-3 个已成熟专项验证）| 低 | 近期 |
| R2 | 重分类建议 JSON 契约落地：写 `reclass_gen.py` 骨架（读建议 JSON → 出调整表，不自动改底稿）| 低 | 中期 |
| R3 | 统一差异 JSON 契约：给 2-3 个核对程序加 `recon_entry()` 包装 | 中 | 中期 |
| R4 | report_gen 骨架（审定数 → 报告，合并规则留 offset.json）| 中 | 后期 |
| R5 | 收编剩余零散程序 | 低（逐个）| 随用随收 |

## 八、安全边界（为什么现在不做完整闭环）

- 审计调整分录涉及专业判断 → 程序只产"建议"、人工确认后才生效（status 字段）
- 报告合并抵消是审计判断核心 → 只读审定数 JSON，不自动生成抵消分录
- 外部证据（对账单/回函/合同）→ 只做核对引用，不修改底稿审定数
- 后续若要做完整闭环，契约已就位：JSON 接口不变，只需把 status 流转接入人工审批

---

**结论**：框架已定（本文档 + §七路线），代码侧从 R1（recon_registry 骨架）开始。
