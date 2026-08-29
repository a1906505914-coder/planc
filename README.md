# 审计底稿自动化小程序 — 操作手册

> 本文档面向使用/维护本系统的审计人员与开发者。覆盖：运行方式、数据包要求、防回退工具、常见问题。

## 一、运行方式（三选一）

| 方式 | 命令 | 适用 |
|---|---|---|
| **一键全套** | `run_audit.bat run_all.py ah`（AH 三集团并行）<br>`run_audit.bat run_all.py <数据包目录>`（单包） | 日常全套生成 |
| **单模块重跑** | `run_audit.bat regress_module.py <模块名> [--allow 白名单]` | 改代码后快速回归 |
| **单程序直跑** | `run_audit.bat <程序>.py <数据包目录>` | 专项调试 |

> ⚠️ **建议一律经 `run_audit.bat` 运行**：它固定 `PYTHONHASHSEED=0`，保证生成器行序稳定（裸 `python` 跑会导致每次重跑行序不同，造成"看似回退"的困扰）。

## 二、数据包要求

- **SAP 账套（AH 形态）**：目录含 `科目余额表/`（按期间列）、`综合查询明细表`、`辅助核算余额表`。
- **U8 账套（XBJ/ADF 形态）**：目录内平铺 每主体三件套（科目余额表 / 综合查询明细表 / 辅助核算余额表）。
- 网银核对另需 `网银/` 目录下的各银行 CSV 流水。

## 三、核心模块入口（13 个主模块 + 专项）

| 模块键 | 程序 | 产出 |
|---|---|---|
| current_account | `current_account_detail.py` | 往来 8 科目底稿（含对方科目核对、WBS 回款追踪）|
| bank | `bank_deposit_detail.py` / `bank_reconcile_detail.py` | 银行存款、网银双向核对 |
| tax / payroll / expense / pl | 对应 `*_detail.py` | 应交税费 / 职工薪酬 / 费用 / 损益 |
| inventory / longterm / loan / equity / rd_expense / revenue | 对应模块 | 存货 / 长期资产 / 借款 / 权益 / 研发 / 收入 |
| 对方科目核对（#875） | `counterparty_recon.py` | 46 科目对方科目核对（已模块内集成注入各底稿）|

专项：`ah_intra_*`（88 家关联往来）、`aging_review.py`（账龄复核）、`confirm_parser.py`（回函）、`fa_movement.py`（固定资产）、`lease_review.py`（租赁复核）等。

## 四、防回退工具链（改代码前必读）

1. **git**：小程序目录已纳入 git。改代码前 `git status` 看改动，误改 `git checkout -- <文件>` 回滚。
2. **指纹基线**：`python working_paper_fingerprint.py base` 存当前底稿指纹；改代码重跑后 `check` 对比，报任何内容差异。
3. **预期差异白名单**：`check --allow <白名单文件>` 自动忽略预期差异（如"新增对方科目核对 sheet"）。
4. **模块级回归**：`run_audit.bat regress_module.py <模块> --allow 白名单` —— 改某模块只重跑该模块 + 指纹对比。
5. **底稿归档**：`python archive_wp.py [--zip]` 按批次归档正式底稿，回退可取回。
6. **年度切换**：`audit_config.py` 集中配置 YEAR，`python audit_config.py scan` 扫描剩余硬编码。
7. **断点续跑（默认开）**：大任务（AH 全量 50 分钟）中途崩溃后，重跑自动跳过「已完成且签名未变」的科目（数据根目录 `.resume_{主体}.jsonl` 记录）。`--no-resume` 强制全量；改代码或重导数据后签名变化 → 自动全量（无需手动清理）。
8. **运行报告**：每次运行自动写 `run_history/run_{时间}.md`（数据包/参数/git commit/代码签名/OK·SKIP·FAIL 明细/失败 traceback）——任何一份底稿都能追溯「哪个版本代码+哪批数据生成的」。
9. **运行前自检（preflight）**：`run_all.py` 跑前自动检查磁盘空间/文件占用/并发/数据完整性，BLOCK 阻断、WARN 提示（`--skip-preflight` 跳过）。
10. **看门狗（卡死检测）**：长任务运行时另开终端 `python watchdog_check.py <数据目录>`，心跳停滞超 30 分钟即告警并写 `.ALERT_*` 标记（preflight 会提示处理）；`--clean` 清理标记。
11. **底稿完整性检查（替代逐张点开）**：`python wp_completeness_check.py` 一键扫全部底稿，产出**完整性总表**（每底稿一行：关键表数/空关键表/ERROR/WARN/状态）+ 异常明细，一眼看出哪些底稿有问题。`run_all.py` 已自动集成（生成后自动出报告到 `completeness_reports/`）。

## 五、常见问题

| 现象 | 原因 | 处理 |
|---|---|---|
| 重跑报 `PermissionError` / 底稿没更新 | 目标文件被 Excel/WPS 打开（`~$` 锁文件） | 关闭文件后重跑；程序已自动跳过并提示 |
| 底稿"整份空白"但 TB 有数据 | 取数未覆盖（如 2430 长期待摊小额发生额） | 对照 `self_check` 提示定位，检查对应模块取数边界 |
| 行序每次不同 | 未用 `run_audit.bat`（PYTHONHASHSEED 随机） | 统一用 run_audit.bat；指纹工具已行序无关不误报 |
| XBJ 等大表 OOM | openpyxl 驻留全部 cell | current_account 已 xw 流式化（#876），大表不再 OOM |
| 专项储备明细表 ≠ 计提核对表 | TB 缺期末日（7/31）过账凭证 | 企业重导科目余额表后重跑对齐 |

## 六、开发约定

- 主流程入口统一走 `run_u8_on_sap.py`（SAP/U8 自适应）；集团并行走 `ah_parallel_run.py`。
- 保存统一走 `audit_common._safe_save`（锁感知）；搬移统一走 `audit_common.move_if_free`。
- 写表优先用 `xw_render.py`（xlsxwriter 渲染层），避免直接 openpyxl 逐 cell 大表。
- 数据源规则（科目码/符号/功能范围/WBS）配置在 `account_profiles` / `sap_subject_map`，新账套先核对配置再跑。
