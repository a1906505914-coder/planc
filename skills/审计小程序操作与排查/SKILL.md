---
name: 审计小程序操作与排查
description: 审计底稿自动化小程序（D:\底稿测试\账套取数审计小程序，201 个 Python 脚本）的操作与排查指南。当用户要求跑底稿/生成底稿/取数/科目核对/回归测试/排查底稿问题/运行审计小程序/调用审计程序时使用。包含唯一入口 run_audit.bat、run_u8_on_sap 参数、数据包要求、防回退验证链（回归测试+指纹+完整性）、常见问题排查、多人协作约束（Git 分支/并发锁/脱敏红线）。
---

# 审计小程序操作与排查

审计底稿自动化小程序：把账套原始数据（TB/GL）自动生成审计底稿的 Python 工具集（13 个科目生成器 + 专项核对程序 + 三件套质检）。

## 一、程序位置与入口

- 程序根目录：`D:\底稿测试\账套取数审计小程序`（成员机器上为各自 clone 的目录，下文以 `<APP>` 代指）。
- **唯一主入口 `run_u8_on_sap.py`**（SAP/U8 自适应）；集团并行走 `ah_parallel_run.py`。

| 场景 | 命令（在 `<APP>` 下） |
|---|---|
| AH 三集团全套 | `run_audit.bat run_all.py ah` |
| 单数据包全套 | `run_audit.bat run_all.py <数据包目录>` |
| 单模块回归 | `run_audit.bat regress_module.py <模块> [--allow 白名单]` |
| 专项调试 | `run_audit.bat <程序>.py <数据包目录>` |
| 集团往来重跑 | `python run_u8_on_sap.py --group --subj current_account` |
| 断点续跑 | `python resume_yy.py`；收口 `python finalize_yy.py` |
| 完整性检查 | `python wp_completeness_check.py` |

run_u8_on_sap 常用参数：`--subj bank,current_account`（只跑指定科目）、`--group`（集团模式）、`--out-dir <输出>`、`--skip-check`（跳过质检）、`--no-resume`（强制全量）。

**数据包要求**：SAP 账套含 `科目余额表/`（按期间列）+ 综合查询明细表 + 辅助核算余额表；U8 账套平铺每主体三件套；网银核对另需 `网银/` CSV 流水。数据路径由 `paths.py` 统一管理（环境变量 `AUDIT_DATA_ROOT` 可覆盖；账套映射见 `paths.ACCT_NAMES`）。

## 二、铁律（不可违反）

1. **一律经 `run_audit.bat` 运行**——它固定 `PYTHONHASHSEED=0` 保证行序稳定；裸 `python` 跑会导致每次重跑行序不同（看似"回退"的假象）。
2. **重跑期间不改代码**（生成器惰性 import，任何改动污染进行中的重跑）；全量只发生在收口。
3. **改公共层前先跑回归测试**：`python tests/test_architecture.py`（架构单测 31 项）+ `python tests/smoke_test.py`（冒烟 64 项），全绿再提交。
4. **数据安全红线**：所有数据进 AI 对话前必须本地脱敏（`desensitizer.py`/`mask_*`），真实主体名称永不进入对话；底稿交付前跑脱敏。
5. **发现问题如实标注**：数据/工具问题在底稿中标注（禁止"报正常"），差异项交人工复核，不自动归一化。
6. 运行前自检（preflight）BLOCK 则停止；同数据包被 `.audit.lock` 占用时等超时（30 分钟）或确认无人运行后处理。

## 三、验证链（改完代码/生成完底稿后的检查）

1. **回归测试**：`python tests/test_architecture.py` + `python tests/smoke_test.py`；
2. **指纹回归**：`working_paper_fingerprint.py base`（存基线）→ 改后 `check` 对比，`--allow` 白名单过滤预期差异；基线只由合并审查者更新；
3. **模块级回归**：`run_audit.bat regress_module.py <模块> --allow 白名单`；
4. **三件套审查**：`audit_checker.py` + `self_check.py` + `recon_self.py`（audit_checker ERROR 0 才交付）；
5. **完整性总表**：`wp_completeness_check.py` 一键扫全部底稿（每底稿一行：关键表数/空表/ERROR/WARN）；
6. **运行报告**：每次运行自动写 `run_history/run_{时间}.md`（可追溯哪个 commit + 哪批数据）。

## 四、常见问题排查

| 现象 | 原因 | 处理 |
|---|---|---|
| `PermissionError` / 底稿没更新 | 目标文件被 Excel/WPS 打开（`~$` 锁文件） | 关闭文件后重跑 |
| 底稿"整份空白"但 TB 有数据 | 取数未覆盖（如小额发生额科目） | 对照 `self_check` 提示，查模块取数边界 |
| 行序每次不同 | 未用 `run_audit.bat`（PYTHONHASHSEED 随机） | 统一 run_audit.bat；指纹工具行序无关 |
| XBJ 等大表 OOM | openpyxl 驻留全部 cell | 走 xw 流式渲染；`.cache` 过大可清 |
| 专项储备明细表 ≠ 计提核对表 | TB 缺期末日（7/31）过账凭证 | 企业重导科目余额表后重跑 |
| 同数据包被占用 | `.audit.lock`（他人正在跑） | 等 30 分钟超时或确认无人后删锁 |

## 五、多人协作约束（2026-08-30 起生效）

- 程序唯一真身在 Git bare 仓库；成员 `git clone` 各自副本，改完 push，**合并需审查**；
- 同数据包并发被 `.audit.lock` BLOCK；定时任务**单主负责制**；
- 缓存隔离：`AUDIT_CACHE_DIR`/`AUDIT_GL_CACHE_DIR` 指向本机本地（不写共享盘）；
- 详细规则见 `<APP>/docs/本机服务器接入指南.md` 与 `docs/审计小程序总览与团队协作说明.md` §6.5/6.6。
