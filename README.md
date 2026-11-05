# 实盘网格指令生成项目（多产品版）

基金经理组合指令 → 可批量导入 Oplus 的 A股/港股网格买卖清单。一次运行覆盖全部产品，每个产品/单元独立计算仓位和净值。

## 核心特性

- **多产品**：一次覆盖全部产品（当前 26 个产品、55 个产品/单元组合），每档网格标注所属产品/单元
- **真实持仓基线（推荐日常用法）**：每天读最新一份《全球产品持仓查询》的真实仓位，仓位 100% 准确，自动覆盖手动交易、货币基金利息、融券、美股 swap；实测 1014/1015 全产品仓位 0 差异
- **swap/融券安全规则**：只做美股/场外互换的 swap 单元不生成 A股/港股网格；纯融券空头不生成多头买入网格
- **人工手改记忆**：在剔除版指令里做的人工修改会被 `manual_patch.py` 记住，下次自动恢复
- **流动性模拟**：用昨日分钟行情模拟每档网格的可成交性，输出可模拟成交清单（3 个子表）
- **净值推断**：从持仓文件逐持仓反推每个产品净资产
- **无冻结窗格**：所有导出 Excel 均不设冻结窗格
- **日期重叠保护**：当持仓已包含当天 oplus 成交、却仍要叠加时，程序会打印醒目提醒

## 目录结构

```
live_grid_project/
├── 日常代码/                      ★ 入口脚本（每天跑的就是这几个）
│   ├── import_daily_data.py       ① 一键归位当天新文件
│   ├── step1_filter.py            ② 剔除已成交 + 恢复历史手改
│   └── step2_grid.py              ④ 学习手改 + 跑网格（默认±30%价格带）
├── _待导入/                      收件箱：每天新文件丢这里
├── live_grid_simulator_sim.py    网格主程序（多产品版）
├── filter_completed_instructions_hk_nosuffix_fixed.py  底层剔除脚本
├── manual_patch.py               手改记忆脚本
├── data_loader.py                Wind 行情加载
├── instruction_reader.py         指令文件解析
├── grid_engine.py / features_daily.py / features_minute.py / gi_score.py / core_grid_weights.py / config.py / utils.py   核心计算模块
├── oplus文件/                    Oplus 成交历史（oplusMMDD.xlsx）
├── 指令/                         经理指令 + 剔除产物 + 人工补丁库
├── 持仓数据/                      全球产品持仓查询文件（仓位权威）
├── state/                        运行状态（审计/去重，固定基线不依赖它算仓位）
├── outputs_live_grid/            最终输出（买卖清单、复核报告）
├── 批量埋单文档/                  Oplus 接口规范参考
├── a股买卖清单导入模板.xlsx / 港股买卖清单导入模板.xlsx
├── 每日操作SOP.md                 ★ 每天照着做
└── requirements.txt
```

## 每日运行流程（四步）

```bat
cd /d C:\Users\gucao\Desktop\live_grid_project

:: 1) 当天新文件（oplus、最新全球持仓、新指令）丢进 _待导入\，自动归位
python 日常代码/import_daily_data.py

:: 2) 剔除已成交 + 自动恢复历史手改
python 日常代码/step1_filter.py

:: 3) 打开 指令\ 下最新的 *_剔除已成交.xlsx，人工核对增删指令，保存（只改这个xlsx）

:: 4) 学习手改 + 生成全产品网格（含净值 + 流动性模拟）
python 日常代码/step2_grid.py --asof-date 2026-10-15 --oplus-date 1014 --position-file "持仓数据/全球产品持仓查询_20261014.xls" --instruction-file "指令/组合指令1013_剔除已成交.xlsx" --no-apply-oplus-trades
```

> `step2_grid.py` 默认加 `--price-band 0.30`（参考价±30% 内的档位才输出），不用手写。

### 每次要改的参数

| 参数 | 填什么 |
|------|--------|
| `--asof-date` | 要跑哪天的网格（盘前跑填当天） |
| `--position-file` | **前一交易日收盘后**的最新全球持仓（仓位权威，以文件内部日期为准） |
| `--oplus-date` | 前一交易日（step1 剔除用；盘前第4步因已加 `--no-apply` 不叠加） |
| `--instruction-file` | step1 生成的最新剔除版指令 |
| `--no-apply-oplus-trades` | **盘前必加**：持仓已含前一天成交，不再叠加 oplus |

## 为什么日常用真实持仓基线，而不是纯增量

纯增量（`--incremental-state-mode`，沿用 state 只加 oplus）在股数层面融券 0 差异，但有三类结构性偏差：
1. **比例漂移**：仓位比例不随价格/净值重算，如振华股份真实 11.17%、纯增量只算出 9.7%；
2. **融券被截断**：融券负债在持仓里「市值/净资产」=0，会被 clip 成 0；
3. **非网格手动交易跟踪不到**：海通单元清仓、海外/广发 swap 美股、货币基金利息等在 oplus 里没有股票成交记录。

直接读每天最新真实持仓，这些问题全部自动消除，且与「每天本来就要下载持仓」的操作一致。`--incremental-state-mode` 仅保留作特殊补算兜底。

## 常用功能命令

以下 `...` 代表上面第 4 步的固定参数（含 `--no-apply-oplus-trades`）。

### 只跑单个产品（不需另存指令文件）

```bat
python 日常代码/step2_grid.py ... --product "净值线3号"
```
多个产品逗号分隔：`--product "净值线3号,申新8号"`。

### 快速版（不要流动性模拟）

```bat
python 日常代码/step2_grid.py ... --disable-liquidity-simulation
```

### 指令没变，只刷新行情

```bat
python 日常代码/step2_grid.py ... --full-refresh
```

### 特殊补算（持仓日早于目标日、中间有 oplus 成交）

去掉 `--no-apply-oplus-trades`，并须用持仓日当天 oplus 先建快照、再只加后续增量，避免跨天累计指令重复。日常盘前用不到。

## 输出文件

`outputs_live_grid/` 下：

| 文件 | 用途 |
|------|------|
| `实盘网格_A股/港股_本次新增网格_买卖清单导入.xlsx` | **主要交付**：仅本次新增档位，导入 Oplus |
| `实盘网格_A股/港股_当前全部网格_买卖清单导入.xlsx` | 当前全部在挂网格 |
| `实盘网格_复核报告.xlsx` | 13 个 sheet：run_meta、product_navs、position_check、current_positions、oplus 增量/累计、流动性、errors |
| `实盘网格_昨日分钟流动性模拟.xlsx` | 3 个子表：可模拟成交网格、逐档成交核实、分钟流动性概览 |
| `实盘网格_基金经理格式_当前全部网格.xlsx` | 按产品/单元/股票分组的经理格式 |
| `实盘网格_A股/港股_昨日分钟模拟成交_可导入.xlsx` | 昨日分钟模拟成交清单 |

## 净值计算

持仓文件无直接「产品净资产」列，程序用 **持仓市值（人民币）÷ 市值/净资产比例** 逐持仓反推，产品内取中位数，结果写入复核报告 `product_navs`（当前 25/26 个产品可推断）。

- 优先使用「持仓市值（人民币）」列，与净资产分母币种一致，避免港币/美元错配
- 剔除占比低于 0.2% 的仓位（比例列仅精确到 0.01%，极小仓位反推失真）
- 两融账户的融券盯市口径 Wind 与系统不一致，故净值一律以持仓文件推断，不自行逐日盯市

## 持仓表多空处理（两融账户）

- 持仓表「多空标志」取值：多仓 / 融券负债 / 空仓。
- `current_pos` 只表示多仓占净资产比例（多头网格用）；融券负债规模记在 `short_qty`（融券股数），多空分开。
- 同一代码同时有多头和融券时，多头网格仍按真实多仓比例生成。
- 纯融券（无多仓）的代码不生成多头买入网格，跳过原因写入复核报告 `position_check`。

## state 目录

- `current_positions.csv`：本次仓位快照（含 `short_qty`），固定基线模式下仅用于审计/查看，不作为下次输入
- `active_grid_master.xlsx`：当前完整网格主表
- `delta_ledger.xlsx`：控制「本次新增网格」去重
- `oplus_trade_state.csv`：oplus 已应用进度（仅增量模式用）
- `instruction_snapshot.csv`：指令摘要

## 环境要求

- Anaconda Python（`D:\anaconda3\python.exe`，含 pywin32 + WindPy）
- 安装依赖：`pip install -r requirements.txt`
- 运行时需连接 Wind 终端；跑批前关闭所有 Excel

## 关键参数速查

| 参数 | 作用 |
|------|------|
| `--no-apply-oplus-trades` | 不叠加 oplus 成交（**盘前日常必加**） |
| `--incremental-state-mode` | 旧的纯增量模式（日常不推荐，特殊补算用） |
| `--full-refresh` | 全量刷新所有股票行情 |
| `--disable-liquidity-simulation` | 禁用分钟级流动性模拟 |
| `--product "名称"` | 只跑指定产品（逗号分隔） |
| `--reset-position-state` / `--reset-oplus-state` | 重建仓位/oplus 基线（首次或重建用） |
| `--price-band 0.30` | 价格带±30%（step2_grid 默认已加） |
| `--liquidity-participation 0.01` | 流动性参与率 1%（默认） |
