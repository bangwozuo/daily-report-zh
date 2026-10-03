# 经营日报员

> **每天打烊后一句话告诉老板"今天生意怎么样、明天注意什么"**

![演示](docs/demo.mp4)

*上面是本仓 5 个代表资产的真实执行录屏（数据采集 → 接入体检 → 异常检测 → 日报生成 → 异常预警，每个镜头均为脚本实跑截图）。*

[![Stage](https://img.shields.io/badge/stage-P0-orange)](https://github.com/bangwozuo)
[![Asset](https://img.shields.io/badge/asset-prompt%20%2B%20script-blueviolet)](#资产形态)
[![NoKey](https://img.shields.io/badge/API%20Key-not%20required-success)](#资产形态)
[![License](https://img.shields.io/badge/license-Apache--2.0-green)](LICENSE)

---

## 它是谁

面向 **本地生活商家** 的数字员工资产包。

| 项目 | 内容 |
|------|------|
| 目标用户 | 全部门店（受众最广） |
| 交付物 | 日报推送准时率 100%（每晚 22:00）；老板阅读耗时 ≤1 分钟；异常预警召回率 |
| 技能数 | 3 |
| 工作流数 | 4 |
| 旧名存档 | `门店经营日报分析师` |

### 数字员工总览

| 项 | 内容 |
|---|---|
| **身份** | 经营日报员——面向餐饮/零售/美业等线下门店，每天 21:30 采集数据、22:00 出日报、越界即预警、周一给方向 |
| **能力边界** | 做：数据体检归集、异常判定、日报文案、预警编排、周度建议；不做：财务记账、投资级建议、资金操作、替代店主决策 |
| **工作方式** | 判定归脚本规则引擎（R1 阈值 / R2 离群 / R3 趋势 / R4 同比），表达归模型；所有对外输出保留 🔒 人工确认 |
| **KPI** | 日报 22:00 准时推送；老板阅读 ≤1 分钟（结论+3 个数+1 个行动）；异常早一天发现 = 少一天失血 |
| **合规底线** | 只读不写、最小必要、脱敏先拦后取；数据只走官方 API 或用户导出，不爬取私域数据 |

---

## 资产矩阵

| 名称 | 一句话 | 类型 | README |
|---|---|---|---|
| 门店数据采集 | 打烊后的数据闸门：合规、完整、干净，BLOCK 时下游不得启动 | 工作流 · T3 | [README](workflows/store-data-collect-flow/README.md) |
| 门店数据对接 | 数据接入体检：白名单/黑名单 + 字段完整率 + 6 类个人信息脱敏 | 技能 · T4 | [README](skills/store-data-sync/README.md) |
| 异常检测 | 4 条量化规则 + 5 类假阳性拦截，24 条预警条条带排查方向 | 技能 · T1 | [README](skills/anomaly-detect/README.md) |
| 日报生成 | 22:00 出日报：一句话结论 + 三个关键数 + 一个行动建议 | 工作流 · T3 | [README](workflows/daily-report-generate-flow/README.md) |
| 日报文案生成 | 数据转人话：正文 ≤ 200 字，关键数 ≤ 3，结论含方向+幅度+归因 | 技能 · T1 | [README](skills/daily-report-copy/README.md) |
| 异常预警 | 筛·排·留：级别过滤 + 免打扰 + 30 分钟升级，脚本不自动发送 | 工作流 · T3 | [README](workflows/anomaly-alert-flow/README.md) |
| 周度经营建议 | 每周一只给一个重点动作：能落地、零预算、一周见反馈 | 工作流 · T3 | [README](workflows/weekly-ops-advice-flow/README.md) |

---

## 资产形态

**提示词 + 可选脚本** —— 这是理解本仓库的关键：

| 特性 | 说明 |
|------|------|
| ✅ 无需 API Key | 提示词模式不需要任何密钥，不调用模型 |
| ✅ 双模式 | 纯提示词粘贴即用；各资产另配确定性 Python 脚本（`--demo` / `--input`），算派生指标、跑判定规则、落盘 Excel/Word/PNG |
| ✅ 零部署 | 没有服务端；脚本模式需 `pip install -r requirements.txt` |
| ✅ 平台无关 | Coze / WorkBuddy / Dify / Claude / ChatGPT 均可 |
| ✅ 用户自备算力 | 模型来自你自己的订阅 |

---

## 快速开始

```text
1. 打开 skills/store-data-sync/prompt.txt
2. 全文复制
3. 粘贴到你常用的 AI 工具（Coze / WorkBuddy / Dify / Claude / ChatGPT）
4. 按 SKILL.md 的输入规格提供数据
```

就这四步。完整指引见 [使用手册](docs/04-usage.md)。

---

## 仓库结构

```text
daily-report-zh/
├── README.md / employee.md / package.yaml     # 入口与 12 字段定义卡
├── docs/01~07                                 # 员工级文档（架构/流程/场景/手册/示例/录像/测试）
├── skills/                                    # 3 个原子技能
│   └── <skill>/
│       ├── README.md  SKILL.md  prompt.txt  schema.json  examples/
│       └── docs/                              # 该技能自己的 10 项文档 + 配图
├── workflows/                                 # 4 条工作流（复合技能）
│   └── <workflow>/
│       ├── README.md  SKILL.md  prompt.txt  schema.json  examples/
│       └── docs/                              # 该工作流自己的 10 项文档 + 配图
├── knowledge/                                 # RAG wiki 知识库
│   ├── README.md  RAG-接入指南.md  template.md
│   └── wiki/(index.md, _template.md, entries/)
├── connectors/                                # 连接器说明 + 合规红线
├── quality/                                   # 效果基线与追踪日志
└── tests/                                     # 资产校验测试（离线，无需密钥）
```

### 每个技能 / 工作流自带的 docs

| 文档 | 内容 |
|------|------|
| `README.md` | 资产速览与快速开始 |
| `docs/01-usage-manual.md` | 安装使用手册 |
| `docs/02-architecture.md` | 业务架构图 |
| `docs/03-flow.md` | 流程图（Mermaid + 配图） |
| `docs/04-examples.md` | 使用示例 |
| `docs/05-media.md` | 截图和录屏（清单 + 分镜脚本） |
| `docs/06-scenarios.md` | 使用场景（适用 / 不适用） |
| `docs/07-audience.md` | 用户群体 |
| `docs/08-value.md` | 解决问题与价值 |
| `docs/09-test-report.md` | 测试报告 |
| `docs/assets/overview.svg` | 自动生成的流程示意图 |

---

## 交付物导航

| 文档 | 内容 |
|------|------|
| [业务架构](docs/01-architecture.md) | 四层架构 + 数据流 + 能力边界 |
| [工作流流程](docs/02-workflow.md) | 4 条工作流的 DAG 可视化 |
| [使用场景](docs/03-scenarios.md) | 3 个真实场景（含前后对比） |
| [使用手册](docs/04-usage.md) | 各平台导入指引 + 常见问题 |
| [示例库](docs/05-examples.md) | 3 组输入输出示例 |
| [录像脚本](docs/06-recording-script.md) | 7 镜头分镜 + 旁白稿 |
| [校验报告](docs/07-test-report.md) | 资产质量校验结果 |

---

## 技能清单（3 个）

| # | 技能 | 能力族 | 复杂度 | 提示词 | 文档 |
|---|------|--------|--------|--------|------|
| 1 | 门店数据对接 | 文案生成 | `M` | [prompt.txt](skills/store-data-sync/prompt.txt) | [docs](skills/store-data-sync/docs/) |
| 2 | 日报文案生成 | 上架优化 | `S` | [prompt.txt](skills/daily-report-copy/prompt.txt) | [docs](skills/daily-report-copy/docs/) |
| 3 | 异常检测 | 文案生成 | `S` | [prompt.txt](skills/anomaly-detect/prompt.txt) | [docs](skills/anomaly-detect/docs/) |

## 工作流清单（4 条）

| # | 工作流 | 阶段 | 复杂度 | 触发 | 定义 | 文档 |
|---|--------|------|--------|------|------|------|
| 1 | 门店数据采集 | `P0` | `M` | 定时（每日 21:30） | [SKILL.md](workflows/store-data-collect-flow/SKILL.md) | [docs](workflows/store-data-collect-flow/docs/) |
| 2 | 日报生成 | `P0` | `S` | 定时（每日 22:00） | [SKILL.md](workflows/daily-report-generate-flow/SKILL.md) | [docs](workflows/daily-report-generate-flow/docs/) |
| 3 | 异常预警 | `P0` | `S` | 事件（指标越界） | [SKILL.md](workflows/anomaly-alert-flow/SKILL.md) | [docs](workflows/anomaly-alert-flow/docs/) |
| 4 | 周度经营建议 | `P1` | `S` | 定时（每周一） | [SKILL.md](workflows/weekly-ops-advice-flow/SKILL.md) | [docs](workflows/weekly-ops-advice-flow/docs/) |

---

## 知识库与连接器

| 目录 | 说明 |
|------|------|
| [`knowledge/`](knowledge/README.md) | RAG wiki 知识库：填入业务信息可显著提升输出质量 |
| [`connectors/`](connectors/README.md) | 连接器说明：数据从哪来、怎么合规地来 |

---

## 资产校验

```bash
pip install -r requirements.txt
pytest tests/ -v
```

校验技能完整性、提示词结构、契约一致性、工作流 DAG、技能级与工作流级 docs 完整性、知识库 wiki 与连接器结构。
**不需要任何 API Key。**

---

## 合规声明

- ✅ 所有输出为 **AI 辅助生成**，交付前须人工审核
- ✅ 提示词内置**违禁词禁止清单**，符合《广告法》要求
- ✅ 遵循《人工智能生成合成内容标识办法》
- ✅ 连接器只走**官方 API** 或**用户导出数据**
- ✅ 所有对外发布动作**保留人工确认环节**

---

## 许可

[Apache-2.0](LICENSE) — 可自由使用、修改、商用

---

*由 bangwozuo 业务库自动生成 · 2026-09-29*
