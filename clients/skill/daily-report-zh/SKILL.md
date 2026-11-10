---
name: daily-report-zh
description: 帮我做经营日报员。读取门店经营资料，检查来源、归集数据、检测异常、生成日报和周建议，或回顾原办理事项；调用已安装数字员工并返回可编辑文件。
---

# 经营日报员

调用已安装员工的原生工作流，不在本地另写计算、调度或模型循环。当前默认资料为明确标记的合成门店数据；实际客户资料须先按原入口导入并传入准确来源引用。

1. 用 `python scripts/product.py contract --connection <连接文件>` 核对安装、输入规格和可用动作。连接只允许当前本地服务；安装变化即停止，不自行换版本。
2. 按用户任务选择 `sync` 来源体检、`collect` 数据归集、`detect` 异常检测、`copy` 日报文案、`daily` 完整日报、`alert` 提醒草稿或 `weekly` 周建议。将 `operation` 和已提供的 `source_ref`、`metrics_ref` 放入输入文件的 `fields`；原任务写在 `text`。缺资料时展示缺项，不补造经营数字。
3. 新事项保存稳定操作编号，只提交一次 `submit --operation <原编号> --input <文件>`。丢响应或状态不明时用 `lookup` 查同一编号，不换号重发。
4. 展示实际来源、缺口、文件和草稿。用 `review --decision revise --expected-revision <修订> --notes <意见>` 保存修改意见；用户已明确同意本版结果时可 `confirm`，明确退回则 `reject`。确认后回读原事项及其原生运行，原稿保持。

脚本参数均需 `--connection <连接文件>`。该文件不含凭证，不读取其他应用 Cookie。文件链接与本地成果确认不代表外发、门店后台写入、客户验收、经营增收或消息送达。触发计划和提醒箱沿用员工实际服务，主Skill不启动新定时器。
