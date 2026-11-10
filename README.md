# 经营日报员 · 原生运行接入版 0.2.0-rc.1

读取获权经营导出，核对来源、归集指标、检测异常，生成可编辑日报和周建议。

本版实际接入 DeepAgents 的既有原生 LangGraph 工作流，保留3项方法Skill、4项工作流Skill和8项知识义务。主Skill已在Codex安装、被官方发现并实际调用；Word、Excel、PNG和JSON已生成、核对、确认，原事项和原稿保留。

这是连接已有部署的预发布候选。当前默认资料为明确标记的合成门店数据；没有连接门店后台、自动外发渠道或客户真实经营资料。当前源框架为 source-quote-compatible，实际环境DeepAgents 0.7.28；不将其资格冒称为另一个公共框架发行版的验收。

## 使用

1. 将 `clients/skill/daily-report-zh` 作为主Skill安装到当前支持的宿主。
2. 服务部署者提供准确连接文件，包含原地址、项目、Agent、安装ID和清单SHA256。仓内连接样例只适用于本机当前原部署，没有账户凭证。
3. 用 `scripts/product.py contract --connection <文件>` 核对版本；按SKILL说明提交一次。丢响应只查询原编号。
4. 核对当前草稿及文件。修改意见、确认或退回均绑定准确修订；确认不代表外部消息已发送。

本仓不捆绑公共框架、业务数据库、客户资料、密钥或另一套Agent循环。独立离线使用仍可参考原脚本；它不等于本版主Skill已连接数字员工服务。

## 工程入口

- `lib/employee.json`：完整员工属性、原范围、运行及分发绑定。
- `clients/skill/daily-report-zh`：主Skill；`clients/sdk/daily.py`为对应SDK。
- `skills`：3项原方法；`workflows/call-daily-*`：4项原生流程调用Skill。
- `knowledge/published`：当前知识服务实际消费的正文；客户配置模板不冒充已提供客户事实。
- `server/bindings`：原七项资产、原知识运行及触发定义的准确关联。
- `quality-report.json`：有限样本、实际证据及未通过范围。

## 当前限制

全1.9.1发版门槛、跨框架发行接收和全38岗验收尚未完成；Trace存储加密R78保留。知识引用的分类/权利元数据在旧消费者中未完整返回，此缺口明确保留。其他宿主未实测，不宣称通用效果、准时推送率或经营收益。

原版说明完整保留于 `docs/original-README-before-191.md`，其中效果及兼容性声明属于原版资料，不替代本版实测范围。采用Apache-2.0，完整许可见LICENSE。
