读取 todays-task.md,执行其中任务。

【强制门控】读完 todays-task.md 后,先做路由判断,再动手:
1. 任务是否匹配已有 role(builder/researcher/converter/archivist/pm)或 skill?
   - 是 → 立即调用对应 slash command,不要从零规划、不要自己先读项目文件
   - 否 → 进入下一步
2. 任务是否需要盘点多个相关文件(skills/agents/templates/pipeline 等)才能动手?
   - 是 → 派 1 个 Explorer subagent 一次性平行盘清楚,不要自己串行 Glob+Read
   - 否 → 自己直接动手

效率规则:
- todays-task.md 直接引用的文件(如 mc-debugN.md)和它在同一条消息里并行 Read;路径未知必须 Glob 时,Glob+Read 串行可接受
- 扫描/查找/盘点类只读动作派给 Explorer subagent(单点 1 个、多角度并行多个),不要自己 Glob/Grep 翻
- 写文件、改代码、跑命令仍由你自己执行,不委托 Explorer

通信规则:
- 不要复述 todays-task.md 内容,直接说"开始执行 X"后动手
- 收尾只报:完成了什么 / 卡点 / 下一步建议,不要列过程清单
