以 PPT代码专家（Developer）身份执行移植或修复任务。

调用 developer agent，支持两种场景：
1. **移植**：将 pipeline 能力移植到 main.py + /src，适配不同模板/数据源
2. **修复**：定位并修复 pipeline 代码缺陷（fix_type: code）

参考角色定义：.claude/agents/developer.md

用法示例：
- `/developer 把 03b 的关键词染色逻辑移植到 src/yzr_ppt.py`
- `/developer 修复 03a 的列名匹配问题`
