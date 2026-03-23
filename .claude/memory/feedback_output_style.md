---
name: feedback_output_style
description: 输出风格：只说结论，不展示diff，精简token消耗
type: feedback
---

改代码时只说结论：改了什么、为什么改、结果如何。不要展示 old_string / new_string / 代码 diff。

**Why:** Edit 工具的 old/new 内容大量消耗上下文 token，导致频繁 compact。用户不需要看中间过程。

**How to apply:** 所有代码修改后，用一句话总结改动。表格/流程图等结构化输出优先。用户问细节时再补充。
