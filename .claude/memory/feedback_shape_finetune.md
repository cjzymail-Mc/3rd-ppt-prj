---
name: Shape 微调四参数完整 + 从标准模板读值
description: fine_tuned 代码块必须含 Left/Top/Width/Height 四参数，基准值从 Template 2.1.pptx 读取
type: feedback
originSessionId: 41aa1118-8152-43ad-bbbe-6a66a3a736cd
---
用户微调 shape 位置时，代码块必须包含 Left/Top/Width/Height 四个参数，标记 `#fine_tuned`。

**Why:** 只设 Left/Top 不设 Width/Height 会导致 shape 尺寸漂移；用户已微调过的值不能被模板原始值覆盖。

**How to apply:**
- 基准值从 `src/Template 2.1.pptx` 标准模板中用 COM 读取，不从已生成的输出文件读
- 如果用户已微调过某些值，保留用户值，只补充缺失参数
- 详见 `skills/fine-tuned-shapes.md`
