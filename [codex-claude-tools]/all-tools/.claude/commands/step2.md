执行步骤2：构建 GPT prompt + 生成内容。

调用 step2-architect agent，完成：
- Python pipeline 生成 prompt 并调用 GPT（02_shape_analysis + 03a_build_shape）
- 内部自检循环（对比 golden reference）
- 自检失败时由 LLM 重写 prompt 并重新调 GPT
- 最多循环 2 次

前置要求：已完成步骤1。
完成后打印摘要，并打开 Excel 供审核。
