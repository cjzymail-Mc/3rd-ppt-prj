执行步骤1：分析 PPT 模板。

调用 step1-analyzer agent，完成：
- Python pipeline 提取 shape 结构（01_shape_detail + 01b_auto_annotate + 02_shape_analysis）
- 内部自检循环（对比模板完整性）
- 自检失败时由 LLM 修复批注
- 最多循环 2 次

完成后打印摘要，并打开 Excel 供审核。
