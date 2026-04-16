执行步骤3：构建 & 交付 PPT。

调用 step3-builder agent，完成：
- 智能检测 prompt 是否更新（F1）
- Python pipeline 通过 COM 写入 PPT（03b_build_ppt_com，内置自检已完善）
- 失败时诊断问题层级，建议回到对应步骤

前置要求：已完成步骤2。
完成后打印摘要，并打开 PPT 供审核。
