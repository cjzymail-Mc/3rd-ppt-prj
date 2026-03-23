---
name: feedback_com_constraints
description: 技术栈硬约束：必须用COM，禁止openpyxl/python-pptx/numpy
type: feedback
---

Excel/PPT 读写必须用 `win32com.client` COM 接口。

**Why:** 本地环境有加密 Excel 文件，openpyxl/pandas 无法读取。python-pptx 无法保留模板的字体/颜色/阴影等格式。

**How to apply:** 任何涉及 xlsx 或 pptx 操作的代码，统一用 COM。禁止 openpyxl、pandas 直接读写 xlsx，禁止 python-pptx，禁止 numpy。GPT 调用复用 `src/Function_030.py` 的 `GPT_5()` 函数。
