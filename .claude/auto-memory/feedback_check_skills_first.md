---
name: Check skills/ before declining capability
description: 涉及用户当前打开文件 / 选中状态 / 屏幕实时数据时，先查 skills/ 和 debug/，不要凭默认能力边界否认
type: feedback
originSessionId: 7de9f536-08a2-4eb4-b494-0d7d0c3ee6e6
---
# Check `skills/` and `debug/` before saying "I can't"

任何涉及"用户当前打开的文件 / 选中的对象 / 屏幕实时状态"的请求，**第一反射是 `Glob skills/* debug/*`**，不是先用默认 Claude 能力边界拒绝。

**Why:** 2026-04-29 用户让我读取 PPT 当前选中的 shape，我直接回答"看不到"，被用户指正——`skills/read_selected_shape.py` 就在仓库里，文件名和注释都写得很白（"读取当前鼠标选中的 PPT shape"），我连 Glob 都没跑就先否认。这是**默认能力假设**击败了**项目实际能力**——本项目通过 win32com 桥接到正在运行的 Office 实例，是**有完整能力**读取实时状态的。

**How to apply:**
- 触发词：用户消息出现 "我选中的"、"我当前打开的"、"屏幕上的"、"刚才那个"、"这个 shape/cell/chart" 等指代实时状态的表述
- 第一步：`Glob skills/read_*`、`Glob debug/read_*`、`Grep "选中\|active\|GetActiveObject"`
- 第二步：找到对应脚本就直接 `python skills/xxx.py` 跑，输出贴回对话
- 第三步：找不到才说"项目里没有现成的，要我写一个吗？"
- **决不**：凭"我没有访问 X 的能力"这种通用边界回答先否认
- 涉及 PPT/Excel COM 时，本项目的工具集 = `skills/*.py`、`debug/*.py`、`src/Function_030.py` 里的 helper、`src/_ppt_shared.py`，扫一遍再下判断
- 输出乱码（GBK/UTF-8 不匹配）时，加 `PYTHONIOENCODING=utf-8` 重跑
