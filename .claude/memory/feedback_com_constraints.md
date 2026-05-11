---
name: feedback_com_constraints
description: 技术栈硬约束：必须用COM，禁止openpyxl/python-pptx/numpy
type: feedback
---

Excel/PPT 读写必须用 `win32com.client` COM 接口。

**Why:** 本地环境有加密 Excel 文件，openpyxl/pandas 无法读取。python-pptx 无法保留模板的字体/颜色/阴影等格式。

**How to apply:** 任何涉及 xlsx 或 pptx 操作的代码，统一用 COM。禁止 openpyxl、pandas 直接读写 xlsx，禁止 python-pptx，禁止 numpy。GPT 调用复用 `src/Function_030.py` 的 `GPT_5()` 函数。

## COM 开发规范

| 场景 | 错误做法 | 正确做法 |
|------|---------|---------|
| 读COM属性 | `getattr(shp,"X",None)` | `try: shp.X except: None` |
| 多步骤开Excel | `Dispatch` 复用实例 | `DispatchEx` + `sleep(0.5)` 强制新进程 |
| 写图表数据 | `ChartData.Workbook` | `SeriesCollection(1).Values/XValues` |
| 插入图片 | `AddPicture(W=w,H=h)` | 先`-1/-1`取原始尺寸,再等比缩放 |
| Clone幻灯片 | 不加sleep | `Copy→sleep(1.5)→Paste(X)→sleep(1.0)` |
| 访问 `Shapes.Paste()` 的 chart | `mc_shape.Chart.X`（抛 -2147352567） | `mc_shape.Item(1).Chart.X`（Paste 返 ShapeRange，.Chart 不 fan-out） |
| tk 顶层窗口 HWND | `win.winfo_id()`（子控件 HWND，FlashWindowEx 静默失败） | `int(win.wm_frame(), 16)` 或封装的 `_get_toplevel_hwnd(win)` |
| 多显示器居中弹窗 | `winfo_screenwidth()`（主屏分辨率，副屏看不见） | `MonitorFromPoint(GetCursorPos) + GetMonitorInfoW.rcWork` |
| chart 主标题隐藏 | 单调 `SetElement(0)`（COM 时序敏感） | `HasTitle=False` + `SetElement(0)` 双保险 |
| GPT 输出文本入 PPT | 直接 `_write_text(shp, gpt_out)` | 先 `clamp_text(gpt_out, max_chars, max_lines)`（自动剔空行） |
| Excel chart Copy/Delete | UI Selection 路径（`Range.Select() → api[0].Copy()`，需 `Excel_zoom` 让 chart 进视口） | 对象引用路径（`mc_chart1 = mc_sht.charts.add()` → `mc_chart1.api[0].Copy()` / `_tmp_chart.delete()`），免疫缩放/滚动/视口可见。详见 `feedback_chart_write.md` |
| 读 PPT 当前选中 shape | 凭"无访问能力"否认，让用户描述 | `python skills/read_selected_shape.py`（项目自带 win32com `GetActiveObject` 桥接）；通过 Bash 工具跑一次性 python 时 stdout 中文乱码套餐见 `feedback_python_stdout_encoding.md` |
| PPT 文本换行符 | `content` 含 `\n` 直接写入 → 整段变 1 段 | `content.replace("\n", "\r")` 再写；TextRange.Text 用 `\r` 分段 |
| PNG 截图（系统加密 Office 输出环境） | `slide.Export("PNG")` / `SaveAs(PDF)` 输出被 DLP 加密，Pillow 读不了 | `slide.Copy()` → `PIL.ImageGrab.grabclipboard()` → `img.save()`（剪贴板数据由 Python 写出，不加密） |
| 写入文本后字体 | 默认接管为系统默认字体（非模板字体） | `_write_text()` 后显式 `tr.Font.Name = "微软雅黑"` |
