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
| Excel/PowerPoint 批量自动化（分析/生成类脚本） | `Dispatch` 复用实例 → attach 到用户活 Office + `.Quit()` 把它一起关（2026-05-28 实际丢用户未保存内容） | `DispatchEx` 强制独立进程（详见下方"Dispatch vs DispatchEx 雷区"）|
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

## Dispatch vs DispatchEx 雷区（2026-05-28 实战补充）

**事故**：跑 `pipeline/01_shape_detail.py` 时 `generate_shape_detail_xlsx` 用 `Dispatch("Excel.Application")` + `app.Visible=False` + finally `excel.Quit()`。`Dispatch` attach 到用户**正在编辑**的 Excel 实例 → 设 `Visible=False` 时被拒（"Property 'Excel.Application.Visible' can not be set"）→ 异常传播到 finally → `excel.Quit()` 把用户的 Excel 整个关掉。用户来不及保存，被迫选"未保存"丢内容。同类雷在 PowerPoint 侧也存在（`Dispatch("PowerPoint.Application")` + `app.Quit()` 同样关用户活 PPT）。

**判据：用 `Dispatch` 还是 `DispatchEx`？看脚本意图**

| 脚本类型 | 选谁 | 例 |
|---|---|---|
| **批量分析 / 批量生成**（开模板/Excel 读数据 → 产单独输出文件 → 退出）| `DispatchEx` 独立进程 | `pipeline/01_shape_detail.py`、`pipeline/03b_build_ppt_com.py`、`pipeline/ppt_pipeline_common.py::{load_excel_rows, generate_shape_detail_xlsx, create_iteration_sheet}`、`skills/inspect-*-template/*.py` |
| **驱动用户活 Office**（往用户开着的 PPT 里写、读用户当前选中 shape）| `Dispatch` 共享实例 | `Main.py`、`src/*_ppt.py` 生产流程、`skills/read_selected_shape.py` |

**规范化收益**：把所有"批量类"脚本里的 `Dispatch("Excel.Application")` / `Dispatch("PowerPoint.Application")` 全部改成 `DispatchEx`——独立进程不会 attach 用户实例，`.Quit()` 也只关自己开的那个进程。

**只读类还可叠加 `Open(ReadOnly=True, WithWindow=False)`**（镜像 `inspect-ppt-template` 的安全开法，保证脚本绝不修改模板、且不弹窗）。详见 Step1 PowerPoint 隔离改动：`pipeline/01_shape_detail.py` line ~165。
