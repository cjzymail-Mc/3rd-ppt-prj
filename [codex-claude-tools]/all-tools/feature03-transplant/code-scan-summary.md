# Main.py + /src 代码扫描摘要

> 移植前的代码结构快照，供 developer agent 参考。

---

## 1. 文件结构

```
src/
├── __init__.py              # 空
├── init.py                  # 空
├── Global_var_030.py        # 全局变量：dic_matrix（矩阵页码映射）、get_value()
├── Class_030.py             # 颜色常量、FONT_REGISTRY 字体表、delay
├── Function_030.py          # 核心函数库（~3300行）
└── codex_ppt.py             # yzr模板移植样板（待重命名为 yzr_ppt.py）
```

---

## 2. Main.py 关键节点

| 行号 | 内容 | 说明 |
|------|------|------|
| L14-98 | import 区 | subprocess, win32com, xlwings, tkinter 等 |
| L107 | `mc_path = os.getcwd()` | 工作目录（已加 IDE 兼容 fallback） |
| L118-120 | `from src.Class_030/Function_030/codex_ppt import *` | 三大 src 模块导入 |
| L231-235 | COM 打开 PPT | `mc_app.Presentations.Open(mc_path + r'\src\Template 2.1.pptx')` |
| L750-760 | 【5】实战测评：问卷模板页 | `questionnaire_ppt(mc_ppt, mc_slide)` |
| L765-772 | 查找问卷 sheet | 遍历 sheets 找含"问卷"的 sheet |
| L778-785 | 问卷解析（GPT） | `questionnaire_Excel(mc_sht, mc_ppt, mc_slide, mc_model, ...)` |
| **L800-809** | **Codex 分析页** | **`make_codex_slide(mc_sht, mc_ppt, mc_slide, sample_name, ...)`** |

**L800-809 是移植的插入点**：在这里加模板选择对话框，路由到 yzr 或 zxh 模块。

---

## 3. 函数调用参数约定

所有问卷相关函数共享统一参数模式：

```python
func(mc_sht, mc_ppt, mc_slide, sample_name, mc_gpt="n", mc_model="openai/gpt-5.4")
```

| 参数 | 类型 | 来源 |
|------|------|------|
| mc_sht | xlwings Sheet | 问卷 Excel sheet |
| mc_ppt | COM Presentation | PowerPoint 文件对象 |
| mc_slide | COM Slide | 当前幻灯片（函数返回更新后的 slide） |
| sample_name | str | 鞋款名称 |
| mc_gpt | "y"/"n" | 是否启用 GPT |
| mc_model | str | GPT 模型名（OpenRouter 格式） |

---

## 4. Function_030.py 关键函数

### GPT 调用
- `GPT_5(mc_prompt, model)` — L173，OpenRouter GPT 调用，pipeline 唯一外部依赖

### 问卷处理
- `questionnaire_ppt(mc_ppt, mc_slide)` — L1057，Clone Slide(4) 生成空白问卷页
- `questionnaire_Excel(mc_sht, mc_ppt, mc_slide, mc_model, ...)` — L1103，逐人生成问卷页
- `parse_survey_data(data_tuple)` — L818，解析问卷原始数据
- `extract_info(questionnaire_data)` — L976，提取关键信息

### 染色 & 格式
- `color_key(text_range, key, color, bold=1)` — L1702，单关键词染色（旧版）
- `smart_color_text(text_range, color_red, color_blue, bold=1)` — L1730，智能染色（旧版）
- `adj(text_range, size, c_color, b_color, bold=0, trs=0)` — L3249，字体属性调整

### 对话框（tkinter 模式，移植对话框的参考）
- `ask_gpt_model()` — L656，GPT 版本选择弹窗
- `center_window(win, width, height)` — L579，窗口居中
- `force_window_front(win)` — L621，窗口置顶（ctypes + tkinter）
- `flash_taskbar(win)` — L588，任务栏闪烁

### 图表 & 矩阵
- `make_chart_for_questionnaire(mc_cell, mc_slide, ...)` — L1982
- `make_chart(mc_sht, mc_slide)` — L2120
- `make_matrix(mc_sht, mc_slide)` — L2693

### 工具
- `search(mc_sht0, target, row_offset, column_offset)` — L1514，Excel 搜索
- `ppt_save(mc_ppt, sample_name, mc_path)` — L3296
- `RGB_to_Hex_to_Dec(rgb)` — L3262

---

## 5. codex_ppt.py（yzr 模板移植样板）

### 公开 API
```python
make_codex_slide(mc_sht, mc_ppt, mc_slide, sample_name, mc_gpt="n", mc_model=_MODEL)
```

### 内部结构
- `CODEX_SHAPES[]` — 9 个 shape 硬编码定义（矩形11/12/17/19/68/77, 图片39, 文本框16, 图表44）
- `_TEMPLATE_SLIDE = 15` — Clone 的模板页码（现已变为14）
- 策略路由：score_10pt / grade_letter / sample_aggregation / extract_column / extract_image / gpt_prompted / mean_extraction
- Excel 读取用 xlwings（`_xlwings_to_rows()`）
- GPT 调用复用 `src.Function_030.GPT_5`

### 与 pipeline 的差距（zxh_ppt.py 需补齐）

| 能力 | codex_ppt.py | pipeline（最新） |
|------|-------------|-----------------|
| `_write_text()` | `tr.Text = content`（无转换） | `content.replace("\n", "\r")` + `Font.Name = "微软雅黑"` |
| 关键词染色 | 单色（调用方传 color_rgb） | section-aware 双色（自动检测优势/劣势段落） |
| 字数限制 | 无 | `clamp_text()` 句子边界硬截断 |
| 列匹配 | 硬编码 `_SCORE_COLS` / `_TEXT_COLS` | 动态 `_classify_columns()` + `_find_col()` |
| 受访者数据块 | 硬编码列名 | 动态列匹配 |
| 自检 | 无（移植不需要） | 属性 + SSIM + 内容 + 字体检查（仅 pipeline 使用） |

---

## 6. Class_030.py 可复用资产

```python
# 颜色常量
black = 0;  white = 16777215;  red = 255;  green = 5287936
dark_blue = 6299648;  light_blue = 15773696;  yellow = 65535

# 字体登记表
FONT_REGISTRY = {"微软雅黑": {"usage": "中文正文/标题"}, "Arial": {"usage": "英文/数字"}}
font_exists_in_registry(font_name) → bool
register_font(font_name, usage="")
```

---

## 7. Template 2.1.pptx 页面布局

| 页码 | 用途 |
|------|------|
| 1-13 | 各类内容模板（矩阵、图表等） |
| **14** | **yzr 空白模板**（原 codex，`make_codex_slide` Clone 此页） |
| **15** | **zxh 空白模板**（`make_zxh_slide` 将 Clone 此页） |
