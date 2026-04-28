---
name: Shape 位置微调工作流
description: 用户指定 shape 名称后，在 xxx_ppt.py 的 make_xxx_slide() 中插入 Left/Top/Width/Height 微调代码块
type: feedback
---

用户经常需要对生成的 PPT 中 1-3 个 shape 做位置/尺寸微调（Left/Top/Width/Height）。

**Why:** Clone Slide 继承的模板位置不一定适合最终输出，需要在代码中硬编码修正值。

## 操作流程

1. **用户说**："帮我微调 Rectangle 68 的位置" 或 "把 XXX shape 往左移一点"
2. **定位文件**：根据 shape 名称在对应的 `XXX_SHAPES` 常量中确认属于哪个 `xxx_ppt.py`
3. **获取基准值**：从**标准模板** `src/Template 2.1.pptx` 中用 COM 读取该 shape 的 Left/Top/Width/Height（不要从已生成的输出文件读取）
4. **插入位置**：在 `make_xxx_slide()` 函数中，Clone Slide 之后、遍历 shapes 之前（`time.sleep(1.0)` 和 `for spec in XXX_SHAPES:` 之间）
5. **代码模式**（四参数完整，标记 `#fine_tuned`）：
   ```python
   # Shape 位置微调 #fine_tuned
   try:
       _shp = new_slide.Shapes("Rectangle 68")
       _shp.Left   = 20.20   # 从标准模板读取，或用户指定值
       _shp.Top    = 260.65
       _shp.Width  = 416.88
       _shp.Height = 247.19
   except Exception:
       pass
   ```
6. **注意**：如果用户已经微调过某些值，只更新用户未指定的参数，不要用模板原始值覆盖用户微调过的值

## Shape 常量命名规范

| 模板文件 | 常量名 | 模板页 |
|----------|--------|--------|
| `src/yzr_ppt.py` | `YZR_SHAPES` | Slide 15 |
| `src/zxh_ppt.py` | `ZXH_SHAPES` | Slide 17 |
| 未来新增 | `{NAME}_SHAPES` | 按模板定 |

## 当前已微调的 shapes

### yzr_ppt.py
| Shape | Left | Top | Width | Height | 备注 |
|-------|------|-----|-------|--------|------|
| Rectangle 68 | 20.20 | 260.65 | 416.88 | 247.19 | 模板 Clone 后原位微调 |
| Rectangle 77 | 450.79 | 260.65 | 274.36 | 225.38 | 模板 Clone 后原位微调 |
| TextBox 16 | 14.83 | 19.02 | 252.0 | 36.35 | 鞋款名称 TextBox（2026-04-27 用户实测） |
| Chart (原 Chart 13) | 242.19 | 21.95 | 467.24 | 224.48 | fix4 路线：覆盖在 chart 分支里（shp.Delete() 后，make_chart_for_yzr 前）；2026-04-24 用户实测值 |

#### yzr Chart 3D 视图参数（`make_chart_for_yzr` 内）

xlwings 默认 3D 视角不等于模板期望视角，必须显式设置。映射 PPT "三维旋转" 面板：

| PPT 面板 | Excel api | 用户实测值 |
|--|--|--|
| X 旋转 | `Elevation` | 20 |
| Y 旋转 | `Rotation` | 15 |
| 透视 | `Perspective` | 0 |
| 直角坐标轴 | `RightAngleAxes` | True |
| 自动缩放 | `AutoScaling` | True |
| 深度 | `DepthPercent` | 100 |
| 高度 | `HeightPercent` | 100 |

**调参工作流**：用户在 PPT 里手动调 → `read_selected_shape.py` 读 L/T/W/H → "设置形状格式 → 效果 → 三维旋转" 面板抄录 → 回写 `make_chart_for_yzr` 的 3D 视图块。详见 `.claude/memory/feedback_chart_write.md`。

### zxh_ppt.py
| Shape | Left | Top | Width | Height | 备注 |
|-------|------|-----|-------|--------|------|
| TextBox 15 | 37.75 | 128.25 | 648.99 | 330.55 | 模板原始值 |
| TextBox 17 | **650** | 146.63 | **280** | 265.11 | Left/Width 为用户微调值 |

## 单独调试入口

每个 `xxx_ppt.py` 均有 `if __name__ == "__main__"` 调试入口：
- 自动打开 `Template 2.1.pptx`（与 Main.py 同方式），不自动保存
- 连接已打开的 Excel（xlwings），自动找"问卷" sheet
- 用法：`python src/yzr_ppt.py` 或 `python src/zxh_ppt.py`
- 关键：顶部 `if __name__ == "__main__": sys.path.insert(...)` 解决直接运行时的导入问题

---

## 用 COM 实时读取当前选中的 Shape

在 IDLE / VS Code 终端中运行以下脚本，可读取**当前 PPT 中手工选中的 shape** 的所有关键属性，用于确认名称、位置、图表类型、数据等，是微调和诊断的核心工具。

```python
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import win32com.client

ppt_app = win32com.client.GetActiveObject('PowerPoint.Application')
win = ppt_app.ActiveWindow
view = win.View
slide = view.Slide
print(f'当前幻灯片: 第 {slide.SlideIndex} 页')

sel = win.Selection
print(f'Selection.Type: {sel.Type}')   # 2=shape选中, 3=文字选中, 1=无

if sel.Type in (2, 3):
    shapes = sel.ShapeRange
    print(f'选中 shape 数量: {shapes.Count}')
    for i in range(1, shapes.Count + 1):
        sh = shapes.Item(i)
        print(f'Shape: Name={sh.Name}, Type={sh.Type}')
        print(f'  位置: Left={sh.Left:.1f}, Top={sh.Top:.1f}, Width={sh.Width:.1f}, Height={sh.Height:.1f}')
        if sh.HasTextFrame:
            txt = sh.TextFrame.TextRange.Text
            print(f'  Text (前100字): {repr(txt[:100])}')
        if sh.Type == 3:   # OLE 嵌入对象（含图表）
            try:
                chart = sh.Chart
                print(f'  Chart.ChartType: {chart.ChartType}')
                sc = chart.SeriesCollection()
                print(f'  SeriesCollection.Count: {sc.Count}')
                for s_idx in range(1, min(sc.Count+1, 4)):
                    series = sc.Item(s_idx)
                    print(f'    Series[{s_idx}] Name={series.Name}, Values={list(series.Values)}')
                ax = chart.Axes(2)
                print(f'  Axes(2): Min={ax.MinimumScale}, Max={ax.MaximumScale}, AutoMin={ax.MinimumScaleIsAuto}, AutoMax={ax.MaximumScaleIsAuto}')
            except Exception as e:
                print(f'  Chart 读取异常: {e}')
else:
    print('当前没有选中任何 shape')
```

**ChartType 常用对照表**：

| 值 | 含义 |
|----|------|
| 57 | `xlBarClustered`（簇状条形图）|
| 60 | `xl3DBarClustered`（三维条形图）|
| -4151 | `xlRadar`（雷达图）|
| 51 | `xlColumnClustered`（簇状柱形图）|
| 4 | `xlLine`（折线图）|

**注意**：`Selection.Type` 需要 PPT 窗口处于前台激活状态，否则返回 1（无选中）。

---

## 新建 `xxx_ppt.py` 完整规范

以 `yzr_ppt.py` 为黄金标准模板。新建文件时按以下 Checklist 逐项确认：

### 1. 文件结构

```
xxx_ppt.py
├── 文件头注释（模板说明、公开 API）
├── sys.path.insert（直接运行时导入修复）
├── GPT_5 / overlay 三重 try/except import
├── _ppt_shared 三重 try/except import（所有共享工具）
├── _MODEL / _COPY_PASTE_DELAY 常量
├── {NAME}_SHAPES 列表（shape 配置）
├── _TEMPLATE_SLIDE 常量（模板页码）
├── 工具函数：_safe_text / _numeric / _com_get / _shoe_name
├── GPT 函数：_build_respondent_block / _build_rich_prompt / _call_gpt
├── _extract_shoe_image（图片提取，即便当前不用也保留）
├── _build_content（按 strategy 路由）
├── COM 写入：_write_text / _write_chart（含调试 print）/ _apply_keyword_color / _replace_image
├── make_xxx_slide()（公开 API，含调试 print）
└── if __name__ == "__main__"（单独调试入口）
```

### 2. `{NAME}_SHAPES` 配置规范

每个 shape 的 dict 必须包含：
- `"name"`：PPT 中的实际 shape 名（区分大小写，英文/中文与模板一致）
- `"strategy"`：见下表
- `"params"`（可选）：strategy 所需参数
- `"budget"`（gpt_prompted 必填）：`{"max_chars": N, "max_lines": N}`
- `"template_text"`（gpt_prompted 建议填）：模板原文，作为 GPT 的 style_anchor

| strategy | 含义 | 典型 shape |
|----------|------|-----------|
| `score_10pt` | 综合评分 X/10 | Rectangle 11 |
| `grade_letter` | 等级 A/B/C | Rectangle 12 |
| `sample_aggregation` | 样本统计文字 | Rectangle 17 |
| `extract_column` | 从 Excel 列读文本 | TextBox 16 |
| `mean_extraction` | 计算均值 → 写图表 | 图表 44 |
| `gpt_prompted` | 调 GPT 生成文本 | Rectangle 68/77 |
| `extract_image` | 提取 Excel 嵌入图 | Picture 39 |
| `skip` | 跳过不处理 | 装饰性 shape |

### 3. `make_xxx_slide()` 必检项

- [ ] 入口打印：`[xxx] 开始生成评测页  sample=...  gpt=...`
- [ ] 读取数据后打印：`[xxx] 读取问卷数据：N 行，M 列`
- [ ] 克隆前打印：`[xxx] 克隆模板第 N 页 → 新建第 X 页...`
- [ ] shape 微调块紧跟 `time.sleep(1.0)` 之后，用 `try/except pass` 包裹，标注 `#fine_tuned`
- [ ] 循环内每个 shape 打印：`[处理] name  strategy=...` 或 `[未找到] name`
- [ ] 末尾打印：`[xxx] 完成！新页在第 N 页`

### 4. `_write_chart()` 必检项

- [ ] `chart is None` 时打印并 return False
- [ ] labels 解析为 0 时打印 content 前 60 字并 return False
- [ ] 写入前打印指标数量和 label/value 列表
- [ ] `ChartData.Activate/BreakLink` 的异常要打印（不能静默吞掉）
- [ ] 写入成功/失败均打印

### 5. `if __name__ == "__main__"` 调试入口必检项

- [ ] `sys.path.insert(0, proj_root)` 放最前面
- [ ] 用 `win32com.client.Dispatch` 打开模板 PPT（不复用已有实例）
- [ ] 用 `xlwings.books.active` 连接已打开的 Excel
- [ ] 自动找"问卷" sheet，找不到时打印提示并 `exit(1)`
- [ ] 从"基础信息" sheet 读 `sample_name`
- [ ] `mc_gpt = "n"`（调试默认关闭 GPT）
- [ ] 末尾打印完成信息和"不要保存"提示

### 6. 不要重复造轮子

以下函数/逻辑**不要**在 `xxx_ppt.py` 中自己重写，直接从 `_ppt_shared` import：

| 函数 | 说明 |
|------|------|
| `_extract_score_means` | 已修复"轮次列混入"bug |
| `_classify_columns` | 动态识别评分列/文本列 |
| `_find_col` / `_col_values` | 列查找 |
| `_score_10pt` / `_score_to_grade` | 评分换算 |
| `_xlwings_to_rows` | Excel → rows |
| `clamp_text` | 文本截断（剔空行 + 字数 / 行数硬上限） |
| `_apply_keyword_color` | per-shape **section context** 染色：扫段头切换 advantage/disadvantage 模式，再把 `【keyword】` 按当前段染红/蓝。适用 yzr/zxh 单 shape 单段语境（每个 shape 只有"优点"或"缺点"一种基调） |
| `_apply_conclusion_color` | per-shape **bracket-typed** 染色：`<keyword>` 红+粗 / `[keyword]` 蓝+粗 / `(keyword)` 仅粗。适用 6.3 结论页这种"单 shape 内多段、多色"场景；剥离 ASCII 标记，保留中文 【】 给 section header |
| `_strip_bullet_on_section_headers` | 段头 `【XX】` 行去掉 ■ bullet（`Result_Bullet` 默认每段都加 ■，段头加 ■ 视觉冗余） |

**两套染色函数的选用决策**：

- 若一个 shape 只写"优点"或"缺点"一种基调 → `_apply_keyword_color` + GPT prompt 用 `【keyword】` 单一标记
- 若一个 shape 内有"优点 / 缺点 / 修改建议"多段 → `_apply_conclusion_color` + GPT prompt 用 `<>` / `[]` / `()` 分类标记
- 详见 `.claude/memory/feedback_conclusion_coloring.md`

`parse_survey_data` 在 `Function_030.py`，已修复"has_one 误删列"bug，新模板通过 `questionnaire_Excel` 间接使用，无需关心。

---

## 已知 Bug 与修复经验（2026-04-17）

### Bug 1：`parse_survey_data` 误删评分列

**现象**：问卷数据清洗后少了某个性能维度（如"抓地性"），条形图比原始数据少 1 列。

**根因**：`Function_030.py` 的 `parse_survey_data` 用 `not has_one` 来排除"第几轮反馈"列——只要某列有任意一行的值 = 1.0，整列被排除。但性能评分完全可以打 1 分。

**修复**：改为检查列**标题**是否含 `["第几轮", "轮次", "轮反馈", "这是第几"]`，不依赖数据值：
```python
_round_keys = ["第几轮", "轮次", "轮反馈", "这是第几"]
is_round_col = any(k in col_header_str for k in _round_keys)
if numeric_count >= ... and not is_round_col:
    score_indices.append(col_idx)
```

**新模板注意**：`parse_survey_data` 是共享函数（`Function_030.py`），新模板直接受益，无需重新处理。

---

### Bug 2：`_extract_score_means` 把"第几轮反馈"混入均值

**现象**：雷达图 / 条形图多出一个值为 1.0 的异常数据点（"这是第几轮反馈"均值）。

**根因**：`_ppt_shared.py` 的 `_extract_score_means` 的 `reject_keys` 里没有"轮"类关键词，"这是第几轮反馈"被当成普通评分列计算均值。

**修复**：在 `reject_keys` 中追加 `["第几轮", "轮次", "轮反馈", "这是第几"]`。

**新模板注意**：`_extract_score_means` 来自共享的 `_ppt_shared.py`，新模板 import 后自动受益。

---

### Bug 3：条形图坐标轴自适应导致视觉失真

**现象**：7-8 分的差距在图表上看起来巨大（坐标轴从 7 开始而非 0）。

**根因**：原代码在格式化后直接 `Axes(2).Delete()`。`Delete()` 会让 Excel 重置为自动量程，之前手动设置的 MinimumScale/MaximumScale 全部失效。

**修复**：**不能先 `Delete()` 再设 scale**。改为：
1. 识别量表范围：取评分列数据，有值 > 5 则为 10 分制，否则 5 分制
2. 设置固定量程：`MinimumScale=0`，`MaximumScale=5或10`
3. 隐藏坐标轴（代替 Delete）：设 `TickLabelPosition=-4142`、`MajorTickMark=-4142`、`Format.Line.Visible=0`

```python
# 识别量表（5分制 or 10分制）
_scale_max = 10 if any(v > 5 for v in flat_vals) else 5

# 固定轴 + 隐藏（不 Delete）
_val_axis = mc_chart1.api[1].Axes(2)
_val_axis.MinimumScaleIsAuto = False
_val_axis.MaximumScaleIsAuto = False
_val_axis.MinimumScale = 0
_val_axis.MaximumScale = _scale_max
_val_axis.TickLabelPosition = -4142   # xlTickLabelPositionNone
_val_axis.MajorTickMark = -4142       # xlTickMarkNone
_val_axis.MinorTickMark = -4142
_val_axis.Format.Line.Visible = 0     # msoFalse
```

**适用范围**：`Function_030.py` 的 `make_chart_for_questionnaire`（问卷条形图）。新模板如果自建图表函数，同样适用此规则。

---

### Bug 4：`_write_chart` 假成功——`Activate` 失败后 `series.Values` 静默失效

**现象**：终端打印 `[图表] 写入成功`，但 PPT 中图表数据是模板原始占位值（坐标轴显示 0-1，bars 不可见）。

**根因**：`ChartData.Activate()` 失败后代码继续执行，`series.Values = tuple(values)` 在部分 Office 版本上不抛异常但写入无效——原始模板数据没有被替换。

**修复**：
1. `ChartData.Activate()` 改为重试 3 次（间隔递增），并打印每次结果
2. 写入后**回读验证**：`actual_vals = list(series.Values)`，若首值与期望误差 > 0.05 则报 `验证失败`，明确提示用户
3. `BreakLink` 仍保持可选（失败不中断主流程）

**新模板注意**：直接复制 `yzr_ppt.py` 的 `_write_chart`，不要用旧版（无验证逻辑的版本）。

**常见触发环境**：
- 同事的 Office 版本与开发者不同（尤其是 PPT 2016 vs 2019/365）
- PPT 文件以只读方式打开
- 某些企业加密环境下 `ChartData.Activate()` 被拦截

---

### 规范：`xxx_ppt.py` 必须包含调试输出

**问题**：`make_xxx_slide()` 和 `_write_chart()` 原本无任何 print，运行时完全不知道进度和报错。

**规范**：新建 `xxx_ppt.py` 时，`make_xxx_slide()` 和 `_write_chart()` 必须包含以下打印：

**`make_xxx_slide()` 入口和 shape 循环**：
```python
print(f"\n[xxx] 开始生成评测页  sample={sample_name}  gpt={'开启' if gpt_enabled else '关闭'}")
print(f"[xxx] 读取问卷数据：{len(rows)} 行（含标题行），{len(rows[0]) if rows else 0} 列")
print(f"[xxx] 克隆模板第 {_TEMPLATE_SLIDE} 页 → 新建第 {X} 页...")
# 循环内：
print(f"  [处理] {name}  strategy={strategy}")
print(f"  [未找到] {name}（模板中不存在此 shape，跳过）")
# 末尾：
print(f"[xxx] 完成！新页在第 {new_slide.SlideIndex} 页")
```

**`_write_chart()` 关键节点**：
```python
print(f"  [图表] 准备写入 {len(labels)} 个指标: {list(zip(labels, values))}")
print(f"  [图表] 写入成功")
# 异常时：
print(f"  [图表] 写入失败: {_e}")
```

**`mean_extraction` 分支**：
```python
print(f"  [均值] 提取到 {len(means)} 个指标均值: {[(k, round(v,2)) for k,v in means[:8]]}")
```

---

### Bug 5：`Shapes.Paste()` 返回 ShapeRange，`.Chart` 静默失败（2026-04-27）

**现象**：`make_chart_for_yzr` 粘贴 chart 到 PPT 后，chart 主标题（数值/Series Name）始终不消失，即便代码里写了 `mc_shape.Chart.SetElement(0)`。

**根因**：`mc_slide.Shapes.Paste()` 返回的是 **ShapeRange**（不是 Shape）。
- ShapeRange 的 `.Left/.Top/.Width/.Height` **会** fan-out 到内部 shape，所以这些代码不报错；
- 但 `.Chart` / `.HasChart` 不在 fan-out 列表，访问时抛 `com_error -2147352567 发生意外`；
- 旧代码用 `try: mc_shape.Chart.SetElement(0) except Exception as _e: print(...)`，错误被静默吞掉，title 永远没被隐藏，且 print 混在大量日志里没被注意到。

**修复**（`_ppt_shared.py::make_chart_for_yzr`）：
```python
try:
    _shape_one = mc_shape.Item(1) if hasattr(mc_shape, "Item") else mc_shape
    _shape_one.Chart.HasTitle = False     # 属性直写
    _shape_one.Chart.SetElement(0)        # UI 命令，双保险
    print("  [yzr-chart] PPT 端主标题已隐藏")    # 成功也要打 print
except Exception as _e:
    print(f"  [yzr-chart] PPT 端隐藏标题失败（{_e!r}）")
```

**衍生规则**（已写入 CLAUDE.md § 3 和 feedback_com_constraints.md）：
- `Shapes.Paste()` 返回 ShapeRange，访问 `.Chart` 必须先 `.Item(1)`
- silent except 反模式：必须 success/failure **都打 print**，否则修了等于没修

**新模板注意**：所有"粘贴后改 chart 属性"的代码都按此模式写。

---

### Bug 6：bar chart 数据标签压在 bar 末端（2026-04-27）

**现象**：score = 量表最大值（10 分制满分=10）时，对应 bar 末端的数据标签被 bar 本身压住、看不清。

**根因**：`make_chart_for_questionnaire` 中 `_val_axis.MaximumScale = _scale_max`（5 分制→5，10 分制→10），导致满分 bar 占据整条数值轴，数据标签没有显示空间。

**修复**（`Function_030.py::make_chart_for_questionnaire:2075`）：
```python
_axis_max = _scale_max + 1   # 5→6，10→11
_val_axis.MaximumScale = _axis_max
```

**新模板注意**：自建 bar chart 时同样适用，`_ppt_shared.py::make_chart_for_yzr` 是 3D bar，目前 `MaximumScale = 10` 硬编码（未分发场景未触发问题）。后续若要分发，统一切到 `_scale_max + 1`。

---

### Bug 7：tk popup 不居中 / 任务栏不闪烁（2026-04-27）

**现象**：第二个 GPT 弹窗（版本选择）+ 第三个弹窗（模板选择）经常不居中、任务栏图标不闪烁，用户根本没注意到弹窗已弹出。

**根因**：
1. `force_window_front` / `flash_taskbar` 用 `win.winfo_id()` 取 HWND——这是 **Tk 子控件 HWND**，`SetWindowPos` / `FlashWindowEx` 对它静默失败。
2. `center_window` 用 `winfo_screenwidth()`——多显示器下返回主屏分辨率，弹窗永远跑到主屏，PPT 在副屏时用户看不到。
3. `flash_taskbar(win)` 调用整体被注释掉了。

**修复**（`Function_030.py`）：
- 新增 `_get_toplevel_hwnd(win)`：用 `int(win.wm_frame(), 16)` 取真正的顶层 HWND，失败回退 `GetParent` 兜底
- `center_window` 改为按光标当前所在屏 `MonitorFromPoint + GetMonitorInfoW.rcWork` 居中（多屏友好，避开任务栏）
- `force_window_front` 默认开启**任务栏闪烁 + 系统蜂鸣**（`MessageBeep`），并在 400ms 后切回 `NOTOPMOST`，避免永久挡 PPT
