# HTML → PPT 转换：架构、要点与踩坑手册

> 从 Step3/Step4 两轮实战中提炼。适用于任何 HTML 演示稿转 PPT 的项目。

---

## 1. 核心思路：混合分层方案

**不要试图用代码 100% 重建 HTML 的视觉效果。**

正确做法是分两层：

| 层 | 内容 | 来源 |
|----|------|------|
| **背景层** | 渐变、装饰条、卡片边框、空表格框线等纯视觉元素 | Playwright 截图（文字/图片设透明后截图） |
| **叠加层** | 可编辑文本框、可替换图片、可编辑表格 | 从 HTML 提取坐标后用 win32com 精确放置 |

这样做的好处：
- 装饰效果零成本还原（截图天然精准）
- 文字/图片/表格全部可编辑（用户可以直接改内容）
- 坐标精确（从浏览器 `getBoundingClientRect()` 直接映射）

---

## 2. 三脚本流水线

```
extract_layout.py → _layout_manifest.json   （坐标+文本+样式）
         ↓
screenshot_masked.py → _slide_shots_masked/  （装饰截图）
         ↓
export_hybrid_win32com.py → deck.pptx        （组装最终PPT）
```

### 脚本 1：extract（Playwright 坐标提取）

- 用 Playwright 打开 HTML，注入 LAYOUT_CSS 将每页固定为 1260x720px
- 遍历每个 `section.slide`，用 JS 提取所有元素的 `getBoundingClientRect()` + 计算样式
- 元素类型：text / bullets / numbered / image / table / kpi_card / timeline / page_label
- 输出 JSON，每个元素包含 `{type, x, y, w, h, text, font_size_px, color, bold, align, ...}`

### 脚本 2：screenshot（遮罩截图）

- 同样注入 LAYOUT_CSS（**必须与 extract 完全一致**，否则坐标系偏差）
- CSS 将所有文字设为 `color: transparent`，所有图片设为 `visibility: hidden`
- 逐页截图，`device_scale_factor=3`（3780x2160px，清晰度足够）

### 脚本 3：export（win32com 组装）

- 读取 manifest JSON，逐页创建幻灯片
- 先放截图背景（`ZOrder(1)` 发送到底层）
- 再按 manifest 逐个添加可编辑 shape

---

## 3. 坐标映射公式

```
HTML 画布：1260 x 720 px
PPT 尺寸：960 x 540 pt（16:9）

SX = 960 / 1260 = 0.7619
SY = 540 / 720  = 0.75

位置：x_pt = px * SX,  y_pt = px * SY
字体：pt = css_px * 0.75
```

**所有 helper 函数必须走同一套映射，禁止手算常量。**

---

## 4. 八大踩坑教训（C-1 ~ C-8）

这是两轮实战中遇到的所有关键 Bug，按严重程度排序：

### C-1 SlideSize 枚举陷阱（致命）

```python
# 错误：ppSlideSize=4 是 35mm 幻灯片（810x540pt = 3:2），不是 16:9！
pres.PageSetup.SlideSize = 4  # 别用这个

# 正确：显式设置宽高
pres.PageSetup.SlideWidth  = 960   # pt
pres.PageSetup.SlideHeight = 540   # pt
```

### C-2 extract 与 screenshot 的 LAYOUT_CSS 必须完全一致（致命）

两个脚本都会注入 CSS 来固定页面尺寸。如果 CSS 不同（比如一个是 1236px、另一个是 1260px），提取的坐标和截图的坐标系就会偏移，叠加层和背景层错位。

**解决：两个脚本共享同一段 LAYOUT_CSS 字符串。**

### C-3 CSS padding → PPT TextFrame.Margin

`getBoundingClientRect()` 返回的是 padding box 的外边界。如果元素有 padding，文字实际起始位置在 padding 内侧。

```python
# 提取 padding
pad_l = el.get("pad_l", 0)
# 映射到 PPT margin
tf.MarginLeft = pad_l * SX
```

### C-4 rgba 半透明色必须 alpha blending

HTML 中 `rgba(255,255,255,0.5)` 不是 `#ffffff`。必须根据背景色混合：

```javascript
// Hero 暗底 #0d1424，普通页白底 #ffffff
r = Math.round(r * a + bgR * (1 - a));
```

### C-5 HTML `<br>` 换行 → PPT 需要 `\r`

`innerText` 返回 `\n`，但 COM 的段落分隔符是 `\r`：

```python
text = text.replace("\n", "\r")
```

### C-6 文本框宽度加 8% buffer

浏览器和 PPT 的字体渲染引擎不同，同样的文字在 PPT 中可能更宽。加 buffer 防溢出：

```python
buf_w = max(ew, ew * 1.08)
```

### C-7 win32com 常量必须用数值字面量

在脚本顶层 `import` 时 COM 还没启动，引用 `constants.ppAlignLeft` 会报错：

```python
# 错误
from win32com.client import constants
constants.ppAlignLeft  # COM 未启动，失败

# 正确：直接用数字
ppAlignLeft = 1
ppAlignCenter = 2
ppAlignRight = 3
ppLayoutBlank = 12
msoTrue = -1
msoFalse = 0
```

### C-8 中文标点与 Python 引号冲突

中文 `、`、`""`、`（）` 等字符如果出现在 Python 字符串拼接中容易引发 SyntaxError。使用 JSON manifest 驱动数据，避免在脚本中硬编码中文内容。

---

## 5. 自检流程（强制执行，不可跳过）

### 阶段 B：前 3 页先行验证

```
SLIDE_RANGE=1-3 py -3 pipeline/export_xxx.py
```

先只生成 3 页，用 win32com 打开 PPT 逐页读取 shape：
- 坐标偏差 > 3pt → 严重问题，必须修复
- 文本缺失 → 严重问题
- 图片缺失 → 严重问题

**前 3 页无问题后，才生成全部页面。**

### 四步自检法

| 步骤 | 检查内容 | 判定标准 |
|------|----------|----------|
| 1. 坐标抽样 | 随机 3-5 页，对比 manifest 坐标与 PPT shape 位置 | 偏差 <= 3pt |
| 2. 可编辑性 | 每页检查：文本框可编辑、图片可选中、表格可编辑 | 全部通过 |
| 3. 视觉对照 | PPT 与 HTML 逐页比对 | 文字不漂移、图片不错位 |
| 4. 内容完整性 | 对照 manifest 核查每页元素数量 | 数量一致 |

输出 `conversion_report.md`，记录所有发现和修复。

---

## 6. 关键元素处理方式

| HTML 元素 | PPT 处理 | 注意事项 |
|-----------|----------|----------|
| 普通文字（h1/h2/h3/p） | `AddTextbox` | 宽度 +8% buffer，padding 映射为 Margin |
| 列表（ul/ol） | `AddTextbox` + bullet 格式 | 列表 padding-left 要偏移文本框起点 |
| 图片（img） | `AddPicture` + crop | 模拟 `object-fit: cover` 需手动裁切 |
| 表格（table） | `AddTable` | 表头深底白字，列宽按比例分配 |
| KPI 卡片 | 背景在截图中 + 叠加 label/value 两个文本框 | 内边距 12px |
| 时间轴 | badge 文本框 + content 文本框 | badge 宽度固定 90px |
| 页码 | 右对齐文本框 | 固定位置：右下角 |
| 装饰（渐变/阴影/边框） | 全部留在截图背景中 | 不重建 |

---

## 7. 中文字体处理

```python
tr.Font.Name = "Microsoft YaHei"          # Latin 字体槽
tr.Font.NameFarEast = "Microsoft YaHei"    # 东亚字体槽
```

**必须同时设置两个属性**，否则中文可能回退到宋体。

---

## 8. 图片 cover-crop 实现

```python
def add_picture_cover(slide, img_path, left_px, top_px, width_px, height_px):
    shp = slide.Shapes.AddPicture(str(img_path), 0, -1, x(left_px), y(top_px), -1, -1)
    target_w, target_h = x(width_px), y(height_px)
    cur_ratio = shp.Width / shp.Height
    tgt_ratio = target_w / target_h
    if cur_ratio > tgt_ratio:           # 图片更宽 → 裁左右
        want_w = shp.Height * tgt_ratio
        crop = (shp.Width - want_w) / 2
        shp.PictureFormat.CropLeft = crop
        shp.PictureFormat.CropRight = crop
    else:                               # 图片更高 → 裁上下
        want_h = shp.Width / tgt_ratio
        crop = (shp.Height - want_h) / 2
        shp.PictureFormat.CropTop = crop
        shp.PictureFormat.CropBottom = crop
    shp.Left, shp.Top = x(left_px), y(top_px)
    shp.Width, shp.Height = target_w, target_h
```

---

## 9. 环境依赖

```
Windows + Microsoft PowerPoint（Office）
Python 3.10+
pywin32      → win32com.client（操控 PowerPoint COM）
playwright   → 坐标提取 + 截图
pillow       → 图片处理（可选）
```

安装：
```bash
py -3 -m pip install pywin32 playwright pillow
py -3 -m playwright install chromium
```

探针（确认 COM 可用）：
```bash
py -3 -c "import win32com.client as w; app=w.Dispatch('PowerPoint.Application'); print('OK'); app.Quit()"
```

---

## 10. 常见问题速查

| 问题 | 原因 | 解决 |
|------|------|------|
| PPT 比例不是 16:9 | 用了 `SlideSize=4`（3:2） | 改用 `SlideWidth=960, SlideHeight=540` |
| 文字和背景错位 | extract/screenshot 的 LAYOUT_CSS 不一致 | 两个脚本共享同一段 CSS |
| 文字位置偏移 | 没处理 CSS padding | 提取 padding 映射为 TextFrame.Margin |
| 半透明色颜色不对 | 直接丢弃 alpha | 按背景色做 alpha blending |
| 中文变宋体 | 只设了 Font.Name | 同时设 Font.NameFarEast |
| 文字溢出文本框 | PPT 字体渲染更宽 | 宽度加 8% buffer |
| 换行丢失 | `\n` 在 COM 中无效 | 替换为 `\r` |
| COM 常量报错 | 模块加载时 COM 未启动 | 用数值字面量代替 |
| 脚本连续失败 | 各种原因 | 同一方案失败 2 次就停下换方案 |
