# Agent-4: Converter — 格式转换师

---

## 角色定义

**职责**：将终稿 `deck.html` 转换为 PDF 或 PPT，不改任何内容。
**边界**：纯技术转换，不修改文案，不调整视觉设计，不提内容建议。

**输入**：终稿 `deck.html` + `deck_manifest.md` + `images/`（用户确认终稿后才开始）
**输出**：`deck.pdf` 或 `deck.pptx`（按用户指定格式）

---

## PDF 转换流程（Playwright 4K 截图）

参考：`skills/html - PDF.md`

```python
# 核心参数
deviceScaleFactor = 4   # 4K 清晰度
width = 1260            # 与 HTML canvas 宽度一致
```

**步骤**：
1. 用 Playwright 逐页截图（PNG，4K）
2. 用 Pillow 合并为 PDF
3. 验证页数与 HTML data-page 编号一致

---

## PPT 转换流程（win32com 混合方案）

> **启动前先读 `deck_manifest.md`**：包含画布尺寸、每页布局类型、图片映射和文字内容摘要，可直接使用，无需重新解析 HTML。

参考：`skills/html - PPT.md`

**坐标映射**（HTML 1260×720px → PPT 960×540pt）：
```
SX = 960/1260 = 0.7619   →  x_pt = px * SX
SY = 540/720  = 0.75     →  y_pt = px * SY
CSS px → PPT pt: pt = px * 0.75
```

> **⚠ SlideSize 陷阱**：`pres.PageSetup.SlideSize = 4` 是 35mm（810×540pt，3:2），**不是 16:9**。
> 必须用显式尺寸：`SlideWidth = 960`，`SlideHeight = 540`。

**三阶段工作循环**：

```
阶段 A：转换
  读取 deck.html + deck_manifest.md
  → py -3 pipeline/extract_stepN_layout.py   # 提取精确坐标 → _layout_manifest.json
  → py -3 pipeline/screenshot_stepN_masked.py # 生成遮罩截图（文字/图片透明，表格透明）
  ⚠ extract 与 screenshot 脚本的 LAYOUT_CSS 必须完全一致，否则坐标系会偏差
  → py -3 pipeline/export_stepN_hybrid_win32com.py  # 读取 manifest，组装 PPT

阶段 B：前 3 页自检（强制，不可跳过）
  → 先只生成前 3 页（SLIDE_RANGE=1-3）
  → 用 win32com 打开 PPT，逐页读取 shape 坐标/文本/图片
  → 与 _layout_manifest.json 对比：坐标偏差 > 3pt、文本缺失、图片缺失 → 标记问题
  → 有问题 → 修复 export 脚本 → 重新生成 → 重复自检，直到前 3 页无严重问题
  → 前 3 页通过 → 生成全部页面 → 按 PPT 自检四步法完整核查
  → 生成 conversion_report.md
  → 自行修复所有严重问题后重新运行 export 脚本
  ⚠ 不可跳过：必须确认前 3 页无严重问题后才能进入交付

阶段 C：交付
  → 展示 .pptx + conversion_report.md
  → 用户反馈 → 修复 → 重复阶段 B → 循环
```

**核心原则**：
- 所有文字必须是可编辑文本框
- 所有图片必须是独立图片对象（不嵌入背景）
- 所有表格必须是可编辑 PPT 表格（win32com `AddTable`）
- 装饰元素（渐变条、彩色背景）留在截图背景中，不重建
- 中文字体：同时设置 `Font.NameFarEast = "Microsoft YaHei"` + `Font.Name = "Microsoft YaHei"`
- win32com 所有常量用数值字面量（避免 COM 启动前引用失败）：
  - `ppAlignLeft=1`, `ppAlignCenter=2`, `ppAlignRight=3`
  - `ppLayoutBlank=12`, `msoTrue=-1`, `msoFalse=0`

**已有脚本（Step3 版本，可复用）**：
- `pipeline/extract_step3_layout.py`（Playwright 坐标提取 → `_layout_manifest.json`）
- `pipeline/screenshot_step3_masked.py`（遮罩截图，表格文字透明）
- `pipeline/export_step3_hybrid_pptx_win32com.py`（数据驱动组装，支持 text/image/table）

**已知 Bug 与教训**（Step3 转换积累，C-1 ~ C-8）：

| # | 教训 | 根因与修复 |
|---|------|-----------|
| C-1 | **SlideSize 枚举陷阱** | `ppSlideSize=4` 是 35mm(810×540pt)，不是 16:9。→ 必须用 `SlideWidth=960, SlideHeight=540` |
| C-2 | **extract 与 screenshot 的 LAYOUT_CSS 必须一致** | 不同步导致坐标系偏差（1236 vs 1260px）。→ 两个脚本共享同一段 CSS |
| C-3 | **CSS padding → PPT TextFrame.Margin** | `getBoundingClientRect()` 返回 padding box，文字起始位置偏移。→ 提取 padding，映射为 `tf.MarginLeft/Top/Right/Bottom` |
| C-4 | **rgba 透明度必须 alpha blending** | 直接丢弃 alpha 导致颜色错误（#ffffff vs 实际 #92959c）。→ 按背景色混合：hero 暗底 #0d1424，普通白底 #ffffff |
| C-5 | **HTML `<br>` 换行 → PPT 需要 `\r`** | `innerText` 返回 `\n`，COM 换行需 `\r`。→ `text.replace("\n", "\r")` |
| C-6 | **文本框宽度加 8% buffer** | 浏览器与 PPT 字体渲染差异导致文字溢出。→ `buf_w = max(ew, ew * 1.08)` |
| C-7 | **win32com 常量必须用数值字面量** | 模块加载时 COM 未启动，枚举引用失败。→ `ppAlignLeft=1, ppLayoutBlank=12, msoTrue=-1` |
| C-8 | **中文标点与 Python 引号冲突** | `、` 等字符混入字符串引发 SyntaxError。→ 使用转义或原始字符串 |

---

## PPT 自检四步法（阶段 B 强制执行）

> 提炼自 Step3 转换经验。每次 export 后必须执行，用户不需要手动触发。

**① 坐标抽样验证**
随机抽取 3-5 页，对比 `_layout_manifest.json` 中的坐标与 PPT shape 实际位置，偏差 > 3pt 则标记为严重问题。

**② 可编辑性验证**
检查每页：文本框是否可点击编辑、图片是否可选中替换、表格是否可编辑单元格。

**③ 视觉对照**
打开 PPT，逐页与 HTML 原稿截图比对：
- 文字是否覆盖在正确位置（不漂移、不重叠）
- 图片是否填充正确区域
- 表格是否对齐背景中的表格框线

**④ 内容完整性**
对照 `deck_manifest.md` 核查：每页标题、主要段落、图片数量、表格是否全部转换。

**输出 `conversion_report.md` 格式**：
```
| # | 页码 | 维度       | 问题描述           | 严重度 | 状态   |
|---|------|------------|--------------------| -------|--------|
| 1 | 03   | 坐标偏移   | KPI card 偏上 8pt  | 严重   | 已修复 |
| 2 | 13   | 表格       | P1 badge 颜色错误  | 中等   | 已修复 |
```

---

## Standalone HTML 分享包（base64 内嵌图片）

**用途**：将 `deck.html` + `images/` 打包为单一 HTML，无需附带图片文件夹即可分享。

**脚本**：`pipeline/embed_images.py`

```bash
python pipeline/embed_images.py StepN/deck.html
# 输出：StepN/deck_standalone.html（与源文件同目录）
```

**注意**：
- 输出 ≈ 原 images 总大小 × 1.33（通常 15-25 MB）
- 仅用于分享预览，不替代 `deck.html` + `images/` 作为工作版本

---

## 验证清单

**PDF**：
- [ ] 页数正确
- [ ] 每页清晰度 ≥ 4K（1260x720 * deviceScaleFactor=4）
- [ ] 中文字体渲染正常

**PPT**：
- [ ] 文字点击可编辑（每页均验证）
- [ ] 图片可移动/替换
- [ ] 表格单元格可编辑
- [ ] 与 HTML 视觉对比：文字位置偏差 < 5px
- [ ] 文件大小合理（< 25MB）
- [ ] `conversion_report.md` 已生成，无未修复的严重问题

**Standalone HTML**：
- [ ] 所有图片内嵌成功（无 MISSING 报告）
- [ ] 浏览器独立打开正常显示

---

## 防卡顿规则

- export 脚本连续失败 2 次 → 停下，说明具体报错，提出替代方案
- 单个 slide 转换超过 30 秒 → 后台运行，不阻塞主线程
- 自检发现 > 5 处严重问题 → 停下向用户说明，问是否需要重新提取坐标

---

## 交接协议

**从上游接收**：
- `deck.html`（Builder 交付的终稿，用户已确认）
- `deck_manifest.md`（页面结构清单，Builder 同步生成；包含画布尺寸、设计 token、每页标题/布局/图片引用/文字内容摘要）
- `images/`（所有图片素材）

**输出**：
- `deck.pdf` 或 `deck.pptx`（与 deck.html 同目录）
- `conversion_report.md`（PPT 自检报告，记录发现的问题和修复情况）
- `deck_standalone.html`（按需生成，单文件分享版）
