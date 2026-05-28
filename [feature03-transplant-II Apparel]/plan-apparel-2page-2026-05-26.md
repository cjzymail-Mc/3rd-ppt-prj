# apparel_ppt 单页扩双页 — 行动计划

**日期**：2026-05-26（rev2，下午）
**任务**：`src/apparel_ppt.py` 当前在 `Main.py` 跑出 1 页（落在 PPT 第 12 页）→ 改为生成 2 页（落在第 13、14 页）
**用户决策已落定**：
- ✅ 第 12 页**废弃**，apparel 改成只生成 13/14（不是 12 保留 + 新增）
- ✅ 13 页新增数据全部来自源 Excel（评分 / 累计跑量 / 适宜温度 / 训练定位）
- ⏳ 技术路线（参数化 vs 双函数）待 Excel 字段定位完后再定

---

## 一、当前进展（截至 rev2）

### ✅ 已完成

#### 1. PPT 现状扫描（inspect-office-template `--active --slides "12-14"`）

产物：`debug/inspect-apparel-p1213/inspect_report.md`

关键差异：

| 维度 | Page 12（旧） | Page 13（新，数据图表型） | Page 14（新，文字型） |
|---|---|---|---|
| shape 数 | 22 | 22 | 7 |
| 圆图标签 | "版型/面料/吸湿排汗/速干"纯文字 | **带评分**「版型\n3.98/5」 | 无 |
| Chart 数 | 4 | **5**（多 Chart 63） | 0 |
| 优缺点 | 短概述+长文本同页 | 无 | **优点 / 缺点 长 bullet 列表** |
| 特有 shape | Rectangle 25 "I 面料信息" | Oval 49 "适宜温度 15~25℃"、Rounded Rectangle "累计跑量km 671"、"定位日常训练 7/9" | 仅标题+受试者信息+长文 |

**结论**：13/14 不是 12 的简单切两半，是重新设计。

#### 2. 代码盘点（Explore agent）

- `make_apparel_slide(mc_sht, mc_ppt, mc_slide, sample_name, mc_gpt, mc_model)` 第 890-925 行
- `APPAREL_SHAPES` 列表第 161-195 行（22 条 shape 元数据）
- `_TEMPLATE_SLIDE = 19` 第 102 行（硬编码 Clone 源页）
- 末页追加：`X = mc_ppt.Slides.Count + 1`，无页码参数
- 3 次 GPT 调用：TextBox 24 受试者信息、TextBox 8 优点短描、TextBox 22 缺点短描
- `Main.py:822-840` 调用点：`elif template_choice == "apparel"` 分支

#### 3. zxh p1p2 蓝本判 NULL（推翻原 plan 假设）

**重要复盘**：CLAUDE.md §5 「zxh_ppt.py：含 p1p2 模式」描述误导。实测 `ZXH_SHAPES` 只是 prompt 内的 P1/P2 文本格式（`"format": "p1p2"`），**不是双页架构拆分**——zxh 本身仍是单页。

→ **不再以 zxh 为蓝本**。apparel 双页架构需要从 0 设计。
→ 待办：CLAUDE.md §5 那行描述本轮结束前要修。

#### 4. 4 字段数据源调研（2026-05-26 已落定）

源文件：`20260521 服装试穿报告  紧身背心 2025 数据 v2.2.xlsx`
源 sheet：`服装试穿问卷--紧身背心`（10 行 × 36 列 → 表头 + **9 名受试者**）

| 字段 | 列 | 表头名 | 数据格式 | 聚合策略 |
|---|---|---|---|---|
| (A) 4 维度评分 | H / O / S / X 等 | 各子项评分 | 1-5 整数 | 复用 `_extract_means_for_category()` |
| **(B) 累计跑量 km** | **G** | `6、测试累计总跑量（km）` | 混合：`63` / `55km` / `120`（部分带 "km" 后缀） | **sum across 9 → "671km"** |
| **(C) 适宜温度** | **AD** | `适合的温度区间（体感温度）` | 枚举：`5℃~15℃` / `15℃~25℃` | **mode（最高频 bin）→ "15~25℃"** |
| **(D) 训练定位** | **AC** | `适宜的穿着场景` | 长枚举：`训练（日常慢跑...）` / `训练/竞速（都可以）` | **count(含"训练") / 9 → "7/9"** |

读取建议：`load_excel_rows(xlsx, sheet_name="服装试穿问卷--紧身背心", fuzzy_keyword="紧身背心")`，列名按上面 G/AC/AD 严格匹配。
B 列 parsing：`re.findall(r"\d+", str(val))` 取数字后求和（兼容 `"55km"` / `120` 两种）。

inspect 报告产物：`debug/inspect-apparel-xlsx/inspect_excel_report.{json,md}`

---

## 二、接下来计划

### 步骤 1：Excel inspect skill 到位（用户造）

需求清单（**用户造、我不动手撸临时脚本**）：
- 入口：`--active`（GetActiveObject Excel）+ 文件模式
- 输出：每 sheet `name / used_range / 表头列名清单 / 行数 / 前 N 行数据预览`
- 过滤：`--sheets "问卷,跑量"` 关键词模糊筛
- 产物：`inspect_excel_report.json` + `.md`
- 复用：`office-com-helpers` 的 `com_get / safe_print / 模糊 sheet 匹配`
- 不需要：chart / pivot / 公式 / formatting

### 步骤 2：扫源 Excel 定位 B/C/D 字段（我做）

跑 `inspect-excel-template --active --out-dir debug/inspect-apparel-xlsx/`，肉眼对照列名，定位：
- 累计跑量 km：哪个 sheet 哪一列？数据格式（数字 / "671km"字串）
- 适宜温度：哪一列？数据格式（"15~25" / "15-25℃"）
- 训练定位：哪一列？数据格式（"日常训练" / 多选枚举）

如果是受试者粒度的列（每个 sample 一个值）→ 聚合策略：均值？众数？比例？
如果是统计型字段（已经是 9/9 形式）→ 直接读单元格。

### 步骤 3：设计双页架构（我做，主 Claude）

两种候选：

| 方案 | 入口签名 | 优点 | 缺点 |
|---|---|---|---|
| A. 参数化分发 | `make_apparel_slide(..., page="p13" \| "p14")` | 单函数，状态共享天然（GPT 总结、Excel 句柄） | if/else 分支膨胀 |
| B. 双函数 | `make_apparel_p13_slide()` + `make_apparel_p14_slide()` | 职责清晰，SHAPES 列表独立 | 共享状态要外提；Main.py 要调两次 |

**预倾向 B**：13/14 两页结构差异大（22 shape vs 7 shape，图表型 vs 文字型），合一个函数 if/else 过深。GPT 调用错峰（13 页有评分+训练定位、14 页有长 bullet 优缺点），共享状态有限。

最终方案待 Excel 字段定位完后拍。

### 步骤 4：转 /developer 落地（我派单）

执行清单交付物（**不再回到 Plan 阶段**）：

| 项 | 内容 |
|---|---|
| A. SHAPES 列表 | `APPAREL_P13_SHAPES`（22 条）+ `APPAREL_P14_SHAPES`（7 条），坐标 + strategy + budget 全标 |
| B. 入口签名 | 按步骤 3 拍板的方案（A 或 B） |
| C. 数据源 | Excel sheet/列名 → 字段映射表（B/C/D 三个新字段） |
| D. Clone 源页 | `_TEMPLATE_P13_SLIDE` + `_TEMPLATE_P14_SLIDE`；本次源 = 用户人工做的当前 page13/14 |
| E. Main.py | `elif template_choice == "apparel"` 分支改成调两次（B 方案）或单次传 page（A 方案） |
| F. GPT prompt | TextBox 23 优点 bullet + TextBox 26 缺点 bullet（page14 的两个长文本）；评分计算复用 `_extract_means_for_category` |

### 步骤 5：交付前自检（我做）

- `ppt-visual-fidelity-check --active-a (新生成 13/14) vs --active-b (用户当前 13/14 标杆)` SSIM
- SSIM < 0.85 → 回炉调坐标 / 字号
- SSIM ≥ 0.85 → 通过

---

## 三、风险 / 注意

| 风险 | 应对 |
|---|---|
| **当前 PPT = 人工制作，非模板** | Clone 源 = 用户当前打开的 ppt 第 13/14 页本身。dev 启动前要确认用户保存了样本快照（如 `template/apparel-page13-14-template.pptx`），否则用户改 ppt 后基准漂移 |
| Chart 63 的数据源 / 系列 | step2 跑完 inspect-excel 后用 `read-selected-shape` 选中 Chart 63 取真 chart 数据 |
| 染色函数选型 | TextBox 23 / 26 长 bullet 是否要 GPT 关键词染色？按 page14 现有样本看，**有蓝/红字** → 用 `_apply_keyword_color` |
| 12 页废弃后旧用户文件兼容 | apparel_ppt 单跑时还会在 Slides.Count+1 处追加；用户旧文件可能还有 12 页样式 → dev 阶段加 `__main__` 调试入口分别跑 page13/page14 |

---

## 四、状态机

```
[✅ 步骤 1] Excel inspect skill 到位（用户已造 inspect-excel-template）
   ↓
[✅ 步骤 2] inspect-excel-template --active 跑完，B/C/D 列名已定位（G/AD/AC）
   ↓
[当前] 步骤 3：架构 A vs B 待用户拍板
   ↓
[步骤 4] 派 /developer agent 按清单落地
   ↓
[步骤 5] ppt-visual-fidelity-check SSIM ≥ 0.85
   ↓
[交付] Main.py + src/apparel_ppt.py + 视觉验收报告
```

---

## 五、待办备忘

- [ ] **CLAUDE.md §5** 移除"含 p1p2 模式"描述（实测无双页架构含义，仅 prompt 文本格式）
- [ ] **auto-memory** 写 `feedback_zxh_p1p2_not_blueprint.md`：zxh 的 p1p2 是 prompt 内 P1/P2 文本格式，不是双页架构蓝本，下次不要再扫
- [ ] 用户决策：扩 inspect-office-template 还是新建 inspect-excel-template
