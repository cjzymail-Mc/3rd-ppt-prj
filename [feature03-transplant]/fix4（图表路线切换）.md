# fix4（图表路线切换）.md — chart 路线切换计划（改模板 → 从零制表）

> **状态**：路线决策已定稿，待实施
> **前置**：fix3（图表写入诊断）.md 已充分诊断"改模板"路径在本项目生产约束下不可行
> **最后更新**：2026-04-24

---

## Context

fix3（图表写入诊断）.md 把 chart 写入 bug 的每一个坑都验证过一轮。结论层面有两个关键事实：

1. **"原地改模板 chart"的两条子实现**（COM in-place / XML surgery）在你的生产约束下**全部不可行**
2. **"从零制表 + OLE 粘贴"路线**（`Function_030.make_chart_for_questionnaire` 范式）在同一套办公环境下已稳定运行数年

所以 fix4 不再继续修 bug，而是**切换路线**：把 `yzr_ppt.py` 的 chart 写入从 `_write_chart`（改模板）改为 `make_chart_for_yzr`（从零制表）。

---

## 路线决策：为什么放弃"改模板"

### 生产约束（用户场景，Mc-debug-4.md line 1775 原文）

> "我将模板 ppt 和 py 代码分享给同事，但数据永远是他们自己提供，所以：默认情况下，ppt chart 数据源丢失是 100% 会发生的事件"

这句话隐含两个目标：

| 目标 | 含义 |
|--|--|
| A. 100% 还原模板视觉 | 原地改 chart 数据，保留所有样式 |
| B. 模板 + 代码分发给同事，同事自己填数据 | chart 内部状态必然在他人机器上漂移 |

**A 与 B 物理上不兼容**。任何"原地改 chart"的实现都要求 chart 内部状态（IsLinked / embedded workbook / numCache）在分发链路里保持稳定，而 B 目标必然破坏这个稳定性。

### 证据链（debug-4 + fix3 已积累）

| 事实 | 出处 |
|--|--|
| 同事机器 Run 1（fresh 模板）：STRAT 1-4 全失败，readback=[] | debug-4 line 1552 |
| 同事机器 Run 2（手工重建 chart 后）：STRAT 1-3 通过，但 STRAT 4 的 BreakLink 立即破坏健康 chart | debug-4 line 1713-1725 |
| BreakLink + Activate 是**凶手**，不是保护伞 | fix3 坑 2 |
| Activate 在 Build 4266 抛 DISP_E_EXCEPTION，且触发 GUI 弹窗阻塞脚本 | fix3 坑 3 |
| XML surgery 路径彻底死：办公室默认加密 pptx → CFB 容器，非 zip | fix3 坑 4，debug-4 line 1911 |
| `make_chart_for_questionnaire` 在办公室多年生产从未报错 | debug-4 line 1777 |

### 结论

- "改模板"路线对**单机自用**成立（chart 不被历史 BreakLink 污染 + 数据在本机）
- "改模板"路线对**当前分发场景**不成立，这不是可修复的 bug，是路线与需求物理不兼容
- `make_chart` 路线已有成熟范本 `make_chart_for_questionnaire`，视觉还原 95%（3D bar 样式需在 xlwings 里复刻）

**放弃 fix3（图表写入诊断）.md 阶段 1 的 fresh 模板 STRAT 1 验证**——即使通过，也不能保证同事机器，是在消耗预算做交叉验证而非生产路径。

---

## 实施清单

### 改动 1：`src/_ppt_shared.py` 新增 `make_chart_for_yzr`

**位置**：`_ppt_shared.py` 文件末尾，紧跟 `clamp_text` 之后。

**函数签名**：

```python
def make_chart_for_yzr(
    mc_cell,        # xlwings Range：指标名+均值两列数据所在区域的锚点
    mc_slide,       # 目标 PPT slide（win32com）
    Left, Top, Width, Height,  # 粘贴后的位置/尺寸（points）
):
    """为 yzr 模板构建 3D 条形图（ChartType=60），OLE 粘贴到 PPT。

    与 make_chart_for_questionnaire 的差异：
      - chart_type: 3d_bar_clustered（对应 ChartType=60）
      - 数据形状：7 指标 × 1 均值列（questionnaire 是 N 人 × M 指标）
      - 量程：固定 0~10（yzr 问卷一律 10 分制）
      - 返回：xlwings chart 对象（外层决定是否 delete）

    参考 Function_030.py:1999 make_chart_for_questionnaire 的框架。
    """
```

**内部流程**（对齐 `make_chart_for_questionnaire` 的骨架）：

```
1. mc_sht = mc_cell.sheet; mc_sht.select(); mc_cell.select()
2. 读 CurrentRegion 形状（行数/列数），确定数据块
3. mc_sht.charts.add(chart_left, chart_top, width=Width, height=Height)
4. chart.chart_type = '3d_bar_clustered'  # 或 api[0].ChartType = 60
5. chart.set_source_data(...)
6. SetElement(100)       # 隐藏图例
7. SetElement(328)       # 隐藏网格线
8. Axes(2) MinimumScale=0, MaximumScale=10
   Axes(2) TickLabelPosition/MajorTickMark/MinorTickMark = -4142
   Axes(2) Format.Line.Visible = 0
9. SeriesCollection(1).ApplyDataLabels()
10. SetElement(0)        # 隐藏主标题
11. api[0].Copy()
12. mc_slide.Shapes.Paste()
13. xlwings.apps.active.api.CutCopyMode = False  # 硬规则 #3
14. mc_shape.Left/Top = 参数传入值
15. return chart
```

**注意事项**（硬规则复用）：

- **xlPicture = -4147** 常量禁用（这里是 OLE 粘贴不是图片粘贴）
- **`CutCopyMode = False`** 必须在 Paste 后立即执行（断 OLE 热链接）
- **删行前先 delete chart**（避免 Excel "错误公式引用"弹窗；但本函数保留 chart 对象供外层决策）

### 改动 2：`src/yzr_ppt.py::make_codex_slide` chart 分支改造

**现状**（yzr_ppt.py:553-554）：

```python
if strategy == "mean_extraction" or bool(_com_get(shp, "HasChart", False)):
    _write_chart(shp, content)
```

**改造后**：

```python
if strategy == "mean_extraction" or bool(_com_get(shp, "HasChart", False)):
    # 记录模板 chart shape 的位置/尺寸
    L, T, W, H = shp.Left, shp.Top, shp.Width, shp.Height
    # 删除模板 chart shape（从零制表路线，不再原地改）
    shp.Delete()
    # 在 Excel 里建临时数据区（7 指标 × 均值），建 chart，OLE 粘贴
    mc_cell = _prepare_yzr_chart_data(mc_sht, content)  # 新增 helper
    make_chart_for_yzr(mc_cell, new_slide, Left=L, Top=T, Width=W, Height=H)
    continue
```

**新增 helper `_prepare_yzr_chart_data(mc_sht, content)`**：
- 解析 content（`"指标名:均值"` 每行一条，共 7 行）
- 在 `mc_sht` 的空白区（约 100 行下方，与 `make_chart_for_questionnaire` 共用临时区）写入 2 列：指标名 / 均值
- 返回左上角单元格（传给 `make_chart_for_yzr` 作锚点）
- 临时数据区**保留**（硬规则：OLE 嵌入 chart 保持对行号引用，删行 = PPT 图表数据消失，参考 debug-4 line 744-760）

### 改动 3：`src/_ppt_shared.py::_write_chart` 注释警告

**不删除** `_write_chart`（zxh_ppt 还在用，且单机场景仍然有效），但在 docstring 最顶部加一段警告：

```python
def _write_chart(shp, content: str) -> bool:
    """Write chart data via SeriesCollection.

    ⚠️ **适用场景警告**：
      仅限"单机自用、模板 + 数据同机"场景。
      分发场景（模板/代码发给他人，数据他人填）下 chart 内部状态会漂移，
      此函数不可靠，请改用 make_chart_for_yzr（从零制表 + OLE 粘贴）。
      参考 [feature03-transplant]/fix4（图表路线切换）.md 路线决策。

    (原 docstring 保留...)
    """
```

### 改动 4：`CLAUDE.md` 硬规则追加

在 §3 "硬规则"末尾加一条：

> **分发场景 chart 必须从零制表**：模板发给他人、数据他人填时，chart 走 `make_chart_for_{template}` 路线（xlwings 新建 + OLE 粘贴），禁用 `_write_chart` 原位改。原因见 fix4（图表路线切换）.md。

### 改动 5：固化经验到 `.claude/memory/feedback_chart_write.md`（新文件）

内容要点：
- BreakLink 是凶手不是保护伞（debug-4 line 1703-1735）
- Activate 在 Build 4266 抛 DISP_E_EXCEPTION，且触发 GUI 弹窗
- 加密办公环境下 XML surgery（zipfile）整条路径不可用（CFB 非 zip）
- chart shape 名中英文差异 + 重建后 COM 名会变（`Chart 13` → `Chart 27`）
- "100% 视觉还原"与"分发给他人+他人填数据"在物理上互斥
- 分发场景 chart 强制走 `make_chart_for_{template}` 路线

---

## 不动哪些代码

- **zxh_ppt.py**：`_write_chart` 依赖保留不动（zxh 目前还没有分发到同事机器的场景，先观望）
- **Pipeline `_write_chart`**（`pipeline/03b_build_ppt_com.py`）：保留（Pipeline 用于新模板分析，本身就是单机场景）
- **模板 pptx 文件**：不动（chart shape 会在运行时被 shp.Delete() 删掉，模板本身不改）
- **`make_chart_for_questionnaire`**：不动（已稳定在产）

---

## 视觉验收

| 要求 | 可接受阈值 |
|--|--|
| 3D bar 样式 | 与模板 chart 视觉相似度 >= 90%（允许 3D 倾角 / 柱宽 / 配色微差） |
| 量程/轴 | 0~10 固定，轴线/刻度/标签隐藏 |
| 数据标签 | 每根柱子末端显示数值（1 位小数） |
| 位置/尺寸 | 与模板 chart 的 L/T/W/H 完全一致（读取 → 删除 → 还原） |
| 跨机一致性 | 用户机器 + 同事机器运行结果一致（bars 正确、样式相同） |

---

## 风险评估

| 风险 | 概率 | 缓解 |
|--|--|--|
| 3D bar 样式调不到模板那么好看（颜色/倾角差异） | 中 | 先跑通功能，视觉再迭代；可接受 95% 相似 |
| Excel 临时数据区与 `make_chart_for_questionnaire` 的临时区冲突 | 低 | 约定 yzr 用不同起始行（e.g. 120 行），避开 questionnaire 的 100 行 |
| OLE 粘贴后 mc_shape 尺寸漂移 | 低 | 粘贴后立即覆写 L/T/W/H（现有范式已处理） |
| 同事机器 xlwings 版本差异 | 低 | xlwings 已在办公室多机稳跑多年 |
| 删除模板 chart shape 后位置信息丢失 | 零 | 先读 L/T/W/H 再删除（顺序已在改动 2 中明确） |

---

## 验收（分阶段）

### 阶段 1 — 单元验证
- [ ] `make_chart_for_yzr` 独立跑通（给定 7 条固定数据，生成 chart 并粘贴到空白 slide）
- [ ] 视觉检查：3D bar + 7 指标 + 0~10 量程 + 无图例/轴/网格

### 阶段 2 — 集成验证（用户机器）
- [ ] `python src/yzr_ppt.py` 端到端跑通
- [ ] 生成的 slide chart 位置/尺寸与模板一致
- [ ] 数据正确（与 `_extract_score_means` 输出一致）

### 阶段 3 — 跨机验证（同事机器）
- [ ] 同事拿到**未处理的模板 pptx** + 新版代码
- [ ] 同事用自己数据跑 `python src/yzr_ppt.py`
- [ ] chart bars 正确显示，样式与用户机器一致

### 阶段 4 — 知识固化
- [ ] `.claude/memory/feedback_chart_write.md` 落地
- [ ] `CLAUDE.md` §3 硬规则条目追加
- [ ] `_write_chart` docstring 警告已加
- [ ] fix3（图表写入诊断）.md 状态更新为 "superseded by fix4"（保留历史档案，不删文件）

---

## 回滚预案

如果 `make_chart_for_yzr` 在同事机器上也出问题（低概率事件）：

1. **第一步**：保留现有 `_write_chart` 调用路径不动（已经在产）
2. **第二步**：检查 xlwings 版本（`xlwings.__version__`）与生产一致
3. **第三步**：检查 Excel COM 是否正常（`make_chart_for_questionnaire` 能否跑）
4. **第四步**：若 Excel COM 本身故障，回退到"用户代跑"手工流程（非技术路径问题）

---

## 当前待办（按时间顺序）

- ⏳ 用户确认 fix4 路线决策（本文档）
- ⏳ 实施改动 1-5（分 3 个 PR，参考"实施清单"节）
- ⏳ 单元 + 集成 + 跨机 3 阶段验收
- ⏳ 固化经验到 memory 并更新 CLAUDE.md

---

## 与 fix3（图表写入诊断）.md 的关系

- fix3（图表写入诊断）.md 的诊断价值：**保留**（踩过的 7 个坑是宝贵现场数据）
- fix3（图表写入诊断）.md 的修复方向：**废弃**（阶段 1 / 阶段 2 的"改模板"路线在生产约束下不可行）
- fix3（图表写入诊断）.md 的阶段 3（候选 B：make_chart 兜底）：**升级为 fix4 主线**

fix4 实施完成后，fix3（图表写入诊断）.md 文件保留作为历史档案，但状态标记为 "superseded by fix4"。

---

## 执行结果（2026-04-24 实施记录）

### 改动落地清单

| # | 文件 | 改动 | 行号范围 |
|--|--|--|--|
| 1 | `src/_ppt_shared.py` | 新增 `_prepare_yzr_chart_data(mc_sht, content)`：解析 `"指标:均值"` 格式 content，写入 Excel 安全区（`origin + rows_count + 40` 行，避开 questionnaire 的 `+8` 行区），返回 xlwings Range 锚点 | 文件末尾新增 |
| 2 | `src/_ppt_shared.py` | 新增 `make_chart_for_yzr(mc_cell, mc_slide, Left, Top, Width, Height)`：xlwings 建 3D 条形图（`api[1].ChartType = 60`）→ 固定量程 0~10 → 隐藏轴/刻度/图例/网格/主标题 → `api[0].Copy()` + `Shapes.Paste()` → `CutCopyMode = False` → 还原 L/T/W/H | 紧跟 _prepare_yzr_chart_data |
| 3 | `src/_ppt_shared.py::_write_chart` | docstring 顶部追加 ⚠️ 适用场景警告（仅单机自用，分发场景用 make_chart_for_yzr） | docstring 开头 |
| 4 | `src/yzr_ppt.py` | 3 处 import 块补全：`make_chart_for_yzr, _prepare_yzr_chart_data`（相对 / `src.` / 扁平三种路径都更新了） | line 47-82 |
| 5 | `src/yzr_ppt.py::make_codex_slide` | chart 分支改造：<br>① 读 `L/T/W/H`<br>② `shp.Delete()` 删除模板 chart shape<br>③ `mc_cell = _prepare_yzr_chart_data(mc_sht, content)`<br>④ `_tmp_chart = make_chart_for_yzr(mc_cell, new_slide, L, T, W, H)`<br>⑤ 清理 Excel：`DisplayAlerts=False` → `_tmp_chart.delete()` → 恢复 DisplayAlerts。**临时数据行保留不删**（见"关键设计决策"第 1 条） | 原 553-554 两行扩展为约 55 行 |
| 6 | `.claude/CLAUDE.md` §3 | 追加硬规则条目："分发场景 chart 强制从零制表（fix4）"，列出禁用 `_write_chart` 原位改的原因和跳转 fix4（图表路线切换）.md | §3 硬规则末尾 |
| 7 | `.claude/memory/feedback_chart_write.md` | 新建经验固化文件，包含 6 条核心教训：BreakLink 是凶手、Activate DISP_E、XML surgery 被 CFB 加密封死、chart 名漂移、目标互斥、make_chart 路线稳定多年 | 新文件 |
| 8 | `.claude/memory/MEMORY.md` | 索引新增一行指向 `feedback_chart_write.md` | Feedback 节末尾 |

### 关键设计决策（实施时的取舍）

1. **临时数据清理 vs 保留**：最终选择**只删 chart，保留临时数据行**（修正后方案）。
   - 最初方案：仿 `make_chart_for_questionnaire` 删 chart + 删 rows。
   - 修正理由（用户实测）：即使 `CutCopyMode = False` 已执行，删 Excel 端临时数据行仍会导致 PPT 端 OLE chart 数据丢失（bars 消失）。对照 Mc-debug-4.md line 744："我决定不折腾了，直接保留临时数据、保留图表吧，优先保证 ppt 图表的稳定性"。
   - 当前实现：`DisplayAlerts = False` → `_tmp_chart.delete()` → 恢复 DisplayAlerts。Excel 端临时数据行**不删**。
   - 副作用可接受：Excel 会残留 7 行临时指标数据，略显凌乱，但保证 PPT chart 稳定。

2. **临时数据区位置**：`origin.offset(row_offset=rows_count + 40, column_offset=0)`
   - 理由：`make_chart_for_questionnaire` 用 `+ 8` 行偏移；yzr 用 `+ 40` 行留大安全间距，避免并行调用时数据区冲突

3. **3D bar 类型设置**：`api[1].ChartType = 60`（xl3DBarClustered），失败时回退 `chart_type = 'bar_clustered'`（2D）
   - 理由：xlwings 的字符串接口对 3D 类型支持不稳定，直接走 COM 层；失败回退确保不崩

4. **scope 严格限定 src/**：未动 `pipeline/03b_build_ppt_com.py::_write_chart`、未动 `orchestrator.py`
   - 理由：Pipeline 是单机分析工具，本身不是分发产物，_write_chart 缺陷不会在 Pipeline 场景暴露；scope creep 会拖延交付

5. **zxh_ppt.py 保留不动**：zxh 当前未分发，单机场景下 `_write_chart` 仍然有效
   - 理由：避免破坏已稳定运行的 zxh；将来若 zxh 要分发，可仿 fix4 模式切换

### 自检结果

| 检查项 | 方法 | 结果 |
|--|--|--|
| `_ppt_shared.py` 语法 | `python -c "ast.parse(...)"` | ✅ OK |
| `yzr_ppt.py` 语法 | `python -c "ast.parse(...)"` | ✅ OK |
| 新函数可导入 | `from src._ppt_shared import make_chart_for_yzr, _prepare_yzr_chart_data, _write_chart` | ✅ 三个符号都能拿到 |

### 待用户验证的阶段（未执行）

fix4 的"实施"已完成，但"验证"需要用户在真实 Excel + PPT 环境下执行（Claude 无 COM 访问）：

- ⏳ **阶段 1（单元）**：`python src/yzr_ppt.py` 单页调试，观察 3D bar 是否正确绘制、位置/尺寸是否与模板 Chart 13 一致
- ⏳ **阶段 2（集成）**：`python Main.py` 端到端运行，确认 slide 15 chart 符合预期
- ⏳ **阶段 3（跨机）**：同事机器（Build 4266）上跑 `python src/yzr_ppt.py` + 同事自己的问卷数据，验证 bars 正确显示、样式与用户机器一致
- ⏳ **阶段 4（定稿）**：验证通过后，将 fix3（图表写入诊断）.md 状态改为 "superseded by fix4"

### 日志锚点（便于调试）

运行时关注以下日志前缀：

- `[yzr-chart] 临时数据已写入：anchor=(行,列)，N 个指标` —— `_prepare_yzr_chart_data` 成功
- `[yzr-chart] 开始 xlwings 建 3D 条形图 → OLE 粘贴` —— 进入制表流程
- `[yzr-chart] ChartType = 60 (3D bar clustered)` —— 3D 类型设置成功
- `[yzr-chart] 3D 视图：Elevation=20, Rotation=15, RightAngleAxes=True, Depth=100, Height=100` —— 3D 旋转设置成功
- `[yzr-chart] 坐标轴已固定 0~10，轴线/刻度/标签已隐藏` —— 轴处理成功
- `[yzr-chart] 已粘贴至 PPT（L=, T=, W=, H=）` —— OLE 粘贴 + 位置还原成功

如任何一步失败，错误会带在 `[警告]` 前缀后，不会崩溃 `make_codex_slide` 整个流程。

### 3D 旋转参数（2026-04-24 补充）

**关键发现**：xlwings 建立的 3D chart 默认视角（Elevation/Rotation）**不等于** PowerPoint 模板 Chart 13 的视角。必须在 `make_chart_for_yzr` 里显式设置，才能让 OLE 粘贴到 PPT 后的视角符合用户期望。

**PPT "三维旋转" 面板 ↔ Excel chart API 映射**：

| PPT 面板字段 | Excel chart.api[1] 属性 | 取值范围 | 用户实测值 |
|--|--|--|--|
| X 旋转 | `Elevation` | -90 ~ 90 | 20 |
| Y 旋转 | `Rotation` | 0 ~ 360 | 15 |
| Z 旋转 | —（Excel 不直接暴露） | — | 0（忽略） |
| 透视 | `Perspective` | 0 ~ 100 | 0 |
| 直角坐标轴 ☑ | `RightAngleAxes` | True/False | True |
| 自动缩放 ☑ | `AutoScaling` | True/False | True |
| 深度 | `DepthPercent` | 20 ~ 2000 | 100 |
| 高度 | `HeightPercent` | 5 ~ 500 | 100 |

**调参工作流**（未来新模板适配可复用）：

1. 用户在 PPT 里手动调到满意 → 选中 chart shape
2. 读坐标：`python skills/read_selected_shape.py` → 拿到 L/T/W/H
3. 读 3D 旋转：PPT "设置形状格式 → 效果 → 三维旋转" 面板逐项抄录
4. 按映射表回写到 `make_chart_for_{name}` 里的 3D 视图参数块

**兜底**：所有 3D 属性设置都在 `try/except` 块内，失败时日志提示 fallback 到 xlwings 默认视角，不让 chart 生成崩溃。

---

## 变更影响面（给未来接手者）

- **yzr_ppt.py**：chart 分支行为发生根本变化（不再原地改，改为删除+重建），视觉差异可能存在（3D 倾角/颜色/柱宽 ≈95% 还原，非 100%）
- **zxh_ppt.py**：无影响
- **Pipeline**：无影响
- **Main.py**：无影响（只通过 `make_codex_slide` 间接使用）
- **Excel 临时数据区**：yzr 运行时会在 `origin + rows + 40` 行临时写入 2 列数据，运行结束后**保留不删**（删 Excel 端数据会导致 PPT chart 失数据；仅删除 Excel 端 chart 对象）
- **Template 2.1.pptx**：模板文件本身不修改（chart shape 在运行时 Copy 后的新 slide 上被删除重建）
