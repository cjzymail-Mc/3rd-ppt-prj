# Plan: 6-Agent 协作构建 PPT 生成脚本 + 迭代验证

## Context

### 用户需求总结
1. **6个 Agents 构建 Python 代码**：通过 architect→tech_lead→developer↔tester 循环，让生成的 py 代码能精确复制模板 PPT 的格式/风格（内容来自 Excel 源数据）
2. **代码生成 PPT**：稳定后，用产出的 py 代码一次性精确生成 PPT
3. **多文件开发阶段**：保留多个 py + JSON/MD 文件便于调试；通过用户最终评判后再考虑整合到 main.py + src/
4. **未来可扩展**：此工作流稳定后，可套用到更多模板

### 核心洞察
- **迭代在 Agent 层**：Agents 修改 Python 代码 → 代码生成 PPT → Tester 检查 → 不合格则 Developer 改代码
- **Python 代码是确定性的**：写好后一次运行精确产出，跟 main.py 一样
- **diff_result.json + fix-ppt.md**：是 Agent 迭代的依据，Tester 产出给 Developer 看

### "复制模板"的含义
- 格式100%保留（字体、颜色、位置、样式、chart类型）
- 内容全新构建（来自 Excel 源数据 `2025 数据 v2.2.xlsx`）
- 模板 = `src/Template 2.1.pptx` 第14页(空白) + 第15页(标准)

---

## 整体架构

```
┌─────────────────────────────────────────────────────┐
│                 orchestrator.py                       │
│  @arch → @tech → @dev ↔ @test (循环) → @opti → @sec │
└───────────┬─────────────────────────────┬────────────┘
            │ Agents 写/改代码             │ Agents 读测试结果
            ▼                             ▼
┌─────────────────────────────────────────────────────┐
│              pipeline/ (5个步骤脚本)                   │
│                                                       │
│  01_shape_detail.py ──→ shape_detail_com.json         │
│         ↓                                             │
│  02_shape_analysis.py ──→ shape_analysis_map.json     │
│         ↓                    prompt_specs.json         │
│         ↓                    readability_budget.json   │
│  03_build_shape.py ──→ build_shape_content.json       │
│         ↓                prompt_trace.json             │
│  03_build_ppt_com.py ──→ codex X.Y.pptx              │
│         ↓                  post_write_readback.json    │
│  04_shape_diff_test.py ──→ diff_result.json           │
│                             fix-ppt.md                 │
└─────────────────────────────────────────────────────┘
            │
            ▼
     用户检查 PPT → 通过 → 整合到 main.py + src/
```

---

## 文件结构

```
project root/
├── pipeline/                           # 新建 - Agent产出的脚本集
│   ├── __init__.py                     # 空文件
│   ├── ppt_pipeline_common.py          # 共享工具（我先创建基础版）
│   ├── 01_shape_detail.py              # Step 1: 模板shape识别
│   ├── 02_shape_analysis.py            # Step 2: 角色推断 + prompt规格
│   ├── 03_build_shape.py              # Step 3A: 内容构建（策略矩阵）
│   ├── 03_build_ppt_com.py            # Step 3B: 模板克隆 + COM写入
│   └── 04_shape_diff_test.py          # Step 4: 三层差异测试
├── src/
│   ├── __init__.py                     # 新建（空文件，使 src 成为 package）
│   ├── Function_030.py                 # 不修改 - 复用 GPT_5/extract_info 等
│   ├── Class_030.py                    # 不修改
│   ├── Global_var_030.py               # 不修改
│   └── Template 2.1.pptx              # 不修改
├── 2025 数据 v2.2.xlsx                # 不修改
├── main.py                            # 不修改（未来整合才改）
├── orchestrator.py                     # 已配置好（上次会话改造完成）
├── new-ppt-workflow.md                 # v4.0 执行规范
│
│   # --- 以下为脚本运行时产出（不预先创建） ---
├── shape_detail_com.json              # Step 1 产出
├── shape_fingerprint_map.json         # Step 1 产出
├── shape_analysis_map.json            # Step 2 产出
├── prompt_specs.json                  # Step 2 产出
├── readability_budget.json            # Step 2 产出
├── build_shape_content.json           # Step 3A 产出
├── prompt_trace.json                  # Step 3A 产出
├── content_validation_report.md       # Step 3A 产出
├── shape_data_gap_report.md           # Step 3A 产出
├── codex X.Y.pptx                     # Step 3B 产出（最终PPT）
├── build-ppt-report.md                # Step 3B 产出
├── post_write_readback.json           # Step 3B 产出
├── diff_result.json                   # Step 4 产出（Agent迭代依据）
├── fix-ppt.md                         # Step 4 产出（修复建议）
└── diff_semantic_report.md            # Step 4 产出
```

**JSON/MD 输出到项目根目录**（不在 pipeline/ 内），原因：
- orchestrator.py 的 `_check_bug_report()` 已配置在根目录查找 `diff_result.json`
- Agent 直接在根目录读取中间产物，无需路径推理

---

## 我现在创建什么 vs Agents 后续做什么

### 我创建（基础设施 + 初始脚本框架）

| 文件 | 内容 | 行数 |
|------|------|------|
| `src/__init__.py` | 空文件 | 0 |
| `pipeline/__init__.py` | 空文件 | 0 |
| `pipeline/ppt_pipeline_common.py` | 共享工具（路径、COM安全、Excel读取、指标提取、文本截断、导入Function_030） | ~220 |
| `pipeline/01_shape_detail.py` | Step 1 完整实现（基于 codex-legacy2 + bug修复） | ~130 |
| `pipeline/02_shape_analysis.py` | Step 2 完整实现（基于 codex-legacy2 + 增强） | ~150 |
| `pipeline/03_build_shape.py` | Step 3A 完整实现（策略矩阵 + GPT prompt） | ~210 |
| `pipeline/03_build_ppt_com.py` | Step 3B 完整实现（模板克隆 + COM写入） | ~140 |
| `pipeline/04_shape_diff_test.py` | Step 4 完整实现（三层评分算法） | ~210 |

**总计约 ~1060 行**，基于 codex-legacy2 代码修复+增强，包含4个已知 bug 的修复。

### Agents 后续迭代

Agent 通过 orchestrator.py 循环：
1. **Tester** 运行 Step 1-4，检查 `diff_result.json`
2. 不达标 → **Developer** 读 `fix-ppt.md`，修改对应脚本
3. 再次运行 Step 1-4 → 再检查
4. 直到 Visual≥98, Readability≥95, Semantic=100

**Agents 可能修改的内容**：
- `02_shape_analysis.py` 的角色推断规则（如果初始启发式不准确）
- `03_build_shape.py` 的 GPT prompt 措辞（如果生成内容不够精准）
- `03_build_ppt_com.py` 的 shape 查找逻辑（如果 name 不匹配）
- `04_shape_diff_test.py` 的评分权重（如果过严/过松）

---

## 各脚本技术方案

### `ppt_pipeline_common.py` — 共享工具

**来源**：codex-legacy2/ppt_pipeline_common.py（216行），修复后版本。

**关键修复**：
1. `ROOT` 路径：`Path(__file__).parent.parent`（指向项目根，非 pipeline/）
2. `load_legacy_functions()`：改用 `from src.Function_030 import ...`（创建 `src/__init__.py` 后可用）
3. 新增 `is_in_group(shp)`：try-except 替代 `getattr(shp, "ParentGroup", None)`
4. `load_excel_rows()`：添加模糊 sheet 匹配（`'问卷' in sheet.name`）

**保留不变**：`now_ts()`, `safe_text()`, `numeric()`, `to_rows()`, `com_get()`, `com_call()`, `write_md()`, `write_json()`, `extract_metrics()`, `extract_score_means()`, `clamp_text()`

### `01_shape_detail.py` — Step 1: 模板 shape 识别

**任务**：对比 Template 第14页(空白) vs 第15页(标准)，提取新增 shape 属性 + 指纹。

**每个 shape 提取**：name, left, top, width, height, shape_type, text, font_name, font_size, has_chart, in_group, z_order

**指纹**：shape_type + rounded(geometry) + name + text_prefix(20) + z_order + in_group

**输出**：`shape_detail_com.json`（shape属性列表）, `shape_fingerprint_map.json`

**COM安全**：
- `has_chart`: `try: bool(shp.HasChart) except: False`
- `in_group`: `try: shp.ParentGroup; True except: False`
- `text`: `try: shp.TextFrame.TextRange.Text except: ""`
- 整体用 try-finally 确保 `app.Quit()`

### `02_shape_analysis.py` — Step 2: 角色推断 + Prompt 规格

**任务**：读取 shape_detail_com.json + Excel 数据，为每个 shape 分配角色、生成 prompt 规格和可读性预算。

**角色推断优先级**：
1. has_chart → chart
2. name 含 "title" → title
3. text 含 "样本/人数/N=" → sample_stat
4. text 含 "建议/结论/总结" → insight
5. len(text) ≤ 18 且非空 → title
6. len(text) ≥ 80 → long_summary
7. 默认 → body

**可读性预算**：`max_chars = max(18, min(240, template_len * 1.2))`；max_lines: title=1, sample_stat=1, insight=4, 其余=6

**输出**：`shape_analysis_map.json`, `prompt_specs.json`, `readability_budget.json`

### `03_build_shape.py` — Step 3A: 按策略矩阵构建内容

**策略矩阵**：

| 角色 | 策略 | GPT? | 数据来源 |
|------|------|------|---------|
| title | 模板原文直出 | 否 | shape_info["text"] |
| sample_stat | 样本量聚合 | 否 | f"有效样本 N={count}" |
| chart | 评分均值提取 | 否 | parse_survey_data() → 列均值 |
| body | extract_info() → GPT fallback | 条件 | 问卷原始数据 |
| long_summary | 模板锚点 + 数据驱动 GPT | 是 | 模板文本 + 统计指标 |
| insight | 模板锚点 + 行动建议 GPT | 是 | 模板文本 + 统计指标 |

**GPT 调用**：复用 `src/Function_030.py` 的 `GPT_5(prompt, model)`
**数据提取**：复用 `parse_survey_data()`, `extract_info()`
**硬截断**：`clamp_text(content, max_chars, max_lines)` 强制执行预算
**Fallback**：每个 GPT 角色有硬编码的兜底文本

**输出**：`build_shape_content.json`, `prompt_trace.json`, `content_validation_report.md`, `shape_data_gap_report.md`

### `03_build_ppt_com.py` — Step 3B: 模板克隆 + COM 写入

**流程**：
1. 打开 PowerPoint，打开 Template
2. 克隆第15页到新 Presentation
3. 读取 `build_shape_content.json`
4. 按 shape name 精确查找（fallback 用几何位置匹配）
5. 文本 shape：`TextFrame.TextRange.Text = content`（不碰格式！）
6. 图表 shape：`Chart.ChartData` → 写入 worksheet → `SetSourceData()`
7. 写后回读验证
8. 保存为 `codex {version}.pptx`

**COM安全**：
- Copy/Paste 后 `time.sleep(delay)`
- Chart workbook 写完后 `wb.Close(False)`
- try-finally 确保资源释放

**输出**：`codex {version}.pptx`, `build-ppt-report.md`, `post_write_readback.json`

### `04_shape_diff_test.py` — Step 4: 三层差异测试

**对比对象**：Template 第15页 vs codex.pptx 第1页

**三层评分**：

| 层级 | 阈值 | 检查项 |
|------|------|--------|
| Visual ≥ 98 | left/top/width/height 偏移(w=10/10/8/8, 容差20px), shape_type(w=8), chart_type(w=16), font(w=4+4) |
| Readability ≥ 95 | 文本相似度(w=0.6, SequenceMatcher), 长度比(w=0.25), 行数比(w=0.15) |
| Semantic = 100 | 关键词覆盖: ["样本", "建议", "反馈", "评分"] |

**Shape 匹配**：按 name 匹配（非按索引），fallback 用几何位置

**输出**：
- `diff_result.json`：结构化评分（Agent 读取判断通过/不通过）
- `fix-ppt.md`：逐 shape 差异明细 + 修复建议（Agent 读取修复代码）
- `diff_semantic_report.md`：语义覆盖分析

---

## 已知 Bug 修复清单

| # | Bug | 来源 | 修复方案 |
|---|-----|------|---------|
| 1 | `getattr(shape, "ParentGroup", None)` 抛 COM 异常 | debug/Mc-debug.md | `is_in_group()` + try-except |
| 2 | `Function_030.py` 相对导入失败 | codex-legacy2 | 创建 `src/__init__.py` + `from src.xxx import` |
| 3 | Chart 数据 `split(":")` 对含冒号值失败 | codex-legacy2 分析 | `rsplit(":", 1)` + workbook 关闭 |
| 4 | Diff 测试按 shape 索引匹配（重排序则错配） | codex-legacy2 分析 | 按 name 匹配 + 几何位置 fallback |

---

## 使用方式

### 方式一：手动逐步运行（调试阶段）
```bash
python pipeline/01_shape_detail.py                          # Step 1
python pipeline/02_shape_analysis.py                        # Step 2
python pipeline/03_build_shape.py                           # Step 3A
python pipeline/03_build_ppt_com.py --version 1.0           # Step 3B
python pipeline/04_shape_diff_test.py --target "codex 1.0.pptx"  # Step 4
```

### 方式二：Agent 驱动（正式构建）
```bash
python orchestrator.py
> 根据 new-ppt-workflow.md，运行 pipeline/ 脚本生成问卷分析PPT，并通过三层差异测试
```

orchestrator.py 会自动：
1. Architect 确认 PLAN.md
2. Developer 运行 pipeline/ 脚本
3. Tester 检查 diff_result.json（Visual≥98, Readability≥95, Semantic=100）
4. 不通过 → Developer 读 fix-ppt.md 修改脚本代码 → 重新运行
5. 通过 → Optimizer + Security

### 方式三：未来整合（用户评审通过后）
```python
# main.py 中新增调用：
from pipeline.build_questionnaire_page import run_pipeline
mc_slide = run_pipeline(mc_sht, mc_ppt, mc_slide, sample_name, mc_model)
```

---

## 实施顺序

```
Phase 1: 基础设施（2个空文件）
  1. 创建 src/__init__.py
  2. 创建 pipeline/ 目录 + __init__.py

Phase 2: 共享工具
  3. 编写 pipeline/ppt_pipeline_common.py

Phase 3: 步骤脚本（按依赖顺序）
  4. 编写 pipeline/01_shape_detail.py
  5. 编写 pipeline/02_shape_analysis.py
  6. 编写 pipeline/03_build_shape.py
  7. 编写 pipeline/03_build_ppt_com.py
  8. 编写 pipeline/04_shape_diff_test.py

Phase 4: 冒烟测试
  9. 验证导入：python -c "from src.Function_030 import GPT_5; print('OK')"
  10. 运行 Step 1 确认 shape_detail_com.json 产出正确
```

**Phase 2-3 总计约 1060 行代码**，基于 codex-legacy2 修复+增强。

Phase 4 之后，用户可以选择：
- 手动逐步运行调试
- 用 orchestrator.py 让 Agents 接管迭代
