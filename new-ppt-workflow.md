# New PPT Workflow（生产版 / 可直接交给 Orchestrator）

> 版本：v4.0  
> 定位：这是**执行规范**，不是建议清单。  
> 目标：基于 `Main.py + src/` 的既有能力，稳定交付 98%+ 视觉保真度的单页成品 PPT。  
> 强约束：PPT 必须用 `pywin32 + win32com.client`；Excel 必须用 `xlwings + COM API(.api)`；严禁 `python-pptx`。

---

## 0. 输入 / 输出契约

### 0.1 输入文件（必须存在）
- `src/Template 2.1.pptx`（第14页空白模板，第15页标准模板）
- `2025 数据 v2.2.xlsx`（`问卷sheet`）
- `Main.py`、`src/Function_030.py`、`src/Class_030.py`（复用能力）

### 0.2 关键输出文件
- 识别层：`01-shape-detail.md`、`shape_detail_com.json`、`shape_fingerprint_map.json`
- 分析层：`02-shape-analysis.md`、`shape_analysis_map.json`、`prompt_specs.json`、`readability_budget.json`
- 构建层：`build_shape_content.json`、`content_validation_report.md`、`prompt_trace.json`、`shape_data_gap_report.md`
- 写入层：`codex <version>.pptx`、`build-ppt-report.md`、`post_write_readback.json`
- 测试层：`fix-ppt.md`、`diff_semantic_report.md`、`diff_result.json`
- 迭代层：`iteration_history.md`

---

## 1. 统一执行入口

### 1.1 主入口
```bash
python 00-ppt.py --start-version 1.0 --max-rounds 3
```

### 1.2 分段执行
- 仅重跑 diff：
```bash
python 00-ppt.py --from-step 4 --to-step 4 --start-version 1.2 --max-rounds 1
```
- 仅重跑构建与写入：
```bash
python 00-ppt.py --from-step 3 --to-step 3 --start-version 1.2 --max-rounds 1
```

---

## 2. Step 1：shape 识别与指纹（`01-shape-detail.py`）

### 2.1 任务
对比模板第14页与第15页，识别新增 shape，并输出结构化信息：
- `name/type/left/top/width/height/z_order/in_group/text/font/has_chart`

### 2.2 指纹策略
每个 shape 生成稳定指纹（用于抗模板轻微变化）：
- `shape_type + rounded geometry + name + text_prefix + z_order + in_group`

### 2.3 验收
- `shape_detail_com.json` 中 `new_shapes` 非空（目标通常为9）
- `shape_fingerprint_map.json` 成功产出

---

## 3. Step 2：shape→源数据映射与 Prompt 规格（`02-shape-analysis.py`）

### 3.1 任务
为每个目标 shape 明确：
1) shape 角色（title/body/long_summary/insight/sample_stat/chart）  
2) 来源字段范围（header 集合）  
3) 输出约束（`max_chars/max_lines/max_bullets`）  
4) Prompt 规格（`style_anchor + instruction + output_constraints`）

### 3.2 Prompt 设计原则（强制）
- Prompt 必须来自“**标准模板文本 + 源数据**”关系，不允许套用旧 prompt 直接复用。
- 每个 shape 一条独立 prompt 规格。
- chart 角色标记为“非GPT，数据聚合生成”。

### 3.3 验收
- `shape_analysis_map.json`、`prompt_specs.json`、`readability_budget.json` 全部存在。

---

## 4. Step 3A：按 shape 角色构建内容（`03-build_shape.py`）

### 4.1 核心原则
**不同 shape 使用不同生成策略，不允许统一 GPT_5。**

### 4.2 标准策略矩阵
- `title`：模板锚点直出（非 GPT）
- `sample_stat`：问卷样本量聚合（非 GPT）
- `chart`：每项评分均值提取（非 GPT）
- `body`：优先 `extract_info()` / regex 提取；失败再 GPT
- `long_summary`：模板锚点 + 数据指标驱动的 GPT prompt
- `insight`：模板锚点 + 数据指标驱动的 GPT prompt（动作建议型）

### 4.3 chart 专项规则
- 必须由 `问卷sheet` 评分列计算每项均值（人数维度平均）
- 输出格式：`指标:均值`（供写入层 chart 更新）

### 4.4 数据缺口上报（强制）
若某 shape 在源数据中找不到必要信息，必须写入 `shape_data_gap_report.md`：
- shape 名称
- role
- strategy
- gap reason（例如：未识别评分列 / GPT fallback）

### 4.5 验收
- `build_shape_content.json` 有每个shape内容
- `content_validation_report.md` 通过预算检查
- `prompt_trace.json` 可追溯每个shape的 prompt 或生成依据
- `shape_data_gap_report.md` 存在且可读

---

## 5. Step 3B：模板克隆+内容写入（`03-build_ppt_com.py`）

### 5.1 写入流程
1. 打开 `PowerPoint.Application`
2. 新建 presentation
3. 克隆模板第15页到目标页
4. 按 shape 名称精确定位并写入

### 5.2 写入约束
- 文本 shape：仅写 `TextFrame.TextRange.Text`
- 图表 shape：仅更新 `ChartData/Series/Categories` 数据
- 不得重置字体、颜色、段落、图表样式壳

### 5.3 写后回读
每个 shape 写入后立即回读并记录：
- 是否更新成功
- 文本写入长度 vs 回读长度
- chart 类型保留情况

### 5.4 验收
- `codex <version>.pptx` 产出
- `build-ppt-report.md` + `post_write_readback.json` 产出

---

## 6. Step 4：严格差异测试（`04-shape_diff_test.py`）

### 6.1 对比对象
- 模板：`Template 2.1.pptx` 第15页
- 目标：`codex <version>.pptx` 第1页

### 6.2 三层门禁
1. **Visual**：几何、shape type、字体、颜色、chart type
2. **Readability**：文本相似度、长度比、行数比
3. **Semantic**：关键语义覆盖（样本/建议/反馈等）

### 6.3 通过标准（强制）
- Visual Score ≥ 98
- Readability Score ≥ 95
- Semantic Coverage = 100
- 且无关键 shape fail

### 6.4 输出
- `fix-ppt.md`（差异明细 + 修复建议）
- `diff_semantic_report.md`
- `diff_result.json`

---

## 7. Step 5：多轮迭代控制（由 `00-ppt.py` 执行）

### 7.1 迭代规则
- Step4 fail 时自动进入下一轮版本（1.0→1.1→1.2...）
- 每轮不覆盖历史文件
- 所有轮次写入 `iteration_history.md`

### 7.2 修复优先级
1. 先改 shape 的 strategy（是否错用了 GPT）
2. 再改 prompt（style anchor / instruction / budget）
3. 再改提取函数（extract_info / 均值提取 / regex）

---

## 8. Agent 执行职责（给 orchestrator.py 使用）

### architect
- 基于本文件输出 `PLAN.md`，明确每步产物与验收阈值。

### tech_lead
- 审核 `shape->strategy` 是否合理，禁止“全量 GPT”路线。

### developer
- 实现脚本与提取函数，确保 per-shape 策略落地。

### tester
- 运行 Step4 并产出 `fix-ppt.md`，不达标则阻断交付。

### optimizer
- 优化执行稳定性与速度，不改变视觉结果。

### security
- 审核 API key、路径、输出产物安全与清理策略。

---

## 9. 非功能要求

- 稳定性：COM 调用要有容错与资源释放。
- 可追溯：每步必须有 JSON/MD 产物。
- 可复现：同输入同版本应可重复生成相近结果。
- 可审计：prompt 与 gap 必须落盘。

---

## 10. 禁止事项

- 禁止 `python-pptx`
- 禁止在文本 shape 上重置样式
- 禁止 chart shape 改写样式壳（只改数据）
- 禁止将“找不到数据”静默吞掉（必须上报至 `shape_data_gap_report.md`）
