# PPT Diff & Fix Report

- 状态: fail
- visual_score: 100.00%
- readability_score: 96.36%
- semantic_coverage: 0.00%
- template_shapes: 26
- target_shapes: 26
- paired: 26
- 时间: 2026-03-18T16:02:33

## Shape对比
|template|target|visual|readability|text_len|match|
|---|---|---|---|---|---|
|Straight Connector 1|Straight Connector 1|100.00|100.00|0/0|name|
|Straight Connector 2|Straight Connector 2|100.00|100.00|0/0|name|
|Straight Connector 3|Straight Connector 3|100.00|100.00|0/0|name|
|Rectangle 4|Rectangle 4|100.00|100.00|0/0|name|
|Straight Connector 5|Straight Connector 5|100.00|100.00|0/0|name|
|Straight Connector 6|Straight Connector 6|100.00|100.00|0/0|name|
|Straight Connector 7|Straight Connector 7|100.00|100.00|0/0|name|
|Straight Connector 8|Straight Connector 8|100.00|100.00|0/0|name|
|Straight Connector 9|Straight Connector 9|100.00|100.00|0/0|name|
|Straight Connector 10|Straight Connector 10|100.00|100.00|0/0|name|
|Rectangle 11|Rectangle 11|100.00|100.00|7/7|name|
|Rectangle 12|Rectangle 12|100.00|100.00|1/1|name|
|Straight Connector 15|Straight Connector 15|100.00|100.00|0/0|name|
|Rectangle 17|Rectangle 17|100.00|100.00|34/40|name|
|Rectangle 19|Rectangle 19|100.00|100.00|0/0|name|
|Straight Connector 32|Straight Connector 32|100.00|100.00|0/0|name|
|Picture 39|Picture 39|100.00|100.00|0/0|name|
|Rectangle 40|Rectangle 40|100.00|100.00|0/0|name|
|TextBox 16|TextBox 16|100.00|100.00|13/13|name|
|Straight Connector 41|Straight Connector 41|100.00|100.00|0/0|name|
|Rectangle 68|Rectangle 68|100.00|50.31|225/38|name|
|Picture 74|Picture 74|100.00|100.00|0/0|name|
|Picture 75|Picture 75|100.00|100.00|0/0|name|
|Rectangle 77|Rectangle 77|100.00|55.00|168/30|name|
|Rectangle 14|Rectangle 14|100.00|100.00|11/11|name|
|图表 44|图表 44|100.00|100.00|0/0|name|

## 修正建议

| shape | 问题 | fix_type | 建议 |
|-------|------|----------|------|
| Rectangle 68 | readability=50.31 < 95 (文本长度/行数偏差) | annotation | 调整 readability_budget 或 prompt 字数约束 |
| Rectangle 77 | readability=55.0 < 95 (文本长度/行数偏差) | annotation | 调整 readability_budget 或 prompt 字数约束 |
| (全局) | 语义关键词缺失: 样本, 建议, 反馈 | annotation | prompt 模板中已要求融入这些关键词，检查 gpt_summary.md 或增强 prompt 约束 |

---

## 补充深度诊断（Reviewer LLM）

### 1. Readability FAIL 根因分析

#### Rectangle 68（score=50.31，template=225字/9行，target=38字/4行）

**模板内容**：3个负面维度（包裹性/稳定性/止滑性），每维度2~3行具体描述，含(X/3)统计，共225字。

**实际生成**：`【舒适性】整体穿着脚感舒适（1/3）【止滑性】抓地表现有提及（1/3）`

**问题**：
- **维度偏差**：模板区域是负面问题汇总（包裹卡脚/稳定性/止滑差），但GPT生成的是正面/模糊关键词（舒适/止滑），完全偏题
- **字数严重不足**：38/225 = 17%，budget设定270字但实际输出仅38字，说明`{target_chars}`未有效传递至GPT
- **`内容描述`缺失维度引导**：无用户批注，GPT只根据数据中抓到的关键词（舒适/抓地）做摘要，而非按模板风格逐维度展开

#### Rectangle 77（score=55.0，template=168字/5行，target=30字/4行）

**模板内容**：2个正面维度（止滑性-场地测试细节 + 场地感-触地回弹反馈），共168字。

**实际生成**：`【止滑性】抓地表现有提及。【舒适性】整体穿着脚感舒适。`

**问题**：
- **内容贫乏**：30/168 = 18%，与R68同样是GPT在数据稀疏下过早截断输出
- **维度替换**：场地感（触地回弹）未出现，被舒适替换；止滑性描述过于笼统

**共同根因**：当前数据只有2-3名测试者，"补充说明"字段内容稀少，GPT在`no_fabrication=true`约束下只能从有限数据中提炼，导致输出远低于`{target_chars}`目标。**`{target_chars}`可能未被正确填入gpt_summary.md的prompt**（需检查`03a_build_shape.py`的模板渲染逻辑）。

---

### 2. Semantic FAIL 根因分析（0% — 样本/建议/反馈全缺失）

**证据**：`gpt_summary.md`第18行已明确要求`结论中请自然融入：'样本'、'反馈'、'建议'`，但PPT全文无任何一处出现这三个词。

**可能原因（按概率排序）**：

1. **级联失效**（最可能）：R68/R77生成内容极短（38/30字），远不够容纳三个额外关键词。GPT在`no_fabrication`约束下，优先删减"无数据支撑"的内容，语义关键词作为"说明性词汇"被最先丢弃。
2. **`{extra}`占位符为空**：xlsx `备注`字段空白，`{extra}`为空字符串，可能造成注意事项列表格式异常，导致第18行在GPT视角中权重降低。
3. **模板渲染异常**：如果`03a_build_shape.py`未正确用`gpt_summary.md`构建prompt（而是直接用`02-prompt_specs.json`中的`instruction`字段），则第18行根本不会发送给GPT。

---

### 3. 精准修复建议（fix_type=annotation）

#### 操作：在xlsx `备注`列填写以下内容

| Shape | 备注建议（填入xlsx "备注"字段） |
|-------|-------------------------------|
| **Rectangle 68** | `必须覆盖3个维度（包裹性、稳定性、止滑性），每维度2~3行引用测试者具体反馈并标注(X/3)统计，总字数不少于200字；文中必须出现"样本"（如"本次3名样本"）和"反馈"两词；末尾必须给出改进建议（含"建议"二字）` |
| **Rectangle 77** | `必须覆盖2个维度（止滑性含室内外场地细节、场地感含触地回弹描述），每维度2~3行具体描述，总字数不少于150字；文中必须出现"反馈"或"样本"；末尾须含"建议"` |

#### 可选代码排查（fix_type=code，低优先级）

- 检查`03a_build_shape.py`的`gpt_prompted`分支：确认`{target_chars}`是否实际填入了`budget.max_chars`（270/201），而非某个默认小值或空值。若该变量未正确传递，`备注`修复后仍可能输出过短。

---

### 4. 验收结论

**FAIL**

| 层级 | 得分 | 阈值 | 状态 |
|------|------|------|------|
| Visual | 100.00% | ≥ 98% | PASS |
| Readability | 96.36% | ≥ 95% | **注意**：总分虽达标，但R68(50.31)和R77(55.0)两个shape严重不达标，被其他全满分shape拉高了均值 |
| Semantic | 0.00% | = 100% | **FAIL** |

> Readability总分96.36%貌似达标，但这是26个shape的均值——R68和R77的50/55分被其余24个100分shape稀释。单独看这两个shape，内容质量不可接受。

**必须修复**：
1. xlsx `备注`字段补充R68/R77的维度和字数约束（见上表）
2. 重跑Pipeline（02b→02→03a→03b→04）
