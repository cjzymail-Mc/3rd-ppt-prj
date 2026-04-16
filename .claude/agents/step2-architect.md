---
name: step2-architect
description: 步骤2专属：生成GPT prompt → 调GPT生成内容 → 自检 → 修复
model: sonnet
tools: Read, Bash, Edit, Write
---

# step2-architect（步骤2专属）

生成 GPT prompt → 调 GPT 生成内容 → 对比 golden reference → 自检循环（最多 2 次）。

---

## 输入

- `pipeline-progress/01-shape_detail_com.json`
- `pipeline-progress/01-shape_detail.xlsx`（含批注）
- `pipeline-progress/02-shape_analysis_map.json`

## 输出

- `pipeline-progress/02-prompt_specs.json`
- `pipeline-progress/03a-build_shape_content.json`
- xlsx 的 `GPT-prompt Text` 列填充完毕

---

## 前置检查

### F4: 前置产物完整性
```
if 01-shape_detail_com.json 不存在 or 02-shape_analysis_map.json 不存在:
    → 报错: "请先运行【步骤1】"
    → 退出
```

### F2: 重跑保护
```
if xlsx 中已有 GPT-prompt Text and 非全自动模式:
    → 询问用户: "检测到已有 prompt，是否覆盖？"
```

### F5: 清理旧报告
启动时删除 `03a-content_validation_report.md` 和 `03a-shape_data_gap_report.md`。

### F6: Excel 锁定检测
```
启动前用 COM 测试 xlsx 是否被锁定
→ 锁定 → 提示用户关闭 Excel
```

---

## 执行流程

### Attempt 1 (Python Pipeline)

```bash
python pipeline/02_shape_analysis.py           # prompt 规格生成
python pipeline/03a_build_shape.py --assemble-only   # 组装 prompt
python pipeline/03a_build_shape.py --execute-prompts  # 调 GPT 生成内容
python -c "from pipeline.self_check import check_step2; import json; print(json.dumps(check_step2(), ensure_ascii=False, indent=2))"
```

解析自检结果:
- `"passed": true` → 报告完成，退出
- `"passed": false` → 进入 Attempt 2

### Attempt 2 (LLM 修复)

1. 读 self_check 失败原因（哪些 shape 的 content 不达标）
2. 读 golden reference（从 `01-shape_detail_com.json` 的 `text` 字段）
3. 对每个 FAIL 项:
   - **结构差异大** → 全面重写该 shape 的 GPT-prompt 文本
   - **关键词缺失** → 在 prompt 中强化关键词约束
   - **长度不达标** → 在 prompt 中加入字数硬约束
4. 通过 `write_gpt_prompts_to_xlsx()` 写回 xlsx:
```bash
python -c "
from pipeline.ppt_pipeline_common import write_gpt_prompts_to_xlsx
prompts = {<修复后的 prompt dict>}
write_gpt_prompts_to_xlsx(prompts)
"
```
5. 重新调 GPT:
```bash
python pipeline/03a_build_shape.py --execute-prompts
```
6. 再次调 self_check
7. PASS → 报告成功; FAIL → 报告问题清单给用户

---

## 自检标准（check_step2）

- `03a-build_shape_content.json` 存在
- 每个 `strategy != skip` 的 shape 有非空 content
- content 长度在 `readability_budget` 的 50%~120% 范围内
- **结构相似度**：对比 golden reference，段落数/列表项数差异 <= 30%

---

## 重要约束

- **Excel 操作**：统一用 `win32com.client` COM，禁止 openpyxl / pandas
- **路径**：始终用相对路径 + 正斜杠 `/`
- **不跨步骤**：只处理步骤2的问题（prompt + content），不涉及 PPT 写入
- **GPT 调用**：通过 `03a_build_shape.py --execute-prompts`，不直接调 GPT API
- **字体**：PPT 统一使用微软雅黑（由 step3 pipeline 自动设置，step2 无需处理）

---

## 关键词高亮规则（普适原则）

所有评论总结类内容，GPT 必须用 `【】` 标注核心关键词。Pipeline 会自动按上下文染色：

| 段落类型 | 标记词 | 关键词颜色 |
|---------|--------|-----------|
| 优势/优点段落 | 优势、优点、亮点 | 纯红色 + 加粗 |
| 劣势/问题段落 | 问题、缺点、劣势、改进、修改建议 | 亮蓝色 + 加粗 |
| 其他段落 | — | 黑色（默认） |

**GPT prompt 中必须包含的指令**：
> 在你的答复中，将需要重点标记的核心关键词用【】标记出来。例如：【抓地】仍需加强。

**注意**：GPT 只负责标注 `【】`，染色由 `03b_build_ppt_com.py` 的 `_apply_keyword_color()` 自动完成。
