---
name: project_4agent_architecture
description: 项目架构：4-Agent混合工作流，含5类fix分流和结构化output_contract
type: project
---

2026-03-17 完成 6-Agent → 4-Agent 架构重构。
2026-03-19 执行 plan1（COM稳定性+路由修复）+ plan4（prompt精度+结构化升级）。

- orchestrator.py 调度 4 个专用 Agent: Analyst / Builder / Reviewer / Developer
- 固定工作流: Analyst → PAUSE → Builder → Reviewer → [Developer条件] → 循环

**Plan1 修复 (2026-03-19):**
- COM: `parse_user_annotations()` 和 `02b` 改用 `DispatchEx` + sleep 避免实例冲突
- 路由: section 9 fallback 改用 rich prompt（不再丢失 respondent data）
- orchestrator 输出过滤器增加 ⚠️ / [WARN]
- 02 新增 guardrail: --sheet 指定但 0 批注时 abort

**Plan4 升级 (2026-03-19):**
- prompt 模板恢复 `每个分类不超过3行` 默认约束（codex_ppt.py 回归修复）
- 01b golden reference: gpt_prompted shapes 自动写完整约束描述
- Analyst/Builder prompt 追加 golden reference few-shot
- 02 新增 `_parse_output_contract()`: 从内容描述自动提取 required_keywords / bracket_highlight / ratio_required / sentiment
- 04 fix_type 从 2 类扩展为 5 类: keyword_missing / budget_overflow / budget_underflow / style_mismatch / code
- 02b 按 fix_type 定向修正（关键词/字数/风格分别处理）
- orchestrator 从 JSON 读取 fix_type，`!= "code"` 统一走 Builder 分支
- 备注字段已废弃，所有指令统一写入内容描述

**Why:** 用户需要稳定的混合工作流，pipeline 为主 + LLM 补语义空缺 + 人工只处理低置信度项。

**How to apply:** 新增/修改 agent 时遵循"Pipeline先行 + LLM精调"模式；内容描述是映射知识入口，需包含来源+方向+关键词+格式约束。
