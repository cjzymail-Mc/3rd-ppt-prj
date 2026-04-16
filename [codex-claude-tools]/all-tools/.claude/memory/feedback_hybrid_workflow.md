---
name: feedback_hybrid_workflow
description: 用户对agent工作流的核心要求：Pipeline先行+LLM精调，不要纯LLM硬啃
type: feedback
---

Agent 应该在每个环节都先调用 pipeline 脚本，而不是从头用 LLM 硬啃。

**Why:** 纯 LLM 执行导致 Analyst 耗时 294s（逐个 shape 推理），而 pipeline 脚本秒级完成。同时纯 pipeline 又无法处理模糊/歧义场景。

**How to apply:** 每个 agent prompt 必须遵循"Phase 1 Pipeline → Phase 2 LLM 精调"模式。LLM 只在 pipeline 无法判断时介入（模糊批注、语义审核、智能修正批注）。判断标准：LLM 介入能提升精确度才保留，否则走纯 pipeline。
