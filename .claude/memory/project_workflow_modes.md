---
name: Workflow Modes Detail
description: Cold start vs hot iteration flow diagrams, mixed mode comparison table, version tracking scheme
type: project
---

## 冷启动流程（选项 0）

```
[Analyst] Pipeline(01+01b) + LLM 增强批注
    ↓
  ⏸️ P1 — 用户校准批注
    ↓
[Builder] 02 → 03a Phase1 → ⏸️ Prompt Review → 03a Phase2 → 03b → PPT
    ↓
  ✅ 初始化完成
```

## 热迭代流程（选项 1-4）

```
[跳过 Analyst LLM]
    ↓
⏸️ PROMPT REVIEW → 03a Phase2 → 03b → PPT          ← 首轮
    ↓ (选项2-4)
[Reviewer] 04验收 → PASS/FAIL
    ↓ FAIL
  02b(sheet-only) → Builder LLM(改prompt) → ⏸️ → 03a Phase2 → 03b
    ↓
  循环至 max_rounds
```

## 混合模式对照

| Agent | 冷启动（选项0） | 热迭代（选项1-4） |
|-------|---------------|-----------------|
| Analyst | 01+01b + LLM增强 | 跳过 LLM（仅确保JSON） |
| Builder首轮 | 02→03a(full)→03b | prompt review → 03a Phase2 → 03b |
| Builder修正轮 | — | 02b(sheet-only) → LLM改prompt → 03a Phase2 → 03b |
| Reviewer | 不进入 | 04测试 + LLM prompt级建议 |
| Developer | — | 读报告+修代码（code类问题时） |

## 版本追溯

| 轮次 | xlsx Sheet | PPT 文件 |
|------|-----------|----------|
| 首轮 | Shape Detail | claude-ppt 1.0.pptx |
| 第2轮 | claude-ppt 1.1 | claude-ppt 1.1.pptx |
| 第3轮 | claude-ppt 1.2 | claude-ppt 1.2.pptx |

## fix_type 分流（5 类）

| fix_type | 含义 | 后续动作 |
|----------|------|---------|
| `keyword_missing` | 语义关键词缺失 | Builder LLM 在 prompt 中追加关键词要求 |
| `budget_overflow` | 文本过长 | Builder LLM 在 prompt 中追加字数上限 |
| `budget_underflow` | 文本过短/空白 | Builder LLM 在 prompt 中要求充实内容 |
| `style_mismatch` | 格式/语调偏离 | Builder LLM 在 prompt 中追加风格约束 |
| `code` | pipeline代码缺陷 | Developer修代码 → Builder重跑 |

> orchestrator 路由逻辑：`code` → Developer，其余全部 → Builder LLM(直接改prompt)
