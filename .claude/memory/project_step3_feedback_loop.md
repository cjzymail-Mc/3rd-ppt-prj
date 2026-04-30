---
name: Step3→Step2 Feedback Loop
description: 2026-04-10 — Step3 content issues auto-loop back to step2, Excel prompt sync pre-check
type: project
originSessionId: cbe4f442-a918-46f9-a462-df9bde5dcbaf
---
## Step3→Step2 反馈循环

Step3 自检发现内容级严重问题（超长、SSIM 差异大）时，自动保存反馈到 `03-feedback_to_step2.json`，orchestrator 回退到 step2 重新生成内容，再重跑 step3。最多循环 1 次。

**Why:** Step3 无法缩短 GPT 生成的超长内容，必须由 step2 用更严格的 budget 约束重新调 GPT。用户不想手动来回切换步骤。
**How to apply:** `_is_content_issue()` 区分内容级 vs 格式级问题。`_apply_step3_feedback()` 向 prompt 注入字数硬约束。反馈文件消费后自动删除。

## Excel Prompt 同步预检

Step3 启动前，`_sync_excel_prompts()` 对比 Excel `GPT-prompt Text` 列 vs `03a-pending_prompts.json`，不一致则自动补跑 `03a --execute-prompts`。

**Why:** 用户在 Excel 中编辑 prompt 后直接选 step3，期望改动自动生效。JSON 不会自动同步 Excel 编辑。
**How to apply:** 这是 step3 的 Phase 0 预检，在 pipeline 和自检之前执行。用户只需：改 Excel → 选 step3 → 全自动。
