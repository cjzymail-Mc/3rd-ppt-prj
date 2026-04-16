---
name: plan3 3+1 Architecture
description: 2026-04-09 restructure — step-based agents with local self-check loops, orchestrator as thin dispatcher
type: project
originSessionId: cbe4f442-a918-46f9-a462-df9bde5dcbaf
---
# plan3: 3+1 Agent Architecture (2026-04-09)

Restructured from global iteration loop (plan2) to local per-step self-check loops.

**Why:** Global iteration was inefficient (one failure reruns everything). User wanted each step to be self-contained with its own agent.

**How to apply:**
- 3 main agents: step1-analyzer, step2-architect, step3-builder
- 1 auxiliary: curator (via /curator slash command only)
- Orchestrator is a thin menu + agent dispatcher (~769 lines, down from ~1700)
- Each agent internally runs: Attempt 1 (Python pipeline) -> self-check -> Attempt 2 (LLM fix)
- Max 2 attempts per step
- Developer agent deleted (code repairs done in Claude Code main conversation)
- pipeline/self_check.py provides check_step1() and check_step2()
- Archived: 02b_iteration_setup.py, 04_shape_diff_test.py, old agents (01-04) in _archive/
