---
name: Stability Priority
description: User explicitly prefers deterministic Python code over LLM-based delegation for workflow orchestration
type: feedback
---

User values stability and predictability in workflow control. Fixed Python orchestrator preferred over LLM-based task delegation.

**Why:** User stated "写死的 py代码好处就是稳定可靠（毕竟稳定性最重要）". Past experience shows deterministic code prevents LLM "brain farts" from disrupting critical pipeline execution.

**How to apply:** When proposing architectural changes, always preserve deterministic Python control for workflow logic, pipeline execution, version management, and state tracking. Only use LLM agents for semantic judgment tasks (annotation enhancement, prompt rewriting, failure analysis, code fixes). Never suggest replacing Python control flow with LLM delegation.
