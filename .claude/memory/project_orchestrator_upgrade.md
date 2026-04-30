---
name: Orchestrator Architecture Decision
description: 2026-04-01 decision to keep Python orchestrator, upgrade modularly instead of replacing with LLM delegation
type: project
---

User evaluated whether to replace orchestrator.py with native Claude Code agent delegation (based on Grok's suggestion).

**Decision**: Keep Python orchestrator, improve via modular refactoring.

**Why**: The orchestrator is a 1622-line workflow engine (not a simple router). It handles deterministic pipeline execution, version arithmetic, 5 interactive pause points, Windows COM integration, and complex conditional routing. LLM delegation cannot reliably replace any of these.

**How to apply**: Future work should improve orchestrator modularity (config extraction, module splitting, --agent flag upgrade, prompt templates) rather than replacing it with LLM-based coordination. Phases 0-4 defined in upgrade plan.
