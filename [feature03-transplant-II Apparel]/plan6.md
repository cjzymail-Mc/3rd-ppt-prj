# Plan: Per-Agent Memory 机制评估（验证专家说法）

## Context

**用户背景**：用户在 auto-memory junction 化完成后，询问"是否要再做 per-agent memory 机制"。专家给的说法（原文）：

> 在项目根目录建`.claude/agents/`文件夹，每个子文件夹放一个专用 Agent（有自己的 prompt 和工具权限）……每个子 Agent 有独立上下文+专属记忆，不会互相污染。

**用户真实需求**：仅验证专家说法是否准确，**没有具体痛点**。不需要架构改动。

**目标**：给出权威的事实核实报告，明确"哪些说法是真"、"哪些是误解"，并落地一份 reference memory 供未来查阅（避免再次重新调研）。

## 事实核实

### 专家说法逐条核对

| 专家说法 | 是否准确 | 实际情况 |
|--|--|--|
| 项目根目录建 `.claude/agents/` 文件夹 | ✅ 准确 | 本项目已有该目录，含 5 个 agent：curator/developer/step1-analyzer/step2-architect/step3-builder |
| 每个子 Agent 有自己的 prompt | ✅ 准确 | 每个 agent 一个 `.md` 文件，含 YAML frontmatter（`name`/`description`/`model`/`tools`）+ 自定义 system prompt |
| 每个子 Agent 有自己的工具权限 | ✅ 准确 | frontmatter 的 `tools:` 字段直接列举允许的工具（如 `Read, Bash, Glob, Grep`） |
| 主 Agent 自动 Delegate 给子 Agent | ✅ 准确 | 主 Claude 通过 Agent 工具 + `subagent_type` 参数调用子 agent |
| 每个子 Agent 有独立上下文 | ✅ 准确 | 每次 Agent 调用启动一个全新隔离的 conversation，不继承父对话历史 |
| **每个子 Agent 有"专属记忆"，不会互相污染** | ❌ **误解** | Claude Code **没有内置** per-agent memory 机制。所有 agent 都从同一份用户级 auto-memory（已 junction 到 `.claude/auto-memory/`）和项目级 memory（`.claude/memory/`）读取 |
| 多 Agent 之间互相 @ / 并行 | ⚠️ 半真 | Agent 工具支持并行调用，但"互相 @"不是内置机制——子 agent 不能直接调用其他 agent |

### 关键澄清：上下文隔离 ≠ 记忆隔离

专家把两件事混在一起说："独立上下文+专属记忆"。实际上：

| 维度 | Claude Code 实际行为 |
|--|--|
| **上下文（conversation）隔离** | ✅ 每次 Agent 调用都是新对话，完全隔离 |
| **System prompt 隔离** | ✅ 每个 agent 有自己的 `.md` 定义 |
| **Tool access 隔离** | ✅ frontmatter 声明的工具子集 |
| **文件系统/memory 隔离** | ❌ **不隔离**——所有 agent 共享同一个 `.claude/memory/` 和 `.claude/auto-memory/` 目录 |

**意味着**：如果一个 agent 把洞察写入 `.claude/auto-memory/MEMORY.md`，那条索引会出现在所有其他 agent 的 system prompt 里（包括主 Claude）。"记忆污染"在 Claude Code 默认机制下**确实可能发生**，但这通常被视为特性而非 bug——"经验跨 agent 共享"是合理默认。

### 项目实际现状（已经验证可工作的设置）

5 个 agent 共享 memory，没有问题，因为：

1. **memory 文件是按需 Read 的**：agent 只在 system prompt 引用了某文件时才会去读，不是全部加载
2. **职责通过 prompt 隔离**：每个 agent 的 system prompt 明确说"你做什么、不做什么"
3. **当前引用关系健康**：
   - `developer.md` → 引用 `feedback_conclusion_coloring.md`（合理：移植 PPT 时确实需要）
   - `step3-builder.md` → 引用 `reference_pipeline_repair.md`（合理：构建失败时诊断）
   - 其他 agent 不显式引用 memory，按需找

### 如果 Claude Code 没有 per-agent memory，能不能"约定层"手工搭？

可以——但纯属约定，无系统支持。需要每个 agent 的 prompt 里手工写 "你只能 Read `.claude/agents/<name>/memory/` 下的文件"，其他 agent 写入 memory 时也手工选择目录。**复杂度高，收益低**——除非有具体的污染痛点。

## 推荐方案

**不动当前架构**。理由：

1. 专家说法的核心结论（"agent 间不污染"）在 Claude Code 默认机制下**已经成立**——通过上下文/prompt/tool 三重隔离，只是机制不是"独立 memory"
2. 用户**没有具体痛点**，没有"某 agent 看了不该看的 memory 导致犯错"的实例
3. 加一层"约定式 per-agent memory"会增加心智负担（每次写 memory 时多一个"放哪个 agent 子目录"的决策），ROI 极低
4. 现有 5 agent + 共享 memory 的设置已经稳定运行，没有迹象表明需要重构

## 实施动作

仅 1 个：**写一份 reference memory，固化本次分析结论**，避免未来重新调研。

### 新建文件

`.claude/memory/reference_agent_memory_design.md`（手工 curator 层 memory）

**内容大纲**：
- Claude Code 的 5 维 agent 隔离表（上下文 / prompt / tool / 文件系统 / memory 各自的隔离粒度）
- "上下文隔离 ≠ 记忆隔离"的常见误解
- 当前项目 5 agent 的引用关系图（各 agent 引用了哪些 memory 文件）
- 何时考虑约定式 per-agent memory（具体痛点列表）+ 当前结论"不需要"

**归属判断**（按 `mc-update.md` 边界规则）：
- 这是"架构设计决策档案"，低频按需查询
- → 进 `.claude/memory/`，**不进** `.claude/auto-memory/`
- 不更新 `.claude/auto-memory/MEMORY.md`（不该自动加载到每次会话）

### 索引更新

- `.claude/memory/MEMORY.md` 加一行：
  ```
  - [reference_agent_memory_design.md](reference_agent_memory_design.md) — Claude Code agent 隔离机制 5 维分析 + per-agent memory 决策档案（不引入）
  ```
- `.claude/CLAUDE.md` Section 6 不动（这是低频查询，不进路由表，避免 Section 6 膨胀）

## 关键文件路径

| 文件 | 动作 |
|--|--|
| `.claude/memory/reference_agent_memory_design.md` | 新建（约 80-100 行） |
| `.claude/memory/MEMORY.md` | 编辑（加一行索引） |

## 复用现有内容

- `.claude/agents/{curator,developer,step1-analyzer,step2-architect,step3-builder}.md` — 用作"5 个 agent 引用关系图"的事实来源
- `.claude/memory/reference_3account_junction.md` — 同类型文档的格式参考（架构决策档案）
- `.claude/commands/mc-update.md` 的"auto-memory vs memory 边界规则" — 用作判断 reference 文件归属的依据

## 验证步骤

1. **内容正确性**：6 个月后读 `reference_agent_memory_design.md`，能否在 1 分钟内回答"我要不要做 per-agent memory"
2. **索引可达**：`grep "agent_memory_design" .claude/memory/MEMORY.md` 命中
3. **不污染 auto-memory**：`grep "agent_memory_design" .claude/auto-memory/MEMORY.md` 应为空
4. **不污染 CLAUDE.md**：`grep "agent_memory_design" .claude/CLAUDE.md` 应为空（除非未来用户明确要求加路由）

## 影响范围

- 不动业务代码
- 不动 5 个 agent 定义
- 不动 auto-memory 目录
- 不动 CLAUDE.md
- 仅在 `.claude/memory/` 加 1 个新文件 + 1 行索引
