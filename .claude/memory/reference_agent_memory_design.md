---
name: Per-Agent Memory 决策档案
description: Claude Code agent 隔离机制 5 维分析 + 项目当前 5 agent 引用关系图 + 决定不引入约定式 per-agent memory 的理由
type: reference
---

# Per-Agent Memory 决策档案（2026-04-29）

## 起因

用户听到专家说法："每个子 Agent 有独立上下文+专属记忆，不会互相污染"。询问是否要在本项目引入 per-agent memory 机制。

调研后结论：**专家说法部分正确，"专属记忆"是误解**。当前项目无需改动。

## Claude Code agent 隔离机制（5 维真相）

| 维度 | 是否隔离 | 实现方式 |
|--|--|--|
| **上下文（conversation）** | ✅ 完全隔离 | 每次 Agent 工具调用启动一个全新 conversation，不继承父对话历史 |
| **System prompt** | ✅ 完全隔离 | 每个 agent 一个 `.claude/agents/<name>.md`，独立定义职责边界 |
| **Tool access** | ✅ 完全隔离 | frontmatter `tools:` 字段声明允许的工具子集 |
| **Model 选择** | ✅ 完全隔离 | frontmatter `model:` 字段（如 sonnet/opus） |
| **文件系统 / memory** | ❌ **不隔离** | 所有 agent 共享 `.claude/memory/` + `.claude/auto-memory/` |

## 常见误解：上下文隔离 ≠ 记忆隔离

专家把这两件事混为一谈。区别如下：

- **上下文隔离**：父对话和子 agent 的对话历史互不可见（Claude Code 内置机制）
- **记忆隔离**：子 agent 写入的 memory 文件对其他 agent / 父对话也可见（**Claude Code 没有这个机制**）

如果一个 agent 把洞察写入 `.claude/auto-memory/MEMORY.md`，那条索引会出现在**所有**其他 agent 的 system prompt 里（包括主 Claude）。这通常被视为特性而非 bug——"经验跨 agent 共享"是合理默认。

## 项目当前 5 个 agent 的 memory 引用关系图

```
                       共享 memory 池
              .claude/memory/ + .claude/auto-memory/
                            |
        ┌─────────┬─────────┼─────────┬───────────┐
        ↓         ↓         ↓         ↓           ↓
    curator  developer  step1     step2      step3-
                       analyzer  architect   builder
        |         |         |         |           |
        |  feedback_       |         |    reference_
        |  conclusion_    无显式引用      pipeline_
        |  coloring.md                  repair.md
        |
       无引用
```

| Agent | 显式引用的 memory 文件 | 备注 |
|--|--|--|
| curator | （无） | 写报告到 `pipeline-progress/05-...md`，不读写 memory |
| developer | `.claude/memory/feedback_conclusion_coloring.md` | PPT 移植时确实需要染色规范 |
| step1-analyzer | （无） | 模板分析自包含 |
| step2-architect | （无） | GPT prompt 生成自包含 |
| step3-builder | `.claude/memory/reference_pipeline_repair.md` | 构建失败时诊断需要 |

**所有引用都指向共享 memory 池**，没有任何 per-agent 子目录。

## 当前架构为什么够用

虽然 memory 全局共享，但实际上不会污染：

1. **memory 文件是按需 Read 的**——agent 只在 system prompt 显式引用了某文件时才会读，而不是全部加载
2. **`.claude/auto-memory/MEMORY.md` 经过精简**——只剩 user_profile / stability / check_skills_first 3 条 P0，不会让其他 agent 拿到不相关的洞察
3. **职责通过 prompt 隔离**——每个 agent 的 system prompt 明确说"你做什么、不做什么"
4. **5 个 agent 当前的引用关系健康**——developer 引用染色规范、step3-builder 引用修复指引，都是"该看的看了，不该看的没引用"

## 何时该考虑约定式 per-agent memory

满足以下**任一条件**才值得引入复杂度：

- [ ] 出现具体的"误用 memory"事故（如 step1-analyzer 看了 fix4 chart 决策导致它做出错误的分析）
- [ ] 某 agent 沉淀的知识有强保密 / 强领域性，不应跨 agent 流出（如 fictitious "财务 agent" 不能让 PPT agent 看见）
- [ ] auto-memory 索引条数膨胀回 10+ 条，且大多与某单一 agent 强相关（这时分流到 per-agent 才有意义）
- [ ] 子 agent 自动捕获的洞察明显应该归属某 agent 而非全局（不容易判断时一般归全局更简单）

**当前都不满足**。引入约定式 per-agent memory 反而带来：
- 每次写 memory 时多一个"放哪个 agent 子目录"的决策
- agent prompt 都要加"只读自己子目录"的约束（容易遗忘）
- Curator 流程要相应改造
- 移植到新项目时多一层抽象

## 如果未来真要引入：约定式 per-agent memory 设计草图

仅作未来参考。**不在本次实施。**

```
.claude/
├── memory/                       # 跨 agent 共享池
│   ├── MEMORY.md
│   ├── feedback_*.md             # 通用经验
│   └── reference_*.md            # 架构档案
├── auto-memory/                  # auto-loaded 用户级偏好
│   ├── MEMORY.md
│   └── ...
└── agents/
    ├── curator.md
    ├── curator/                  # ← curator 专属 memory
    │   ├── MEMORY.md
    │   └── feedback_*.md
    ├── developer.md
    ├── developer/                # ← developer 专属
    │   └── ...
    └── ...
```

约束：
- 每个 agent 的 prompt 末尾加"你只读 `.claude/agents/<your_name>/memory/` 和 `.claude/memory/` 共享池"
- 跨 agent 知识由 Curator 显式 promote 到共享池
- auto-memory 仍然全局（用户偏好不分 agent）

## 验证步骤（确认本结论是否还成立）

每隔半年或新增 agent 时跑一次：

1. 数 `.claude/auto-memory/MEMORY.md` 的索引行数——超过 5 条要重新评估
2. grep 各 agent 引用的 memory 文件路径，看是否出现不该看的引用
3. 回顾近期使用日志，是否有"某 agent 因看了不相关 memory 出错"的事故
4. 全部正常 → 维持当前架构

## 相关文档

- 上层规则（写新 memory 时的归属判断）：`.claude/commands/mc-update.md` 中"auto-memory vs memory 知识边界规则"段
- 同类决策档案格式参考：`.claude/memory/reference_3account_junction.md`
- agent 定义文件：`.claude/agents/{curator,developer,step1-analyzer,step2-architect,step3-builder}.md`
