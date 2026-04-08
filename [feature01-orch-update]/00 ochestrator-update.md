# Orchestrator 升级改造方案

## Context

用户问了 Grok：固定 Python 代码的 orchestrator.py 是否已过时？Grok 建议删除 orchestrator.py，让 Claude Code 原生子 Agent 系统自动调度。用户希望保留固定菜单选项（工作模式、最大轮次），同时获得系统性的改造建议。

---

## 第一部分：对 Grok 建议的系统性评估

### 建议 1："迁移到原生 .claude/agents/ 系统"
**评价：已经做了。** 项目已有 4 个 agent 文件在 `.claude/agents/` 中，Grok 没有看到这一点。

### 建议 2："删除 orchestrator.py，让 Claude 自动委派"
**评价：对本系统而言是危险的错误建议。**

orchestrator.py 不是简单的"路由器"，而是 1622 行的**工作流引擎**，包含：
- **确定性 Pipeline 执行**：按精确顺序调用 8 个 Python 脚本（subprocess.run）
- **版本自动检测**：扫描 `claude-ppt X.Y.pptx` + `.version_tracker.json`，做版本算术
- **两阶段 03a 执行**：`--assemble-only` → 暂停人工审核 → `--execute-prompts`
- **复杂路由逻辑**：冷启动 vs 热迭代、续接检测、fix_type 分流（code→Developer，其余→Builder）
- **5 个交互暂停点**：打开 Excel（`os.startfile()`）让用户校准
- **700+ 行 Prompt 构建**：包含修正数据、golden reference、内容片段、sheet 名等上下文
- **Windows 特殊处理**：cmd 长度限制（>4000 字符写临时文件）、COM 集成

删除 orchestrator.py = 丧失以上全部能力。Claude Code 的自动委派设计用于"模糊请求的智能路由"，而非"精确 8 步 Pipeline + 条件分支 + 交互暂停 + 版本算术"。

### 建议 3："保留 orchestrator.py 作为薄路由，用 `claude --agent`"
**评价：部分有效。** `--agent` 标志替代当前的 `--append-system-prompt` 机制是合理改进（见 Phase 2）。但说 orchestrator 是"薄路由"是误解——agent 调用只占全部代码的 ~100/1622 行。

### 建议 4："创建 coordinator-agent.md 做任务分解"
**评价：不适合本系统。** 工作流是固定的（Analyst→Builder→Reviewer→Developer→循环），不需要"任务分解"。Coordinator Agent 会引入不确定性。

### 总结
Grok 把本项目当作通用"多 Agent 代码项目"给了通用建议，未理解：
- 系统是**混合 Pipeline**（确定性 Python + LLM Agent），不是纯 LLM 工作流
- Agent 只做 4 个特定语义任务，不是通用工人
- 稳定性 > 新颖性

**用户的直觉是对的**：写死的 Python 提供稳定性。升级方向不是用 LLM 委派替代 Python，而是**改善 Python 代码的模块化、可配置性和可维护性**。

---

## 第二部分：当前真实问题

| 问题 | 影响 |
|------|------|
| 1622 行单文件，混合 UI/路由/Pipeline执行/Prompt构建/状态管理 | 改任何一处都有连带风险 |
| Agent 调用用 `claude -p` + `--append-system-prompt` + 手写 YAML 解析 | 40 行自定义解析器，frontmatter 中 tools/description 解析了但没用 |
| 配置硬编码（阈值、模型、路径、退避时间、显示名） | 调参需改代码 |
| 新增 Agent 需改 `AGENT_CONFIGS` dict + 写 prompt builder 方法 + 改路由 | 扩展成本高 |
| Prompt 构建器是 Python f-string，700+ 行混中英文 | 调 prompt 需改代码，有语法风险 |

---

## 第三部分：分阶段升级方案

### Phase 0：配置外部化（风险：极低 | 价值：高）

**创建 `pipeline_config.json`**，抽取：
- 质量阈值：`visual_threshold: 98`, `readability_threshold: 95`, `semantic_threshold: 100`
- Agent 设置：每个 agent 的 model、max_budget、max_turns
- 退避时间：`[5, 10, 20]`
- 路径：template、source data、progress dir
- 显示名：中文 agent 名称映射

在 `PPTOrchestrator.__init__` 中加载，当前值作为默认值（无配置文件也能跑）。

### Phase 2：Agent 调用升级（风险：低 | 价值：中）

**将 `claude -p <prompt> --append-system-prompt <body>` 替换为 `claude --agent <name> -p <prompt>`**

改动点：
- `AgentExecutor.run_agent()` 中删除 `_parse_agent_file()` 方法
- 命令构建改为：
  ```python
  cmd = ["claude", "--agent", agent_name, "-p", prompt,
         "--output-format", "stream-json", "--verbose",
         "--max-turns", "20", "--max-budget-usd", str(budget),
         "--session-id", sid, "--no-chrome", "--dangerously-skip-permissions"]
  ```
- Claude CLI 自动读取 `.claude/agents/XX-name.md`，应用 model/tools/description
- Windows 临时文件 workaround 保留不变
- 重试逻辑保留不变

### Phase 1：模块拆分（风险：中 | 价值：极高）

**将 orchestrator.py 拆为专注模块**：

```
orchestrator/
  __init__.py           # re-export main()
  cli.py                # 菜单系统、argparse、账户选择 (~120行)
  config.py             # 配置加载、默认值 (~60行)
  agent_executor.py     # AgentExecutor 类 (~200行)
  state_manager.py      # StateManager 类 (~30行)
  error_handler.py      # ErrorHandler 类 (~40行)
  progress_monitor.py   # ProgressMonitor + 显示名 (~50行)
  version_tracker.py    # 版本检测、记录 (~60行)
  pipeline_runner.py    # Pipeline 脚本执行 (~100行)
  prompt_builder.py     # 所有 prompt 构建方法 (~250行)
  workflow.py           # PPTOrchestrator.run() 核心逻辑 (~400行)
  main.py               # 入口
```

提取顺序（每步独立可验证）：
1. config.py → 2. state_manager.py → 3. error_handler.py → 4. progress_monitor.py → 5. version_tracker.py → 6. agent_executor.py → 7. pipeline_runner.py → 8. prompt_builder.py → 9. cli.py → 10. workflow.py

每次提取后跑 `python -m py_compile` 验证。

### Phase 3：动态 Agent 注册（风险：低 | 价值：中）

**创建 `agent_registry.json`** 替代硬编码 `AGENT_CONFIGS`：
```json
{
  "analyst": {
    "file": "01-analyst.md",
    "output_files": ["pipeline-progress/01-shape_detail_com.json"],
    "display_name": "PPT模板分析师",
    "phase": "analysis"
  },
  ...
}
```

不用扩展 agent.md frontmatter（避免 Claude CLI 未来校验冲突）。工作流顺序仍硬编码在 workflow.py（固定流程是设计意图）。

### Phase 4：Prompt 模板外部化（风险：中 | 价值：高）

**将 prompt 构建从 Python f-string 迁移到模板文件**：

```
pipeline/prompt_templates/orchestrator/
  analyst_full.md           # 全量批注增强
  analyst_targeted.md       # 定向修复模式
  builder_prompt_optimizer.md  # prompt 重写
  reviewer_llm_only.md      # 语义审核
  developer.md              # 代码修复
```

用 `{{variable}}` 占位符 + `str.replace()`（无 Jinja2 依赖）。保留 Python 硬编码版本作为 fallback。

### Phase 5（可选/延后）：探索原生 Agent Tool

当前评估：**不建议近期实施**。原因：
- orchestrator 需要运行确定性 Python 脚本（Agent Tool 只能 spawn LLM）
- 需要 `input()` 交互暂停（Agent Tool 上下文不支持）
- 版本追踪等文件系统操作在 Python 中更可靠

待 Phase 0-4 稳定后再评估。

---

## 优先级矩阵

| Phase | 风险 | 价值 | 工作量 | 依赖 |
|-------|------|------|--------|------|
| **0: 配置外部化** | 极低 | 高 | ~2h | 无 |
| **2: --agent 标志** | 低 | 中 | ~2h | 无 |
| **1: 模块拆分** | 中 | 极高 | ~6-8h | Phase 0(可选) |
| **3: Agent 注册** | 低 | 中 | ~3h | Phase 1 |
| **4: Prompt 模板** | 中 | 高 | ~4-5h | Phase 1 |
| 5: Native Agent | 高 | 低 | 未知 | 全部 |

**推荐执行顺序**：Phase 0 → Phase 2 → Phase 1 → Phase 3 → Phase 4

---

## 必须保留的设计

1. **固定菜单系统**（选项 0-5）
2. **工作模式选择**（冷启动 vs 热迭代）
3. **最大轮次选择**
4. **交互暂停点**（批注审核、prompt 审核、PPT 审核）
5. **Pipeline-first 架构**（确定性 Python 脚本优先，LLM 只做语义任务）
6. **版本管理**（自动检测 + 追踪）
7. **Windows COM 集成**
8. **账户选择**（mc/xh）

---

## 验证方案

每个 Phase 完成后：
1. `python -m py_compile orchestrator.py`（或各模块）
2. 运行菜单选项 5（review-only）验证基础流程
3. 运行菜单选项 0（initialize）验证冷启动
4. 运行菜单选项 2 验证热迭代 + 修正循环

## 关键文件

- `orchestrator.py` — 1622 行主文件，所有改造的目标
- `.claude/agents/01-analyst.md` ~ `04-developer.md` — Agent 规格
- `pipeline/ppt_pipeline_common.py` — 共享工具（878行）
- `pipeline_config.json` — 新建的配置文件
- `agent_registry.json` — 新建的 Agent 注册表
