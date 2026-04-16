# Workflow 第二轮优化方案：借鉴 mc-ppt 经验

## Context

用户刚完成另一个项目（mc-ppt，md→html→ppt 路线），希望从中借鉴两点经验来优化当前 3rd-ppt-prj 工作流：

1. **slash commands + agents.md 的轻量调用模式**
2. **CLAUDE.md 瘦身与文档分层**

同时确认 `.claude/agents/archive/` 历史遗留文件夹是否可以删除。

**前置约束**（来自上一轮决策）：保留 Python orchestrator 作为工作流引擎，不用 LLM 委派替代它。

---

## 第一部分：mc-ppt 参考项目的核心做法

读取了 `[feature01-orch-update]/` 中的两个文件，得出以下事实（注意：这是一个**完全不同**的 PPT 构建路线）：

### 1. mc-ppt 的工作流是 md→html→ppt（与当前项目不同）
- 用 Playwright 从 HTML 提取坐标 → 截图（文字/图片遮罩）→ win32com 叠加可编辑元素
- 三脚本流水线：`extract_layout.py` → `screenshot_masked.py` → `export_hybrid_win32com.py`
- **不适用于当前项目**——本项目走的是 PPT 模板克隆 + COM 写入路线，无 HTML 中间层

### 2. mc-ppt 的 CLAUDE.md 结构（约 65 行，4 个章节）

```
0. 防卡顿规范（失败 2 次停下、>2min 用 Agent 后台、不确定先问）
1. 文件组织约定（StepN/ 子目录、agents/ 角色、commands/ 命令）
2. 三条核心禁止规则（HTML 专用，不适合本项目）
3. 自定义命令表（/today, /role-pm, /role-researcher, /role-builder, /role-converter）
4. 变更记录
```

**关键观察**：mc-ppt 把所有角色专属内容**迁移到 `agents/` 目录**，CLAUDE.md 只留通用规范 + 命令索引 + 变更记录。

### 3. mc-ppt 的 slash commands 模式

`.claude/commands/role-pm.md` 之类的文件 = 一句话指令，让 Claude **以该角色身份执行**当前对话任务。

mc-ppt 没有 Python orchestrator——它的工作流是 user 输入 `/role-pm` → Claude 加载该 agent 的 system prompt → 直接在主对话执行。

---

## 第二部分：哪些可以借鉴，哪些不能

| mc-ppt 做法 | 是否借鉴 | 理由 |
|-------------|----------|------|
| HTML 中间层 + Playwright 截图 | ❌ | 完全不同路线，本项目用 PPT 模板克隆 |
| Slash commands 替代 orchestrator | ❌ | 本项目 orchestrator 1622 行确定性逻辑无法替代 |
| Slash commands **包装** orchestrator 入口 | ✅ | 用户体验提升：少打几个字 + 无需选菜单 |
| CLAUDE.md 瘦身（只留通用规范 + 索引） | ✅ | 当前 CLAUDE.md 179 行，可降到 ~80 行 |
| 详细规则迁移到 `.claude/agents/` 或 `.claude/memory/` | ✅ | 已部分实施，可继续推进 |
| 变更记录章节 | ✅ | 当前 CLAUDE.md 没有，建议加上 |
| 防卡顿规范（失败 2 次停下） | ✅ | 通用最佳实践，建议加入 |
| StepN/ 子目录结构 | ❌ | 本项目用 `pipeline-progress/` 平铺更适合 8 步流水线 |

---

## 第三部分：当前 CLAUDE.md 现状分析

`.claude/CLAUDE.md`（179 行）目前的内容：

| 区块 | 行数估计 | 是否应留在 CLAUDE.md |
|------|---------|-------------------|
| 项目结构图 | ~25 | ✅ 留下（导航必备） |
| 关键规则（路径、最小改动、Excel/PPT） | ~10 | ✅ 留下（通用） |
| 启动命令 + 两种模式表 | ~15 | ✅ 留下（核心入口） |
| 冷启动流程图 | ~10 | ⚠️ 简化或移到 `agents/` |
| 热迭代流程图 | ~12 | ⚠️ 简化或移到 `agents/` |
| 混合模式 4-列对照表 | ~10 | ❌ 移到 `memory/project_*.md` |
| 版本追溯表 | ~6 | ❌ 移到 `memory/` |
| 三层门禁表 | ~6 | ❌ 移到 `03-reviewer.md`（已部分存在） |
| fix_type 5 类分流表 | ~10 | ❌ 移到 `02-builder.md` 或 `memory/` |
| 手动 Pipeline 命令 | ~10 | ❌ 移到 `commands/manual-pipeline.md` |
| 用户批注字段表 + golden reference | ~15 | ❌ 移到 `01-analyst.md`（已部分存在） |
| 关键配置 | ~5 | ✅ 留下 |
| COM 开发规范表 | ~10 | ❌ 移到 `memory/feedback_com_constraints.md` |
| src/ 目录附录 | ~5 | ✅ 留下 |

**目标：从 179 行 → 约 80-90 行**，CLAUDE.md 变成"导航 + 通用规范 + 命令索引"。

---

## 第四部分：分阶段优化方案

### Phase A：archive 文件夹清理（风险：零 | 工作量：1 分钟）

**已验证**：`.claude/agents/archive/` 包含 7 个文件（v6 6-agent 系统遗留）：
- `01-arch.md`, `02-tech.md`, `03-dev.md`, `04-test.md`, `05-opti.md`, `06-secu.md`, `CLAUDE-6-Agents.md`

**引用情况**：
- ✅ 当前 `orchestrator.py` 不引用任何 archive 文件
- ✅ 当前 4 个 active agent 不引用
- ✅ `.claude/CLAUDE.md`、`new-ppt-workflow.md`、`USAGE_GUIDE.md` 不引用
- ⚠️ 仅 `orchestrator_v6_legacy.py` + `debug/mc-dir-v6.py` + `debug/Mc-debug-*.md` 引用

**结论：可以安全删除 `.claude/agents/archive/`**。

附带建议（可选）：
- `orchestrator_v6_legacy.py` 也可移到 `debug/` 或删除（已经是 legacy 后缀）
- `debug/` 目录已经是隔离的历史区，无需动

---

### Phase B：Slash Commands 包装 orchestrator 入口（风险：低 | 工作量：30 分钟 | 价值：高）

**目的**：让用户从 `python orchestrator.py` + 选菜单 → 一句 `/init` 或 `/iter2` 直接启动。

**新建 6 个 slash commands** 在 `.claude/commands/`：

```
.claude/commands/
├── init.md          # 冷启动（菜单选项 0）
├── iter1.md         # 1 轮热迭代（菜单选项 1）
├── iter2.md         # 2 轮热迭代（菜单选项 2）
├── iter3.md         # 3 轮热迭代（菜单选项 3）
├── auto2.md         # 自动 2 轮（菜单选项 4）
└── review.md        # 仅验收（菜单选项 5）
```

**每个文件内容示例**（`init.md`）：

```markdown
请运行 orchestrator 的冷启动模式（菜单选项 0），完成全新 PPT 的初始化分析与构建。

执行命令：
\`\`\`bash
python orchestrator.py
\`\`\`

启动后在菜单中选择 "0"。等待 Analyst 完成批注后，会暂停让我校准 Excel 黄色单元格。
```

或者更精简版（`iter2.md`）：

```markdown
运行 orchestrator 进入 2 轮热迭代模式（菜单选项 2）。
\`\`\`bash
python orchestrator.py
\`\`\`
菜单中输入 2。
```

**进阶版**（如果 orchestrator 支持命令行参数直接传模式，未来可改成）：

```markdown
\`\`\`bash
python orchestrator.py --mode 2
\`\`\`
```

**当前限制**：orchestrator.py 现在是交互式菜单（`input()`）。如果想真正实现 `/iter2` 一键启动，需要给 orchestrator 加 `--mode N` 参数（约 10 行代码改动）。

**两种实施路径**：
1. **保守路径**：slash command 只是文档化命令 + 提示用户在菜单选哪个数字。零代码修改。
2. **激进路径**：给 orchestrator 加 `--mode N` 参数，slash command 直接传参。10 行 Python 改动 + 6 个 md 文件。

**推荐保守路径**（先验证用户体验，再决定是否加参数）。

---

### Phase C：CLAUDE.md 瘦身（风险：低 | 工作量：1-2 小时 | 价值：高）

**目标结构**（约 80-90 行）：

```markdown
# CLAUDE.md - PPT Pipeline + Agent 项目规范

> 通用规范 + 入口索引。详情参见各 agent 与 memory 文件。

## 0. 防卡顿规范
- 同一方案失败 2 次 → 停下来说明并提替代
- 预计 >2min 操作 → 用 Agent(run_in_background)
- 不确定的技术 → 先问，不要默默试 >3min

## 1. 项目结构（保留当前的目录树，~25 行）

## 2. 关键规则
- 路径：相对路径 + 正斜杠
- 最小改动：先说明再动手
- 输出：只说结论，不展示 diff
- Excel：win32com COM（禁 openpyxl/pandas）
- PPT：Clone 模板页（禁 python-pptx）

## 3. 启动方式

| Slash Command | 等价菜单 | 说明 |
|---------------|---------|------|
| `/init` | 0 | 冷启动初始化 |
| `/iter1` ~ `/iter3` | 1-3 | 热迭代（1-3 轮） |
| `/auto2` | 4 | 自动 2 轮 |
| `/review` | 5 | 仅验收 |

或手动：`python orchestrator.py`

## 4. 工作流模式（简表）

| 模式 | 何时用 |
|------|------|
| 冷启动 | Excel 不存在或全新 PPT |
| 热迭代 | 已有 prompt，调 GPT 出 PPT |
| 验收 | 检查最新 PPT |

## 5. 关键配置
- 模板：`pipeline/standard and empty template.pptx`
- 数据：`pipeline/source data.xlsx`
- GPT：`openai/gpt-5.4`（OpenRouter）

## 6. 详情索引

| 主题 | 位置 |
|------|------|
| Agent 角色定义 | `.claude/agents/01-analyst.md` ~ `04-developer.md` |
| 三层门禁 + fix_type | `.claude/agents/03-reviewer.md` |
| 用户批注字段 + golden reference | `.claude/agents/01-analyst.md` |
| COM 开发规范 | `.claude/memory/feedback_com_constraints.md` |
| 混合工作流 Pipeline→LLM | `.claude/memory/feedback_hybrid_workflow.md` |
| 4-Agent 架构决策 | `.claude/memory/project_4agent_architecture.md` |
| 完整手动 Pipeline 命令 | `pipeline/README.md`（新建）或 `commands/manual.md` |

## 7. 变更记录
| 日期 | 变更 |
|------|------|
| 2026-04-08 | CLAUDE.md 瘦身：详情迁移到 agents/memory，新增 slash commands |
| 2026-04-01 | 4-Agent 架构定型，弃用 6-Agent v6 |
```

**迁移目标位置**：

| 当前 CLAUDE.md 内容 | 迁移到 |
|-------------------|--------|
| 三层门禁表 | `.claude/agents/03-reviewer.md`（验证已存在或追加） |
| fix_type 5 类表 | `.claude/agents/02-builder.md`（验证已存在或追加） |
| 用户批注 golden reference | `.claude/agents/01-analyst.md`（验证已存在或追加） |
| COM 开发规范表 | 新建 `.claude/memory/feedback_com_constraints.md` |
| 手动 Pipeline 命令 | 新建 `pipeline/README.md` 或 `.claude/commands/manual-pipeline.md` |
| 版本追溯表 | 新建 `.claude/memory/project_versioning.md` |
| 冷启动/热迭代详细流程图 | 新建 `.claude/memory/project_workflow_modes.md` |

**迁移原则**：
- 信息**不丢失**，只搬家
- 搬家时检查目标文件是否已经有相同内容（避免重复）
- CLAUDE.md 留一个"详情索引"表格，把搬走的内容指向新位置

---

### Phase D：CLAUDE.md 位置标准化（可选 | 风险：低）

**当前**：CLAUDE.md 在 `.claude/CLAUDE.md`
**建议**：移到项目根 `./CLAUDE.md`（官方推荐位置，工具发现更标准）

注意：两个位置都能被 Claude Code 自动加载。如果不想动，保持现状也可以。

---

## 第五部分：执行优先级

| Phase | 工作量 | 风险 | 价值 | 建议执行 |
|-------|-------|------|------|---------|
| **A: 删除 archive** | 1 分钟 | 零 | 低（清理） | ⭐ 立即 |
| **B: Slash commands（保守版）** | 30 分钟 | 低 | 高 | ⭐ 立即 |
| **C: CLAUDE.md 瘦身** | 1-2 小时 | 低 | 高 | ⭐ 推荐 |
| B-激进版（加 --mode 参数） | +10 分钟 | 低 | 中 | 可选 |
| D: CLAUDE.md 位置移动 | 5 分钟 | 低 | 低 | 可选 |

**推荐顺序**：A → B（保守）→ C → 用一段时间 → 决定是否做 B-激进 + D。

---

## 第六部分：与上一轮 ochestrator-update.md 的关系

上一轮制定的 5 阶段方案（Phase 0-4：配置外部化、`--agent` 标志、模块拆分、Agent 注册、Prompt 模板）是**重型重构**。

本轮 A/B/C 是**轻型优化**，与上一轮**互补不冲突**：
- 本轮 A/B/C 改善**用户体验 + 文档结构**，不动 orchestrator 内部
- 上一轮 0-4 改善**代码质量 + 可维护性**，需要拆解 orchestrator

**建议执行顺序**：
1. 先做本轮 A + B + C（快速见效，~3 小时）
2. 视情况启动上一轮 Phase 0（配置外部化，~2 小时）
3. 上一轮 Phase 1（模块拆分）作为长期目标，等本轮稳定后再启动

---

## 第七部分：必须保留的设计（不变）

1. Python orchestrator 作为工作流引擎（不被 slash commands 替代）
2. 4-Agent 架构（analyst/builder/reviewer/developer）
3. Pipeline-first 混合架构
4. 5 个交互暂停点
5. 版本追踪、Excel COM 集成、Windows 特殊处理

---

## 关键文件清单

### 需要修改
- `D:/Technique Support/Claude Code Learning/3rd-ppt-prj/.claude/CLAUDE.md` — 瘦身（179 行 → ~85 行）

### 需要新建
- `.claude/commands/init.md`, `iter1.md`, `iter2.md`, `iter3.md`, `auto2.md`, `review.md` — 6 个 slash command
- `.claude/memory/feedback_com_constraints.md` — COM 开发规范（从 CLAUDE.md 迁出）
- `.claude/memory/project_workflow_modes.md` — 冷启动/热迭代详细流程（从 CLAUDE.md 迁出）
- `pipeline/README.md` 或 `.claude/commands/manual-pipeline.md` — 手动 Pipeline 命令（从 CLAUDE.md 迁出）

### 需要删除
- `.claude/agents/archive/` 整个文件夹（已验证零引用）

### 可能需要追加内容
- `.claude/agents/01-analyst.md` — 用户批注 golden reference（验证后追加）
- `.claude/agents/02-builder.md` — fix_type 5 类表（验证后追加）
- `.claude/agents/03-reviewer.md` — 三层门禁阈值（验证后追加）

---

## 验证方案

完成 A+B+C 后：
1. 用 `/init` 启动一次冷启动验证 slash command 工作
2. 用 `/iter2` 跑一次热迭代验证菜单提示
3. 让 Claude 读取一个 agent 任务，确认它能从 CLAUDE.md 索引找到 `.claude/agents/` 中的详细规则
4. 跑 `python -m py_compile orchestrator.py` 确保 orchestrator 未受影响
