# CLAUDE.md - PPT Pipeline + Agent 项目规范

> 通用规范 + 入口索引。详情参见各 agent 与 memory 文件。

---

## 0. 防卡顿规范

- 同一方案连续失败 2 次 → 停下来向用户说明原因，提出替代方案
- 预计超过 2 分钟的操作 → 用 Agent(run_in_background) 分流
- 遇到不确定的技术选型 → 先问用户，不要默默试超过 3 分钟

---

## 1. 项目结构

```
项目根目录/
├── orchestrator.py                  # 5-Agent 调度（Pipeline先行 + LLM精调）
├── pipeline/
│   ├── ppt_pipeline_common.py       # 公共工具（路径、COM、Excel、批注解析）
│   ├── 01_shape_detail.py           # Step 1: 提取模板shape
│   ├── 01b_auto_annotate.py         # Step 1B: 规则表自动批注
│   ├── 02_shape_analysis.py         # Step 2: 角色推断 + prompt生成
│   ├── 02b_iteration_setup.py       # Step 2B: 修正轮sheet创建 + 基础修正
│   ├── 03a_build_shape.py           # Step 3A: 内容生成（Python/GPT）
│   ├── 03b_build_ppt_com.py         # Step 3B: COM写入PPT
│   ├── 04_shape_diff_test.py        # Step 4: 三层验收 + fix_type分类
│   └── prompt_templates/gpt_summary.md  # GPT prompt 模板
├── pipeline-progress/               # 中间产物（01-/02-/03a-/03b-/04- 前缀）
├── .claude/agents/                  # 5个Agent配置
│   ├── 01-analyst.md                # 分析师：Pipeline推断 + LLM审核模糊项
│   ├── 02-builder.md                # 构建师：Pipeline生成 + LLM精调批注(修正轮)
│   ├── 03-reviewer.md               # 验收师：Pipeline测试 + LLM语义审核
│   ├── 04-developer.md              # 代码专家：LLM修复pipeline代码
│   └── 05-curator.md               # 知识固化师：经验积累 + 规则沉淀
└── src/Function_030.py              # GPT_5 函数（不修改，直接import）
```

---

## 2. 关键规则

- **路径**: 始终用相对路径 + 正斜杠 `/`
- **最小改动**: 只改必要的部分，先说明再动手
- **输出**: 改代码时只说结论（改了什么、为什么、结果），不展示 diff
- **Excel**: 统一用 `win32com.client` COM（加密环境，禁 openpyxl/pandas）
- **PPT**: Clone 模板页，不新建 shape；禁 `python-pptx`

---

## 3. 启动方式

```bash
python orchestrator.py          # 交互选择模式(0-5)
```

| 菜单选项 | 说明 |
|---------|------|
| 0 初始化 | 全新 PPT 分析，从零构建结构和 prompt |
| 1-4 热迭代 | prompt 已存在，编辑 prompt → 调 GPT → 出 PPT（1-4 对应轮次） |
| 5 验收 | 只跑验收，检查最新 PPT 质量 |

> Excel 不存在时，无论选什么都强制路由到选项 0

Curator Agent 独立于 orchestrator，通过 `/role-curator` 手动调用。

---

## 4. 关键配置

- 模板目录: `template/`（支持多套 pptx + xlsx，orchestrator 启动时选择）
- 默认模板: `template/standard and empty template.pptx`
- 默认数据: `template/source data.xlsx`
- GPT: `openai/gpt-5.4`（OpenRouter），`from src.Function_030 import GPT_5`

---

## 5. 详情索引

| 主题 | 位置 |
|------|------|
| Agent 角色定义（Analyst/Builder/Reviewer/Developer） | `.claude/agents/01~04-*.md` |
| 知识固化师（Curator） | `.claude/agents/05-curator.md` |
| 三层门禁 + fix_type 分类 | `.claude/agents/03-reviewer.md` |
| 用户批注字段 + golden reference | `.claude/agents/01-analyst.md` |
| 冷启动/热迭代流程图 + 版本追溯 | `.claude/memory/project_workflow_modes.md` |
| COM 开发规范 | `.claude/memory/feedback_com_constraints.md` |
| 混合工作流 Pipeline→LLM | `.claude/memory/feedback_hybrid_workflow.md` |
| GPT 数据稀疏时截断问题 | `.claude/memory/feedback_gpt_sparse_data.md` |
| 手动 Pipeline 命令 + 批注字段 | `.claude/memory/reference_manual_pipeline.md` |
| 架构决策记录 | `.claude/memory/project_4agent_architecture.md` |
| `src/` 目录 | 历史遗留 main.py 相关模块，与 Pipeline/Agent 工作流无关 |

---

## 6. 变更记录

| 日期 | 变更 |
|------|------|
| 2026-04-08 | CLAUDE.md 瘦身：179行→索引式，详情迁移到 agents/memory |
| 2026-04-08 | 新增 05-curator.md（知识固化师），独立于 orchestrator |
| 2026-04-08 | 修复：STRATEGY_CODES 补 extract_image、gpt_rich→gpt_prompted 统一、contract_section 占位符 |
| 2026-04-01 | 4-Agent 架构定型，弃用 6-Agent v6 |
