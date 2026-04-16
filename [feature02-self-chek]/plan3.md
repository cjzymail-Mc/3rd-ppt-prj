# Plan 3: 局部循环 + 3+1 Agent 架构 + Pipeline 轻度清理

> 本方案取代 plan2。核心变化：**按步骤切分 agent**（而非按能力）、**Pipeline 瘦身**、**Orchestrator 下放执行权给 agent**。

---

## 1. 背景

### 1.1 当前痛点

1. **整体循环效率低**：一次 PPT 不合格 → 全流程重跑 → 全自动模式也无法托管
2. **Agent 按能力切分导致碎片化**：analyst/builder/reviewer 各管一段，跨步骤协作复杂
3. **自检散落在 orchestrator**：pipeline 和 agent 都依赖外部编排，缺乏自治
4. **Pipeline 含遗留的多-sheet 迭代代码**（`02b`, `04`），与新的"局部循环"模型不兼容
5. **developer agent 冗余**：Claude Code 主对话本身就有 Read/Edit/Write/Bash 能力，无需单独 agent

### 1.2 设计目标

| 目标 | 实现手段 |
|------|---------|
| 局部自检循环 | 每个步骤内部两阶段（Python → Agent），最多 2 次 |
| 按步骤切分 agent | 3 个主线 agent + 1 个辅助 agent (curator) |
| 代码修复回归主对话 | 删除 developer agent，pipeline 代码 bug 由 Claude Code 主对话处理 |
| Pipeline 瘦身 | 删除废弃脚本，新增 `self_check.py` |
| Orchestrator 瘦身 | 只做菜单 + agent 调度，不再直接跑 pipeline |
| 多入口 | orchestrator 菜单 / slash command（不再用 @ mention）|
| 一键全自动 | 菜单 0 串联 1→2→3，agent 内部全自治 |

---

## 2. 新菜单（一字不改）

```
🎯 请选择运行模式:

  0️⃣  <全自动> ── 分析 → 构建 → 交付ppt
  1️⃣  步骤1 —— 分析（新）PPT 模板
  2️⃣  步骤2 —— 构建 prompt
  3️⃣  步骤3 —— 构建 & 交付 ppt
```

每个选项行为：
- **0 全自动**：依次串联步骤1→2→3，中间不暂停，最后自动打开 PPT
- **1 步骤1**：运行 → 自检循环 → 通过后弹 Excel 供审核
- **2 步骤2**：运行 → 自检循环 → 通过后弹 Excel 供审核
- **3 步骤3**：运行 → 自检循环 → 通过后自动打开 PPT 供审核

---

## 3. 新架构总览

### 3.1 角色分工

```
┌─────────────────────────────────────────────────────────────┐
│                  入口层（2 种入口）                           │
│       orchestrator.py 菜单         slash command              │
└───────────┬─────────────────────────┬────────────────────────┘
            │                         │
            ↓                         ↓
┌─────────────────────────────────────────────────────────────┐
│              主线 Agent（3 个，一步一 agent）                 │
│  step1-analyzer    step2-architect    step3-builder          │
│       ↕                  ↕                  ↕                │
│    自检自修复        自检自修复         自检自修复             │
└───────────┬─────────────────┬─────────────────┬────────────┘
            │                 │                 │
            ↓                 ↓                 ↓
┌─────────────────────────────────────────────────────────────┐
│              Pipeline 层（工具脚本，被 agent 调用）            │
│  01,01b / 02,03a / 03b  +  ppt_pipeline_common.py           │
│                         +  self_check.py【新】               │
└─────────────────────────────────────────────────────────────┘

┌────────────────────────────────────────────────────────────┐
│  辅助 Agent: curator                                        │
│  通过 /curator 调用 → 一轮结束后沉淀经验                     │
└────────────────────────────────────────────────────────────┘

┌────────────────────────────────────────────────────────────┐
│  代码修复 → Claude Code 主对话                              │
│  pipeline 代码 bug 直接在主对话中处理（无需独立 agent）      │
└────────────────────────────────────────────────────────────┘
```

### 3.2 两阶段自检循环（所有步骤统一）

```
step-agent 被调用 (来自 orchestrator / slash command):

┌─ Attempt 1: Python Pipeline ──────────────────┐
│  • Bash 调用对应 pipeline 脚本                 │
│  • 从零建立结构性产出（稳定、确定）              │
│  • 调 self_check 函数做自检                    │
│  └─ PASS → 返回成功                            │
│  └─ FAIL → 进入 Attempt 2                      │
└────────────────────────────────────────────────┘

┌─ Attempt 2: LLM Agent 修复 ───────────────────┐
│  • 读自检失败项 + golden reference             │
│  • LLM 语义分析 + 生成修复                     │
│  • 通过 COM 写回 xlsx / 直接改 JSON            │
│  • 再次调 self_check                           │
│  └─ PASS → 返回成功                            │
│  └─ FAIL → 返回失败 + 问题清单 (用户介入)       │
└────────────────────────────────────────────────┘
```

最大循环次数：**2**（Attempt 1 + Attempt 2）。超限则弹出 Excel/PPT 让用户手动介入。

---

## 4. Agent 设计

### 4.1 `step1-analyzer`（步骤1专属）

**职责**：分析 PPT 模板 → 提取 shape 结构 → 生成批注 → 自检 → 修复

**工具**：`Read, Bash, Edit, Write`

**输入**：
- 用户选的标准模板 `{template_path}`
- 用户选的数据 `{xlsx_path}`

**输出**：
- `pipeline-progress/01-shape_detail_com.json`
- `pipeline-progress/01-shape_detail.xlsx`（含完整批注）

**执行流程**：

```
## Attempt 1 (Python Pipeline)
1. Bash: python pipeline/01_shape_detail.py
2. Bash: python pipeline/01b_auto_annotate.py
3. Bash: python -c "from pipeline.self_check import check_step1; check_step1()"
4. 解析自检结果:
   - PASS → 报告完成 + 退出
   - FAIL → 进入 Attempt 2

## Attempt 2 (LLM 修复)
1. 读 01-shape_detail_com.json 每个 shape 的属性
2. 读 xlsx 当前批注（通过 parse_user_annotations）
3. 对每个 FAIL 项:
   - strategy 为空/(必填) → 根据 shape text 推断正确 strategy
   - description 为空/(必填) → 根据 shape text 生成描述
4. 通过 COM 写回 xlsx
5. 再次调 check_step1
6. PASS → 报告成功; FAIL → 报告问题清单
```

**自检标准**（`check_step1`）：
- `01-shape_detail_com.json` 存在且 shapes 数组非空
- 每个 shape 的 `strategy` 字段已赋值（非空、非 "(必填)"）
- `gpt_prompted` 类 shape 的 `description` 已赋值
- shape 数量与用户选的标准模板页数匹配

---

### 4.2 `step2-architect`（步骤2专属）

**职责**：生成 GPT prompt → 调 GPT 生成内容 → 对比 golden reference → 自检 → 修复

**工具**：`Read, Bash, Edit, Write`

**输入**：
- `pipeline-progress/01-shape_detail_com.json`
- `pipeline-progress/01-shape_detail.xlsx`

**输出**：
- `pipeline-progress/02-prompt_specs.json`
- `pipeline-progress/03a-build_shape_content.json`
- xlsx 的 `GPT-prompt Text` 列填充完毕

**执行流程**：

```
## Attempt 1 (Python Pipeline)
1. 前置检查: 01-shape_detail_com.json + xlsx 必须存在
2. Bash: python pipeline/02_shape_analysis.py
3. Bash: python pipeline/03a_build_shape.py --assemble-only
4. Bash: python pipeline/03a_build_shape.py --execute-prompts
5. Bash: python -c "from pipeline.self_check import check_step2; check_step2()"
6. 解析自检结果:
   - PASS → 退出
   - FAIL → 进入 Attempt 2

## Attempt 2 (LLM 修复)
1. 读 self_check 失败原因（哪些 shape 的 content 不达标）
2. 读 golden reference（用户所选标准模板的原始文本）
3. 对每个 FAIL 项:
   - 结构差异大 → 全面重写该 shape 的 GPT-prompt 文本
   - 关键词缺失 → 在 prompt 中强化关键词约束
   - 长度不达标 → 在 prompt 中加入字数硬约束
4. 通过 write_gpt_prompts_to_xlsx() 写回 xlsx
5. Bash: python pipeline/03a_build_shape.py --execute-prompts (用新 prompt 重新调 GPT)
6. 再次调 check_step2
7. PASS → 报告成功; FAIL → 报告问题清单
```

**自检标准**（`check_step2`）：
- `03a-build_shape_content.json` 存在
- 每个 `strategy≠skip` 的 shape 有非空 content
- content 长度在 `readability_budget` 的 50%~120% 范围内
- `gpt_prompted` shape 的 `required_keywords` 出现在 content 中
- **结构相似度检查**：对比 golden reference（从 `01-shape_detail_com.json[*].text` 提取），段落数/列表项数差异 ≤ 30%

---

### 4.3 `step3-builder`（步骤3专属）

**职责**：通过 COM 写入 PPT → 视觉/属性自检 → 失败时要么自修复要么报告失败类型

**工具**：`Read, Bash, Edit, Write`

**输入**：
- `pipeline-progress/03a-build_shape_content.json`
- `pipeline-progress/01-shape_detail.xlsx`
- 用户选的标准模板

**输出**：
- `output/claude-ppt N.N.pptx`

**执行流程**：

```
## 前置智能检测 (F1)
if xlsx.mtime > 03a-build_shape_content.json.mtime:
    print("[智能检测] xlsx 中 prompt 已更新，重新调 GPT")
    Bash: python pipeline/03a_build_shape.py --execute-prompts

## Attempt 1 (Python Pipeline)
1. 版本号计算
2. Bash: python pipeline/03b_build_ppt_com.py --version X.X
   (03b 内置 4 步自检 + MAX_SELF_FIX=2 自动修复 → 已有局部循环)
3. 读 03b-self_check_report.md 判断是否通过
4. PASS → 退出
5. FAIL → 进入 Attempt 2

## Attempt 2 (分类处理)
1. 分析 03b-self_check_report.md 中的失败类型:
   - 视觉/属性异常 → 代码层问题，返回 "建议在 Claude Code 主对话中修复 pipeline 代码"（可参考 .claude/memory/reference_pipeline_repair.md）
   - 文本长度不达标 → prompt 层问题，返回 "suggest 回到步骤2"
   - shape 匹配错误 → 批注层问题，返回 "suggest 回到步骤1"
2. 不在本 agent 内跨层修复（避免超出职责）
3. 报告失败 + 诊断建议给用户
```

**自检标准**：直接复用 `03b_build_ppt_com.py` 的内置自检报告。

**重要设计**：step3-builder 的 Attempt 2 **不做跨层修复**——因为步骤3 的失败通常意味着上游（步骤1/2）有问题。如果强行在步骤3 修，会污染整个流程的数据一致性。正确做法是给出诊断建议，让用户回到对应步骤。

---

### 4.4 `curator`（辅助，唯一保留的辅助 agent）

**职责**：知识固化 + 经验沉淀（不绑定任何步骤）

**调用方式**：仅 `/curator` slash command（不再支持 @ mention）

**调用时机**：
- 一轮完整工作结束后用户手动触发
- 想沉淀经验/规则/教训时

**输出**：`pipeline-progress/05-solidification-*.md`

---

### 4.5 代码修复的处理方式（不再有独立 agent）

**场景**：pipeline 代码 bug、`_SCORE_COLS` 列名错误、COM API 异常等。

**处理路径**：

```
用户在 Claude Code 主对话直接说:
  "修复 pipeline/03a_build_shape.py 中的 _SCORE_COLS 列名错误"
       ↓
Claude 主对话:
  - 读 CLAUDE.md 第 5 节，找到索引: "Pipeline 代码修复指引"
  - 读 .claude/memory/reference_pipeline_repair.md 获取详细知识
  - 用 Read/Edit/Bash 工具完成修复
  - 自检 (py_compile + 重跑相关 pipeline)
```

**为什么删除 developer agent**：
- Claude Code 主对话本身具备完整代码能力（Read/Edit/Write/Bash）
- developer 提示词无独家知识，全是通用代码修复指引
- Claude Code sub-agent 不能调用其他 sub-agent，所以 step3-builder 无法主动调 developer
- 主对话通过 CLAUDE.md 索引能按需读取 `reference_pipeline_repair.md`，自然继承 pipeline 知识

**知识存储设计原则**：
- **CLAUDE.md** 保持精简：只放索引，不放细节
- **`.claude/memory/reference_pipeline_repair.md`** 存放完整代码修复知识
- 主对话默认加载 CLAUDE.md → 从索引找到 memory 文件 → 按需读取
- 这与项目现有的 memory 组织模式（`feedback_com_constraints.md` 等）一致

---

## 5. Orchestrator 瘦身设计

### 5.1 新 `main()`

```python
def main():
    _select_account()
    project_root = find_project_root()
    _select_template(project_root)  # 用户选标准模板 + 数据

    print("\n🎯 请选择运行模式:\n")
    print("  0️⃣  <全自动> ── 分析 → 构建 → 交付ppt")
    print("  1️⃣  步骤1 —— 分析（新）PPT 模板")
    print("  2️⃣  步骤2 —— 构建 prompt")
    print("  3️⃣  步骤3 —— 构建 & 交付 ppt\n")

    choice = input("请输入 [0-3]（直接回车=0）: ").strip() or "0"
    step = int(choice)

    orch = PPTOrchestrator(project_root, auto_mode=(step == 0))
    success = asyncio.run(orch.run(step))
    sys.exit(0 if success else 1)
```

### 5.2 新 `PPTOrchestrator.run()`

```python
async def run(self, step: int) -> bool:
    if step == 0:
        # 全自动：串联 1→2→3
        if not await self._call_agent("step1-analyzer"): return False
        if not await self._call_agent("step2-architect"): return False
        if not await self._call_agent("step3-builder"): return False
        self._open_latest_pptx()
        return True
    elif step == 1:
        ok = await self._call_agent("step1-analyzer")
        if ok: self._open_xlsx()
        return ok
    elif step == 2:
        ok = await self._call_agent("step2-architect")
        if ok: self._open_xlsx()
        return ok
    elif step == 3:
        ok = await self._call_agent("step3-builder")
        if ok: self._open_latest_pptx()
        return ok
```

### 5.3 删除的方法

```
_run_pipeline, _run_pipeline_step         → 移到 agent 内部（agent 用 Bash 直接调）
_run_builder_pipeline                     → 删除（逻辑进 step2-architect）
_run_prompt_only_pipeline                 → 删除
_run_03a_with_prompt_review               → 删除
_check_review_passed                      → 删除（04 废弃）
_archive_round                            → 删除（无多轮）
_reviewer_llm_only_prompt                 → 删除
_developer_prompt                         → 删除（agent 自带 prompt）
_builder_prompt_optimizer_prompt          → 删除
_builder_llm_only_prompt                  → 删除
_prompts_exist                            → 删除
_analyst_phase2_prompt                    → 删除（agent 自带 prompt）
_detect_next_version_index                → 移到 agent 或保留
_idx_to_version / _record_version         → 移到 agent 或保留
```

### 5.4 保留/新增的方法

```
_call_agent(agent_name: str) → bool      【新】薄封装 AgentExecutor 调用
_open_xlsx()                             【保留】用户审核用
_open_latest_pptx()                      【保留】用户审核用
_select_template()                       【保留】启动时选模板
_verify_pptx_exists()                    【保留】可能由 step3 agent 自己做
AgentExecutor                            【保留】底层 subprocess 调用
StateManager, ErrorHandler, ProgressMonitor  【保留】基础设施
```

### 5.5 `__init__` 简化

```python
def __init__(self, project_root: Path, auto_mode: bool = False):
    self.project_root = project_root
    self.auto_mode = auto_mode
    self.executor = AgentExecutor(project_root)
    self.error_handler = ErrorHandler()
    self.monitor = ProgressMonitor()
    self.state_manager = StateManager(project_root)
    # 去掉: max_rounds, skip_analyst_first_round, init_mode, results{}
```

---

## 6. Pipeline 变动

### 6.1 删除（2 个文件）

| 文件 | 原因 |
|------|------|
| `pipeline/02b_iteration_setup.py` | 多-sheet 迭代模式已废弃，局部循环无需跨 sheet |
| `pipeline/04_shape_diff_test.py` | 用户明确表示由人工审核替代 |

### 6.2 新增（1 个文件）

**`pipeline/self_check.py`** — 自检函数库，供 agent 通过 `python -c` 或 `Bash` 调用。

```python
"""Self-check functions for step1/step2/step3 local iteration loops."""

from pathlib import Path
from typing import Dict, List, Any
import json

def check_step1(progress_dir: Path = None) -> Dict[str, Any]:
    """Step 1 自检: shape 提取完整性 + 批注覆盖.
    
    Returns:
        {
            "passed": bool,
            "issues": [{"shape": str, "problem": str}, ...],
            "summary": str
        }
    """
    # 实现细节:
    # 1. 读 01-shape_detail_com.json
    # 2. 调 parse_user_annotations() 读 xlsx 批注
    # 3. 检查每个 shape:
    #    - strategy_exact 非空且非 "(必填)"
    #    - gpt_prompted 类 description 非空
    # 4. 返回问题列表
    ...


def check_step2(progress_dir: Path = None, template_path: Path = None) -> Dict[str, Any]:
    """Step 2 自检: 内容生成质量 + golden reference 对比.
    
    Returns:
        {
            "passed": bool,
            "issues": [{"shape": str, "problem": str, "fix_hint": str}, ...],
            "summary": str
        }
    """
    # 实现细节:
    # 1. 读 03a-build_shape_content.json
    # 2. 对每个 strategy != "skip" 的 shape:
    #    - content 非空
    #    - len(content) 在 budget * [0.5, 1.2] 范围内
    #    - required_keywords 都出现在 content 中
    # 3. 加载 golden reference（从 01-shape_detail_com.json 的 text 字段）
    # 4. 对 gpt_prompted shape 做结构相似度比对:
    #    - 段落数差异 <= 30%
    #    - 列表项数差异 <= 30%
    # 5. 返回问题列表 + 修复建议
    ...


def load_golden_reference(progress_dir: Path = None) -> Dict[str, str]:
    """从 01-shape_detail_com.json 加载每个 shape 的原始文本作为 golden reference."""
    ...
```

**设计原则**：
- 纯 Python 函数，无 LLM 调用
- 可通过命令行调用：`python -c "from pipeline.self_check import check_step1; import json; print(json.dumps(check_step1()))"`
- 返回 JSON 结构，agent 可解析

### 6.3 保留（轻度简化）

| 文件 | 动作 |
|------|------|
| `ppt_pipeline_common.py` | 保留；去除 `02b` 相关的 helper（如果有）|
| `01_shape_detail.py` | 保留 |
| `01b_auto_annotate.py` | 保留 |
| `02_shape_analysis.py` | 保留；移除与 `02b` 相关的多-sheet 逻辑（如果有）|
| `03a_build_shape.py` | 保留；移除 `--sheet` 参数（多-sheet 遗留）|
| `03b_build_ppt_com.py` | 保留（内置自检已完善）|
| `fix_chart_link.py` | 保留 |
| `prompt_templates/` | 保留 |

### 6.4 Pipeline 变动验证

```bash
# 删除后确认 orchestrator / agent 不再引用
grep -r "02b_iteration_setup" . --include="*.py"
grep -r "04_shape_diff_test" . --include="*.py"
grep -r "02b_iteration_setup" .claude/
grep -r "04_shape_diff_test" .claude/
```

---

## 7. Slash Command 设计

### 7.1 新增 3 个步骤命令

**`.claude/commands/step1.md`**

```markdown
执行步骤1：分析 PPT 模板。

调用 step1-analyzer agent，完成：
- Python pipeline 提取 shape 结构（01_shape_detail + 01b_auto_annotate）
- 内部自检循环（对比模板完整性）
- 自检失败时由 LLM 修复批注
- 最多循环 2 次

完成后打印摘要，并打开 Excel 供审核。
```

**`.claude/commands/step2.md`**

```markdown
执行步骤2：构建 GPT prompt + 生成内容。

调用 step2-architect agent，完成：
- Python pipeline 生成 prompt 并调用 GPT（02 + 03a）
- 内部自检循环（对比 golden reference）
- 自检失败时由 LLM 重写 prompt 并重新调 GPT
- 最多循环 2 次

前置要求：已完成步骤1。
完成后打印摘要，并打开 Excel 供审核。
```

**`.claude/commands/step3.md`**

```markdown
执行步骤3：构建 & 交付 PPT。

调用 step3-builder agent，完成：
- 智能检测 prompt 是否更新（F1）
- Python pipeline 通过 COM 写入 PPT（03b，内置自检已完善）
- 失败时诊断问题层级，建议回到对应步骤

前置要求：已完成步骤2。
完成后打印摘要，并打开 PPT 供审核。
```

### 7.2 重命名 curator 命令

**当前**：`.claude/commands/role-curator.md`
**新名**：`.claude/commands/curator.md`（更简洁）

调用方式：`/curator`（不再宣传 `@curator`）

### 7.3 旧命令处理

**保留不动**：`c-pr, c-psh, md-update, safe-commit, today`
**重命名**：`role-curator.md` → `curator.md`

---

## 8. 用户场景覆盖（S1~S9）

> 以下 9 个场景从 plan2 继承，在新架构下重新验证。

### S1: 首次使用（全新模板）
```
用户: 选 0(全自动) → orchestrator 调 step1→step2→step3 → 出 PPT
```
✅ 无断层。每个 agent 内部自检通过才进入下一步。

### S2: 改 prompt 后重跑步骤3（最常见）
```
用户: 审核 PPT → 发现问题 → 改 xlsx prompt → 选 3
```
✅ **F1 内嵌在 step3-builder**：Attempt 1 开始前比较 xlsx.mtime vs content.json.mtime，若 xlsx 更新则自动重跑 03a。

### S3: 改批注后重跑步骤2
```
用户: 选 1 → 审核 → 改 strategy/description → 选 2
```
✅ **F2 内嵌在 step2-architect**：检测 xlsx 是否已有 GPT-prompt Text；有 + 手动模式 → 询问；无 → 完整执行。

### S4: 步骤1 重跑
```
用户: 跑过步骤1 → 手工改批注 → 又选步骤1
```
✅ **F3 内嵌在 step1-analyzer**：比较模板 mtime vs JSON mtime，模板未变 → 仅重跑 01b + Attempt 2，保护批注。

### S5: 步骤0 在已有进度上运行
```
用户: 有部分进度 → 选 0(全自动)
```
✅ step1/2 内部的 F2/F3 自动生效。

### S6: 跳步运行
```
用户: 直接选 3，但没跑过 1 和 2
```
✅ **F4 内嵌在各 agent**：前置检查缺失 → 报错 + 引导 "请先运行【步骤X】"。

### S7: 残留报告误导
```
用户: 多次运行 → 看到旧报告
```
✅ **F5 内嵌在各 agent**：每个 agent 启动时清理自己的输出报告。

### S8: PPT 效果差，根因在步骤1
```
用户: 出 PPT → 发现 strategy 推断错 → 回到步骤1
```
✅ 单步回退天然支持 + F3 保护。

### S9: Excel 未关闭
```
用户: 步骤1 后打开 Excel 审核 → 忘记关 → 选步骤2
```
✅ **F6 内嵌在 step2/step3**：启动前用 COM 测试锁定，锁定 → 提示关闭。

---

## 9. 断层修复总表（F1~F6）

| 编号 | 场景 | 修复位置 | 实现方 |
|------|------|---------|--------|
| F1 | prompt 更新 → step3 用旧 content | `step3-builder` Attempt 1 前 | Agent 内嵌 |
| F2 | 重跑 step2 覆盖手工 prompt | `step2-architect` 启动时 | Agent 内嵌 |
| F3 | 重跑 step1 覆盖手工批注 | `step1-analyzer` 启动时 | Agent 内嵌 |
| F4 | 跳步运行缺前置产物 | 各 agent 前置检查 | Agent 内嵌 |
| F5 | 残留报告误导 | 各 agent 启动时清理 | Agent 内嵌 |
| F6 | Excel 被锁定 | step2/step3 启动时 | Agent 内嵌 |

**新设计的优势**：所有断层修复都在 agent 内部，orchestrator 不再关心这些细节。

---

## 10. 职责边界

### 10.1 Orchestrator 的职责（最小化）

```
✅ 菜单交互
✅ 模板/数据选择（_select_template）
✅ 调用对应 agent
✅ 打开 Excel/PPT 供审核
❌ 不跑 pipeline 脚本
❌ 不写自检逻辑
❌ 不处理断层修复
```

### 10.2 Agent 的职责（完整自治）

```
✅ 跑 pipeline 脚本（Bash 调用）
✅ 内部自检循环
✅ LLM 修复
✅ 前置检查 (F4)
✅ 清理旧报告 (F5)
✅ 锁定检测 (F6)
✅ 智能跳过/保护 (F1/F2/F3)
✅ 报告生成
```

### 10.3 用户的职责

```
✅ 选模板/数据
✅ 选菜单
✅ 审核 Excel/PPT
✅ Pipeline 代码 bug 时直接在 Claude Code 主对话中说明修复需求
✅ 一轮结束后 /curator 沉淀经验（可选）
```

---

## 11. 实施顺序

### Phase 1: Pipeline 层改动
1. **新增** `pipeline/self_check.py`
   - `check_step1()`
   - `check_step2()`
   - `load_golden_reference()`
2. **删除** `pipeline/02b_iteration_setup.py`
3. **删除** `pipeline/04_shape_diff_test.py`
4. **清理** `03a_build_shape.py` 中的 `--sheet` 参数分支（如有）
5. **清理** `02_shape_analysis.py` 中的多-sheet 逻辑（如有）
6. **py_compile** 验证所有 pipeline 脚本

### Phase 2: Agent 定义
7. **新建** `.claude/agents/step1-analyzer.md`
8. **新建** `.claude/agents/step2-architect.md`
9. **新建** `.claude/agents/step3-builder.md`
10. **重命名** `.claude/agents/05-curator.md` → `.claude/agents/curator.md`
11. **归档** `.claude/agents/01-analyst.md` `02-builder.md` `03-reviewer.md` `04-developer.md` → `.claude/agents/_archive/`
   - 注意：developer 也归档（不再保留）

### Phase 3: Slash Commands
12. **新建** `.claude/commands/step1.md`
13. **新建** `.claude/commands/step2.md`
14. **新建** `.claude/commands/step3.md`
15. **重命名** `.claude/commands/role-curator.md` → `.claude/commands/curator.md`

### Phase 4: Orchestrator 瘦身
16. **重写** `main()` 使用新菜单
17. **重写** `PPTOrchestrator.run(step)` 为 agent 调度器
18. **新增** `_call_agent(agent_name)` 方法
19. **简化** `__init__`
20. **删除** 11 个废弃方法（见 5.3 节）
21. **精简** `AGENT_CONFIGS` 和 `AGENT_DISPLAY`
22. **py_compile** 验证

### Phase 5: CLAUDE.md 精简 + 知识外置
23. **新建** `.claude/memory/reference_pipeline_repair.md`（替代 developer agent 的知识载体）
    - pipeline 文件清单和职责
    - 常见修复类型表（列名错误 / 策略路由 / COM 写入）
    - 技术栈约束（pywin32 / 禁用 python-pptx / 禁用 openpyxl-pandas）
    - 修复后自检要求（py_compile + 重跑 pipeline）
    - 引用 ppt_pipeline_common.py 的关键 helper 函数清单

24. **CLAUDE.md 最小改动**（保持精简，不内嵌细节）：
    - 第 1 节"项目结构"：移除 `02b_iteration_setup.py` 和 `04_shape_diff_test.py`，新增 `self_check.py`
    - 第 1 节"项目结构"：更新 agent 清单（删除 analyst/builder/reviewer/developer，新增 step1/step2/step3-*）
    - 第 3 节"启动方式"：菜单选项从 6 项更新为 4 项
    - 第 5 节"详情索引"表：
      - 删除已归档 agent 的索引行
      - 新增 1 行：`\| Pipeline 代码修复指引 \| .claude/memory/reference_pipeline_repair.md \|`
      - 新增 1 行：`\| Step1/2/3 Agent 定义 \| .claude/agents/step1-analyzer.md (等) \|`
    - 第 6 节"变更记录"：追加 plan3 重构记录

25. **关键原则**：CLAUDE.md 不内嵌任何代码修复细节，只通过索引指向 memory 文件

### Phase 6: 验证
26. `python orchestrator.py` 启动 → 菜单正确
27. 选步骤1 → step1-analyzer 被调用 → 内部循环成功
28. 选步骤2 → step2-architect 被调用 → 自检对比 golden reference
29. 选步骤3 → step3-builder 被调用 → 出 PPT
30. 选全自动 → 三个 agent 串联成功
31. **场景验证**：S2/S3/S4/S6 各验证一次

---

## 12. 风险与回退

### 12.1 主要风险

| 风险 | 缓解措施 |
|------|---------|
| Agent 内自检逻辑不稳定 | `self_check.py` 纯 Python 实现，agent 只调用不重写 |
| Agent Bash 调用失败 | 保留 orchestrator 的 AgentExecutor 错误处理 |
| 删除 02b/04 影响未知代码 | Phase 1 最后全局 grep 确认无引用 |
| LLM 修复导致数据污染 | 每次修复后强制再自检；超限则报告失败不继续 |

### 12.2 回退方案

- plan2.md 保留（不删除）
- plan3.md 执行前 commit 一次，Phase 1 后再 commit 一次
- 旧 agent 文件归档到 `_archive/`，不直接删除
- 02b 和 04 先移到 `pipeline/_archive/`，确认无引用再彻底删除

---

## 13. 预期收益

| 指标 | 旧架构 | 新架构 |
|------|--------|--------|
| 菜单选项 | 6 | 4 |
| Agent 总数 | 5 (analyst/builder/reviewer/developer/curator) | **4** (step1/step2/step3-* + curator) |
| 主线 agent | 3 (analyst/builder/reviewer) | 3 (step1/step2/step3-*) |
| 辅助 agent | 2 (developer/curator) | **1** (curator) |
| 循环模型 | 整体循环 | 局部循环 |
| Agent 职责 | 按能力切分 | 按步骤切分 |
| Orchestrator 代码量 | ~1700 行 | 预计 < 700 行 |
| Pipeline 文件数 | 8 | 7（-2 +1） |
| 自检位置 | orchestrator + 03b | self_check.py + 03b + agent |
| 入口数 | 1（菜单） | **2**（菜单 + slash command）|
| 代码修复路径 | developer agent | Claude Code 主对话 |
| 一键全自动 | ✅ 有但效果差 | ✅ 预期效果显著提升 |

---

## 14. 待用户确认的点

以下问题在 plan3 中已做默认决策，如需调整请提出：

1. **旧 agent 文件处理**：默认归档到 `.claude/agents/_archive/`，不直接删除（包括 developer）
2. **旧 pipeline 文件处理**：默认先移到 `pipeline/_archive/`，验证无引用后再彻底删除
3. **agent 命名**：已确认 `step1-analyzer / step2-architect / step3-builder`
4. **orchestrator 保留**：已确认保留
5. **developer agent 删除**：已确认删除（代码修复回归 Claude Code 主对话）
6. **curator 调用方式**：已确认仅 slash command (`/curator`)，不再支持 `@curator`
7. **`02_shape_analysis.py` / `03a_build_shape.py` 的深度简化**：plan3 默认只做轻度清理（去除多-sheet 参数），不改核心逻辑

---

## 15. 总结

**plan3 vs plan2 核心差异**：

| 维度 | plan2 | plan3 |
|------|-------|-------|
| Agent 架构 | 5 个能力型 agent 原封不动 | **3+1 步骤型架构**，主线彻底重命名，删除 developer |
| 自检位置 | orchestrator 内的 `_self_check_step1/2` | `pipeline/self_check.py` + agent 内部 |
| 修复策略 | 先 inject_fix_constraints 再重跑 | 第一轮 Python，第二轮 Agent LLM |
| Orchestrator | 瘦身但仍跑 pipeline | 完全下放，只做调度 |
| Pipeline | 不动 | 删 2 个 + 新增 1 个 |
| 入口 | 仅菜单 | **菜单 + slash command**（不再用 @ mention）|
| 代码修复 | developer agent | **Claude Code 主对话** |

**plan3 是 plan2 的彻底进化版**，两者不兼容，选 plan3 就废弃 plan2。
