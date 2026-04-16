# Orchestrator 优化：Pipeline-First, LLM Fallback

## Context

用户运行 orchestrator 后发现耗时严重超出预期：
- Step1 仅分析 PPT 模板就花了 **217s**
- Step2 连续失败 3 次，共花费 **710s+** 后仍失败

根因：当前每个步骤都通过 `claude -p` 子进程执行。即使 Python pipeline 脚本只需几秒，Claude 会话的启动、理解角色、逐 turn 调用 Bash、推理等开销使每步膨胀到 200-400s。

**优化思路**：orchestrator 直接跑 Python pipeline + self_check。只在自检失败时才启动 Claude agent 做 LLM 修复。

---

## 优劣势分析

| | 优势 | 劣势 |
|--|------|------|
| **速度** | Happy path: 每步 10-30s（step2 含 GPT 调用约 30-90s） | - |
| **成本** | Pipeline 成功时零 LLM token 消耗 | - |
| **可靠性** | 减少 subprocess 启动失败风险 | Pipeline 脚本 crash 时需 fallback 到全量 agent |
| **调试** | 直接看 Python 脚本 stdout/stderr，不被 Claude stream-json 包裹 | - |
| **复杂度** | - | orchestrator 新增 ~150 行代码（3 个新方法） |
| **LLM 修复** | 传入精确的失败上下文，agent 跳过 Attempt 1 直奔修复 | agent 需理解 "REPAIR MODE" 指令（通过 task_prompt） |

---

## 实现方案

### 修改文件：`orchestrator.py`（唯一需要改的文件）

#### 1. 新增 `_run_pipeline(self, step: int) -> Tuple[bool, str]` (~50 行)

用 `subprocess.run()` 直接执行 Python 脚本，返回 `(success, error_detail)`。

各步骤执行的脚本：

| Step | 脚本序列 |
|------|---------|
| 1 | `01_shape_detail.py` → `01b_auto_annotate.py` → `02_shape_analysis.py` |
| 2 | `02_shape_analysis.py` → `03a_build_shape.py --assemble-only` → `03a_build_shape.py --execute-prompts` |
| 3 | `03b_build_ppt_com.py --version X.X` |

关键细节：
- 用 `sys.executable` 确保同一 Python 解释器
- `env` 继承 `os.environ`（包含 `PPT_TEMPLATE_PATH` / `PPT_EXCEL_PATH`）
- 每个脚本 `timeout=300`（5 分钟，主要为 step2 的 GPT 调用留余量）
- 任一脚本 returncode != 0 → 立即返回 `(False, error_detail)`
- 打印每个脚本的运行状态和耗时

#### 2. 新增 `_run_self_check(self, step: int) -> Tuple[bool, Dict]` (~30 行)

| Step | 自检方式 |
|------|---------|
| 1 | 直接 import `check_step1()` — 纯 JSON 读取，无 COM 依赖，毫秒级 |
| 2 | 直接 import `check_step2()` — 同上 |
| 3 | 读取 `03b-self_check_report.md`，查找 `"结论：PASS"` 或 `"结论：FAIL"` |

Step 3 特殊：`03b_build_ppt_com.py` 内置 `MAX_SELF_FIX=2` 自检循环，且**始终返回 exit code 0**（L641），pass/fail 只体现在报告文件中。

返回结构与 `self_check.py` 一致：`{"passed": bool, "issues": [...], "summary": str}`

#### 3. 新增 `_run_step(self, step: int) -> bool` (~40 行) — 核心调度

```
Pipeline 直跑 → 成功？
  ├─ 脚本 crash (returncode!=0) → _call_agent(完整模式)
  └─ 脚本成功 → self_check
       ├─ PASS → 记录 synthetic result，返回 True
       └─ FAIL → _call_agent(REPAIR MODE，传入失败详情)
```

- 成功时创建 synthetic `ExecutionResult`（cost=0, tokens=0）供 `display_summary()` 使用
- 记录实际 pipeline + self_check 耗时到 `duration` 字段

#### 4. 修改 `_call_agent(self, agent_name, failure_context=None)` (~15 行改动)

新增可选参数 `failure_context: Optional[str]`：
- **有值时**：在 task_prompt 中追加 `## REPAIR MODE` 区块，包含自检失败详情，指示 agent 跳过 Attempt 1 直接执行 Attempt 2
- **无值时**：保持原有行为（完整流程）

不修改 agent spec 文件（`.claude/agents/step*.md`），所有上下文通过 task_prompt 传递。

#### 5. 修改 `run(self, step: int)` (~20 行改动)

将 `_call_agent()` 调用替换为 `_run_step()`：

```python
# 原来:
if not await self._call_agent("step1-analyzer"):
# 改为:
if not await self._run_step(1):
```

全自动模式（step=0）遍历 [1, 2, 3]，任一步失败则终止。

---

## 预期效果

| 场景 | 当前耗时 | 优化后耗时 |
|------|---------|-----------|
| Step 1 (pipeline 成功) | ~217s | ~5-10s |
| Step 2 (pipeline + GPT 成功) | ~400s | ~30-90s (GPT API 为主) |
| Step 3 (COM 写入成功) | ~200s | ~10-30s |
| Step 2 (pipeline 失败，需 LLM 修复) | ~400s × 3 = 1200s | ~60s + 一次 agent 会话 |

---

## 风险与缓解

| 风险 | 缓解 |
|------|------|
| COM 句柄泄露（脚本 crash 时） | `03b_build_ppt_com.py` 已有 `try/finally` 清理 COM；其他脚本短暂使用 COM |
| `self_check.py` import 副作用 | 已验证：只用 json/re/pathlib，无 COM 依赖 |
| Step 3 report 文件不存在 | 03b 被阻断（如无 content JSON）时不生成 report → `_run_self_check(3)` 视为 pipeline crash |
| Agent 不理解 REPAIR MODE | task_prompt 中明确写"跳过 Attempt 1"+ 附带完整失败信息 |

---

## 不改的文件

- `pipeline/*.py` — 所有 pipeline 脚本不变
- `pipeline/self_check.py` — 不变
- `.claude/agents/step*.md` — agent spec 不变
- `ppt_pipeline_common.py` — 不变

---

## 验证方式

1. `python -m py_compile orchestrator.py` — 语法检查
2. Step 1 单步运行：选模板 → 步骤1 → 确认 ~10s 内完成 + xlsx 生成
3. Step 2 单步运行：步骤2 → 确认 GPT 调用正常 + self_check 结果输出
4. Step 3 单步运行：步骤3 → 确认 PPT 生成 + report 解析
5. 全自动模式（0）：确认 1→2→3 串联正常
6. 模拟 self_check 失败：临时修改 self_check 返回 False，确认 LLM agent 被正确启动并收到 REPAIR MODE 上下文
