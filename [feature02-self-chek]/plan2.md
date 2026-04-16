# Plan: Orchestrator 重构 — 全局循环 → 局部自检循环

## Context

用户使用全局循环（分析→构建→检验→下一轮）发现体验差、质量不高。借鉴 HTML→PPT 项目经验，改为**局部循环**：每步内部自检通过后才进入下一步。同时简化菜单（6选项→4选项），移除 Step 4 自动验收（改为用户人工审核）。

---

## 新菜单（一字不改）

```
🎯 请选择运行模式:

  0️⃣  <全自动> ── 分析 → 构建 → 交付ppt
  1️⃣  步骤1 —— 分析（新）PPT 模板
  2️⃣  步骤2 —— 构建 prompt
  3️⃣  步骤3 —— 构建 & 交付 ppt
```

---

## 用户场景分析 & 断层修复

### 场景 S1: 首次使用（全新模板）

```
用户操作: 放入新 pptx + xlsx → 选 0(全自动) 或 选 1
期望: 从零提取 shape → 生成 prompt → 出 PPT
```

无断层。Step 1→2→3 正常串联。

### 场景 S2: 改 prompt 后重跑（最常见的迭代场景）

```
用户操作: 跑完步骤3 → 审核PPT发现内容问题 → 手工改 Excel 里的 prompt → 选 3
期望: 用新 prompt 重新调 GPT + 出 PPT
```

**断层**: Step 3 只跑 03b，读的是旧的 `03a-build_shape_content.json`，不会重新调 GPT。

**修复 F1**: Step 3 启动时比较 `xlsx.mtime` vs `03a-build_shape_content.json.mtime`。
如果 xlsx 更新 → 自动先跑 `03a --execute-prompts` 再跑 `03b`。
控制台打印: `[智能检测] prompt 已更新，自动重新调用 GPT ...`

### 场景 S3: 改批注后重跑 Step 2

```
用户操作: 跑完步骤1 → 审核 → 手工改了 strategy/description → 选 2
期望: 用新批注重新生成 prompt + 调 GPT
```

**断层**: Step 2 跑 `03a --assemble-only` 会用 02 的分析结果重新组装 prompt，**覆盖**用户手工改的 `GPT-prompt Text`。

**修复 F2**: Step 2 启动时检测 xlsx 中 `GPT-prompt Text` 是否已有内容（调 `read_gpt_prompts_from_xlsx()`）。
- 如果已有 prompt → 询问用户: `检测到已有 prompt，是否保留？[Y=保留并直接执行GPT / n=重新组装]`
- 全自动模式下默认保留（不覆盖）

### 场景 S4: Step 1 重跑（重新分析同一模板）

```
用户操作: 跑过步骤1 → 手工改了批注 → 又选了步骤1
期望: 重新提取 shape（可能模板有改动），但不丢失手工批注
```

**断层**: `01_shape_detail.py` 会重新生成 xlsx，覆盖用户的手工批注。

**修复 F3**: Step 1 检测模板文件 mtime vs JSON mtime:
- 模板更新了 → 重新提取（警告: `⚠️ 模板已更新，将重新提取 shape，已有批注可能被重置`）
- 模板未更新 → 跳过 01 提取，仅重跑 `01b + Analyst LLM`（保护现有批注）
- 控制台明确告知用户跳过原因

### 场景 S5: Step 0（全自动）在已有进度上运行

```
用户操作: 已跑过步骤1 → 手工改了批注 → 选 0(全自动)
期望: 跳过重复的提取，用现有批注直接走完全流程
```

**修复**: Step 0 调用的是同样的 Step 1/2/3 方法，F3 和 F2 的保护逻辑自动生效:
- Step 1 内部: 模板未变 → 跳过提取，保护批注
- Step 2 内部: 全自动模式默认保留已有 prompt
- 无需额外处理

### 场景 S6: 跳步运行（缺前置产物）

```
用户操作: 直接选步骤3，但没跑过步骤1和2
期望: 清晰报错，告知先跑哪一步
```

**修复 F4**: 每个 Step 启动时做前置检查，失败时给出明确引导:
- Step 2: `❌ 缺少 01-shape_detail_com.json，请先运行【步骤1】`
- Step 3: `❌ 缺少 03a-build_shape_content.json，请先运行【步骤2】`
  - 但如果 xlsx 有 prompt 且 content JSON 不存在 → `❌ 缺少 GPT 生成内容，请先运行【步骤2】`

### 场景 S7: 残留报告误导

```
用户操作: 跑过步骤3 → 修改了内容 → 再跑步骤3
期望: 看到的是本次运行的报告，不是上次残留的
```

**修复 F5**: 每个 Step 启动时清理自己的输出报告:
- Step 1: 删旧的 `01-shape_fingerprint_map.json`（如模板变了）
- Step 2: 删旧的 `03a-build_shape_content.json`, `03a-pending_prompts.json`
- Step 3: 删旧的 `03b-self_check_report.md`, `03b-build_ppt_report.md`, `04-diff_*.json/md`

### 场景 S8: PPT 效果差，根因在 Step 1 的批注

```
用户操作: 跑完全流程 → PPT 质量差 → 意识到是 strategy 或 description 推断有误
期望: 回到步骤1修正，然后重新走步骤2→3
```

无断层。用户选 Step 1 → 修正批注 → 选 Step 2 → 选 Step 3。
新菜单天然支持单步回退。加上 F3 的保护，Step 1 不会覆盖未变的模板提取。

### 场景 S9: Excel 未关闭导致 COM 写入失败

```
用户操作: 步骤1结束打开了 Excel → 用户忘记关闭 → 选了步骤2
期望: 不会静默失败
```

**修复 F6**: Step 2 和 Step 3 启动前，尝试以只读方式打开 xlsx 测试是否被锁定。
如果锁定 → `⚠️ 01-shape_detail.xlsx 正在被其他程序打开，请关闭后按 Enter 继续...`

---

## 改动范围

**唯一主改文件**: `orchestrator.py`

| 区域 | 动作 | 说明 |
|------|------|------|
| `main()` (L1572-1701) | **重写** | 新菜单 + 简化 dispatch |
| `PPTOrchestrator.__init__` (L461-479) | **简化** | 去掉 `max_rounds`, `skip_analyst_first_round`, `init_mode` |
| `run()` (L961-1467) | **重写** | 拆成 `run(step)` 调度 + 3 个 step 方法 |
| `AGENT_CONFIGS` (L77-108) | **精简** | 移除 `reviewer`, `developer` |
| `AGENT_DISPLAY` (L412-417) | **精简** | 移除 `reviewer`, `developer` |

### 删除的方法（死代码）
- `_check_review_passed()` (L899) — 不再作为流程门禁（04 降级为可选诊断）
- `_archive_round()` (L944) — 多轮归档，不再有全局多轮
- `_reviewer_llm_only_prompt()` (L857) — Reviewer agent 移除
- `_developer_prompt()` (L876) — Developer agent 移除
- `_builder_prompt_optimizer_prompt()` (L676) — fix-based 优化，被 Step 2 内循环取代
- `_builder_llm_only_prompt()` (L834) — 不再需要
- `_run_builder_pipeline()` (L638) — 被 step 方法取代
- `_run_prompt_only_pipeline()` (L668) — 被 step 方法取代
- `_run_03a_with_prompt_review()` (L597) — 逻辑内联到 Step 2
- `_prompts_exist()` (L659) — 热迭代检测不再需要

### 保留的方法
- `_run_pipeline()` / `_run_pipeline_step()` — 基础 subprocess 执行器
- `_detect_next_version_index()` / `_idx_to_version()` / `_record_version()` — Step 3 用
- `_verify_pptx_exists()` — Step 3 用
- `_analyst_phase2_prompt()` — Step 1 Analyst LLM 用（保留 full mode，去掉 targeted mode）
- `_load_enhanced_list()` / `_mark_enhanced()` — 可选保留
- `AgentExecutor`, `StateManager`, `ErrorHandler`, `ProgressMonitor` 类 — 保留

---

## 新架构

### `run(step)` — 薄调度层

```python
async def run(self, step: int) -> bool:
    if step == 0:
        ok = await self._run_step1_analyze()
        if ok: ok = await self._run_step2_build_prompt()
        if ok: ok = await self._run_step3_build_ppt()
        # 全自动完成后打开 PPT
        return ok
    elif step == 1:
        return await self._run_step1_analyze()
    elif step == 2:
        return await self._run_step2_build_prompt()
    elif step == 3:
        return await self._run_step3_build_ppt()
```

### Step 1: `_run_step1_analyze()` — 分析 PPT 模板

```
【F5】清理旧报告: 删 01-shape_fingerprint_map.json（如果模板变了）

【F3】智能跳过:
  if 模板 mtime > JSON mtime → 重新提取（01_shape_detail.py + 01b + Analyst）
  if 模板未变且 xlsx 存在 → 跳过 01 提取，仅重跑 01b + Analyst（保护已有批注）
  if 首次运行 → 完整提取

Pipeline: 01_shape_detail.py → 01b_auto_annotate.py
LLM: Analyst 增强批注
自检循环 (max 2 次):
  _self_check_step1():
    ① 01-shape_detail_com.json 存在且 new_shapes 非空
    ② 每个 shape 的 strategy 已赋值（非空、非 "(必填)"）
    ③ gpt_prompted shape 的 description 已赋值
  失败 → 重跑 01b + Analyst LLM
结束 → 非全自动时 os.startfile(xlsx) 供用户审核
```

### Step 2: `_run_step2_build_prompt()` — 构建 prompt

```
【F4】前置检查: 01-shape_detail_com.json 和 xlsx 必须存在
  缺失 → ❌ 请先运行【步骤1】
【F6】检测 xlsx 是否被锁定 → 提示关闭
【F5】清理旧报告: 删 03a-build_shape_content.json, 03a-pending_prompts.json

【F2】Prompt 保护:
  检测 xlsx 中 GPT-prompt Text 是否已有内容
  - 已有 + 手动模式 → 询问: 保留已有 prompt 还是重新组装？
  - 已有 + 全自动 → 默认保留，跳过 --assemble-only
  - 无 prompt → 完整执行 02 + 03a 两阶段

Pipeline（完整模式）: 02_shape_analysis.py → 03a --assemble-only → 03a --execute-prompts
Pipeline（保留模式）: 仅 03a --execute-prompts

自检循环 (max 2 次):
  _self_check_step2():
    ① 03a-build_shape_content.json 存在
    ② 每个 strategy≠skip 的 shape 有非空 content
    ③ content 长度在 budget 容差范围内（50%~120%）
    ④ gpt_prompted shape 的 required_keywords 出现在 content 中
    ⑤ 对标模板原始文本：从 01-shape_detail_com.json 提取模板原始 text 作为 golden reference，
       比对 GPT 生成内容与模板原文的结构相似度（段落数、列表项数、关键短语覆盖率）
  失败 → 将自检失败原因注入 prompt 约束（_inject_fix_constraints），
         再重跑 03a --execute-prompts
结束 → 非全自动时 os.startfile(xlsx) 供用户审核
```

### Step 3: `_run_step3_build_ppt()` — 构建 & 交付 PPT

```
【F4】前置检查: 03a-build_shape_content.json 必须存在
  缺失 → ❌ 请先运行【步骤2】
【F6】检测 xlsx 是否被锁定 → 提示关闭
【F5】清理旧报告: 删 03b-self_check_report.md, 03b-build_ppt_report.md,
     04-diff_result.json, 04-diff_semantic_report.md

【F1】智能检测 prompt 更新:
  if xlsx.mtime > 03a-build_shape_content.json.mtime:
    print("[智能检测] prompt 已更新，自动重新调用 GPT ...")
    运行 03a --execute-prompts（用最新 prompt 重新生成内容）

版本: _detect_next_version_index() → _record_version()
Pipeline: 03b_build_ppt_com.py --version X.X
  （03b 内部已有 4 步自检 + MAX_SELF_FIX=2 自动修复，无需外层循环）
读取 03b-self_check_report.md 显示结果

可选诊断: 运行 04_shape_diff_test.py 生成诊断报告（不阻断流程），
  用户审核 PPT 时可参考 04-diff_semantic_report.md 发现肉眼难察觉的精细问题
结束 → 非全自动时 os.startfile(pptx) 供用户审核
```

### `main()` 简化

```python
def main():
    _select_account()
    project_root = find_project_root()
    _select_template(project_root)

    # 新菜单（精确文字）
    print("\n🎯 请选择运行模式:\n")
    print("  0️⃣  <全自动> ── 分析 → 构建 → 交付ppt")
    print("  1️⃣  步骤1 —— 分析（新）PPT 模板")
    print("  2️⃣  步骤2 —— 构建 prompt")
    print("  3️⃣  步骤3 —— 构建 & 交付 ppt\n")

    choice → step (0-3), default='0'
    auto_mode = (step == 0)

    orch = PPTOrchestrator(project_root, max_budget, auto_mode)
    success = asyncio.run(orch.run(step))
    sys.exit(0 if success else 1)
```

### `__init__` 简化

```python
def __init__(self, project_root: Path, max_budget: float = 10.0, auto_mode: bool = False):
    # 去掉 max_rounds, skip_analyst_first_round, init_mode
```

---

## 自检函数详情

### `_self_check_step1()` → `list[dict]`
- 读 `01-shape_detail_com.json` → 计数 `new_shapes`
- 调 `parse_user_annotations()` 从 xlsx 读批注
- 检查每个 shape: `strategy_exact` 非空且非 `(必填)`
- 检查 `gpt_prompted` shape: `description` 非空且非 `(必填)`
- 返回问题列表，空列表 = 通过

### `_self_check_step2()` → `list[dict]`
- 读 `03a-build_shape_content.json`
- 遍历 items: 跳过 `strategy=="skip"`
- 检查 `content` 非空
- 检查 `len(content)` 在 `budget * 0.5` ~ `budget * 1.2` 范围内
- 检查 `required_keywords` 都出现在 content 中
- 对标模板原文：从 `01-shape_detail_com.json` 的 `new_shapes[*].text` 提取原文，
  比对段落数、列表项数、关键短语覆盖率；差异过大标记为问题
- 返回问题列表，空列表 = 通过

### `_inject_fix_constraints()`
- 输入: `_self_check_step2()` 返回的问题列表
- 对每个失败 shape，在 xlsx 的 `GPT-prompt Text` 末尾追加约束行（如 "字数不超过 X"、"必须包含关键词 Y"）
- 调用 `write_gpt_prompts_to_xlsx()` 写回
- 这样下一次 `03a --execute-prompts` 使用的就是增强后的 prompt

### `_check_xlsx_locked()` 【F6 新增】
- 尝试以只读方式通过 COM 打开 xlsx
- 成功 → 关闭，返回 False (未锁定)
- 失败 → 返回 True (被锁定)
- 调用处: Step 2 和 Step 3 启动时

### `_clean_stale_reports(step)` 【F5 新增】
- Step 1: 条件清理 fingerprint_map（仅模板变化时）
- Step 2: 删 03a-build_shape_content.json, 03a-pending_prompts.json, 03a-prompt_trace.json
- Step 3: 删 03b-self_check_report.md, 03b-build_ppt_report.md, 04-diff_result.json, 04-diff_semantic_report.md

---

## 断层修复总表

| 编号 | 场景 | 问题 | 修复 | 涉及 Step |
|------|------|------|------|-----------|
| F1 | 改 prompt → 选步骤3 | Step 3 用旧 content JSON | 比较 mtime，自动重跑 03a | Step 3 |
| F2 | 改批注 → 选步骤2 | 03a --assemble-only 覆盖手工 prompt | 检测已有 prompt，询问/默认保留 | Step 2 |
| F3 | 重跑步骤1 | 01 重新生成 xlsx 覆盖手工批注 | 模板未变则跳过提取，仅重跑 01b+LLM | Step 1 |
| F4 | 跳步运行 | 缺前置产物静默失败 | 明确报错 + 引导先跑哪步 | Step 2, 3 |
| F5 | 多次运行 | 残留报告误导 | 每步启动时清理自己的旧报告 | Step 1, 2, 3 |
| F6 | Excel 未关闭 | COM 写入失败 | 启动前检测锁定，提示关闭 | Step 2, 3 |

---

## 职责边界: Orchestrator vs Agent

### 设计原则

Orchestrator 只管**主流程自动化**（0→1→2→3 + 步骤内自检）。
超出主流程的问题，由用户通过 slash command / @ agent 手动介入。
**不在 orchestrator 里堆砌边缘 case 处理逻辑。**

### 边界划分

| 问题类型 | 谁处理 | 用户操作 |
|---------|--------|---------|
| strategy/description/params 有误 | 用户 → Excel 手工编辑 | 改 Excel → 重跑步骤 2 |
| GPT prompt 不满意 | 用户 → Excel 手工编辑 | 改 Excel → 重跑步骤 3 (F1 自动检测) |
| JSON 结构数据有误（原始文本提取错误、shape 遗漏） | **Agent** (analyst) | `@analyst` 或 `/role-analyst` 手动修复 JSON |
| Pipeline 代码 bug | **Agent** (developer) | `@developer` 或 `/role-developer` 修复代码 |
| PPT 格式/样式异常（非内容问题） | **Agent** (builder) | `@builder` 或 `/role-builder` 定向修复 |

### Agent 自检要求

手动调用的 Agent 同样需要内置自检循环，确保修复有效：

**Analyst Agent 自检**（修复 JSON / xlsx 后）:
1. 重新读取修改后的 JSON/xlsx
2. 对比修改前后：确认目标 shape 已修正
3. 运行 `_self_check_step1()` 同等逻辑验证整体一致性
4. 打印修改摘要 + 自检结果

**Builder Agent 自检**（修复批注/prompt 后）:
1. 读取修改后的 xlsx prompt
2. 验证目标 shape 的 prompt 包含所需约束
3. 可选：调 GPT 生成一次内容，对比预期
4. 打印修改摘要 + 自检结果

**Developer Agent 自检**（修复代码后）:
1. `py_compile` 验证修改的文件
2. 运行受影响的 pipeline step，验证不报错
3. 对比修复前后的输出差异
4. 打印修复摘要 + 测试结果

> 具体实现：在 `.claude/agents/01-analyst.md` 等 agent 定义文件中追加自检指令段落。
> 这是 agent 配置的改动，不是 orchestrator 代码的改动。

### 典型用户干预流程

```
场景: PPT 中某个 shape 内容完全错误，追溯到 JSON 原始文本提取有误

1. 用户审核 PPT → 发现问题
2. 打开 01-shape_detail.xlsx 对照 → 确认是提取问题（非 prompt 问题）
3. 调用: @analyst "Shape X 的原始文本提取有误，JSON 中记录的是 '...'，
         实际模板中应该是 '...'，请修正 01-shape_detail_com.json"
4. Analyst Agent:
   a. 通过 COM 读取模板 PPT 验证用户描述
   b. 修正 JSON 中对应 shape 的 text 字段
   c. 自检: 重新读取 JSON 确认修正生效
   d. 报告: "已修正 Shape X 的 text 字段，自检通过"
5. 用户: 重跑步骤 2 → 步骤 3
```

---

## 实施顺序

### Part 1: Orchestrator 重构 (orchestrator.py)
1. **添加** 工具方法: `_check_xlsx_locked()`, `_clean_stale_reports(step)`
2. **添加** `_self_check_step1()` 和 `_self_check_step2()` 方法
3. **添加** `_inject_fix_constraints()` 方法
4. **添加** `_run_step1_analyze()`, `_run_step2_build_prompt()`, `_run_step3_build_ppt()` 三个 step 方法
   - Step 1 内嵌 F3 模板变化检测
   - Step 2 内嵌 F2 prompt 保护 + F6 锁定检测
   - Step 3 内嵌 F1 智能 prompt 更新检测 + F6 锁定检测 + 可选 04 诊断
   - 每个 Step 启动时调 `_clean_stale_reports()` (F5)
   - 每个 Step 启动时做前置检查 (F4)
5. **重写** `run()` 为 step 调度器
6. **重写** `main()` 为新菜单
7. **简化** `__init__`，去掉废弃参数
8. **删除** 上述列出的死代码方法
9. **精简** `AGENT_CONFIGS` 和 `AGENT_DISPLAY`

### Part 2: Agent 自检指令 (`.claude/agents/*.md`)
10. **更新** `01-analyst.md`: 追加自检段落（修复后验证 JSON/xlsx 一致性）
11. **更新** `02-builder.md`: 追加自检段落（修复后验证 prompt 完整性）
12. **更新** `04-developer.md`: 追加自检段落（修复后 py_compile + 运行验证）

---

## 验证

1. `py_compile` 验证语法
2. `python orchestrator.py` 启动，确认新菜单显示正确（精确匹配用户要求的文字）
3. 选择步骤 1，确认 01 + 01b pipeline 执行 + 自检循环运行
4. 选择步骤 2，确认 02 + 03a pipeline 执行 + 自检循环运行
5. 选择步骤 3，确认 03b pipeline 执行 + 内置自检
6. 选择全自动，确认 1→2→3 串联执行无暂停
7. **场景验证**:
   - S2: 改 prompt → 选步骤3 → 确认自动重跑 GPT
   - S3: 选步骤2 → 确认询问是否保留已有 prompt
   - S4: 模板未变 → 选步骤1 → 确认跳过提取
   - S6: 直接选步骤3 → 确认报错引导
