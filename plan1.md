# Plan: Prompt-Centric 架构升级 — 冷启动/热迭代分离

## Context

**核心痛点**：当前修正轮通过 02b 修改"内容描述"（中间变量），再经 02→03a 重新组装 prompt。但 prompt 一旦生成，注释的使命就完成了。后续修正应直接编辑 prompt（终端变量），跳过中间环节。

**架构洞察**：完整 pipeline（01+01b+LLM增强+02+03a Phase1）是一次性的"冷启动"。后续所有轮次都是"热迭代"——只改 prompt → 调 GPT → 出 PPT。将两者混在同一流程是低效的根源。

**目标**：
- 新增选项 0（初始化）：冷启动专用，全新 PPT 分析
- Excel 不存在时：无论用户选什么，强制路由到选项 0
- 选项 1-4：统一为热迭代路径，跳过 Analyst LLM，直接操作 prompt
- 修正轮：fix.md → Builder LLM 直接修改 prompt → 用户审核 → 03a Phase 2 → 03b
- Reviewer：建议从"改注释"升级为"改 prompt"

**设计决策**：不新建 agent，复用 Builder（新 prompt-optimizer prompt）+ Reviewer（升级 fix 建议格式）

## 菜单设计

```
🎯 请选择运行模式:

  0️⃣  🆕 初始化 ── 全新 PPT 分析，从零构建结构和 prompt
  1️⃣  快速出图 ── 跑一轮就交付，适合赶进度
  2️⃣  标准打磨 ── 生成 → 验收 → 修正，两轮收工
  3️⃣  精雕细琢 ── 三轮迭代反复打磨，追求极致
  4️⃣  🤖 挂机托管 ── 全自动两轮，泡杯咖啡等结果
  5️⃣  🔍 单独验收 ── 只跑验收，检查最新 PPT 质量
```

### 路由逻辑

| 条件 | 行为 |
|------|------|
| Excel 不存在 | 无论用户选什么 → 强制走选项 0，打印提示 |
| 用户选 0 | 冷启动：01+01b+LLM增强+P1暂停+02+03a(full)+03b |
| 用户选 1-4 | 热迭代：跳过 Analyst LLM，prompt-centric 路径 |
| 用户选 5 | 单独验收（不变） |

### 选项 0 流程（冷启动）

```
[Analyst] Pipeline(01+01b) + LLM 增强注释
    ↓
  ⏸️ P1 — auto-open Excel，校准批注
    ↓
[Builder] 02 → 03a Phase1（组装 prompt）
    ↓
  ⏸️ PROMPT REVIEW — auto-open Excel，审核 prompt
    ↓
  03a Phase2（调 GPT）→ 03b → PPT
    ↓
  ✅ 初始化完成（不进 Reviewer）
```

### 选项 1-4 流程（热迭代）

```
[跳过 Analyst LLM] 仅确保 01 JSON 存在
    ↓
[Builder] ⏸️ PROMPT REVIEW → 03a Phase2 → 03b → PPT    ← 首轮
    ↓
[Reviewer] 04验收 → PASS/FAIL                            ← 选项2-4
    ↓ FAIL
  02b(sheet-only) → Builder LLM(改prompt) → ⏸️ → 03a Phase2 → 03b
    ↓
  循环至 max_rounds
```

## 文件清单

| 文件 | 改动量 | 说明 |
|------|--------|------|
| `orchestrator.py` | 大 | 菜单重构 + 选项0 + 强制路由 + 热迭代路径 |
| `pipeline/02b_iteration_setup.py` | 小 | 新增 `--sheet-only` flag |
| `.claude/agents/02-builder.md` | 小 | 修正轮改为编辑 prompt |
| `.claude/agents/03-reviewer.md` | 小 | fix 建议改为 prompt 级别 |
| `.claude/agents/01-analyst.md` | 小 | 标注一次性/冷启动角色 |
| `.claude/CLAUDE.md` | 中 | 更新工作流文档 |

## 改动详情

### 1. orchestrator.py（核心）

#### 1A. 新增 `_prompts_exist()` 方法

```python
def _prompts_exist(self) -> bool:
    """Check if GPT-prompt Text cells are already populated in Excel."""
    from pipeline.ppt_pipeline_common import read_gpt_prompts_from_xlsx
    prompts = read_gpt_prompts_from_xlsx()
    return len(prompts) > 0
```

#### 1B. 菜单新增选项 0 + 强制路由

```python
print("  0️⃣  🆕 初始化 ── 全新 PPT 分析，从零构建结构和 prompt")
print("  1️⃣  快速出图 ── 跑一轮就交付，适合赶进度")
# ... 2-5 不变
```

强制路由逻辑：
```python
if not SHAPE_DETAIL_XLSX.exists():
    if choice != "0":
        safe_print(f"⚠️  Excel 不存在，自动切换到 [0-初始化] 模式")
    choice = "0"
```

选项 0 处理：
```python
if choice == "0":
    init_mode = True
    max_rounds = 1  # 初始化只跑一轮完整流程
    # 走完整冷启动路径（含 Analyst LLM + P1 暂停）
```

#### 1C. 冷启动 vs 热迭代分支

主循环入口根据 `init_mode` 分流：

**冷启动（选项 0）**：
- 完整 Analyst：pipeline(01+01b) + LLM 增强
- P1 暂停：用户校准批注
- 完整 Builder：02 → 03a(full, assemble+execute) → 03b
- 不进 Reviewer，直接结束

**热迭代（选项 1-4）**：
- 跳过 Analyst LLM（仅确保 01 JSON 产物存在，不存在则跑 pipeline 01+01b）
- Builder 首轮：prompt review 暂停 → 03a Phase 2 → 03b
- 选项 2-4 进 Reviewer → 修正轮

#### 1D. 新增 `_run_prompt_only_pipeline(version, sheet_name)` 方法

```python
def _run_prompt_only_pipeline(self, version: str, sheet_name: str):
    """Hot iteration: prompts in Excel, skip 02 + 03a Phase 1."""
    self._run_pipeline("03a_build_shape.py", ["--execute-prompts"])
    self._run_pipeline("03b_build_ppt_com.py", ["--version", version])
```

#### 1E. 新增 `_builder_prompt_optimizer_prompt(sheet_name, fix_data)` 方法

指导 Builder LLM 直接修改 Excel 中的 GPT-prompt Text：

```
你是 PPT 内容优化师。根据验收报告，直接修改 Excel 中的 GPT prompt。

## 任务
读取 01-shape_detail.xlsx 最新 sheet 的 "GPT-prompt Text" 单元格，
根据以下问题修改对应 shape 的 prompt 文本。

## 问题清单
{fix_items}

## 规则
- 只改有问题的 shape 的 prompt，其余不动
- 用 write_gpt_prompts_to_xlsx() 写回
- 不要修改"内容描述"等注释字段
```

#### 1F. 更新 `_reviewer_llm_only_prompt()` 方法

- 去掉"修改注释"措辞，改为"直接修改 GPT-prompt Text"
- fix item 增加 `prompt_fix_suggestion` 字段
- 保留 fix_type 分类不变

#### 1G. 修正轮流程重写

```python
# 1. 02b --sheet-only（仅创建新 sheet，继承上轮 prompt）
self._run_pipeline("02b_iteration_setup.py", ["--version", version, "--sheet-only"])

# 2. Builder LLM（直接改 prompt）
builder_prompt = self._builder_prompt_optimizer_prompt(sheet_name, fix_data)
self._call_agent("builder", builder_prompt)

# 3. Prompt 审核暂停（非 auto_mode 时）
if not self.auto_mode:
    os.startfile(str(xlsx_path))
    print("⏸️  PROMPT REVIEW — 请审核修改后的 GPT prompt...")
    input()

# 4. prompt-only pipeline（03a Phase 2 + 03b）
self._run_prompt_only_pipeline(version, sheet_name)
```

### 2. pipeline/02b_iteration_setup.py

#### 2A. 新增 `--sheet-only` CLI flag

```python
ap.add_argument("--sheet-only", action="store_true",
                help="Only create new sheet (inherit prompts), skip annotation fixes")
```

#### 2B. main() 中提前返回

```python
if args.sheet_only:
    safe_print(f"[OK] Sheet-only 模式：{sheet_name} 已创建，跳过注释修正")
    return 0
```

### 3. .claude/agents/02-builder.md

修正轮职责：~~"通过 COM 精调 xlsx 批注"~~ → "直接修改 xlsx 中的 GPT-prompt Text"

新增：
```
## 修正轮 Prompt 编辑
- 用 read_gpt_prompts_from_xlsx() 读取当前 prompt
- 根据 fix 报告修改有问题的 prompt
- 用 write_gpt_prompts_to_xlsx() 写回
- 不修改"内容描述"等注释字段
```

### 4. .claude/agents/03-reviewer.md

fix 建议格式：~~"在「内容描述」中添加 XXX"~~ → "将 GPT-prompt Text 中的 XXX 改为 YYY"

### 5. .claude/agents/01-analyst.md

添加说明：
```
> Analyst 是冷启动角色（选项 0）。热迭代模式（选项 1-4）自动跳过。
```

### 6. .claude/CLAUDE.md

#### 6A. 更新混合工作流表格

| Agent | 冷启动（选项0） | 热迭代（选项1-4） |
|-------|---------------|-----------------|
| Analyst | 01+01b + LLM增强 | 跳过 LLM（仅确保JSON） |
| Builder首轮 | 02→03a(full)→03b | prompt review → 03a Phase2 → 03b |
| Builder修正轮 | — | 02b(sheet-only) → LLM改prompt → 03a Phase2 → 03b |
| Reviewer | 不进入 | 04测试 + LLM prompt级建议 |

#### 6B. 更新流程图

```
━━━ 选项 0: 冷启动 ━━━
[Analyst] 01+01b + LLM → ⏸️ P1 → [Builder] 02→03a→03b → ✅

━━━ 选项 1-4: 热迭代 ━━━
⏸️ Prompt Review → 03a Phase2 → 03b → PPT
    ↓ (选项2-4)
[Reviewer] → FAIL → 02b(sheet-only) → Builder LLM(改prompt) → ⏸️ → 03a Phase2 → 03b
```

## 最终对比

| | 现有流程 | 新流程 |
|--|---------|--------|
| 冷启动/迭代 | 混合在一起，每轮都判断 | 菜单分离，选项 0 vs 1-4 |
| Analyst | 每轮可能运行 | 仅选项 0 运行 LLM |
| 修正轮 | 改注释→02→03a→03b (4步) | 改prompt→03a Phase2→03b (3步) |
| LLM 关注点 | 内容描述（中间变量） | GPT-prompt Text（终端变量） |
| Excel 不存在 | 可能报错 | 强制路由到选项 0 |

## 实现顺序

1. `pipeline/02b_iteration_setup.py` — `--sheet-only` flag（最小改动）
2. `.claude/agents/` — 3 个 agent spec 更新（文档级）
3. `orchestrator.py` — 菜单 + 路由 + 冷启动/热迭代分支 + 修正轮重写
4. `.claude/CLAUDE.md` — 文档同步

## 验证

1. `python -m py_compile pipeline/02b_iteration_setup.py orchestrator.py`
2. 删除 Excel → `python orchestrator.py` 选 1 → 自动路由到选项 0，跑冷启动
3. Excel 已存在 → `python orchestrator.py` 选 0 → 冷启动（重新分析）
4. Excel 已存在 → `python orchestrator.py` 选 1 → 跳过 Analyst LLM，prompt review → PPT
5. `python orchestrator.py` 选 2 → 首轮同上；修正轮：02b --sheet-only → Builder 改 prompt → 03a Phase 2 → 03b
6. 手动确认：修正轮不修改"内容描述"列，只修改"GPT-prompt Text"列
