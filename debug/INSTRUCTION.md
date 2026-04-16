# 运行模式解析

> `python orchestrator.py` 启动后选择模式 0-3

---

## 菜单总览

| 选项 | 名称 | 说明 |
|------|------|------|
| **0** | 全自动 | Step1 → Step2 → Step3 串联，含自动回退循环 |
| **1** | 步骤1 | 分析 PPT 模板 |
| **2** | 步骤2 | 构建 prompt + 调 GPT 生成内容 |
| **3** | 步骤3 | 写入 PPT + 自检修复 |

---

## 每个步骤的通用执行流程

所有步骤（1/2/3）共享同一套 `_run_step()` 调度逻辑：

```
Phase 0  预检（仅 Step3）
Phase 1  直接运行 Python Pipeline 脚本
Phase 2  自检（self-check）
Phase 3  自动修复（仅 Step2）
Phase 4  严重度分级 → 决定是否放行
```

### Phase 1: Pipeline 直跑

用 `subprocess.run()` 依次执行对应脚本，**不启动 LLM**：

| Step | 脚本序列 |
|------|---------|
| 1 | `01_shape_detail.py` → `01b_auto_annotate.py` → `02_shape_analysis.py` |
| 2 | `02_shape_analysis.py` → `03a_build_shape.py --assemble-only` → `03a_build_shape.py --execute-prompts` |
| 3 | `03b_build_ppt_com.py --version X.X` |

- 脚本 crash（returncode != 0）→ 跳过自检，直接启动 LLM Agent 完整流程

### Phase 2: 自检

| Step | 自检方式 |
|------|---------|
| 1 | `self_check.check_step1()` — 纯 JSON 校验 |
| 2 | `self_check.check_step2()` — 纯 JSON 校验 + 填充内容检测 |
| 3 | 读取 `03b-self_check_report.md` — 解析 "结论：PASS/FAIL" |

**注意：Step3 的 pipeline 脚本（03b）内部已经包含一个独立的自检+自动修复循环（最多 3 次），完成后输出 report。Orchestrator 再读取这份 report 做二次判断。**

- 自检通过 → 步骤完成，继续下一步
- 自检失败 → 进入 Phase 3/4

### Phase 3: 自动修复（仅 Step2）

Step2 自检发现段落数/列表项不匹配时：
1. 向 prompt 注入结构约束（"输出必须包含恰好 N 个段落"）
2. 重新调用 GPT（`03a --execute-prompts`）
3. 再次自检 → 通过则完成

### Phase 4: 严重度分级

将所有问题分为两级：

| 级别 | 判定条件 | 处理方式 |
|------|---------|---------|
| **严重** | 缺失/为空/未知策略/report 中 "严重" 标记 | 阻断流程，启动 LLM Agent 修复 |
| **轻微** | 段落数偏差/填充内容等 | 警告输出，不阻断 |

Step3 严重问题进一步分类：
- **内容级**（超长/SSIM 差异/关键词缺失）→ 保存反馈文件，**回退 Step2** 重新生成
- **格式级**（字体/颜色/排版）→ 启动 LLM Agent 修复

---

## 各模式详细流程

### 模式 0：全自动

```
Step1 ──→ Step2 ──→ Step3
                      │
                      ├─ PASS → 打开 PPT，完成
                      │
                      └─ FAIL(内容问题) → 保存反馈
                           │
                           ↓
                       回退 Step2（带字数硬约束）
                           │
                           ↓
                       重跑 Step3（仅循环 1 次）
                           │
                           ├─ PASS → 打开 PPT，完成
                           └─ FAIL → 终止
```

任一步骤失败且无法回退 → 工作流终止。

### 模式 1：步骤1（分析 PPT 模板）

```
Pipeline: 01_shape_detail → 01b_auto_annotate → 02_shape_analysis
    ↓
自检 → PASS → 打开 Excel（用户编辑黄色单元格）
    ↓ FAIL
严重度分级 → 轻微放行 / 严重启动 LLM 修复
```

产物：`01-shape_detail.xlsx` + `01-shape_detail_com.json` + `02-shape_analysis_map.json`

### 模式 2：步骤2（构建 prompt + 调 GPT）

```
Pipeline: 02_shape_analysis → 03a --assemble-only → (继承上轮约束) → 03a --execute-prompts
    ↓
自检 → PASS → 打开 Excel（用户可编辑 prompt）
    ↓ FAIL
自动修复（注入段落/列表约束 → 重跑 GPT）
    ↓ 仍失败
严重度分级 → 轻微放行 / 严重启动 LLM 修复
```

产物：`02-prompt_specs.json` + `03a-build_shape_content.json`

### 模式 3：步骤3（构建 PPT）

```
Phase 0 预检:
  ├─ 对比 Excel prompt vs JSON → 有变化则自动补跑 GPT
  └─ 显示 Step2 遗留问题

Pipeline: 03b_build_ppt_com（内含自检循环 ×3 + 自动修复）
    ↓
Orchestrator 读取 report → PASS → 打开 PPT，完成
    ↓ FAIL
严重度分级:
  ├─ 内容问题 → 保存反馈 → 回退 Step2 → 重跑 Step3（循环 1 次）
  ├─ 格式问题 → 启动 LLM Agent 修复
  └─ 轻微问题 → 警告放行
```

产物：`output/YYYY-MM-DD claude-ppt X.X.pptx` + `03b-self_check_report.md`

---

## 03b 内部自检循环（Step3 Pipeline 内置）

03b 脚本在生成 PPT 后，**自身**执行最多 3 轮（1 次初始 + 2 次修复）自检：

```
[写入 PPT] → 检查(属性 + 视觉SSIM + 内容 + 字体) → 有严重问题？
                                                        ├─ 否 → 输出 PASS report
                                                        └─ 是 → 自动修复(字体/染色/文本) → 保存 → 再检查
                                                              (最多修复 2 次)
```

检查维度：
1. **属性检查** — shape 位置/大小/可编辑性
2. **视觉对比** — 模板 vs 生成页 SSIM（剪贴板截图，绕过加密）
3. **内容检查** — 文本是否为空/截断/超长、段落数、字体名称
4. **关键词染色** — 优势段落红色、劣势段落蓝色（写入时自动完成）

---

## 反馈与回退机制

| 机制 | 触发条件 | 行为 |
|------|---------|------|
| **Step3 → Step2 回退** | 内容超长 | 保存 `03-feedback_to_step2.json`，Step2 消费后三管齐下注入约束（见下），重跑 GPT |
| **Excel prompt 同步** | 用户在 Excel 编辑了 prompt | Step3 启动前自动检测差异，补跑 `03a --execute-prompts` |
| **结构约束继承** | Step2 上轮自检发现段落/列表不匹配 | `--assemble-only` 后自动注入 "必须包含 N 个段落" 约束 |
| **字数硬截断** | GPT 输出超过 max_chars | `clamp_text()` 在句子边界截断，防止溢出穿透到 PPT |

### Step3 反馈注入（三管齐下）

当 Step3 检测到"内容超长"时，反馈消费阶段同时修改三个文件：

| # | 文件 | 修改内容 | 作用 |
|---|------|---------|------|
| 1 | `02-readability_budget.json` | 降低 `max_chars` | `clamp_text()` 硬截断 + GPT prompt 中的 `target_chars` |
| 2 | `02-shape_analysis_map.json` | `user_instruction` 追加字数约束 | GPT prompt 中的"用户指令"段 |
| 3 | `03a-pending_prompts.json` | prompt 末尾追加【硬约束】 | override prompt 也包含约束 |

### SSIM 分类说明

SSIM（模板 vs 生成页视觉相似度）**不触发 Step2 回退**。原因：模板页有占位文本，生成页有实际内容，SSIM 必然偏低，Step2 无法改善。SSIM 严重问题归为"格式级"，交由 LLM Agent 处理或人工检查。
