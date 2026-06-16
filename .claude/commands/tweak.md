微调现有模板的闭环编排（plan-2026-05-28 §6 Step 4）。

> **本命令 ≠ /developer**：`/developer` = 单次代码改动；`/tweak` = "明确目标 → 必要时刷契约基线 → /developer 改 → 主 Claude 跑 acceptance" 的完整微调闭环。**只跑一次，不自动重试**——自动优化闭环属 Step 5（未实装），任何形式的"循环改到通过"都触发 plan §5 红旗。

## 何时用 /tweak（命中任一）

- 已知模板加新 shape / 改新字段 / 改字号字体颜色
- prompt 文案调优后想验证生产 PPT 没回退
- shape 微调（位置/尺寸/AutoSize）想验证不破坏既有验收
- 用户已自己改了 `template/*.pptx`，需要刷新契约基线再改代码

## 反模式（不要走 /tweak）

- ❌ 完全新模板首跑 → 走 `orchestrator.py` 全流程（CLAUDE.md §1 决策点表"完全新模板"）
- ❌ chart 路线问题 / 多轮 pivot 类 → 走主 Claude 兜底（fix4 类）
- ❌ 单纯改 prompt 文案、不在乎是否回归 → 直接 `/developer`，不需要 acceptance

## 流程（6 步，主 Claude 编排）

### 步骤 1：解析需求

- 微调对象（哪个模板 + 哪个生产代码？`src/{name}_ppt.py` / pipeline / 都改？）
- 微调类型分类：
  - **A. 改代码、不动模板** → 跳过步骤 2，直接复用既有 `acceptance/{name}.json`
  - **B. 改了模板（.pptx 排版/字号/颜色）** → 必须先跑步骤 2 刷新契约基线
  - **C. 加新 shape / 改 shape 期望格式** → 走 B（模板视为已变）
- 不确定 A/B → 问用户一句"本次有动 template/*.pptx 吗？"

### 步骤 2（条件触发）：刷新契约基线

只在步骤 1 判定 B/C 时跑：

```powershell
# 跑 inspect 抓最新模板的 paragraph-aware run 矩阵
python "C:\Users\<USER>\.claude\skills\inspect-ppt-template\inspect_ppt_template.py" `
    --file "template/<your>.pptx" --slides "<N>" --full `
    --out "pipeline-progress/_inspect_probe/"
```

**或**复用 Step1 已经实现的"烤草稿契约"逻辑（推荐——它一次性产 `01-acceptance_draft.json`）：

```powershell
$env:PPT_TEMPLATE_PATH="template/<your>.pptx"
python pipeline/01_shape_detail.py --force
```

产物：`pipeline-progress/01-acceptance_draft.json`（草稿契约）。

**护栏（plan §5 护栏 1）**：草稿契约只是基线参考。**期望值真相只能来自外部**——若用户的微调目标态超出模板（如要把字号 16→18），Step1 烤的是旧值，必须人工把目标态写进契约才能用，**禁止把"模板默认值"当成期望值跑验收**（红旗 4 重演）。

### 步骤 3：生成微调 plan md

- 位置：贴着任务归属——
  - 若属现有 feature：`[feature{XX}-...]/tweak-{YYYY-MM-DD}-{topic}.md`
  - 否则：`pipeline-progress/tweak-{YYYY-MM-DD}-{topic}.md`
- 必含字段：
  - 目标（一句话）
  - 范围（动哪个文件、动哪几个 shape）
  - 不动什么（明确边界，防 developer 顺手改飞）
  - 验收维度（哪几条 acceptance 规则会变 / 哪几条必须保绿）
  - 完成定义（DOD）：`must_fix=0` 且未引入新的 warn

### 步骤 4：调 /developer 跑代码改动

```
/developer 见 <plan md 路径>，按 plan 的范围/边界改
```

- developer 4 禁照旧（CLAUDE.md §3 apparel-fix4 那条 + `feedback_acceptance_gate.md`）：①不跑 acceptance ②不改 contract ③不自创 trace event 名 ④不 hardcode 期望值"回读自证"
- developer 回报后，主 Claude **必 `git diff` 看新增"验证"逻辑里有没有 hardcode 常量**

### 步骤 5：主 Claude 跑 ppt-acceptance-check

**前置**（沿用 plan-acceptance-gate-split-2026-05-27 Step A 约定）：
- PPT 开着（或用 `--active-new`）
- trace 已落 `acceptance/{name}_trace.jsonl`（pipeline 跑由 `PPT_PIPELINE_TRACE=1` 触发，src 跑由 `_TRACE` 模块级 logger 触发）
- 契约 `acceptance/{name}.json` 存在

```powershell
python "C:\Users\<USER>\.claude\skills\ppt-acceptance-check\ppt_acceptance_check.py" `
    --new "<生产 PPT 路径>" `
    --template "template/<your>.pptx" `
    --slide-pairs "<src:dst>" `
    --contract "acceptance/<name>.json" `
    --pipeline-trace "acceptance/<name>_trace.jsonl" `
    --out-dir "acceptance/<name>-out/"
```

判读：
- `must_fix=0` → ✅ PASS，可交付
- `must_fix>0` → 列根因清单 + 给用户两条候选（继续 `/tweak` 再跑一轮 / 升级到主 Claude 兜底），**不自动重跑**

### 步骤 6：回报

只报三件事：
- 完成了什么（一句话）
- acceptance 结论（PASS / FAIL + must_fix/warn 计数）
- 下一步建议（PASS → 是否 mc-update / 是否 commit；FAIL → 根因 + 路线候选）

## 硬约束（plan §5 三护栏）

| # | 护栏 | 如果违反 |
|---|---|---|
| 1 | 契约期望值只能来自外部真相（excel / inspect 目标态 / 用户人工） | 退化成红旗 4 自动化版 |
| 2 | **/tweak 单次跑，不自动重试**（自动闭环属 Step 5，未实装） | 触发 CLAUDE.md §0「连续失败 2 次熔断」 |
| 3 | 验收编排权留主 Claude（developer 不跑、不改契约） | 退化回 acceptance-gate-split 之前的旧坑 |

## 与其他命令的关系

| 场景 | 命令 |
|---|---|
| 完全新模板首跑 | `orchestrator.py` 全流程（Pipeline 系统） |
| 已知模板单次改代码、不在乎回归 | `/developer <task>` |
| **已知模板微调 + 想要回归保证** | **`/tweak <需求>`** ← 本命令 |
| 修复 pipeline 沉默 bug / 多轮 pivot | 主 Claude 兜底（不走 slash command） |
| 任务收尾固化文档 | `/mc-update` |

## 反例

- ❌ 走 /tweak 但不跑 acceptance（退化成 /developer，丢回归保证）
- ❌ 走 /tweak 跑了 acceptance 但 must_fix>0 还自动再调 /developer（触发护栏 2）
- ❌ developer 在 /tweak 流程里自跑 acceptance（违反 plan-2026-05-27 责任拆分）
- ❌ 把模板默认值当成期望值塞进契约跑验收（红旗 4 自动化版）
- ❌ 跳过步骤 3 plan md 直接喊 /developer（developer 没明确边界容易改飞）
