# Plan：工作流简化 —— 验收能力贯通 Pipeline，收敛到一份"共享真相产物"

**日期**：2026-05-28
**触发事件**：用户复盘三重混合工作流时提出三个命题——①验收 skill 越来越强，但只长在 src/developer 侧；②"解析新模板"阶段粗糙，形成「粗略解析 + 严苛验收」的不对称；③Pipeline 自检还停在旧标准。要求系统回顾 repo 后评估"工作流能否进一步简化"。
**性质**：前瞻型简化方案。**本 plan 只记录路线，不含已完成改动**。经一轮 Explore 实锤 + 主 Claude 综合，等用户点头后再开工。
**2026-05-28 二次细化**：用户复盘时强调对排版要求高、模板常见"同一 shape 多字体/颜色/字号 + 每行不同格式"，要求解析能识别、验收能按行设门槛。主 Claude 实锤 apparel 模板后发现 §3 "同一个数据结构" premise 有缺口（inspect 与 acceptance 用两套不同 walker、颗粒度不兼容），故把主杠杆从"flat `expected_runs`"升级为"paragraph-aware + 段内合并"模型（详见新增 §3.5）。用户已拍板两条边界 + 五条护栏，批准开工。
**关联**：建立在 [plan-acceptance-gate-split-2026-05-27.md](plan-acceptance-gate-split-2026-05-27.md)（验收权从 developer 收回主 Claude）之上；红旗复盘见 `.claude/memory/feedback_acceptance_gate.md`。

---

## 0. 背景一句话

强大的验收能力（`ppt-acceptance-check` L0-L5 + `runs_match_signature`、`inspect-ppt-template --full`）在 2026-05 全部长在 **src/developer 这一侧**；Pipeline 体系在更早就冻结了，至今用另一套旧标准（LLM 主观评分 + 手写 shape 级遍历）。两侧用的是**两套互不相通的"真相"**，中间没有桥。简化的本质 = **合并重复工具（3 探针→1、2 验收→1、N 份期望态→1 份契约）**，不是减少工作流步骤。

---

## 1. 评估输入（Explore 实锤的 repo ground truth）

> 一轮 Explore subagent 通读 orchestrator.py + pipeline/*.py + Step1/2/3 agents 后的事实核实。行号为盘点时快照，开工前以实际为准。

### 1.1 用户三个命题全部被实锤，且比描述更彻底

| 用户命题 | repo 事实 | 证据 |
|---|---|---|
| "粗略地解析新模板" | Step1 手写 COM 遍历，**只到 shape 级**（name/L/T/W/H/text 纯字符串），**完全没 paragraph/run**，更没调 `inspect-ppt-template`。新升级的 `--full` run 矩阵能力一点没吃到 | `pipeline/01_shape_detail.py` L74-93（`shape_obj()` 仅 shape 级字段）、L110-206（手写双页遍历）、L26-40（import 全是项目内 com_get，无 inspect skill） |
| "严苛地验收" | `ppt-acceptance-check` L0-L5 + `runs_match_signature` 仅在 src/developer 侧用 | `acceptance/apparel.json` + `ppt-acceptance-check/SKILL.md` |
| "pipeline 自检还是旧标准" | orchestrator 自检 = `04-fix_ppt.md`（LLM 主观 visual/readability/semantic 评分）+ `pipeline/self_check.py`（paragraph/bullet 计数）。**全文搜不到一处调用 `ppt-acceptance-check`** | orchestrator.py L468、L776-817、L1018-1022；`generate_self_check_report` 在 ppt_pipeline_common.py L972 |
| （隐藏第四处） | Step2/3 **没接 TraceLogger**，pipeline 产出的 PPT **天生无法跑 L4 行为层**（chart silent failure 这类在 pipeline 侧根本测不出） | `pipeline/03b_build_ppt_com.py` 无 trace 落盘、无 COM 调用成败事件记录 |
| （隐藏第五处） | `pipeline-self-check-loop` skill **只有方法论文档，零代码落地**；orchestrator 的自检内循环是项目自有逻辑，不复用 skill 引擎 | `grep "pipeline-self-check-loop\|structural_check\|ppt_visual_check"` 全 repo 零匹配 |

### 1.2 核心根因

现状不是"两端能力不对称"那么简单，而是**两端用的是两套互不相通的真相产物**：

- Pipeline 端产出：`04-fix_ppt.md`（主观评分 md）+ `03b-self_check_report.md`（计数表）
- developer/acceptance 端要的：结构化断言契约 `acceptance/{name}.json`（含 `expected_runs` / `expected_from: excel:`）
- **中间没有桥** → "解析新模板"和"验收"各写各的期望态，且解析这一侧根本没有 run 级粒度可写

---

## 2. 核心命题：能简化，但简化 = 收敛，不是砍步骤

工作流的"步数"压不下去——build + verify 是必须的。能压的是**并行重复的机制数**：

```
现状（发散）：                          简化后（收敛）：
3 个"看 shape 长啥样"的探针             1 个探针：inspect-ppt-template --full
  - Step1 手写 COM（shape 级）            （read-selected-shape 作交互式变体）
  - inspect-ppt-template（--full）              │ 产出
  - read-selected-shape（--full）               ▼
2 套验收                               1 份真相产物：contract（含 expected_runs）
  - pipeline self_check.py（计数）              │ 消费
  - ppt-acceptance-check（L0-L5）               ▼
N 份"期望态"各写各的                    1 个验收：ppt-acceptance-check
```

净效果：探针 3→1、验收 2→1、期望态描述 N→1 份契约。Pipeline 和 src 两侧**第一次共享同一份真相**。

---

## 3. 简化主杠杆：一份"共享真相产物"

**关键洞察**：`inspect-ppt-template --full` 吐的 `paragraphs[].runs[]{text,size,rgb,bold,...}`，和 `ppt-acceptance-check` 的 `runs_match_signature` 吃的 `expected_runs`，**是同一个数据结构**——一个生产、一个消费。现在它们没接上，所以上一轮评估里"手工 copy 桥"才存在（见 [feedback_acceptance_gate.md] 红旗 5 续）。

**把它接上，整条线就塌缩**，一步治三个病：

1. **不对称消失**：Step1 解析改用 `inspect --full`，"解析"和"验收"用同一份 run 矩阵，解析自动变得和验收一样严苛
2. **手工桥消失**：Step1 直接把 `expected_runs` 落进契约，无需人肉从 inspect 输出 copy 到 contract
3. **Pipeline 产物升级**：从"主观评分 md"升级成"可直接喂验收的契约"，developer 移植完直接复用，不用重新 inspect

---

## 3.5 二次细化：共享真相产物必须升到 paragraph-aware（2026-05-28 会话）

**触发**：用户复盘时强调——他对排版要求高，提供的模板**常见**复杂格式（同一 shape 内多字体/多颜色/多字号；某 shape 分多行、每行不同字号颜色）。要求：①解析能识别这些细节；②验收能按"实际情况"设合理门槛（如"某 shape 2 行，每行不同字号颜色"）。这是把 §3 的主杠杆从"flat run 签名"顶到"按行多格式"的目标值钉死。

### 3.5.1 实锤：§3 的"同一个数据结构"premise 有缺口

§3 原断言"`inspect --full` 的 `paragraphs[].runs[]` 和 `runs_match_signature` 的 `expected_runs` 是同一个数据结构"——**字段兼容，颗粒度不兼容**：

| 维度 | inspect `extract_paragraphs` | acceptance `_walk_runs` |
|---|---|---|
| 结构 | paragraph 嵌套 | 扁平（整 shape 一条序列）|
| 合并 | **不合并**（忠实 paragraph.Runs）| 按 `(rgb,bold,italic,size)` **合并** |
| 空白 | 保留 `\r` 噪音 run | 可滤空白 run（ignore_whitespace_runs）|
| 分行信息 | **保留** | **丢失** |

直接把 inspect 的 run 倾倒进 `expected_runs` → run 数对不上 → `runs_match_signature` 判 FAIL。**两个"共享"工具其实用了两套不同 walker。**

实锤数据（`inspect --full` 跑 `template/apparel-page13-14-template.pptx` p13-14）：
- **TextBox 6**：行1 `品质` 20pt 黑 / 行2 `3.98 / 5` 16pt 红 ← 用户"每行不同字号颜色"的真实样本
- **TextBox 50**：行2 `15~25℃` 被 PPT **拆成 `15~25`+`℃` 两段**（样式完全相同）← 证明"必须合并"才不误报
- **RR 53**：行1 `累计跑量km` 11pt 白（拆 2 段）/ 行2 `671` 24pt 白

### 3.5.2 方向决策：paragraph-aware + 段内合并 = 唯一共享真相

取 inspect 的"分行结构" + `_walk_runs` 的"段内合并"，合成一个模型：

- **canonical extractor `extract_paragraph_runs(shape)`** → `[{para_idx, alignment, text, runs:[{rgb,bold,size,...}]}]`，**每段内**按属性合并相邻同样式 run、丢 `\r`/`\n`-only run。
- **Step1（感知）**：用它替换现有 shape 级 COM 遍历 → 解析能看见"按行多字号/多颜色"，并把 `expected_paragraphs` 烤进**草稿** `acceptance/{name}.json`。
- **acceptance（新检查）**：加 `paragraphs_match_signature`，消费 `expected_paragraphs`，做**按行** run 签名断言（= 用户要的"每行设门槛"）。保留现有 flat `runs_match_signature` / `runs_match_template` 向后兼容（apparel 保绿）。

效果：inspect 产出 = acceptance 输入，§3 premise **由"把 inspect 降维"改成"把 acceptance 升维到 inspect 的丰富度"** 而成立。

**两条边界（用户已拍板）**：
1. **动 skill 但只加不改**：给 acceptance skill 加 `paragraphs_match_signature` + 共享 walker，全部加法、向后兼容；现有 flat 检查与 `acceptance/apparel.json` 不动。
2. **契约 skill 为权威 walker，Step1 import**：canonical walker 定义在 acceptance skill 侧（验收读 live PPT，walker 必须在 skill 侧），Step1 跨边界 import（sys.path 挂 skill 目录）。保证"一个 walker 一份真相"——**不 vendor 镜像**（镜像会和 skill 那份漂移）。

### 3.5.3 五条护栏（写 developer 任务时作为约束）

1. **"同一套准则" = 同一 walker + 同一组维度**。维度清单（rgb/bold/size/italic/font_name 选哪几个合并 + 比对）必须两侧同源、写死一处；否则 walker 按 size 合并、验收却查 font_name，又对不上。
2. **Step1 烤的契约只能是草稿，期望值真相来源守死**（红旗 4/5 护栏）。模板=目标态 → 模板提取合法；目标态**超出**模板（如 RR53/55 升级）→ Step1 烤的是旧值，必须人工用外部真相覆盖，禁止自动化把错误期望固化成门禁。
3. **`\r`（真段落）vs `\n`（段内软换行）的"行"定义先约定**。apparel TextBox 26 里两者并存，直接影响"每行"怎么数、`paragraphs_match_signature` 按什么切行。
4. **严格度分级**。固定标签（评分/数值/温度区间）→ `must_fix` 严格按行；GPT 自由文本（优缺点 bullet）run 随【】关键词浮动 → `warn` 或只断言"标题行 + 正文 pattern"，禁刚性 run 数断言（否则天天误报）。apparel.json 现有分级是参考范式。
5. **先拿 apparel 做回归基线再上新模板**。新 walker + 新检查写完，先重提 apparel 的 `expected_paragraphs`，确认复现已知正确格式 + 对现有 apparel.pptx 判 PASS，再跑全新模板。

---

## 4. 两条目标工作流（用户提案 + 细化）

### 4.1 流程 1：全新模板

> Pipeline 冷启动 → developer 移植 → acceptance 验收 → 自动优化 → 交付

**补丁（缺的地基）**：**"Pipeline 冷启动"那一步必须额外产出 acceptance 契约**——Step1 调 `inspect --full` → 把 `expected_runs` 落进 `acceptance/{name}.json`。否则后续"跑 acceptance"无契约可跑。这样冷启动的价值从"给 prompt 语料 + 视觉基线"升级为"**额外交付一份契约**"，移植 → 验收自动闭环。

### 4.2 流程 2：微调现有模板 —— 要不要新 skill？

> 读需求 → 生成任务 plan md → developer 改代码 → acceptance 验收 → 自动优化 → 交付

**结论：不需要从零造重型 skill，需要一根薄胶水。**

零件都在：developer（改代码）、ppt-acceptance-check（验收）、inspect-ppt-template（刷新基线）。缺的只是一个 thin orchestration command，把"读需求 → inspect 当前模板刷新契约基线 → 生成 plan md → developer 改 → 主 Claude 跑 acceptance"串起来。建议做成 `/tweak` slash command 而非重型 skill。

**依赖顺序硬约束**：必须先有"共享真相产物"地基（§3），否则 `/tweak` 里的"跑 acceptance"同样没有严苛契约可比。

---

## 5. ⚠️ 关键警告：「自动优化」是自动化版的红旗 4

用户两条流程都有"acceptance 验收 → **自动优化** → 交付"。这一步最危险：

> 让生成器在闭环里反复改自己直到通过验收，**如果验收标准是生成器侧能改的，它必然收敛到"作弊通过"**——这正是 2026-05-27 封禁的 hardcode 回读自证（红旗 4），只是变成了自动化版本。

**护栏三条，缺一不可**：

| # | 护栏 | 依据 |
|---|---|---|
| 1 | 契约期望值**只能**来自外部真相（Excel 真实数据 `expected_from: excel:` / inspect 目标态），生成器/developer **无权改契约** | 延续 developer.md「4 禁」第②③④条 |
| 2 | 自动优化**硬上限 2 轮**，不过就 escalate 主 Claude，禁止无限刷 | 对齐 CLAUDE.md §0「连续失败 2 次熔断」 |
| 3 | 验收编排权留在主 Claude（审查者 ≠ 被审查者），自动优化只能改"被审查物"，不能碰"审查标准" | 延续 [plan-acceptance-gate-split-2026-05-27.md] Step A |

---

## 6. 落地排序（有依赖，不能并行）

| 序 | 动作 | 改动位置 | 性质 | 为什么是这个顺序 |
|---|---|---|---|---|
| 1a | acceptance skill 加 canonical `extract_paragraph_runs` walker（段内合并）+ `paragraphs_match_signature` 检查（只加不改，§3.5.2 边界1）| acceptance skill（用户级）| 加 skill 能力 | walker 是权威，Step1 要 import，故先落 skill 侧；1a 必先于 1b |
| 1b | Step1 import 权威 walker，替换 shape 级 COM 遍历做 paragraph-aware 感知，把 `expected_paragraphs` 烤进**草稿** `acceptance/{name}.json`（受 §3.5.3 五护栏约束）| `pipeline/01_shape_detail.py` | 改 pipeline 代码 | **地基**。不做这步，后面全是空中楼阁 |
| 2 | Step3 接 TraceLogger（事件落 `acceptance/{name}_trace.jsonl`，复用 src 侧标准 event 名） | `pipeline/03b_build_ppt_com.py` | 改 pipeline 代码 | 让 pipeline 产物能跑 L4 行为层 |
| 3 | orchestrator 末步用 `ppt-acceptance-check` 替/补 `04-fix_ppt.md` 作为 Step3 后置门禁 | `orchestrator.py` | 改 orchestrator | 1+2 就绪后才有东西可验 |
| 4 | 做 `/tweak` 薄命令（流程 2 微调路径） | 新 slash command | 新增编排 | 依赖 1-3 的契约地基 |
| 5 | 自动优化闭环 + §5 三条护栏 | 编排逻辑 | 编排 + 护栏 | 最后做，风险最高，要前面都稳了 |

---

## 7. 已知风险 / 注意事项

| 风险 | 应对 |
|---|---|
| Step1 改 paragraph-aware 后，旧模板的 shape_detail_com.json schema 变了，下游 Step2/3 可能读不到老字段 | 加字段而非改字段；`paragraphs`/`expected_paragraphs` 作为新增键，老键（text/font_name/...）保留兼容 |
| `inspect --active` 依赖 PPT 开着；pipeline 全自动跑时可能没开 | Step1 已自己 `Dispatch` 打开 PPT，改成复用该实例传给 inspect，或 inspect 支持 `--file` 路径模式 |
| 自动优化闭环若护栏没上就先做（贪图省事） | §6 排序硬约束：第 5 步必须最后；护栏（§5）与闭环同批上线，禁分离 |
| Pipeline 改造把 src 侧已稳的范式带歪 | src/apparel_ppt.py 的 trace 接法是参考范式，Step2/3 接 TraceLogger 时**对齐它的 event 名**，不另起一套 |
| `pipeline-self-check-loop` skill 始终没落地 | 不急着实装（orchestrator 内循环够用）；但 §3 之后 orchestrator 末步直接调 ppt-acceptance-check，等于事实上让 acceptance 成为统一验收入口，self-check-loop 可标注"由 acceptance-check 承担" |

---

## 8. 不做的事 / 边界

- **不动 src/ 已稳代码**：apparel_ppt.py / Main.py 现状合规，仅作 Step2/3 接 trace 的范式参考
- **不重写 orchestrator 内循环**：只在末步挂 acceptance 门禁，不推翻它自有的迭代逻辑
- **不造重型新 skill**：流程 2 用薄 `/tweak` 命令，不新建 agent
- **不在本 plan 阶段碰代码**：本文件仅为路线固化，待用户批准

---

## 9. 决策记录

| 日期 | 决策 | 决策者 | 理由 |
|---|---|---|---|
| 2026-05-28 | 确认"解析粗/验收严"不对称有代码根因（非主观感受） | Explore 实锤 + 主 Claude | Step1 仅 shape 级、零 inspect 复用、pipeline 自检零 acceptance 调用 |
| 2026-05-28 | 简化主杠杆定为"共享真相产物（expected_runs 贯通 inspect→contract→acceptance）" | 主 Claude | 一步治三病；inspect 产出与 acceptance 消费是同一数据结构 |
| 2026-05-28 | 流程 2 用薄 `/tweak` 命令，不造重型 skill | 主 Claude | 零件齐全，只缺胶水；避免过早抽象 |
| 2026-05-28 | 「自动优化」必须配 §5 三条护栏，否则等于自动化红旗 4 | 主 Claude | 闭环自优化 = 被审查者改审查标准的极端形态 |
| 2026-05-28 | 落地分 5 步且有硬依赖顺序，Step1 是地基 | 主 Claude | 契约地基不先建，后面验收/微调/自优化全空转 |
| 2026-05-28 | 共享真相产物升到 paragraph-aware + 段内合并模型（非 §3 原设 flat `expected_runs`）| 用户 + 主 Claude | 实锤 inspect 与 `_walk_runs` 颗粒度不兼容；用户模板常见"每行不同字号颜色"需按行断言 |
| 2026-05-28 | 动 acceptance/inspect 用户级 skill，但只加不改（向后兼容）| 用户 | 加 `paragraphs_match_signature` 是唯一能表达"每行设门槛"的路 |
| 2026-05-28 | canonical walker 归 acceptance skill（权威），Step1 import，不 vendor 镜像 | 用户 | 验收读 live PPT，walker 必在 skill 侧；import 保证"一个 walker 一份真相"，镜像会漂移 |
| 2026-05-28 | 落地第 1 步拆 1a（skill 加 walker+检查）→ 1b（Step1 import 烤草稿契约），1a 先于 1b | 主 Claude | Step1 import 依赖 walker 已在 skill 侧存在 |
| 2026-05-28 | 批准开工：先把本次二次细化补进 plan，再从 Step 1a 起 | 用户 | 方向 + 边界 + 五护栏均已拍板 |

---

## 10. 几周后（或开工后）回看时该问自己的问题

1. 第 1 步（Step1 接 inspect --full）落地后，新模板 `acceptance/{name}.json` 是否真的自动带上了 `expected_runs`？数据：__________
2. 手工 copy 桥是否真的消失了（不再需要人肉从 inspect 输出抄到 contract）？数据：__________
3. Pipeline 末步跑 acceptance-check 后，是否抓到过 04-fix_ppt.md 主观评分漏掉的结构性 bug？案例：__________
4. `/tweak` 命令是否真的把"微调现有模板"的链路缩短了？对比改造前工作量：__________
5. 自动优化闭环上线后，§5 三条护栏是否拦住过"作弊收敛"？案例：__________

→ 把上面 5 个空填了，本 plan 寿命终结，写入 [STATE.md](../STATE.md) §1 变更日志归档。

---

## 11. 落地状态（2026-05-28 单会话完成 Step 1-3，同日 Step 4 加挂）

> 本次单执行会话扛完 §6 落地排序的 Step 1-3（pipeline 管线改造段），主 Claude 编排 + 2 名 cold-start developer 协作，主 Claude 跑端到端验收。Step 4（`/tweak` 薄编排）同日加挂完成。**剩 Step 5（自动优化闭环 + §5 三护栏）—— 风险最高、放最后。下次开工从 §11.5 第 2 项接力。**

### 11.1 已完成清单（§6 表内项）

| 序 | 状态 | 落地位置 |
|---|---|---|
| 1a | ✅ **DONE** | `C:/Users/xy24/.claude/skills/ppt-acceptance-check/paragraph_runs.py`（新建，权威 walker `extract_paragraph_runs` + `MERGE_DIMS=(rgb,bold,size)`）；`layers/runs.py` 加 `paragraphs_match_signature` L3 检查（纯加分支，现有 4 个 elif 未动）|
| 1b | ✅ **DONE** | `pipeline/01_shape_detail.py`：`Path.home()` 挂 sys.path import 权威 walker；`shape_obj` 加 `paragraphs` 键（diff key 元组不动）；`main()` 烤 `pipeline-progress/01-acceptance_draft.json`（默认 warn 严重度）|
| 2 | ✅ **DONE** | `pipeline/03b_build_ppt_com.py`：接 TraceLogger（镜像 `src/apparel_ppt.py` 第 90-135 行）；3 处 `except` 发 `com_api_failed_but_continued`（`_write_text`/`_write_chart`/`_replace_image`）；`apply_shape` 路由块包 `_trace_shape` 上下文；`PPT_PIPELINE_TRACE` 环境变量触发 |
| 3 | ✅ **DONE** | `orchestrator.py`：新增 `_wrap_draft_contract`（静态方法）+ `_run_acceptance_gate`（never-raises）；`_run_pipeline_scripts` step==3 分支注入 trace env + 清旧 trace + 跑完调 acceptance；MVP 信息性报 `[GATE] PASS/FAIL`，不阻断 orchestrator |
| 4 | ✅ **DONE** | `.claude/commands/tweak.md`（6 步薄编排：解析需求 → 必要时刷契约基线 → 生成 plan md → /developer 改 → 主 Claude 跑 acceptance → 回报）。硬约束沿用 §5 三护栏 + plan-2026-05-27 责任拆分；**单次跑、不自动重试**（自动闭环属 Step 5）。CLAUDE.md §1 决策点速查表"shape 微调"+"已知模板加新 shape"两行加 `/tweak` 路径；STATE.md §1 变更日志 +1 行。 |

### 11.2 顺手拆掉的 COM 安全雷（不在原 §6，但 2026-05-28 实战必修）

事故：`pipeline/01_shape_detail.py` 跑 `generate_shape_detail_xlsx` 时用 `Dispatch("Excel.Application")` + finally `excel.Quit()`，**attach 用户活 Excel + 关闭**，用户丢未保存内容。同类雷 PowerPoint 侧也存在。

修复（全部归 `feedback_com_constraints.md` 末尾节）：
- `pipeline/ppt_pipeline_common.py` 3 处 `Dispatch("Excel.Application")` → `DispatchEx`（`load_excel_rows` / `generate_shape_detail_xlsx` / `create_iteration_sheet`）
- `pipeline/01_shape_detail.py` PowerPoint `Dispatch` → `DispatchEx` + `Open(ReadOnly=True, WithWindow=False)`（镜像 inspect skill 安全开法）
- `pipeline/03b_build_ppt_com.py` PowerPoint `Dispatch` → `DispatchEx`（保 `Visible=True` 不破自检视觉对比）

**判据**（写新 win32com 脚本前必查）：批量分析/生成类 → `DispatchEx` 隔离；驱动用户活 Office 类（`Main.py`/`src/*_ppt.py` 生产流程）→ 才用 `Dispatch`。详见 `feedback_com_constraints.md`。

### 11.3 健壮性顺手补

- `pipeline/01_shape_detail.py`：草稿契约挪到 `generate_shape_detail_xlsx` **之前**；xlsx 包 try/except 不致命——Excel 抽风不阻断关键产物（JSON + 草稿契约）

### 11.4 端到端验证 PASS（下次跑回归基线参考）

**self-compare 闭环测试**：模板（`template/empty and standard-apparel.pptx`）副本作为"产出 PPT"，slide-pairs `2:2`，喂自动 wrap 的契约：

```
L0 配对  : passed=22 must_fix=0
L1 数据  : passed=0 must_fix=0
L2 格式  : passed=38 must_fix=0
L3 染色  : passed=10 must_fix=0  ← paragraphs_match_signature 全过（共享真相产物链闭合证据）
L4 行为  : degraded（未喂 trace，预期）
L5 视觉  : passed=1 must_fix=0
结论    : PASS（must_fix 0 / warn 0 / tolerate 0）
报告    : pipeline-progress/acceptance-out/acceptance_report.md
```

**复现命令**（PowerShell，假设两个 fix 都已落）：
```powershell
# 1. Step1 跑富格式模板
$env:PPT_TEMPLATE_PATH="template/empty and standard-apparel.pptx"
python pipeline/01_shape_detail.py --force

# 2. wrap 草稿成完整契约
python -c "import json; d=json.load(open('pipeline-progress/01-acceptance_draft.json',encoding='utf-8')); c={'version':1,'shape_pairing':{'rules':['strip_clone_suffix'],'manual_overrides':{}},'tolerances':{'geometry_px':2.0,'ssim_threshold':0.85,'font_size_pt':1.0},'rules':d['rules']}; open('pipeline-progress/_acceptance_contract.auto.json','w',encoding='utf-8').write(json.dumps(c,ensure_ascii=False,indent=2))"

# 3. self-compare 跑 acceptance（用 2:2 slide-pairs；真 pipeline 跑完 03b 后用 1:2）
$tpl="$PWD/template/empty and standard-apparel.pptx"
cp "$tpl" "$PWD/pipeline-output/claude-ppt 1.0.pptx"
python "$HOME/.claude/skills/ppt-acceptance-check/ppt_acceptance_check.py" `
    --new "$PWD/pipeline-output/claude-ppt 1.0.pptx" `
    --template "$tpl" `
    --slide-pairs 2:2 `
    --contract "$PWD/pipeline-progress/_acceptance_contract.auto.json" `
    --out-dir "$PWD/pipeline-progress/acceptance-out/"
```

### 11.5 留给下次的（按依赖顺序）

1. ~~**Step 4：`/tweak` 薄命令**~~ ✅ **DONE（2026-05-28 同日完成）**
   - 落地位置：`.claude/commands/tweak.md`
   - 6 步流程：解析需求 → 必要时刷契约基线（条件触发，B/C 类微调才跑）→ 生成 plan md → /developer 改 → 主 Claude 跑 ppt-acceptance-check → 回报
   - 硬约束：plan §5 三护栏 + plan-2026-05-27 责任拆分（developer 不跑 / 不改契约 / 不 hardcode 期望值；验收编排权留主 Claude）
   - **单次跑、不自动重试**——任何形式的"循环改到通过"都触发自动闭环红旗，归 Step 5

2. **Step 5：自动优化闭环 + §5 三护栏**（最危险，放最后）
   - 性质：闭环自动化，**必须配 §5 三护栏防"作弊收敛"**（详见 §5 警告）
     - 护栏 1：契约期望值只能来自外部真相（excel/inspect），生成器无权改
     - 护栏 2：硬上限 2 轮，不过 escalate 主 Claude
     - 护栏 3：验收编排权留主 Claude（审查者 ≠ 被审查者）
   - 实现位置：orchestrator.py 加 acceptance gate 闭环（当 must_fix>0 时自动回头跑 step2/step3，2 轮上限）
   - 依赖：§11.1 全部 + Step 4（建议先有 /tweak 的微调体验再上自动闭环）
   - **实装前必读**（2026-05-28 新沉淀）：`.claude/memory/feedback_acceptance_gate.md` 末节「自动闭环 = 自动化版红旗 4」—— 三护栏的失效模式、为什么人工 git diff 复查兜不住"结构性绕道"、`/tweak` 当前 0 轮硬上限设计意图。任何"自动化把 acceptance 跑通"的设计提案先按这一节逐条 challenge：契约期望值能否被闭环写？重试上限多少？审查标准是不是被审查物的一部分？三条全过才允许动 orchestrator.py。

---

### 11.5.1 下次开工 checklist（开 Step 5 / 继续微调时按序走）

| 序 | 动作 | 触发条件 |
|---|---|---|
| 1 | 跑几次真实 `/tweak` 取 §10 第 4 题数据（"`/tweak` 是否真的缩短了微调链路"）| 在正常微调任务遇到时顺手取证；不要为取证造任务 |
| 2 | 评估 Step 5 是否真有必要 | 取证结论 + 用户当前痛点（是否真的需要自动闭环 / 还是 `/tweak` 单次跑已够用）|
| 3 | 若评估通过，**先读** `feedback_acceptance_gate.md` 末节「自动闭环 = 自动化版红旗 4」三护栏 | 触发 Step 5 设计前 |
| 4 | 起新 plan `plan-2026-XX-XX-自动闭环.md`，用三护栏逐条对照本项目场景做适配（特别是"闭环代码访问 contract / trace 白名单 / walker 维度"的物理隔离怎么做）| Step 5 真要动手 |
| 5 | 落地分两阶段：阶段 A 只跑闭环 dry-run（mock must_fix>0，看会不会去改契约/trace）；阶段 B 才接真 step2/step3 重跑 | 阶段 A 不过禁阶段 B |
| 6 | 上线后立刻取 §10 第 5 题数据（"§5 三护栏是否拦住过作弊收敛"）| Step 5 上线后第 1-2 轮真实任务里取证 |

### 11.6 临时验证物清单（可保留 or 清理）

| 路径 | 用途 | 处置建议 |
|---|---|---|
| `pipeline-progress/_inspect_probe/_mc_verify_walker.py` | 主 Claude 独立验证 1a walker | 保留（回归基线）|
| `pipeline-progress/_inspect_probe/_mc_verify_trace.py` | 主 Claude 独立验证 2 trace 埋点（fake shape）| 保留（回归基线）|
| `pipeline-progress/_inspect_probe/probe_paragraph_runs.py` | developer 1a 自验证脚本 | 可删 |
| `pipeline-progress/_inspect_probe/inspect_report.{json,md}` | 早期 inspect 探针输出 | 可删 |
| `pipeline-progress/_acceptance_contract.auto.json` | wrap 后的完整契约 | 每次跑会重写，保留 |
| `pipeline-progress/01-acceptance_draft.json` | Step1 烤的草稿契约 | **保留**（Step1 每次跑会更新）|
| `pipeline-output/claude-ppt 1.0.pptx` | self-compare 用的模板副本 | 可删（真跑 03b 会重生成）|
| `pipeline-progress/acceptance-out/acceptance_report.{md,json}` | acceptance 报告 | 保留作回归基线 |

### 11.7 §10 的 5 个回看问题，本次能填的先填了

| # | 问题 | 当前数据 |
|---|---|---|
| 1 | Step1 改完后，`expected_paragraphs`（plan 升级版的 `expected_runs`）是否真的自动落进契约？| ✅ `01-acceptance_draft.json` 自动 10 条 `paragraphs_match_signature` 规则，包含多段 shape（TextBox 24 = 5 段 / TextBox 8 = 3 段） |
| 2 | 手工 copy 桥消失？| ✅ Step1 烤草稿后 0 处人肉 copy |
| 3 | Pipeline 末步跑 acceptance 是否抓到过 04-fix_ppt.md 漏掉的结构性 bug？| 待真 pipeline 跑 / 真分发任务取证 |
| 4 | `/tweak` 是否缩短了微调链路？| ✅ 命令已落地（`.claude/commands/tweak.md`），数据待真实微调任务取证 |
| 5 | 自动优化闭环 + §5 三护栏是否拦住过作弊收敛？| Step 5 未做 |
