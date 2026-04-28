# plan3（工作流5阶段定稿）.md — 工作流定稿（用户决策版）

> **状态**：✅ 用户已决策（2026-04-27），待实施
> **前置档案**：plan2（三重混合机制再评估）.md（三重混合机制再评估）、fix2（三重混合架构整改）.md / fix3（图表写入诊断）.md / fix4（图表路线切换）.md
> **写入时间**：2026-04-27（基于 trash-bin/pipeline-progress-yzr 实测档案修订）

---

## 0. 决策结论（用户原话）

1. **Pipeline 保留**——确实有价值（反复讨论多次确认）
2. **新模板首次接入：必跑一遍完整 Pipeline**（token 成本可接受）
3. **首跑完成后：视 PPT 效果与复杂度评估**——继续 Pipeline 迭代 OR 直接进入移植
4. **移植阶段：首选 `/developer` 调用 Agent**
5. **复杂问题：在主 Claude 对话窗口处理**

---

## 1. 实测档案验证（修正前面分析的错误）

之前的评估有一处错误："yzr 历史上没用过 Pipeline → Developer 衔接"。在 `【trash-bin】/pipeline-progress-yzr/` 中找到了**完整 Pipeline 跑批的产物**，证明 yzr 实际走过完整流程。

### 实测证据

| 时间 | 阶段 | 产物 |
|--|--|--|
| 2026-03-23 09:12~11:18 | Step 1 + LLM 增强 | `01-shape_detail.xlsx` + `.json` + `.analyst_enhanced.json` |
| 2026-03-23 11:27 | Step 2 | `02-prompt_specs.json` + `02-shape_analysis_map.json` |
| 2026-03-23 16:47 | Step 3a + 3b | `03a-build_shape_content.json`、`03b-post_write_readback.json` |
| 2026-03-23 16:36 | Step 4 自检 | `04-fix_ppt.md`（visual=100%, readability=100%, semantic=66.67%） |
| 2026-04-16 | fix2 整改 | Pipeline prompt 同步到 `src/yzr_ppt.py`（注释 `prompt_src/synced_at/synced_by` 为证） |

### 关键发现

- yzr 早期接入（3-23）确实走过完整 Pipeline + Developer 衔接，产物清单齐全
- `02-shape_analysis_map.json` 显示用户对 9 个 shape 做过 Excel 手工标注（驱动 Step 2 strategy 推断）
- fix2 阶段（4-16）由 Developer Agent 把 Pipeline 的 prompt 同步到 `_build_rich_prompt()`
- **但 fix3 / fix4 chart 修复阶段（4-23 ~ 4-27）完全没动 Pipeline**——这是后期 bug 修复，不是初次移植

**这印证了用户决策的合理性**：Pipeline 在"初次接入"时有真实价值；在"已知模板修 bug"时不参与。

---

## 2. 用户工作流程图（5 阶段）

```
新模板 / 数据源到手
       │
       ▼
┌─────────────────────────────┐
│  阶段 1：Pipeline 首跑      │  ← 必做（按用户决策）
│  - orchestrator.py 全流程   │
│  - 产出 pipeline-progress/  │
│    01~04 全套档案           │
└──────────────┬──────────────┘
               │
               ▼
┌─────────────────────────────┐
│  阶段 2：评估 PPT 效果      │  ← 用户决策点
│  - 看 03b 输出的 PPT 视觉   │
│  - 读 04-fix_ppt.md 自检    │
│  - 判断：是否值得继续迭代？ │
└──────────────┬──────────────┘
               │
       ┌───────┴───────┐
       │               │
       ▼               ▼
  视觉满意度≥80%   视觉满意度<80%
  且 shape 角色清晰  或 shape 复杂
       │               │
       ▼               ▼
┌─────────────┐  ┌──────────────────┐
│ 阶段 3a     │  │ 阶段 3b          │
│ 跳过迭代，  │  │ 继续 Pipeline    │
│ 进入移植    │  │ 迭代（修标注 →   │
│             │  │ 重跑 Step2/3）   │
└──────┬──────┘  └────────┬─────────┘
       │                  │
       └────────┬─────────┘
                │
                ▼
┌─────────────────────────────┐
│  阶段 4：/developer 移植    │  ← 默认路径
│  - 复制 yzr_ppt.py 骨架     │
│  - import _ppt_shared       │
│  - 同步 prompt（02-*.json）│
│  - 处理 chart 等组件        │
│  - 接入 Main.py             │
└──────────────┬──────────────┘
               │
       ┌───────┴───────┐
       │               │
       ▼               ▼
   一切顺利         遇到复杂问题
   通过冒烟测试    （路线决策 / 沉默 bug）
       │               │
       ▼               ▼
   完成             阶段 5：主 Claude 对话
                    - 路线讨论
                    - 多轮战略 pivot
                    - 复杂 bug 复盘
                         │
                         ▼
                       完成
```

---

## 3. 各阶段动作清单

### 阶段 1：Pipeline 首跑（必做）

**入口**：`python orchestrator.py` → 选项 ① 全自动

**期望产出**：
- `pipeline-progress/01-shape_detail.xlsx` — shape 清单 + 用户可填的"内容描述"列
- `pipeline-progress/02-prompt_specs.json` — 每 shape 最终 prompt
- `pipeline-progress/03a-build_shape_content.json` — GPT 生成内容
- `pipeline-progress/04-fix_ppt.md` — 自检报告

**用户参与点**：
- Step 1 后查看 `01-shape_detail.xlsx`，对关键 shape 填"内容描述"列（黄色单元格）
- 这个标注会驱动 Step 2 的 strategy 推断（参考 yzr 实测：9 个 shape 标注过）

### 阶段 2：评估 PPT 效果（决策节点）

**评估输入**：
- 03b 阶段产出的 PPT 文件
- `04-fix_ppt.md` 中的三项分数

**评估标准**（建议）：

| visual | readability | semantic | 决策 |
|--|--|--|--|
| ≥80% | ≥80% | ≥60% | ✅ 进入移植（阶段 3a） |
| <80% 任一项 | — | — | 🔄 继续 Pipeline 迭代（阶段 3b） |
| visual=100% but semantic<50% | — | — | ⚠️ shape 角色识别有误，回头改 .xlsx 标注 |

**用户主观判断**：肉眼看生成的 PPT，如果整体视觉接近模板就可以进移植。

### 阶段 3a：跳过迭代，进入移植

直接到阶段 4。

### 阶段 3b：继续 Pipeline 迭代（可选）

**做法**：
- 修改 `01-shape_detail.xlsx` 的"内容描述" / `strategy` / `params` 列
- 重跑 Step 2 + Step 3（可在 `orchestrator.py` 选项 ②③ 分步执行）
- 视情况 2-3 轮迭代，达到满意度后进阶段 4

### 阶段 4：`/developer` 移植（首选）

**调用方式**：在 Claude Code 里输入 `/developer 把 yzr 模板的 Pipeline 产物移植到 src/`

**Agent 应做的事**（按 `developer.md` 场景 2 Checklist）：

```
□ 新建 src/{template}_ppt.py（复制 yzr_ppt.py 骨架）
□ 替换 SHAPES 列表 + _TEMPLATE_SLIDE
□ 从 02-prompt_specs.json 提取 prompt → 写入 _build_rich_prompt()
□ 加 prompt_src / synced_at 注释（fix2 范式）
□ import _ppt_shared 共享工具
□ 处理图表分支（决策树）：
   - 系列固定 + 模板已含 chart shape → _write_chart() 原位（仅单机自用场景）
   - 分发场景 / 跨机 / 加密 → make_chart_for_{name}（fix4 范式，强制）
□ 接入 Main.py 选择逻辑
□ 跑 debug/test_src_smoke.py
□ 语法检查 + 端到端运行验证
```

**Agent 模型**：Sonnet（developer.md 指定，token 成本约为 Opus 的 1/3）

### 阶段 5：主 Claude 对话（复杂问题兜底）

**何时切换到主 Claude**：

| 触发信号 | 例子 |
|--|--|
| 路线决策 | fix3 → fix4 切换"改模板"到"从零制表" |
| 多轮战略 pivot | 同一 bug 试 ≥2 个方案都失败 |
| 沉默失败 bug | chart bars 消失但无异常（fix3 chart 写入） |
| 跨机/分发场景适配 | yzr 同事机器 Build 4266 兼容性 |
| 涉及 Pipeline ↔ src/ 边界 | fix2 的 _ppt_shared 抽取 |

**做什么**：路线讨论 / 候选方案对比 / fix{N}.md 决策档案 / memory 固化经验

---

## 4. 关键工具索引（按阶段）

| 阶段 | 工具 | 文件 |
|--|--|--|
| 1（Pipeline 首跑） | `orchestrator.py` | 项目根目录 |
| 1（手工标注）| Excel 编辑 | `pipeline-progress/01-shape_detail.xlsx` |
| 2（视觉检查） | 直接打开 PPT | Pipeline 产出文件 |
| 4（移植 checklist） | `developer.md` 场景 2 | `.claude/agents/developer.md` |
| 4（fine_tuned 微调） | `read_selected_shape.py` + `fine-tuned-shapes.md` | `skills/` |
| 4（chart 决策） | fix4 路线 + `feedback_chart_write.md` | `[feature03-transplant]/fix4（图表路线切换）.md`、`.claude/memory/` |
| 5（主 Claude 反射） | `feedback_debug_protocol.md` | `.claude/memory/` |

---

## 5. 落地建议（按优先级）

| 优先级 | 改动 | 预估成本 |
|--|--|--|
| ★★★ | **CLAUDE.md §1 双轨架构补一段"工作流路由"**：把上面阶段 1-5 流程图嵌入 §1，让每次接到新模板任务时第一时间能查到路径 | 30 分钟 |
| ★★★ | **`.claude/agents/developer.md` 增加"Pipeline 产物消费手册"小节**：明确告诉 Agent 拿到 `02-prompt_specs.json` / `02-shape_analysis_map.json` 时怎么用（已经有 Checklist，再具体一些） | 30 分钟 |
| ★★ | **`orchestrator.py` 跑完后追加一段提示**："Pipeline 已完成，建议下一步：(a) 检查 04-fix_ppt.md，(b) 视效果选择继续迭代或调用 /developer 移植" | 15 分钟 |
| ★★ | **`.claude/memory/` 新建 `feedback_workflow_routing.md`**：把"5 阶段流程 + 触发信号 + 工具索引"做成 LLM 可检索的记忆，每次相关场景自动加载 | 30 分钟 |
| ★ | **写一个 `skills/port_handoff_checklist.md`**：物理列出"Pipeline 跑到哪一步 → Developer 接手时需要哪些产物 → 每个产物的字段含义"。供 Developer Agent 启动时读取 | 1 小时 |
| ★ | **prompt 自动 diff 工具（来自 fix2）**：当 Pipeline 的 `gpt_summary.md` 有变化时，提示 src/ 哪些模板需要 sync。低优先级，等下次 Pipeline 升级时再做 | 半天 |

---

## 6. 不做的事（避免 scope creep）

- ❌ 不删 Pipeline 代码
- ❌ 不重构 orchestrator.py
- ❌ 不改 yzr_ppt.py / zxh_ppt.py 的现有架构（已稳定）
- ❌ 不引入新的"自动 Pipeline → src/ 同步"工具（保留手工同步 + 注释追溯，符合 fix2 决策）
- ❌ 不强求 /developer 处理路线决策类任务（保留主 Claude 对话作为复杂问题兜底）

---

## 7. 决策点速查表（每次新任务用）

| 任务类型 | 默认路径 |
|--|--|
| **完全新模板** | 阶段 1（必跑）→ 2（评估）→ 3a/3b → 4（/developer）→ 5（如卡住） |
| **已知模板加新 shape** | 跳过 1-3，直接阶段 4（/developer 改 SHAPES 列表） |
| **已知模板 bug 修复** | 跳过 1-3，直接阶段 5（主 Claude，因为 bug 通常涉及路线判断） |
| **prompt 文案调优** | 跳过 1-3，阶段 4（/developer 改 _build_rich_prompt）|
| **shape 微调** | 跳过 1-3，阶段 4（/developer + skills/fine-tuned-shapes.md） |
| **chart 路线问题（如 fix4 类）** | 跳过 1-3，阶段 5（主 Claude） |

---

## 8. 一句话总结

> **三重混合机制不推翻、不降级，但流程标准化**：新模板必跑 Pipeline 首轮 → 视效果决定迭代或移植 → 移植走 /developer → 复杂问题回主 Claude。
>
> Pipeline 的价值在初次接入；/developer 的价值在批量执行；主 Claude 的价值在路线决策。三者各司其职，无重叠浪费。

---

## 9. 参考档案

- `[feature03-transplant]/plan2（三重混合机制再评估）.md` —— 三重混合机制再评估
- `[feature03-transplant]/fix2（三重混合架构整改）.md` / `fix3（图表写入诊断）.md` / `fix4（图表路线切换）.md` —— 历次决策档案
- `【trash-bin】/pipeline-progress-yzr/` —— yzr 早期 Pipeline 完整跑批的实测产物（已 stash）
- `.claude/agents/developer.md` —— Developer Agent 角色与移植 Checklist
- `.claude/memory/feedback_chart_write.md` —— chart 写入经验（fix4 落地）
- `.claude/memory/feedback_debug_protocol.md` —— 主 Claude 调试反射动作（fix3→fix4 教训）
- `debug/Mc-debug-5-三重混合机制.md` —— 三重混合机制设计原始讨论
