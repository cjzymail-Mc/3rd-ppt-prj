---
name: feedback_workflow_routing
description: 接到新模板/新任务时的 5 阶段工作流路由（plan3 定稿）
type: feedback
---

每次接到 PPT 模板相关任务时，按以下路由判断走哪条路。这是 plan3（工作流5阶段定稿）.md §0 用户决策的固化。

**Why:** yzr 历史经历证明：Pipeline 在初次接入有真实价值（fix2 阶段把 prompt 同步到 src/ 是有效的），但已知模板修 bug 阶段（fix3/fix4）不需要 Pipeline。**没有路由意识就会每次重复判断走哪条路**——浪费时间 + 容易选错。

**How to apply:** 看任务类型 → 查表选路径 → 按 5 阶段执行。

---

## 5 阶段工作流（plan3 定稿）

```
新模板/数据源到手
       ↓
① Pipeline 首跑（必做）        orchestrator.py 全流程
       ↓
② 评估 PPT 效果（决策点）      看 04-fix_ppt.md
       ↓
   ┌───┴───┐
   ↓       ↓
③a 跳过   ③b 继续 Pipeline 迭代
   ↓       ↓
   └───┬───┘
       ↓
④ /developer 移植（默认路径）  Sonnet + 自动加载 developer.md
       ↓
⑤ 主 Claude 兜底（复杂问题）   路线决策类才回主对话
```

---

## 决策点速查表（每次新任务先查）

| 任务类型 | 默认路径 | 备注 |
|--|--|--|
| **完全新模板** | ① → ② → ③a/③b → ④ → 卡住时 ⑤ | 必跑 Pipeline 首轮（用户决策原话） |
| **已知模板加新 shape** | 直接 ④ | /developer 改 SHAPES 列表 |
| **已知模板 bug 修复** | 直接 ⑤ | bug 通常涉及路线判断，主 Claude |
| **prompt 文案调优** | 直接 ④ | /developer 改 _build_rich_prompt |
| **shape 微调** | 直接 ④ | /developer + skills/fine-tuned-shapes.md |
| **chart 路线问题（fix4 类）** | 直接 ⑤ | 主 Claude，因为是路线决策 |

---

## 阶段触发信号速查

### 何时停留在 Pipeline（继续 ③b 迭代）

- 视觉满意度 < 80%
- shape 角色被错误识别（semantic 分数 < 50%）
- 用户标注（01-shape_detail.xlsx）需要修改

### 何时进 ④ /developer

- Pipeline 视觉满意度 ≥ 80%
- shape 角色已基本对位
- 准备同步 prompt 到 src/

### 何时切换到 ⑤ 主 Claude

| 触发信号 | 例子 |
|--|--|
| 路线决策 | fix3 → fix4 切换"改模板"到"从零制表" |
| 多轮战略 pivot | 同一 bug 试 ≥2 个方案都失败 |
| 沉默失败 bug | chart bars 消失但无异常 |
| 跨机/分发场景适配 | yzr 同事机器 Build 4266 兼容性 |
| 涉及 Pipeline ↔ src/ 边界 | fix2 的 _ppt_shared 抽取 |

---

## 工具索引（按阶段）

| 阶段 | 工具 |
|--|--|
| ① Pipeline 首跑 | `orchestrator.py` 全流程 |
| ① 手工标注 | 编辑 `pipeline-progress/01-shape_detail.xlsx`（黄色"内容描述"列） |
| ② 视觉检查 | 直接打开 03b 输出的 PPT |
| ② 自检报告 | 读 `pipeline-progress/04-fix_ppt.md` |
| ③b 迭代 | `orchestrator.py` 选项 ②③ 分步执行 |
| ④ 移植 | `/developer` slash command（自动加载 `.claude/agents/developer.md`） |
| ④ 微调 | `skills/read_selected_shape.py` + `skills/fine-tuned-shapes.md` |
| ④ chart 决策 | fix4 路线 + `feedback_chart_write.md` |
| ⑤ 主 Claude 反射 | `feedback_debug_protocol.md`（grep 优先 / 质疑约定 / 2 次失败熔断） |

---

## 反模式（避免）

- ❌ **跳过 Pipeline 首跑直接移植**：除非是已知模板。新模板首跑能省下 Developer 大量手工读 shape 的时间
- ❌ **Pipeline 跑出来不看就移植**：必须读 04-fix_ppt.md 决定继续迭代还是直接进 ④
- ❌ **复杂问题在 /developer 里反复 pivot**：子 Agent context 隔离，多轮战略讨论应该回主 Claude
- ❌ **简单任务用主 Claude（Opus）**：批量改文件、机械实施请用 /developer（Sonnet 省 token）

---

## 参考档案

- `plan3（工作流5阶段定稿）.md` —— 完整工作流定稿（含 5 阶段流程图 + 实测档案验证）
- `[feature03-transplant]/plan2（三重混合机制再评估）.md` —— 三重混合机制再评估
- `.claude/agents/developer.md` —— /developer 角色 + Pipeline 产物消费手册
- `.claude/memory/feedback_debug_protocol.md` —— 主 Claude 调试反射动作（fix3→fix4 教训）
- `.claude/memory/feedback_chart_write.md` —— chart 写入路线（fix4 落地）
