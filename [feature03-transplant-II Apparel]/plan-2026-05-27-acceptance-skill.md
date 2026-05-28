# 计划：ppt-acceptance-check —— 完整的 PPT 验收 skill

**日期**：2026-05-27（rev2，重写）
**触发**：昨天 apparel 双页移植验收漏掉 6 类问题中的 4 类；用户要求一个**完整的整体验收 skill**，不是"修补昨天 bug"的副产品
**作用域**：`~/.claude/skills/ppt-acceptance-check/`（用户级）
**作者**：Claude
**状态**：等用户拍板 §九 决策点

---

## 一、设计原则（先定原则、再定能力）

| 原则 | 含义 | 反例（昨天踩到的） |
|---|---|---|
| **完备性优先** | skill 必须覆盖 PPT 验收的所有维度，不为某一个 bug 设计 | 昨天 diff probe 只看 shape 几何+字体，漏 chart 数据、漏 GPT 行为 |
| **配对鲁棒性是地基** | shape 配对失败 → 后面所有层都白做。这是 L0，不是 feature | 昨天 `common = a ∩ b` 严格按 Name，Clone 自动改名直接漏检 |
| **三态分类** | 每条规则都标 `must_fix / warn / tolerate_if_*`，自动判 PASS/FAIL | 昨天 diff 报告所有差异一锅出，得人审，等于半自动 |
| **项目契约外置** | "哪些 shape 是数据型、阈值多少、预期值哪来" 写在 JSON 契约，skill 是引擎 | 昨天 diff probe 写死 ±2 几何容差，其他字段全严格 |
| **一份报告一次决策** | 跑一次 = 一份 md + 一份 json，PASS/FAIL 二元结论，AI 不用跨工具拼信息 | 昨天要看 SSIM、看 diff、看 runs、看人工选中 shape，4 份产物拼不出结论 |
| **不修复，只断言** | skill 不改代码、不重生成 PPT。fail 了交回项目层（developer / 人）处理 | —— |

---

## 二、PPT 验收的 6 层维度（昨天只覆盖 2-3 层）

| 层 | 名称 | 检查的是 | 昨天 diff probe 覆盖？ | 昨天踩到的典型漏 |
|---|---|---|---|---|
| **L0** | 结构 / 配对 | shape 存在性 + 名字一致性 + 数量 + 类型 + 仅 A / 仅 B 清单 | ❌ 严格 Name 配对，Clone 改名静默漏 | Chart 63（模板） vs Chart 8/12/15/18（新生成）— diff 报告"仅在 A/B"列表抓到了，但没标 must_fix；runs 报告全空就是配对失败 |
| **L1** | 数据 | TextBox 文本值、Chart series.Values、Picture 源/裁剪、Table 单元格 vs 期望值 | ❌ 完全没读 chart 数据 | Chart 63 series 是模板默认值 `[5,15]/[20,17]/[10,3]`、TextBox 50 mode 取错 |
| **L2** | 格式 | shape 几何 (L/T/W/H)、fill、line、paragraph 级 font/size/color/align/spacing/AutoSize | ✅ 部分（diff_shape_detail.py） | TextBox 24 高度差 100pt、7 个 shape AutoSize=0 vs 1 |
| **L3** | 染色 / Runs | character-run 级 font/color/bold 变化、"应该有 N 个染色 runs" 结构断言 | ⚠️ 有脚本但跑出空报告（配对失败） | TextBox 23/26 染色 bullet 缺失 |
| **L4** | 行为 / Trace | GPT 调用日志、fallback 路径检测、COM API 失败检测、pipeline 执行 trace | ❌ 完全没有 | Chart 63 `ChartData.Activate()` 失败 3 次，代码继续走、smoke test gpt=n 走 fallback，3 个 GPT 槽实际没调 |
| **L5** | 视觉 | PNG 渲染 + SSIM + （可选）像素 diff 热力图 | ✅ ppt-visual-fidelity-check 已有 | p14 SSIM 0.7998，但给数字不给位置 |

**关键观察**：

- 昨天 diff probe 实际只覆盖 L2 部分 + L3 坏的，**L0/L1/L4 完全没有**。
- L0 是地基——它失败时 L2/L3/L4/L5 都失真。runs 报告 83 字节全空就是因为 L0 配对失败把所有 shape 跳过了。
- L4「行为层」是昨天**根本没有意识到要存在**的维度。Chart 63 Activate 失败但代码继续走，这件事 PPT 产物上看不出来（数据是模板默认值，视觉正常），**只能从 pipeline 执行 trace 里捞**。

---

## 三、统一入口

```bash
python -m ppt_acceptance_check \
    --new "v1.6.pptx" \
    --template "apparel-page13-14-template.pptx" \
    --slide-pairs "12:13,13:14" \
    --contract acceptance/apparel.json \
    --pipeline-trace debug/apparel_trace.jsonl \
    --layers L0,L1,L2,L3,L4,L5 \
    --out-dir debug/acceptance-apparel/
```

| 参数 | 必需 | 说明 |
|---|---|---|
| `--new` | 是 | 新生成 PPT；支持 `--active` 接管打开中的 PowerPoint |
| `--template` | 是 | 标杆 PPT |
| `--slide-pairs` | 是 | `new_idx:template_idx,...`（apparel = `12:13,13:14`） |
| `--contract` | 否 | 项目契约 JSON；无则全严格 + 默认配对策略 |
| `--pipeline-trace` | 否 | pipeline 跑时落盘的 jsonl trace（**L4 必需**）；无则 L4 降级到"只查产物，不查行为" |
| `--layers` | 否 | 默认 6 层全跑 |
| `--out-dir` | 否 | 默认 `debug/acceptance-<timestamp>/` |

**入口设计原则**：参数最小、契约外置、layer 可裁剪（debug 时单跑某层）。

---

## 四、契约文件（项目特化外置）

**`acceptance_contract.json` 是 skill 通用化的核心**——把"这个项目的具体规则"从 skill 引擎里拆出去。

```json
{
  "version": 1,
  "shape_pairing": {
    "strategy": "fuzzy",
    "rules": [
      "strip_clone_suffix",
      "ignore_chart_auto_renumber",
      "match_by_position_if_name_lost"
    ],
    "manual_overrides": {
      "p13": {
        "Chart 8 [new]": "Chart 63 [template]"
      }
    }
  },
  "tolerances": {
    "geometry_px": 2,
    "font_size_pt": 0.5,
    "color_rgb_distance": 0,
    "ssim_threshold": 0.96
  },
  "rules": [
    {
      "id": "chart63_temp_range",
      "shape": "Chart 63",
      "layer": "L1",
      "check": "chart_series_values",
      "expected_from": "excel:服装试穿问卷--紧身背心:AD,AE:temp_range",
      "severity": "must_fix"
    },
    {
      "id": "tbx50_temp_mode",
      "shape": "TextBox 50",
      "layer": "L1",
      "check": "text_equals_computed",
      "expected_from": "excel:服装试穿问卷--紧身背心:AD:mode",
      "severity": "must_fix"
    },
    {
      "id": "gpt_strengths_color_runs",
      "shape": "TextBox 23",
      "layer": "L3",
      "check": "has_color_runs",
      "expected_min": 2,
      "severity": "must_fix",
      "skip_if_trace": "gpt_strengths.called == false"
    },
    {
      "id": "gpt_must_be_called",
      "layer": "L4",
      "check": "trace_event_present",
      "expected_events": ["gpt_strengths", "gpt_drawbacks", "gpt_subject_info"],
      "severity": "must_fix",
      "scope": "production"
    },
    {
      "id": "no_silent_com_failure",
      "layer": "L4",
      "check": "no_trace_event",
      "forbidden_events": ["com_api_failed_but_continued"],
      "severity": "must_fix"
    },
    {
      "id": "geometry_global",
      "shape": "*",
      "layer": "L2",
      "check": "geometry_within_tolerance",
      "severity": "must_fix"
    },
    {
      "id": "autosize_global",
      "shape": "*",
      "layer": "L2",
      "check": "autosize_matches_template",
      "severity": "warn"
    }
  ],
  "modes": {
    "smoke": {
      "rule_overrides": {
        "gpt_must_be_called": "tolerate",
        "gpt_strengths_color_runs": "tolerate"
      }
    },
    "production": {}
  }
}
```

**三态 severity → exit code**：

| severity | 含义 | exit |
|---|---|---|
| `must_fix` | 必修 | 1 |
| `warn` | 记录不阻断 | 0 |
| `tolerate` | 豁免（条件或模式） | 0 |

**两个关键机制**：

1. **`skip_if_trace`**：规则可以引用 trace 事件做条件豁免（"GPT 没调 → 不查 GPT 染色"）
2. **`modes`**：smoke 模式自动把一批规则降级为 tolerate，避免 dev 阶段误报

---

## 五、L4 行为层的契约：pipeline trace 格式

L4 是新增维度，需要 pipeline 在跑的时候**主动落盘 trace**。约定 jsonl 格式：

```jsonl
{"ts":"2026-05-27T10:00:01","event":"shape_write_start","slide":12,"shape":"Chart 63","strategy":"bar_stacked_temp_range"}
{"ts":"2026-05-27T10:00:01","event":"com_api_failed_but_continued","slide":12,"shape":"Chart 63","api":"ChartData.Activate","attempts":3,"continued":true}
{"ts":"2026-05-27T10:00:02","event":"shape_write_end","slide":12,"shape":"Chart 63","ok":false}
{"ts":"2026-05-27T10:00:03","event":"gpt_strengths","slide":13,"called":false,"reason":"mc_gpt=n"}
```

**对项目侧的要求**（小、可选、不破坏现有代码）：
- 在 `office-com-helpers` 加一个 `TraceLogger`（上下文管理器）
- pipeline 代码用 `with trace.shape_write("Chart 63", "bar_stacked_temp_range"):` 包一下关键调用
- COM 失败兜底处用 `trace.event("com_api_failed_but_continued", api=..., attempts=...)`

**渐进迁移**：trace 不接入也能跑 L4 ——只是"只查产物（fallback 文本判定）"，不能查"COM 失败但继续走"这类暗坑。

---

## 六、报告样式

主报告 `acceptance_report.md`（人看）+ `acceptance_report.json`（AI 自检 / CI 读）。

```markdown
# PPT 验收报告

**结论：FAIL** （必修 3 / 警告 5 / 容忍 2）
**模式**：production
**契约**：acceptance/apparel.json
**slide pairs**：12:13, 13:14

## 摘要
| 层 | 通过 | 警告 | 容忍 | 必修违反 |
|---|---|---|---|---|
| L0 配对 | 26 | 0 | 0 | 0 |
| L1 数据 | 4 | 0 | 0 | 2 |
| L2 格式 | 20 | 7 | 0 | 0 |
| L3 染色 | 6 | 0 | 2 | 1 |
| L4 行为 | 3 | 0 | 0 | 0 |
| L5 视觉 | 1 | 0 | 0 | 0 |

## L0 配对：fuzzy 命中
- Chart 63 [template:p13] → Chart 8 [new:p12]（按 manual_overrides）
- TextBox 6 [template:p13] → TextBox 6 [new:p12]（严格命中）
- ... (24 more)

## 必修违反清单（先看这里）

### [L1] chart63_temp_range
- shape: Chart 63 [template] / Chart 8 [new]
- 实际 series.Values: [[5,15],[20,17],[10,3]] = 模板默认值
- 期望: [...由 Excel AD/AE 算] = [[6,8],[18,5],[11,2]]
- 修复方向: pipeline 的 _write_chart63 BreakLink 路径

### [L1] tbx50_temp_mode
- 实际: "5~15℃"  期望: "15~25℃"（mode of AD = 15~25, 4/5 票）

### [L3] gpt_strengths_color_runs
- shape: TextBox 23 [p14]
- 实际 color runs 数: 0  期望: ≥2

## 警告清单（不阻断）
- [L2] 7 个 shape autosize=0 vs 模板 1 ...

## 容忍清单（已豁免）
- [L3] gpt_drawbacks_color_runs：因 trace.gpt_drawbacks.called=false 豁免

## L5 视觉层
- p12 vs p13 SSIM = 0.958
- p13 vs p14 SSIM = 0.799（阈值 0.96，失败已计入必修）
- PNG 输出：debug/acceptance-apparel/slide_12_*.png, slide_13_*.png
```

---

## 七、skill 落地阶段（独立于任何具体项目）

| 阶段 | 内容 | 关键设计点 | 预估 |
|---|---|---|---|
| **S0 地基** | 骨架 + 契约引擎 + **L0 配对（fuzzy + manual_overrides）** + 三态 severity 解析 + exit code | L0 不是 layer 之一，是地基；其他 layer 跑之前先过 L0 | 1.5h |
| **S1 L1 数据层** | chart series 回读 + text 回读 + Picture/Table；`expected_from` resolver（先支持硬编码 + excel: 两种） | excel: DSL 第一版只支持 `sheet:col:agg(mode/sum/mean/count_contains)`；其他后续加 | 2h |
| **S2 L2 格式层** | 搬 `diff_shape_detail.py`，按契约 tolerances 过滤；分离 must_fix / warn 字段 | autosize 默认 warn 不 must_fix（容易刷屏） | 1h |
| **S3 L3 染色层** | 搬 `diff_shape_runs.py` + **修配对漏（基于 S0 L0）** + `has_color_runs` / `runs_match_template` 两种断言 | runs 报告别再空了；要么有差异要么有"shape 配对失败"的明确报错 | 1h |
| **S4 L4 行为层** | `TraceLogger`（office-com-helpers 加）+ trace jsonl 解析 + `trace_event_present` / `no_trace_event` 两种断言 | trace 不接入 → L4 降级"只查产物 fallback 文本判定"；要明确给降级原因 | 1.5h |
| **S5 L5 视觉层** | 吸收 `ppt-visual-fidelity-check` 的 slide→PNG + SSIM；老 skill 标 deprecated | SSIM 阈值从契约读，不写死 | 0.5h |
| **S6 报告聚合** | 6 层结果汇总 → md + json；exit code 决策 | 报告头摆"结论"，必修清单永远在最上 | 1h |
| **S7 文档** | SKILL.md + 一份 `acceptance_contract.example.json` + 一份"如何为新项目写 contract"指南 | 帮其他人/未来 AI 上手 | 1h |

**总预算 ~9.5h**（比 rev1 翻倍，主要在 S0 配对地基 + S4 行为层）。每步独立可测、可暂停。

**S0 之后是骨架**：跑 `python -m ppt_acceptance_check --new x.pptx --template y.pptx --slide-pairs 1:1 --layers L0` 就能验证 L0 跑得通；后续 S1-S5 一层一层补。

---

## 八、对现有 skill / 工具的处置

| 对象 | 处置 |
|---|---|
| `~/.claude/skills/ppt-visual-fidelity-check/` | 保留 + SKILL.md 标 deprecated，指向 `ppt-acceptance-check --layers L5` |
| `3rd-ppt-prj/debug/diff_shape_detail.py` | 搬进 skill 的 `layers/format.py`；原位留 redirect 注释 |
| `3rd-ppt-prj/debug/diff_shape_runs.py` | 搬进 skill 的 `layers/runs.py`；原位留 redirect 注释 |
| `~/.claude/skills/pipeline-self-check-loop/SKILL.md` | 反射文档补一句"L0-L5 自检统一调 ppt-acceptance-check" |
| `~/.claude/skills/office-com-helpers/` | 增加 `TraceLogger` 上下文管理器供 L4 用 |
| `~/.claude/skills/inspect-office-template/` | 不变；它是开工前探查，与验收正交 |

---

## 九、3 个用户拍板的决策点

| # | 决策 | 我的倾向 | 原因 |
|---|---|---|---|
| 1 | skill 作用域 | **用户级** `~/.claude/skills/` | feedback memory「reviewer 默认通用 + 用户级」；apparel/zxh/yzr + 未来 PPT 项目通用 |
| 2 | L4 trace 接入方式 | **TraceLogger 加在 office-com-helpers** | 单点改动，所有项目零成本接入；不接入则 L4 自动降级 |
| 3 | 是否吸收 ppt-visual-fidelity-check | **吸收（标 deprecated 不删）** | 不吸收则违背"整体验收一个 skill 一个入口"诉求；不删保护存量项目 |

---

## 十、Out of Scope（不在本 skill 范围）

明确说不做的，避免范围漂移：

- ❌ **自动修复**：skill 只断言不修；fail 了交 developer agent / 人
- ❌ **代码层 review**：`.py` 源码审查走 `/code-review` 或 `code-reviewer` agent
- ❌ **PPT 之外**：Word / Excel 验收不在 v1 范围（设计上可扩，v2 再说）
- ❌ **跨页流程检查**：例如"目录页页码必须和正文一致"——业务规则，不是验收
- ❌ **模板探查 / 设计验证**：那是 `inspect-office-template` 的事，开工前跑

---

## 十一、应用：apparel 是首个 dogfood，不是设计目标

**关系澄清**：
- skill 完整落地（S0-S7）是**独立项目**，不为 apparel 让步
- apparel 收尾用 skill 来做，是**验证手段**，跑出报告后 apparel 项目自己决定怎么修（developer agent 或人）
- 如果 apparel 报告显示 skill 漏检了什么 → 改 skill；如果显示真实 bug → 改 apparel 代码

**apparel dogfood 路径**（skill S0-S7 完成后）：

1. 在 `3rd-ppt-prj/` 写 `acceptance/apparel.json`（参考 §四样板）
2. 用户跑真实 GPT 模式 Main.py 出 v1.6（不再 smoke test 模式）
3. 跑 `ppt-acceptance-check --new v1.6 --template apparel-page13-14-template.pptx --slide-pairs 12:13,13:14 --contract acceptance/apparel.json --mode production`
4. 看必修清单 → developer 修 → 重跑 → 必修 0 + SSIM ≥0.96 → 收工
5. 收尾 `[feature03-transplant-II Apparel]/fix5（视觉细节调优）.md` + STATE.md

**预期**：报告会直接复现 6 类问题（Chart 配对漂移 / Chart 63 数据未生效 / TextBox 50 mode 错 / GPT fallback / autosize / SSIM 0.80）；这就是 dogfood 验证 skill 设计合理性的方式。

---

## 十二、风险

| 风险 | 应对 |
|---|---|
| S0 fuzzy 配对策略复杂度被低估 | 第一版只做 `strip_clone_suffix` + `manual_overrides`；按位置匹配 (`match_by_position_if_name_lost`) 留 v2 |
| S1 `expected_from` DSL（`excel:sheet:col:agg`）抽象成本高 | 第一版只做"硬编码值"+ 一种 `excel:sheet:col:agg(mode\|sum\|mean\|count_contains)`；其他后续 |
| L4 trace 接入推不动各项目 | trace 不接入时 L4 自动降级，报告里**显式标"L4: degraded, no trace"**——把 visibility 给到用户，让他自己决定要不要补 |
| 契约文件复杂、新项目门槛高 | S7 写"如何为新项目写 contract"指南 + `acceptance_contract.example.json`；按 80/20 给"开箱即用"的全严格模式 |
| skill 跑得慢（6 层 + COM 调用） | layer 可裁剪（`--layers L0,L1,L4`），开发时只跑关键层；L5 SSIM 最慢，可单独跑 |
| 用户每次跑 Main.py 都要手动跑 skill | 后续可考虑 Stop hook 自动触发；先不做，避免过度工程 |

---

## 十三、依赖与引用

- `[[feedback_reviewer_scope]]` — reviewer 默认通用 + 用户级
- `[[feedback_agent_handoff_gate]]` — 探查→@agent 之间插人工确认门
- `3rd-ppt-prj/CLAUDE.md` — apparel 项目硬规则
- `3rd-ppt-prj/debug/Mc-debug-6-apparel修复.md` — 昨天踩坑原始记录
- `3rd-ppt-prj/debug/diff_p13p14_detail.md` — 昨天 diff probe 实际产出（确认有效）
- `3rd-ppt-prj/debug/diff_p13p14_runs.md` — 昨天 runs probe 全空（确认配对失败漏）
- `3rd-ppt-prj/plan-apparel-2page-2026-05-26.md` — 昨天 apparel 双页移植 plan
- `Mc-emoji/[feature-03-take-a-nap]/Claude-工作模式优化评估-当前版-2026-05-25.md` §四 — 8 个用户级 skill 盘点

---

## 修订记录

- **rev2 (2026-05-27)**：用户指出 rev1 把 skill 绑给 apparel 收尾，违背"整体验收"诉求。重写：
  - 新增 §一设计原则
  - 维度从 4 层（数据/格式/染色/视觉）扩到 6 层（+L0 配对 + L4 行为）
  - skill 落地阶段 S0-S7 解耦 apparel
  - 新增 §五 L4 trace 格式约定
  - 新增 §十 Out of Scope
  - 新增 §十一 apparel dogfood 路径（不再是设计目标）
- **rev1 (2026-05-27)**：基于 Wave 1/2/3 修复 + skill 草案整合（已废，被 rev2 替代）
