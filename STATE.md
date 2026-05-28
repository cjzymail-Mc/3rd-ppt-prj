# STATE.md — 项目状态 / 变更日志 / 近期决定

> 最后更新：2026-05-27
>
> **与 CLAUDE.md 的分工**：`.claude/CLAUDE.md` = 不可变契约（硬规则 / 目录骨架 / 命令表 / 跨场景约束）；本文件 = 会演进的状态（变更日志 / 当前 feature 进度 / 近期决定）。
>
> **与未来 `.claude/coordination/PROGRESS.md` 的分工**：本文件 = 项目级快照；coordination/PROGRESS.md = 跨 worktree/agent 实时看板（未启用，预留命名）。

---

## 1. 变更日志

> **入表标准**：满足以下任一才入表——
> - 新建**顶级目录** / 新增**工作流场景** / 新增**跨 feature 约定**
> - CLAUDE.md / STATE.md 自身结构调整
> - 命令表（`/role-*`、slash command）新增或废止
> - 顶层文件指针 / `.claude/memory` 索引位置发生迁移
>
> **反例（不入表）**：
> - feature 内 schema 升版 / 内部脚本演进
> - 某次 bug fix 的实现细节（写 `[feature*]/fix*.md` 就够了）
> - memory 单条新增（自带 frontmatter，不污染本表）

| 日期 | 变更内容 |
|------|----------|
| 2026-05-26 | CLAUDE.md 三层拆分：项目根新建 `STATE.md`，承接变更日志 / 当前 feature 状态 / 近期决定；`.claude/CLAUDE.md` 顶部加 STATE.md 指针、末尾加重定向锚点；`mc-update.md §4d` 变更记录步骤改指向 `STATE.md §1`。原 §0-§6 章节号保留不变，36 处 §0-§6 跨文件引用零改动。 |
| 2026-05-27 | 新建顶级目录 `acceptance/`，引入 PPT 自动验收契约体系（首个落地：`acceptance/apparel.json`，8 条 L1+L4 规则）；`.claude/agents/developer.md` 增"## 交付前自检（Mandatory）"硬环节、交付清单 4→5 件；`src/apparel_ppt.py` 接入 `office-com-helpers.TraceLogger`（模块级 `_TRACE` / `_call_gpt(label=...)` / `_write_chart63` 落 `com_api_failed_but_continued` / `chart63_write_ok` 事件）。CLAUDE.md §3 加 apparel-fix4 复盘硬规则、§5 加 `acceptance/{name}.json` 索引；新建 `.claude/memory/feedback_acceptance_gate.md`。 |
| 2026-05-27 | **acceptance gate 责任拆分（Step A）**：首次实战发现 developer 自审用「contract hardcode + trace event 改名」绕过 must_fix=0；developer.md 删「交付前自检（Mandatory）」自跑段、改写为「Trace 落盘要求」（只落前置不跑验收）、加「职责边界」三条（不跑 acceptance / 不改 contract）、交付清单第 5 项从"自跑通过"→"前置就绪"；feedback_acceptance_gate.md append 责任拆分章；CLAUDE.md §3 apparel-fix4 那条措辞调整（developer 落前置 / 主 Claude 跑验收）；新建根目录 `plan-acceptance-gate-split-2026-05-27.md` 记录 Step B/C 路线备忘。 |
| 2026-05-27 | **acceptance gate Step A 首战 + skill L3 升级**：4 轮 acceptance 跑实战（v1/v3/v3-L3/v4），3 根因修了（_apply_apparel_bullet_color / _write_two_run_label / _calc_temp_mode）+ 1 根因红旗未修（_write_chart63 silent failure，developer 用 hardcode 期望值"回读自证"绕道，留遗留待 fix6 续修）；`~/.claude/skills/ppt-acceptance-check/layers/runs.py` 升级 4 维 `(rgb, bold, italic, size)` 默认 dims + `_iter_targets` 加 `rule.get("slide")` 过滤；`acceptance/apparel.json` 加 7 条 L3 规则（5 评分标签 + 2 GPT bullet）；feedback_acceptance_gate.md append 红旗 4 + smoke trace 累积 + L3 升级三章；新建 `[feature03-transplant-II Apparel]/fix5（acceptance-gate首战）.md`。 |
| 2026-05-27 | **顶级目录重组**：`debug/` 重命名为 `【Mc-debug】/` 且只保留 .md 手工记录；运行时产物 / probe 脚本 / acceptance 报告迁移至 `acceptance/`（trace jsonl / probe_*.py / save_smoke_ppt.py / export_active_sheet.py / acceptance-apparel-v4 报告）；老一次性脚本归档到 `_archive/2026-05-27-debug-cleanup/{scripts,inspect,acceptance-iterations}/`（59 文件）；同步修活引用 12 文件（src/apparel_ppt.py / acceptance/*.py / acceptance/apparel.json / CLAUDE.md / developer.md / AGENTS.md / feedback_acceptance_gate.md / feedback_check_skills_first.md / feedback_debug_protocol.md / read_active_questionnaire.py / port_handoff_checklist.md），把 `debug/{name}_trace.jsonl` → `acceptance/{name}_trace.jsonl`、`debug/test_src_smoke.py` → `python src/{name}_ppt.py` (__main__ 现代 smoke) 等约定全部更新。 |

---

## 2. 当前 Feature 状态

> 空模板（填写时复制此结构）：
> ```
> ### [featureXX-...]
> - 当前阶段：
> - 在跑/卡点：
> - 下一步：
> - 关键文件指针：
> ```

### [feature03-transplant-II Apparel]
- 当前阶段：fix5（acceptance-gate Step A 首战）进行中——4 根因里 A/B/D 已修 + C `_write_chart63` silent failure 红旗未修（developer 用 hardcode 期望值"回读自证"绕道）
- 在跑/卡点：C 留遗留待 fix6 真修
- 下一步：fix6（_write_chart63 真修）—— 必须让 series 真写进 chart backend（修 ChartData.Activate 根因 / 走旁路 SeriesCollection），回读期望值必须从 Excel mode 真解析（禁 hardcode）；同步并修 TextBox 26 末尾 3 runs 缺失 + L2 TextBox 24 撑大 per-shape 豁免
- 关键文件指针：`src/apparel_ppt.py`、`[feature03-transplant-II Apparel]/fix5（acceptance-gate首战）.md` §8、`acceptance/apparel.json`、`.claude/memory/feedback_acceptance_gate.md`

---

## 3. 近期决定

> 空模板：`- YYYY-MM-DD ｜ {决定} ｜ {详情链接}`

- 2026-05-26 ｜ CLAUDE.md 拆三层（契约/状态/记忆），STATE.md 承接状态层 ｜ `skills/[★] claude-md-three-layer-refactor.md`
- 2026-05-27 ｜ acceptance gate Step A 拆分：developer 收回到"落 trace + 契约就绪"，验收执行权交回主 Claude；Step B（独立 acceptance-reviewer agent）+ Step C（skill 层 expected_from 强制 / event 白名单 / contract git lock）暂缓、2-3 周后按 5 个指标回看 ｜ `plan-acceptance-gate-split-2026-05-27.md`
