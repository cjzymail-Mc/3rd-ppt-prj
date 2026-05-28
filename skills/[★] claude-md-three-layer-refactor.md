# CLAUDE.md 三层拆分改造规范

> 用途：把单文件 `CLAUDE.md`（同时承担契约 + 状态 + 历史）拆成「契约层 + 状态层 + 记忆层」三层结构。
> 适用：任何使用 Claude Code 且 `CLAUDE.md` 已超 100 行 / 混杂演进史 / 章节引用分散的 repo。
> 来源：2026-05-26 本 repo 落地实战（[feature02-方彩霞ppt整理]/plan07.md）。

---

## 1. 目的与适用判据

**什么时候要拆**（任一命中即建议拆）：
- `CLAUDE.md` 超 100 行
- 含变更日志/changelog 类表格（>5 条历史条目）
- 某节出现「2026-XX-XX 后...、plan0X §Y 已停...」这类带日期的演进描述
- 跨文件引用 `CLAUDE.md §N` 超 20 处（grep `CLAUDE\.md.*§` 统计）
- 想找硬规则时被状态噪音淹没

**目标三层**：
| 层 | 文件 | 职能 | 加载时机 |
|---|---|---|---|
| 契约 | `CLAUDE.md` | 不可变硬规则、目录骨架、命令表、跨场景约束 | 每会话 system prompt |
| 状态 | `STATE.md`（项目根） | 变更日志、当前 feature 状态、近期决定 | 按需 Read |
| 记忆 | `.claude/memory/MEMORY.md` + `.claude/auto-memory/MEMORY.md` | 反射记忆（用户偏好/高频反射/技术细节） | auto-memory 每会话，memory 按需 |

---

## 2. 拆分判据：什么留 CLAUDE.md / 什么迁 STATE.md

**留 CLAUDE.md（契约层）**：
- 硬规则 / 禁止事项（如「禁用 position:absolute」「染色双闸门」）
- 顶层目录骨架（**不含**每个 feature 的内部脚本演进史）
- 自定义命令表 / 角色清单
- 跨场景执行行为规范（防卡顿、记忆纪律、超时硬约束等）
- 任何「换会话/换协作者也不变」的约束

**迁 STATE.md（状态层）**：
- 变更日志（changelog）整张表 + 入表标准说明
- 当前活跃 feature 的进度/卡点/下一步
- 近期架构决定（带日期，可能被未来决定推翻）
- "X 已停 / Y 已废 / Z 已归档" 这类生命周期标注

**判别问句**：「这条内容 6 个月后还会一字不改吗？」
- ✅ 是 → 契约 → CLAUDE.md
- ❌ 否 → 状态 → STATE.md

---

## 3. 命名约定（重要）

**用 `STATE.md`，不要用 `PROGRESS.md`**——理由：
- `STATE.md` = 项目级状态（changelog + 当前 feature 状态 + 近期决定）
- `PROGRESS.md` 留给未来 `.claude/coordination/PROGRESS.md`（跨 worktree/agent 实时看板）
- 两者语义不同：STATE 是项目快照，PROGRESS 是实时进度。一上来叫 PROGRESS.md 会与未来协调层同名冲突。

**预留未来结构**（不在本次创建）：
```
.claude/coordination/
├── PROGRESS.md          # 全局看板（哪个窗口在干啥）
└── plans/feature-XX/
    ├── plan.md          # 任务规格
    └── status.md        # 该窗口当前状态
```

---

## 4. 迁移操作步骤（按顺序）

### Step 1：盘点引用
```powershell
# PowerShell（grep 命令分两路径，brace expansion 不支持）
grep -rn "CLAUDE\.md.*§" .
grep -rn "变更记录\|changelog" .
```
统计 §引用数量、识别哪些是「操作指引」（要改）、哪些是「历史事实陈述」（不改）。

### Step 2：创建 STATE.md（项目根）
模板：
```markdown
# STATE.md — 项目状态 / 变更日志 / 近期决定

> 最后更新：YYYY-MM-DD
>
> **与 CLAUDE.md 的分工**：CLAUDE.md = 不可变契约；本文件 = 会演进的状态。
> **与未来 `.claude/coordination/PROGRESS.md` 的分工**：本文件 = 项目级状态；coordination/PROGRESS.md = 跨 worktree/agent 实时看板（未启用）。

## 1. 变更日志
> 入表标准：[从原 CLAUDE.md changelog 节搬过来 + 反例]
| 日期 | 变更内容 |
|------|---------|
| [迁原表全部行] | ... |
| YYYY-MM-DD | CLAUDE.md 拆三层：项目根新增 STATE.md... |

## 2. 当前 Feature 状态
> 空模板：
> ```
> ### [featureXX-...]
> - 当前阶段：
> - 在跑/卡点：
> - 下一步：
> - 关键文件指针：
> ```

## 3. 近期决定
> 空模板：
> - YYYY-MM-DD ｜ {决定} ｜ {详情链接}
```

### Step 3：瘦身 CLAUDE.md
- 顶部「最后更新」下方加一行：`> 项目状态 / 变更日志 / 近期决定已迁至 [STATE.md](./STATE.md)`
- 删 changelog 整节，替换为：`> 📌 项目状态 / 变更日志 / 近期决定 → 见 [STATE.md](./STATE.md)`
- §1 文件结构图：保留顶层骨架，删每节带日期的演进史长注释，改一行指针「演进详见 [feature*]/plan*.md 与 STATE.md」
- §1 结构图加 STATE.md 条目（与 CLAUDE.md 并列）
- 注意：保留 §0/§1/§2/§3/§4 章节号不变 → 36 处 §0-§4 引用零改动

### Step 4：mc-update.md 同步（如果项目用 mc-update 流程）
若 `.claude/commands/mc-update.md` 硬编码了「CLAUDE.md §N 变更记录表 +1 行」：
- 第 4 步标题：「CLAUDE.md / AGENTS.md 同步检查」→ 「CLAUDE.md / STATE.md 同步检查」（如不维护 AGENTS.md 则删）
- 变更日志步骤：指向 `STATE.md §1 变更日志`
- 结构性变更触发条件精化：「新建**顶级目录** / 新增**工作流场景** / 新增**跨 feature 约定**」（避免 feature 内 schema 升版误入表）
- grep 命令路径分列（PowerShell 不支持 `{a,b}` brace expansion）

### Step 5：引用迁移（关键决策）
**只改「当前操作指引」**，**不改「历史事实陈述」**——
- ✅ 改：`"用 mc-update 更新 CLAUDE.md 变更记录"` → `"用 mc-update 更新 STATE.md §1 变更日志"`
- ✅ 改：fix-*.md 里「`CLAUDE.md §5 变更记录加 1 行`」这类待办指令
- ❌ 不改：「PDF_CATEGORY 与 CLAUDE.md §5 4 大类硬约定四源对齐」这类陈述当时事实的话
- ❌ 不改：debug-*/mc-debug-*.md 等凝固态会话档案（按「凝固态档案不回溯篡改」原则）

判别原则：CLAUDE.md 末尾加的「项目状态/变更日志 → 见 STATE.md」指针为所有 dangling 旧引用提供重定向锚点，所以"§5"在历史档案里的引用仍可解析（语义上指「曾经在 §5、现在在 STATE.md §1 的同一份内容」）。

---

## 5. 验证清单

执行后逐项过：

- [ ] **结构验证**：`ls` 项目根，确认 `STATE.md` 存在；`wc -l CLAUDE.md` 比改前明显下降（典型下降 20-40%）
- [ ] **引用验证 A**：grep `"CLAUDE\.md.*§5\|CLAUDE\.md.*变更记录"`（替换 §N 为原章节号）——残留命中应只在：(a) STATE.md 自己的迁移说明条目；(b) plan/debug 凝固态档案；(c) 本次改造的方案文档自身
- [ ] **引用验证 B**：grep `"STATE\.md"` 应见 5+ 处新引用（CLAUDE.md 顶部 + CLAUDE.md 结构图 + CLAUDE.md 末尾指针 + STATE.md 自身 + mc-update.md）
- [ ] **章节稳定性**：grep `"CLAUDE\.md.*§[0-4]"` 数量应等于改前（章节号兼容，零改动）
- [ ] **mc-update 子检查**（如有该流程）：(a) 第 4 步标题已去 AGENTS.md（或保留正确）；(b) 变更日志步骤指向 `STATE.md §1`；(c) 结构性变更触发条件已精化；(d) grep 命令路径分列
- [ ] **AGENTS.md 残留**（如已剥离 codex 维护）：grep `"AGENTS\.md"` 在 mc-update.md 应 0 命中
- [ ] **冒烟读**：新开会话加载 CLAUDE.md，确认无残留状态描述、目录树注释干净

---

## 6. 常见坑

1. **全角 vs 半角标点**：中文 CLAUDE.md 里「上限，」是全角逗号，Edit 工具的 old_string 写成 ASCII 逗号会匹配失败 → 先 Read 确切字节，再粘贴
2. **直接 Set-Content/Out-File 覆盖 CLAUDE.md** 会被 Claude Code auto-mode 拦截（in-place destruction not created in session）→ 改用 Edit 工具逐段替换
3. **凝固态档案误改**：plan-*/debug-*/mc-debug-*.md 是历史会话归档，里面的「§5 4 大类硬约定」陈述当时事实，**不要改成 STATE.md** → 会引入虚假回溯
4. **changelog 入表标准漏迁**：原 §5 表头那段「入表标准 + 反例」必须连表一起迁到 STATE.md §1，否则下次 mc-update 时不知道什么该入表
5. **§1 文件结构图 STATE.md 漏加节点**：拆完忘了把 STATE.md 加进结构图，下次别的 agent 不知道这文件存在
6. **预留 PROGRESS.md 命名误用**：不要一开始就把状态文件叫 PROGRESS.md，未来 coordination 层会与之冲突

---

## 7. 落地参考

本 repo 实战留存：
- 方案文档：`[feature02-方彩霞ppt整理]/plan07.md`
- 改后契约：`CLAUDE.md`（127 行，原 160 行）
- 状态文件：`STATE.md`（71 行）
- 流程同步：`.claude/commands/mc-update.md`（P0+P1+P2 三档清理）
- 引用迁移：`[feature02-方彩霞ppt整理]/fix-pre-archive-auto-classify.md` L10/L171、`plan05.md` L372
