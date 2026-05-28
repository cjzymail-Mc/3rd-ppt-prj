任务完成后的文档更新流程。

## 第 0 步：要不要固化？

> 本流程是固化记忆的**唯一入口**（CLAUDE.md §0.2）：只在任务完成后走，任务途中一律不碰
> `.claude/{auto-memory,memory}/`、任何 `MEMORY.md`、harness memory 目录。途中发现的可固化
> 经验只在收尾口头提示「建议 mc-update 固化 X」，由用户在本流程里拍板。

**两道闸门，全过才写**：

1. **频次闸**：这条经验 1 年内会被反复需要 ≥ 3 次吗？
   - ❌ 否（一次性反思 / git log 已是真相）→ **不写**，本次结束
2. **去重闸**：先 grep `.claude/{auto-memory,memory}/`，有没有同主题条目已覆盖？
   - ✅ 已被现有条目完整覆盖 → **不写**（哪怕够 ≥3 次——再写只是近似重复条目，徒增索引拥挤）
   - 仅部分覆盖 → 走 append（第 3 步），**不另起新条目**
   - 完全没有 → 才允许新建

两道都过才继续。任务完成 ≠ memory 必有更新——很多时候 git log + Mc-debug-N.md 已是真相，强行固化只污染索引层。

## 第 1 步：列清单等用户审核

**列出候选文件，按重要性排序展示，等用户挑选。** 不要直接动手。

## 第 2 步：归属判断

**memory 归属**（核心问句：没这条知识，新会话前 2-3 回复会出错吗？）：

| 答案 | 归属 | 路径 |
|--|--|--|
| ✅ 用户偏好 / 跨主题元规则 / 高频反射 | auto-memory | `.claude/auto-memory/` |
| ❌ 特定任务 / 技术细节 / 一次性踩坑 / 架构档案 | memory | `.claude/memory/` |

**改动归属**：

| 改动类型 | 写到 |
|--|--|
| 技术 bug + 代码修复 + 实证 | `[feature*]/fix*.md` |
| 跨主题讨论 / 架构决策 / 架构精简 | `[feature*]/Mc-debug-*.md` 续写 |
| 反复用得上的元规则 / 高频反射 | `.claude/memory/feedback_*.md` |
| 无新规则可提炼（meta / process 改动） | 啥都不写（git log 即真相） |

## 第 3 步：执行（仅对用户挑中的候选）

1. 按第 0 步去重闸结论落位：部分覆盖 → append 现有文件；完全没有 → 才新建条目
2. 写文件（frontmatter：`name` / `description` / `type`）
3. 更新 MEMORY.md 索引（auto-memory 索引尤其精简，每会话都加载）
4. **CLAUDE.md / STATE.md 同步检查**（必做，不是"如有"）：
   - 4a 既有指针：grep 是否对得上新 memory 路径（auto-memory vs memory 路径不要错）
   - 4b 结构性变更（**严判**：仅以下三类触发）：新建**顶级目录** / 新增**工作流场景** / 新增**跨 feature 约定** → CLAUDE.md §6 文件结构/详情索引加节点
     - 反例（**不触发**）：feature 内 schema 升版、内部脚本重命名、单个 memory 条目新增
   - 4c 命令表：本次有没有新增 `/role-*` 或 slash command？→ 命令表 +1 行
   - 4d 变更记录：4a-4c 任一触发 → 在 `STATE.md §1 变更日志` 表 +1 行（日期 + 简述）。**不是** CLAUDE.md（CLAUDE.md 是契约层，不维护 changelog；STATE.md 是状态层）
5. 挪/删文件后，grep dangling 引用并修复（凝固态档案除外：debug-*/plan-*/fix-* 不回溯篡改；CLAUDE.md 末尾已加重定向锚点为旧 §引用兜底）

## 反例

- ❌ 任务完成必写 memory（git log 已是真相）
- ❌ 任务途中主动写 / 改记忆文件（CLAUDE.md §0.2 禁止；只在收尾走本流程）
- ❌ 踩坑都塞 auto-memory（每次会话 system prompt 拥挤）
- ❌ 用户偏好放 memory（按需加载导致前 2-3 回复风格不对）
- ❌ 不 grep 同主题就写新 memory（产生重复 / 碎片）；同主题已有条目还新建近似条目（该 append 就 append）
- ❌ 删/移文件后不更新索引（产生 dangling 引用）
- ❌ meta/process 改动塞进 fix*.md（拉低 fix 文档信噪比）
- ❌ 新建顶级目录或文件类型，但忘改 CLAUDE.md 文件结构图（下次会话拿到过时的项目地图）
- ❌ 把 "如有 CLAUDE.md 指针就更新" 误读为 "没有就跳过"（实际：常常是该新增一条）
- ❌ 把 changelog 行写进 CLAUDE.md（契约层不维护变更日志 → 写到 `STATE.md §1`）
