任务完成后的文档更新流程。

## 第 0 步：要不要固化？

**核心问句**：这条经验 1 年内会被反复需要 ≥ 3 次吗？

- ✅ 是 → 继续
- ❌ 否（一次性反思 / git log 已是真相）→ **不写**，本次结束

任务完成 ≠ memory 必有更新。很多时候 git log + Mc-debug-N.md 已是真相，强行固化只污染索引层。

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

1. grep 已有同主题文件，重叠 → append 现有文件，不另起
2. 写文件（frontmatter：`name` / `description` / `type`）
3. 更新 MEMORY.md 索引（auto-memory 索引尤其精简，每会话都加载）
4. **CLAUDE.md / AGENTS.md 同步检查**（必做，不是"如有"）：
   - 4a 既有指针：grep 是否对得上新 memory 路径（auto-memory vs memory 路径不要错）
   - 4b 结构性变更：本次任务有没有新建顶级目录 / 新增 metadata schema / 新工作流场景？→ 文件结构图加节点
   - 4c 命令表：本次有没有新增 `/role-*` 或 slash command？→ 命令表 +1 行
   - 4d 变更记录：4a-4c 任一触发 → 「## 4. 变更记录」表 +1 行（日期 + 简述）
5. 挪/删文件后，grep dangling 引用并修复（凝固态档案除外：debug-*/plan-* 不回溯篡改）

## 反例

- ❌ 任务完成必写 memory（git log 已是真相）
- ❌ 踩坑都塞 auto-memory（每次会话 system prompt 拥挤）
- ❌ 用户偏好放 memory（按需加载导致前 2-3 回复风格不对）
- ❌ 不 grep 同主题就写新 memory（产生重复 / 碎片）
- ❌ 删/移文件后不更新索引（产生 dangling 引用）
- ❌ meta/process 改动塞进 fix*.md（拉低 fix 文档信噪比）
- ❌ 新建顶级目录或文件类型，但忘改 CLAUDE.md 文件结构图（下次会话拿到过时的项目地图）
- ❌ 把 "如有 CLAUDE.md 指针就更新" 误读为 "没有就跳过"（实际：常常是该新增一条）
