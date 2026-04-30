很好，你非常顺利地完成了任务。接下来你更新相关文档、总结相关经验（更新范围包括：CLAUDE.md 和其他此次更新涉及到的修改）
你首先列出需要更新的文件，按重要性排序，展示给我。我会告诉你更新哪些

---

## auto-memory vs memory 的知识边界规则（**每次更新都必须遵守**）

涉及 memory 文件的写入/更新时，必须先判断这条知识的归属层，再决定写到哪个目录：

### 判断标准

**核心问句**：没有这条知识，我在新会话的前 2-3 个回复会做错决策吗？

| 答案 | 归属 | 路径 |
|--|--|--|
| ✅ 是（影响每次对话的响应风格 / 路由决策 / 反射动作） | **auto-memory**（每次会话自动加载索引） | `.claude/auto-memory/` |
| ❌ 否（仅特定任务触发；技术细节；一次性踩坑；架构决策档案） | **memory**（按需加载，通过 CLAUDE.md / MEMORY.md 索引发现） | `.claude/memory/` |

### 典型分类

**应进 auto-memory**：
- 用户画像 / 角色 / 协作偏好（每次响应都用得上）
- 跨主题的元规则（"稳定性优先 over 新颖性"等）
- 高频反射动作（已与 CLAUDE.md Section 0/2/3 deduped 后的精简核心）

**应进 memory**：
- 特定模板 / 特定文件的工作流知识
- 一次性踩坑（"某次 fix 解决了某具体问题"）
- 架构决策档案（"YYYY-MM-DD 决定改用 X 路线"）
- 历史优化记录（已落实到代码、归档性质）
- 技术细节对照表（COM API 用法、字段含义等）

### 冗余规则（适度冗余）

- **P0-P1 级关键反射动作**：可以在 CLAUDE.md + auto-memory 两处冗余保留（双保险）
- **其他普通知识**：单一存放即可，不要冗余架构

### 每次更新动作清单

每次写新 memory / 移动 memory 时按这个顺序处理：

1. **判断归属**：用上面"核心问句"判断 auto-memory 还是 memory
2. **检查重叠**：grep 已有 memory 文件，避免和现存内容重复（如有重叠，append 进现有文件，不另起新文件）
3. **写文件**：frontmatter 含 `name` / `description` / `type`
4. **更新索引**：相应目录的 `MEMORY.md` 加一行（auto-memory 索引尽量精简，因为它每次会话都进 system prompt）
5. **更新 CLAUDE.md 指针**（如果有引用此文件的硬规则）：路径写对（`.claude/auto-memory/` vs `.claude/memory/`）
6. **审查 dangling**：如果挪了 / 删了文件，检查所有 MEMORY.md 索引和 CLAUDE.md 指针是否还在引用旧路径

### 反例（不要这样做）

- ❌ 把所有踩坑记录都塞 auto-memory（每次会话 system prompt 拥挤，索引信噪比下降）
- ❌ 用户偏好放 memory（按需加载导致前 2-3 回复风格不对）
- ❌ 写新 memory 时不查 grep 已有同主题文件（产生重复 / 内容碎片化）
- ❌ 删除 / 移动 memory 后不更新索引（产生 dangling 引用）

### 参考实现

本项目 2026-04-29 完成了一次 auto-memory ↔ memory 的边界整理：12 条 auto-memory 压到 3 条（仅 user_profile / stability / check_skills_first），其余 8 条按上述标准挪到 memory 或合并到 com_constraints。可 `git log --grep "memory junction"` 查相关 commit 看实战案例。