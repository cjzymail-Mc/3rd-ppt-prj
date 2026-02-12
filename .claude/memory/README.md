# Agent Memory Files（Agent 记忆文件）

> **重要**：本目录通过 Git 同步，换电脑后仍然有效

---

## 📁 文件结构

```
.claude/memory/
├── README.md                 # 本文件（说明文档）
├── ARCHITECT_RULES.md        # ✅ 已创建（230 行）
├── TECH_LEAD_RULES.md        # ⚠️ 待创建
├── DEVELOPER_RULES.md        # ⚠️ 待创建
├── TESTER_RULES.md           # ⚠️ 待创建
├── OPTIMIZER_RULES.md        # ⚠️ 待创建
└── SECURITY_RULES.md         # ⚠️ 待创建
```

---

## 🎯 设计理念

### 为什么需要 RULES 文件？

**问题**：所有 agents 都会犯重复性错误
- Developer: 路径问题、中文符号、安全漏洞
- Tester: 只测理想情况、遗漏边界测试
- Tech Lead: 过度修改 PLAN.md
- Optimizer: 过度优化
- Security: 误报过多

**解决方案**：为每个 agent 创建独立的 RULES 文件，记录反复出现的错误和经验教训

---

## 📋 RULES 文件模板

每个 RULES 文件应包含：

### 1. 反复出现的错误

```markdown
### 错误1：{错误描述} ⭐⭐⭐

**问题描述**：{具体问题}

**发生频率**：⚠️ 高频/中频/低频

**根因**：{为什么会犯这个错误}

**❌ 错误示范**：
{具体的错误代码/操作}

**✅ 正确示范**：
{正确的代码/操作}

**规则**：
- ✅ {必须做的事}
- ❌ {绝对不能做的事}
```

### 2. 关键发现与最佳实践

```markdown
### 发现1：{洞察标题}

**案例**：{具体案例}

**原因**：{分析}

**教训**：{可复用的经验}
```

### 3. 检查清单

```markdown
## 📋 每次任务前检查清单

开始任务前，必须确认：
- [ ] {检查项1}
- [ ] {检查项2}

任务结束前，必须确认：
- [ ] {检查项3}
- [ ] {检查项4}
```

---

## 🔄 更新机制

### 何时更新？

- ✅ 发现新的重复性错误
- ✅ 用户反馈问题
- ✅ 修复 bug 后总结根因

### 如何更新？

1. 当前 agent 在 `claude-progress.md` 中详细记录错误
2. 用户查阅后，决定是否值得记入 RULES
3. 用户手动更新对应的 RULES 文件（或请 agent 更新）

### 更新流程

```
Agent 工作 → 犯错 → 记录到 progress.md
  ↓
用户查阅 progress.md → 判断是否重复性错误
  ↓
如果是 → 更新到 {AGENT}_RULES.md
  ↓
Git commit → 团队共享
  ↓
下次运行 → Agent 读取 RULES → 避免重复犯错
```

---

## 📊 当前状态

| Agent | RULES 文件 | 状态 | 错误记录数 | 最后更新 |
|-------|-----------|------|-----------|---------|
| **Architect** | ARCHITECT_RULES.md | ✅ 已创建 | 4 个 | 2026-02-06 |
| Tech Lead | TECH_LEAD_RULES.md | ⚠️ 待创建 | - | - |
| Developer | DEVELOPER_RULES.md | ⚠️ 待创建 | - | - |
| Tester | TESTER_RULES.md | ⚠️ 待创建 | - | - |
| Optimizer | OPTIMIZER_RULES.md | ⚠️ 待创建 | - | - |
| Security | SECURITY_RULES.md | ⚠️ 待创建 | - | - |

---

## 🎯 下一步行动

1. **逐步创建其他 RULES 文件**：
   - 当对应的 agent 开始工作时，创建其 RULES 文件
   - 不需要一次性全部创建（空文件没有意义）

2. **持续维护**：
   - 每次 agent 犯重复性错误时，及时记录
   - 定期审查和更新 RULES 文件

3. **团队协作**：
   - 通过 Git 同步，确保所有成员看到最新的经验教训
   - 换电脑后，RULES 文件自动生效

---

**创建时间**：2026-02-06
**最后更新**：2026-02-06
