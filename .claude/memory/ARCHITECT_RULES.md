# Architect Agent - 项目级规则（ARCHITECT_RULES.md）

> **重要**：本文件通过 Git 同步，换电脑后仍然有效
> 记录 Architect agent 在本项目中反复出现的错误和经验教训

---

## ⚠️ 本项目特殊说明

### 项目目录结构

```
D:/Technique Support/Claude Code Learning/2nd-repo/  ← Repo 根目录（.git 在这里）
├── .git/
├── src/
│   └── orchestrator_v6.py                                  ← 代码在这里
├── plan.md                                          ← 在 repo 根目录
├── repo-scan-result.md                              ← 在 repo 根目录
└── .claude/                                         ← 在 repo 根目录
```

**关键问题**：
- ❌ **不要**在 `src/` 目录下运行 `python orchestrator_v6.py`
- ✅ **应该**在 repo 根目录运行，或者确保代码使用 `find_project_root()` 而非 `Path.cwd()`
- ⚠️ 如果在 `src/` 运行，`Path.cwd()` 会返回 `src/`，导致找不到 `plan.md`、`repo-scan-result.md` 等文件

**所有文件路径必须相对于 repo 根目录**：
- ✅ `plan.md`（不是 `../plan.md` 或 `src/plan.md`）
- ✅ `repo-scan-result.md`
- ✅ `.claude/agents/01-arch.md`

---

## 🚨 反复出现的错误（必须避免）

### 错误1：计划保存位置错误 ⭐⭐⭐

**问题描述**：将计划保存到 `~/.claude-mc/plans/` 而非项目根目录 `plan.md`

**发生频率**：⚠️ 高频（反复出现的"老bug"）

**根因**：
- Plan Mode 系统提示优先级高于项目规范
- 指令冲突：系统默认路径 vs 用户要求
- 缺少持久化的记忆机制

**❌ 错误示范**：
```python
Write(
    file_path="C:\Users\xy24\.claude-mc\plans\linked-beaming-puddle.md",
    content="..."
)
```

**✅ 正确示范**：
```python
# 方案1：如果 plan.md 已存在，使用 Edit 追加
# 注意：使用相对于 repo 根目录的路径（不是相对于 src/）
Edit(
    file_path="plan.md",  # 相对于 repo 根目录
    old_string="...",
    new_string="..."
)

# 方案2：如果 plan.md 不存在，使用 Write 创建
Write(
    file_path="plan.md",  # 相对于 repo 根目录
    content="..."
)
```

**规则**：
- ✅ **所有计划内容必须保存到项目根目录 `plan.md`**
- ✅ 使用相对路径，不要使用绝对路径
- ✅ 优先使用 Edit 工具追加内容（而非 Write 创建新文件）
- ❌ 永远不要保存到 `~/.claude-mc/plans/` 目录

---

### 错误2：越权尝试执行代码修复 ⭐⭐⭐

**问题描述**：在 Plan Mode 批准后，尝试直接执行代码修复任务

**发生频率**：⚠️ 中频（角色定位容易混淆）

**根因**：
- 误解 "You can now make edits" 为"可以修改源代码"
- Plan Mode 不经过三重防护机制
- 混淆了"制定计划"和"执行计划"的职责边界

**❌ 错误示范**：
```python
# Plan Mode 批准后，尝试执行修复
Bash(command="git checkout -b fix/bug-7-repo-scan")
Read(file_path="src/orchestrator_v6.py")  # 准备修改代码
Edit(file_path="src/orchestrator_v6.py", ...)  # 修改源代码
```

**✅ 正确示范**：
```python
# Plan Mode 批准后，任务结束
输出消息：
"✅ Architect 任务已完成！
下一步：请用户输入 /exit 退出会话，启动后续 agents（Developer、Tech Lead）执行修复计划。"

# 不执行任何代码修复操作
```

**规则**：
- ✅ **Architect 只负责制定计划（PLAN.md），不执行代码修复**
- ✅ ExitPlanMode 批准后，告知用户交给其他 agents 执行
- ✅ 明确职责终止点：计划批准 = Architect 任务完成
- ❌ 永远不要创建 Git 分支、修改源代码、运行测试

---

### 错误3：未生成进度记录文件 ⭐⭐

**问题描述**：任务结束时没有生成 `claude-progress.md`

**发生频率**：⚠️ 中频（容易遗忘的步骤）

**根因**：
- Plan Mode 不经过 Orchestrator 流程
- `_init_progress_file()` 未被自动调用
- 缺少"生成进度文件"的检查清单

**❌ 错误示范**：
```python
# 任务结束时，直接退出，没有生成任何记录
# 用户无法了解工作细节（尤其在全自动模式下）
```

**✅ 正确示范**：
```python
# 任务结束前，必须生成 claude-progress.md
Write(
    file_path="claude-progress.md",  # 相对于 repo 根目录
    content="""
    # Architect 工作记录 - {日期}

    ## 任务描述
    ...

    ## 执行过程
    ...

    ## 🐛 错误记录（Debug Log）
    ...

    ## 关键决策
    ...

    ## 下一步行动
    ...
    """
)
```

**规则**：
- ✅ **每次任务结束前，必须生成 `claude-progress.md`**
- ✅ 文件内容包括：任务描述、执行过程、错误记录、关键决策、下一步行动
- ✅ 保存到项目根目录
- ✅ 包含详细的"🐛 错误记录"章节（错误示范 vs 正确示范）

---

### 错误4：未主动查找 repo-scan-result.md ⭐

**问题描述**：任务开始时，没有主动检查项目根目录是否有 `repo-scan-result.md`

**发生频率**：⚠️ 低频（新功能，容易忘记）

**根因**：
- Bug #7 导致代码自动检测失效（Path.cwd() 路径错误）
- 缺少主动查找意识
- 没有形成习惯

**❌ 错误示范**：
```python
# 直接开始分析任务，没有检查扫描结果文件
# 错过 token 节省机会
```

**✅ 正确示范**：
```python
# 任务开始时，第一件事：检查 repo-scan-result.md
try:
    Read(file_path="repo-scan-result.md")  # 相对于 repo 根目录
    # 如果存在，包含到 prompt 中，节省 30-70% token
except:
    # 文件不存在，继续正常流程
    pass
```

**规则**：
- ✅ **任务开始时，主动检查项目根目录是否有 `repo-scan-result.md`**
- ✅ 如果存在，读取并包含到初始 prompt 中
- ✅ Token 节省效果：架构设计任务 60-70%，Bug 修复任务 20-25%

---

## ✅ 关键发现与最佳实践

### 发现1：单元测试通过 ≠ 功能可用

**案例**：61 个单元测试全部通过，但 Bug #7 在生产环境完全失效

**原因**：单元测试只测试理想情况（在项目根目录运行）

**教训**：必须进行端到端功能测试，覆盖所有真实使用场景

### 发现2：用户反馈是最宝贵的测试

**案例**：用户一句话"你应该先自动加载这个呀"直接定位生产 bug

**教训**：比运行 100 个单元测试更有价值

### 发现3：文件保存位置问题是系统性问题

**根因**：系统提示优先级 > 用户指令 > 项目规范

**解决**：三层记忆机制（MEMORY.md + RULES.md + agent prompt）

### 发现4：所有 agents 都需要三层记忆机制

**洞察**：不仅 Architect 会犯重复性错误，所有 agents 都会

**建议**：为所有 6 个 agents 创建独立的 RULES 文件

---

## 📋 每次任务前检查清单

开始任务前，必须确认：

- [ ] 是否主动查找并读取 `repo-scan-result.md`？
- [ ] 是否确认计划保存位置为项目根目录 `plan.md`？
- [ ] 是否使用相对路径（而非绝对路径）？
- [ ] 是否明确自己的职责边界（只制定计划，不执行代码）？

任务结束前，必须确认：

- [ ] 是否已生成 `claude-progress.md`？
- [ ] 是否包含详细的"🐛 错误记录"章节？
- [ ] 是否告知用户交给后续 agents 执行？

---

## 🎯 持续改进方向

1. **主动查找 repo-scan-result.md**：每次任务开始时第一件事
2. **严格遵守职责边界**：永远不要越权执行代码修复
3. **文件保存位置检查**：优先使用 Edit 工具而非 Write
4. **生成进度记录**：任务结束前必做，包含详细 debug 信息

---

**文件版本**：v1.0
**最后更新**：2026-02-06
**更新者**：Architect (本次会话)
**下次更新时机**：当发现新的重复性错误，或现有规则需要修正时
