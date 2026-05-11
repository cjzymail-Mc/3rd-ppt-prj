# Plan: 3 账号 auto-memory junction 化 + CLAUDE.md 压缩

## Context

**问题**：用户在 Windows 用 3 个 Claude Pro 账号（mc / yk / xh）轮换工作于同一项目（PPT pipeline）。Claude Code 的 auto-memory 系统把会话洞察自动写入 `<账号根>/projects/<项目>/memory/`，3 个账号的根目录物理隔离 → 切账号就"失忆"另一边的最新洞察。

**漂移现状（2026-04-29 实测）**：
- 账号根：mc=`C:\Users\xy24\.claude-mc\`，yk=`C:\Users\xy24\.claude\`，xh=`C:\Users\xy24\.claude-xh\`
- 8 个 feedback/project 文件 3 账号一致（旧的、共有的）
- `MEMORY.md` 严格 subset：mc(8 条) ⊂ yk(9 条) ⊂ xh(10 条)，但 yk 独有 1 条 + xh 独有 2 条互不交叉，**没有任何账号有完整 11 条 union**
- 3 个 unique 文件：`feedback_check_skills_first.md`(yk-only)、`feedback_skip_vs_clear.md`(xh-only)、`feedback_unit_normalize_bmi.md`(xh-only)
- `claude_migrate.py` 是覆盖式整目录迁移（line 110 `shutil.rmtree(tgt_project)`），用它做 mc→yk 会丢 yk 独有那条；不能解决 union 问题

**第二个问题**：repo 内 `.claude/CLAUDE.md`（146 行）的 Section 3「硬规则」19 条 bullet 占近半篇幅，多条已写"详见 `.claude/memory/xxx.md`"——内容在两边重复存。这两个问题相关：CLAUDE.md 想用 memory 文件作单一权威源，前提是 memory 在所有账号都存在。

**目标**：
1. 用 NTFS junction 把 3 个账号 auto-memory 目录物理合一到 repo 内 `.claude/auto-memory/`，进 git，永久杜绝漂移
2. 压缩 CLAUDE.md 到 ~70 行，硬规则详情下沉到 memory 单一源
3. 一次 commit 全做完，跨账号验证一次到位

## Junction 工作原理（架构理解）

**核心机制**：NTFS junction 是 OS 文件系统层的目录别名（reparse point）。3 个账号的 `memory\` 入口在 Windows 看来是 3 个独立目录，但实际上都被透明重定向到 D 盘 repo 内的同一个物理目录。

```
Claude Code 写入路径（它认为的）            实际物理路径（OS 重定向后）
┌──────────────────────────────────┐
│ C:\Users\xy24\.claude-mc\        │ ──┐
│   projects\<proj>\memory\        │   │
├──────────────────────────────────┤   │   ┌──────────────────────────────┐
│ C:\Users\xy24\.claude\           │ ──┼─→ │ D:\...\3rd-ppt-prj\          │
│   projects\<proj>\memory\        │   │   │   .claude\auto-memory\       │
├──────────────────────────────────┤   │   │   ├── MEMORY.md      ←本体   │
│ C:\Users\xy24\.claude-xh\        │ ──┘   │   ├── feedback_*.md  ←本体   │
│   projects\<proj>\memory\        │       │   └── ...                    │
└──────────────────────────────────┘       └──────────────────────────────┘
```

**关键事实**：

| 维度 | C: 上（账号侧入口） | D: 上（repo 侧本体） |
|--|--|--|
| 目录入口 | ✅ 一条 reparse point（~100 字节） | ✅ 真实目录 |
| 文件本体（.md 内容） | ❌ 完全不存在 | ✅ 唯一物理副本 |
| 磁盘占用 | 3 × 几百字节（仅 3 条入口） | 11 个 .md 文件总和 |
| Claude Code 感知 | 无（看到的是普通目录） | 无（看到的是普通目录） |
| inode / mtime | 由 reparse point 转发 | 真实归属 |

**为什么不需要改 Claude Code 默认行为**：OS 在 `CreateFile` / `WriteFile` 调用层就完成透明重定向，Claude Code 完全无感知。它仍然按 hardcoded 路径 `<account>/projects/<proj>/memory/` 读写，但实际命中 repo 内的物理文件——3 个账号同时指向同一物理本体，自动同步。

## 未来维护警告（关键安全规则）

Junction 引入了几个反直觉的危险操作，**必须**记入项目永久知识：

| 操作 | 行为 | 安全性 |
|--|--|--|
| `os.rmdir(memory_path)` 或 `rmdir <path>`（无 `/s`） | 只删 reparse point 入口，本体安全 | ✅ 安全（rollback 用） |
| `shutil.rmtree(memory_path)` 或 `rmdir /s <path>` | **穿透 junction 杀 D 盘本体** | ⚠️ 危险（这是 `claude_migrate.py` line 110 必须改的根因） |
| `del <path>\*` 或 `rm -rf <path>/*` | 同上，杀本体 | ⚠️ 危险 |
| 删除账号根目录 `C:\Users\xy24\.claude-mc\` 等 | 递归穿透，**直接杀 D 盘本体** | ⚠️ 必须先解 junction 再删账号 |
| 文件管理器拖拽删除 `memory` 文件夹 | 走回收站，不穿透 | ✅ 安全（回收站可恢复） |
| 第三方备份工具（File History 等） | 默认不进入 reparse point | ✅ 安全（不会重复备份 3 份） |

**未来要彻底删除某个账号时的正确流程**：
1. 先跑 `python tools/memory_junction_rollback.py --account mc` 解 junction
2. 此时 mc 账号下 `memory\` 已经是空目录或恢复了备份内容（不再是 reparse point）
3. 再删账号根目录 `C:\Users\xy24\.claude-mc\`

这些安全规则必须同步落到 `.claude/memory/reference_3account_junction.md`，作为未来唯一的查阅入口。

## 推荐方案

### 阶段 A：Union merge + Junction 化

**新建 4 个工具脚本**：

1. **`tools/memory_union_merge.py`**（CLI，独立非 GUI）
   - `--dry-run`（默认）：扫描 3 账号 auto-memory，输出三栏决策表 `文件名 | 来源 | 决策`，让用户人工 review
   - `--apply`：执行 union merge，写入 `D:\...\3rd-ppt-prj\.claude\auto-memory\`
   - **MEMORY.md 特殊处理**：解析 3 版的 `- [...](file.md) — ...` bullet 行，按 `(file.md)` key 求 union → 11 行写出（mc-yk-xh 共有 8 条 + yk 独 1 条 + xh 独 2 条）
   - 其他 .md 文件：按文件名 union；同名文件取最新 mtime 版本（实测仅 `MEMORY.md` 跨账号同名内容不同，其他重名都已 mtime 一致）
   - **复用** `claude_migrate.py:detect_project_name()` (line 33) 和 `ACCOUNTS` (line 14)

2. **`tools/memory_junction_setup.py`**（CLI，一次性）
   - 校验 `.claude/auto-memory/` 已有 11 文件
   - 备份：3 账号原 `memory/` 目录复制到 `.claude/auto-memory/.pre-junction-backup/{mc,yk,xh}-{timestamp}/`
   - `shutil.rmtree` 删 3 账号原 `memory/` 目录
   - `subprocess.run(["cmd", "/c", "mklink", "/J", account_path, repo_target], check=True)` 建 junction（`mklink /J` 创建 NTFS 目录联接，**不需要管理员权限**，跨盘可用）
   - 验证：`os.stat(path).st_file_attributes & stat.FILE_ATTRIBUTE_REPARSE_POINT` 为非零

3. **`tools/memory_junction_rollback.py`**（保险）
   - `os.rmdir(account_memory_path)` 删除 junction（**绝对不能用 `shutil.rmtree`**——会跟进 junction 删 repo 真实文件）
   - 从 `.pre-junction-backup` 用 `shutil.copytree` 恢复

4. **`tools/memory_junction_verify.py`**（端到端测试，约 30 行）
   - 在 mc 账号 memory 目录写 `_test_<timestamp>.md` → 验证 yk、xh、`.claude/auto-memory/` 四处都能 `os.path.exists`
   - 删除 → 验证四处同步消失

**修改 `claude_migrate.py`**（最小改动，约 20 行）：
- 新增辅助函数 `is_junction(path)` 用 reparse point 属性检测
- 在 `run_migration()` line 109-122 间插入 junction 检测：复制目标 `projects/<project>/memory` 子目录前先检查 source/target 任一为 junction → 跳过该子目录的 `rmtree` 和 `copytree`，避免穿透 junction 破坏 repo
- 在 `_show_confirm()` line 312 的"迁移文件"列表加注：`(memory 子目录已 junction，自动跳过)`

**新建 `.claude/memory/reference_3account_junction.md`**（手工 curator 层 memory）：
- 3 账号路径表
- Junction 工作原理 + 拓扑图（C: 上仅 reparse point 入口，D: 上是文件本体）
- 4 个工具脚本用法
- 自检命令：`fsutil reparsepoint query "<path>"` 报告 `Mount Point` 即生效
- **未来维护警告表**（rmtree 穿透 vs rmdir 安全；删账号目录会杀 D 盘本体；正确卸载流程）
- **理由**：未来 junction 失效（重装系统、目录移动、误删账号）时第一反应应该是查这份文档；新机器重建账号时必读

**修改 `.gitignore`**：追加 `.claude/auto-memory/.pre-junction-backup/` 和 `.claude/auto-memory/_test_*.md`

### 阶段 B：CLAUDE.md 压缩（与 A 同 commit）

**Section 3 处理规则**：

| 处理 | 数量 | 规则 |
|--|--|--|
| **完全删除** | 2 条 | 第 89 行 yzr/zxh 共享工具（`_ppt_shared.py` 已落地）、第 94 行 bar chart max+1（`Function_030.py` 已落地） |
| **压缩到 1 行 + 指针** | 13 条 | `(YYYY-MM) 触发场景 + 一句话结论 → memory文件名` 三件套；fix4/fix5/Shapes.Paste()/tk popup HWND/chart Copy-Delete/GPT 风格锚等 |
| **新增 3 行指针** | 3 条 | 指向 `.claude/auto-memory/feedback_{check_skills_first, skip_vs_clear, unit_normalize_bmi}.md`（junction 后所有账号都能读到） |
| **保留不动** | Section 0/1/2/4/5/6 | 路由表 + 高频铁律，每次新对话主 Claude 必读 |

预计 146 行 → 70 行左右。

### 阶段 C（不在本轮做）

把 `.claude/auto-memory/` 中价值高的洞察 promote 到 `.claude/memory/` 项目层。**不强制做**——auto-memory 进 git 后两层差异主要是"自动 vs 手工"，不必急于消除。

## 执行依赖图

```
[准备] 关闭所有 3 个 Claude Code 会话（避免 auto-memory 文件锁）
   ↓
[A1] memory_union_merge.py --dry-run     → 用户 review 三栏决策表
   ↓
[A2] memory_union_merge.py --apply        → .claude/auto-memory/ 11 文件
   ↓
[A3] memory_junction_setup.py             → 备份 + rmtree + mklink
   ↓
[A4] memory_junction_verify.py            → 端到端验证四处同步
   ↓                                      ↓
[A5] 改 claude_migrate.py              [B] CLAUDE.md 压缩
   ↓                                      ↓
[A6] 写 reference_3account_junction.md
   ↓
[A7+B] git add + commit（一次性）
```

A1→A2→A3→A4 严格串行；A5/B 可并行；A6 在 A4 完成后做（要记录最终状态）；A7+B 一次 commit。

## 关键文件路径

| 文件 | 动作 | 说明 |
|--|--|--|
| `tools/memory_union_merge.py` | 新建 | 独立 CLI |
| `tools/memory_junction_setup.py` | 新建 | 一次性 |
| `tools/memory_junction_rollback.py` | 新建 | 保险 |
| `tools/memory_junction_verify.py` | 新建 | 测试 |
| `.claude/auto-memory/` | 新建（目录） | 11 union 文件 |
| `.claude/memory/reference_3account_junction.md` | 新建 | 架构 reference |
| `.claude/CLAUDE.md` | 编辑 | 压缩 146→~70 行 |
| `claude_migrate.py` | 编辑 | line 109+312 |
| `.gitignore` | 编辑 | 排除备份+测试 |

## 复用现有代码

- `claude_migrate.py:ACCOUNTS` (line 14)：3 账号路径配置 → 4 个新工具脚本复用
- `claude_migrate.py:detect_project_name()` (line 33)：项目名探测 → 复用
- `claude_migrate.py:next_backup_number()` (line 62)：备份编号生成 → `memory_junction_setup.py` 的 `.pre-junction-backup` 复用思路（按 `pre-junction#001` 编号）

## 风险与回滚

| 风险 | 检测 | 回滚 |
|--|--|--|
| Junction 创建失败（目标盘非 NTFS / 路径已存在） | A3 脚本 `os.stat` 验证失败 | 跳过后续，从 `.pre-junction-backup` 恢复账号目录 |
| Union merge 写错 | A1 dry-run 阶段三栏表暴露 | A2 只写新目录，不动账号原文件，删 `.claude/auto-memory/` 重跑 |
| `claude_migrate.py` A5 未改完先跑迁移（穿透 junction 删 repo） | 必须人工 verify A5 完成才允许跑迁移 | git 历史恢复（前提：`.claude/auto-memory/` 已 commit） |
| 会话中文件锁（Claude Code 在改 auto-memory） | 执行前关闭所有 3 个 Claude Code 会话 | — |
| CLAUDE.md 压缩丢信息 | B 阶段验证脚本 grep 删除行 → memory 文件 | git 历史恢复 |

## 验证步骤

1. **A2 验证**：`ls .claude/auto-memory/` 出现 11 文件；`MEMORY.md` 含 11 行 union 索引
2. **A3 验证**：3 处 `fsutil reparsepoint query "C:\Users\xy24\.claude{-mc,,-xh}\projects\D--...\memory"` 都报告 `Reparse Tag Value: 0xa0000003` 或 `Mount Point`
3. **A4 自动测试**：在 mc 账号 memory 写时间戳文件，yk、xh、repo 三处立即可见；删除后四处同步消失
4. **A5 验证**：手工 `python claude_migrate.py` 跑一次 mc → yk 迁移，迁完后 `fsutil reparsepoint query "C:\Users\xy24\.claude\projects\D--...\memory"` 仍报告是 junction（未被穿透）；`.claude/auto-memory/` 11 文件未损
5. **B 验证**：脚本 `git diff CLAUDE.md` 提取删除行 → 对每行在 `.claude/memory/*.md` + `.claude/auto-memory/*.md` 中 grep 等价语义；脚本判定后人工 review 一遍
6. **整体验证**：`git status` 应只有预期的修改/新增；`git log -1 --stat` commit 内容覆盖所有 9 个文件改动

## 影响范围

- **不动**：src/、pipeline/、orchestrator.py、Main.py 等业务代码
- **不动**：`.claude/agents/`、skills/、debug/
- **新增**：4 个工具脚本（约 300 行）+ 1 个 reference memory + 11 个 auto-memory 文件
- **修改**：CLAUDE.md（压缩）、claude_migrate.py（约 20 行）、.gitignore（约 2 行）
