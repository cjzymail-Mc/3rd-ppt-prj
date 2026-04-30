---
name: 3 账号 auto-memory junction 架构
description: 3 个 Claude Pro 账号（mc/yk/xh）通过 NTFS junction 共享同一 memory 物理目录的架构记录 + 维护警告
type: reference
---

## 背景

用户用 3 个 Claude Pro 账号轮换工作于本项目（切 token 上限时换账号）。Claude Code 的 auto-memory 自动写入 `<账号根>/projects/<项目>/memory/`，3 个账号根目录物理隔离 → 切账号"失忆"另一边洞察。

2026-04-29 用 NTFS junction 把 3 个账号 memory 目录物理合并到 repo 内 `.claude/auto-memory/`，进 git，永久杜绝漂移。

## 3 账号路径表

| 账号 | 根目录 | memory 路径（junction 入口） |
|--|--|--|
| mc | `C:\Users\xy24\.claude-mc\` | `<root>\projects\<proj>\memory\` |
| yk | `C:\Users\xy24\.claude\` | `<root>\projects\<proj>\memory\` |
| xh | `C:\Users\xy24\.claude-xh\` | `<root>\projects\<proj>\memory\` |

`<proj>` 由 `claude_migrate.py:detect_project_name()` 自动生成（路径转字符串），当前是 `D--Technique-Support-Claude-Code-Learning-3rd-ppt-prj`。

## Junction 工作原理

3 个账号的 `memory\` 入口在 OS 看来是 3 个独立目录，实际都是 NTFS reparse point，全部重定向到 D 盘 repo 内的 `D:\...\3rd-ppt-prj\.claude\auto-memory\`。

```
Claude Code 写入路径（它认为的）        实际物理路径（OS 重定向后）
┌──────────────────────────┐
│ <mc>\...\memory\         │ ──┐
├──────────────────────────┤   │   ┌─────────────────────────────┐
│ <yk>\...\memory\         │ ──┼─→ │ D:\...\.claude\auto-memory\ │
├──────────────────────────┤   │   │   ├── MEMORY.md      ←本体  │
│ <xh>\...\memory\         │ ──┘   │   ├── feedback_*.md  ←本体  │
└──────────────────────────┘       │   └── ...                   │
                                   └─────────────────────────────┘
```

**关键事实**：
- C: 上 3 个 `memory\` 入口各自只是一条 ~100 字节的 reparse point；**没有任何 .md 文件物理存在于 C:**
- 文件本体只在 D: 上有唯一物理副本
- Claude Code 在 OS 文件系统层完全无感知，按默认 hardcoded 路径读写即可
- 任何账号的写入立即对其他 2 个账号可见（同一物理文件）

## 工具脚本

全部位于 `skills/`：

| 脚本 | 用途 | 调用 |
|--|--|--|
| `memory_union_merge.py` | 合并 3 账号 memory 到 `.claude/auto-memory/`（一次性，dry-run / apply） | `python skills/memory_union_merge.py --apply` |
| `memory_junction_setup.py` | 备份 3 账号 memory + rmtree + mklink /J 建 junction | `python skills/memory_junction_setup.py --apply` |
| `memory_junction_rollback.py` | 解 junction（os.rmdir）+ 从 `.pre-junction-backup/` 恢复 | `python skills/memory_junction_rollback.py --apply [--account mc]` |
| `memory_junction_verify.py` | 端到端验证 junction 状态 + 写/删传播测试 | `python skills/memory_junction_verify.py` |

## 自检命令

任何时候都可以快速验证 junction 是否生效：

```bash
fsutil reparsepoint query "C:\Users\xy24\.claude-mc\projects\D--Technique-Support-Claude-Code-Learning-3rd-ppt-prj\memory"
```

输出含 `Reparse Tag Value: 0xa0000003` 或 `Mount Point` 即生效。

或跑 `python skills/memory_junction_verify.py`（自动测 3 账号 + 写传播）。

## 未来维护警告（关键安全规则）

Junction 引入了几个反直觉的危险操作：

| 操作 | 行为 | 安全性 |
|--|--|--|
| `os.rmdir(memory_path)` 或 `rmdir <path>`（无 `/s`） | 只删 reparse point 入口，本体安全 | ✅ 安全（rollback 用） |
| `shutil.rmtree(memory_path)` 或 `rmdir /s <path>` | **穿透 junction 杀 D 盘本体** | ⚠️ 危险 |
| `del <path>\*` 或 `rm -rf <path>/*` | 同上，杀本体 | ⚠️ 危险 |
| 删除账号根目录 `C:\Users\xy24\.claude-mc\` 整个 | 递归穿透，**直接杀 D 盘本体** | ⚠️ 必须先解 junction |
| 文件管理器拖拽删除 `memory` 文件夹 | 走回收站，不穿透 | ✅ 安全（回收站可恢复） |
| 第三方备份工具（File History 等） | 默认不进入 reparse point | ✅ 安全（不会 3 倍备份） |

`claude_migrate.py` 已通过 `is_junction()` + `_safe_clean_target_project()` + `_ignore_memory_junction` 处理这个问题（line ~75 区域）；新写任何账号目录操作脚本都必须类似处理。

## 未来要彻底删除某个账号时的正确流程

1. `python skills/memory_junction_rollback.py --apply --account mc` 解 mc 的 junction
2. 此时 mc 账号下 `memory\` 已不再是 reparse point（恢复了备份内容或为空）
3. 再删账号根目录 `C:\Users\xy24\.claude-mc\`，安全

## 备份位置

`.claude/auto-memory/.pre-junction-backup/<account>-<num>/`，已 `.gitignore` 排除。

## 失效场景与恢复

- **重装系统**：junction 是 NTFS 卷的属性，跟卷走；如果 D 盘重装则丢，需重跑 `memory_junction_setup.py`
- **目录移动**：移动 repo 路径后 junction target 失效，需 rollback + 重做
- **新机器**：需重跑 `memory_union_merge.py` + `memory_junction_setup.py`
- **某账号 memory 被误用 `rmtree` 穿透**：D 盘本体已被破坏，从 git 历史恢复 `.claude/auto-memory/`；不要用 `.pre-junction-backup`（那是 setup 前的旧版本）
