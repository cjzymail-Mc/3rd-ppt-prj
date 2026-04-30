---
name: 3 账号 auto-memory junction 同步方案
description: 多个 Claude Pro 账号轮换工作时，用 NTFS junction 把各账号 auto-memory 物理合一到 repo 内，永久杜绝漂移；附移植到其他项目/其他电脑的复制清单
type: skill
---

# 3 账号 auto-memory junction 同步方案（可移植）

## 一句话总结

把 N 个 Claude Pro 账号的 `<账号根>/projects/<项目>/memory/` 目录通过 **NTFS junction** 全部重定向到 repo 内某个物理目录（建议 `.claude/auto-memory/`），让 OS 在文件系统层透明合一，Claude Code 完全无感知。git 自动跨设备同步。

## 适用场景

- 在 Windows 上用多个 Claude Pro 账号轮换工作（碰 token 上限就切账号）
- 同一个项目下，每次切账号 auto-memory 就"失忆"另一边的洞察
- 不想改 Claude Code 默认行为（CLAUDE.md 强制要求也未必 100% 遵守）
- 想要 auto-memory 进 git，跨机自动同步

**不适用**：单账号工作 / 非 Windows / 非 NTFS 文件系统。

## 架构原理（简略）

```
Claude Code 写入路径（它认为的）           实际物理路径
┌──────────────────────────┐
│ <accountA>\...\memory\   │ ──┐
├──────────────────────────┤   │   ┌─────────────────────────────┐
│ <accountB>\...\memory\   │ ──┼─→ │ <repo>\.claude\auto-memory\ │
├──────────────────────────┤   │   │   ├── MEMORY.md     ←本体   │
│ <accountC>\...\memory\   │ ──┘   │   ├── feedback_*.md ←本体   │
└──────────────────────────┘       │   └── ...                   │
                                   └─────────────────────────────┘
```

**关键事实**：
- 账号侧的 `memory\` 入口在 OS 看来是一条 ~100 字节的 NTFS reparse point；**没有任何 .md 文件物理存在于 C 盘**
- 文件本体只在 repo 物理目录有唯一副本
- 任何账号的写入立即对其他账号可见（同一物理文件）
- Claude Code 在 `CreateFile` / `WriteFile` 调用层就被透明重定向，应用层完全无感知

## 移植到新项目（同一台机器）

适用：你已经在这台机器的某个项目用过这套方案，想给另一个项目（同一批账号）也启用。

**前置**：4 个 `memory_*.py` 脚本已经在源项目的 `skills/` 里。

```bash
# 假设源项目：D:\proj-A\，目标项目：D:\proj-B\
cd D:\proj-B

# 1. 拷贝 4 个 Python 脚本到目标项目
mkdir skills 2>nul
copy D:\proj-A\skills\memory_union_merge.py skills\
copy D:\proj-A\skills\memory_junction_setup.py skills\
copy D:\proj-A\skills\memory_junction_rollback.py skills\
copy D:\proj-A\skills\memory_junction_verify.py skills\
copy D:\proj-A\skills\memory-junction-3account.md skills\

# 2. 拷贝 reference memory（架构记录）
copy D:\proj-A\.claude\memory\reference_3account_junction.md .claude\memory\

# 3. .gitignore 追加 2 行（如还没有）
echo .claude/auto-memory/.pre-junction-backup/ >> .gitignore
echo .claude/auto-memory/_test_*.md >> .gitignore

# 4. 跑 union merge（dry-run 看决策表，apply 实际执行）
python skills\memory_union_merge.py
python skills\memory_union_merge.py --apply

# 5. 关闭所有 Claude Code 会话，跑 junction setup
python skills\memory_junction_setup.py
python skills\memory_junction_setup.py --apply

# 6. 端到端验证
python skills\memory_junction_verify.py

# 7. git add + commit
```

**`memory_*.py` 脚本无需任何代码修改**——它们用 `os.path.dirname(__file__)` 自动推断项目根，用 `path.replace(":/\\\\ ", "-")` 自动推断 Claude Code 的项目目录名。

## 移植到新机器（同一批账号）

适用：你换了一台 Windows 机器（或重装系统），希望恢复同样的方案。

```bash
# 1. clone 项目（auto-memory 已经在 git 里）
git clone <repo-url>
cd <repo>

# 2. 验证账号目录是否存在（看路径是否对）
dir "%USERPROFILE%\.claude-mc"
dir "%USERPROFILE%\.claude"
dir "%USERPROFILE%\.claude-xh"

# 3. 跑一次 dry-run（如果 3 账号都没有该项目的 memory 目录，会提示"源目录不存在，跳过"，这是正常的）
python skills\memory_junction_setup.py

# 4. apply 建 junction
python skills\memory_junction_setup.py --apply

# 5. 验证
python skills\memory_junction_verify.py
```

**特别注意**：跨机器时 NTFS junction 的"重定向"信息是写在源盘 NTFS 元数据里的，**不会跟着 git 走**。所以必须在新机器上重新跑 `memory_junction_setup.py --apply` 来重建 junction。git 同步的只有 `.claude/auto-memory/` 这个物理目录的文件本体。

## 移植到不同账号配置

如果你的账号不是 mc/yk/xh 三个、不是 `~/.claude-{xx}/` 这种命名，需要改 4 个脚本顶部的 `ACCOUNTS` 字典：

```python
ACCOUNTS = {
    "你的别名1": os.path.join(USERPROFILE, ".claude-xxx"),
    "你的别名2": os.path.join(USERPROFILE, ".claude"),
    # 加更多账号...
}
```

4 个脚本都引用同一份 ACCOUNTS（可以 import 共享，但当前是各自独立的副本——简单一致，不抽公共文件）。

## 4 个工具脚本清单

全部在 `skills/`：

| 脚本 | 用途 | 何时用 |
|--|--|--|
| `memory_union_merge.py` | 扫描 N 账号 memory，求 union 写入 repo 物理目录 | 一次性首次设置（或 rollback 后重新合并） |
| `memory_junction_setup.py` | 备份 + rmtree + `mklink /J` 建 junction | 一次性首次设置 / 移植到新机器 |
| `memory_junction_rollback.py` | `os.rmdir`（**不是 rmtree**）解 junction + 从备份恢复 | 误操作恢复 / 卸载某账号前 |
| `memory_junction_verify.py` | 检测 junction 状态 + 端到端写/删传播测试 | 任何时候自检 |

每个脚本都有 `--dry-run`（默认）和 `--apply` 模式；`rollback` 还支持 `--account <key>` 单独回滚一个账号。

## 自检命令

随时可以验证 junction 是否健在：

```bash
# 单点检测
fsutil reparsepoint query "<account_path>\projects\<proj>\memory"
# 输出含 "Reparse Tag Value: 0xa0000003" 或 "Mount Point" 即生效

# 全套自动测试
python skills\memory_junction_verify.py
```

## 维护警告（关键安全规则）

Junction 引入了几个反直觉的危险操作。**任何写在新工具脚本里的"账号目录清理"代码都必须遵守这些规则**：

| 操作 | 行为 | 安全性 |
|--|--|--|
| `os.rmdir(memory_path)` 或 `rmdir <path>`（无 `/s`） | 只删 reparse point 入口，本体安全 | ✅ 安全（rollback 用） |
| `shutil.rmtree(memory_path)` 或 `rmdir /s <path>` | **穿透 junction 杀 repo 本体** | ⚠️ 危险 |
| `del <path>\*` 或 `rm -rf <path>/*` | 同上，杀本体 | ⚠️ 危险 |
| 删除账号根目录 `C:\Users\xy24\.claude-xxx\` | 递归穿透，**杀 repo 本体** | ⚠️ 必须先解 junction |
| 文件管理器拖拽删除 `memory` 文件夹 | 走回收站，不穿透 | ✅ 安全 |
| 第三方备份工具（File History 等） | 默认不进入 reparse point | ✅ 安全（不会重复备份 N 份） |

**写"清理目标账号目录"代码时的标准写法**：
```python
def is_junction(path):
    if not os.path.isdir(path): return False
    attrs = ctypes.windll.kernel32.GetFileAttributesW(path)
    if attrs == 0xFFFFFFFF: return False
    return bool(attrs & stat.FILE_ATTRIBUTE_REPARSE_POINT)

# 清理时必须按 entry 遍历，跳过 junction
for entry in os.listdir(target_project):
    full = os.path.join(target_project, entry)
    if entry == "memory" and is_junction(full):
        continue  # 关键：跳过 junction
    if os.path.isdir(full) and not is_junction(full):
        shutil.rmtree(full)
    else:
        os.remove(full)
```

参考实现：本项目 `claude_migrate.py:_safe_clean_target_project()`。

## 卸载某个账号的正确流程

未来要彻底删除某账号（比如不再用了）：

```bash
# 1. 先解该账号的 junction
python skills\memory_junction_rollback.py --apply --account <account_key>

# 2. 此时该账号 memory 已不再是 reparse point（恢复了备份内容或为空）
# 验证：
fsutil reparsepoint query "<account_path>\...\memory"
# 应报告 "The file or directory is not a reparse point."

# 3. 才能安全地删账号根目录
rmdir /s /q "%USERPROFILE%\.claude-xxx"
```

## 故障恢复

| 故障 | 原因 | 恢复 |
|--|--|--|
| `mklink` 报"已存在" | 之前 setup 没清干净 | `os.rmdir` 删除残留入口（不要 rmtree！）后重跑 |
| `mklink` 报"无效路径" | 目标盘非 NTFS（FAT/exFAT 不支持 reparse） | 换一个 NTFS 盘做物理目录 |
| `fsutil reparsepoint query` 报"不是 reparse point" | junction 失效（系统重装、目录被覆盖） | 重跑 `memory_junction_setup.py --apply` |
| Claude 写入 memory 报权限错误 | 某账号 memory 是 junction，但 repo 物理目录被 git stash / 删了 | 检查 repo 内 `.claude/auto-memory/` 是否存在 |
| 账号侧"看不到 memory 内容" | repo 物理目录里没文件 | 跑 `memory_junction_verify.py` 看 junction 状态；若 junction 是好的就重跑 union merge |
| 所有账号 memory 内容被穿透 rmtree 杀掉 | 有代码用 `shutil.rmtree` 走过 junction（违反安全规则） | git 历史恢复 `.claude/auto-memory/`；**不要**用 `.pre-junction-backup`（那是 setup 前的旧版本） |

## 该方案不解决的问题

- **多人协作冲突**：仍然要靠 git merge 解决文件层面的冲突；junction 只解决"同一人多账号"漂移
- **跨操作系统**：本方案 Windows 专属；macOS/Linux 用 symlink 实现思路类似但脚本要重写（`os.symlink` + 不同的 reparse 检测）
- **Claude 项目目录名变化**：如果 Claude Code 升级改了项目目录命名规则，4 个脚本的 `detect_project_name()` 要跟着改

## 参考实现

本项目的具体实现：
- 4 个脚本：`skills/memory_union_merge.py` / `memory_junction_setup.py` / `memory_junction_rollback.py` / `memory_junction_verify.py`
- 架构记录：`.claude/memory/reference_3account_junction.md`
- migrate 工具的安全升级：`claude_migrate.py:is_junction()` + `_safe_clean_target_project()` + `_ignore_memory_junction`
