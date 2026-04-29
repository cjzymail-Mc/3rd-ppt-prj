同步本项目在 3 个 Claude Pro 账户间的 memory 文件。

## 3 个账户路径（与 claude_migrate.py 一致）

| key | 路径 |
|--|--|
| mc | `~/.claude-mc` |
| yk | `~/.claude` |
| xh | `~/.claude-xh` |

## memory 路径

```
{account_dir}/projects/{project_folder}/memory/
```

`project_folder` = 当前工作目录绝对路径，把 `:` `\` `/` `空格` 全部替换为 `-`，再 rstrip("-")。
（与 `claude_migrate.py::detect_project_name()` 保持一致；可直接 `python -c "import claude_migrate as m; print(m.detect_project_name())"` 验证。）

## 执行步骤

1. **检测**：列出 3 个账户 memory 目录下所有 .md 文件 + md5 + mtime（任一账户该路径不存在则跳过并标注）
2. **对账表**（必须先打印给用户看）：
   - ✓ 三方 md5 一致 → 跳过
   - △ 某账户独有 → 待补（来源 → 其余）
   - ✗ 三方 md5 不同 → 冲突
3. **决策**：
   - 独有文件：来源账户 → 其余账户（cp -p 保留 mtime）
   - MEMORY.md（纯索引）冲突：用 `diff` 严格判定超集；若有版本是其他全部的超集 → 取超集；否则 **停下来问用户**
   - 其他文件冲突：默认取 mtime 最新，但**先打印 diff 给用户看再执行**
4. **执行**：仅 `cp -p`，不改写文件内容；不动 memory 目录之外的任何文件
5. **验证**：重新 md5sum，3 账户应完全一致；输出最终对账表

## 硬规则

- **只动 memory 目录下的 .md**，不要碰 `projects/{project_folder}/` 下的其他文件（如 .jsonl 会话历史、settings.json）
- **不要做内容级合并**：只做文件级同步 + MEMORY.md 严格超集；任何"模型来判断哪几行该保留"的情形必须问用户
- 同步前先输出对账表，得到非冲突项的同步清单后立即执行（不需逐条问）
- 冲突项必须显式停下来，列出 diff 让用户裁决
- 同步后必须用 md5sum 三方对比验证一致
