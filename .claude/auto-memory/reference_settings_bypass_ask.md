---
name: settings.local.json bypass+ask pattern
description: Validated pattern — allow:* with ask:rm 拉回确认；ask 优先级高于 allow wildcard（官方未明说，实测2026-05-11）
type: reference
originSessionId: abc1ac05-7447-4878-bf6d-580b269037f3
---
`skills/[★] bypass-permission + ask.md` 是该项目验证过的 `.claude/settings.local.json` 模板：用 `allow: ["Bash(*)", "PowerShell(*)", ...]` 通配静默执行，靠 `ask: ["Bash(rm *)", "PowerShell(Remove-Item *)", "PowerShell(rm *)"]` 把删除命令拉回弹确认。

**关键非显事实**：
- `ask` 列表的优先级**高于** `allow` 通配（官方文档没明说，2026-05-11 在本项目实测验证：`git status` 静默通过，`rm empty.md` 弹确认）
- 不能用 `defaultMode: "bypassPermissions"`——它会跳过 allow/deny/ask 全部三类规则，只保留 `rm -rf /` 和 `rm -rf ~` 两个硬编码兜底，ask 规则失效
- 用户手动重启 Claude Code 会话后配置才生效

**何时引用**：用户要求"静默执行 / 减少弹窗 / 但保留 rm 确认"这类权限调整时，直接套该 skill 文档的"目标内容"整文件替换。
