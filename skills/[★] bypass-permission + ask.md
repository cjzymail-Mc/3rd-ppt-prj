# Settings.local.json 修改计划

> 给 Claude 的指令：按本文档**只做一件事**——把当前项目的 `.claude/settings.local.json` 替换为下面的"目标内容"，备份原文件，**不要重启 Claude Code**（用户自己来）。

---

## 目标

把 `.claude/settings.local.json` 改成"自动通过几乎所有命令，但 `rm` / `Remove-Item` 仍弹确认"的配置。

## 操作步骤

1. **备份**：将当前 `.claude/settings.local.json` 复制一份为 `.claude/settings.local.json.bak`
2. **整文件替换**：用下面的"目标内容"**完整覆盖** `.claude/settings.local.json`
   - ⚠️ 旧 `allow` 列表里的所有条目（包括 40+ 条具体 `PowerShell(...)`）**整个废弃**，**不要保留任何一条**
   - ⚠️ 旧的 `Mcp(*)` 写法是错的，新配置里用的是 `mcp__*`
   - ⚠️ 如果原文件里有 `defaultMode` 字段（不管在哪个层级），**全部删掉**——不要保留 `bypassPermissions`，它会让下面的 ask 规则失效
3. **停止**：保存后告诉用户已完成，等用户自己重启 Claude Code 会话

## 目标内容（完整替换为这个）

```json
{
  "permissions": {
    "allow": [
      "Bash(*)",
      "PowerShell(*)",
      "Read(*)",
      "Write(*)",
      "Edit(*)",
      "WebSearch(*)",
      "WebFetch(*)",
      "mcp__*"
    ],
    "ask": [
      "Bash(rm *)",
      "PowerShell(Remove-Item *)",
      "PowerShell(rm *)"
    ]
  }
}
```

## 配置说明（供 Claude 理解，不需重复给用户）

- `allow` 通配让绝大多数命令静默执行
- `ask` 列表的优先级**高于** `allow`，所以 `rm` / `Remove-Item` 会被拉回来弹确认
- 这个 ask 优先级的行为是**实测验证过**的（官方文档没明说），未来 Claude Code 版本如果改变此行为，需要重新评估
- 不能用 `defaultMode: "bypassPermissions"`：它会完全跳过 allow/deny/ask 三类规则，只保留 `rm -rf /` 和 `rm -rf ~` 两个硬编码兜底，达不到"普通 rm 也要确认"的需求

## 用户的验证步骤（Claude 不需要做，仅供参考）

用户重启 Claude Code 会话后，会自己跑：
1. 任意非 rm 命令（如 `git status`）→ 应静默通过
2. `rm 某测试文件` → 应弹确认

如两条都符合，配置生效。
