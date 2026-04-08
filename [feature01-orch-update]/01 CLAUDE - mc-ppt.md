# CLAUDE.md — 通用规范（所有 Agent）

> 最后更新：2026-04-03

---

## 0. 防卡顿规范
- 同一方案连续失败 2 次 → 停下来向用户说明原因，提出替代方案
- 预计超过 2 分钟的操作 → 用 Agent(run_in_background) 分流
- 遇到不确定的技术选型 → 先问用户，不要默默试超过 3 分钟

---

## 1. 文件组织约定

```
项目根目录/
├── pipeline/              # Python 检测工具
├── StepN/                 # 每个子任务独立文件夹
│   ├── brief.md           # 需求文档
│   ├── research_pack.md   # 调研资料
│   ├── deck.md            # 内容大纲（中间产物）
│   ├── deck.html          # 最终演示稿
│   ├── deck_manifest.md   # 页面结构清单（Builder 生成，Converter 消费）
│   ├── review_report.md   # 诊断式自检报告
│   └── images/            # 图片资源
├── Debug/                 # 历史备份
├── .claude/commands/      # 自定义命令
├── [work-flow-rebuild]/agents/  # Agent 角色定义文件
├── skills/                # 技能文档（PDF/PPT 转换等）
├── CLAUDE.md              # 本文件（通用规范）
└── todays-task.md         # 每日任务入口
```

---

## 2. 三条核心禁止规则（HTML）

1. **禁止弯引号**：HTML 属性必须用 ASCII 直引号 `"`，不得用 `""`
2. **禁止绝对定位**：内容区域禁用 `position: absolute`，改用 CSS Grid/Flexbox
3. **禁止 overflow:hidden**：内容容器不设裁切，内容自然撑开

---

## 3. 自定义命令

| 命令 | 用途 |
|------|------|
| `/today` | 读取 todays-task.md 并执行 |
| `/role-pm` | 激活 Agent-1 PM 角色 |
| `/role-researcher` | 激活 Agent-2 Researcher 角色 |
| `/role-builder` | 激活 Agent-3 Builder 角色 |
| `/role-converter` | 激活 Agent-4 Converter 角色 |


---

## 4. 变更记录

| 日期 | 变更内容 |
|------|---------|
| 2026-03-30 | 初版创建 |
| 2026-04-02 | 大幅精简：通用规范保留，角色专属内容迁移至 [work-flow-rebuild]/agents/ |
| 2026-04-03 | 新增 deck_manifest.md 为 StepN 标准产物（Builder 生成，Converter 消费） |
| 2026-04-03 | PPT 转换教训固化：agent-4-converter.md 新增「已知 Bug 与教训」C-1~C-8；阶段 B 强制前 3 页自检 |
