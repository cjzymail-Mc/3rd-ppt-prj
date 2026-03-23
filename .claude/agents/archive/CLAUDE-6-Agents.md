# CLAUDE-6-Agents.md - 多Agent调度系统（Orchestrator）

> 本文件记录 orchestrator 多Agent调度系统的规范，与 pipeline 独立。

---

## 概览

- **主文件**: `src/orchestrator_v6.py`（~3900行，多Agent调度系统）
- **实际运行**: `mc-dir-v6.py`（根目录备份，需与 src/ 同步）
- **Agent 配置**: `.claude/agents/01-arch.md` ~ `06-secu.md`
- **Hook**: `.claude/hooks/architect_guard.py` + `.claude/settings.json`
- **测试**: `tests/unit/`（61 unit tests）

```
项目根目录/
├── src/orchestrator_v6.py   # 源码（主文件）
├── mc-dir-v6.py             # 运行入口（备份）
├── .claude/
│   ├── agents/              # 6个Agent配置
│   ├── hooks/               # Hook脚本 + 调试日志
│   └── settings.json        # Hook配置（启动时缓存，改后需重启）
├── tests/unit/              # 单元测试
├── PLAN.md                  # Architect 生成的实施计划
└── claude-progressXX.md     # Agent 工作记录
```

### 常用命令
```bash
python mc-dir-v6.py              # 运行调度系统
pytest tests/unit/ -v            # 单元测试
cat .claude/hooks/guard_debug.log  # 查看Hook调试日志
```

---

## Hook 调试经验（重要！）

### 调试日志优先
Hook 不生效时**第一时间加日志**，不要盲猜：
- 日志用**绝对路径**（基于 `os.path.abspath(__file__)` 推算）
- 记录: tool_name、env var、lock file path、cwd、拦截/放行决策
- 查看: `cat .claude/hooks/guard_debug.log`

### settings.json 修改后必须重启会话
Claude Code 启动时缓存配置，中途修改不生效。

### Hook 自锁恢复
修改 hook 引入 bug 导致无差别拦截时：
1. 删除 `.claude/settings.json`
2. **重启 Claude Code 会话**
3. 修复 hook 代码
4. 恢复 settings.json

### Hook 格式
- **exit code 2** = 阻止（stderr 显示为错误）
- **exit code 0** = 放行
- ~~`{"continue": false}`~~ 旧格式无效

---

## Agent 开发工作流

### 阶段 1 — 代码开发（orchestrator + agents）
1. 用 orchestrator 指派 developer agent 生成 `src/新模块_ppt.py`（参考 codex_ppt.py 结构）
2. 硬编码 SHAPES 列表来自 pipeline Step 1 输出的 `01-shape_detail.xlsx`
3. architect agent 审查 shape 策略矩阵；tester agent 跑 diff test

### 阶段 2 — 集成 main.py
- 在 `from src.XXX import *` 行后加 `from src.新模块_ppt import make_xxx_slide`
- 在对应 section 末尾加 ~8 行调用块（参考 `【5.6】Codex 分析页` 的写法）

### 新 PPT 模块结构（参考 src/codex_ppt.py）
```
XXXX_SHAPES = [硬编码 shape 规格列表]   # 来自 01-shape_detail.xlsx 批注
make_xxxx_slide(mc_sht, mc_ppt, mc_slide, ...)  # 唯一公开 API
```
main.py 只需 `from src.xxxx_ppt import make_xxxx_slide` + ~8 行调用。
