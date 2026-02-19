# CLAUDE.md - 多Agent调度系统 项目规范

> 本文件每次会话自动加载。保持精简，避免浪费 token。

---

## 项目概览

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
│   ├── settings.json        # Hook配置（启动时缓存，改后需重启）
│   └── CLAUDE.md            # 本文件
├── tests/unit/              # 单元测试
├── PLAN.md                  # Architect 生成的实施计划
└── claude-progressXX.md     # Agent 工作记录
```

---

## 关键规则

### 路径规范
- 始终使用**相对路径** + **正斜杠 `/`**
- ✅ `src/orchestrator_v6.py` ❌ `C:\Users\...\src\orchestrator_v6.py`
- 遇到 "File has been unexpectedly modified" → 重新读取文件，用相对路径重试

### 工作方式
- **最小改动原则**：只改必要的部分
- **先说明再动手**：修改前简述要改哪些文件
- **不确定就问**：不要猜测路径、环境、配置

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

## PPT Pipeline（pipeline/）

与 orchestrator 独立的确定性脚本，直接运行生成 PPT：

```bash
python pipeline/02_shape_analysis.py   # 读取 shape_detail.md 批注
python pipeline/03_build_shape.py      # GPT 生成内容（需 openai httpx 包）
python pipeline/03_build_ppt_com.py --version 1.x
python pipeline/04_shape_diff_test.py --target "codex 1.x.pptx"
```

- GPT 模型：`openai/gpt-5.2`（OpenRouter）
- 不修改 `src/Function_030.py`，直接 `from src.Function_030 import GPT_5`
- 新模板：改 `pipeline/ppt_pipeline_common.py` 中 TEMPLATE_PATH/EXCEL_PATH → 跑 Step 1 → 填 shape_detail.md 批注

---

## 其他注意事项

- Python 代码中不要用中文全角符号（`（）` → `()`）
- Excel 处理用 `pandas` + `openpyxl`，不要用 Read 读 .xlsx
- 自定义命令文件名须与 JSON 内 `name` 字段一致
