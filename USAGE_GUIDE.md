# PPT工程化交付系统 — 使用指南

## 系统概览

本系统通过6个专业Agent协作，以软件工程流程生成高保真PPT（98%+视觉保真度）。

```
用户需求 → Architect(规划) → TechLead(审核) → Developer(构建) → Tester(测试)
                                                      ↑            ↓
                                                      └── fix-ppt.md ──┘  (不通过则循环)
                                                → Optimizer(优化) → Security(审计) → 交付
```

---

## 前置要求

- Windows OS + Microsoft Office（Excel + PowerPoint 必须安装）
- Python 3.10+
- 依赖：`pywin32`、`xlwings`、`openai`
- Claude Code CLI 已安装（`claude` 命令可用）
- 项目根目录包含：
  - `src/Template 2.1.pptx`（模板文件）
  - `2025 数据 v2.2.xlsx`（数据源）
  - `main.py` + `src/Function_030.py` + `src/Class_030.py`（既有能力）

---

## 使用方式

### 方式一：全自动模式（推荐首次使用）

```bash
python orchestrator.py
```

启动后进入交互式REPL，直接输入需求即可：

```
📝 请输入任务（输入 quit 退出）：
> 根据问卷数据生成篮球试穿分析PPT
```

系统会自动：
1. 识别复杂度（PPT关键词 → COMPLEX → 6个agent全上）
2. 创建feature分支
3. 按阶段执行：Architect → TechLead → Developer↔Tester循环 → Optimizer → Security
4. 输出最终PPT和所有中间产物

### 方式二：手动指定Agent（精确控制）

在REPL中使用 `@agent` 语法：

```bash
# 单个agent执行
> @arch 根据new-ppt-workflow.md生成PLAN.md

# 串行执行（→ 分隔）
> @arch 规划 -> @dev 实现 -> @test 验证

# 并行执行（&& 分隔）
> @dev 构建shape内容 && @opti 优化COM稳定性

# 混合模式
> @arch 规划 -> @tech 审核 -> (@dev 实现 && @opti 优化) -> @test 验证
```

**Agent别名：**

| 全名 | 英文简称 | 中文简称 |
|------|---------|---------|
| architect | @arch | @架构 |
| tech_lead | @tech | @技术 |
| developer | @dev | @开发 |
| tester | @test | @测试 |
| optimizer | @opti | @优化 |
| security | @sec | @安全 |

### 方式三：从PLAN.md恢复执行

如果Architect已完成规划，可跳过规划阶段直接执行：

```bash
python orchestrator.py --from-plan
```

### 方式四：带多轮迭代的执行

```bash
python orchestrator.py --from-plan --max-rounds 3
```

Developer和Tester会循环最多3轮，直到diff_result.json全部达标。

### 方式五：自动Architect模式（无需交互确认PLAN.md）

```bash
python orchestrator.py --auto-architect --max-rounds 3
```

---

## 执行流程详解

### Phase 1: 规划和设计

**Architect** 读取 `new-ppt-workflow.md` + `repo-scan-result.md`，生成 `PLAN.md`：
- 5步脚本链路（输入/输出/验收门槛）
- per-shape策略矩阵（哪些用GPT、哪些不用）
- 可读性预算（每个shape的max_chars/max_lines/max_bullets）
- 三层测试阈值

**TechLead** 审核 PLAN.md：
- 检查策略矩阵是否完整（title/sample_stat/chart必须非GPT）
- 检查脚本链路是否完整
- 发现"全量GPT"路线则打回重做

### Phase 2: 开发和测试（循环）

**Developer** 按PLAN.md实现5步脚本：

| Step | 脚本 | 产出 |
|------|------|------|
| 1 | `01-shape-detail.py` | shape_detail_com.json, shape_fingerprint_map.json |
| 2 | `02-shape-analysis.py` | shape_analysis_map.json, prompt_specs.json, readability_budget.json |
| 3A | `03-build_shape.py` | build_shape_content.json, content_validation_report.md, prompt_trace.json |
| 3B | `03-build_ppt_com.py` | **codex X.Y.pptx**, build-ppt-report.md, post_write_readback.json |
| 4 | `04-shape_diff_test.py` | fix-ppt.md, **diff_result.json**, diff_semantic_report.md |

**Tester** 运行Step4差异测试，检查三层门禁：

| 层级 | 阈值 | 检查内容 |
|------|------|---------|
| Visual | >= 98 | 几何位置、shape type、字体、颜色、chart type |
| Readability | >= 95 | 文本长度比、行数比、字符相似度 |
| Semantic | = 100 | 关键语义词全覆盖 |

**不通过** → 归档本轮报告 → Developer读取fix-ppt.md修复 → 下一轮（codex 1.0 → 1.1 → 1.2）

**通过** → 进入Phase 3

### Phase 3: 优化和安全

- **Optimizer**: 优化COM稳定性、缓存中间产物、减少重复读取
- **Security**: 审计API key、输出路径、COM资源释放，产出 SECURITY_AUDIT.md

---

## 关键文件说明

### 输入文件（不可修改）

| 文件 | 说明 |
|------|------|
| `src/Template 2.1.pptx` | PPT模板（第14页=空白基准，第15页=标准模板） |
| `2025 数据 v2.2.xlsx` | 唯一数据源（问卷sheet） |
| `main.py` + `src/*.py` | 既有PPT能力（仅复用，不修改） |

### 配置文件

| 文件 | 说明 |
|------|------|
| `new-ppt-workflow.md` | v4.0执行规范（最高优先级参考） |
| `repo-scan-result.md` | 代码库能力分析（供Agent理解既有代码） |
| `.claude/agents/01~06-*.md` | 6个Agent的角色定义 |
| `.claude/hooks/architect_guard.py` | Architect权限守卫（禁止写非.md文件） |

### 输出产物

| 文件 | 产出阶段 | 说明 |
|------|---------|------|
| `PLAN.md` | Architect | 实施计划 |
| `shape_detail_com.json` | Step1 | shape属性和指纹 |
| `shape_analysis_map.json` | Step2 | shape到数据源映射 |
| `prompt_specs.json` | Step2 | 每个shape的prompt规格 |
| `readability_budget.json` | Step2 | 可读性预算 |
| `build_shape_content.json` | Step3A | 生成的shape内容 |
| `prompt_trace.json` | Step3A | prompt追踪记录 |
| `shape_data_gap_report.md` | Step3A | 数据缺口报告 |
| `codex X.Y.pptx` | Step3B | **最终PPT产物** |
| `build-ppt-report.md` | Step3B | 构建日志 |
| `post_write_readback.json` | Step3B | 写后回读确认 |
| `diff_result.json` | Step4 | 三层评分（JSON） |
| `fix-ppt.md` | Step4 | 修复建议 |
| `SECURITY_AUDIT.md` | Security | 安全审计报告 |

---

## 迭代机制

当diff测试不通过时，orchestrator自动进入下一轮：

```
Round 1: Developer生成 codex 1.0.pptx → Tester测试 → 不通过
         ↓ 归档 fix-ppt.md → fix-ppt-round1.md
Round 2: Developer读取fix-ppt.md修复建议 → 生成 codex 2.0.pptx → Tester测试 → 不通过
         ↓ 归档
Round 3: Developer继续修复 → 生成 codex 3.0.pptx → Tester测试 → 通过!
         ↓
Phase 3: Optimizer + Security
```

**修复优先级路由**（fix-ppt.md中会标注）：
1. 先检查shape策略是否正确（是否错用了GPT？）
2. 再调整prompt（style anchor/instruction/budget）
3. 再改提取函数（extract_info/均值提取/regex）

---

## 常见操作速查

### 只运行特定Agent

```bash
# 只让Architect生成计划
> @arch 根据new-ppt-workflow.md为问卷分析PPT生成PLAN.md

# 只让Developer实现
> @dev 按PLAN.md实现所有步骤脚本

# 只让Tester测试
> @test 运行04-shape_diff_test.py测试codex 1.0.pptx
```

### 从失败中恢复

```bash
# 查看上次执行状态
cat .claude/state.json

# 从断点恢复
python orchestrator.py --resume
```

### 查看调试信息

```bash
# Hook调试日志
cat .claude/hooks/guard_debug.log

# Agent执行错误日志
cat .claude/error_log.json

# 进度文件
cat claude-progress*.md
```

### 手动运行步骤脚本（不通过orchestrator）

```bash
# 运行完整流水线
python 00-ppt.py --start-version 1.0 --max-rounds 3

# 仅重跑diff测试
python 00-ppt.py --from-step 4 --to-step 4 --start-version 1.0 --max-rounds 1

# 仅重跑构建+写入
python 00-ppt.py --from-step 3 --to-step 3 --start-version 1.1 --max-rounds 1
```

---

## 注意事项

1. **settings.json修改后必须重启** — Claude Code启动时缓存配置，中途修改不生效
2. **COM对象必须正确释放** — 异常退出后检查是否有残留的 POWERPNT.EXE / EXCEL.EXE 进程
3. **剪贴板操作需要延时** — 如果遇到粘贴失败，增加delay参数
4. **不要手动修改 src/Template 2.1.pptx** — 这是只读基准模板
5. **所有路径使用正斜杠 `/`** — Windows环境下也使用正斜杠
