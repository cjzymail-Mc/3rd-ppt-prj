# CLAUDE.md - PPT Pipeline + Agent 项目规范

> 本文件每次会话自动加载。保持精简。

---

## 项目结构

```
项目根目录/
├── orchestrator.py                  # 4-Agent 调度（Pipeline先行 + LLM精调）
├── pipeline/
│   ├── ppt_pipeline_common.py       # 公共工具（路径、COM、Excel、批注解析）
│   ├── 01_shape_detail.py           # Step 1: 提取模板shape
│   ├── 01b_auto_annotate.py         # Step 1B: 规则表自动批注
│   ├── 02_shape_analysis.py         # Step 2: 角色推断 + prompt生成
│   ├── 02b_iteration_setup.py       # Step 2B: 修正轮sheet创建 + 基础修正
│   ├── 03a_build_shape.py           # Step 3A: 内容生成（Python/GPT）
│   ├── 03b_build_ppt_com.py         # Step 3B: COM写入PPT
│   ├── 04_shape_diff_test.py        # Step 4: 三层验收 + fix_type分类
│   └── prompt_templates/gpt_summary.md  # GPT prompt 模板
├── pipeline-progress/               # 中间产物（01-/02-/03a-/03b-/04- 前缀）
├── .claude/agents/                  # 4个Agent配置
│   ├── 01-analyst.md                # 分析师：Pipeline推断 + LLM审核模糊项
│   ├── 02-builder.md                # 构建师：Pipeline生成 + LLM精调批注(修正轮)
│   ├── 03-reviewer.md               # 验收师：Pipeline测试 + LLM语义审核
│   └── 04-developer.md              # 代码专家：LLM修复pipeline代码
└── src/Function_030.py              # GPT_5 函数（不修改，直接import）
```

---

## 关键规则

- **路径**: 始终用相对路径 + 正斜杠 `/`
- **最小改动**: 只改必要的部分，先说明再动手
- **输出**: 改代码时只说结论（改了什么、为什么、结果），不展示 diff
- **Excel**: 统一用 `win32com.client` COM（加密环境，禁 openpyxl/pandas）
- **PPT**: Clone 模板页，不新建 shape；禁 `python-pptx`

---

## 混合工作流（Pipeline + Agent）

### 启动

```bash
python orchestrator.py          # 交互选择轮次(1-3)
python orchestrator.py --max-rounds 2   # 或直接指定
```

### 流程

```
[Analyst] Pipeline(01+01b) → LLM增强所有批注 → 填写xlsx
    ↓
  PAUSE — 用户校准xlsx → Enter继续
    ↓
[Builder] 直接Pipeline(02→03a→03b) → claude-ppt 1.0.pptx    ← 首轮无LLM
    ↓
[Reviewer] 直接Pipeline(04验收) → PASS/FAIL
    ↓ FAIL → LLM语义审核 → fix_type分流
    ├─ annotation → 直接Pipeline(02b) → [Builder] LLM精调xlsx → 直接Pipeline(02→03a→03b)
    └─ code → [Developer] LLM修代码 → Builder重跑
    ↓
[Reviewer] 重新验收 → 循环至 max_rounds
```

### 混合模式：Pipeline 由 orchestrator 直接执行，LLM 只做智能任务

| Agent | orchestrator 直接执行 | LLM 负责 | 何时跳过LLM |
|-------|---------------------|---------|------------|
| Analyst | 01提取 + 01b规则推断 | 增强所有shape批注 | 从不跳过 |
| Builder首轮 | 02→03a→03b全链路 | 无 | 始终 |
| Builder修正轮 | 02b + 02→03a→03b | 仅精调xlsx批注 | — |
| Reviewer | 04三层测试 | 语义审核,补充精准建议 | PASS时 |
| Developer | 无 | 读报告+修代码 | 无code问题时 |

### 版本追溯

| 轮次 | xlsx Sheet | PPT 文件 |
|------|-----------|----------|
| 首轮 | Shape Detail | claude-ppt 1.0.pptx |
| 第2轮 | claude-ppt 1.1 | claude-ppt 1.1.pptx |
| 第3轮 | claude-ppt 1.2 | claude-ppt 1.2.pptx |

### 三层门禁（全部达标=PASS）

| 层级 | 阈值 | 检查内容 |
|------|------|---------|
| Visual | >= 98 | 几何位置、字体、颜色、ChartType |
| Readability | >= 95 | 文本长度比、行数比 |
| Semantic | = 100 | 关键词覆盖：样本、建议、反馈 |

### fix_type 分流（5 类）

| fix_type | 含义 | 后续动作 |
|----------|------|---------|
| `keyword_missing` | 语义关键词缺失 | 02b 追加关键词要求 → 重跑 pipeline |
| `budget_overflow` | 文本过长 | 02b 追加字数约束 → 重跑 pipeline |
| `budget_underflow` | 文本过短/空白 | 02b 要求充实内容 → 重跑 pipeline |
| `style_mismatch` | 格式/语调偏离 | 02b 追加风格约束 → 重跑 pipeline |
| `code` | pipeline代码缺陷 | Developer修代码 → Builder重跑 |

> orchestrator 路由逻辑：`code` → Developer，其余全部 → Builder(02b→pipeline)

---

## 手动 Pipeline（不走 Orchestrator）

```bash
python pipeline/01_shape_detail.py                                # → xlsx + JSON
python pipeline/01b_auto_annotate.py                              # → 自动填写xlsx批注
# 用户编辑 01-shape_detail.xlsx 黄色单元格
python pipeline/02_shape_analysis.py                              # → 02-*.json
python pipeline/03a_build_shape.py                                # → 03a-*.json
python pipeline/03b_build_ppt_com.py --version 1.0                # → claude-ppt 1.0.pptx
python pipeline/04_shape_diff_test.py --target "claude-ppt 1.0.pptx"  # → 04-* reports
```

### 用户批注字段（01-shape_detail.xlsx）

| 字段 | 必填 | 说明 |
|------|------|------|
| **内容描述** | 是(黄色) | 映射知识入口：来源+方向+关键词要求+格式约束（见下方 golden reference） |
| strategy | 否 | 精确策略代码，覆盖自动识别 |
| params | 否 | `source=补充说明, filter=缺点` |

> **备注字段已废弃**，所有指令统一写入「内容描述」。02 会自动解析 output_contract 子字段。

#### 内容描述 golden reference（gpt_prompted 类）

```
缺点: 从补充说明总结缺点。必须包含'建议'、'反馈'、'样本'关键词，用【】括起关键性能词，每段结论后注明(X/N)比例
优点: 从补充说明总结优点。必须包含'建议'、'反馈'、'样本'关键词，用【】括起关键性能词，每段结论后注明(X/N)比例
```

### 关键配置

- 模板: `pipeline/standard and empty template.pptx`（Slide1=空白, Slide2=标准）
- 数据: `pipeline/source data.xlsx`
- GPT: `openai/gpt-5.4`（OpenRouter），`from src.Function_030 import GPT_5`

---

## COM 开发规范

| 场景 | 错误做法 | 正确做法 |
|------|---------|---------|
| 读COM属性 | `getattr(shp,"X",None)` | `try: shp.X except: None` |
| 多步骤开Excel | `Dispatch` 复用实例 | `DispatchEx` + `sleep(0.5)` 强制新进程 |
| 写图表数据 | `ChartData.Workbook` | `SeriesCollection(1).Values/XValues` |
| 插入图片 | `AddPicture(W=w,H=h)` | 先`-1/-1`取原始尺寸,再等比缩放 |
| Clone幻灯片 | 不加sleep | `Copy→sleep(1.5)→Paste(X)→sleep(1.0)` |

---

## 附：src/ 目录（非 Pipeline 核心）

- `src/Function_030.py` — GPT_5 函数，Pipeline 通过 `import` 调用
- `src/` 下其他文件为历史遗留的 main.py 相关模块，与 Pipeline/Agent 工作流无关
