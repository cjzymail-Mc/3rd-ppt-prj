---
name: Pipeline Code Repair Guide
description: Pipeline file inventory, common fix types, tech stack constraints, and self-check requirements for code repairs done in Claude Code main conversation
type: reference
---

# Pipeline 代码修复指引

> 当 step3-builder 诊断出代码层问题时，用户会在 Claude Code 主对话中说明修复需求。本文件提供修复所需的完整知识。

---

## 1. Pipeline 文件清单与职责

| 文件 | 职责 |
|------|------|
| `pipeline/ppt_pipeline_common.py` | 公共工具：路径常量、COM Excel 读写、批注解析、shape detail xlsx 生成 |
| `pipeline/01_shape_detail.py` | 提取 PPT 模板 shape 结构 -> JSON + xlsx |
| `pipeline/01b_auto_annotate.py` | 规则表自动批注（strategy/params 推断）|
| `pipeline/02_shape_analysis.py` | 角色推断 + prompt 规格生成 |
| `pipeline/03a_build_shape.py` | 内容生成：组装 prompt (`--assemble-only`) + 调 GPT (`--execute-prompts`) |
| `pipeline/03b_build_ppt_com.py` | COM 写入 PPT，内置 4 步自检 + MAX_SELF_FIX=2 自动修复 |
| `pipeline/self_check.py` | 自检函数库：`check_step1()`, `check_step2()`, `load_golden_reference()` |
| `pipeline/fix_chart_link.py` | 图表链接修复工具 |

---

## 2. 常见修复类型

| 修复类型 | 典型症状 | 常见根因 |
|---------|---------|---------|
| 列名错误 | `_SCORE_COLS` / `_ANNO_KEYS` KeyError | Excel 表头变更，代码中硬编码列名未同步 |
| 策略路由 | shape 使用了错误的 build 函数 | `STRATEGY_CODES` 未注册新策略，或 `02_shape_analysis.py` 映射逻辑有误 |
| COM 写入 | 字体/位置/大小不对 | `03b_build_ppt_com.py` 中 COM API 调用参数错误 |
| 图表处理 | chart shape 写入失败 | `has_chart=true` 的 shape 需要特殊 COM 路径 |
| 编码问题 | 中文乱码 | 缺少 `encoding='utf-8'` 或 COM 返回值未正确处理 |

---

## 3. 技术栈约束

- **Excel 操作**：统一用 `win32com.client` COM（加密环境，禁 openpyxl / pandas）
- **PPT 操作**：Clone 模板页，不新建 shape；禁 `python-pptx`
- **路径**：始终用相对路径 + 正斜杠 `/`
- **GPT 调用**：`from src.Function_030 import GPT_5`，模型 `openai/gpt-5.4`

---

## 4. 修复后自检要求

1. `python -m py_compile pipeline/<modified_file>.py` — 语法检查
2. 重跑相关 pipeline 脚本验证功能正确
3. 如修改了 `self_check.py`，运行 `python pipeline/self_check.py step1` 和 `step2` 确认无报错

---

## 5. ppt_pipeline_common.py 关键 Helper 函数

| 函数 | 用途 |
|------|------|
| `parse_user_annotations()` | 从 xlsx 读取用户批注 -> `{shape_name: {key: value}}` |
| `write_gpt_prompts_to_xlsx(prompts)` | 将 GPT prompt 写入 xlsx 的 `GPT-prompt Text` 列 |
| `read_gpt_prompts_from_xlsx()` | 从 xlsx 读取已有的 GPT prompt |
| `generate_shape_detail_xlsx(shapes)` | 生成 shape detail xlsx（含批注区域）|
| `load_excel_rows(sheet_name)` | 加载数据源 Excel 行数据 |
| `has_user_annotations()` | 快速检查是否有用户批注 |
| `STRATEGY_CODES` | 有效策略码集合 |
| `PROGRESS_DIR` | `pipeline-progress/` 路径常量 |
| `SHAPE_DETAIL_XLSX` | `pipeline-progress/01-shape_detail.xlsx` 路径常量 |
