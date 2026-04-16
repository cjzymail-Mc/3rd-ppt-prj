---
name: Manual Pipeline Commands
description: Step-by-step commands to run the pipeline without orchestrator
type: reference
---

## 手动 Pipeline（不走 Orchestrator）

```bash
python pipeline/01_shape_detail.py                                # → xlsx + JSON
python pipeline/01b_auto_annotate.py                              # → 自动填写xlsx批注
# 用户编辑 01-shape_detail.xlsx 黄色单元格
python pipeline/02_shape_analysis.py                              # → 02-*.json
python pipeline/03a_build_shape.py                                # → 03a-*.json
python pipeline/03b_build_ppt_com.py --version 1.0                # → claude-ppt 1.0.pptx
# 自检: python pipeline/self_check.py step1  /  python pipeline/self_check.py step2
```

## 用户批注字段（01-shape_detail.xlsx）

| 字段 | 必填 | 说明 |
|------|------|------|
| **内容描述** | 是(黄色) | 映射知识入口：来源+方向+关键词要求+格式约束 |
| strategy | 否 | 精确策略代码，覆盖自动识别 |
| params | 否 | `source=补充说明, filter=缺点` |

> **备注字段已废弃**，所有指令统一写入「内容描述」。02 会自动解析 output_contract 子字段。

## Step3 预检与反馈文件

| 文件 | 用途 |
|------|------|
| `03-feedback_to_step2.json` | Step3 内容超长反馈 → step2 消费后自动删除 |
| `03a-pending_prompts.json` | Step3 启动前与 Excel prompt 对比的基准 |
| `03b-baseline_page.png` | 模板截图（剪贴板→Pillow，绕过加密） |
| `03b-generated_page.png` | 生成 PPT 截图（同上） |

Step3 启动时自动执行：
1. `_sync_excel_prompts()` — Excel prompt 变化则补跑 GPT
2. 显示 step2 遗留问题
3. 运行 03b pipeline
4. 自检（属性 + 结构 + SSIM + 内容 + 字体）

## 模板文件位置

模板和数据文件统一放在 `template/` 目录。支持多套文件，orchestrator 启动时会提示选择。
也可通过环境变量覆盖：`PPT_TEMPLATE_PATH` 和 `PPT_EXCEL_PATH`。
