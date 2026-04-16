# AGENTS.md

Last updated: 2026-04-16

## 0. 技术栈速览（最小必要）
- 主语言: `Python`
- 主入口:
  - `orchestrator.py`（Pipeline/Agent 工作流）
  - `Main.py`（src 生产工作流）
- Office 自动化: `pywin32 / win32com.client`（Excel/PPT 主读写）
- Excel 辅助: `xlwings`（主要在 `Main.py` / `src`）
- LLM 调用: `openai` + `httpx`，通过 OpenRouter（`src/Function_030.py::GPT_5`）
- 产物格式: `xlsx / pptx / json / md / png`

## 1. 架构定位（双轨 + 三重混合）
- 生产主线A: `Main.py + src/*`（已知模板、日常交付、Clone保真）
- 分析主线B: `orchestrator.py + pipeline/*`（新模板分析、prompt迭代、自检修复）
- 三重混合建议权重:
  - Pipeline: 负责 shape分析 + prompt管理 + GPT内容生成 + 图表写入经验沉淀
  - Agents: 负责语义修正、自检回路、失败兜底
  - Developer: 负责模板移植与最终10~30%工程收口（不是纯收尾）

## 2. 何时走哪条路
- 已知模板改bug: 直接改 `src/*`
- 新模板但相似: 先用 Pipeline Step1/2 产出结构与prompt，再移植到 `src/{template}_ppt.py`
- 完全陌生模板: 先跑 Pipeline 全流程，再进入 Developer 移植

## 3. 硬规则（必须遵守）
- Excel/PPT 仅用 COM（`win32com.client`）
- 禁 `python-pptx` / `openpyxl` / `pandas` 直接读写
- PPT 页面通过 Clone 模板页生成，不重建 shape
- 文本写入使用 `\r` 分段（不是 `\n`）
- 字体统一 `微软雅黑`
- 关键词高亮依赖 `【】` 标记 + 段落上下文红/蓝染色
- 遇到类似问题先查 `[codex-claude-tools]/INDEX.md` 索引复用现成工具；无匹配再新建，避免重复造轮子

## 4. 图表与复制粘贴坑位
- OLE图表粘贴后要断热链接（`CutCopyMode=False`）
- `CopyPicture` 使用 `xlPicture=-4147`（EMF）
- 删行前先删除图表，避免公式引用弹窗
- 两套图表机制不可混淆:
  - Pipeline `_write_chart`: 向模板已有chart注数
  - `Function_030.make_chart*`: Excel建图后OLE粘贴

## 5. 移植交付最小清单（Developer）
- 新建 `src/{template}_ppt.py`
- 确认 clone 页码 + shape清单
- 复用 Pipeline prompt 结果，不手工散落重写
- `Main.py` 中接入模板选择与调用
- 至少跑一次 `debug/test_src_smoke.py`（如存在）+ 端到端验证

## 6. fix2 执行优先级（精简）
1. 先清理陈旧引用（`codex_ppt.py -> yzr_ppt.py`）
2. 先补 src 冒烟测试，保护生产
3. 共享模块仅抽“纯数据工具”，视觉写入函数保持 yzr/zxh 独立
4. prompt 不做强共享；保留独立并标记 `prompt_src/synced_at`

## 7. 当前原则
- 不追求 orchestrator 覆盖 100% 模板
- Pipeline 追求通用能力；src 追求可交付与可控微调
- 优先交付，再做架构清债
