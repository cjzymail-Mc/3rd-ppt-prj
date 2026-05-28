# Memory Index

## User
- [user_profile.md](user_profile.md) — 用户角色、技术背景、协作偏好

## Feedback
- [feedback_hybrid_workflow.md](feedback_hybrid_workflow.md) — Pipeline先行+LLM精调，不要纯LLM硬啃
- [feedback_com_constraints.md](feedback_com_constraints.md) — 必须用COM，禁openpyxl/python-pptx/numpy + COM规范表（含 \n→\r、PNG加密绕过、字体显式）
- [feedback_output_style.md](feedback_output_style.md) — 只说结论，不展示diff
- [feedback_chart_write.md](feedback_chart_write.md) — 分发场景chart强制从零制表，禁用_write_chart原位改（fix4）
- [feedback_debug_protocol.md](feedback_debug_protocol.md) — COM/OLE/模板 bug 调试流程：grep优先、质疑约定、2次失败熔断（fix3→fix4 血的教训）
- [feedback_workflow_routing.md](feedback_workflow_routing.md) — 5阶段工作流路由：Pipeline首跑→评估→/developer→主Claude兜底（plan3 定稿）
- [feedback_conclusion_coloring.md](feedback_conclusion_coloring.md) — 6.3 结论页 bracket-typed 染色：<>红 / []蓝 / ()仅粗体；【】专给 section header
- [feedback_popup_ui.md](feedback_popup_ui.md) — tk 弹窗约定：iOS systemGroupedBackground + 白卡 + Indigo 强调；不走 header band + CTA 路线
- [feedback_summary_sink.md](feedback_summary_sink.md) — 多阶段 GPT 累积模式：summary_sink: list | None 可选参数订阅 completion
- [feedback_debug_entry.md](feedback_debug_entry.md) — xxx_ppt.py 需有 __main__ 单页调试入口，3 层 import fallback
- [feedback_shape_finetune.md](feedback_shape_finetune.md) — fine_tuned 四参数完整，基准值从标准模板读取
- [feedback_skip_vs_clear.md](feedback_skip_vs_clear.md) — SHAPES skip 策略不清模板预置文字；移植新模板必查源 shape 是否预留空（apparel-fix1）
- [feedback_unit_normalize_bmi.md](feedback_unit_normalize_bmi.md) — 100KG 实为 100 斤等填错，粗修 m→cm + 斤→kg 后 BMI∈[16,32] 交叉验证（apparel-fix1）
- [feedback_python_stdout_encoding.md](feedback_python_stdout_encoding.md) — Bash 工具跑 python 输出中文要 PYTHONIOENCODING + io.TextIOWrapper 双保险；chcp/set 是 cmd 语法在 git bash 失效
- [feedback_perf_rewrite_validate.md](feedback_perf_rewrite_validate.md) — 重写老 selection 链路前先 print 边界单元格 diff，end('up'/'down') 在连续非空区会越过表头（test_detail off-by-one 教训）
- [feedback_acceptance_gate.md](feedback_acceptance_gate.md) — PPT 交付必过 ppt-acceptance-check（L0+L1+L4）；**责任分离（2026-05-27）**：developer 只落 trace + 契约就绪，主 Claude 编排者跑验收 + 判读 report（防自审绕道）

## Project
- [project_4agent_architecture.md](project_4agent_architecture.md) — 4-Agent 混合工作流 + fix_type/output_contract/COM DispatchEx（2026-03-17/19）
- [project_workflow_modes.md](project_workflow_modes.md) — 冷启动/热迭代流程图 + 版本追溯 + fix_type 分流表
- [project_orchestrator_upgrade.md](project_orchestrator_upgrade.md) — Python orchestrator 路线确认（2026-04-01 决策档案）
- [project_plan3_architecture.md](project_plan3_architecture.md) — plan3 重构：3+1 step-based agents + 局部自检循环（2026-04-09，supersedes 4agent）
- [project_step3_feedback_loop.md](project_step3_feedback_loop.md) — Step3→Step2 反馈循环 + Excel prompt 同步预检（2026-04-10）

## Reference
- [reference_manual_pipeline.md](reference_manual_pipeline.md) — 手动Pipeline命令 + 用户批注字段说明
- [reference_pipeline_repair.md](reference_pipeline_repair.md) — Pipeline 代码修复指引：文件清单 + fix 类型 + 技术栈约束 + 自检要求
- [reference_3account_junction.md](reference_3account_junction.md) — 3 账号 auto-memory junction 架构 + 维护警告
- [reference_agent_memory_design.md](reference_agent_memory_design.md) — Claude Code agent 隔离机制 5 维分析 + per-agent memory 决策档案（不引入）
