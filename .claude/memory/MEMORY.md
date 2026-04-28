# Memory Index

## User
- [user_profile.md](user_profile.md) — 用户角色、技术背景、协作偏好

## Feedback
- [feedback_hybrid_workflow.md](feedback_hybrid_workflow.md) — Pipeline先行+LLM精调，不要纯LLM硬啃
- [feedback_com_constraints.md](feedback_com_constraints.md) — 必须用COM，禁openpyxl/python-pptx/numpy + COM规范表
- [feedback_output_style.md](feedback_output_style.md) — 只说结论，不展示diff
- [feedback_gpt_sparse_data.md](feedback_gpt_sparse_data.md) — GPT数据稀疏时截断，需维度引导+字数下限+强硬关键词
- [feedback_chart_write.md](feedback_chart_write.md) — 分发场景chart强制从零制表，禁用_write_chart原位改（fix4）
- [feedback_debug_protocol.md](feedback_debug_protocol.md) — COM/OLE/模板 bug 调试流程：grep优先、质疑约定、2次失败熔断（fix3→fix4 血的教训）
- [feedback_workflow_routing.md](feedback_workflow_routing.md) — 5阶段工作流路由：Pipeline首跑→评估→/developer→主Claude兜底（plan3 定稿）
- [feedback_conclusion_coloring.md](feedback_conclusion_coloring.md) — 6.3 结论页 bracket-typed 染色：<>红 / []蓝 / ()仅粗体；【】专给 section header
- [feedback_popup_ui.md](feedback_popup_ui.md) — tk 弹窗约定：iOS systemGroupedBackground + 白卡 + Indigo 强调；不走 header band + CTA 路线
- [feedback_summary_sink.md](feedback_summary_sink.md) — 多阶段 GPT 累积模式：summary_sink: list | None 可选参数订阅 completion

## Project
- [project_4agent_architecture.md](project_4agent_architecture.md) — 5-Agent混合工作流 + fix_type/output_contract/COM DispatchEx
- [project_workflow_modes.md](project_workflow_modes.md) — 冷启动/热迭代流程图 + 版本追溯 + fix_type分流表

## Reference
- [reference_manual_pipeline.md](reference_manual_pipeline.md) — 手动Pipeline命令 + 用户批注字段说明
