---
name: feedback_summary_sink
description: 多阶段 GPT 流水线累积——内层函数加 summary_sink: list | None = None 参数，外层订阅 completion
type: feedback
---

外层（如 `Main.py` 6.3 结论页）需要内层函数（如 `questionnaire_Excel` 多 runner 循环）每轮 GPT completion 时，给内层函数加 **`summary_sink: list | None = None`** 可选参数 + 内部 `summary_sink.append(mc_completion)`，外层传一个 list 进去就拿到全部累积。

**Why:**

1. **不破坏 return 签名** — `questionnaire_Excel` 已经 `return mc_sht, mc_slide`，加第三个 return 会冲击其他调用方（虽然当前只 1 处调用，但保持向后兼容是反射动作）。
2. **不破坏内层逻辑** — 累积只是"顺手 append"，不影响主流程；list 没传就完全跳过。
3. **解耦订阅 vs 不订阅** — 同一个函数既能在"只生成 PPT"场景跑，也能在"需要汇总"场景跑，调用方说了算。
4. **避免全局 state** — 把 `messages` 这种 module-level list 当订阅源是历史教训：跨 import / 跨调用很难追踪。显式参数让数据流可见。
5. **plan4 验证** — 6.3 结论页 prompt 必须显式注入"先前结论"作为 GPT 上下文（不能只靠 `messages` 累积，因为 `messages` 在某些分支可能被 reset）。`summary_sink` 让外层完整拿到 sheet 循环 + questionnaire 循环两批 completion，传给 `gen_result_prompt(sheet_summaries=..., questionnaire_summaries=...)`。

**How to apply:**

- **内层函数签名末尾加 `summary_sink: list | None = None`**（保持位置参数兼容；Python 类型提示在 `Function_030.py` 风格里可写可不写，写出来更清晰）
- **append 用 try/except 包**：`summary_sink.append(...)` 失败不能阻塞主流程
  ```python
  if summary_sink is not None:
      try:
          summary_sink.append(mc_completion)
      except Exception:
          pass
  ```
- **外层调用方在循环前 init list**：`all_questionnaire_summaries = []`，循环结束后整个 list 可用
- **跟 `gen_*_prompt` 的注入参数对齐**：被订阅函数产出的 list 直接喂给 `sheet_summaries=` / `questionnaire_summaries=` 这种命名参数，避免位置参数错位
- **不要塞太多内容进 sink** — 只 append 最终 completion 字符串，不要 append 中间状态 dict。Sink 是给外层"读结果"的，不是给内层"传状态"的

**反例（不要这样做）**：

- ❌ 改 return 签名加第三个值（破坏向后兼容，且内层函数有可能在循环中提前 return）
- ❌ 用 module-level global list 累积（跨调用混乱，难追踪谁在写）
- ❌ append dict 含 `mc_prompt + mc_completion + meta`（订阅方只要文本时变成解耦不彻底）

**Code anchor**：`src/Function_030.py::questionnaire_Excel(summary_sink=...)` + `Main.py` 顶部 `all_questionnaire_summaries = []` init + 调用 `questionnaire_Excel(..., summary_sink=all_questionnaire_summaries)` + 6.3 处的 `gen_result_prompt(..., questionnaire_summaries=all_questionnaire_summaries)`。
