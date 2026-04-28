---
name: feedback_conclusion_coloring
description: 6.3 结论页 bracket-typed 染色——GPT 用 <>/[]/() 标关键词，由 _apply_conclusion_color 按括号类型决定红/蓝/仅粗体
type: feedback
---

6.3 最终结论页（`Main.py` 末尾的总结 TextBox）的染色路线：**GPT 用 ASCII 半角括号自标关键词，外层按括号类型决定颜色**，不要复用 `_apply_keyword_color`。

| 标记 | 段位 | 视觉 |
|--|--|--|
| `<keyword>` | 【优点】 | 红 (#FF0000) + 加粗 |
| `[keyword]` | 【缺点】 | 蓝 (#F0B400 light_blue) + 加粗 |
| `(keyword)` | 【修改建议】 | 仅加粗，不染色 |

中文 **【】** 保留给 section header（如 `【优点】` / `【缺点】` / `【修改建议】`），由 `_strip_bullet_on_section_headers`（`_ppt_shared.py`）单独识别去 ■。

**Why:** `_apply_keyword_color` 的逻辑是"扫描段落上下文（看到'优点'切到 advantage 模式 → 后续 【】 染红）"，假设 **整个 shape 只有一个段落基调**。yzr/zxh 各 shape 是这种情况（一个 shape 写优点 / 另一个 shape 写缺点）。但 6.3 结论页是 **同一 shape 内三段全有**——section context 跟踪虽然能跑，但 GPT 只用单一标记 `【】` 时，"优点段中提到的缺点关键词"会被错染。bracket-typed 把"该词属于哪类"的语义直接编码进标记本身，与段头位置解耦。todays-task 用户原话："新建一个【结论染色函数】来处理三种染色：优点红色、缺点蓝色、建议加粗即可无需染色。"

**How to apply:**

- **6.3 结论页**：用 `_apply_conclusion_color(Text.shape)`（不要用 `_apply_keyword_color`）
- **yzr/zxh 模板 per-shape**：继续用 `_apply_keyword_color`（GPT prompt 里仍用 `【】` 单一标记，section context 染色）
- **GPT prompt 写作**：bracket scheme 的 prompt 必须显式列三条规则 + 例子（`例如：<回弹性能>`），并强调"仅包词本身、不含标点"。`gen_result_prompt`（`Function_030.py`）已落地范例
- **clamp_text 顺序**：先 `clamp_text` 再 `_apply_conclusion_color`，否则截断可能砍掉一半的 `<...>` 留下孤悬括号
- **Result_Bullet 写入后顺序**：`_strip_bullet_on_section_headers(Text.tr)` → `_apply_conclusion_color(Text.shape)` → 旧逻辑 `color_key(sample_name, red)`（保留 sample_name 整体染红）

**Code anchor**：`src/_ppt_shared.py::_apply_conclusion_color` + `src/Function_030.py::gen_result_prompt` + `Main.py` 【6.3】 块。
