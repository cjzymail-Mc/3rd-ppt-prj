# plan4 — 升级最终页总结（gen_result_prompt + Main.py【6.3】）

## Context

- 用户当前选中的 shape 是 **slide 11 的 `TextBox 5`**（Width=902, Height=152, Arial 16），由 `Main.py` `【6】结论部分` 流程生成。
- 写入路径：`Main.py:881` 调 `gen_result_prompt(sample_name)` → `Main.py:886` 调 `GPT_5(...)` → `Main.py:901` 通过 `Result_Bullet` 写入。
- 现状 `gen_result_prompt`（`src/Function_030.py:549-557`）只有 3 行指令、只插入 `sample_name`，**不传任何数据**，完全靠 `GPT_5` 维护的 `messages` 上下文。
- 用户反馈：
  1. prompt 有问题、结论散乱；
  2. 发送给 GPT 的内容混乱；
  3. 期望最终页总结涵盖 **优点 / 缺点 / 修改建议** 三段式。
- 这是 src/ 生产线（`Main.py` + `src/Function_030.py`），**不是 Pipeline**；只动这两个 + 必要时 `_ppt_shared.py`，不要碰 `pipeline/`。

---

## 根因清单

| 症状 | 根因 |
|--|--|
| 结论散乱无结构 | prompt 没要求三段式（优/缺/建议），GPT 自由发挥 |
| GPT 输入"混乱" | 全部依赖 `messages` 累积（数十 K tokens 的 Excel + 多名问卷），prompt 自身不画重点 |
| 没有关键词高亮 | 缺 `【】` 标注指令，无法被 `_apply_keyword_color` 染色 |
| 文本溢出 TextBox 5 | prompt 写"≤300 字"但实际可视空间约 ~150-200 字；写入前没 `clamp_text` |
| 单纯靠 messages 不稳健 | 若 `messages` 被 reset、或某个分支提前失败，summary 就是无据生成 |

---

## 决策（已与用户确认，2026-04-27）

1. **数据注入策略 = 显式注入先前结论（更鲁棒）**
   - `Main.py` 在每轮 GPT 调用后把 `mc_completion` 累积进一个 list；
   - `gen_result_prompt` 新增参数显式接收这些先前结论；
   - GPT 不再仅靠会话记忆，而是看到结构化、已浓缩好的素材。

2. **段内项目符号 = 段头不加 bullet、条目用 1/2/3**
   - `Result_Bullet` 默认每段都 ■，需要在写入后把段头行（如 `【优点】`）的 `Bullet.Visible` 单独设 0。

---

## 改动 1：`src/Function_030.py:549-557` 重写 `gen_result_prompt`

新签名：

```python
def gen_result_prompt(
    sample_name: str,
    sheet_summaries: list[str] | None = None,        # 各测试方法 sheet 的 GPT 结论
    questionnaire_summaries: list[str] | None = None, # 各运动员问卷反馈的 GPT 结论
) -> str:
```

`sheet_summaries` 默认 None → 退化成现状（兼容旧调用）；新调用方在 `Main.py` 里传值。

prompt 模板要点：
- 显式打包 sheet + 问卷的"先前结论"区块到 `【已知信息】`；
- 显式要求 **三段式**：`【优点】` / `【缺点】` / `【修改建议】`，每段 1/2/3 编号；
- 每段下面用 `1、2、3...` 编号条目；
- 强制 `【】` 标关键词（仅括词本身、不含标点）；
- 字数 ≤ **150 汉字**、行数 ≤ **9 行**（含 3 段头）；
- 若某段确实无内容 → 保留段头并写"暂无显著XX"；
- 只能基于已有数据，禁推测；
- 直接输出，不重述题面、不展示分析过程。

骨架伪代码：

```python
def gen_result_prompt(sample_name, sheet_summaries=None, questionnaire_summaries=None):
    sheet_block = ""
    if sheet_summaries:
        sheet_block = "\n".join(
            f"【测试方法 {i+1}】\n{s}"
            for i, s in enumerate(sheet_summaries) if s
        )
    qn_block = ""
    if questionnaire_summaries:
        qn_block = "\n".join(
            f"【运动员 {i+1}】\n{s}"
            for i, s in enumerate(questionnaire_summaries) if s
        )

    info_section = ""
    if sheet_block or qn_block:
        info_section = (
            f"【已知信息】\n"
            f"以下是先前已经分析得出的结论，请直接基于此做总结，"
            f"不要再展开重复分析：\n\n"
            f"{sheet_block}\n\n{qn_block}\n\n"
        )

    return (
        f"{info_section}"
        f"【你的任务】\n"
        f"请综合上述信息，对【{sample_name}】给出一份最终评测总结。\n\n"
        f"严格按以下三段式输出，每段用编号条目（1、2、3...）：\n"
        f"【优点】\n1、...\n2、...\n\n"
        f"【缺点】\n1、...\n2、...\n\n"
        f"【修改建议】\n1、...\n2、...\n\n"
        f"硬性要求：\n"
        f"- 每段至少 1 条；若确实无内容，保留段头并写\"暂无显著XX\"。\n"
        f"- 每条结论中把最核心的关键性能词用【】括起来"
        f"（仅括词本身、不含标点），后续会自动高亮。\n"
        f"- 总字数 ≤ 150 汉字，总行数 ≤ 9 行。\n"
        f"- 只能基于已有数据，不允许推测或编造。\n"
        f"- 直接输出结论，不重述题面、不展示分析过程。\n"
    )
```

---

## 改动 2：`Main.py` 累积先前结论 + 写入后处理

### 2A. 累积 sheet 循环的结论（约 line 540-712 区域）

在循环外初始化：

```python
all_sheet_summaries = []        # 各 sheet 的 GPT 结论
all_questionnaire_summaries = []  # 各运动员问卷的 GPT 结论
```

在 sheet 循环里现有的 `Result_Bullet(...)` 调用之前（line 697 附近）：

```python
all_sheet_summaries.append(mc_completion)
```

### 2B. 累积问卷循环的结论（约 line 779-813 区域）

类似地，问卷分支里调完 GPT 拿到 completion 之后追加：

```python
all_questionnaire_summaries.append(mc_completion)
```

### 2C. 改写 `【6.3】` 块（line 881-905）

```python
# 新签名调用
mc_prompt = gen_result_prompt(
    sample_name,
    sheet_summaries=all_sheet_summaries,
    questionnaire_summaries=all_questionnaire_summaries,
)

mc_completion = GPT_5(mc_prompt, model=mc_model)

# === 新增：写入前 clamp 字数/行数 ===
mc_completion = clamp_text(mc_completion, max_chars=200, max_lines=10)

# 旧逻辑：删 hint
hint_shape.tr.Text = '成功获取 [OPENAI] Reply ！'
time.sleep(1)
hint_shape.shape.Delete()

Left = 51
Top = 281
Text = Result_Bullet(mc_slide, Left, Top, mc_completion, scale=1)

# === 新增：段头行去掉 ■ bullet ===
_strip_bullet_on_section_headers(Text.tr)

# === 新增：自动按段落上下文红/蓝染色 + 去【】 ===
_apply_keyword_color(Text.shape)

# 保留旧逻辑：sample_name 红色
color_key(Text.tr, sample_name, red)
```

### 2D. 新增段头去 bullet 工具

放 `src/_ppt_shared.py`（与 `_apply_keyword_color` 同处，下文一并搬迁）：

```python
def _strip_bullet_on_section_headers(tr) -> None:
    """段头行（含【优点】/【缺点】/【修改建议】等）去掉 ■ bullet。

    Result_Bullet 默认每段都加 ■，但段头加 ■ 视觉冗余。
    """
    try:
        paragraphs = tr.Paragraphs()
        n = int(paragraphs.Count)
        for i in range(1, n + 1):
            p = tr.Paragraphs(i, 1)
            line = (p.Text or "").strip()
            # 段头识别：以【开头、含】，整行内容近似只是段标
            if line.startswith("【") and "】" in line and len(line) <= 10:
                p.ParagraphFormat.Bullet.Visible = 0
    except Exception:
        pass
```

---

## 改动 3：把 `_apply_keyword_color` 搬到 `src/_ppt_shared.py`

当前 `_apply_keyword_color` 在 `src/yzr_ppt.py:393`，`Main.py` 不该依赖一个 per-template 文件。**搬到 `_ppt_shared.py`**，`yzr_ppt.py` 改 import 它。

注意点：
- 搬迁时需要随之搬 `_RED / _BLUE / _BLACK`（`_ppt_shared.py:14-16` 已有）和 `_ADVANTAGE_MARKERS / _DISADVANTAGE_MARKERS`（`_ppt_shared.py:32-33` 已有），只搬函数即可，常量已经在 shared 里。
- `_DISADVANTAGE_MARKERS` 已包含 `"修改建议"`（`src/_ppt_shared.py:33`），所以 `【修改建议】` 段会被识别为 disadvantage 并染蓝，**符合预期**，无需扩词表。

---

## 关键文件清单

| 文件 | 改动 |
|--|--|
| `src/Function_030.py` | 重写 `gen_result_prompt` 接收两个 list 参数（lines 549-557） |
| `Main.py` | (a) 顶部 import 三个工具；(b) 循环外初始化两个累积 list；(c) sheet 循环和问卷循环各 `.append(mc_completion)` 一行；(d) 【6.3】块按改动 2C 重写 |
| `src/_ppt_shared.py` | 新增 `_apply_keyword_color`（从 yzr_ppt 搬迁）+ 新增 `_strip_bullet_on_section_headers` |
| `src/yzr_ppt.py` | 删本地 `_apply_keyword_color`，改 `from ._ppt_shared import _apply_keyword_color` |

---

## 复用清单（不要重复造）

| 工具 | 位置 | 用途 |
|--|--|--|
| `_ADVANTAGE_MARKERS` / `_DISADVANTAGE_MARKERS` | `src/_ppt_shared.py:32-33` | "修改建议" 已在 disadvantage 列表 ✓ |
| `clamp_text(text, max_chars, max_lines)` | `src/_ppt_shared.py:411` | 剔空行 + 收口字数/行数 |
| `_apply_keyword_color(shp)` | `src/yzr_ppt.py:393`（待搬） | 段落上下文 → 优势段红色 / 劣势段蓝色 + 去 `【】` |
| `Result_Bullet(slide, L, T, text, scale)` | `src/Class_030.py:545` | 已用，不变 |
| `color_key` | `Function_030` | 把 sample_name 染红，保留 |

---

## 验证方式（端到端）

1. **正常路径**：`python Main.py` 跑完整流程到 slide 11；
   - TextBox 5 应有三段：`【优点】 / 【缺点】 / 【修改建议】`；
   - 文本不溢出 152 高度；
   - 段头无 ■ bullet，条目有 ■ bullet；
   - 优点段关键词红色加粗、缺点 / 建议段关键词蓝色加粗；
   - sample_name 整体红色（旧逻辑保留）。

2. **隔离调试**（推荐）：在 `gen_result_prompt` 写一个 `__main__`，用 mock 的 `sheet_summaries` / `questionnaire_summaries` 调一次 GPT，肉眼看输出格式是否符合三段式。

3. **边界用例**：
   - `sheet_summaries=[]` 且 `questionnaire_summaries=[]` → 退化成"仅靠题面"，应仍能输出占位"暂无显著XX"；
   - GPT 偶吐过长 → `clamp_text` 应卡到 200 字 / 10 行内；
   - 反馈无明显缺点 → 应输出 `【缺点】\n暂无显著缺点`；

4. **回归用例**：跑一次 yzr 模板（`python src/yzr_ppt.py`），确认 `_apply_keyword_color` 搬位之后 yzr 流程未坏。

---

## 不在本次范围

- 不改 `gen_questionnaire_prompt` / `gen_mc_prompt` 自身（只读它们的 completion）；
- 不改 Pipeline 端的 `gpt_summary.md`；
- 不改 `Result_Bullet` 类本身（只在外部对其 TextRange 后处理）；
- 不调整 TextBox 5 在模板（slide 10）上的 Width/Height/Top/Left。

---

## 实际落地与 plan 偏差（2026-04-27 完工记录）

plan4 落地后又叠加了 todays-task 的两轮迭代，最终偏离 plan 的关键点：

### 1. 染色方案：从单一【】 → bracket-typed

- **plan 原版**：所有关键词统一用 `【keyword】` 标记，由 `_apply_keyword_color` 按 section context 染色（优点段染红、缺点段染蓝）
- **实际落地**：todays-task 改为 **bracket-typed**——`<>` 红+粗（优点）、`[]` 蓝+粗（缺点）、`(...)` 仅粗（修改建议）；新建 `_apply_conclusion_color`（不复用 `_apply_keyword_color`），中文 `【】` 保留给 section header
- **触发原因**：用户 todays-task 反馈"染色逻辑错误"——单一 `【】` 标记 + section context 染色在"同一 shape 内多段多色"场景下，跨段引用同一关键词会错染

### 2. 字数/行数预算：200/10 → 280/13

- **plan 原版**：`clamp_text(max_chars=200, max_lines=10)`，prompt 限定 ≤150 字 / ≤9 行
- **实际落地**：扩到 `clamp_text(max_chars=280, max_lines=13)`，prompt 限定 ≤270 字 / ≤12 行
- **触发原因**：用户 todays-task 反馈"文字总长度太短了"——TextBox 5 模板高度 152 不是硬上限，`Class_030.Text_Box` 没设 `tf.AutoSize=0`，PPT 默认 `msoAutoSizeShapeToFitText` 接管，shape 自动撑高
- **新认知（已写入 CLAUDE.md 硬规则）**：`Result_Bullet` / `Text_Box` 子类自动 auto-grow，`clamp_text` 上限不受模板 shape 高度束缚（仅受 slide 高度束缚）

### 3. 共享工具搬迁

- **plan 原版**：搬 `_apply_keyword_color` 到 `_ppt_shared.py`
- **实际落地**：除了 `_apply_keyword_color`，还新增 `_apply_conclusion_color` + `_strip_bullet_on_section_headers` 一并放 `_ppt_shared.py`（两套染色函数适用场景互补，都需要共享）

### 4. summary_sink 模式

- **plan 原版**：直接在 sheet 循环 / 问卷循环里 `all_*_summaries.append(mc_completion)`
- **实际落地**：sheet 循环在 `Main.py` 内直接 append；但问卷循环在 `questionnaire_Excel`（Function_030）内部，**给函数加了 `summary_sink: list | None = None` 参数** + 内部 append，外层传 list 进去订阅
- **沉淀**：见 `.claude/memory/feedback_summary_sink.md`

### 5. 弹窗样式（追加任务，超出 plan4 范围）

- 同一会话用户追加要求"优化弹窗排版字体"，对 `_ask_with_countdown` 做了 iOS systemGroupedBackground 风格重写
- **沉淀**：见 `.claude/memory/feedback_popup_ui.md`
