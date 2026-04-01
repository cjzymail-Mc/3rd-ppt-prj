# Plan3: Prompt 冗余重构 + 文本截断修复

## 问题分析

### 问题 1: Prompt 反复优化后越来越冗余

**现象**：每轮 Builder LLM 都在原 prompt 尾部追加约束（"必须包含XX关键词"、"字数控制在XX字"...），多轮后 prompt 变成一堆互相矛盾的补丁堆砌。

**根因**：`_builder_prompt_optimizer_prompt()` (orchestrator.py L675-698) 中的指令明确写了"追加"：
```
keyword_missing → 在 prompt 中追加关键词要求
budget_overflow → 在 prompt 中追加字数上限约束
```
Builder LLM 被训练为只加不删，永不重构。

**根因链**：
```
Round 1: 原始 prompt → GPT 生成 → 验收失败(缺"建议")
Round 2: Builder 追加 "必须包含建议" → 但原 prompt 仍说"末尾给出建议" → 两条指令冲突
Round 3: Builder 再追加 "控制在200字" → 但之前追加的"不少于250字"仍在 → 三条矛盾
...
Round N: prompt 变成指令垃圾场，GPT 产出反而更差
```

**答：不需要扩大 Reviewer 权限**。Reviewer 的职责是分析失败+给建议，它做得没问题。问题在 Builder 的编辑策略（追加 vs 重写）。

### 问题 2: 最终 PPT 文本被粗暴截断

**现象**：GPT 生成的完整总结被 `clamp_text(t[:max_chars])` 硬切，句子断裂、关键词丢失。

**根因**：`clamp_text()` (ppt_pipeline_common.py L346-354) 执行 `t[:max_chars]` 硬截断。

**根因链**：
```
prompt 告诉 GPT "总字数控制在270字左右"
  → GPT 输出 ~300 字（"左右"理解偏宽）
    → clamp_text(t[:270]) 硬切到 270 字
      → 句子断裂，末尾的"建议"关键词被切掉
        → 04 验收: semantic_coverage=66.67%
          → fix.md: "追加必须包含'建议'"
            → Builder 在 prompt 追加约束
              → GPT 再次输出 ~300 字，"建议"仍在末尾
                → 又被截断 → 无限循环
```

**用户立场**：
- prompt 中保留精确字数限制（引导 GPT 产出合理长度）
- 最终文本不硬截断（信息完整性 > 精确字数匹配）
- "我可以接受文字超过字数，但 prompt 中还是要用精确字数来限制"

**readability 影响评估**：
当前 readability 公式 `len_score = min(1.0, len_b / (len_a * 0.5))` 只惩罚文本过短（<50%），不惩罚过长。因此移除截断不会影响 readability 分数。验证：template=225字, generated=300字 → 300/(225*0.5)=2.67 → capped=1.0 → 满分。

---

## 改动规划

### 改动 1: Builder LLM — 从"追加"到"全面重写"

**文件**: `orchestrator.py` — `_builder_prompt_optimizer_prompt()` 方法 (L675-698)

**当前问题**：
- 指令用"追加"语气，Builder 只加不删
- 没有提供原始 prompt 模板作为参考，Builder 无法判断什么是"干净的 prompt"
- 没有附带当前生成内容和 budget 信息，Builder 缺乏上下文

**改为**：
```python
def _builder_prompt_optimizer_prompt(self, sheet_name: str, fix_data: list[dict]) -> str:
    """Builder LLM: 根据 fix 报告全面重写 GPT-prompt Text。"""

    # 1. 读取原始 prompt 模板（clean baseline）
    tpl_path = self.project_root / "pipeline" / "prompt_templates" / "gpt_summary.md"
    original_template = ""
    if tpl_path.exists():
        original_template = tpl_path.read_text(encoding="utf-8")[:800]  # 截取核心部分

    # 2. 读取当前生成内容（供 Builder 了解实际产出）
    content_path = self.project_root / "pipeline-progress" / "03a-build_shape_content.json"
    content_snippets = {}
    if content_path.exists():
        try:
            data = json.loads(content_path.read_text(encoding="utf-8"))
            for item in data.get("items", []):
                sn = item.get("shape_name", "")
                content_snippets[sn] = item.get("content", "")[:200]
        except Exception:
            pass

    # 3. 构建 fix 清单（附带当前内容片段）
    fix_lines = []
    for fx in fix_data:
        if fx.get("fix_type") == "code":
            continue
        shape = fx["shape"]
        line = f"  - {shape}: [{fx.get('fix_type','')}] {fx.get('issue','')}"
        line += f"\n    建议: {fx.get('suggestion','')}"
        if shape in content_snippets:
            line += f"\n    当前生成内容(前200字): {content_snippets[shape]}"
        fix_lines.append(line)
    fix_items = "\n".join(fix_lines)

    return (
        f"你是 PPT 内容优化师。根据验收报告，**全面重写** Excel 中的 GPT prompt。\n\n"
        f"## 核心原则\n"
        f"**不要在原 prompt 上打补丁！** 每次都基于原始模板 + fix 建议，重新编写一个干净、完整的 prompt。\n"
        f"原因：反复追加约束会导致 prompt 冗余膨胀、指令矛盾，GPT 产出质量反而下降。\n\n"
        f"## 原始 prompt 模板（clean baseline）\n"
        f"```\n{original_template}\n```\n\n"
        f"## 工作 sheet: 「{sheet_name}」\n\n"
        f"## 验收失败清单\n{fix_items}\n\n"
        f"## 重写策略\n"
        f"1. 先读取当前 prompt 全文\n"
        f"2. 对照原始模板，识别哪些是有效约束、哪些是冗余补丁\n"
        f"3. 以原始模板为骨架，融入 fix 建议中的有效约束，生成一个干净的新 prompt\n"
        f"4. 具体 fix_type 处理：\n"
        f"   - keyword_missing: 在 prompt **前半段**明确要求包含缺失关键词（不要放在末尾）\n"
        f"   - budget_overflow: 降低目标字数（当前 budget 的 85%），让 GPT 输出更紧凑\n"
        f"   - budget_underflow: 提高字数下限，要求充实内容\n"
        f"   - style_mismatch: 加入格式/语调约束\n\n"
        f"## 任务\n"
        f"1. 通过 Python COM 打开 `pipeline-progress/01-shape_detail.xlsx`「{sheet_name}」sheet\n"
        f"2. 找到上述 shape 的「GPT-prompt Text」单元格\n"
        f"3. **全面重写** prompt（不是追加！），保持干净、无冗余\n"
        f"4. 保存并关闭 xlsx\n"
        f"5. 打印修改摘要（列出每个 shape 的改动要点）\n\n"
        f"## 规则\n"
        f"- 只改有问题的 shape 的 prompt，其余不动\n"
        f"- **不要修改「内容描述」「strategy」「params」等注释字段**\n"
        f"- ⚠️ 不要运行任何 pipeline 脚本\n"
    )
```

**关键变化**：
- "追加" → "全面重写"
- 提供原始 prompt 模板（clean baseline）
- 提供当前生成内容片段（上下文）
- 明确重写策略："以原始模板为骨架，融入有效约束"

### 改动 2: Builder Agent 定义更新

**文件**: `.claude/agents/02-builder.md`

在"你的唯一任务"步骤 3 修改：

当前：
```
3. 根据 fix 建议，修改有问题的 shape 的 prompt 文本：
   - 添加关键词要求（如「必须融入'建议'一词」）
   - 调整字数约束（如「控制在 180-220 字」）
   - 修正风格偏差（如「严格按照参考文本的格式和语调」）
```

改为：
```
3. 根据 fix 建议，**全面重写**有问题的 shape 的 prompt 文本：
   - ⚠️ 不要在原 prompt 上追加补丁！基于 orchestrator 提供的原始模板重写
   - 将 fix 建议中的有效约束融入新 prompt，保持干净、无冗余
   - 如果多条 fix 建议有冲突，以最新一条为准
```

### 改动 3: clamp_text 移除字符硬截断

**文件**: `pipeline/ppt_pipeline_common.py` L346-354

当前：
```python
def clamp_text(text: str, max_chars: int, max_lines: int) -> str:
    """Hard-truncate text to budget constraints."""
    t = safe_text(text)
    if max_chars > 0 and len(t) > max_chars:
        t = t[:max_chars]          # ← 硬切句子
    if max_lines > 0:
        lines = t.splitlines() or [t]
        t = "\n".join(lines[:max_lines])
    return t
```

改为：
```python
def clamp_text(text: str, max_chars: int, max_lines: int) -> str:
    """Soft-clamp: 只限制行数（保护PPT版面），不截断字符（保护信息完整性）。

    字符限制通过 prompt 引导 GPT 控制，不在后处理中强制执行。
    """
    t = safe_text(text)
    # 仅行数限制（防止 PPT 版面溢出）
    if max_lines > 0:
        lines = t.splitlines() or [t]
        t = "\n".join(lines[:max_lines])
    return t
```

**原理**：
- 字符限制 → 由 prompt "总字数控制在X字左右" 引导 GPT，不在代码中强制
- 行数限制 → 保留，保护 PPT shape 版面不溢出
- `max_chars` 参数保留（不改函数签名），但函数内部不使用

**影响**：
- 03a 调用 clamp_text 后，文本可能超过 max_chars → `valid` 标记为 False → 仅供记录，不阻断流程
- PPT 中会出现完整文本，关键词不再被切掉
- readability 不受影响（公式只惩罚过短，不惩罚过长）

### 改动 4: Prompt 模板关键词前置

**文件**: `pipeline/prompt_templates/gpt_summary.md` L18 / L40

当前：
```
- 结论中请自然融入：'样本'（如'本次{n}名样本'）、'反馈'（如'样本反馈'）、'建议'（末尾给出改进建议）
```

改为：
```
- 【关键词要求】开头第一句必须同时包含'样本'和'反馈'（如"本次{n}名样本反馈显示"），正文中自然融入'建议'（如"建议关注..."）
```

同步更新硬编码 fallback：`pipeline/03a_build_shape.py` L470 的同一行。

**原理**：即使文本略超 budget，关键词都在前半段，不会丢失。"末尾给出改进建议" → "正文中自然融入'建议'"，避免关键词集中在尾部。

### 改动 5（可选）: Reviewer 检测 prompt 膨胀

**文件**: `.claude/agents/03-reviewer.md`

在 Step 2 增加一项检查：

```
- **prompt 膨胀检测**：如果 GPT-prompt Text 超过 500 字或包含 3 条以上"必须包含"/"控制在"类约束，
  在 fix 建议中标注 fix_type="prompt_bloat"，建议 Builder 全面重构 prompt
```

**说明**：这是辅助手段。核心修复在改动 1（Builder 默认全面重写），此项仅作为 Reviewer 的额外诊断信号。如果不想增加 Reviewer 复杂度，可以跳过。

---

## 改动量汇总

| 文件 | 改动量 | 说明 |
|------|--------|------|
| `orchestrator.py` | 大 (~50行) | `_builder_prompt_optimizer_prompt()` 全面重写 |
| `.claude/agents/02-builder.md` | 小 (~5行) | 编辑策略：追加→重写 |
| `pipeline/ppt_pipeline_common.py` | 小 (~3行) | clamp_text 移除字符截断 |
| `pipeline/prompt_templates/gpt_summary.md` | 小 (2行) | 关键词前置 |
| `pipeline/03a_build_shape.py` | 小 (1行) | fallback prompt 同步 |
| `.claude/agents/03-reviewer.md` | 小 (可选) | prompt 膨胀检测 |

## 实施顺序

1. `ppt_pipeline_common.py` — 移除截断（独立，无依赖）
2. `prompt_templates/gpt_summary.md` + `03a_build_shape.py` — 关键词前置（独立）
3. `orchestrator.py` — Builder LLM prompt 全面重写（核心改动）
4. `02-builder.md` — Agent 定义同步（依赖改动3的设计）
5. `03-reviewer.md` — prompt 膨胀检测（可选，独立）

## 验证

1. `python -m py_compile pipeline/ppt_pipeline_common.py pipeline/03a_build_shape.py orchestrator.py`
2. 选 0 初始化 → 检查生成内容：文本应完整，不再有半截句子
3. 选 5 验收 → 语义覆盖率应提升（关键词前置 + 不再截断）
4. 选 2 多轮 → 检查 Builder LLM 是否"重写"而非"追加"（观察 prompt 长度是否稳定）

## 与原 Plan3 对比

| 原 Plan3 | 更新后 | 变化原因 |
|----------|--------|---------|
| clamp_text 智能截断（句末断句） | 直接移除字符截断 | 用户"可以接受超字数" |
| long_summary budget 1.5x | 删除（不需要） | 不截断则 budget 仅用于 prompt 引导，当前 1.2x 作为引导值合理 |
| prompt 关键词前置 | **保留** | 仍有保险价值 |
| 04 截断诊断 | 删除（不需要） | 不截断则不存在"截断丢失 vs 未生成"的区分问题 |
| Builder LLM prompt 增强 | **升级为全面重写** | 用户新需求：解决 prompt 冗余膨胀 |
| — | Builder Agent 定义更新 | 新增 |
| — | Reviewer prompt 膨胀检测 | 新增（可选） |
