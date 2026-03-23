# claude-plan.md — 全自动模式优化 + 选项4

## 目标

让 `orchestrator.py` 支持"选项4: max=3 全自动跳过所有暂停"，并优化上游生成精度 + 下游修正精度，确保 3 轮内自动收敛通过三层门禁。

## 实现顺序

1. **03a: output_contract 注入 GPT prompt**（首轮命中率↑）
2. **02b: 量化修正 + 约束替换**（修正精度↑，防震荡）
3. **01b: fallback 默认内容描述**（消除全自动盲区）
4. **02: gpt_prompted 默认 contract**（确保验收覆盖）
5. **orchestrator: 选项4 全自动模式**（流程控制）

---

## 改动 1: 03a — output_contract 注入 GPT prompt

### 文件: `pipeline/03a_build_shape.py`

### 问题

02 已经从内容描述解析出 `output_contract`（required_keywords / bracket_highlight / ratio_required），但 03a 组装 GPT prompt 时**完全没用它**。contract 只在 04 验收时使用。GPT 不知道自己必须满足什么合约。

### 改动

#### 1A. `_build_rich_prompt()` 读取 output_contract 并生成约束段

在 `_build_rich_prompt()` 函数签名中新增参数 `output_contract: dict = None`。

在函数体中（`user_section` 构建之后、template format 之前），生成合约约束段：

```python
# Build output contract section
contract_section = ""
if output_contract:
    lines = []
    kw = output_contract.get("required_keywords")
    if kw:
        lines.append(f"- 必须包含关键词: {'、'.join(kw)}")
    if output_contract.get("bracket_highlight"):
        lines.append("- 关键性能词用【】括起（仅括词本身，不含标点）")
    if output_contract.get("ratio_required"):
        lines.append("- 每段结论后注明 (X/N) 比例")
    if lines:
        contract_section = "\n【输出合约 — 必须满足】\n" + "\n".join(lines) + "\n"
```

将 `contract_section` 拼入 prompt（放在 `user_section` 之后）。同时更新 template format 调用，传入新变量。

#### 1B. 所有调用 `_build_rich_prompt()` 的地方传入 output_contract

在 `build_content()` 中，从 mapping 数据（`m.get("output_contract", {})`）获取 contract，传给 `_build_rich_prompt()`。

需要修改 `build_content()` 签名，新增 `output_contract: dict = None` 参数。

涉及 3 处调用：
- L550-553: exact dispatch `gpt_prompted`
- L624-627: hint-based `gpt_prompted`
- L677-682: role-based `long_summary/body/insight`

#### 1C. `_build_all()` 循环中读取 output_contract 并传入

在 `_build_all()` 的 for 循环中：
```python
oc = m.get("output_contract", {})
```

传入 `build_content(..., output_contract=oc)`。

#### 1D. prompt template 更新（可选）

如果使用外部模板 `gpt_summary.md`，在模板中添加 `{contract_section}` 占位符。
如果模板中没有此占位符，fallback 硬编码 prompt 中直接拼入即可。

### 验证

```bash
python -m py_compile pipeline/03a_build_shape.py
# 运行 03a --assemble-only，检查 03a-pending_prompts.json 中的 prompt 是否包含【输出合约】段
```

---

## 改动 2: 02b — 量化修正 + 约束替换

### 文件: `pipeline/02b_iteration_setup.py`

### 问题

1. 修正指令只给方向（"宁短勿长"），不给量化目标
2. 约束只追加不替换，多轮后可能出现矛盾（"宁短勿长" + "内容需更充实"）

### 改动

#### 2A. 从 04-diff_result.json 读取实际数值

当前 04 的 fix 条目已包含实际字数信息，格式如：
```
"issue": "readability=85 < 95 (文本过长: 320/200字)"
```

在 `apply_annotation_fixes()` 中，解析 issue 字段提取实际字数和目标字数：

```python
import re

def _extract_char_counts(issue: str) -> tuple[int, int]:
    """从 issue 描述中提取 (实际字数, 目标字数)。"""
    m = re.search(r"(\d+)/(\d+)字", issue)
    if m:
        return int(m.group(1)), int(m.group(2))
    return 0, 0
```

#### 2B. 量化修正指令

将当前的方向性 suffix 改为包含具体数字：

```python
if ft == "keyword_missing":
    suffix = "。必须包含'样本'、'建议'、'反馈'关键词"
elif ft == "budget_overflow":
    actual, target = _extract_char_counts(fx.get("issue", ""))
    if actual and target:
        lo = int(target * 0.9)
        hi = int(target * 1.1)
        suffix = f"。字数严格控制在{lo}-{hi}字（当前{actual}字，目标{target}字）"
    else:
        suffix = "。严格控制总字数，宁短勿长"
elif ft == "budget_underflow":
    actual, target = _extract_char_counts(fx.get("issue", ""))
    if actual and target:
        lo = int(target * 0.9)
        hi = int(target * 1.1)
        suffix = f"。内容需更充实，字数控制在{lo}-{hi}字（当前{actual}字，目标{target}字）"
    else:
        suffix = "。内容需要更充实，涵盖更多测试者反馈"
elif ft == "style_mismatch":
    suffix = "。严格按照参考文本的格式和语调"
```

#### 2C. 约束替换（而非追加）

当前逻辑：`if suffix and suffix not in existing: new_val = f"{existing}{suffix}"`

改为：先移除上一轮的同类约束，再追加新约束。

识别上一轮约束的方式：通过前缀模式匹配。定义约束模式：

```python
# 约束模式 — 同类约束只保留最新版本
_CONSTRAINT_PATTERNS = {
    "budget_overflow":  re.compile(r"[。；;]\s*(?:字数严格控制|严格控制总字数|宁短勿长)[^。；;]*"),
    "budget_underflow": re.compile(r"[。；;]\s*(?:内容需[要更]充实|字数控制在)[^。；;]*"),
    "keyword_missing":  re.compile(r"[。；;]\s*必须包含[^。；;]*关键词[^。；;]*"),
    "style_mismatch":   re.compile(r"[。；;]\s*严格按照参考文本[^。；;]*"),
}
```

在追加新 suffix 之前，用对应模式清除 existing 中的旧约束：

```python
pattern = _CONSTRAINT_PATTERNS.get(ft)
if pattern:
    existing = pattern.sub("", existing)
```

### 验证

```bash
python -m py_compile pipeline/02b_iteration_setup.py
# 模拟场景：手动设置 04-diff_result.json 包含 budget_overflow fix（320/200字）
# 运行 02b，检查 xlsx 中的内容描述是否包含 "字数严格控制在180-220字（当前320字，目标200字）"
```

---

## 改动 3: 01b — fallback 默认内容描述

### 文件: `pipeline/01b_auto_annotate.py`

### 问题

`infer_annotation()` 末尾的 fallback（L175）返回 `desc = "（必填）"`。在全自动模式下，这些 shape 没有有效内容描述，GPT prompt 质量最差。

### 改动

#### 3A. 为 fallback shape 生成基于 role 的默认描述

当前 fallback 代码（函数末尾）：

```python
# Fallback: non-trivial text that doesn't match rules
desc = "（必填）"
return {"description": desc, "strategy": strategy, "params": params}
```

改为根据 text 长度和内容特征给出合理默认值：

```python
# Fallback: non-trivial text that doesn't match rules → auto-classify
if len(text) > 60 and ("【" in text or "】" in text or "（" in text):
    # 长文本有结构化标记 → 大概率需要 GPT 生成
    desc = ("从补充说明总结要点。"
            "必须包含'建议'、'反馈'、'样本'关键词，"
            "用【】括起关键性能词，每段结论后注明(X/N)比例")
    strategy = "gpt_prompted"
    params = "source=补充说明"
elif len(text) > 30:
    # 中等长度文本 → GPT 基础生成
    desc = "从补充说明总结要点"
    strategy = "gpt_prompted"
    params = "source=补充说明"
else:
    # 短文本但无法归类 → 保留原文
    desc = "（自动保留原文）"
    strategy = "template_direct"
return {"description": desc, "strategy": strategy, "params": params}
```

这样全自动模式下不会有 `"（必填）"` 的空白盲区。

### 验证

```bash
python -m py_compile pipeline/01b_auto_annotate.py
# 运行 01b，确认不再输出 "（必填）" 描述（除非模板确实有无法归类的新 shape）
```

---

## 改动 4: 02 — gpt_prompted 默认 contract

### 文件: `pipeline/02_shape_analysis.py`

### 问题

`_parse_output_contract()` 从自由文本解析 contract。如果内容描述没有写 `'样本'`、`'建议'` 等引号关键词，contract 的 `required_keywords` 为空。这导致 04 语义检查对这些 shape 不生效。

### 改动

#### 4A. 为 gpt_prompted shape 注入默认 contract

在 mapping 循环中（~L225），当 `strategy_exact == "gpt_prompted"` 且 contract 为空或缺少关键字段时，注入默认值：

```python
oc = _parse_output_contract(desc)

# 对 gpt_prompted shape 注入默认合约（确保语义检查覆盖）
if strategy_exact == "gpt_prompted" or "gpt_prompted" in strategy_hint_l:
    if not oc.get("required_keywords"):
        oc["required_keywords"] = ["样本", "建议", "反馈"]
    if "bracket_highlight" not in oc:
        oc["bracket_highlight"] = True
    if "ratio_required" not in oc:
        oc["ratio_required"] = True
```

在 prompt_specs 循环中（~L251）执行同样的注入。

### 验证

```bash
python -m py_compile pipeline/02_shape_analysis.py
# 运行 02，检查 02-shape_analysis_map.json 中 gpt_prompted shape 的 output_contract 是否都有默认值
```

---

## 改动 5: orchestrator — 选项4 全自动模式

### 文件: `orchestrator.py`

### 改动

#### 5A. 启动菜单新增选项4

当前代码（~L1214-1226）：

```python
print("  1 — 单轮（仅生成，不修正）")
print("  2 — 两轮（生成 + 1轮修正）")
print("  3 — 三轮（生成 + 2轮修正）")
```

新增：

```python
print("  4 — 全自动三轮（跳过所有暂停）")
```

选项解析：

```python
if choice in ('1', '2', '3', '4'):
    if choice == '4':
        max_rounds = 3
        auto_mode = True
    else:
        max_rounds = int(choice)
        auto_mode = False
    break
```

#### 5B. 新增 `auto_mode` 属性

在 `PPTOrchestrator.__init__()` 中新增 `self.auto_mode = auto_mode` 参数。

在实例化时传入：

```python
orch = PPTOrchestrator(
    project_root=project_root,
    max_rounds=max_rounds,
    max_budget=args.max_budget,
    skip_analyst_first_round=skip_analyst,
    auto_mode=auto_mode,  # 新增
    ...
)
```

#### 5C. auto_mode 跳过所有暂停

需要跳过的 3 个暂停点：

**P1 — Analyst 后批注审核（~L890）：**

```python
if self.skip_analyst_first_round:
    print(f"\n  [跳过 PAUSE] ...")
elif self.auto_mode:
    print(f"\n  [全自动] 跳过批注校准暂停")
else:
    # 原有的 os.startfile + input() 暂停逻辑
```

**Prompt Review — 03a Phase 1 后（在 `_run_03a_with_prompt_review()` 中）：**

```python
if has_pending:
    if self.auto_mode:
        print(f"  [全自动] 跳过 prompt 审核，直接执行 GPT")
    else:
        # 原有的 os.startfile + input() 暂停逻辑
```

**P2 — PPT 生成后验收确认（~L1050）：**

```python
if self.auto_mode:
    print(f"  [全自动] 跳过 PPT 审核，直接进入验收")
else:
    # 原有的 "是否进入验收？[Y/n]" 暂停逻辑
```

#### 5D. auto_mode 的 Analyst LLM 策略

选项4 = max_rounds=3，所以 `skip_analyst = (max_rounds == 1)` → `skip_analyst = False`。
Analyst LLM 正常运行，不跳过。这是正确的：全自动模式更需要 LLM 增强批注。

#### 5E. auto_mode 结束信息

所有轮次完成后，打印全自动模式专用总结：

```python
if self.auto_mode:
    print(f"\n{'=' * 60}")
    print(f"🤖 全自动模式完成 — {self.max_rounds} 轮迭代")
    print(f"   最终产物: claude-ppt {final_version}.pptx")
    print(f"{'=' * 60}")
```

### 验证

```bash
python -m py_compile orchestrator.py
# 运行 orchestrator → 选 4 → 确认无任何暂停，全程自动跑完
```

---

## 最终流程图

```
[S1] 选账户 → [S2] 选模式

━━━ 选项1: max=1 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Pipeline(01+01b)，跳过 LLM → 02 → 03a(prompt审核暂停) → 03b → 完成

━━━ 选项2/3: max=2/3 ━━━━━━━━━━━━━━━━━━━━━━━━
Pipeline(01+01b) + Analyst LLM → P1暂停 → 02 → 03a(prompt审核暂停) → 03b
→ P2暂停 → 04验收 → [修正循环] → 完成

━━━ 选项4: max=3 全自动 ━━━━━━━━━━━━━━━━━━━━━
Pipeline(01+01b) + Analyst LLM → 02 → 03a(直接调GPT) → 03b
→ 04验收 → [自动修正 × 2轮] → 完成
（全程无暂停）
```

## 文件改动汇总

| 文件 | 改动 | 复杂度 |
|------|------|--------|
| `pipeline/03a_build_shape.py` | output_contract 注入 prompt | 中 |
| `pipeline/02b_iteration_setup.py` | 量化修正 + 约束替换 | 中 |
| `pipeline/01b_auto_annotate.py` | fallback 默认描述 | 低 |
| `pipeline/02_shape_analysis.py` | 默认 contract 注入 | 低 |
| `orchestrator.py` | 选项4 + auto_mode | 中 |

## 执行检查清单

1. `python -m py_compile` 全部 5 个文件
2. `python pipeline/01b_auto_annotate.py` → 确认无 "（必填）" 输出
3. `python pipeline/02_shape_analysis.py` → 确认 gpt_prompted shape 有默认 contract
4. `python pipeline/03a_build_shape.py --assemble-only` → 确认 prompt 包含【输出合约】段
5. `python orchestrator.py` → 选 1 → 确认单轮 + 1次暂停（prompt审核）
6. `python orchestrator.py` → 选 4 → 确认全自动 3 轮无暂停
