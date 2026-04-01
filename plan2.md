# Plan2: 热迭代闭环增强（3 项改动）

## 改动总览

| # | 改动 | 目的 | 涉及文件 |
|---|------|------|---------|
| A | 版本号语义化 | 主版本=冷启动，副版本=热迭代 | orchestrator.py |
| B | 选0 末尾自动验收 | 冷启动产出 fix.md，衔接后续热迭代 | orchestrator.py |
| C | 热迭代 Round 1 继承 fix.md | Builder LLM 自动优化 prompt | orchestrator.py |

三项改动仅涉及 `orchestrator.py`，无需改 pipeline 脚本或 agent spec。

---

## A. 版本号语义化

### 当前问题

版本号只做递增防碰撞，没有语义。当前状态：
- pptx: 1.0, 1.2, 1.4, 1.5, 1.6, 1.9（有空洞）
- tracker: 1.1~1.9（9个版本）
- 下次无论选什么，都从 2.0 开始 — 选0"初始化"得到 2.0，语义混乱

### 设计

**主版本号 = 冷启动会话，副版本号 = 热迭代轮次**

| 操作 | 版本规则 | 示例 |
|------|---------|------|
| 选 0（冷启动） | 跳到下一个 X.0 | 1.x → 2.0，2.x → 3.0 |
| 选 1-4（热迭代） | 递增副版本 | 2.0 → 2.1 → 2.2 |
| 选 5（验收） | 不产生新版本 | — |

示例时间线：
```
选0 → 1.0          ← 第一次分析
选2 → 1.1, 1.2     ← 迭代
选0 → 2.0          ← 重新分析（新模板/新数据）
选1 → 2.1          ← 微调
选2 → 2.2, 2.3     ← 迭代
选0 → 3.0          ← 又一次全新分析
```

### 改动：`run()` 中 base_idx 计算（~L983）

```python
base_idx = self._detect_next_version_index()

# ▸ 新增：冷启动跳到下一个主版本 X.0
if self.init_mode:
    base_idx = (base_idx + 9) // 10 * 10   # 向上取整到下一个十的倍数
    # 10→10(1.0), 16→20(2.0), 20→20(2.0), 21→30(3.0)
```

验证算术：
- 首次运行：raw=10 → (10+9)//10*10 = 10 → 版本 1.0 ✓
- 上次到 1.5(idx15)：raw=16 → (16+9)//10*10 = 20 → 版本 2.0 ✓
- 上次到 1.9(idx19)：raw=20 → (20+9)//10*10 = 20 → 版本 2.0 ✓
- 上次到 2.0(idx20)：raw=21 → (21+9)//10*10 = 30 → 版本 3.0 ✓

### 改动：冷启动强制 `is_continuation = False`（~L985）

```python
# 现有
is_continuation = (base_idx > 10) and fix_report.exists()
# 改为
is_continuation = (base_idx > 10) and fix_report.exists() and not self.init_mode
```

原因：冷启动是全新分析，不应继承上一个会话的 fix.md 或走 02b 续跑路径。

---

## B. 选0 末尾自动验收

### 当前问题

选0 产出 PPT 后直接结束，用户不知道质量如何，也没有 fix.md。后续热迭代无法自动优化。

### 设计

选0 的 `max_rounds == 1` 结束点，增加自动验收（仅生成报告 + fix.md，不进修正轮）。

```
当前选0:  01+01b → Analyst → ⏸️批注 → 02→03a→03b → PPT → 结束
改后选0:  01+01b → Analyst → ⏸️批注 → 02→03a→03b → PPT → 04验收(仅报告) → 结束
```

### 改动：`max_rounds == 1` 分支（~L1253）

```python
if self.max_rounds == 1:
    if self.init_mode:
        # 冷启动：自动验收（仅报告，不进修正轮）
        print(f"\n  [验收] 自动检查 claude-ppt {version}.pptx 质量...")
        r = self._run_pipeline(
            "pipeline/04_shape_diff_test.py",
            ["--target", f"claude-ppt {version}.pptx"]
        )
        for line in r.stdout.splitlines():
            stripped = line.strip()
            if stripped:
                print(f"    {stripped}")
        passed, fix_type = self._check_review_passed()
        if passed:
            print(f"\n✅ claude-ppt {version}.pptx 初始化完成，验收通过！")
        else:
            print(f"\n⚠️  claude-ppt {version}.pptx 初始化完成，验收未通过 (fix_type={fix_type})")
            print(f"   已生成 fix.md — 下次选1或选2时 Builder LLM 会自动优化 prompt")
    else:
        print(f"\n✅ claude-ppt {version}.pptx 已生成，请人工审核。")
    self.monitor.display_summary(self.results, time.time() - start_time)
    self.state_manager.clear_state()
    return True
```

### 效果

选0结束后：
- 用户立即看到质量报告（三层分数）
- fix.md 自动生成（如果未通过）
- 下次选1/2时，改动C的 `fix_is_fresh` 逻辑自动继承

---

## C. 热迭代 Round 1 继承 fix.md

### 当前问题

选2 的 Round 1 不读 fix.md，直接走 PROMPT REVIEW → 出 PPT。验收结果被浪费。

### 设计

在热迭代 Round 1 的 PROMPT REVIEW **之前**，检查 fix.md 是否新鲜。若新鲜，先让 Builder LLM 自动优化 prompt。

```
当前 Round 1:  prerequisites → ⏸️ PROMPT REVIEW → 03a → 03b
改后 Round 1:  prerequisites → [if fix fresh] Builder LLM 改 prompt → ⏸️ PROMPT REVIEW → 03a → 03b
```

### 时间戳陷阱

02b --sheet-only 会修改 xlsx（新建 sheet），导致 xlsx.mtime > fix.mtime。
**必须在 02b 之前捕获时间戳比较结果。**

### 改动 1：在 02b 之前捕获 `fix_is_fresh`（~L986，紧接改动A之后）

```python
is_continuation = (base_idx > 10) and fix_report.exists() and not self.init_mode

# ▸ 新增：在 02b 修改 xlsx 之前，捕获 fix.md 新鲜度
fix_is_fresh = (
    is_continuation
    and fix_report.stat().st_mtime > xlsx_path.stat().st_mtime
)
if fix_is_fresh:
    print("  [INFO] fix.md 比 xlsx 更新，Round 1 将自动优化 prompt")

if is_continuation:
    # 02b --sheet-only ...（现有代码不变）
```

判定逻辑：
- fix.md 比 xlsx 更新 → 用户跑了验收但没手动改 Excel → **读取 fix.md**
- fix.md 比 xlsx 更旧 → 用户已手动编辑过 Excel → **忽略**
- init_mode → **永远忽略**（冷启动不继承旧会话的 fix）

### 改动 2：热迭代 Round 1 插入 Builder LLM（~L1155）

在 prerequisites 通过后、PROMPT REVIEW 之前：

```python
else:
    # ▸ 新增：如果 fix.md 新鲜，先让 Builder LLM 自动优化 prompt
    if fix_is_fresh:
        fix_data = []
        diff_path = self.project_root / "pipeline-progress" / "04-diff_result.json"
        if diff_path.exists():
            try:
                fix_data = json.loads(diff_path.read_text(encoding="utf-8")).get("fixes", [])
            except (json.JSONDecodeError, IOError):
                pass
        if fix_data:
            sheet_name = f"claude-ppt {version}"
            print(f"  [Agent] Builder LLM 根据 fix.md 优化 prompt ...")
            self.monitor.display_agent_start("builder")
            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent, AGENT_CONFIGS["builder"],
                self._builder_prompt_optimizer_prompt(sheet_name, fix_data)
            )
            self.monitor.display_agent_complete(result)
            if result.status == AgentStatus.FAILED:
                print(f"\n⚠️  Builder LLM prompt 优化失败，继续手动审核")

    # 现有 PROMPT REVIEW 逻辑不变
    if not self.auto_mode:
        ...
```

Builder LLM 失败时**不终止流程**，用户仍可在 PROMPT REVIEW 中手动修改。

---

## 完整用户工作流示例

### 典型流程：选0 → 选2

```
选0:
  01+01b → Analyst LLM → ⏸️批注审核 → 02→03a→03b → PPT 2.0
  → 04自动验收 → fix.md（⚠️ 未通过，keyword_missing）

选2 (max_rounds=2):
  Round 1:
    [Hot] 跳过01+01b，跳过Analyst LLM
    02b --sheet-only → sheet「claude-ppt 2.1」
    [fix_is_fresh] Builder LLM 读 fix.md → 自动改 prompt
    ⏸️ PROMPT REVIEW（用户可进一步调整）
    03a → 03b → PPT 2.1
    → 04 验收 → fix.md2
    → Reviewer LLM 补充诊断
  Round 2:
    02b --sheet-only → sheet「claude-ppt 2.2」
    Builder LLM 读 fix.md2 → 改 prompt
    ⏸️ PROMPT REVIEW
    03a → 03b → PPT 2.2
    → 04 验收 → ✅ PASS
```

### 断点续传：选0 → 选1 → 选5 → 选1

```
选0: → PPT 2.0 → fix.md
选1: [fix_is_fresh] Builder LLM → ⏸️ PROMPT REVIEW → PPT 2.1（不验收）
选5: → 04 验收 PPT 2.1 → fix.md2
选1: [fix_is_fresh] Builder LLM(fix.md2) → ⏸️ PROMPT REVIEW → PPT 2.2
```

---

## 改动量汇总

| 改动 | 位置 | 行数 | 说明 |
|------|------|------|------|
| A: 版本语义化 | ~L983 + L985 | ~5行 | base_idx 取整 + is_continuation 排除 init |
| B: 选0自动验收 | ~L1253 | ~15行 | max_rounds==1 分支内加 04 调用 |
| C: fix.md继承 | ~L986 + ~L1155 | ~20行 | fix_is_fresh 捕获 + Builder LLM 调用 |
| **合计** | | **~40行** | 仅改 orchestrator.py |

## 验证

1. `python -m py_compile orchestrator.py`
2. 选 0：确认版本号为下一个 X.0，末尾自动验收并生成 fix.md
3. 选 0 → 选 2：确认 Round 1 打印 "[Agent] Builder LLM 根据 fix.md 优化 prompt ..."
4. 选 0 → 手动改 xlsx → 选 2：确认 Round 1 跳过 Builder LLM（fix.md 不新鲜）
5. 选 0 → 选 0：确认主版本号递增（1.0 → 2.0 → 3.0）
6. 选 2 内部：确认 Round 1 → 2.1，Round 2 → 2.2（副版本递增）
