# fix2: 三重混合制架构修复计划

## Context

三重混合制（Pipeline 50% + Agents 40% + Developer 10%）概念方向正确，但实现层面存在严重问题：

- **代码 3-way 重复**：`_write_chart`、`_write_text`、`_apply_keyword_color` 各 3 份拷贝
- **yzr/zxh 95% 重叠**：20+ 函数在两个文件间复制粘贴
- **Prompt 4-5 套独立系统**：Pipeline 模板、Pipeline 硬编码回退、yzr 硬编码、zxh 硬编码、Function_030 遗留
- **Pipeline↔src/ 零连接**：完整的架构隔离墙，无自动化移植路径
- **src/ 无测试**：Pipeline 有 self_check.py，src/ 无任何验证
- **陈旧引用**：developer.md 多处指向已重命名的 codex_ppt.py

---

## 现状事实

### 图表方案差异（并非"重复"，而是两套不同机制）

| | Pipeline `_write_chart` | Function_030 `make_chart*` |
|--|--|--|
| 机制 | 往模板已有 chart shape 原位注入数据 | 在 Excel 中新建 chart → OLE 粘贴到 PPT |
| 库 | win32com 直接 COM | xlwings (.api[] 访问 COM) |
| 格式化 | 零格式化（继承模板） | 50+ 行手动格式化 |
| 适用场景 | 模板已含 chart 占位符 | 模板无 chart，需从 Excel 数据动态生成 |

结论：两种图表方案**解决不同问题，不应合并**，但 Pipeline 的 `_write_chart` 与 yzr/zxh 的 `_write_chart` 是同一个函数的 3 份拷贝，应去重。

### 代码重复清单（确认级别）

| 函数 | pipeline/03b | src/yzr_ppt.py | src/zxh_ppt.py | 差异 |
|------|:-:|:-:|:-:|------|
| `_write_chart` | L123 | L580 | L606 | pipeline 返回 dict，yzr/zxh 返回 bool |
| `_write_text` | L90 | L559 | L585 | pipeline 有 readback 验证，yzr/zxh 无 |
| `_apply_keyword_color` | L193 | L616 | L642 | pipeline 用 `com_get`，yzr/zxh 用本地 `_com_get` |
| `_build_respondent_block` | L357 | L336 | L338 | 近乎完全相同 |
| `_classify_columns` | L315 | L183 | L191 | 完全相同 |
| `_find_col` | L305 | L173 | L181 | 完全相同 |
| `clamp_text` | common | L310 | L312 | 完全相同 |
| `_col_values` | L168 | L229 | L237 | 完全相同 |
| `_score_10pt/_score_to_grade` | L186 | L247 | L255 | 完全相同 |
| `_sample_stat_text` | L213 | L273 | L278 | 完全相同 |
| `_extract_score_means` | — | L125 | L133 | yzr/zxh 之间完全相同 |
| 颜色常量 `_RED/_BLUE/_BLACK` | ✓ | ✓ | ✓ | 值相同，变量名不同 |
| `_ADVANTAGE/_DISADVANTAGE_MARKERS` | ✓ | ✓ | ✓ | 完全相同 |

### Prompt 系统对比

| 系统 | 存储方式 | 模板加载 | 版本管理 | 独有能力 |
|------|---------|---------|---------|---------|
| Pipeline `gpt_summary.md` | MD 文件 + `.format()` 占位 | ✓ 运行时加载 | 无(每次覆盖) | output_contract, user_instruction |
| Pipeline 硬编码回退 | Python 字符串 | — | — | 同上(冗余) |
| yzr_ppt `_build_rich_prompt` | Python 字符串 | — | — | 无 style_anchor（始终传空） |
| zxh_ppt `_build_rich_prompt` | Python 字符串 | — | — | `fmt='p1p2'` 独有模式 |
| Function_030 `gen_*_prompt` | Python f-string | — | — | 完全不同的 prompt 设计 |

---

## 修复计划（5 项，按优先级排序）

> **优先级说明（2026-04-16 修订）**
>
> 原始 fix2 的顺序是 Fix1→Fix2→Fix3→Fix4→Fix5，偏向"架构债一次性清理"。
> 经过讨论后，改为**按生产风险优先、架构债靠后**的顺序：
>
> | 新优先级 | 任务 | 重新定位的理由 |
> |--|--|--|
> | ★★★ | **Fix 1** 陈旧引用 | 零成本，先做 |
> | ★★★ | **Fix 5** src/ 冒烟测试 | 保护当前生产（yzr/zxh 零测试） |
> | ★★ | **Fix 3** Developer Playbook | 架构债，但新模板移植是低频操作，不紧急 |
> | ★★ | **Fix 2 (partial)** 仅提取纯数据工具 | 降低跨文件 bug 风险，但不动写入/prompt 逻辑 |
> | ★ | **Fix 4** prompt 版本注释 | 保留独立副本 + 版本追溯，不做文件级共享 |
>
> **关键决策：**
> - yzr/zxh **独立性优先**：`_write_text/_write_chart/_build_rich_prompt` 等影响视觉输出的函数**保留各自独立**
> - 纯数据工具（`_find_col/_score_10pt/clamp_text` 等）可以提取到 `_ppt_shared.py`，不影响独立微调
> - Prompt 不做文件级共享（选项 A 否决），采用版本注释 + 独立 copy（选项 B 改良版）

---

### Fix 1: 陈旧引用清理（5 分钟，零风险）

`codex_ppt.py` 已改名为 `yzr_ppt.py`，但多处文档未同步。

**修改文件：**
- `.claude/agents/developer.md`：L49, L92, L102, L105 → 替换 codex_ppt.py 为 yzr_ppt.py
- `.claude/commands/developer.md`：L10 → 同上

---

### Fix 2 (partial): 共享工具模块 `src/_ppt_shared.py`（仅纯数据工具）

> **范围收窄**：原计划提取 20+ 函数，现在只提取**纯数据工具**（不涉及 PPT 写入、不影响视觉输出）。
> 理由：用户要求 yzr/zxh 保持独立微调能力，视觉输出相关函数必须各自独立。

**目标**：降低跨文件 bug 风险（例如 `_find_col` 在一侧修 bug，另一侧忘记同步），但**不牺牲**两套模板的独立可微调性。

**新建文件 `src/_ppt_shared.py`，只提取以下函数：**

```
# —— Excel 数据提取（纯读，不涉及 PPT） ——
_find_col, _classify_columns, _col_values
_extract_score_means, _xlwings_to_rows

# —— 评分/统计（纯计算） ——
_score_10pt, _score_to_grade, _sample_stat_text

# —— 文本处理（纯字符串） ——
clamp_text

# —— 常量（值完全相同） ——
_RED, _BLUE, _BLACK, _ADVANTAGE_MARKERS, _DISADVANTAGE_MARKERS
```

**修改 yzr_ppt.py / zxh_ppt.py**：删除以上函数，改为 `from src._ppt_shared import ...`（显式 import，非 `*`）。

**明确保留在各模板文件中（不提取）：**

| 函数 | 保留原因 |
|--|--|
| `_write_text` | 影响视觉输出；pipeline/yzr/zxh 已有分歧（readback 验证差异） |
| `_write_chart` | 影响视觉输出；返回值格式差异（dict vs bool） |
| `_apply_keyword_color` | 影响视觉输出；内部 `_com_get` 引用差异 |
| `_build_rich_prompt` | prompt 是 per-template 差异点（zxh 有 p1p2 模式） |
| `_build_content` | 策略路由因模板不同 |
| `_build_respondent_block` | 虽相似但涉及模板特定字段 |
| `_replace_image / _extract_shoe_image` | 图片处理可能因模板差异化 |
| `_call_gpt` | 可能需要 per-template 调参 |
| `_safe_text / _numeric / _com_get / _to_rows` | 包装函数，保留在各文件便于调试 |
| `_shoe_name` | 业务特定字段处理 |
| Shape 定义：`CODEX_SHAPES` / `ZXH_SHAPES` | 模板结构 |
| `make_xxx_slide()` | 公共 API |
| 模板特有函数（如 zxh 的 `_color_section_headers`） | 定制化 |

**Pipeline 03b 不改**：保留独立，避免动已稳定的 pipeline。3-way 重复接受现状。

---

### Fix 3: Developer 移植 Playbook 更新

**修改 `.claude/agents/developer.md` 场景2**：明确"移植交付物"和步骤。

```
新模板移植 Checklist:
━━━━━━━━━━━━━━━━━━━━━
输入:
  □ 模板 .pptx（template/ 目录）
  □ 配套 .xlsx 数据文件
  □ Pipeline 达到 ~80% 视觉满意度（如有的话）

Developer 工作:
  □ 新建 src/{template}_ppt.py（复制 yzr_ppt.py 骨架）
  □ 替换 shape 定义（SHAPES 列表）
  □ 修改 slide 编号（clone 哪页）
  □ 从 Pipeline 提取最终 prompt（02-prompt_specs.json / gpt_summary.md）
     → 写入 _build_rich_prompt()（或未来从共享模板文件加载）
  □ 处理图表 shape（决策树）：
     ┌─ 系列数固定 + 模板已含该图表 shape
     │   → 用 _write_chart()（保留在模板文件，从 yzr 复制）
     │   → 适用于复杂图表（雷达、散点、气泡等），只要模板预置了 shape
     │
     └─ 系列数动态 或 模板无图表 shape
         → 用 Function_030.make_chart*()（Excel OLE 粘贴）
         → 目前仅支持简单柱状/折线，复杂图表需扩展 Function_030
         → 警告：雷达图数据范围形状严格（N 行维度 × M 列系列），注入前校验
  □ 图表方案选择理由（写入代码注释，便于未来维护者判断）
  □ 接入 Main.py：
     - ask_template_choice() 增加选项
     - import + 调用 make_{template}_slide()
  □ 语法检查 + 至少 1 次端到端运行验证

不需要 Developer 做:
  × 重写 prompt（从 Pipeline 产物提取）
  × 重建 shape 格式/字体（Clone 继承）
  × 复制通用工具函数（import _ppt_shared）
```

---

### Fix 4: Prompt 版本追溯（独立副本 + 版本注释）

> **决策**：否决选项 A（文件级共享）和选项 C（覆盖机制），采用选项 B 的改良版。
> 理由：用户明确要求两套模板最大独立性，prompt 改一个影响两套的风险不可接受。

**做法：**

1. `_build_rich_prompt()` **保留**在各模板文件中（yzr_ppt.py / zxh_ppt.py 各自独立）
2. 在每个 `_build_rich_prompt()` 上方添加**版本注释**：

   ```python
   # prompt_src: pipeline/prompt_templates/gpt_summary.md
   # synced_at: 2026-04-16
   # synced_by: Developer（移植时从 pipeline 拷贝）
   def _build_rich_prompt(...):
       ...
   ```

3. 在 Developer Playbook（Fix 3）中增加移植 checklist 项：
   - ☐ 从 `pipeline/prompt_templates/gpt_summary.md` 拷贝最新 prompt
   - ☐ 更新 `synced_at` 日期
   - ☐ 如做 per-template 定制（如 zxh 的 p1p2），在注释中标明

**放弃的能力**：
- 不做自动 diff / 脚本化同步工具（留给未来，现在不紧急）
- 不做文件级共享（稳定性优先于 DRY）

**接受的风险**：
- Pipeline prompt 改进不会自动流到 src/
- 依赖 Developer 在移植 / 大版本升级时主动 check `synced_at`
- 版本注释提供审计基线，发现漂移时能定位

---

### Fix 5: src/ 冒烟测试（保护当前生产）

> **优先级上调**：原计划 Fix 5 是"保底"，现在是 ★★★ 与 Fix 1 并列。
> 理由：yzr/zxh 零测试覆盖，是两套生产代码的**当前**风险，不是未来风险。

**新建 `debug/test_src_smoke.py`**

针对每个已移植模板（yzr / zxh）的冒烟测试：

```
for template in [yzr, zxh]:
    1. 打开模板 PPT（template/*.pptx）
    2. Clone 目标 slide
    3. 遍历 Shapes，确认数量与 SHAPES 定义一致
    4. 验证每个 shape 可被 _write_text 读取（不实际改写）
    5. 调用 _build_content(mock_data)，验证 prompt 构建不抛异常（可 mock GPT）
    6. 打印 shape name/type/text 摘要 → 存为 baseline
    7. 关闭 PPT（不保存）
```

**触发时机：**
- 每次修改 yzr_ppt.py / zxh_ppt.py 后手动跑一次
- 每次 Fix 2 共享模块变更后**必须**跑
- 可选：接入 pre-commit hook（但不强制，避免开发摩擦）

**不追求：**
- 视觉 diff（太贵，留给 reviewer agent）
- 完整 GPT 调用（慢 + 花钱）
- Pipeline self_check.py 级别的完整度

**只要捕获：**
- Python 语法错误 / import 失败
- Shape 名称在模板中不存在（重命名回归）
- COM 接口调用异常
- `_build_content` / `_build_rich_prompt` 抛异常

---

## 实施顺序（2026-04-16 修订）

```
第一波（当日完成，保护生产）:
  Fix 1（陈旧引用）    → 5min   ★★★
  Fix 5（冒烟测试）    → 30min  ★★★

第二波（下一轮，降低架构债）:
  Fix 3（Playbook）    → 30min  ★★
  Fix 2 partial（纯数据工具共享） → 40min  ★★

第三波（补漏，不紧急）:
  Fix 4（prompt 版本注释）→ 15min  ★
```

**第一波完成后验证：**
- yzr / zxh 两套模板能正常生成 PPT（冒烟测试 pass）
- developer.md / commands/developer.md 无 codex_ppt.py 遗留

**第二波完成后验证：**
1. `python -m py_compile src/_ppt_shared.py` — 语法检查
2. `python -m py_compile src/yzr_ppt.py` — import 正确性
3. `python -m py_compile src/zxh_ppt.py` — import 正确性
4. `python debug/test_src_smoke.py` — 冒烟测试通过（两套模板）
5. `python Main.py` 端到端运行 — 选 yzr/zxh 模板，确认 PPT 正常生成
6. `python orchestrator.py` 运行 — 确认 pipeline 不受影响

**第三波完成后验证：**
- yzr_ppt.py / zxh_ppt.py 的 `_build_rich_prompt` 上方均有 `prompt_src` / `synced_at` 注释
- developer.md Playbook 中 prompt 同步 checklist 可用
