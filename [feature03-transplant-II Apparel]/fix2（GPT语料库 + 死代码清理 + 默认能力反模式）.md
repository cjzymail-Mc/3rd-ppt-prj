# fix2 — GPT 语料库注入 + 死代码清理 + 默认能力假设反模式

**日期**：2026-04-29
**范围**：`apparel_ppt.py` / `Function_030.py` / 对话流程教训
**关联文档**：
- `.claude/memory/feedback_check_skills_first.md`（用户 memory）
- `.claude/memory/feedback_debug_protocol.md` "默认能力假设反模式" 节
- `.claude/memory/feedback_chart_write.md` "对象引用 vs Selection 路径" 节
- `.claude/CLAUDE.md` §0 第 4 条反射 + §3 chart 路径硬规则 + §3 GPT 风格锚硬规则

---

## 本轮 4 件事

### 1. apparel chart 数值轴答疑（已记）

用户问"chart value bar 最大值是自适应还是固定？"。答：**固定 0~6**，遵循 CLAUDE.md "bar chart max = 量表 max + 1" 硬规则——5 分制 → 6，避免 score=5 的 bar 末端数据标签被压。已落地 `apparel_ppt.py::make_chart_for_apparel:706` `_val_axis.MaximumScale = _SCALE_MAX + 1`。

### 2. Excel zoom 技术答疑（无代码改动，但产生固化经验）

用户问"为什么 yzr/zxh/apparel 的 make_chart 不需要 Excel_zoom，但 Function_030.make_chart 必须？"。

**根因**：Excel COM chart 操作有两套路径——
- 老 `make_chart` 走 **UI Selection** 路径（`Range.Select() → Selection.Copy`），强依赖视口可见
- 新 `make_chart_for_*` 走 **对象引用** 路径（`mc_sht.charts.add() → api[0].Copy()`），免疫可见性

`ChartObject.Copy()` 本身不读 ActiveSelection；老代码报错的真正原因是它前面那串 `selection.end('down')` 导航逻辑需要视口可见——Copy 是被 Select 拖累的。

**固化位置**：`feedback_chart_write.md` 新增 "对象引用 vs Selection 路径" 节 + CLAUDE.md §3 加硬规则。

### 3. 死代码清理（已落地）

`Function_030.py:241-256` 删除 3 个函数：
- `run_com_template_analysis`
- `run_com_build_final_ppt`
- `run_com_verify_fidelity`

**判定无用依据**：
1. 全仓 `Grep` 无任何调用方
2. 三个函数 import 的模块（`analyze_templates_com.py` / `build_codex_ppt_com.py` / `verify_ppt_fidelity_com.py`）全部位于 `【trash-bin】/codex work/`，本来就是 ImportError 状态
3. 是 Pipeline 早期 codex 实验路线的入口包装，核心实现已搬到 trash-bin，剩下的 wrapper 是僵尸代码

### 4. apparel GPT 风格语料库注入（已落地）

**问题**：`_build_rich_prompt` 的 `style_anchor` 槽位之前用 `fallback_map[focus]`（短文本）当风格示例——信息密度不够，且和 fallback 兜底语义混用。

**修复**：
- `apparel_ppt.py` 顶部加 `_STYLE_REFERENCE_CORPUS`（13 条"问题 // 优势"对照式专业试穿评价语料）
- line 613 `style_anchor=fallback` → `style_anchor=_STYLE_REFERENCE_CORPUS`
- `fallback_map` 完整保留作为 GPT 失败兜底（职责分离）

**为什么不外置 txt**：13 行常量足够清晰；语料和 prompt 模板高度耦合（focus 模式/content_source 模式的措辞依赖它），放一起方便联动调整；与项目 `fallback_map` / `_HEIGHT_CM_MAX` / `_BMI_OK` 模块级常量惯例一致。

---

## 默认能力假设反模式（本轮最重要的教训）

### 现场

用户：「服装试穿测试报告 PPT 我选中的这个shape 是结论，由 apparel_ppt.py 中的总结函数生成，为什么没有修改建议？少了一部分」

我（错）：「我没有访问你 PPT 当前选中状态的能力——既没有屏幕截图，也没有 COM 接口连到你打开的 PowerPoint。麻烦你告诉我 shape 名 / 贴文字 / 截图...」

用户：「奇怪，你看不到我选中的shape吗？我当前选中的shape」

我（错）：试图直接写 win32com 脚本读 ActiveWindow.Selection（被用户中断）

用户：「你什么情况，不是skills里有 read_selected_shape.py吗」

我（最终）：跑 `python skills/read_selected_shape.py` 拿到完整信息（TextBox 4, Slide 11, 230 字, 缺【修改建议】段）

### 真相

本项目通过 `win32com.client.GetActiveObject("PowerPoint.Application")` 桥接到运行中的 Office，**有完整能力**读取实时选中、当前 slide、所有 shape 属性。`skills/read_selected_shape.py` 早就存在，注释"读取当前鼠标选中的 PPT shape 的完整信息"白纸黑字。

我没去查 `skills/`，凭"通用 Claude 默认能力边界"判断当下能不能做——这是错误的。

### 反模式特征

- 用户消息出现"我选中的"、"我当前打开的"、"屏幕上的 X"等指代实时状态的表述
- 我没跑 `Glob skills/* debug/*` 就先回"我看不到 / 我做不到 / 我没有 X 能力"
- 用户反复纠正后才去查工具

### 正确反射（已固化到 3 处）

1. **`.claude/CLAUDE.md` §0 第 4 条反射**：触发 → `Glob skills/read_* debug/read_*`，禁止默认能力否认
2. **`.claude/memory/feedback_debug_protocol.md`** 反射动作表加一行 + "默认能力假设反模式" 详细复盘
3. **`C:\Users\xy24\.claude\projects\D--Technique-Support-Claude-Code-Learning-3rd-ppt-prj\memory\feedback_check_skills_first.md`**：用户 memory，跨会话生效

### 额外发现

- 跑 Python 脚本输出乱码（`���ŵ㡿`）→ 加 `PYTHONIOENCODING=utf-8` 重跑可解
- `skills/read_selected_shape.py` 输出全面：shape name / type / position / text full + paragraphs / fill / line + chart series / picture crop。是项目里**最重要的 PPT 调试工具之一**，应该在所有 PPT 调试任务中默认先跑

---

## 6.3 结论页缺【修改建议】根因（用户表示先不深究）

**现象**：Slide 11 / TextBox 4 / 230 字 / 8 行，只有【优点】3 条 + 【缺点】3 条，**没有【修改建议】段头**。

**生成路径**：不是 `apparel_ppt.py`，是 `Main.py:894-941` 的 6.3 结论页代码——`gen_result_prompt` → `GPT_5` → `clamp_text(280, 13)` → `Result_Bullet` → `_apply_conclusion_color`。

**排查表**：

| 候选 | 评估 |
|--|--|
| ❌ prompt 没要求第三段 | 排除（`Function_030.py:567-569` 写死了三段式） |
| ❌ `clamp_text` 截掉了 | 排除（budget 280 字/13 行 > 当前 230 字/8 行，远没触限） |
| ❌ `_apply_conclusion_color` 删了 | 排除（只染色，不删段） |
| ❌ `_strip_bullet_on_section_headers` 删了 | 排除（只去 ■ bullet） |
| ✅ **GPT 模型自己偷懒，没生成第三段** | **唯一剩下的可能** |

**最硬证据**：【修改建议】段头本身都没出现。如果是被截断，至少能看到段头 + 残段；现在直接没段头 = GPT 根本没生成。

**未来修复方向**（用户暂不动手，记此为 backlog）：
1. **强约束 + 重排限制**：把 prompt 字数限制改 "总字数控制在 250-280 字之间"（区间），把"硬性要求"改 "缺一段视为输出失败"
2. **后置校验 + 自动补段**：completion 拿回后检查 "【修改建议】" 是否存在，若缺则追加 "暂无显著改进项" 兜底段头
3. **拆成两次 GPT 调用**：先优缺点，再单独建议（重，不推荐）

推荐路线：方案 1 + 方案 2 组合（prompt 层减少漏段概率 + 代码层硬兜底）。

---

## 修改文件清单

| 文件 | 改动 |
|--|--|
| `src/apparel_ppt.py` | 顶部加 `_STYLE_REFERENCE_CORPUS` 常量（13 行）；line 613 `style_anchor=_STYLE_REFERENCE_CORPUS` |
| `src/Function_030.py` | 删除 line 241-256 的 3 个死函数 |
| `.claude/CLAUDE.md` | §0 反射动作表加第 4 条；§3 加 "chart Copy/Delete 走对象引用"、"GPT 风格锚 ≠ fallback" 两条硬规则 |
| `.claude/memory/feedback_debug_protocol.md` | 反射动作表加第 5 条；新增 "默认能力假设反模式" 节 |
| `.claude/memory/feedback_chart_write.md` | 新增 "对象引用 vs Selection 路径" 节 |
| `C:\Users\xy24\.claude\projects\...\memory\feedback_check_skills_first.md` | 用户 memory 新建（跨会话生效） |
| `[feature03-transplant-II Apparel]/fix2（...）.md` | 本文件（固化报告） |
