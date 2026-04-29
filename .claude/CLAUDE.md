# CLAUDE.md - PPT Pipeline 项目规范

## 0. 防卡顿规范

- 同一方案连续失败 2 次 → 停下来说明原因，提出替代方案
- 预计超过 2 分钟的操作 → 用 Agent(run_in_background) 分流
- 遇到不确定的技术选型 → 先问用户，不要默默试超过 3 分钟

### 4 条反射动作（fix3→fix4 + 2026-04-29 血的教训）

| 触发 | 反射 |
|--|--|
| 接到涉及 COM / OLE / 模板 / 分发的 bug | **第一步 grep 项目看有没有已解决同类问题的生产代码**，不是第一步改代码 |
| 用户用"我们之前约定"开头 | **立刻问"这个约定是在什么假设下达成的？当前场景假设还成立吗？"**——区分偏好 vs 硬需求 |
| 同一技术类别连续失败 2 次 | **停下来写 3 个候选路线**，不要再换变体继续第 3 次尝试 |
| 用户提"我选中的 / 我当前打开的 / 屏幕上的 X" | **先 `Glob skills/read_* debug/read_*`**，找现成桥接工具直接跑（如 `skills/read_selected_shape.py`），**禁止凭"默认 Claude 能力边界"先否认**。本项目通过 win32com `GetActiveObject` 桥接到正在运行的 Office，是有完整能力读取实时状态的 |

详见 `.claude/memory/feedback_debug_protocol.md`（7 步流程 + 4 条具体错误复盘）。

---

## 1. 双轨架构（三重混合制）

本项目存在**两套并行生产系统**，职责不同，不应混淆：

| | Pipeline / Orchestrator | src/ / Main |
|--|--|--|
| 入口 | `orchestrator.py` | `Main.py` |
| 机制 | Step1→2→3 + LLM Agents 自检 | 手工 Python + GPT 直调 |
| 适用场景 | 新模板分析、通用内容生成 | 已知模板的日常生产运行 |
| 核心文件 | `pipeline/*.py` | `src/Function_030.py` + `src/yzr_ppt.py` + `src/zxh_ppt.py` |

**新模板移植路径**：Pipeline 跑到 ~80% 视觉满意度 → Developer 写 `src/{name}_ppt.py`
（Clone 模板页继承格式，工具函数从 `src/_ppt_shared.py` import，prompt 从 Pipeline 产物提取）

### 5 阶段工作流（plan3 定稿）

```
新模板/数据源到手
       ↓
① Pipeline 首跑（必做）        orchestrator.py 全流程
       ↓
② 评估 PPT 效果（决策点）      看 04-fix_ppt.md 的 visual/readability/semantic
       ↓
   ┌───┴───┐
   ↓       ↓
③a 跳过   ③b 继续 Pipeline 迭代（修 .xlsx 标注 → 重跑 Step2/3）
   ↓       ↓
   └───┬───┘
       ↓
④ /developer 移植（默认路径）  Sonnet + 自动加载 developer.md Checklist
       ↓
⑤ 主 Claude 兜底（复杂问题）   路线决策、多轮 pivot、沉默 bug 这类才回主对话
```

### 决策点速查表（每次新任务先查）

| 任务类型 | 默认路径 |
|--|--|
| 完全新模板 | ① → ② → ③a/③b → ④ → 卡住时 ⑤ |
| 已知模板加新 shape | 直接 ④（/developer 改 SHAPES 列表） |
| 已知模板 bug 修复 | 直接 ⑤（路线判断类，主 Claude） |
| prompt 文案调优 | 直接 ④（/developer 改 _build_rich_prompt） |
| shape 微调 | 直接 ④（/developer + skills/fine-tuned-shapes.md） |
| chart 路线问题（fix4 类） | 直接 ⑤（主 Claude） |

完整流程图、各阶段动作清单、工具索引详见 `plan3（工作流5阶段定稿）.md` 与 `.claude/memory/feedback_workflow_routing.md`。

---

## 2. 核心代码规则

- **路径**: 始终用相对路径 + 正斜杠 `/`
- **最小改动**: 只改必要的部分，先说明再动手
- **Excel**: 统一 `win32com.client` COM（加密环境，禁 openpyxl/pandas）
- **PPT**: Clone 模板页，不新建 shape；禁 `python-pptx`
- **字体**: 统一微软雅黑（`_write_text` 自动设置）
- **换行**: PPT COM 用 `\r` 分段，`\n` 无效
- **染色**: GPT 用 `【】` 标注关键词 → `_apply_keyword_color` 按段落上下文红/蓝染色
- **截图**: 系统加密 PPT 导出图片，改用剪贴板→Pillow 方案绕过

---

## 3. 硬规则（反复踩过的坑）

- **OLE 图表粘贴**：`Shapes.Paste()` 后必须 `CutCopyMode = False` 断热链接，否则删行后 PPT 图表失数据
- **CopyPicture 常量**：xlPicture = **-4147**（矢量 EMF），`4` 是无效值会退化为位图
- **删行前先 delete chart**：否则 chart 公式引用失效时 Excel 弹"错误公式引用"弹窗
- **yzr_ppt / zxh_ppt 共享工具**：两文件 95% 重复，工具函数统一放 `src/_ppt_shared.py`，不要在各自文件中复制粘贴
- **图表两套机制勿混淆**：Pipeline `_write_chart` = 原位注入模板 chart 数据；`Function_030.make_chart*` = Excel 新建 chart → OLE 粘贴，两者解决不同问题
- **分发场景 chart 强制从零制表**（fix4）：模板 / 代码分发给他人、数据由他人填的场景，chart 必须走 `make_chart_for_{template}` 路线（xlwings 新建 + OLE 粘贴），**禁用** `_write_chart` 原位改。原因：chart 内部状态（IsLinked / embedded workbook / numCache）在分发链路里必然漂移；加密办公环境下 XML surgery 也不可用（CFB 非 zip）。详见 `[feature03-transplant]/fix4（图表路线切换）.md`
- **xlwings 3D chart 必须显式设置 3D 视图**（fix4）：xlwings 建立的 3D chart 默认 Elevation/Rotation 不等于 PPT 模板期望视角，OLE 粘贴后会视觉漂移。`make_chart_for_{name}` 必须显式设 `Elevation / Rotation / RightAngleAxes / AutoScaling / Perspective / DepthPercent / HeightPercent`。PPT "三维旋转" 面板 ↔ Excel chart API 映射表见 `.claude/memory/feedback_chart_write.md`
- **`Shapes.Paste()` 返回 ShapeRange，不是 Shape**（2026-04-27）：`mc_shape = mc_slide.Shapes.Paste()` 拿到的是 ShapeRange；`.Left/.Top/.Width/.Height` 会 fan-out 到内部 shape 所以能直接用，但 **`.Chart`/`.HasChart` 不在 fan-out 列表，会抛 `-2147352567 发生意外`**。访问 chart 必须先 `mc_shape.Item(1)` 取真正的 Shape。隐藏 chart 主标题双保险写法：`Item(1).Chart.HasTitle = False` + `Item(1).Chart.SetElement(0)`。详见 `[feature03-transplant]/fix5（chart-title-hide）.md` 假设记录或 `_ppt_shared.py::make_chart_for_yzr`
- **bar chart 数值轴 max = 量表 max + 1**（2026-04-27）：5 分制 → 6，10 分制 → 11。原因：`MaximumScale = scale_max` 时 score=max 的 bar 末端会被数据标签压住、看不清。已落地 `Function_030.py::make_chart_for_questionnaire`
- **tk popup HWND 必须用 `wm_frame()`，不是 `winfo_id()`**（2026-04-27）：`winfo_id()` 返回 Tk 子控件 HWND，`SetWindowPos` / `FlashWindowEx` 对它静默失败 —— 这是"任务栏不闪烁/弹窗不居中"的根因。统一用 `_get_toplevel_hwnd(win)`（在 `Function_030.py`）。多显示器居中也别用 `winfo_screenwidth()`，要按光标所在屏 `MonitorFromPoint + GetMonitorInfoW.rcWork`
- **GPT 输出文本必经 `clamp_text` 自动剔空行**（2026-04-27）：GPT 偶尔在段落间多吐空行，`splitlines` 直接 join 会让 PPT TextFrame 行数翻倍、超出 shape Height。`clamp_text`（`_ppt_shared.py`）入口已内置：剔纯空白行 + 每行 strip。新写 `gpt_prompted` 分支调用 GPT 后**必须**走 `clamp_text`
- **结论页用 bracket-typed 染色，不要复用 `_apply_keyword_color`**（2026-04-27 todays-task）：6.3 最终结论页有"优点 / 缺点 / 修改建议"三段；GPT 用半角 `<keyword>` 标优点（红+粗）、`[keyword]` 标缺点（蓝+粗）、`(keyword)` 标建议（仅粗），由 `_apply_conclusion_color`（`_ppt_shared.py`）统一处理 + 剥离 ASCII 标记。中文 **【】 保留为 section header 标记**（`_strip_bullet_on_section_headers` 用它识别段头去 ■）。两套染色函数适用场景不同：`_apply_keyword_color` 用 section context（per-shape，yzr/zxh 各 shape 单独标注）；`_apply_conclusion_color` 用 bracket type（单 shape 内多段、多色，结论页专用）。详见 `.claude/memory/feedback_conclusion_coloring.md`
- **`Result_Bullet` 自动 auto-grow 高度**（2026-04-27）：`Class_030.Text_Box` 不设 `tf.AutoSize=0`，PPT 默认 `msoAutoSizeShapeToFitText` 接管，shape 高度随文字自动撑高。意味着 `clamp_text(max_chars/max_lines)` 可以放心扩量到模板 shape 几何之上（例如 6.3 结论页从 plan4 的 200 字 / 10 行扩到 280 字 / 13 行）；硬上限是 slide 高度，不是 shape 默认高度。**只对 `Result_Bullet` / `Text_Box` 子类成立**；`_write_text` 显式置 `tf.AutoSize=0` 锁定模板几何，写到模板预置 shape 时不要混淆
- **多阶段 GPT 结论的累积用 `summary_sink: list | None = None` 参数**（2026-04-27 plan4）：当外层（`Main.py` 6.3）需要内层循环（`questionnaire_Excel` 多 runner 循环）每轮的 GPT completion 时，给内层函数加 `summary_sink=None` 可选参数 + 内部 `summary_sink.append(mc_completion)`。优点：不改 return 签名、不破坏既有调用、外部传 list 即可订阅。详见 `.claude/memory/feedback_summary_sink.md`
- **tk 弹窗样式约定**（2026-04-27）：`_ask_with_countdown` 用 iOS systemGroupedBackground (`#F2F2F7`) 窗口 + 纯白卡片按钮 + `#4A6CF7` Indigo 描边标记默认按钮；字体统一 `Microsoft YaHei UI`；按钮用 `highlightthickness=2` + `highlightbackground` 实现描边而不是 `relief`；自然尺寸用 `winfo_reqwidth/reqheight()` + `width` 入参作下限，不要 hardcode `height = 80 + 60 * len(options)`。**不要尝试 "蓝 header band + solid CTA" 路线**——tk 没圆角和阴影，做出来反而粗糙。详见 `.claude/memory/feedback_popup_ui.md`
- **`skip` 策略 ≠ 清旧文本**（2026-04-28 apparel-fix1）：`SHAPES` spec 里 `strategy: "skip"` 只表示"代码不写新值"，**不会清除模板里预置的文字**。clone 模板时旧文本会随 shape 一起带过来。所以装饰性 shape（如 apparel 4 个 Oval 虚线圈）想真空：要么改源模板把那 shape 文字清掉（推荐，一次到位），要么显式加 `clear` strategy 在 `_build_content` 返回 `""` 后强制 `tf.TextRange.Text = ""`。今天 Oval 3/13/16/19 误用 `score_category_mean` 写分数被发现，改 skip 后清空模板源 shape 即可永久修复。**新模板移植时检查清单**：所有 `skip` shape 在源模板里也要是空的，否则 clone 会残留
- **测试者身高/体重单位混淆用 BMI 交叉验证**（2026-04-28 apparel-fix1）：填问卷常见 100 KG 实际是 100 斤、1 CM 实际是 1 m。单一阈值（如 weight > 110 → 斤）会漏掉 100 这种边界值。正确套路：(1) 粗修 m→cm（< 3 视为 m）+ 斤→kg（> 110 视为斤），(2) 算 BMI，越界（不在 [16, 32]）则试 weight ÷ 2，新 BMI 落回区间才采纳。例：(160cm, 100kg) BMI=39 → 试 50kg → BMI=19.5 ✓ → 用 50kg；(180cm, 90kg) BMI=27.8 ✓ → 保留 90kg（不误伤真胖子）。已落地 `apparel_ppt.py::_normalize_person`，prompt 和 fallback 都先洗再喂下游 GPT
- **chart Copy/Delete 走对象引用，不走 UI Selection**（2026-04-29）：Excel COM 里 chart 有两套访问入口：(A) Selection 路径 `Range.Select() → Selection.Copy`，**强依赖 ActiveWindow + 视口 + 选中状态**；(B) 对象引用路径 `worksheet.charts.add()` 拿到的 xlwings chart 对象，`mc_chart1.api[0].Copy()` / `_tmp_chart.delete()`，**仅依赖 Worksheet 对象本身**，免疫可见性/缩放/滚动。`Function_030.py::make_chart` 走 (A) 所以必须 `Excel_zoom(mc_sht, 30)` 把 chart 缩进视口；`yzr/zxh/apparel_ppt.py::make_chart_for_*` 全部走 (B) 所以无需任何缩放——chart 在不在屏幕上、sheet 切没切到当前页都不影响 Copy/Delete。新写 chart 函数一律走 (B)；`Function_030.make_chart` 是技术债，重构方向是按 (B) 改写后丢弃 `Excel_zoom`。详见 `.claude/memory/feedback_chart_write.md` "对象引用 vs Selection 路径" 节
- **GPT 风格锚 ≠ fallback，要分两个槽位**（2026-04-29 apparel）：早期 `_build_rich_prompt` 把 `style_anchor` 用 fallback 文本充当——结果 1）GPT 拿到的"风格示例"信息密度不够，2）fallback 失败兜底语义被风格槽污染。正确做法：模块级常量 `_STYLE_REFERENCE_CORPUS`（专业语料）喂给 `style_anchor`；`fallback_map` 只保留作 GPT 失败时的硬编码兜底。两个槽位职责互不干扰。已落地 `apparel_ppt.py:_STYLE_REFERENCE_CORPUS`（13 条"问题 // 优势"对照式参考文本）+ line 613 `style_anchor=_STYLE_REFERENCE_CORPUS`

---

## 4. 入口命令

```bash
python orchestrator.py    # Pipeline 系统（菜单 0=全自动 / 1/2/3 分步）
python Main.py            # src/ 生产系统
python src/yzr_ppt.py     # yzr 单页调试（需先打开 Excel）
python src/zxh_ppt.py     # zxh 单页调试（需先打开 Excel）
```

---

## 5. 核心文件索引

| 文件 | 作用 |
|------|------|
| `orchestrator.py` | Pipeline 调度入口（1425行） |
| `pipeline/03a_build_shape.py` | GPT 内容生成 + prompt 管理 |
| `pipeline/03b_build_ppt_com.py` | COM 写入 PPT（_write_chart / _write_text） |
| `pipeline/prompt_templates/gpt_summary.md` | GPT prompt 模板（Pipeline 专用） |
| `src/Function_030.py` | 生产核心库（3504行）：GPT_5、问卷、图表、Excel COM |
| `src/yzr_ppt.py` | 杨祖锐模板：Clone Slide 15（含 `__main__` 单页调试） |
| `src/zxh_ppt.py` | 之行模板：Clone Slide 17（含 p1p2 模式 + `__main__` 单页调试） |
| `src/_ppt_shared.py` | 共享工具模块（fix2 计划新建，消除 yzr/zxh 重复） |
| `Main.py` | src/ 生产入口（1055行） |

---

## 6. 详情索引

| 主题 | 位置 |
|------|------|
| Step1/2/3 Agent 定义 | `.claude/agents/step1-analyzer.md` 等 |
| Developer 移植规范 + Checklist | `.claude/agents/developer.md` |
| 知识固化师（Curator） | `.claude/agents/curator.md` |
| COM 开发规范 | `.claude/memory/feedback_com_constraints.md` |
| 混合工作流 Pipeline→LLM | `.claude/memory/feedback_hybrid_workflow.md` |
| 手动 Pipeline 命令 + 批注字段 | `.claude/memory/reference_manual_pipeline.md` |
| 架构修复计划（fix2） | `[feature03-transplant]/fix2（三重混合架构整改）.md` |
| Shape 微调工作流 + 调试入口 | `skills/fine-tuned-shapes.md` |
