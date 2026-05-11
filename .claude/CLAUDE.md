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

格式：`(YYYY-MM 触发场景) 结论 → 详情位置`

**短规则（独立成立，无外链）**：
- **OLE 图表粘贴**：`Shapes.Paste()` 后必须 `CutCopyMode = False` 断热链接，否则删行后 PPT 图表失数据
- **CopyPicture 常量**：xlPicture = **-4147**（矢量 EMF），`4` 是无效值会退化为位图
- **删行前先 delete chart**：否则 chart 公式引用失效时 Excel 弹"错误公式引用"弹窗
- **图表两套机制勿混淆**：Pipeline `_write_chart` = 原位注入模板 chart 数据；`make_chart_for_*` = Excel 新建 chart → OLE 粘贴；两者解决不同问题

**链接到 memory 的详情规则**：
- `(2026-04 fix4 分发场景)` chart 必须走 `make_chart_for_*`（xlwings + OLE 粘贴），禁 `_write_chart` 原位改 → `[feature03-transplant]/fix4（图表路线切换）.md`
- `(2026-04 fix4 3D chart)` xlwings 建的 3D chart 默认 Elevation/Rotation 错位，必须显式设 7 个 3D 参数 → `.claude/memory/feedback_chart_write.md`
- `(2026-04 fix5)` `Shapes.Paste()` 返回 ShapeRange，访问 `.Chart` 必须先 `.Item(1)` 拿真 Shape；隐藏 chart 标题用 `HasTitle=False + SetElement(0)` → `[feature03-transplant]/fix5（chart-title-hide）.md`
- `(2026-04 tk popup)` HWND 必须用 `wm_frame()`，`winfo_id()` 拿子控件 HWND 让 `SetWindowPos`/`FlashWindowEx` 静默失败；统一 `_get_toplevel_hwnd` → `Function_030.py`
- `(2026-04 GPT 输出)` 必经 `clamp_text` 剔空行+strip，否则 splitlines+join 让 PPT TextFrame 行数翻倍超 shape 高度 → `src/_ppt_shared.py::clamp_text`
- `(2026-04 结论页染色)` 用 `_apply_conclusion_color`（`<>`红 / `[]`蓝 / `()`粗），不要复用 `_apply_keyword_color`（后者按 section context per-shape）→ `.claude/memory/feedback_conclusion_coloring.md`
- `(2026-04 Result_Bullet)` `Class_030.Text_Box` 默认 `msoAutoSizeShapeToFitText`，`clamp_text` 可超模板 shape 几何（硬上限 = slide 高度）；`_write_text` 显式 `AutoSize=0` 锁定 → `src/_ppt_shared.py`
- `(2026-04 多阶段 GPT 累积)` 用 `summary_sink: list | None = None` 参数订阅内层每轮 completion，不破坏 return 签名 → `.claude/memory/feedback_summary_sink.md`
- `(2026-04 tk 弹窗样式)` iOS systemGroupedBackground + 白卡片 + Indigo 描边；`highlightthickness` 不用 `relief`；尺寸用 `winfo_reqwidth/reqheight` 不要 hardcode → `.claude/memory/feedback_popup_ui.md`
- `(2026-04 apparel-fix1 skip)` `SHAPES strategy: "skip"` **不清**模板预置文字，新模板移植必查源 shape 是否预留空 → `.claude/memory/feedback_skip_vs_clear.md`
- `(2026-04 apparel-fix1 BMI)` 100KG/1cm 等单位混淆，粗修 m→cm + 斤→kg 后用 `BMI∈[16,32]` 交叉验证识别误填 → `.claude/memory/feedback_unit_normalize_bmi.md`
- `(2026-04 fix4 chart 引用)` Copy/Delete 走 `worksheet.charts.add()` 对象引用，不走 `Range.Select+Selection`；后者强依赖 ActiveWindow/视口/选中态 → `.claude/memory/feedback_chart_write.md`
- `(2026-04 apparel GPT 槽)` `style_anchor` 用 `_STYLE_REFERENCE_CORPUS`（专业语料）；`fallback_map` 只作 GPT 失败兜底；两槽位职责互不污染 → `src/apparel_ppt.py:_STYLE_REFERENCE_CORPUS`

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
| `src/_ppt_shared.py` | 共享工具模块（已建立，消除 yzr/zxh 重复） |
| `Main.py` | src/ 生产入口（1055行） |

---

## 6. 详情索引

| 主题 | 位置 |
|------|------|
| Step1/2/3 Agent 定义 | `.claude/agents/step1-analyzer.md` 等 |
| Developer 移植规范 + Checklist | `.claude/agents/developer.md` |
| COM 开发规范 | `.claude/memory/feedback_com_constraints.md` |
| 混合工作流 Pipeline→LLM | `.claude/memory/feedback_hybrid_workflow.md` |
| 手动 Pipeline 命令 + 批注字段 | `.claude/memory/reference_manual_pipeline.md` |
| 架构修复计划（fix2） | `[feature03-transplant]/fix2（三重混合架构整改）.md` |
| Shape 微调工作流 + 调试入口 | `skills/fine-tuned-shapes.md` |
| 3 账号 auto-memory junction 架构 | `.claude/memory/reference_3account_junction.md` |
| 3 账号 junction 移植方案（新项目/新机器复用） | `skills/memory-junction-3account.md` |
