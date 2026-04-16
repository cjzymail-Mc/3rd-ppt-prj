# CLAUDE.md - PPT Pipeline 项目规范

## 0. 防卡顿规范

- 同一方案连续失败 2 次 → 停下来说明原因，提出替代方案
- 预计超过 2 分钟的操作 → 用 Agent(run_in_background) 分流
- 遇到不确定的技术选型 → 先问用户，不要默默试超过 3 分钟

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
| 架构修复计划（fix2） | `[feature03-transplant]/fix2.md` |
| Shape 微调工作流 + 调试入口 | `skills/fine-tuned-shapes.md` |
