# Skills 文件夹索引

开发/调试 PPT 模板时的辅助工具与规范文档。

---

## 文件一览

| 文件 | 类型 | 作用 |
|--|--|--|
| `read_selected_shape.py` | Python 脚本 | 读取鼠标选中 shape 的完整信息（名称/坐标/文本/图表/图片） |
| `(legacy) diagnose_chart_write.py` | Python 脚本（历史） | 诊断 chart COM 写入跨机兼容性（STRAT 1-6 分级测试）；fix4 之后不再是主路径 |
| `fine-tuned-shapes.md` | 规范文档 | Shape 位置微调工作流（何时/如何在 xxx_ppt.py 插入硬编码坐标） |
| `port_handoff_checklist.md` | 规范文档 | Pipeline → /developer 移植衔接 Checklist（plan3 5 阶段流程的阶段 ④ 详细手册） |
| `memory-junction-3account.md` | 规范文档 | 3 账号 auto-memory NTFS junction 同步方案（含移植到新项目/新机器的步骤） |
| `memory_union_merge.py` | Python 脚本 | 扫描 N 账号 auto-memory，求 union 写入 repo 内 `.claude/auto-memory/`（dry-run / apply） |
| `memory_junction_setup.py` | Python 脚本 | 备份 + rmtree + `mklink /J` 给 N 账号 memory 建 junction（一次性） |
| `memory_junction_rollback.py` | Python 脚本 | `os.rmdir` 解 junction（**不是 rmtree**）+ 从备份恢复，可单账号回滚 |
| `memory_junction_verify.py` | Python 脚本 | 检测 junction 状态 + 端到端写/删传播测试 |

---

## 1. `read_selected_shape.py`

**用途**：开发新模板适配时，快速查 shape 的 COM 名称、坐标、文本内容、图表/图片元数据。

**用法**：
```bash
# 1) 在 PowerPoint 里选中一个或多个 shape（可多选）
# 2) 运行：
python skills/read_selected_shape.py
```

**输出**：纯文本到 stdout，包含：
- Shape 类型（AutoShape / Chart / Picture / TextBox 等）
- Name（**COM 内部名**，与中文 UI "选择窗格" 显示名可能不同）
- Left / Top / Width / Height（单位 points）
- 文本框：段落/字体/颜色
- Chart：ChartType / 系列数据 / IsLinked
- Picture：原生尺寸 + 裁剪框

**典型场景**：
- 写新模板的 `{NAME}_SHAPES` 常量时确认 shape 名
- 微调坐标前读取模板基准值（配合 `fine-tuned-shapes.md`）
- 调 bug 时检查 PPT 内存状态（chart IsLinked / 文本字体等）
- 读 3D chart 的坐标后，配合 PPT "设置形状格式 → 效果 → 三维旋转" 面板抄录 Elevation/Rotation 等视角参数，回写到 `make_chart_for_{name}` 的 3D 视图块（详见 `.claude/memory/feedback_chart_write.md`）

---

## 2. `(legacy) diagnose_chart_write.py`

**用途**：诊断 PPT chart 写入的跨机兼容性（Office Build 版本差异 / 加密环境差异）。

**注意（fix4 之后）**：生产已切到"从零制表 + OLE 粘贴"路线（`make_chart_for_yzr`），此脚本**不再是修 bug 主路径**，保留作为历史诊断工具。

**用法**：
```bash
# 推荐（最小污染）：只跑 STRAT 1 裸写入
python skills/diagnose_chart_write.py --strat1

# 完整诊断（会污染 chart 状态，慎用）
python skills/diagnose_chart_write.py --all
```

**前置**：PPT 打开 fresh 模板 + 选中 chart shape。

**STRAT 清单**：
| STRAT | 写入手段 | 备注 |
|--|--|--|
| 1 | `series.Values = tuple([...])` 裸写 | 最小化，fresh chart 上 work |
| 2 | VARIANT SAFEARRAY | 同上 |
| 3 | 写入 + `chart.Refresh()` | 同上 |
| 4 | `BreakLink` → 写入 | ⚠️ 会破坏 healthy chart（fix3 坑 2） |
| 5 | `Activate` → `Workbook.Sheets(1)` 写 cell | ⚠️ Build 4266 抛 DISP_E + GUI 弹窗 |
| 6 | XML surgery（zipfile） | 🔴 加密 pptx 不可用（fix3 坑 4） |

**历史结论**：见 `[feature03-transplant]/fix3（图表写入诊断）.md` + `fix4（图表路线切换）.md`。

---

## 3. `fine-tuned-shapes.md`

**类型**：规范文档（不是脚本）。

**用途**：定义"用户要求微调某个 shape 位置"时的标准工作流——在 `xxx_ppt.py::make_xxx_slide()` 中插入硬编码的 Left/Top/Width/Height。

**触发场景**：用户说"帮我把 Rectangle 68 往左移一点"、"XX shape 尺寸再大些"。

**关键点**：
- 代码块必须加 `#fine_tuned` 注释标记
- 插入位置：Clone Slide 之后、`for spec in XXX_SHAPES:` 循环之前
- 基准值**从标准模板读**（`src/Template 2.1.pptx`），不从已生成的输出文件读
- 保护用户已微调过的值，不要用模板值覆盖

详细流程、当前已微调 shape 表、单独调试入口说明见文档本体。

---

## 4. `memory-junction-3account.md` + `memory_*.py`（4 脚本套件）

**类型**：规范文档 + Python 脚本套件。

**用途**：解决多个 Claude Pro 账号轮换工作时，每个账号下 `<账号根>/projects/<项目>/memory/` 各自独立漂移的问题。用 NTFS junction 把 N 个账号的 memory 目录物理合一到 repo 内 `.claude/auto-memory/`，进 git，永久杜绝漂移。

**触发场景**：在新项目 / 新机器上想复用本项目的 3 账号 memory 同步方案。

**关键点**：
- 4 个 Python 脚本**项目无关**——`detect_project_name()` 自动从脚本所在路径推断项目名，复制到新项目无需改代码
- 账号配置在脚本顶部的 `ACCOUNTS` 字典，账号数 / 别名 / 路径都可改
- Junction 是 OS 文件系统层重定向，Claude Code 完全无感知，无需改它的默认行为
- **关键安全规则**：`shutil.rmtree(junction_path)` 会穿透 junction 删 repo 本体；删除时必须用 `os.rmdir`（参考 `claude_migrate.py:_safe_clean_target_project()`）
- `auto-memory` 进 git 后跨设备自动同步；新机器只需重跑 `memory_junction_setup.py --apply` 重建本机 junction

完整移植清单（新项目 / 新机器 / 不同账号配置）+ 故障恢复表见文档本体。

---

## 新增 skill 的约定

- Python 脚本放 `skills/*.py`，顶部 docstring 写清楚：用途 / 前置条件 / 用法 / 输出
- 规范文档放 `skills/*.md`，采用 frontmatter（`name` / `description` / `type`）便于记忆系统引用
- 新增后**在本索引文件登记一条**（文件/类型/作用 + 一段简述）
