# Repo Scan Result

## 🔍 快速索引（Quick Index）

**核心技术特征（Core Tech Strategy）**：
- [cite_start]**PowerPoint 控制**: **纯 COM 自动化** (`win32com`) [cite: 1, 2, 5]。
    - [cite_start]*特点*: 摒弃 `python-pptx`，直接映射 VBA 对象模型。实现了对 PPT 对象的像素级排版 (`Left/Top`)、层级控制 (`ZOrder`) 和剪贴板交互 (`Copy/Paste`) [cite: 1, 2]。
- [cite_start]**Excel 控制**: **Hybrid 模式** (`xlwings` + `COM API`) [cite: 2, 5]。
    - [cite_start]*特点*: 使用 `xlwings` 处理数据读取/写入，但通过 `.api` 属性直接调用底层 COM 接口处理图表细节（如图例隐藏、坐标轴删除、系列染色） 。

**核心类（按功能分类）**：
- [cite_start]`Text_Box` (基类) @ `src/Class_030.py` — 封装 PPT COM 对象 (`Shape.TextFrame.TextRange`)，实现对微软雅黑/Arial 字体、行距的精细控制 。
- [cite_start]`Line_Shape / Circle_Shape` @ `src/Class_030.py` — 直接调用 `Slide.Shapes.AddLine/AddShape` 绘制矢量图形，用于矩阵图的辅助线和标记 。

**核心函数（关键逻辑）**：
- [cite_start]`make_chart` @ `src/Function_030.py` — **COM 图表引擎**。先用 `xlwings` 创建图表，随即用 `.api` (VBA) 清洗格式（删除网格线/标题），最后调用 `.Copy()` 经剪贴板粘贴至 PPT 。
- [cite_start]`make_matrix` @ `src/Function_030.py` — **坐标映射引擎**。不依赖 Excel 图表，而是读取数据后，直接在 PPT Slide 上通过计算像素坐标，用 `mc_pic` 和 `Text_small` 绘制散点和图片 。
- [cite_start]`mc_pic` @ `src/Function_030.py` — **图片排版引擎**。从 Excel 复制图片 (`Copy`)，粘贴至 PPT (`Paste`)，并操作 COM 对象的 `ScaleHeight` 属性防止失真 。
- [cite_start]`GPT_5` @ `src/Function_030.py` — LLM 接口封装，含 Proxy 自动检测、OpenRouter 适配、对话历史管理 。

**主流程**：
- [cite_start]`Main Loop` @ `030 PPT Robot...py` — 顶层编排：初始化 COM 应用 -> 生成封面/目录 -> 遍历 Sheet (基础测试) -> 矩阵图 -> 问卷分析 -> 总结 -> 保存 。

## 📋 核心接口定义（API Interfaces）

- [cite_start]`make_chart(mc_sht, mc_slide) -> (Left, Top, Height, Width)` — 混合调用：xlwings 创建 -> COM 格式化 -> 剪贴板传输 -> PPT COM 定位 。
- [cite_start]`GPT_5(mc_prompt, model) -> str` — LLM 接口，封装了 OpenRouter 适配和系统代理自动检测 (`detect_system_proxy`) 。
- [cite_start]`search(mc_sht0, target, row_offset, column_offset) -> Range` — xlwings 封装的查找功能，用于在 Excel 中定位 "关键锚点" 。
- [cite_start]`content_slide(mc_ppt) -> None` — 目录生成器，直接复制 PPT Slide 对象并修改 TextRange 内容 。

## 🔁 常见模式（Common Patterns）

- [cite_start]**VBA-to-Python Translation**: 代码逻辑几乎是 VBA 的 Python 移植版。例如使用 `Shape.Line.Visible = 0` 或 `Shape.Fill.Transparency = 1` [cite: 1, 2]。
- [cite_start]**Clipboard Bridge**: 数据/图表/图片从 Excel 到 PPT 的流转完全依赖系统剪贴板 (`.api.Copy()` -> `Slides.Paste()`)，而非文件保存读取 。
- [cite_start]**Pixel-Perfect Layout**: 摒弃自动布局，所有元素（图表、文本、形状）均通过硬编码的像素坐标（如 `Left=225, Top=100`）或相对偏移量进行绝对定位 。
- [cite_start]**Hybrid Excel Manipulation**: 读取数据用 `range.value` (xlwings)，操作图表样式用 `range.api...` (COM) 。

## 🛠️ 技术栈

- [cite_start]**PPT 引擎**: `pywin32` (`win32com.client`) — **关键路径，无第三方封装库依赖** 。
- [cite_start]**Excel 引擎**: `xlwings` (高层封装) + `COM API` (底层控制) 。
- [cite_start]**AI 核心**: `openai` (GPT-5/OpenRouter 接口) 。
- [cite_start]**环境依赖**: Windows OS, Microsoft Office (必须安装，因为依赖 COM 组件) [cite: 2, 5]。

## 📁 项目结构

```
项目根目录/
├── orchestrator.py              # Agent调度系统（6个PPT专业agent的协调器）
├── main.py                      # PPT生成主入口（原始脚本，初始化COM应用，调度主循环）
├── new-ppt-workflow.md          # PPT流水线执行规范（v4.0，最高优先级参考）
├── repo-scan-result.md          # 本文件（代码库分析）
├── 00-ppt.py                    # 统一执行入口（串联Step1~4 + 多轮迭代）
├── 01-shape-detail.py           # Step1: shape识别与指纹
├── 02-shape-analysis.py         # Step2: shape->源数据映射
├── 03-build_shape.py            # Step3A: 按shape角色构建内容
├── 03-build_ppt_com.py          # Step3B: 模板克隆+内容写入
├── 04-shape_diff_test.py        # Step4: 严格差异测试（三层门禁）
├── 2025 数据 v2.2.xlsx          # 数据源（问卷sheet）
├── src/
│   ├── Class_030.py             # PPT组件库：封装 win32com 的 Shape/TextRange 对象
│   ├── Function_030.py          # 核心逻辑：混合 xlwings 与 COM API 的图表/图片/GPT 处理
│   ├── Global_var_030.py        # 配置：定义 PPT 模板页码映射
│   └── Template 2.1.pptx        # 模板（第14页=空白，第15页=标准）
├── .claude/
│   ├── agents/                  # 6个PPT专业agent配置（01-arch ~ 06-secu）
│   ├── hooks/                   # Architect权限守卫
│   └── CLAUDE.md                # 项目规范
└── codex-legacy/                # 历史工作记录
```


## 🧩 核心模块

1.  **COM Orchestrator**: 主程序直接操作 `PowerPoint.Application` 和 `xlwings.books.active`，控制两个应用程序的交互 。
2.  **Hybrid Chart Formatter**: 在 Python 中写 VBA 逻辑 (`SetElement`, `Axes(2).Delete`)，实现对 Excel 图表样式的极致控制 。
3.  **Direct Shape Drawer**: 不通过图表引擎，直接在 PPT 画布上用 COM 接口 (`Shapes.AddLine/AddShape`) 绘制矩阵分析图 。
4.  **Content Injector**: 遍历 Excel 单元格，通过 GPT 生成文本，再通过 COM 接口写入 PPT 文本框 (`TextFrame.TextRange.Text`) 。

## 🏗️ 代码风格与架构

- **风格**: **Scripting over Engineering**。为了追求对 Office 的极致控制，牺牲了跨平台性，选择了直接操作 Windows COM 接口。
- **排版逻辑**: "硬编码坐标" + "相对偏移计算"。例如矩阵图通过计算 Excel 数值与像素的比例 (`delt_l`, `delt_t`) 来确定落点 。
- **异常处理**: 针对 COM 调用可能出现的剪贴板占用或未响应，包含 `time.sleep(delay)` 进行延时缓冲 。

## 🔗 依赖关系

- 脚本强依赖本地安装的 **Microsoft Excel** 和 **PowerPoint** (COM 接口提供方) 。
- `Function_030` 强依赖 `Class_030` 提供的 PPT 元素封装类来执行绘制操作 。

## 💼 关键业务逻辑

这是一个 **"人肉操作自动化"** (RPA-like) 的系统。它不生成 PPT XML 文件，而是像一个隐形的人：
1. 打开 Excel 和 PPT 软件 。
2. 在 Excel 中选中数据，让 GPT 分析 。
3. 在 Excel 中生成临时图表，用 VBA 命令修饰它 。
4. **复制**图表/图片 。
5. 切换到 PPT，**粘贴**，然后用鼠标（坐标代码）把它拖到准确位置 。
6. 甚至在 PPT 里直接画线、画圈（矩阵图逻辑） 。