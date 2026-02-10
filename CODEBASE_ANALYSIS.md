# 代码库分析报告

## 1. 项目结构
```
.
├── Main.py
├── src/
│   ├── Class_030.py
│   ├── Function_030.py
│   ├── Global_var_030.py
│   ├── Template 2.1.pptx
│   └── init.py
├── ppt模板 - 1.pptx
├── 问卷模板 - 1.xlsx
├── 篮球试穿问卷ppt - 1.pptx
├── repo-scan-result.md
└── your-role.md
```

## 2. 技术栈
- **语言**：Python（脚本式自动化）。
- **Office 自动化**：`win32com.client` 直接驱动 PowerPoint/Excel COM。 
- **Excel 读写**：`xlwings` + COM `.api` 混合调用。 
- **LLM 接入**：`openai` 客户端（含 httpx 依赖）。
- **桌面交互**：`tkinter` 弹窗提示。

## 3. 代码风格与架构模式
- **风格**：脚本化/RPA 风格，强调“像人一样操作 Office”的流程编排。
- **架构**：以 `Main.py` 为入口调度，`Class_030.py` 抽象 PPT 组件，`Function_030.py` 封装核心图表/排版/剪贴板操作。
- **排版方式**：硬编码像素坐标 + 相对偏移计算，依赖剪贴板将 Excel 图表/图片粘贴到 PPT。

## 4. 关键文件清单
1. `Main.py` - 主入口与流程编排。
2. `src/Function_030.py` - 核心业务逻辑（图表、矩阵图、内容生成）。
3. `src/Class_030.py` - PPT 组件/形状封装。
4. `src/Global_var_030.py` - 模板页码/全局配置。
5. `ppt模板 - 1.pptx` - 现有 PPT 模板。
6. `问卷模板 - 1.xlsx` - 现有 Excel 问卷模板。

## 5. 模块依赖关系
- `Main.py` → `src/Class_030.py`、`src/Function_030.py`、`src/Global_var_030.py`。
- `src/Function_030.py` → `src/Class_030.py`、`src/Global_var_030.py`。
- 运行时依赖：Windows + Microsoft Office（COM 自动化要求）。
