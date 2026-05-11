---
name: Python stdout 中文乱码（Windows + Bash tool）
description: 在 Bash 工具里跑 python 输出中文（接管 Excel/PPT 诊断脚本），stdout 默认 GBK 会乱码；不能用 chcp/set 等 cmd 语法；正确做法是 PYTHONIOENCODING + io.TextIOWrapper 双保险
type: feedback
---

在本项目通过 Bash 工具跑一次性 python 脚本（典型场景：`win32com.GetActiveObject` 接管已打开的 Excel/PPT 做诊断）时，**输出含中文必须显式处理 stdout 编码**，否则会被 Windows 默认 cp936 / GBK 截断成乱码（`������Ϣ` 这种）。

**Why:** 历史踩坑（2026-05 修 `test_detail` off-by-one bug 时连吃 2 次失败才走通）：
1. 第 1 次直接 `python -c "..."` → stdout GBK，中文全乱码
2. 第 2 次试 `set PYTHONIOENCODING=utf-8 && chcp 65001 >nul && python ...` → `chcp: command not found`。**Bash 工具实际是 git bash，不是 cmd**，`set` / `chcp` 是 cmd 语法，在 git bash 里全部失效
3. 第 3 次双保险才成功

**How to apply:**

正确写法（git bash 兼容 + Python 内层兜底）：

```bash
PYTHONIOENCODING=utf-8 python -c "
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
# ... 你的代码 ...
"
```

两层都不能省：
- **外层 `PYTHONIOENCODING=utf-8`**：bash 前置环境变量语法，覆盖 Python 启动时的默认编码探测
- **内层 `io.TextIOWrapper(...)`**：兜底重新包装 `sys.stdout.buffer`，防止某些 Python 版本/启动路径下外层环境变量被忽略；`errors='replace'` 让无法编码的字符变 `?` 而不是抛 `UnicodeEncodeError` 中断脚本

**禁止**：
- `chcp 65001` —— cmd 命令，git bash 没这玩意
- `set PYTHONIOENCODING=utf-8` —— cmd 语法（git bash 用 `export` 或前置赋值）
- `$env:PYTHONIOENCODING="utf-8"` —— PowerShell 语法，Bash 工具默认不走 PS

**触发场景速判**：本项目里几乎所有"接管运行中的 Excel/PPT 跑诊断脚本"的命令都需要这个套餐——一旦看到 `xlwings.books.active` / `win32com.client.GetActiveObject` + 输出含中文 sheet 名 / 单元格值，立刻上双保险，不要先试裸命令再补救。

**Skills 里的等价工具**：`skills/read_selected_shape.py` 等已自带 stdout 包装；如果是写一次性诊断 `python -c`，用上面的双保险模板。
