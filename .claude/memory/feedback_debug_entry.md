---
name: xxx_ppt.py 单页调试入口
description: 每个模板 ppt.py 需有 __main__ 调试入口，自动打开模板不保存，连接已打开的 Excel
type: feedback
originSessionId: 41aa1118-8152-43ad-bbbe-6a66a3a736cd
---
每个 `xxx_ppt.py` 都必须有 `if __name__ == "__main__"` 调试入口，用户习惯单独调试一页而非跑整个 Main.py（5分钟）。

**Why:** Main.py 全流程太慢，调试单页功能时需要快速迭代。

**How to apply:**
- 文件顶部 `if __name__ == "__main__": sys.path.insert(0, ...)` 解决直接运行时的导入问题
- 用 `Dispatch` + `Presentations.Open` 打开模板（与 Main.py 同方式），不自动保存
- 用 `xlwings.books.active` 连接已打开的 Excel
- import fallback 需三层：`.xxx` → `src.xxx` → `xxx`（裸导入）
- 详见 `skills/fine-tuned-shapes.md`
