---
name: 性能重写遗留代码先验证隐式行为
description: 重写老 xlwings/COM 函数做性能优化前，先实测验证原代码的隐式副作用（end('up')/selection 链路等），不要只看代码表面语义；典型表现是"歪打正着"——原作者没注释、新写法靠"清晰语义"反而漏数据
type: feedback
---

重写本项目老 xlwings / COM 函数做性能优化时（替换 `selection.end()` / `selection.offset()` 链路为显式行号 + bulk read），**改动前必须实测旧代码的边界单元格是否被处理**，不能只看代码"语义"推断它干了什么。

**Why:** 2026-05 `test_detail` 升级踩坑：commit `c9e248f` 把旧的 `selection.end('up')` + `offset(1, -1)` 链路替换成显式 `start_cell.row` 起步的 bulk read，commit message 写"速度提升 ~7N 倍"。但旧代码有个**没人注释过**的隐式副作用：

- 旧代码：`selection.end('up')` 从底部回到 col 5 顶部，由于 col 5 表头行（"测试项目"）也非空，`end('up')` 会**越过第一数据行直接到表头**，再 `offset(1, -1)` 落到 (first_data_row, marker_col)。这是"歪打正着"——靠表头非空帮它对上了起点
- 新代码：把 `start_cell.row` 当成"表头行" + 1，结果**首条数据行被吞掉**

后果：用户每次跑 Main.py 都漏处理首项「跑鞋基本参数测试」对应的 3 张 sheet（被 `if all_test_detail[2][s] in c.name` 静默跳过，不报错），半年后才发现。

**How to apply:**

接到「重写老 xlwings/COM 函数做性能优化」类任务时（典型：把 `selection.end()` / `mc_book.selection` 链路改成 `mc_sht.range((r, c)).value` bulk read），**改完先做边界实测，不要相信代码语义直觉**：

1. **找到原代码的迭代起止边界**（loop range、selection 起点），用 print 把每次迭代实际读取的 `(row, col, value)` 打出来
2. **改完新代码也打一份**，逐项 diff
3. 重点查首条/末条数据是否都在新版输出里——本项目老代码很多用 `selection.end('up'/'down')` 链路，**连续非空区段会越过表头/末尾**，新代码用显式行号时极易差 1
4. 不要被 commit message 的"~Nx 提速"诱导跳过验证步骤

**项目里的高风险候选（未来可能被同样重写）**：
- `Function_030.py::test_detail`（已踩 + 已修复）
- `Function_030.py::make_chart`（也用 selection.end 链路找 chart 区段）
- `Function_030.py::questionnaire_Excel`（行级遍历 + selection）

任何老 selection 链路重写前，**先按上面 1-3 步打数据对照表**，再 diff，再合并。不允许"看着改对了就提交"。

**配套规则**：见 `feedback_com_constraints.md` 的 COM 性能模式（bulk read > 单元格逐次读），和 `feedback_python_stdout_encoding.md`（边界实测打 print 时记得带 stdout 编码双保险，不然中文 sheet 名/cell 值会乱码看不出 diff）。
