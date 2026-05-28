---
name: feedback_debug_protocol
description: 未来调试涉及 COM/OLE/模板的 bug 时，避免重蹈 fix3→fix4 绕弯路的实战流程
type: feedback
---

涉及 COM / OLE / 模板 / 多机分发的 bug，按本协议行动，避免"修 bug 变绕路"。

**Why:** fix3 用数日反复修 `_write_chart`（STRAT 1-6 全试），结果 fix4 切路线到 `make_chart_for_yzr` 几小时搞定。绕路的根因不是技术难度，是**流程缺失**——没有先搜项目、没有质疑用户"约定"的前提、同类技术连续失败仍在继续、用户给的路线信号被当成干扰。

**How to apply:** 本协议适用于任何"修改 PPT / Excel COM 对象" "涉及加密/跨机/分发" "有内部状态的组件（Chart/Table/SmartArt/OLE）"类 bug。

---

## 这次为什么绕弯路（4 条具体错误）

1. **没先搜项目有没有同类问题解决过**。`make_chart_for_questionnaire` 在 `Function_030.py` 已稳跑多年，接手 chart bug 的第一步本应是 `grep chart` 全项目，而不是直接改 `_write_chart`。
2. **把"用户的偏好"当成"硬约束"**。用户说的"约定 100% 还原模板"是偏好；"模板分发给同事 + 数据同事填 + 加密环境"是需求。需求 > 偏好。我把偏好当起点，没先做第一性判断。
3. **同一技术类别连续失败 ≥3 次仍在继续**。STRAT 1-6 本质都是"让 COM 写入成功"，每次失败只换变体。连续失败 2 次就应该跳出整个类别。
4. **把用户 pivot 当成干扰而不是证据**。用户提过"make_chart 多年稳跑"（证据），也问过"你怎么又走从零制表了"（困惑）。我读错信号，把证据当干扰。

---

## 未来遇到类似情况，按这个顺序做

### 🔵 拿到新 bug 的头 10 分钟（诊断前）

1. **grep 项目**：关键字 = 相同对象类型 + 相同操作动词。看有没有已被解决过的同类问题。
2. **读那段生产代码**：它怎么解决的？它的假设前提是什么？和当前 bug 场景的前提是否一致？
3. **问用户 3 个问题**：
   - 这个模板 / 代码会分发给别人吗？
   - 数据源在哪？会不会漂移 / 丢失？
   - 环境有什么特殊约束（加密、Office 版本、离线）？
4. **列约束清单（写下来）**：区分"偏好 vs 硬需求"，硬需求之间互斥的要立刻报警。

### 🟡 调试中的熔断器

5. **失败计数器**：同一技术类别（COM 某组 API / XML surgery / 某个库的一系列调用）失败 ≥2 次 → 停下来问"是不是类别错了"，不再换变体。
6. **"用户提到的另一个方案"是证据不是干扰**：写进候选列表，不是听过就算。

### 🟢 决策后

7. **路线变更必须书面记录**：为什么从 A 切到 B、放弃了 A 的什么、B 的代价是什么——写进 `fix{N}.md`。下次回滚时有依据，不会转圈。

---

## 5 条必须养成的反射动作（最小集）

| 触发 | 反射 |
|--|--|
| 接到一个涉及 COM / OLE / 模板的 bug | **第一步 grep，不是第一步改代码** |
| 用户用"我们之前约定"开头 | **立刻问"这个约定是在什么假设下达成的？当前场景假设还成立吗？"** |
| 同一方案连续失败 2 次 | **停下来写 3 个候选路线，不是继续第 3 次尝试** |
| 写 `try/except` 包裹关键操作 | **success 路径必须也打 print**，不能只在 except 里打——否则 except 静默吞错时，看日志根本不知道操作没执行（详见下方"silent except 反模式"） |
| 用户提"我选中的 / 我当前打开的 / 屏幕上的 X" | **先 `Glob skills/read_*`** 找现成桥接工具，禁止凭"默认 Claude 能力"先否认（详见下方"默认能力假设反模式"） |

---

## silent except 反模式（2026-04-27 chart title hide bug 复盘）

**触发场景**：`_ppt_shared.py::make_chart_for_yzr` 修复 chart title 不隐藏，加了：
```python
try:
    mc_shape.Chart.SetElement(0)   # 调用方
except Exception as _e:
    print(f"  PPT 端隐藏标题失败（{_e}）")
```

测试时认为"修好了"，第二天用户报告 title 仍然显示。

**真相**：`Shapes.Paste()` 返回 ShapeRange，`.Chart` 抛 `com_error -2147352567`。except 把错误吞了，但日志里那行 print 混在大量 yzr 调试输出里没被注意，看起来"什么都没发生 = 看似成功"。

**反模式特征**：
- try 里是关键操作（不是装饰性的）
- except 只打 print，不抛、不重试
- 没有 success 路径的 print
- 验证只看 PPT 视觉结果，不看日志

**正确写法**：
```python
try:
    mc_shape.Item(1).Chart.SetElement(0)
    print("  [yzr-chart] PPT 端主标题已隐藏")    # ← 关键：成功也打 print
except Exception as _e:
    print(f"  [yzr-chart] 隐藏失败（{_e!r}）")    # 用 !r 暴露 com_error 完整码
```

**额外收益**：日志成对出现"开始 → 成功/失败"，下次 grep 日志快速定位是哪一步漏了。

---

## 默认能力假设反模式（2026-04-29 read_selected_shape bug 复盘）

**触发场景**：用户问"读取我当前选中的 PPT shape"，我直接回答"我没有访问 PPT 实时状态的能力"，被用户当场指正——`skills/read_selected_shape.py` 早就在仓库里，注释"读取当前鼠标选中的 PPT shape 的完整信息"白纸黑字。

**真相**：本项目通过 `win32com.client.GetActiveObject("PowerPoint.Application")` 桥接到运行中的 Office 实例，**有完整能力**读取实时选中、当前 slide、所有 shape 属性。不去 grep 项目自带工具、不去 ls `skills/` 目录，凭"通用 Claude 默认能力边界"判断当下能不能做——这是错误的。

**反模式特征**：
- 用户消息出现"我选中的"、"我当前打开的"、"屏幕上的 X"等指代实时状态的表述
- 我没跑 `Glob` / `Grep` 就先回"我看不到 / 我做不到 / 我没有 X 能力"
- 用户反复纠正后才去查工具

**正确反射**：
1. **触发词**：消息里出现"我选中的"、"我当前的"、"刚才那个"、"屏幕上的"、"打开的"
2. **第一步**：`Glob skills/read_*` + `Grep "GetActiveObject\|active\|选中"`
3. **第二步**：找到对应脚本就直接 `python skills/xxx.py`，输出贴回对话
4. **第三步**：找不到才说"项目里没现成的，要我写一个吗？"
5. **绝对禁止**：在 step 1 之前用"默认能力边界"否认

**额外**：通过 Bash 工具跑 python 脚本输出中文乱码时，单加 `PYTHONIOENCODING=utf-8` 在某些 git bash 环境下还不够，需要 + `io.TextIOWrapper` 双保险，详见 `feedback_python_stdout_encoding.md`。

---

## 参考案例

- `[feature03-transplant]/fix3（图表写入诊断）.md` —— 绕弯路的全过程档案（7 个踩过的坑）
- `[feature03-transplant]/fix4（图表路线切换）.md` —— 路线切换后的决策与落地
- `【Mc-debug】/Mc-debug-4.md` line 1775-1827 —— 用户给出技术证据（make_chart 稳跑）但我没接住的现场
- `[feature03-transplant-II Apparel]/fix2（GPT语料库 + 死代码清理）.md` —— 2026-04-29 默认能力假设反模式现场记录
