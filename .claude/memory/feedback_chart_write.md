---
name: feedback_chart_write
description: PPT chart 写入——分发场景强制从零制表，禁用改模板路线
type: feedback
---

分发场景（模板 / 代码发给他人，数据他人填）下，PPT chart 必须走"从零制表 + OLE 粘贴"路线（`make_chart_for_{template}`），**禁用** `_write_chart` 原位改。

**Why:** fix3（图表写入诊断）.md 多轮双机诊断 + fix4（图表路线切换）.md 路线决策：

1. `BreakLink` 不是保护伞，是凶手。它会把健康 chart 写成僵尸态（readback=[], bars 消失），且不可逆。
2. `ChartData.Activate()` 在 Office Build 4266 抛 `DISP_E_EXCEPTION(-2147352567)`，且触发 GUI 弹窗"链接文件不可用"，阻塞脚本。
3. XML surgery 路径彻底死：办公室默认加密 pptx → CFB 复合文件容器，`zipfile` 无法打开。
4. chart COM 名会在重建后漂移（`Chart 13` → `Chart 27`），硬编码名字查找不可靠。
5. "100% 视觉还原" 与 "分发给他人 + 他人填数据" 物理上互斥——任何原地改 chart 实现都要求 chart 内部状态在分发链路里保持稳定，而他人填数据这个动作必然破坏这个稳定性。
6. `make_chart_for_questionnaire`（xlwings 新建 3D bar → `api[0].Copy()` → `mc_slide.Shapes.Paste()` → `CutCopyMode = False`）在办公室多年生产零报错。

**How to apply:**

- **yzr_ppt.py（已落地 fix4）**：chart 分支用 `_prepare_yzr_chart_data` + `make_chart_for_yzr`，不走 `_write_chart`
- **新模板移植**：若模板要分发给同事，chart 分支仿 `make_chart_for_yzr` 写专属 `make_chart_for_{name}`，放在 `src/_ppt_shared.py`
- **zxh_ppt.py**：当前仍用 `_write_chart`（未分发场景）；若未来要分发给同事，必须切到从零制表路线
- **Pipeline `_write_chart`**（`pipeline/03b_build_ppt_com.py`）：保留（仅用于新模板分析，本身是单机场景）
- **chart shape 查找**：优先按 `HasChart=True` 定位，不写死 `"Chart 13"` 这种会漂移的名字

**关键硬规则复用：**

- `Shapes.Paste()` 后必须 `CutCopyMode = False` 断 OLE 热链接（CLAUDE.md §3 规则 1）
- 删行前先 `chart.delete()`，否则 Excel 弹"错误公式引用"弹窗（CLAUDE.md §3 规则 3）
- **临时数据行保留不删**（用户实测经验）：即使 `CutCopyMode = False` 已执行，删除 Excel 端临时数据行仍会导致 PPT 端 OLE chart 数据丢失（bars 消失）。只删 Excel 端 chart 对象即可。对照 Mc-debug-4.md line 744。
- 清理顺序：`DisplayAlerts = False` → `chart.delete()` → 恢复 DisplayAlerts（**不**包含删 rows 步骤）

**参考实现：**

- `src/_ppt_shared.py::make_chart_for_yzr`（fix4 新增）
- `src/Function_030.py::make_chart_for_questionnaire`（范式来源，line 1999）
- 调用清理模式：`src/Function_030.py` line 399-427

---

## 3D chart 视角必须显式设置（2026-04-24）

xlwings 创建的 3D chart（ChartType=60）**默认视角 ≠ 用户期望视角**。即使 ChartType 相同，xlwings 默认 Elevation/Rotation 与 PPT 模板原 chart 不一致，导致 OLE 粘贴后视觉漂移。

**必须在 make_chart_for_{name} 里显式设置以下属性：**

```python
_ch = mc_chart1.api[1]
_ch.RightAngleAxes = True    # 直角坐标轴
_ch.AutoScaling = True       # 自动缩放
_ch.Elevation = 20           # X 旋转
_ch.Rotation = 15            # Y 旋转
_ch.Perspective = 0          # 透视
_ch.DepthPercent = 100
_ch.HeightPercent = 100
```

**PPT "三维旋转" 面板 ↔ Excel chart API 映射表：**

| PPT 面板 | Excel api | 方向 |
|--|--|--|
| X 旋转 | `Elevation` | 绕水平轴翻转（正值=俯视） |
| Y 旋转 | `Rotation` | 绕垂直轴旋转（正值=右旋） |
| Z 旋转 | —（不暴露） | — |
| 透视 | `Perspective` | RightAngleAxes=True 时通常忽略 |
| 直角坐标轴 | `RightAngleAxes` | True/False |
| 自动缩放 | `AutoScaling` | True/False |
| 深度 | `DepthPercent` | 20-2000 |
| 高度 | `HeightPercent` | 5-500 |

**调参工作流**：用户手工在 PPT 调到满意 → `read_selected_shape.py` 读 L/T/W/H → "设置形状格式 → 效果 → 三维旋转" 面板抄录 → 按映射表回写 Python。

**兜底**：所有 3D 属性用 `try/except` 包裹，失败时日志告警但不崩溃。

---

## `Shapes.Paste()` 返回 ShapeRange 陷阱（2026-04-27）

`mc_shape = mc_slide.Shapes.Paste()` 返回的是 **ShapeRange**，不是 Shape。

- `.Left/.Top/.Width/.Height` 会 fan-out 到内部 shape，所以这些代码不报错
- 但 `.Chart` / `.HasChart` **不在 fan-out 列表**，访问时抛 `com_error -2147352567 发生意外`
- 这就是为什么 `mc_shape.Chart.SetElement(0)` 看似有效但 chart title 一直不消失——错误被外层 `try/except` 静默吞掉了

**正确写法**：
```python
_shape_one = mc_shape.Item(1) if hasattr(mc_shape, "Item") else mc_shape
_shape_one.Chart.HasTitle = False
_shape_one.Chart.SetElement(0)
```

---

## chart 主标题双保险隐藏（2026-04-27）

OLE 粘贴到 PPT 后，单点调用经常因 COM 时序问题失败。**双保险写法**：
```python
chart.HasTitle = False     # 属性直写，最直接
chart.SetElement(0)        # UI 命令，等价于点击"图表元素 → 标题 → 无"
```

实测两者中任意一个能成功就解决，二选一不可靠（不同环境/不同 chart_type 表现不同）。

---

## bar chart 数值轴 max = 量表 max + 1（2026-04-27）

`MaximumScale = _scale_max`（5 分制→5，10 分制→10）会让 score=max 的 bar 末端被数据标签压住。统一改为 `_scale_max + 1`：
- 5 分制 → MaximumScale = 6
- 10 分制 → MaximumScale = 11

已落地：`Function_030.py::make_chart_for_questionnaire`。
未落地：`_ppt_shared.py::make_chart_for_yzr`（当前硬编码 10，未分发未触发问题；分发前改）。
