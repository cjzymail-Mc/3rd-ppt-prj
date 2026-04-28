# fix3（图表写入诊断）.md — chart 写入 bug 修复计划

> **状态**：凶手已定位（BreakLink/Activate），待 fresh 模板双机验证 STRAT 1
> **最后更新**：2026-04-24（经多轮诊断 + 路线重估，XML surgery 因加密约束废弃）

---

## Context

`yzr_ppt.py` 生成评测页时需要把问卷 7 个指标的均值写入模板 slide 15 上的 3D 条形图
（ChartType=60，COM 名 `Chart 13`，中文 UI 显示"图表 44"→ 后重建为 `Chart 27`）。

约束（用户场景）：
- 办公室电脑**默认加密所有 Office 文件**（CFB 容器，非 zip）
- 模板 pptx + 代码分发给同事，**数据源永远由同事自己填**（即 embedded workbook 外链 100% 丢失）
- 项目约定：保留模板 chart 视觉样式（3D / 颜色 / 坐标轴），只改数据

---

## 踩过的坑（时间顺序）

### 坑 1 — 误判"Build 4266 的 COM 接口坏了"
- **现象**：同事机器（Office Build 4266）上 `series.Values = tuple` 写入后 readback=[]，chart 清空
- **错误结论**：Office 旧版损坏，只能走 XML surgery
- **真相**：chart 已被**我们自己代码历史上的 BreakLink 污染成僵尸态**，之后任何 COM 写入都静默失败

### 坑 2 — BreakLink 是凶手，不是保护伞
- 生产代码 `_write_chart` 里 `if is_linked: chart.ChartData.BreakLink() + Activate×3`
- 实测：**BreakLink 会把 healthy chart 弄成僵尸**（readback=[]，bars 消失）
- 双机诊断证据：
  - 同事 Run 2（手工重建后的 fresh chart）：STRAT 1/2/3 ✅，紧跟的 STRAT 4 BreakLink → 立即清空
  - 用户机器：同样 STRAT 1-3 ✅，STRAT 4 BreakLink 也破坏

### 坑 3 — Activate 是次凶，且会触发 GUI 弹窗
- 双机 `ChartData.Activate()` 均抛 `DISP_E_EXCEPTION(-2147352567, '发生意外')`
- 同事机器还会弹"链接文件不可用"对话框 → 脚本被阻塞

### 坑 4 — XML surgery 路径整条废掉
- 原计划：`zipfile` 直接改 `ppt/charts/chart1.xml` 的 `<c:numCache>/<c:strCache>`
- 实测同事机器 STRAT 6：`zipfile.BadZipFile: File is not a zip file`
- 初以为是 Save 后磁盘刷新延迟，后来用户提醒**办公室默认加密**
- 加密 pptx 是 **CFB 复合文件**，不是 zip → `zipfile` 本质上读不了，**这条路死**

### 坑 5 — 诊断脚本自污染
- 早期 `diagnose_chart_write.py` 默认顺序跑 STRAT 1 → 5
- STRAT 4 的 BreakLink 会把 chart 弄坏，导致后续 STRAT 5 / 6 的结果不可信
- 修复：默认模式改成 `--strat1`（只跑裸写入，不污染）

### 坑 6 — 模板 chart 的 shape 名随重建变化
- 原 COM 名 `Chart 13`（对应中文 UI "图表 44"）
- 用户手工删除重建后 → `Chart 13` / `Chart 27`（不固定）
- 生产代码若用硬名 `YZR_SHAPES[{"name": "图表 44"}]` 会找不到 shape
- 修复方向：按 slide 上"第一个 HasChart=True 的 shape"定位，不写死名字

### 坑 7 — 中文 UI 下"选择窗格"名和 COM 内部名不同
- 中文 UI 显示 "图表 13"，COM `Shape.Name` 返回 "Chart 13"
- 调试要以 COM 名为准，不能照抄选择窗格

---

## 路线 / 技术手段评估

| 手段 | 视觉还原度 | 加密兼容 | 跨机稳定 | 状态 |
|--|--|--|--|--|
| **A. COM 纯 STRAT 1**（`series.Values = tuple`，无 BreakLink/Activate） | 100%（保留模板） | ✅ | **待 fresh 模板验证** | 🟡 候选 1 |
| **B. make_chart + OLE 粘贴**（`Function_030.make_chart_*` 路线） | 95%（xlwings 复刻 3D bar 样式） | ✅ | ✅ 办公室多年稳跑 | 🟢 候选 2（兜底） |
| ~~C. COM BreakLink + Activate~~ | — | — | — | 🔴 废弃（凶手） |
| ~~D. XML surgery（zipfile）~~ | 100% | ❌ 加密 CFB | — | 🔴 废弃（加密约束不可克服） |
| ~~E. 手动 patch 步骤（方案 X）~~ | 100% | ❌ 加密同上 | — | 🔴 废弃 |

---

## 修复计划（分阶段，按候选 A → B 决策）

### 阶段 0 — 诊断脚本就绪（已完成）
- ✅ `skills/diagnose_chart_write.py` 默认 `--strat1` 模式（最小污染）
- ✅ 日志分离：写入前状态 / 写入后 readback / 视觉验收标准

### 阶段 1 — 验证候选 A（进行中）

**实验条件**：
1. 用 **fresh 模板**（未被历史 BreakLink 污染）—— 从 git HEAD 重新签出 `src/Template 2.1.pptx`
2. 打开 → slide 15 → 选中 3D 条形图
3. 双机分别跑 `python skills/diagnose_chart_write.py`

**验收**：
- ✅ 通过：readback=[1..7]，bars 肉眼可见 1/2/3/4/5/6/7 → 进入阶段 2
- ❌ 失败：chart 清空或 readback 为空 → 跳到阶段 3（候选 B）

### 阶段 2 — 候选 A 通过时的生产代码改造

**改动清单**：

| 文件 | 改动 |
|--|--|
| `src/_ppt_shared.py::_write_chart` | 删除整段 `if is_linked: BreakLink + Activate×3` 分支，只保留 `series.Values = tuple(values); series.XValues = tuple(labels)` |
| `src/yzr_ppt.py::YZR_SHAPES` | Chart 查找逻辑改为"按 HasChart 找第一个"，去掉硬编码名字 |
| `src/_ppt_shared.py::_write_chart` | 验证逻辑简化：只看 readback 不为空即视为成功，不再卡"首值误差 0.05" |

**注意**：同事拿到模板第一次跑前，**模板必须是 fresh 状态**。生产文档需要加一条说明"若 chart bars 消失，复制一份 git HEAD 的模板覆盖"。

### 阶段 3 — 候选 A 失败时的候选 B 兜底

**设计**：仿 `Function_030.make_chart_for_questionnaire`，为 yzr 写 `make_chart_for_yzr`。

**改动清单**：

| 文件 | 改动 |
|--|--|
| `src/_ppt_shared.py` | 新增 `make_chart_for_yzr(mc_cell, mc_slide, Left, Top, Width, Height)`：xlwings 建 3D 条形图（`ChartType = 60` 或 `'3d_bar_clustered'`），7 指标，0-10 量程，隐藏图例/网格/坐标轴 |
| `src/yzr_ppt.py::make_codex_slide` | Chart 处理改为：读模板 Chart shape 的 L/T/W/H → 删除 shape → 调 `make_chart_for_yzr` → OLE paste → 还原 L/T/W/H → `CutCopyMode = False` |
| `src/_ppt_shared.py::_write_chart` | yzr 不再走此函数（zxh 保留兼容） |

**风险**：xlwings 3D bar 样式需调试对齐模板视觉（颜色、3D 倾角、柱宽）。首次迭代成本中等。

---

## 不动哪些代码

- 不动模板 pptx 文件（用 git HEAD 版本）
- 不做机器检测分支（双机走同一代码路径）
- 不引入 `python-pptx`（项目禁用）
- zxh_ppt.py 的 chart 流程暂不动（等 yzr 定稿后再评估）

---

## 验收

- **候选 A 双机视觉一致** = 最佳验收
- **候选 B 双机视觉一致 + 样式 95%+ 相似模板** = 可接受兜底
- 无论哪条路，最终 `yzr_ppt.py` 端到端跑一次 + 同事机器端到端跑一次，bars 正确 = 验收通过

---

## 当前待办

- ⏳ 用户：git 签出 fresh 模板，双机跑 `python skills/diagnose_chart_write.py`
- ⏳ 根据结果走阶段 2 或阶段 3
- （未来）固化经验到 `.claude/memory/feedback_chart_write.md`：
  - BreakLink 是凶手不是保护伞
  - Activate 在 Build 4266 抛 DISP_E_EXCEPTION，且触发 GUI 弹窗
  - 加密环境下 XML surgery（zipfile）路径整体不可用
  - chart shape 名中英文差异 + 重建后名字会变
