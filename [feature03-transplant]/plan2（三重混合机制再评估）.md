# plan2（三重混合机制再评估）.md — 三重混合机制再评估（基于 yzr_ppt 移植任务）

> **状态**：评估稿，待用户决策是否实施
> **评估依据**：fix3（图表写入诊断）.md / fix4（图表路线切换）.md / Mc-debug-4.md / Mc-debug-5-三重混合机制.md
> **写入时间**：2026-04-27

---

## Context

`Mc-debug-5-三重混合机制.md` 中提出了"Pipeline 50% + Agents 40% + Developer 10%"的加权合作设计。本文档基于刚刚完成的 yzr_ppt chart 移植任务（fix3→fix4 + 3D 旋转手动调参）的真实经历，做一次**不护航**的再评估。

---

## 1. yzr 任务的实际工作量来源（数据驱动）

| 工作量来源 | 占比 | 备注 |
|--|--|--|
| Pipeline 代码 | **-10%（负贡献）** | `_write_chart` 流入 src/ 后误导路线，是 fix3 绕弯路的结构性根因 |
| Agents 自检 | **0%** | chart 写入沉默失败（readback=[]、bars 消失）压根不抛异常，自检全程未触发 |
| Developer（Claude + 用户） | **110%** | 路线决策、`make_chart_for_yzr` 实现、3D 旋转手动调参全部在这层 |
| `Function_030.py` 多年生产代码 | **核心范式来源** | `make_chart_for_questionnaire` 才是真正的地基，比 Pipeline 早数年 |

**这个数据直接打脸 50/40/10 的权重设计。** 已知模板 + bug 修复场景里，Pipeline 不是地基，是绊脚石。

---

## 2. 三重混合制的真实价值（诚实清单）

### 仍然有价值的部分

1. **`src/_ppt_shared.py` 共享纯数据工具**（Fix 2 partial 的产物）—— 这次 `_extract_score_means` 等被 yzr 直接复用，确实减少了重复
2. **`fine-tuned-shapes.md` 的微调流程文档** —— 直接指导了 Chart 位置 + 3D 旋转参数的回写
3. **`read_selected_shape.py` 调试工具** —— 这次读 chart 坐标 + 3D 视角参数都靠它
4. **Developer Agent 的 Playbook（Fix 3）** —— 让"copy yzr 改 zxh"的移植变成有章可循

### 价值打折扣的部分

5. **Pipeline 步骤 1-2（shape 分析 + prompt 生成）** —— 仅在"全新陌生模板"时有用，已知模板（如 yzr/zxh）跑它是浪费
6. **Agents 自检循环** —— 对沉默失败类 bug（chart bars 消失）完全无效；只能抓"长度超限/字体错误"等浅层问题

### 暴露的设计漏洞

7. **Pipeline 代码流入 src/ 没有"适配性评审"机制** —— `_write_chart` 从 Pipeline 直接复制到 `_ppt_shared.py`，没人审过它在分发场景下是否合理。**fix3 绕弯路的结构性根因**
8. **三重混合制没有"分发场景"维度** —— 框架只考虑"开发-验证-移植"，没考虑"代码 + 模板分发给同事，他人填数据"。yzr 是分发链路里第一个真实案例
9. **没有"复杂度评估前置"机制** —— fix3 把 chart 修复当成几小时的事，结果耗数天。已在 `feedback_debug_protocol.md` 中用 7 步流程 + 3 条反射动作补上

---

## 3. 结论：框架不需要重构，需要重新解释

### 原版（误导性）

> 三重混合制 = Pipeline 50% + Agents 40% + Developer 10% **加权合作**

### 修正后（贴合现实）

> 三重混合制 = **三种工作模式的路由器**，根据任务类型分流

### 5 种任务场景 → 工具组合映射表

| 任务类型 | 主路径 | Pipeline 角色 | Agents 角色 |
|--|--|--|--|
| **已知模板 bug 修复**（fix3/fix4 这类） | Developer 100% | 不参与 | 不参与 |
| **已知模板日常运行**（Main.py 跑批） | src/ 100% | 不参与 | 不参与 |
| **新模板接入**（杨祖锐式） | Developer 主线 | 步骤 1 提供 shape 清单 | 不参与 |
| **完全陌生模板首次分析** | Pipeline 全流程 | 全流程 | 自检循环 |
| **批量生成多模板** | Pipeline 主线 | 全流程 | 自检循环 |

**这个分流逻辑其实 CLAUDE.md §1 双轨架构已经隐含了，但没有显式打成路由表**——导致每次接到任务时仍然要重新判断走哪条路。

---

## 4. 优化建议（按优先级）

### ★★★ 优化 1：CLAUDE.md §1 升级为"任务路由表"

把双轨架构从"两套并行系统"扩成"5 种任务场景 → 工具组合"显式映射表（如第 3 节所示）。

**收益**：接到新任务时第一件事是查表，不再猜。  
**成本**：~30 分钟，改动 CLAUDE.md §1，新增 5×3 表格。  
**紧急度**：高。每次任务开头都受益。

### ★★ 优化 2：Pipeline → src/ 流入路径加"适配性 checklist"

Fix 2 把 `_write_chart` 抽到 `_ppt_shared.py` 时漏了一步：**评审它的假设前提是否在 src/ 场景仍成立**。

具体做法：在 `developer.md` 里加一条移植 checklist：
- [ ] 这个函数在 Pipeline 里的假设是什么？（单机？fresh 模板？）
- [ ] 它流入 src/ 后的场景是什么？（分发？跨机？加密？）
- [ ] 假设有 mismatch 吗？

如果有 mismatch，要么改写，要么加 docstring 警告（fix4 给 `_write_chart` 加的警告是个好范例）。

**收益**：未来 Pipeline 代码流入不会再出 fix3 这种结构性误导。  
**成本**：~1 小时，改 `developer.md` + 在历史共享函数 docstring 上补 1-2 处场景警告。  
**紧急度**：中。下次有 Pipeline 代码流入 src/ 前必须做。

### ★ 优化 3：Agents 自检覆盖"视觉级回归"

只是登记，不紧急。当前自检是文本规则级（长度/字体），抓不住沉默失败。理想是加截图比对，但成本高。

**收益**：抓住 chart bars 消失之类的沉默 bug。  
**成本**：高（需引入 SSIM 或视觉 diff 工具，跨机加密环境下还要绕过截图限制）。  
**紧急度**：低。先记着，等出第二次类似 bug 再上。

---

## 5. 一句话结论

> 三重混合机制**仍然有价值**，但价值不在"三层加权合作"，在**"五种任务场景的路由分流"**。50/40/10 的权重设定误导性强，应当抛弃。
>
> 这次 yzr 任务的真正主力是 Developer（Claude + 用户），Pipeline 是绊脚石，Agents 缺席——这是**三重混合制最常见的实际配置**，不是异常。

---

## 6. 当前待办（按本计划）

- ⏳ 用户决策是否采纳本评估
- ⏳ 如采纳，按优先级实施优化 1-3
- ⏳ 完成后，更新本文件状态为 "implemented"

---

## 7. 参考档案

- `[feature03-transplant]/fix2（三重混合架构整改）.md` —— Fix 2 partial 的产物：`src/_ppt_shared.py` 由来
- `[feature03-transplant]/fix3（图表写入诊断）.md` —— chart 写入 bug 多轮诊断（绕弯路全过程）
- `[feature03-transplant]/fix4（图表路线切换）.md` —— 路线切换决策与落地（make_chart_for_yzr）
- `debug/Mc-debug-4.md` —— chart bug 现场对话记录（line 1775-1827 是关键节点）
- `debug/Mc-debug-5-三重混合机制.md` —— 三重混合机制原版描述与早期评估
- `.claude/memory/feedback_debug_protocol.md` —— fix3→fix4 血的教训固化（7 步流程 + 3 反射动作）
- `.claude/memory/feedback_chart_write.md` —— chart write 经验固化（含 3D 视图必须显式设置）
