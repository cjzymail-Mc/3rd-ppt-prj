# New PPT Workflow（高精度工程化 v3.0）

> 目标：把“能生成”升级为“稳定生成且高保真交付（98%+）”。

---

## 一、为什么你已经拆得很细，但效果仍不理想？

你上一版流程已经覆盖了主链路（识别→分析→构建→测试→迭代），但实际效果不理想通常来自以下 **4个系统性缺口**：

1. **形状匹配策略不够严格**
   - 仅靠 `shape.Name` 或枚举顺序匹配，容易在模板变化后错位。
   - 某些 shape 名称为空、重复、或在组内，导致替换对象漂移。

2. **“内容构建”与“版式约束”没有绑定**
   - 构建函数只关心“语义正确”，没有强制“字符预算/行数预算/段落预算”。
   - 结果就是：内容可读性下降、溢出、换行异常、视觉破坏。

3. **Diff 测试缺少“业务可读性”指标**
   - 只比几何/字体/颜色，仍可能出现“文本很长但样式一致”而被误判为通过。
   - 没有把“文本长度偏差、行数偏差、关键信息覆盖率”纳入硬门槛。

4. **迭代闭环缺“自动修复决策层”**
   - 有 diff，但没有把“差异 -> 应该改哪个函数 -> 怎么改”结构化。
   - 导致每轮修复仍偏人工经验，收敛慢。

---

## 二、建议继续拆分的环节（重点）

下面是我建议再拆分的任务包（你现在的流程在这些地方还可以更细）：

### A. Step1（shape识别）再拆成 3 个包

#### A1. shape 指纹构建包（Fingerprint）
为每个目标 shape 生成稳定指纹：
- `name`
- `type`
- `left/top/width/height`（四舍五入）
- `z-order`
- `text_prefix`
- `in_group`

> 用指纹匹配替代“单一 name 匹配”。

#### A2. 组内 shape 解析包（Group Resolver）
- 单独展开 group 内部元素并建立路径标识（如 `GroupA/TitleText`）。
- 避免 ParentGroup/GroupItems 导致的定位漂移。

#### A3. 模板漂移检测包（Template Drift Gate）
- 先检测模板页关键锚点是否变化（新增/删除/重排）。
- 漂移超过阈值直接阻断，防止后续“误替换”。

---

### B. Step2（shape分析）再拆成 3 个包

#### B1. 字段映射包（Field Mapping）
- 每个 shape 明确：来源字段、过滤规则、聚合方式。
- 产出结构化映射：`shape -> source columns -> transform`。

#### B2. Prompt 规格包（Prompt Spec）
- 给每个 AI shape 单独定义 prompt 模板：
  - 输入字段
  - 口径
  - 字数上限
  - 禁止词
  - 输出格式

#### B3. 可读性预算包（Readability Budget）
- 每个 shape 定义硬预算：
  - max chars
  - max lines
  - max bullets
- 构建函数必须遵守预算，不满足则自动压缩重写。

---

### C. Step3（内容构建 + 生成PPT）再拆成 4 个包

#### C1. 9函数构建包（Builder Functions）
- 9个函数只负责“返回内容对象”，不直接写PPT。
- 返回统一结构：`{text, bullets, chart_data, metadata}`。

#### C2. 内容校验包（Content Validator）
- 写入前先校验：空值、过长、非法字符、预算超限。
- 不通过则触发自动降噪/压缩。

#### C3. 写入执行包（PPT Writer）
- 专门负责 COM 写入：
  - 文本仅写 `TextFrame.TextRange.Text`
  - 图表只更新 series/categories
  - 保留格式不重置

#### C4. 写后回读包（Post-write Readback）
- 写入后立刻回读 shape 的实际文本/图表数据，确认“写入成功且未被PPT自动改写”。

---

### D. Step4（diff测试）再拆成 3 个包

#### D1. 几何样式 diff 包（Visual Diff）
- 位置/大小/字体/颜色/图表类型（你已有基础）。

#### D2. 可读性 diff 包（Readability Diff）
- 文本长度比
- 行数比
- bullet 数量比
- 文本相似度（带阈值）

#### D3. 业务语义 diff 包（Semantic Diff）
- 检查关键字段是否覆盖：样本量、关键词、建议项是否都出现。
- 不满足则直接 fail。

---

### E. Step5（迭代修复）再拆成 2 个包

#### E1. 修复路由包（Fix Router）
- 将 diff 项自动映射到具体函数：
  - 文本超长 -> `build_shape_x`
  - 图表异常 -> `chart_update_x`
  - 位置偏差 -> 写入定位函数

#### E2. 收敛控制包（Convergence Control）
- 每轮限定可改动范围，避免“修 A 伤 B”。
- 记录每轮得分曲线，判断是否进入稳定收敛。

---

## 三、建议的新标准流程（可直接执行）

### Phase 0：预检
1. 环境预检（Office/COM/xlwings）
2. 输入文件预检（模板、Excel、sheet）
3. 模板漂移预检（锚点）

### Phase 1：识别与映射
4. 生成 `shape_detail_com.json`
5. 生成 `shape_fingerprint_map.json`
6. 生成 `shape_analysis_map.json`
7. 生成 `prompt_specs.json`
8. 生成 `readability_budget.json`

### Phase 2：内容构建与写入
9. 运行 `build_shape_content`（9函数）
10. 运行内容校验器
11. 运行 COM 写入器（克隆标准页 + 替换内容）
12. 写后回读校验

### Phase 3：质量门禁
13. 视觉样式 diff
14. 可读性 diff
15. 语义覆盖 diff
16. 汇总总分与 fail 列表

### Phase 4：迭代与交付
17. 自动修复路由
18. 定向修复并重跑 13~16
19. 版本递增保存（1.1/1.2/1.3）
20. 人工验收与交付归档

---

## 四、关键产物清单（建议新增）

除了你当前已有文件，建议新增：

- `shape_fingerprint_map.json`
- `prompt_specs.json`
- `readability_budget.json`
- `content_validation_report.md`
- `post_write_readback.json`
- `diff_semantic_report.md`
- `iteration_history.md`

---

## 五、验收门槛（建议）

必须同时满足：

1. **Visual Score ≥ 98%**
2. **Readability Score ≥ 95%**（长度/行数/结构）
3. **Semantic Coverage = 100%**（关键字段不缺失）
4. **无高优先级 fail 项**

任一不满足，禁止进入“最终交付”。

---

## 六、执行建议（落地优先级）

优先级建议：

1. 先做 `readability_budget + content_validator`（最快改善“不可读”）
2. 再做 `diff_semantic`（避免“看起来像但信息错”）
3. 最后做 `fix_router`（加速多轮收敛）

这样你能最短路径提升产出质量，而不是一次性大改全部链路。
