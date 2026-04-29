# fix1：chart 标签精简 + 测试者单位归一 + Oval skip 修正

> 日期：2026-04-28
> 范围：`src/apparel_ppt.py`
> 触发：apparel 模板移植后首轮视觉验收（用户调试 `apparel_ppt.py`）

apparel 模板移植后用户用单页调试方式核对每个 shape，发现 3 个独立 bug：chart 数据标签太冗余、测试者基本信息单位混淆、4 个装饰圆圈被误写分数。本文记录三者的根因和修法。

---

## Bug 1：Chart value 标签冗余 「【腰围】版型」→ 「腰围」

### 现象
4 个 chart 都按分类（版型 / 面料 / 吸湿排汗 / 速干）切片，左下角已有大字标题"版型 / 面料 / ..."，但 chart 内每条 bar 的 value 标签仍然是「【腰围】版型」「【衣领】版型」等 —— 后缀的"版型"和左下角标题完全重复。

### 根因
`_extract_means_for_category` 的标签清理只去掉了"评分"二字和括号说明，保留了 `【...】+ 分类后缀` 的完整组合：

```python
# 旧逻辑
clean_label = re.sub(r'^\d+[、.]\s*', '', h)
clean_label = re.sub(r'（[^）]*）', '', clean_label).strip()
clean_label = clean_label.replace("评分", "").strip()
# 结果: "1、【腰围】版型评分（说明）" → "【腰围】版型"
```

### 修法
chart 已按 category 分组，value 标签只需保留差异点 —— 直接抓 `【...】` 内的字。无 `【】` 时回退旧逻辑兜底。

```python
m = re.search(r'【([^】]+)】', h)
if m:
    clean_label = m.group(1).strip()
else:
    # 兜底：无 【】 时退回旧清理逻辑
    clean_label = re.sub(r'^\d+[、.]\s*', '', h)
    clean_label = re.sub(r'（[^）]*）', '', clean_label).strip()
    clean_label = clean_label.replace("评分", "").strip()
```

`apparel_ppt.py:238-252`

---

## Bug 2：测试者基本信息单位混淆 (KG ↔ 斤, CM ↔ M)

### 现象
TextBox 24 显示 `A: 1CM / 100 KG`，明显两个单位都填错：1 应是 1m（=100cm），100 应是 100 斤（=50kg）。

### 第一版修法（不够稳）
单纯阈值：`weight > 110 → ÷2`、`height < 3 → ×100`。问题：100 ≤ 110 不触发，weight 100 没被修复。简单降阈值到 80 又会误伤真 80kg 男性测试者。

### 第二版修法：BMI 交叉验证
两步走：
1. **粗修**：身高 < 3 视 m，体重 > 110 视斤
2. **BMI 反推（细修）**：算 BMI；越界（不在 `[16, 32]`）则试 `weight ÷ 2`，若新 BMI 落入区间才采纳

测试用例（已通过）：

| 输入 | 输出 | 说明 |
|---|---|---|
| `(160, 100)` | `(160, 50)` | BMI 39 → ÷2 → BMI 19.5 ✓ |
| `(180, 90)` | `(180, 90)` | BMI 27.8 ✓，真胖子不误伤 |
| `(170, 130)` | `(170, 65)` | 粗修阶段已解决 |
| `(1.65, 52)` | `(165, 52)` | m → cm |
| `(1, 100)` | `(100, 100)` | 双错无解，保留 |

代码：`apparel_ppt.py::_normalize_height_cm / _normalize_weight_kg / _cross_validate_bmi / _normalize_person`

**关键**：prompt builder 和 fallback **都要先洗再喂下游**，否则 GPT 拿到脏数据会乱编（GPT 对 100kg 不觉得有问题）。

---

## Bug 3：4 个装饰 Oval 残留分数

### 现象
4 个虚线圆圈（Oval 3 / 13 / 16 / 19）按模板设计是装饰元素，**不该有文字**（分数已在分类标题旁单独显示）。但 PPT 输出页里每个圈内都有 4.2 / 4.9 / 5.0 / 5.0。

### 根因
APPAREL_SHAPES 早期把它们设成 `score_category_mean` 写"整体均值"，注释还说"Pipeline 标 skip 但视觉上是分类总分圆环"。这是 dev override 错了 —— 模板设计本意是空圈装饰。

### 修法
spec 改回 `skip`：

```python
{"name": "Oval 3",  "strategy": "skip"},  # 装饰虚线圆，不放文字
# 同样改 Oval 13 / 16 / 19
```

`apparel_ppt.py:130-149`

### 二次坑：skip 不清旧文本
改完发现 PPT 里 Ovals 仍有数字 —— 因为 `skip` 的语义是"代码不写新值"，**不会清空 shape**。
- 检查源模板 `Template 2.1.pptx` 的 slide 19，4 个 Oval 都是空的 ✓
- 之前那一轮 `score_category_mean` 已经把数字写到模板上 → 当前输出页里残留
- 通过 COM 一次性清空所有 Oval 文字（已执行）

**已沉淀到 CLAUDE.md 硬规则 + auto-memory `feedback_skip_vs_clear.md`**：移植新模板必查所有 `skip` shape 在源模板里也是空的。

---

## 验收

- Chart value 标签：`整体 / 衣领 / 袖口 / 胸围 / 腰围` 等单字段 ✓
- TextBox 24：`A: 100CM / 50 KG`（A 的源数据 1m 部分自动 ×100；100kg 自动识别为斤）
- 4 个 Oval：空 ✓
- 影响范围：仅 `src/apparel_ppt.py`（增/改约 80 行），无 shared 模块改动

## 经验沉淀

| 经验 | 沉淀位置 |
|---|---|
| skip ≠ clear，移植必查源 shape | `CLAUDE.md` 硬规则 + `feedback_skip_vs_clear.md` |
| BMI 交叉验证识别单位混淆 | `CLAUDE.md` 硬规则 + `feedback_unit_normalize_bmi.md` |
