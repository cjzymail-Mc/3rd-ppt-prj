# zxh_ppt.py 修复计划 — 代码级修复

> 记录于 2026-04-15
> 所有修复均在 `src/zxh_ppt.py` 中完成，不改模板 .pptx 文件。
> 内容全部由 GPT 生成 + COM 写入，因此从 prompt 和代码两个层面修复。

---

## 问题诊断（COM 实测 + pipeline 自检）

| # | 问题 | 证据来源 | 根本原因 |
|---|------|---------|---------|
| 1 | TextBox 17 超出右边界 70pt | COM 实测 R=1030 > 960 | 模板原始 L=550 W=480，克隆后未矫正 |
| 2 | TextBox 15/17 重叠 137pt | COM 实测 TB15.R=687, TB17.L=550 | 同上 |
| 3 | TextBox 17 格式错误 | 模板占位文字 vs 实际输出 | `filter="缺点"` 触发缺点摘要分支，应为 P1/P2 行动建议 |
| 4 | `style_anchor` 始终为空 | 代码 L527 硬编码 `""` | `_build_content()` 未从 spec 读取模板参考文字 |
| 5 | 章节标题（"优势"/"问题"）无颜色 | COM 实测全部黑色 | `_apply_keyword_color` 只染关键词，不染标题行 |

### 补充诊断数据

```
Slide 9 实测（COM）：
  TextBox 15: Left=38  Width=649  Right=687  ← 正常
  TextBox 17: Left=550 Width=480  Right=1030 ← 溢出 70pt

pipeline 03b-self_check_report.md：
  TextBox 15 内容超长 268>159 字 → clamp_text() 已在 zxh_ppt.py 中解决（当前 113 字）
  TextBox 17 内容超长 183>97  字 → 同上已解决（当前 63 字）
  SSIM=0.25（layout bug 是视觉差异主因）

02-shape_analysis_map.json 模板占位文字：
  TextBox 15 (133字): "优势\r包裹性表现较好...\r\r问题\r抓地不足（4/7）..."
  TextBox 17 (81字):  "修改建议\rP1：优化抓地\v核查橡胶硬度...\rP2：修正细节体验\v加长鞋带"
```

---

## 修复清单（5 处改动，全在 src/zxh_ppt.py）

---

### Fix 1+2：Layout 溢出 — `make_zxh_slide()` 加后处理

克隆模板 slide 后、写内容前，用代码直接矫正 TextBox 17 的位置和宽度。

**插入位置**：`time.sleep(1.0)` (line 719) 之后，`# === Per-shape content` 注释之前

```python
    # Fix 1+2: 矫正 TextBox 17 布局
    # 模板原始 L=550 W=480 → R=1030，溢出 slide 右边界(960pt)，且与 TB15(R=687)重叠 137pt
    # 矫正后: L=700 W=240 → R=940（不溢出，与 TB15 间距 13pt）
    try:
        _tb17 = new_slide.Shapes("TextBox 17")
        _tb17.Left  = 700
        _tb17.Width = 240
    except Exception:
        pass
```

---

### Fix 3：TextBox 17 Prompt → P1/P2 行动建议结构

模板占位文字是 "P1：优化抓地 / P1：优化后跟锁定 / P2：修正细节体验" 这样的行动建议格式。
当前 `filter="缺点"` 让 GPT 走自由摘要分支，格式完全不对。
需要新增一个 p1p2 格式分支。

#### 3a — `ZXH_SHAPES` 配置变更（line 67-72）

```python
# 改前
"params": {"source": "补充说明", "filter": "缺点"},

# 改后
"params": {"source": "补充说明", "filter": "修改建议", "format": "p1p2"},
```

#### 3b — `_build_rich_prompt()` 新增 `fmt` 参数和 p1p2 分支

**签名变更**（line 373-378）：

```python
def _build_rich_prompt(
    budget: dict, rows: List[List[Any]],
    focus: str = "", fmt: str = "",          # ← 新增 fmt
    content_source: str = "补充说明",
    style_anchor: str = "",
) -> str:
```

**新增分支**（插入 line 387 `extra = "每个分类不超过3行"` 之后、line 389 `if focus:` 之前）：

```python
    if fmt == "p1p2":
        task_line = (
            f"请从{n}名测试者的反馈中，提炼2-3条优先级最高的修改行动建议。\n"
            f"格式严格按照：\n"
            f"修改建议\n"
            f"P1：[问题简称]\n"
            f"[具体建议措施，1-2行]\n"
            f"P2：[问题简称]\n"
            f"[具体建议措施，1-2行]\n"
            f"每条建议聚焦一个可落地的改进点，P1 为最重要的优先改进项。\n"
            f"每段结论中，请将最核心的1-2个关键性能词用【】括起来（仅括词本身，不含标点），"
            f"这些关键词后续会自动高亮显示。"
        )
        format_note = "- 严格按照 P1/P2 优先级格式输出，P1 为最重要改进\n"
        extra = "每条建议不超过2行"
    elif focus:
        # ← 原有 focus 分支保持不变
```

注意：p1p2 分支将 `extra` 覆盖为 "每条建议不超过2行"，因此必须在 `extra = "每个分类不超过3行"` 之后插入。

#### 3c — `_build_content()` 提取 `fmt` 并透传

`gpt_prompted` 分支（line 520-532）：

```python
    if strategy == "gpt_prompted":
        focus = params.get("filter", "")
        fmt   = params.get("format", "")             # ← 新增
        src   = params.get("source", "补充说明")
        fallback_map = {
            "优点":    "样本反馈总体稳定，核心指标表现均衡。",
            "缺点":    "反馈集中，建议围绕关键指标继续优化。",
            "修改建议": "P1：优化核心指标\r根据样本反馈重点改进",   # ← 新增
        }
        fallback = fallback_map.get(focus, "样本反馈总体稳定，核心指标表现均衡。")
        prompt = _build_rich_prompt(budget, rows, focus=focus, fmt=fmt,
                                    content_source=src,
                                    style_anchor=style_anchor)  # ← Fix 4 一并改
        result = _call_gpt(prompt, fallback, gpt_enabled, model)
        return clamp_text(result, budget.get("max_chars", 200), budget.get("max_lines", 6))
```

---

### Fix 4：`style_anchor` 传入模板参考文字

`_build_rich_prompt` 已有 `style_anchor` 参数，但 `_build_content()` 始终传空字符串。
目标：将各 shape 的模板占位文字作为 GPT 的格式参考。

#### 4a — `ZXH_SHAPES` 加 `template_text` 字段（整体替换 line 60-73）

```python
ZXH_SHAPES = [
    {
        "name": "TextBox 15",
        "strategy": "gpt_prompted",
        "params": {"source": "补充说明", "filter": ""},
        "budget": {"max_chars": 159, "max_lines": 9},
        "template_text": (
            "优势\r包裹性表现较好，鞋脚一体性明显 \r支撑与刚性在线，实战稳定性较好 "
            "\r整体舒适度不错\r\r问题\r抓地不足（4/7）：木地板、急停、横移时更明显 "
            "\r后跟/脚踝问题（3/7）：掉跟、外踝卡脚 \r缓震偏硬/偏薄（2/7）\r鞋带偏短（2/7）"
        ),
    },
    {
        "name": "TextBox 17",
        "strategy": "gpt_prompted",
        "params": {"source": "补充说明", "filter": "修改建议", "format": "p1p2"},
        "budget": {"max_chars": 97, "max_lines": 8},
        "template_text": (
            "修改建议\rP1：优化抓地\v核查橡胶硬度，如果硬度正确考虑调软"
            "\r更换为普通橡胶或者止滑橡胶\rP1：优化后跟锁定\v调整后跟杯、领口泡棉"
            " \rP2：修正细节体验\v加长鞋带"
        ),
    },
]
```

#### 4b — `_build_content()` 签名加 `style_anchor` 参数

```python
def _build_content(spec: dict, rows: List[List[Any]],
                   gpt_enabled: bool, model: str,
                   style_anchor: str = "") -> str:    # ← 新增
```

#### 4c — `make_zxh_slide()` 调用时传入模板文字

```python
        content = _build_content(spec, rows, gpt_enabled, mc_model,
                                 style_anchor=spec.get("template_text", ""))
```

---

### Fix 5：章节标题行染色

`_apply_keyword_color` 只处理 【关键词】，不处理 "优势"/"问题" 这类章节标题行。
新增 `_color_section_headers()` 函数，复用 `tr.Find()` 模式（与 `_apply_keyword_color` 一致）。

#### 新函数（插入 `_apply_keyword_color` 之后，约 line 659）

```python
def _color_section_headers(shp) -> None:
    """Bold+color section header words: advantage markers → red, disadvantage → blue.

    Must be called AFTER _apply_keyword_color (which resets all text to black first).
    Uses tr.Find() for consistency with the rest of the codebase.
    """
    try:
        tf = _com_get(shp, "TextFrame", None)
        if tf is None:
            return
        tr = tf.TextRange
        for marker in _ADVANTAGE_MARKERS:
            start = 1
            while start <= tr.Length:
                found = tr.Find(marker, start)
                if found is None:
                    break
                found.Font.Bold  = True
                found.Font.Color = _RED
                start = found.Start + found.Length
        for marker in _DISADVANTAGE_MARKERS:
            start = 1
            while start <= tr.Length:
                found = tr.Find(marker, start)
                if found is None:
                    break
                found.Font.Bold  = True
                found.Font.Color = _BLUE
                start = found.Start + found.Length
    except Exception:
        pass  # coloring is cosmetic — never fail the build
```

#### 调用点（`make_zxh_slide()` 写入循环末尾）

```python
            if ok and strategy == "gpt_prompted":
                _apply_keyword_color(shp)
                _color_section_headers(shp)    # ← Fix 5
```

---

## 编辑顺序（从文件底部向上，避免行号偏移影响后续编辑）

| 步骤 | 改动位置 | 说明 |
|------|---------|------|
| 1 | line 659 后 | 新增 `_color_section_headers()` 函数 |
| 2 | `make_zxh_slide()` L719 后 | 插入 TextBox 17 layout 矫正代码 |
| 3 | `make_zxh_slide()` L746 | `_build_content` 调用加 `style_anchor=` |
| 4 | `make_zxh_slide()` L755 | 加 `_color_section_headers(shp)` 调用 |
| 5 | `_build_content()` L486 | 加 `style_anchor` 参数 + 提取 `fmt` + 改调用 |
| 6 | `_build_rich_prompt()` L373 | 加 `fmt` 参数 + 新增 p1p2 分支 |
| 7 | `ZXH_SHAPES` L60 | 整体替换：加 `template_text`，改 TB17 params |

---

## 验证

```bash
python -m py_compile src/zxh_ppt.py
```

运行 Main.py → 选 zxh 模板 → 检查：
- [ ] TextBox 17 不再溢出/与 TB15 不再重叠
- [ ] TextBox 17 内容为 "修改建议\rP1:...\rP2:..." 结构
- [ ] "优势" 标题红色加粗，"问题"/"修改建议" 标题蓝色加粗
- [ ] TextBox 15/17 中关键词有红/蓝高亮
