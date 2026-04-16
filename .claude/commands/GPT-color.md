# GPT 关键词自动染色

## 技术路径

GPT 标注 `【】` → `03b _apply_keyword_color()` 去括号 + 按段落上下文染色

## 分工

| 环节 | 负责 | 做什么 |
|------|------|--------|
| 标注 | GPT | 在 content 中用 `【关键词】` 标记核心词 |
| 染色 | 03b pipeline | 去 `【】`，按段落类型 bold + 染色 |

## 染色规则

| 段落类型 | 识别标记 | 颜色 | RGB 值 |
|---------|---------|------|--------|
| 优势/优点 | 优势、优点、亮点 | 纯红 | `255` (255,0,0) |
| 劣势/问题 | 问题、缺点、劣势、改进、修改建议 | 亮蓝 | `15773696` (0,176,240) |
| 其他 | — | 黑色 | `0` |

## 核心代码位置

- 染色函数: `pipeline/03b_build_ppt_com.py` → `_apply_keyword_color(shp)`
- GPT prompt 指令: `pipeline/prompt_templates/gpt_summary.md` 第9行
- 参考实现: `src/Function_030.py` → `smart_color_text()`

## 字体规范

统一使用 **微软雅黑**。`_write_text()` 写入文本后自动设置 `tr.Font.Name = "微软雅黑"`。自检会验证字体一致性，`_auto_fix` 可自动修复。

## 执行时机

`apply_shape()` 写入文本后自动调用，无需 `color_hint` 字段触发。
