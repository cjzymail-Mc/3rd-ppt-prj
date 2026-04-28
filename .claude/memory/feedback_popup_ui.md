---
name: feedback_popup_ui
description: tk 弹窗（_ask_with_countdown）样式约定——iOS systemGroupedBackground + 白色卡片 + Indigo 强调
type: feedback
---

`_ask_with_countdown`（`src/Function_030.py`）的视觉约定：**iOS 系统分组背景灰底 + 纯白卡片按钮 + 饱和 Indigo 强调描边**，字体统一 `Microsoft YaHei UI`。

| 角色 | 颜色 | 用途 |
|--|--|--|
| `_UI_BG` | `#F2F2F7` | 窗口底色（iOS systemGroupedBackground） |
| `_UI_BTN_BG` | `#FFFFFF` | 按钮静态：纯白卡片，与窗口底色拉开层次 |
| `_UI_BTN_BG_HOVER` | `#E3F2FD` | 按钮 hover：浅蓝填充 |
| `_UI_BTN_BORDER` | `#E5E5EA` | 非默认按钮静态描边（iOS separator gray） |
| `_UI_BTN_BORDER_HOVER` | `#64B5F6` | hover 描边：稍深蓝 |
| `_UI_ACCENT` | `#4A6CF7` | 默认按钮强调描边（饱和 Indigo，参考用户给的 chat dialog 截图） |
| `_UI_FG` | `#1C1C1E` | 主文本（iOS label primary） |
| `_UI_FG_MUTED` | `#8E8E93` | 描述文本（iOS secondaryLabel） |

**Why:**

1. **白底 + 白按钮没层次**（用户原话："浅灰色和弹窗窗体的颜色太接近了"）。iOS 设置页 / 通讯录的"浅灰底 + 白色卡片"反向分层是 tk 能复刻的最具高级感的方案，**不需要圆角和阴影**也能撑起视觉层次。
2. **不要尝试"蓝 header band + solid CTA"路线**（曾走过弯路）：reference 的截图（chat dialog with blue header）依赖圆角 + drop shadow，tk 没这两样东西，做出来反而粗糙。用户原话："这版效果不行。你直接换回上一版、蓝色参考新的截图即可"——保留 iOS 结构 + 只把红色 accent 换成截图里的蓝色才是正解。
3. **`Microsoft YaHei UI` 而非 `Arial`**：原版 `Arial 12pt` 渲染中文走系统 fallback，块状难看；YaHei UI 是 Win 自带，免依赖。
4. **`highlightthickness=2 + highlightbackground` 实现描边而不是 `relief`**：`relief='groove'` 这类 3D 边框在 flat 设计下违和；`highlightthickness` 是真正的"flat outline"。
5. **自然尺寸 `winfo_reqwidth/reqheight()`** 替代 hardcode `height = 80 + 60 * len(options)`：内容驱动尺寸，加 padding / 多行描述时不需要重新调常数。`width` 入参作下限。

**How to apply:**

- 新加弹窗优先复用 `_ask_with_countdown`（已支持 `options=[(label, value, description, fg)]` 多选项 + 倒计时 + 偏好记忆）
- 自定义新弹窗时引用上表常量（在 `Function_030.py` 弹窗常量块）
- Hover 同时改 `bg + highlightbackground + highlightcolor` 三项（filling + ring 一起变更，层次感更强）
- `cursor="hand2"` 是必须，强化"可点击"反馈
- 横向 padding ~28 / 纵向 padding ~24 / 选项行间 14 是已调试好的舒适值
- `pack(fill="x")` 让按钮填满 row，配合自然尺寸自动横向对齐
- **不要做 `overrideredirect(True)` 全自定义 chrome**——丢失 OS 拖拽 / 关闭按钮，得不偿失

**Code anchor**：`src/Function_030.py::_ask_with_countdown` + 顶部 `_UI_*` 常量块（弹窗视觉常量）。
