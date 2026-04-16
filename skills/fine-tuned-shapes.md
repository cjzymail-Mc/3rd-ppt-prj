---
name: Shape 位置微调工作流
description: 用户指定 shape 名称后，在 xxx_ppt.py 的 make_xxx_slide() 中插入 Left/Top/Width/Height 微调代码块
type: feedback
---

用户经常需要对生成的 PPT 中 1-3 个 shape 做位置/尺寸微调（Left/Top/Width/Height）。

**Why:** Clone Slide 继承的模板位置不一定适合最终输出，需要在代码中硬编码修正值。

## 操作流程

1. **用户说**："帮我微调 Rectangle 68 的位置" 或 "把 XXX shape 往左移一点"
2. **定位文件**：根据 shape 名称在对应的 `XXX_SHAPES` 常量中确认属于哪个 `xxx_ppt.py`
3. **获取基准值**：从**标准模板** `src/Template 2.1.pptx` 中用 COM 读取该 shape 的 Left/Top/Width/Height（不要从已生成的输出文件读取）
4. **插入位置**：在 `make_xxx_slide()` 函数中，Clone Slide 之后、遍历 shapes 之前（`time.sleep(1.0)` 和 `for spec in XXX_SHAPES:` 之间）
5. **代码模式**（四参数完整，标记 `#fine_tuned`）：
   ```python
   # Shape 位置微调 #fine_tuned
   try:
       _shp = new_slide.Shapes("Rectangle 68")
       _shp.Left   = 20.20   # 从标准模板读取，或用户指定值
       _shp.Top    = 260.65
       _shp.Width  = 416.88
       _shp.Height = 247.19
   except Exception:
       pass
   ```
6. **注意**：如果用户已经微调过某些值，只更新用户未指定的参数，不要用模板原始值覆盖用户微调过的值

## Shape 常量命名规范

| 模板文件 | 常量名 | 模板页 |
|----------|--------|--------|
| `src/yzr_ppt.py` | `YZR_SHAPES` | Slide 15 |
| `src/zxh_ppt.py` | `ZXH_SHAPES` | Slide 17 |
| 未来新增 | `{NAME}_SHAPES` | 按模板定 |

## 当前已微调的 shapes

### yzr_ppt.py
| Shape | Left | Top | Width | Height |
|-------|------|-----|-------|--------|
| Rectangle 68 | 20.20 | 260.65 | 416.88 | 247.19 |
| Rectangle 77 | 450.79 | 260.65 | 274.36 | 225.38 |

### zxh_ppt.py
| Shape | Left | Top | Width | Height | 备注 |
|-------|------|-----|-------|--------|------|
| TextBox 15 | 37.75 | 128.25 | 648.99 | 330.55 | 模板原始值 |
| TextBox 17 | **650** | 146.63 | **280** | 265.11 | Left/Width 为用户微调值 |

## 单独调试入口

每个 `xxx_ppt.py` 均有 `if __name__ == "__main__"` 调试入口：
- 自动打开 `Template 2.1.pptx`（与 Main.py 同方式），不自动保存
- 连接已打开的 Excel（xlwings），自动找"问卷" sheet
- 用法：`python src/yzr_ppt.py` 或 `python src/zxh_ppt.py`
- 关键：顶部 `if __name__ == "__main__": sys.path.insert(...)` 解决直接运行时的导入问题
