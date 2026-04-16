# Pipeline 移植到 src/zxh_ppt.py + 模板选择对话框

## Context

用户的 PPT pipeline（orchestrator + pipeline/*.py）已生成接近 90% 质量的 PPT。现在进入移植阶段：将 pipeline 核心能力整合为独立 Python 文件放入 `/src`，在 Main.py 中通过对话框路由到不同模板处理模块。

已有参考：`src/codex_ppt.py` 是早期移植样板（针对 yzr 模板），但缺少本轮 pipeline 的多项改进。

---

## 三件事

| # | 任务 | 文件 |
|---|------|------|
| 1 | 重命名 codex_ppt.py → yzr_ppt.py，更新引用 | `src/`, `Main.py` |
| 2 | 创建 zxh_ppt.py（移植 pipeline 核心能力） | `src/zxh_ppt.py`（新建） |
| 3 | Main.py 添加模板选择对话框 | `Main.py`, `src/Function_030.py` |

> **文件名注意**：用户写的 `zxh-ppt.py` 含连字符，Python 无法 import。实际命名为 `zxh_ppt.py`（下划线）。

---

## 任务 1：重命名 codex_ppt.py → yzr_ppt.py

**操作：**
1. `git mv src/codex_ppt.py src/yzr_ppt.py`
2. `Main.py:120` — `from src.codex_ppt import make_codex_slide` → `from src.yzr_ppt import make_codex_slide`

函数名 `make_codex_slide` 暂不改（避免破坏其他调用链）。
同时更新 yzr_ppt.py 中 `_TEMPLATE_SLIDE` 从 15 → 14（yzr 空白模板现在在第14页）。

---

## 任务 2：创建 src/zxh_ppt.py

### 策略：以 codex_ppt.py 为脚手架，升级 5 项能力

以 codex_ppt.py 结构为基础 copy，然后逐项升级：

| # | 升级项 | 来源 | 改动要点 |
|---|--------|------|---------|
| 1 | `_write_text()` | `pipeline/03b:108-110` | 加 `content.replace("\n", "\r")` + `tr.Font.Name = "微软雅黑"` |
| 2 | `_apply_keyword_color()` | `pipeline/03b:188-247` | 单色 → section-aware 双色（自动检测优势/劣势段落） |
| 3 | `clamp_text()` | `pipeline/ppt_pipeline_common:351-373` | 新增函数，在 GPT 返回后、写入前调用 |
| 4 | `_classify_columns()` + `_find_col()` | `pipeline/03a:305-354` | 替代硬编码 _SCORE_COLS/_TEXT_COLS |
| 5 | `_build_respondent_block()` | `pipeline/03a:357-398` | 用动态列匹配替代固定列名 |

### 模块结构

```python
"""zxh_ppt.py — 杨祖锐模板 PPT 生成（零 pipeline 依赖）"""

# 1. Constants
_RED, _BLUE, _BLACK = 255, 15773696, 0
_MODEL = "openai/gpt-5.4"
_TEMPLATE_SLIDE = 15  # zxh 空白模板在 Template 2.1.pptx 第15页

# 2. Section markers (for dual-color keyword coloring)
_ADVANTAGE_MARKERS = ["优势", "优点", "亮点", "表现较好", "表现突出"]
_DISADVANTAGE_MARKERS = ["问题", "缺点", "劣势", "不足", "改进", "修改建议", "待优化"]

# 3. Shape specs (hardcoded, from pipeline step1 分析结果)
ZXH_SHAPES = [
    {"name": "TextBox 15", "strategy": "gpt_prompted",
     "params": {"source": "补充说明", "filter": ""},
     "budget": {"max_chars": 159, "max_lines": 9}},
    {"name": "TextBox 17", "strategy": "gpt_prompted",
     "params": {"source": "补充说明", "filter": "缺点"},
     "budget": {"max_chars": 97, "max_lines": 8}},
]

# 4. Utility functions (from codex_ppt.py, 不变)
# 5. Dynamic column helpers (from pipeline/03a, 新增)
# 6. Data extraction (from codex_ppt.py, 不变)
# 7. GPT prompt builder (升级为动态列匹配版)
# 8. clamp_text() (from pipeline/ppt_pipeline_common, 新增)
# 9. xlwings helpers (from codex_ppt.py, 不变)
# 10. Content builder (from codex_ppt.py, 集成 clamp_text)
# 11. COM writers (升级版 _write_text + _apply_keyword_color)
# 12. Public API: make_zxh_slide()
```

### 公开 API 签名

```python
def make_zxh_slide(mc_sht, mc_ppt, mc_slide, sample_name: str,
                   mc_gpt: str = "n", mc_model: str = _MODEL):
```

与 `make_codex_slide()` 签名一致，Main.py 可无缝路由。

### 关键代码片段

**`_write_text()` 升级版：**
```python
def _write_text(shp, content: str) -> bool:
    try:
        tf = shp.TextFrame
        tr = tf.TextRange
        tr.Text = content.replace("\n", "\r")  # PPT 段落分隔符
        tr.Font.Name = "微软雅黑"
        return True
    except Exception:
        return False
```

**`_apply_keyword_color()` section-aware 版：**
- 不接受 color_rgb 参数，自动分析段落所属 section
- 遍历每个段落，跟踪 current_section（优势/劣势/其它）
- 【keyword】在优势段落 → 红色加粗，劣势段落 → 蓝色加粗，其它 → 黑色
- 源码来自 `pipeline/03b_build_ppt_com.py:188-247`

**`clamp_text()` 安全网：**
- 行数硬限 + 字数在句子边界截断
- 在 `_build_content()` 中 GPT 返回后立即调用
- 源码来自 `pipeline/ppt_pipeline_common.py:351-373`

---

## 任务 3：Main.py 模板选择对话框

### 新增函数：`ask_template_choice()`

**位置**：`src/Function_030.py`（紧接 `ask_gpt_model()` 之后，约 L717）

```python
def ask_template_choice():
    """弹窗选择问卷模板，返回 'yzr' 或 'zxh'。"""
    choice = None
    def select(v):
        nonlocal choice
        choice = v
        win.quit()
        win.destroy()

    win = tk.Tk()
    win.title("选择问卷模板")
    win.resizable(False, False)
    win.protocol("WM_DELETE_WINDOW", lambda: select("yzr"))
    tk.Label(win, text="请选择问卷模板:", font=("Arial", 12)).pack(pady=10)
    tk.Button(win, text="yzr模板", width=20, command=lambda: select("yzr")).pack(pady=5)
    tk.Button(win, text="zxh模板", width=20, command=lambda: select("zxh")).pack(pady=5)
    force_window_front(win)
    center_window(win, 300, 160)
    win.mainloop()
    return choice or "yzr"
```

沿用 `ask_gpt_model()` 的 tkinter 模式：nonlocal 捕获 + force_window_front + center_window。

### Main.py 改动

**imports (L120)：**
```python
from src.yzr_ppt import make_codex_slide    # 重命名后
from src.zxh_ppt import make_zxh_slide      # 新模块
```

**call site (L800-809)：**
```python
# 【5.6】Codex 分析页：高保真评测汇总
if mc_sht is not None:
    template_choice = ask_template_choice()
    if template_choice == "zxh":
        mc_slide = make_zxh_slide(
            mc_sht, mc_ppt, mc_slide, sample_name,
            mc_gpt=mc_gpt, mc_model=mc_model,
        )
    else:
        mc_slide = make_codex_slide(
            mc_sht, mc_ppt, mc_slide, sample_name,
            mc_gpt=mc_gpt, mc_model=mc_model,
        )
```

---

## 任务 4：同步升级 yzr_ppt.py

> 在任务 2（zxh_ppt.py 移植）完成后执行。将同样的 4 项能力升级同步到 yzr_ppt.py。

### 改动清单

| # | 升级项 | 当前代码 | 改为 |
|---|--------|---------|------|
| 1 | `_write_text()` | `tr.Text = content`（L469） | 加 `content.replace("\n", "\r")` + `tr.Font.Name = "微软雅黑"` |
| 2 | `_apply_keyword_color()` | 单色，接受 `color_rgb` 参数（L524） | section-aware 双色，从 zxh_ppt.py 复制同名函数 |
| 3 | `clamp_text()` | 不存在 | 新增函数，在 `_build_content()` 的 gpt_prompted 分支中调用 |
| 4 | 列匹配 | 硬编码 `_SCORE_COLS` / `_TEXT_COLS` | 替换为 `_classify_columns()` + `_find_col()`，从 zxh_ppt.py 复制 |

### 注意事项

- `CODEX_SHAPES` 和 `_TEMPLATE_SLIDE = 14` 不动（yzr 模板专属）
- 函数签名 `make_codex_slide()` 不改名
- 改动后直接从 zxh_ppt.py 复制对应函数即可，两个文件的辅助函数保持一致

---

## 执行顺序

1. 任务 1（重命名）— 最先做，风险最低
2. 任务 3（对话框）— 先搭好路由框架，zxh 分支可以先 placeholder
3. 任务 2（zxh_ppt.py）— 主体工作，逐项从 pipeline 移植
4. 任务 4（yzr_ppt.py 同步升级）— 从 zxh_ppt.py 复制 4 项改进

---

## 验证方式

1. `python -m py_compile src/yzr_ppt.py` — 重命名 + 升级后语法检查
2. `python -m py_compile src/zxh_ppt.py` — 新文件语法检查
3. `python -m py_compile Main.py` — import 和调用正确性
4. 运行 Main.py → 到问卷阶段 → 弹出模板选择对话框 → 选 yzr → 验证染色/字体/截断生效
5. 选 zxh → 调用 make_zxh_slide() → 生成 PPT 页面 → 检查字体/染色/截断是否生效
