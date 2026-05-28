#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_ppt_shared.py — yzr / zxh 共享工具（数据计算 + 通用 COM 写入）.

共享范围：
  - 纯数据/纯计算函数（_extract_score_means、_classify_columns 等）
  - 通用 COM 写入函数（_com_get、_write_text、_write_chart）
    这些函数在所有模板中行为完全一致，集中维护避免重复修 bug。

per-template 保留在各自文件：
  - {NAME}_SHAPES 配置、_TEMPLATE_SLIDE
  - _build_rich_prompt（prompt 结构因模板而异）
  - _apply_keyword_color / _color_section_headers（染色规则可能不同）
  - make_{name}_slide()（含 #fine_tuned shape 微调块）

详见 skills/fine-tuned-shapes.md。
"""

from __future__ import annotations

import re
import time
from typing import Any, Dict, List, Optional, Tuple


# ===========================================================================
# 常量（yzr / zxh 两套值完全相同）
# ===========================================================================
_RED   = 255        # RGB(255, 0, 0)   — advantage keywords
_BLUE  = 15773696   # RGB(0, 176, 240) — disadvantage keywords
_BLACK = 0          # RGB(0, 0, 0)     — default text

_ADVANTAGE_MARKERS = ["优势", "优点", "亮点", "表现较好", "表现突出"]
_DISADVANTAGE_MARKERS = ["问题", "缺点", "劣势", "不足", "改进", "修改建议", "待优化"]


# ===========================================================================
# 内部微型 helper（本模块内部用；yzr/zxh 仍各自保留自己的 _safe_text/_numeric
# 以方便调试，故此处复制无害）
# ===========================================================================
def _safe_text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _numeric(v: Any):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


def _to_rows(value: Any) -> List[List[Any]]:
    """Convert xlwings used_range.value to List[List[Any]]."""
    if value is None:
        return []
    if isinstance(value, tuple):
        rows = []
        for row in value:
            if isinstance(row, tuple):
                rows.append(list(row))
            else:
                rows.append([row])
        return rows
    if isinstance(value, list):
        if value and isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


# ===========================================================================
# Excel 数据提取（纯读；零 COM 副作用）
# ===========================================================================
_NAME_KEYWORDS = ["姓名", "name"]
_WEIGHT_KEYWORDS = ["体重", "weight"]
_SKIP_KEYWORDS = ["鞋款", "轮次", "身高", "累计", "场次"]


def _find_col(headers: List[str], keywords: List[str]) -> Optional[str]:
    """Find first header containing any keyword (case-insensitive)."""
    for h in headers:
        hl = h.lower()
        for kw in keywords:
            if kw.lower() in hl:
                return h
    return None


def _classify_columns(headers: List[str], rows: List[List[Any]],
                      extra_skip_keywords: Optional[List[str]] = None,
                      ) -> Tuple[List[str], List[str]]:
    """Dynamically classify columns into score (numeric) and text (feedback).

    Score columns: >70% of data rows have numeric values in 0-10 range.
    Text columns:  >30% of data rows have string values longer than 5 chars.

    Args:
        extra_skip_keywords: per-template additional skip keywords; merged into
            the default _SKIP_KEYWORDS set. Used by apparel to filter new
            metadata cols (data_id / 提交时间 / 三围 / 温度...) that would
            otherwise be misclassified as text feedback. See fix3.
    """
    name_col = _find_col(headers, _NAME_KEYWORDS)
    weight_col = _find_col(headers, _WEIGHT_KEYWORDS)
    extra = list(extra_skip_keywords or [])
    skip_cols = {name_col, weight_col} | {
        h for h in headers
        if any(kw in h for kw in (_SKIP_KEYWORDS + extra))
    }

    score_cols, text_cols = [], []
    n_data = len(rows) - 1
    if n_data < 1:
        return score_cols, text_cols

    for ci, h in enumerate(headers):
        if h in skip_cols:
            continue
        nums, texts = 0, 0
        for row in rows[1:]:
            val = row[ci] if ci < len(row) else None
            if val is None:
                continue
            try:
                v = float(val)
                if 0 <= v <= 10:
                    nums += 1
            except (ValueError, TypeError):
                if len(_safe_text(val)) > 5:
                    texts += 1
        if nums / n_data > 0.7:
            score_cols.append(h)
        elif texts / n_data > 0.3:
            text_cols.append(h)

    return score_cols, text_cols


def _col_values(rows: List[List[Any]], *keywords: str) -> List[Any]:
    """Return all non-None values from the first column whose header
    contains any of the given keywords."""
    if not rows:
        return []
    headers = [_safe_text(h) for h in rows[0]]
    for kw in keywords:
        for idx, h in enumerate(headers):
            if kw in h:
                return [
                    row[idx]
                    for row in rows[1:]
                    if idx < len(row) and row[idx] is not None
                    and _safe_text(row[idx]) != ""
                ]
    return []


def _extract_score_means(rows: List[List[Any]]) -> List[Tuple[str, float]]:
    """Extract per-column score means for bar charts."""
    if not rows or len(rows) < 2:
        return []

    headers = [_safe_text(h) for h in rows[0]]
    data = rows[1:]
    ncol = max(len(r) for r in rows)

    score_like = []
    backup_numeric = []
    score_keys = [
        "评分", "分数", "打分", "满意", "体验", "表现",
        "减震", "回弹", "稳定", "抓地", "舒适", "透气", "支撑",
    ]
    reject_keys = ["姓名", "昵称", "电话", "联系方式", "地址", "微信", "备注", "日期", "时间",
                   "第几轮", "轮次", "轮反馈", "这是第几"]  # 排除"第几轮反馈"等非评分列

    for c in range(ncol):
        header = headers[c] if c < len(headers) else f"指标{c+1}"
        if any(k in header for k in reject_keys):
            continue

        vals = []
        for r in data:
            if c >= len(r):
                continue
            n = _numeric(r[c])
            if n is not None:
                vals.append(float(n))

        if not vals:
            continue

        mean_val = sum(vals) / len(vals)
        in_score_range = all(0 <= v <= 20 for v in vals)
        if any(k in header for k in score_keys) and in_score_range:
            score_like.append((header, round(mean_val, 3)))
        elif in_score_range:
            backup_numeric.append((header, round(mean_val, 3)))

    return score_like + backup_numeric


def _xlwings_to_rows(mc_sht) -> List[List[Any]]:
    """xlwings Sheet → List[List[Any]], using CurrentRegion of the data anchor."""
    try:
        import importlib
        fn030 = importlib.import_module("src.Function_030")
        mc_cell0 = fn030.get_range(mc_sht)
        if mc_cell0 is not None:
            raw = mc_cell0.api.CurrentRegion.Value
            if raw is not None:
                return _to_rows(raw)
    except Exception:
        pass
    return _to_rows(mc_sht.used_range.value)


# ===========================================================================
# 评分/统计（纯计算）
# ===========================================================================
def _score_10pt(rows: List[List[Any]]):
    """Calculate overall mean score, auto-detects 5-scale vs 10-scale,
    returns a 10-point float or None."""
    means = _extract_score_means(rows)
    if not means:
        return None
    overall = sum(v for _, v in means) / len(means)
    max_mean = max(v for _, v in means)
    if max_mean <= 5.5:
        return round(overall * 2, 2)
    return round(overall, 2)


def _score_to_grade(score_10: float) -> str:
    """Convert a 10-point score to a letter grade."""
    s = score_10 * 10
    if s >= 95: return "S+"
    if s >= 90: return "S-"
    if s >= 85: return "A+"
    if s >= 80: return "A-"
    if s >= 75: return "B+"
    if s >= 70: return "B-"
    if s >= 65: return "C+"
    return "C-"


def _score_to_grade_letter(score_10: float) -> str:
    """Return only the letter portion (S/A/B/C) of the grade."""
    return _score_to_grade(score_10)[0]


def _score_to_grade_modifier(score_10: float) -> str:
    """Return only the +/- modifier of the grade."""
    return _score_to_grade(score_10)[1]


def _sample_stat_text(rows: List[List[Any]]) -> str:
    """Build sample stat text: trial count / avg weight / court position."""
    count = max(0, len(rows) - 1)

    weights = _col_values(rows, "体重", "Weight", "重量")
    valid_w = [float(w) for w in weights if _numeric(w) is not None]
    avg_w = round(sum(valid_w) / len(valid_w), 1) if valid_w else None

    positions = _col_values(rows, "球场定位", "打法", "定位", "位置")
    pos_clean: List[str] = []
    seen: set = set()
    for p in positions:
        s = _safe_text(p)
        if s and s not in seen:
            seen.add(s)
            pos_clean.append(s)

    lines = [f"试穿人数：{count}人"]
    if avg_w is not None:
        lines.append(f"测试者平均体重：{avg_w}KG")
    if pos_clean:
        lines.append(f"测试者球场定位：{'、'.join(pos_clean)}")
    return "\n".join(lines)


# ===========================================================================
# 通用 COM 写入工具（PPT shape 写入，所有模板共用）
# ===========================================================================

def _com_get(obj, attr: str, default=None):
    """Safe getattr for COM objects (getattr raises on COM objects)."""
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def _write_text(shp, content: str) -> bool:
    """Write text to PPT shape with \\n→\\r conversion and 微软雅黑 font."""
    if not bool(_com_get(shp, "HasTextFrame", 0)):
        return False
    tf = _com_get(shp, "TextFrame", None)
    tr = _com_get(tf, "TextRange", None) if tf is not None else None
    if tr is None:
        return False
    try:
        tf.AutoSize = 0  # ppAutoSizeNone — preserve template geometry
    except Exception:
        pass
    try:
        # PPT COM 使用 \r 作为段落分隔符；\n 会被忽略
        tr.Text = content.replace("\n", "\r")
        tr.Font.Name = "微软雅黑"
        return True
    except Exception:
        return False


def _write_chart(shp, content: str) -> bool:
    """Write chart data via SeriesCollection.

    ⚠️ **适用场景警告（fix4）**：
      此函数仅限"单机自用、模板 + 数据同机"场景。
      分发场景（模板 / 代码发给他人，数据他人填）下 chart 内部状态必然漂移，
      此函数不可靠。请改用 make_chart_for_yzr（从零制表 + OLE 粘贴）。
      路线决策与论据见 [feature03-transplant]/fix4（图表路线切换）.md。
      当前仍保留此函数，是为 zxh_ppt 在单机场景下的兼容性，以及未来 Pipeline
      新模板分析场景的使用。

    策略分叉（关键决策）：
      - IsLinked=False（inline-cache chart，无 embedded workbook）：
        直接 series.Values = tuple，**不调 Activate / BreakLink**。
        原因：无后端 workbook 的 chart，Activate() 会即兴拉起临时 workbook，
        部分机器（用户 + 同事）上该过程不稳定，制造 "null link 幽灵"，
        后续 SeriesCollection 写入静默失效 → bars 消失。
        inline cache 可直接接收 COM 写入，跳过 Activate 反而稳定。

      - IsLinked=True（linked to external workbook）：
        走原有流程 BreakLink → Activate(×3) → 再 BreakLink → write。
    """
    chart = _com_get(shp, "Chart", None)
    if chart is None:
        print(f"  [图表] 未找到 Chart 对象，跳过")
        return False

    lines = [x.strip() for x in (content or "").splitlines() if x.strip()]
    labels, values = [], []
    for line in lines[:10]:
        if ":" in line:
            k, v = line.rsplit(":", 1)
            labels.append(k.strip())
            try:
                values.append(float(v.strip()))
            except Exception:
                values.append(0.0)

    if not labels:
        print(f"  [图表] content 解析出 0 条数据，跳过写入（content={repr(content[:60])}）")
        return False

    print(f"  [图表] 准备写入 {len(labels)} 个指标: {list(zip(labels, values))}")

    # 探测是否真的有外部链接
    is_linked = False
    try:
        is_linked = bool(chart.ChartData.IsLinked)
        print(f"  [图表] ChartData.IsLinked = {is_linked}")
    except Exception as _e:
        print(f"  [图表] IsLinked 探测异常（按 False 处理）: {_e}")
        is_linked = False

    if is_linked:
        # 链接型 chart：保留原流程（BreakLink → Activate ×3 → 再 BreakLink）
        try:
            print(f"  [图表] 检测到外部链接，提前 BreakLink...")
            chart.ChartData.BreakLink()
            time.sleep(0.8)
        except Exception as _e:
            print(f"  [图表] 前置 BreakLink 异常: {_e}")

        for _attempt in range(1, 4):
            try:
                chart.ChartData.Activate()
                time.sleep(0.8)
                print(f"  [图表] Activate 成功（第{_attempt}次）")
                break
            except Exception as _e:
                print(f"  [图表] Activate 第{_attempt}次失败: {_e}")
                time.sleep(0.4 * _attempt)

        try:
            chart.ChartData.BreakLink()
            time.sleep(0.3)
        except Exception:
            pass
    else:
        # inline-cache chart：跳过 Activate/BreakLink，直接写
        # 这是修复 "bars 消失" 的关键——Activate 是幽灵 null link 的触发源
        print(f"  [图表] inline-cache chart，跳过 Activate/BreakLink，直接写入")

    # 写入 + 回读验证
    try:
        series = chart.SeriesCollection(1)
        series.Values = tuple(values)
        series.XValues = tuple(labels)
        time.sleep(0.3)

        actual_vals = list(series.Values)
        if actual_vals and abs(float(actual_vals[0]) - values[0]) < 0.05:
            print(f"  [图表] 写入并验证成功（首值 期望={values[0]:.2f} 实际={float(actual_vals[0]):.2f}）")
            return True
        else:
            print(f"  [图表] 写入后验证失败！期望首值={values[0]:.2f}，实际={actual_vals[0] if actual_vals else 'N/A'}")
            print(f"  [图表] 如持续失败，考虑用 AddChart2 重建 chart（方案 D）")
            return False
    except Exception as _e:
        print(f"  [图表] 写入失败: {_e}")
        return False


# ===========================================================================
# 文本处理（纯字符串）
# ===========================================================================
def clamp_text(text: str, max_chars: int, max_lines: int) -> str:
    """Clamp text to fit PPT shape: enforce both line count and character count.

    Line clamp: hard cut at max_lines.
    Char clamp: hard cut at sentence boundary when exceeding max_chars.

    Pre-clean: strip blank/whitespace-only lines —— GPT 偶尔会在段落间吐空行，
    一旦保留空行，PPT TextFrame 行数翻倍、超出 shape Height。
    （旧 bug：用户多次反馈"每行之间多一行空行"，根因就在这里——
     原版 splitlines 直接 join 不剔空行）
    """
    t = _safe_text(text)
    # 去除空行 / 纯空白行 + 去掉每行前后空白
    if t:
        cleaned = [ln.strip() for ln in t.splitlines()]
        cleaned = [ln for ln in cleaned if ln]
        t = "\n".join(cleaned)
    if max_lines > 0:
        lines = t.splitlines() or [t]
        t = "\n".join(lines[:max_lines])
    if max_chars > 0 and len(t) > max_chars:
        truncated = t[:max_chars]
        for sep in ['。', '！', '？', '\n', '；', '，', '、', ' ']:
            idx = truncated.rfind(sep)
            if idx > max_chars * 0.5:
                t = truncated[:idx + 1].rstrip()
                break
        else:
            t = truncated
    return t


# ===========================================================================
# 关键词染色 + 段头 bullet 去除（plan4 — Main.py 6.3 也复用，从 yzr_ppt 提升）
# ===========================================================================

def _apply_keyword_color(shp) -> None:
    """Remove 【】 brackets, then bold+color keywords by section context.

    Rules:
      - Keywords in advantage sections (优势/优点/...) → red + bold
      - Keywords in disadvantage sections (问题/缺点/修改建议/...) → blue + bold
      - All other text → black
    """
    try:
        tf = _com_get(shp, "TextFrame", None)
        if tf is None:
            return
        tr = tf.TextRange
        full_text = tr.Text

        keywords = list(dict.fromkeys(re.findall(r'【([^】]+)】', full_text)))
        if not keywords:
            return

        # Build keyword→color map based on section context
        kw_color: Dict[str, int] = {}
        current_section = "neutral"
        for line in full_text.split('\r'):
            line_stripped = line.strip()
            if any(m in line_stripped for m in _ADVANTAGE_MARKERS):
                current_section = "advantage"
            elif any(m in line_stripped for m in _DISADVANTAGE_MARKERS):
                current_section = "disadvantage"
            for kw in re.findall(r'【([^】]+)】', line):
                if current_section == "advantage":
                    kw_color[kw] = _RED
                elif current_section == "disadvantage":
                    kw_color[kw] = _BLUE

        # Remove all 【】 brackets
        tr.Text = re.sub(r'[【】]', '', full_text)

        # Reset entire shape to black first (clear any inherited colors)
        tr.Font.Color = _BLACK

        # Bold + color each keyword
        for kw, color in kw_color.items():
            start = 1
            while start <= tr.Length:
                found = tr.Find(kw, start)
                if found is None:
                    break
                found.Font.Bold = True
                found.Font.Color = color
                start = found.Start + found.Length
    except Exception:
        pass  # coloring is cosmetic — never fail the build


def _apply_conclusion_color(shp) -> None:
    """6.3 结论页专用染色（todays-task：bracket-typed keyword scheme）.

    GPT 自标记关键词，按括号类型决定染色：
      - <keyword>  → 红色 + 加粗   （优点段关键词）
      - [keyword]  → 蓝色 + 加粗   （缺点段关键词）
      - (keyword)  → 仅加粗、不染色 （修改建议段关键词）

    染色完成后：
      - 剥离全部 ASCII <>[]() 标记括号；
      - 保留中文 【】 段头括号（由 _strip_bullet_on_section_headers 单独处理）；
      - 其余文本 reset 为黑色非粗。
    """
    try:
        tf = _com_get(shp, "TextFrame", None)
        if tf is None:
            return
        tr = tf.TextRange
        full_text = tr.Text or ""

        red_kws = list(dict.fromkeys(re.findall(r'<([^<>\r\n]+)>', full_text)))
        blue_kws = list(dict.fromkeys(re.findall(r'\[([^\[\]\r\n]+)\]', full_text)))
        bold_kws = list(dict.fromkeys(re.findall(r'\(([^()\r\n]+)\)', full_text)))

        if not (red_kws or blue_kws or bold_kws):
            return

        # 仅剥离 ASCII 标记括号；中文【】保持不变
        cleaned = re.sub(r'[<>\[\]()]', '', full_text)
        tr.Text = cleaned

        # 整段先 reset 为黑色 + 非粗（清除继承样式）
        try:
            tr.Font.Color = _BLACK
        except Exception:
            pass
        try:
            tr.Font.Bold = False
        except Exception:
            pass

        def _mark(keywords, color, bold):
            for kw in keywords:
                if not kw:
                    continue
                start = 1
                while start <= tr.Length:
                    found = tr.Find(kw, start)
                    if found is None:
                        break
                    try:
                        if bold:
                            found.Font.Bold = True
                        if color is not None:
                            found.Font.Color = color
                    except Exception:
                        pass
                    start = found.Start + found.Length

        # 顺序：红 → 蓝 → 仅粗体（三类互不重叠时顺序无影响）
        _mark(red_kws, _RED, bold=True)
        _mark(blue_kws, _BLUE, bold=True)
        _mark(bold_kws, None, bold=True)
    except Exception:
        pass  # coloring is cosmetic — never fail the build


def _strip_bullet_on_section_headers(tr) -> None:
    """段头行（如 【优点】/【缺点】/【修改建议】）去掉 ■ bullet。

    Result_Bullet 默认每段都加 ■，但段头本身加 ■ 视觉冗余。
    识别规则：行首是【，含】，整行近似只是段标（≤10 字符）。
    """
    try:
        paragraphs = tr.Paragraphs()
        n = int(paragraphs.Count)
        for i in range(1, n + 1):
            p = tr.Paragraphs(i, 1)
            line = (p.Text or "").strip()
            if line.startswith("【") and "】" in line and len(line) <= 10:
                try:
                    p.ParagraphFormat.Bullet.Visible = 0
                except Exception:
                    pass
    except Exception:
        pass


# ===========================================================================
# fix4 — 分发场景 chart 从零制表（xlwings 新建 → OLE 粘贴）
# 路线决策见 [feature03-transplant]/fix4（图表路线切换）.md
# 范式参考 src/Function_030.py::make_chart_for_questionnaire（已在产多年）
# ===========================================================================

def _prepare_yzr_chart_data(mc_sht, content: str):
    """Parse mean_extraction content and write to Excel as 2-column table.

    content format (每行一条):
        指标名:均值
        抓地性:8.33
        ...

    在 mc_sht 的安全区域（远离 make_chart_for_questionnaire 的临时区）写入：
        | 指标   | 均值 |
        | 抓地性 | 8.33 |
        | ...    | ...  |

    Returns the xlwings Range pointing to the table's header cell
    (for make_chart_for_yzr 的 data anchor).
    """
    import importlib

    # 获取数据原点（与 make_chart_for_questionnaire 共享同一判定逻辑）
    try:
        fn030 = importlib.import_module("src.Function_030")
        origin = fn030.get_range(mc_sht)
    except Exception:
        try:
            fn030 = importlib.import_module("Function_030")
            origin = fn030.get_range(mc_sht)
        except Exception:
            origin = mc_sht.range("A1")

    # 原数据行数
    try:
        rows_count = origin.api.CurrentRegion.Rows.Count
    except Exception:
        rows_count = 10

    # 安全偏移：questionnaire 用 i+8，yzr 用 i+40，远离以避免冲突
    anchor = origin.offset(row_offset=rows_count + 40, column_offset=0)

    # 解析 content
    parsed: List[Tuple[str, float]] = []
    for line in (content or "").splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        k, v = line.rsplit(":", 1)
        try:
            parsed.append((k.strip(), float(v.strip())))
        except Exception:
            parsed.append((k.strip(), 0.0))

    if not parsed:
        parsed = [("占位", 0.0)]

    table = [("指标", "均值")] + parsed
    anchor.value = table
    print(f"  [yzr-chart] 临时数据已写入：anchor=({anchor.row},{anchor.column})，{len(parsed)} 个指标")
    return anchor


def make_chart_for_yzr(mc_cell, mc_slide, Left, Top, Width, Height):
    """为 yzr 模板构建 3D 条形图（ChartType=60），OLE 粘贴到 PPT。

    参数：
      mc_cell: xlwings Range —— 2 列表格锚点（由 _prepare_yzr_chart_data 返回）
      mc_slide: PPT slide (win32com)
      Left, Top, Width, Height: 粘贴到 PPT 后的位置/尺寸（points）

    返回：xlwings chart 对象（外层决定是否 delete 以清理 Excel 端）。

    与 make_chart_for_questionnaire 的差异：
      - ChartType = 60（xl3DBarClustered；questionnaire 用 bar_clustered 2D）
      - 固定量程 0~10（yzr 问卷一律 10 分制）
      - 数据形状：N 指标 × 1 均值列

    硬规则依赖：
      - `CutCopyMode = False`：Paste 后必须执行，断 OLE 热链接（规则 #1）
      - 不使用 xlPicture 常量（这里是 OLE 粘贴，不是图片粘贴）
    """
    import random
    import xlwings

    print("[yzr-chart] 开始 xlwings 建 3D 条形图 → OLE 粘贴")

    mc_sht = mc_cell.sheet
    try:
        mc_sht.select()
    except Exception:
        pass
    mc_cell.select()

    i0 = mc_cell.api.CurrentRegion.Row
    i = mc_cell.api.CurrentRegion.Rows.Count
    j0 = mc_cell.api.CurrentRegion.Column
    j = mc_cell.api.CurrentRegion.Columns.Count

    # Excel 里 chart 的位置（临时——后面会 OLE 粘贴到 PPT）
    chart_left = mc_sht.cells(i0 + i - 2, j0 + 3).left
    chart_top = mc_sht.cells(i0 + i - 2, j0 + 3).top

    mc_chart1 = mc_sht.charts.add(chart_left, chart_top, width=Width, height=Height)

    # 3D 条形图：ChartType = 60 (xl3DBarClustered)
    try:
        mc_chart1.api[1].ChartType = 60
        print("  [yzr-chart] ChartType = 60 (3D bar clustered)")
    except Exception as _e:
        print(f"  [yzr-chart] 设置 ChartType=60 失败（{_e}），回退 2D bar")
        mc_chart1.chart_type = "bar_clustered"

    mc_chart1.set_source_data(
        mc_sht.range((i0, j0), (i0 + i - 1, j0 + j - 1))
    )

    # 三维视图/旋转（与 PPT"三维旋转"面板对齐；2026-04-24 用户实测值）
    # 映射：X 旋转 → Elevation，Y 旋转 → Rotation
    try:
        _ch = mc_chart1.api[1]
        _ch.RightAngleAxes = True    # 直角坐标轴 ☑
        _ch.AutoScaling = True       # 自动缩放 ☑
        _ch.Elevation = 20           # X 旋转 20°
        _ch.Rotation = 15            # Y 旋转 15°
        _ch.Perspective = 0          # 透视 0°（RightAngleAxes=True 时通常忽略，为完整性保留）
        _ch.DepthPercent = 100       # 深度 100%
        _ch.HeightPercent = 100      # 高度 100%
        print("  [yzr-chart] 3D 视图：Elevation=20, Rotation=15, RightAngleAxes=True, Depth=100, Height=100")
    except Exception as _e:
        print(f"  [yzr-chart] 设置 3D 视图失败（{_e}），使用 xlwings 默认视角")

    # 隐藏图例
    mc_chart1.api[1].SetElement(100)
    # 隐藏网格线
    mc_chart1.api[1].SetElement(328)

    # 固定数值轴量程 0~10，隐藏轴线/刻度/标签
    _val_axis = mc_chart1.api[1].Axes(2)
    _val_axis.MinimumScaleIsAuto = False
    _val_axis.MaximumScaleIsAuto = False
    _val_axis.MinimumScale = 0
    _val_axis.MaximumScale = 10
    _val_axis.TickLabelPosition = -4142
    _val_axis.MajorTickMark = -4142
    _val_axis.MinorTickMark = -4142
    try:
        _val_axis.Format.Line.Visible = 0
    except Exception:
        pass
    print("  [yzr-chart] 坐标轴已固定 0~10，轴线/刻度/标签已隐藏")

    # 数据标签
    try:
        mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()
    except Exception:
        pass

    # 隐藏主标题（SetElement(0) 调用两次，与 questionnaire 一致）
    # 注意：必须保留双调用 —— Main.py 流程下单调用偶尔生效，
    # 但 yzr_ppt.py 单页调试 (__main__) 流程下 COM 时序不同，
    # 单调用会失败、标题不隐藏。双调用是防御性写法，不要再注释掉。
    mc_chart1.api[1].SetElement(0)
    mc_chart1.api[1].SetElement(0)

    # OLE 复制（保留交互性和最高显示质量）
    mc_cell.select()
    mc_chart1.api[0].Copy()
    time.sleep(0.5 + random.random() * 0.3)

    mc_shape = mc_slide.Shapes.Paste()
    time.sleep(0.5)  # 等待 PPT 完成 OLE embed 渲染

    # 断 OLE 热链接（硬规则 #1）
    # 粘贴后 Excel 仍保持 CutCopyMode=True，此时 PPT 与 Excel chart 之间存在
    # COM 热链接；Excel 任何变动（删行/删 chart）都会刷新 PPT。清除后 OLE embed
    # 进入独立显示状态，后续删行/删 chart 不再影响 PPT 图表。
    try:
        xlwings.apps.active.api.CutCopyMode = False
    except Exception:
        pass

    # 还原 PPT 端位置/尺寸
    mc_shape.Left = Left
    mc_shape.Top = Top
    try:
        mc_shape.Width = Width
        mc_shape.Height = Height
    except Exception:
        pass

    # 隐藏主标题（PPT 端再补一刀）—— 上一版 bug 复盘：
    #   旧代码写 `mc_shape.Chart.SetElement(0)`，但 `Shapes.Paste()` 返回的是
    #   **ShapeRange**（不是 Shape），ShapeRange.Chart 直接抛 COM 错
    #   `-2147352567 发生意外`，被外层 except 静默吞掉，title 永远没被隐藏。
    #   .Left/.Top/.Width/.Height 可以在 ShapeRange 上直接设是因为 ShapeRange
    #   会把这些属性 fan-out 到子 shape，但 .Chart 不在 fan-out 列表里。
    # 修法：先 .Item(1) 取真正的单个 Shape，再访问 .Chart；
    #   双保险：HasTitle=False（属性直写） + SetElement(0)（UI 命令）。
    try:
        _shape_one = mc_shape.Item(1) if hasattr(mc_shape, "Item") else mc_shape
        _shape_one.Chart.HasTitle = False
        _shape_one.Chart.SetElement(0)
        print("  [yzr-chart] PPT 端主标题已隐藏")
    except Exception as _e:
        print(f"  [yzr-chart] PPT 端隐藏标题失败（{_e!r}）")

    print(f"  [yzr-chart] 已粘贴至 PPT（L={Left}, T={Top}, W={Width}, H={Height}）")
    return mc_chart1
