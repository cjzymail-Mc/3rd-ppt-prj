#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""_ppt_shared.py — yzr / zxh 共享的纯数据/纯计算工具.

范围严格限定：**不涉及 PPT 写入、不影响视觉输出的纯函数**。
目的：降低 yzr_ppt.py / zxh_ppt.py 间的跨文件 bug 风险（如 _find_col 改一边忘了另一边），
但保留视觉写入、prompt、模板 shape 等 per-template 可独立微调的函数在各自文件中。

详见 [feature03-transplant]/fix2.md Fix 2 (partial)。
"""

from __future__ import annotations

from typing import Any, List, Optional, Tuple


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


def _classify_columns(headers: List[str], rows: List[List[Any]]
                      ) -> Tuple[List[str], List[str]]:
    """Dynamically classify columns into score (numeric) and text (feedback).

    Score columns: >70% of data rows have numeric values in 0-10 range.
    Text columns:  >30% of data rows have string values longer than 5 chars.
    """
    name_col = _find_col(headers, _NAME_KEYWORDS)
    weight_col = _find_col(headers, _WEIGHT_KEYWORDS)
    skip_cols = {name_col, weight_col} | {
        h for h in headers
        if any(kw in h for kw in _SKIP_KEYWORDS)
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
    reject_keys = ["姓名", "昵称", "电话", "联系方式", "地址", "微信", "备注", "日期", "时间"]

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
# 文本处理（纯字符串）
# ===========================================================================
def clamp_text(text: str, max_chars: int, max_lines: int) -> str:
    """Clamp text to fit PPT shape: enforce both line count and character count.

    Line clamp: hard cut at max_lines.
    Char clamp: hard cut at sentence boundary when exceeding max_chars.
    """
    t = _safe_text(text)
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
