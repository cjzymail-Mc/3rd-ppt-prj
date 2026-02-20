#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PPT pipeline common helpers.

Fixed vs codex-legacy2:
  1. ROOT = parent.parent (project root, not pipeline/)
  2. load_legacy_functions() uses src package import (requires src/__init__.py)
  3. is_in_group(shp) uses try-except (COM getattr workaround)
  4. load_excel_rows() adds fuzzy sheet matching
"""

from __future__ import annotations

import importlib
import json
import statistics
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple


def setup_console_encoding() -> None:
    """No-op stub — kept for call-site compatibility."""
    pass


def safe_print(*args, **kwargs) -> None:
    """print() that survives Windows cp1252 consoles.

    Encodes the line as UTF-8 and replaces any unprintable characters,
    then falls back to a fully ASCII representation.
    Uses regular print() (not buffer.write) to avoid MINGW64 double-output.
    """
    end = kwargs.get("end", "\n")
    line = " ".join(str(a) for a in args)
    try:
        print(line, end=end)
    except UnicodeEncodeError:
        # Replace characters that can't be encoded by the current console
        ascii_line = line.encode(sys.stdout.encoding or "ascii", "replace").decode(
            sys.stdout.encoding or "ascii"
        )
        print(ascii_line, end=end)

# ---- paths ----
ROOT = Path(__file__).resolve().parent.parent          # project root
SRC_DIR = ROOT / "src"
EXCEL_PATH = ROOT / "2025 数据 v2.2.xlsx"
TEMPLATE_PATH = ROOT / "src" / "Template 2.1.pptx"
PROGRESS_DIR = ROOT / "pipeline-progress"
PROGRESS_DIR.mkdir(parents=True, exist_ok=True)        # create on first import


# ---- tiny helpers ----

def now_ts() -> str:
    return datetime.now().isoformat(timespec="seconds")


def safe_text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def to_rows(value: Any) -> List[List[Any]]:
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


def numeric(v: Any):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


# ---- COM safety helpers ----

def com_get(obj, attr: str, default=None):
    """Safe getattr for COM objects."""
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def com_call(obj, method: str, *args, **kwargs):
    """Safe method call on COM objects."""
    try:
        fn = getattr(obj, method)
        return fn(*args, **kwargs)
    except Exception:
        return None


def is_in_group(shp) -> bool:
    """Check if shape is inside a group.

    BUG FIX: getattr(shape, 'ParentGroup', None) raises COM exception
    instead of returning None. Must use try-except.
    """
    try:
        shp.ParentGroup
        return True
    except Exception:
        return False


# ---- I/O helpers ----

def write_md(path: Path, lines: List[str]) -> None:
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# ---- legacy function loader ----

def load_legacy_functions() -> Dict[str, Any]:
    """Load GPT_5 / extract_info from src.Function_030.

    Uses package import (src.__init__.py must exist).
    Falls back to sys.path insertion if package import fails.
    """
    try:
        fn030 = importlib.import_module("src.Function_030")
        return {
            "GPT_5": getattr(fn030, "GPT_5", None),
            "extract_info": getattr(fn030, "extract_info", None),
            "gen_questionnaire_prompt": getattr(fn030, "gen_questionnaire_prompt", None),
            "gen_result_prompt": getattr(fn030, "gen_result_prompt", None),
            "gen_mc_prompt": getattr(fn030, "gen_mc_prompt", None),
        }
    except Exception:
        # fallback: add src/ to path directly
        if str(SRC_DIR) not in sys.path:
            sys.path.insert(0, str(SRC_DIR))
        try:
            fn030 = importlib.import_module("Function_030")
            return {
                "GPT_5": getattr(fn030, "GPT_5", None),
                "extract_info": getattr(fn030, "extract_info", None),
                "gen_questionnaire_prompt": getattr(fn030, "gen_questionnaire_prompt", None),
                "gen_result_prompt": getattr(fn030, "gen_result_prompt", None),
                "gen_mc_prompt": getattr(fn030, "gen_mc_prompt", None),
            }
        except Exception:
            return {"GPT_5": None, "extract_info": None}


# ---- Excel loader ----

def load_excel_rows(sheet_name: str = "问卷sheet") -> Tuple[List[List[Any]], str, List[str]]:
    """Read Excel with pandas + openpyxl (no COM, no xlwings App lifecycle).

    Avoids the xlwings App()/quit() COM event loop that causes double-print
    and exit-code-1 on Windows. pandas reads .xlsx directly via openpyxl.

    Fuzzy sheet matching: exact → contains '问卷' → first sheet.
    Returns all cell values as-is (including None for empty cells).
    """
    notes: List[str] = []
    try:
        import pandas as pd  # type: ignore
    except ImportError as e:
        raise RuntimeError(f"pandas unavailable: {e}")

    # Read sheet names without loading data
    xl = pd.ExcelFile(str(EXCEL_PATH), engine="openpyxl")
    all_sheets = xl.sheet_names

    # 1) exact match
    matched = sheet_name if sheet_name in all_sheets else None
    # 2) fuzzy: sheet name contains '问卷'
    if matched is None:
        for s in all_sheets:
            if "问卷" in s:
                matched = s
                notes.append(f"exact '{sheet_name}' not found, fuzzy match '{s}'")
                break
    # 3) fallback to first sheet
    if matched is None:
        matched = all_sheets[0]
        notes.append(f"sheet '{sheet_name}' not found, fallback '{matched}'")

    # Read with header=None so row 0 = original headers
    df = pd.read_excel(str(EXCEL_PATH), sheet_name=matched,
                       header=None, engine="openpyxl")

    if df.empty:
        raise RuntimeError("Excel sheet is empty")

    # Convert to List[List[Any]], preserving None for empty cells
    rows = []
    for _, row in df.iterrows():
        rows.append([None if pd.isna(v) else v for v in row])

    return rows, matched, notes


# ---- data extraction ----

def extract_metrics(rows: List[List[Any]]) -> Dict[str, Any]:
    headers = [safe_text(v) for v in rows[0]] if rows else []
    data = rows[1:] if len(rows) > 1 else []

    nums = []
    text_cells = []
    for r in data:
        for c in r:
            n = numeric(c)
            if n is not None:
                nums.append(n)
            t = safe_text(c)
            if t:
                text_cells.append(t)

    kws: Dict[str, int] = {}
    hit_words = ["舒适", "稳定", "回弹", "抓地", "缓震", "支撑", "透气"]
    for t in text_cells:
        for w in hit_words:
            if w in t:
                kws[w] = kws.get(w, 0) + 1

    metrics = {
        "respondent_count": len(data),
        "headers": headers,
        "numeric_mean": round(statistics.mean(nums), 3) if nums else None,
        "numeric_median": round(statistics.median(nums), 3) if nums else None,
        "numeric_min": round(min(nums), 3) if nums else None,
        "numeric_max": round(max(nums), 3) if nums else None,
        "keywords": sorted(kws.items(), key=lambda x: x[1], reverse=True),
        "text_preview": text_cells[:40],
    }
    return metrics


def extract_score_means(rows: List[List[Any]]) -> List[Tuple[str, float]]:
    """Extract per-column score means for bar charts.

    Strategy:
    1) Prefer columns whose header contains score-like keywords
    2) Fallback to all numeric columns in score range
    3) Return (header, mean) pairs
    """
    if not rows or len(rows) < 2:
        return []

    headers = [safe_text(h) for h in rows[0]]
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
            n = numeric(r[c])
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

    # BUG FIX: score_like and backup_numeric are mutually exclusive (elif branches),
    # so returning only score_like when non-empty drops all unmatched score columns.
    # E.g. "缓震性", "包裹性", "抗扭转性", "防侧翻性", "耐久性" have no keyword match
    # and land in backup_numeric — their means must be included in the overall average.
    # Return all in-range columns: keyword-matched first, then unmatched.
    return score_like + backup_numeric


# ---- text clamping ----

def clamp_text(text: str, max_chars: int, max_lines: int) -> str:
    """Hard-truncate text to budget constraints."""
    t = safe_text(text)
    if max_chars > 0 and len(t) > max_chars:
        t = t[:max_chars]
    if max_lines > 0:
        lines = t.splitlines() or [t]
        t = "\n".join(lines[:max_lines])
    return t


# ---- human-in-the-loop: shape_detail.md annotation parser ----

SHAPE_DETAIL_MD = PROGRESS_DIR / "01-shape_detail.md"

# Annotation field keys the user can fill in
_ANNO_KEYS = {
    "内容来源": "content_source",      # where the content comes from
    "生成方式": "build_strategy",       # free-text description (human-readable, preserved)
    "修正说明": "fix_notes",            # other corrections / details
    "角色覆盖": "role_override",        # override auto-inferred role
    "prompt覆盖": "prompt_override",    # override default prompt instruction
    # --- structured fields (machine-readable, added alongside natural language) ---
    "strategy": "strategy_exact",       # exact strategy code: score_10pt / grade_letter / ...
    "params": "params",                 # key=value pairs: column=X, filter=Y, format=Z
}

# Valid strategy codes (for documentation / validation)
STRATEGY_CODES = frozenset({
    "score_10pt",       # compute mean → normalize to 10pt → "X.XX/10"
    "grade_letter",     # compute mean → normalize to 100pt → letter grade
    "sample_aggregation",  # extract stats from Excel (no GPT)
    "extract_column",   # pull value from a specific Excel column
    "gpt_prompted",     # GPT with full questionnaire text in prompt
    "mean_extraction",  # bar chart means (chart shapes only)
    "template_direct",  # copy template text verbatim
    "skip",             # decorative / image shape — write nothing
})


def generate_shape_detail_md(shapes: list, existing_annos: dict = None) -> list[str]:
    """Generate shape_detail.md lines with per-shape annotation placeholders.

    Called by Step 1 after extracting shapes. The user can later edit the
    '用户批注' section under each shape to guide subsequent steps.

    existing_annos: dict returned by parse_user_annotations() — if provided,
    each shape's annotation fields are pre-filled from the old md instead of
    left as empty placeholders. Shapes not found in existing_annos get blanks.
    Pass None (or omit) for a fully fresh md (--force mode).
    """
    existing_annos = existing_annos or {}

    lines = [
        "# Shape Detail Report",
        "",
        "> 本文件由 Step 1 自动生成，供用户核验和批注。",
        "> 如果 agents 生成的 PPT 不符合预期，您可以在每个 shape 下方的",
        "> '用户批注' 区域填写修改意见，然后重新启动 agents 工作流。",
        "> Step 2 会自动读取您的批注并作为优先指令。",
        "",
        "---",
        "",
    ]

    for i, s in enumerate(shapes, 1):
        shape_name = s.get("name", f"shape_{i}")
        text_preview = (s.get("text") or "")[:120].replace("\n", "\\n")
        # Restore existing annotations for this shape (empty string if not found)
        anno = existing_annos.get(shape_name, {})

        lines += [
            f"## {i}. {shape_name}",
            "",
            f"- shape_type: {s.get('shape_type', 0)}",
            f"- has_chart: {s.get('has_chart', False)}",
            f"- in_group: {s.get('in_group', False)}",
            f"- left/top: {s.get('left', 0):.1f} / {s.get('top', 0):.1f}",
            f"- width/height: {s.get('width', 0):.1f} / {s.get('height', 0):.1f}",
            f"- font: {s.get('font_name', '')} {s.get('font_size', 0)}",
            f"- z_order: {s.get('z_order', 0)}",
            f"- text: {text_preview}",
            "",
            "### 用户批注",
            "",
            f"- 内容来源: {anno.get('content_source', '')}",
            f"- 生成方式: {anno.get('build_strategy', '')}",
            f"- 修正说明: {anno.get('fix_notes', '')}",
            f"- 角色覆盖: {anno.get('role_override', '')}",
            f"- prompt覆盖: {anno.get('prompt_override', '')}",
            f"- strategy: {anno.get('strategy_exact', '')}",
            f"- params: {anno.get('params', '')}",
            "",
        ]

    lines += [
        "---",
        "",
        "> 填写完毕后，保存此文件，重新运行 agents 工作流即可。",
        "> Step 1 检测到批注后会跳过重新提取，Step 2 会合并您的批注。",
    ]
    return lines


def parse_user_annotations() -> dict[str, dict[str, str]]:
    """Parse user annotations from shape_detail.md.

    Returns: {shape_name: {content_source, build_strategy, fix_notes, ...}}
    Only includes shapes that have at least one non-empty annotation.
    """
    if not SHAPE_DETAIL_MD.exists():
        return {}

    text = SHAPE_DETAIL_MD.read_text(encoding="utf-8")
    result: dict[str, dict[str, str]] = {}

    current_shape = None
    in_annotation = False

    for line in text.splitlines():
        stripped = line.strip()

        # Detect shape header: "## 1. 矩形 11" or "## 2. 图表 44"
        if stripped.startswith("## ") and not stripped.startswith("### "):
            # Extract shape name: everything after "## N. "
            parts = stripped[3:].split(". ", 1)
            if len(parts) == 2:
                current_shape = parts[1].strip()
            else:
                current_shape = stripped[3:].strip()
            in_annotation = False
            continue

        # Detect annotation section
        if stripped == "### 用户批注":
            in_annotation = True
            continue

        # Parse annotation fields
        if in_annotation and current_shape and stripped.startswith("- "):
            field_line = stripped[2:]
            for cn_key, en_key in _ANNO_KEYS.items():
                if field_line.startswith(cn_key + ":") or field_line.startswith(cn_key + ": "):
                    value = field_line[len(cn_key) + 1:].strip()
                    if value:
                        if current_shape not in result:
                            result[current_shape] = {}
                        result[current_shape][en_key] = value
                    break

    return result


def has_user_annotations() -> bool:
    """Check if shape_detail.md exists and contains at least one annotation."""
    return bool(parse_user_annotations())


def parse_params(params_str: str) -> dict:
    """Parse 'key=val, key2=val2' annotation string into a dict.

    Example:
        "source=补充说明, filter=缺点"  ->  {"source": "补充说明", "filter": "缺点"}
        "column=鞋款名称"               ->  {"column": "鞋款名称"}
        ""                             ->  {}
    """
    result: dict = {}
    if not params_str:
        return result
    for part in params_str.split(","):
        part = part.strip()
        if "=" in part:
            k, _, v = part.partition("=")
            result[k.strip()] = v.strip()
    return result
