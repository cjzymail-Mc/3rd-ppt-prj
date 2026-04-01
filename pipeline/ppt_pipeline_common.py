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
import time
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
EXCEL_PATH = ROOT / "pipeline" / "source data.xlsx"
TEMPLATE_PATH = ROOT / "pipeline" / "standard and empty template.pptx"
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


def _rgb(hex_color: str) -> int:
    """Convert 'RRGGBB' hex string to COM RGB long integer."""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + g * 256 + b * 65536


def _set_thin_border(rng) -> None:
    """Apply thin border to all edges of a COM Range."""
    for edge in (7, 8, 9, 10):  # xlLeft, xlTop, xlBottom, xlRight
        rng.Borders(edge).LineStyle = 1  # xlContinuous
        rng.Borders(edge).Weight = 2     # xlThin


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
    """Read Excel via COM (supports encrypted files).

    Fuzzy sheet matching: exact -> contains '问卷' -> first sheet.
    Returns all cell values as-is (including None for empty cells).
    """
    import win32com.client

    notes: List[str] = []
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(EXCEL_PATH.resolve()), 0, True)  # ReadOnly

        # 1) exact match
        matched = None
        matched_name = ""
        for i in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(i).Name == sheet_name:
                matched = wb.Sheets(i)
                matched_name = sheet_name
                break
        # 2) fuzzy: contains '问卷'
        if matched is None:
            for i in range(1, wb.Sheets.Count + 1):
                sname = wb.Sheets(i).Name
                if "问卷" in sname:
                    matched = wb.Sheets(i)
                    matched_name = sname
                    notes.append(f"exact '{sheet_name}' not found, fuzzy match '{sname}'")
                    break
        # 3) fallback to first sheet
        if matched is None:
            matched = wb.Sheets(1)
            matched_name = matched.Name
            notes.append(f"sheet '{sheet_name}' not found, fallback '{matched_name}'")

        raw = matched.UsedRange.Value

        if raw is None:
            raise RuntimeError("Excel sheet is empty")

        # COM returns: tuple of tuples (multi-row), tuple (single row), or scalar
        rows: List[List[Any]] = []
        if isinstance(raw, tuple):
            for row_data in raw:
                if isinstance(row_data, tuple):
                    rows.append(list(row_data))
                else:
                    rows.append([row_data])
        else:
            rows.append([raw])

        if not rows:
            raise RuntimeError("Excel sheet is empty")

        wb.Close(False)
        return rows, matched_name, notes
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


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
    """Soft-clamp: 只限制行数（保护PPT版面），不截断字符（保护信息完整性）。

    字符限制通过 prompt 引导 GPT 控制，不在后处理中强制执行。
    """
    t = safe_text(text)
    # 仅行数限制（防止 PPT 版面溢出）
    if max_lines > 0:
        lines = t.splitlines() or [t]
        t = "\n".join(lines[:max_lines])
    return t


# ---- human-in-the-loop: shape_detail.xlsx annotation parser ----

SHAPE_DETAIL_XLSX = PROGRESS_DIR / "01-shape_detail.xlsx"

# Annotation field keys the user can fill in
# Primary: "内容描述" (natural language); optional: strategy/params/备注
_ANNO_KEYS = {
    # --- primary field (shown in xlsx, pure yellow) ---
    "内容描述": "description",           # natural language: what generates this shape
    # --- optional fields (shown in xlsx, no fill) ---
    "strategy": "strategy_exact",       # exact strategy code: score_10pt / grade_letter / ...
    "params": "params",                 # key=value pairs: column=X, filter=Y, format=Z
    "GPT-prompt Text": "gpt_prompt_text",  # assembled GPT prompt for review/edit
    "备注": "fix_notes",                # legacy: merged into 内容描述, still parsed for old sheets
    # --- legacy keys (still parsed if present, not generated in new xlsx) ---
    "内容来源": "content_source",
    "生成方式": "build_strategy",
    "修正说明": "fix_notes",
    "角色覆盖": "role_override",
    "prompt覆盖": "prompt_override",
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


def generate_shape_detail_xlsx(
    shapes: list,
    existing_annos: dict = None,
    sheet_name: str = "Shape Detail",
) -> None:
    """Generate shape_detail.xlsx via COM (supports encrypted environments).

    Writes directly to SHAPE_DETAIL_XLSX. Each shape block has:
    - Header row (blue): "Shape #N" | shape_name
    - Property rows with borders
    - Annotation header (green): "用户批注"
    - Primary field (yellow): "内容描述"
    - Optional fields: strategy / params / 备注
    - 4 blank rows gap between shapes

    sheet_name: Name of the sheet to create (default "Shape Detail").
    """
    import win32com.client

    existing_annos = existing_annos or {}
    xlsx_path = str(SHAPE_DETAIL_XLSX.resolve())

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Add()
        ws = wb.Sheets(1)
        ws.Name = sheet_name
        ws.Columns(1).ColumnWidth = 18
        ws.Columns(2).ColumnWidth = 75

        # Colors
        BLUE = _rgb("4472C4")
        GREEN = _rgb("E2EFDA")
        YELLOW = _rgb("FFFF00")
        RED = _rgb("FF0000")
        WHITE = _rgb("FFFFFF")
        GREY = _rgb("999999")
        DARK_GREY = _rgb("666666")

        r = 1
        # Row 1: title
        c = ws.Cells(r, 1)
        c.Value = "Shape Detail Report"
        c.Font.Bold = True
        c.Font.Size = 14
        r += 1

        # Row 2: subtitle
        c = ws.Cells(r, 1)
        c.Value = '编辑"用户批注"区域的B列，保存后运行 Step 2 即可生效'
        c.Font.Size = 10
        c.Font.Italic = True
        r += 2  # skip blank row

        # --- Instruction: 内容描述 (primary) ---
        ws.Cells(r, 1).Value = "填写说明"
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 11
        r += 1

        ws.Cells(r, 1).Value = "内容描述"
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 10
        ws.Cells(r, 2).Value = "必填。描述内容来源和生成方式，可追加详细的 GPT 约束（如关键词、字数、格式要求）。"
        ws.Cells(r, 2).Font.Size = 10
        r += 1

        ws.Cells(r, 2).Value = "系统会自动识别关键词来确定生成策略，无需手动指定代码。"
        ws.Cells(r, 2).Font.Size = 10
        r += 1

        ws.Cells(r, 2).Value = "示例:"
        ws.Cells(r, 2).Font.Size = 10
        r += 1

        for example in [
            "  \u00b7 评分均值10分制              \u2192 自动计算均值，输出 X.XX/10",
            "  \u00b7 鞋款名称                    \u2192 自动从Excel提取鞋款名称",
            "  \u00b7 从补充说明总结缺点           \u2192 GPT读取问卷原文，总结缺点",
            "  \u00b7 不走GPT统计人数体重          \u2192 Python直接统计，不调用GPT",
            "  \u00b7 (留空)                      \u2192 系统根据shape类型自动推断",
        ]:
            ws.Cells(r, 2).Value = example
            ws.Cells(r, 2).Font.Size = 10
            ws.Cells(r, 2).Font.Color = DARK_GREY
            r += 1

        r += 1  # blank row

        # --- Instruction: optional params ---
        ws.Cells(r, 1).Value = "可选参数"
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 11
        ws.Cells(r, 1).Font.Color = GREY
        ws.Cells(r, 2).Value = "高级选项，可不填。填写后优先级高于内容描述。"
        ws.Cells(r, 2).Font.Size = 10
        ws.Cells(r, 2).Font.Color = GREY
        r += 1

        ws.Cells(r, 1).Value = "strategy"
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 10
        ws.Cells(r, 2).Value = "精确策略代码，优先级最高，覆盖内容描述的自动识别。"
        ws.Cells(r, 2).Font.Size = 10
        r += 1
        ws.Cells(r, 2).Value = "可选值: score_10pt / grade_letter / sample_aggregation / extract_column"
        ws.Cells(r, 2).Font.Size = 10
        ws.Cells(r, 2).Font.Color = DARK_GREY
        r += 1
        ws.Cells(r, 2).Value = "        gpt_prompted / mean_extraction / template_direct / skip"
        ws.Cells(r, 2).Font.Size = 10
        ws.Cells(r, 2).Font.Color = DARK_GREY
        r += 1

        ws.Cells(r, 1).Value = "params"
        ws.Cells(r, 1).Font.Bold = True
        ws.Cells(r, 1).Font.Size = 10
        ws.Cells(r, 2).Value = "策略参数，key=value 逗号分隔。"
        ws.Cells(r, 2).Font.Size = 10
        r += 1
        ws.Cells(r, 2).Value = "示例: source=补充说明, filter=缺点, column=鞋款名称"
        ws.Cells(r, 2).Font.Size = 10
        ws.Cells(r, 2).Font.Color = DARK_GREY
        r += 1

        r += 1  # blank row before shapes

        # --- Per-shape blocks ---
        for i, s in enumerate(shapes, 1):
            shape_name = s.get("name", f"shape_{i}")
            text_preview = (s.get("text") or "")[:120].replace("\n", " ")
            anno = existing_annos.get(shape_name, {})

            # Shape header (blue)
            for col in (1, 2):
                c = ws.Cells(r, col)
                c.Interior.Color = BLUE
                c.Font.Bold = True
                c.Font.Size = 11
                c.Font.Color = WHITE
                _set_thin_border(c)
            ws.Cells(r, 1).Value = f"Shape #{i}"
            ws.Cells(r, 2).Value = shape_name
            r += 1

            # Properties
            props = [
                ("shape_type", str(s.get("shape_type", 0))),
                ("has_chart", str(s.get("has_chart", False))),
                ("in_group", str(s.get("in_group", False))),
                ("left/top", f"{s.get('left', 0):.1f} / {s.get('top', 0):.1f}"),
                ("width/height", f"{s.get('width', 0):.1f} / {s.get('height', 0):.1f}"),
                ("font", f"{s.get('font_name', '')} {s.get('font_size', 0)}"),
                ("z_order", str(s.get("z_order", 0))),
                ("text", text_preview),
            ]
            for key, val in props:
                ca = ws.Cells(r, 1)
                ca.Value = key
                _set_thin_border(ca)
                cb = ws.Cells(r, 2)
                cb.Value = val
                _set_thin_border(cb)
                if key == "text":
                    cb.Font.Bold = True
                    cb.Font.Color = RED
                r += 1

            # Annotation header (green)
            for col in (1, 2):
                c = ws.Cells(r, col)
                c.Interior.Color = GREEN
                c.Font.Bold = True
                c.Font.Size = 10
                _set_thin_border(c)
            ws.Cells(r, 1).Value = "用户批注"
            ws.Cells(r, 2).Value = "← 该内容生成的原理是什么？请在下方填写说明"
            r += 1

            # Primary field: 内容描述 (yellow)
            ca = ws.Cells(r, 1)
            ca.Value = "内容描述"
            _set_thin_border(ca)
            cb = ws.Cells(r, 2)
            cb.Value = anno.get("description", "")
            cb.Interior.Color = YELLOW
            _set_thin_border(cb)
            r += 1

            # Optional fields (border only, no fill, grey label)
            for key, val in [
                ("strategy", anno.get("strategy_exact", "")),
                ("params", anno.get("params", "")),
                ("GPT-prompt Text", anno.get("gpt_prompt_text", "")),
            ]:
                ca = ws.Cells(r, 1)
                ca.Value = key
                ca.Font.Color = GREY
                ca.Font.Bold = True
                _set_thin_border(ca)
                cb = ws.Cells(r, 2)
                cb.Value = val
                cb.WrapText = True
                _set_thin_border(cb)
                r += 1

            # 4 blank rows gap
            r += 4

        # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.SaveAs(xlsx_path, 51)
        wb.Close(False)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def parse_user_annotations(sheet_name: str = None) -> dict[str, dict[str, str]]:
    """Parse user annotations from shape_detail.xlsx via COM (supports encrypted files).

    sheet_name: If provided, read from the sheet with this name.
                If None, read from the first sheet (default, backward-compatible).

    Returns: {shape_name: {content_source, build_strategy, fix_notes, ...}}
    Only includes shapes that have at least one non-empty annotation.
    """
    if not SHAPE_DETAIL_XLSX.exists():
        return {}

    import win32com.client
    import time as _time

    # Use DispatchEx to force a NEW Excel instance, avoiding conflict with
    # any lingering COM process from previous pipeline steps.
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    _time.sleep(0.5)  # let COM server stabilize

    result: dict[str, dict[str, str]] = {}

    try:
        wb = excel.Workbooks.Open(str(SHAPE_DETAIL_XLSX.resolve()), 0, True)
        if sheet_name:
            ws = None
            for i in range(1, wb.Sheets.Count + 1):
                if wb.Sheets(i).Name == sheet_name:
                    ws = wb.Sheets(i)
                    break
            if ws is None:
                safe_print(f"[WARN] parse_user_annotations: sheet '{sheet_name}' not found, using first sheet")
                ws = wb.Sheets(1)
        else:
            ws = wb.Sheets(1)
        max_row = ws.UsedRange.Rows.Count

        current_shape = None
        in_annotation = False

        for r in range(1, max_row + 1):
            a_val = safe_text(ws.Cells(r, 1).Value)
            b_val = safe_text(ws.Cells(r, 2).Value)

            if a_val.startswith("Shape #"):
                current_shape = b_val
                in_annotation = False
                continue

            if a_val == "用户批注":
                in_annotation = True
                continue

            if in_annotation and current_shape and a_val in _ANNO_KEYS:
                en_key = _ANNO_KEYS[a_val]
                if b_val:
                    if current_shape not in result:
                        result[current_shape] = {}
                    result[current_shape][en_key] = b_val

        wb.Close(False)
    except Exception as e:
        safe_print(f"[WARN] parse_user_annotations COM error: {e}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

    return result


def write_gpt_prompts_to_xlsx(
    prompts: dict[str, str],
    sheet_name: str = None,
) -> None:
    """Write assembled GPT prompts to 'GPT-prompt Text' cells in shape_detail.xlsx.

    prompts: {shape_name: prompt_text}
    """
    if not prompts or not SHAPE_DETAIL_XLSX.exists():
        return

    import win32com.client
    import time as _time

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    _time.sleep(0.5)

    try:
        wb = excel.Workbooks.Open(str(SHAPE_DETAIL_XLSX.resolve()))
        if sheet_name:
            ws = None
            for i in range(1, wb.Sheets.Count + 1):
                if wb.Sheets(i).Name == sheet_name:
                    ws = wb.Sheets(i)
                    break
            if ws is None:
                ws = wb.Sheets(wb.Sheets.Count)
        else:
            ws = wb.Sheets(wb.Sheets.Count)

        max_row = ws.UsedRange.Rows.Count
        current_shape = None
        written = 0

        for r in range(1, max_row + 1):
            a_val = safe_text(ws.Cells(r, 1).Value)
            b_val_raw = ws.Cells(r, 2).Value

            if a_val.startswith("Shape #"):
                current_shape = safe_text(b_val_raw)
                continue

            if a_val == "GPT-prompt Text" and current_shape and current_shape in prompts:
                cell = ws.Cells(r, 2)
                cell.Value = prompts[current_shape]
                cell.WrapText = True
                written += 1

        wb.Save()
        wb.Close(False)
        safe_print(f"[OK] 写入 {written} 个 GPT prompt 到 xlsx")
    except Exception as e:
        safe_print(f"[WARN] write_gpt_prompts_to_xlsx COM error: {e}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def read_gpt_prompts_from_xlsx(sheet_name: str = None) -> dict[str, str]:
    """Read 'GPT-prompt Text' cell values from shape_detail.xlsx.

    Returns: {shape_name: prompt_text} (only non-empty entries).
    """
    if not SHAPE_DETAIL_XLSX.exists():
        return {}

    import win32com.client
    import time as _time

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    _time.sleep(0.5)

    result: dict[str, str] = {}
    try:
        wb = excel.Workbooks.Open(str(SHAPE_DETAIL_XLSX.resolve()), 0, True)
        if sheet_name:
            ws = None
            for i in range(1, wb.Sheets.Count + 1):
                if wb.Sheets(i).Name == sheet_name:
                    ws = wb.Sheets(i)
                    break
            if ws is None:
                ws = wb.Sheets(wb.Sheets.Count)
        else:
            ws = wb.Sheets(wb.Sheets.Count)

        max_row = ws.UsedRange.Rows.Count
        current_shape = None

        for r in range(1, max_row + 1):
            a_val = safe_text(ws.Cells(r, 1).Value)
            b_val = safe_text(ws.Cells(r, 2).Value)

            if a_val.startswith("Shape #"):
                current_shape = b_val
                continue

            if a_val == "GPT-prompt Text" and current_shape and b_val:
                result[current_shape] = b_val

        wb.Close(False)
    except Exception as e:
        safe_print(f"[WARN] read_gpt_prompts_from_xlsx COM error: {e}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

    return result


def has_user_annotations(sheet_name: str = None) -> bool:
    """Check if shape_detail.xlsx exists and contains at least one annotation."""
    return bool(parse_user_annotations(sheet_name=sheet_name))


def create_iteration_sheet(new_sheet_name: str) -> str:
    """Copy the last sheet in shape_detail.xlsx to a new sheet with the given name.

    Used for multi-round traceability: each iteration round gets its own sheet
    (e.g. "claude-ppt 1.1", "claude-ppt 1.2") so Builder can update annotations
    without overwriting previous rounds.

    Returns the new sheet name on success, or "" on failure.
    """
    if not SHAPE_DETAIL_XLSX.exists():
        safe_print("[WARN] create_iteration_sheet: xlsx not found")
        return ""

    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(SHAPE_DETAIL_XLSX.resolve()))
        # Check if sheet already exists
        for i in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(i).Name == new_sheet_name:
                safe_print(f"[INFO] create_iteration_sheet: sheet '{new_sheet_name}' already exists")
                wb.Close(False)
                return new_sheet_name

        # Copy the last sheet (most recent iteration) to create a new one
        last_sheet = wb.Sheets(wb.Sheets.Count)
        last_sheet.Copy(After=last_sheet)
        new_ws = wb.Sheets(wb.Sheets.Count)
        new_ws.Name = new_sheet_name

        wb.Save()
        wb.Close(False)
        safe_print(f"[OK] create_iteration_sheet: created sheet '{new_sheet_name}'")
        return new_sheet_name
    except Exception as e:
        safe_print(f"[WARN] create_iteration_sheet COM error: {e}")
        return ""
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


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
