#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PPT pipeline common helpers."""

from __future__ import annotations

import importlib
import json
import statistics
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

ROOT = Path(__file__).resolve().parent
SRC_DIR = ROOT / "src"
EXCEL_PATH = ROOT / "2025 数据 v2.2.xlsx"
TEMPLATE_PATH = ROOT / "src" / "Template 2.1.pptx"


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


def com_get(obj, attr: str, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def com_call(obj, method: str, *args, **kwargs):
    try:
        fn = getattr(obj, method)
        return fn(*args, **kwargs)
    except Exception:
        return None


def write_md(path: Path, lines: List[str]) -> None:
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_json(path: Path, data: Any) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_legacy_functions():
    """Load GPT_5 / extract_info from existing Function_030.py."""
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))
    try:
        fn030 = importlib.import_module("Function_030")
        return {
            "GPT_5": getattr(fn030, "GPT_5", None),
            "extract_info": getattr(fn030, "extract_info", None),
        }
    except Exception:
        return {"GPT_5": None, "extract_info": None}


def load_excel_rows(sheet_name: str = "问卷sheet") -> Tuple[List[List[Any]], str, List[str]]:
    """Read Excel with xlwings + COM API."""
    notes: List[str] = []
    try:
        import xlwings as xw  # type: ignore
    except Exception as e:
        raise RuntimeError(f"xlwings unavailable: {e}")

    app = xw.App(visible=False, add_book=False)
    wb = None
    try:
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(EXCEL_PATH))
        try:
            sht = wb.sheets[sheet_name]
        except Exception:
            sht = wb.sheets[0]
            notes.append(f"sheet '{sheet_name}' not found, fallback '{sht.name}'")

        rows = to_rows(sht.api.UsedRange.Value)
        if not rows:
            raise RuntimeError("Excel UsedRange is empty")
        return rows, sht.name, notes
    finally:
        if wb is not None:
            wb.close()
        app.quit()


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


def clamp_text(text: str, max_chars: int, max_lines: int) -> str:
    t = safe_text(text)
    if max_chars > 0 and len(t) > max_chars:
        t = t[:max_chars]
    if max_lines > 0:
        lines = t.splitlines() or [t]
        t = "\n".join(lines[:max_lines])
    return t
