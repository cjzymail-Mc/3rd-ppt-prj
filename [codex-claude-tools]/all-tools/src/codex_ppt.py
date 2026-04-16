#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Codex evaluation slide builder — zero pipeline/ dependency.

Self-contained: all helpers are copied/adapted from pipeline/03a and 03b.
Only public API: make_codex_slide().
"""

from __future__ import annotations

import re
import time
import tempfile
from pathlib import Path
from typing import Any, List, Tuple

# The only src-internal dependency: GPT_5 (relative import when used as package,
# absolute import when run standalone)
GPT_5 = None
try:
    from .Function_030 import GPT_5  # type: ignore
except Exception:
    try:
        from src.Function_030 import GPT_5  # type: ignore
    except Exception:
        GPT_5 = None

# ---------------------------------------------------------------------------
# Colors (must match main.py globals)
# ---------------------------------------------------------------------------
_RED  = 255        # red   = RGB(255, 0, 0)
_BLUE = 15773696   # light_blue = RGB(0, 176, 240)

# Default GPT model
_MODEL = "openai/gpt-5.2"

# Clipboard copy-paste COM buffer (seconds)
_COPY_PASTE_DELAY = 1.5

# ---------------------------------------------------------------------------
# Hardcoded shape specs (from 01-shape_detail.md annotations)
# ---------------------------------------------------------------------------
CODEX_SHAPES = [
    {"name": "矩形 11",   "strategy": "score_10pt",        "color_hint": ""},
    {"name": "矩形 12",   "strategy": "grade_letter",       "color_hint": ""},
    {"name": "矩形 17",   "strategy": "sample_aggregation", "color_hint": ""},
    {"name": "矩形 19",   "strategy": "skip",               "color_hint": ""},
    {"name": "图片 39",   "strategy": "extract_image",      "color_hint": ""},
    {"name": "文本框 16", "strategy": "extract_column",     "color_hint": "",
     "params": {"column": "鞋款名称"}},
    {"name": "矩形 68",   "strategy": "gpt_prompted",       "color_hint": "blue",
     "params": {"source": "补充说明", "filter": "缺点"},
     "budget": {"max_chars": 270, "max_lines": 9}},
    {"name": "矩形 77",   "strategy": "gpt_prompted",       "color_hint": "red",
     "params": {"source": "补充说明", "filter": "优点"},
     "budget": {"max_chars": 201, "max_lines": 5}},
    {"name": "图表 44",   "strategy": "mean_extraction",    "color_hint": ""},
]

# ---------------------------------------------------------------------------
# Tiny helper utilities (self-contained, no pipeline import)
# ---------------------------------------------------------------------------

def _safe_text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _numeric(v: Any):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


def _com_get(obj, attr: str, default=None):
    """Safe getattr for COM objects (getattr raises on COM objects)."""
    try:
        return getattr(obj, attr)