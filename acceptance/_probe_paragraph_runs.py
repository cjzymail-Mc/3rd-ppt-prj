#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Self-check probe for 1a/1c changes (throwaway, safe to delete after use).

Tests:
  1. import paragraph_runs; MERGE_DIMS == ("rgb","bold","size")
  2. extract_paragraph_runs(shape) on slide 20 TextBox 6 — 4 new keys present
  3. extract_paragraph_runs(shape, merge_dims=("rgb","bold","size","underline")) — no exception
"""
import sys
import os
from pathlib import Path

# Inject skill dir
_SKILL_DIR = Path.home() / ".claude" / "skills" / "ppt-acceptance-check"
if str(_SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(_SKILL_DIR))

import paragraph_runs
from paragraph_runs import extract_paragraph_runs, MERGE_DIMS

# --- Check 1: MERGE_DIMS unchanged ---
assert MERGE_DIMS == ("rgb", "bold", "size"), f"MERGE_DIMS changed! got {MERGE_DIMS}"
print(f"[PASS] MERGE_DIMS == {MERGE_DIMS}")

# --- Open PPTX via COM ---
import win32com.client

pptx_path = str(
    Path(__file__).resolve().parent / "apparel_v4_smoke.pptx"
)
assert Path(pptx_path).exists(), f"PPTX not found: {pptx_path}"

app = win32com.client.DispatchEx("PowerPoint.Application")
app.DisplayAlerts = 0
pres = app.Presentations.Open(pptx_path, ReadOnly=True, Untitled=False, WithWindow=False)

try:
    slide = pres.Slides(20)

    # Find TextBox 6
    target_shape = None
    for i in range(1, int(slide.Shapes.Count) + 1):
        s = slide.Shapes(i)
        try:
            if str(s.Name) == "TextBox 6":
                target_shape = s
                break
        except Exception:
            pass

    if target_shape is None:
        print("[WARN] TextBox 6 not found on slide 20 — printing all shape names:")
        for i in range(1, int(slide.Shapes.Count) + 1):
            try:
                print(f"  {slide.Shapes(i).Name}")
            except Exception:
                pass
        # Still run the checks on the first text shape we find
        for i in range(1, int(slide.Shapes.Count) + 1):
            s = slide.Shapes(i)
            try:
                if int(s.HasTextFrame) == -1:
                    target_shape = s
                    print(f"[INFO] Using fallback shape: {s.Name}")
                    break
            except Exception:
                pass

    assert target_shape is not None, "No usable shape found on slide 20"

    # --- Check 2: 4 new keys present in first run ---
    paras = extract_paragraph_runs(target_shape)
    assert paras, "extract_paragraph_runs returned empty list"
    first_run = None
    for p in paras:
        if p.get("runs"):
            first_run = p["runs"][0]
            break
    assert first_run is not None, "No non-empty run found"

    print(f"[INFO] First run keys: {list(first_run.keys())}")
    for key in ("underline", "baseline_offset", "font_name_ascii", "font_name_fareast"):
        assert key in first_run, f"Key '{key}' missing from run dict"
        print(f"[PASS] key '{key}' present, value={first_run[key]!r}")

    # --- Check 3: merge_dims override does not raise ---
    paras2 = extract_paragraph_runs(
        target_shape,
        merge_dims=("rgb", "bold", "size", "underline"),
    )
    # Should return list (possibly empty) without raising
    assert isinstance(paras2, list), "Expected list return"
    print(f"[PASS] merge_dims=('rgb','bold','size','underline') -> {len(paras2)} paras, no exception")

finally:
    try:
        pres.Close()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass

print("[ALL PASS] Self-check probe completed successfully.")
