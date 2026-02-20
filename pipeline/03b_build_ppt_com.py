#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 3B: Clone template slide 15 and write content via COM.

Flow:
  1. Open PowerPoint, open Template
  2. Clone slide 15 to new Presentation
  3. Read build_shape_content.json
  4. Match shapes by name (fallback: geometry position)
  5. Text shapes: TextFrame.TextRange.Text = content (preserve formatting)
  6. Chart shapes: write to ChartData worksheet
  7. Post-write readback verification
  8. Save as codex {version}.pptx

Fixed vs codex-legacy2:
  - Chart data: rsplit(":", 1) instead of split(":")
  - Chart workbook: wb.Close(False) after writing
  - Copy/Paste: time.sleep(delay) buffer
  - Geometry fallback for shape matching
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    ROOT,
    PROGRESS_DIR,
    TEMPLATE_PATH,
    com_call,
    com_get,
    now_ts,
    safe_print,
    safe_text,
    setup_console_encoding,
    write_json,
    write_md,
)

CONTENT_JSON = PROGRESS_DIR / "03a-build_shape_content.json"
OUT_REPORT   = PROGRESS_DIR / "03b-build_ppt_report.md"
OUT_READBACK = PROGRESS_DIR / "03b-post_write_readback.json"

COPY_PASTE_DELAY = 1.5  # seconds buffer for COM clipboard operations


def _find_shape_by_name(slide, name: str):
    """Find shape by exact name match."""
    try:
        return slide.Shapes(name)
    except Exception:
        return None


def _find_shape_by_geometry(slide, item: dict, tolerance: float = 15.0):
    """Fallback: find shape by closest geometry match."""
    target_left = item.get("left", 0)
    target_top = item.get("top", 0)
    best = None
    best_dist = float("inf")

    for i in range(1, int(slide.Shapes.Count) + 1):
        shp = slide.Shapes(i)
        left = float(com_get(shp, "Left", 0) or 0)
        top = float(com_get(shp, "Top", 0) or 0)
        dist = abs(left - target_left) + abs(top - target_top)
        if dist < best_dist:
            best_dist = dist
            best = shp

    if best is not None and best_dist <= tolerance:
        return best
    return None


def _write_text(shp, content: str) -> dict:
    """Write text to shape, preserving formatting."""
    name = safe_text(com_get(shp, "Name", ""))
    if not bool(com_get(shp, "HasTextFrame", 0)):
        return {"shape_name": name, "updated": False, "reason": "no text frame"}

    tf = com_get(shp, "TextFrame", None)
    tr = com_get(tf, "TextRange", None) if tf is not None else None
    if tr is None:
        return {"shape_name": name, "updated": False, "reason": "no text range"}

    # Lock shape size to preserve template geometry (prevent AutoFit resize)
    try:
        tf.AutoSize = 0  # ppAutoSizeNone = 0
    except Exception:
        pass

    try:
        tr.Text = content
        rb_text = com_get(tr, "Text", "") or ""
        return {
            "shape_name": name,
            "updated": True,
            "mode": "text",
            "written_len": len(content),
            "readback_len": len(rb_text),
        }
    except Exception as e:
        return {"shape_name": name, "updated": False, "reason": str(e)}


def _write_chart(shp, content: str) -> dict:
    """Write chart data via SeriesCollection — bypasses ChartData.Workbook entirely.

    Strategy:
    1. Activate() + BreakLink() in an isolated try block to clear external link refs
       (Activate is required for BreakLink; both may fail on non-interactive sessions
       but the data write below will still succeed)
    2. Parse content ("label:value" per line) into labels/values lists
    3. Set series.Values and series.XValues directly on SeriesCollection(1)
       — no workbook open/close, no "linked file" dialog risk
    """
    name = safe_text(com_get(shp, "Name", ""))
    chart = com_get(shp, "Chart", None)
    if chart is None:
        return {"shape_name": name, "updated": False, "reason": "no chart object"}

    # Parse content first (fail fast if nothing to write)
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
        return {"shape_name": name, "updated": False, "reason": "no parseable data in content"}

    try:
        # Step 1: Break external link.
        # Activate() is required for BreakLink(), but throws on non-interactive sessions.
        # Isolated try — failure here does not abort data write.
        try:
            chart.ChartData.Activate()
            time.sleep(0.5)
            chart.ChartData.BreakLink()
            time.sleep(0.3)
        except Exception:
            pass

        # Step 2: Write data via SeriesCollection — no Workbook access, no linked file risk.
        series = chart.SeriesCollection(1)
        series.Values = tuple(values)
        series.XValues = tuple(labels)

        ctype = int(com_get(chart, "ChartType", 0) or 0)
        return {
            "shape_name": name,
            "updated": True,
            "mode": "chart",
            "chart_type": ctype,
            "data_rows": len(labels),
        }
    except Exception as e:
        return {"shape_name": name, "updated": False, "reason": str(e)}


# Color map: color_hint string → PowerPoint RGB integer (same values as Function_030.py)
_COLOR_MAP = {
    "red":  255,        # red  = RGB(255, 0, 0)
    "blue": 15773696,   # light_blue = RGB(0, 176, 240)
}


def _apply_keyword_color(shp, color_rgb: int) -> None:
    """Remove 【】 brackets from shape text, then bold+color the bracketed keywords.

    Logic mirrors color_key / smart_color_text from Function_030.py, simplified
    for single-color shapes (each shape is either 优点 OR 缺点, not mixed).
    """
    try:
        tf = com_get(shp, "TextFrame", None)
        if tf is None:
            return
        tr = tf.TextRange
        full_text = tr.Text

        # Extract unique keywords (preserve order)
        keywords = list(dict.fromkeys(re.findall(r'【([^】]+)】', full_text)))
        if not keywords:
            return

        # Remove all 【】 brackets from displayed text
        tr.Text = re.sub(r'[【】]', '', full_text)

        # Bold + color each keyword using Find loop (same as color_key in Function_030.py)
        for kw in keywords:
            start = 1
            while start <= tr.Length:
                found = tr.Find(kw, start)
                if found is None:
                    break
                found.Font.Bold = True
                found.Font.Color = color_rgb
                start = found.Start + found.Length
    except Exception:
        pass  # coloring is cosmetic — never fail the build


def _replace_image(slide, shp, img_path: str) -> dict:
    """Replace an image shape with a picture loaded from img_path.

    Records the original geometry, deletes the old shape, then inserts a
    new picture at exactly the same Left/Top/Width/Height.
    """
    name = safe_text(com_get(shp, "Name", ""))
    try:
        left   = float(com_get(shp, "Left",   0))
        top    = float(com_get(shp, "Top",    0))
        width  = float(com_get(shp, "Width",  100))
        height = float(com_get(shp, "Height", 100))
        shp.Delete()
        new_shp = slide.Shapes.AddPicture(
            FileName=img_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left, Top=top,
            Width=width, Height=height,
        )
        try:
            new_shp.Name = name  # restore original name for diff-test matching
        except Exception:
            pass
        return {
            "shape_name": name,
            "updated": True,
            "mode": "image_replace",
            "img_path": img_path,
        }
    except Exception as e:
        return {"shape_name": name, "updated": False, "reason": str(e)}


def apply_shape(slide, item: dict, shape_detail: dict = None) -> dict:
    """Find and update a single shape on the slide."""
    name = item.get("shape_name", "")
    content = item.get("content", "")
    role = item.get("role", "")
    strategy = item.get("strategy", "")

    # Skip shapes explicitly marked as skip — preserve cloned template geometry
    if strategy == "skip":
        return {"shape_name": name, "updated": False, "reason": "strategy_skip", "match_method": "n/a"}

    # Try name match first, then geometry fallback
    shp = _find_shape_by_name(slide, name)
    match_method = "name"
    if shp is None and shape_detail:
        shp = _find_shape_by_geometry(slide, shape_detail)
        match_method = "geometry"
    if shp is None:
        return {"shape_name": name, "updated": False, "reason": "shape not found", "match_method": "none"}

    # Route to appropriate writer
    color_hint = item.get("color_hint", "")
    if strategy == "extract_image":
        if content.startswith("IMAGE:"):
            result = _replace_image(slide, shp, content[6:].strip())
        else:
            result = {"shape_name": name, "updated": False, "reason": "extract_image: no IMAGE: path in content"}
    elif role == "chart" or bool(com_get(shp, "HasChart", False)):
        result = _write_chart(shp, content)
    elif bool(com_get(shp, "HasTextFrame", 0)):
        result = _write_text(shp, content)
        # Post-write keyword coloring (优点=red, 缺点=blue)
        if color_hint in _COLOR_MAP and result.get("updated"):
            _apply_keyword_color(shp, _COLOR_MAP[color_hint])
    else:
        result = {"shape_name": name, "updated": False, "reason": "unsupported shape type"}

    result["match_method"] = match_method
    return result


def main() -> int:
    setup_console_encoding()
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default="1.0")
    args = ap.parse_args()

    out_ppt = ROOT / f"codex {args.version}.pptx"

    if not TEMPLATE_PATH.exists():
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", "- 原因: 模板文件缺失"])
        safe_print("[BLOCKED] Template missing")
        return 0

    if not CONTENT_JSON.exists():
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", "- 原因: build_shape_content.json缺失"])
        safe_print("[BLOCKED] Content JSON missing")
        return 0

    content_data = json.loads(CONTENT_JSON.read_text(encoding="utf-8"))
    items = content_data.get("items", [])

    # Load shape_detail for geometry fallback
    shape_detail_path = PROGRESS_DIR / "01-shape_detail_com.json"
    shape_details = {}
    if shape_detail_path.exists():
        sd = json.loads(shape_detail_path.read_text(encoding="utf-8"))
        for s in sd.get("new_shapes", []):
            shape_details[s.get("name", "")] = s

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", f"- 原因: {e}"])
        safe_print(f"[BLOCKED] {e}")
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    src = app.Presentations.Open(str(TEMPLATE_PATH))
    dst = app.Presentations.Add()

    results = []
    try:
        # Clone slide 15 to new presentation
        src.Slides(15).Copy()
        time.sleep(COPY_PASTE_DELAY)  # COM clipboard buffer
        dst.Slides.Paste()
        time.sleep(COPY_PASTE_DELAY)
        slide = dst.Slides(1)

        for item in items:
            detail = shape_details.get(item.get("shape_name", ""))
            results.append(apply_shape(slide, item, detail))

        dst.SaveAs(str(out_ppt))

        updated = sum(1 for r in results if r.get("updated"))
        write_json(OUT_READBACK, {
            "generated_at": now_ts(),
            "target": out_ppt.name,
            "results": results,
        })
        write_md(OUT_REPORT, [
            "# Build PPT Report",
            "",
            "- 状态: ok",
            f"- 产物: {out_ppt.name}",
            f"- 更新shape: {updated}/{len(results)}",
            f"- 时间: {now_ts()}",
            "",
            "## Details",
            "",
            "|shape|updated|mode|match|reason|",
            "|---|---|---|---|---|",
        ] + [
            f"|{r.get('shape_name','')}|{r.get('updated','')}|{r.get('mode','')}|{r.get('match_method','')}|{r.get('reason','')}|"
            for r in results
        ])
        safe_print(f"[OK] Generated {out_ppt.name}, updated {updated}/{len(results)} shapes")
        return 0
    finally:
        com_call(src, "Close")
        com_call(dst, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
