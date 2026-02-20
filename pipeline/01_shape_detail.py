#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 1: Compare Template page 14 (blank) vs page 15 (standard).

Extracts new shapes' attributes and fingerprints.
Outputs:
  - shape_detail_com.json   (for agents / code)
  - shape_fingerprint_map.json
  - shape_detail.md          (for user review + annotation)

Human-in-the-loop:
  If shape_detail.md already contains user annotations, Step 1 re-extracts
  shapes from the template and merges existing annotations into the new md.
  Shapes with annotations have them restored; new shapes get empty placeholders.
  Use --force to clear all annotations and start fresh.
"""

from __future__ import annotations

import sys
from pathlib import Path

# ensure project root is on sys.path so pipeline package can find ppt_pipeline_common
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    PROGRESS_DIR,
    SHAPE_DETAIL_MD,
    TEMPLATE_PATH,
    com_get,
    com_call,
    generate_shape_detail_md,
    parse_user_annotations,
    is_in_group,
    now_ts,
    safe_print,
    safe_text,
    setup_console_encoding,
    write_json,
    write_md,
)

OUT_JSON = PROGRESS_DIR / "01-shape_detail_com.json"
OUT_FP = PROGRESS_DIR / "01-shape_fingerprint_map.json"
OUT_MD = SHAPE_DETAIL_MD  # pipeline-progress/01-shape_detail.md


def _safe_has_chart(shp) -> bool:
    """COM-safe check for HasChart."""
    try:
        return bool(shp.HasChart)
    except Exception:
        return False


def _safe_text(shp) -> str:
    """COM-safe extraction of shape text."""
    try:
        return shp.TextFrame.TextRange.Text or ""
    except Exception:
        return ""


def _safe_font(shp):
    """COM-safe extraction of font name and size."""
    try:
        font = shp.TextFrame.TextRange.Font
        name = safe_text(com_get(font, "Name", ""))
        size = float(com_get(font, "Size", 0.0) or 0.0)
        return name, size
    except Exception:
        return "", 0.0


def shape_obj(slide_index: int, shp, idx: int) -> dict:
    name = com_get(shp, "Name", "") or f"shape_{slide_index}_{idx}"
    text = _safe_text(shp)
    font_name, font_size = _safe_font(shp)

    return {
        "slide_index": slide_index,
        "name": safe_text(name),
        "left": float(com_get(shp, "Left", 0.0) or 0.0),
        "top": float(com_get(shp, "Top", 0.0) or 0.0),
        "width": float(com_get(shp, "Width", 0.0) or 0.0),
        "height": float(com_get(shp, "Height", 0.0) or 0.0),
        "shape_type": int(com_get(shp, "Type", 0) or 0),
        "text": text,
        "font_name": font_name,
        "font_size": font_size,
        "has_chart": _safe_has_chart(shp),
        "in_group": is_in_group(shp),
        "z_order": int(com_get(shp, "ZOrderPosition", 0) or 0),
    }


def fingersafe_print(item: dict) -> dict:
    return {
        "name": item["name"],
        "shape_type": item["shape_type"],
        "left": round(item["left"], 1),
        "top": round(item["top"], 1),
        "width": round(item["width"], 1),
        "height": round(item["height"], 1),
        "z_order": item["z_order"],
        "text_prefix": (item["text"] or "")[:20],
        "in_group": item["in_group"],
    }


def main() -> int:
    setup_console_encoding()
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true",
                    help="Force re-extraction even if user annotations exist")
    args = ap.parse_args()

    # --- Read existing annotations before re-extraction (merge mode) ---
    # --force: start fresh, clear all annotations
    # default: re-extract shapes, restore existing annotations into new md
    if args.force:
        existing_annos: dict = {}
        safe_print("[INFO] --force: re-extracting shapes, annotations will be cleared.")
    else:
        existing_annos = parse_user_annotations()
        if existing_annos:
            safe_print(f"[INFO] Found {len(existing_annos)} annotated shape(s) — "
                       f"will merge into new md.")
        else:
            safe_print("[INFO] No existing annotations — generating fresh md.")

    if not TEMPLATE_PATH.exists():
        write_json(OUT_JSON, {"status": "blocked", "reason": "template missing", "new_shapes": []})
        safe_print("[BLOCKED] Template file not found")
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_json(OUT_JSON, {"status": "blocked", "reason": str(e), "new_shapes": []})
        safe_print(f"[BLOCKED] {e}")
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    pres = app.Presentations.Open(str(TEMPLATE_PATH))

    try:
        blank = pres.Slides(14)
        std = pres.Slides(15)

        # Collect blank page shape keys for diffing
        blank_keys = set()
        for i in range(1, int(blank.Shapes.Count) + 1):
            s = shape_obj(14, blank.Shapes(i), i)
            key = (
                s["shape_type"],
                round(s["left"], 1),
                round(s["top"], 1),
                round(s["width"], 1),
                round(s["height"], 1),
                s["name"],
            )
            blank_keys.add(key)

        # Find shapes on page 15 that are NOT on page 14
        new_shapes = []
        for i in range(1, int(std.Shapes.Count) + 1):
            s = shape_obj(15, std.Shapes(i), i)
            key = (
                s["shape_type"],
                round(s["left"], 1),
                round(s["top"], 1),
                round(s["width"], 1),
                round(s["height"], 1),
                s["name"],
            )
            if key not in blank_keys:
                new_shapes.append(s)

        fp_items = [fingersafe_print(x) for x in new_shapes]

        # Write JSON artifacts (for agents / code)
        write_json(OUT_JSON, {
            "status": "ok",
            "generated_at": now_ts(),
            "template_slide": 15,
            "blank_slide": 14,
            "blank_shape_count": int(blank.Shapes.Count),
            "std_shape_count": int(std.Shapes.Count),
            "new_shapes": new_shapes,
        })
        write_json(OUT_FP, {
            "generated_at": now_ts(),
            "fingerprints": fp_items,
        })

        # Write human-readable MD — merge existing annotations if present
        md_lines = generate_shape_detail_md(new_shapes, existing_annos=existing_annos)
        write_md(OUT_MD, md_lines)

        safe_print(f"[OK] {len(new_shapes)} new shapes. "
              f"Wrote {OUT_JSON.name}, {OUT_FP.name}, {OUT_MD.name}")
        return 0
    finally:
        com_call(pres, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
