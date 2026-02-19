#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step1: 对比模板14/15页，提取新增shape详情与指纹。"""

from __future__ import annotations

from pathlib import Path

from ppt_pipeline_common import ROOT, TEMPLATE_PATH, com_call, com_get, now_ts, write_json, write_md

OUT_MD = ROOT / "01-shape-detail.md"
OUT_JSON = ROOT / "shape_detail_com.json"
OUT_FP = ROOT / "shape_fingerprint_map.json"


def shape_obj(slide_index: int, shp, idx: int):
    name = com_get(shp, "Name", "") or f"shape_{slide_index}_{idx}"
    has_tf = bool(com_get(shp, "HasTextFrame", 0))
    text = ""
    font_name = ""
    font_size = 0.0
    if has_tf:
        tf = com_get(shp, "TextFrame", None)
        tr = com_get(tf, "TextRange", None) if tf is not None else None
        text = com_get(tr, "Text", "") or ""
        font = com_get(tr, "Font", None) if tr is not None else None
        font_name = com_get(font, "Name", "") if font is not None else ""
        font_size = float(com_get(font, "Size", 0.0) or 0.0) if font is not None else 0.0

    parent_group = com_get(shp, "ParentGroup", None)
    zorder = com_get(shp, "ZOrderPosition", 0)
    return {
        "slide_index": slide_index,
        "name": name,
        "left": float(com_get(shp, "Left", 0.0) or 0.0),
        "top": float(com_get(shp, "Top", 0.0) or 0.0),
        "width": float(com_get(shp, "Width", 0.0) or 0.0),
        "height": float(com_get(shp, "Height", 0.0) or 0.0),
        "shape_type": int(com_get(shp, "Type", 0) or 0),
        "text": text,
        "font_name": font_name,
        "font_size": font_size,
        "has_chart": bool(com_get(shp, "HasChart", False)),
        "in_group": parent_group is not None,
        "z_order": int(zorder or 0),
    }


def fingerprint(item: dict) -> dict:
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
    if not TEMPLATE_PATH.exists():
        write_md(OUT_MD, ["# 01-shape-detail", "", "- 状态: blocked", "- 原因: 模板文件不存在"])
        write_json(OUT_JSON, {"status": "blocked", "reason": "template missing", "new_shapes": []})
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_md(OUT_MD, ["# 01-shape-detail", "", "- 状态: blocked", f"- 原因: {e}"])
        write_json(OUT_JSON, {"status": "blocked", "reason": str(e), "new_shapes": []})
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    pres = app.Presentations.Open(str(TEMPLATE_PATH))

    try:
        blank = pres.Slides(14)
        std = pres.Slides(15)

        blank_keys = set()
        for i in range(1, int(blank.Shapes.Count) + 1):
            s = shape_obj(14, blank.Shapes(i), i)
            blank_keys.add((s["shape_type"], round(s["left"], 1), round(s["top"], 1), round(s["width"], 1), round(s["height"], 1), s["name"]))

        new_shapes = []
        for i in range(1, int(std.Shapes.Count) + 1):
            s = shape_obj(15, std.Shapes(i), i)
            key = (s["shape_type"], round(s["left"], 1), round(s["top"], 1), round(s["width"], 1), round(s["height"], 1), s["name"])
            if key not in blank_keys:
                new_shapes.append(s)

        fp_items = [fingerprint(x) for x in new_shapes]

        write_json(OUT_JSON, {"status": "ok", "generated_at": now_ts(), "new_shapes": new_shapes})
        write_json(OUT_FP, {"generated_at": now_ts(), "fingerprints": fp_items})

        lines = [
            "# 01-shape-detail",
            "",
            f"- 状态: ok",
            f"- 时间: {now_ts()}",
            f"- 新增shape数量: {len(new_shapes)}",
            "",
        ]
        for i, s in enumerate(new_shapes, 1):
            lines += [
                f"## {i}. {s['name']}",
                f"- type: {s['shape_type']}  has_chart: {s['has_chart']}  in_group: {s['in_group']}",
                f"- left/top: {s['left']:.1f}/{s['top']:.1f}",
                f"- width/height: {s['width']:.1f}/{s['height']:.1f}",
                f"- font: {s['font_name']} {s['font_size']}",
                f"- text: {(s['text'] or '')[:120]}",
                "",
            ]
        write_md(OUT_MD, lines)
        print(f"[OK] wrote {OUT_MD.name}, {OUT_JSON.name}, {OUT_FP.name}")
        return 0
    finally:
        com_call(pres, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
