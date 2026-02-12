#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step4: 严格diff（visual + readability + semantic）。"""

from __future__ import annotations

import argparse
import json
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List

from ppt_pipeline_common import ROOT, TEMPLATE_PATH, com_call, com_get, now_ts, safe_text, write_json, write_md

FIX_MD = ROOT / "fix-ppt.md"
SEM_MD = ROOT / "diff_semantic_report.md"

CONTENT_JSON = ROOT / "build_shape_content.json"


def sim(a: str, b: str) -> float:
    return SequenceMatcher(None, a or "", b or "").ratio()


def shape_props(shp) -> Dict:
    has_tf = bool(com_get(shp, "HasTextFrame", 0))
    text = ""
    font_name = ""
    font_size = 0.0
    font_color = None
    if has_tf:
        tf = com_get(shp, "TextFrame", None)
        tr = com_get(tf, "TextRange", None) if tf is not None else None
        text = safe_text(com_get(tr, "Text", "")) if tr is not None else ""
        font = com_get(tr, "Font", None) if tr is not None else None
        font_name = safe_text(com_get(font, "Name", "")) if font is not None else ""
        font_size = float(com_get(font, "Size", 0.0) or 0.0) if font is not None else 0.0
        fc = com_get(com_get(font, "Color", None), "RGB", None) if font is not None else None
        font_color = fc

    return {
        "name": safe_text(com_get(shp, "Name", "")),
        "left": float(com_get(shp, "Left", 0.0) or 0.0),
        "top": float(com_get(shp, "Top", 0.0) or 0.0),
        "width": float(com_get(shp, "Width", 0.0) or 0.0),
        "height": float(com_get(shp, "Height", 0.0) or 0.0),
        "shape_type": int(com_get(shp, "Type", 0) or 0),
        "text": text,
        "font_name": font_name,
        "font_size": font_size,
        "font_color": font_color,
        "has_chart": bool(com_get(shp, "HasChart", False)),
        "chart_type": int(com_get(com_get(shp, "Chart", None), "ChartType", 0) or 0),
    }


def visual_score(a: Dict, b: Dict) -> float:
    score = 0.0
    w = 0.0

    def add(v, wt):
        nonlocal score, w
        score += v * wt
        w += wt

    add(max(0.0, 1 - abs(a["left"] - b["left"]) / 20), 10)
    add(max(0.0, 1 - abs(a["top"] - b["top"]) / 20), 10)
    add(max(0.0, 1 - abs(a["width"] - b["width"]) / 20), 8)
    add(max(0.0, 1 - abs(a["height"] - b["height"]) / 20), 8)
    add(1.0 if a["shape_type"] == b["shape_type"] else 0.0, 8)
    if a["has_chart"] or b["has_chart"]:
        add(1.0 if a["chart_type"] == b["chart_type"] else 0.0, 16)

    return (score / w) * 100 if w else 0.0


def readability_score(a: Dict, b: Dict) -> float:
    if not a["text"] and not b["text"]:
        return 100.0
    text_sim = sim(a["text"], b["text"]) * 100
    len_a = max(1, len(a["text"]))
    len_b = max(1, len(b["text"]))
    len_ratio = min(len_a, len_b) / max(len_a, len_b)
    line_a = max(1, len(a["text"].splitlines()))
    line_b = max(1, len(b["text"].splitlines()))
    line_ratio = min(line_a, line_b) / max(line_a, line_b)
    return (text_sim * 0.6) + (len_ratio * 100 * 0.25) + (line_ratio * 100 * 0.15)


def semantic_report(target_slide) -> Dict:
    required = ["样本", "建议", "反馈"]
    text_blob = []
    for i in range(1, int(target_slide.Shapes.Count) + 1):
        shp = target_slide.Shapes(i)
        if bool(com_get(shp, "HasTextFrame", 0)):
            tf = com_get(shp, "TextFrame", None)
            tr = com_get(tf, "TextRange", None) if tf is not None else None
            text_blob.append(safe_text(com_get(tr, "Text", "")))
    all_text = "\n".join(text_blob)
    hits = {k: (k in all_text) for k in required}
    coverage = sum(1 for v in hits.values() if v) / len(required) * 100
    return {"coverage": coverage, "hits": hits}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--target", default="codex 1.0.pptx")
    args = ap.parse_args()
    target = ROOT / args.target

    if not TEMPLATE_PATH.exists() or not target.exists():
        write_md(FIX_MD, ["# PPT Diff & Fix Report", "", "- 状态: blocked", "- 原因: 模板或目标PPT缺失"])
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_md(FIX_MD, ["# PPT Diff & Fix Report", "", "- 状态: blocked", f"- 原因: {e}"])
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    p1 = app.Presentations.Open(str(TEMPLATE_PATH))
    p2 = app.Presentations.Open(str(target))

    try:
        s1 = p1.Slides(15)
        s2 = p2.Slides(1)

        n = min(int(s1.Shapes.Count), int(s2.Shapes.Count))
        rows = []
        visual_scores = []
        read_scores = []
        fails = []

        for i in range(1, n + 1):
            a = shape_props(s1.Shapes(i))
            b = shape_props(s2.Shapes(i))
            vs = visual_score(a, b)
            rs = readability_score(a, b)
            visual_scores.append(vs)
            read_scores.append(rs)
            if vs < 98 or rs < 95:
                fails.append((i, a["name"], vs, rs))

            rows.append(
                f"|{i}|{a['name']}|{b['name']}|{vs:.2f}|{rs:.2f}|{len(a['text'])}/{len(b['text'])}|"
            )

        sem = semantic_report(s2)
        visual_avg = sum(visual_scores) / len(visual_scores) if visual_scores else 0.0
        read_avg = sum(read_scores) / len(read_scores) if read_scores else 0.0
        status = "ok" if visual_avg >= 98 and read_avg >= 95 and sem["coverage"] >= 100 and not fails else "fail"

        lines = [
            "# PPT Diff & Fix Report",
            "",
            f"- 状态: {status}",
            f"- visual_score: {visual_avg:.2f}%",
            f"- readability_score: {read_avg:.2f}%",
            f"- semantic_coverage: {sem['coverage']:.2f}%",
            f"- template_shapes: {int(s1.Shapes.Count)}",
            f"- target_shapes: {int(s2.Shapes.Count)}",
            f"- 时间: {now_ts()}",
            "",
            "## Shape对比",
            "|#|template|target|visual|readability|text_len|",
            "|---|---|---|---|---|---|",
        ] + rows

        if fails:
            lines += ["", "## 差异与修复建议"]
            for i, nm, vs, rs in fails[:40]:
                lines.append(f"- shape[{i}] {nm}: visual={vs:.2f}, readability={rs:.2f} -> 调整03-build_shape对应函数prompt/预算")

        write_md(FIX_MD, lines)
        write_md(SEM_MD, [
            "# diff_semantic_report",
            "",
            f"- coverage: {sem['coverage']:.2f}%",
            f"- hits: {sem['hits']}",
        ])
        write_json(ROOT / "diff_result.json", {
            "status": status,
            "visual_score": visual_avg,
            "readability_score": read_avg,
            "semantic_coverage": sem["coverage"],
            "fails": fails,
            "generated_at": now_ts(),
        })

        print(f"[OK] diff done status={status} visual={visual_avg:.2f} read={read_avg:.2f}")
        return 0 if status == "ok" else 1
    finally:
        com_call(p1, "Close")
        com_call(p2, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
