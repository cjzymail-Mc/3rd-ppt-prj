#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step3B: 克隆标准页并写入内容；写后回读校验。"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from ppt_pipeline_common import ROOT, TEMPLATE_PATH, com_call, com_get, now_ts, write_json, write_md

CONTENT_JSON = ROOT / "build_shape_content.json"
OUT_REPORT = ROOT / "build-ppt-report.md"
OUT_READBACK = ROOT / "post_write_readback.json"


def apply_shape(slide, item: dict) -> dict:
    name = item.get("shape_name", "")
    content = item.get("content", "")
    shp = com_call(slide.Shapes, "Item", name)
    if shp is None:
        return {"shape_name": name, "updated": False, "reason": "shape not found"}

    if bool(com_get(shp, "HasTextFrame", 0)):
        tf = com_get(shp, "TextFrame", None)
        tr = com_get(tf, "TextRange", None) if tf is not None else None
        if tr is not None:
            try:
                tr.Text = content
                rb_text = com_get(tr, "Text", "") or ""
                return {"shape_name": name, "updated": True, "mode": "text", "written_len": len(content), "readback_len": len(rb_text)}
            except Exception as e:
                return {"shape_name": name, "updated": False, "reason": str(e)}

    if bool(com_get(shp, "HasChart", False)):
        chart = com_get(shp, "Chart", None)
        try:
            wb = chart.ChartData.Workbook
            ws = wb.Worksheets(1)
            ws.Cells(1, 1).Value = "类别"
            ws.Cells(1, 2).Value = "值"

            lines = [x.strip() for x in (content or "").splitlines() if x.strip()]
            if not lines:
                lines = ["A:1", "B:2"]

            row = 2
            for line in lines[:8]:
                if ":" in line:
                    k, v = line.split(":", 1)
                else:
                    k, v = line, "1"
                ws.Cells(row, 1).Value = k.strip()
                try:
                    ws.Cells(row, 2).Value = float(v.strip())
                except Exception:
                    ws.Cells(row, 2).Value = 1
                row += 1

            chart.SetSourceData(ws.Range(f"A1:B{max(3, row-1)}"))
            ctype = int(com_get(chart, "ChartType", 0) or 0)
            return {"shape_name": name, "updated": True, "mode": "chart", "chart_type": ctype}
        except Exception as e:
            return {"shape_name": name, "updated": False, "reason": str(e)}

    return {"shape_name": name, "updated": False, "reason": "unsupported shape"}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default="1.0")
    args = ap.parse_args()

    out_ppt = ROOT / f"codex {args.version}.pptx"

    if not TEMPLATE_PATH.exists() or not CONTENT_JSON.exists():
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", "- 原因: 模板或build_shape_content.json缺失"])
        return 0

    items = json.loads(CONTENT_JSON.read_text(encoding="utf-8")).get("items", [])

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", f"- 原因: {e}"])
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    src = app.Presentations.Open(str(TEMPLATE_PATH))
    dst = app.Presentations.Add()

    results = []
    try:
        src.Slides(15).Copy()
        dst.Slides.Paste()
        slide = dst.Slides(1)

        for item in items:
            results.append(apply_shape(slide, item))

        dst.SaveAs(str(out_ppt))

        updated = sum(1 for r in results if r.get("updated"))
        write_json(OUT_READBACK, {"generated_at": now_ts(), "target": out_ppt.name, "results": results})
        write_md(
            OUT_REPORT,
            [
                "# Build PPT Report",
                "",
                "- 状态: ok",
                f"- 产物: {out_ppt.name}",
                f"- 更新shape: {updated}/{len(results)}",
                f"- 时间: {now_ts()}",
            ],
        )
        print(f"[OK] generated {out_ppt.name}, updated {updated}/{len(results)}")
        return 0
    finally:
        com_call(src, "Close")
        com_call(dst, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
