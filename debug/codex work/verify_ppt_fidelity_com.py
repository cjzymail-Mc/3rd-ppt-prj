"""PPT 保真度对比脚本（COM）。"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List

import win32com.client

STD_PPT = "1-标准 ppt 模板.pptx"
GEN_PPT = "gemini-jules.pptx"
OUT_JSON = "fidelity_diff_report.json"


def open_powerpoint(visible: bool = False):
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = visible
    app.DisplayAlerts = 0
    return app


def _safe_get(func, default=None):
    try:
        return func()
    except Exception:
        return default


def _shape_dict(shp) -> Dict[str, Any]:
    row = {
        "name": _safe_get(lambda: shp.Name, ""),
        "left": _safe_get(lambda: float(shp.Left), None),
        "top": _safe_get(lambda: float(shp.Top), None),
        "width": _safe_get(lambda: float(shp.Width), None),
        "height": _safe_get(lambda: float(shp.Height), None),
        "type": _safe_get(lambda: int(shp.Type), None),
        "font": {},
        "chart": {},
    }

    if _safe_get(lambda: bool(shp.HasTextFrame), False) and _safe_get(lambda: bool(shp.TextFrame.HasText), False):
        tr = shp.TextFrame.TextRange
        row["font"] = {
            "name": _safe_get(lambda: tr.Font.Name, None),
            "size": _safe_get(lambda: tr.Font.Size, None),
            "bold": _safe_get(lambda: tr.Font.Bold, None),
            "color_rgb": _safe_get(lambda: tr.Font.Color.RGB, None),
        }

    if _safe_get(lambda: bool(shp.HasChart), False):
        chart = shp.Chart
        row["chart"] = {
            "chart_type": _safe_get(lambda: chart.ChartType, None),
            "series_count": _safe_get(lambda: chart.SeriesCollection().Count, None),
        }

    return row


def _compare_value(field: str, a, b, diffs: List[Dict[str, Any]], tol: float = 0.3):
    if a is None and b is None:
        return
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        if abs(float(a) - float(b)) > tol:
            diffs.append({"field": field, "expected": a, "actual": b})
    elif a != b:
        diffs.append({"field": field, "expected": a, "actual": b})


def verify_ppt_fidelity(std_ppt: str = STD_PPT, gen_ppt: str = GEN_PPT, output_json: str = OUT_JSON):
    ppt = None
    prs_std = None
    prs_gen = None

    abs_std = os.path.abspath(std_ppt)
    abs_gen = os.path.abspath(gen_ppt)
    abs_out = os.path.abspath(output_json)

    report: Dict[str, Any] = {
        "meta": {
            "standard": abs_std,
            "generated": abs_gen,
        },
        "slides": [],
        "summary": {},
    }

    try:
        if not os.path.exists(abs_gen):
            raise FileNotFoundError(f"未找到生成文件：{abs_gen}")

        ppt = open_powerpoint(visible=False)
        prs_std = ppt.Presentations.Open(abs_std, WithWindow=False)
        prs_gen = ppt.Presentations.Open(abs_gen, WithWindow=False)

        slide_count = min(prs_std.Slides.Count, prs_gen.Slides.Count)
        total_diffs = 0

        for sidx in range(1, slide_count + 1):
            s_std = prs_std.Slides(sidx)
            s_gen = prs_gen.Slides(sidx)

            std_shapes = {_shape_dict(shp)["name"]: _shape_dict(shp) for shp in s_std.Shapes}
            gen_shapes = {_shape_dict(shp)["name"]: _shape_dict(shp) for shp in s_gen.Shapes}

            all_names = sorted(set(std_shapes.keys()) | set(gen_shapes.keys()))
            slide_diffs = []

            for name in all_names:
                a = std_shapes.get(name)
                b = gen_shapes.get(name)

                if a is None or b is None:
                    slide_diffs.append(
                        {
                            "shape": name,
                            "field": "existence",
                            "expected": "exists_in_both",
                            "actual": "missing_in_generated" if b is None else "extra_in_generated",
                        }
                    )
                    continue

                diffs = []
                for f in ["left", "top", "width", "height", "type"]:
                    _compare_value(f, a.get(f), b.get(f), diffs)

                for f in ["name", "size", "bold", "color_rgb"]:
                    _compare_value(f"font.{f}", a.get("font", {}).get(f), b.get("font", {}).get(f), diffs)

                for f in ["chart_type", "series_count"]:
                    _compare_value(f"chart.{f}", a.get("chart", {}).get(f), b.get("chart", {}).get(f), diffs)

                if diffs:
                    slide_diffs.append({"shape": name, "diffs": diffs})

            total_diffs += len(slide_diffs)
            report["slides"].append({"slide_index": sidx, "diff_count": len(slide_diffs), "diffs": slide_diffs})

        report["summary"] = {
            "slide_count_compared": slide_count,
            "total_shape_diff_items": total_diffs,
            "result": "PASS" if total_diffs == 0 else "HAS_DIFF",
        }

        with open(abs_out, "w", encoding="utf-8") as f:
            json.dump(report, f, ensure_ascii=False, indent=2)

        return report

    finally:
        try:
            if prs_std is not None:
                prs_std.Close()
        except Exception:
            pass
        try:
            if prs_gen is not None:
                prs_gen.Close()
        except Exception:
            pass
        try:
            if ppt is not None:
                ppt.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    r = verify_ppt_fidelity()
    print(f"对比完成：{r['summary']}")
