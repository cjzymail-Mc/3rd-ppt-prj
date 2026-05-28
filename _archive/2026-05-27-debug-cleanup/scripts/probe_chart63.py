"""Probe Chart 63 on slide 13: categories / series / values.

给 /developer 用于决定 _write_chart 注入策略：是替换全部 X/Y 还是只替换某列。
桥接 ActivePresentation，不动用户文件。
"""
from __future__ import annotations

import json
import os
import sys

import win32com.client

SLIDE_IDX = 13
TARGET_SHAPE = "Chart 63"
OUT_JSON = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "inspect-apparel-p1213", "chart63_data.json")
)


def safe(val):
    try:
        json.dumps(val)
        return val
    except Exception:
        return str(val)


def main() -> int:
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as exc:
        print(f"FAIL: 没有运行中的 PowerPoint: {exc}")
        return 1

    pres = ppt.ActivePresentation
    slide = pres.Slides(SLIDE_IDX)
    chart_shape = None
    for sh in slide.Shapes:
        if sh.Name == TARGET_SHAPE:
            chart_shape = sh
            break
    if chart_shape is None:
        print(f"FAIL: slide {SLIDE_IDX} 找不到 '{TARGET_SHAPE}'")
        return 2

    chart = chart_shape.Chart
    report = {
        "shape_name": chart_shape.Name,
        "L": chart_shape.Left,
        "T": chart_shape.Top,
        "W": chart_shape.Width,
        "H": chart_shape.Height,
        "chart_type": chart.ChartType,
        "has_title": bool(chart.HasTitle),
        "title": chart.ChartTitle.Text if chart.HasTitle else None,
        "series": [],
    }

    # 系列遍历
    sc = chart.SeriesCollection()
    for i in range(1, sc.Count + 1):
        s = sc.Item(i)
        try:
            xvals = list(s.XValues) if s.XValues else []
        except Exception:
            xvals = None
        try:
            yvals = list(s.Values) if s.Values else []
        except Exception:
            yvals = None
        try:
            sname = s.Name
        except Exception:
            sname = None
        report["series"].append({
            "index": i,
            "name": safe(sname),
            "x_values": [safe(v) for v in (xvals or [])],
            "y_values": [safe(v) for v in (yvals or [])],
        })

    # ChartData workbook（如果可读）
    try:
        cd = chart.ChartData
        report["chart_data_activated_at_dump"] = False
        # 不强制 Activate，避免触发 Excel 跳出
    except Exception as exc:
        report["chart_data_error"] = str(exc)

    os.makedirs(os.path.dirname(OUT_JSON), exist_ok=True)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=2)

    print(f"[probe] saved -> {OUT_JSON}")
    print(f"[probe] chart_type={report['chart_type']}, series_count={len(report['series'])}")
    for s in report["series"]:
        print(f"  Series#{s['index']} name={s['name']!r}  x={s['x_values']}  y={s['y_values']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
