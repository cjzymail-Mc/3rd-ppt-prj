"""
基于 PowerPoint COM 的模板分析脚本。

要求与约束：
1) 必须使用 win32com.client.Dispatch("PowerPoint.Application")。
2) PowerPoint 后台运行：Visible=False。
3) 同时打开空白模板与标准模板，逐页分析并比较。
4) 递归遍历组内 shape。
5) 输出 shape_detail_com.json，供后续按 shape.Name 精准替换。

⚠️ 关键约定：
- 为了后续 build 脚本能“精准替换且不破坏样式”，建议先在标准模板中为关键目标 shape 手动命名
  （如：Title、MainChart、Q1_Bar、SummaryText 等）。
- 本脚本会在 meta.notes 中再次写入该约定提示。
"""

from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional

import win32com.client

try:
    from src.Class_030 import font_exists_in_registry
except Exception:
    def font_exists_in_registry(_font_name):
        return False


def open_powerpoint(visible: bool = False):
    """启动 PowerPoint COM 应用。"""
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = visible
    app.DisplayAlerts = 0
    return app


def _safe_get(callable_obj, default=None):
    try:
        return callable_obj()
    except Exception:
        return default


def _extract_text_info(shape) -> Dict[str, Any]:
    info: Dict[str, Any] = {
        "text": "",
        "font": {},
        "paragraph": {},
    }

    has_text_frame = _safe_get(lambda: bool(shape.HasTextFrame), False)
    if not has_text_frame:
        return info

    has_text = _safe_get(lambda: bool(shape.TextFrame.HasText), False)
    if not has_text:
        return info

    tr = shape.TextFrame.TextRange
    info["text"] = _safe_get(lambda: tr.Text, "")

    font = tr.Font
    font_name = _safe_get(lambda: font.Name, None)
    info["font"] = {
        "name": font_name,
        "name_far_east": _safe_get(lambda: font.NameFarEast, None),
        "size": _safe_get(lambda: font.Size, None),
        "bold": _safe_get(lambda: font.Bold, None),
        "italic": _safe_get(lambda: font.Italic, None),
        "color_rgb": _safe_get(lambda: font.Color.RGB, None),
        "font_registered": font_exists_in_registry(font_name),
    }

    pf = tr.ParagraphFormat
    info["paragraph"] = {
        "alignment": _safe_get(lambda: pf.Alignment, None),
        "line_spacing": _safe_get(lambda: pf.LineRuleWithin, None),
        "space_within": _safe_get(lambda: pf.SpaceWithin, None),
        "space_before": _safe_get(lambda: pf.SpaceBefore, None),
        "space_after": _safe_get(lambda: pf.SpaceAfter, None),
    }
    return info


def _extract_chart_info(shape) -> Dict[str, Any]:
    info: Dict[str, Any] = {
        "has_chart": False,
        "chart_type": None,
        "has_data_table": None,
        "series_count": None,
        "series": [],
    }

    has_chart = _safe_get(lambda: bool(shape.HasChart), False)
    if not has_chart:
        return info

    chart = shape.Chart
    info["has_chart"] = True
    info["chart_type"] = _safe_get(lambda: chart.ChartType, None)
    info["has_data_table"] = _safe_get(lambda: bool(chart.HasDataTable), None)
    series_count = _safe_get(lambda: chart.SeriesCollection().Count, None)
    info["series_count"] = series_count

    if isinstance(series_count, int):
        for idx in range(1, series_count + 1):
            s = chart.SeriesCollection(idx)
            info["series"].append(
                {
                    "index": idx,
                    "name": _safe_get(lambda: s.Name, None),
                    "points_count": _safe_get(lambda: s.Points().Count, None),
                }
            )

    return info


def _shape_identity(shape, slide_index: int, path_prefix: str = "") -> Dict[str, Any]:
    shape_name = _safe_get(lambda: shape.Name, "") or ""
    unique_name = f"{path_prefix}/{shape_name}" if path_prefix else shape_name

    base = {
        "slide_index": slide_index,
        "shape_name": unique_name,
        "shape_name_raw": shape_name,
        "left": _safe_get(lambda: float(shape.Left), None),
        "top": _safe_get(lambda: float(shape.Top), None),
        "width": _safe_get(lambda: float(shape.Width), None),
        "height": _safe_get(lambda: float(shape.Height), None),
        "type": _safe_get(lambda: int(shape.Type), None),
        "is_group": _safe_get(lambda: int(shape.Type) == 6, False),
        "group_path": path_prefix,
    }

    base["text_info"] = _extract_text_info(shape)
    base["chart_info"] = _extract_chart_info(shape)
    return base


def _walk_shape(shp, slide_index: int, path_prefix: str, rows: List[Dict[str, Any]]):
    row = _shape_identity(shp, slide_index=slide_index, path_prefix=path_prefix)
    rows.append(row)

    if row["is_group"]:
        group_name = row["shape_name_raw"] or "Group"
        child_prefix = f"{path_prefix}/{group_name}" if path_prefix else group_name
        group_items = _safe_get(lambda: shp.GroupItems, None)
        if group_items is not None:
            for i in range(1, group_items.Count + 1):
                child = group_items.Item(i)
                _walk_shape(child, slide_index=slide_index, path_prefix=child_prefix, rows=rows)


def _flatten_shapes(slide, slide_index: int, path_prefix: str = "") -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for shp in slide.Shapes:
        _walk_shape(shp, slide_index=slide_index, path_prefix=path_prefix, rows=rows)
    return rows


def _shape_signature(row: Dict[str, Any], tolerance: float = 0.3) -> tuple:
    def q(v):
        if v is None:
            return None
        return round(float(v) / tolerance) * tolerance

    return (
        row.get("type"),
        q(row.get("left")),
        q(row.get("top")),
        q(row.get("width")),
        q(row.get("height")),
    )


def analyze_templates(blank_path: str, standard_path: str, output_path: str = "shape_detail_com.json") -> Dict[str, Any]:
    ppt = None
    prs_blank = None
    prs_std = None

    abs_blank = os.path.abspath(blank_path)
    abs_std = os.path.abspath(standard_path)
    abs_output = os.path.abspath(output_path)

    try:
        ppt = open_powerpoint(visible=False)
        prs_blank = ppt.Presentations.Open(abs_blank, WithWindow=False)
        prs_std = ppt.Presentations.Open(abs_std, WithWindow=False)

        blank_map: Dict[int, List[Dict[str, Any]]] = {}
        std_map: Dict[int, List[Dict[str, Any]]] = {}

        for sidx in range(1, prs_blank.Slides.Count + 1):
            blank_map[sidx] = _flatten_shapes(prs_blank.Slides(sidx), slide_index=sidx)

        for sidx in range(1, prs_std.Slides.Count + 1):
            std_map[sidx] = _flatten_shapes(prs_std.Slides(sidx), slide_index=sidx)

        changed_shapes: List[Dict[str, Any]] = []

        for sidx, std_rows in std_map.items():
            blank_rows = blank_map.get(sidx, [])
            blank_sigs = {_shape_signature(r) for r in blank_rows}

            for row in std_rows:
                sig = _shape_signature(row)
                if sig not in blank_sigs:
                    changed_shapes.append(row)

        result = {
            "meta": {
                "blank_template": abs_blank,
                "standard_template": abs_std,
                "slide_count_blank": len(blank_map),
                "slide_count_standard": len(std_map),
                "changed_shape_count": len(changed_shapes),
                "notes": [
                    "后续替换逻辑基于 shape.Name 精准定位。",
                    "请在标准模板中先手动命名关键 shape（如 Title、MainChart、SummaryText）。",
                    "若未命名，将影响 build 脚本稳定定位，建议先做模板命名治理。",
                ],
            },
            "changed_shapes": changed_shapes,
        }

        with open(abs_output, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)

        return result

    finally:
        try:
            if prs_blank is not None:
                prs_blank.Close()
        except Exception:
            pass
        try:
            if prs_std is not None:
                prs_std.Close()
        except Exception:
            pass
        try:
            if ppt is not None:
                ppt.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    output = analyze_templates(
        blank_path="0-空白 ppt 模板.pptx",
        standard_path="1-标准 ppt 模板.pptx",
        output_path="shape_detail_com.json",
    )
    print(f"分析完成，输出 changed_shapes={output['meta']['changed_shape_count']}")
