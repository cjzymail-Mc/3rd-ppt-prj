#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""read_selected_shape.py — 读取当前鼠标选中的 PPT shape 的完整信息.

用法：
    1. 在 PowerPoint 里选中一个或多个 shape
    2. 运行：python skills/read_selected_shape.py

用途：
    - 调试 per-template 适配时查 shape 名称（中/英文内部名）
    - 获取微调所需坐标（Left/Top/Width/Height）
    - 查看 shape 文本内容（含段落/字体/颜色）
    - 图表类 shape 打印图表类型、系列数据、IsLinked 状态
    - 图片类 shape 打印原生尺寸 + 裁剪框

输出到 stdout（纯文本），方便直接贴回对话。
"""
from __future__ import annotations

import sys
import win32com.client

# ---- PowerPoint MsoShapeType 常量（仅列常见的） -----------------------------
MSO_SHAPE_TYPE = {
    1: "AutoShape", 2: "Callout", 3: "Chart", 4: "Comment",
    5: "Freeform", 6: "Group", 7: "EmbeddedOLEObject", 8: "FormControl",
    9: "Line", 10: "LinkedOLEObject", 11: "LinkedPicture", 12: "OLEControlObject",
    13: "Picture", 14: "Placeholder", 15: "TextEffect", 16: "Media",
    17: "TextBox", 18: "ScriptAnchor", 19: "Table", 20: "Canvas",
    21: "Diagram", 22: "Ink", 23: "InkComment", 24: "IgxGraphic",
}
SELECTION_TYPE = {
    0: "None", 1: "Slides", 2: "Shapes", 3: "Text"
}


def _safe(fn, default=None):
    """Call a COM accessor, swallow any exception."""
    try:
        v = fn()
        return v
    except Exception:
        return default


def _rgb(color_long):
    """Convert COM BGR long int → 'R,G,B' string."""
    try:
        c = int(color_long)
        return f"{c & 0xFF},{(c >> 8) & 0xFF},{(c >> 16) & 0xFF}"
    except Exception:
        return str(color_long)


def _print_text_details(sh):
    """Print per-paragraph/run text + font details."""
    try:
        tf = sh.TextFrame
        if not int(_safe(lambda: tf.HasText, 0)):
            print("  TEXT=（空）")
            return
        tr = tf.TextRange
        full = str(_safe(lambda: tr.Text, ""))
        print(f"  TEXT_FULL={full.replace(chr(13), '[CR]').replace(chr(10), '[LF]')[:500]}")
        print(f"  TEXT_LEN={len(full)}")

        # paragraph-level
        paragraphs = _safe(lambda: tr.Paragraphs(), None)
        pcount = int(_safe(lambda: paragraphs.Count, 0)) if paragraphs else 0
        print(f"  PARAGRAPHS={pcount}")
        for i in range(1, min(pcount, 20) + 1):
            p = tr.Paragraphs(i, 1)
            ptext = str(_safe(lambda: p.Text, "")).rstrip("\r\n")
            fname = _safe(lambda: p.Font.Name, "")
            fsize = _safe(lambda: p.Font.Size, "")
            fbold = _safe(lambda: p.Font.Bold, "")
            fcolor = _safe(lambda: p.Font.Color.RGB, "")
            print(f"    p{i}: text={ptext!r} font={fname} size={fsize} "
                  f"bold={fbold} color=RGB({_rgb(fcolor) if fcolor != '' else '-'})")
    except Exception as e:
        print(f"  TEXT_READ_ERROR={e}")


def _print_chart_details(sh):
    """Print chart type, series data, linkage status."""
    try:
        chart = sh.Chart
        ctype = _safe(lambda: chart.ChartType, "")
        print(f"  CHART_TYPE={ctype}")
        is_linked = _safe(lambda: chart.ChartData.IsLinked, "")
        print(f"  CHART_IS_LINKED={is_linked}")
        sc = _safe(lambda: chart.SeriesCollection(), None)
        scount = int(_safe(lambda: sc.Count, 0)) if sc else 0
        print(f"  SERIES_COUNT={scount}")
        for si in range(1, scount + 1):
            s = chart.SeriesCollection(si)
            name = _safe(lambda: s.Name, "")
            try:
                vals = list(s.Values)
            except Exception:
                vals = []
            try:
                xvals = list(s.XValues)
            except Exception:
                xvals = []
            print(f"    series{si}: name={name!r} values={vals} xvalues={xvals}")
    except Exception as e:
        print(f"  CHART_READ_ERROR={e}")


def _print_picture_details(sh):
    """Print picture format details (cropping, original size)."""
    try:
        pf = sh.PictureFormat
        for attr in ("CropLeft", "CropTop", "CropRight", "CropBottom",
                     "Brightness", "Contrast"):
            v = _safe(lambda a=attr: getattr(pf, a), "")
            print(f"  PICTURE_{attr}={v}")
    except Exception as e:
        print(f"  PICTURE_READ_ERROR={e}")


def _print_fill_line(sh):
    """Print fill + line info (useful for Rectangle-like shapes)."""
    try:
        f = sh.Fill
        vis = _safe(lambda: f.Visible, "")
        ftype = _safe(lambda: f.Type, "")
        fcolor = _safe(lambda: f.ForeColor.RGB, "")
        transp = _safe(lambda: f.Transparency, "")
        print(f"  FILL: visible={vis} type={ftype} "
              f"color=RGB({_rgb(fcolor) if fcolor != '' else '-'}) transparency={transp}")
    except Exception:
        pass
    try:
        ln = sh.Line
        lvis = _safe(lambda: ln.Visible, "")
        lweight = _safe(lambda: ln.Weight, "")
        lcolor = _safe(lambda: ln.ForeColor.RGB, "")
        print(f"  LINE: visible={lvis} weight={lweight} "
              f"color=RGB({_rgb(lcolor) if lcolor != '' else '-'})")
    except Exception:
        pass


def _print_shape(sh, idx):
    print(f"========== SHAPE {idx} ==========")
    name = _safe(lambda: sh.Name, "")
    typ = _safe(lambda: int(sh.Type), -1)
    typ_name = MSO_SHAPE_TYPE.get(typ, f"Unknown({typ})")
    print(f"  NAME={name!r}   （COM 内部名，和'选择窗格'的本地化显示可能不同）")
    print(f"  TYPE={typ}  ({typ_name})")
    print(f"  SHAPE_ID={_safe(lambda: sh.Id, '')}")
    print(f"  POSITION: Left={_safe(lambda: sh.Left, '')} "
          f"Top={_safe(lambda: sh.Top, '')} "
          f"Width={_safe(lambda: sh.Width, '')} "
          f"Height={_safe(lambda: sh.Height, '')}")
    print(f"  ROTATION={_safe(lambda: sh.Rotation, '')}")
    print(f"  Z_ORDER={_safe(lambda: sh.ZOrderPosition, '')}")
    print(f"  VISIBLE={_safe(lambda: sh.Visible, '')}")
    print(f"  HAS_TEXT_FRAME={_safe(lambda: int(sh.HasTextFrame), 0)}")
    print(f"  HAS_CHART={_safe(lambda: 1 if bool(sh.HasChart) else 0, 0)}")
    print(f"  HAS_TABLE={_safe(lambda: 1 if bool(sh.HasTable) else 0, 0)}")
    print(f"  AUTO_SHAPE_TYPE={_safe(lambda: sh.AutoShapeType, '')}")

    if int(_safe(lambda: sh.HasTextFrame, 0)):
        _print_text_details(sh)

    if _safe(lambda: bool(sh.HasChart), False):
        _print_chart_details(sh)

    if typ in (11, 13):  # Picture / LinkedPicture
        _print_picture_details(sh)

    if typ in (1, 17):  # AutoShape / TextBox
        _print_fill_line(sh)

    print()


def main():
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        print(f"PPT_NOT_ACTIVE: {e}")
        return 1

    try:
        sel = app.ActiveWindow.Selection
        stype = int(sel.Type)
    except Exception as e:
        print(f"NO_SELECTION: {e}")
        return 1

    print(f"SELECTION_TYPE={stype} ({SELECTION_TYPE.get(stype, '?')})")
    if stype not in (2, 3):
        print("请先在 PPT 里选中一个或多个 shape 再运行。")
        return 0

    try:
        sr = sel.ShapeRange
        cnt = int(sr.Count)
    except Exception as e:
        print(f"NO_SHAPERANGE: {e}")
        return 1

    print(f"SHAPE_COUNT={cnt}")
    try:
        slide = app.ActiveWindow.View.Slide
        print(f"SLIDE_INDEX={_safe(lambda: slide.SlideIndex, '?')}")
    except Exception:
        pass
    print()

    for i in range(1, cnt + 1):
        _print_shape(sr.Item(i), i)
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
