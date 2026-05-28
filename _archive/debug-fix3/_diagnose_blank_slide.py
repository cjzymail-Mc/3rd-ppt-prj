#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fix3 诊断：连接当前打开的 PPT，dump 最后一页全部 shape，看为什么是空白。

排查矩阵：
  - 总 slide 数 vs 期望（应该 +1）
  - 最后页 shape 数 + 名字列表（看是否 clone 成功）
  - 每个 shape 的 text 内容（看 _write_text 是否生效）
  - chart shape 是否被删（fix4 路线会先删模板 chart）
"""
from __future__ import annotations

import io
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import win32com.client

MSO_SHAPE_TYPE = {
    1: "AutoShape", 3: "Chart", 6: "Group", 7: "EmbeddedOLEObject",
    9: "Line", 13: "Picture", 14: "Placeholder",
    17: "TextBox", 19: "Table",
}


def _safe(fn, default=""):
    try:
        return fn()
    except Exception as e:
        return f"<ERR: {e}>"


def main():
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        print(f"PPT_NOT_ACTIVE: {e}")
        return 1

    presos = app.Presentations
    print(f"=== Open presentations: {presos.Count}")
    for i in range(1, presos.Count + 1):
        p = presos.Item(i)
        print(f"  [{i}] {p.Name}  slides={p.Slides.Count}")
    print()

    # 取 active presentation
    ppt = app.ActivePresentation
    print(f"=== Active: {ppt.Name}")
    total = ppt.Slides.Count
    print(f"=== Total slides: {total}")

    # 看最后 3 页
    start = max(1, total - 2)
    for idx in range(start, total + 1):
        slide = ppt.Slides.Item(idx)
        shapes = slide.Shapes
        n = shapes.Count
        print(f"\n========== Slide {idx} (shape count = {n}) ==========")
        if n == 0:
            print("  (空 slide — 没有任何 shape)")
            continue
        for j in range(1, n + 1):
            sh = shapes.Item(j)
            name = _safe(lambda: sh.Name)
            t = _safe(lambda: int(sh.Type), -1)
            tname = MSO_SHAPE_TYPE.get(t, f"?{t}")
            L = _safe(lambda: round(float(sh.Left), 1))
            T = _safe(lambda: round(float(sh.Top), 1))
            W = _safe(lambda: round(float(sh.Width), 1))
            H = _safe(lambda: round(float(sh.Height), 1))
            print(f"  [{j:02d}] name={name!r} type={t}({tname}) "
                  f"L={L} T={T} W={W} H={H}")

            # text
            has_tf = _safe(lambda: int(sh.HasTextFrame), 0)
            if has_tf:
                txt = _safe(lambda: sh.TextFrame.TextRange.Text, "")
                if isinstance(txt, str):
                    disp = txt.replace("\r", "[CR]").replace("\n", "[LF]")
                    if len(disp) > 200:
                        disp = disp[:200] + "..."
                    print(f"       TEXT={disp!r}")

            # chart
            has_ch = _safe(lambda: bool(sh.HasChart), False)
            if has_ch:
                ct = _safe(lambda: sh.Chart.ChartType, "")
                sc = _safe(lambda: sh.Chart.SeriesCollection().Count, 0)
                print(f"       CHART type={ct} series_count={sc}")
                try:
                    s1 = sh.Chart.SeriesCollection(1)
                    vals = list(s1.Values)
                    xs = list(s1.XValues)
                    print(f"       SERIES1 X={xs} Vals={vals}")
                except Exception as e:
                    print(f"       SERIES read err: {e}")

    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)
