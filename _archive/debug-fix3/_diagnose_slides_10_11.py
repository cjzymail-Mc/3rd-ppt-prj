#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读当前 PPT 第 10、11 页全部 shape + chart 数据。"""
from __future__ import annotations
import io, sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import win32com.client

MSO = {1:"AutoShape",3:"Chart",6:"Group",7:"OLE",13:"Picture",17:"TextBox"}

def _s(fn, d=""):
    try: return fn()
    except Exception as e: return f"<ERR:{e}>"

def main():
    app = win32com.client.GetActiveObject("PowerPoint.Application")
    ppt = app.ActivePresentation
    print(f"=== {ppt.Name}  total slides={ppt.Slides.Count}\n")

    for idx in (10, 11):
        slide = ppt.Slides.Item(idx)
        n = slide.Shapes.Count
        print(f"\n========== Slide {idx}  ({n} shapes) ==========")
        for j in range(1, n + 1):
            sh = slide.Shapes.Item(j)
            name = _s(lambda: sh.Name)
            t = _s(lambda: int(sh.Type), -1)
            L = _s(lambda: round(float(sh.Left), 1))
            T = _s(lambda: round(float(sh.Top), 1))
            W = _s(lambda: round(float(sh.Width), 1))
            H = _s(lambda: round(float(sh.Height), 1))
            print(f"[{j:02d}] {name!r} type={t}({MSO.get(t,'?')}) L={L} T={T} W={W} H={H}")
            if _s(lambda: int(sh.HasTextFrame), 0):
                txt = _s(lambda: sh.TextFrame.TextRange.Text, "")
                if isinstance(txt, str) and txt.strip():
                    disp = txt.replace("\r","\\r").replace("\n","\\n")
                    if len(disp) > 200: disp = disp[:200] + "..."
                    print(f"     TEXT={disp!r}")
            if _s(lambda: bool(sh.HasChart), False):
                ct = _s(lambda: sh.Chart.ChartType, "")
                il = _s(lambda: sh.Chart.ChartData.IsLinked, "")
                print(f"     CHART_TYPE={ct}  IS_LINKED={il}")
                try:
                    sc_count = sh.Chart.SeriesCollection().Count
                    print(f"     SERIES_COUNT={sc_count}")
                    for si in range(1, min(sc_count, 3) + 1):
                        s1 = sh.Chart.SeriesCollection(si)
                        sname = _s(lambda: s1.Name, "")
                        vals = list(s1.Values)
                        xs = list(s1.XValues)
                        xs_short = xs if len(str(xs)) < 200 else (xs[:5] + ['...'])
                        print(f"     series{si}: name={sname!r}")
                        print(f"               X={xs_short}")
                        print(f"               V={vals}")
                except Exception as e:
                    print(f"     CHART read err: {e}")
            print()

if __name__ == "__main__":
    main()
