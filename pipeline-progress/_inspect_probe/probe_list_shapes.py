#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""快速列出模板所有 slide 的 shape 名称，用于确认 p13 索引和 shape name。"""
import sys, time
import win32com.client

TEMPLATE_PATH = (
    r"D:\Technique Support\Claude Code Learning\3rd-ppt-prj"
    r"\template\apparel-page13-14-template.pptx"
)

def safe_print(*args):
    line = " ".join(str(a) for a in args)
    try:
        print(line)
    except UnicodeEncodeError:
        print(line.encode(sys.stdout.encoding or "ascii", "replace")
              .decode(sys.stdout.encoding or "ascii"))

app = prs = None
try:
    app = win32com.client.DispatchEx("PowerPoint.Application")
    prs = app.Presentations.Open(TEMPLATE_PATH, ReadOnly=True, Untitled=False, WithWindow=False)
    time.sleep(0.5)
    safe_print(f"Total slides: {prs.Slides.Count}")
    for si in range(1, prs.Slides.Count + 1):
        sld = prs.Slides(si)
        safe_print(f"\n--- Slide {si} (shape count={sld.Shapes.Count}) ---")
        for i in range(1, sld.Shapes.Count + 1):
            shp = sld.Shapes(i)
            try:
                txt = str(shp.TextFrame.TextRange.Text)[:40].replace("\r","\\r").replace("\n","\\n")
            except Exception:
                txt = "(no text)"
            safe_print(f"  [{i}] name={shp.Name!r}  txt={txt!r}")
finally:
    if prs:
        try: prs.Close()
        except: pass
    if app:
        try: app.Quit()
        except: pass
    safe_print("Done.")
