"""快速探测 src/Template 2.1.pptx 当前页数（确认 merge 前基线）。"""
import os
import sys

import win32com.client

PROJ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGET = os.path.join(PROJ, "src", "Template 2.1.pptx")


def main():
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    pres = app.Presentations.Open(TARGET, ReadOnly=True)
    try:
        print(f"Template 2.1.pptx slides: {pres.Slides.Count}")
        for i in range(max(1, pres.Slides.Count - 4), pres.Slides.Count + 1):
            sld = pres.Slides(i)
            print(f"  slide {i}: {sld.Shapes.Count} shapes")
    finally:
        pres.Close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
