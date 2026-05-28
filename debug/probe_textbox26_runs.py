"""Probe template TextBox 26 runs to see the literal text of trailing size-16 runs.

Opens template/apparel-page13-14-template.pptx, slide 2 (p14), finds TextBox 26,
prints each run's text + (rgb, bold, size).
"""
import sys
from pathlib import Path

import win32com.client

TEMPLATE = (Path(__file__).resolve().parent.parent
            / "template" / "apparel-page13-14-template.pptx")

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True
pres = ppt.Presentations.Open(str(TEMPLATE), WithWindow=False, ReadOnly=True)
try:
    targets_by_name = {"TextBox 23": None, "TextBox 26": None}
    for sidx in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(sidx)
        for i in range(1, slide.Shapes.Count + 1):
            s = slide.Shapes(i)
            nm = str(s.Name)
            if nm in targets_by_name:
                targets_by_name[nm] = (sidx, s)

    print(f"Total slides: {pres.Slides.Count}")
    for nm, info in targets_by_name.items():
        print(f"  {nm}: slide {info[0] if info else 'NOT FOUND'}")
    print()
    # Probe both
    for nm in ("TextBox 23", "TextBox 26"):
        info = targets_by_name[nm]
        if info is None:
            continue
        target_slide_idx, target = info
        print(f"========== {nm} (slide {target_slide_idx}) ==========")

        tr = target.TextFrame.TextRange
        full = tr.Text
        print(f"FULL TEXT: {full!r}")
        n_runs = tr.Runs().Count
        n_para = tr.Paragraphs().Count
        print(f"Total runs: {n_runs}; total paragraphs: {n_para}")
        sizes_seen = set()
        for i in range(1, n_runs + 1):
            r = tr.Runs(i)
            try: sz = float(r.Font.Size)
            except Exception: sz = 0.0
            sizes_seen.add(sz)
        print(f"Sizes used: {sorted(sizes_seen)}")
        # Last paragraph detail
        if n_para > 0:
            last_para = tr.Paragraphs(n_para)
            print(f"Last para: {last_para.Text!r}")
            try:
                last_size = float(last_para.Font.Size)
            except Exception:
                last_size = 0.0
            print(f"Last para Font.Size: {last_size}")
        print()
finally:
    pres.Close()
