"""Probe a specific slide on the active presentation. Pass slide index as argv[1]."""
import sys
import win32com.client

idx = int(sys.argv[1]) if len(sys.argv) > 1 else 12
ppt = win32com.client.GetActiveObject("PowerPoint.Application")
pres = ppt.ActivePresentation
print(f"Active: {pres.Name}  slides={pres.Slides.Count}  asking slide {idx}")
sld = pres.Slides(idx)
print(f"Slide {idx}: {sld.Shapes.Count} shapes")
for sh in sld.Shapes:
    name = sh.Name
    has_text = bool(sh.HasTextFrame) and bool(sh.TextFrame.HasText)
    text_preview = ""
    if has_text:
        text_preview = sh.TextFrame.TextRange.Text[:80].replace("\r", "[CR]")
    print(f"  {name:30s}  L={sh.Left:.0f} T={sh.Top:.0f}  text={text_preview!r}")
