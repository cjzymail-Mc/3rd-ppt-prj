"""列出所有 slide 的 shape 数 + 前一两个 TextBox 文本。"""
import win32com.client

ppt = win32com.client.GetActiveObject("PowerPoint.Application")
pres = ppt.ActivePresentation
print(f"Active: {pres.Name}  slides={pres.Slides.Count}")
for sld in pres.Slides:
    idx = sld.SlideIndex
    n = sld.Shapes.Count
    titles = []
    for sh in sld.Shapes:
        if sh.HasTextFrame and sh.TextFrame.HasText:
            t = sh.TextFrame.TextRange.Text.replace("\r", " | ")[:40]
            if t.strip():
                titles.append(f"{sh.Name}={t!r}")
        if len(titles) >= 2:
            break
    print(f"slide {idx}: {n} shapes  {' / '.join(titles)}")
