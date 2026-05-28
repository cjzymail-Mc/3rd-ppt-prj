"""Probe character-level color runs for the 5 selected shapes.

5 shapes: TextBox 6 / TextBox 14 / TextBox 17 / TextBox 20 / TextBox 50
对照：模板 apparel-page13-14-template.pptx slide 13 vs v3 active slide 12.
"""
from __future__ import annotations
import sys
from pathlib import Path

HELPERS = Path.home() / ".claude" / "skills" / "office-com-helpers"
sys.path.insert(0, str(HELPERS))
from office_com_helpers import safe_print

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import win32com.client

SHAPES = ["TextBox 6", "TextBox 14", "TextBox 17", "TextBox 20", "TextBox 50"]


def walk_runs(shape):
    runs = []
    try:
        if int(getattr(shape, "HasTextFrame", 0) or 0) != -1:
            return runs
        tr = shape.TextFrame.TextRange
        n = int(tr.Length)
        if n <= 0:
            return runs

        def attrs(i):
            ch = tr.Characters(i, 1)
            try:
                rgb = int(ch.Font.Color.RGB)
            except Exception:
                rgb = None
            try:
                bold = int(ch.Font.Bold)
            except Exception:
                bold = None
            try:
                size = float(ch.Font.Size)
            except Exception:
                size = None
            try:
                txt = str(ch.Text)
            except Exception:
                txt = ""
            return rgb, bold, size, txt

        cur_rgb, cur_bold, cur_size, cur_text = attrs(1)
        cur_start, cur_len = 1, 1
        for i in range(2, n + 1):
            rgb, bold, size, txt = attrs(i)
            if (rgb, bold, size) == (cur_rgb, cur_bold, cur_size):
                cur_len += 1
                cur_text += txt
            else:
                runs.append({"start": cur_start, "len": cur_len,
                             "rgb": cur_rgb, "bold": cur_bold, "size": cur_size, "text": cur_text})
                cur_start, cur_len = i, 1
                cur_rgb, cur_bold, cur_size, cur_text = rgb, bold, size, txt
        runs.append({"start": cur_start, "len": cur_len,
                     "rgb": cur_rgb, "bold": cur_bold, "size": cur_size, "text": cur_text})
    except Exception as e:
        safe_print(f"  [walk_runs err] {e}")
    return runs


def report(label, runs):
    if not runs:
        safe_print(f"  [{label}] (empty)")
        return
    rgbs = set(r["rgb"] for r in runs if r["rgb"] is not None)
    sizes = set(r["size"] for r in runs if r["size"] is not None)
    safe_print(f"  [{label}] runs={len(runs)} distinct_rgb={len(rgbs)} sizes={sizes}")
    for i, r in enumerate(runs):
        txt = r["text"].replace("\r", "\\r").replace("\n", "\\n")
        rgb_hex = hex(r["rgb"]) if r["rgb"] is not None else "?"
        safe_print(f"    run{i}: rgb={rgb_hex} bold={r['bold']} size={r['size']} len={r['len']} '{txt}'")


def find_shape(slide, name):
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        if str(s.Name) == name:
            return s
    return None


def main():
    ppt = win32com.client.GetActiveObject("PowerPoint.Application")
    pres_v3 = ppt.ActivePresentation
    safe_print(f"V3 PPT: {pres_v3.Name}  slide=12")

    tpl_path = str(Path(r"D:/Technique Support/Claude Code Learning/3rd-ppt-prj/template/apparel-page13-14-template.pptx").resolve())
    pres_tpl = ppt.Presentations.Open(tpl_path, ReadOnly=True, WithWindow=False)
    safe_print(f"TPL PPT: {pres_tpl.Name}  slide=13\n")

    try:
        tpl_slide = pres_tpl.Slides(13)
        v3_slide = pres_v3.Slides(12)
        for name in SHAPES:
            safe_print(f"========== {name} ==========")
            tpl_shape = find_shape(tpl_slide, name)
            v3_shape = find_shape(v3_slide, name)
            if not tpl_shape:
                safe_print(f"  [TPL] missing")
            else:
                report("TPL", walk_runs(tpl_shape))
            if not v3_shape:
                safe_print(f"  [V3] missing")
            else:
                report("V3", walk_runs(v3_shape))
            safe_print("")
    finally:
        try:
            pres_tpl.Close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
