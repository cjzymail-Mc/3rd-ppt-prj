"""Probe color/bold runs for 4 GPT-driven shapes in template vs v3 PPT.

模板：apparel-page13-14-template.pptx (slide 13 + 14)
v3 PPT：当前 active PowerPoint 的 slide 12 + 13

输出每个 shape 的 run 数 / 不同 rgb 数 / bold run 数 / 前 3 个 run 文本预览。
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

TARGETS = [
    ("p13", 13, 12, "TextBox 24"),  # 受试者信息（gpt_prompted）
    ("p14", 14, 13, "TextBox 24"),  # 同上 reused
    ("p14", 14, 13, "TextBox 23"),  # strengths（gpt_strengths_bullet）
    ("p14", 14, 13, "TextBox 26"),  # drawbacks（gpt_drawbacks_bullet）
]


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
                txt = str(ch.Text)
            except Exception:
                txt = ""
            return rgb, bold, txt

        cur_rgb, cur_bold, cur_text = attrs(1)
        cur_start, cur_len = 1, 1
        for i in range(2, n + 1):
            rgb, bold, txt = attrs(i)
            if (rgb, bold) == (cur_rgb, cur_bold):
                cur_len += 1
                cur_text += txt
            else:
                runs.append({"start": cur_start, "len": cur_len,
                             "rgb": cur_rgb, "bold": cur_bold, "text": cur_text})
                cur_start, cur_len = i, 1
                cur_rgb, cur_bold, cur_text = rgb, bold, txt
        runs.append({"start": cur_start, "len": cur_len,
                     "rgb": cur_rgb, "bold": cur_bold, "text": cur_text})
    except Exception as e:
        safe_print(f"  [walk_runs err] {e}")
    return runs


def report(label, runs):
    safe_print(f"\n  [{label}] {len(runs)} runs")
    if not runs:
        return
    rgbs = set(r["rgb"] for r in runs if r["rgb"] is not None)
    bolds = [r for r in runs if r["bold"] == -1]
    safe_print(f"    distinct_rgb={len(rgbs)} → {[hex(c) for c in rgbs]}")
    safe_print(f"    bold_runs={len(bolds)}")
    for i, r in enumerate(runs[:6]):
        txt = r["text"].replace("\r", "\\r").replace("\n", "\\n")[:40]
        rgb_hex = hex(r["rgb"]) if r["rgb"] is not None else "?"
        safe_print(f"    run{i}: rgb={rgb_hex} bold={r['bold']} len={r['len']} '{txt}'")


def find_shape(slide, name):
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        if str(s.Name) == name:
            return s
    return None


def main():
    ppt = win32com.client.GetActiveObject("PowerPoint.Application")

    # 找 v3 PPT（active）
    pres_v3 = ppt.ActivePresentation
    safe_print(f"v3 PPT: {pres_v3.Name}")

    # 打开模板（read-only）
    tpl_path = str(Path(r"D:/Technique Support/Claude Code Learning/3rd-ppt-prj/template/apparel-page13-14-template.pptx").resolve())
    pres_tpl = ppt.Presentations.Open(tpl_path, ReadOnly=True, WithWindow=False)
    safe_print(f"Template: {pres_tpl.Name}")

    try:
        for label, tpl_idx, v3_idx, shape_name in TARGETS:
            safe_print(f"\n========== [{label}] {shape_name} (tpl slide {tpl_idx} / v3 slide {v3_idx}) ==========")
            tpl_slide = pres_tpl.Slides(tpl_idx)
            v3_slide = pres_v3.Slides(v3_idx)
            tpl_shape = find_shape(tpl_slide, shape_name)
            v3_shape = find_shape(v3_slide, shape_name)
            if not tpl_shape:
                safe_print(f"  [TPL] shape '{shape_name}' not found on slide {tpl_idx}")
            else:
                report("TPL", walk_runs(tpl_shape))
            if not v3_shape:
                safe_print(f"  [V3] shape '{shape_name}' not found on slide {v3_idx}")
            else:
                report("V3", walk_runs(v3_shape))
    finally:
        try:
            pres_tpl.Close()
        except Exception:
            pass


if __name__ == "__main__":
    main()
