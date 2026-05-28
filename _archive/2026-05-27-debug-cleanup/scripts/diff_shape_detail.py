"""深度对比 p13/p14 shape 细节 diff（新生成 vs 模板）。

A: v1.0.pptx slide 12 (新 p13) + slide 13 (新 p14)
B: apparel-page13-14-template.pptx slide 13 (模板 p13) + slide 14 (模板 p14)

每个 shape 取：name / L / T / W / H / fill / line / 每段文本的 font / size / bold / color / align / 段落行距
对比 by shape name，输出 diff 报告。
"""
import os
import sys
from pathlib import Path
import win32com.client


def safe_get(obj, attr, default=None):
    try:
        v = getattr(obj, attr)
        return v
    except Exception:
        return default


def long_to_rgb(c):
    if c is None:
        return None
    try:
        v = int(c)
        return (v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF)
    except Exception:
        return None


def shape_detail(shp):
    d = {
        "name": safe_get(shp, "Name", ""),
        "L": round(float(safe_get(shp, "Left", 0) or 0), 2),
        "T": round(float(safe_get(shp, "Top", 0) or 0), 2),
        "W": round(float(safe_get(shp, "Width", 0) or 0), 2),
        "H": round(float(safe_get(shp, "Height", 0) or 0), 2),
        "has_text": False,
        "paragraphs": [],
        "fill_color": None,
        "fill_visible": None,
        "line_color": None,
        "line_visible": None,
        "line_weight": None,
        "text_frame": {},
    }

    # fill
    try:
        fill = shp.Fill
        d["fill_visible"] = bool(fill.Visible)
        try:
            d["fill_color"] = long_to_rgb(fill.ForeColor.RGB)
        except Exception:
            pass
    except Exception:
        pass

    # line
    try:
        line = shp.Line
        d["line_visible"] = bool(line.Visible)
        try:
            d["line_color"] = long_to_rgb(line.ForeColor.RGB)
        except Exception:
            pass
        try:
            d["line_weight"] = round(float(line.Weight), 2)
        except Exception:
            pass
    except Exception:
        pass

    # text
    try:
        if shp.HasTextFrame and shp.TextFrame.HasText:
            d["has_text"] = True
            tf = shp.TextFrame
            d["text_frame"] = {
                "margin_l": round(float(safe_get(tf, "MarginLeft", 0) or 0), 2),
                "margin_r": round(float(safe_get(tf, "MarginRight", 0) or 0), 2),
                "margin_t": round(float(safe_get(tf, "MarginTop", 0) or 0), 2),
                "margin_b": round(float(safe_get(tf, "MarginBottom", 0) or 0), 2),
                "word_wrap": bool(safe_get(tf, "WordWrap", False) or False),
                "auto_size": int(safe_get(tf, "AutoSize", 0) or 0),
            }
            tr = tf.TextRange
            paras = tr.Paragraphs()
            for i in range(1, paras.Count + 1):
                p = paras.Item(i)
                font = p.Font
                pf = p.ParagraphFormat
                para = {
                    "text": p.Text.replace("\r", "").replace("\v", "")[:60],
                    "font_name": safe_get(font, "Name", ""),
                    "size": float(safe_get(font, "Size", 0) or 0),
                    "bold": int(safe_get(font, "Bold", 0) or 0),
                    "italic": int(safe_get(font, "Italic", 0) or 0),
                    "underline": int(safe_get(font, "Underline", 0) or 0),
                    "color": long_to_rgb(safe_get(font.Color, "RGB", None)),
                    "align": int(safe_get(pf, "Alignment", 0) or 0),
                    "space_before": float(safe_get(pf, "SpaceBefore", 0) or 0),
                    "space_after": float(safe_get(pf, "SpaceAfter", 0) or 0),
                    "space_within": float(safe_get(pf, "SpaceWithin", 0) or 0),
                    "bullet_type": int(safe_get(safe_get(pf, "Bullet", None), "Type", 0) or 0)
                                   if safe_get(pf, "Bullet", None) is not None else 0,
                }
                d["paragraphs"].append(para)
    except Exception as e:
        d["text_error"] = str(e)
    return d


def diff_dicts(a, b, path="", out=None):
    """递归 diff，只输出有差异的字段。"""
    if out is None:
        out = []
    if isinstance(a, dict) and isinstance(b, dict):
        keys = set(a.keys()) | set(b.keys())
        for k in sorted(keys):
            if k in ("L", "T", "W", "H"):
                # 几何上 ±2 之内忽略
                va, vb = a.get(k), b.get(k)
                if va is not None and vb is not None and abs(float(va) - float(vb)) < 2:
                    continue
            diff_dicts(a.get(k), b.get(k), f"{path}.{k}" if path else k, out)
    elif isinstance(a, list) and isinstance(b, list):
        for i in range(max(len(a), len(b))):
            va = a[i] if i < len(a) else None
            vb = b[i] if i < len(b) else None
            diff_dicts(va, vb, f"{path}[{i}]", out)
    else:
        if a != b:
            out.append((path, a, b))
    return out


def main():
    proj = Path(__file__).resolve().parent.parent
    snapshot = str(proj / "template" / "apparel-page13-14-template.pptx")

    app = win32com.client.GetActiveObject("PowerPoint.Application")

    # A = 用户当前打开的 v1.0
    pres_a = None
    for i in range(1, app.Presentations.Count + 1):
        p = app.Presentations(i)
        if "v 1.0" in p.Name or "v1.0" in p.Name:
            pres_a = p
            break
    if pres_a is None:
        pres_a = app.ActivePresentation
    print(f"[A] {pres_a.Name}  slides={pres_a.Slides.Count}")

    pres_b = app.Presentations.Open(snapshot, ReadOnly=True, WithWindow=True)
    print(f"[B] {pres_b.Name}  slides={pres_b.Slides.Count}")

    pairs = [
        ("p13", pres_a.Slides(12), pres_b.Slides(13)),
        ("p14", pres_a.Slides(13), pres_b.Slides(14)),
    ]

    out_md = ["# Shape 细节 diff 报告\n\n新生成 (A) vs 模板 (B)\n"]
    try:
        for label, slide_a, slide_b in pairs:
            out_md.append(f"\n## {label}\n")
            print(f"\n========= {label} =========")
            shapes_a = {s.Name: s for s in slide_a.Shapes}
            shapes_b = {s.Name: s for s in slide_b.Shapes}
            common = set(shapes_a.keys()) & set(shapes_b.keys())
            only_a = set(shapes_a.keys()) - common
            only_b = set(shapes_b.keys()) - common

            if only_a:
                out_md.append(f"\n**仅在 A**: {sorted(only_a)}\n")
            if only_b:
                out_md.append(f"\n**仅在 B**: {sorted(only_b)}\n")

            for name in sorted(common):
                a_det = shape_detail(shapes_a[name])
                b_det = shape_detail(shapes_b[name])
                diffs = diff_dicts(a_det, b_det)
                if not diffs:
                    continue
                out_md.append(f"\n### {name}\n")
                out_md.append("| field | A (new) | B (template) |\n|---|---|---|")
                for path, va, vb in diffs:
                    out_md.append(f"| `{path}` | `{va}` | `{vb}` |")
                out_md.append("")
                print(f"  {name}: {len(diffs)} diffs")
    finally:
        try:
            pres_b.Close()
        except Exception:
            pass

    report = "\n".join(out_md)
    out_path = proj / "debug" / "diff_p13p14_detail.md"
    out_path.write_text(report, encoding="utf-8")
    print(f"\n报告: {out_path}")


if __name__ == "__main__":
    main()
