"""深度对比 p13/p14 shape 的 runs 级字体细节。

paragraph 内 character runs（字符片段）级别才能看到染色 / 局部 bold / 局部字号变化。
按 shape -> paragraph -> run 三级遍历，按 run.Text/Font 取属性。
"""
import json
from pathlib import Path
import win32com.client


def safe_get(obj, attr, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def long_to_rgb(c):
    if c is None:
        return None
    try:
        v = int(c)
        return [v & 0xFF, (v >> 8) & 0xFF, (v >> 16) & 0xFF]
    except Exception:
        return None


def runs_of_paragraph(para):
    """把 paragraph 拆成 character runs（按 Font 变化分段）。"""
    runs = []
    try:
        tr = para
        n = tr.Length
        i = 1
        while i <= n:
            run = tr.Characters(i, 1)
            font = run.Font
            attr = {
                "name": safe_get(font, "Name", ""),
                "size": float(safe_get(font, "Size", 0) or 0),
                "bold": int(safe_get(font, "Bold", 0) or 0),
                "italic": int(safe_get(font, "Italic", 0) or 0),
                "color": long_to_rgb(safe_get(font.Color, "RGB", None)),
            }
            if runs and runs[-1]["attr"] == attr:
                runs[-1]["text"] += run.Text
            else:
                runs.append({"text": run.Text, "attr": attr})
            i += 1
    except Exception as e:
        runs.append({"error": str(e)})
    return runs


def shape_runs(shp):
    """每个 shape 输出 paragraphs -> runs。"""
    name = safe_get(shp, "Name", "")
    if not safe_get(shp, "HasTextFrame", False):
        return None
    try:
        tf = shp.TextFrame
        if not tf.HasText:
            return None
        tr = tf.TextRange
        paras = tr.Paragraphs()
        out = []
        for i in range(1, paras.Count + 1):
            p = paras.Item(i)
            text = p.Text.replace("\r", "").replace("\v", "")
            if not text:
                continue
            out.append({
                "p_text": text[:80],
                "runs": runs_of_paragraph(p),
            })
        return {"name": name, "paragraphs": out}
    except Exception as e:
        return {"name": name, "error": str(e)}


def runs_diff(a_runs, b_runs):
    """对比两个 shape 的 paragraphs/runs，返回差异列表。"""
    diffs = []
    if a_runs is None and b_runs is None:
        return diffs
    if a_runs is None or b_runs is None:
        diffs.append(("has_text", a_runs is not None, b_runs is not None))
        return diffs
    a_paras = a_runs.get("paragraphs", [])
    b_paras = b_runs.get("paragraphs", [])
    if len(a_paras) != len(b_paras):
        diffs.append(("paragraph_count", len(a_paras), len(b_paras)))
    for i in range(min(len(a_paras), len(b_paras))):
        ap, bp = a_paras[i], b_paras[i]
        a_text = ap.get("p_text", "")
        b_text = bp.get("p_text", "")
        a_runs_list = ap.get("runs", [])
        b_runs_list = bp.get("runs", [])

        # 文字内容差异不报（聚焦字体/颜色/大小）
        # 但 runs 数差异要报
        if len(a_runs_list) != len(b_runs_list):
            diffs.append((f"p{i+1}.runs_count", len(a_runs_list), len(b_runs_list)))
            # 列出 B 模板的 runs 细节
            diffs.append((f"p{i+1}.B_runs_summary",
                         [(r['text'], r['attr'].get('color'), r['attr'].get('bold')) for r in b_runs_list[:5]],
                         None))
            continue
        for j in range(len(a_runs_list)):
            a_attr = a_runs_list[j].get("attr", {})
            b_attr = b_runs_list[j].get("attr", {})
            for k in ("name", "size", "bold", "italic", "color"):
                if a_attr.get(k) != b_attr.get(k):
                    diffs.append(
                        (f"p{i+1}.run{j+1}.{k}  ({a_runs_list[j].get('text','')[:20]})",
                         a_attr.get(k), b_attr.get(k))
                    )
    return diffs


def main():
    proj = Path(__file__).resolve().parent.parent
    snapshot = str(proj / "template" / "apparel-page13-14-template.pptx")
    app = win32com.client.GetActiveObject("PowerPoint.Application")

    pres_a = None
    for i in range(1, app.Presentations.Count + 1):
        p = app.Presentations(i)
        if "v 1.0" in p.Name or "v1.0" in p.Name:
            pres_a = p
            break
    if pres_a is None:
        pres_a = app.ActivePresentation
    print(f"[A] {pres_a.Name}")

    pres_b = app.Presentations.Open(snapshot, ReadOnly=True, WithWindow=True)
    print(f"[B] {pres_b.Name}")

    pairs = [
        ("p13", pres_a.Slides(12), pres_b.Slides(13)),
        ("p14", pres_a.Slides(13), pres_b.Slides(14)),
    ]

    md = ["# Shape Runs 细节 diff（字体/颜色/字号/染色）\n"]
    try:
        for label, slide_a, slide_b in pairs:
            md.append(f"\n## {label}\n")
            print(f"\n========= {label} =========")
            a_map = {s.Name: shape_runs(s) for s in slide_a.Shapes if safe_get(s, "HasTextFrame", False)}
            b_map = {s.Name: shape_runs(s) for s in slide_b.Shapes if safe_get(s, "HasTextFrame", False)}
            common = set(a_map.keys()) & set(b_map.keys())
            for name in sorted(common):
                diffs = runs_diff(a_map[name], b_map[name])
                if not diffs:
                    continue
                md.append(f"\n### {name}\n")
                md.append("| field | A (new) | B (template) |\n|---|---|---|")
                for path, va, vb in diffs:
                    md.append(f"| `{path}` | `{va}` | `{vb}` |")
                md.append("")
                print(f"  {name}: {len(diffs)} runs-level diffs")
    finally:
        try:
            pres_b.Close()
        except Exception:
            pass

    out = proj / "debug" / "diff_p13p14_runs.md"
    out.write_text("\n".join(md), encoding="utf-8")
    print(f"\n报告: {out}")


if __name__ == "__main__":
    main()
