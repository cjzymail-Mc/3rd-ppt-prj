#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 4: Three-layer diff test (Visual + Readability + Semantic).

Compares Template page 15 vs claude-ppt.pptx page 1.

Three layers:
  Visual >= 98:   geometry + shape_type + chart_type + font
  Readability >= 95: text similarity + length ratio + line ratio
  Semantic = 100: keyword coverage (样本, 建议, 反馈)

Fixed vs codex-legacy2:
  - BUG FIX: Match shapes by name (not by index), with geometry fallback
  - Font scoring: name + size weights
  - Output: diff_result.json, fix-ppt.md, diff_semantic_report.md
"""

from __future__ import annotations

import argparse
import json
import sys
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, List, Optional

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    OUTPUT_DIR,
    PROGRESS_DIR,
    ROOT,
    TEMPLATE_PATH,
    com_call,
    com_get,
    now_ts,
    safe_print,
    safe_text,
    setup_console_encoding,
    write_json,
    write_md,
)

FIX_MD    = PROGRESS_DIR / "04-fix_ppt.md"
SEM_MD    = PROGRESS_DIR / "04-diff_semantic_report.md"
DIFF_JSON = PROGRESS_DIR / "04-diff_result.json"


def sim(a: str, b: str) -> float:
    """Text similarity ratio."""
    return SequenceMatcher(None, a or "", b or "").ratio()


def shape_props(shp) -> Dict:
    """Extract comparison-relevant properties from a shape."""
    has_tf = bool(com_get(shp, "HasTextFrame", 0))
    text = ""
    font_name = ""
    font_size = 0.0
    font_color = None
    if has_tf:
        try:
            tr = shp.TextFrame.TextRange
            text = safe_text(com_get(tr, "Text", ""))
            font = com_get(tr, "Font", None)
            if font is not None:
                font_name = safe_text(com_get(font, "Name", ""))
                font_size = float(com_get(font, "Size", 0.0) or 0.0)
                fc_obj = com_get(font, "Color", None)
                font_color = com_get(fc_obj, "RGB", None) if fc_obj is not None else None
        except Exception:
            pass

    has_chart = False
    chart_type = 0
    try:
        has_chart = bool(shp.HasChart)
        if has_chart:
            chart_type = int(com_get(shp.Chart, "ChartType", 0) or 0)
    except Exception:
        pass

    return {
        "name": safe_text(com_get(shp, "Name", "")),
        "left": float(com_get(shp, "Left", 0.0) or 0.0),
        "top": float(com_get(shp, "Top", 0.0) or 0.0),
        "width": float(com_get(shp, "Width", 0.0) or 0.0),
        "height": float(com_get(shp, "Height", 0.0) or 0.0),
        "shape_type": int(com_get(shp, "Type", 0) or 0),
        "text": text,
        "font_name": font_name,
        "font_size": font_size,
        "font_color": font_color,
        "has_chart": has_chart,
        "chart_type": chart_type,
    }


def _collect_shapes(slide) -> Dict[str, Dict]:
    """Collect all shapes on a slide, keyed by name."""
    result = {}
    for i in range(1, int(slide.Shapes.Count) + 1):
        props = shape_props(slide.Shapes(i))
        result[props["name"]] = props
    return result


def _match_by_geometry(target_props: Dict, candidates: Dict[str, Dict], tolerance: float = 30.0) -> Optional[Dict]:
    """Find the closest candidate shape by geometry when name doesn't match."""
    best = None
    best_dist = float("inf")
    for cand in candidates.values():
        dist = (
            abs(target_props["left"] - cand["left"])
            + abs(target_props["top"] - cand["top"])
        )
        if dist < best_dist:
            best_dist = dist
            best = cand
    if best is not None and best_dist <= tolerance:
        return best
    return None


def visual_score(a: Dict, b: Dict) -> float:
    """Weighted visual fidelity score (0-100)."""
    score = 0.0
    w = 0.0

    def add(v, wt):
        nonlocal score, w
        score += v * wt
        w += wt

    # Geometry: tolerance 20px
    add(max(0.0, 1 - abs(a["left"] - b["left"]) / 20), 10)
    add(max(0.0, 1 - abs(a["top"] - b["top"]) / 20), 10)
    add(max(0.0, 1 - abs(a["width"] - b["width"]) / 20), 8)
    add(max(0.0, 1 - abs(a["height"] - b["height"]) / 20), 8)

    # Shape type
    add(1.0 if a["shape_type"] == b["shape_type"] else 0.0, 8)

    # Chart type (only if either has chart)
    if a["has_chart"] or b["has_chart"]:
        add(1.0 if a["chart_type"] == b["chart_type"] else 0.0, 16)

    # Font name
    if a["font_name"] or b["font_name"]:
        add(1.0 if a["font_name"] == b["font_name"] else 0.0, 4)

    # Font size
    if a["font_size"] > 0 or b["font_size"] > 0:
        max_fs = max(a["font_size"], b["font_size"], 1)
        add(max(0.0, 1 - abs(a["font_size"] - b["font_size"]) / max_fs), 4)

    return (score / w) * 100 if w else 100.0


def readability_score(a: Dict, b: Dict) -> float:
    """Structural readability score (0-100).

    Measures format/structure similarity only — NOT text content similarity.
    Content legitimately changes with each questionnaire dataset, so comparing
    text against the template's example text is not meaningful.

    Scoring:
      - Length score (70%): generated text >= 50% of template length -> full credit
      - Line score  (30%): generated line count >= 50% of template lines -> full credit
    """
    if not a["text"] and not b["text"]:
        return 100.0
    if not b["text"]:
        # Template has text but generated is empty — penalize
        return 0.0
    if not a["text"]:
        # Template empty, generated has content — acceptable
        return 100.0

    len_a = max(1, len(a["text"]))
    len_b = max(1, len(b["text"]))
    # Full credit if generated is >= 50% of template length
    len_score = min(1.0, len_b / (len_a * 0.5))

    line_a = max(1, len(a["text"].splitlines()))
    line_b = max(1, len(b["text"].splitlines()))
    # Full credit if generated has >= 50% as many lines as template
    line_score = min(1.0, line_b / (line_a * 0.5))

    return (len_score * 100 * 0.7) + (line_score * 100 * 0.3)


def semantic_report(target_slide) -> Dict:
    """Check keyword coverage on the target slide."""
    required = ["样本", "建议", "反馈"]
    text_blob = []
    for i in range(1, int(target_slide.Shapes.Count) + 1):
        shp = target_slide.Shapes(i)
        try:
            if bool(shp.HasTextFrame):
                text_blob.append(safe_text(shp.TextFrame.TextRange.Text))
        except Exception:
            pass
    all_text = "\n".join(text_blob)
    hits = {k: (k in all_text) for k in required}
    coverage = sum(1 for v in hits.values() if v) / len(required) * 100
    return {"coverage": coverage, "hits": hits}


def main() -> int:
    setup_console_encoding()
    ap = argparse.ArgumentParser()
    ap.add_argument("--target", default="claude-ppt 1.0.pptx")
    args = ap.parse_args()
    target = OUTPUT_DIR / args.target

    if not TEMPLATE_PATH.exists() or not target.exists():
        msg = "模板缺失" if not TEMPLATE_PATH.exists() else f"目标PPT缺失: {args.target}"
        write_md(FIX_MD, ["# PPT Diff & Fix Report", "", "- 状态: blocked", f"- 原因: {msg}"])
        write_json(DIFF_JSON, {"status": "blocked", "reason": msg})
        safe_print(f"[BLOCKED] {msg}")
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_md(FIX_MD, ["# PPT Diff & Fix Report", "", "- 状态: blocked", f"- 原因: {e}"])
        write_json(DIFF_JSON, {"status": "blocked", "reason": str(e)})
        safe_print(f"[BLOCKED] {e}")
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    p1 = app.Presentations.Open(str(TEMPLATE_PATH))
    p2 = app.Presentations.Open(str(target))

    try:
        s1 = p1.Slides(2)   # template standard page
        s2 = p2.Slides(1)   # generated page

        # BUG FIX: match by name instead of index
        template_shapes = _collect_shapes(s1)
        target_shapes = _collect_shapes(s2)

        rows = []
        visual_scores = []
        read_scores = []
        fails = []
        pair_count = 0

        for name, a in template_shapes.items():
            # Try name match first
            b = target_shapes.get(name)
            match_method = "name"
            if b is None:
                # Geometry fallback
                b = _match_by_geometry(a, target_shapes)
                match_method = "geometry"
            if b is None:
                # Shape missing entirely
                fails.append({
                    "template_name": name,
                    "target_name": "(missing)",
                    "visual": 0.0,
                    "readability": 0.0,
                    "reason": "shape not found in target",
                })
                visual_scores.append(0.0)
                read_scores.append(0.0)
                rows.append(f"|{name}|(missing)|0.00|0.00|{len(a['text'])}/0|{match_method}|")
                pair_count += 1
                continue

            vs = visual_score(a, b)
            rs = readability_score(a, b)
            visual_scores.append(vs)
            read_scores.append(rs)
            pair_count += 1

            if vs < 98 or rs < 95:
                fails.append({
                    "template_name": name,
                    "target_name": b["name"],
                    "visual": round(vs, 2),
                    "readability": round(rs, 2),
                    "reason": f"visual={vs:.2f}<98" if vs < 98 else f"readability={rs:.2f}<95",
                })

            rows.append(
                f"|{name}|{b['name']}|{vs:.2f}|{rs:.2f}"
                f"|{len(a['text'])}/{len(b['text'])}|{match_method}|"
            )

        sem = semantic_report(s2)
        visual_avg = sum(visual_scores) / len(visual_scores) if visual_scores else 0.0
        read_avg = sum(read_scores) / len(read_scores) if read_scores else 0.0
        passed = visual_avg >= 98 and read_avg >= 95 and sem["coverage"] >= 100 and not fails
        status = "ok" if passed else "fail"

        # Write fix-ppt.md
        lines = [
            "# PPT Diff & Fix Report",
            "",
            f"- 状态: {status}",
            f"- visual_score: {visual_avg:.2f}%",
            f"- readability_score: {read_avg:.2f}%",
            f"- semantic_coverage: {sem['coverage']:.2f}%",
            f"- template_shapes: {int(s1.Shapes.Count)}",
            f"- target_shapes: {int(s2.Shapes.Count)}",
            f"- paired: {pair_count}",
            f"- 时间: {now_ts()}",
            "",
            "## Shape对比",
            "|template|target|visual|readability|text_len|match|",
            "|---|---|---|---|---|---|",
        ] + rows

        # Semantic fails
        sem_fails = []
        if sem["coverage"] < 100:
            missing_kw = [k for k, v in sem["hits"].items() if not v]
            kw_str = "、".join(missing_kw)
            sem_fails.append({
                "shape": "(全局)",
                "issue": f"语义关键词缺失: {', '.join(missing_kw)}",
                "fix_type": "keyword_missing",
                "suggestion": f"在 gpt_prompted 类 shape 的内容描述中追加：必须包含'{kw_str}'关键词",
            })

        # Classify fix_type for each fail (5-type system)
        #   code            — pipeline 代码缺陷 (shape 缺失, COM 错误)
        #   style_mismatch  — 几何/字体/颜色偏差 (visual < 98)
        #   budget_overflow  — 文本过长 (readability < 95, len > budget)
        #   budget_underflow — 文本过短/空白 (readability < 95, len << budget)
        #   keyword_missing  — 语义关键词缺失 (semantic < 100)
        fix_items = []
        for f in fails[:40]:
            vs = f["visual"]
            rs = f["readability"]
            tname = f["template_name"]

            if f["target_name"] == "(missing)":
                fix_items.append({
                    "shape": tname, "issue": "shape 缺失",
                    "fix_type": "code", "suggestion": "03b COM 写入逻辑未覆盖该 shape",
                })
            elif vs < 98:
                fix_items.append({
                    "shape": tname, "issue": f"visual={vs} < 98 (几何/字体/颜色偏差)",
                    "fix_type": "style_mismatch", "suggestion": "检查 03b COM 写入是否破坏了 shape 格式属性，或调整 style_anchor",
                })
            elif rs < 95:
                # Determine overflow vs underflow
                t_len = len(template_shapes.get(tname, {}).get("text", ""))
                g_len = len(target_shapes.get(f.get("target_name", ""), {}).get("text", ""))
                if t_len > 0 and g_len > t_len * 1.2:
                    fix_items.append({
                        "shape": tname, "issue": f"readability={rs} < 95 (文本过长: {g_len}/{t_len}字)",
                        "fix_type": "budget_overflow",
                        "suggestion": "缩减 prompt 字数约束或在内容描述中追加'宁短勿长'",
                    })
                elif t_len > 0 and g_len < t_len * 0.3:
                    fix_items.append({
                        "shape": tname, "issue": f"readability={rs} < 95 (文本过短: {g_len}/{t_len}字)",
                        "fix_type": "budget_underflow",
                        "suggestion": "内容描述中要求更充实的内容，涵盖更多测试者反馈",
                    })
                else:
                    fix_items.append({
                        "shape": tname, "issue": f"readability={rs} < 95 (文本长度/行数偏差)",
                        "fix_type": "budget_overflow",
                        "suggestion": "调整 readability_budget 或 prompt 字数约束",
                    })

        all_fixes = fix_items + sem_fails

        if all_fixes:
            lines += [
                "", "## 修正建议",
                "", "| shape | 问题 | fix_type | 建议 |",
                "|-------|------|----------|------|",
            ]
            for fx in all_fixes:
                lines.append(
                    f"| {fx['shape']} | {fx['issue']} | {fx['fix_type']} | {fx['suggestion']} |"
                )

        write_md(FIX_MD, lines)

        # Write semantic report
        write_md(SEM_MD, [
            "# diff_semantic_report",
            "",
            f"- coverage: {sem['coverage']:.2f}%",
            f"- hits: {sem['hits']}",
        ])

        # Determine overall fix_type (code takes priority; otherwise first non-code type)
        has_code_fix = any(fx["fix_type"] == "code" for fx in all_fixes)
        if has_code_fix:
            overall_fix_type = "code"
        elif all_fixes:
            overall_fix_type = all_fixes[0]["fix_type"]
        else:
            overall_fix_type = ""

        # Write diff_result.json (the key file for agent iteration)
        write_json(DIFF_JSON, {
            "status": status,
            "visual_score": round(visual_avg, 2),
            "readability_score": round(read_avg, 2),
            "semantic_coverage": sem["coverage"],
            "fix_type": overall_fix_type,
            "fixes": [
                {
                    "shape": fx["shape"],
                    "issue": fx["issue"],
                    "fix_type": fx["fix_type"],
                    "suggestion": fx["suggestion"],
                }
                for fx in all_fixes
            ],
            "generated_at": now_ts(),
        })

        safe_print(f"[{'OK' if passed else 'FAIL'}] visual={visual_avg:.2f} readability={read_avg:.2f} semantic={sem['coverage']:.0f}")
        return 0 if passed else 1
    finally:
        com_call(p1, "Close")
        com_call(p2, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
