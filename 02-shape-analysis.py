#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step2: 生成 shape->数据映射、prompt规格、可读性预算。"""

from __future__ import annotations

import json
from pathlib import Path

from ppt_pipeline_common import ROOT, extract_metrics, load_excel_rows, now_ts, safe_text, write_json, write_md

SHAPE_JSON = ROOT / "shape_detail_com.json"
OUT_MD = ROOT / "02-shape-analysis.md"
OUT_MAP = ROOT / "shape_analysis_map.json"
OUT_PROMPTS = ROOT / "prompt_specs.json"
OUT_BUDGET = ROOT / "readability_budget.json"


def infer_role(item: dict) -> str:
    txt = safe_text(item.get("text", ""))
    name = safe_text(item.get("name", "")).lower()
    if item.get("has_chart"):
        return "chart"
    if "title" in name or len(txt) <= 18:
        return "title"
    if len(txt) >= 80:
        return "long_summary"
    if any(k in txt for k in ["建议", "结论", "总结"]):
        return "insight"
    if any(k in txt for k in ["样本", "人数", "n="]):
        return "sample_stat"
    return "body"


def main() -> int:
    if not SHAPE_JSON.exists():
        write_md(OUT_MD, ["# 02-shape-analysis", "", "- 状态: blocked", "- 原因: 缺少shape_detail_com.json"])
        return 0

    data = json.loads(SHAPE_JSON.read_text(encoding="utf-8"))
    shapes = data.get("new_shapes", []) if isinstance(data, dict) else []
    if not shapes:
        write_md(OUT_MD, ["# 02-shape-analysis", "", "- 状态: blocked", "- 原因: new_shapes为空"])
        return 0

    try:
        rows, sheet, notes = load_excel_rows("问卷sheet")
    except Exception as e:
        write_md(OUT_MD, ["# 02-shape-analysis", "", "- 状态: blocked", f"- 原因: {e}"])
        return 0

    metrics = extract_metrics(rows)
    headers = metrics["headers"]

    mapping = []
    prompts = []
    budgets = []
    for i, shp in enumerate(shapes[:9], 1):
        role = infer_role(shp)
        shape_name = shp.get("name", f"shape_{i}")
        template_text = safe_text(shp.get("text", ""))
        max_chars = max(18, min(220, int(len(template_text) * 1.15) if template_text else 80))
        max_lines = 1 if role == "title" else (8 if role in {"long_summary", "insight", "body"} else 4)

        mapping.append({
            "index": i,
            "shape_name": shape_name,
            "role": role,
            "source_sheet": sheet,
            "source_headers": headers[:20],
            "build_function": f"build_shape_{i}",
            "template_text": template_text,
        })

        prompts.append({
            "shape_name": shape_name,
            "role": role,
            "model": "openai/gpt-5-mini",
            "instruction": (
                "请输出PPT可直接粘贴中文文本；客观、简短、可读性强；"
                "严格遵守字数与行数限制；不要markdown，不要解释。"
            ),
            "context_headers": headers[:20],
        })

        budgets.append({
            "shape_name": shape_name,
            "role": role,
            "max_chars": max_chars,
            "max_lines": max_lines,
            "max_bullets": 4,
        })

    write_json(OUT_MAP, {"generated_at": now_ts(), "mapping": mapping, "metrics": metrics})
    write_json(OUT_PROMPTS, {"generated_at": now_ts(), "prompts": prompts})
    write_json(OUT_BUDGET, {"generated_at": now_ts(), "budgets": budgets})

    lines = [
        "# 02-shape-analysis",
        "",
        f"- 状态: ok",
        f"- 时间: {now_ts()}",
        f"- 数据sheet: {sheet}",
        f"- notes: {notes}",
        f"- mapping_count: {len(mapping)}",
        "",
        "## 输出文件",
        "- shape_analysis_map.json",
        "- prompt_specs.json",
        "- readability_budget.json",
        "",
    ]
    for m in mapping:
        lines += [
            f"### {m['index']}. {m['shape_name']}",
            f"- role: {m['role']}",
            f"- build_function: {m['build_function']}",
            f"- template_text_len: {len(m['template_text'])}",
            "",
        ]
    write_md(OUT_MD, lines)
    print(f"[OK] wrote {OUT_MD.name}, {OUT_MAP.name}, {OUT_PROMPTS.name}, {OUT_BUDGET.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
