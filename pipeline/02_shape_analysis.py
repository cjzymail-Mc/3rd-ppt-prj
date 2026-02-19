#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 2: Role inference + prompt specs + readability budget.

Reads shape_detail_com.json + Excel data.
Assigns each shape a role, generates prompt specs and readability budgets.

Human-in-the-loop:
  If shape_detail.md contains user annotations, they override the auto-inferred
  values. Supported annotation fields:
    - 内容来源   -> injected into prompt as explicit data source
    - 生成方式   -> overrides build strategy (template_direct / gpt_prompted / ...)
    - 修正说明   -> appended to prompt as additional constraint
    - 角色覆盖   -> overrides auto-inferred role
    - prompt覆盖 -> replaces default instruction entirely
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    ROOT,
    extract_metrics,
    load_excel_rows,
    now_ts,
    parse_user_annotations,
    safe_print,
    safe_text,
    setup_console_encoding,
    write_json,
    write_md,
)

SHAPE_JSON = ROOT / "shape_detail_com.json"
OUT_MAP = ROOT / "shape_analysis_map.json"
OUT_PROMPTS = ROOT / "prompt_specs.json"
OUT_BUDGET = ROOT / "readability_budget.json"


def infer_role(item: dict) -> str:
    """Assign role with strict priority order per strategy matrix."""
    txt = safe_text(item.get("text", ""))
    name = safe_text(item.get("name", "")).lower()

    # 1. has_chart -> chart
    if item.get("has_chart"):
        return "chart"

    # 2. name contains "title" -> title
    if "title" in name:
        return "title"

    # 3. text contains sample/count keywords -> sample_stat
    if any(k in txt for k in ["样本", "人数", "N=", "n="]):
        return "sample_stat"

    # 4. text contains insight keywords -> insight
    if any(k in txt for k in ["建议", "结论", "总结"]):
        return "insight"

    # 5. short non-empty text -> title
    if 0 < len(txt) <= 18:
        return "title"

    # 6. long text -> long_summary
    if len(txt) >= 80:
        return "long_summary"

    # 7. default -> body
    return "body"


def prompt_rule(role: str) -> str:
    rules = {
        "title": "输出一个标题，不超过1行，保持和模板语气一致。",
        "sample_stat": "输出样本量/统计一句话，不超过1行。",
        "long_summary": "输出2~4行总结，保持模板句式和信息密度，禁止空话。",
        "insight": "输出2条建议，每条单行，先结论后动作。",
        "chart": "图表不走GPT，使用评分均值数据。",
        "body": "输出2~3行正文，语言客观，贴近模板风格。",
    }
    return rules.get(role, rules["body"])


def main() -> int:
    setup_console_encoding()
    if not SHAPE_JSON.exists():
        safe_print(f"[BLOCKED] Missing {SHAPE_JSON.name}")
        return 0

    data = json.loads(SHAPE_JSON.read_text(encoding="utf-8"))
    shapes = data.get("new_shapes", []) if isinstance(data, dict) else []
    if not shapes:
        safe_print("[BLOCKED] new_shapes is empty")
        return 0

    try:
        rows, sheet, notes = load_excel_rows("问卷sheet")
    except Exception as e:
        safe_print(f"[BLOCKED] Excel load failed: {e}")
        return 0

    metrics = extract_metrics(rows)
    headers = metrics["headers"]

    # --- Load user annotations (human-in-the-loop) ---
    annotations = parse_user_annotations()
    if annotations:
        safe_print(f"[INFO] Found user annotations for {len(annotations)} shapes")

    mapping = []
    prompts = []
    budgets = []

    for i, shp in enumerate(shapes, 1):
        shape_name = shp.get("name", f"shape_{i}")
        template_text = safe_text(shp.get("text", ""))
        anno = annotations.get(shape_name, {})

        # Role: user override > auto inference
        if anno.get("role_override"):
            role = anno["role_override"]
        else:
            role = infer_role(shp)

        # Instruction: user override > default rule
        if anno.get("prompt_override"):
            instruction = anno["prompt_override"]
        else:
            instruction = prompt_rule(role)

        # Append fix_notes as additional constraint
        if anno.get("fix_notes"):
            instruction += f"\n[用户修正] {anno['fix_notes']}"

        # Inject content_source into prompt context
        content_source_note = anno.get("content_source", "")

        # Build strategy hint (informational, used by Step 3A)
        strategy_hint = anno.get("build_strategy", "")

        max_chars = max(18, min(240, int(len(template_text) * 1.2) if template_text else 100))
        # max_lines: derive from template text line count when available
        if template_text:
            natural = len([l for l in template_text.replace("\r", "\n").splitlines() if l.strip()])
            max_lines = max(natural, 1)
        else:
            max_lines = 1 if role == "title" else (4 if role == "insight" else 6)

        # Structured machine-readable fields (new)
        strategy_exact = anno.get("strategy_exact", "")
        params_raw = anno.get("params", "")

        m = {
            "index": i,
            "shape_name": shape_name,
            "role": role,
            "source_sheet": sheet,
            "source_headers": headers[:30],
            "build_function": f"build_shape_{i}",
            "template_text": template_text,
            "template_text_len": len(template_text),
        }
        if strategy_hint:
            m["user_strategy_hint"] = strategy_hint
        if content_source_note:
            m["user_content_source"] = content_source_note
        if strategy_exact:
            m["strategy_exact"] = strategy_exact      # exact code, preferred over hint
        if params_raw:
            m["params"] = params_raw                  # raw string, parsed by Step 3A
        if anno:
            m["has_user_annotation"] = True
        mapping.append(m)

        p = {
            "shape_name": shape_name,
            "role": role,
            "model": "openai/gpt-5-mini",
            "goal": "根据源数据提炼与模板风格匹配的内容",
            "style_anchor": template_text[:200],
            "instruction": instruction,
            "output_constraints": {
                "max_chars": max_chars,
                "max_lines": max_lines,
                "no_markdown": True,
                "no_fabrication": True,
            },
            "context_headers": headers[:30],
        }
        if content_source_note:
            p["user_content_source"] = content_source_note
        prompts.append(p)

        budgets.append({
            "shape_name": shape_name,
            "role": role,
            "max_chars": max_chars,
            "max_lines": max_lines,
            "max_bullets": 4,
        })

    write_json(OUT_MAP, {
        "generated_at": now_ts(),
        "has_user_annotations": bool(annotations),
        "annotated_shapes": list(annotations.keys()),
        "mapping": mapping,
        "metrics": metrics,
    })
    write_json(OUT_PROMPTS, {
        "generated_at": now_ts(),
        "prompts": prompts,
    })
    write_json(OUT_BUDGET, {
        "generated_at": now_ts(),
        "budgets": budgets,
    })

    anno_msg = f" ({len(annotations)} user-annotated)" if annotations else ""
    safe_print(f"[OK] {len(mapping)} shapes analyzed{anno_msg}. "
          f"Wrote {OUT_MAP.name}, {OUT_PROMPTS.name}, {OUT_BUDGET.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
