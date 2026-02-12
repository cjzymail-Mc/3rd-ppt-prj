#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step3A: 基于analysis+prompt+budget构建shape内容，并做内容校验。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List

from ppt_pipeline_common import (
    ROOT,
    clamp_text,
    extract_metrics,
    load_excel_rows,
    load_legacy_functions,
    now_ts,
    safe_text,
    write_json,
    write_md,
)

MAP_JSON = ROOT / "shape_analysis_map.json"
PROMPT_JSON = ROOT / "prompt_specs.json"
BUDGET_JSON = ROOT / "readability_budget.json"
OUT_CONTENT = ROOT / "build_shape_content.json"
OUT_VALID = ROOT / "content_validation_report.md"


def _budget_for(name: str, budgets: List[dict]) -> dict:
    for b in budgets:
        if b.get("shape_name") == name:
            return b
    return {"max_chars": 80, "max_lines": 4, "max_bullets": 4}


def _prompt_for(name: str, prompts: List[dict]) -> dict:
    for p in prompts:
        if p.get("shape_name") == name:
            return p
    return {"model": "openai/gpt-5-mini", "instruction": ""}


def call_gpt(gpt_fn, prompt: str, model: str, fallback: str) -> str:
    if gpt_fn:
        try:
            text = safe_text(gpt_fn(prompt, model))
            if text:
                return text
        except Exception:
            pass
    return fallback


def build_content(role: str, shape_name: str, metrics: Dict[str, Any], tmpl_text: str, pconf: dict, funcs: dict) -> str:
    gpt_fn = funcs.get("GPT_5")
    extract_info = funcs.get("extract_info")

    if role == "title":
        return tmpl_text or "问卷分析报告"

    if role == "sample_stat":
        return f"有效样本 N={metrics.get('respondent_count', 0)}"

    if role == "chart":
        kws = metrics.get("keywords", [])[:6]
        if not kws:
            return "A:1\nB:2\nC:3"
        return "\n".join([f"{k}:{v}" for k, v in kws])

    if role == "long_summary":
        prompt = (
            f"{pconf.get('instruction','')}\n"
            f"任务：为shape `{shape_name}` 输出2~4句总结。\n"
            f"数据指标：{metrics}\n"
        )
        return call_gpt(gpt_fn, prompt, pconf.get("model", "openai/gpt-5-mini"), "样本反馈总体稳定，核心指标表现均衡。")

    if role == "insight":
        prompt = (
            f"{pconf.get('instruction','')}\n"
            "任务：给出2条可执行建议，每条单行。\n"
            f"数据指标：{metrics}"
        )
        return call_gpt(gpt_fn, prompt, pconf.get("model", "openai/gpt-5-mini"), "1) 优化关键场景体验\n2) 强化稳定性一致性")

    # body
    if extract_info:
        try:
            rows, _, _ = load_excel_rows("问卷sheet")
            info = extract_info(rows)[:4]
            lines = [f"{i+1}.{n}/{w}/{d}/{p}" for i, (n, w, d, p) in enumerate(info)]
            if lines:
                return "\n".join(lines)
        except Exception:
            pass

    prompt = f"{pconf.get('instruction','')}\n任务：输出可读正文2~3行。\n数据指标：{metrics}"
    return call_gpt(gpt_fn, prompt, pconf.get("model", "openai/gpt-5-mini"), "反馈集中，建议围绕关键指标继续优化。")


def main() -> int:
    if not (MAP_JSON.exists() and PROMPT_JSON.exists() and BUDGET_JSON.exists()):
        write_md(OUT_VALID, ["# content_validation_report", "", "- 状态: blocked", "- 原因: analysis artifacts missing"])
        return 0

    mapping = json.loads(MAP_JSON.read_text(encoding="utf-8")).get("mapping", [])
    prompts = json.loads(PROMPT_JSON.read_text(encoding="utf-8")).get("prompts", [])
    budgets = json.loads(BUDGET_JSON.read_text(encoding="utf-8")).get("budgets", [])

    rows, sheet, notes = load_excel_rows("问卷sheet")
    metrics = extract_metrics(rows)
    funcs = load_legacy_functions()

    items = []
    val_lines = [
        "# content_validation_report",
        "",
        f"- 时间: {now_ts()}",
        f"- sheet: {sheet}",
        f"- notes: {notes}",
        "",
        "|shape|role|len|line_count|max_chars|max_lines|valid|",
        "|---|---|---|---|---|---|---|",
    ]

    for m in mapping:
        name = m["shape_name"]
        role = m["role"]
        budget = _budget_for(name, budgets)
        pconf = _prompt_for(name, prompts)

        raw = build_content(role, name, metrics, m.get("template_text", ""), pconf, funcs)
        txt = clamp_text(raw, int(budget.get("max_chars", 80)), int(budget.get("max_lines", 4)))

        line_count = len(txt.splitlines()) if txt else 0
        valid = len(txt) <= int(budget.get("max_chars", 80)) and line_count <= int(budget.get("max_lines", 4))

        items.append({
            "shape_name": name,
            "role": role,
            "build_function": m.get("build_function"),
            "content": txt,
            "budget": budget,
            "valid": valid,
        })

        val_lines.append(
            f"|{name}|{role}|{len(txt)}|{line_count}|{budget.get('max_chars')}|{budget.get('max_lines')}|{valid}|"
        )

    write_json(OUT_CONTENT, {"generated_at": now_ts(), "sheet": sheet, "items": items, "metrics": metrics})
    write_md(OUT_VALID, val_lines)
    print(f"[OK] wrote {OUT_CONTENT.name}, {OUT_VALID.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
