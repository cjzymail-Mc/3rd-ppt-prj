#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step3A: 基于analysis+prompt+budget构建shape内容，并做内容校验。"""

from __future__ import annotations

import json
from typing import Any, Dict, List

from ppt_pipeline_common import (
    ROOT,
    clamp_text,
    extract_metrics,
    extract_score_means,
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
OUT_PROMPT_TRACE = ROOT / "prompt_trace.json"
OUT_GAP = ROOT / "shape_data_gap_report.md"


def _budget_for(name: str, budgets: List[dict]) -> dict:
    for b in budgets:
        if b.get("shape_name") == name:
            return b
    return {"max_chars": 80, "max_lines": 4, "max_bullets": 4}


def _prompt_for(name: str, prompts: List[dict]) -> dict:
    for p in prompts:
        if p.get("shape_name") == name:
            return p
    return {"model": "openai/gpt-5-mini", "instruction": "", "style_anchor": ""}


def call_gpt(gpt_fn, prompt: str, model: str, fallback: str) -> str:
    if gpt_fn:
        try:
            text = safe_text(gpt_fn(prompt, model))
            if text:
                return text
        except Exception:
            pass
    return fallback


def build_prompt_from_template(shape_name: str, role: str, pconf: dict, budget: dict, metrics: Dict[str, Any], rows: List[List[Any]]) -> str:
    style_anchor = safe_text(pconf.get("style_anchor", ""))
    instruction = safe_text(pconf.get("instruction", ""))
    headers = pconf.get("context_headers", [])
    data_preview = rows[1:1 + min(6, max(1, len(rows) - 1))]

    return (
        "你是一个极其严谨的PPT内容工程师。\n"
        "你的任务：依据源数据生成用于单个shape的最终文案。\n"
        "硬约束：不得编造、不得输出解释、不得使用markdown。\n"
        f"shape名称：{shape_name}\n"
        f"shape角色：{role}\n"
        f"风格锚点（来自标准模板原文）：{style_anchor}\n"
        f"源数据字段头：{headers}\n"
        f"源数据片段：{data_preview}\n"
        f"数据统计：{metrics}\n"
        f"写作要求：{instruction}\n"
        f"输出约束：max_chars={budget.get('max_chars', 80)}, max_lines={budget.get('max_lines', 4)}\n"
        "请直接输出最终文本。"
    )


def build_content(role: str, shape_name: str, metrics: Dict[str, Any], tmpl_text: str, pconf: dict, budget: dict, funcs: dict, rows: List[List[Any]]) -> tuple[str, str, str, str]:
    """return: text, prompt_or_reason, strategy, gap_reason"""
    gpt_fn = funcs.get("GPT_5")
    extract_info = funcs.get("extract_info")

    # 1) chart shape -> deterministic mean extraction
    if role == "chart":
        means = extract_score_means(rows)
        if not means:
            return "减震:0\n回弹:0\n稳定:0", "chart_mean_fallback", "mean_extraction", "未识别到可用评分列"
        return "\n".join([f"{k}:{v:.2f}" for k, v in means[:8]]), "chart_mean_from_source", "mean_extraction", ""

    # 2) title/sample_stat -> deterministic no GPT
    if role == "title":
        title = tmpl_text or "问卷分析报告"
        return title, "title_from_template", "template_direct", "" if tmpl_text else "模板标题为空，使用默认标题"

    if role == "sample_stat":
        return f"有效样本 N={metrics.get('respondent_count', 0)}", "sample_count_from_source", "numeric_aggregation", ""

    # 3) body 优先 extract_info
    if role == "body" and extract_info:
        try:
            info = extract_info(rows)[:4]
            lines = [f"{i+1}.{n}/{w}/{d}/{p}" for i, (n, w, d, p) in enumerate(info)]
            if lines:
                return "\n".join(lines), "extract_info(rows)", "regex_extract_info", ""
        except Exception:
            pass

    # 4) long_summary / insight / body fallback -> GPT with template-source prompt
    prompt = build_prompt_from_template(shape_name, role, pconf, budget, metrics, rows)

    if role == "insight":
        fallback = "1) 优化关键场景体验\n2) 强化稳定性一致性"
    elif role == "body":
        fallback = "反馈集中，建议围绕关键指标继续优化。"
    else:
        fallback = "样本反馈总体稳定，核心指标表现均衡。"

    txt = call_gpt(gpt_fn, prompt, pconf.get("model", "openai/gpt-5-mini"), fallback)
    gap = "" if txt and txt != fallback else "GPT未返回有效结果，已使用兜底文本"
    return txt, prompt, "gpt_prompted", gap


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
    prompt_trace = []
    gap_lines = [
        "# shape_data_gap_report",
        "",
        f"- 时间: {now_ts()}",
        f"- sheet: {sheet}",
        "",
        "|shape|role|strategy|gap|",
        "|---|---|---|---|",
    ]

    val_lines = [
        "# content_validation_report",
        "",
        f"- 时间: {now_ts()}",
        f"- sheet: {sheet}",
        f"- notes: {notes}",
        "",
        "|shape|role|strategy|len|line_count|max_chars|max_lines|valid|",
        "|---|---|---|---|---|---|---|---|",
    ]

    for m in mapping:
        name = m["shape_name"]
        role = m["role"]
        budget = _budget_for(name, budgets)
        pconf = _prompt_for(name, prompts)

        raw, used_prompt_or_reason, strategy, gap = build_content(role, name, metrics, m.get("template_text", ""), pconf, budget, funcs, rows)
        txt = clamp_text(raw, int(budget.get("max_chars", 80)), int(budget.get("max_lines", 4)))

        line_count = len(txt.splitlines()) if txt else 0
        valid = len(txt) <= int(budget.get("max_chars", 80)) and line_count <= int(budget.get("max_lines", 4))

        items.append({
            "shape_name": name,
            "role": role,
            "strategy": strategy,
            "build_function": m.get("build_function"),
            "content": txt,
            "budget": budget,
            "valid": valid,
            "prompt_model": pconf.get("model", "openai/gpt-5-mini"),
        })

        prompt_trace.append({
            "shape_name": name,
            "role": role,
            "strategy": strategy,
            "model": pconf.get("model", "openai/gpt-5-mini"),
            "prompt_or_reason": used_prompt_or_reason,
        })

        gap_lines.append(f"|{name}|{role}|{strategy}|{gap}|")
        val_lines.append(
            f"|{name}|{role}|{strategy}|{len(txt)}|{line_count}|{budget.get('max_chars')}|{budget.get('max_lines')}|{valid}|"
        )

    write_json(OUT_CONTENT, {"generated_at": now_ts(), "sheet": sheet, "items": items, "metrics": metrics})
    write_json(OUT_PROMPT_TRACE, {"generated_at": now_ts(), "prompts": prompt_trace})
    write_md(OUT_VALID, val_lines)
    write_md(OUT_GAP, gap_lines)
    print(f"[OK] wrote {OUT_CONTENT.name}, {OUT_VALID.name}, {OUT_PROMPT_TRACE.name}, {OUT_GAP.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
