#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 3A: Build shape content per strategy matrix.

Routes each shape to the correct builder based on user annotations
(生成方式 field from shape_detail.xlsx) and role.

Strategy routing (priority order):
  annotation "10分" + "评分均值"  -> score_10pt   (Python, no GPT)
  annotation "100分制" or "档"    -> grade_letter  (Python, no GPT)
  annotation "numeric_aggregation"/"不走GPT" -> sample_aggregation (Python, no GPT)
  annotation "鞋款名称"           -> extract_shoe_name (Python, no GPT)
  annotation "gpt_prompted"       -> gpt_rich      (GPT_5 with full questionnaire text)
  role == "chart"                 -> mean_extraction (Python, no GPT)
  role == "title"                 -> template_direct (Python, no GPT)
  role == "sample_stat"           -> sample_aggregation (Python, no GPT)
  default                         -> gpt_prompted  (GPT_5 with basic prompt)

GPT_5 is imported directly from src.Function_030 (not through load_legacy_functions).
Model: openai/gpt-5.2 (OpenRouter)
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Tuple

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    EXCEL_PATH,
    PROGRESS_DIR,
    SHAPE_DETAIL_XLSX,
    clamp_text,
    extract_metrics,
    extract_score_means,
    load_excel_rows,
    now_ts,
    numeric,
    parse_params,
    read_gpt_prompts_from_xlsx,
    safe_print,
    safe_text,
    setup_console_encoding,
    write_gpt_prompts_to_xlsx,
    write_json,
    write_md,
)

# ---------------------------------------------------------------------------
# Direct import of GPT_5 from src.Function_030 — show error if it fails
# ---------------------------------------------------------------------------
GPT_5 = None
_GPT_AVAILABLE = False
try:
    from src.Function_030 import GPT_5  # type: ignore
    _GPT_AVAILABLE = True
except Exception as _e:
    safe_print(f"[WARN] GPT_5 import failed: {_e}")
    safe_print("[WARN] GPT-dependent shapes will use fallback text.")

MODEL = "openai/gpt-5.4"

# ---------------------------------------------------------------------------
# File paths
# ---------------------------------------------------------------------------
MAP_JSON          = PROGRESS_DIR / "02-shape_analysis_map.json"
PROMPT_JSON       = PROGRESS_DIR / "02-prompt_specs.json"
BUDGET_JSON       = PROGRESS_DIR / "02-readability_budget.json"
SHAPE_DETAIL_JSON = PROGRESS_DIR / "01-shape_detail_com.json"
OUT_CONTENT         = PROGRESS_DIR / "03a-build_shape_content.json"
OUT_VALID           = PROGRESS_DIR / "03a-content_validation_report.md"
OUT_PROMPT_TRACE    = PROGRESS_DIR / "03a-prompt_trace.json"
OUT_GAP             = PROGRESS_DIR / "03a-shape_data_gap_report.md"
OUT_PENDING_PROMPTS = PROGRESS_DIR / "03a-pending_prompts.json"


def _load_shape_types() -> dict:
    """Load {shape_name: shape_type} from shape_detail_com.json.
    Used to skip image/non-text shapes (shape_type 13 = picture)."""
    if not SHAPE_DETAIL_JSON.exists():
        return {}
    try:
        data = json.loads(SHAPE_DETAIL_JSON.read_text(encoding="utf-8"))
        shapes = data.get("new_shapes", [])
        return {s.get("name", ""): s.get("shape_type", 1) for s in shapes}
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Image extraction helper
# ---------------------------------------------------------------------------

def _extract_excel_image(sheet_hint: str = "问卷") -> str:
    """Extract the first embedded image from the Excel questionnaire sheet via COM.

    Supports encrypted files. Uses CopyPicture + temporary ChartObject to export.
    Returns the absolute path to the saved image file (in PROGRESS_DIR),
    or "" if no image is found.
    """
    import win32com.client
    import time as _time

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(EXCEL_PATH.resolve()), 0, True)

        # Find sheet
        ws = None
        for i in range(1, wb.Sheets.Count + 1):
            if sheet_hint in wb.Sheets(i).Name:
                ws = wb.Sheets(i)
                break
        if ws is None:
            safe_print("[WARN] extract_image: 未找到问卷sheet")
            wb.Close(False)
            return ""

        # Find first picture shape (Type 13 = msoPicture)
        pic = None
        for i in range(1, ws.Shapes.Count + 1):
            try:
                if ws.Shapes(i).Type == 13:
                    pic = ws.Shapes(i)
                    break
            except Exception:
                continue

        if pic is None:
            safe_print("[WARN] extract_image: 问卷sheet中没有嵌入图片")
            wb.Close(False)
            return ""

        out_path = str((PROGRESS_DIR / "03a-shoe_image.png").resolve())

        # Export via temporary ChartObject
        pic.CopyPicture(1, 2)  # xlScreen=1, xlBitmap=2
        _time.sleep(0.3)
        co = ws.ChartObjects().Add(0, 0, pic.Width + 10, pic.Height + 10)
        co.Chart.Paste()
        _time.sleep(0.3)
        co.Chart.Export(out_path, "PNG")
        co.Delete()

        wb.Close(False)
        safe_print("[OK] extract_image: 已提取图片 -> 03a-shoe_image.png")
        return out_path
    except Exception as e:
        safe_print(f"[WARN] extract_image COM error: {e}")
        return ""
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Data extraction helpers
# ---------------------------------------------------------------------------

def _col_values(rows: List[List[Any]], *keywords: str) -> List[Any]:
    """Return all non-None values from the first column whose header
    contains any of the given keywords."""
    if not rows:
        return []
    headers = [safe_text(h) for h in rows[0]]
    for kw in keywords:
        for idx, h in enumerate(headers):
            if kw in h:
                return [
                    row[idx]
                    for row in rows[1:]
                    if idx < len(row) and row[idx] is not None
                    and safe_text(row[idx]) != ""
                ]
    return []


def _score_10pt(rows: List[List[Any]]) -> float | None:
    """Calculate the overall mean score from score columns,
    auto-detects 5-scale vs 10-scale, returns a 10-point value."""
    means = extract_score_means(rows)
    if not means:
        return None
    overall = sum(v for _, v in means) / len(means)
    # If all column means are <= 5.5, assume 5-scale → multiply by 2
    max_mean = max(v for _, v in means)
    if max_mean <= 5.5:
        return round(overall * 2, 2)
    return round(overall, 2)


def _score_to_grade(score_10: float) -> str:
    """Convert a 10-point score to a letter grade (100-point scale internally)."""
    s = score_10 * 10  # convert to 100-point
    if s >= 95: return "S+"
    if s >= 90: return "S-"
    if s >= 85: return "A+"
    if s >= 80: return "A-"
    if s >= 75: return "B+"
    if s >= 70: return "B-"
    if s >= 65: return "C+"
    return "C-"


def _sample_stat_text(rows: List[List[Any]]) -> str:
    """Build sample stat text: 试穿人数 / 平均体重 / 球场定位."""
    count = max(0, len(rows) - 1)

    # 平均体重
    weights = _col_values(rows, "体重", "Weight", "重量")
    valid_w = [float(w) for w in weights if numeric(w) is not None]
    avg_w = round(sum(valid_w) / len(valid_w), 1) if valid_w else None

    # 球场定位 / 打法
    positions = _col_values(rows, "球场定位", "打法", "定位", "位置")
    # Flatten compound values like "速度型后卫" → keep unique
    pos_clean: List[str] = []
    seen = set()
    for p in positions:
        s = safe_text(p)
        if s and s not in seen:
            seen.add(s)
            pos_clean.append(s)

    lines = [f"试穿人数：{count}人"]
    if avg_w is not None:
        lines.append(f"测试者平均体重：{avg_w}KG")
    if pos_clean:
        lines.append(f"测试者球场定位：{'、'.join(pos_clean)}")
    return "\n".join(lines)


def _shoe_name(rows: List[List[Any]], col_hint: str = "") -> str:
    """Extract shoe name from Excel column.

    col_hint: column keyword from params (e.g. "鞋款名称"). Tried first if provided.
    """
    keywords = ([col_hint] if col_hint else []) + ["鞋款名称", "鞋款", "试穿"]
    names = _col_values(rows, *keywords)
    unique = list(dict.fromkeys(safe_text(n) for n in names if n))
    return unique[0] if unique else ""


def _questionnaire_comments(rows: List[List[Any]]) -> List[str]:
    """Extract all non-empty values from 补充说明 column."""
    raw = _col_values(rows, "补充说明", "备注", "意见", "评价")
    return [safe_text(c) for c in raw if safe_text(c)]


# ---------------------------------------------------------------------------
# Budget / prompt lookup
# ---------------------------------------------------------------------------

def _budget_for(name: str, budgets: List[dict]) -> dict:
    for b in budgets:
        if b.get("shape_name") == name:
            return b
    return {"max_chars": 80, "max_lines": 4, "max_bullets": 4}


def _prompt_for(name: str, prompts: List[dict]) -> dict:
    for p in prompts:
        if p.get("shape_name") == name:
            return p
    return {"model": MODEL, "instruction": "", "style_anchor": ""}


# ---------------------------------------------------------------------------
# GPT call wrapper
# ---------------------------------------------------------------------------

def _call_gpt(prompt: str, fallback: str) -> Tuple[str, bool]:
    """Call GPT_5 with MODEL. Returns (text, used_fallback)."""
    if not _GPT_AVAILABLE or GPT_5 is None:
        return fallback, True
    try:
        result = safe_text(GPT_5(prompt, MODEL))
        if result:
            return result, False
    except Exception as e:
        safe_print(f"[WARN] GPT_5 call error: {e}")
    return fallback, True


# ---------------------------------------------------------------------------
# GPT rich prompt builder (for long_summary / body with gpt_prompted)
# Follows gen_mc_prompt style: 参考文本 first, structured per-respondent data,
# explicit (X/N) ratio requirement, strict format/tone instructions.
# ---------------------------------------------------------------------------

# Column matching keywords (dynamic — works across different Excel formats)
_NAME_KEYWORDS = ["姓名", "name"]
_WEIGHT_KEYWORDS = ["体重", "weight"]
_SKIP_KEYWORDS = ["鞋款", "轮次", "身高", "累计", "场次"]  # metadata, not feedback


def _find_col(headers: List[str], keywords: List[str]) -> Optional[str]:
    """Find first header containing any keyword (case-insensitive)."""
    for h in headers:
        hl = h.lower()
        for kw in keywords:
            if kw.lower() in hl:
                return h
    return None


def _classify_columns(headers: List[str], rows: List[List[Any]]
                      ) -> Tuple[List[str], List[str]]:
    """Dynamically classify columns into score (numeric) and text (feedback).

    Score columns: >70% of data rows have numeric values in 1-10 range.
    Text columns:  >30% of data rows have string values longer than 5 chars.
    """
    name_col = _find_col(headers, _NAME_KEYWORDS)
    weight_col = _find_col(headers, _WEIGHT_KEYWORDS)
    skip_cols = {name_col, weight_col} | {
        h for h in headers
        if any(kw in h for kw in _SKIP_KEYWORDS)
    }

    score_cols, text_cols = [], []
    n_data = len(rows) - 1
    if n_data < 1:
        return score_cols, text_cols

    for ci, h in enumerate(headers):
        if h in skip_cols:
            continue
        nums, texts = 0, 0
        for row in rows[1:]:
            val = row[ci] if ci < len(row) else None
            if val is None:
                continue
            try:
                v = float(val)
                if 0 <= v <= 10:
                    nums += 1
            except (ValueError, TypeError):
                if len(str(val).strip()) > 5:
                    texts += 1
        if nums / n_data > 0.7:
            score_cols.append(h)
        elif texts / n_data > 0.3:
            text_cols.append(h)

    return score_cols, text_cols


def _build_respondent_block(rows: List[List[Any]]) -> Tuple[str, int]:
    """Build a per-respondent data block for inclusion in GPT prompt.
    Dynamically matches columns — works with any Excel format.
    Returns (block_text, respondent_count)."""
    if not rows or len(rows) < 2:
        return "（无数据）", 0

    headers = [safe_text(h) for h in rows[0]]
    n = len(rows) - 1

    name_col = _find_col(headers, _NAME_KEYWORDS)
    weight_col = _find_col(headers, _WEIGHT_KEYWORDS)
    score_cols, text_cols = _classify_columns(headers, rows)

    blocks = []
    for i, row in enumerate(rows[1:], 1):
        fd: Dict[str, str] = {}
        for j, h in enumerate(headers):
            if j < len(row):
                fd[h] = safe_text(row[j])

        name = fd.get(name_col, f"受访者{i}") if name_col else f"受访者{i}"
        weight = fd.get(weight_col, "") if weight_col else ""

        scores = ", ".join(
            f"{h.split('（')[0] if '（' in h else h}={fd[h]}"
            for h in score_cols if fd.get(h)
        )
        feedbacks = "\n  ".join(
            f"{h.split('（')[0] if '（' in h else h}: {fd[h]}"
            for h in text_cols if fd.get(h)
        )

        weight_str = f"  体重:{weight}KG" if weight else ""
        parts = [f"【受访者{i}】{name}{weight_str}"]
        if scores:
            parts.append(f"  各项评分: {scores}")
        if feedbacks:
            parts.append(f"  试穿反馈:\n  {feedbacks}")
        blocks.append("\n".join(parts))

    return "\n\n".join(blocks), n


_PROMPT_TEMPLATE_PATH = Path(__file__).resolve().parent / "prompt_templates" / "gpt_summary.md"


def _load_prompt_template(mode: str) -> str | None:
    """Load a prompt template section from gpt_summary.md by mode name.

    Returns the template string (with {placeholders}), or None if file
    is missing or the mode section is not found.
    """
    if not _PROMPT_TEMPLATE_PATH.exists():
        return None
    try:
        raw = _PROMPT_TEMPLATE_PATH.read_text(encoding="utf-8")
    except Exception:
        return None

    marker = f"<!-- MODE: {mode} -->"
    start = raw.find(marker)
    if start < 0:
        return None
    start += len(marker)

    # Find the next MODE marker (or end of file)
    next_marker = raw.find("<!-- MODE:", start)
    section = raw[start:next_marker].strip() if next_marker >= 0 else raw[start:].strip()
    return section if section else None


def _build_rich_prompt(
    shape_name: str,
    role: str,
    style_anchor: str,
    content_source: str,
    fix_notes: str,
    budget: dict,
    rows: List[List[Any]],
    focus: str = "",
    user_instruction: str = "",
    output_contract: dict = None,
) -> str:
    """Build GPT prompt for body/long_summary shapes.

    focus (str): if set (e.g. '优点' or '缺点'), use free-form summarization mode —
        GPT decides the category structure. If empty, use the legacy categorized mode
        where content_source dictates what to summarize.

    Reads prompt template from pipeline/prompt_templates/gpt_summary.md.
    Falls back to hardcoded prompt if the template file is missing.
    """
    respondent_block, n = _build_respondent_block(rows)
    max_chars = budget.get("max_chars", 200)
    max_lines = budget.get("max_lines", 6)
    target_chars = max_chars

    # Build user instruction section (independent paragraph, high weight)
    user_section = ""
    if user_instruction:
        user_section = f"\n用户指令：\n{user_instruction}\n"

    # Build output contract section
    contract_section = ""
    if output_contract:
        lines = []
        kw = output_contract.get("required_keywords")
        if kw:
            lines.append(f"- 必须包含关键词: {'、'.join(kw)}")
        if output_contract.get("bracket_highlight"):
            lines.append("- 关键性能词用【】括起（仅括词本身，不含标点）")
        if output_contract.get("ratio_required"):
            lines.append("- 每段结论后注明 (X/N) 比例")
        if lines:
            contract_section = "\n输出合约（必须满足）：\n" + "\n".join(lines) + "\n"

    # Try loading external template
    mode = "free_form" if focus else "categorized"
    template = _load_prompt_template(mode)
    if template is not None:
        try:
            return template.format(
                style_anchor=style_anchor,
                n=n,
                focus=focus,
                content_source=content_source,
                target_chars=target_chars,
                max_lines=max_lines,
                user_instruction=user_section,
                contract_section=contract_section,
                respondent_block=respondent_block,
            )
        except (KeyError, IndexError) as e:
            safe_print(f"[WARN] prompt template format error: {e}, using fallback")

    # Fallback: hardcoded prompt (identical to original logic)
    if focus:
        task_line = (
            f"请从{n}名测试者的实际反馈中，自由归纳这款篮球鞋的【{focus}】。\n"
            f"根据实际反馈内容自行决定分段维度，不要按固定性能类别（如包裹性、止滑性等）分类。\n"
            f"在每条结论后注明（X/{n}）表示几分之几的测试者有此体验。\n"
            f"每段结论中，请将最核心的1-2个关键性能词用【】括起来（仅括词本身，不含标点），"
            f"例如：【止滑性】在室内场地表现稳定（2/{n}）。这些关键词后续会自动高亮显示。"
        )
        format_note = "- 参考文本仅作语调参考，不必复制其分类结构\n"
    else:
        task_line = (
            f"下面是{n}名测试者对这款篮球鞋的原始试穿反馈，"
            f"请帮我按分类汇总其中的【{content_source}】。\n"
            f"在每条结论后注明（X/{n}）表示几分之几的测试者有此反馈。"
        )
        format_note = "- 严格按照参考文本的格式、语调、陈述方式\n"

    return (
        f"1、参考文本（参考语调和信息密度）\n{style_anchor}\n\n"
        f"2、你的任务\n{task_line}\n"
        f"{user_section}"
        f"{contract_section}\n"
        f"注意：\n"
        f"- 你只能分析已有数据，不能推测或编造\n"
        f"- 直接给出结论，不要展示分析过程\n"
        f"- {format_note}"
        f"- 总字数控制在{target_chars}字左右\n"
        f"- 不超过{max_lines}行\n"
        f"- 每个分类不超过3行\n"
        f"- 关键词要求：开头第一句必须同时包含'样本'和'反馈'（如\"本次{n}名样本反馈显示\"），正文中自然融入'建议'（如\"建议关注...\"）\n\n"
        f"3、{n}名测试者原始反馈\n{respondent_block}\n\n"
        f"直接输出结论，不需要任何前言。"
    )


# ---------------------------------------------------------------------------
# Main content builder — routes by strategy_hint first, then role
# ---------------------------------------------------------------------------

def build_content(
    role: str,
    shape_name: str,
    metrics: Dict[str, Any],
    tmpl_text: str,
    pconf: dict,
    budget: dict,
    strategy_hint: str,
    content_source: str,
    rows: List[List[Any]],
    shape_type: int = 1,
    strategy_exact: str = "",
    params: dict = None,
    user_instruction: str = "",
    dry_run: bool = False,
    override_prompt: str = "",
    output_contract: dict = None,
) -> Tuple[str, str, str, str]:
    """Build content for one shape.

    Returns: (text, prompt_or_reason, strategy_label, gap_reason)

    Routing priority:
      1. strategy_exact (structured field from shape_detail.xlsx) — exact dispatch
      2. strategy_hint keyword matching (legacy natural-language fallback)
      3. role-based defaults
    """
    if params is None:
        params = {}

    # Skip non-text shapes (pictures, media, etc.) unless extract_image is set
    # shape_type 13 = picture, 15 = video, 16 = audio, 24 = web frame
    NON_TEXT_TYPES = {13, 15, 16, 24}
    if shape_type in NON_TEXT_TYPES and strategy_exact != "extract_image":
        return "", "skip_non_text_shape", "skip", f"shape_type={shape_type} has no text frame"

    hint = strategy_hint  # original (not lowercased — Chinese)
    hint_l = hint.lower()
    style_anchor = safe_text(pconf.get("style_anchor", ""))
    fix_notes = ""
    instruction = safe_text(pconf.get("instruction", ""))
    # Extract fix_notes from instruction if it was appended by Step 2
    if "[用户修正]" in instruction:
        parts = instruction.split("[用户修正]", 1)
        fix_notes = parts[1].strip()

    # -----------------------------------------------------------------------
    # 0. EXACT dispatch — strategy field takes absolute priority
    #    No keyword matching needed; code is an unambiguous instruction.
    # -----------------------------------------------------------------------
    if strategy_exact:
        s = strategy_exact

        if s == "skip":
            return "", "explicit_skip", "skip", ""

        if s == "template_direct":
            return tmpl_text or "", "template_direct", "template_direct", ""

        if s == "grade_letter":
            score = _score_10pt(rows)
            if score is not None:
                return _score_to_grade(score), "grade_letter_calc", "grade_letter", ""
            return tmpl_text or "B+", "grade_letter_fallback", "grade_letter", "无法计算评分均值"

        if s == "score_10pt":
            score = _score_10pt(rows)
            if score is not None:
                return f"{score}/10", "score_10pt_calc", "score_10pt", ""
            return tmpl_text or "-.--/10", "score_10pt_fallback", "score_10pt", "无法计算评分均值"

        if s == "sample_aggregation":
            text = _sample_stat_text(rows)
            gap = "" if text else "无法提取样本统计数据"
            return text or f"试穿人数：{metrics.get('respondent_count', 0)}人", \
                   "sample_stat_aggregation", "sample_aggregation", gap

        if s == "extract_column":
            col = params.get("column", "")
            name = _shoe_name(rows, col_hint=col)
            gap = "" if name else f"未找到列: {col or '鞋款名称'}"
            return name or tmpl_text or "未知鞋款", "extract_column_exact", "extract_column", gap

        if s == "gpt_prompted":
            src = params.get("source", content_source or "补充说明")
            flt = params.get("filter", "")
            # When filter (优点/缺点) is set, use free-form summarization mode.
            # Pass focus=flt; effective_src is just the data source without category suffix.
            effective_src = src
            fallback_map = {
                "long_summary": "样本反馈总体稳定，核心指标表现均衡。",
                "body": "反馈集中，建议围绕关键指标继续优化。",
            }
            fallback = fallback_map.get(role, fallback_map["body"])
            prompt = override_prompt or _build_rich_prompt(
                shape_name, role, style_anchor, effective_src, fix_notes, budget, rows,
                focus=flt, user_instruction=user_instruction, output_contract=output_contract,
            )
            if dry_run:
                return fallback, prompt, "gpt_prompted", ""
            txt, used_fb = _call_gpt(prompt, fallback)
            gap = "GPT未返回有效结果，已使用兜底文本" if used_fb else ""
            return txt, prompt, "gpt_prompted", gap

        if s == "mean_extraction":
            means = extract_score_means(rows)
            if not means:
                return "减震:0\n回弹:0\n稳定:0", "chart_mean_fallback", "mean_extraction", "未识别到可用评分列"
            return "\n".join(f"{k}:{v:.2f}" for k, v in means[:8]), \
                   "chart_mean_from_source", "mean_extraction", ""

        if s == "extract_image":
            sheet_hint = params.get("sheet", "问卷")
            img_path = _extract_excel_image(sheet_hint)
            if img_path:
                return f"IMAGE:{img_path}", "extract_image_ok", "extract_image", ""
            return "IMAGE_MISSING:请手动插入产品图片", "extract_image_failed", "extract_image", "未能从Excel问卷sheet提取图片"

        # Unknown strategy_exact code — fall through to keyword matching below
        safe_print(f"[WARN] Unknown strategy_exact='{s}' for '{shape_name}', falling back to hint")

    # -----------------------------------------------------------------------
    # 1. Grade letter  e.g. "A-", "S+"
    # NOTE: must check BEFORE score_10pt because "100分制" contains "10分"
    # -----------------------------------------------------------------------
    if "100分制" in hint or ("档" in hint and "评分" in hint):
        score = _score_10pt(rows)
        if score is not None:
            return _score_to_grade(score), "grade_letter_calc", "grade_letter", ""
        return tmpl_text or "B+", "grade_letter_fallback", "grade_letter", "无法计算评分均值"

    # -----------------------------------------------------------------------
    # 2. Score mean → 10-point format  e.g. "8.29/10"
    # -----------------------------------------------------------------------
    if "10分" in hint and "评分均值" in hint:
        score = _score_10pt(rows)
        if score is not None:
            return f"{score}/10", "score_10pt_calc", "score_10pt", ""
        return tmpl_text or "-.--/10", "score_10pt_fallback", "score_10pt", "无法计算评分均值"

    # -----------------------------------------------------------------------
    # 3. Sample stat aggregation (no GPT)
    # -----------------------------------------------------------------------
    if "numeric_aggregation" in hint_l or "不走gpt" in hint_l or "不走GPT" in hint:
        text = _sample_stat_text(rows)
        gap = "" if text else "无法提取样本统计数据"
        return text or f"试穿人数：{metrics.get('respondent_count', 0)}人", \
               "sample_stat_aggregation", "sample_aggregation", gap

    # -----------------------------------------------------------------------
    # 4. Extract shoe name from column
    # -----------------------------------------------------------------------
    if "鞋款名称" in hint or ("提取" in hint and "gpt" not in hint_l and "gpt_prompted" not in hint_l):
        name = _shoe_name(rows)
        gap = "" if name else "未找到鞋款名称列"
        return name or tmpl_text or "未知鞋款", "shoe_name_from_excel", "extract_column", gap

    # -----------------------------------------------------------------------
    # 5. GPT rich prompt (long_summary with actual questionnaire text)
    # -----------------------------------------------------------------------
    if "gpt_prompted" in hint_l:
        fallback_map = {
            "long_summary": "样本反馈总体稳定，核心指标表现均衡。",
            "body": "反馈集中，建议围绕关键指标继续优化。",
        }
        fallback = fallback_map.get(role, fallback_map["body"])
        prompt = override_prompt or _build_rich_prompt(
            shape_name, role, style_anchor, content_source, fix_notes, budget, rows,
            user_instruction=user_instruction, output_contract=output_contract,
        )
        if dry_run:
            return fallback, prompt, "gpt_prompted", ""
        txt, used_fb = _call_gpt(prompt, fallback)
        gap = "GPT未返回有效结果，已使用兜底文本" if used_fb else ""
        return txt, prompt, "gpt_prompted", gap

    # -----------------------------------------------------------------------
    # 6. chart role → deterministic mean extraction
    # -----------------------------------------------------------------------
    if role == "chart":
        means = extract_score_means(rows)
        if not means:
            return "减震:0\n回弹:0\n稳定:0", "chart_mean_fallback", "mean_extraction", "未识别到可用评分列"
        return "\n".join(f"{k}:{v:.2f}" for k, v in means[:8]), \
               "chart_mean_from_source", "mean_extraction", ""

    # -----------------------------------------------------------------------
    # 7. Empty template + no annotation → leave blank (structural/decorative shape)
    # -----------------------------------------------------------------------
    if not tmpl_text and not hint and role in {"body", "title"}:
        return "", "empty_template_no_hint", "skip_empty", ""

    # -----------------------------------------------------------------------
    # 8. title role → template direct (no GPT)
    # -----------------------------------------------------------------------
    if role == "title":
        title = tmpl_text or "问卷分析报告"
        gap = "" if tmpl_text else "模板标题为空，使用默认标题"
        return title, "title_from_template", "template_direct", gap

    # -----------------------------------------------------------------------
    # 8. sample_stat role (no annotation) → basic aggregation
    # -----------------------------------------------------------------------
    if role == "sample_stat":
        text = _sample_stat_text(rows)
        return text or f"试穿人数：{metrics.get('respondent_count', 0)}人", \
               "sample_stat_aggregation", "sample_aggregation", ""

    # -----------------------------------------------------------------------
    # 9. body / long_summary / insight → rich prompt with respondent data
    #    (matches codex_ppt.py behavior — always use full questionnaire data)
    # -----------------------------------------------------------------------
    if role in ("long_summary", "body", "insight"):
        fallback_map = {
            "long_summary": "样本反馈总体稳定，核心指标表现均衡。",
            "body": "反馈集中，建议围绕关键指标继续优化。",
            "insight": "1) 优化关键场景体验\n2) 强化稳定性一致性",
        }
        fallback = fallback_map.get(role, fallback_map["body"])
        prompt = override_prompt or _build_rich_prompt(
            shape_name, role, style_anchor,
            content_source or "补充说明",
            fix_notes, budget, rows,
            user_instruction=user_instruction, output_contract=output_contract,
        )
        if dry_run:
            return fallback, prompt, "gpt_prompted", ""
        txt, used_fb = _call_gpt(prompt, fallback)
        gap = "GPT未返回有效结果，已使用兜底文本" if used_fb else ""
        return txt, prompt, "gpt_prompted", gap

    # -----------------------------------------------------------------------
    # 10. Other roles → basic prompt (no questionnaire data needed)
    # -----------------------------------------------------------------------
    prompt = override_prompt or (
        "你是PPT内容工程师。依据源数据生成用于单个shape的最终文案。\n"
        "硬约束：不得编造、不得输出解释、不得使用markdown。\n"
        f"shape角色：{role}\n"
        f"风格锚点（模板原文）：{style_anchor}\n"
        f"写作要求：{instruction}\n"
        f"输出约束：max_chars={budget.get('max_chars', 80)}, "
        f"max_lines={budget.get('max_lines', 4)}\n"
        f"数据统计：respondent_count={metrics.get('respondent_count')}, "
        f"keywords={metrics.get('keywords', [])}\n"
        "请直接输出最终文本。"
    )
    fallback = "反馈集中，建议围绕关键指标继续优化。"
    if dry_run:
        return fallback, prompt, "gpt_prompted", ""
    txt, used_fb = _call_gpt(prompt, fallback)
    gap = "GPT未返回有效结果，已使用兜底文本" if used_fb else ""
    return txt, prompt, "gpt_prompted", gap


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def _build_all(mapping, prompts_spec, budgets, rows, metrics, shape_types,
               dry_run=False, prompt_overrides=None):
    """Core build loop. Returns (items, prompt_trace, gap_lines, val_lines, pending).

    dry_run=True: GPT shapes return fallback text + assembled prompt (no GPT call).
    prompt_overrides: {shape_name: prompt_text} — use these instead of assembling.
    pending: list of dicts for GPT shapes (only populated when dry_run=True).
    """
    prompt_overrides = prompt_overrides or {}
    sheet = ""
    items = []
    prompt_trace = []
    pending = []
    gap_lines = [
        "# shape_data_gap_report", "",
        f"- 时间: {now_ts()}", "",
        "|shape|role|strategy|gap|", "|---|---|---|---|",
    ]
    val_lines = [
        "# content_validation_report", "",
        f"- 时间: {now_ts()}",
        f"- GPT_5_available: {_GPT_AVAILABLE}", "",
        "|shape|role|strategy|len|lines|max_chars|max_lines|valid|",
        "|---|---|---|---|---|---|---|---|",
    ]

    for m in mapping:
        name           = m["shape_name"]
        role           = m["role"]
        tmpl_text      = safe_text(m.get("template_text", ""))
        strategy_hint  = safe_text(m.get("user_strategy_hint", ""))
        content_source = safe_text(m.get("user_content_source", ""))
        strategy_exact = safe_text(m.get("strategy_exact", ""))
        params         = parse_params(safe_text(m.get("params", "")))
        ui             = safe_text(m.get("user_instruction", ""))
        budget         = _budget_for(name, budgets)
        pconf          = _prompt_for(name, prompts_spec)
        shape_type     = shape_types.get(name, 1)

        override = prompt_overrides.get(name, "")
        oc = m.get("output_contract", {})

        raw, used_prompt_or_reason, strategy, gap = build_content(
            role, name, metrics, tmpl_text,
            pconf, budget, strategy_hint, content_source, rows,
            shape_type=shape_type,
            strategy_exact=strategy_exact,
            params=params,
            user_instruction=ui,
            dry_run=dry_run,
            override_prompt=override,
            output_contract=oc,
        )

        # Chart data and image paths are structured — never clamp them.
        if strategy in {"mean_extraction", "extract_image"}:
            txt = raw
        else:
            txt = clamp_text(raw, int(budget.get("max_chars", 80)), int(budget.get("max_lines", 4)))

        line_count = len(txt.splitlines()) if txt else 0
        valid = (
            len(txt) <= int(budget.get("max_chars", 80))
            and line_count <= int(budget.get("max_lines", 4))
        )

        # Determine post-write keyword color (优点=red, 缺点=blue)
        color_hint = ""
        if strategy_exact == "gpt_prompted":
            flt = params.get("filter", "")
            if flt == "优点":
                color_hint = "red"
            elif flt == "缺点":
                color_hint = "blue"

        items.append({
            "shape_name": name,
            "role": role,
            "strategy": strategy,
            "build_function": m.get("build_function"),
            "content": txt,
            "color_hint": color_hint,
            "budget": budget,
            "valid": valid,
            "prompt_model": MODEL,
        })

        prompt_trace.append({
            "shape_name": name,
            "role": role,
            "strategy": strategy,
            "strategy_hint": strategy_hint,
            "model": MODEL,
            "prompt_or_reason": used_prompt_or_reason
            if len(str(used_prompt_or_reason)) < 200
            else str(used_prompt_or_reason)[:200] + "...",
        })

        # Collect pending GPT prompts (dry_run only)
        if dry_run and strategy == "gpt_prompted":
            pending.append({
                "shape_name": name,
                "role": role,
                "strategy": strategy,
                "prompt": used_prompt_or_reason,  # full prompt text
                "fallback": txt,
                "budget": budget,
            })

        gap_lines.append(f"|{name}|{role}|{strategy}|{gap}|")
        val_lines.append(
            f"|{name}|{role}|{strategy}|{len(txt)}|{line_count}"
            f"|{budget.get('max_chars')}|{budget.get('max_lines')}|{valid}|"
        )

        safe_print(f"  [{strategy:22s}] {name}: {txt[:60].replace(chr(10), ' ')}")

    return items, prompt_trace, gap_lines, val_lines, pending


def main() -> int:
    import argparse
    setup_console_encoding()

    ap = argparse.ArgumentParser(description="Step 3A: Build shape content")
    ap.add_argument("--assemble-only", action="store_true",
                    help="Phase 1: assemble GPT prompts without calling GPT")
    ap.add_argument("--execute-prompts", action="store_true",
                    help="Phase 2: read reviewed prompts and call GPT")
    ap.add_argument("--interactive", action="store_true",
                    help="Standalone: Phase 1 + open Excel + pause + Phase 2")
    args = ap.parse_args()

    if not (MAP_JSON.exists() and PROMPT_JSON.exists() and BUDGET_JSON.exists()):
        safe_print("[BLOCKED] Analysis artifacts missing. Run 02_shape_analysis.py first.")
        return 0

    mapping = json.loads(MAP_JSON.read_text(encoding="utf-8")).get("mapping", [])
    prompts_spec = json.loads(PROMPT_JSON.read_text(encoding="utf-8")).get("prompts", [])
    budgets = json.loads(BUDGET_JSON.read_text(encoding="utf-8")).get("budgets", [])

    rows, sheet, notes = load_excel_rows("问卷sheet")
    metrics = extract_metrics(rows)
    shape_types = _load_shape_types()

    safe_print(f"[INFO] GPT_5 available: {_GPT_AVAILABLE}  model: {MODEL}")
    safe_print(f"[INFO] Excel sheet: {sheet}, rows: {len(rows)-1}")

    # ------------------------------------------------------------------
    # Mode: --assemble-only  (Phase 1)
    # ------------------------------------------------------------------
    if args.assemble_only:
        items, pt, gl, vl, pending = _build_all(
            mapping, prompts_spec, budgets, rows, metrics, shape_types,
            dry_run=True,
        )
        # Save pending prompts (full text, no truncation)
        write_json(OUT_PENDING_PROMPTS, {
            "generated_at": now_ts(),
            "pending": pending,
        })
        # Write prompts to Excel "GPT-prompt Text" cells
        if pending:
            xlsx_prompts = {p["shape_name"]: p["prompt"] for p in pending}
            write_gpt_prompts_to_xlsx(xlsx_prompts)
        # Write Phase 1 content (fallback for GPT shapes)
        write_json(OUT_CONTENT, {"generated_at": now_ts(), "sheet": sheet,
                                 "items": items, "metrics": metrics})
        write_json(OUT_PROMPT_TRACE, {"generated_at": now_ts(), "model": MODEL, "prompts": pt})
        write_md(OUT_VALID, vl)
        write_md(OUT_GAP, gl)
        safe_print(f"[OK] Phase 1: {len(items)} shapes built, {len(pending)} GPT prompts pending review.")
        return 0

    # ------------------------------------------------------------------
    # Mode: --execute-prompts  (Phase 2)
    # ------------------------------------------------------------------
    if args.execute_prompts:
        # Load original prompts from JSON
        prompt_overrides: dict[str, str] = {}
        if OUT_PENDING_PROMPTS.exists():
            pd = json.loads(OUT_PENDING_PROMPTS.read_text(encoding="utf-8"))
            for p in pd.get("pending", []):
                prompt_overrides[p["shape_name"]] = p["prompt"]
        # Read back from Excel (user may have edited)
        xlsx_prompts = read_gpt_prompts_from_xlsx()
        for name, edited in xlsx_prompts.items():
            if edited.strip():
                prompt_overrides[name] = edited  # user edit takes priority

        if prompt_overrides:
            safe_print(f"[INFO] Phase 2: {len(prompt_overrides)} GPT prompts loaded (Excel edits take priority)")

        items, pt, gl, vl, _ = _build_all(
            mapping, prompts_spec, budgets, rows, metrics, shape_types,
            dry_run=False,
            prompt_overrides=prompt_overrides,
        )
        write_json(OUT_CONTENT, {"generated_at": now_ts(), "sheet": sheet,
                                 "items": items, "metrics": metrics})
        write_json(OUT_PROMPT_TRACE, {"generated_at": now_ts(), "model": MODEL, "prompts": pt})
        write_md(OUT_VALID, vl)
        write_md(OUT_GAP, gl)
        safe_print(f"[OK] Phase 2: {len(items)} shapes built. Wrote {OUT_CONTENT.name}")
        return 0

    # ------------------------------------------------------------------
    # Mode: --interactive  (Phase 1 + pause + Phase 2, for manual pipeline)
    # ------------------------------------------------------------------
    if args.interactive:
        import os
        # Phase 1
        items, pt, gl, vl, pending = _build_all(
            mapping, prompts_spec, budgets, rows, metrics, shape_types,
            dry_run=True,
        )
        write_json(OUT_PENDING_PROMPTS, {
            "generated_at": now_ts(), "pending": pending,
        })
        if pending:
            xlsx_prompts = {p["shape_name"]: p["prompt"] for p in pending}
            write_gpt_prompts_to_xlsx(xlsx_prompts)
            safe_print(f"\n[INFO] {len(pending)} 个 GPT prompt 已写入 Excel")
            # Open Excel for user review
            try:
                os.startfile(str(SHAPE_DETAIL_XLSX))
                safe_print("[已打开] Excel — 请编辑「GPT-prompt Text」单元格")
            except Exception:
                safe_print(f"[手动打开] {SHAPE_DETAIL_XLSX}")
            safe_print("编辑完成后，保存并关闭 Excel，然后按 Enter 继续...")
            input()
            # Phase 2: read back and call GPT
            prompt_overrides: dict[str, str] = {}
            for p in pending:
                prompt_overrides[p["shape_name"]] = p["prompt"]
            edited = read_gpt_prompts_from_xlsx()
            for name, text in edited.items():
                if text.strip():
                    prompt_overrides[name] = text
            items, pt, gl, vl, _ = _build_all(
                mapping, prompts_spec, budgets, rows, metrics, shape_types,
                dry_run=False,
                prompt_overrides=prompt_overrides,
            )
        # Write final outputs
        write_json(OUT_CONTENT, {"generated_at": now_ts(), "sheet": sheet,
                                 "items": items, "metrics": metrics})
        write_json(OUT_PROMPT_TRACE, {"generated_at": now_ts(), "model": MODEL, "prompts": pt})
        write_md(OUT_VALID, vl)
        write_md(OUT_GAP, gl)
        safe_print(f"[OK] {len(items)} shapes built. Wrote {OUT_CONTENT.name}")
        return 0

    # ------------------------------------------------------------------
    # Default mode: single-pass (backward compatible, no pause)
    # ------------------------------------------------------------------
    items, pt, gl, vl, _ = _build_all(
        mapping, prompts_spec, budgets, rows, metrics, shape_types,
        dry_run=False,
    )
    write_json(OUT_CONTENT, {"generated_at": now_ts(), "sheet": sheet,
                             "items": items, "metrics": metrics})
    write_json(OUT_PROMPT_TRACE, {"generated_at": now_ts(), "model": MODEL, "prompts": pt})
    write_md(OUT_VALID, vl)
    write_md(OUT_GAP, gl)
    safe_print(f"[OK] {len(items)} shapes built. Wrote {OUT_CONTENT.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
