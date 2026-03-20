# -*- coding: utf-8 -*-
"""Step 1B: Auto-annotate shape_detail.xlsx based on rule table.

Reads 01-shape_detail_com.json, applies deterministic rules to infer
annotation fields (description / strategy / params), then writes them
directly into the xlsx via COM.

This replaces the slow LLM-based analysis that the Analyst agent
previously did shape-by-shape.

Rules (from analyst spec):
  text "X.XX/10"           -> score_10pt
  single uppercase letter  -> grade_letter
  "试穿人数"/"体重"/"球场"  -> sample_aggregation
  has_chart = true         -> mean_extraction
  shape_type = 13 (picture)-> extract_image
  product/shoe name        -> extract_column
  long text with 【】+ pos  -> gpt_prompted (filter=优点)
  long text with 【】+ neg  -> gpt_prompted (filter=缺点)
  title-like (short, big)  -> template_direct
  empty + no feature       -> skip
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    PROGRESS_DIR,
    SHAPE_DETAIL_XLSX,
    safe_print,
    setup_console_encoding,
)

DETAIL_JSON = PROGRESS_DIR / "01-shape_detail_com.json"

# Negative-sentiment keywords (complaints / issues)
_NEG_KEYWORDS = re.compile(
    r"不够|不足|偏[硬软重]|缺[少乏点]|差|问题|不[好佳舒适]|挤压|打滑|卡脚|异物感|磨脚|闷热|缺点"
)
# Positive-sentiment keywords
_POS_KEYWORDS = re.compile(
    r"稳定|舒适|澎湃|回弹|支撑|包裹|锁定|反馈|优[点势秀良]|出色|优秀|充足"
)
# Negation patterns after keywords: "支撑略低" → positive negated; "异物感不明显" → negative negated
_POS_NEGATION = re.compile(r"[略不]|低|差|不[够足]")
_NEG_NEGATION = re.compile(r"不[明严]|尚[能可]")


def _has_positive(text: str) -> bool:
    """Check for positive keywords NOT in a negative context."""
    for m in _POS_KEYWORDS.finditer(text):
        after = text[m.end():m.end() + 3]
        if _POS_NEGATION.match(after):
            continue
        return True
    return False


def _has_negative(text: str) -> bool:
    """Check for negative keywords NOT in a negation context."""
    for m in _NEG_KEYWORDS.finditer(text):
        after = text[m.end():m.end() + 4]
        if _NEG_NEGATION.search(after):
            continue
        return True
    return False


def infer_annotation(shape: dict) -> dict:
    """Apply rule table to a single shape, return annotation dict.

    Returns: {description, strategy, params} — empty strings if not applicable.
    """
    name = shape.get("name", "")
    text = shape.get("text", "").strip()
    shape_type = shape.get("shape_type", 0)
    has_chart = shape.get("has_chart", False)
    font_size = shape.get("font_size", 0)

    desc = ""
    strategy = ""
    params = ""

    # Rule 1: Chart shape
    if has_chart:
        strategy = "mean_extraction"
        desc = "（自动计算均值）"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 2: Picture shape (type 13)
    if shape_type == 13:
        strategy = "extract_image"
        desc = "（自动提取图片）"
        params = "sheet=问卷"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 3: Score pattern "X.XX/10"
    if re.search(r"\d+\.\d+/10", text):
        desc = "评分均值10分制"
        strategy = "score_10pt"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 4: Single uppercase letter (grade), large font
    if re.match(r"^[A-S]$", text) and font_size >= 36:
        desc = "评分均值100分制档"
        strategy = "grade_letter"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 5: Sample aggregation (试穿人数 / 体重 / 球场定位)
    if any(kw in text for kw in ("试穿人数", "体重", "球场定位", "球场")):
        desc = "不走GPT统计人数体重"
        strategy = "sample_aggregation"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 6: Product name (short text, has quotes or product-like pattern)
    if text and len(text) < 30 and ("'" in text or "'" in text) and font_size >= 16:
        desc = "鞋款名称"
        strategy = "extract_column"
        params = "column=鞋款名称"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 7: Long text with 【】brackets -> GPT prompted (pos/neg)
    if len(text) > 60 and "【" in text and "】" in text:
        # Strip category headers like 【包裹性】【稳定性】 before sentiment matching
        text_no_headers = re.sub(r"【[^】]{1,10}】", "", text)
        has_neg = _has_negative(text_no_headers)
        has_pos = _has_positive(text_no_headers)

        if has_neg and not has_pos:
            desc = ("从补充说明总结缺点。"
                    "必须包含'建议'、'反馈'、'样本'关键词，"
                    "用【】括起关键性能词，每段结论后注明(X/N)比例")
            strategy = "gpt_prompted"
            params = "source=补充说明, filter=缺点"
        elif has_pos and not has_neg:
            desc = ("从补充说明总结优点。"
                    "必须包含'建议'、'反馈'、'样本'关键词，"
                    "用【】括起关键性能词，每段结论后注明(X/N)比例")
            strategy = "gpt_prompted"
            params = "source=补充说明, filter=优点"
        else:
            # Mixed or ambiguous — default to free-form GPT
            desc = ("从补充说明总结优缺点。"
                    "必须包含'建议'、'反馈'、'样本'关键词，"
                    "用【】括起关键性能词")
            strategy = "gpt_prompted"
            params = "source=补充说明"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 8: Title-like (short text, big font, not a score)
    if text and len(text) < 20 and font_size >= 14:
        strategy = "template_direct"
        desc = "（自动保留原文）"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 9: Empty text + no features -> skip
    if not text:
        strategy = "skip"
        desc = "（自动跳过）"
        return {"description": desc, "strategy": strategy, "params": params}

    # Rule 10: Short text that doesn't match above -> template_direct
    if len(text) < 20:
        strategy = "template_direct"
        desc = "（自动保留原文）"
        return {"description": desc, "strategy": strategy, "params": params}

    # Fallback: non-trivial text that doesn't match rules → auto-classify
    if len(text) > 60 and ("【" in text or "】" in text or "（" in text):
        # Long text with structural markers → GPT generation
        desc = ("从补充说明总结要点。"
                "必须包含'建议'、'反馈'、'样本'关键词，"
                "用【】括起关键性能词，每段结论后注明(X/N)比例")
        strategy = "gpt_prompted"
        params = "source=补充说明"
    elif len(text) > 30:
        # Medium-length text → GPT basic generation
        desc = "从补充说明总结要点"
        strategy = "gpt_prompted"
        params = "source=补充说明"
    else:
        # Short text that can't be classified → keep original
        desc = "（自动保留原文）"
        strategy = "template_direct"
    return {"description": desc, "strategy": strategy, "params": params}


def write_annotations_to_xlsx(annotations: dict[str, dict]) -> None:
    """Write inferred annotations into shape_detail.xlsx via COM.

    annotations: {shape_name: {description, strategy, params}}
    """
    if not SHAPE_DETAIL_XLSX.exists():
        safe_print("[ERROR] 01-shape_detail.xlsx not found, run Step 1 first")
        return

    import win32com.client

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(SHAPE_DETAIL_XLSX.resolve()))
        ws = wb.Sheets(1)
        max_row = ws.UsedRange.Rows.Count

        current_shape = None
        in_annotation = False
        written = 0

        for r in range(1, max_row + 1):
            a_val = str(ws.Cells(r, 1).Value or "").strip()
            b_val = str(ws.Cells(r, 2).Value or "").strip()

            if a_val.startswith("Shape #"):
                current_shape = b_val
                in_annotation = False
                continue

            if a_val == "用户批注":
                in_annotation = True
                continue

            if in_annotation and current_shape and current_shape in annotations:
                if a_val == "GPT-prompt Text":
                    continue  # never overwrite user-editable GPT prompt
                anno = annotations[current_shape]
                if a_val == "内容描述" and anno.get("description"):
                    # Only write if cell is currently empty
                    if not b_val:
                        cell = ws.Cells(r, 2)
                        desc_val = anno["description"]
                        cell.Value = desc_val
                        if desc_val == "（必填）":
                            cell.Font.Color = 0x0000FF  # Red (BGR)
                            cell.Font.Bold = True
                        elif desc_val.startswith("（自动"):
                            cell.Font.Color = 0x999999  # Gray (BGR)
                            cell.Font.Bold = False
                            cell.Interior.ColorIndex = 0  # Remove yellow fill
                        written += 1
                elif a_val == "strategy" and anno.get("strategy"):
                    if not b_val:
                        ws.Cells(r, 2).Value = anno["strategy"]
                elif a_val == "params" and anno.get("params"):
                    if not b_val:
                        ws.Cells(r, 2).Value = anno["params"]

        wb.Save()
        wb.Close(False)
        safe_print(f"[OK] 写入 {written} 个 shape 的批注到 xlsx")

    except Exception as e:
        safe_print(f"[ERROR] COM write failed: {e}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass


def main() -> int:
    setup_console_encoding()

    if not DETAIL_JSON.exists():
        safe_print("[ERROR] 01-shape_detail_com.json not found, run Step 1 first")
        return 1

    with open(DETAIL_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    shapes = data.get("new_shapes", [])
    if not shapes:
        safe_print("[WARN] No shapes found in JSON")
        return 0

    # Infer annotations
    annotations: dict[str, dict] = {}
    for s in shapes:
        name = s.get("name", "")
        if not name:
            continue
        anno = infer_annotation(s)
        annotations[name] = anno

    # Print summary
    ambiguous = [n for n, a in annotations.items()
                 if a.get("description") in ("（必填）", "从补充说明总结优缺点")
                 or not a.get("strategy")]
    safe_print(f"\n{'=' * 60}")
    safe_print(f"自动批注推断结果 ({len(annotations)} shapes, {len(ambiguous)} 项待LLM审核)")
    safe_print(f"{'=' * 60}")
    for i, (name, anno) in enumerate(annotations.items(), 1):
        strat = anno.get("strategy", "") or "—"
        desc = anno.get("description", "") or "—"
        par = anno.get("params", "")
        flag = " ⚠️" if name in ambiguous else ""
        line = f"  #{i:2d} [{name}]: {strat}"
        if desc and desc != "—":
            line += f" — {desc}"
        if par:
            line += f" ({par})"
        line += flag
        safe_print(line)
    if ambiguous:
        safe_print(f"\n⚠️  需要LLM审核: {', '.join(ambiguous)}")
    else:
        safe_print(f"\n✅ 全部推断明确，无需LLM审核")
    safe_print(f"{'=' * 60}\n")

    # Write to xlsx
    write_annotations_to_xlsx(annotations)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
