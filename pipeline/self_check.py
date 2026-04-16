# -*- coding: utf-8 -*-
"""Self-check functions for step1/step2 local iteration loops.

Step3 reuses 03b_build_ppt_com.py's built-in self-check report directly.

Usage (from agent via Bash):
    python -c "from pipeline.self_check import check_step1; import json; print(json.dumps(check_step1()))"
    python -c "from pipeline.self_check import check_step2; import json; print(json.dumps(check_step2()))"
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Dict, List

# ── paths ──────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parent.parent
PROGRESS_DIR = ROOT / "pipeline-progress"
SHAPE_JSON = PROGRESS_DIR / "01-shape_detail_com.json"
CONTENT_JSON = PROGRESS_DIR / "03a-build_shape_content.json"
BUDGET_JSON = PROGRESS_DIR / "02-readability_budget.json"
ANALYSIS_MAP = PROGRESS_DIR / "02-shape_analysis_map.json"

# Valid strategy codes (mirror of ppt_pipeline_common.STRATEGY_CODES)
_VALID_STRATEGIES = frozenset({
    "score_10pt", "grade_letter", "sample_aggregation", "extract_column",
    "extract_image", "gpt_prompted", "mean_extraction", "template_direct", "skip",
})


# ── Step 1 self-check ──────────────────────────────────────────────

def check_step1(progress_dir: Path = None) -> Dict[str, Any]:
    """Step 1 self-check: shape extraction completeness + annotation coverage.

    Checks:
      1. 01-shape_detail_com.json exists and new_shapes is non-empty
      2. Each shape has a valid strategy_exact (from xlsx annotation)
      3. gpt_prompted shapes have non-empty description

    Returns:
        {"passed": bool, "issues": [...], "summary": str}
    """
    pdir = Path(progress_dir) if progress_dir else PROGRESS_DIR
    shape_json = pdir / "01-shape_detail_com.json"
    analysis_map = pdir / "02-shape_analysis_map.json"

    issues: List[Dict[str, str]] = []

    # --- Check 1: shape JSON exists and has shapes ---
    if not shape_json.exists():
        return _fail("01-shape_detail_com.json not found", issues)

    data = json.loads(shape_json.read_text(encoding="utf-8"))
    shapes = data.get("new_shapes", []) if isinstance(data, dict) else []
    if not shapes:
        issues.append({"shape": "(all)", "problem": "new_shapes is empty"})
        return _result(False, issues)

    # --- Check 2 & 3: need analysis_map for strategy info ---
    if not analysis_map.exists():
        issues.append({"shape": "(all)", "problem": "02-shape_analysis_map.json not found — run 02_shape_analysis.py first"})
        return _result(False, issues)

    amap = json.loads(analysis_map.read_text(encoding="utf-8"))
    mapping = amap.get("mapping", [])

    # Build lookup: shape_name -> mapping entry
    by_name = {m["shape_name"]: m for m in mapping}

    for s in shapes:
        name = s.get("name", "?")
        m = by_name.get(name)
        if not m:
            # Shape exists in JSON but not in analysis map — not yet analyzed
            issues.append({"shape": name, "problem": "not found in analysis_map (step 02 not run?)"})
            continue

        strategy = m.get("strategy_exact", "")
        if not strategy or strategy == "(必填)":
            issues.append({"shape": name, "problem": f"strategy is empty or placeholder: '{strategy}'"})
        elif strategy not in _VALID_STRATEGIES:
            issues.append({"shape": name, "problem": f"unknown strategy: '{strategy}'"})

        # gpt_prompted shapes need a description
        if strategy == "gpt_prompted":
            desc = m.get("user_instruction", "") or m.get("user_strategy_hint", "")
            if not desc:
                issues.append({"shape": name, "problem": "gpt_prompted but no user description/instruction"})

    return _result(len(issues) == 0, issues)


# ── Step 2 self-check ──────────────────────────────────────────────

def check_step2(progress_dir: Path = None) -> Dict[str, Any]:
    """Step 2 self-check: content generation quality + golden reference comparison.

    Checks:
      1. 03a-build_shape_content.json exists
      2. Each non-skip shape has non-empty content
      3. Content length within budget * [0.5, 1.2]
      4. gpt_prompted shapes: structural similarity vs golden reference (paragraph/bullet count)

    Returns:
        {"passed": bool, "issues": [...], "summary": str}
    """
    pdir = Path(progress_dir) if progress_dir else PROGRESS_DIR
    content_json = pdir / "03a-build_shape_content.json"
    budget_json = pdir / "02-readability_budget.json"
    shape_json = pdir / "01-shape_detail_com.json"

    issues: List[Dict[str, str]] = []

    # --- Check 1: content JSON exists ---
    if not content_json.exists():
        return _fail("03a-build_shape_content.json not found", issues)

    cdata = json.loads(content_json.read_text(encoding="utf-8"))
    items = cdata.get("items", [])
    if not items:
        issues.append({"shape": "(all)", "problem": "items is empty"})
        return _result(False, issues)

    # Load budgets for length check
    budgets = {}
    if budget_json.exists():
        bdata = json.loads(budget_json.read_text(encoding="utf-8"))
        for b in bdata.get("budgets", []):
            budgets[b["shape_name"]] = b

    # Load golden reference (template text from shape JSON)
    golden = load_golden_reference(pdir)

    for item in items:
        name = item.get("shape_name", "?")
        strategy = item.get("strategy", "")
        content = item.get("content", "")

        if strategy == "skip":
            continue

        # --- Check 2: non-empty content ---
        if not content.strip():
            issues.append({"shape": name, "problem": "content is empty",
                           "fix_hint": "re-run GPT or check prompt"})
            continue

        # --- Check 3: length within budget ---
        budget = budgets.get(name) or item.get("budget", {})
        max_chars = budget.get("max_chars", 0)
        if max_chars > 0:
            clen = len(content)
            lo, hi = int(max_chars * 0.5), int(max_chars * 1.2)
            if clen < lo:
                issues.append({"shape": name,
                               "problem": f"content too short: {clen} chars (min {lo})",
                               "fix_hint": "add detail or loosen constraints in prompt"})
            elif clen > hi:
                issues.append({"shape": name,
                               "problem": f"content too long: {clen} chars (max {hi})",
                               "fix_hint": "add char limit to prompt"})

        # --- Check 4: structural similarity for gpt_prompted ---
        if strategy == "gpt_prompted" and name in golden:
            ref_text = golden[name]
            _check_structure(name, content, ref_text, issues)

    return _result(len(issues) == 0, issues)


# ── Golden reference loader ───────────────────────────────────────

def load_golden_reference(progress_dir: Path = None) -> Dict[str, str]:
    """Load each shape's original template text as golden reference.

    Source: 01-shape_detail_com.json[*].text
    Returns: {shape_name: template_text}
    """
    pdir = Path(progress_dir) if progress_dir else PROGRESS_DIR
    shape_json = pdir / "01-shape_detail_com.json"

    if not shape_json.exists():
        return {}

    data = json.loads(shape_json.read_text(encoding="utf-8"))
    shapes = data.get("new_shapes", []) if isinstance(data, dict) else []

    return {s["name"]: s.get("text", "") for s in shapes if s.get("name")}


# ── Helpers ────────────────────────────────────────────────────────

def _check_structure(name: str, content: str, ref_text: str,
                     issues: List[Dict[str, str]]) -> None:
    """Compare paragraph / bullet counts between content and golden reference."""
    ref_paras = _count_paragraphs(ref_text)
    gen_paras = _count_paragraphs(content)

    if ref_paras > 0:
        diff_ratio = abs(gen_paras - ref_paras) / ref_paras
        if diff_ratio > 0.3:
            issues.append({
                "shape": name,
                "problem": f"paragraph count mismatch: generated {gen_paras} vs template {ref_paras} (>{30}% diff)",
                "fix_hint": "adjust prompt to match template structure",
            })

    ref_bullets = _count_bullets(ref_text)
    gen_bullets = _count_bullets(content)
    if ref_bullets > 0:
        diff_ratio = abs(gen_bullets - ref_bullets) / ref_bullets
        if diff_ratio > 0.3:
            issues.append({
                "shape": name,
                "problem": f"bullet count mismatch: generated {gen_bullets} vs template {ref_bullets} (>{30}% diff)",
                "fix_hint": "adjust prompt to produce similar number of bullet points",
            })


def _count_paragraphs(text: str) -> int:
    """Count non-empty paragraphs (split by \\n\\n or \\r\\r)."""
    blocks = re.split(r'[\r\n]{2,}', text.strip())
    return len([b for b in blocks if b.strip()])


def _count_bullets(text: str) -> int:
    """Count lines starting with bullet-like patterns."""
    lines = re.split(r'[\r\n]+', text)
    bullet_re = re.compile(r'^\s*[-•·▪■□※★☆►▶P\d]')
    return sum(1 for l in lines if bullet_re.match(l))


def _fail(msg: str, issues: List) -> Dict[str, Any]:
    issues.insert(0, {"shape": "(all)", "problem": msg})
    return _result(False, issues)


def _result(passed: bool, issues: List) -> Dict[str, Any]:
    total = len(issues)
    if passed:
        summary = "All checks passed"
    else:
        summary = f"{total} issue(s) found"
    return {"passed": passed, "issues": issues, "summary": summary}


# ── CLI entry ──────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    fn = sys.argv[1] if len(sys.argv) > 1 else "step1"
    if fn == "step1":
        result = check_step1()
    elif fn == "step2":
        result = check_step2()
    else:
        print(f"Usage: python self_check.py [step1|step2]")
        sys.exit(1)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    sys.exit(0 if result["passed"] else 1)
