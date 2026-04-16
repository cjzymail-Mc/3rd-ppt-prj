# -*- coding: utf-8 -*-
"""Step 2B: Prepare xlsx for iteration round.

Reads 04-fix_ppt.md + 04-diff_result.json, creates a new sheet in
01-shape_detail.xlsx (copied from the last sheet), and applies basic
annotation fixes for fix_type=annotation items.

Usage:
    python pipeline/02b_iteration_setup.py --version 1.1

This replaces the LLM-based "parse fix report → create sheet → update
annotations" steps that the Builder agent previously did manually.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    PROGRESS_DIR,
    SHAPE_DETAIL_XLSX,
    create_iteration_sheet,
    safe_print,
    setup_console_encoding,
)

DIFF_JSON = PROGRESS_DIR / "04-diff_result.json"

# Constraint patterns — same-type constraints only keep the latest version
_CONSTRAINT_PATTERNS = {
    "budget_overflow":  re.compile(r"[。；;]\s*(?:字数严格控制|严格控制总字数|宁短勿长)[^。；;]*"),
    "budget_underflow": re.compile(r"[。；;]\s*(?:内容需[要更]充实|字数控制在)[^。；;]*"),
    "keyword_missing":  re.compile(r"[。；;]\s*必须包含[^。；;]*关键词[^。；;]*"),
    "style_mismatch":   re.compile(r"[。；;]\s*严格按照参考文本[^。；;]*"),
}


def _extract_char_counts(issue: str) -> tuple[int, int]:
    """Extract (actual_chars, target_chars) from issue description."""
    m = re.search(r"(\d+)/(\d+)字", issue)
    if m:
        return int(m.group(1)), int(m.group(2))
    return 0, 0


def apply_annotation_fixes(sheet_name: str, fixes: list[dict]) -> int:
    """Apply basic annotation fixes to the new sheet via COM.

    For readability issues: appends suggestion to 内容描述.
    For semantic issues: appends keyword reminder to 内容描述.

    Returns number of shapes updated.
    """
    if not fixes:
        return 0

    # Accept all non-code fix types (keyword_missing, budget_overflow, budget_underflow, style_mismatch, annotation)
    annotation_fixes = [f for f in fixes if f.get("fix_type", "") != "code"]
    if not annotation_fixes:
        return 0

    import win32com.client
    import time as _time

    # Use DispatchEx to force a NEW Excel instance, avoiding conflict with
    # any lingering COM process from previous pipeline steps.
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    _time.sleep(0.5)  # let COM server stabilize

    updated = 0
    try:
        wb = excel.Workbooks.Open(str(SHAPE_DETAIL_XLSX.resolve()))

        # Find the target sheet
        ws = None
        for i in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(i).Name == sheet_name:
                ws = wb.Sheets(i)
                break
        if ws is None:
            safe_print(f"[WARN] Sheet '{sheet_name}' not found")
            wb.Close(False)
            return 0

        max_row = ws.UsedRange.Rows.Count

        # Build fix lookup: shape_name -> fix info
        # For semantic (global) fixes, apply to all gpt_prompted shapes
        shape_fixes = {}
        global_fixes = []
        for fx in annotation_fixes:
            if fx["shape"] == "(全局)":
                global_fixes.append(fx)
            else:
                shape_fixes[fx["shape"]] = fx

        # Walk through xlsx rows
        current_shape = None
        in_annotation = False

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

            if not in_annotation or not current_shape:
                continue

            # Apply shape-specific fixes to 内容描述 field (targeted by fix_type)
            if a_val == "内容描述" and current_shape in shape_fixes:
                fx = shape_fixes[current_shape]
                ft = fx.get("fix_type", "")
                existing = b_val

                # Quantitative suffix based on fix_type
                if ft == "keyword_missing":
                    suffix = "。必须包含'样本'、'建议'、'反馈'关键词"
                elif ft == "budget_overflow":
                    actual, target = _extract_char_counts(fx.get("issue", ""))
                    if actual and target:
                        lo = int(target * 0.9)
                        hi = int(target * 1.1)
                        suffix = f"。字数严格控制在{lo}-{hi}字（当前{actual}字，目标{target}字）"
                    else:
                        suffix = "。严格控制总字数，宁短勿长"
                elif ft == "budget_underflow":
                    actual, target = _extract_char_counts(fx.get("issue", ""))
                    if actual and target:
                        lo = int(target * 0.9)
                        hi = int(target * 1.1)
                        suffix = f"。内容需更充实，字数控制在{lo}-{hi}字（当前{actual}字，目标{target}字）"
                    else:
                        suffix = "。内容需要更充实，涵盖更多测试者反馈"
                elif ft == "style_mismatch":
                    suffix = "。严格按照参考文本的格式和语调"
                else:
                    suffix = f"; {fx.get('suggestion', '')}" if fx.get("suggestion") else ""

                if suffix:
                    # Remove old same-type constraint before appending new one
                    pattern = _CONSTRAINT_PATTERNS.get(ft)
                    if pattern:
                        existing = pattern.sub("", existing)
                    if suffix not in existing:
                        new_val = f"{existing}{suffix}" if existing else suffix.lstrip("。; ")
                        ws.Cells(r, 2).Value = new_val
                        updated += 1

            # Apply global keyword fixes to 内容描述 of gpt_prompted shapes
            if a_val == "内容描述" and global_fixes:
                # Check if this shape uses gpt_prompted strategy
                # Look back a few rows for strategy field
                for check_r in range(max(1, r - 3), r):
                    check_a = str(ws.Cells(check_r, 1).Value or "").strip()
                    check_b = str(ws.Cells(check_r, 2).Value or "").strip()
                    if check_a == "strategy" and "gpt" in check_b.lower():
                        existing = b_val
                        for gfx in global_fixes:
                            ft = gfx.get("fix_type", "")
                            if ft == "keyword_missing":
                                suffix = "。必须包含'样本'、'建议'、'反馈'关键词"
                            else:
                                suffix = f"; {gfx.get('suggestion', '')}" if gfx.get("suggestion") else ""
                            if suffix:
                                pattern = _CONSTRAINT_PATTERNS.get(ft)
                                if pattern:
                                    existing = pattern.sub("", existing)
                                if suffix not in existing:
                                    existing = f"{existing}{suffix}" if existing else suffix.lstrip("。; ")
                        if existing != b_val:
                            ws.Cells(r, 2).Value = existing
                            updated += 1
                        break

        _time.sleep(0.5)  # stabilize before save
        wb.Save()
        wb.Close(False)

    except Exception as e:
        safe_print(f"[ERROR] COM write failed: {e}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass
        _time.sleep(0.3)  # let Excel process fully exit

    return updated


def main() -> int:
    setup_console_encoding()

    ap = argparse.ArgumentParser()
    ap.add_argument("--version", required=True, help="Version string, e.g. '1.1'")
    ap.add_argument("--sheet-only", action="store_true",
                    help="Only create new sheet (inherit prompts), skip annotation fixes")
    args = ap.parse_args()

    sheet_name = f"claude-ppt {args.version}"

    # 1. Create new sheet FIRST (always needed — inherits annotations from prev version)
    result = create_iteration_sheet(sheet_name)
    if result:
        safe_print(f"[OK] 创建新 sheet: {sheet_name}")
    else:
        safe_print(f"[ERROR] 创建 sheet 失败: {sheet_name}")
        return 1

    # Sheet-only mode: just create the sheet, skip annotation fixes
    if args.sheet_only:
        safe_print(f"[OK] Sheet-only 模式：{sheet_name} 已创建，跳过注释修正")
        return 0

    # 2. Read diff result and apply fixes (optional — sheet is already usable without fixes)
    if not DIFF_JSON.exists():
        safe_print("[WARN] 04-diff_result.json not found, skipping fix application")
        return 0

    with open(DIFF_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    fixes = data.get("fixes", [])
    non_code_fixes = [f for f in fixes if f.get("fix_type", "") != "code"]
    code_fixes = [f for f in fixes if f.get("fix_type", "") == "code"]

    if not fixes:
        safe_print("[WARN] No fixes found in diff result")
        return 0

    safe_print(f"修正项: {len(fixes)} 个 (non-code={len(non_code_fixes)}, code={len(code_fixes)})")

    # 3. Apply annotation fixes
    updated = apply_annotation_fixes(sheet_name, fixes)
    safe_print(f"[OK] 更新 {updated} 个 shape 的批注")

    safe_print(f"\n下一步: python pipeline/02_shape_analysis.py --sheet \"{sheet_name}\"")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
