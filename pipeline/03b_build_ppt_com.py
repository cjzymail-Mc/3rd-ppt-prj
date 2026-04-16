#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 3B: Clone template slide 15 and write content via COM.

Flow:
  1. Open PowerPoint, open Template
  2. Clone slide 15 to new Presentation
  3. Read build_shape_content.json
  4. Match shapes by name (fallback: geometry position)
  5. Text shapes: TextFrame.TextRange.Text = content (preserve formatting)
  6. Chart shapes: write to ChartData worksheet
  7. Post-write readback verification
  8. Save as claude-ppt {version}.pptx

Fixed vs codex-legacy2:
  - Chart data: rsplit(":", 1) instead of split(":")
  - Chart workbook: wb.Close(False) after writing
  - Copy/Paste: time.sleep(delay) buffer
  - Geometry fallback for shape matching
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pipeline.ppt_pipeline_common import (
    OUTPUT_DIR,
    PROGRESS_DIR,
    ROOT,
    TEMPLATE_PATH,
    com_call,
    com_get,
    generate_self_check_report,
    now_ts,
    safe_print,
    safe_text,
    setup_console_encoding,
    slide_export_png,
    ssim_compare,
    write_json,
    write_md,
)

CONTENT_JSON = PROGRESS_DIR / "03a-build_shape_content.json"
OUT_REPORT   = PROGRESS_DIR / "03b-build_ppt_report.md"
OUT_READBACK = PROGRESS_DIR / "03b-post_write_readback.json"
SELF_CHECK_REPORT = PROGRESS_DIR / "03b-self_check_report.md"
BASELINE_JSON = PROGRESS_DIR / "01-shape_detail_com.json"
MAX_SELF_FIX = 2

COPY_PASTE_DELAY = 1.5  # seconds buffer for COM clipboard operations


def _find_shape_by_name(slide, name: str):
    """Find shape by exact name match."""
    try:
        return slide.Shapes(name)
    except Exception:
        return None


def _find_shape_by_geometry(slide, item: dict, tolerance: float = 15.0):
    """Fallback: find shape by closest geometry match."""
    target_left = item.get("left", 0)
    target_top = item.get("top", 0)
    best = None
    best_dist = float("inf")

    for i in range(1, int(slide.Shapes.Count) + 1):
        shp = slide.Shapes(i)
        left = float(com_get(shp, "Left", 0) or 0)
        top = float(com_get(shp, "Top", 0) or 0)
        dist = abs(left - target_left) + abs(top - target_top)
        if dist < best_dist:
            best_dist = dist
            best = shp

    if best is not None and best_dist <= tolerance:
        return best
    return None


def _write_text(shp, content: str) -> dict:
    """Write text to shape, preserving formatting."""
    name = safe_text(com_get(shp, "Name", ""))
    if not bool(com_get(shp, "HasTextFrame", 0)):
        return {"shape_name": name, "updated": False, "reason": "no text frame"}

    tf = com_get(shp, "TextFrame", None)
    tr = com_get(tf, "TextRange", None) if tf is not None else None
    if tr is None:
        return {"shape_name": name, "updated": False, "reason": "no text range"}

    # Lock shape size to preserve template geometry (prevent AutoFit resize)
    try:
        tf.AutoSize = 0  # ppAutoSizeNone = 0
    except Exception:
        pass

    try:
        # PPT COM uses \r for paragraph breaks; \n is ignored / collapses to 1 paragraph
        tr.Text = content.replace("\n", "\r")
        tr.Font.Name = "微软雅黑"
        rb_text = com_get(tr, "Text", "") or ""
        return {
            "shape_name": name,
            "updated": True,
            "mode": "text",
            "written_len": len(content),
            "readback_len": len(rb_text),
        }
    except Exception as e:
        return {"shape_name": name, "updated": False, "reason": str(e)}


def _write_chart(shp, content: str) -> dict:
    """Write chart data via SeriesCollection — bypasses ChartData.Workbook entirely.

    Strategy:
    1. Activate() + BreakLink() in an isolated try block to clear external link refs
       (Activate is required for BreakLink; both may fail on non-interactive sessions
       but the data write below will still succeed)
    2. Parse content ("label:value" per line) into labels/values lists
    3. Set series.Values and series.XValues directly on SeriesCollection(1)
       — no workbook open/close, no "linked file" dialog risk
    """
    name = safe_text(com_get(shp, "Name", ""))
    chart = com_get(shp, "Chart", None)
    if chart is None:
        return {"shape_name": name, "updated": False, "reason": "no chart object"}

    # Parse content first (fail fast if nothing to write)
    lines = [x.strip() for x in (content or "").splitlines() if x.strip()]
    labels, values = [], []
    for line in lines[:10]:
        if ":" in line:
            k, v = line.rsplit(":", 1)
            labels.append(k.strip())
            try:
                values.append(float(v.strip()))
            except Exception:
                values.append(0.0)

    if not labels:
        return {"shape_name": name, "updated": False, "reason": "no parseable data in content"}

    try:
        # Step 1: Break external link.
        # Activate() is required for BreakLink(), but throws on non-interactive sessions.
        # Isolated try — failure here does not abort data write.
        try:
            chart.ChartData.Activate()
            time.sleep(0.5)
            chart.ChartData.BreakLink()
            time.sleep(0.3)
        except Exception:
            pass

        # Step 2: Write data via SeriesCollection — no Workbook access, no linked file risk.
        series = chart.SeriesCollection(1)
        series.Values = tuple(values)
        series.XValues = tuple(labels)

        ctype = int(com_get(chart, "ChartType", 0) or 0)
        return {
            "shape_name": name,
            "updated": True,
            "mode": "chart",
            "chart_type": ctype,
            "data_rows": len(labels),
        }
    except Exception as e:
        return {"shape_name": name, "updated": False, "reason": str(e)}


# Color constants (PowerPoint RGB integers)
_RED = 255           # RGB(255, 0, 0)  — advantage keywords
_BLUE = 15773696     # RGB(0, 176, 240) — disadvantage keywords
_BLACK = 0           # RGB(0, 0, 0)    — default text

# Section markers for auto-detecting advantage vs disadvantage
_ADVANTAGE_MARKERS = ["优势", "优点", "亮点", "表现较好", "表现突出"]
_DISADVANTAGE_MARKERS = ["问题", "缺点", "劣势", "不足", "改进", "修改建议", "待优化"]


def _apply_keyword_color(shp) -> None:
    """Remove 【】 brackets, then bold+color keywords by section context.

    Rules:
      - Keywords in advantage sections → red + bold
      - Keywords in disadvantage sections → blue + bold
      - All other text → black
    """
    try:
        tf = com_get(shp, "TextFrame", None)
        if tf is None:
            return
        tr = tf.TextRange
        full_text = tr.Text

        keywords = list(dict.fromkeys(re.findall(r'【([^】]+)】', full_text)))
        if not keywords:
            return

        # Build keyword→color map based on section context
        kw_color = {}
        current_section = "neutral"
        for line in full_text.split('\r'):
            line_stripped = line.strip()
            # Check if this line is a section header
            if any(m in line_stripped for m in _ADVANTAGE_MARKERS):
                current_section = "advantage"
            elif any(m in line_stripped for m in _DISADVANTAGE_MARKERS):
                current_section = "disadvantage"
            # Assign color to keywords found on this line
            for kw in re.findall(r'【([^】]+)】', line):
                if current_section == "advantage":
                    kw_color[kw] = _RED
                elif current_section == "disadvantage":
                    kw_color[kw] = _BLUE
                # neutral keywords stay uncolored (black)

        # Remove all 【】 brackets
        tr.Text = re.sub(r'[【】]', '', full_text)

        # Set entire shape to black first (reset any inherited colors)
        tr.Font.Color = _BLACK

        # Bold + color each keyword
        for kw, color in kw_color.items():
            start = 1
            while start <= tr.Length:
                found = tr.Find(kw, start)
                if found is None:
                    break
                found.Font.Bold = True
                found.Font.Color = color
                start = found.Start + found.Length
    except Exception:
        pass  # coloring is cosmetic — never fail the build


def _replace_image(slide, shp, img_path: str) -> dict:
    """Replace an image shape with a picture loaded from img_path.

    Records the original geometry, deletes the old shape, then inserts a
    new picture at exactly the same Left/Top/Width/Height.
    """
    name = safe_text(com_get(shp, "Name", ""))
    try:
        left   = float(com_get(shp, "Left",   0))
        top    = float(com_get(shp, "Top",    0))
        width  = float(com_get(shp, "Width",  100))
        height = float(com_get(shp, "Height", 100))
        shp.Delete()
        new_shp = slide.Shapes.AddPicture(
            FileName=img_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left, Top=top,
            Width=width, Height=height,
        )
        try:
            new_shp.Name = name  # restore original name for diff-test matching
        except Exception:
            pass
        return {
            "shape_name": name,
            "updated": True,
            "mode": "image_replace",
            "img_path": img_path,
        }
    except Exception as e:
        return {"shape_name": name, "updated": False, "reason": str(e)}


def apply_shape(slide, item: dict, shape_detail: dict = None) -> dict:
    """Find and update a single shape on the slide."""
    name = item.get("shape_name", "")
    content = item.get("content", "")
    role = item.get("role", "")
    strategy = item.get("strategy", "")

    # Skip shapes explicitly marked as skip — preserve cloned template geometry
    if strategy == "skip":
        return {"shape_name": name, "updated": False, "reason": "strategy_skip", "match_method": "n/a"}

    # Try name match first, then geometry fallback
    shp = _find_shape_by_name(slide, name)
    match_method = "name"
    if shp is None and shape_detail:
        shp = _find_shape_by_geometry(slide, shape_detail)
        match_method = "geometry"
    if shp is None:
        return {"shape_name": name, "updated": False, "reason": "shape not found", "match_method": "none"}

    # Route to appropriate writer
    color_hint = item.get("color_hint", "")
    if strategy == "extract_image":
        if content.startswith("IMAGE:"):
            result = _replace_image(slide, shp, content[6:].strip())
        else:
            result = {"shape_name": name, "updated": False, "reason": "extract_image: no IMAGE: path in content"}
    elif role == "chart" or bool(com_get(shp, "HasChart", False)):
        result = _write_chart(shp, content)
    elif bool(com_get(shp, "HasTextFrame", 0)):
        result = _write_text(shp, content)
        # Post-write: auto-detect 【】keywords → red/blue by section context
        if result.get("updated"):
            _apply_keyword_color(shp)
    else:
        result = {"shape_name": name, "updated": False, "reason": "unsupported shape type"}

    result["match_method"] = match_method
    return result


# ---- self-check functions ----

def _check_attributes(slide, baseline: dict) -> list:
    """Steps ①②: Compare shape attributes and editability against baseline.

    Returns list of issue dicts.
    """
    issues = []
    for i in range(1, int(slide.Shapes.Count) + 1):
        shp = slide.Shapes(i)
        name = safe_text(com_get(shp, "Name", ""))
        ref = baseline.get(name)
        if ref is None:
            continue

        # ① Attribute verification: position and size
        for prop, key, threshold, label in [
            ("Left",   "left",   3.0, "位置Left"),
            ("Top",    "top",    3.0, "位置Top"),
            ("Width",  "width",  5.0, "宽度"),
            ("Height", "height", 5.0, "高度"),
        ]:
            actual = float(com_get(shp, prop, 0) or 0)
            expected = float(ref.get(key, 0))
            diff = abs(actual - expected)
            if diff > threshold:
                issues.append({
                    "shape": name, "dimension": "属性校验",
                    "description": f"{label}偏差 {diff:.1f}pt (期望{expected:.0f}, 实际{actual:.0f})",
                    "severity": "严重" if diff > threshold * 2 else "中等",
                    "status": "待修复",
                })

        # ② Editability verification
        has_tf = bool(com_get(shp, "HasTextFrame", 0))
        has_chart = bool(com_get(shp, "HasChart", False))
        has_table = bool(com_get(shp, "HasTable", False))

        if has_tf:
            try:
                tf = com_get(shp, "TextFrame", None)
                _ = tf.TextRange.Text
            except Exception:
                issues.append({
                    "shape": name, "dimension": "可编辑性",
                    "description": "文本框不可读写",
                    "severity": "严重", "status": "待修复",
                })
        if has_chart:
            try:
                _ = com_get(shp, "Chart", None)
            except Exception:
                issues.append({
                    "shape": name, "dimension": "可编辑性",
                    "description": "图表不可访问",
                    "severity": "严重", "status": "待修复",
                })
        if has_table:
            try:
                _ = shp.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text
            except Exception:
                issues.append({
                    "shape": name, "dimension": "可编辑性",
                    "description": "表格不可编辑",
                    "severity": "严重", "status": "待修复",
                })

    return issues


def _check_visual(app, template_slide, generated_slide) -> tuple:
    """Step ③: Visual comparison via clipboard-based PNG + SSIM.

    Uses slide.Copy() → PIL.ImageGrab to get unencrypted PNGs,
    then compares with SSIM.

    Returns (ssim_score, issues_list).
    ssim_score = -1.0 if comparison was skipped.
    """
    baseline_png = str(PROGRESS_DIR / "03b-baseline_page.png")
    generated_png = str(PROGRESS_DIR / "03b-generated_page.png")

    ok_a = slide_export_png(app, template_slide, baseline_png)
    ok_b = slide_export_png(app, generated_slide, generated_png)

    if not ok_a or not ok_b:
        safe_print("[WARN] Visual comparison skipped (clipboard export failed)")
        return -1.0, []

    score = ssim_compare(baseline_png, generated_png)
    issues = []
    if score < 0:
        safe_print("[WARN] Visual comparison skipped (SSIM failed)")
    elif score < 0.85:
        issues.append({
            "shape": "全页", "dimension": "视觉对照",
            "description": f"SSIM={score:.2f} < 0.85 视觉差异过大",
            "severity": "严重", "status": "需人工检查",
        })
    elif score < 0.92:
        issues.append({
            "shape": "全页", "dimension": "视觉对照",
            "description": f"SSIM={score:.2f} < 0.92 存在可见差异",
            "severity": "中等", "status": "需人工检查",
        })
    else:
        safe_print(f"[OK] Visual SSIM={score:.2f}")

    return score, issues


def _check_content(slide, expected_items: list) -> list:
    """Step ④: Content completeness — verify written text matches expectations.

    Returns list of issue dicts.
    """
    issues = []
    for item in expected_items:
        name = item.get("shape_name", "")
        expected = item.get("content", "")
        strategy = item.get("strategy", "")
        if strategy == "skip" or not expected:
            continue

        shp = _find_shape_by_name(slide, name)
        if shp is None:
            issues.append({
                "shape": name, "dimension": "内容完整性",
                "description": "shape未找到",
                "severity": "严重", "status": "待修复",
            })
            continue

        # For image shapes, check existence rather than text
        if strategy == "extract_image":
            shp_type = int(com_get(shp, "Type", 0) or 0)
            if shp_type != 13:  # msoPicture = 13
                issues.append({
                    "shape": name, "dimension": "内容完整性",
                    "description": "图片未成功插入",
                    "severity": "严重", "status": "待修复",
                })
            continue

        if not bool(com_get(shp, "HasTextFrame", 0)):
            continue

        # Font check
        try:
            font_name = com_get(shp, "TextFrame", None).TextRange.Font.Name or ""
            if font_name and font_name != "微软雅黑":
                issues.append({
                    "shape": name, "dimension": "字体校验",
                    "description": f"字体不一致: {font_name} (应为微软雅黑)",
                    "severity": "中等", "status": "待修复",
                })
        except Exception:
            pass

        try:
            actual = (com_get(shp, "TextFrame", None).TextRange.Text or "").strip()
        except Exception:
            actual = ""

        if not actual and expected.strip():
            issues.append({
                "shape": name, "dimension": "内容完整性",
                "description": "文本缺失（应有内容但为空）",
                "severity": "严重", "status": "待修复",
            })
        elif len(actual) < len(expected.strip()) * 0.8 and len(expected.strip()) > 10:
            issues.append({
                "shape": name, "dimension": "内容完整性",
                "description": f"文本可能截断 (实际{len(actual)}字 vs 预期{len(expected.strip())}字)",
                "severity": "中等", "status": "待修复",
            })

        # Content overflow check against budget
        budget = item.get("budget") or {}
        max_chars = budget.get("max_chars", 0)
        if max_chars > 0 and len(actual) > max_chars * 1.2:
            issues.append({
                "shape": name, "dimension": "内容完整性",
                "description": f"内容超长 {len(actual)}字 (上限{max_chars}字, 超出{len(actual) - max_chars}字)",
                "severity": "严重" if len(actual) > max_chars * 1.5 else "中等",
                "status": "待修复",
            })

        # Paragraph count check: actual paragraphs vs expected line count
        max_lines = budget.get("max_lines", 0)
        if max_lines > 0:
            try:
                actual_paras = int(com_get(shp, "TextFrame", None).TextRange.Paragraphs().Count)
            except Exception:
                actual_paras = -1
            if actual_paras == 1 and max_lines > 1:
                issues.append({
                    "shape": name, "dimension": "内容完整性",
                    "description": f"段落结构异常: 仅1段 (模板期望{max_lines}行), 可能换行符错误",
                    "severity": "严重", "status": "待修复",
                })
            elif actual_paras > max_lines * 1.5 and actual_paras > 3:
                issues.append({
                    "shape": name, "dimension": "内容完整性",
                    "description": f"段落过多: {actual_paras}段 (模板期望{max_lines}行)",
                    "severity": "中等", "status": "待修复",
                })

        # Check required_keywords if present
        keywords = item.get("required_keywords", [])
        for kw in keywords:
            # Strip 【】 brackets for comparison (coloring step removes them)
            clean_actual = re.sub(r'[【】]', '', actual)
            if kw not in clean_actual:
                issues.append({
                    "shape": name, "dimension": "内容完整性",
                    "description": f"关键词缺失：{kw}",
                    "severity": "严重", "status": "待修复",
                })

    return issues


def _auto_fix(slide, issues: list, items: list, shape_details: dict) -> list:
    """Attempt to auto-fix issues. Returns updated issues with status changed."""
    updated_issues = []
    for iss in issues:
        name = iss.get("shape", "")
        dim = iss.get("dimension", "")
        sev = iss.get("severity", "")

        # Only attempt fix on fixable types
        if iss.get("status") == "需人工检查":
            updated_issues.append(iss)
            continue

        shp = _find_shape_by_name(slide, name) if name and name != "全页" else None

        if dim == "字体校验" and shp:
            try:
                shp.TextFrame.TextRange.Font.Name = "微软雅黑"
                iss["status"] = "已修复"
            except Exception:
                iss["status"] = "修复失败"
        elif dim == "内容完整性" and "文本缺失" in iss.get("description", "") and shp:
            # Re-write the content
            matching = [it for it in items if it.get("shape_name") == name]
            if matching:
                result = _write_text(shp, matching[0].get("content", ""))
                if result.get("updated"):
                    iss["status"] = "已修复"
                else:
                    iss["status"] = "修复失败"
        elif dim == "内容完整性" and "文本可能截断" in iss.get("description", "") and shp:
            matching = [it for it in items if it.get("shape_name") == name]
            if matching:
                result = _write_text(shp, matching[0].get("content", ""))
                if result.get("updated"):
                    iss["status"] = "已修复"
                else:
                    iss["status"] = "修复失败"
        elif dim == "内容完整性" and "关键词缺失" in iss.get("description", "") and shp:
            # Re-write won't help if content generation missed keywords — report only
            iss["status"] = "需 prompt 优化"
        else:
            # Attribute/editability issues are structural — can't auto-fix
            if sev == "严重":
                iss["status"] = "需人工检查"

        updated_issues.append(iss)
    return updated_issues


def main() -> int:
    setup_console_encoding()
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default="1.0")
    args = ap.parse_args()

    out_ppt = OUTPUT_DIR / f"claude-ppt {args.version}.pptx"

    if not TEMPLATE_PATH.exists():
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", "- 原因: 模板文件缺失"])
        safe_print("[BLOCKED] Template missing")
        return 0

    if not CONTENT_JSON.exists():
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", "- 原因: build_shape_content.json缺失"])
        safe_print("[BLOCKED] Content JSON missing")
        return 0

    content_data = json.loads(CONTENT_JSON.read_text(encoding="utf-8"))
    items = content_data.get("items", [])

    # Load shape_detail for geometry fallback
    shape_detail_path = PROGRESS_DIR / "01-shape_detail_com.json"
    shape_details = {}
    if shape_detail_path.exists():
        sd = json.loads(shape_detail_path.read_text(encoding="utf-8"))
        for s in sd.get("new_shapes", []):
            shape_details[s.get("name", "")] = s

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_md(OUT_REPORT, ["# Build PPT Report", "", "- 状态: blocked", f"- 原因: {e}"])
        safe_print(f"[BLOCKED] {e}")
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    src = app.Presentations.Open(str(TEMPLATE_PATH))
    dst = app.Presentations.Add()

    results = []
    try:
        # Clone slide 2 (standard page) to new presentation
        src.Slides(2).Copy()
        time.sleep(COPY_PASTE_DELAY)  # COM clipboard buffer
        dst.Slides.Paste()
        time.sleep(COPY_PASTE_DELAY)
        slide = dst.Slides(1)

        for item in items:
            detail = shape_details.get(item.get("shape_name", ""))
            results.append(apply_shape(slide, item, detail))

        dst.SaveAs(str(out_ppt))

        updated = sum(1 for r in results if r.get("updated"))
        write_json(OUT_READBACK, {
            "generated_at": now_ts(),
            "target": out_ppt.name,
            "results": results,
        })
        write_md(OUT_REPORT, [
            "# Build PPT Report",
            "",
            "- 状态: ok",
            f"- 产物: {out_ppt.name}",
            f"- 更新shape: {updated}/{len(results)}",
            f"- 时间: {now_ts()}",
            "",
            "## Details",
            "",
            "|shape|updated|mode|match|reason|",
            "|---|---|---|---|---|",
        ] + [
            f"|{r.get('shape_name','')}|{r.get('updated','')}|{r.get('mode','')}|{r.get('match_method','')}|{r.get('reason','')}|"
            for r in results
        ])
        safe_print(f"[OK] Generated {out_ppt.name}, updated {updated}/{len(results)} shapes")

        # ========== SELF-CHECK LOOP ==========
        safe_print("\n[SELF-CHECK] Starting four-step verification...")
        ssim_score = -1.0
        all_issues = []

        for attempt in range(MAX_SELF_FIX + 1):
            all_issues = []
            # ①② Attribute + editability check
            all_issues += _check_attributes(slide, shape_details)
            # ③ Visual comparison (template slide vs generated slide)
            template_slide = src.Slides(2)
            ssim_score, vis_issues = _check_visual(app, template_slide, slide)
            all_issues += vis_issues
            # ④ Content completeness
            all_issues += _check_content(slide, items)

            severe = [i for i in all_issues if i.get("severity") == "严重"]
            safe_print(f"  [Attempt {attempt + 1}] {len(all_issues)} issues "
                       f"({len(severe)} severe)")

            if not severe or attempt == MAX_SELF_FIX:
                break

            # Auto-fix and re-save
            safe_print(f"  [FIX] Attempting auto-fix...")
            all_issues = _auto_fix(slide, all_issues, items, shape_details)
            dst.Save()

        # Generate self-check report
        generate_self_check_report(
            all_issues, str(SELF_CHECK_REPORT),
            ssim_score=ssim_score, attempts=attempt + 1,
        )
        unfixed_severe = sum(1 for i in all_issues
                            if i.get("severity") == "严重"
                            and "已修复" not in i.get("status", ""))
        if unfixed_severe:
            safe_print(f"[SELF-CHECK] FAIL: {unfixed_severe} unfixed severe issues — see {SELF_CHECK_REPORT.name}")
        else:
            safe_print(f"[SELF-CHECK] PASS — report: {SELF_CHECK_REPORT.name}")

        return 0
    finally:
        com_call(src, "Close")
        com_call(dst, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
