#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 1: Compare Template page 14 (blank) vs page 15 (standard).

Extracts new shapes' attributes and fingerprints.
Outputs:
  - shape_detail_com.json   (for agents / code)
  - shape_fingerprint_map.json
  - shape_detail.xlsx        (for user review + annotation in Excel)

Human-in-the-loop:
  If shape_detail.xlsx already contains user annotations, Step 1 re-extracts
  shapes from the template and merges existing annotations into the new xlsx.
  Shapes with annotations have them restored; new shapes get empty placeholders.
  Use --force to clear all annotations and start fresh.
"""

from __future__ import annotations

import sys
from pathlib import Path

# ensure project root is on sys.path so pipeline package can find ppt_pipeline_common
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# ── import canonical paragraph-aware walker (Step 1b, plan §3.5.2 边界2) ──
# 权威 walker 定义在 acceptance skill 侧；Step1 跨边界 import（不 vendor 镜像）。
# 用 Path.home() 不 hardcode 用户名。
_SKILL_DIR = Path.home() / ".claude" / "skills" / "ppt-acceptance-check"
if str(_SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(_SKILL_DIR))
try:
    from paragraph_runs import extract_paragraph_runs, MERGE_DIMS  # noqa: F401
    _WALKER_AVAILABLE = True
except ImportError:
    _WALKER_AVAILABLE = False

    def extract_paragraph_runs(shape):  # type: ignore[misc]
        """Fallback stub if skill not installed — returns empty list."""
        return []

from pipeline.ppt_pipeline_common import (
    PROGRESS_DIR,
    SHAPE_DETAIL_XLSX,
    TEMPLATE_PATH,
    com_get,
    com_call,
    generate_shape_detail_xlsx,
    parse_user_annotations,
    is_in_group,
    now_ts,
    safe_print,
    safe_text,
    setup_console_encoding,
    write_json,
)

OUT_JSON = PROGRESS_DIR / "01-shape_detail_com.json"
OUT_FP = PROGRESS_DIR / "01-shape_fingerprint_map.json"
OUT_XLSX = SHAPE_DETAIL_XLSX  # pipeline-progress/01-shape_detail.xlsx


def _safe_has_chart(shp) -> bool:
    """COM-safe check for HasChart."""
    try:
        return bool(shp.HasChart)
    except Exception:
        return False


def _safe_text(shp) -> str:
    """COM-safe extraction of shape text."""
    try:
        return shp.TextFrame.TextRange.Text or ""
    except Exception:
        return ""


def _safe_font(shp):
    """COM-safe extraction of font name and size."""
    try:
        font = shp.TextFrame.TextRange.Font
        name = safe_text(com_get(font, "Name", ""))
        size = float(com_get(font, "Size", 0.0) or 0.0)
        return name, size
    except Exception:
        return "", 0.0


def shape_obj(slide_index: int, shp, idx: int) -> dict:
    name = com_get(shp, "Name", "") or f"shape_{slide_index}_{idx}"
    text = _safe_text(shp)
    font_name, font_size = _safe_font(shp)

    # plan §3.5.2 / §7: 新增 paragraphs 键（paragraph-aware感知），不改已有键。
    # paragraphs 不进 diff key 元组（diff key 用在下方 blank_keys/key，不含此字段）。
    paragraphs = extract_paragraph_runs(shp)

    return {
        "slide_index": slide_index,
        "name": safe_text(name),
        "left": float(com_get(shp, "Left", 0.0) or 0.0),
        "top": float(com_get(shp, "Top", 0.0) or 0.0),
        "width": float(com_get(shp, "Width", 0.0) or 0.0),
        "height": float(com_get(shp, "Height", 0.0) or 0.0),
        "shape_type": int(com_get(shp, "Type", 0) or 0),
        "text": text,
        "font_name": font_name,
        "font_size": font_size,
        "has_chart": _safe_has_chart(shp),
        "in_group": is_in_group(shp),
        "z_order": int(com_get(shp, "ZOrderPosition", 0) or 0),
        "paragraphs": paragraphs,  # NEW: paragraph-aware run 矩阵（空 list = 无文本）
    }


def fingersafe_print(item: dict) -> dict:
    return {
        "name": item["name"],
        "shape_type": item["shape_type"],
        "left": round(item["left"], 1),
        "top": round(item["top"], 1),
        "width": round(item["width"], 1),
        "height": round(item["height"], 1),
        "z_order": item["z_order"],
        "text_prefix": (item["text"] or "")[:20],
        "in_group": item["in_group"],
    }


def main() -> int:
    setup_console_encoding()
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true",
                    help="Force re-extraction even if user annotations exist")
    args = ap.parse_args()

    # --- Read existing annotations before re-extraction (merge mode) ---
    # --force: start fresh, clear all annotations
    # default: re-extract shapes, restore existing annotations into new md
    if args.force:
        existing_annos: dict = {}
        safe_print("[INFO] --force: re-extracting shapes, annotations will be cleared.")
    else:
        existing_annos = parse_user_annotations()
        if existing_annos:
            safe_print(f"[INFO] Found {len(existing_annos)} annotated shape(s) — "
                       f"will merge into new md.")
        else:
            safe_print("[INFO] No existing annotations — generating fresh md.")

    if not TEMPLATE_PATH.exists():
        write_json(OUT_JSON, {"status": "blocked", "reason": "template missing", "new_shapes": []})
        safe_print("[BLOCKED] Template file not found")
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_json(OUT_JSON, {"status": "blocked", "reason": str(e), "new_shapes": []})
        safe_print(f"[BLOCKED] {e}")
        return 0

    # 隔离进程打开模板（DispatchEx，绝不 attach 用户正在编辑的 PowerPoint）。
    # Step1 只读模板，故 ReadOnly + WithWindow=False（镜像 inspect-ppt-template skill 的安全开法）。
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.DisplayAlerts = 0
    pres = app.Presentations.Open(str(TEMPLATE_PATH), ReadOnly=True, Untitled=False, WithWindow=False)

    try:
        blank = pres.Slides(1)
        std = pres.Slides(2)

        # Collect blank page shape keys for diffing
        blank_keys = set()
        for i in range(1, int(blank.Shapes.Count) + 1):
            s = shape_obj(1, blank.Shapes(i), i)
            key = (
                s["shape_type"],
                round(s["left"], 1),
                round(s["top"], 1),
                round(s["width"], 1),
                round(s["height"], 1),
                s["name"],
            )
            blank_keys.add(key)

        # Find shapes on page 2 that are NOT on page 1
        new_shapes = []
        for i in range(1, int(std.Shapes.Count) + 1):
            s = shape_obj(2, std.Shapes(i), i)
            key = (
                s["shape_type"],
                round(s["left"], 1),
                round(s["top"], 1),
                round(s["width"], 1),
                round(s["height"], 1),
                s["name"],
            )
            if key not in blank_keys:
                new_shapes.append(s)

        fp_items = [fingersafe_print(x) for x in new_shapes]

        # Write JSON artifacts (for agents / code)
        write_json(OUT_JSON, {
            "status": "ok",
            "generated_at": now_ts(),
            "template_slide": 2,
            "blank_slide": 1,
            "blank_shape_count": int(blank.Shapes.Count),
            "std_shape_count": int(std.Shapes.Count),
            "new_shapes": new_shapes,
        })
        write_json(OUT_FP, {
            "generated_at": now_ts(),
            "fingerprints": fp_items,
        })

        # ── 烤草稿契约（plan §3.5 / §6 Step1b）─────────────────────────────
        # 放在 generate_shape_detail_xlsx 之前：草稿契约是关键产物，xlsx 是人审便利产物。
        # xlsx 崩溃不应阻断草稿生成。
        # 对每个有 runs 的文本 shape 生成 paragraphs_match_signature 草稿规则。
        # 全部默认 severity: "warn"（护栏#4：严格度由人工 curate 决定，草稿不擅自定 must_fix）。
        # expected_paragraphs.runs 只保留 {rgb, bold, size}（MERGE_DIMS，护栏#1）。
        # 禁止写/改真契约（护栏#2）；草稿只落 pipeline-progress/。
        import re

        def _slug(shape_name: str) -> str:
            """shape 名 → id 合法 slug（非字母数字下划线替换为 _）。"""
            return re.sub(r"[^\w]", "_", shape_name).strip("_")

        draft_rules = []
        for s in new_shapes:
            paras = s.get("paragraphs") or []
            # 只取"至少一段含 run"的 shape
            has_runs = any(p.get("runs") for p in paras)
            if not has_runs:
                continue

            # expected_paragraphs: 只留 {rgb, bold, size}（MERGE_DIMS 三键）+ alignment
            ep = []
            for p in paras:
                runs_slim = [
                    {"rgb": r.get("rgb"), "bold": r.get("bold"), "size": r.get("size")}
                    for r in p.get("runs", [])
                ]
                ep.append({
                    "alignment": p.get("alignment"),
                    "runs": runs_slim,
                })

            draft_rules.append({
                "_comment": (
                    "AUTO-DRAFT by Step1 — 草稿需人审："
                    "① severity（升级类 shape 期望值可能是模板旧态，要用外部真相覆盖，护栏#2）"
                    "② 自由文本类应降 warn 或删，禁刚性 run 数断言（护栏#4）"
                ),
                "id": f"{_slug(s['name'])}_paras",
                "shape": s["name"],
                "layer": "L3",
                "slide": s["slide_index"],
                "check": "paragraphs_match_signature",
                "expected_paragraphs": ep,
                "severity": "warn",
            })

        out_draft = PROGRESS_DIR / "01-acceptance_draft.json"
        write_json(out_draft, {
            "_comment": (
                "Step1 自动生成的草稿契约片段。"
                "人工 curate 后并入 acceptance/{name}.json。禁直接当门禁。"
            ),
            "rules": draft_rules,
        })
        safe_print(f"[OK] Draft contract: {out_draft.name} ({len(draft_rules)} rules)")
        # ─────────────────────────────────────────────────────────────────────

        # Write human-readable XLSX — merge existing annotations if present
        # xlsx 是人审便利产物；失败时 warning，不阻断（JSON + 草稿已落地）。
        try:
            generate_shape_detail_xlsx(new_shapes, existing_annos=existing_annos)
        except Exception as _xlsx_err:
            safe_print(f"[WARN] xlsx 未生成（JSON 和草稿契约已落地）: {_xlsx_err}")

        safe_print(f"[OK] {len(new_shapes)} new shapes. "
              f"Wrote {OUT_JSON.name}, {OUT_FP.name}")
        return 0
    finally:
        com_call(pres, "Close")
        com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
