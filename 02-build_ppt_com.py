#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 2: build ppt by COM (blocked report in non-Windows Office env)."""
import argparse
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PPT_TEMPLATE = ROOT / "src" / "Template 2.1.pptx"
SHAPE_JSON = ROOT / "shape_detail_com.json"
REPORT = ROOT / "build-ppt-report.md"


def write_blocked(reason: str) -> None:
    REPORT.write_text(
        "# Build PPT Report\n\n"
        f"- 状态: blocked\n- 原因: {reason}\n"
        f"- 时间: {datetime.now().isoformat(timespec='seconds')}\n",
        encoding="utf-8",
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default="1.0")
    args = ap.parse_args()

    out_ppt = ROOT / f"codex {args.version}.pptx"

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_blocked(f"win32com unavailable: {e}")
        print("[WARN] COM unavailable; wrote build report.")
        return 0

    if not PPT_TEMPLATE.exists() or not SHAPE_JSON.exists():
        write_blocked("template or shape json missing")
        print("[WARN] Missing inputs; wrote build report.")
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 0
    src = app.Presentations.Open(str(PPT_TEMPLATE), WithWindow=False)
    dst = app.Presentations.Add()
    try:
        src.Slides(15).Copy()
        dst.Slides.Paste()
        dst.SaveAs(str(out_ppt))
        REPORT.write_text(
            "# Build PPT Report\n\n"
            f"- 状态: ok\n- 产物: `{out_ppt.name}`\n"
            f"- 时间: {datetime.now().isoformat(timespec='seconds')}\n",
            encoding="utf-8",
        )
        print(f"[OK] Generated {out_ppt.name}")
        return 0
    finally:
        src.Close()
        dst.Close()
        app.Quit()


if __name__ == "__main__":
    raise SystemExit(main())
