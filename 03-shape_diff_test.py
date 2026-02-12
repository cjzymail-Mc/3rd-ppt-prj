#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 3: shape diff test with COM; output fix-ppt.md."""
import argparse
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "src" / "Template 2.1.pptx"
FIX = ROOT / "fix-ppt.md"


def write_blocked(reason: str) -> None:
    FIX.write_text(
        "# PPT Diff & Fix Report\n\n"
        f"- 状态: blocked\n- 原因: {reason}\n"
        f"- 时间: {datetime.now().isoformat(timespec='seconds')}\n"
        "\n> 当前环境不可完成 COM 级差异比对，请在 Windows + Office 环境运行。\n",
        encoding="utf-8",
    )


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--target", default="codex 1.0.pptx")
    args = ap.parse_args()
    target = ROOT / args.target

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_blocked(f"win32com unavailable: {e}")
        print("[WARN] COM unavailable; wrote fix-ppt.md")
        return 0

    if not TEMPLATE.exists() or not target.exists():
        write_blocked("template or target ppt missing")
        print("[WARN] Missing input ppt; wrote fix-ppt.md")
        return 0

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = 0
    p1 = app.Presentations.Open(str(TEMPLATE), WithWindow=False)
    p2 = app.Presentations.Open(str(target), WithWindow=False)
    try:
        s1 = p1.Slides(15)
        s2 = p2.Slides(1)
        c1 = s1.Shapes.Count
        c2 = s2.Shapes.Count
        FIX.write_text(
            "# PPT Diff & Fix Report\n\n"
            f"- 状态: ok\n- template_shapes: {c1}\n- target_shapes: {c2}\n"
            f"- 时间: {datetime.now().isoformat(timespec='seconds')}\n"
            "\n> 详细字段级diff可继续扩展。\n",
            encoding="utf-8",
        )
        print("[OK] Wrote fix-ppt.md")
        return 0
    finally:
        p1.Close()
        p2.Close()
        app.Quit()


if __name__ == "__main__":
    raise SystemExit(main())
