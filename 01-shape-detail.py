#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 1: extract新增shape详情（COM优先，环境不支持时写阻断报告）"""
import json
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUT_MD = ROOT / "shape-detail.md"
OUT_JSON = ROOT / "shape_detail_com.json"
PPT_PATH = ROOT / "src" / "Template 2.1.pptx"


def write_blocked(reason: str) -> None:
    OUT_JSON.write_text(json.dumps({
        "status": "blocked",
        "reason": reason,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "new_shapes": []
    }, ensure_ascii=False, indent=2), encoding="utf-8")
    OUT_MD.write_text(
        "# Shape Detail Report\n\n"
        f"- 状态: blocked\n- 原因: {reason}\n"
        f"- 时间: {datetime.now().isoformat(timespec='seconds')}\n"
        "\n> 当前环境不具备 Windows Office COM，无法读取 Template 2.1.pptx 的 shape 详情。\n",
        encoding="utf-8",
    )


def main() -> int:
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_blocked(f"win32com unavailable: {e}")
        print("[WARN] COM unavailable; wrote blocked reports.")
        return 0

    if not PPT_PATH.exists():
        write_blocked(f"template not found: {PPT_PATH}")
        print("[WARN] Template missing; wrote blocked reports.")
        return 0

    # 真正COM实现（Windows+Office环境下执行）
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True
    pres = app.Presentations.Open(str(PPT_PATH))
    try:
        blank = pres.Slides(14)
        std = pres.Slides(15)
        blank_names = set()
        for i in range(1, blank.Shapes.Count + 1):
            s = blank.Shapes(i)
            blank_names.add((s.Name or "", int(s.Type)))

        new_shapes = []
        for i in range(1, std.Shapes.Count + 1):
            s = std.Shapes(i)
            key = ((s.Name or ""), int(s.Type))
            if key in blank_names:
                continue
            name = s.Name or f"auto_shape_{i}"
            item = {
                "slide_index": 15,
                "name": name,
                "left": float(s.Left),
                "top": float(s.Top),
                "width": float(s.Width),
                "height": float(s.Height),
                "shape_type": int(s.Type),
                "text": "",
                "in_group": bool(getattr(s, "ParentGroup", None)),
            }
            if getattr(s, "HasTextFrame", 0) and s.TextFrame.HasText:
                item["text"] = s.TextFrame.TextRange.Text
            new_shapes.append(item)

        OUT_JSON.write_text(json.dumps({"status": "ok", "new_shapes": new_shapes}, ensure_ascii=False, indent=2), encoding="utf-8")
        lines = ["# Shape Detail Report", "", f"- 新增shape数量: {len(new_shapes)}", ""]
        for idx, it in enumerate(new_shapes, 1):
            lines += [f"## {idx}. {it['name']}", f"- left/top: {it['left']:.2f}/{it['top']:.2f}", f"- width/height: {it['width']:.2f}/{it['height']:.2f}", f"- type: {it['shape_type']}", f"- text: {it['text']}", ""]
        OUT_MD.write_text("\n".join(lines), encoding="utf-8")
        print(f"[OK] Wrote {OUT_MD.name} and {OUT_JSON.name}")
        return 0
    finally:
        pres.Close()
        app.Quit()


if __name__ == "__main__":
    raise SystemExit(main())
