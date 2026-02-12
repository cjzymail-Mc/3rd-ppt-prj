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


def com_get(obj, attr: str, default=None):
    """安全读取COM属性：PowerPoint属性不适用时会直接抛com_error。"""
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def com_call(obj, method: str, *args, **kwargs):
    """安全调用COM方法。"""
    try:
        fn = getattr(obj, method)
        return fn(*args, **kwargs)
    except Exception:
        return None


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
            blank_names.add((com_get(s, "Name", "") or "", int(com_get(s, "Type", 0) or 0)))

        new_shapes = []
        for i in range(1, std.Shapes.Count + 1):
            s = std.Shapes(i)
            key = ((com_get(s, "Name", "") or ""), int(com_get(s, "Type", 0) or 0))
            if key in blank_names:
                continue

            name = com_get(s, "Name", "") or f"auto_shape_{i}"
            has_text_frame = bool(com_get(s, "HasTextFrame", 0))
            has_text = False
            if has_text_frame:
                tf = com_get(s, "TextFrame", None)
                if tf is not None:
                    has_text = bool(com_get(tf, "HasText", 0))

            text_value = ""
            if has_text:
                tf = com_get(s, "TextFrame", None)
                tr = com_get(tf, "TextRange", None) if tf is not None else None
                text_value = com_get(tr, "Text", "") or ""

            # ParentGroup 在很多shape上会直接抛COM异常，不可直接 getattr
            parent_group = com_get(s, "ParentGroup", None)
            in_group = parent_group is not None

            item = {
                "slide_index": 15,
                "name": name,
                "left": float(com_get(s, "Left", 0.0) or 0.0),
                "top": float(com_get(s, "Top", 0.0) or 0.0),
                "width": float(com_get(s, "Width", 0.0) or 0.0),
                "height": float(com_get(s, "Height", 0.0) or 0.0),
                "shape_type": int(com_get(s, "Type", 0) or 0),
                "text": text_value,
                "in_group": in_group,
            }
            new_shapes.append(item)

        OUT_JSON.write_text(
            json.dumps({"status": "ok", "new_shapes": new_shapes}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        lines = ["# Shape Detail Report", "", f"- 新增shape数量: {len(new_shapes)}", ""]
        for idx, it in enumerate(new_shapes, 1):
            lines += [
                f"## {idx}. {it['name']}",
                f"- left/top: {it['left']:.2f}/{it['top']:.2f}",
                f"- width/height: {it['width']:.2f}/{it['height']:.2f}",
                f"- type: {it['shape_type']}",
                f"- text: {it['text']}",
                f"- in_group: {it['in_group']}",
                "",
            ]
        OUT_MD.write_text("\n".join(lines), encoding="utf-8")
        print(f"[OK] Wrote {OUT_MD.name} and {OUT_JSON.name}")
        return 0
    finally:
        if pres is not None:
            com_call(pres, "Close")
        if app is not None:
            com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
