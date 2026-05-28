"""Snapshot active PowerPoint to template/apparel-page13-14-template.pptx.

GetActiveObject 桥接当前打开的 v1.4，用 SaveCopyAs 另存一份给 dev 做基准。
不动用户文件，不 Close 不 Quit。
"""
from __future__ import annotations

import os
import sys

import win32com.client

OUT_PATH = os.path.abspath(
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "template",
        "apparel-page13-14-template.pptx",
    )
)


def main() -> int:
    try:
        ppt = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as exc:
        print(f"[snapshot] FAIL: 没有运行中的 PowerPoint: {exc}")
        return 1

    pres = ppt.ActivePresentation
    src = pres.FullName
    print(f"[snapshot] 源 PPT: {src}")
    print(f"[snapshot] 目标:   {OUT_PATH}")

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    pres.SaveCopyAs(OUT_PATH)
    print("[snapshot] SaveCopyAs 完成")
    return 0


if __name__ == "__main__":
    sys.exit(main())
