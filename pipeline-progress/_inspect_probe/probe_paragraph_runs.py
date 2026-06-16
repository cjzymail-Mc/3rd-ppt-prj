#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自验证脚本：对 apparel-page13-14-template.pptx p13 的 4 个指定 shape
跑 extract_paragraph_runs，打印每段的 merged runs。

运行方式（无需 PPT 已打开）：
    python pipeline-progress/_inspect_probe/probe_paragraph_runs.py

COM 策略：DispatchEx 新进程 ReadOnly 打开，finally 无条件 Quit。
"""
from __future__ import annotations

import sys
import os
import time

# ── 把 acceptance skill 根目录加入 sys.path，让 paragraph_runs 可 import ──
SKILL_ROOT = r"C:\Users\xy24\.claude\skills\ppt-acceptance-check"
if SKILL_ROOT not in sys.path:
    sys.path.insert(0, SKILL_ROOT)

from paragraph_runs import extract_paragraph_runs, MERGE_DIMS  # noqa: E402

# ── 模板路径 ──
TEMPLATE_PATH = (
    r"D:\Technique Support\Claude Code Learning\3rd-ppt-prj"
    r"\template\apparel-page13-14-template.pptx"
)

# ── 目标 shape（均在 p13，即 Slides(1)）──
TARGET_SHAPES = [
    "TextBox 6",
    "TextBox 50",
    "Rounded Rectangle 53",
    "Rounded Rectangle 55",
]

# ── 期望摘要（仅供肉眼核对，脚本不做 assert 避免 hardcode 自证）──
EXPECTED_NOTES = {
    "TextBox 6":            "2段: p1=[20pt 黑] p2=[16pt 红]，各1 merged run",
    "TextBox 50":           "2段: p2 '15~25℃' 拆成2 run → 合并后应是1个 16pt 红 run",
    "Rounded Rectangle 53": "2段: p1=[11pt 白] p2=[24pt 白]，各1 merged run",
    "Rounded Rectangle 55": "1段: 段内2 run（11pt白 + 24pt白，size不同不合并）",
}


def safe_print(*args):
    line = " ".join(str(a) for a in args)
    try:
        print(line)
    except UnicodeEncodeError:
        print(line.encode(sys.stdout.encoding or "ascii", "replace")
              .decode(sys.stdout.encoding or "ascii"))


def shape_by_name(slide, name: str):
    try:
        for i in range(1, slide.Shapes.Count + 1):
            s = slide.Shapes(i)
            if str(s.Name) == name:
                return s
    except Exception:
        pass
    return None


def main():
    import win32com.client

    app = None
    prs = None
    try:
        safe_print(f"[probe] DispatchEx PowerPoint.Application (新进程)")
        app = win32com.client.DispatchEx("PowerPoint.Application")

        safe_print(f"[probe] ReadOnly open: {TEMPLATE_PATH}")
        prs = app.Presentations.Open(
            TEMPLATE_PATH,
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )
        time.sleep(0.5)

        slide = prs.Slides(13)  # 模板 p13 = Slides(13)
        safe_print(f"[probe] Slide 13（模板 p13），Shape count = {slide.Shapes.Count}")
        safe_print(f"[probe] MERGE_DIMS = {MERGE_DIMS}\n")

        for shape_name in TARGET_SHAPES:
            safe_print("=" * 60)
            safe_print(f"Shape: {shape_name}")
            note = EXPECTED_NOTES.get(shape_name, "")
            if note:
                safe_print(f"  预期: {note}")

            shp = shape_by_name(slide, shape_name)
            if shp is None:
                safe_print(f"  [ERROR] shape not found")
                continue

            paras = extract_paragraph_runs(shp)
            safe_print(f"  段落数 (para count) = {len(paras)}")
            for p in paras:
                safe_print(f"  -- para_idx={p['para_idx']}  alignment={p['alignment']}"
                           f"  text={p['text']!r}")
                runs = p["runs"]
                safe_print(f"     merged runs count = {len(runs)}")
                for r in runs:
                    safe_print(
                        f"     run: rgb={r['rgb']}  bold={r['bold']}  size={r['size']}"
                        f"  italic={r['italic']}  font={r['font_name']!r}"
                        f"  text={r['text']!r}"
                    )
            safe_print("")

    finally:
        if prs is not None:
            try:
                prs.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        safe_print("[probe] PowerPoint Quit. Done.")


if __name__ == "__main__":
    main()
