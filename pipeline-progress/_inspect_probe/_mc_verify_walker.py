#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""主 Claude 独立验证：直接 import 权威 walker，跑 apparel 模板真 COM，确认段内合并生效。
不复用 developer 的 probe 脚本（防 hardcode 自证）。"""
import sys, os
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

SKILL = r"C:\Users\xy24\.claude\skills\ppt-acceptance-check"
sys.path.insert(0, SKILL)
from paragraph_runs import extract_paragraph_runs, MERGE_DIMS

import win32com.client as win32
TPL = os.path.abspath("template/apparel-page13-14-template.pptx")
app = win32.DispatchEx("PowerPoint.Application")
try:
    app.DisplayAlerts = 0
except Exception:
    pass
pres = app.Presentations.Open(TPL, ReadOnly=True, Untitled=False, WithWindow=False)
try:
    slide = pres.Slides(13)
    targets = ["TextBox 6", "TextBox 50", "Rounded Rectangle 53", "Rounded Rectangle 55"]
    by_name = {}
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        try:
            nm = str(s.Name)
        except Exception:
            continue
        if nm in targets:
            by_name[nm] = s
    print("MERGE_DIMS =", MERGE_DIMS)
    for nm in targets:
        s = by_name.get(nm)
        if s is None:
            print(f"[MISS] {nm} not found")
            continue
        paras = extract_paragraph_runs(s)
        print(f"\n=== {nm} : {len(paras)} paragraphs ===")
        for p in paras:
            runs = p["runs"]
            print(f"  p{p['para_idx']} align={p['alignment']} runs={len(runs)} text={p['text']!r}")
            for r in runs:
                print(f"      run rgb={r['rgb']} size={r['size']} bold={r['bold']} text={r['text']!r}")
finally:
    try:
        pres.Close()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass
