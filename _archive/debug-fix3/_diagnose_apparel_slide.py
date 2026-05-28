#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""扫描当前活动 PPT 全部 slide，找出 apparel 评测页（含 TextBox 24 + 4 个 Chart）
并 dump 详细 shape 状态，验证 fix3 各项。"""
from __future__ import annotations
import io, sys
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import win32com.client

MSO = {1:"AutoShape",3:"Chart",6:"Group",7:"OLE",13:"Picture",14:"Placeholder",17:"TextBox"}

def _safe(fn, d=""):
    try: return fn()
    except Exception as e: return f"<ERR:{e}>"

def is_apparel_slide(slide):
    """apparel 评测页特征：含 TextBox 24 + 至少 1 个 Chart"""
    names = []
    for j in range(1, slide.Shapes.Count + 1):
        try:
            names.append(slide.Shapes.Item(j).Name)
        except Exception:
            pass
    has_tb24 = any("24" in n for n in names)
    has_chart = False
    for j in range(1, slide.Shapes.Count + 1):
        try:
            if bool(slide.Shapes.Item(j).HasChart):
                has_chart = True; break
        except Exception:
            pass
    return has_tb24, has_chart, names

def main():
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        print(f"PPT_NOT_ACTIVE: {e}"); return 1
    ppt = app.ActivePresentation
    total = ppt.Slides.Count
    print(f"=== Active: {ppt.Name}\n=== Total slides: {total}\n")

    # 找 apparel 候选页（含 TextBox 24 + Chart 标记）
    candidates = []
    for i in range(1, total + 1):
        slide = ppt.Slides.Item(i)
        has_tb24, has_chart, names = is_apparel_slide(slide)
        score = (1 if has_tb24 else 0) + (1 if has_chart else 0)
        if score >= 1:
            candidates.append((i, score, has_tb24, has_chart, names))

    print("=== Candidate slides ===")
    for i, score, tb24, ch, names in candidates:
        print(f"  Slide {i}: score={score} TB24={tb24} Chart={ch} count={len(names)}")
    print()

    # 选最高分的页（apparel 应该 score=2，有 TextBox 24 + Chart）
    if not candidates:
        print("!! 没找到候选 apparel 评测页（没有 TextBox 24 + Chart 组合）")
        # dump 最后 3 页结构帮诊断
        print("\n=== 最后 3 页 shape 名字 ===")
        for i in range(max(1, total - 2), total + 1):
            slide = ppt.Slides.Item(i)
            names = []
            for j in range(1, slide.Shapes.Count + 1):
                try: names.append(slide.Shapes.Item(j).Name)
                except: pass
            print(f"  Slide {i} ({slide.Shapes.Count} shapes): {names}")
        return 0

    best = max(candidates, key=lambda x: (x[1], x[0]))
    slide_idx = best[0]
    slide = ppt.Slides.Item(slide_idx)
    print(f"=== 选中分析: Slide {slide_idx} ({slide.Shapes.Count} shapes) ===\n")

    for j in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes.Item(j)
        name = _safe(lambda: sh.Name)
        t = _safe(lambda: int(sh.Type), -1)
        L = _safe(lambda: round(float(sh.Left), 1))
        T = _safe(lambda: round(float(sh.Top), 1))
        W = _safe(lambda: round(float(sh.Width), 1))
        H = _safe(lambda: round(float(sh.Height), 1))
        print(f"[{j:02d}] {name!r} type={t}({MSO.get(t,'?')}) L={L} T={T} W={W} H={H}")
        if _safe(lambda: int(sh.HasTextFrame), 0):
            txt = _safe(lambda: sh.TextFrame.TextRange.Text, "")
            if isinstance(txt, str) and txt:
                disp = txt.replace("\r", "\\r").replace("\n", "\\n")
                if len(disp) > 300:
                    disp = disp[:300] + "..."
                print(f"     TEXT={disp!r}")
                # 统计 paragraph 行数（按 \r 切）
                lines = txt.split("\r")
                print(f"     LINES={len(lines)}")
        if _safe(lambda: bool(sh.HasChart), False):
            try:
                s1 = sh.Chart.SeriesCollection(1)
                vals = list(s1.Values)
                xs = list(s1.XValues)
                print(f"     CHART X={xs} Vals={vals}")
            except Exception as e:
                print(f"     CHART read err: {e}")
        print()

if __name__ == "__main__":
    sys.exit(main() or 0)
