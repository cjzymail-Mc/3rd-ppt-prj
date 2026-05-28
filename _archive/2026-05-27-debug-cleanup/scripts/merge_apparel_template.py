#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""merge_apparel_template.py — 一次性预处理：把 apparel-page13-14-template.pptx
的 slide 13 和 slide 14 合并（InsertFromFile）到 **Main.py 实际打开的模板**：
src/Template 2.1.pptx。

合并后 src/Template 2.1.pptx 总页数：原有 19 页 + 新增 2 页 = 21 页。
新函数 _TEMPLATE_P13_SLIDE = 20、_TEMPLATE_P14_SLIDE = 21。

用法（一次性运行；用户当前打开的 v1.4.pptx 不会受影响）：
    python debug/merge_apparel_template.py

运行后会打印最终 Slides.Count 和每页 shape 数，用于核验。
"""

import sys
import time
from pathlib import Path

# 确保项目根在 sys.path
_PROJ_ROOT = str(Path(__file__).resolve().parent.parent)
if _PROJ_ROOT not in sys.path:
    sys.path.insert(0, _PROJ_ROOT)

import win32com.client

_TEMPLATE_BASE = _PROJ_ROOT + r"\src\Template 2.1.pptx"
_TEMPLATE_SRC  = _PROJ_ROOT + r"\template\apparel-page13-14-template.pptx"


def main():
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True

    print(f"打开基础模板: {_TEMPLATE_BASE}")
    base_ppt = app.Presentations.Open(_TEMPLATE_BASE)
    print(f"  当前 Slides.Count = {base_ppt.Slides.Count}")

    # 打开源模板（只读）
    print(f"打开源模板: {_TEMPLATE_SRC}")
    src_ppt = app.Presentations.Open(_TEMPLATE_SRC, ReadOnly=True)
    src_count = src_ppt.Slides.Count
    print(f"  源模板 Slides.Count = {src_count}（应 >= 14）")

    if src_count < 14:
        print(f"错误：源模板只有 {src_count} 页，需要至少 14 页（含 slide 13 / 14）")
        src_ppt.Close()
        base_ppt.Close()
        return

    # InsertFromFile 将源模板的特定 slide 复制到 base_ppt 末尾
    # 参数：FileName, Index, SlideStart, SlideEnd
    # 注：InsertFromFile 的 SlideStart/SlideEnd 是 1-based
    base_count_before = base_ppt.Slides.Count

    print(f"\n插入 slide 13 (源模板第 13 页)...")
    base_ppt.Slides.InsertFromFile(_TEMPLATE_SRC, base_count_before, 13, 13)
    time.sleep(1.0)

    base_count_mid = base_ppt.Slides.Count
    print(f"  插入后 Slides.Count = {base_count_mid}  → slide {base_count_mid} 是新 p13")
    p13_idx = base_count_mid

    print(f"\n插入 slide 14 (源模板第 14 页)...")
    base_ppt.Slides.InsertFromFile(_TEMPLATE_SRC, base_count_mid, 14, 14)
    time.sleep(1.0)

    base_count_after = base_ppt.Slides.Count
    print(f"  插入后 Slides.Count = {base_count_after}  → slide {base_count_after} 是新 p14")
    p14_idx = base_count_after

    # 保存
    print(f"\n保存 {_TEMPLATE_BASE} ...")
    base_ppt.Save()
    time.sleep(0.5)

    # 验证
    print(f"\n=== 合并结果 ===")
    for i in range(1, base_count_after + 1):
        sld = base_ppt.Slides(i)
        n_shapes = sld.Shapes.Count
        print(f"  Slide {i}: {n_shapes} shapes")

    print(f"\n请将以下常量写入 src/apparel_ppt.py:")
    print(f"  _TEMPLATE_P13_SLIDE = {p13_idx}")
    print(f"  _TEMPLATE_P14_SLIDE = {p14_idx}")

    src_ppt.Close()
    print("\n完成。基础模板已保存，源模板已关闭。")


if __name__ == "__main__":
    main()
