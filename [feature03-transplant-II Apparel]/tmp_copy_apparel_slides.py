#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""tmp_copy_apparel_slides.py — 一次性脚本

把 template/empty and standard-apparel.pptx 的所有 slide
追加到 src/Template 2.1.pptx 的末尾，保留原始美工 / 母版 / 动画。

策略：win32com COM 跨文件 Slides.Copy() → Slides.Paste(idx)
（python-pptx 禁用，COM 才能保留母版与原始格式）
"""

from __future__ import annotations

import os
import sys
import time
import random

import win32com.client  # type: ignore

ROOT = os.path.abspath(os.path.dirname(__file__))
SRC_PPT = os.path.join(ROOT, "template", "empty and standard-apparel.pptx")
DST_PPT = os.path.join(ROOT, "src", "Template 2.1.pptx")

COPY_PASTE_DELAY = 0.6  # COM 剪贴板缓冲


def main() -> int:
    if not os.path.isfile(SRC_PPT):
        print(f"[ERR] 源文件不存在: {SRC_PPT}")
        return 1
    if not os.path.isfile(DST_PPT):
        print(f"[ERR] 目标文件不存在: {DST_PPT}")
        return 1

    print(f"[INFO] 源:  {SRC_PPT}")
    print(f"[INFO] 目标: {DST_PPT}")

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0
    app.Visible = True

    # 必须以可写模式打开目标；ReadOnly=False
    src = app.Presentations.Open(SRC_PPT, ReadOnly=True, WithWindow=True)
    dst = app.Presentations.Open(DST_PPT, ReadOnly=False, WithWindow=True)

    src_count = src.Slides.Count
    dst_count_before = dst.Slides.Count
    print(f"[INFO] 源 slide 数: {src_count}")
    print(f"[INFO] 目标 slide 数（追加前）: {dst_count_before}")

    pasted = 0
    try:
        for i in range(1, src_count + 1):
            src.Slides(i).Copy()
            time.sleep(COPY_PASTE_DELAY + random.random() * 0.2)

            # 追加到末尾：paste(目标当前长度 + 1)
            target_idx = dst.Slides.Count + 1
            dst.Slides.Paste(target_idx)
            time.sleep(COPY_PASTE_DELAY + random.random() * 0.2)
            pasted += 1
            print(f"  [OK] 复制 src.slide({i}) → dst.slide({target_idx})")

        dst.Save()
        print(f"[DONE] 已追加 {pasted} 张 slide，目标当前共 {dst.Slides.Count} 张。")
    finally:
        try:
            src.Close()
        except Exception:
            pass
        # 不关闭 dst，让用户能直接看到结果；用户手动关或下游脚本处理
        # 如果想自动关，取消下面注释
        # dst.Close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
