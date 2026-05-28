#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""read_selected_shape.py — 项目入口 wrapper。

真身在系统 skill 目录（单源真相）：
    C:\\Users\\xy24\\.claude\\skills\\read-selected-shape\\read_selected_shape.py

为什么保留这个 wrapper：
    - CLAUDE.md / fine-tuned-shapes.md / settings.local.json.bak 等 22 处旧引用
      都写 `python skills/read_selected_shape.py`，wrapper 让它们不用改路径
    - settings.local.json 的 Bash 白名单也匹配相对路径
    - 换机器 / 换账号时只需要改下面这一行 SYSTEM_PATH，不用动 repo 内 22 处

用法（兼容 / 原样保留）：
    python skills/read_selected_shape.py            # 精简版
    python skills/read_selected_shape.py --full     # 全参数版（run 级字体 + shadow + 3D + chart/table）
"""
from __future__ import annotations

import os
import runpy
import sys

# ---- 系统 skill 路径（机器级单源） ----------------------------------------
# 换机器时只改这里。优先用环境变量 override，方便 CI / 多账号场景。
SYSTEM_PATH = os.environ.get(
    "READ_SELECTED_SHAPE_IMPL",
    r"C:\Users\xy24\.claude\skills\read-selected-shape\read_selected_shape.py",
)

if not os.path.isfile(SYSTEM_PATH):
    print(
        f"[wrapper] 系统 skill 真身不存在: {SYSTEM_PATH}\n"
        f"  请检查路径，或设环境变量 READ_SELECTED_SHAPE_IMPL 指向正确位置。",
        file=sys.stderr,
    )
    sys.exit(2)

# 用 runpy 执行系统版脚本，保留 __main__ 语义 + sys.argv 透传
runpy.run_path(SYSTEM_PATH, run_name="__main__")
