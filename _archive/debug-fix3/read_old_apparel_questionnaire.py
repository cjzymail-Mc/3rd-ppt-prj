#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""debug/read_old_apparel_questionnaire.py — dump 旧 apparel 问卷模板列结构.

为对比新问卷 schema 用。
打开独立 xlwings App，不影响用户当前 Excel 会话。
"""
from __future__ import annotations

import io
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import xlwings as xw

_proj = Path(__file__).resolve().parent.parent
old_path = _proj / "template" / "source data-apparel.xlsx"

app = xw.App(visible=False, add_book=False)
try:
    book = app.books.open(str(old_path), update_links=False, read_only=True)
    for sht in book.sheets:
        print(f"\n==== Sheet: {sht.name} ====")
        used = sht.used_range
        print(f"used_range: {used.address}  rows={used.rows.count}  cols={used.columns.count}")
        if used.rows.count < 1:
            continue
        raw = used.value
        if isinstance(raw, tuple):
            rows = [list(r) if isinstance(r, tuple) else [r] for r in raw]
        elif isinstance(raw, list):
            rows = raw if (raw and isinstance(raw[0], list)) else [raw]
        else:
            rows = [[raw]]
        if not rows:
            continue
        headers = [("" if h is None else str(h).strip()) for h in rows[0]]
        print("Headers:")
        for i, h in enumerate(headers):
            print(f"  [{i:02d}] {h}")
        # 第一行数据样例
        if len(rows) > 1:
            print("First data row:")
            for i, v in enumerate(rows[1][:len(headers)]):
                s = "" if v is None else str(v).strip()
                if len(s) > 40:
                    s = s[:40] + "…"
                print(f"  [{i:02d}] {s}")
        print(f"data rows: {len(rows) - 1}")
    book.close()
finally:
    app.quit()
