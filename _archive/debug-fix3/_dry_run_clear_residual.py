#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dry-run：桥接当前 Excel，计算 _clear_residual_outside_valid_region 会清的范围，
**只算不动**，给用户确认范围合理性。
"""
from __future__ import annotations
import io, sys
from pathlib import Path
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import xlwings as xw
from src.Function_030 import get_range


def main():
    book = xw.books.active
    sht = None
    for s in book.sheets:
        if "问卷" in s.name:
            sht = s; break
    if sht is None:
        sht = book.sheets.active
    print(f"=== Active workbook: {book.name}")
    print(f"=== Active sheet:    {sht.name}\n")

    mc_cell0 = get_range(sht)
    print(f"--- mc_cell0 = {mc_cell0.address}  (row={mc_cell0.row}, col={mc_cell0.column})")

    cr = mc_cell0.api.CurrentRegion
    valid_last_row = mc_cell0.row + cr.Rows.Count - 1
    valid_last_col = mc_cell0.column + cr.Columns.Count - 1
    print(f"--- 有效数据区: row {mc_cell0.row}~{valid_last_row}, col {mc_cell0.column}~{valid_last_col}  "
          f"({cr.Rows.Count} rows × {cr.Columns.Count} cols)")

    used = sht.used_range
    used_last_row = used.last_cell.row
    used_last_col = used.last_cell.column
    print(f"--- used_range : {used.address}  (last row={used_last_row}, last col={used_last_col})\n")

    print("--- dry-run 清理范围（**只算不动**）:")
    will_clear_rows = 0
    will_clear_cols = 0

    if used_last_row > valid_last_row:
        end_col = max(used_last_col, valid_last_col)
        print(f"  下方残留: 清 row {valid_last_row + 1}~{used_last_row}, col 1~{end_col}")
        print(f"            = {used_last_row - valid_last_row} 行 × {end_col} 列")
        will_clear_rows = used_last_row - valid_last_row
    else:
        print(f"  下方残留: 无（used_range 没超出有效区下边界）")

    if used_last_col > valid_last_col:
        print(f"  右侧残留: 清 row 1~{valid_last_row}, col {valid_last_col + 1}~{used_last_col}")
        print(f"            = {valid_last_row} 行 × {used_last_col - valid_last_col} 列")
        will_clear_cols = used_last_col - valid_last_col
    else:
        print(f"  右侧残留: 无")

    print()
    if will_clear_rows or will_clear_cols:
        # 抽样下方残留前 3 行的 col A-B 内容，让用户看清要清什么
        print(f"--- 下方残留抽样（前 3 行 col A-D）:")
        for r in range(valid_last_row + 1, min(valid_last_row + 4, used_last_row + 1)):
            vals = []
            for c in range(1, 5):
                try:
                    vals.append(str(sht.cells(r, c).value)[:15])
                except Exception:
                    vals.append("?")
            print(f"  row {r}: {vals}")
    else:
        print("--- sheet 已干净，无残留")


if __name__ == "__main__":
    main()
