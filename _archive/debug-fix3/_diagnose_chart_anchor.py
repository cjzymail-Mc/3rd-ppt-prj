#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""桥接当前 Excel sheet，复现 _prepare_apparel_chart_data 的 anchor 计算 +
读 anchor 周边实际数据，看 CurrentRegion 为什么吞主问卷区。

复现 src/apparel_ppt.py:636-676 + 700-706 的逻辑。
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
from src.apparel_ppt import _extract_means_for_category, _xlwings_to_rows


def main():
    book = xw.books.active
    sht = None
    for s in book.sheets:
        if "问卷" in s.name:
            sht = s
            break
    if sht is None:
        sht = book.sheets.active
    print(f"=== Active workbook: {book.name}")
    print(f"=== Active sheet:    {sht.name}\n")

    # 1) origin（apparel 调 get_range(mc_sht)）
    origin = get_range(sht)
    print(f"--- origin = {origin.address}  (row={origin.row}, col={origin.column})")

    # 2) origin.CurrentRegion 看主问卷数据区范围
    cr = origin.api.CurrentRegion
    cr_addr = cr.Address
    cr_rows = cr.Rows.Count
    cr_cols = cr.Columns.Count
    print(f"--- origin.CurrentRegion = {cr_addr}")
    print(f"    rows={cr_rows}, cols={cr_cols}")
    print(f"    (主问卷数据区在 row {origin.row} ~ {origin.row + cr_rows - 1}, "
          f"col {origin.column} ~ {origin.column + cr_cols - 1})\n")

    # 3) 复现 apparel 4 个 chart 的 anchor offset 累积
    rows = _xlwings_to_rows(sht)
    print(f"--- _xlwings_to_rows: {len(rows)} 行 × {len(rows[0]) if rows else 0} 列\n")

    print("--- 复现 4 个 chart 的 anchor 位置 ---")
    chart_anchor_offset = 50
    for cat in ["版型", "面料", "吸湿排汗", "速干"]:
        means = _extract_means_for_category(rows, cat)
        content_lines = len(means) if means else 1
        anchor_row = origin.row + cr_rows + chart_anchor_offset
        anchor_col = origin.column
        anchor_cell = sht.cells(anchor_row, anchor_col)
        print(f"  [{cat}] offset={chart_anchor_offset} → anchor={anchor_cell.address}  "
              f"(row={anchor_row}, col={anchor_col})  content={content_lines} 行")

        # 看 anchor 周边的实际数据
        # 写入会占 anchor_row ~ anchor_row + content_lines 行 × 2 列
        try:
            top_value = sht.cells(anchor_row, anchor_col).value
            top_right = sht.cells(anchor_row, anchor_col + 1).value
            print(f"      anchor cell value={top_value!r}, right={top_right!r}")
        except Exception as e:
            print(f"      anchor cell read err: {e}")

        # 关键：模拟 _prepare_apparel_chart_data 写完后取 CurrentRegion
        # 注意：当前 anchor 可能还没数据（如果上次跑被清理）或有残留数据
        try:
            anchor_cr = anchor_cell.api.CurrentRegion
            print(f"      → 当前 anchor 的 CurrentRegion = {anchor_cr.Address}  "
                  f"rows={anchor_cr.Rows.Count}  cols={anchor_cr.Columns.Count}")
        except Exception as e:
            print(f"      anchor CR err: {e}")

        chart_anchor_offset += content_lines + 5
        print()

    # 4) 看 sheet 总用量
    used = sht.used_range
    print(f"--- used_range: {used.address}  rows={used.rows.count}  cols={used.columns.count}\n")

    # 5) 关键诊断：origin.CurrentRegion 的最后一行 + 50 是否跟主数据区相隔够远
    first_anchor_row = origin.row + cr_rows + 50
    gap = first_anchor_row - (origin.row + cr_rows - 1)
    print(f"--- 第 1 个 anchor 与主问卷数据区间距: {gap} 空行")
    print(f"    主数据区末行 = row {origin.row + cr_rows - 1}")
    print(f"    第 1 anchor  = row {first_anchor_row}")
    print(f"    > 间距 >= 1 行通常足够阻断 CurrentRegion，但要看 anchor 那一带是否有别的残留数据")


if __name__ == "__main__":
    main()
