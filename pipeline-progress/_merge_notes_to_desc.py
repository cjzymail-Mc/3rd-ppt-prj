# -*- coding: utf-8 -*-
"""One-shot migration: merge 备注 into 内容描述, then delete 备注 rows.

Only operates on the LAST sheet of 01-shape_detail.xlsx.
Run once, then delete this script.
"""
import sys
import time
import win32com.client

XLSX = r"D:\Technique Support\Claude Code Learning\3rd-ppt-prj\pipeline-progress\01-shape_detail.xlsx"

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(XLSX)
    ws = wb.Sheets(wb.Sheets.Count)
    sheet_name = ws.Name
    print(f"[INFO] Operating on last sheet: '{sheet_name}'")

    max_row = ws.UsedRange.Rows.Count
    print(f"[INFO] Total rows: {max_row}")

    # Pass 1: collect 内容描述 and 备注 row pairs per shape
    current_shape = None
    desc_rows = {}   # shape_name -> row number of 内容描述
    note_rows = {}   # shape_name -> row number of 备注
    in_annotation = False

    for r in range(1, max_row + 1):
        a = str(ws.Cells(r, 1).Value or "").strip()
        if a.startswith("Shape #"):
            current_shape = str(ws.Cells(r, 2).Value or "").strip()
            in_annotation = False
            continue
        if a == "用户批注":
            in_annotation = True
            continue
        if not in_annotation or not current_shape:
            continue
        if a == "内容描述":
            desc_rows[current_shape] = r
        elif a == "备注":
            note_rows[current_shape] = r

    print(f"[INFO] Found {len(desc_rows)} shapes with 内容描述, {len(note_rows)} with 备注")

    # Pass 2: merge 备注 into 内容描述
    merged = 0
    for shape_name, note_r in note_rows.items():
        note_val = str(ws.Cells(note_r, 2).Value or "").strip()
        if not note_val:
            continue
        desc_r = desc_rows.get(shape_name)
        if desc_r is None:
            continue
        desc_val = str(ws.Cells(desc_r, 2).Value or "").strip()
        if note_val in desc_val:
            print(f"  [SKIP] {shape_name}: 备注 already in 内容描述")
            continue
        new_desc = f"{desc_val}；{note_val}" if desc_val else note_val
        ws.Cells(desc_r, 2).Value = new_desc
        merged += 1
        print(f"  [MERGE] {shape_name}: +'{note_val[:40]}...'")

    print(f"[INFO] Merged {merged} 备注 into 内容描述")

    # Pass 3: delete 备注 rows (from bottom to top to avoid row shift)
    rows_to_delete = sorted(note_rows.values(), reverse=True)
    for r in rows_to_delete:
        ws.Rows(r).Delete()
    print(f"[INFO] Deleted {len(rows_to_delete)} 备注 rows")

    wb.Save()
    wb.Close(False)
    print("[OK] Done.")

except Exception as e:
    print(f"[ERROR] {e}")
finally:
    try:
        excel.Quit()
    except Exception:
        pass
