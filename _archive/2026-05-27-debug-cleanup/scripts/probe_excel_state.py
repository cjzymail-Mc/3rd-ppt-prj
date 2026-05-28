"""列出当前 Excel 进程里所有打开的 Workbook + 当前 active 的 sheet 清单。
   UTF-8 safe print（避免中文乱码）。"""
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import win32com.client

try:
    app = win32com.client.GetActiveObject("Excel.Application")
except Exception as e:
    print(f"GetActiveObject FAIL: {e}  --> Excel 进程不存在")
    raise SystemExit(0)

print(f"Excel.Application.Workbooks.Count = {app.Workbooks.Count}")
for i in range(1, app.Workbooks.Count + 1):
    wb = app.Workbooks(i)
    try:
        full_name = wb.FullName
    except Exception:
        full_name = "<unknown>"
    print(f"  [{i}] Name={wb.Name}")
    print(f"      FullName={full_name}")
    print(f"      Saved={wb.Saved}")
    try:
        print(f"      Sheets ({wb.Sheets.Count}):")
        for j in range(1, wb.Sheets.Count + 1):
            sh = wb.Sheets(j)
            ur = sh.UsedRange
            rows = ur.Rows.Count if ur else 0
            cols = ur.Columns.Count if ur else 0
            print(f"        - [{j}] {sh.Name}  (UsedRange {rows}×{cols})")
    except Exception as e:
        print(f"      sheets probe FAIL: {e}")

try:
    active = app.ActiveWorkbook
    if active is not None:
        print(f"\nActive workbook = {active.Name}")
        try:
            print(f"Active sheet    = {active.ActiveSheet.Name}")
        except Exception:
            pass
except Exception:
    pass
