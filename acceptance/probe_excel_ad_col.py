"""Probe column AD (温度区间) raw strings from acceptance/data-apparel.xlsx.

用 DispatchEx 起隔离 Excel 进程，不动用户当前打开的 active workbook。
"""
from __future__ import annotations
import sys
from pathlib import Path

HELPERS = Path.home() / ".claude" / "skills" / "office-com-helpers"
sys.path.insert(0, str(HELPERS))
from office_com_helpers import safe_print

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import win32com.client

XLSX = str(Path(r"D:/Technique Support/Claude Code Learning/3rd-ppt-prj/acceptance/data-apparel.xlsx").resolve())

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(XLSX, ReadOnly=True)
    try:
        for sh in wb.Sheets:
            safe_print(f"\n=== Sheet: {sh.Name} ===")
            used = sh.UsedRange
            rows = used.Rows.Count
            cols = used.Columns.Count
            safe_print(f"  UsedRange rows={rows} cols={cols}")
            # 列 AD = 第 30 列
            ad_col = 30
            safe_print(f"  Column AD (idx={ad_col}) values:")
            for r in range(1, min(rows + 1, 15)):
                cell = sh.Cells(r, ad_col)
                v = cell.Value
                safe_print(f"    row{r}: {v!r}")
            # 顺便看下 row1 列 28-32 是哪些表头
            safe_print(f"  Headers row1 col 28-34:")
            for c in range(28, 35):
                v = sh.Cells(1, c).Value
                safe_print(f"    col{c}: {v!r}")
    finally:
        wb.Close(SaveChanges=False)
finally:
    excel.Quit()
