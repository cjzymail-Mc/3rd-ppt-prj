"""把用户当前 active sheet 复制到新 workbook 并 SaveAs 到指定路径。
   不动用户原工作簿（只在新临时 workbook 上 SaveAs+Close）。

用法:
    python acceptance/export_active_sheet.py <输出路径.xlsx>
"""
import sys
import io
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import win32com.client

if len(sys.argv) < 2:
    print("usage: python acceptance/export_active_sheet.py <out_path.xlsx>")
    raise SystemExit(2)

out_path = os.path.abspath(sys.argv[1])
out_dir = os.path.dirname(out_path)
os.makedirs(out_dir, exist_ok=True)

app = win32com.client.GetActiveObject("Excel.Application")
src_wb = app.ActiveWorkbook
src_sheet = src_wb.ActiveSheet

print(f"Source workbook: {src_wb.Name}")
print(f"Source sheet   : {src_sheet.Name}")
print(f"Target path    : {out_path}")

src_sheet.Copy()  # 无参数 -> 复制到新 workbook（自动成为 active）
new_wb = app.ActiveWorkbook
assert new_wb.Name != src_wb.Name, "Copy 未生成新 workbook"

# 51 = xlOpenXMLWorkbook (.xlsx)
new_wb.SaveAs(out_path, FileFormat=51)
new_wb.Close(SaveChanges=False)  # 关掉这份临时 workbook（已 SaveAs 落盘）

# 切回原 workbook（防止 active 状态变化）
src_wb.Activate()
print(f"OK: {out_path}  ({os.path.getsize(out_path)} bytes)")
print(f"Active 重置为: {app.ActiveWorkbook.Name}")
