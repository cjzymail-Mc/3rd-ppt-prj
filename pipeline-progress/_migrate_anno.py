#!/usr/bin/env python3
"""One-shot: migrate annotations from backup MD into 01-shape_detail.xlsx via COM."""
import win32com.client
import os

xlsx_path = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "01-shape_detail.xlsx")
)

# Annotations extracted from "01-shape_detail (back-up).md"
# Names mapped to actual COM names in the xlsx
annos = {
    "Rectangle 11": {
        "\u5185\u5bb9\u63cf\u8ff0": "\u6240\u6709\u7528\u6237\u8bc4\u5206\u7684\u5747\u503c",
        "strategy": "score_10pt",
        "params": "scale=auto, format=X.XX/10",
        "\u5907\u6ce8": "",
    },
    "Rectangle 12": {
        "\u5185\u5bb9\u63cf\u8ff0": "\u6240\u6709\u7528\u6237\u8bc4\u5206\u7684\u5747\u503c",
        "strategy": "grade_letter",
        "params": "scale=auto",
        "\u5907\u6ce8": "",
    },
    "Rectangle 17": {
        "\u5185\u5bb9\u63cf\u8ff0": "Excel \u95ee\u5377sheet\uff0c\u63d0\u53d6\u8bd5\u7a7f\u4eba\u6570\u3001\u5e73\u5747\u4f53\u91cd\u3001\u7403\u573a\u5b9a\u4f4d\u5217",
        "strategy": "sample_aggregation",
        "params": "fields=\u8bd5\u7a7f\u4eba\u6570|\u5e73\u5747\u4f53\u91cd|\u7403\u573a\u5b9a\u4f4d",
        "\u5907\u6ce8": "\u683c\u5f0f\u4fdd\u6301\u548c\u6a21\u677f\u4e00\u81f4\uff0c\u6bcf\u9879\u72ec\u5360\u4e00\u884c",
    },
    "Rectangle 19": {
        "\u5185\u5bb9\u63cf\u8ff0": "\u88c5\u9970\u6027\u7ec6\u6761\uff0c\u65e0\u5185\u5bb9",
        "strategy": "skip",
        "params": "",
        "\u5907\u6ce8": "",
    },
    "Picture 39": {
        "\u5185\u5bb9\u63cf\u8ff0": "Excel \u95ee\u5377sheet\uff0c\u7b2c\u4e00\u5f20\u5d4c\u5165\u56fe\u7247(\u978b\u6b3e\u7167\u7247)",
        "strategy": "extract_image",
        "params": "sheet=\u95ee\u5377",
        "\u5907\u6ce8": "\u4fdd\u6301\u539f\u59cb\u5c3a\u5bf8\u548c\u4f4d\u7f6e\u4e0d\u53d8",
    },
    "TextBox 16": {
        "\u5185\u5bb9\u63cf\u8ff0": "\u978b\u6b3e\u540d\u79f0",
        "strategy": "extract_column",
        "params": "column=\u978b\u6b3e\u540d\u79f0",
        "\u5907\u6ce8": "",
    },
    "Rectangle 68": {
        "\u5185\u5bb9\u63cf\u8ff0": "\u95ee\u5377\u8865\u5145\u8bf4\u660e\uff0c\u5f52\u7eb3\u4ea7\u54c1\u7f3a\u70b9",
        "strategy": "gpt_prompted",
        "params": "source=\u8865\u5145\u8bf4\u660e, filter=\u7f3a\u70b9",
        "\u5907\u6ce8": "\u63a7\u5236\u5728280\u5b57\u5de6\u53f3\uff1bGPT\u81ea\u884c\u51b3\u5b9a\u5206\u6bb5\u7ef4\u5ea6\uff0c\u4e0d\u8981\u6309\u56fa\u5b9a\u6027\u80fd\u7c7b\u522b\u5206\u7c7b\uff1b\u4fdd\u7559(X/N)\u6bd4\u4f8b\u6570\u636e",
    },
    "Rectangle 77": {
        "\u5185\u5bb9\u63cf\u8ff0": "\u95ee\u5377\u8865\u5145\u8bf4\u660e\uff0c\u5f52\u7eb3\u4ea7\u54c1\u4f18\u70b9",
        "strategy": "gpt_prompted",
        "params": "source=\u8865\u5145\u8bf4\u660e, filter=\u4f18\u70b9",
        "\u5907\u6ce8": "\u63a7\u5236\u5728220\u5b57\u5de6\u53f3\uff1bGPT\u81ea\u884c\u51b3\u5b9a\u5206\u6bb5\u7ef4\u5ea6\uff0c\u4e0d\u8981\u6309\u56fa\u5b9a\u6027\u80fd\u7c7b\u522b\u5206\u7c7b\uff1b\u4fdd\u7559(X/N)\u6bd4\u4f8b\u6570\u636e",
    },
    "\u56fe\u8868 44": {
        "\u5185\u5bb9\u63cf\u8ff0": "Excel \u95ee\u5377sheet \u5404\u8bc4\u5206\u5217\u5747\u503c",
        "strategy": "mean_extraction",
        "params": "",
        "\u5907\u6ce8": "",
    },
}

print(f"[INFO] Opening: {xlsx_path}")
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

try:
    wb = excel.Workbooks.Open(xlsx_path)
    ws = wb.Sheets(1)
    max_row = ws.UsedRange.Rows.Count
    print(f"[INFO] Sheet rows: {max_row}")

    current_shape = None
    in_anno = False
    written = 0

    for r in range(1, max_row + 1):
        a_val = str(ws.Cells(r, 1).Value or "").strip()

        if a_val.startswith("Shape #"):
            current_shape = str(ws.Cells(r, 2).Value or "").strip()
            in_anno = False
            continue

        if a_val == "\u7528\u6237\u6279\u6ce8":  # 用户批注
            in_anno = True
            continue

        if in_anno and current_shape and current_shape in annos:
            shape_anno = annos[current_shape]
            if a_val in shape_anno:
                val = shape_anno[a_val]
                if val:
                    ws.Cells(r, 2).Value = val
                    written += 1
                    # safe print
                    try:
                        print(f"  [WRITE] {current_shape} / {a_val} = {val[:60]}")
                    except UnicodeEncodeError:
                        print(f"  [WRITE] {current_shape} / (field) = (value written)")

    wb.Save()
    print(f"\n[OK] Written {written} annotation cells.")
finally:
    wb.Close(False)
    excel.Quit()
    print("[OK] Excel COM released.")
