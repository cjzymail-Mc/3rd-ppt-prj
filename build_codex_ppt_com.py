"""
基于 PowerPoint COM 的高保真构建脚本（严禁 python-pptx）。

核心函数：
- open_powerpoint()
- duplicate_slide()
- update_text_shape()
- update_chart_data()
- build_final_ppt()
"""

from __future__ import annotations

import os
from collections import Counter
from typing import Dict, List, Tuple

import win32com.client
import xlwings as xw

TEMPLATE_PPT = "1-标准 ppt 模板.pptx"
DATA_EXCEL = "2025 数据 v2.2.xlsx"
OUTPUT_PPT = "gemini-jules.pptx"

# 这里按“语义名字 -> 默认 shape 名”映射；建议先在模板中手工命名语义名后直接改成语义名
SHAPE_MAPPING = {
    "ProductName": "文本框 16",
    "Score": "矩形 11",
    "Grade": "矩形 12",
    "TestInfo": "矩形 17",
    "MainChart": "图表 44",
}

ATTR_COLS = {
    "抓地性": 4,
    "缓震性": 5,
    "包裹性": 6,
    "抗扭转性": 7,
    "重量&透气性": 8,
    "防侧翻性": 9,
    "耐久性": 10,
}


def open_powerpoint(visible: bool = False):
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = visible
    app.DisplayAlerts = 0
    return app


def duplicate_slide(source_ppt, target_ppt, slide_index: int):
    """从 source 复制 slide 到 target 尾部并返回新 slide。"""
    source_ppt.Slides(slide_index).Copy()
    pasted = target_ppt.Slides.Paste(target_ppt.Slides.Count + 1)
    return pasted


def _find_shape(slide, shape_name: str):
    try:
        return slide.Shapes(shape_name)
    except Exception:
        return None


def update_text_shape(slide, shape_name: str, new_text: str) -> bool:
    shp = _find_shape(slide, shape_name)
    if shp is None:
        return False

    try:
        if shp.HasTextFrame:
            shp.TextFrame.TextRange.Text = str(new_text)
            return True
    except Exception:
        return False

    return False


def update_chart_data(slide, shape_name: str, excel_data_dict: Dict[str, float]) -> bool:
    """通过 ChartData.Workbook 更新图表数据，尽量保留样式。"""
    shp = _find_shape(slide, shape_name)
    if shp is None:
        return False

    try:
        if not shp.HasChart:
            return False

        chart = shp.Chart
        chart.ChartData.Activate()
        wb = chart.ChartData.Workbook
        ws = wb.Worksheets(1)

        categories = list(excel_data_dict.keys())
        values = list(excel_data_dict.values())

        ws.Range("A1").Value = "指标"
        ws.Range("B1").Value = "均值"

        for i, (cat, val) in enumerate(zip(categories, values), start=2):
            ws.Range(f"A{i}").Value = cat
            ws.Range(f"B{i}").Value = float(val)

        last_row = len(categories) + 1
        chart.SetSourceData(ws.Range(f"A1:B{last_row}"))
        wb.Application.Quit()
        return True

    except Exception:
        return False


def _grade_from_score(score: float) -> str:
    if score >= 9:
        return "S"
    if score >= 8:
        return "A"
    if score >= 7:
        return "B"
    return "C"


def _safe_float(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.strip())
        except Exception:
            return None
    return None


def _read_excel_data(excel_path: str) -> Dict[str, object]:
    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        wb = app.books.open(os.path.abspath(excel_path))

        target_sheet = None
        for sht in wb.sheets:
            if "问卷" in sht.name:
                target_sheet = sht
                break
        if target_sheet is None:
            target_sheet = wb.sheets[0]

        data = target_sheet.range("A1").expand().value
        rows = data[1:] if data and len(data) > 1 else []

        valid_rows = [r for r in rows if isinstance(r, list)]
        if not valid_rows:
            return {}

        attr_avg = {}
        for name, col_idx in ATTR_COLS.items():
            values = []
            for row in valid_rows:
                if len(row) > col_idx:
                    num = _safe_float(row[col_idx])
                    if num is not None:
                        values.append(num)
            attr_avg[name] = round(sum(values) / len(values), 2) if values else 0.0

        total = [v for v in attr_avg.values() if v is not None]
        score = round(sum(total) / len(total), 2) if total else 0.0

        weights = [_safe_float(r[2]) for r in valid_rows if len(r) > 2]
        weights = [x for x in weights if x is not None]
        avg_weight = round(sum(weights) / len(weights), 1) if weights else 0.0

        positions = [r[14] for r in valid_rows if len(r) > 14 and r[14]]
        common_pos = Counter(positions).most_common(1)[0][0] if positions else "未知"

        product_name = None
        for r in valid_rows:
            if len(r) > 3 and r[3]:
                product_name = r[3]
                break
        product_name = product_name or "Unknown"

        return {
            "count": len(valid_rows),
            "avg_weight": avg_weight,
            "position": common_pos,
            "product_name": product_name,
            "final_score": score,
            "grade": _grade_from_score(score),
            "attr_avg": attr_avg,
        }
    finally:
        try:
            if wb is not None:
                wb.close()
        except Exception:
            pass
        try:
            if app is not None:
                app.quit()
        except Exception:
            pass


def build_final_ppt(
    template_ppt: str = TEMPLATE_PPT,
    data_excel: str = DATA_EXCEL,
    output_ppt: str = OUTPUT_PPT,
    slide_index: int = 1,
):
    ppt_app = None
    prs_template = None
    prs_out = None

    excel_data = _read_excel_data(data_excel)
    if not excel_data:
        raise RuntimeError("Excel 数据读取失败，无法构建 PPT。")

    try:
        ppt_app = open_powerpoint(visible=False)
        prs_template = ppt_app.Presentations.Open(os.path.abspath(template_ppt), WithWindow=False)
        prs_out = ppt_app.Presentations.Add()

        target_slide = duplicate_slide(prs_template, prs_out, slide_index=slide_index)

        update_text_shape(target_slide, SHAPE_MAPPING["ProductName"], excel_data["product_name"])
        update_text_shape(target_slide, SHAPE_MAPPING["Score"], f"{excel_data['final_score']:.2f}/10")
        update_text_shape(target_slide, SHAPE_MAPPING["Grade"], excel_data["grade"])
        info_text = (
            f"试穿人数：{excel_data['count']}人\n"
            f"测试者平均体重：{excel_data['avg_weight']}KG\n"
            f"测试者球场定位：{excel_data['position']}"
        )
        update_text_shape(target_slide, SHAPE_MAPPING["TestInfo"], info_text)

        update_chart_data(target_slide, SHAPE_MAPPING["MainChart"], excel_data["attr_avg"])

        prs_out.SaveAs(os.path.abspath(output_ppt))
        return os.path.abspath(output_ppt)

    finally:
        try:
            if prs_template is not None:
                prs_template.Close()
        except Exception:
            pass
        try:
            if prs_out is not None:
                prs_out.Close()
        except Exception:
            pass
        try:
            if ppt_app is not None:
                ppt_app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    out = build_final_ppt()
    print(f"构建完成：{out}")
