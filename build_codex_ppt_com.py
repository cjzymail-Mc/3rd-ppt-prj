import win32com.client
import xlwings as xw
import os
import sys
import json
from collections import Counter

# Configuration
TEMPLATE_PPT = "1-标准 ppt 模板.pptx"
DATA_EXCEL = "2-问卷模板.xlsx"
OUTPUT_PPT = "gemini-jules.pptx"

# Shape Mapping (Inferred from Analysis)
# The user might rename shapes, but we provide defaults based on our analysis.
SHAPE_MAPPING = {
    "Score": "矩形 11",          # "8.29/10"
    "Grade": "矩形 12",          # "A"
    "Hyphen": "矩形 13",         # "-"
    "TestInfo": "矩形 17",       # "试穿人数..."
    "ProductImage": "图片 39",   # Image
    "ProductName": "文本框 16",  # "SUPER GRASS"
    "MainChart": "图表 44",      # Radar/Bar Chart
    "Feedback1": "矩形 68",      # "【包裹性】..."
    "Feedback2": "矩形 77"       # "【止滑性】..."
}

# Attribute Columns in Excel (Indices 0-based, assuming header is row 0)
# Columns: '抓地性' (4), '缓震性' (5), '包裹性' (6), '抗扭转性' (7), '重量&透气性' (8), '防侧翻性' (9), '耐久性' (10)
ATTR_COLS = {
    "抓地性": 4,
    "缓震性": 5,
    "包裹性": 6,
    "抗扭转性": 7,
    "重量&透气性": 8,
    "防侧翻性": 9,
    "耐久性": 10
}

def get_excel_data():
    """Reads data using xlwings and calculates stats."""
    try:
        app = xw.App(visible=False)
        wb = app.books.open(os.path.abspath(DATA_EXCEL))
        sheet = wb.sheets[0]

        # Read all data
        data = sheet.range("A1").expand().value
        header = data[0]
        rows = data[1:]

        wb.close()
        app.quit()

        if not rows:
            return None

        # Stats
        count = len(rows)

        # Weight (Col 2)
        weights = [r[2] for r in rows if r[2] is not None]
        avg_weight = sum(weights) / len(weights) if weights else 0

        # Position (Col 14: "你认为这双鞋更加适合什么打法的球员穿")
        positions = [r[14] for r in rows if r[14] is not None]
        common_pos = Counter(positions).most_common(1)[0][0] if positions else "未知"

        # Product Name (Col 3)
        product_name = rows[0][3] if rows[0][3] else "Unknown"

        # Attributes Scores
        attr_sums = {k: 0 for k in ATTR_COLS}
        attr_counts = {k: 0 for k in ATTR_COLS}

        for row in rows:
            for attr, col_idx in ATTR_COLS.items():
                val = row[col_idx]
                if val is not None and isinstance(val, (int, float)):
                    attr_sums[attr] += val
                    attr_counts[attr] += 1

        attr_avgs = {}
        total_score_sum = 0
        total_score_count = 0

        for attr in ATTR_COLS:
            if attr_counts[attr] > 0:
                avg = attr_sums[attr] / attr_counts[attr]
                attr_avgs[attr] = avg
                total_score_sum += avg
                total_score_count += 1
            else:
                attr_avgs[attr] = 0

        final_score = total_score_sum / total_score_count if total_score_count > 0 else 0

        # Grade Logic
        if final_score >= 9: grade = "S"
        elif final_score >= 8: grade = "A"
        elif final_score >= 7: grade = "B"
        else: grade = "C"

        # Feedbacks (Col 15: "补充说明")
        feedbacks = [r[15] for r in rows if r[15] is not None]

        return {
            "count": count,
            "avg_weight": avg_weight,
            "position": common_pos,
            "product_name": product_name,
            "final_score": final_score,
            "grade": grade,
            "attr_avgs": attr_avgs,
            "feedbacks": feedbacks
        }

    except Exception as e:
        print(f"Error reading Excel: {e}")
        # Kill excel if needed?
        try: app.quit()
        except: pass
        return None

def build_ppt(data):
    """Updates PPT using win32com."""
    ppt_app = None
    prs = None
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True # Usually needed for some operations, but requested Hidden
        # Prompt said "Visible = False".
        # However, selecting shapes sometimes fails if not visible.
        # But we access by Name/Index, so should be fine.
        ppt_app.Visible = False

        abs_template = os.path.abspath(TEMPLATE_PPT)
        abs_output = os.path.abspath(OUTPUT_PPT)

        # Open Template
        prs = ppt_app.Presentations.Open(abs_template)

        # Save As New
        prs.SaveAs(abs_output)

        # We work on the first slide (assuming single slide template or first slide is target)
        slide = prs.Slides(1)

        # Helper to find shape
        def get_shape(name):
            try:
                return slide.Shapes(name)
            except:
                print(f"Warning: Shape '{name}' not found by name.")
                return None

        # 1. Update Product Name
        s_name = get_shape(SHAPE_MAPPING["ProductName"])
        if s_name and s_name.HasTextFrame:
            s_name.TextFrame.TextRange.Text = data["product_name"]

        # 2. Update Score
        s_score = get_shape(SHAPE_MAPPING["Score"])
        if s_score and s_score.HasTextFrame:
            s_score.TextFrame.TextRange.Text = f"{data['final_score']:.2f}/10"

        # 3. Update Grade
        s_grade = get_shape(SHAPE_MAPPING["Grade"])
        if s_grade and s_grade.HasTextFrame:
            s_grade.TextFrame.TextRange.Text = data["grade"]

        # 4. Update Test Info
        s_info = get_shape(SHAPE_MAPPING["TestInfo"])
        if s_info and s_info.HasTextFrame:
            info_text = (f"试穿人数：{data['count']}人\n"
                         f"测试者平均体重：{data['avg_weight']:.1f}KG\n"
                         f"测试者球场定位：{data['position']}")
            s_info.TextFrame.TextRange.Text = info_text

        # 5. Update Feedbacks (Split into 2 boxes simply for now)
        # Template has "Feedback1" and "Feedback2"
        # We just join all feedbacks and split?
        # Or categorization logic?
        # For simplicity: Join all unique feedbacks.
        all_fb = "\n".join(set(data["feedbacks"]))
        # Split roughly
        mid = len(all_fb) // 2
        fb1 = all_fb[:mid]
        fb2 = all_fb[mid:]

        # Actually template has sections like 【包裹性】.
        # If we can't parse sections easily, just dump text.
        # Ideally, we should parse the feedback text to find keywords.
        # But given constraints, we just put raw text.
        s_fb1 = get_shape(SHAPE_MAPPING["Feedback1"])
        if s_fb1 and s_fb1.HasTextFrame:
            # Keep the header if exists? No, replace content.
            # Maybe keep "【用户反馈】" prefix if we want.
            s_fb1.TextFrame.TextRange.Text = "【综合反馈-1】\n" + fb1

        s_fb2 = get_shape(SHAPE_MAPPING["Feedback2"])
        if s_fb2 and s_fb2.HasTextFrame:
            s_fb2.TextFrame.TextRange.Text = "【综合反馈-2】\n" + fb2

        # 6. Update Chart
        s_chart = get_shape(SHAPE_MAPPING["MainChart"])
        if s_chart and s_chart.HasChart:
            chart = s_chart.Chart
            # Update Data
            # Attributes order in chart?
            # We assume the order matches our ATTR_COLS keys or we need to read Categories.
            # Reading categories from chart is tricky if not linked to Excel range.
            # But we can try to set them.

            categories = list(data["attr_avgs"].keys())
            values = list(data["attr_avgs"].values())

            try:
                # Attempt to set values directly
                # Series 1
                series = chart.SeriesCollection(1)
                series.Values = values
                series.XValues = categories
            except Exception as e:
                print(f"Error updating chart data: {e}")
                # Fallback: Open chart data (heavy)
                try:
                    wb_chart = chart.ChartData.Workbook
                    ws_chart = wb_chart.Worksheets(1)
                    # Write data to A2:B8
                    for i, (cat, val) in enumerate(zip(categories, values)):
                        ws_chart.Range(f"A{i+2}").Value = cat
                        ws_chart.Range(f"B{i+2}").Value = val
                    wb_chart.Close()
                except Exception as e2:
                    print(f"Error updating chart data via Workbook: {e2}")

        # Save
        prs.Save()
        prs.Close()
        ppt_app.Quit()
        print(f"Successfully generated {OUTPUT_PPT}")

    except Exception as e:
        print(f"Error building PPT: {e}")
        if prs: prs.Close()
        if ppt_app: ppt_app.Quit()

if __name__ == "__main__":
    if not os.name == 'nt':
        print("This script must be run on Windows.")
        # sys.exit(1) # Allow syntax check

    print("Reading data...")
    data = get_excel_data()
    if data:
        print("Data loaded. Building PPT...")
        build_ppt(data)
    else:
        print("Failed to load data.")
