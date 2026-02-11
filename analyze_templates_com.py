import win32com.client
import json
import os
import sys

# Ensure win32com is available (will fail on Linux, but script is for Windows)
try:
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
except Exception as e:
    print("This script requires Windows and PowerPoint installed.")
    # sys.exit(1) # Commented out to allow import on Linux for syntax check

def analyze_templates(blank_path, standard_path, output_path="shape_detail_com.json"):
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.Visible = True # Analysis usually needs visible for better debugging, but prompt said hidden.
        # But for analysis, Hidden is fine.
        ppt.Visible = False # As requested generally

        abs_blank = os.path.abspath(blank_path)
        abs_std = os.path.abspath(standard_path)

        prs_blank = ppt.Presentations.Open(abs_blank)
        prs_std = ppt.Presentations.Open(abs_std)

        added_shapes = []

        # Assuming slides match 1-to-1 or standard has more
        # We iterate standard slides
        for i in range(1, prs_std.Slides.Count + 1):
            slide_std = prs_std.Slides(i)

            # Get corresponding blank slide if exists
            slide_blank = None
            if i <= prs_blank.Slides.Count:
                slide_blank = prs_blank.Slides(i)

            print(f"Analyzing Slide {i}...")

            # Get shapes from blank slide (for comparison)
            blank_shapes_info = []
            if slide_blank:
                for s in slide_blank.Shapes:
                    blank_shapes_info.append({
                        "left": s.Left,
                        "top": s.Top,
                        "width": s.Width,
                        "height": s.Height,
                        "type": s.Type
                    })

            # Check added shapes in Standard
            for shape in slide_std.Shapes:
                is_new = True
                # Simple check: match by position and size with small tolerance (1 point)
                for b_info in blank_shapes_info:
                    if (abs(shape.Left - b_info["left"]) < 1 and
                        abs(shape.Top - b_info["top"]) < 1 and
                        abs(shape.Width - b_info["width"]) < 1 and
                        abs(shape.Height - b_info["height"]) < 1):
                        is_new = False
                        break

                if is_new:
                    # Detailed extraction
                    details = {
                        "slide_index": i,
                        "name": shape.Name,
                        "left": shape.Left,
                        "top": shape.Top,
                        "width": shape.Width,
                        "height": shape.Height,
                        "type": shape.Type, # 1=msoAutoShape, 17=msoTextBox, etc.
                        "text": "",
                        "font_name": "",
                        "font_size": 0,
                        "font_bold": False,
                        "font_color": "",
                        "chart_type": "",
                        "is_group": False
                    }

                    # Check for Text
                    if shape.HasTextFrame:
                        if shape.TextFrame.HasText:
                            details["text"] = shape.TextFrame.TextRange.Text
                            # Font info from first run
                            font = shape.TextFrame.TextRange.Runs(1).Font
                            details["font_name"] = font.Name
                            details["font_size"] = font.Size
                            details["font_bold"] = font.Bold
                            details["font_color"] = font.Color.RGB

                    # Check for Chart
                    # msoChart = 3
                    if shape.Type == 3:
                        try:
                            details["chart_type"] = shape.Chart.ChartType
                        except:
                            pass

                    # Check for Group
                    # msoGroup = 6
                    if shape.Type == 6:
                        details["is_group"] = True

                    added_shapes.append(details)
                    print(f"  Found added shape: {shape.Name}")

        prs_blank.Close()
        prs_std.Close()
        ppt.Quit()

        with open(output_path, "w", encoding='utf-8') as f:
            json.dump(added_shapes, f, indent=4, ensure_ascii=False)

        print(f"Analysis complete. Found {len(added_shapes)} added shapes. Saved to {output_path}")

    except Exception as e:
        print(f"Error during analysis: {e}")
        # Ensure cleanup
        try:
            prs_blank.Close()
            prs_std.Close()
            ppt.Quit()
        except:
            pass

if __name__ == "__main__":
    analyze_templates("0-空白 ppt 模板.pptx", "1-标准 ppt 模板.pptx")
