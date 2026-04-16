import json
from pathlib import Path
import win32com.client

out_path = Path(__file__).resolve().parent / "selected_shape.json"

result = {"ok": False, "error": None, "selection_type": None, "shape_count": 0, "shapes": []}

try:
    app = win32com.client.GetActiveObject("PowerPoint.Application")
    sel = app.ActiveWindow.Selection
    if sel is None:
        result["error"] = "NO_ACTIVE_SELECTION"
    else:
        st = int(sel.Type)
        result["selection_type"] = st
        if st not in (2, 3):
            result["error"] = "SELECTION_NOT_SHAPE"
        else:
            sr = sel.ShapeRange
            cnt = int(sr.Count)
            result["shape_count"] = cnt
            for i in range(1, cnt + 1):
                sh = sr.Item(i)
                item = {
                    "index": i,
                    "name": "",
                    "type": None,
                    "has_text": 0,
                    "has_chart": 0,
                    "text": "",
                }
                try: item["name"] = str(sh.Name)
                except Exception: pass
                try: item["type"] = int(sh.Type)
                except Exception: pass
                try: item["has_text"] = int(sh.HasTextFrame)
                except Exception: pass
                if item["has_text"] == -1:
                    try:
                        if int(sh.TextFrame.HasText) == -1:
                            item["text"] = str(sh.TextFrame.TextRange.Text or "")
                    except Exception:
                        pass
                try: item["has_chart"] = 1 if bool(sh.HasChart) else 0
                except Exception: pass
                result["shapes"].append(item)
            result["ok"] = True
except Exception as e:
    result["error"] = str(e)

out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(str(out_path))
