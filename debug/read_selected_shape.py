import win32com.client

try:
    app = win32com.client.GetActiveObject("PowerPoint.Application")
except Exception as e:
    print(f"PPT_NOT_ACTIVE: {e}")
    raise SystemExit(0)

try:
    sel = app.ActiveWindow.Selection
except Exception:
    print("NO_ACTIVE_SELECTION")
    raise SystemExit(0)

if sel is None:
    print("NO_ACTIVE_SELECTION")
    raise SystemExit(0)

try:
    st = int(sel.Type)
except Exception:
    print("SELECTION_TYPE_UNKNOWN")
    raise SystemExit(0)

print(f"SELECTION_TYPE={st}")
if st not in (2, 3):
    print("SELECTION_NOT_SHAPE")
    raise SystemExit(0)

try:
    sr = sel.ShapeRange
    cnt = int(sr.Count)
except Exception:
    print("NO_SHAPERANGE")
    raise SystemExit(0)

print(f"SHAPE_COUNT={cnt}")
for i in range(1, cnt + 1):
    sh = sr.Item(i)
    name = ""
    typ = ""
    has_text = 0
    has_chart = 0
    text = ""

    try:
        name = str(sh.Name)
    except Exception:
        pass
    try:
        typ = int(sh.Type)
    except Exception:
        pass
    try:
        has_text = int(sh.HasTextFrame)
    except Exception:
        pass
    if has_text == -1:
        try:
            if int(sh.TextFrame.HasText) == -1:
                text = str(sh.TextFrame.TextRange.Text or "")
        except Exception:
            pass
    try:
        has_chart = 1 if bool(sh.HasChart) else 0
    except Exception:
        pass

    text = text.replace("\r", "[CR]").replace("\n", "[LF]")
    if len(text) > 400:
        text = text[:400] + "..."

    print(f"--- SHAPE {i} ---")
    print(f"NAME={name}")
    print(f"TYPE={typ}")
    print(f"HAS_TEXT={has_text}")
    print(f"HAS_CHART={has_chart}")
    print(f"TEXT={text}")
