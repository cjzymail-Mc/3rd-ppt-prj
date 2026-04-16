import win32com.client

try:
    app = win32com.client.GetActiveObject("PowerPoint.Application")
except Exception as e:
    print(f"PPT_NOT_ACTIVE: {e}")
    raise SystemExit(0)

try:
    vis = app.Visible
except Exception:
    vis = 'ERR'
print(f"APP_VISIBLE={vis}")

try:
    cnt = int(app.Presentations.Count)
    print(f"PRESENTATIONS_COUNT={cnt}")
    for i in range(1, cnt + 1):
        p = app.Presentations.Item(i)
        name = ''
        path = ''
        try: name = str(p.Name)
        except Exception: pass
        try: path = str(p.FullName)
        except Exception: pass
        print(f"PRES_{i}={name}|{path}")
except Exception as e:
    print(f"PRESENTATIONS_ERR={e}")

try:
    aw = app.ActiveWindow
    print("ACTIVE_WINDOW=OK")
except Exception as e:
    print(f"ACTIVE_WINDOW_ERR={e}")
    raise SystemExit(0)

try:
    vp = aw.ViewType
except Exception:
    vp = 'ERR'
print(f"VIEW_TYPE={vp}")

try:
    sel = aw.Selection
except Exception as e:
    print(f"SELECTION_ERR={e}")
    raise SystemExit(0)

if sel is None:
    print("SELECTION_NONE_OBJECT")
    raise SystemExit(0)

try:
    st = int(sel.Type)
except Exception as e:
    print(f"SELECTION_TYPE_ERR={e}")
    raise SystemExit(0)

print(f"SELECTION_TYPE={st}")

try:
    sld = aw.View.Slide
    print(f"CURRENT_SLIDE_INDEX={int(sld.SlideIndex)}")
except Exception as e:
    print(f"CURRENT_SLIDE_ERR={e}")
