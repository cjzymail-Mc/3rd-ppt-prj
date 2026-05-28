"""列出当前 PowerPoint 进程里所有打开的 Presentation。"""
import win32com.client

try:
    app = win32com.client.GetActiveObject("PowerPoint.Application")
except Exception as e:
    print(f"GetActiveObject FAIL: {e}")
    raise SystemExit(1)

print(f"PowerPoint Presentations.Count = {app.Presentations.Count}")
for i in range(1, app.Presentations.Count + 1):
    p = app.Presentations(i)
    print(f"  [{i}] {p.Name}  | slides={p.Slides.Count}  | dirty={p.Saved == 0}")
