"""把 active PPT 列表里 '演示文稿1' SaveCopyAs 到 acceptance/ 下，方便 acceptance 用文件路径跑."""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path.home() / ".claude" / "skills" / "office-com-helpers"))
from office_com_helpers import safe_print

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

import win32com.client

ppt = win32com.client.GetActiveObject("PowerPoint.Application")
target_name = "演示文稿1"
out = str(Path(r"D:/Technique Support/Claude Code Learning/3rd-ppt-prj/acceptance/apparel_v4_smoke.pptx").resolve())

for p in ppt.Presentations:
    if p.Name == target_name:
        safe_print(f"Found {p.Name} ({p.Slides.Count} slides), saving copy to {out}")
        try:
            Path(out).unlink(missing_ok=True)
        except Exception:
            pass
        p.SaveCopyAs(out)
        safe_print("SaveCopyAs OK")
        break
else:
    safe_print(f"Not found: {target_name}")
