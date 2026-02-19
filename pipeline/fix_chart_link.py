#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""One-time utility: break broken external chart link in Template 2.1.pptx.

The chart on slide 15 has an external Excel link that no longer exists,
causing a popup dialog every time the file is opened. This script:
  1. Opens the template via COM (DisplayAlerts=0 suppresses the popup)
  2. Finds every chart shape across all slides
  3. Calls ChartData.BreakLink() to sever the external Excel connection
  4. Saves and closes the template

Run once, then the popup will be gone permanently.

Usage:
    python pipeline/fix_chart_link.py
"""

import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "src" / "Template 2.1.pptx"


def _p(msg: str):
    """Print with ASCII fallback for non-UTF8 consoles."""
    try:
        print(msg)
    except UnicodeEncodeError:
        print(msg.encode("ascii", errors="replace").decode("ascii"))


def fix_chart_link():
    try:
        import win32com.client
    except ImportError:
        print("[ERROR] pywin32 not installed. Run: pip install pywin32")
        return 1

    if not TEMPLATE.exists():
        _p(f"[ERROR] Template not found: {TEMPLATE}")
        return 1

    _p(f"[INFO] Opening: {TEMPLATE}")
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = 0   # suppress "linked file not available" popup
    app.Visible = True

    try:
        prs = app.Presentations.Open(str(TEMPLATE), ReadOnly=False)
    except Exception as e:
        _p(f"[ERROR] Cannot open presentation: {e}")
        app.Quit()
        return 1

    # Give PowerPoint a moment to settle
    time.sleep(1)

    fixed = 0
    errors = []

    for slide_idx in range(1, prs.Slides.Count + 1):
        slide = prs.Slides(slide_idx)
        for shape_idx in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(shape_idx)
            if not shp.HasChart:
                continue

            shape_name = shp.Name
            _p(f"[INFO] Found chart: slide {slide_idx}, shape '{shape_name}'")

            chart_data = shp.Chart.ChartData

            # Strategy 1: ChartData.BreakLink() — works without Activate(),
            # severs the external Excel connection directly.
            try:
                chart_data.BreakLink()
                _p(f"[OK]   BreakLink() succeeded on slide {slide_idx}")
                fixed += 1
                continue
            except Exception as e1:
                _p(f"[INFO] BreakLink() returned: {e1}")
                _p("[INFO] Trying Activate() + workbook approach...")

            # Strategy 2: Activate embedded workbook, then break via wb.BreakLink()
            # This opens an Excel-like editor; we suppress its alerts too.
            try:
                # Suppress Excel dialogs if Excel opens
                try:
                    xl = win32com.client.GetActiveObject("Excel.Application")
                    xl.DisplayAlerts = False
                except Exception:
                    pass

                chart_data.Activate()
                time.sleep(0.5)

                wb = chart_data.Workbook
                try:
                    links = wb.LinkSources(1)   # 1 = xlExcelLinks
                except Exception:
                    links = None

                if links:
                    _p(f"[INFO]   Found {len(links)} link(s), breaking...")
                    for link in links:
                        try:
                            wb.BreakLink(link, 1)
                            _p(f"[OK]    Broke link: {link}")
                            fixed += 1
                        except Exception as e2:
                            msg = f"wb.BreakLink failed: {e2}"
                            _p(f"[WARN]  {msg}")
                            errors.append(msg)
                    try:
                        wb.Save()
                    except Exception:
                        pass
                else:
                    _p("[INFO]   No xlExcelLinks found in workbook.")

            except Exception as e3:
                msg = f"slide {slide_idx} '{shape_name}': {e3}"
                _p(f"[WARN] Both strategies failed - {msg}")
                errors.append(msg)

    _p("\n[INFO] Saving template...")
    try:
        prs.Save()
        _p("[OK]   Template saved.")
    except Exception as e:
        _p(f"[ERROR] Save failed: {e}")
        errors.append(str(e))

    prs.Close()
    app.Quit()

    if fixed > 0:
        _p(f"\n[DONE] Fixed {fixed} chart(s). Popup should be gone now.")
        _p("       Open the template to verify no popup appears.")
    elif not errors:
        _p("\n[DONE] No external links found via either strategy.")
        _p("       If popup persists, the link may be embedded as OLE — check manually.")
    else:
        _p(f"\n[PARTIAL] Fixed={fixed}, errors={len(errors)}:")
        for err in errors:
            _p(f"  - {err}")

    return 0 if not errors else 1


if __name__ == "__main__":
    raise SystemExit(fix_chart_link())
