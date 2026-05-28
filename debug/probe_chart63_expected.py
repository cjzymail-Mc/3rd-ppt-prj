"""Compute expected Chart 63 values from acceptance/data-apparel.xlsx.

Run this whenever data-apparel.xlsx changes; copy the printed values into
acceptance/apparel.json L1 rule p13_chart63_temp_range.expected.
"""
import sys
from pathlib import Path

# Ensure src/ on path so apparel_ppt._calc_chart63_data is importable
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

helpers = Path.home() / ".claude" / "skills" / "office-com-helpers"
sys.path.insert(0, str(helpers))
from office_com_helpers import load_excel_rows

from apparel_ppt import _calc_chart63_data

XLSX = ROOT / "acceptance" / "data-apparel.xlsx"
rows, _, _ = load_excel_rows(
    excel_path=str(XLSX),
    sheet_name="服装试穿问卷",
    fuzzy_keyword="服装试穿问卷",
)
print(f"Loaded {len(rows)} rows (incl header) from {XLSX.name}")
data = _calc_chart63_data(rows)
print()
print("Computed chart63 data:")
for k, v in data.items():
    print(f"  {k}: {v}")
print()
print("L1 chart_series_values expected (3 series × 2 cats):")
expected = [data["s1_values"], data["s2_values"], data["s3_values"]]
print(f"  {expected}")
