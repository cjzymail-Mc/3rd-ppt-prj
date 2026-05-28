"""读 Excel 全部 9 个 sample 的列 AD（适合温度）+ 列 AE（实际温度），
打印 Counter，看 mode 选哪个。"""
import sys
from collections import Counter
from pathlib import Path
import xlwings

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def main():
    book = xlwings.books.active
    sht = None
    for s in book.sheets:
        if "问卷" in s.name or "紧身背心" in s.name:
            sht = s
            break
    if sht is None:
        print("FAIL: 没找到问卷 sheet")
        return 1
    print(f"sheet: {sht.name}")
    used = sht.used_range
    print(f"used_range: {used.address}  rows={used.last_cell.row}  cols={used.last_cell.column}")

    rows = used.value
    headers = [str(h) if h else "" for h in rows[0]]
    print("\n列 AD/AE 数据：")
    print(f"  AD header: {headers[29] if len(headers) > 29 else '(out of range)'}")  # AD = 30, idx 29
    print(f"  AE header: {headers[30] if len(headers) > 30 else '(out of range)'}")

    print(f"\nrow count (含表头): {len(rows)}")
    bins_ad = []
    bins_ae = []
    for i, row in enumerate(rows[1:], 2):
        ad_val = row[29] if len(row) > 29 else None
        ae_val = row[30] if len(row) > 30 else None
        print(f"  row {i}: AD={ad_val!r}  AE={ae_val!r}")
        if ad_val:
            bins_ad.append(str(ad_val).strip())
        if ae_val:
            bins_ae.append(str(ae_val).strip())

    print(f"\nAD Counter: {Counter(bins_ad).most_common()}")
    print(f"AE Counter: {Counter(bins_ae).most_common()}")


if __name__ == "__main__":
    main()
