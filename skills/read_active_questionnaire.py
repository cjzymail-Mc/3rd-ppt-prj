#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""skills/read_active_questionnaire.py — 读取当前活动 Excel 当前 sheet 的问卷结构.

用法：
    1. 在 Excel 里打开目标问卷并选中目标 sheet
    2. 运行：python skills/read_active_questionnaire.py
    3. JSON 摘要落到 acceptance/active_questionnaire_dump.json（每次覆盖）

桥接现已打开的 Excel（xlwings GetActiveObject），dump：
  - sheet 名 / 总行列数
  - 全部列名（带索引）
  - 前 2 行样例数据（脱敏：长文本截断到 40 字）
  - 数值列识别（>50% 数据 0~10 范围）
  - 各类关键词命中情况：
      * _CATEGORY_KEYWORDS（版型/面料/吸湿排汗/速干）
      * _NAME_KEYWORDS / _WEIGHT_KEYWORDS / 身高
      * 文本反馈列锚点（"补充说明" / "文字描述"）

用途：评估 src/apparel_ppt.py 对新问卷的兼容性；新问卷适配时复跑确认 schema。
"""
from __future__ import annotations

import io
import json
import sys
from pathlib import Path

# Force UTF-8 stdout so terminal printing of Chinese works on GBK consoles.
try:
    sys.stdout.reconfigure(encoding="utf-8")  # py3.7+
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import xlwings as xw

# --- 复用 apparel_ppt.py 的关键字配置（保持评估口径与生产代码一致）-----
_CATEGORY_KEYWORDS = {
    "版型":     ["版型"],
    "面料":     ["面料", "亲肤", "轻量"],
    "吸湿排汗": ["吸湿", "排汗"],
    "速干":     ["速干"],
}
_NAME_KEYWORDS = ["姓名", "name"]
_WEIGHT_KEYWORDS = ["体重", "weight"]
_HEIGHT_KEYWORDS = ["身高", "height"]
_TEXT_MARKERS = ["文字描述", "描述）", "留言", "提交人", "提交时间", "更新时间", "标题"]
_SOURCE_KEYWORDS = ["补充说明"]   # apparel "优点/缺点" gpt_prompted 的 source 槽

def _safe_text(v):
    return "" if v is None else str(v).strip()

def _is_numeric_05(v):
    try:
        if v is None or v == "":
            return False
        f = float(v)
        return 0 <= f <= 10
    except Exception:
        return False

def main():
    book = xw.books.active
    sht = book.sheets.active
    print(f"=== Active workbook: {book.name}")
    print(f"=== Active sheet:    {sht.name}")

    used = sht.used_range
    print(f"=== used_range:      {used.address}  rows={used.rows.count}  cols={used.columns.count}")

    raw = used.value
    if raw is None:
        print("(used_range 为空)")
        return
    # 规范化为 list of list
    if isinstance(raw, tuple):
        rows = [list(r) if isinstance(r, tuple) else [r] for r in raw]
    elif isinstance(raw, list):
        rows = raw if (raw and isinstance(raw[0], list)) else [raw]
    else:
        rows = [[raw]]

    if not rows:
        print("(rows 为空)")
        return

    headers = [_safe_text(h) for h in rows[0]]
    data = rows[1:]
    n_data = len(data)
    print(f"=== headers count:   {len(headers)}")
    print(f"=== data rows:       {n_data}\n")

    # --- 1. 列清单 ---
    print("---- 全部列名 ----")
    for i, h in enumerate(headers):
        print(f"  [{i:02d}] {h}")
    print()

    # --- 2. 样例数据（前 2 行）---
    print("---- 前 2 行样例（长文本截断 40 字）----")
    for ri, row in enumerate(data[:2], 1):
        print(f"  Row {ri}:")
        for ci, h in enumerate(headers):
            v = row[ci] if ci < len(row) else None
            s = _safe_text(v)
            if len(s) > 40:
                s = s[:40] + "…"
            print(f"    [{ci:02d}] {h}: {s}")
    print()

    # --- 3. 数值列识别 ---
    print("---- 数值列识别（>50% 数据落在 [0,10]）----")
    numeric_cols = []
    for ci, h in enumerate(headers):
        hits = sum(1 for r in data if ci < len(r) and _is_numeric_05(r[ci]))
        if n_data and hits / n_data > 0.5:
            sample_vals = [r[ci] for r in data[:3] if ci < len(r)]
            numeric_cols.append((ci, h, hits, sample_vals))
            print(f"  [{ci:02d}] {h}  ({hits}/{n_data})  样例={sample_vals}")
    print(f"  共 {len(numeric_cols)} 个数值列\n")

    # --- 4. _CATEGORY_KEYWORDS 命中情况 ---
    print("---- _CATEGORY_KEYWORDS 命中情况（apparel_ppt._extract_means_for_category）----")
    for cat, kws in _CATEGORY_KEYWORDS.items():
        own = []
        for ci, h in enumerate(headers):
            if any(kw in h for kw in kws):
                excluded = any(m in h for m in _TEXT_MARKERS)
                tag = "  [文本列 已排除]" if excluded else ""
                # 计算数值占比
                hits = sum(1 for r in data if ci < len(r) and _is_numeric_05(r[ci]))
                num_ratio = hits / n_data if n_data else 0
                own.append((ci, h, num_ratio, excluded, tag))
        print(f"  {cat}: {len(own)} 个匹配列")
        for ci, h, ratio, excluded, tag in own:
            mark = "[Y]" if (not excluded and ratio > 0.5) else "[N]"
            print(f"    {mark} [{ci:02d}] {h}  num_ratio={ratio:.0%}{tag}")
    print()

    # --- 5. 人物信息列 ---
    def _find(label, kws):
        for ci, h in enumerate(headers):
            hl = h.lower()
            for kw in kws:
                if kw.lower() in hl:
                    return ci, h
        return None, None
    print("---- 人物信息列 ----")
    for label, kws in [("姓名", _NAME_KEYWORDS),
                       ("身高", _HEIGHT_KEYWORDS),
                       ("体重", _WEIGHT_KEYWORDS)]:
        ci, h = _find(label, kws)
        if ci is None:
            print(f"  [N] {label}: 未找到")
        else:
            sample = [data[r][ci] for r in range(min(3, n_data)) if ci < len(data[r])]
            print(f"  [Y] {label}: [{ci:02d}] {h}  样例={sample}")
    print()

    # --- 6. 文本反馈列（"补充说明"）---
    print("---- 文本反馈列（apparel gpt_prompted source='补充说明'）----")
    found_src = False
    for ci, h in enumerate(headers):
        if any(kw in h for kw in _SOURCE_KEYWORDS):
            found_src = True
            sample = [data[r][ci] for r in range(min(2, n_data)) if ci < len(data[r])]
            sample_s = [(_safe_text(s)[:30] + "…") if len(_safe_text(s)) > 30 else _safe_text(s) for s in sample]
            print(f"  [Y] [{ci:02d}] {h}  样例={sample_s}")
    if not found_src:
        print("  [N] 未找到包含 '补充说明' 的列  ← apparel 优点/缺点 GPT prompt 会拿不到原料！")
    print()

    # --- 7. 输出 JSON 摘要 ---
    summary = {
        "sheet": sht.name,
        "rows": n_data,
        "cols": len(headers),
        "headers": headers,
        "numeric_cols": [{"idx": ci, "name": h} for ci, h, _, _ in numeric_cols],
        "category_hits": {
            cat: [h for ci, h in enumerate(headers) if any(kw in h for kw in kws)]
            for cat, kws in _CATEGORY_KEYWORDS.items()
        },
    }
    out_path = Path(__file__).resolve().parent.parent / "acceptance" / "active_questionnaire_dump.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"=== JSON 摘要已写入: {out_path}")

if __name__ == "__main__":
    main()
