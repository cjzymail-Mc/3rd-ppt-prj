#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fix3 透气性遗漏修复后的双 schema 烟测：新旧 apparel 问卷都能正确分组。"""
from __future__ import annotations
import io, sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.apparel_ppt import _section_boundary, _extract_means_for_category

# === 新问卷 headers（来自 debug/active_questionnaire_dump.json）===
NEW_HEADERS = [
    "data_id", "标题", "1、跑者姓名", "2、身高（cm）", "3、体重（kg）",
    "4、三围（胸、腰、臀）（cm）", "6、测试累计总跑量（km）",
    "1、【整体】版型评分（得分越高评价越合身，下同）",
    "2、【领口】版型评分", "2、【袖口】版型评分", "3、【肩部】版型评分",
    "4、【胸围】版型评分", "5、【腰围】版型评分",
    "版型评价（文字描述）",
    "1、面料手感/亲肤度（得分越高评价越好，下同）",
    "2、面料轻量化评分", "3、透气性评分",
    "面料评价（文字描述）",
    "1、【整体】吸湿排汗性能评分（得分越高评价越好，下同）",
    "2、【腋下】吸湿排汗性能评分", "3、【后腰】吸湿排汗性能评分",
    "4、【前胸】吸湿排汗性能评分",
    "吸湿排汗性能评价（文字描述）",
    "5、【整体】速干性能评分（得分越高评价越好，下同）",
    "6、【腋下】速干性能评分", "7、【后腰】速干性能评分",
    "8、【前胸】速干性能评分",
    "速干性能评价（文字描述）",
    "适宜的穿着场景", "适合的温度区间（体感温度）", "实际穿着温度区间（℃）",
    "1、该样品可以达到的支撑效果", "其他补充内容",
    "提交人", "提交时间", "更新时间",
]

# === 旧问卷 headers（来自 debug/read_old_apparel_questionnaire.py 实跑）===
OLD_HEADERS = [
    "data_id", "标题", "1、跑者姓名", "2、身高（cm）", "3、体重（kg）",
    "5、试穿产品品牌",
    "1、【整体】版型评分（得分越高评价越合身，下同）",
    "2、【衣领】版型评分", "3、【袖口】版型评分", "4、【胸围】版型评分",
    "5、【腰围】版型评分",
    "版型评价（文字描述）",
    "1、面料手感/亲肤度（得分越高评价越好，下同）",
    "2、面料透气性", "3、轻量化",
    "面料评价（文字描述）",
    "1、【整体】吸湿排汗性能评分（得分越高评价越好，下同）",
    "2、【前胸】吸湿排汗性能评分", "3、【后背】吸湿排汗性能评分",
    "吸湿排汗性能评价（文字描述）",
    "4、【整体】速干性能评分（得分越高评价越好，下同）",
    "5、【后背】速干性能评分", "6、【前胸】速干性能评分",
    "速干性能评价（文字描述）",
    "如果您对此件跑步背心还有其他试穿感受，请留言：",
    "提交人", "提交时间", "更新时间",
]

EXPECTED = {
    "new": {"版型": 6, "面料": 3, "吸湿排汗": 4, "速干": 4},
    "old": {"版型": 5, "面料": 3, "吸湿排汗": 3, "速干": 3},
}


def mk_rows(headers, n_samples=9):
    """生成 n_samples 行模拟数据：每个评分列填 4（合理评分），其他填字符串。"""
    rows = [headers[:]]
    for i in range(n_samples):
        row = []
        for h in headers:
            if "评分" in h or "亲肤度" in h or "透气性" in h or "轻量化" in h or "面料透气" in h:
                row.append(4)  # 合理评分
            elif "身高" in h:
                row.append("170")
            elif "体重" in h:
                row.append("60")
            elif "姓名" in h:
                row.append(f"跑者{chr(65+i)}")
            else:
                row.append(f"value{i}")
        rows.append(row)
    return rows


def test_schema(name, headers, expected):
    print(f"\n========== {name} schema ==========")
    rows = mk_rows(headers, n_samples=9)
    all_ok = True
    for cat, exp_count in expected.items():
        # _section_boundary 检查
        bounds = _section_boundary(headers, cat)
        # _extract_means_for_category 实际结果
        means = _extract_means_for_category(rows, cat)
        actual = len(means)
        ok = actual == exp_count
        if not ok:
            all_ok = False
        flag = "✓" if ok else "✗ MISMATCH!"
        print(f"  [{cat}] bounds={bounds} → 实际 {actual} 个，期望 {exp_count}  {flag}")
        if means:
            labels = [m[0] for m in means]
            print(f"          labels: {labels}")
    return all_ok


ok_new = test_schema("新问卷 (active_questionnaire_dump)", NEW_HEADERS, EXPECTED["new"])
ok_old = test_schema("旧问卷 (source data-apparel)", OLD_HEADERS, EXPECTED["old"])

print(f"\n=== overall: {'PASS' if ok_new and ok_old else 'FAIL'}")
