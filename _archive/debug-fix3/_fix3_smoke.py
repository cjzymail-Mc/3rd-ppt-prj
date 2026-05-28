#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fix3 smoke test：验证 _build_respondent_block 已剔除新问卷噪声列。"""
import io
import sys
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.apparel_ppt import _build_respondent_block, _build_content, APPAREL_SHAPES

hdr = ['data_id', '标题', '1、跑者姓名', '2、身高（cm）', '3、体重（kg）',
       '4、三围（胸、腰、臀）（cm）', '6、测试累计总跑量（km）',
       '1、【整体】版型评分', '版型评价（文字描述）',
       '适宜的穿着场景', '适合的温度区间（体感温度）',
       '1、该样品可以达到的支撑效果', '其他补充内容',
       '提交人', '提交时间', '更新时间']
def mk(i):
    return ['id-' + str(i), '试穿报告', '跑者' + chr(64 + i), '170', '60',
            '86-72-88', '500', '4', '版型挺合身但腋下偏松',
            '日常通勤', '5-15℃', '中等支撑',
            '面料柔软透气，长跑舒适，吸湿性强',
            '张经理', '2026-01-01 12:00', '2026-01-02 09:00']

rows = [hdr] + [mk(i) for i in range(1, 10)]

block, n = _build_respondent_block(rows)
print(f"=== n = {n}")
print("=== 1st respondent block ===")
print(block.split("\n\n")[0])
print()

print("=== 不应泄漏的噪声列 ===")
all_ok = True
for kw in ['三围', '86-72-88', '提交时间', 'data_id', '更新时间', '5-15℃', '中等支撑', '日常通勤', '500', '张经理']:
    leaked = kw in block
    flag = "LEAK!" if leaked else "clean"
    if leaked:
        all_ok = False
    print(f"  {kw}: {flag}")
print()

print("=== 应保留的（评分 / 文本反馈 / 人物信息） ===")
for kw in ['身高:170CM', '体重:60KG', '版型评分=4', '版型评价', '其他补充内容', '面料柔软']:
    present = kw in block
    flag = "present" if present else "MISSING!"
    if not present:
        all_ok = False
    print(f"  {kw}: {flag}")
print()

# 测试 fix3 Step 4 budget 动态化
print("=== TextBox 24 spec budget 动态化 ===")
spec_t24 = next(s for s in APPAREL_SHAPES if s["name"] == "TextBox 24")
print(f"  original budget: {spec_t24['budget']}")
# 不调 GPT，直接看 fallback
content = _build_content(spec_t24, rows, gpt_enabled=False, model="dummy")
lines = content.split("\r")
print(f"  fallback 输出 lines = {len(lines)}（含标题，n=9 应 = 10）")
for ln in lines:
    print(f"    | {ln}")
print()

# 测试 fix3 Step 5 fallback_map 动态化
print("=== gpt_prompted 优点 fallback 动态化 ===")
spec_t8 = next(s for s in APPAREL_SHAPES if s["name"] == "TextBox 8")
content = _build_content(spec_t8, rows, gpt_enabled=False, model="dummy")
print(f"  优点 fallback (n={n}) →")
for ln in content.split("\r"):
    print(f"    | {ln}")
has_n_ratio = f"({n}/{n})" in content
print(f"  含 ({n}/{n})? {has_n_ratio}")
print()

print("=== overall ===", "PASS" if all_ok else "FAIL")
