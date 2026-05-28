# PPT 验收报告

**结论：FAIL**  （必修 4 / 警告 0 / 容忍 0）

- 模式：`production`
- 契约：`acceptance/apparel.json`
- slide pairs：12:13, 13:14

## 摘要

| 层 | 通过 | 警告 | 容忍 | 必修违反 | 降级 |
|---|---|---|---|---|---|
| L0 配对 | 29 | 0 | 0 | 0 |  |
| L1 数据 | 2 | 0 | 0 | 4 |  |

## L0 配对命中

- [p13] Chart 63 ↔ Chart 63 via exact
- [p13] Straight Connector 4 ↔ Straight Connector 4 via exact
- [p13] Straight Connector 5 ↔ Straight Connector 5 via exact
- [p13] TextBox 1 ↔ TextBox 1 via exact
- [p13] Oval 3 ↔ Oval 3 via exact
- [p13] TextBox 6 ↔ TextBox 6 via exact
- [p13] Oval 13 ↔ Oval 13 via exact
- [p13] TextBox 14 ↔ TextBox 14 via exact
- [p13] Oval 16 ↔ Oval 16 via exact
- [p13] TextBox 17 ↔ TextBox 17 via exact
- [p13] Oval 19 ↔ Oval 19 via exact
- [p13] TextBox 20 ↔ TextBox 20 via exact
- [p13] TextBox 24 ↔ TextBox 24 via exact
- [p13] TextBox 32 ↔ TextBox 32 via exact
- [p13] Oval 49 ↔ Oval 49 via exact
- [p13] TextBox 50 ↔ TextBox 50 via exact
- [p13] Rounded Rectangle 53 ↔ Rounded Rectangle 53 via exact
- [p13] Rounded Rectangle 55 ↔ Rounded Rectangle 55 via exact
- [p13] Chart 8 ↔ Chart 7 via ignore_chart_auto_renumber(dist=0.0)
- [p13] Chart 12 ↔ Chart 9 via ignore_chart_auto_renumber(dist=0.0)
- [p13] Chart 15 ↔ Chart 10 via ignore_chart_auto_renumber(dist=0.0)
- [p13] Chart 18 ↔ Chart 11 via ignore_chart_auto_renumber(dist=0.0)
- [p14] Straight Connector 4 ↔ Straight Connector 4 via exact
- [p14] Straight Connector 5 ↔ Straight Connector 5 via exact
- [p14] TextBox 1 ↔ TextBox 1 via exact
- [p14] TextBox 23 ↔ TextBox 23 via exact
- [p14] TextBox 24 ↔ TextBox 24 via exact
- [p14] TextBox 26 ↔ TextBox 26 via exact
- [p14] TextBox 32 ↔ TextBox 32 via exact

## 必修违反清单（先看这里）

### [L1] p13_chart63_temp_range [p13] Chart 63
- chart series (excel DSL 执行失败: (-2146827864, 'OLE error 0x800a01a8', None, None)): expected 为 None
- 详情：`{"actual": [[5.0, 15.0], [20.0, 17.0], [10.0, 3.0]], "expected": null, "expected_from": "excel DSL 执行失败: (-2146827864, 'OLE error 0x800a01a8', None, None)"}`

### [L1] p13_chart63_temp_range [p14] Chart 63
- shape 'Chart 63' 在 new 侧找不到对应（L0 配对失败）

### [L1] p13_temp_mode_label [p14] TextBox 50
- shape 'TextBox 50' 在 new 侧找不到对应（L0 配对失败）

### [L1] p13_total_km_label [p14] Rounded Rectangle 53
- shape 'Rounded Rectangle 53' 在 new 侧找不到对应（L0 配对失败）

## 详细清单（含通过项）

### L0 配对
- ✓ [must_fix] L0_pair_Chart 63 [p13] — Chart 63 ↔ Chart 63 via exact
- ✓ [must_fix] L0_pair_Straight Connector 4 [p13] — Straight Connector 4 ↔ Straight Connector 4 via exact
- ✓ [must_fix] L0_pair_Straight Connector 5 [p13] — Straight Connector 5 ↔ Straight Connector 5 via exact
- ✓ [must_fix] L0_pair_TextBox 1 [p13] — TextBox 1 ↔ TextBox 1 via exact
- ✓ [must_fix] L0_pair_Oval 3 [p13] — Oval 3 ↔ Oval 3 via exact
- ✓ [must_fix] L0_pair_TextBox 6 [p13] — TextBox 6 ↔ TextBox 6 via exact
- ✓ [must_fix] L0_pair_Oval 13 [p13] — Oval 13 ↔ Oval 13 via exact
- ✓ [must_fix] L0_pair_TextBox 14 [p13] — TextBox 14 ↔ TextBox 14 via exact
- ✓ [must_fix] L0_pair_Oval 16 [p13] — Oval 16 ↔ Oval 16 via exact
- ✓ [must_fix] L0_pair_TextBox 17 [p13] — TextBox 17 ↔ TextBox 17 via exact
- ✓ [must_fix] L0_pair_Oval 19 [p13] — Oval 19 ↔ Oval 19 via exact
- ✓ [must_fix] L0_pair_TextBox 20 [p13] — TextBox 20 ↔ TextBox 20 via exact
- ✓ [must_fix] L0_pair_TextBox 24 [p13] — TextBox 24 ↔ TextBox 24 via exact
- ✓ [must_fix] L0_pair_TextBox 32 [p13] — TextBox 32 ↔ TextBox 32 via exact
- ✓ [must_fix] L0_pair_Oval 49 [p13] — Oval 49 ↔ Oval 49 via exact
- ✓ [must_fix] L0_pair_TextBox 50 [p13] — TextBox 50 ↔ TextBox 50 via exact
- ✓ [must_fix] L0_pair_Rounded Rectangle 53 [p13] — Rounded Rectangle 53 ↔ Rounded Rectangle 53 via exact
- ✓ [must_fix] L0_pair_Rounded Rectangle 55 [p13] — Rounded Rectangle 55 ↔ Rounded Rectangle 55 via exact
- ✓ [must_fix] L0_pair_Chart 8 [p13] — Chart 8 ↔ Chart 7 via ignore_chart_auto_renumber(dist=0.0)
- ✓ [must_fix] L0_pair_Chart 12 [p13] — Chart 12 ↔ Chart 9 via ignore_chart_auto_renumber(dist=0.0)
- ✓ [must_fix] L0_pair_Chart 15 [p13] — Chart 15 ↔ Chart 10 via ignore_chart_auto_renumber(dist=0.0)
- ✓ [must_fix] L0_pair_Chart 18 [p13] — Chart 18 ↔ Chart 11 via ignore_chart_auto_renumber(dist=0.0)
- ✓ [must_fix] L0_pair_Straight Connector 4 [p14] — Straight Connector 4 ↔ Straight Connector 4 via exact
- ✓ [must_fix] L0_pair_Straight Connector 5 [p14] — Straight Connector 5 ↔ Straight Connector 5 via exact
- ✓ [must_fix] L0_pair_TextBox 1 [p14] — TextBox 1 ↔ TextBox 1 via exact
- ✓ [must_fix] L0_pair_TextBox 23 [p14] — TextBox 23 ↔ TextBox 23 via exact
- ✓ [must_fix] L0_pair_TextBox 24 [p14] — TextBox 24 ↔ TextBox 24 via exact
- ✓ [must_fix] L0_pair_TextBox 26 [p14] — TextBox 26 ↔ TextBox 26 via exact
- ✓ [must_fix] L0_pair_TextBox 32 [p14] — TextBox 32 ↔ TextBox 32 via exact

### L1 数据
- ✗ [must_fix] p13_chart63_temp_range [p13] — chart series (excel DSL 执行失败: (-2146827864, 'OLE error 0x800a01a8', None, None)): expected 为 None
- ✗ [must_fix] p13_chart63_temp_range [p14] — shape 'Chart 63' 在 new 侧找不到对应（L0 配对失败）
- ✓ [must_fix] p13_temp_mode_label [p13] — text contains '': True (actual snippet='适宜温度\r5~15℃')
- ✗ [must_fix] p13_temp_mode_label [p14] — shape 'TextBox 50' 在 new 侧找不到对应（L0 配对失败）
- ✓ [must_fix] p13_total_km_label [p13] — text contains '': True (actual snippet='累计跑量km\r571')
- ✗ [must_fix] p13_total_km_label [p14] — shape 'Rounded Rectangle 53' 在 new 侧找不到对应（L0 配对失败）

