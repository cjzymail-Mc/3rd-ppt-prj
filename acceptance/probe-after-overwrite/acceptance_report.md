# PPT 验收报告

**结论：PASS**  （必修 0 / 警告 2 / 容忍 0）

- 模式：`production`
- 契约：`acceptance/apparel.json`
- slide pairs：9:13

## 摘要

| 层 | 通过 | 警告 | 容忍 | 必修违反 | 降级 |
|---|---|---|---|---|---|
| L0 配对 | 22 | 2 | 0 | 0 |  |
| L3 染色 | 7 | 0 | 0 | 0 |  |

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

## 警告清单（不阻断）

- [L0] L0_only_new_Rounded Rectangle 2 [p13] Rounded Rectangle 2 — 仅在 new 中：Rounded Rectangle 2（template 无对应 shape）
- [L0] L0_only_new_Rounded Rectangle 7 [p13] Rounded Rectangle 7 — 仅在 new 中：Rounded Rectangle 7（template 无对应 shape）

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
- ⚠ [warn] L0_only_new_Rounded Rectangle 2 [p13] — 仅在 new 中：Rounded Rectangle 2（template 无对应 shape）
- ⚠ [warn] L0_only_new_Rounded Rectangle 7 [p13] — 仅在 new 中：Rounded Rectangle 7（template 无对应 shape）

### L3 染色
- ✓ [must_fix] p13_textbox6_runs::TextBox 6 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_textbox14_runs::TextBox 14 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_textbox17_runs::TextBox 17 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_textbox20_runs::TextBox 20 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_textbox50_runs::TextBox 50 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_rr53_two_run_signature::Rounded Rectangle 53 [p13] — runs signature (dims=['rgb', 'bold', 'size'], ignore_ws=True): new=2 runs vs expected=2 runs
- ✓ [must_fix] p13_rr55_two_run_signature::Rounded Rectangle 55 [p13] — runs signature (dims=['rgb', 'bold', 'size'], ignore_ws=True): new=2 runs vs expected=2 runs

