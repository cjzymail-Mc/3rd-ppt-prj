# PPT 验收报告

**结论：PASS**  （必修 0 / 警告 0 / 容忍 0）

- 模式：`production`
- 契约：`D:/Technique Support/Claude Code Learning/3rd-ppt-prj/pipeline-progress/_acceptance_contract.auto.json`
- slide pairs：2:2

## 摘要

| 层 | 通过 | 警告 | 容忍 | 必修违反 | 降级 |
|---|---|---|---|---|---|
| L0 配对 | 22 | 0 | 0 | 0 |  |
| L1 数据 | 0 | 0 | 0 | 0 |  |
| L2 格式 | 38 | 0 | 0 | 0 |  |
| L3 染色 | 10 | 0 | 0 | 0 |  |
| L4 行为 | 0 | 0 | 0 | 0 | ⚠ 无 pipeline trace（path=None）；L4 规则全部降级为 warn，无法验证 COM 失败 / GPT 调用类暗坑。 |
| L5 视觉 | 1 | 0 | 0 | 0 |  |

## L0 配对命中

- [p2] Straight Connector 4 ↔ Straight Connector 4 via exact
- [p2] Straight Connector 5 ↔ Straight Connector 5 via exact
- [p2] TextBox 1 ↔ TextBox 1 via exact
- [p2] Oval 3 ↔ Oval 3 via exact
- [p2] TextBox 6 ↔ TextBox 6 via exact
- [p2] Chart 12 ↔ Chart 12 via exact
- [p2] Oval 13 ↔ Oval 13 via exact
- [p2] TextBox 14 ↔ TextBox 14 via exact
- [p2] Chart 15 ↔ Chart 15 via exact
- [p2] Oval 16 ↔ Oval 16 via exact
- [p2] TextBox 17 ↔ TextBox 17 via exact
- [p2] Chart 18 ↔ Chart 18 via exact
- [p2] Oval 19 ↔ Oval 19 via exact
- [p2] TextBox 20 ↔ TextBox 20 via exact
- [p2] Chart 21 ↔ Chart 21 via exact
- [p2] TextBox 23 ↔ TextBox 23 via exact
- [p2] TextBox 24 ↔ TextBox 24 via exact
- [p2] TextBox 26 ↔ TextBox 26 via exact
- [p2] TextBox 32 ↔ TextBox 32 via exact
- [p2] TextBox 8 ↔ TextBox 8 via exact
- [p2] TextBox 22 ↔ TextBox 22 via exact
- [p2] Rectangle 25 ↔ Rectangle 25 via exact

## 详细清单（含通过项）

### L0 配对
- ✓ [must_fix] L0_pair_Straight Connector 4 [p2] — Straight Connector 4 ↔ Straight Connector 4 via exact
- ✓ [must_fix] L0_pair_Straight Connector 5 [p2] — Straight Connector 5 ↔ Straight Connector 5 via exact
- ✓ [must_fix] L0_pair_TextBox 1 [p2] — TextBox 1 ↔ TextBox 1 via exact
- ✓ [must_fix] L0_pair_Oval 3 [p2] — Oval 3 ↔ Oval 3 via exact
- ✓ [must_fix] L0_pair_TextBox 6 [p2] — TextBox 6 ↔ TextBox 6 via exact
- ✓ [must_fix] L0_pair_Chart 12 [p2] — Chart 12 ↔ Chart 12 via exact
- ✓ [must_fix] L0_pair_Oval 13 [p2] — Oval 13 ↔ Oval 13 via exact
- ✓ [must_fix] L0_pair_TextBox 14 [p2] — TextBox 14 ↔ TextBox 14 via exact
- ✓ [must_fix] L0_pair_Chart 15 [p2] — Chart 15 ↔ Chart 15 via exact
- ✓ [must_fix] L0_pair_Oval 16 [p2] — Oval 16 ↔ Oval 16 via exact
- ✓ [must_fix] L0_pair_TextBox 17 [p2] — TextBox 17 ↔ TextBox 17 via exact
- ✓ [must_fix] L0_pair_Chart 18 [p2] — Chart 18 ↔ Chart 18 via exact
- ✓ [must_fix] L0_pair_Oval 19 [p2] — Oval 19 ↔ Oval 19 via exact
- ✓ [must_fix] L0_pair_TextBox 20 [p2] — TextBox 20 ↔ TextBox 20 via exact
- ✓ [must_fix] L0_pair_Chart 21 [p2] — Chart 21 ↔ Chart 21 via exact
- ✓ [must_fix] L0_pair_TextBox 23 [p2] — TextBox 23 ↔ TextBox 23 via exact
- ✓ [must_fix] L0_pair_TextBox 24 [p2] — TextBox 24 ↔ TextBox 24 via exact
- ✓ [must_fix] L0_pair_TextBox 26 [p2] — TextBox 26 ↔ TextBox 26 via exact
- ✓ [must_fix] L0_pair_TextBox 32 [p2] — TextBox 32 ↔ TextBox 32 via exact
- ✓ [must_fix] L0_pair_TextBox 8 [p2] — TextBox 8 ↔ TextBox 8 via exact
- ✓ [must_fix] L0_pair_TextBox 22 [p2] — TextBox 22 ↔ TextBox 22 via exact
- ✓ [must_fix] L0_pair_Rectangle 25 [p2] — Rectangle 25 ↔ Rectangle 25 via exact

### L2 格式
- ✓ [must_fix] geometry_global::Straight Connector 4 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Straight Connector 5 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 1 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Oval 3 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 6 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Chart 12 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Oval 13 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 14 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Chart 15 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Oval 16 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 17 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Chart 18 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Oval 19 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 20 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Chart 21 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 23 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 24 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 26 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 32 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 8 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::TextBox 22 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global::Rectangle 25 [p2] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [warn] autosize_global::TextBox 1 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::Oval 3 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 6 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::Oval 13 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 14 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::Oval 16 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 17 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::Oval 19 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 20 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 23 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 24 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 26 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 32 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 8 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::TextBox 22 [p2] — autosize: new=1 template=1
- ✓ [warn] autosize_global::Rectangle 25 [p2] — autosize: new=1 template=1

### L3 染色
- ✓ [warn] TextBox_6_paras::TextBox 6 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 1 paras, all ok
- ✓ [warn] TextBox_14_paras::TextBox 14 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 1 paras, all ok
- ✓ [warn] TextBox_17_paras::TextBox 17 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 1 paras, all ok
- ✓ [warn] TextBox_20_paras::TextBox 20 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 1 paras, all ok
- ✓ [warn] TextBox_23_paras::TextBox 23 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 1 paras, all ok
- ✓ [warn] TextBox_24_paras::TextBox 24 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 5 paras, all ok
- ✓ [warn] TextBox_26_paras::TextBox 26 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 1 paras, all ok
- ✓ [warn] TextBox_8_paras::TextBox 8 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 3 paras, all ok
- ✓ [warn] TextBox_22_paras::TextBox 22 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 2 paras, all ok
- ✓ [warn] Rectangle_25_paras::Rectangle 25 [p2] — paragraphs_match_signature (dims=['rgb', 'bold', 'size']): 1 paras, all ok

### L5 视觉
- ✓ [must_fix] L5_ssim_p2 [p2] — SSIM=1.0000 (threshold=0.85)

