# PPT 验收报告

**结论：FAIL**  （必修 4 / 警告 9 / 容忍 0）

- 模式：`production`
- 契约：`acceptance/apparel.json`
- slide pairs：22:13, 23:14

## 摘要

| 层 | 通过 | 警告 | 容忍 | 必修违反 | 降级 |
|---|---|---|---|---|---|
| L0 配对 | 29 | 0 | 0 | 0 |  |
| L1 数据 | 3 | 0 | 0 | 0 |  |
| L2 格式 | 40 | 9 | 0 | 0 |  |
| L3 染色 | 4 | 0 | 0 | 3 |  |
| L4 行为 | 5 | 0 | 0 | 0 |  |
| L5 视觉 | 1 | 0 | 0 | 1 |  |

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

### [L3] p13_textbox14_runs::TextBox 14 [p13] TextBox 14
- runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- 详情：`{"dims": ["rgb", "bold", "italic", "size"], "new_seq": [[0, -1, 0, 20.0], [255, -1, 0, 16.0]], "template_seq": [[0, -1, 0, 20.0], [255, -1, 0, 14.0]]}`

### [L3] p14_textbox23_runs::TextBox 23 [p14] TextBox 23
- runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=14 runs vs template=12 runs
- 详情：`{"dims": ["rgb", "bold", "italic", "size"], "new_seq": [[192, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0]], "template_seq": [[192, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0], [255, -1, 0, 14.0], [0, 0, 0, 14.0]]}`

### [L3] p14_textbox26_runs::TextBox 26 [p14] TextBox 26
- runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=19 runs vs template=23 runs
- 详情：`{"dims": ["rgb", "bold", "italic", "size"], "new_seq": [[12611584, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [0, 0, 0, 16.0], [15773696, -1, 0, 16.0], [0, 0, 0, 16.0], [15773696, -1, 0, 16.0], [0, 0, 0, 16.0]], "template_seq": [[12611584, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [0, -1, 0, 14.0], [0, 0, 0, 14.0], [15773696, -1, 0, 14.0], [0, 0, 0, 14.0], [0, 0, 0, 16.0], [0, -1, 0, 16.0], [0, 0, 0, 16.0]]}`

### [L5] L5_ssim_p14 [p14]
- SSIM=0.7035 (threshold=0.85)
- 详情：`{"ssim": 0.7035109327468748, "threshold": 0.85, "png_new": "D:\\Technique Support\\Claude Code Learning\\3rd-ppt-prj\\acceptance\\acceptance-apparel-fix6b\\visual\\new_023.png", "png_template": "D:\\Technique Support\\Claude Code Learning\\3rd-ppt-prj\\acceptance\\acceptance-apparel-fix6b\\visual\\template_014.png"}`

## 警告清单（不阻断）

- [L2] autosize_global_apparel::TextBox 6 [p13] TextBox 6 — autosize: new=0 template=1
- [L2] autosize_global_apparel::TextBox 14 [p13] TextBox 14 — autosize: new=0 template=1
- [L2] autosize_global_apparel::TextBox 17 [p13] TextBox 17 — autosize: new=0 template=1
- [L2] autosize_global_apparel::TextBox 20 [p13] TextBox 20 — autosize: new=0 template=1
- [L2] autosize_global_apparel::TextBox 50 [p13] TextBox 50 — autosize: new=0 template=1
- [L2] autosize_global_apparel::Rounded Rectangle 53 [p13] Rounded Rectangle 53 — autosize: new=0 template=1
- [L2] autosize_global_apparel::Rounded Rectangle 55 [p13] Rounded Rectangle 55 — autosize: new=0 template=1
- [L2] autosize_global_apparel::TextBox 23 [p14] TextBox 23 — autosize: new=0 template=1
- [L2] autosize_global_apparel::TextBox 26 [p14] TextBox 26 — autosize: new=0 template=1

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
- ✓ [must_fix] p13_chart63_temp_range [p13] — chart series (inline): match
- ✓ [must_fix] p13_temp_mode_label [p13] — text contains '5℃~15℃': True (actual snippet='适宜温度\r5℃~15℃')
- ✓ [must_fix] p13_total_km_label [p13] — text contains '': True (actual snippet='累计跑量km\r571')

### L2 格式
- ✓ [must_fix] geometry_global_apparel::Chart 63 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Straight Connector 4 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Straight Connector 5 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 1 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Oval 3 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 6 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Oval 13 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 14 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Oval 16 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 17 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Oval 19 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 20 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 32 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Oval 49 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 50 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Rounded Rectangle 53 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Rounded Rectangle 55 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Chart 8 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Chart 12 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Chart 15 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Chart 18 [p13] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Straight Connector 4 [p14] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::Straight Connector 5 [p14] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 1 [p14] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 23 [p14] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 26 [p14] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [must_fix] geometry_global_apparel::TextBox 32 [p14] — geom diff worst=left=0.00 (tol=2.0)
- ✓ [warn] p13_textbox24_geometry_relaxed::TextBox 24 [p13] — geom diff worst=height=98.96 (tol=200.0)
- ✓ [warn] p13_textbox24_geometry_relaxed::TextBox 24 [p14] — geom diff worst=height=98.96 (tol=200.0)
- ✓ [warn] autosize_global_apparel::TextBox 1 [p13] — autosize: new=1 template=1
- ✓ [warn] autosize_global_apparel::Oval 3 [p13] — autosize: new=1 template=1
- ⚠ [warn] autosize_global_apparel::TextBox 6 [p13] — autosize: new=0 template=1
- ✓ [warn] autosize_global_apparel::Oval 13 [p13] — autosize: new=1 template=1
- ⚠ [warn] autosize_global_apparel::TextBox 14 [p13] — autosize: new=0 template=1
- ✓ [warn] autosize_global_apparel::Oval 16 [p13] — autosize: new=1 template=1
- ⚠ [warn] autosize_global_apparel::TextBox 17 [p13] — autosize: new=0 template=1
- ✓ [warn] autosize_global_apparel::Oval 19 [p13] — autosize: new=1 template=1
- ⚠ [warn] autosize_global_apparel::TextBox 20 [p13] — autosize: new=0 template=1
- ✓ [warn] autosize_global_apparel::TextBox 24 [p13] — autosize: new=0 template=0
- ✓ [warn] autosize_global_apparel::TextBox 32 [p13] — autosize: new=1 template=1
- ✓ [warn] autosize_global_apparel::Oval 49 [p13] — autosize: new=1 template=1
- ⚠ [warn] autosize_global_apparel::TextBox 50 [p13] — autosize: new=0 template=1
- ⚠ [warn] autosize_global_apparel::Rounded Rectangle 53 [p13] — autosize: new=0 template=1
- ⚠ [warn] autosize_global_apparel::Rounded Rectangle 55 [p13] — autosize: new=0 template=1
- ✓ [warn] autosize_global_apparel::TextBox 1 [p14] — autosize: new=1 template=1
- ⚠ [warn] autosize_global_apparel::TextBox 23 [p14] — autosize: new=0 template=1
- ✓ [warn] autosize_global_apparel::TextBox 24 [p14] — autosize: new=0 template=0
- ⚠ [warn] autosize_global_apparel::TextBox 26 [p14] — autosize: new=0 template=1
- ✓ [warn] autosize_global_apparel::TextBox 32 [p14] — autosize: new=1 template=1

### L3 染色
- ✓ [must_fix] p13_textbox6_runs::TextBox 6 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✗ [must_fix] p13_textbox14_runs::TextBox 14 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_textbox17_runs::TextBox 17 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_textbox20_runs::TextBox 20 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✓ [must_fix] p13_textbox50_runs::TextBox 50 [p13] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=2 runs vs template=2 runs
- ✗ [must_fix] p14_textbox23_runs::TextBox 23 [p14] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=14 runs vs template=12 runs
- ✗ [must_fix] p14_textbox26_runs::TextBox 26 [p14] — runs structure (dims=['rgb', 'bold', 'italic', 'size']): new=19 runs vs template=23 runs

### L4 行为
- ✓ [must_fix] no_silent_com_failure — no forbidden events: ['com_api_failed_but_continued']
- ✓ [must_fix] p14_gpt_strengths_called — trace events present: 1/1; missing=[]
- ✓ [must_fix] p14_gpt_drawbacks_called — trace events present: 1/1; missing=[]
- ✓ [must_fix] gpt_respondent_info_called — event gpt_respondent_info count=2 (min=1 max=2)
- ✓ [must_fix] chart63_write_ok — trace events present: 1/1; missing=[]

### L5 视觉
- ✓ [must_fix] L5_ssim_p13 [p13] — SSIM=0.9722 (threshold=0.85)
- ✗ [must_fix] L5_ssim_p14 [p14] — SSIM=0.7035 (threshold=0.85)

