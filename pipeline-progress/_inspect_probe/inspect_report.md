# PPT Inspect 报告

> 由 `inspect-ppt-template` skill 生成。

> 用于写 `template_mapping.py` / `*_mapping.json` 时定位 pptx shape names。


## PPTX: `D:\Technique Support\Claude Code Learning\3rd-ppt-prj\template\apparel-page13-14-template.pptx`

- 模式：`open`（active=桥接运行中 PowerPoint / open=ReadOnly 新进程）
- 总页数：17
- 页面过滤：`[13, 14]`（仅扫这些 slide）

### Slide 13（22 shapes）

| idx | name | type | text? | chart? | L/T/W/H | text 预览 |
|-----|------|------|-------|--------|---------|----------|
| 1 | `Chart 63` | 3 |  | ✓ | 163/410/427/113 |  |
| 2 | `Straight Connector 4` | 9 |  |  | 340/87/285/0 |  |
| 3 | `Straight Connector 5` | 9 |  |  | 55/87/285/0 |  |
| 4 | `TextBox 1` | 17 | ✓ |  | 43/53/250/32 | 服装试穿反馈结果 |
| 5 | `Oval 3` | 1 | ✓ |  | 36/102/113/113 |  |
| 6 | `TextBox 6` | 17 | ✓ |  | 30/122/125/73 | 版型↵3.98 / 5 |
| 7 | `Oval 13` | 1 | ✓ |  | 332/107/113/113 |  |
| 8 | `TextBox 14` | 17 | ✓ |  | 327/125/125/69 | 面料↵3.96 / 5 |
| 9 | `Oval 16` | 1 | ✓ |  | 36/260/113/113 |  |
| 10 | `TextBox 17` | 17 | ✓ |  | 30/282/125/73 | 吸湿排汗↵3.61 / 5 |
| 11 | `Oval 19` | 1 | ✓ |  | 328/260/113/113 |  |
| 12 | `TextBox 20` | 17 | ✓ |  | 327/280/125/73 | 速干↵3.52 / 5 |
| 13 | `TextBox 24` | 17 | ✓ |  | 831/54/125/148 | 受试者信息 Information↵A: 162CM / 53 KG↵B: 161CM / 53 KG↵C: 170CM / 56 KG↵D: 156CM / 45 KG↵E: 168CM / 60 KG↵F: 163CM / 55 KG↵... |
| 14 | `TextBox 32` | 17 | ✓ |  | 21/9/578/29 | 试穿反馈【 Athletes’ Feedback】 |
| 15 | `Chart 7` | 3 |  | ✓ | 126/84/190/150 |  |
| 16 | `Chart 9` | 3 |  | ✓ | 406/106/224/98 |  |
| 17 | `Chart 10` | 3 |  | ✓ | 145/271/190/98 |  |
| 18 | `Chart 11` | 3 |  | ✓ | 436/268/190/98 |  |
| 19 | `Oval 49` | 1 | ✓ |  | 34/402/113/113 |  |
| 20 | `TextBox 50` | 17 | ✓ |  | 26/422/125/73 | 适宜温度↵15~25℃ |
| 21 | `Rounded Rectangle 53` | 1 | ✓ |  | 848/227/92/55 | 累计跑量km↵671 |
| 22 | `Rounded Rectangle 55` | 1 | ✓ |  | 849/309/92/55 | 定位日常训练7/9 |

**Run-level detail（--full 模式）：**

- `TextBox 1` (idx=4, paragraphs=1):
  - p1 align=1 runs=1: "服装试穿反馈结果"
    - 20.0pt #000000 B font=Arial: "服装试穿反馈结果"
- `TextBox 6` (idx=6, paragraphs=2):
  - p1 align=2 runs=2: "版型"
    - 20.0pt #000000 B font=Arial: "版型"
    - 20.0pt #000000 B font=Arial: "↵"
  - p2 align=2 runs=1: "3.98 / 5"
    - 16.0pt #FF0000 B font=Arial: "3.98 / 5"
- `TextBox 14` (idx=8, paragraphs=2):
  - p1 align=2 runs=2: "面料"
    - 20.0pt #000000 B font=Arial: "面料"
    - 20.0pt #000000 B font=Arial: "↵"
  - p2 align=2 runs=1: "3.96 / 5"
    - 14.0pt #FF0000 B font=Arial: "3.96 / 5"
- `TextBox 17` (idx=10, paragraphs=2):
  - p1 align=2 runs=2: "吸湿排汗"
    - 20.0pt #000000 B font=Arial: "吸湿排汗"
    - 20.0pt #000000 B font=Arial: "↵"
  - p2 align=2 runs=1: "3.61 / 5"
    - 16.0pt #FF0000 B font=Arial: "3.61 / 5"
- `TextBox 20` (idx=12, paragraphs=2):
  - p1 align=2 runs=2: "速干"
    - 20.0pt #000000 B font=Arial: "速干"
    - 20.0pt #000000 B font=Arial: "↵"
  - p2 align=2 runs=1: "3.52 / 5"
    - 16.0pt #FF0000 B font=Arial: "3.52 / 5"
- `TextBox 24` (idx=13, paragraphs=10):
  - p1 align=2 runs=2: "受试者信息 Information"
    - 9.0pt #000000 B font=微软雅黑: "受试者信息 "
    - 9.0pt #000000 B font=微软雅黑: "Information↵"
  - p2 align=2 runs=1: "A: 162CM / 53 KG"
    - 9.0pt #000000 B font=微软雅黑: "A: 162CM / 53 KG↵"
  - p3 align=2 runs=1: "B: 161CM / 53 KG"
    - 9.0pt #000000 B font=微软雅黑: "B: 161CM / 53 KG↵"
  - p4 align=2 runs=1: "C: 170CM / 56 KG"
    - 9.0pt #000000 B font=微软雅黑: "C: 170CM / 56 KG↵"
  - p5 align=2 runs=1: "D: 156CM / 45 KG"
    - 9.0pt #000000 B font=微软雅黑: "D: 156CM / 45 KG↵"
  - p6 align=2 runs=1: "E: 168CM / 60 KG"
    - 9.0pt #000000 B font=微软雅黑: "E: 168CM / 60 KG↵"
  - p7 align=2 runs=1: "F: 163CM / 55 KG"
    - 9.0pt #000000 B font=微软雅黑: "F: 163CM / 55 KG↵"
  - p8 align=2 runs=1: "G: 160CM / 50 KG"
    - 9.0pt #000000 B font=微软雅黑: "G: 160CM / 50 KG↵"
  - p9 align=2 runs=1: "H: 160CM / 50 KG"
    - 9.0pt #000000 B font=微软雅黑: "H: 160CM / 50 KG↵"
  - p10 align=2 runs=1: "I: 162CM / 58 KG"
    - 9.0pt #000000 B font=微软雅黑: "I: 162CM / 58 KG"
- `TextBox 32` (idx=14, paragraphs=1):
  - p1 align=1 runs=4: "试穿反馈【 Athletes’ Feedback】"
    - 18.0pt #FFFFFF B font=微软雅黑: "试穿反馈"
    - 14.0pt #FFC000 B font=微软雅黑: "【"
    - 14.0pt #FFFFFF B font=Helvetica: " "
    - 14.0pt #FFC000 B font=微软雅黑: "Athletes’ Feedback】"
- `TextBox 50` (idx=20, paragraphs=2):
  - p1 align=2 runs=2: "适宜温度"
    - 20.0pt #000000 B font=Arial: "适宜温度"
    - 20.0pt #000000 B font=Arial: "↵"
  - p2 align=2 runs=2: "15~25℃"
    - 16.0pt #FF0000 B font=Arial: "15~25"
    - 16.0pt #FF0000 B font=Arial: "℃"
- `Rounded Rectangle 53` (idx=21, paragraphs=2):
  - p1 align=2 runs=2: "累计跑量km"
    - 11.0pt #FFFFFF B font=Arial: "累计跑量"
    - 11.0pt #FFFFFF B font=Arial: "km↵"
  - p2 align=2 runs=1: "671"
    - 24.0pt #FFFFFF B font=Arial: "671"
- `Rounded Rectangle 55` (idx=22, paragraphs=1):
  - p1 align=2 runs=2: "定位日常训练7/9"
    - 11.0pt #FFFFFF B font=Arial: "定位日常训练"
    - 24.0pt #FFFFFF B font=Arial: "7/9"

### Slide 14（7 shapes）

| idx | name | type | text? | chart? | L/T/W/H | text 预览 |
|-----|------|------|-------|--------|---------|----------|
| 1 | `Straight Connector 4` | 9 |  |  | 340/87/285/0 |  |
| 2 | `Straight Connector 5` | 9 |  |  | 55/87/285/0 |  |
| 3 | `TextBox 1` | 17 | ✓ |  | 43/53/250/32 | 服装试穿反馈结果 |
| 4 | `TextBox 23` | 17 | ✓ |  | 34/128/556/134 | 优点 strengths↵整体版型合身、修身显身材，上身贴合，动作舒展不受限（9/9）。↵具备一定支撑性，覆盖中低到高强度训练（7/9）。↵面料有亲肤性与一定耐用性，多名反馈不起球不勾丝↵短距离及日常训练场景下舒适性评价较好。 |
| 5 | `TextBox 24` | 17 | ✓ |  | 831/54/125/148 | 受试者信息 Information↵A: 162CM / 53 KG↵B: 161CM / 53 KG↵C: 170CM / 56 KG↵D: 156CM / 45 KG↵E: 168CM / 60 KG↵F: 163CM / 55 KG↵... |
| 6 | `TextBox 26` | 17 | ✓ |  | 35/306/554/189 | 缺点 drawbacks↵面料弹性不够（2/9）↵前胸闷热、面料偏厚（4/9）↵腋下摩擦、副乳硌感、胸下磨皮较集中，长距离或出汗后更明显（8/9）。↵透气排汗不足较突出，局部有只吸不排、贴身感，速干较差（6/9）。↵后背口袋不便，补给袋定位... |
| 7 | `TextBox 32` | 17 | ✓ |  | 21/9/578/29 | 试穿反馈【 Athletes’ Feedback】 |

**Run-level detail（--full 模式）：**

- `TextBox 1` (idx=3, paragraphs=1):
  - p1 align=1 runs=1: "服装试穿反馈结果"
    - 20.0pt #000000 B font=Arial: "服装试穿反馈结果"
- `TextBox 23` (idx=4, paragraphs=3):
  - p1 align=2 runs=2: "优点 strengths"
    - 14.0pt #C00000 B font=微软雅黑: "优点 "
    - 14.0pt #C00000 B font=微软雅黑: "strengths↵"
  - p2 align=2 runs=16: "整体版型合身、修身显身材，上身贴合，动作舒展不受限（9/9）。↵具备一定支撑性，覆盖中低到高强度训练（7/9）。↵面料有..."
    - 14.0pt #000000 font=Arial: "整体"
    - 14.0pt #FF0000 B font=Arial: "版型合身"
    - 14.0pt #000000 font=Arial: "、"
    - 14.0pt #FF0000 B font=Arial: "修身显身材"
    - 14.0pt #000000 font=Arial: "，上身贴合，动作舒展不受限（"
    - 14.0pt #000000 font=Arial: "9/9"
    - 14.0pt #000000 font=Arial: "）。↵具备一定"
    - 14.0pt #FF0000 B font=Arial: "支撑性"
    - 14.0pt #000000 font=Arial: "，覆盖中低到高强度训练（"
    - 14.0pt #000000 font=Arial: "7/9"
    - 14.0pt #000000 font=Arial: "）。↵面料有"
    - 14.0pt #FF0000 B font=Arial: "亲肤性"
    - 14.0pt #000000 font=Arial: "与一定"
    - 14.0pt #FF0000 B font=Arial: "耐用性"
    - 14.0pt #000000 font=Arial: "，多名反馈不起球不勾丝"
    - 14.0pt #000000 font=Arial: "↵"
  - p3 align=2 runs=1: "短距离及日常训练场景下舒适性评价较好。"
    - 14.0pt #000000 font=Arial: "短距离及日常训练场景下舒适性评价较好。"
- `TextBox 24` (idx=5, paragraphs=10):
  - p1 align=2 runs=2: "受试者信息 Information"
    - 9.0pt #000000 B font=微软雅黑: "受试者信息 "
    - 9.0pt #000000 B font=微软雅黑: "Information↵"
  - p2 align=2 runs=1: "A: 162CM / 53 KG"
    - 9.0pt #000000 B font=微软雅黑: "A: 162CM / 53 KG↵"
  - p3 align=2 runs=1: "B: 161CM / 53 KG"
    - 9.0pt #000000 B font=微软雅黑: "B: 161CM / 53 KG↵"
  - p4 align=2 runs=1: "C: 170CM / 56 KG"
    - 9.0pt #000000 B font=微软雅黑: "C: 170CM / 56 KG↵"
  - p5 align=2 runs=1: "D: 156CM / 45 KG"
    - 9.0pt #000000 B font=微软雅黑: "D: 156CM / 45 KG↵"
  - p6 align=2 runs=1: "E: 168CM / 60 KG"
    - 9.0pt #000000 B font=微软雅黑: "E: 168CM / 60 KG↵"
  - p7 align=2 runs=1: "F: 163CM / 55 KG"
    - 9.0pt #000000 B font=微软雅黑: "F: 163CM / 55 KG↵"
  - p8 align=2 runs=1: "G: 160CM / 50 KG"
    - 9.0pt #000000 B font=微软雅黑: "G: 160CM / 50 KG↵"
  - p9 align=2 runs=1: "H: 160CM / 50 KG"
    - 9.0pt #000000 B font=微软雅黑: "H: 160CM / 50 KG↵"
  - p10 align=2 runs=1: "I: 162CM / 58 KG"
    - 9.0pt #000000 B font=微软雅黑: "I: 162CM / 58 KG"
- `TextBox 26` (idx=6, paragraphs=5):
  - p1 align=2 runs=2: "缺点 drawbacks"
    - 14.0pt #0070C0 B font=微软雅黑: "缺点 "
    - 14.0pt #0070C0 B font=微软雅黑: "drawbacks↵"
  - p2 align=2 runs=5: "面料弹性不够（2/9）"
    - 14.0pt #000000 font=Arial: "面料"
    - 14.0pt #00B0F0 B font=Arial: "弹性不够（"
    - 14.0pt #00B0F0 B font=Arial: "2/9"
    - 14.0pt #00B0F0 B font=Arial: "）"
    - 14.0pt #00B0F0 B font=Arial: "↵"
  - p3 align=2 runs=8: "前胸闷热、面料偏厚（4/9）"
    - 14.0pt #000000 font=Arial: "前胸"
    - 14.0pt #00B0F0 B font=Arial: "闷热"
    - 14.0pt #000000 font=Arial: "、"
    - 14.0pt #00B0F0 B font=Arial: "面料偏厚"
    - 14.0pt #000000 font=Arial: "（"
    - 14.0pt #000000 font=Arial: "4/9"
    - 14.0pt #000000 font=Arial: "）"
    - 14.0pt #00B0F0 B font=Arial: "↵"
  - p4 align=2 runs=18: "腋下摩擦、副乳硌感、胸下磨皮较集中，长距离或出汗后更明显（8/9）。↵透气排汗不足较突出，局部有只吸不排、贴身感，速干较..."
    - 14.0pt #000000 font=Arial: "腋下"
    - 14.0pt #00B0F0 B font=Arial: "摩擦"
    - 14.0pt #000000 font=Arial: "、"
    - 14.0pt #00B0F0 B font=Arial: "副乳硌感"
    - 14.0pt #000000 font=Arial: "、"
    - 14.0pt #00B0F0 B font=Arial: "胸下磨皮"
    - 14.0pt #000000 font=Arial: "较集中，长距离或出汗后更明显（"
    - 14.0pt #000000 font=Arial: "8/9"
    - 14.0pt #000000 font=Arial: "）。"
    - 14.0pt #000000 B font=Arial: "↵"
    - 14.0pt #000000 font=Arial: "透气"
    - 14.0pt #00B0F0 B font=Arial: "排汗不足"
    - 14.0pt #000000 font=Arial: "较突出，局部有只吸不排、贴身感，速干较差（"
    - 14.0pt #000000 font=Arial: "6/9"
    - 14.0pt #000000 font=Arial: "）。↵后背口袋不便，补给袋定位与开口方向影响取用（"
    - 14.0pt #000000 font=Arial: "4/9"
    - 14.0pt #000000 font=Arial: "）；"
    - 14.0pt #000000 font=Arial: "↵"
  - p5 align=2 runs=6: "建议改为侧向开口/更低位置设计。"
    - 14.0pt #000000 font=Arial: "建议"
    - 16.0pt #000000 font=微软雅黑: "改为"
    - 16.0pt #000000 B font=微软雅黑: "侧向开口"
    - 16.0pt #000000 B font=微软雅黑: "/"
    - 16.0pt #000000 B font=微软雅黑: "更低位置"
    - 16.0pt #000000 font=微软雅黑: "设计。"
- `TextBox 32` (idx=7, paragraphs=1):
  - p1 align=1 runs=4: "试穿反馈【 Athletes’ Feedback】"
    - 18.0pt #FFFFFF B font=微软雅黑: "试穿反馈"
    - 14.0pt #FFC000 B font=微软雅黑: "【"
    - 14.0pt #FFFFFF B font=Helvetica: " "
    - 14.0pt #FFC000 B font=微软雅黑: "Athletes’ Feedback】"
