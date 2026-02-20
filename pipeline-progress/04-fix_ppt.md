# PPT Diff & Fix Report

- 状态: fail
- visual_score: 98.06%
- readability_score: 92.44%
- semantic_coverage: 33.33%
- template_shapes: 26
- target_shapes: 26
- paired: 26
- 时间: 2026-02-19T13:25:23

## Shape对比
|template|target|visual|readability|text_len|match|
|---|---|---|---|---|---|
|直接连接符 1|直接连接符 1|100.00|100.00|0/0|name|
|直接连接符 2|直接连接符 2|100.00|100.00|0/0|name|
|直接连接符 3|直接连接符 3|100.00|100.00|0/0|name|
|矩形 4|矩形 4|100.00|100.00|0/0|name|
|直接连接符 5|直接连接符 5|100.00|100.00|0/0|name|
|直接连接符 6|直接连接符 6|100.00|100.00|0/0|name|
|直接连接符 7|直接连接符 7|100.00|100.00|0/0|name|
|直接连接符 8|直接连接符 8|100.00|100.00|0/0|name|
|直接连接符 9|直接连接符 9|100.00|100.00|0/0|name|
|直接连接符 10|直接连接符 10|100.00|100.00|0/0|name|
|矩形 11|矩形 11|100.00|82.86|7/7|name|
|矩形 12|矩形 12|84.62|67.50|1/2|name|
|直接连接符 15|直接连接符 15|100.00|100.00|0/0|name|
|矩形 17|矩形 17|100.00|84.90|34/40|name|
|矩形 19|矩形 19|95.80|100.00|0/0|name|
|直接连接符 32|直接连接符 32|100.00|100.00|0/0|name|
|图片 39|图片 39|100.00|100.00|0/0|name|
|矩形 40|矩形 40|100.00|100.00|0/0|name|
|文本框 16|文本框 16|100.00|91.15|13/11|name|
|直接连接符 41|直接连接符 41|100.00|100.00|0/0|name|
|图表 44|图表 44|100.00|100.00|0/0|name|
|矩形 68|矩形 68|84.62|42.97|225/130|name|
|图片 74|图片 74|100.00|100.00|0/0|name|
|图片 75|图片 75|100.00|100.00|0/0|name|
|矩形 77|矩形 77|84.62|34.16|168/80|name|
|矩形 14|矩形 14|100.00|100.00|11/11|name|

## 差异与修复建议
- 矩形 11 -> 矩形 11: visual=100.0, readability=82.86 -> 调整03a_build_shape.py对应shape的prompt/预算
- 矩形 12 -> 矩形 12: visual=84.62, readability=67.5 -> 调整03a_build_shape.py对应shape的prompt/预算
- 矩形 17 -> 矩形 17: visual=100.0, readability=84.9 -> 调整03a_build_shape.py对应shape的prompt/预算
- 矩形 19 -> 矩形 19: visual=95.8, readability=100.0 -> 调整03a_build_shape.py对应shape的prompt/预算
- 文本框 16 -> 文本框 16: visual=100.0, readability=91.15 -> 调整03a_build_shape.py对应shape的prompt/预算
- 矩形 68 -> 矩形 68: visual=84.62, readability=42.97 -> 调整03a_build_shape.py对应shape的prompt/预算
- 矩形 77 -> 矩形 77: visual=84.62, readability=34.16 -> 调整03a_build_shape.py对应shape的prompt/预算
