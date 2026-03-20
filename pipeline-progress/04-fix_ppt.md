# PPT Diff & Fix Report

- 状态: fail
- visual_score: 100.00%
- readability_score: 98.96%
- semantic_coverage: 66.67%
- template_shapes: 26
- target_shapes: 26
- paired: 26
- 时间: 2026-03-20T17:17:38

## Shape对比
|template|target|visual|readability|text_len|match|
|---|---|---|---|---|---|
|Straight Connector 1|Straight Connector 1|100.00|100.00|0/0|name|
|Straight Connector 2|Straight Connector 2|100.00|100.00|0/0|name|
|Straight Connector 3|Straight Connector 3|100.00|100.00|0/0|name|
|Rectangle 4|Rectangle 4|100.00|100.00|0/0|name|
|Straight Connector 5|Straight Connector 5|100.00|100.00|0/0|name|
|Straight Connector 6|Straight Connector 6|100.00|100.00|0/0|name|
|Straight Connector 7|Straight Connector 7|100.00|100.00|0/0|name|
|Straight Connector 8|Straight Connector 8|100.00|100.00|0/0|name|
|Straight Connector 9|Straight Connector 9|100.00|100.00|0/0|name|
|Straight Connector 10|Straight Connector 10|100.00|100.00|0/0|name|
|Rectangle 11|Rectangle 11|100.00|100.00|7/7|name|
|Rectangle 12|Rectangle 12|100.00|100.00|1/2|name|
|Straight Connector 15|Straight Connector 15|100.00|100.00|0/0|name|
|Rectangle 17|Rectangle 17|100.00|100.00|34/40|name|
|Rectangle 19|Rectangle 19|100.00|100.00|0/0|name|
|Straight Connector 32|Straight Connector 32|100.00|100.00|0/0|name|
|Picture 39|Picture 39|100.00|100.00|0/0|name|
|Rectangle 40|Rectangle 40|100.00|100.00|0/0|name|
|TextBox 16|TextBox 16|100.00|73.08|13/4|name|
|Straight Connector 41|Straight Connector 41|100.00|100.00|0/0|name|
|Rectangle 68|Rectangle 68|100.00|100.00|225/254|name|
|Picture 74|Picture 74|100.00|100.00|0/0|name|
|Picture 75|Picture 75|100.00|100.00|0/0|name|
|Rectangle 77|Rectangle 77|100.00|100.00|168/187|name|
|Rectangle 14|Rectangle 14|100.00|100.00|11/11|name|
|图表 44|图表 44|100.00|100.00|0/0|name|

## 修正建议

| shape | 问题 | fix_type | 建议 |
|-------|------|----------|------|
| TextBox 16 | readability=73.08 < 95 (文本长度/行数偏差) | budget_overflow | 调整 readability_budget 或 prompt 字数约束 |
| (全局) | 语义关键词缺失: 建议 | keyword_missing | 在 gpt_prompted 类 shape 的内容描述中追加：必须包含'建议'关键词 |
