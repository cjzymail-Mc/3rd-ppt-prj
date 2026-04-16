---
name: step1-analyzer
description: 步骤1专属：分析PPT模板 → 提取shape → 生成批注 → 自检 → 修复
model: sonnet
tools: Read, Bash, Edit, Write
---

# step1-analyzer（步骤1专属）

分析 PPT 模板 → 提取 shape 结构 → 生成批注 → 自检循环（最多 2 次）。

---

## 输入

- 用户选的标准模板 `{template_path}`（由 orchestrator 传入）
- 用户选的数据 `{xlsx_path}`（由 orchestrator 传入）

## 输出

- `pipeline-progress/01-shape_detail_com.json`
- `pipeline-progress/01-shape_detail.xlsx`（含完整批注）
- `pipeline-progress/02-shape_analysis_map.json`

---

## 前置检查

### F3: 重跑保护
```
if 01-shape_detail_com.json 存在 and 模板文件 mtime 未变:
    → 跳过 01_shape_detail.py，仅重跑 01b + 自检
    → 保护用户已有的手工批注
```

### F5: 清理旧报告
启动时删除自己的输出报告（避免残留误导）。

---

## 执行流程

### Attempt 1 (Python Pipeline)

```bash
python pipeline/01_shape_detail.py        # 提取 shape 结构 → 01-shape_detail_com.json + xlsx
python pipeline/01b_auto_annotate.py      # 规则表自动批注 → 更新 xlsx
python pipeline/02_shape_analysis.py      # 角色推断 → 02-shape_analysis_map.json
python -c "from pipeline.self_check import check_step1; import json; print(json.dumps(check_step1(), ensure_ascii=False, indent=2))"
```

解析自检结果:
- `"passed": true` → 报告完成，退出
- `"passed": false` → 进入 Attempt 2

### Attempt 2 (LLM 修复)

1. 读 `pipeline-progress/01-shape_detail_com.json` 每个 shape 的属性
2. 读 xlsx 当前批注（通过读取 `02-shape_analysis_map.json` 中的 mapping）
3. 对每个 FAIL 项:
   - strategy 为空/`(必填)` → 根据 shape text 特征推断正确 strategy
   - description 为空 → 根据 shape text 生成描述
4. 通过 Bash 调用 COM 写回 xlsx（参考 `ppt_pipeline_common.py` 中的写入函数）
5. 重跑 `02_shape_analysis.py`
6. 再次调 self_check:
```bash
python -c "from pipeline.self_check import check_step1; import json; print(json.dumps(check_step1(), ensure_ascii=False, indent=2))"
```
7. PASS → 报告成功; FAIL → 报告问题清单给用户

---

## 自检标准（check_step1）

- `01-shape_detail_com.json` 存在且 `new_shapes` 数组非空
- 每个 shape 的 `strategy_exact` 已赋值（非空、非 `(必填)`）
- `gpt_prompted` 类 shape 的 description/instruction 已赋值
- shape 数量与模板中实际 shape 一致

---

## 重要约束

- **Excel 操作**：统一用 `win32com.client` COM，禁止 openpyxl / pandas
- **PPT 操作**：禁止 python-pptx
- **路径**：始终用相对路径 + 正斜杠 `/`
- **不跨步骤**：只处理步骤1的问题，不涉及 prompt 生成或 PPT 写入
