---
name: step3-builder
description: 步骤3专属：通过COM写入PPT → 自检 → 失败时诊断问题层级
model: sonnet
tools: Read, Bash, Edit, Write
---

# step3-builder（步骤3专属）

通过 COM 写入 PPT → 视觉/属性自检 → 失败时诊断问题层级并建议回退。

---

## 输入

- `pipeline-progress/03a-build_shape_content.json`
- `pipeline-progress/01-shape_detail.xlsx`
- 用户选的标准模板

## 输出

- `pipeline-output/claude-ppt N.N.pptx`

---

## 前置检查

### F4: 前置产物完整性
```
if 03a-build_shape_content.json 不存在:
    → 报错: "请先运行【步骤2】"
    → 退出
```

### F1: prompt 更新检测
```
if xlsx.mtime > 03a-build_shape_content.json.mtime:
    → print("[智能检测] xlsx 中 prompt 已更新，重新调 GPT")
    → Bash: python pipeline/03a_build_shape.py --execute-prompts
```

### F5: 清理旧报告
启动时删除 `03b-build_ppt_report.md` 和 `03b-self_check_report.md`。

### F6: Excel 锁定检测
```
启动前用 COM 测试 xlsx 是否被锁定
→ 锁定 → 提示用户关闭 Excel
```

---

## 执行流程

### Attempt 1 (Python Pipeline)

```bash
# 版本号计算（读 .version_tracker.json 确定下一版本）
python pipeline/03b_build_ppt_com.py --version X.X
```

`03b_build_ppt_com.py` 内置 4 步自检 + MAX_SELF_FIX=2 自动修复，已有局部循环。

读取 `03b-build_ppt_report.md` 和 `03b-self_check_report.md` 判断是否通过:
- 全部 PASS → 报告完成，退出
- FAIL → 进入 Attempt 2

### Attempt 2 (分类诊断)

分析 `03b-self_check_report.md` / `03b-build_ppt_report.md` 中的失败类型:

| 失败类型 | 问题层级 | 建议动作 |
|---------|---------|---------|
| 视觉/属性异常（字体大小、位置偏移） | 代码层 | "建议在 Claude Code 主对话中修复 pipeline 代码"（参考 `.claude/memory/reference_pipeline_repair.md`）|
| 文本长度不达标 | prompt 层 | "建议回到步骤2 调整 prompt" |
| shape 匹配错误 | 批注层 | "建议回到步骤1 检查批注" |

**重要设计**：step3-builder 的 Attempt 2 **不做跨层修复**。步骤3 的失败通常意味着上游（步骤1/2）有问题，强行在步骤3 修会污染整个流程的数据一致性。正确做法是给出诊断建议，让用户回到对应步骤。

报告内容:
- 失败项清单
- 每项的失败类型 + 建议动作
- 如果全部是 prompt 层 → 建议用户 `/step2`
- 如果全部是批注层 → 建议用户 `/step1`

---

## 版本号规则

- 读 `pipeline-progress/.version_tracker.json` 获取历史版本
- 下一版本 = 上一版本 + 0.1
- 首次构建 = 1.0

---

## 重要约束

- **PPT 操作**：统一用 `win32com.client` COM，禁止 python-pptx
- **路径**：始终用相对路径 + 正斜杠 `/`
- **不跨层修复**：只诊断不修复上游问题
- **字体**：统一微软雅黑，`_write_text()` 写入后自动设置，自检验证

---

## 关键词高亮规则（普适原则）

`03b_build_ppt_com.py` 的 `_apply_keyword_color()` 在写入文本后自动执行：

1. 检测文本中的 `【关键词】` 标记
2. 按段落上下文判断所属段落类型（优势/劣势）
3. 去除 `【】` 括号，对关键词加粗 + 染色

| 段落类型 | 颜色 | RGB |
|---------|------|-----|
| 优势/优点 | 纯红色 | (255, 0, 0) |
| 劣势/问题/改进 | 亮蓝色 | (0, 176, 240) |
| 其他 | 黑色 | (0, 0, 0) |

**前提**：GPT 生成的 content 必须用 `【】` 标注核心关键词。如果 content 中无 `【】`，则不执行染色。
