---
name: skip 策略不会清除模板预置文本
description: PPT 移植时 SHAPES spec 中的 skip 策略只是不写新值，clone 模板预置文字会残留；移植新模板必查项
type: feedback
originSessionId: f09ee6c8-e7f2-49bd-a8e1-4dbf22fd69c8
---
`SHAPES` spec 里 `strategy: "skip"` 的含义是"代码这一轮不写新内容"，**不等于清空 shape 文本**。Clone 模板（如 `apparel_ppt.py` 走 slide 19 → slide 20）时，模板上 shape 里的文字会原样带过来。

**Why:** 2026-04-28 apparel 移植发现 4 个装饰性 Oval 圈在 PPT 里残留 "4.9 / 5.0 / 5.0" 等分数 —— 根因是早期 dev 把它们设成了 `score_category_mean` 写分数；改回 `skip` 后旧分数仍在（因为之前那一轮已经写到模板/输出页里了）。用户连续两次说"shape 文字没清"才定位到是 `skip` 不清旧值的认知盲区。

**How to apply:**

1. **新模板移植时**：所有 `strategy: "skip"` 的 shape，去源模板（如 `Template 2.1.pptx` 末尾的 standard slide）确认那 shape 里**本来就是空的**。如果源模板里有预置文字而你想空 → 改源模板，不要靠代码每次清。
2. **想用代码强制清空**：要么显式 `sh.TextFrame.TextRange.Text = ""`；要么扩展 `_build_content` 加一个 `clear` strategy，返回 sentinel 让 `_write_text` 走清除分支。
3. **诊断 "shape 残留文字" 类 bug**：第一步问"这 shape 是 skip 还是 write 了？源模板那位置是不是空的？" —— 这两个问题能 80% 定位根因。

类比：yzr/zxh 模板里所有 `skip` shape 都已经是模板预置好的最终文字（分类标题"版型/面料"等），所以一直没暴露这个坑；apparel 装饰圈是新场景才触发。
