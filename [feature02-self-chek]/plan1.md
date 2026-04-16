# Feature 02: 自检机制移植方案

> 将 HTML→PPT 工作流中的 Builder 自检 + Converter 四步法，移植到当前 PPT Pipeline 工作流。

---

## 1. 现状问题

当前工作流：`03b 生成 PPT` → `orchestrator 串联` → `04 验收`

- 03b 写完就退出，不验证产物质量
- 04 是独立步骤，发现问题后需走完整迭代轮次才能修复
- 缺少视觉对照（模板 PPT vs 生成 PPT）
- 修复反馈环路太长：build → test → 下一轮 prompt 优化 → rebuild → retest

**目标**：将 build + check + fix 紧耦合在同一步骤内，交付前自行修复，减少迭代轮次。

---

## 2. 移植架构

```
当前流程:
  03b(build) ──→ orchestrator ──→ 04(test) ──→ 下一轮迭代

目标流程:
  03b(build) ──→ 03b(self-check 四步法) ──→ 03b(auto-fix) ──→ 03b(re-check)
       │                                                              │
       │              ← 循环直到无严重问题或达到 max_retry ←           │
       │                                                              │
       └──────────────────→ 交付 ──→ 04(终验，仅报告，不再修复)
```

04 从"发现问题+驱动修复"降级为"终验报告"——相当于 QA 的最终签字，不再承担修复职责。

---

## 3. 四步法详细设计

### 步骤 ① 属性校验（坐标/字体/颜色）

**基准**：`pipeline-progress/01-shape_detail_com.json`（模板 shape 属性）
**方法**：03b 写入 PPT 后，COM 回读每个 shape 的实际属性，与基准对比

| 检查项 | 判定标准 | 严重度 |
|--------|---------|--------|
| 位置 (Left/Top) | 偏差 > 3pt | 严重 |
| 尺寸 (Width/Height) | 偏差 > 5pt | 严重 |
| 字号 (Font.Size) | 不一致 | 中等 |
| 字体 (Font.Name) | 不一致 | 中等 |
| 颜色 (Font.Color) | 不一致 | 轻微 |
| 文本对齐 (Alignment) | 不一致 | 中等 |

**实现**：新增函数 `_readback_and_verify(ppt_path, baseline_json) → list[Issue]`

### 步骤 ② 可编辑性验证

**方法**：COM 遍历生成 PPT 的所有 shape，检查关键属性

| 检查项 | 方法 | 判定标准 |
|--------|------|---------|
| 文本框可编辑 | `shape.HasTextFrame` + `shape.TextFrame.TextRange.Text` 可读写 | 不可访问 = 严重 |
| 图片可选中 | `shape.Type == msoPicture(13)` | 图片丢失 = 严重 |
| 表格可编辑 | `shape.HasTable` + 单元格可读写 | 不可访问 = 严重 |
| shape 未锁定 | `shape.LockAspectRatio` 等属性可读 | 锁定 = 中等 |

**实现**：集成在 `_readback_and_verify()` 内

### 步骤 ③ 视觉对照（截图对比）

**基准**：标准模板 PPT 的目标页截图
**方法**：

```python
# 1. 导出模板页截图（只需执行一次，缓存到 pipeline-progress/）
template_slide.Export("pipeline-progress/03b-baseline_page.png", "PNG", 1920, 1080)

# 2. 导出生成页截图
generated_slide.Export("pipeline-progress/03b-generated_page.png", "PNG", 1920, 1080)

# 3. 像素级差异检测
#    方案 A：Pillow + numpy 计算 SSIM（轻量，无额外依赖）
#    方案 B：生成差异图供人工/LLM 审查
```

| 检查项 | 判定标准 | 严重度 |
|--------|---------|--------|
| 整体相似度 (SSIM) | < 0.85 | 严重 |
| 局部差异区域面积 | > 页面 10% | 严重 |
| 文字区域偏移 | 可见漂移 | 中等 |

**输出物**：
- `pipeline-progress/03b-visual_diff.png`（差异热力图）
- 差异区域坐标列表

**实现**：新增函数 `_visual_compare(template_slide, generated_slide) → list[Issue]`

> 注意：`slide.Export()` 需要 PPT 在前台打开才能截图。如果 COM 在后台模式运行时截图失效，降级为仅属性校验（步骤①②④），视觉对照标记为 SKIPPED。

### 步骤 ④ 内容完整性

**基准**：`pipeline-progress/03a-build_shape_content.json`（预期写入内容）+ Excel 数据源
**方法**：COM 回读每个 shape 的实际文本，与预期内容对比

| 检查项 | 判定标准 | 严重度 |
|--------|---------|--------|
| 文本缺失 | shape 应有文本但为空 | 严重 |
| 关键词缺失 | required_keywords 未出现 | 严重 |
| 文本截断 | 实际文本长度 < 预期的 80% | 中等 |
| 多余文本 | shape 出现非预期内容 | 轻微 |

**实现**：新增函数 `_content_completeness_check(ppt_path, content_json) → list[Issue]`

---

## 4. 自修复循环

```python
MAX_SELF_FIX_RETRIES = 2

for attempt in range(MAX_SELF_FIX_RETRIES + 1):
    # build PPT (首次) 或 fix PPT (重试)
    if attempt == 0:
        build_ppt(content_json, template_path, out_path)
    else:
        apply_fixes(out_path, issues)

    # 四步法自检
    issues = []
    issues += readback_and_verify(out_path, baseline_json)    # ①②
    issues += visual_compare(template_slide, generated_slide)  # ③
    issues += content_completeness_check(out_path, content_json)  # ④

    severe = [i for i in issues if i.severity == "严重"]
    if not severe:
        break  # 通过，交付

    if attempt == MAX_SELF_FIX_RETRIES:
        # 达到重试上限，生成报告，交给 04 终验
        break

# 输出自检报告
generate_self_check_report(issues, out_path="pipeline-progress/03b-self_check_report.md")
```

**可自动修复的问题类型**：
| 问题 | 修复方式 |
|------|---------|
| 文本缺失/截断 | 重新 COM 写入对应 shape |
| 字体/字号不一致 | COM 重设 Font 属性 |
| 关键词缺失 | 重新写入包含关键词的文本 |

**不可自动修复的问题类型**（报告给 04/用户）：
| 问题 | 原因 |
|------|------|
| 坐标严重偏移 | 可能是 clone 页面的结构性问题 |
| shape 不可编辑 | 可能是模板保护或 COM 异常 |
| 视觉大面积差异 | 可能需要调整 content 策略 |

---

## 5. 自检报告格式

输出到 `pipeline-progress/03b-self_check_report.md`：

```markdown
# 03b Self-Check Report

| # | Shape | 维度 | 问题描述 | 严重度 | 状态 |
|---|-------|------|---------|--------|------|
| 1 | Slide1.Shape3 | 内容完整性 | 文本缺失："建议" | 严重 | 已修复(attempt 2) |
| 2 | Slide1.Shape7 | 属性校验 | 字号 12pt→10pt | 中等 | 已修复(attempt 2) |
| 3 | Slide1 | 视觉对照 | SSIM=0.92 通过 | — | PASS |

## 总结
- 自检轮次：2
- 发现问题：2 严重 / 0 中等
- 已修复：2 / 未修复：0
- 视觉相似度：SSIM 0.92
- 结论：✅ 通过，交付终验
```

---

## 6. 修改文件清单

| 文件 | 改动 | 工作量 |
|------|------|--------|
| `pipeline/03b_build_ppt_com.py` | 新增四步法自检 + 自修复循环 + 报告生成 | 主要 |
| `pipeline/ppt_pipeline_common.py` | 新增 `slide_export_png()` 截图工具函数 + `ssim_compare()` 图像对比 | 中等 |
| `pipeline/04_shape_diff_test.py` | 降级为终验报告（逻辑不变，角色定位调整） | 微调 |
| `orchestrator.py` | 读取 `03b-self_check_report.md` 判断是否需要进入终验 | 微调 |
| `.claude/agents/02-builder.md` | 更新角色定义，加入自检职责说明 | 文档 |
| `.claude/agents/03-reviewer.md` | 更新角色定义，明确"终验"定位 | 文档 |

**不改的**：01/01b/02/02b/03a — 上游步骤不受影响

---

## 7. 依赖与风险

| 项目 | 说明 | 降级方案 |
|------|------|---------|
| `slide.Export()` 截图 | 需要 PowerPoint 进程可见窗口 | 后台模式时跳过步骤③，标记 SKIPPED |
| Pillow/numpy（SSIM） | 需额外 pip 安装 | 降级为仅输出截图，人工/LLM 对比 |
| 自修复写入冲突 | COM 写入可能因 PPT 未关闭而失败 | 重试前确保 close+reopen |

---

## 8. 验证方式

1. **单元测试**：用现有模板 + 已知数据，跑 03b → 检查报告是否正确识别差异
2. **人工对照**：比对自检报告 vs 人眼发现的问题，评估覆盖率
3. **回归测试**：确认 04 终验结果与 03b 自检结果一致（不应出现 03b 漏检的严重问题）
4. **编译检查**：`python -m py_compile pipeline/03b_build_ppt_com.py`
