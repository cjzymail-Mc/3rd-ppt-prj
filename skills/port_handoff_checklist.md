---
name: Pipeline → Developer 移植衔接 Checklist
description: 跑完 Pipeline 后，调用 /developer 做移植时的产物消费清单
type: feedback
---

跑完 Pipeline（部分或全部）后，准备做 src/{template}_ppt.py 移植时使用本 Checklist。

**目标**：避免每次移植都重新摸索"哪些 Pipeline 产物有用、字段怎么映射"。

---

## 1. Pipeline 产物清单（按阶段）

跑完不同阶段后会产出以下文件，全部位于 `pipeline-progress/`：

### Step 1 产物（必读）

| 文件 | 大小级别 | 用途 |
|--|--|--|
| `01-shape_detail.xlsx` | ~30KB | **主产物**：shape 清单 + 用户填的"内容描述"标注列 |
| `01-shape_detail_com.json` | ~5KB | 同上 JSON 版本，便于程序读取 |
| `01-shape_fingerprint_map.json` | ~2KB | shape 指纹（跨轮次匹配用） |

### Step 2 产物（移植主要消费对象）

| 文件 | 大小级别 | 用途 |
|--|--|--|
| `02-prompt_specs.json` | ~15KB | **每 shape 最终 prompt** —— 移植时直接提取 |
| `02-shape_analysis_map.json` | ~17KB | **每 shape strategy 推断** —— 决定 SHAPES 列表的 strategy 字段 |
| `02-readability_budget.json` | ~1.5KB | **字数/行数预算** —— 写入 SHAPES 列表的 budget 字段 |

### Step 3 产物（参考用）

| 文件 | 用途 |
|--|--|
| `03a-build_shape_content.json` | GPT 生成的实际内容样本（验证 prompt 效果） |
| `03b-post_write_readback.json` | PPT 写入后回读（确认写入成功） |

### Step 4 产物（健康度判断）

| 文件 | 用途 |
|--|--|
| `04-fix_ppt.md` | 自检报告：visual / readability / semantic 三项分数 + 修正建议 |

---

## 2. 字段映射规范（02-prompt_specs.json → src/{template}_ppt.py）

```python
# Pipeline JSON 字段 → src/ 代码位置
{
  "shape_name": "Rectangle 68",        →  SHAPES 列表的 "name" 字段
  "role": "advantage",                 →  分支判断 / 或写到 prompt 的角色提示
  "model": "openai/gpt-5-mini",        →  _MODEL 常量（或 spec["model"]）
  "instruction": "...",                →  _build_rich_prompt() 的核心 instruction
  "output_constraints": {              →  SHAPES 列表的 "budget" 字段
    "max_chars": 270,                  →  budget["max_chars"]
    "max_lines": 9,                    →  budget["max_lines"]
    "no_markdown": true                →  prompt 里加"禁 markdown"约束
  },
  "context_headers": [...],            →  Excel 列名清单（用于 _classify_columns）
  "user_content_source": "...",        →  spec["params"]["source"]
  "user_instruction": "..."            →  prompt 拼接到 instruction 末尾
}
```

---

## 3. 移植 Checklist（按顺序勾选）

### 输入准备
- [ ] 模板 `.pptx` 文件（`template/{template}.pptx`）
- [ ] 配套数据 `.xlsx`
- [ ] Pipeline 产物（至少 Step 1，推荐到 Step 2）
- [ ] 视觉效果初步认可（视觉 ≥ 80% 或可接受）

### 创建新模板文件
- [ ] 复制 `src/yzr_ppt.py` 为 `src/{template}_ppt.py`
- [ ] 修改 `_TEMPLATE_SLIDE` 为模板的实际页码
- [ ] 修改 `{NAME}_SHAPES` 列表（替换 shape 名 + strategy + budget）

### 同步 Pipeline 产物
- [ ] 读 `02-shape_analysis_map.json` → 确定每 shape 的 `strategy_exact` → 写入 SHAPES 列表
- [ ] 读 `02-prompt_specs.json` → 提取 `instruction` / `user_instruction` → 写入 `_build_rich_prompt()`
- [ ] 读 `02-readability_budget.json` → 写入 SHAPES 列表的 `budget`
- [ ] 添加 prompt 追溯注释（fix2 范式）：

```python
# prompt_src:  pipeline/prompt_templates/gpt_summary.md
# synced_at:   YYYY-MM-DD  ← 同步当天日期
# synced_by:   Developer
def _build_rich_prompt(...):
    ...
```

### 处理特殊组件
- [ ] **图表分支决策**（按 fix4 路线）：
  - 单机自用 + 模板预置 chart shape → `_write_chart()` 原位
  - 分发场景（模板/代码发给同事，他人填数据）→ `make_chart_for_{name}()` 从零制表 + OLE 粘贴
  - 详见 `.claude/memory/feedback_chart_write.md`
- [ ] **图片处理**：用 `_replace_image()`，按模板 shape 的 L/T/W/H 等比缩放居中
- [ ] **shape 微调**：参考 `skills/fine-tuned-shapes.md` 在 `make_{name}_slide()` 头部插入 `#fine_tuned` 块

### 接入 Main.py
- [ ] `Main.py::ask_template_choice()` 增加新选项
- [ ] import `make_{name}_slide` + 在分支里调用

### 验证
- [ ] `python -c "import ast; ast.parse(open('src/{template}_ppt.py').read())"` 语法通过
- [ ] `python src/{template}_ppt.py` 单页调试跑通（需先打开 Excel）
- [ ] 在 `src/{template}_ppt.py` `__main__` 块加单页 smoke
- [ ] 跑 `python src/{template}_ppt.py` 全通（需 Excel + PPT 打开）

---

## 4. 仅跑了 Step 1 的降级方案

如果用户只跑了 Step 1（评估后觉得不需要继续 Pipeline 迭代），手头没有 02-*.json，按以下方式处理：

| 缺失产物 | 替代方案 |
|--|--|
| `02-shape_analysis_map.json`（strategy） | 看 01-shape_detail.xlsx 的"内容描述"列 + Developer 经验判断 |
| `02-prompt_specs.json`（prompt） | **复制 yzr_ppt.py 现成 `_build_rich_prompt()` 改造**（最快路径） |
| `02-readability_budget.json`（budget） | 用 yzr/zxh 的 budget 作为初值，运行后视效果调整 |

---

## 5. 不要做的事

- ❌ 跳过 prompt 追溯注释（未来 Pipeline 升级时无法追溯哪些模板需要 sync）
- ❌ 把 `02-*.json` 的内容硬编码进 Python 字符串（保留 JSON 形态）
- ❌ 在 src/ 里重新做 shape 角色判断（Step 2 已经做完）
- ❌ 直接 import `pipeline/03b_build_ppt_com.py` 的代码（Pipeline 假设是单机场景，分发场景要适配性评审）

---

## 6. 参考档案

- `.claude/agents/developer.md` —— `/developer` Agent 完整角色定义 + 移植 Checklist
- `plan3（工作流5阶段定稿）.md` —— 5 阶段工作流定稿
- `.claude/memory/feedback_workflow_routing.md` —— 工作流路由记忆
- `.claude/memory/feedback_chart_write.md` —— chart 路线决策（分发 → make_chart_for_{name}）
- `[feature03-transplant]/fix2（三重混合架构整改）.md` —— Pipeline → src/ 衔接的同步机制
- `【trash-bin】/pipeline-progress-yzr/` —— yzr 早期完整 Pipeline 跑批的产物（参考样本）
