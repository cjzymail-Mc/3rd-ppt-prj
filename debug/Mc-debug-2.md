

-------------------
首先，考虑优化 agent
-------------------


我的工作要求ppt的制作精度非常高，因此你必须将ppt当成一个架构精密的软件系统来对待
我已构建了软件工程通用的6个agents，位于 .claude/agents 目录中
我还构建了专业的agents控制器，位于根目录的 orchestrator.py

接下来我希望将它们改造成专业ppt工作团队，按照软件工程的逻辑，来生成ppt
即，用严谨的开发流程来生成ppt【规划、构建代码、测试、反馈、修改...（重复多轮后）→ 交付】

我原有的ppt生成脚本由【根目录main.py + /src】脚本程序构成，
我的脚本解析结果在【codex-legacy】的repo-sanc-result.md中
我原有的脚本虽然效率不高，但能够实现ppt的精准控制，你必须严格遵循我的原有思路，来进行接下来的工作
不要小看这份工作，我用codex构建目标ppt，结果已经连续失败许多次，你必须全力以赴
codex的工作记录位于 【codex-legacy】文件夹中，你仔细阅读，然后再继续工作
帮我完成改写 【6个agents 的md文件】  和 【orchestrator.py】
将它们打造成专业的 ppt 软件工程师团队，来满足我对ppt的严苛要求





---------- 02-17 night Claude -----------


完成。总结一下你的三个问题：                                                  
  1. repo-scan-result.md — 已小幅更新

  原文件的项目结构部分过时（引用 030 PPT Robot...v3.0.py），已更新为完整的当前目  录结构，包含orchestrator、6个step脚本、agent配置等。核心技术分析部分（COM自动
  化、xlwings混合模式、核心函数/类）无需改动——main.py和src/没有变。

  2. new-ppt-workflow.md — 不需要更新

  它是 v4.0 执行规范，是上游源头。我们所有改动（agent
  MD、orchestrator.py）都是下游对齐它。改下游不需要动上游。

  3. 使用指南 — 已创建 USAGE_GUIDE.md

  涵盖以下内容：

  ┌──────────────┬─────────────────────────────────────────────────────────────┐
  │     章节     │                            内容                             │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 系统概览     │ 6个Agent协作流程图                                          │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 前置要求     │ Windows+Office+Python+依赖                                  │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 5种使用方式  │ 全自动 / 手动@agent / 从PLAN恢复 / 多轮迭代 / 自动Architect │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 执行流程详解 │ Phase 1→2→3 的每步细节                                      │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 关键文件说明 │ 输入/配置/输出三类文件的完整清单                            │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 迭代机制     │ Round循环 + 修复优先级路由                                  │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 常见操作速查 │ 单Agent运行、失败恢复、调试信息、手动步骤脚本               │
  ├──────────────┼─────────────────────────────────────────────────────────────┤
  │ 注意事项     │ 5个关键注意点                                               │
  └──────────────┴─────────────────────────────────────────────────────────────┘










------------------------------------------
接下来，考虑优化py代码 / 讨论 py 实现的可能性
------------------------------------------

好的，你已经完成 agents md 文件的改写和 orchestrator.py 的升级

接下来考虑用纯 python 代码 实现上述工作流程的可能性

我之前使用codex 尝试过这个可能性，效果很差。

之前的代码 / debug 信息都放在 【codex-legacy2】文件夹中

注意，我的源代码中包含一个 Function_030.py文件，里面的 GPT_5 函数可以实现 GPT_5 模型的调用

前提是你需要自己构建合适的 prompt 提示词。你也可以在 Function_030.py 中看到我构建的 mc_prompt 历史

注意事项：
    我之前的脚本代码【根目录main.py + /src】，里面的所有工具你都可以调用，不需要重新造轮子
    例如，你可以动态加载 src/Function_030.py 里的 GPT_5 与 extract_info，用于 AI 总结与复杂字段提取（并带兜底逻辑）

我就是通过这个方式，实现从 原始问卷中 ，提取关键信息的

你先评估下，用纯 python 代码 实现 agents + new-ppt-workflow.md 工作流程的可能性










----------------------------------------

Feasibility Assessment: Pure Python Implementation of PPT Workflow                                                                                              Based on thorough analysis of codex-legacy2/ failure history, Function_030.py   (37 functions), main.py execution flow, and Class_030.py shape hierarchy,       here's my assessment:                                                                                                                                           ---                                                                             Verdict: 高度可行，且推荐这么做                                                                                                                                 ---                                                                             1. codex-legacy2 为什么失败？                                                                                                                                   不是代码架构有问题，而是环境问题：所有脚本在 import win32com.client             时直接崩溃，因为 Codex 的运行环境没有安装 pywin32 + Microsoft Office。

  [BLOCKED] win32com unavailable: No module named 'win32com'


  实际上 codex-legacy2 的 5 步脚本架构设计是合理的，策略矩阵、三层测试、版本迭代
  逻辑都写好了，只是完全没跑起来。

  ---
  2. 现有可复用资源清单

  类别: GPT 调用
  可复用工具: GPT_5(prompt, model)
  来自: Function_030.py
  ────────────────────────────────────────
  类别: Prompt 构建
  可复用工具: gen_mc_prompt(), gen_result_prompt(), gen_questionnaire_prompt(),
    build_questionnaire_summary_prompt()
  来自: Function_030.py
  ────────────────────────────────────────
  类别: 数据提取
  可复用工具: extract_info(), parse_survey_data()
  来自: Function_030.py
  ────────────────────────────────────────
  类别: Excel 搜索
  可复用工具: search(sheet, target, row_offset, col_offset)
  来自: Function_030.py
  ────────────────────────────────────────
  类别: 图表生成
  可复用工具: make_chart_for_questionnaire(), make_chart()
  来自: Function_030.py
  ────────────────────────────────────────
  类别: 文本格式化
  可复用工具: color_key(), smart_color_text(), adj()
  来自: Function_030.py
  ────────────────────────────────────────
  类别: 图片粘贴
  可复用工具: mc_pic()
  来自: Function_030.py
  ────────────────────────────────────────
  类别: Shape 类
  可复用工具: Text_Box, Title_Max, Title_1, Title_2, Result_Bullet 等13个类
  来自: Class_030.py
  ────────────────────────────────────────
  类别: PPT 保存
  可复用工具: ppt_save()
  来自: Function_030.py
  ────────────────────────────────────────
  类别: COM 包装
  可复用工具: 剪贴板延时、Slide.Copy/Paste 模式
  来自: main.py

  结论：造轮子的工作量接近零。所有 COM 操作、GPT
  调用、数据提取、格式化工具都已存在。

  ---
  3. 五步脚本与现有能力的对应关系

  Step: Step 1
  脚本: 01-shape-detail.py
  核心能力: COM 读取 shape 属性（name/type/position/font/text）
  已有?: main.py 已有遍历 shape 模式
  ────────────────────────────────────────
  Step: Step 2
  脚本: 02-shape-analysis.py
  核心能力: shape 角色推断 + prompt 规格生成
  已有?: 需要新写，但纯 Python 逻辑
  ────────────────────────────────────────
  Step: Step 3A
  脚本: 03-build_shape.py
  核心能力: 按策略矩阵生成内容
  已有?: GPT_5() + extract_info() + parse_survey_data() 全部可复用
  ────────────────────────────────────────
  Step: Step 3B
  脚本: 03-build_ppt_com.py
  核心能力: 克隆模板 + COM 写入
  已有?: main.py 的 Slide.Copy/Paste + TextRange.Text 写入模式直接照搬
  ────────────────────────────────────────
  Step: Step 4
  脚本: 04-shape_diff_test.py
  核心能力: 三层对比测试
  已有?: 需要新写，但属性读取 COM 模式已有

  需要从零写的部分：
  - Step 2 的角色推断启发式逻辑（通过 shape name/text 判断 role）
  - Step 4 的三层评分算法（几何对比、文本相似度、语义覆盖）
  - 00-ppt.py 版本迭代控制器

  需要适配的部分：
  - Step 3A 需要为 long_summary 和 insight 构建新 prompt（可参考现有的
  gen_mc_prompt() 风格）
  - Step 3B 需要按 shape name 精确定位写入（现有代码按 shape 类型创建，需改为按
  name 查找）

  ---
  4. 纯 Python vs Agent 方案的优劣

  ┌──────────────┬──────────────────────────┬──────────────────────────────────┐
  │     维度     │        纯 Python         │            Agent 方案            │
  ├──────────────┼──────────────────────────┼──────────────────────────────────┤
  │ 确定性       │ 高 — 同输入同输出        │ 低 — LLM 每次可能生成不同代码    │
  ├──────────────┼──────────────────────────┼──────────────────────────────────┤
  │ 调试         │ 标准 Python debug        │ 需要看 agent 日志                │
  ├──────────────┼──────────────────────────┼──────────────────────────────────┤
  │ 速度         │ 快 — 无 agent 启动开销   │ 慢 — 每步需要 Claude 思考+写代码 │
  ├──────────────┼──────────────────────────┼──────────────────────────────────┤
  │ 可靠性       │ 高 — 写好就能反复用      │ 低 — agent 可能遗漏策略矩阵约束  │
  ├──────────────┼──────────────────────────┼──────────────────────────────────┤
  │ 灵活性       │ 低 — 新模板需改代码      │ 高 — agent 可适应新需求          │
  ├──────────────┼──────────────────────────┼──────────────────────────────────┤
  │ 首次开发成本 │ 中 — 需要手动实现5个脚本 │ 低 — agent 自动生成              │
  └──────────────┴──────────────────────────┴──────────────────────────────────┘

  关键判断：你的需求是固定模板 + 固定数据源 → 高保真
  PPT，确定性和可靠性远比灵活性重要。纯 Python 是正确选择。

  ---
  5. 风险点与应对

  风险: COM 不稳定（剪贴板占用）
  影响: Step 3B 写入失败
  应对: 已有 time.sleep(random.random()*delay) 模式，加重试机制
  ────────────────────────────────────────
  风险: shape name 不稳定
  影响: Step 1/3B 定位失败
  应对: 用指纹（type+position+text_prefix）做备用匹配
  ────────────────────────────────────────
  风险: GPT 输出超长/格式不对
  影响: Step 3A 内容不可用
  应对: 加 readability_budget 硬截断 + 格式校验
  ────────────────────────────────────────
  风险: 图表数据源映射
  影响: Step 3A chart 生成
  应对: 复用 parse_survey_data() + make_chart_for_questionnaire()

  ---
  6. 建议的实施路径

  Phase 1: 基础脚本（纯确定性路径先行）
    01-shape-detail.py    → COM 读取 shape 属性
    02-shape-analysis.py  → 角色推断 + prompt 规格
    03-build_shape.py     → 先只做 title/sample_stat/chart（无 GPT）
    03-build_ppt_com.py   → 克隆模板 + 写入确定性内容
    04-shape_diff_test.py → 三层评分算法

  Phase 2: GPT 路径
    03-build_shape.py     → 加入 long_summary/insight 的 GPT 生成
    Prompt 构建           → 参考 gen_mc_prompt() 风格

  Phase 3: 编排与迭代
    00-ppt.py             → 版本管理 + 迭代循环
    fix-ppt.md 解析       → 修复路由自动化


  你觉得这个方向可以推进吗？如果确认，我可以进入计划模式，详细设计每个脚本的实现
  方案。












 你选了单页问卷总结。第二个问题需要我补充说明吗？或者你有额外的背景想分享？  
 > 你先将它设计成独立的接口，待调试稳定后，我希望将它融入main的流程中            
● 明白了。设计目标很清晰：                                                      
  1. 单页：问卷总结页（Template page 15 风格）                                    
  2. 独立接口：先作为独立脚本，接口设计要干净
  3. 未来融入 main.py：稳定后可无缝集成
  4. 一次精确产出：没有迭代循环，跟 main.py 一样确定性执行

  这比之前的5步流水线简单得多。让我重新设计方案。














我系统性整理了下我的总体需求：  
1、我需要6个agents来帮我完成py代码的构建，并通过严苛agents循环检验，让生成的py代码能够精确复制模板ppt；注意，我所谓的【复制ppt模板】，指的是格式完全复制、风格统一，但内容是全新构建的，ppt内容始终源于【Excel源数据】  
2、然后我会使用6个agents协作产生的py代码，来帮我制作ppt // 如果制作的ppt符合我的要求，我会考虑将多个py文件整合到原有的【main.py + src文件夹的function等py文件中】，但在ppt通过用户（我）的最终评判之前，目前维持多个py文件状态没问题吧？ 而且也可以保留多个md/json文件输出，来方便调试。
3、如果这个总的工作流程顺利，我可以提供更多的模板，让agents这个工作流程批量产生更多的代码，实现多模板ppt的制作
4、目前仅需专注搞定这一套模板，【Template 2.1.pptx】 已包含 【空白 ppt 模板】和【标准 ppt 模板】，分别位于【Template 2.1.pptx】文件 的第14页 和 15页

根据我的需求，你综合评估一个解决方案给我



















Plan: 6-Agent 协作构建 PPT 生成脚本 + 迭代验证

 Context

 用户需求总结

 1. 6个 Agents 构建 Python 代码：通过 architect→tech_lead→developer↔tester
 循环，让生成的 py 代码能精确复制模板 PPT 的格式/风格（内容来自 Excel 源数据）
 2. 代码生成 PPT：稳定后，用产出的 py 代码一次性精确生成 PPT
 3. 多文件开发阶段：保留多个 py + JSON/MD
 文件便于调试；通过用户最终评判后再考虑整合到 main.py + src/
 4. 未来可扩展：此工作流稳定后，可套用到更多模板

 核心洞察

 - 迭代在 Agent 层：Agents 修改 Python 代码 → 代码生成 PPT → Tester 检查 →
 不合格则 Developer 改代码
 - Python 代码是确定性的：写好后一次运行精确产出，跟 main.py 一样
 - diff_result.json + fix-ppt.md：是 Agent 迭代的依据，Tester 产出给 Developer
 看

 "复制模板"的含义

 - 格式100%保留（字体、颜色、位置、样式、chart类型）
 - 内容全新构建（来自 Excel 源数据 2025 数据 v2.2.xlsx）
 - 模板 = src/Template 2.1.pptx 第14页(空白) + 第15页(标准)

 ---
 整体架构

 ┌─────────────────────────────────────────────────────┐
 │                 orchestrator.py                       │
 │  @arch → @tech → @dev ↔ @test (循环) → @opti → @sec │
 └───────────┬─────────────────────────────┬────────────┘
             │ Agents 写/改代码             │ Agents 读测试结果
             ▼                             ▼
 ┌─────────────────────────────────────────────────────┐
 │              pipeline/ (5个步骤脚本)                   │
 │                                                       │
 │  01_shape_detail.py ──→ shape_detail_com.json         │
 │         ↓                                             │
 │  02_shape_analysis.py ──→ shape_analysis_map.json     │
 │         ↓                    prompt_specs.json         │
 │         ↓                    readability_budget.json   │
 │  03_build_shape.py ──→ build_shape_content.json       │
 │         ↓                prompt_trace.json             │
 │  03_build_ppt_com.py ──→ codex X.Y.pptx              │
 │         ↓                  post_write_readback.json    │
 │  04_shape_diff_test.py ──→ diff_result.json           │
 │                             fix-ppt.md                 │
 └─────────────────────────────────────────────────────┘
             │
             ▼
      用户检查 PPT → 通过 → 整合到 main.py + src/


 ---
 文件结构   ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️

 project root/
 ├── pipeline/                           # 新建 - Agent产出的脚本集
 │   ├── __init__.py                     # 空文件
 │   ├── ppt_pipeline_common.py          # 共享工具（我先创建基础版）
 │   ├── 01_shape_detail.py              # Step 1: 模板shape识别
 │   ├── 02_shape_analysis.py            # Step 2: 角色推断 + prompt规格
 │   ├── 03_build_shape.py              # Step 3A: 内容构建（策略矩阵）
 │   ├── 03_build_ppt_com.py            # Step 3B: 模板克隆 + COM写入
 │   └── 04_shape_diff_test.py          # Step 4: 三层差异测试
 ├── src/
 │   ├── __init__.py                     # 新建（空文件，使 src 成为 package）
 │   ├── Function_030.py                 # 不修改 - 复用 GPT_5/extract_info 等
 │   ├── Class_030.py                    # 不修改
 │   ├── Global_var_030.py               # 不修改
 │   └── Template 2.1.pptx              # 不修改
 ├── 2025 数据 v2.2.xlsx                # 不修改
 ├── main.py                            # 不修改（未来整合才改）
 ├── orchestrator.py                     # 已配置好（上次会话改造完成）
 ├── new-ppt-workflow.md                 # v4.0 执行规范
 │
 │   # --- 以下为脚本运行时产出（不预先创建） ---
 ├── shape_detail_com.json              # Step 1 产出
 ├── shape_fingerprint_map.json         # Step 1 产出
 ├── shape_analysis_map.json            # Step 2 产出
 ├── prompt_specs.json                  # Step 2 产出
 ├── readability_budget.json            # Step 2 产出
 ├── build_shape_content.json           # Step 3A 产出
 ├── prompt_trace.json                  # Step 3A 产出
 ├── content_validation_report.md       # Step 3A 产出
 ├── shape_data_gap_report.md           # Step 3A 产出
 ├── codex X.Y.pptx                     # Step 3B 产出（最终PPT）
 ├── build-ppt-report.md                # Step 3B 产出
 ├── post_write_readback.json           # Step 3B 产出
 ├── diff_result.json                   # Step 4 产出（Agent迭代依据）
 ├── fix-ppt.md                         # Step 4 产出（修复建议）
 └── diff_semantic_report.md            # Step 4 产出


 JSON/MD 输出到项目根目录（不在 pipeline/ 内），原因：
 - orchestrator.py 的 _check_bug_report() 已配置在根目录查找 diff_result.json
 - Agent 直接在根目录读取中间产物，无需路径推理

 ---
 我现在创建什么 vs Agents 后续做什么

 我创建（基础设施 + 初始脚本框架）

 文件: src/__init__.py
 内容: 空文件
 行数: 0
 ────────────────────────────────────────
 文件: pipeline/__init__.py
 内容: 空文件
 行数: 0
 ────────────────────────────────────────
 文件: pipeline/ppt_pipeline_common.py
 内容:
 共享工具（路径、COM安全、Excel读取、指标提取、文本截断、导入Function_030）
 行数: ~220
 ────────────────────────────────────────
 文件: pipeline/01_shape_detail.py
 内容: Step 1 完整实现（基于 codex-legacy2 + bug修复）
 行数: ~130
 ────────────────────────────────────────
 文件: pipeline/02_shape_analysis.py
 内容: Step 2 完整实现（基于 codex-legacy2 + 增强）
 行数: ~150
 ────────────────────────────────────────
 文件: pipeline/03_build_shape.py
 内容: Step 3A 完整实现（策略矩阵 + GPT prompt）
 行数: ~210
 ────────────────────────────────────────
 文件: pipeline/03_build_ppt_com.py
 内容: Step 3B 完整实现（模板克隆 + COM写入）
 行数: ~140
 ────────────────────────────────────────
 文件: pipeline/04_shape_diff_test.py
 内容: Step 4 完整实现（三层评分算法）
 行数: ~210

 总计约 ~1060 行，基于 codex-legacy2 代码修复+增强，包含4个已知 bug 的修复。

 Agents 后续迭代

 Agent 通过 orchestrator.py 循环：
 1. Tester 运行 Step 1-4，检查 diff_result.json
 2. 不达标 → Developer 读 fix-ppt.md，修改对应脚本
 3. 再次运行 Step 1-4 → 再检查
 4. 直到 Visual≥98, Readability≥95, Semantic=100

 Agents 可能修改的内容：
 - 02_shape_analysis.py 的角色推断规则（如果初始启发式不准确）
 - 03_build_shape.py 的 GPT prompt 措辞（如果生成内容不够精准）
 - 03_build_ppt_com.py 的 shape 查找逻辑（如果 name 不匹配）
 - 04_shape_diff_test.py 的评分权重（如果过严/过松）

 ---
 各脚本技术方案

 ppt_pipeline_common.py — 共享工具

 来源：codex-legacy2/ppt_pipeline_common.py（216行），修复后版本。

 关键修复：
 1. ROOT 路径：Path(__file__).parent.parent（指向项目根，非 pipeline/）
 2. load_legacy_functions()：改用 from src.Function_030 import ...（创建
 src/__init__.py 后可用）
 3. 新增 is_in_group(shp)：try-except 替代 getattr(shp, "ParentGroup", None)
 4. load_excel_rows()：添加模糊 sheet 匹配（'问卷' in sheet.name）

 保留不变：now_ts(), safe_text(), numeric(), to_rows(), com_get(), com_call(),
 write_md(), write_json(), extract_metrics(), extract_score_means(),
 clamp_text()

 01_shape_detail.py — Step 1: 模板 shape 识别

 任务：对比 Template 第14页(空白) vs 第15页(标准)，提取新增 shape 属性 + 指纹。

 每个 shape 提取：name, left, top, width, height, shape_type, text, font_name,
 font_size, has_chart, in_group, z_order

 指纹：shape_type + rounded(geometry) + name + text_prefix(20) + z_order +
 in_group

 输出：shape_detail_com.json（shape属性列表）, shape_fingerprint_map.json

 COM安全：
 - has_chart: try: bool(shp.HasChart) except: False
 - in_group: try: shp.ParentGroup; True except: False
 - text: try: shp.TextFrame.TextRange.Text except: ""
 - 整体用 try-finally 确保 app.Quit()

 02_shape_analysis.py — Step 2: 角色推断 + Prompt 规格

 任务：读取 shape_detail_com.json + Excel 数据，为每个 shape 分配角色、生成
 prompt 规格和可读性预算。

 角色推断优先级：
 1. has_chart → chart
 2. name 含 "title" → title
 3. text 含 "样本/人数/N=" → sample_stat
 4. text 含 "建议/结论/总结" → insight
 5. len(text) ≤ 18 且非空 → title
 6. len(text) ≥ 80 → long_summary
 7. 默认 → body

 可读性预算：max_chars = max(18, min(240, template_len * 1.2))；max_lines:
 title=1, sample_stat=1, insight=4, 其余=6

 输出：shape_analysis_map.json, prompt_specs.json, readability_budget.json

 03_build_shape.py — Step 3A: 按策略矩阵构建内容

 策略矩阵：

 角色: title
 策略: 模板原文直出
 GPT?: 否
 数据来源: shape_info["text"]
 ────────────────────────────────────────
 角色: sample_stat
 策略: 样本量聚合
 GPT?: 否
 数据来源: f"有效样本 N={count}"
 ────────────────────────────────────────
 角色: chart
 策略: 评分均值提取
 GPT?: 否
 数据来源: parse_survey_data() → 列均值
 ────────────────────────────────────────
 角色: body
 策略: extract_info() → GPT fallback
 GPT?: 条件
 数据来源: 问卷原始数据
 ────────────────────────────────────────
 角色: long_summary
 策略: 模板锚点 + 数据驱动 GPT
 GPT?: 是
 数据来源: 模板文本 + 统计指标
 ────────────────────────────────────────
 角色: insight
 策略: 模板锚点 + 行动建议 GPT
 GPT?: 是
 数据来源: 模板文本 + 统计指标

 GPT 调用：复用 src/Function_030.py 的 GPT_5(prompt, model)
 数据提取：复用 parse_survey_data(), extract_info()
 硬截断：clamp_text(content, max_chars, max_lines) 强制执行预算
 Fallback：每个 GPT 角色有硬编码的兜底文本

 输出：build_shape_content.json, prompt_trace.json,
 content_validation_report.md, shape_data_gap_report.md

 03_build_ppt_com.py — Step 3B: 模板克隆 + COM 写入

 流程：
 1. 打开 PowerPoint，打开 Template
 2. 克隆第15页到新 Presentation
 3. 读取 build_shape_content.json
 4. 按 shape name 精确查找（fallback 用几何位置匹配）
 5. 文本 shape：TextFrame.TextRange.Text = content（不碰格式！）
 6. 图表 shape：Chart.ChartData → 写入 worksheet → SetSourceData()
 7. 写后回读验证
 8. 保存为 codex {version}.pptx

 COM安全：
 - Copy/Paste 后 time.sleep(delay)
 - Chart workbook 写完后 wb.Close(False)
 - try-finally 确保资源释放

 输出：codex {version}.pptx, build-ppt-report.md, post_write_readback.json

 04_shape_diff_test.py — Step 4: 三层差异测试

 对比对象：Template 第15页 vs codex.pptx 第1页

 三层评分：

 层级: Visual ≥ 98
 阈值: left/top/width/height 偏移(w=10/10/8/8, 容差20px), shape_type(w=8),
   chart_type(w=16), font(w=4+4)
 检查项:
 ────────────────────────────────────────
 层级: Readability ≥ 95
 阈值: 文本相似度(w=0.6, SequenceMatcher), 长度比(w=0.25), 行数比(w=0.15)
 检查项:
 ────────────────────────────────────────
 层级: Semantic = 100
 阈值: 关键词覆盖: ["样本", "建议", "反馈", "评分"]
 检查项:

 Shape 匹配：按 name 匹配（非按索引），fallback 用几何位置

 输出：
 - diff_result.json：结构化评分（Agent 读取判断通过/不通过）
 - fix-ppt.md：逐 shape 差异明细 + 修复建议（Agent 读取修复代码）
 - diff_semantic_report.md：语义覆盖分析

 ---
 已知 Bug 修复清单

 #: 1
 Bug: getattr(shape, "ParentGroup", None) 抛 COM 异常
 来源: debug/Mc-debug.md
 修复方案: is_in_group() + try-except
 ────────────────────────────────────────
 #: 2
 Bug: Function_030.py 相对导入失败
 来源: codex-legacy2
 修复方案: 创建 src/__init__.py + from src.xxx import
 ────────────────────────────────────────
 #: 3
 Bug: Chart 数据 split(":") 对含冒号值失败
 来源: codex-legacy2 分析
 修复方案: rsplit(":", 1) + workbook 关闭
 ────────────────────────────────────────
 #: 4
 Bug: Diff 测试按 shape 索引匹配（重排序则错配）
 来源: codex-legacy2 分析
 修复方案: 按 name 匹配 + 几何位置 fallback

 ---
 使用方式

 方式一：手动逐步运行（调试阶段）

 python pipeline/01_shape_detail.py                          # Step 1
 python pipeline/02_shape_analysis.py                        # Step 2
 python pipeline/03_build_shape.py                           # Step 3A
 python pipeline/03_build_ppt_com.py --version 1.0           # Step 3B
 python pipeline/04_shape_diff_test.py --target "codex 1.0.pptx"  # Step 4


 方式二：Agent 驱动（正式构建）

 python orchestrator.py
 > 根据 new-ppt-workflow.md，运行 pipeline/
 脚本生成问卷分析PPT，并通过三层差异测试


 orchestrator.py 会自动：
 1. Architect 确认 PLAN.md
 2. Developer 运行 pipeline/ 脚本
 3. Tester 检查 diff_result.json（Visual≥98, Readability≥95, Semantic=100）
 4. 不通过 → Developer 读 fix-ppt.md 修改脚本代码 → 重新运行
 5. 通过 → Optimizer + Security

 方式三：未来整合（用户评审通过后）

 # main.py 中新增调用：
 from pipeline.build_questionnaire_page import run_pipeline
 mc_slide = run_pipeline(mc_sht, mc_ppt, mc_slide, sample_name, mc_model)

 ---
 实施顺序

 Phase 1: 基础设施（2个空文件）
   1. 创建 src/__init__.py
   2. 创建 pipeline/ 目录 + __init__.py

 Phase 2: 共享工具
   3. 编写 pipeline/ppt_pipeline_common.py

 Phase 3: 步骤脚本（按依赖顺序）
   4. 编写 pipeline/01_shape_detail.py
   5. 编写 pipeline/02_shape_analysis.py
   6. 编写 pipeline/03_build_shape.py
   7. 编写 pipeline/03_build_ppt_com.py
   8. 编写 pipeline/04_shape_diff_test.py

 Phase 4: 冒烟测试
   9. 验证导入：python -c "from src.Function_030 import GPT_5; print('OK')"
   10. 运行 Step 1 确认 shape_detail_com.json 产出正确


 Phase 2-3 总计约 1060 行代码，基于 codex-legacy2 修复+增强。

 Phase 4 之后，用户可以选择：
 - 手动逐步运行调试
 - 用 orchestrator.py 让 Agents 接管迭代


----------- 上述计划内容已保存至 plan.md ----------
♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️










我希望预留一个辅助入口：

假设我执行了agents工作流程（假设agents自动迭代了1-3轮，根据用户设定的最大循环来确定），发现生成的ppt跟我设想的差距仍然比较大
假设agents生成了 shape-detail.md（给用户核验） 和 shape-detai.json（给其他agent编程）
那么我希望可以人工编辑shape-detail.md
我会在每个对应的shape下面用文字描述：
1、这个shape的内容具体来自哪里
2、建议用什么方式生成这个shape的内容
3、其他要修正的地方和细节

然后我会再重新启动agents工作流程，让agents基于现有的shape-detail.md（给用户核验） 和 shape-detai.json（给其他agent编程）
在这个基础上继续工作，而不是生成全新的shape-detail  md/json文件








我还是有个疑问。我将这个问题想太复杂了吗？  需要6个agent来协作完成吗？

  还是说你可以直接完成？我最终的需求是，我提供【标准模板（含内容+格式）】         【空白模板（仅格式）】【数据源Excel】三个文件，然后得到一套准确的py代码，       可以稳定地批量生产同样格式的ppt，但内容可能经常变化，只有格式保持稳定不变       。                                                                              但我对ppt的排版精确度要求非常高。
  6个agents完成这项工作，跟直接交给你完成相比，agents流程有什么优势可言吗？
  我的确有大量ppt模板需要实现自动化，每个模板都详细沟通、交流、debug、不太
  现实，我希望尽可能自动化实现这个【自动化的任务】

● 这是个很好的问题，我直接说结论。                                              
  ---
  6 agents vs 直接跟我协作
                                                                                  对你这个任务，6 agents 没有明显优势。

  ┌────────────┬────────────────────────────┬───────────────────────────────┐
  │            │          直接跟我          │           6 agents            │
  ├────────────┼────────────────────────────┼───────────────────────────────┤
  │ 适合场景   │ 需求清晰、任务可迭代       │ 架构未知、需多视角审查        │
  ├────────────┼────────────────────────────┼───────────────────────────────┤
  │ 调试效率   │ 高（你直接看到我改了什么） │ 低（agents 互相调用，难追踪） │
  ├────────────┼────────────────────────────┼───────────────────────────────┤
  │ Token 消耗 │ 低                         │ 高（6个角色各自消耗）         │
  ├────────────┼────────────────────────────┼───────────────────────────────┤
  │ 精度控制   │ 高（你逐步确认）           │ 低（agents 可能过度改动）     │
  ├────────────┼────────────────────────────┼───────────────────────────────┤
  │ 你这个需求 │ ✅ 合适                    │ ❌ 过度设计                   │
  └────────────┴────────────────────────────┴───────────────────────────────┘

  ---
  真正的问题：多模板怎么做到自动化

  你的瓶颈不是"代码生成"，而是每个模板的 shape 业务逻辑定义——每个 shape
  装什么内容、数据从哪里来。这是人机协作无法完全跳过的一步。

  但可以大幅提速。 我建议这样的流程：♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️

  新模板进来
      ↓
  Step 1 自动跑 → 生成 shape_detail.md（已实现）
      ↓
  你用 10-15分钟 填批注（生成方式/内容来源）← 不可跳过，但可以快
      ↓
  Step 2-4 全自动跑 → 出 PPT
      ↓
  你看结果，如有问题在 shape_detail.md 里改批注 → 重跑

  可以进一步优化的地方：Step 1 之后，让我看一眼 shape_detail.md
  帮你预填批注（根据 shape 的模板文字 + Excel
  列名推断），你只需要审核、微调，不用从头写。

  ---
  实际建议

  1. 当前调试阶段：继续跟我直接协作，pipeline 跑稳为止
  2. 新模板上线时：运行 Step 1 → 把 shape_detail.md 发给我 → 我帮你填批注 →
  你确认 → 跑 Step 2-4
  3. orchestrator + 6 agents：留给真正复杂的任务（比如未来需要给 pipeline
  加新功能模块时）

  一个模板上线的工作量大概是：你 10 分钟 + 我 2 分钟 + 自动跑 3
  分钟，不需要每次都从头 debug 大量代码。














♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 

咱们抛开这个任务，从agents本质来讨论。agents + orchestrator
  的工作流程的优势在哪里？ 什么任务适合直接交给你？什么任务适合用agents +
  orchestrator的流程？ 你系统性回答下这个问题，并详细解释说明、对比     
  
   ● 本质区别：单次对话 vs 自主执行系统                                            
  ---
  你（Claude）是什么
  每次对话，我本质上是一个无状态的函数：

  输入（上下文窗口） → 思考 → 输出

  - 上下文窗口有限（~200K tokens）
  - 无法自主执行下一步——需要你触发
  - 无法并行——我是单线程的
  - 遗忘：对话结束后状态清空

  ---
  Agents + Orchestrator 解决了什么

  本质是：把"一次大请求"拆成"多个自主执行的小任务"，并允许并行、循环、状态持久化
  。

  orchestrator
      ├── Agent A（独立上下文）→ 产出物 A
      ├── Agent B（独立上下文）→ 产出物 B  ← 并行
      └── Agent C 读取 A+B 的产出物 → 产出物 C
           └── 如果不达标 → 重新执行 A/B（循环）

  ---
  核心优势对比

  优势 1：上下文隔离 / 不受窗口限制

  ┌──────────────┬──────────────────────────┬───────────────────────────────┐
  │              │         直接跟我         │            Agents             │
  ├──────────────┼──────────────────────────┼───────────────────────────────┤
  │ 上下文       │ 单一窗口，越聊越拥挤     │ 每个 agent 有独立干净的上下文 │
  ├──────────────┼──────────────────────────┼───────────────────────────────┤
  │ 适合任务规模 │ 中小型（几百行代码级别） │ 大型（上万行、跨多文件改动）  │
  └──────────────┴──────────────────────────┴───────────────────────────────┘

  典型场景：让 Agent A 只看 Function_030.py，Agent B 只看
  03_build_shape.py，二者不互相污染。

  ---
  优势 2：并行执行

  直接跟我：串行——你问一步我答一步。

  Agents：并行——多个 agent 同时工作。

  例：分析10个PPT模板的shapes
    直接跟我：分析模板1 → 分析模板2 → ... → 分析模板10（串行，慢）
    Agents：  10个agent同时分析各自的模板（并行，快10倍）

  ---
  优势 3：自主循环迭代（无需人在场）

  这是最重要的优势：

  直接跟我：
    我改代码 → 你运行 → 你看结果 → 你告诉我哪里不对 → 我再改
    ← 每一步都需要你在场 →

  Agents：
    Developer改代码 → Tester自动运行 → Tester自动判断是否达标
    → 不达标 → Developer自动读报告再改 → 再测
    ← 你去喝杯咖啡，回来看结果 →

  关键点：agents 可以在你不在场的情况下完成"改→测→改→测"循环。

  ---
  优势 4：角色分工降低错误率

  单一 Claude 视角：容易"自我满足"（我改了觉得对就结束了）

  多 Agent 角色分工：
  Developer（写代码） ≠ Tester（挑毛病） ≠ Security（找安全漏洞）
  不同角色有不同的"职责偏见"，互相制衡。类似代码审查（Code Review）的效果。

  ---
  什么任务适合直接跟我

  核心判断标准：任务是否需要你全程参与判断

  ✅ 适合直接跟我的任务：

  ┌──────────────────────────────────────┬──────────────────────────────┐
  │                 场景                 │             原因             │
  ├──────────────────────────────────────┼──────────────────────────────┤
  │ 需求不断变化、边做边想清楚           │ 需要快速人机对话调整方向     │
  ├──────────────────────────────────────┼──────────────────────────────┤
  │ 中小规模代码（< 500行）              │ 上下文容得下，不需要多 agent │
  ├──────────────────────────────────────┼──────────────────────────────┤
  │ 调试已有代码的具体 bug               │ 精确定位，不需要并行         │
  ├──────────────────────────────────────┼──────────────────────────────┤
  │ 一次性任务（写个脚本、解释一段代码） │ 无需状态持久化               │
  ├──────────────────────────────────────┼──────────────────────────────┤
  │ 探索性工作（不知道该怎么做）         │ 需要边试边确认方向           │
  ├──────────────────────────────────────┼──────────────────────────────┤
  │ 你现在的任务                         │ 需求在对话中逐步明确         │
  └──────────────────────────────────────┴──────────────────────────────┘

  ---
  什么任务适合 Agents + Orchestrator

  核心判断标准：任务是否可以被明确定义、自主执行、客观验证

  ✅ 适合 Agents 的任务：

  场景: 大型代码库重构（几千行、跨多文件）
  原因: 需要隔离上下文 + 并行分析
  ────────────────────────────────────────
  场景: 批量处理（10个模板同时分析）
  原因: 天然并行
  ────────────────────────────────────────
  场景: 有明确验收标准的任务（测试通过率、diff 评分）
  原因: 可以自主判断完成与否
  ────────────────────────────────────────
  场景: 需要长时间无人值守运行
  原因: 循环迭代不需要人在场
  ────────────────────────────────────────
  场景: 需要多视角审查（安全/性能/功能同时检查）
  原因: 角色分工
  ────────────────────────────────────────
  场景: CI/CD 流水线集成
  原因: 触发条件明确，结果可量化

  ---
  一句话总结

  直接跟我 = 结对编程（你在旁边，随时调整方向）

  Agents   = 雇了一个小团队去执行你写好的需求文档
             （你不在场，他们自己跑完交付结果）

  前提是：agents 只有在"需求文档写得足够清晰、验收标准可量化"的情况下才能发挥优
  势。如果需求模糊，agents 会很快跑偏，修正成本反而更高。

  ---
  你的情况的实际建议

  探索期（现在）：  直接跟我  ← 需求在变、边做边想
                       ↓
  稳定期：         定义好每类模板的标准规范文档
                       ↓
  批量生产期：      Agents  ← 输入规范，输出可验证，无需人在场

  当你有了 10 个模板的经验，知道哪些问题会出现、验收标准是什么，那时候交给
  Agents 才能发挥最大价值。现在还在"搞清楚要怎么做"的阶段，Agents 帮不了你这个。






 当你有了 10 个模板的经验，知道哪些问题会出现、验收标准是什么，那时候交给         Agents 才能发挥最大价值。现在还在"搞清楚要怎么做"的阶段，Agents
  帮不了你这个。   我的工作对ppt精度要求非常高，目前我找到的工作流程就是现        在这些。你也非常清楚我的需求了，对于接下来的开发，你有什么好的建议？我希
  望最终能清晰地交付给 agents完成
   ● 现状评估    当前这个模板还没跑完（Step 3B/4 还没验收）。先说清楚终点，再说路径。          
  
  ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 





  ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 
  ---
  你真正需要的终态

  输入：标准模板 + 空白模板 + Excel
    ↓
  Agent 自动完成：
    Step 1 提取 shapes
    Step 2 分析角色
    Step 2.5 ← 【唯一需要人参与的节点】审核 shape_detail.md 批注
    Step 3A/3B 生成内容 + 写入 PPT
    Step 4 验收
    → 通过则交付，不通过则 Agent 自动修复循环

  人的不可替代职责只有一个：判断每个 shape
  应该装什么内容。这个语义判断不能被自动化，但可以被极度简化。

  ---
  我的分阶段建议

  Phase 1：先把这个模板跑完（现在）

  目标：Step 3B/4 完整跑通，打开 PPT 眼睛确认。

  这一步最重要，因为：
  - Step 3B 的 COM 写入逻辑还没验证过
  - diff 测试的评分阈值（Visual ≥ 98）是否合理还不知道
  - 你还没亲眼看到一份完整的输出

  不要跳过这步就去做 Phase 2。

  ---
  Phase 2：收集 2-3 个模板的经验后，做一件事

  把 shape_detail.md 的批注格式从"自由文本"变成"结构化表单"

  现在批注是这样的：
  - 生成方式: 问卷评分均值小于5分，说明问卷是5分制；否则就是10分制...

  改成这样：
  - 生成方式: score_10pt
  - 生成参数: scale=auto, format=X.XX/10

  为什么：Agent 解析结构化字段比解析自然语言可靠 10 倍。自由文本需要 NLP
  推断，结构化字段是精确指令。

  ---
  Phase 3：写一份"Agent 上线手册"

  当你有 3-5 个模板的经验后，你会发现 shape 类型其实就那几种：

  评分显示类（score_10pt / grade_letter）
  统计摘要类（sample_aggregation）
  提名/名称类（extract_column）
  长文总结类（gpt_rich：优点/缺点/综合）
  图表类（mean_extraction）
  装饰类（skip）

  把每种类型的标准批注格式 + 示例 + 验收标准写成文档，Agent 拿到新模板后：

  1. 运行 Step 1
  2. 对照手册填批注（已知类型直接套用）
  3. 你只需要确认"这个 shape 属于哪种类型"

  ---
  Phase 4：Agent 化（10 个模板后）

  此时你有了：
  - 稳定的 pipeline 代码
  - 结构化批注格式
  - Agent 上线手册
  - 已知的 diff 评分基准

  可以把这个流程交给 Agent：

  Agent 接收三个文件 →
    自动跑 Step 1 →
    根据手册预填批注 →
    你 10 分钟确认批注 →
    Agent 跑 Step 2-4 →
    diff 不通过则自动迭代 →
    通过则交付

  ---
  一个关键原则

  精度来自格式克隆，不来自 GPT

  Step 3B 是精度保障核心——它从模板克隆格式，GPT 只填内容。你对 PPT
  精度的高要求，实际上已经被这个架构设计兜底了。

  真正影响精度的风险点只有两个：
  1. COM 写入破坏原有格式（字体/颜色被覆盖）→ tr.Text = content 而不是重建
  TextRange
  2. 内容溢出 shape 边界→ clamp_text() + 字数 budget

  这两点在当前代码里已经处理，但需要你眼睛看一遍实际 PPT 来确认。

  ---
  ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 
  现在最重要的一步：跑完 Step 3B/4，打开 PPT 看效果，把发现的问题告诉我。







 ---
  什么时候才需要 --force：

  ┌──────────────────────────────┬────────────────────────────┐
  │             情况             │        需要 --force        │
  ├──────────────────────────────┼────────────────────────────┤
  │ 换了新的 PPT 模板            │ ✓                          │
  ├──────────────────────────────┼────────────────────────────┤
  │ 模板里增加/删除了 shape      │ ✓                          │
  ├──────────────────────────────┼────────────────────────────┤
  │ 批注全部重写，不想保留旧版本 │ ✓                          │
  ├──────────────────────────────┼────────────────────────────┤
  │ 模板没变，只是调整批注内容   │ ✗ 直接编辑 shape_detail.md │
  ├──────────────────────────────┼────────────────────────────┤
  │ 像现在这样，批注已完善       │ ✗                          │
  └──────────────────────────────┴────────────────────────────┘

  ---
  你现在的状态： 所有文件都是最新的

  shape_detail.md          ✓ 有完整批注 + 新结构化字段
  shape_analysis_map.json  ✓ 刚刚 Step 2 已更新（含 strategy_exact/params）
  build_shape_content.json ✓ Step 3A 已生成

  直接继续 Step 3B 写入 PPT：
  PYTHONIOENCODING=utf-8 python pipeline/03_build_ppt_com.py --version 1.2








> 所以我每次更新批注，都需要你帮我更新【修正说明】吧？毕竟我自己不会写这种        结构化的描述。。。  pipeline现在不会再去读我的自然语言描述了吧？



两个好问题，直接回答：                                                                                                                                          ---
  Q1：pipeline 还会读自然语言吗？
                                                                                  不再读了（对这9个 shape）。路由逻辑是：
  strategy_exact 有值 → 直接分发，自然语言完全跳过
  strategy_exact 为空 → 才会 keyword 匹配 build_strategy（自然语言）

  你的自然语言 生成方式 现在只是给你自己看的历史记录，代码不碰它。

  ---
  Q2：以后改批注需要我帮你吗？

  大多数情况你自己能改，params 词汇表很小：

  你想改什么: 改筛选关键词
  直接编辑: params 里的 filter=
  例子: filter=缺点 → filter=优缺点
  ────────────────────────────────────────
  你想改什么: 换列名
  直接编辑: params 里的 column=
  例子: column=鞋款名称 → column=产品型号
  ────────────────────────────────────────
  你想改什么: 换策略
  直接编辑: strategy: 字段
  例子: strategy: score_10pt → strategy: template_direct

  需要找我的情况：
  - 新模板有新的 shape 类型，现有8个 strategy code 都不匹配
  - 想新增一种 params key（比如 format= 的行为要改）
  - 任何需要改 Python 代码的需求

  ---
  实际上，自然语言应该降级为纯注释

  现在的 生成方式
  字段既像"操作指令"又像"注释"，容易混淆。建议把它改个名，明确语义：

  ### 用户批注

  - 内容来源: 问卷原始评价文本，按【缺点】进行分段汇总   ← 你自己看的说明
  - 策略备注: gpt_prompted，对应矩形68的【缺点】汇总     ← 改个名，明确是注释
  - 修正说明: 每段不超过3行，保留括号里的测试比例数据
  - strategy: gpt_prompted        ← 机器读这行
  - params: source=补充说明, filter=缺点  ← 机器读这行

  要不要我把 generate_shape_detail_md()
  里的模板占位符改成这个结构？下次新模板上线，Step 1 自动生成的格式就会更清晰。





claude-yk --resume    今晚收获不小，明天继续！





 Step 1 shape_detail.md 迭代更新 — 原地合并而非版本文件

     Context

     用户希望当 shape_detail.md 已存在时，01_shape_detail.py
     能处理模板变更，
     而不是强制覆盖（丢失批注）或跳过（无法感知新 shape）。

     提出两种方案：
     - A：生成版本文件 (shape_detail_v1.1.md, v1.2.md…)
     - B：原地合并——重新提取 shapes，保留已有批注，新 shape 给空占位符

     ---

方案 B：原地合并（推荐）

     实现难度实际很低（约 20 行新代码）：
     - parse_user_annotations() 已经能提取所有批注 → 直接复用
     - generate_shape_detail_md() 加一个 existing_annos
     参数，有则填入，无则留空
     - 两个 JSON 文件（shape_detail_com.json,
     shape_fingerprint_map.json）本来就总是覆盖
   












很好，右上角原本2个识别错误的shape，已经纠正过来啦。

● 很好！全部走通了：

  - 矩形 12 → A- ✓（等级字母修正）                                                - 矩形 19 → 跳过 ✓（装饰条不再写内容）
  - 矩形 68/77 → GPT 生成缺点/优点 ✓
  - updated 7/9：另外 2 个是 skip 策略（矩形 19 + 图片 39），正常

  下一步跑 Step 4 看质量评分：

  pipeline/04_shape_diff_test.py

  跑完看 diff_result.json 或终端输出，会给出 visual / readability
  分数。如果分数低，再针对性调整。

> 03b 为什么显示 updated 7/9 shapes?   为什么有9个，只更新7个？  另外，PS         C:\Video Games\Claude Code\3rd-ppt-prj> &
  C:\Users\HQ\AppData\Local\Programs\Python\Python312\python.exe "c:/Video        Games/Claude Code/3rd-ppt-prj/pipeline/04_shape_diff_test.py"
  [FAIL] visual=98.06 readability=92.44 semantic=33                               PS C:\Video Games\Claude Code\3rd-ppt-prj>
  ● 7/9 的原因很简单：                                                              - 矩形 19 → skip（装饰条，空内容不写）
  - 图片 39 → skip（图片 shape，无文字框）                                      
  9 个 shape，2 个跳过 = 7 个实际写入。完全正常。

  ---
  分数解读：

  ┌─────────────┬───────┬──────┬───────────────┐
  │    指标     │ 得分  │ 门槛 │     状态      │
  ├─────────────┼───────┼──────┼───────────────┤
  │ visual      │ 98.06 │ ≥98  │ ✅ 通过       │
  ├─────────────┼───────┼──────┼───────────────┤
  │ readability │ 92.44 │ ≥95  │ ❌ 差 2.56 分 │
  ├─────────────┼───────┼──────┼───────────────┤
  │ semantic    │ 33    │ =100 │ ❌ 严重不足   │
  └─────────────┴───────┴──────┴───────────────┘

  先看 diff_result.json 找具体原因：

● Read 1 file (ctrl+o to expand)

● 問題清晰了。讓我看下 Step 4 的評分邏輯，再給出診斷：                          
● Explore(Diagnose diff_result failures)                                          ⎿  Done (3 tool uses · 23.7k tokens · 26s)                                      (ctrl+o to expand)                             














矩形 68: 【包裹性】 弯折处........  以及矩形 77:
  【止滑性/耐磨性】 鞋底橡胶在发力与.........  这两个矩阵的生成逻辑错了。矩
  阵68是评论中关于这个产品的优点总结，控制在280字左右，不一定每次的优点都是       【包裹性】【舒适性】等内容，需要gpt自己去总结提炼。同样，矩阵77是产品的缺       点总结，控制在220字左右。帮我完成相应的文件更新







未来如果交给 agents + orchestrator 工作流程，当我碰到这种逻辑错误时，我应该如何让agents帮我修正？它如何才能自动修正？ 


 这是个很好的架构问题。核心结论是：这类错误有两个层次，自动化程度不同。        
  ---
  错误分类
  你描述的错误属于规格错误（Specification Error），不是代码错误：

  ┌──────────┬────────────────────────────┬──────────────────────────┐
  │ 错误类型 │            示例            │       自动检测难度       │
  ├──────────┼────────────────────────────┼──────────────────────────┤
  │ 代码 Bug │ return score_like 少返回列 │ 容易（输出值可量化验证） │
  ├──────────┼────────────────────────────┼──────────────────────────┤
  │ 格式错误 │ 字数超限、行数不对         │ 容易（readability 分数） │
  ├──────────┼────────────────────────────┼──────────────────────────┤
  │ 规格错误 │ 矩形 68 该写缺点却写成优点 │ 困难（需要人判断意图）   │
  └──────────┴────────────────────────────┴──────────────────────────┘

  规格错误的难点：代码严格按照 01-shape_detail.md 的 filter=优点
  执行，代码没错，是 spec 本身错了。Agents 无法在没有人告诉它的情况下，知道 spec
   是错的。

  ---
  两种接入 Orchestrator 的方式

  方式 A：人工识别 + Orchestrator 执行修复（推荐）

  你发现问题后，用自然语言告诉 orchestrator：

  矩形 68 应该是缺点总结，矩形 77 应该是优点总结，
  当前两者内容反了。请修正相关配置并重新跑 Step 3A→3B→4。

  Orchestrator 的执行路径：

  Architect → 分析哪些文件控制这个逻辑
              → pipeline-progress/01-shape_detail.md (filter= 字段)
              → pipeline/03a_build_shape.py (gpt_prompted 分支)

  Developer → 找到 filter=优点 改为 filter=缺点 (矩形 68)
             → 找到 filter=缺点 改为 filter=优点 (矩形 77)

  Tester    → python pipeline/03a_build_shape.py
             → python pipeline/03b_build_ppt_com.py --version x.y
             → python pipeline/04_shape_diff_test.py --target "codex x.y.pptx"
             → 检查 04-diff_result.json → 报告通过/失败

  你只需一句话，agents 完成剩余所有机械操作。这是当前架构下最实用的方式。

  ---
  方式 B：在 Step 4 加入"内容意图验证"（更自动，但有代价）

  在 04_shape_diff_test.py 里加一个 intent_check：

  # 读取 01-shape_detail.md 中每个 shape 的 filter 参数
  # 用 GPT 判断生成的文本是否符合 filter 意图（优点 vs 缺点）
  # 如果不符合 → 记录到 04-fix_ppt.md → tester 将其标记为 semantic 失败

  然后 tester agent 读取 04-fix_ppt.md 的修复建议，自动触发 developer 修复
  spec。

  代价：每次 Step 4 多一次 GPT 调用 + 误判风险（GPT
  有时分不清"鞋底耐磨"是优点还是缺点）。

  ---
  实际建议

  规格错误：你发现 → 一句话告诉 orchestrator → agents 自动修复+验证
  代码 Bug：agents 自动检测（量化分数不达标）→ 自动迭代修复

  最大价值点：你不需要知道该改哪个文件、改哪行、跑哪条命令——这些机械操作全交给
  agents。你只需要保留「判断内容对不对」这个人类判断环节。

  要实现方式 A，orchestrator 需要能接受自然语言问题单、调度
  architect→developer→tester 这条链路。这正是当前 src/orchestrator_v6.py
  设计的用途——你可以考虑在下一个迭代中加这条链路的测试。








♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 

Pipeline 5个脚本说明

 python pipeline/01_shape_detail.py  ------------------- 【建筑承重墙施工】
  提取模板 shape 信息 → 扫描 Template 2.1.pptx，把每个 shape
  的位置、字体、文字、类型写入 JSON 和 MD 文件。MD 文件是给你填批注用的（"这个
  shape 应该填什么、用什么策略"）。 
                      ♥️ ♥️ ♥️ 注意：我填写完批注后，可以让Claude帮我更新成【结构化指令】（这不就是中文编程嘛。。。。） ♥️ ♥️ ♥️  




  ---
  python pipeline/02_shape_analysis.py ----------------- 【室内设计师】
  分析 + 分配策略 → 读 Step 1 的 JSON + 你填的批注 + Excel 数据，为每个 shape
  决定：角色是什么（title/chart/long_summary…）、用什么策略生成内容、字数/行数预
  算是多少。

                              --- ♥️ ♥️ ♥️ 
                              01 和 02 为什么拆分，而不是合并成一步？

                              不是为了提高精度，是为了节省时间：

                              你花了 2 周填写的批注（每个 shape 该怎么做、用什么策略），不应该每次生成 PPT
                              都要重新填一遍。Step 1 只跑一次，批注只填一次，之后每次换新数据只需要从 Step 2
                              或 Step 3 跑。

                              简单说：
                              模板变了 → 从 Step 1 开始
                              只改批注/策略 → 从 Step 2 开始
                              只改数据/代码 → 从 Step 3 开始

                              所以01 和 02的区别是，02相当于说明书，比01更精确，可以用来指导后面的工序，如何精确加工shape的内容？ 
                               01负责说明：有几个shape/是什么/基本信息；02负责：shape的内容是如何产生的/怎么加工                                                                                              ● 对，你理解得很准确。                                                          
                                  Step 1：扫描模板，回答「有什么」
                                  - 这个 PPT 有几个 shape
                                  - 每个 shape 在哪、多大、什么字体、现在写了什么文字 
                                  - 纯物理信息，跟业务无关

                                  Step 2：读批注+数据，回答「怎么做」
                                  - 这个 shape 的内容应该从哪里来（Excel / GPT / 模板直抄）
                                  - 用什么策略生成（score_10pt / mean_extraction / gpt_prompted…）
                                  - 字数/行数上限是多少
                                  - 是业务决策，跟模板的物理结构无关
                                              
                                              --- ♥️ ♥️ ♥️ 
                                              那么，01就是建筑施工，02是室内精装修。难怪两者完全隔离，这样拆分是有必要的
                                                ● 对，这个比喻更贴切。
                                                01 = 建筑施工：看模板PPT，把结构打出来——墙在哪、门在哪、柱子多粗，跟装什么内容无关。
                                                02 = 室内精装修方案：拿着01的结构图 + 业主需求（批注）+材料清单（Excel），规划每个房间怎么装、用什么材料、预算多少。

                                                ---
                                                隔离的好处你已经感受到了：

                                                - 换业主（换一份 Excel 数据）→ 只需要从 Step 3 跑
                                                - 改装修方案（改批注策略）→ 从 Step 2 跑
                                                - 推倒重建（换模板 PPT）→ 从 Step 1 跑

                                                三层变更频率完全不同，所以分开是值得的。                             


  ---
  python pipeline/03a_build_shape.py ----------------- 【装修：工厂加工板材 —— GPT】
  生成文字内容 → 按照 Step 2 的策略，从 Excel 提取数据或调用 GPT，为每个 shape
  生成最终文字。产出一个 JSON，记录每个 shape 该填什么内容。
                        
                        --- ♥️ ♥️ ♥️ 

                        等等，03 为啥又拆分成了 a / b ? 继续装修的例子，解释下

                        ● 03A = 备料：按装修方案，提前把所有材料加工好——木板切成指定尺寸、墙漆调成指定颜色
                         （从 Excel 提数据、调用 GPT生成文字）。结果放在仓库里（03a-build_shape_content.json）。
                          
                          03B = 施工安装：拿着备好的材料，进场逐一装进去——把木板钉上墙、把漆刷上去（用
                          COM 把内容写入 PPT）。

                          ---
                          隔离的好处：

                          - 备料失败（GPT 返回空/Excel 数据错）→ 不用开工，先在仓库检查材料
                          - 材料没问题但安装有问题（COM 写入失败）→ 不用重新备料，直接重跑 03B
                          - 想看看材料对不对（人工审核内容）→ 打开 JSON 检查，确认后再装

                          你刚才就用到了这个好处：chart 数据的 truncation bug 在 03A，COM 写入 bug 在
                          03B，两个可以独立修复、独立重跑。

  ---
  python pipeline/03b_build_ppt_com.py --version 1.x  ----------------- 【装修：现场施工安装】
  写入 PPT → 读 Step 3A 的 JSON，用 COM 自动化打开模板，把内容逐一写入每个
  shape，保存为 codex 1.x.pptx。

  ---
  python pipeline/04_shape_diff_test.py --target "codex 1.x.pptx"
  质量验证 → 对比模板和生成的
  PPT，检查三项指标：visual（形状位置/尺寸）≥98、readability（文字长度/行数）≥95
  、semantic（关键词覆盖）=100。不达标则写入修复建议。

  ---
  一句话记忆法：提取 → 分析 → 生成内容 → 写入PPT → 验证质量

♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 




































 1. orchestrator.py 会自动调用所有 agents 吗？ 
 
 是的，但注意它本身是 Claude 多 Agent 调度器（非直接运行 pipeline 脚本）。调用链为：

  orchestrator.py → 6 个 Claude 子 Agent（arch/tech/dev/test/opti/secu）
      → 每个 Agent 依次运行 pipeline/0X_*.py 脚本
          → 最终输出 codex X.Y.pptx

  运行方式：
  python orchestrator.py new-ppt-workflow.md --auto-architect
  # 或者全自动模式 + 指定循环轮数
  python orchestrator.py new-ppt-workflow.md --auto-architect --max-rounds 3

  ---
  2. 能指定模板/Excel 路径吗？♥️ ♥️ ♥️ ♥️ ♥️ 

  不能，orchestrator.py 没有 --template / --excel 参数。路径硬编码在
  pipeline/ppt_pipeline_common.py：

  EXCEL_PATH   = ROOT / "2025 数据 v2.2.xlsx"
  TEMPLATE_PATH = ROOT / "src" / "Template 2.1.pptx"

  如需改路径，直接编辑 ppt_pipeline_common.py 这两行即可。

  ---
  3. 运行过程会产生哪些文件？

  步骤: Step 1
  输出文件: shape_detail_com.json, shape_fingerprint_map.json, shape_detail.md
  ────────────────────────────────────────
  步骤: Step 2
  输出文件: shape_analysis_map.json, prompt_specs.json, readability_budget.json
  ────────────────────────────────────────
  步骤: Step 3A
  输出文件: build_shape_content.json, prompt_trace.json,
    content_validation_report.md
  ────────────────────────────────────────
  步骤: Step 3B
  输出文件: codex X.Y.pptx, build-ppt-report.md, post_write_readback.json
  ────────────────────────────────────────
  步骤: Step 4
  输出文件: diff_result.json, fix-ppt.md, diff_semantic_report.md

  ---
  4. 最终得到什么？

  根目录下的 codex X.Y.pptx（X.Y 为版本号），即基于模板 + Excel
  问卷数据生成的幻灯片。若 diff 分数不达标（Visual < 98 或 Readability <
  95），fix-ppt.md 会给出修复建议，Agent 会自动迭代（受 --max-rounds 限制）。



所以，如果我检查完自动生成的ppt，如果觉得满意，我可以让你将最终的代码整合
  进入原来的程序中【main + /src函数文件】中？

 ● 可以的。整合思路有两种，你可以选其一：                                                                                                                          ---  方案 A：main.py 作为总入口，顺序调用 pipeline 步骤（推荐，改动最小）

  main.py
    ├── 调用 pipeline/01_shape_detail.py（提取 shapes）
    ├── 调用 pipeline/02_shape_analysis.py（角色分析）
    ├── 调用 pipeline/03_build_shape.py（内容生成）
    ├── 调用 pipeline/03_build_ppt_com.py（COM 写入 PPT）
    └── 调用 pipeline/04_shape_diff_test.py（差异验证）

  优点：pipeline/ 代码几乎不动，只改 main.py 入口；Excel/PPT 路径可以在 main.py
  里统一传入。

  ---
  方案 B：将 pipeline 函数合并进 src/Function_030.py

  把各步骤的核心函数搬进 Function_030.py，按 build_shape_X()
  命名规范注册，main.py 直接调用。

  优点：文件结构更干净；缺点：改动量较大，需要处理原有 xlwings 和新 pandas
  逻辑的共存。

  ---
  我的建议

  先用方案 A（你确认 PPT
  效果满意之后再做），改动可控，风险低。需要我做的时候告诉我，我会：

  1. 读现有 main.py 结构
  2. 在最小改动原则下把调用链接进去
  3. 确保路径参数可配置（不硬编码）



那ppt模板我仍然需要精确指定在ppt中的哪几页？       ♥️ ♥️ ♥️ ♥️ ♥️ 

● 是的，目前页码是硬编码在 pipeline/01_shape_detail.py 里的：                   
  BLANK_SLIDE_IDX    = 13   # 第14页（0-based）= 空白参照页                       
  STANDARD_SLIDE_IDX = 14   # 第15页（0-based）= 标准模板页
 逻辑是：用"标准页 - 空白页 = 新增 shapes"来确定要填写的目标区域。

  ---
  如果你的模板换了页码，有两种处理方式：

  A. 直接改这两行常量（当前推荐，最简单）

  B. 整合进 main.py 时做成可配置参数（整合时一并处理）

  # main.py 入口处统一配置
  CONFIG = {
      "template": "src/Template 2.1.pptx",
      "excel":    "2025 数据 v2.2.xlsx",
      "blank_slide":    14,   # 1-based 页码
      "standard_slide": 15,
  }

  ---
  简单来说：你现在不需要动任何东西，当前模板的 14/15 页是固定匹配的。等到你确认
  PPT 效果满意、准备整合进 main.py 时，我们一并把页码做成可配置参数。




我已经混淆了。你的意思是，你已经完成了 本应该由 orchestrator.py 完成的工作？你直接将py代码都写好了是吗？  我现在可以直接调试 pipeline了？
   pipeline到底是什么意思，我是说中文翻译是什么？我经常听到这个词      
    ● 对，说清楚一点：                                                              
  ---
  我已经写好了什么：
  上一个会话里我直接把代码都写好了：

  pipeline/
  ├── ppt_pipeline_common.py   # 公共工具函数
  ├── 01_shape_detail.py       # 提取模板 shapes
  ├── 02_shape_analysis.py     # 分析角色 + 生成 prompt 规格
  ├── 03_build_shape.py        # 生成每个 shape 的文字内容
  ├── 03_build_ppt_com.py      # 用 COM 把内容写入 PPT
  └── 04_shape_diff_test.py    # 对比模板 vs 生成结果打分

  orchestrator.py 是另一套东西，是用来让 Claude AI 子 Agent
  互相协作写代码用的。你不需要它——因为代码我已经写好了。

  ---
  现在可以直接调试吗？

  可以，直接从第一步跑起：

  python pipeline/01_shape_detail.py

  ---
  Pipeline 是什么意思？

  字面翻译是流水线。

  就像工厂流水线一样，每道工序处理完交给下一道：

  提取模板形状 → 分析角色 → 生成内容 → 写入PPT → 验证结果
     Step 1    →  Step 2  →  Step 3  → Step 3B →  Step 4

  每一步的输出是下一步的输入，整体串联起来就叫 pipeline（流水线/管道流程）。









好的。另外，与标准模板相比，你生成的ppt的【优点/缺点】字数偏少，你按模板
  的字数，来精确要求gpt并构建准确的prompt。  后续我采用orchestrator.py +
  6agents 流程时，如果需要持续迭代py代码，我应该从 01-agent 开始吗？              我如何才能反馈需要改正的点，让它继续迭代代码？ 你的这套pipeline流程，能适       应其他的ppt模板吗?好像不行吧，其他模板要重新构建大量代码
  ● 三个问题逐一回答，然后修字数。                                                
  ---   
  
  Q2：orchestrator + 6 agents 怎么用来迭代代码

  不用从 01-arch 开始，直接告诉 orchestrator 做什么就行：

  # 直接描述问题，跳过规划阶段
  python orchestrator.py "矩形68/77字数偏少，请修改 pipeline/03_build_shape.py
  中的 _build_rich_prompt 函数，让GPT生成字数更贴近模板" --from-plan

  # 或者让 tester 看 diff 结果自动迭代
  python orchestrator.py "根据 diff_result.json 不达标项修复 pipeline/ 脚本"
  --from-plan --max-rounds 3

  提供反馈的方式：把你发现的问题直接写进 task
  描述，越具体越好——比如"矩形68模板225字，生成只有80字"。

  ---
  Q3：能适应其他 PPT 模板吗

  架构层面可以复用，内容层面需要重新批注：

  层: Step 1-4 脚本本身
  是否需要改动: ✅ 不用改，完全通用
  ────────────────────────────────────────
  层: shape_detail.md 批注
  是否需要改动: ❌ 每个新模板都要重新填写
  ────────────────────────────────────────
  层: 路由关键词（评分均值/100分制…）
  是否需要改动: ✅ 已经是中文通用关键词，大多数场景可复用
  ────────────────────────────────────────
  层: Excel 列名
  是否需要改动: ⚠️ 如果问卷列名变了，_col_values 的关键词可能需要补充

  总结：新模板只需运行 Step 1 生成新的
  shape_detail.md，然后手工填批注，不需要重写代码。

  ---
  Q1：修字数 — 用目标字数代替上限

  当前 prompt 只说"不超过X字"，GPT
  会偷懒输出少量文字。改为"控制在X字左右，与参考文本信息密度相近"：

● Update(pipeline\03_build_shape.py)
  ⎿  Added 7 lines, removed 1 line
      277      max_chars = budget.get("max_chars", 200)                               278      max_lines = budget.get("max_lines", 6)
      279
      280 +    # Target length = template text length (style_anchor), so
          +GPT matches
      281 +    # information density of the reference — not just "don't e
          +xceed limit"
      282 +    template_len = len(style_anchor) if style_anchor else max_
          +chars
      283 +    target_chars = max(template_len, 80)
      284 +
      285      extra = fix_notes if fix_notes else "每个分类不超过3行"
      286
      287      return (
     ...
      294          f"- 你只能分析已有数据，不能推测或编造\n"
      295          f"- 直接给出结论，不要展示分析过程\n"
      296          f"- 严格按照参考文本的格式、语调、陈述方式\n"
      292 -        f"- 总字数不超过{max_chars}字，不超过{max_lines}行\n"
      297 +        f"- 总字数控制在{target_chars}字左右（参考文本约{templ
          +ate_len}字），与参考文本保持相近的信息密度\n"
      298 +        f"- 不超过{max_lines}行\n"
      299          f"- {extra}\n\n"
      300          f"【{n}名测试者原始反馈】\n{respondent_block}\n\n"
      301          f"直接输出结论，不需要任何前言。"

● 修完了。再跑一次：                                                            
  python pipeline/03_build_shape.py

  这次矩形 68 prompt 会告诉 GPT "控制在约 225 字左右"，矩形 77 约 168             字，和模板字数对齐。












这套逻辑是我最近2星期不断迭代的结果。还有优化空间吗？我是指针对我的任务：       
高精度复制ppt格式、填充新内容，从而批量制作精美的ppt这类型任务，这个框架p
  ipeline还有优化空间吗？                                                                                                                                       ● 还有优化空间，而且有几个值得认真考虑的方向。我按影响程度排：                            
  ---                                                                             高影响

  1. 缺少批量循环（最大缺口）

  当前架构是 1个数据集 →
  1个PPT。如果你要给10款球鞋各生成一份报告，现在要手动跑10次全流程。

  优化方向：在 03b 之前加一个批量调度层——遍历 Excel 中每款产品的行，自动依次跑
  03a→03b，输出 codex_鞋款A_1.0.pptx、codex_鞋款B_1.0.pptx…

  2. 03a→03b 之间没有内容审核门

  GPT 生成内容后直接写入 PPT，没有人工确认步骤。内容错误只有到 Step 4
  才能发现，此时 PPT 已生成，还要再跑一遍。

  优化方向：03a 输出内容后，打印一份"待写入内容预览"，让你确认后再执行
  03b。对批量生产尤其重要（错一个，10个都错）。

  ---
  中影响

  3. Visual 分数验证的是"几何"，不是"视觉"

  Step 4 的 visual score 检查 shape
  的位置/尺寸数值，但不检查字体大小、颜色、加粗是否被 COM 操作破坏。AutoSize fix
   解决了尺寸问题，但段落级格式漂移仍可能被忽略。

  优化方向：在 Step 4 加一个"文字格式回读"——读取生成 PPT 的
  font_size/bold/color，与模板对比。

  4. 单 shape 更新要跑全流程

  只改一个 shape 的内容，也要重跑 03a（所有 shape 重新生成）+ 03b（全量写入）。

  优化方向：给 03a/03b 加 --shape-name 参数，支持单 shape 增量更新。

  ---
  低影响（锦上添花）

  5. Semantic 关键词是硬编码的

  "样本"、"建议"、"反馈"
  只适合当前这个篮球鞋模板。换模板（比如服装测试报告）就需要手动改代码。

  优化方向：在 01-shape_detail.md 的批注里加 semantic_keywords:
  字段，让用户配置，Step 4 动态读取。

  6. GPT 无重试 + 无成本追踪

  03a 里如果 GPT 超时或返回空内容，直接写空值进 PPT，没有重试逻辑。调用了多少
  token 也不知道。

  ---
  一句话总结

  当前框架：单次生产质量高（接近 98%+ visual），但缺批量能力和内容预审门。
  最值得投入：批量循环 + 内容预审。其余是稳定性加固。

  你目前的框架已经是「会做」的阶段，下一步是「能规模化」。












  我的原工作流程是自动 COM 打开模板ppt，然后直接在模板上修改（新增ppt页面），完成后删除模板所有页面再另存为ppt文件，这样就不会改动【模板ppt】了。你新增的功能需要融入这个流程，
  因此可能需要改写一些代码

● 明白。我需要先读懂 main.py 的 COM 流程（开头打开模板、结尾删除并另存），以及    Function_030.py 现有的工具函数。
                                      


Pipeline → main.py 深度融合：src/codex_ppt.py（无 pipeline 依赖）

 Context

 Pipeline 03a/03b 已验证可生成高质量"codex"评测分析页。
 用户要求将所有逻辑直接移入 src/codex_ppt.py，不依赖任何 pipeline/ 文件，
 确保项目复制到其他电脑后可独立运行。

 main.py COM 流程要点：
 - mc_ppt = Presentations.Open('src/Template 2.1.pptx') — 直接操作模板
 - 每个 section 的模式：mc_ppt.Slides(N).Copy() → mc_ppt.Slides.Paste(X)
 - slide_all 记录初始模板页数；最后 for p in range(1,slide_all+1):
 mc_ppt.Slides(1).Delete()
 - 新功能必须嵌入此 Copy→Paste 模式，Slide(15) 是"标准评测模板页"

 ---
 修改范围

 ┌─────────────────────┬────────────────────────────────────┐
 │        文件         │                改动                │
 ├─────────────────────┼────────────────────────────────────┤
 │ src/codex_ppt.py    │ 新建，含所有逻辑，零 pipeline 依赖 │
 ├─────────────────────┼────────────────────────────────────┤
 │ main.py             │ 2 处微改：1 行 import + ~8 行调用  │
 ├─────────────────────┼────────────────────────────────────┤
 │ src/Function_030.py │ 不改                               │
 ├─────────────────────┼────────────────────────────────────┤
 │ pipeline/ 任何文件  │ 不改                               │
 └─────────────────────┴────────────────────────────────────┘

 ---
 注意事项

 1. Slide(15) 索引：main.py 始终向末尾追加，原始模板 slide 15 不受影响。
 2. COPY_PASTE_DELAY = 1.5s：与 03b 和 main.py 其他 section 保持一致。
 3. GPT_5 guard：GPT_5 = None 时 _call_gpt 直接返回 fallback，不报错。
 4. openpyxl 读图：mc_sht.book.fullname 取 Excel 完整路径，避免硬编码；
 openpyxl 只读模式不干扰 xlwings 已打开的 COM 实例。
 5. _extract_score_means 需处理"所有在范围内"的数值列（含 backup_numeric）
 — 直接从 ppt_pipeline_common.py 复制完整版本，保留注释中的 BUG FIX 说明。

 ---
 验证方式

 # 1. 运行 main.py（Excel 已打开），确认无报错，生成 PPT 含 codex 分析页
 python main.py

 # 2. 打开生成的 PPT：
 #    - codex 页出现在【5.5】问卷汇总页 之后、【6】结论页 之前
 #    - 条形图有评分数据、矩形 68/77 有优缺点文字、颜色正确（蓝/红）
 #    - 图片 39 显示鞋款图片（如果 Excel 内有嵌入图片）

 # 3. GPT 关闭时（mc_gpt='n'）：文字 shape 显示 fallback 文本，不报错
╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌


原来pipeline 流程运行过程中会产生很多md / json
  文件，是为了调试ppt效果、不断迭代修正用的。  pipeline 的工作流程可以保持
  不变（继续产生中间文件），但为了融入到我的系统，你可能需要将这些已调试好        的 md/json文件，同样挪到src文件夹中？                                           或者直接写入codex-ppt.py文件？你认为哪种比较合适？我之前的方案是直接写死        在py文件中。毕竟现在效果我已经基本满意了    




 直接写死在 codex_ppt.py 是更合适的方案。理由如下：

  ---                                                                             
  推荐硬编码的原因：
  1. pipeline 的中间文件是调试工具，不是运行时依赖
    - 01-shape_detail.md, 02-budget.json 等文件的价值是"让你迭代调整"
    - 你已经满意结果了 → 这些文件的使命已完成，可以固化
  2. 真正的自包含
    - 复制到新电脑：只需 src/ 文件夹，不需要 pipeline-progress/ 目录
    - 不存在"文件找不到"的运行时风险
  3. pipeline 流程可以继续独立存在
    - 如果将来换了新模板/新问卷，重新跑 pipeline → 验证效果 → 再更新
  codex_ppt.py 中的常量
    - 两者解耦：pipeline 是"研发工具"，codex_ppt.py 是"生产代码"

  ---
  唯一需要注意的：style anchor（风格锚点）

  02-prompt_specs.json 里每个 shape 有一段"参考文本"用于 GPT
  语调。这段文字可能较长。

  两种处理方式：
  - A（推荐）：在 codex_ppt.py 里作为字符串常量内联，每个 shape 一行 —
  已有的模板原文本身不长
  - B（简化）：gpt_prompted 的两个 shape（矩形 68/77）直接把 style_anchor
  留空，GPT 根据数据自由发挥 — 你调好的 prompt 本身已足够强，style_anchor
  只是锦上添花

  我倾向于 A，把 style_anchor 也硬编码进 CODEX_SHAPES 的每条配置里。这样行为
  100% 可预测。







♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 
生成高精度ppt这么难吗？为什么我花了这么长时间才调试ok？

● 确实很难，原因是几层叠加的：                                                  
  COM 本身就反人类
  - PowerPoint COM 对象行为不一致，getattr 在普通 Python 对象上返回 None，在 COM
   上直接抛异常                                                                  
  - Clipboard 操作需要 sleep buffer，否则 Paste 拿到空内容
  - Chart 数据写入有两条路（ChartData.Workbook vs
  SeriesCollection），前者会触发"链接文件"弹窗，后者才是正解——但文档几乎不提

  "高保真"的门槛比想象高得多
  - 用代码新建 shape 根本达不到模板级别的字体/颜色/阴影精度
  - 唯一可靠的方案是 Clone 模板页再覆写内容，但这又要求模板 shape
  的命名规范（否则找不到）
  - python-pptx 看起来方便，实际上对复杂格式支持很差，COM 才是唯一真正能控制
  PowerPoint 的路

  Windows 环境叠加了额外噪音
  - MINGW64 bash 的 double-print 问题掩盖了真实 exit code
  - sys.stdout.reconfigure() 在 Python 3.12 / Windows 上行为异常
  - xlwings App()/quit() 的 COM 事件循环坑（你的 MEMORY.md
  里记录了好几个这类问题）

  调试周期长
  - 每次改动都要跑完整个 main.py 才能看到 PPT 结果，没有快速反馈
  - pipeline
  的分步拆解（01→02→03a→03b→04）本质上就是为了解决这个问题——让每步可以独立验证

  总结一句话：PowerPoint COM + Windows 环境 +
  高视觉还原度，三个难点叠在一起，调试慢是正常的，不是你的问题。











我计划将目前这套成功经验，推广到更多ppt模板。你说的这些避坑指南，帮我更新
  到claude.md / memory 和其他需要沉淀知识的地方，避免后续继续犯错。另外，我
  希望后续使用agent自动化流程【orchestrator.py +                                  agents】来完成新ppt模板的开发；并沿用pipeline                                   工作流程来完成类似ppt模板的自动化生成。你帮我更新这些避坑指南，然后告诉我       更新了哪几个文件即可（简要解释下）




  更新了 3 个文件，简要说明：                                                   
  ---
  1. memory/MEMORY.md
  - COM Gotchas 新增三条：图表用 SeriesCollection、图片等比缩放插入、Clone        幻灯片的正确 sleep 节奏
  - 新增 "高保真 PPT 开发模式" 章节：核心是"Clone 模板页而非新建 shape"，以及
  codex_ppt.py 作为参考实现
  - 新增 "新模板扩展路线" 章节：orchestrator 写代码 → pipeline 验证 → main.py 8
  行集成的三步路线

  2. .claude/CLAUDE.md
  - 新增 "高保真 PPT COM 开发规范" 章节：一张 COM
  陷阱对照表（属性读取/图表/图片/Clone 四个场景）+ 新模块代码结构
  - 新增 "新模板扩展流程" 章节：三阶段（orchestrator 开发 → pipeline 验证 →
  main.py 集成），每步有具体命令

  3. memory/pipeline-debug.md
  - 新增 Bug 5（图表 ChartData.Workbook → SeriesCollection）、Bug 6（图片拉伸 →
  等比缩放代码片段）、Bug 7（新旧函数双重生成同一页）
  - 扩展"新模板上线流程"：区分 pipeline 验证路线、main.py 集成、以及
  orchestrator+agents 路线三条路径







我还有许多新的ppt模板需要开发（写新代码），为了提高成功率、尽可能高效地迭
  代出优质代码，来实现对新的ppt模板的复制和批量生成，我需要注意哪些细节呢？
  你已将避坑指南写入 claude.md/memory等文档，用来指导agents避坑，那我需要在哪些方面避坑、实现更高效的开发呢？   


  ------------- 结束语 -------------
♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ ♥️ 

从你过去这个项目的完整经历来看，效率瓶颈主要集中在三个阶段：模板准备 → shape    批注 → 验收迭代。以下是每个阶段的核心注意事项。

  ---
  一、PowerPoint 模板准备（最容易忽略）                                         
  这是整个流程的地基，做不好后面都白费。

  必须提前在 PPT 里命名所有 shape
  在 PowerPoint → 选中 shape → 格式 → 选择窗格 → 重命名。
  - 命名规则要有语义，如 score_10pt、grade_letter、chart_main
  - 装饰性/背景 shape 不需要命名（strategy=skip）
  - 图表 shape 必须是 PPT 内置图表（不是图片截图），否则无法用 SeriesCollection
  写数据


  确认哪一页是"标准模板页"
  - 当前项目：Slide(15) = 标准评测页，Slide(14) = 空白页
  - 新项目开始前先数清楚，并在 codex_ppt.py 和 main.py 里写死正确的页码
  - 标准模板页的 shape 布局要完整（包含所有需要写入的 shape），不要用空白页

  图表 shape 要预先有数据
  - 新建一个 PPT 图表时，里面必须有占位数据（否则 SeriesCollection(1) 不存在）
  - 占位数据的行列数和最终数据要一致，否则 SeriesCollection 写入可能 index 越界

  ---
  二、Excel 数据结构（直接影响策略路由）

  列名要和 CODEX_SHAPES 里的 params.column / _col_values 关键词匹配
  - _col_values 用关键词包含匹配，所以"鞋款名称"能匹配"待测鞋款名称"，但"shoe
  name"不行
  - 新模板开发前先列出所有需要提取的列名，确认 Excel 里的实际列名能被关键词匹配

  每次开发新模板前手动清理 Excel 的 temp 区域
  - 用 main.py 跑完一次后，检查问卷 sheet 数据区以下是否干净
  - 不干净就手动删行，避免 CurrentRegion 污染

  ---
  三、shape_detail.md 批注（最关键的人工环节）

  这是 agents 工作质量的核心输入，批注质量直接决定 PPT 质量。

  每个 shape 批注前先问自己三个问题：

  ┌───────────────────────────┬──────────────────────────────┐
  │           问题            │           对应字段            │
  ├───────────────────────────┼──────────────────────────────┤
  │ 这个 shape 的内容从哪来？   │ 内容来源                     │
  ├───────────────────────────┼──────────────────────────────┤
  │ 用什么方式生成？            │ strategy                     │
  ├───────────────────────────┼──────────────────────────────┤
  │ 有什么格式限制？            │ budget.max_chars / max_lines │
  └───────────────────────────┴──────────────────────────────┘

  strategy 选择原则
  有固定值（标题/装饰）      → template_direct 或 skip
  从 Excel 提取数值计算      → score_10pt / grade_letter / sample_aggregation
  从 Excel 提取特定列文字    → extract_column
  需要 GPT 归纳文字反馈      → gpt_prompted（必须指定 filter: 优点/缺点）
  有图表                    → mean_extraction
  有图片                    → extract_image

  budget 要从模板实测
  - 不要估，直接在 PPT 里往 shape 里打字，看装得下多少字、多少行
  - max_chars 设成实测值的 90%（留余量给 GPT 的标点符号）

  grade_letter 和 score_10pt 顺序问题
  - 已在代码里固定：grade_letter 检查在 score_10pt 之前
  - 批注时不需要关心，但要确认 Excel 评分列的量程（0-5 还是
  0-10），会影响自动换算

  ---
  四、pipeline 验证节奏（避免返工）

  不要跑完 01→04 再看结果，要逐步验收：

  Step 01 → 打开 shape_detail.md，逐个确认 shape 信息是否正确
  Step 02 → 看 strategy 路由是否符合预期（特别是 chart/image/skip 分类）
  Step 03a → 看 build_shape_content.json，核对每个 shape 的文字内容
  Step 03b → 打开生成的 PPT，肉眼检查布局和颜色
  Step 04 → 看 diff_result.json，Visual < 98 的 shape 看 fix_ppt.md 的建议

  每步出问题就在当前步修，不要跳过。Step 03a 的文字内容如果不对，Step 03b/04
  再跑也是白跑。

  ---
  五、agents 开发效率

  给 architect agent 的输入要足够具体

  不要只发 shape_detail.md，要同时告诉他：
  - 这个模板的 Slide 编号（克隆哪一页）
  - 每个 shape 的 strategy（从 shape_detail.md 批注来）
  - GPT 需要的 focus（优点/缺点/其他）
  - budget 数字（max_chars, max_lines）

  developer agent 的验收标准要量化

  不要说"生成正确的 PPT"，要说：
  - 矩形 11 显示 "8.50/10" 格式的评分
  - 矩形 12 显示字母等级 "A-"
  - 图表 44 有 7 条数据，数值在 6-10 之间
  - 矩形 68 文字 ≤ 270 字，关键词有蓝色高亮

  tester agent 的门禁要卡紧
  - Visual < 98 → 必须返回 developer 重做，不允许合并
  - 允许 tester 直接查 04-diff_result.json 做决策，不需要人工介入

  ---
  六、一次新模板开发的标准节奏

  1. PPT 里命名好所有 shape（30分钟）
  2. 确认 Excel 列名能被关键词匹配（10分钟）
  3. 跑 Step 01，填写 shape_detail.md 批注（20分钟）
  4. 让 architect agent 生成策略矩阵 + developer 生成 src/xxx_ppt.py
  5. 跑 Step 02-04 验证，Visual < 98 的问题让 developer 迭代
  6. 验收通过后，main.py 加 1 行 import + 8 行调用

  最容易节省时间的地方：Step 3 的批注质量。批注写得越准，agents
  迭代次数越少。这是你作为人工环节最值得投入时间的地方。






















───────────────────





claude-yk --resume 



 python pipeline/01_shape_detail.py  ------------------- 【建筑承重墙施工】
  提取模板 shape 信息 → 扫描 Template 2.1.pptx，把每个 shape
  的位置、字体、文字、类型写入 JSON 和 MD 文件。MD 文件是给你填批注用的（"这个
  shape 应该填什么、用什么策略"）。 
  ---
  python pipeline/02_shape_analysis.py ----------------- 【室内设计师】
  分析 + 分配策略 → 读 Step 1 的 JSON + 你填的批注 + Excel 数据，为每个 shape
  决定：角色是什么（title/chart/long_summary…）、用什么策略生成内容、字数/行数预
  算是多少。

PowerPoint 模板准备（最容易忽略）                                         
  这是整个流程的地基，做不好后面都白费。

  必须提前在 PPT 里命名所有 shape
  在 PowerPoint → 选中 shape → 格式 → 选择窗格 → 重命名。
  - 命名规则要有语义，如 score_10pt、grade_letter、chart_main



为shape添加名称，有什么用途？  
你所谓的给shape命名，其实是为了更好地定义shape，帮助后面的pipeline精确识别shape内容。
shape-detail.md 里我也会干预并手工给各种shape写注释。这个其实和为shape命名是同一个道理吧？
你新增的这个pipeline，似乎替代不了人工识别shape、人工为shape添加定义和描述的功能吧
因为简单的py程序，根本无法分析 shape 的内容和源数据Excel之间的关系，更加无法推测出shape内容的生成逻辑
这道工序（定义shape、推测shape内容的生成逻辑），目前仅能由人工（我）来完成
这道工序即使交给agents，或者交给你，估计完成的精确度应该也不够，还是得人工（我）来审核校对




我最终目的是希望实现agent自动迭代代码，让py代码的模板精度通过迭代越来越高
那么每一轮迭代，agents对ppt shape 内容的识别准确度要越来越高

































--------------------
Mc-ppt 标准工作流程总结
--------------------



我有一个现成的脚本，代码为【根目录main.py + /src】，我的脚本解析结果在repo-sanc-result.md中。

你现在是一个极度严谨的 PowerPoint COM 自动化专家 + Office自动化专家，
你精通xlwings 和 pywin32 对应的office Application COM API（PowerPoint / Excel 的 COM API）
在控制Excel时，你必须使用 xlwings + COM API 二者结合的方式，实现精准控制；
在控制PowePpoint时，你必须使用 pywin32 + win32com.client 通过真正的 PowerPoint.Application COM 接口来操作，
注意，**严禁使用 python-pptx** ！！！


当前repo中的根目录包含几个重要文件： 
【0-空白 ppt 模板.pptx】：仅保留基本的框架和装饰元素 
【1-标准 ppt 模板.pptx】：最终要实现的ppt效果（内容、排版都很完美） 
【2025 数据 v2.2.xlsx】文件中的【问卷sheet】：唯一的数据源，用来生成标准ppt模板中的文字、图表等所有内容
【Template 2.1.pptx】 已包含 【空白 ppt 模板】和【标准 ppt 模板】，分别位于【Template 2.1.pptx】文件 的第14页 和 15页


现在需要在原来的基础上开发新的功能。


------第1步----- 
首先，你需要写一个测试文件（01-shape-detail.py），分别读取【Template 2.1.pptx】中的（14页）【空白 ppt 模板】和（15页）【标准 ppt 模板】
 注意，你必须遍历读取所有元素shape，并识别出【标准 ppt 模板】中新增的元素（哪些元素是标准模板中新增的、空白模板中不存在的元素？）
  识别出新增的元素后，详细读取这些元素的所有信息，包括：

{
页面编号（slide_index），名称（name，如果名称为空，你根据它的特征，简短给它生成一个名字，并记住它）、
位置（top & left）、 大小（height & width）、 
字体（并检查该字体class文件是否存在，如果不存在就需要更新class_030.py文件，遵循class文件中的惯例新增）、 
属性（shape.Type，文本/图片/chart？）、 
内容（文字内容）
是否在组里（Parent.Group） }

 tips: tool 1 - shape top&left v1.1.py是我之前使用的一个小工具，你可以参考它，但我本地拥有 office 环境方便调用，你可能需要自己想办法
然后，将这些新增元素的信息，汇总成01-shape-detail.md（假设你识别出一共有9个新增元素），详细列出上面所有信息，并提供给我






------第2步----- 
分析 shapes 内容：
如果你成功识别出了这些【新增元素】，这些新增元素的内容都是源于【2025 数据 v2.2.xlsx】文件中的【问卷sheet】里的内容 
你需要写一个（02-shape-analysis.py）文件，通过它来逐一对比每个shape【新增元素】与【源数据】之间的关系
    （标准ppt模板里新增的9个shape的内容，完全源于 问卷sheet）
并输出一份标准文档：02-shape-analysis.md ，以便于其他agents根据这个标准文档，构建对应的函数工具，实现内容的精确提取和重构
注意事项：
    我之前的脚本代码【根目录main.py + /src】，里面的所有工具你都可以调用，不需要重新造轮子
    例如，你可以动态加载 src/Function_030.py 里的 GPT_5 与 extract_info，用于 AI 总结与复杂字段提取（并带兜底逻辑）






------第3步----- 
构建 shapes 内容：
    根据02-shape-analysis.md的要求，并结合我之前的脚本代码【根目录main.py + /src】里面的所有工具
    你帮我写一个py测试脚本（03-build_shape.py），主要包含9个函数（假设你识别出一共有9个新增shape元素）
    这9个函数的主要功能，是从【源数据-问卷sheet】中提取对应的信息，加工成对应9个shape的精确内容
    
    再次提醒，我原来脚本中包含【GPT-5通讯函数】，可以用来连接gpt服务器，实现问卷信息的总结和提取，你仅需构建对应的 prompt 即可
    你可以在新的脚本中（02-build_ppt_com.py）重复利用（import）这些现成的函数，不需要重写。





构建 ppt文件：
    另外，完成9个shape内容的构建后，
    你的这个（03-build_ppt_com.py）这个脚本，
    还需要完全按照【标准 ppt 模板】的样式，重新生成1个ppt文件，命名中包含【你的名字 + 版本号】
    （假设你是codex，那命名就叫做【codex 1.0.pptx】）
    你必须严格按照我的工作流程，来完成你的ppt制作（注意： 严禁使用 python-pptx ！！）：

    1. 启动 PowerPoint.Application（Visible=False）
    2. 创建一个新 Presentation 文件（假设你是codex，那命名就叫做【codex 1.0.pptx】）
    3. 复制【标准ppt模板（Template 2.1.pptx 第15页）】，source_slide.Duplicate() 或 Copy + Paste 的方式完整克隆 
        并保留源格式（这是关键：直接基于完美模板）（保留标准ppt模板 所有美工、动画、母版效果）
    4. shape 元素操作
    - 克隆后，根据shape_detail_com.json 精准找到9个目标形状（假设你识别出一共9个shape需要更新）
    - **只做内容替换**：
            - 文本形状 → 更新 .TextFrame.TextRange.Text（保留原有字体、颜色、段落格式）
            - 图表形状 → 读取【2-问卷模板.xlsx】对应数据，并通过 Chart.ChartData 或直接修改ppt-chart的 Series.Values / Categories 来更新数据（**保留图表所有样式、颜色、数据标签、轴格式**）
            - 数据仅需仅需简单的统计：例如均值、频率、汇总等，无需引入numpy等大型计算工具
            - 图片（如有）→ 如果需要替换则使用 .PictureFormat 或 Replace
    5. 所有数据必须来自 Excel（建议使用 xlwings 打开 Excel 读取）




------第4步----- 
你完成ppt文件生成后，需要再写一个测试脚本（04-shape_diff_test.py），
使用 COM 打开【标准模板（位于Template 2.1.pptx第15页）】和【codex.pptx（第1页，你生成的ppt应该只有1页）】，
逐 shape 对比【位置、大小、字体、内容（例如文字数量是否极为接近）颜色、图表类型等所有关键属性，】
给出详细差异报告。如果差异明显，必须自动调整代码并重新生成。

**最高优先级**：
- 视觉保真度必须达到 98%+（几乎和手动复制 slide 后改数据一样）
- 严禁使用 python-pptx 的任何函数
- 代码必须健壮、带错误处理和资源清理
- shape_detail_com.json 精确定位需要更新的shapes
- 生成fix-ppt.md文档，详细标注差异点，以及修复建议（修改哪个函数？如何修改？）



------第5步----- 
假设上一步的测试结果不满意，需要重复上面的【调整ppt - 测试ppt】步骤，
生成新的带编号ppt文件【codex 1.1/1.2/1.3.pptx 】，不要覆盖掉之前的ppt文件
直至你生成的最新ppt文件能够通过（04-shape_diff_test.py）的验证（视觉保真度必须达到 98%+（几乎和手动复制 slide 后改数据一样）
当然，我也会设置一个最大重复限制（例如最大重复3轮）



------第5步----- 
我作为用户，会人工检查最终生成的ppt文件，肉眼对比最新版【codex.pptx】和【标准模板】的差异
根据人工对比效果，再考虑是否继续下一轮的优化（直接进入【调整ppt - 测试ppt】步骤）














------最后的步骤：整合（待ppt效果ok之后再说，不急着整合）----- 
接下来你需要继续完成深度整合，将上面的py代码整合进入原来的【根目录main.py + /src】脚本程序中。注意：
接下来，我希望直接在 main 文件中你上面的功能，
插入到（【5】实战测评部分 ）和（【6】结论部分）之间，插入位置大概位于main.py的 785行-800行之间。























----------- codex：结构优化建议 -------------

接下来讨论我最新的工作流程（最近一次发给你的【我期望的目标】那一大段文字，注意是最新版的）：
我已经尽可能将制作ppt的过程拆解成详细步骤，为什么效果似乎还是不理想？  
你认为哪几个环节可以需要继续拆分成更细的任务包？   
你给出你的调整建议，然后帮我更新成一份 new-ppt-workflow.md



你这个问题非常关键：不是“流程有没有”，而是“流程有没有足够细的质量控制点”。

我已经按你的最新版目标，给出拆分建议并落到新文档 new-ppt-workflow.md 了。

我对“为什么效果不理想”的结论（简版）
最核心是 4 个缺口：

shape 匹配策略不够稳（仅 name/顺序容易漂移）。

内容构建缺少“可读性预算”（字数/行数/bullet 没被硬约束）。

diff 缺少语义与可读性门槛（只看样式不够）。

迭代没有“自动修复路由”，导致收敛慢。

我建议继续拆细的任务包（已写进文档）
重点新增拆分：

Step1：指纹构建 / 组内解析 / 模板漂移门禁。

Step2：字段映射 / Prompt规格 / 可读性预算。

Step3：9函数构建 / 内容校验 / COM写入 / 写后回读。

Step4：视觉diff / 可读性diff / 语义diff 三层门禁。

Step5：修复路由 / 收敛控制。

新流程产物
我还把新版标准流程整理成 Phase 0~4（20个任务），并补了建议新增产物与验收门槛（Visual≥98%、Readability≥95%、Semantic=100%）。









非常好，你按你提供的最新 new-ppt-workflow.md 和 最新拆分计划，
重新调整所有代码（可以删除旧文件，提供全新的一批 py 脚本文件给我）。
我期望仍然拥有统一的控制入口，到时候我直接运行第一个文件（假设是00-ppt.py），
就能实现你上面的全新工作流。至于过程控制需要拆分成几个文件，你自己决定（我认为可以适当简化）。  

另外，
我之前的脚本代码【根目录main.py + /src】，里面的所有工具你都可以调用，不需要重新造轮子
例如，你可以动态加载 src/Function_030.py 里的 GPT_5 与 extract_info 等一些列函数，
用于 AI 总结与复杂字段提取（并带兜底逻辑）。
你也可以参考我的 function.py 文件中的 prompt，来精确控制 GPT_5 输出的文字内容，来匹配shape 内容的生成需要
你也可以重构新的extract_info函数，来提取更多关键信息，用于 shape内容的生成、问卷评分的提取、均值计算、频次统计等
当shape内容对比测试不通过时，
    你可能要优先调整 相关提示词（如果内容是由 GPT_5 函数生成），也可以考虑修改信息提取函数（如果内容是直接从问卷提取）



