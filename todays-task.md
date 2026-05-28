# Today's Task
## 你继续读取 todays-task.md ，完成相应的任务

最近我更新了系统skill（C:\Users\xy24\.claude\skills），

我们来回顾下 3rd-ppt-prj 的工作流程和三重混合机制：

1、拿到全新ppt模板，首先跑 orchestrator，生成 pipeline 进行文件固化；
2、移植，用developer完成代码移植，生成固化的py代码
3、针对部分细节，直接在 claude 主会话中解决最后的 corner case

现在，随着系统skill持续升级，你再根据最近几轮工作经验、踩坑历史，客观评估下上面的工作流程，然后提出优化建议（暂时先不用改代码）


--------------

 开始执行：盘点 inspect/验收 skill 对"同 shape 多字号/多颜色（含同段多
  run）"的识别/检验/修改能力。

● Explore(盘点多run样式能力)
  ⎿  Done (25 tool uses · 65.3k tokens · 10m 7s)
  (ctrl+o to expand)

● 能力盘点（同 shape 多字号/多颜色 / 含同段多 run）

  能力: 识别 run-size
  A read-selected: ✓ L126
  B inspect-ppt: ✗
  C acceptance 引擎: ✓ runs.py L32-95
  D apparel.json: —
  E apparel_ppt: ✓ L1276
  F _ppt_shared: ✗
  ────────────────────────────────────────
  能力: 识别 run-color
  A read-selected: ✓ L132-135
  B inspect-ppt: ✗
  C acceptance 引擎: ✓ runs.py
  D apparel.json: —
  E apparel_ppt: ✓ L1279-89
  F _ppt_shared: ✗
  ────────────────────────────────────────
  能力: 断言 run 差异
  A read-selected: —
  B inspect-ppt: —
  C acceptance 引擎: ✓ runs_match_template L195-214
  D apparel.json: △ 只对老 TextBox 用了
  E apparel_ppt: —
  F _ppt_shared: —
  ────────────────────────────────────────
  能力: 写入 run 差异
  A read-selected: —
  B inspect-ppt: —
  C acceptance 引擎: —
  D apparel.json: —
  E apparel_ppt: ✓ _write_two_run_label L1210-1310
  F _ppt_shared: ✗

  关键结论

  两个真空白：

  1. inspect-ppt-template —— 扫描只到 shape 级（text_preview），不吐
  runs:[{text,size,rgb,bold},...]。今天升级 read-selected-shape 之后，inspect
  反而成了"看模板靠肉眼"的瓶颈
  2. acceptance/apparel.json —— 引擎层（runs_match_template）已经能查同段 N-run
  的 size+rgb 组合，但 RR 53 / RR 55 / TextBox 50 这三个今天新升级的视觉契约
  shape 在 contract 里只有 text_contains 数据断言，没挂
  runs_match_template。引擎在等契约喂养，不在等代码升级

  性价比排序（补盲优先级）：

  ┌───────┬─────────────────────┬───────┬──────────────────────────────────┐
  │ 优先  │       补哪里        │ 工作  │               收益               │
  │  级   │                     │  量   │                                  │
  ├───────┼─────────────────────┼───────┼──────────────────────────────────┤
  │       │ acceptance/apparel. │       │ 直接守住今天的视觉契约，下次     │
  │ P0    │ json 给 RR          │ 5 行  │ GPT/数据源动了不会静默退回模板默 │
  │       │ 53/55/TB50 加       │ JSON  │ 认样式                           │
  │       │ runs_match_template │       │                                  │
  ├───────┼─────────────────────┼───────┼──────────────────────────────────┤
  │       │ inspect-ppt-templat │ ~30   │ 以后调试新模板时直接看 run       │
  │ P1    │ e 加 run 字段输出   │ 行 Py │ 矩阵，不用手动选中再跑           │
  │       │                     │ thon  │ read-selected-shape              │
  ├───────┼─────────────────────┼───────┼──────────────────────────────────┤
  │       │ _ppt_shared 抽      │       │ 仅当 zxh/yzr 也要做同段多 run    │
  │ P2    │ _write_run_styled   │ 中等  │ 时才回本，现在留 apparel 私有 OK │
  │       │ 通用工具            │       │                                  │
  └───────┴─────────────────────┴───────┴──────────────────────────────────┘

  顺手发现的细节

  - contract 已有 runs_match_template 能力 → 不是"引擎不够"，是"契约没写"。这和
  todays-task 末尾"如果以后再加类似 '同段大数字' 的视觉契约，需要在 acceptance
  contract 里加 run_size_differs_in_paragraph
  类断言"的结论方向一致，但比你预期的乐观：根本不用加新断言键，直接用现成的
  runs_match_template 就行
  - read-selected-shape 升级到 --full 之后，inspect
  反而是体系里最旧的一环，跟今天的"单源真相"整合方向逆行（系统 skill
  升级了，模板盘点能力没跟上）

--------------
今日改动总结：

  改动清单（6 处文件）

  层级: 引擎
  文件: ~/.claude/skills/ppt-acceptance-check/layers/runs.py
  改动: 加 runs_match_signature check（内嵌 expected_runs，不依赖
    template，默认过滤 CR/空白 run）
  ────────────────────────────────────────
  层级: 引擎文档
  文件: ~/.claude/skills/ppt-acceptance-check/SKILL.md
  改动: check 表加新行 + 标 runs_match_template 盲区
  ────────────────────────────────────────
  层级: 项目契约
  文件: acceptance/apparel.json
  改动: RR 53 / RR 55 各加 1 条 must_fix runs_match_signature
  ────────────────────────────────────────
  层级: Inspect skill
  文件: ~/.claude/skills/inspect-ppt-template/inspect_ppt_template.py
  改动: 加 extract_paragraphs() + --full 参数 + MD 详情渲染（同
    read-selected-shape --full 约定）
  ────────────────────────────────────────
  层级: Inspect 文档
  文件: ~/.claude/skills/inspect-ppt-template/SKILL.md
  改动: 加 --full 章节 + description 更新
  ────────────────────────────────────────
  层级: 项目 memory
  文件: .claude/memory/feedback_acceptance_gate.md
  改动: 加"2026-05-28 续：runs_match_template 模板=旧/代码=新盲区（红旗 5）"

  闭环验证（一个契约跑 2 次）

  ┌────────────────┬───────┬────────────────────────────────────────────────┐
  │      时点      │  L3   │                      含义                      │
  │                │ 结果  │                                                │
  ├────────────────┼───────┼────────────────────────────────────────────────┤
  │                │ 5     │ 旧 5 条 runs_match_template（模板=旧、shape=旧 │
  │ Pre-overwrite  │ PASS  │  → 全 PASS 假阳性） + 新 2 条                  │
  │                │ + 2   │ runs_match_signature（FAIL 抓到 silent         │
  │                │ FAIL  │ regression）                                   │
  ├────────────────┼───────┼────────────────────────────────────────────────┤
  │                │ 7     │ 真跑 apparel_ppt.py --overwrite-slide 9 后     │
  │ Post-overwrite │ PASS  │ new_seq 真的变 [white-11pt, white-24pt]，不是  │
  │                │ + 0   │ hardcode 自证                                  │
  │                │ FAIL  │                                                │
  └────────────────┴───────┴────────────────────────────────────────────────┘

  知识固化点（沉淀进 memory）

  - 红旗 5（shape 名错位）：runs_match_template 会把"代码没动
  shape"和"代码改对了 shape"都算 PASS——出现条件是"模板=旧、代码=新"
  - 选用规则：模板=期望态时用 runs_match_template；视觉升级超模板态时用
  runs_match_signature + 内嵌 expected_runs
  - expected_runs 取值来源：inspect-ppt-template --active --slides N --full
  跑出目标态后直接复制字段；绝不 hardcode 回读自证（红旗 4 已封禁的延续）
  - _walk_runs 合并同属性 run 不是 bug：是 L3 验收用的视觉等价建模，inspect
  走单独的 extract_paragraphs 保留原始 PPT paragraph.Runs() 结构
