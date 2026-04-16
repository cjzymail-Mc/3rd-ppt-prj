# codex-work1（结构化深扫版）

## 1) 扫描范围与方式
- 扫描时间: 2026-04-16 09:21:25 CST (+0800)
- 仓库路径: `/mnt/d/Technique Support/Claude Code Learning/3rd-ppt-prj`
- 扫描对象:
  - 关键文档: `INSTRUCTION.md`、`feature03-transplant/*`、`[feature02-self-chek]/*`、`.claude/CLAUDE.md`、`.claude/memory/*`、`.claude/agents/*`、`.claude/commands/*`
  - 关键源码: `orchestrator.py`、`pipeline/*.py`、`src/yzr_ppt.py`、`src/zxh_ppt.py`、`src/Function_030.py`、`Main.py`
- 目标: 从“仅分支状态扫描”升级为“可执行的代码结构理解快照”。

## 2) 仓库状态快照（当前）
- 当前分支: `claude-zxh`
- 最近提交: `a3343d9` (`stash 暂存 杨祖锐的ppt模板，后续找机会再来移植到 src`)
- Git 状态:
  - `status_entries=106`
  - `tracked_changed=50`
  - `untracked_entries=56`
  - `modified=27, deleted=23, renamed=1, untracked=56`
- 规模快照:
  - `tracked_files=82`
  - `all_files(不含.git)=227`
  - `all_dirs=265`

结论: 工作区是“重度迭代中”状态，不适合作为干净基线直接发布。

## 3) 项目主结构（已确认）
本仓库目前并行存在两条执行主线：

### A. `orchestrator.py + pipeline/*`（当前主干）
- 这是当前迭代的主工作流（步骤化、可回退、含自检和修复）。
- `orchestrator.py` 不是纯菜单壳：已经实现 **Pipeline-first + Claude fallback**。

### B. `Main.py + src/*`（遗留主程序 + 移植集成）
- `Main.py` 是历史主程序（大量 COM 操作、顺序式构建页面）。
- 已接入模板路由:
  - `ask_template_choice()` in `src/Function_030.py`
  - `make_codex_slide()` in `src/yzr_ppt.py`
  - `make_zxh_slide()` in `src/zxh_ppt.py`
- 这条线用于把 pipeline 能力移植回 `/src` 体系。

## 4) 运行链路理解（代码级）

### 4.1 `orchestrator.py` 实际流程
`main()` -> 选择账号 -> 选择 template/xlsx -> 菜单 `0~3` -> `PPTOrchestrator.run(step)`。

`_run_step(step)` 关键逻辑：
1. Step3 前置处理: Excel prompt 同步检测（变更则补跑 `03a --execute-prompts`）
2. 直接跑 pipeline 脚本（`subprocess.run`）
3. 运行自检（step1/2 用 `pipeline/self_check.py`，step3 读 `03b-self_check_report.md`）
4. Step2 自动修复（结构约束注入 + 重跑 GPT）
5. 严重度门禁:
  - 轻微问题放行
  - 严重问题触发 LLM 修复
  - Step3 内容级严重问题会保存 `03-feedback_to_step2.json` 并回退 Step2

`_call_agent()` 关键点：
- 通过 `claude -p ... --append-system-prompt ...` 调用 `.claude/agents/step*.md`
- 失败时带 `REPAIR MODE` 上下文，要求 agent 跳过 Attempt1，直接做修复。

### 4.2 `pipeline/*` 职责分层
- `01_shape_detail.py`: 模板差分提取 shape（当前代码基于 slide 1 vs 2）
- `01b_auto_annotate.py`: 按规则自动填 strategy/params（替代慢速逐 shape LLM）
- `02_shape_analysis.py`: 角色推断 + prompt 规格 + readability budget + output_contract 解析
- `03a_build_shape.py`: 两阶段内容构建
  - `--assemble-only` 组 prompt
  - `--execute-prompts` 调 GPT（支持 Excel prompt 回写编辑）
  - 含动态列分类、`clamp_text()`、关键词/结构约束
- `03b_build_ppt_com.py`: COM 写入 PPT + 四维自检 + 自动修复循环（最多2次修复）
- `self_check.py`: step1/2 自检函数库
- `ppt_pipeline_common.py`: 路径常量、COM 安全访问、Excel 读写、prompt 写回、截图SSIM、自检报告生成

### 4.3 `src/*` 当前移植状态
- `src/yzr_ppt.py`: 已从旧 `codex_ppt.py` 重命名并升级；公开 API `make_codex_slide()`
- `src/zxh_ppt.py`: 新增模板模块；公开 API `make_zxh_slide()`；包含 layout 矫正和 section 染色增强
- 两者都内置:
  - GPT 生成 + `clamp_text`
  - `\n -> \r` 写入转换
  - `微软雅黑` 强制字体
  - 关键词染色

## 5) “当前主要工作由 Claude 完成”的证据
从文档与代码双侧可确认，当前主干协作是 Claude-first：

1. `.claude/agents/step1-analyzer.md / step2-architect.md / step3-builder.md` 明确了步骤职责和 Attempt1/Attempt2 机制。
2. `.claude/commands/step1.md~step3.md` 已建立标准化入口说明。
3. `.claude/memory/feedback_hybrid_workflow.md` 明确要求: Pipeline先行，LLM精调。
4. `orchestrator.py` 代码层直接调用 Claude CLI，并将失败上下文注入 agent 修复。
5. feature 文档（如 `[feature01-orch-update]`、`[feature02-self-chek]`、`feature03-transplant`）均围绕该工作流持续迭代。

结论: 当前版本的“流程设计与改造主线”确实由 Claude 工作流驱动，Codex 目前更适合承担结构梳理、代码修复、落地执行。

## 6) 发现的关键不一致/风险
1. 文档与代码存在页码描述不一致:
  - 文档有“14/15页模板”说法；`pipeline/01_shape_detail.py` 当前是 `slide 1 vs 2`。
2. 模板默认名与现有文件名有偏差风险:
  - 默认常量仍含 `standard and empty template.pptx`，当前 `template/` 实际主要文件是 `empty and standard*.pptx`；
  - 走 orchestrator 可通过模板选择规避，手动直跑 pipeline 时需注意环境变量。
3. 仓库混合了“流程代码 + 产物文件 + 历史草稿”，上下文噪音高。
4. `pipeline-progress/04-*` 相关文件大量删除，若后续任务依赖旧第4步验收，需要先确认新验收链路。

## 7) 当前可直接接手的任务入口
基于上述理解，后续任务可按三类下发：
1. **流程层**: `orchestrator.py` 的回退策略、门禁规则、模板选择逻辑。
2. **内容层**: `02/03a` 的 prompt 结构、约束注入、数据提取策略。
3. **写入层**: `03b` + `src/yzr_ppt.py` + `src/zxh_ppt.py` 的 COM 写入、自检与样式一致性。

---

这份版本可作为后续任务派工基线使用（已经从“文件状态快照”升级为“代码结构快照”）。
