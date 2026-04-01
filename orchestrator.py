# -*- coding: utf-8 -*-
"""
orchestrator.py - PPT 专业制作团队调度系统 (v2)

4-agent 固定工作流:
  Analyst → PAUSE → Builder → Reviewer → [Developer] → 循环

Agents:
  analyst   - 运行 Step 1 + 自动填写 xlsx 批注
  builder   - 运行 Steps 2-3B 生成 PPT
  reviewer  - 运行 Step 4 验收 PPT + 根因诊断
  developer - 条件触发：修复 pipeline 代码缺陷
"""

import asyncio
import json
import argparse
import subprocess
import sys
import time
import uuid
import os
import re
from pathlib import Path
from enum import Enum
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from datetime import datetime

# Claude 账户配置目录
CLAUDE_CONFIG_DIRS = {
    'mc': os.path.expanduser('~/.claude-mc'),
    'xh': os.path.expanduser('~/.claude-xh'),
}

# Windows 控制台 UTF-8
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')


# ============================================================
# Data structures
# ============================================================

class AgentStatus(Enum):
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"


@dataclass
class AgentConfig:
    name: str
    role_file: str
    output_files: List[str]


@dataclass
class ExecutionResult:
    agent_name: str
    status: AgentStatus
    session_id: str
    exit_code: int
    duration: float
    cost: float
    tokens: int
    output_files: List[str]
    error_message: Optional[str] = None


# ============================================================
# 4 Agent configs (fixed)
# ============================================================

AGENT_CONFIGS = {
    "analyst": AgentConfig(
        name="analyst",
        role_file=".claude/agents/01-analyst.md",
        output_files=[
            "pipeline-progress/01-shape_detail_com.json",
            "pipeline-progress/01-shape_detail.xlsx",
        ],
    ),
    "builder": AgentConfig(
        name="builder",
        role_file=".claude/agents/02-builder.md",
        output_files=[
            "pipeline-progress/02-shape_analysis_map.json",
            "pipeline-progress/03a-build_shape_content.json",
            "pipeline-progress/03b-build_ppt_report.md",
        ],
    ),
    "reviewer": AgentConfig(
        name="reviewer",
        role_file=".claude/agents/03-reviewer.md",
        output_files=[
            "pipeline-progress/04-diff_result.json",
            "pipeline-progress/04-fix_ppt.md",
        ],
    ),
    "developer": AgentConfig(
        name="developer",
        role_file=".claude/agents/04-developer.md",
        output_files=[],
    ),
}


# ============================================================
# AgentExecutor — subprocess runner
# ============================================================

class AgentExecutor:
    """Execute claude -p subprocess and parse output."""

    def __init__(self, project_root: Path, max_budget: float = 10.0):
        self.project_root = project_root
        self.max_budget = max_budget

    def _parse_agent_file(self, content: str) -> Tuple[Dict, str]:
        """Parse YAML frontmatter + body from agent spec file."""
        content = content.strip()
        if not content.startswith('---'):
            return {}, content

        patterns = [
            r'^---[ \t]*[\r\n]+(.*?)[\r\n]+---[ \t]*[\r\n]+(.*)$',
            r'^---[ \t]*[\r\n]+(.*?)[\r\n]+---[ \t]*$',
            r'^---[ \t]*[\r\n]+---[ \t]*[\r\n]+(.*)$',
        ]

        metadata = {}
        body = content

        for i, pattern in enumerate(patterns):
            match = re.match(pattern, content, re.DOTALL)
            if match:
                if i == 2:
                    body = match.group(1).strip() if match.lastindex >= 1 else ""
                elif i == 1:
                    frontmatter_str = match.group(1)
                    body = ""
                    for line in frontmatter_str.split('\n'):
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            key, value = line.split(':', 1)
                            metadata[key.strip()] = value.strip()
                else:
                    frontmatter_str = match.group(1)
                    body = match.group(2).strip()
                    for line in frontmatter_str.split('\n'):
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            key, value = line.split(':', 1)
                            metadata[key.strip()] = value.strip()
                break

        return metadata, body

    def _parse_stream_json(self, stdout: str) -> Tuple[float, int]:
        """Parse cost/tokens from stream-json output."""
        if not stdout or not stdout.strip():
            return 0.0, 0

        lines = stdout.strip().split('\n')
        best_cost = 0.0
        best_tokens = 0

        for line in reversed(lines):
            line = line.strip()
            if not line:
                continue
            try:
                data = json.loads(line)
                if data.get('type') == 'result':
                    cost = data.get('cost_usd', data.get('cost', 0))
                    tokens = data.get('total_tokens', data.get('tokens', 0))
                    if cost > 0 or tokens > 0:
                        return float(cost), int(tokens)

                cost = data.get('cost_usd') if 'cost_usd' in data else data.get('cost', 0)
                tokens = data.get('tokens', 0)
                if tokens == 0:
                    tokens = data.get('total_tokens', 0)
                if tokens == 0 and 'usage' in data:
                    usage = data['usage']
                    tokens = usage.get('total_tokens', 0)
                    if tokens == 0:
                        tokens = usage.get('input_tokens', 0) + usage.get('output_tokens', 0)

                if cost > best_cost:
                    best_cost = float(cost)
                if tokens > best_tokens:
                    best_tokens = int(tokens)

                if best_cost > 0 or best_tokens > 0:
                    return best_cost, best_tokens

            except (json.JSONDecodeError, TypeError, ValueError, AttributeError):
                continue

        return best_cost, best_tokens

    def _check_output_files(self, expected_files: List[str]) -> List[str]:
        return [f for f in expected_files if (self.project_root / f).exists()]

    async def run_agent(
        self,
        config: AgentConfig,
        task_prompt: str,
        timeout: int = 600,
        session_id: Optional[str] = None,
    ) -> ExecutionResult:
        """Execute a single agent as async subprocess."""
        if session_id is None:
            session_id = str(uuid.uuid4())
        start_time = time.time()

        # Read agent spec
        role_file = self.project_root / config.role_file
        try:
            with open(role_file, 'r', encoding='utf-8') as f:
                content = f.read()
            metadata, role_prompt = self._parse_agent_file(content)
        except FileNotFoundError:
            return ExecutionResult(
                agent_name=config.name, status=AgentStatus.FAILED,
                session_id=session_id, exit_code=1, duration=0,
                cost=0, tokens=0, output_files=[],
                error_message=f"Agent spec not found: {config.role_file}",
            )

        agent_model = metadata.get('model', 'sonnet')

        # Windows cmd length workaround: write long prompts to temp file
        prompt_temp_file = None
        actual_prompt = task_prompt
        if len(task_prompt) > 4000:
            prompt_temp_file = self.project_root / ".claude" / f"prompt_{config.name}_{session_id[:8]}.txt"
            prompt_temp_file.parent.mkdir(parents=True, exist_ok=True)
            prompt_temp_file.write_text(task_prompt, encoding='utf-8')
            actual_prompt = f"请先使用 Read 工具读取文件 `{prompt_temp_file}` 获取完整任务指令，然后严格按照指令执行你的职责。"

        cmd = [
            "claude", "-p", actual_prompt,
            "--append-system-prompt", role_prompt,
            "--output-format", "stream-json",
            "--verbose",
            "--model", agent_model,
            "--max-turns", "20",
            "--max-budget-usd", str(self.max_budget),
            "--session-id", session_id,
            "--no-chrome",
            "--dangerously-skip-permissions",
        ]

        # Progress indicator
        async def progress_indicator(agent_name: str, start: float):
            indicators = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
            idx = 0
            while True:
                elapsed = time.time() - start
                print(f"\r      {indicators[idx]} {agent_name} 工作中... ({elapsed:.0f}s)", end="", flush=True)
                idx = (idx + 1) % len(indicators)
                await asyncio.sleep(1)

        try:
            env = os.environ.copy()
            env['ORCHESTRATOR_RUNNING'] = 'true'
            env['ORCHESTRATOR_AGENT'] = config.name

            process = await asyncio.create_subprocess_exec(
                *cmd,
                cwd=str(self.project_root),
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                env=env,
            )

            progress_task = asyncio.create_task(progress_indicator(config.name, start_time))

            try:
                stdout, stderr = await asyncio.wait_for(
                    process.communicate(), timeout=timeout
                )
            except asyncio.TimeoutError:
                progress_task.cancel()
                print()
                process.kill()
                try:
                    await asyncio.wait_for(process.wait(), timeout=5.0)
                except asyncio.TimeoutError:
                    pass
                return ExecutionResult(
                    agent_name=config.name, status=AgentStatus.FAILED,
                    session_id=session_id, exit_code=-1,
                    duration=time.time() - start_time,
                    cost=0, tokens=0, output_files=[],
                    error_message=f"Timeout ({timeout}s)",
                )
            finally:
                progress_task.cancel()
                try:
                    await progress_task
                except asyncio.CancelledError:
                    pass
                print()

            cost, tokens = self._parse_stream_json(stdout.decode('utf-8', errors='replace'))
            duration = time.time() - start_time
            output_files = self._check_output_files(config.output_files)
            status = AgentStatus.COMPLETED if process.returncode == 0 else AgentStatus.FAILED

            return ExecutionResult(
                agent_name=config.name, status=status,
                session_id=session_id, exit_code=process.returncode,
                duration=duration, cost=cost, tokens=tokens,
                output_files=output_files,
                error_message=stderr.decode('utf-8', errors='replace') if process.returncode != 0 else None,
            )

        except Exception as e:
            return ExecutionResult(
                agent_name=config.name, status=AgentStatus.FAILED,
                session_id=session_id, exit_code=1,
                duration=time.time() - start_time,
                cost=0, tokens=0, output_files=[],
                error_message=str(e),
            )
        finally:
            if prompt_temp_file and prompt_temp_file.exists():
                try:
                    prompt_temp_file.unlink()
                except (OSError, PermissionError):
                    pass


# ============================================================
# StateManager
# ============================================================

class StateManager:
    def __init__(self, project_root: Path):
        self.state_file = project_root / ".claude" / "state.json"
        self.state_file.parent.mkdir(parents=True, exist_ok=True)

    def save_state(self, state: Dict) -> None:
        temp = self.state_file.with_suffix('.tmp')
        with open(temp, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2, ensure_ascii=False)
        temp.replace(self.state_file)

    def load_state(self) -> Optional[Dict]:
        if self.state_file.exists():
            with open(self.state_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None

    def clear_state(self) -> None:
        if self.state_file.exists():
            self.state_file.unlink()


# ============================================================
# ErrorHandler
# ============================================================

class ErrorHandler:
    def __init__(self, project_root: Path, max_retries: int = 3):
        self.max_retries = max_retries
        self.backoff_seconds = [5, 10, 20]
        self.error_log_file = project_root / ".claude" / "error_log.json"
        self.error_log_file.parent.mkdir(parents=True, exist_ok=True)

    async def retry_with_backoff(self, func, *args, **kwargs) -> ExecutionResult:
        for attempt in range(self.max_retries):
            result = await func(*args, **kwargs)
            if result.status == AgentStatus.COMPLETED:
                return result
            if attempt < self.max_retries - 1:
                wait = self.backoff_seconds[attempt]
                print(f"  [重试] {result.agent_name} 失败，{wait}s 后重试 ({attempt + 1}/{self.max_retries})")
                await asyncio.sleep(wait)

        self._log_error(result)
        return result

    def _log_error(self, result: ExecutionResult) -> None:
        entry = {
            "timestamp": datetime.now().isoformat(),
            "agent": result.agent_name,
            "exit_code": result.exit_code,
            "error": result.error_message,
            "session_id": result.session_id,
        }
        errors = []
        if self.error_log_file.exists():
            try:
                errors = json.loads(self.error_log_file.read_text(encoding='utf-8'))
            except (json.JSONDecodeError, IOError):
                errors = []
        errors.append(entry)
        self.error_log_file.write_text(json.dumps(errors, indent=2, ensure_ascii=False), encoding='utf-8')


# ============================================================
# ProgressMonitor
# ============================================================

AGENT_DISPLAY = {
    "analyst":  "PPT模板分析师",
    "builder":  "PPT构建师",
    "reviewer": "PPT验收师",
    "developer": "PPT代码专家",
}


class ProgressMonitor:
    def display_agent_start(self, agent_name: str) -> None:
        print(f"  [启动] {AGENT_DISPLAY.get(agent_name, agent_name)}")

    def display_agent_complete(self, result: ExecutionResult) -> None:
        icon = "✅" if result.status == AgentStatus.COMPLETED else "❌"
        if result.cost > 0:
            cost_info = f"${result.cost:.4f}"
        elif result.tokens > 0:
            cost_info = f"{result.tokens:,} tokens"
        else:
            cost_info = "Pro 订阅"
        name = AGENT_DISPLAY.get(result.agent_name, result.agent_name)
        print(f"  {icon} {name} — {result.status.value} ({result.duration:.0f}s, {cost_info})")
        if result.status == AgentStatus.FAILED and result.error_message:
            print(f"      错误: {result.error_message[:120]}")

    def display_summary(self, results: Dict[str, ExecutionResult], total_duration: float) -> None:
        total_cost = sum(r.cost for r in results.values())
        total_tokens = sum(r.tokens for r in results.values())

        print(f"\n{'=' * 60}")
        print(f"执行完成 — 总耗时 {total_duration:.0f}s")
        if total_cost > 0:
            print(f"总成本: ${total_cost:.4f}")
        if total_tokens > 0:
            print(f"总 tokens: {total_tokens:,}")
        print(f"{'=' * 60}")

        for name, r in results.items():
            icon = "✅" if r.status == AgentStatus.COMPLETED else "❌"
            print(f"  {icon} {name:12s} — {r.status.value}")


# ============================================================
# PPTOrchestrator — core workflow
# ============================================================

class PPTOrchestrator:
    """Fixed 4-agent workflow: Analyst → PAUSE → Builder → Reviewer → loop."""

    def __init__(
        self,
        project_root: Path,
        max_rounds: int = 3,
        max_budget: float = 10.0,
        skip_analyst_first_round: bool = False,
        auto_mode: bool = False,
        init_mode: bool = False,
    ):
        self.project_root = project_root
        self.max_rounds = max_rounds
        self.skip_analyst_first_round = skip_analyst_first_round
        self.auto_mode = auto_mode
        self.init_mode = init_mode
        self.executor = AgentExecutor(project_root, max_budget)
        self.error_handler = ErrorHandler(project_root)
        self.state_manager = StateManager(project_root)
        self.monitor = ProgressMonitor()
        self.results: Dict[str, ExecutionResult] = {}

    # ------------------------------------------------------------------
    # Version auto-detection
    # ------------------------------------------------------------------

    def _detect_next_version_index(self) -> int:
        """Scan pptx files + version tracker, return the next version index.

        Version mapping: 1.0=10, 1.1=11, ..., 1.9=19, 2.0=20, ...
        Returns 10 (=1.0) if no existing versions found.
        """
        max_idx = 9  # so first version = 10 = "1.0"
        ver_pattern = re.compile(r"(\d+)\.(\d+)")

        # Source 1: existing pptx files
        pptx_pattern = re.compile(r"^claude-ppt (\d+)\.(\d+)\.pptx$")
        for f in self.project_root.glob("claude-ppt *.pptx"):
            m = pptx_pattern.match(f.name)
            if m:
                idx = int(m.group(1)) * 10 + int(m.group(2))
                if idx > max_idx:
                    max_idx = idx

        # Source 2: version tracker (records xlsx sheets + past attempts)
        tracker = self.project_root / "pipeline-progress" / ".version_tracker.json"
        if tracker.exists():
            try:
                versions = json.loads(tracker.read_text(encoding="utf-8"))
                for v in versions:
                    m = ver_pattern.match(str(v))
                    if m:
                        idx = int(m.group(1)) * 10 + int(m.group(2))
                        if idx > max_idx:
                            max_idx = idx
            except (json.JSONDecodeError, IOError):
                pass

        return max_idx + 1

    def _record_version(self, version: str) -> None:
        """Append version to tracker so future runs skip it."""
        tracker = self.project_root / "pipeline-progress" / ".version_tracker.json"
        versions = []
        if tracker.exists():
            try:
                versions = json.loads(tracker.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, IOError):
                versions = []
        if version not in versions:
            versions.append(version)
            tracker.write_text(json.dumps(versions, ensure_ascii=False), encoding="utf-8")

    @staticmethod
    def _idx_to_version(idx: int) -> str:
        """Convert version index to string: 10→'1.0', 19→'1.9', 20→'2.0'."""
        return f"{idx // 10}.{idx % 10}"

    # ------------------------------------------------------------------
    # Analyst enhanced marker
    # ------------------------------------------------------------------

    ENHANCED_MARKER = "pipeline-progress/.analyst_enhanced.json"

    def _load_enhanced_list(self) -> list[str]:
        """Load list of LLM-enhanced sheet names."""
        marker = self.project_root / self.ENHANCED_MARKER
        if marker.exists():
            try:
                return json.loads(marker.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, IOError):
                pass
        return []

    def _mark_enhanced(self, sheet_name: str) -> None:
        """Record that a sheet has been LLM-enhanced."""
        marker = self.project_root / self.ENHANCED_MARKER
        enhanced = self._load_enhanced_list()
        if sheet_name not in enhanced:
            enhanced.append(sheet_name)
            marker.write_text(json.dumps(enhanced, ensure_ascii=False), encoding="utf-8")

    # ------------------------------------------------------------------
    # Pipeline direct execution
    # ------------------------------------------------------------------

    def _run_pipeline(self, script: str, args: list[str] | None = None) -> subprocess.CompletedProcess:
        """Run a pipeline script directly via subprocess (no agent overhead)."""
        cmd = [sys.executable, script] + (args or [])
        env = os.environ.copy()
        env["PYTHONUTF8"] = "1"
        return subprocess.run(
            cmd, cwd=str(self.project_root),
            capture_output=True, text=True, encoding="utf-8", env=env,
        )

    def _run_pipeline_step(self, script: str, args: list[str] | None = None) -> bool:
        """Run a pipeline step, print filtered output, return True on success."""
        script_name = Path(script).name
        print(f"  [Pipeline] {script_name} ...")
        t0 = time.time()
        r = self._run_pipeline(script, args)
        dur = time.time() - t0
        if r.returncode != 0:
            print(f"\n❌ {script_name} 失败 ({dur:.0f}s):\n{r.stderr[:500]}")
            return False
        for line in r.stdout.splitlines():
            stripped = line.strip()
            if stripped and ("✅" in stripped or "❌" in stripped or "⚠️" in stripped
                            or "[WARN]" in stripped or "写入" in stripped
                            or "生成" in stripped or "pending" in stripped
                            or "Phase" in stripped or "[OK]" in stripped
                            or "GPT prompt" in stripped):
                print(f"    {stripped}")
        print(f"  ✅ {script_name} ({dur:.0f}s)")
        return True

    def _run_03a_with_prompt_review(self) -> bool:
        """Run 03a in two phases with prompt review pause between them."""
        # Phase 1: assemble prompts (no GPT calls)
        if not self._run_pipeline_step("pipeline/03a_build_shape.py", ["--assemble-only"]):
            return False

        # Check if there are pending GPT prompts
        pending_path = self.project_root / "pipeline-progress" / "03a-pending_prompts.json"
        has_pending = False
        if pending_path.exists():
            try:
                pd = json.loads(pending_path.read_text(encoding="utf-8"))
                has_pending = bool(pd.get("pending"))
            except Exception:
                pass

        if has_pending:
            if self.auto_mode:
                print(f"  [全自动] 跳过 prompt 审核，直接执行 GPT")
            else:
                xlsx_path = str((self.project_root / "pipeline-progress" / "01-shape_detail.xlsx").resolve())
                try:
                    os.startfile(xlsx_path)
                    print(f"  [已打开] Excel — 可在「GPT-prompt Text」单元格中编辑 prompt")
                except Exception:
                    print(f"  [手动打开] {xlsx_path}")
                print(f"\n{'=' * 60}")
                print(f"⏸️  GPT Prompt 审核")
                print(f"   请检查 01-shape_detail.xlsx 中「GPT-prompt Text」列")
                print(f"   编辑完成后保存并关闭 Excel，然后按 Enter 继续...")
                print(f"{'=' * 60}")
                input()
        else:
            print(f"  [INFO] 无 GPT 待审核 prompt，跳过审核暂停")

        # Phase 2: call GPT with (possibly edited) prompts
        if not self._run_pipeline_step("pipeline/03a_build_shape.py", ["--execute-prompts"]):
            return False

        return True

    def _run_builder_pipeline(self, version: str, sheet_arg: list[str] | None = None) -> bool:
        """Run 02 → 03a (two-phase with prompt review) → 03b. Returns True on success."""
        # 02
        args_02 = sheet_arg
        if not self._run_pipeline_step("pipeline/02_shape_analysis.py", args_02):
            return False

        # 03a two-phase
        if not self._run_03a_with_prompt_review():
            return False

        # 03b
        if not self._run_pipeline_step("pipeline/03b_build_ppt_com.py", ["--version", version]):
            return False

        return True

    # ------------------------------------------------------------------
    # Hot iteration helpers
    # ------------------------------------------------------------------

    def _prompts_exist(self) -> bool:
        """Check if GPT-prompt Text cells are already populated in Excel."""
        try:
            from pipeline.ppt_pipeline_common import read_gpt_prompts_from_xlsx
            prompts = read_gpt_prompts_from_xlsx()
            return len(prompts) > 0
        except Exception:
            return False

    def _run_prompt_only_pipeline(self, version: str) -> bool:
        """Hot iteration: prompts already in Excel, skip 02 + 03a Phase 1."""
        if not self._run_pipeline_step("pipeline/03a_build_shape.py", ["--execute-prompts"]):
            return False
        if not self._run_pipeline_step("pipeline/03b_build_ppt_com.py", ["--version", version]):
            return False
        return True

    def _builder_prompt_optimizer_prompt(self, sheet_name: str, fix_data: list[dict]) -> str:
        """Builder LLM: 根据 fix 报告全面重写 GPT-prompt Text。"""

        # 1. 读取原始 prompt 模板（clean baseline）
        tpl_path = self.project_root / "pipeline" / "prompt_templates" / "gpt_summary.md"
        original_template = ""
        if tpl_path.exists():
            try:
                original_template = tpl_path.read_text(encoding="utf-8")[:800]
            except Exception:
                pass

        # 2. 读取当前生成内容（供 Builder 了解实际产出）
        content_path = self.project_root / "pipeline-progress" / "03a-build_shape_content.json"
        content_snippets = {}
        if content_path.exists():
            try:
                data = json.loads(content_path.read_text(encoding="utf-8"))
                for item in data.get("items", []):
                    sn = item.get("shape_name", "")
                    content_snippets[sn] = item.get("content", "")[:200]
            except Exception:
                pass

        # 3. 构建 fix 清单（附带当前内容片段）
        fix_lines = []
        for fx in fix_data:
            if fx.get("fix_type") == "code":
                continue
            shape = fx["shape"]
            line = f"  - {shape}: [{fx.get('fix_type','')}] {fx.get('issue','')}"
            line += f"\n    建议: {fx.get('suggestion','')}"
            if shape in content_snippets:
                line += f"\n    当前生成内容(前200字): {content_snippets[shape]}"
            fix_lines.append(line)
        fix_items = "\n".join(fix_lines)

        tpl_section = ""
        if original_template:
            tpl_section = (
                f"## 原始 prompt 模板（clean baseline）\n"
                f"```\n{original_template}\n```\n\n"
            )

        return (
            f"你是 PPT 内容优化师。根据验收报告，**全面重写** Excel 中的 GPT prompt。\n\n"
            f"## 核心原则\n"
            f"**不要在原 prompt 上打补丁！** 每次都基于原始模板 + fix 建议，重新编写一个干净、完整的 prompt。\n"
            f"原因：反复追加约束会导致 prompt 冗余膨胀、指令矛盾，GPT 产出质量反而下降。\n\n"
            f"{tpl_section}"
            f"## 工作 sheet: 「{sheet_name}」\n\n"
            f"## 验收失败清单\n{fix_items}\n\n"
            f"## 重写策略\n"
            f"1. 先读取当前 prompt 全文\n"
            f"2. 对照原始模板，识别哪些是有效约束、哪些是冗余补丁\n"
            f"3. 以原始模板为骨架，融入 fix 建议中的有效约束，生成一个干净的新 prompt\n"
            f"4. 具体 fix_type 处理：\n"
            f"   - keyword_missing: 在 prompt **前半段**明确要求包含缺失关键词（不要放在末尾）\n"
            f"   - budget_overflow: 降低目标字数（当前 budget 的 85%），让 GPT 输出更紧凑\n"
            f"   - budget_underflow: 提高字数下限，要求充实内容\n"
            f"   - style_mismatch: 加入格式/语调约束\n\n"
            f"## 任务\n"
            f"1. 通过 Python COM 打开 `pipeline-progress/01-shape_detail.xlsx`「{sheet_name}」sheet\n"
            f"2. 找到上述 shape 的「GPT-prompt Text」单元格\n"
            f"3. **全面重写** prompt（不是追加！），保持干净、无冗余\n"
            f"4. 保存并关闭 xlsx\n"
            f"5. 打印修改摘要（列出每个 shape 的改动要点）\n\n"
            f"## 规则\n"
            f"- 只改有问题的 shape 的 prompt，其余不动\n"
            f"- **不要修改「内容描述」「strategy」「params」等注释字段**\n"
            f"- ⚠️ 不要运行任何 pipeline 脚本\n"
        )

    # ------------------------------------------------------------------
    # Prompt builders
    # ------------------------------------------------------------------

    def _get_latest_version(self) -> str | None:
        """Return the latest existing version string, or None if fresh."""
        idx = self._detect_next_version_index() - 1
        if idx < 10:
            return None
        return self._idx_to_version(idx)

    def _analyst_phase2_prompt(self, target_sheet: str, fix_shapes: list[dict] | None = None) -> str:
        """Analyst LLM: 增强批注。

        fix_shapes: 如果是续跑，传入失败 shape 列表 [{shape, issue, fix_type, suggestion}]。
                    Analyst 仅聚焦这些 shape，其余跳过。
                    如果为 None，则全量分析所有 shape。
        """
        golden_ref = (
            "【golden reference 示例】\n"
            "好的「内容描述」格式：\n"
            "  缺点总结: 从补充说明总结缺点。必须包含'建议'、'反馈'、'样本'关键词，用【】括起关键性能词，每段结论后注明(X/N)比例\n"
            "  优点总结: 从补充说明总结优点。必须包含'建议'、'反馈'、'样本'关键词，用【】括起关键性能词，每段结论后注明(X/N)比例\n"
            "差的「内容描述」：\n"
            "  ❌ 从补充说明总结缺点 (缺少关键词和格式约束)\n"
            "  ❌ 总结一下缺点 (太模糊)"
        )

        # Targeted mode: only fix specific shapes
        if fix_shapes:
            fix_detail = "\n".join(
                f"  - {fx['shape']}: [{fx.get('fix_type','')}] {fx.get('issue','')} → {fx.get('suggestion','')}"
                for fx in fix_shapes
            )
            shape_names = [fx['shape'] for fx in fix_shapes if fx['shape'] != '(全局)']
            global_fixes = [fx for fx in fix_shapes if fx['shape'] == '(全局)']

            scope_note = ""
            if shape_names:
                scope_note += f"需要修改的 shape: {', '.join(shape_names)}\n"
            if global_fixes:
                scope_note += "另有全局修正（影响所有 gpt_prompted shape）\n"

            return (
                f"你是 PPT 模板分析师。上一轮验收发现以下问题，需要你精准修正批注。\n\n"
                f"⚠️ 本次是【定向修正】，只改有问题的 shape，其余 shape 不要动！\n\n"
                f"【验收反馈】\n{fix_detail}\n\n"
                f"{scope_note}\n"
                f"你的任务：\n"
                f"1. 读取 `pipeline-progress/01-shape_detail_com.json` 了解 shape 属性\n"
                f"2. 通过 COM 打开 `pipeline-progress/01-shape_detail.xlsx`「{target_sheet}」sheet\n"
                f"   ⚠️ 必须操作「{target_sheet}」sheet\n"
                f"3. 只修改上述 shape 的「内容描述」，根据验收反馈精准调整：\n"
                f"   - keyword_missing → 在「内容描述」中追加缺失关键词要求\n"
                f"   - budget_overflow → 追加具体字数上限约束\n"
                f"   - budget_underflow → 追加内容充实要求 + 字数下限\n"
                f"   - style_mismatch → 追加格式/语调约束\n"
                f"   - 全局修正 → 所有 gpt_prompted shape 的「内容描述」都追加对应要求\n"
                f"4. 保存并关闭 xlsx\n"
                f"5. 打印修改摘要\n\n"
                f"{golden_ref}"
            )

        # Full mode: enhance all shapes
        return (
            "你是 PPT 模板分析师。Pipeline 已完成自动推断。\n\n"
            "你的任务：增强 xlsx 中所有 shape 的批注质量。\n\n"
            f"1. 读取 `pipeline-progress/01-shape_detail_com.json` 了解每个 shape 的属性\n"
            f"2. 通过 COM 读取 `pipeline-progress/01-shape_detail.xlsx`「{target_sheet}」sheet 的所有 shape 批注\n"
            f"   ⚠️ 必须操作「{target_sheet}」sheet，不是其他 sheet\n"
            "3. 对每个 shape 评估并改进：\n"
            "   - 「内容描述」: 更具体化，将所有 GPT 约束（关键词、字数、格式）直接写入此字段\n"
            "     例如：「从补充说明总结缺点。必须包含'建议'、'反馈'、'样本'关键词，每段不超过3行，总字数200字以内」\n"
            "   - 「strategy」: 验证自动推断是否匹配 shape 的文本特征\n"
            "   - 「params」: 补充缺失参数\n"
            "   ⚠️ 不再使用「备注」字段，所有指令统一写入「内容描述」\n"
            "4. 特别关注 gpt_prompted 类 shape：\n"
            "   - 读取 shape 的原始 text 确认 filter 方向（缺点/优点）正确\n"
            "   - 在「内容描述」中追加 GPT 输出必须包含的关键词（如「建议」「反馈」）\n"
            "5. 对 strategy 为空或 description 为「（必填）」的 shape，根据原始 text 推断正确的 strategy\n"
            f"6. 通过 Python COM 写入所有改进到 xlsx 的「{target_sheet}」sheet\n"
            "7. 打印修改摘要（列出每个 shape 的改动内容）\n\n"
            f"{golden_ref}"
        )

    def _builder_llm_only_prompt(self, version: str) -> str:
        """Builder LLM: 仅修改 xlsx 批注，不运行 pipeline 脚本。"""
        return (
            f"你是 PPT 构建师。Orchestrator 已运行 02b_iteration_setup.py，"
            f"在 xlsx 中创建了新 sheet「claude-ppt {version}」并应用了基础修正。\n\n"
            f"你的唯一任务：通过 Python COM 精调该 sheet 的批注。\n\n"
            f"1. 读取 `pipeline-progress/04-fix_ppt.md` 获取修正建议\n"
            f"2. 对每个非 code 类 fix 条目（如 keyword_missing / budget_overflow / budget_underflow / style_mismatch），打开 xlsx 修改「claude-ppt {version}」sheet：\n"
            f"   - 精化「内容描述」使 prompt 更精确（将关键词要求、字数约束等直接追加到内容描述中）\n"
            f"   - 调整 strategy/params（如切换 filter=缺点 → filter=优点）\n"
            f"   ⚠️ 不再使用「备注」字段，所有修正统一追加到「内容描述」\n"
            f"3. 保存并关闭 xlsx\n"
            f"4. 打印修改摘要（列出每个 shape 的改动）\n\n"
            f"⚠️ 不要运行任何 pipeline 脚本，orchestrator 会处理。\n\n"
            f"【golden reference 示例】\n"
            f"好的「内容描述」格式：\n"
            f"  缺点总结: 从补充说明总结缺点。必须包含'建议'、'反馈'、'样本'关键词，用【】括起关键性能词，每段结论后注明(X/N)比例\n"
            f"  优点总结: 从补充说明总结优点。必须包含'建议'、'反馈'、'样本'关键词，用【】括起关键性能词，每段结论后注明(X/N)比例\n"
            f"差的「内容描述」：\n"
            f"  ❌ 从补充说明总结缺点 (缺少关键词和格式约束)\n"
            f"  ❌ 总结一下缺点 (太模糊)"
        )

    def _reviewer_llm_only_prompt(self, version: str) -> str:
        """Reviewer LLM: 语义审核 + prompt 级别修复建议。"""
        return (
            f"你是 PPT 验收师。Pipeline 已运行 04_shape_diff_test.py，测试结果为 FAIL。\n\n"
            f"你的任务：分析失败原因，补充精准修复建议（面向 GPT-prompt Text）。\n\n"
            f"1. 读取 `pipeline-progress/04-fix_ppt.md` 和 `pipeline-progress/04-diff_result.json`\n"
            f"2. 对每个失败项做深入分析：\n"
            f"   - 语义覆盖不达标：找出哪些 shape 缺少关键词（样本/建议/反馈），"
            f"建议如何修改对应 shape 的「GPT-prompt Text」\n"
            f"   - Readability 不达标：判断是文本过长/过短/偏题，"
            f"给出具体的 prompt 修改建议（如「在 prompt 中追加：控制在 180-220 字」）\n"
            f"   - Visual 不达标：判断是 COM 写入格式错误还是 shape 匹配错误\n"
            f"3. 将补充诊断追加到 `pipeline-progress/04-fix_ppt.md`\n"
            f"4. 每个 fix 建议必须说明如何修改 GPT-prompt Text\n"
            f"5. 打印验收结论（FAIL + 三层分数 + 具体 prompt 修复建议摘要）\n\n"
            f"⚠️ 修复建议面向 GPT-prompt Text（终端变量），不要建议修改「内容描述」等注释字段。\n"
            f"⚠️ 不要运行任何 pipeline 脚本，orchestrator 会处理。"
        )

    def _developer_prompt(self) -> str:
        fix_content = ""
        fix_path = self.project_root / "pipeline-progress" / "04-fix_ppt.md"
        if fix_path.exists():
            try:
                fix_content = fix_path.read_text(encoding='utf-8')[:3000]
            except (IOError, OSError):
                pass

        return (
            f"你是 PPT 代码专家。Reviewer 诊断出 pipeline 代码缺陷需要修复。\n\n"
            f"修正报告：\n\n---\n{fix_content}\n---\n\n"
            f"请：\n"
            f"1. 找到报告中 `fix_type: code` 的条目\n"
            f"2. 定位并修复对应的 pipeline 脚本\n"
            f"3. 运行语法检查确认无误\n"
            f"4. 简述修改内容"
        )

    # ------------------------------------------------------------------
    # Result checking
    # ------------------------------------------------------------------

    def _check_review_passed(self) -> Tuple[bool, str]:
        """Read diff_result.json, return (passed, fix_type).

        fix_type: 'code' | 'keyword_missing' | 'budget_overflow' | 'budget_underflow' | 'style_mismatch' | 'annotation' | '' (empty if passed).
        """
        diff_path = self.project_root / "pipeline-progress" / "04-diff_result.json"
        if not diff_path.exists():
            return False, "annotation"

        try:
            data = json.loads(diff_path.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, IOError):
            return False, "annotation"

        # Check thresholds
        status = data.get("status", "fail")
        if status == "ok":
            return True, ""

        visual = data.get("visual_score_avg", data.get("visual_score", 0))
        readability = data.get("readability_score_avg", data.get("readability_score", 0))
        semantic = data.get("semantic_coverage", 0)

        if visual >= 98 and readability >= 95 and semantic >= 100:
            return True, ""

        # Determine fix_type from diff_result.json (more reliable than MD text search)
        overall_ft = data.get("fix_type", "")
        if overall_ft == "code":
            return False, "code"
        if overall_ft:
            return False, overall_ft

        # Fallback: check individual fixes for any code type
        for fx in data.get("fixes", []):
            if fx.get("fix_type") == "code":
                return False, "code"

        return False, "annotation"

    def _verify_pptx_exists(self, version: str) -> bool:
        """Check that the expected pptx file was actually created."""
        pptx_path = self.project_root / f"claude-ppt {version}.pptx"
        return pptx_path.exists()

    def _archive_round(self, round_num: int) -> None:
        """Archive fix report and diff result for this round."""
        progress = self.project_root / "pipeline-progress"
        for fname in ["04-fix_ppt.md", "04-diff_result.json"]:
            src = progress / fname
            if src.exists():
                stem, ext = fname.rsplit('.', 1)
                dst = progress / f"{stem}-round{round_num}.{ext}"
                try:
                    dst.write_bytes(src.read_bytes())
                except (IOError, OSError):
                    pass

    # ------------------------------------------------------------------
    # Main workflow
    # ------------------------------------------------------------------

    async def run(self) -> bool:
        """Execute the full PPT production workflow."""
        start_time = time.time()
        state = {
            "task_id": str(uuid.uuid4()),
            "started_at": datetime.now().isoformat(),
            "max_rounds": self.max_rounds,
            "current_round": 0,
            "agents_status": {},
            "results": {},
        }

        print(f"\n{'=' * 60}")
        print(f"PPT 专业制作团队 — 启动")
        print(f"最大迭代轮次: {self.max_rounds}")
        print(f"{'=' * 60}")

        # ---- Phase 1: Analyst (Pipeline direct + Agent if needed) ----
        print(f"\n--- Phase 1: 模板分析 ---")
        phase1_start = time.time()

        # Step 1: Extract template shapes
        json_path = self.project_root / "pipeline-progress" / "01-shape_detail_com.json"
        xlsx_path = self.project_root / "pipeline-progress" / "01-shape_detail.xlsx"
        template_path = self.project_root / "pipeline" / "standard and empty template.pptx"

        # Hot iteration: NEVER overwrite existing xlsx (protects annotations + prompts)
        _skip_pipeline_01 = (
            not self.init_mode and xlsx_path.exists() and json_path.exists()
        )

        if _skip_pipeline_01:
            print("  [Hot] xlsx + JSON 已存在，跳过 01+01b（保护已有 prompt）")
        else:
            # 01: Extract template shapes (skip if JSON is fresh)
            skip_extract = (
                json_path.exists() and xlsx_path.exists() and template_path.exists()
                and json_path.stat().st_mtime > template_path.stat().st_mtime
            )

            if skip_extract:
                print("  [Cache] JSON/xlsx 已是最新，跳过 PPT 提取")
            else:
                print("  [Pipeline] 01_shape_detail.py — 提取模板 shape ...")
                r1 = self._run_pipeline("pipeline/01_shape_detail.py")
                if r1.returncode != 0:
                    print(f"\n❌ 01_shape_detail.py 失败:\n{r1.stderr[:500]}")
                    return False
                print(f"  ✅ 01_shape_detail.py 完成 ({time.time() - phase1_start:.0f}s)")

            # 01b: Auto-annotate
            step1b_start = time.time()
            print("  [Pipeline] 01b_auto_annotate.py — 规则推断批注 ...")
            r2 = self._run_pipeline("pipeline/01b_auto_annotate.py")
            if r2.returncode != 0:
                print(f"\n❌ 01b_auto_annotate.py 失败:\n{r2.stderr[:500]}")
                return False
            for line in r2.stdout.splitlines():
                if line.strip():
                    print(f"    {line}")
            print(f"  ✅ 01b_auto_annotate.py 完成 ({time.time() - step1b_start:.0f}s)")

        # Step 2: Detect continuation state + prepare target sheet
        ambiguous_count = 0
        if not _skip_pipeline_01:
            m = re.search(r"(\d+) 项待LLM审核", r2.stdout)
            if m:
                ambiguous_count = int(m.group(1))
            if ambiguous_count > 0:
                print(f"  [INFO] {ambiguous_count} 个模糊项需要 LLM 重点审核")

        base_idx = self._detect_next_version_index()

        # Init mode: jump to next major version X.0
        if self.init_mode:
            base_idx = (base_idx + 9) // 10 * 10
            # 10→10(1.0), 16→20(2.0), 20→20(2.0), 21→30(3.0)

        fix_report = self.project_root / "pipeline-progress" / "04-fix_ppt.md"
        is_continuation = (base_idx > 10) and fix_report.exists() and not self.init_mode

        # Capture fix.md freshness BEFORE 02b modifies xlsx
        fix_is_fresh = (
            is_continuation
            and fix_report.stat().st_mtime > xlsx_path.stat().st_mtime
        )
        if fix_is_fresh:
            print("  [INFO] fix.md 比 xlsx 更新，Round 1 将自动优化 prompt")

        if is_continuation:
            # Continuation: create new sheet via 02b BEFORE Analyst
            # Retry with incrementing version if sheet already exists
            version_idx = base_idx
            while True:
                next_version = self._idx_to_version(version_idx)
                target_sheet = f"claude-ppt {next_version}"

                print(f"\n  [续跑] 创建新 sheet「{target_sheet}」...")
                t0 = time.time()
                args_02b = ["--version", next_version]
                if not self.init_mode:
                    args_02b.append("--sheet-only")
                r = self._run_pipeline("pipeline/02b_iteration_setup.py", args_02b)
                dur = time.time() - t0
                if r.returncode != 0:
                    print(f"\n❌ 02b_iteration_setup.py 失败 ({dur:.0f}s):\n{r.stderr[:500]}")
                    return False

                # Check if sheet already existed (02b prints "already exists")
                if "already exists" in r.stdout:
                    print(f"    sheet「{target_sheet}」已存在，尝试下一个版本...")
                    self._record_version(next_version)  # record so future runs skip
                    version_idx += 1
                    continue

                # New sheet created successfully
                self._record_version(next_version)
                base_idx = version_idx  # update for builder loop
                for line in r.stdout.splitlines():
                    stripped = line.strip()
                    if stripped:
                        print(f"    {stripped}")
                print(f"  ✅ 02b_iteration_setup.py ({dur:.0f}s) — sheet「{target_sheet}」已创建")
                break
        else:
            target_sheet = "Shape Detail"

        # Step 3: LLM enhances annotations
        # Priority: hot iteration > user override > auto-detection > default (run)
        skip_analyst_llm = False

        # Highest priority: hot iteration always skips Analyst LLM
        if not self.init_mode:
            skip_analyst_llm = True
            print(f"\n  [Hot] 跳过 Analyst LLM 增强（prompt 已存在）")
            self.results["analyst"] = ExecutionResult(
                agent_name="analyst", status=AgentStatus.COMPLETED,
                session_id="skipped-hot-iteration", exit_code=0,
                duration=0.0, cost=0.0, tokens=0, output_files=[],
            )

        # User chose to skip at startup
        elif self.skip_analyst_first_round:
            skip_analyst_llm = True
            print(f"\n  [跳过 Agent] 用户选择跳过 Analyst LLM 增强")
            self.results["analyst"] = ExecutionResult(
                agent_name="analyst", status=AgentStatus.COMPLETED,
                session_id="skipped-user", exit_code=0,
                duration=0.0, cost=0.0, tokens=0, output_files=[],
            )

        # Auto-detection: skip if source sheet already enhanced and no new fix report
        if not skip_analyst_llm and is_continuation:
            prev_version = self._idx_to_version(version_idx - 1)
            prev_sheet = f"claude-ppt {prev_version}"
            marker_path = self.project_root / self.ENHANCED_MARKER
            if prev_sheet in self._load_enhanced_list():
                # Check if fix report is newer than marker → new reviewer feedback
                fix_newer = (
                    fix_report.exists() and marker_path.exists()
                    and fix_report.stat().st_mtime > marker_path.stat().st_mtime
                )
                if fix_newer:
                    print(f"\n  [INFO] 04-fix_ppt.md 有新反馈，Analyst LLM 需根据反馈重新增强")
                else:
                    skip_analyst_llm = True
                    print(f"\n  [跳过 Agent] 批注已从增强的「{prev_sheet}」继承，无需重跑 LLM")
                    self._mark_enhanced(target_sheet)
                    self.results["analyst"] = ExecutionResult(
                        agent_name="analyst", status=AgentStatus.COMPLETED,
                        session_id="skipped-inherited", exit_code=0,
                        duration=0.0, cost=0.0, tokens=0, output_files=[],
                    )

        if not skip_analyst_llm:
            # Load fix data for targeted mode (continuation with fix report)
            fix_shapes = None
            diff_path = self.project_root / "pipeline-progress" / "04-diff_result.json"
            if is_continuation and diff_path.exists():
                try:
                    diff_data = json.loads(diff_path.read_text(encoding="utf-8"))
                    fixes = [f for f in diff_data.get("fixes", []) if f.get("fix_type") != "code"]
                    if fixes:
                        fix_shapes = fixes
                        shape_names = [f["shape"] for f in fixes if f["shape"] != "(全局)"]
                        print(f"\n  [Agent] Analyst 定向修正「{target_sheet}」中 {len(shape_names)} 个问题 shape ...")
                    else:
                        print(f"\n  [Agent] Analyst 增强「{target_sheet}」中所有 shape 批注 ...")
                except (json.JSONDecodeError, IOError):
                    print(f"\n  [Agent] Analyst 增强「{target_sheet}」中所有 shape 批注 ...")
            else:
                print(f"\n  [Agent] Analyst 增强「{target_sheet}」中所有 shape 批注 ...")

            self.monitor.display_agent_start("analyst")
            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent, AGENT_CONFIGS["analyst"],
                self._analyst_phase2_prompt(target_sheet, fix_shapes=fix_shapes)
            )
            self.results["analyst"] = result
            self.monitor.display_agent_complete(result)
            if result.status == AgentStatus.FAILED:
                print("\n❌ Analyst LLM 增强失败，工作流终止。")
                return False
            self._mark_enhanced(target_sheet)

        state["agents_status"]["analyst"] = self.results["analyst"].status.value

        self.state_manager.save_state(state)

        # ---- PAUSE for human review ----
        if not self.init_mode:
            pass  # Hot iteration: no annotation PAUSE (user reviews prompts later)
        elif self.skip_analyst_first_round:
            print(f"\n  [跳过 PAUSE] 用户选择跳过，直接进入构建阶段")
        elif self.auto_mode:
            print(f"\n  [全自动] 跳过批注校准暂停")
        else:
            xlsx_path = str((self.project_root / "pipeline-progress" / "01-shape_detail.xlsx").resolve())
            try:
                os.startfile(xlsx_path)
                print(f"  [已打开] Excel: 01-shape_detail.xlsx")
            except Exception:
                print(f"  [手动打开] {xlsx_path}")
            print(f"\n{'=' * 60}")
            print(f"⏸️  PAUSE — 请校准批注  <sheet: {target_sheet}>")
            print(f"   检查「{target_sheet}」中黄色「内容描述」单元格，确认批注是否正确。")
            print("   修改完成后保存并关闭 Excel，然后按 Enter 继续...")
            print(f"{'=' * 60}")
            input()

        for round_num in range(1, self.max_rounds + 1):
            version = self._idx_to_version(base_idx + round_num - 1)
            state["current_round"] = round_num
            round_start = time.time()
            builder_failed = False

            # Three builder modes:
            #   fresh_build:       first ever build (02→03a→03b, no --sheet)
            #   continuation_r1:   continuation round 1 (02b+Analyst already done, pipeline with --sheet)
            #   correction:        round 2+, or continuation round 2+ (02b→LLM→pipeline)
            is_fresh_build = (round_num == 1 and not is_continuation)
            is_continuation_r1 = (round_num == 1 and is_continuation)

            print(f"\n--- Round {round_num}: 构建 claude-ppt {version}.pptx ---")
            if not is_continuation_r1:
                # continuation_r1 already recorded version before Analyst
                self._record_version(version)

            # ---- Builder ----
            if not self.init_mode and round_num == 1:
                # Hot iteration first round: prompt review + prompt-only pipeline
                # Check prerequisites
                mapping_json = self.project_root / "pipeline-progress" / "02-shape_analysis_map.json"
                if not mapping_json.exists() or not self._prompts_exist():
                    print(f"\n  ⚠️  前置产物不完整（02 artifacts 或 prompt 缺失）")
                    print(f"      请先运行选项 [0-初始化] 构建完整结构")
                    builder_failed = True
                else:
                    # Auto-optimize prompts if fix.md is fresh
                    if fix_is_fresh:
                        fix_data = []
                        diff_path = self.project_root / "pipeline-progress" / "04-diff_result.json"
                        if diff_path.exists():
                            try:
                                fix_data = json.loads(diff_path.read_text(encoding="utf-8")).get("fixes", [])
                            except (json.JSONDecodeError, IOError):
                                pass
                        if fix_data:
                            sheet_name = f"claude-ppt {version}"
                            print(f"  [Agent] Builder LLM 根据 fix.md 优化 prompt ...")
                            self.monitor.display_agent_start("builder")
                            result = await self.error_handler.retry_with_backoff(
                                self.executor.run_agent, AGENT_CONFIGS["builder"],
                                self._builder_prompt_optimizer_prompt(sheet_name, fix_data)
                            )
                            self.monitor.display_agent_complete(result)
                            if result.status == AgentStatus.FAILED:
                                print(f"\n⚠️  Builder LLM prompt 优化失败，继续手动审核")

                    if not self.auto_mode:
                        xlsx_path = str((self.project_root / "pipeline-progress" / "01-shape_detail.xlsx").resolve())
                        try:
                            os.startfile(xlsx_path)
                            print(f"  [已打开] Excel — 可在「GPT-prompt Text」中编辑 prompt")
                        except Exception:
                            print(f"  [手动打开] {xlsx_path}")
                        print(f"\n⏸️  PROMPT REVIEW — 审核/编辑 prompt 后按 Enter 继续...")
                        input()
                    else:
                        print(f"  [全自动] 跳过 prompt 审核")
                    builder_failed = not self._run_prompt_only_pipeline(version)
            elif is_fresh_build:
                # Cold start first build: pure pipeline, no --sheet
                builder_failed = not self._run_builder_pipeline(version)
            elif is_continuation_r1:
                # Continuation round 1: 02b + Analyst already done, just run pipeline
                print(f"  [续跑] sheet「claude-ppt {version}」已由 Analyst 增强，直接生成 PPT")
                builder_failed = not self._run_builder_pipeline(
                    version, sheet_arg=["--sheet", f"claude-ppt {version}"])
            else:
                # Correction round: prompt-centric
                sheet_name = f"claude-ppt {version}"

                # Phase 1: create new sheet (inherit prompts, skip annotation fixes)
                if not self._run_pipeline_step("pipeline/02b_iteration_setup.py",
                                               ["--version", version, "--sheet-only"]):
                    builder_failed = True

                # Phase 2: Builder LLM edits prompts directly
                if not builder_failed:
                    fix_data = []
                    diff_path = self.project_root / "pipeline-progress" / "04-diff_result.json"
                    if diff_path.exists():
                        try:
                            fix_data = json.loads(diff_path.read_text(encoding="utf-8")).get("fixes", [])
                        except (json.JSONDecodeError, IOError):
                            pass
                    print(f"  [Agent] Builder LLM 优化 prompt ...")
                    self.monitor.display_agent_start("builder")
                    result = await self.error_handler.retry_with_backoff(
                        self.executor.run_agent, AGENT_CONFIGS["builder"],
                        self._builder_prompt_optimizer_prompt(sheet_name, fix_data)
                    )
                    self.monitor.display_agent_complete(result)
                    if result.status == AgentStatus.FAILED:
                        print(f"\n❌ Builder LLM prompt 优化失败 (round {round_num})。")
                        builder_failed = True

                # Phase 3: prompt review pause
                if not builder_failed and not self.auto_mode:
                    xlsx_path = str((self.project_root / "pipeline-progress" / "01-shape_detail.xlsx").resolve())
                    try:
                        os.startfile(xlsx_path)
                        print(f"  [已打开] Excel — 请审核修改后的 GPT-prompt Text")
                    except Exception:
                        print(f"  [手动打开] {xlsx_path}")
                    print(f"\n⏸️  PROMPT REVIEW — 审核 prompt 后按 Enter 继续...")
                    input()

                # Phase 4: prompt-only pipeline (03a Phase 2 + 03b)
                if not builder_failed:
                    builder_failed = not self._run_prompt_only_pipeline(version)

            # Record builder result
            key = f"builder_round{round_num}"
            builder_dur = time.time() - round_start
            if builder_failed:
                self.results[key] = ExecutionResult(
                    agent_name="builder", status=AgentStatus.FAILED,
                    session_id="pipeline-direct", exit_code=1,
                    duration=builder_dur, cost=0.0, tokens=0, output_files=[],
                    error_message=f"Pipeline script failed in round {round_num}",
                )
                state["agents_status"][key] = "failed"
                self.state_manager.save_state(state)
                print(f"\n❌ Builder 失败 (round {round_num})，工作流终止。")
                break

            self.results[key] = ExecutionResult(
                agent_name="builder", status=AgentStatus.COMPLETED,
                session_id="pipeline-direct", exit_code=0,
                duration=builder_dur, cost=0.0, tokens=0,
                output_files=[f"claude-ppt {version}.pptx"],
            )
            state["agents_status"][key] = "completed"
            self.state_manager.save_state(state)

            # Verify pptx actually exists
            if not self._verify_pptx_exists(version):
                print(f"\n❌ Builder 完成但 claude-ppt {version}.pptx 未生成。")
                print(f"   请检查 pipeline 脚本输出。工作流终止。")
                break

            print(f"  ✅ Builder 完成 ({builder_dur:.0f}s)")

            # ---- max_rounds=1: skip reviewer, end directly ----
            if self.max_rounds == 1:
                if self.init_mode:
                    # Cold start: auto-verify (report only, no correction round)
                    print(f"\n  [验收] 自动检查 claude-ppt {version}.pptx 质量...")
                    r = self._run_pipeline(
                        "pipeline/04_shape_diff_test.py",
                        ["--target", f"claude-ppt {version}.pptx"]
                    )
                    for line in r.stdout.splitlines():
                        stripped = line.strip()
                        if stripped:
                            print(f"    {stripped}")
                    passed, fix_type = self._check_review_passed()
                    if passed:
                        print(f"\n✅ claude-ppt {version}.pptx 初始化完成，验收通过！")
                    else:
                        print(f"\n⚠️  claude-ppt {version}.pptx 初始化完成，验收未通过 (fix_type={fix_type})")
                        print(f"   已生成 fix.md — 下次选1或选2时 Builder LLM 会自动优化 prompt")
                else:
                    print(f"\n✅ claude-ppt {version}.pptx 已生成，请人工审核。")
                self.monitor.display_summary(self.results, time.time() - start_time)
                self.state_manager.clear_state()
                return True

            # ---- max_rounds>=2: pause for human review + confirm continue ----
            if self.auto_mode:
                print(f"\n  [全自动] 跳过 PPT 审核，直接进入验收")
            else:
                print(f"\n{'=' * 60}")
                print(f"⏸️  PPT 已生成: claude-ppt {version}.pptx")
                print(f"   请先人工审核 PPT。")
                cont = input("   审核完成后，是否进入验收+修正流程？[Y/n]（直接回车=是）: ").strip().lower()
                print(f"{'=' * 60}")
                if cont in ('n', 'no'):
                    print("用户选择终止，工作流结束。")
                    self.monitor.display_summary(self.results, time.time() - start_time)
                    self.state_manager.clear_state()
                    return True

            # ---- Reviewer: direct pipeline + conditional LLM ----
            print(f"\n  [Pipeline] 04_shape_diff_test.py --target \"claude-ppt {version}.pptx\" ...")
            t0 = time.time()
            r = self._run_pipeline(
                "pipeline/04_shape_diff_test.py",
                ["--target", f"claude-ppt {version}.pptx"]
            )
            review_dur = time.time() - t0
            # Show test output (04 returns 1 for FAIL, which is expected)
            for line in r.stdout.splitlines():
                stripped = line.strip()
                if stripped:
                    print(f"    {stripped}")
            print(f"  ✅ 04_shape_diff_test.py ({review_dur:.0f}s)")

            # Check pass/fail
            passed, fix_type = self._check_review_passed()
            if passed:
                key = f"reviewer_round{round_num}"
                self.results[key] = ExecutionResult(
                    agent_name="reviewer", status=AgentStatus.COMPLETED,
                    session_id="pipeline-direct", exit_code=0,
                    duration=review_dur, cost=0.0, tokens=0,
                    output_files=["pipeline-progress/04-diff_result.json"],
                )
                state["agents_status"][key] = "completed"
                print(f"\n✅ claude-ppt {version}.pptx 验收通过！")
                self.monitor.display_summary(self.results, time.time() - start_time)
                self.state_manager.clear_state()
                return True

            # FAIL: LLM semantic review for enhanced fix suggestions
            print(f"\n  [Agent] Reviewer LLM 语义审核 ...")
            self.monitor.display_agent_start("reviewer")
            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent, AGENT_CONFIGS["reviewer"],
                self._reviewer_llm_only_prompt(version)
            )
            key = f"reviewer_round{round_num}"
            self.results[key] = result
            self.monitor.display_agent_complete(result)
            state["agents_status"][key] = result.status.value
            self.state_manager.save_state(state)

            print(f"\n⚠️  Round {round_num} 未通过 (fix_type={fix_type})")

            if round_num >= self.max_rounds:
                print(f"已达最大轮次 ({self.max_rounds})，工作流结束。请检查 04-fix_ppt.md。")
                break

            # Archive this round's reports
            self._archive_round(round_num)

            # Conditional: Developer for code fixes
            if fix_type == "code":
                print(f"\n--- Developer 介入修复代码 ---")
                self.monitor.display_agent_start("developer")
                result = await self.error_handler.retry_with_backoff(
                    self.executor.run_agent, AGENT_CONFIGS["developer"], self._developer_prompt()
                )
                key = f"developer_round{round_num}"
                self.results[key] = result
                self.monitor.display_agent_complete(result)
                state["agents_status"][key] = result.status.value
                self.state_manager.save_state(state)

                if result.status == AgentStatus.FAILED:
                    print(f"\n❌ Developer 失败，工作流终止。")
                    break

        if self.auto_mode:
            latest_idx = base_idx + min(round_num, self.max_rounds) - 1
            final_ver = self._idx_to_version(latest_idx)
            print(f"\n{'=' * 60}")
            print(f"🤖 全自动模式完成 — {self.max_rounds} 轮迭代")
            print(f"   最终产物: claude-ppt {final_ver}.pptx")
            print(f"{'=' * 60}")
        self.monitor.display_summary(self.results, time.time() - start_time)
        self.state_manager.save_state(state)
        return False


# ============================================================
# Utilities
# ============================================================

def find_project_root() -> Path:
    """Find project root by looking for .git directory, starting from script location."""
    current = Path(__file__).resolve().parent
    for _ in range(10):
        if (current / '.git').exists():
            return current
        parent = current.parent
        if parent == current:
            break
        current = parent
    return Path(__file__).resolve().parent


def _select_account() -> str:
    """Select Claude account (mc or xh)."""
    print("\n可用账户: mc / xh")
    while True:
        choice = input("请选择账户 [直接回车=xh]: ").strip().lower()
        if not choice:
            choice = 'xh'
        if choice in CLAUDE_CONFIG_DIRS:
            config_dir = CLAUDE_CONFIG_DIRS[choice]
            if not os.path.exists(config_dir):
                print(f"⚠️ 配置目录不存在: {config_dir}")
                continue
            os.environ['CLAUDE_CONFIG_DIR'] = config_dir
            print(f"✓ 账户: {choice} ({config_dir})\n")
            return choice
        print(f"❌ 无效选择: {choice}")


# ============================================================
# CLI entry point
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="PPT 专业制作团队调度系统")
    parser.add_argument("--max-rounds", type=int, default=3, help="最大迭代轮次 (default: 3)")
    parser.add_argument("--max-budget", type=float, default=10.0, help="每个 agent 最大预算 USD (default: 10.0)")
    parser.add_argument("--resume", action="store_true", help="从上次中断处恢复")
    args = parser.parse_args()

    _select_account()
    project_root = find_project_root()
    print(f"项目目录: {project_root}")

    # Interactive round selection (overrides --max-rounds if user chooses)
    max_rounds = args.max_rounds
    auto_mode = False
    review_only = False
    init_mode = False
    xlsx_exists = (project_root / "pipeline-progress" / "01-shape_detail.xlsx").exists()

    if not any(a.startswith('--max-rounds') for a in sys.argv[1:]):
        print("\n🎯 请选择运行模式:\n")
        print("  0️⃣  <初始化> ── 全新 PPT 分析，从零构建结构和 prompt")
        print("  1️⃣  1轮 ── 直接生成ppt，⚠️ 不自动优化 prompt / 不验收⚠️")
        print("  2️⃣  2轮 ── 完整流程×2：自动优化 prompt → 暂停人工审核 → 生成ppt → 自动验收 → 下一轮")
        print("  3️⃣  3轮 ── 完整流程×3：自动优化 prompt → 暂停人工审核 → 生成ppt → 自动验收 → 下一轮 × 2")
        print("  4️⃣  🚗全自动×2轮🚗  ── 自动跳过所有Pause (选项2的增强版)")
        print("  5️⃣  🔍单独验收ppt🔍 ── 仅验收，检查最新 PPT 质量\n")
        while True:
            choice = input("请选择 [直接回车=1]: ").strip()
            if not choice:
                choice = '1'
            if choice == '0':
                init_mode = True
                max_rounds = 1
                break
            if choice in ('1', '2', '3'):
                max_rounds = int(choice)
                break
            if choice == '4':
                max_rounds = 2
                auto_mode = True
                break
            if choice == '5':
                review_only = True
                max_rounds = 1
                break
            print("❌ 请输入 0-5")

        # Force routing: Excel 不存在时强制初始化
        if not init_mode and not review_only and not xlsx_exists:
            print(f"  ⚠️  Excel 不存在，自动切换到 [0-初始化] 模式")
            init_mode = True
            max_rounds = 1

        if review_only:
            print(f"✓ 🔍 单独验收模式")
        elif init_mode:
            print(f"✓ 🆕 初始化模式：完整分析 + 构建结构 + 生成 prompt")
        elif auto_mode:
            print(f"✓ 🤖 挂机托管: 全自动 2 轮，跳过所有暂停")
        else:
            labels = {1: "快速出图", 2: "标准打磨", 3: "精雕细琢"}
            print(f"✓ {labels[max_rounds]}（{max_rounds} 轮）")

    # Option 5: review-only mode — find latest pptx + run 04
    if review_only:
        orch = PPTOrchestrator(
            project_root=project_root, max_rounds=1,
            max_budget=args.max_budget,
        )
        # Find latest pptx
        latest_idx = orch._detect_next_version_index() - 1
        if latest_idx < 10:
            print("❌ 未找到任何 claude-ppt *.pptx 文件")
            sys.exit(1)
        version = orch._idx_to_version(latest_idx)
        pptx_name = f"claude-ppt {version}.pptx"
        if not (project_root / pptx_name).exists():
            print(f"❌ 文件不存在: {pptx_name}")
            sys.exit(1)
        print(f"\n🔍 验收目标: {pptx_name}")
        print(f"{'=' * 60}")
        t0 = time.time()
        r = orch._run_pipeline(
            "pipeline/04_shape_diff_test.py",
            ["--target", pptx_name]
        )
        dur = time.time() - t0
        for line in r.stdout.splitlines():
            stripped = line.strip()
            if stripped:
                print(f"  {stripped}")
        passed, fix_type = orch._check_review_passed()
        print(f"{'=' * 60}")
        if passed:
            print(f"✅ {pptx_name} 验收通过！({dur:.0f}s)")
        else:
            print(f"⚠️  {pptx_name} 未通过 (fix_type={fix_type}) ({dur:.0f}s)")
            print(f"   报告: pipeline-progress/04-fix_ppt.md")
            print(f"   数据: pipeline-progress/04-diff_result.json")
        sys.exit(0 if passed else 1)

    # Analyst LLM: auto-decide by mode
    skip_analyst = (max_rounds == 1 and not auto_mode and not init_mode)
    if init_mode:
        print("✓ 初始化模式：Analyst LLM 将增强注释")
    elif skip_analyst:
        print("✓ 热迭代：跳过 Analyst LLM，直接操作 prompt")
    elif auto_mode:
        print("✓ 🤖 全自动模式：热迭代，全程无暂停")
    else:
        print("✓ 热迭代模式：直接操作 prompt")

    orch = PPTOrchestrator(
        project_root=project_root,
        max_rounds=max_rounds,
        max_budget=args.max_budget,
        skip_analyst_first_round=skip_analyst,
        auto_mode=auto_mode,
        init_mode=init_mode,
    )

    success = asyncio.run(orch.run())
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
