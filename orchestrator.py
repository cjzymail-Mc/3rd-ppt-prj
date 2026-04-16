# -*- coding: utf-8 -*-
"""
orchestrator.py - PPT 制作调度系统 (v3 — 3+1 Agent, 局部循环)

按步骤切分 agent，每个 agent 内置自检循环（Python -> LLM 修复，最多 2 次）。
Orchestrator 只做菜单 + agent 调度，不直接跑 pipeline。

Agents:
  step1-analyzer  - 分析 PPT 模板 + 自检
  step2-architect - 构建 prompt + 调 GPT + 自检
  step3-builder   - COM 写入 PPT + 诊断

Menu:
  0 全自动  1 步骤1  2 步骤2  3 步骤3
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
    'yk': os.path.expanduser('~/.claude'),
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
# 3+1 Agent configs (step-based architecture)
# ============================================================

AGENT_CONFIGS = {
    "step1-analyzer": AgentConfig(
        name="step1-analyzer",
        role_file=".claude/agents/step1-analyzer.md",
        output_files=[
            "pipeline-progress/01-shape_detail_com.json",
            "pipeline-progress/01-shape_detail.xlsx",
            "pipeline-progress/02-shape_analysis_map.json",
        ],
    ),
    "step2-architect": AgentConfig(
        name="step2-architect",
        role_file=".claude/agents/step2-architect.md",
        output_files=[
            "pipeline-progress/02-prompt_specs.json",
            "pipeline-progress/03a-build_shape_content.json",
        ],
    ),
    "step3-builder": AgentConfig(
        name="step3-builder",
        role_file=".claude/agents/step3-builder.md",
        output_files=[],  # dynamic: pipeline-output/claude-ppt N.N.pptx
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

            stdout_text = stdout.decode('utf-8', errors='replace')
            stderr_text = stderr.decode('utf-8', errors='replace')
            cost, tokens = self._parse_stream_json(stdout_text)
            duration = time.time() - start_time
            output_files = self._check_output_files(config.output_files)
            status = AgentStatus.COMPLETED if process.returncode == 0 else AgentStatus.FAILED

            # Diagnostic log on failure
            if status == AgentStatus.FAILED:
                diag_path = self.project_root / "debug" / f"agent-{config.name}-{session_id[:8]}.log"
                diag_path.parent.mkdir(parents=True, exist_ok=True)
                diag_path.write_text(
                    f"exit_code: {process.returncode}\n"
                    f"duration: {duration:.0f}s\n\n"
                    f"=== STDERR ===\n{stderr_text}\n\n"
                    f"=== STDOUT (last 3000 chars) ===\n{stdout_text[-3000:]}\n",
                    encoding='utf-8',
                )
                print(f"      诊断日志: {diag_path.name}")

            return ExecutionResult(
                agent_name=config.name, status=status,
                session_id=session_id, exit_code=process.returncode,
                duration=duration, cost=cost, tokens=tokens,
                output_files=output_files,
                error_message=stderr_text if process.returncode != 0 else None,
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
    def __init__(self, project_root: Path, max_retries: int = 1):
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
    "step1-analyzer":  "步骤1-分析师",
    "step2-architect": "步骤2-架构师",
    "step3-builder":   "步骤3-构建师",
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
    """Step-based 3+1 agent dispatcher: each step has its own agent with self-check loop."""

    def __init__(
        self,
        project_root: Path,
        auto_mode: bool = False,
        max_budget: float = 10.0,
    ):
        self.project_root = project_root
        self.auto_mode = auto_mode
        self.executor = AgentExecutor(project_root, max_budget)
        self.error_handler = ErrorHandler(project_root)
        self.state_manager = StateManager(project_root)
        self.monitor = ProgressMonitor()
        self.results: Dict[str, ExecutionResult] = {}

    # ------------------------------------------------------------------
    # Version helpers (used by step3-builder agent via env vars)
    # ------------------------------------------------------------------

    def _detect_next_version_index(self) -> int:
        """Scan pptx files + version tracker, return the next version index.

        Version mapping: 1.0=10, 1.1=11, ..., 1.9=19, 2.0=20, ...
        Returns 10 (=1.0) if no existing versions found.
        """
        max_idx = 9  # so first version = 10 = "1.0"
        ver_pattern = re.compile(r"(\d+)\.(\d+)")

        # Source 1: existing pptx files in pipeline-output/
        pptx_pattern = re.compile(r"^claude-ppt (\d+)\.(\d+)\.pptx$")
        output_dir = self.project_root / "pipeline-output"
        for f in output_dir.glob("claude-ppt *.pptx") if output_dir.exists() else []:
            m = pptx_pattern.match(f.name)
            if m:
                idx = int(m.group(1)) * 10 + int(m.group(2))
                if idx > max_idx:
                    max_idx = idx

        # Source 2: version tracker
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

    @staticmethod
    def _idx_to_version(idx: int) -> str:
        """Convert version index to string: 10->'1.0', 19->'1.9', 20->'2.0'."""
        return f"{idx // 10}.{idx % 10}"

    # ------------------------------------------------------------------
    # Pipeline-first execution (plan4: fast path)
    # ------------------------------------------------------------------

    def _run_pipeline(self, step: int) -> Tuple[bool, str]:
        """Run deterministic pipeline scripts directly via subprocess.

        Returns (success, error_detail).
        success=True means all scripts exited with code 0.
        """
        next_ver = self._idx_to_version(self._detect_next_version_index())

        scripts_map = {
            1: [
                [sys.executable, "pipeline/01_shape_detail.py"],
                [sys.executable, "pipeline/01b_auto_annotate.py"],
                [sys.executable, "pipeline/02_shape_analysis.py"],
            ],
            2: [
                [sys.executable, "pipeline/02_shape_analysis.py"],
                [sys.executable, "pipeline/03a_build_shape.py", "--assemble-only"],
                [sys.executable, "pipeline/03a_build_shape.py", "--execute-prompts"],
            ],
            3: [
                [sys.executable, "pipeline/03b_build_ppt_com.py", "--version", next_ver],
            ],
        }

        scripts = scripts_map.get(step, [])
        env = os.environ.copy()
        env['ORCHESTRATOR_RUNNING'] = 'true'

        for i, cmd in enumerate(scripts):
            label = " ".join(cmd[1:])
            print(f"  [PIPELINE] {label} ...", flush=True)
            start = time.time()
            try:
                proc = subprocess.run(
                    cmd,
                    cwd=str(self.project_root),
                    capture_output=True,
                    text=True,
                    encoding='utf-8',
                    errors='replace',
                    env=env,
                    timeout=300,
                )
            except subprocess.TimeoutExpired:
                return False, f"Timeout (300s): {label}"

            elapsed = time.time() - start
            if proc.returncode != 0:
                detail = (f"Script failed: {label} (exit={proc.returncode}, {elapsed:.1f}s)\n"
                          f"stderr: {proc.stderr[:500]}\nstdout: {proc.stdout[-500:]}")
                print(f"  [PIPELINE] FAIL {label} ({elapsed:.1f}s)")
                return False, detail

            print(f"  [PIPELINE] OK   {label} ({elapsed:.1f}s)")

            # Step 2: after --assemble-only, re-apply saved structural constraints
            if step == 2 and "--assemble-only" in cmd:
                self._reapply_saved_constraints()

        return True, ""

    def _reapply_saved_constraints(self) -> None:
        """Re-apply structural constraints from previous self_check to freshly generated prompt_specs."""
        prev_check = self.project_root / "pipeline-progress" / "02-self_check_result.json"
        if not prev_check.exists():
            return

        try:
            saved = json.loads(prev_check.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, IOError):
            return

        issues = saved.get("issues", [])
        structural = [i for i in issues
                      if "paragraph count mismatch" in i.get("problem", "")
                      or "bullet count mismatch" in i.get("problem", "")]
        if not structural:
            return

        specs_path = self.project_root / "pipeline-progress" / "02-prompt_specs.json"
        if not specs_path.exists():
            return
        specs = json.loads(specs_path.read_text(encoding='utf-8'))

        applied = 0
        for issue in structural:
            shape = issue["shape"]
            problem = issue["problem"]
            prompt = next((p for p in specs["prompts"] if p["shape_name"] == shape), None)
            if not prompt:
                continue

            constraint = ""
            m = re.search(r"paragraph count mismatch: generated (\d+) vs template (\d+)", problem)
            if m:
                constraint = f" 输出必须包含恰好 {int(m.group(2))} 个段落（用空行分隔）。"
            m = re.search(r"bullet count mismatch: generated (\d+) vs template (\d+)", problem)
            if m:
                constraint = f" 输出必须包含恰好 {int(m.group(2))} 个列表项（每项单独一行，以序号或符号开头）。"

            if constraint and constraint not in prompt.get("instruction", ""):
                prompt["instruction"] = prompt.get("instruction", "") + constraint
                applied += 1

        if applied:
            specs_path.write_text(json.dumps(specs, ensure_ascii=False, indent=2), encoding='utf-8')
            print(f"  [INHERIT] 继承上轮 {applied} 条结构约束")

        # Also apply step3 feedback (content overflow → tighter budget)
        self._apply_step3_feedback(specs_path)

    def _apply_step3_feedback(self, specs_path: Path) -> None:
        """Read step3 feedback and inject tighter char limits.

        Three-pronged injection:
        1. Reduce max_chars in budget JSON (hard clamp safety net)
        2. Append constraint to user_instruction in mapping JSON (GPT sees it)
        3. Append constraint to pending prompts JSON (override prompts also get it)
        """
        feedback_path = self.project_root / "pipeline-progress" / "03-feedback_to_step2.json"
        if not feedback_path.exists():
            return

        try:
            feedback = json.loads(feedback_path.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, IOError):
            return

        # Parse overflow issues → {shape_name: max_chars}
        overflows = {}
        for issue in feedback.get("issues", []):
            problem = issue.get("problem", "")
            m = re.search(r'\|\s*(\w[\w\s]*?)\s*\|\s*内容完整性\s*\|\s*内容超长\s*(\d+)字\s*\(上限(\d+)字', problem)
            if m:
                overflows[m.group(1).strip()] = int(m.group(3))

        if not overflows:
            # Clean up even if no parseable issues
            try:
                feedback_path.unlink()
            except OSError:
                pass
            return

        applied = 0

        # --- Prong 1: Reduce max_chars in budget ---
        budget_path = self.project_root / "pipeline-progress" / "02-readability_budget.json"
        if budget_path.exists():
            try:
                bdata = json.loads(budget_path.read_text(encoding='utf-8'))
                for b in bdata.get("budgets", []):
                    if b["shape_name"] in overflows:
                        old = b.get("max_chars", 999)
                        new_limit = overflows[b["shape_name"]]
                        if old > new_limit:
                            b["max_chars"] = new_limit
                            print(f"  [FEEDBACK] {b['shape_name']}: budget max_chars {old} → {new_limit}")
                budget_path.write_text(json.dumps(bdata, ensure_ascii=False, indent=2), encoding='utf-8')
            except (json.JSONDecodeError, IOError):
                pass

        # --- Prong 2: Append constraint to user_instruction in mapping ---
        map_path = self.project_root / "pipeline-progress" / "02-shape_analysis_map.json"
        if map_path.exists():
            try:
                mdata = json.loads(map_path.read_text(encoding='utf-8'))
                for m in mdata.get("mapping", []):
                    if m["shape_name"] in overflows:
                        limit = overflows[m["shape_name"]]
                        constraint = f"严格限制总字数不超过{limit}字（含标点），超长会导致排版溢出。"
                        ui = m.get("user_instruction", "")
                        if constraint not in ui:
                            m["user_instruction"] = (ui + " " + constraint).strip()
                            applied += 1
                map_path.write_text(json.dumps(mdata, ensure_ascii=False, indent=2), encoding='utf-8')
            except (json.JSONDecodeError, IOError):
                pass

        # --- Prong 3: Append constraint to pending prompts ---
        pending_path = self.project_root / "pipeline-progress" / "03a-pending_prompts.json"
        if pending_path.exists():
            try:
                pdata = json.loads(pending_path.read_text(encoding='utf-8'))
                for p in pdata.get("pending", []):
                    if p["shape_name"] in overflows:
                        limit = overflows[p["shape_name"]]
                        suffix = f"\n\n【硬约束】总字数不得超过{limit}字（含标点），超出部分会被截断。请精简表达。"
                        if suffix not in p.get("prompt", ""):
                            p["prompt"] = p.get("prompt", "") + suffix
                pdata["feedback_applied"] = True
                pending_path.write_text(json.dumps(pdata, ensure_ascii=False, indent=2), encoding='utf-8')
            except (json.JSONDecodeError, IOError):
                pass

        if applied:
            print(f"  [FEEDBACK] 注入 step3 反馈: {applied} 条字数约束")

        # Clean up feedback file after consumption
        try:
            feedback_path.unlink()
        except OSError:
            pass

    def _check_filler_content(self) -> List[Dict]:
        """Detect filler/placeholder content in generated shape content."""
        content_json = self.project_root / "pipeline-progress" / "03a-build_shape_content.json"
        if not content_json.exists():
            return []

        filler_phrases = ["暂无", "无有效", "无法汇总", "不可用", "无可归纳", "无可供"]
        issues = []
        try:
            cdata = json.loads(content_json.read_text(encoding='utf-8'))
            for item in cdata.get("items", []):
                if item.get("strategy") == "skip":
                    continue
                text = item.get("content", "")
                for phrase in filler_phrases:
                    if phrase in text:
                        issues.append({
                            "shape": item.get("shape_name", "?"),
                            "problem": f"content contains filler: '{phrase}' — source data may be empty",
                            "fix_hint": "check source xlsx data or re-run step 2 with real data",
                        })
                        break
        except (json.JSONDecodeError, IOError):
            pass
        return issues

    def _run_self_check(self, step: int) -> Tuple[bool, Dict]:
        """Run self-check for given step.

        Returns (passed, result_dict) where result_dict has 'passed', 'issues', 'summary'.
        """
        if step in (1, 2):
            from pipeline.self_check import check_step1, check_step2
            result = check_step1() if step == 1 else check_step2()

            # Step 2: supplement with filler content detection
            if step == 2:
                filler_issues = self._check_filler_content()
                if filler_issues:
                    result["issues"].extend(filler_issues)
                    result["passed"] = False
                    result["summary"] = f"{len(result['issues'])} issue(s) found"

            return result["passed"], result

        if step == 3:
            report_path = self.project_root / "pipeline-progress" / "03b-self_check_report.md"
            if not report_path.exists():
                return False, {"passed": False, "issues": [{"shape": "(all)", "problem": "03b-self_check_report.md not found"}],
                               "summary": "Report file missing — pipeline may have crashed"}
            content = report_path.read_text(encoding='utf-8')
            passed = "结论：PASS" in content
            issues = []
            if not passed:
                for line in content.split('\n'):
                    if '|' in line and '严重' in line and '已修复' not in line:
                        issues.append({"shape": "see report", "problem": line.strip()})

            # Step 3: also check for filler content (catch what step 2 missed)
            filler_issues = self._check_filler_content()
            if filler_issues:
                issues.extend(filler_issues)
                passed = False

            return passed, {"passed": passed, "issues": issues,
                            "summary": "PASS" if passed else f"{len(issues)} issue(s) — check content quality"}

        return False, {"passed": False, "issues": [], "summary": f"Unknown step {step}"}

    def _auto_fix_prompts(self, check_result: Dict) -> bool:
        """Fix structural issues by adding constraints to prompts and re-running GPT.

        Handles paragraph count and bullet count mismatches.
        Returns True if fixes were applied and GPT re-run succeeded.
        """
        issues = check_result.get("issues", [])
        structural = [i for i in issues
                      if "paragraph count mismatch" in i.get("problem", "")
                      or "bullet count mismatch" in i.get("problem", "")]
        if not structural:
            return False

        specs_path = self.project_root / "pipeline-progress" / "02-prompt_specs.json"
        if not specs_path.exists():
            return False
        specs = json.loads(specs_path.read_text(encoding='utf-8'))

        modified = False
        for issue in structural:
            shape = issue["shape"]
            problem = issue["problem"]

            prompt = next((p for p in specs["prompts"] if p["shape_name"] == shape), None)
            if not prompt:
                continue

            constraint = ""
            m = re.search(r"paragraph count mismatch: generated (\d+) vs template (\d+)", problem)
            if m:
                target = int(m.group(2))
                constraint = f" 输出必须包含恰好 {target} 个段落（用空行分隔）。"

            m = re.search(r"bullet count mismatch: generated (\d+) vs template (\d+)", problem)
            if m:
                target = int(m.group(2))
                constraint = f" 输出必须包含恰好 {target} 个列表项（每项单独一行，以序号或符号开头）。"

            if constraint and constraint not in prompt.get("instruction", ""):
                prompt["instruction"] = prompt.get("instruction", "") + constraint
                modified = True
                print(f"  [AUTO-FIX] {shape}: 添加结构约束")

        if not modified:
            return False

        specs_path.write_text(json.dumps(specs, ensure_ascii=False, indent=2), encoding='utf-8')

        print(f"  [AUTO-FIX] 重新调用 GPT...")
        try:
            proc = subprocess.run(
                [sys.executable, "pipeline/03a_build_shape.py", "--execute-prompts"],
                cwd=str(self.project_root),
                capture_output=True, text=True, encoding='utf-8', errors='replace',
                env=os.environ.copy(), timeout=300,
            )
            if proc.returncode == 0:
                print(f"  [AUTO-FIX] GPT 重跑完成")
                return True
            print(f"  [AUTO-FIX] GPT 重跑失败 (exit={proc.returncode})")
            return False
        except subprocess.TimeoutExpired:
            print(f"  [AUTO-FIX] GPT 重跑超时")
            return False

    @staticmethod
    def _classify_issue(issue: Dict) -> str:
        """Classify a self-check issue as 'severe' or 'minor'.

        Severe = would cause step 3 to fail or produce unusable output.
        Minor  = structural/cosmetic (paragraph/bullet count mismatch).
        """
        problem = issue.get("problem", "")
        # Step 1/2 severe keywords
        severe_keywords = ["not found", "is empty", "strategy is empty",
                           "unknown strategy", "no user description"]
        # Step 3 report lines contain "| 严重 |" for severe issues
        if "| 严重 |" in problem:
            return "severe"
        # Filler content is a data quality issue, not a pipeline blocker
        if "filler" in problem:
            return "minor"
        for kw in severe_keywords:
            if kw in problem:
                return "severe"
        return "minor"

    @staticmethod
    def _is_content_issue(issue: Dict) -> bool:
        """Check if a step3 severe issue is content-level (needs step2 redo).

        SSIM is NOT a content issue: template vs generated SSIM is always low
        (different text), and step2 cannot improve it. SSIM issues are treated
        as format issues for LLM agent repair.
        """
        problem = issue.get("problem", "")
        content_markers = ["内容超长", "超出", "content too long", "关键词缺失"]
        return any(m in problem for m in content_markers)

    def _has_step3_feedback(self) -> bool:
        """Check if step3 feedback file exists (indicates content-level loop needed)."""
        return (self.project_root / "pipeline-progress" / "03-feedback_to_step2.json").exists()

    def _save_step3_feedback(self, content_issues: list) -> None:
        """Save step3 content issues as feedback for step2 to consume."""
        feedback_path = self.project_root / "pipeline-progress" / "03-feedback_to_step2.json"
        feedback = {
            "generated_at": datetime.now().isoformat(),
            "source": "step3_self_check",
            "issues": content_issues,
        }
        feedback_path.write_text(json.dumps(feedback, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f"  [SAVE] Step3 反馈 → {feedback_path.name}")

    def _sync_excel_prompts(self) -> None:
        """Compare Excel GPT prompts with JSON; re-run GPT if user edited Excel."""
        pending_path = self.project_root / "pipeline-progress" / "03a-pending_prompts.json"
        if not pending_path.exists():
            return

        try:
            from pipeline.ppt_pipeline_common import read_gpt_prompts_from_xlsx
            excel_prompts = read_gpt_prompts_from_xlsx()
        except Exception as e:
            print(f"  [WARN] 读取 Excel prompt 失败: {e}")
            return

        if not excel_prompts:
            return

        # Load JSON prompts for comparison
        try:
            pd_data = json.loads(pending_path.read_text(encoding='utf-8'))
            json_prompts = {p["shape_name"]: p["prompt"] for p in pd_data.get("pending", [])}
        except (json.JSONDecodeError, IOError):
            return

        # Compare: strip whitespace for robust matching
        changed = []
        for name, excel_text in excel_prompts.items():
            json_text = json_prompts.get(name, "")
            if excel_text.strip() != json_text.strip():
                changed.append(name)

        if not changed:
            return

        print(f"  [SYNC] 检测到 Excel prompt 被编辑 ({len(changed)} 个shape: {', '.join(changed)})")
        print(f"  [SYNC] 自动补跑 GPT 生成新内容...")

        # Re-run 03a --execute-prompts to pick up Excel edits
        ok, err = self._run_pipeline_single("03a_build_shape.py", "--execute-prompts")
        if ok:
            print(f"  [SYNC] GPT 内容已更新")
        else:
            print(f"  [WARN] GPT 重跑失败: {err[:200]}")

    def _run_pipeline_single(self, script: str, *args) -> Tuple[bool, str]:
        """Run a single pipeline script with args. Returns (success, error_detail)."""
        cmd = [sys.executable, f"pipeline/{script}"] + list(args)
        label = f"pipeline/{script} {' '.join(args)}"
        print(f"  [PIPELINE] {label} ...")
        t0 = time.time()
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              cwd=str(self.project_root), timeout=300)
        elapsed = time.time() - t0
        if proc.returncode != 0:
            detail = f"exit={proc.returncode}, stderr: {proc.stderr[:500]}"
            print(f"  [PIPELINE] FAIL {label} ({elapsed:.1f}s)")
            return False, detail
        print(f"  [PIPELINE] OK   {label} ({elapsed:.1f}s)")
        return True, ""

    def _save_self_check_result(self, step: int, check_result: Dict) -> None:
        """Save self-check result to file for later use."""
        save_path = self.project_root / "pipeline-progress" / f"0{step}-self_check_result.json"
        check_result["saved_at"] = datetime.now().isoformat()
        save_path.write_text(json.dumps(check_result, ensure_ascii=False, indent=2), encoding='utf-8')
        print(f"  [SAVE] 自检结果 → {save_path.name}")

    def _make_success_result(self, agent_name: str, duration: float, session_tag: str) -> ExecutionResult:
        """Create a synthetic ExecutionResult for the fast path."""
        return ExecutionResult(
            agent_name=agent_name,
            status=AgentStatus.COMPLETED,
            session_id=session_tag,
            exit_code=0,
            duration=duration,
            cost=0.0,
            tokens=0,
            output_files=[f for f in AGENT_CONFIGS[agent_name].output_files
                          if (self.project_root / f).exists()],
        )

    async def _run_step(self, step: int) -> bool:
        """Run a single step: pipeline -> self-check -> auto-fix -> severity gate.

        Severe issues block the flow; minor issues are warnings (continue).
        """
        agent_names = {1: "step1-analyzer", 2: "step2-architect", 3: "step3-builder"}
        agent_name = agent_names[step]
        step_start = time.time()

        print(f"\n  [FAST] 步骤{step} — 直接运行 Python Pipeline...")

        # Phase 0: Step 3 pre-checks
        if step == 3:
            # 0a: Detect Excel prompt edits → auto re-run GPT if needed
            self._sync_excel_prompts()

            # 0b: Show previous step warnings
            prev_check = self.project_root / "pipeline-progress" / "02-self_check_result.json"
            if prev_check.exists():
                try:
                    prev = json.loads(prev_check.read_text(encoding='utf-8'))
                    prev_issues = prev.get("issues", [])
                    if prev_issues:
                        print(f"  [INFO] Step2 遗留 {len(prev_issues)} 个问题:")
                        for pi in prev_issues:
                            print(f"         - {pi.get('shape','?')}: {pi.get('problem','?')}")
                except (json.JSONDecodeError, IOError):
                    pass

        # Phase 1: Run deterministic pipeline
        pipeline_ok, pipeline_error = self._run_pipeline(step)

        if not pipeline_ok:
            print(f"  [CRASH] Pipeline 脚本异常，启动 LLM Agent 完整流程...")
            return await self._call_agent(agent_name)

        # Phase 2: Self-check
        print(f"  [CHECK] 运行自检...")
        check_passed, check_result = self._run_self_check(step)

        if check_passed:
            duration = time.time() - step_start
            print(f"  [PASS] 步骤{step} 自检通过！({duration:.1f}s)")
            self.results[agent_name] = self._make_success_result(agent_name, duration, "direct-pipeline")
            return True

        # Phase 3: Auto-fix structural issues (step 2 only)
        issues = check_result.get("issues", [])
        print(f"  [FAIL] 自检发现 {len(issues)} 个问题")

        if step == 2 and self._auto_fix_prompts(check_result):
            print(f"  [CHECK] 自动修复后重新自检...")
            check_passed, check_result = self._run_self_check(step)
            if check_passed:
                duration = time.time() - step_start
                print(f"  [PASS] 步骤{step} 自动修复后通过！({duration:.1f}s)")
                self.results[agent_name] = self._make_success_result(agent_name, duration, "direct-autofix")
                return True
            issues = check_result.get("issues", [])
            print(f"  [FAIL] 自动修复后仍有 {len(issues)} 个问题")

        # Phase 4: Severity gate — save results, classify, decide
        self._save_self_check_result(step, check_result)

        severe = [i for i in issues if self._classify_issue(i) == "severe"]
        minor  = [i for i in issues if self._classify_issue(i) == "minor"]

        if minor:
            print(f"  [WARN] {len(minor)} 个轻微问题（不阻断流程）:")
            for m in minor:
                print(f"         - {m['shape']}: {m['problem']}")

        if not severe:
            # Only minor issues — warn and continue
            duration = time.time() - step_start
            print(f"  [CONTINUE] 无严重问题，继续下一步 ({duration:.1f}s)")
            self.results[agent_name] = self._make_success_result(agent_name, duration, "direct-warn")
            return True

        # Has severe issues
        print(f"  [SEVERE] {len(severe)} 个严重问题:")
        for s in severe:
            print(f"           - {s['shape']}: {s['problem']}")

        # Step 3: classify severe issues into content-level (needs step2) vs format-level
        if step == 3:
            content_issues = [s for s in severe if self._is_content_issue(s)]
            format_issues = [s for s in severe if not self._is_content_issue(s)]

            if content_issues:
                # Save issues for step2 feedback, signal caller to loop back
                self._save_step3_feedback(content_issues)
                print(f"  [FEEDBACK] {len(content_issues)} 个内容问题需要 step2 重新生成")
                duration = time.time() - step_start
                self.results[agent_name] = self._make_success_result(agent_name, duration, "needs-step2")
                return False  # signal run() to loop back to step2

            if format_issues:
                # Format issues — try LLM agent
                issues_text = json.dumps(issues, ensure_ascii=False, indent=2)
                failure_context = (f"Self-check result: {check_result.get('summary', '')}\n\n"
                                   f"Issues:\n{issues_text}")
                print(f"  [LLM] 启动 Agent 修复格式问题...")
                return await self._call_agent(agent_name, failure_context=failure_context)

        # Step 1/2: try LLM repair
        issues_text = json.dumps(issues, ensure_ascii=False, indent=2)
        failure_context = (f"Self-check result: {check_result.get('summary', '')}\n\n"
                           f"Issues:\n{issues_text}")
        print(f"  [LLM] 启动 Agent 修复...")
        return await self._call_agent(agent_name, failure_context=failure_context)

    # ------------------------------------------------------------------
    # Agent caller (LLM repair path)
    # ------------------------------------------------------------------

    async def _call_agent(self, agent_name: str, failure_context: Optional[str] = None) -> bool:
        """Call a step agent, return True on success."""
        config = AGENT_CONFIGS.get(agent_name)
        if not config:
            print(f"\n  [ERROR] Unknown agent: {agent_name}")
            return False

        # Build task prompt with runtime context
        template_path = os.environ.get("PPT_TEMPLATE_PATH", "")
        xlsx_path = os.environ.get("PPT_EXCEL_PATH", "")
        next_ver = self._idx_to_version(self._detect_next_version_index())

        context_lines = [
            f"## Runtime Context",
            f"- template_path: {template_path}",
            f"- xlsx_path: {xlsx_path}",
            f"- auto_mode: {self.auto_mode}",
            f"- next_version: {next_ver}",
            f"- project_root: {self.project_root}",
            "",
        ]

        if failure_context:
            context_lines += [
                "## !! REPAIR MODE !!",
                "Attempt 1 (Python Pipeline) 已由 orchestrator 直接执行完毕。",
                "Pipeline 脚本运行成功，但自检发现问题。请跳过 Attempt 1，直接执行 Attempt 2 (LLM 修复)。",
                "",
                "### 自检失败详情：",
                failure_context,
                "",
                "请根据上述自检失败信息，直接执行 Attempt 2 (LLM 修复) 流程。不要重跑 Attempt 1 的 Python Pipeline。",
            ]
        else:
            context_lines.append(
                "请按照你的角色定义执行完整流程（包含 Attempt 1 Python Pipeline + Attempt 2 LLM 修复）。"
            )
        task_prompt = "\n".join(context_lines)

        self.monitor.display_agent_start(agent_name)
        result = await self.error_handler.retry_with_backoff(
            self.executor.run_agent, config, task_prompt
        )
        self.results[agent_name] = result
        self.monitor.display_agent_complete(result)

        return result.status == AgentStatus.COMPLETED

    # ------------------------------------------------------------------
    # File openers
    # ------------------------------------------------------------------

    def _open_xlsx(self) -> None:
        """Open shape_detail.xlsx for user review."""
        xlsx = self.project_root / "pipeline-progress" / "01-shape_detail.xlsx"
        if xlsx.exists():
            try:
                os.startfile(str(xlsx.resolve()))
                print(f"  [已打开] Excel: 01-shape_detail.xlsx")
            except Exception:
                print(f"  [手动打开] {xlsx}")

    def _open_latest_pptx(self) -> None:
        """Open the latest claude-ppt *.pptx for user review."""
        idx = self._detect_next_version_index() - 1
        if idx < 10:
            print("  [WARN] 未找到 claude-ppt *.pptx")
            return
        version = self._idx_to_version(idx)
        pptx = self.project_root / "pipeline-output" / f"claude-ppt {version}.pptx"
        if pptx.exists():
            try:
                os.startfile(str(pptx.resolve()))
                print(f"  [已打开] PPT: claude-ppt {version}.pptx")
            except Exception:
                print(f"  [手动打开] {pptx}")
        else:
            print(f"  [WARN] 文件不存在: claude-ppt {version}.pptx")

    # ------------------------------------------------------------------
    # Main workflow
    # ------------------------------------------------------------------

    async def run(self, step: int) -> bool:
        """Execute workflow for the given step (0=full auto, 1/2/3=single step)."""
        start_time = time.time()

        print(f"\n{'=' * 60}")
        if step == 0:
            print(f"PPT 全自动模式 — 分析 -> 构建 -> 交付")
        else:
            step_names = {1: "分析 PPT 模板", 2: "构建 prompt", 3: "构建 & 交付 PPT"}
            print(f"步骤{step} — {step_names.get(step, '?')}")
        print(f"{'=' * 60}")

        success = False

        if step == 0:
            # Full auto: step1 -> step2 -> step3, with step3→step2 loop
            for s in [1, 2, 3]:
                if not await self._run_step(s):
                    if s == 3 and self._has_step3_feedback():
                        print(f"\n  [LOOP] Step3 内容问题 → 回退 Step2 重新生成...")
                        if await self._run_step(2):
                            print(f"\n  [LOOP] Step2 完成 → 重跑 Step3...")
                            if await self._run_step(3):
                                success = True
                                self._open_latest_pptx()
                        break
                    step_names_zh = {1: "step1-analyzer", 2: "step2-architect", 3: "step3-builder"}
                    print(f"\n  {step_names_zh[s]} 失败，工作流终止。")
                    break
            else:
                success = True
                self._open_latest_pptx()

        elif step in (1, 2):
            success = await self._run_step(step)
            if success:
                self._open_xlsx()

        elif step == 3:
            success = await self._run_step(step)
            if not success and self._has_step3_feedback():
                print(f"\n  [LOOP] Step3 内容问题 → 回退 Step2 重新生成...")
                if await self._run_step(2):
                    print(f"\n  [LOOP] Step2 完成 → 重跑 Step3...")
                    success = await self._run_step(3)
            if success:
                self._open_latest_pptx()

        # Summary
        self.monitor.display_summary(self.results, time.time() - start_time)
        self.state_manager.clear_state()
        return success


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
    """Select Claude account (yk, mc, or xh)."""
    print("\n可用账户: yk / mc / xh")
    while True:
        choice = input("请选择账户 [直接回车=yk]: ").strip().lower()
        if not choice:
            choice = 'yk'
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
# Template selection
# ============================================================

def _load_last_template_choice(project_root: Path) -> Optional[Dict]:
    """Load last template choice from cache file."""
    cache = project_root / ".claude" / "last_template_choice.json"
    if cache.exists():
        try:
            return json.loads(cache.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, IOError):
            pass
    return None


def _save_last_template_choice(project_root: Path, pptx_name: str, xlsx_name: str) -> None:
    """Save template choice to cache file."""
    cache = project_root / ".claude" / "last_template_choice.json"
    cache.parent.mkdir(parents=True, exist_ok=True)
    cache.write_text(json.dumps({"pptx": pptx_name, "xlsx": xlsx_name}, ensure_ascii=False), encoding='utf-8')


def _select_template(project_root: Path) -> None:
    """When template/ has multiple files, let user pick 1 pptx + 1 xlsx.
    Sets PPT_TEMPLATE_PATH and PPT_EXCEL_PATH env vars.
    If only one of each exists, auto-selects without prompting.
    Remembers last choice as default for next run."""
    template_dir = project_root / "template"
    if not template_dir.exists():
        return  # fall back to defaults in ppt_pipeline_common

    pptx_files = sorted(template_dir.glob("*.pptx"))
    xlsx_files = sorted(template_dir.glob("*.xlsx"))

    if not pptx_files or not xlsx_files:
        print("⚠️ template/ 中缺少 .pptx 或 .xlsx 文件")
        return

    # Auto-select if only one of each
    if len(pptx_files) == 1 and len(xlsx_files) == 1:
        os.environ["PPT_TEMPLATE_PATH"] = str(pptx_files[0])
        os.environ["PPT_EXCEL_PATH"] = str(xlsx_files[0])
        print(f"  模板: {pptx_files[0].name}")
        print(f"  数据: {xlsx_files[0].name}")
        return

    # Multiple files — unified numbered list, user picks 2 numbers
    all_files = pptx_files + xlsx_files
    print("\n📂 template/ 中发现多套文件，请选择 1个PPT + 1个Excel:\n")
    for i, f in enumerate(all_files, 1):
        tag = "[PPT]  " if f.suffix == ".pptx" else "[Excel]"
        print(f"  {i}. {tag} {f.name}")

    # Resolve last choice to current indices as default
    default_hint = ""
    last = _load_last_template_choice(project_root)
    if last:
        name_to_idx = {f.name: i + 1 for i, f in enumerate(all_files)}
        pptx_idx = name_to_idx.get(last.get("pptx"))
        xlsx_idx = name_to_idx.get(last.get("xlsx"))
        if pptx_idx and xlsx_idx:
            default_hint = f"{pptx_idx} {xlsx_idx}"

    while True:
        if default_hint:
            raw = input(f"\n请输入2个编号（直接回车={default_hint}）: ").strip()
            if not raw:
                raw = default_hint
        else:
            raw = input(f"\n请输入2个编号（用逗号或空格分隔，如 1,3）: ").strip()
        parts = [s.strip() for s in raw.replace(",", " ").split() if s.strip()]
        if len(parts) != 2:
            print("❌ 请输入恰好2个编号")
            continue
        try:
            idx_a, idx_b = int(parts[0]) - 1, int(parts[1]) - 1
        except ValueError:
            print("❌ 请输入数字编号")
            continue
        if not (0 <= idx_a < len(all_files)) or not (0 <= idx_b < len(all_files)):
            print(f"❌ 编号范围 1-{len(all_files)}")
            continue
        fa, fb = all_files[idx_a], all_files[idx_b]
        exts = {fa.suffix, fb.suffix}
        if exts != {".pptx", ".xlsx"}:
            print("❌ 请选择 1个PPT(.pptx) + 1个Excel(.xlsx)")
            continue
        selected_pptx = fa if fa.suffix == ".pptx" else fb
        selected_xlsx = fa if fa.suffix == ".xlsx" else fb
        break

    os.environ["PPT_TEMPLATE_PATH"] = str(selected_pptx)
    os.environ["PPT_EXCEL_PATH"] = str(selected_xlsx)
    _save_last_template_choice(project_root, selected_pptx.name, selected_xlsx.name)
    print(f"\n  ✓ 模板: {selected_pptx.name}")
    print(f"  ✓ 数据: {selected_xlsx.name}")


# ============================================================
# CLI entry point
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="PPT 制作调度系统 (3+1 Agent)")
    parser.add_argument("--max-budget", type=float, default=10.0, help="每个 agent 最大预算 USD (default: 10.0)")
    args = parser.parse_args()

    _select_account()
    project_root = find_project_root()
    print(f"项目目录: {project_root}")

    # Template selection (before mode selection)
    _select_template(project_root)

    # New 4-option menu
    print("\n\U0001f3af 请选择运行模式:\n")
    print("  0\ufe0f\u20e3  <全自动> \u2500\u2500 分析 \u2192 构建 \u2192 交付ppt")
    print("  1\ufe0f\u20e3  步骤1 \u2500\u2500 分析（新）PPT 模板")
    print("  2\ufe0f\u20e3  步骤2 \u2500\u2500 构建 prompt")
    print("  3\ufe0f\u20e3  步骤3 \u2500\u2500 构建 & 交付 ppt\n")

    while True:
        choice = input("请输入 [0-3]（直接回车=0）: ").strip() or "0"
        if choice in ('0', '1', '2', '3'):
            step = int(choice)
            break
        print("\u274c 请输入 0-3")

    orch = PPTOrchestrator(
        project_root=project_root,
        auto_mode=(step == 0),
        max_budget=args.max_budget,
    )
    success = asyncio.run(orch.run(step))
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
