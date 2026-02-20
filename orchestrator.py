# -*- coding: utf-8 -*-
"""
Orchestrator.py - PPT工程化多Agent调度系统

实现自动化调度6个agents，支持：
- 面向PPT任务的复杂度识别
- 规划→构建→测试→反馈→迭代→交付流水线
- 并发执行（asyncio）
- 失败自动重试（最多3次）
- 实时进度监控和成本控制
- 状态持久化和错误日志
"""

import asyncio
import subprocess
import json
import argparse
import sys
import time
import uuid
import os
import re
from pathlib import Path
from enum import Enum
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass, asdict
from datetime import datetime

# Claude 账户配置目录
CLAUDE_CONFIG_DIRS = {
    'mc': os.path.expanduser('~/.claude-mc'),  # 账户1: mc
    'xh': os.path.expanduser('~/.claude-xh')   # 账户2: xh
}

# Windows 控制台 UTF-8 编码支持
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')


# ============================================================
# 数据结构定义
# ============================================================

class AgentStatus(Enum):
    """Agent执行状态"""
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"


class TaskComplexity(Enum):
    """任务复杂度"""
    MINIMAL = "minimal"      # 2个agents (developer + tester)：微调或单点修复
    SIMPLE = "simple"        # 3个agents：需求澄清→实现→验证
    MODERATE = "moderate"    # 4-5个agents：含评审与迭代
    COMPLEX = "complex"      # 完整6个agents：全流程高精度交付


@dataclass
class AgentConfig:
    """Agent配置"""
    name: str
    role_file: str           # .claude/agents/xx.md
    output_files: List[str]  # 预期输出文件（如PLAN.md）


@dataclass
class ExecutionResult:
    """Agent执行结果"""
    agent_name: str
    status: AgentStatus
    session_id: str
    exit_code: int
    duration: float          # 执行时长（秒）
    cost: float              # 成本（USD）
    tokens: int              # 总tokens
    output_files: List[str]  # 实际生成的文件
    error_message: Optional[str] = None


# ============================================================
# 1. TaskParser - 任务解析器
# ============================================================

class TaskParser:
    """解析用户需求、评估复杂度"""

    def __init__(self, project_root: Path):
        self.project_root = project_root

    def parse(self, user_input: str) -> Tuple[str, TaskComplexity]:
        """根据关键词评估复杂度"""
        user_input_lower = user_input.lower()

        # 复杂任务关键词
        complex_keywords = [
            "ppt", "powerpoint", "com", "模板", "版式", "图表",
            "视觉", "保真", "自动化", "多轮"
        ]

        # 简单任务关键词
        simple_keywords = [
            "修复", "bug", "错别字", "文本替换", "单页", "微调"
        ]

        if any(kw in user_input_lower for kw in complex_keywords):
            return user_input, TaskComplexity.COMPLEX
        elif any(kw in user_input_lower for kw in simple_keywords):
            return user_input, TaskComplexity.SIMPLE
        else:
            return user_input, TaskComplexity.MODERATE

    def is_existing_project(self) -> bool:
        """检测是否是现有项目（有源码）"""
        # 检查常见源码目录
        source_dirs = ['src', 'lib', 'app', 'components', 'packages']
        for dir_name in source_dirs:
            if (self.project_root / dir_name).exists():
                return True

        # 检查配置文件
        config_files = [
            'package.json', 'requirements.txt', 'pom.xml',
            'Cargo.toml', 'go.mod', 'composer.json'
        ]
        for file_name in config_files:
            if (self.project_root / file_name).exists():
                return True

        # 检查是否有 git 提交历史
        try:
            result = subprocess.run(
                ['git', 'log', '--oneline', '-1'],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )
            if result.returncode == 0 and result.stdout.strip():
                return True
        except Exception:
            pass

        return False

    def generate_initial_prompt(self, user_request: str, agent_name: str = None, progress_file: str = None) -> str:
        """生成初始任务提示词"""
        base_prompt = f"""用户需求：{user_request}

你正在执行"PPT 软件工程化交付"任务，请严格按角色职责工作。

🚨 **硬性约束（最高优先级）**
- 以 `new-ppt-workflow.md` 作为流程基线（v4.0执行规范）
- 必须遵循现有脚本架构：`Main.py` + `src/` + `repo-scan-result.md`
- PowerPoint 操作使用 `pywin32 + win32com.client`，Excel 使用 `xlwings + COM API`
- **严禁使用 `python-pptx`**

📊 **PPT流水线（5步）**
- Step1: shape识别与指纹 (`01-shape-detail.py`)
- Step2: shape->源数据映射与Prompt规格 (`02-shape-analysis.py`)
- Step3A: 按shape角色构建内容 (`03-build_shape.py`)
- Step3B: 模板克隆+内容写入 (`03-build_ppt_com.py`)
- Step4: 严格差异测试 (`04-shape_diff_test.py`)

🎯 **per-shape策略矩阵（禁止统一GPT）**
- title: 模板锚点直出（非GPT）
- sample_stat: 问卷样本量聚合（非GPT）
- chart: 每项评分均值提取（非GPT）
- body: extract_info/regex优先，GPT fallback
- long_summary: 模板锚点+数据驱动GPT
- insight: 模板锚点+行动建议GPT

📏 **三层测试阈值**
- Visual Score >= 98 | Readability Score >= 95 | Semantic Coverage = 100

📁 **文件路径规范**
- 所有输出文档保存在项目根目录
- 使用相对路径（如 `PLAN.md`）
- 始终使用正斜杠 `/` 作为路径分隔符
"""

        # 如果是 architect 且是现有项目，添加代码库分析指令
        if agent_name == "architect" and self.is_existing_project():
            # 优先检查是否有现成的代码库扫描结果（节省 token）
            scan_file = self.project_root / "repo-scan-result.md"
            if scan_file.exists():
                try:
                    scan_content = scan_file.read_text(encoding='utf-8')
                    base_prompt += f"""

⚠️ 重要提示：这是一个现有项目！

✅ 已检测到代码库扫描结果文件 `repo-scan-result.md`，你可以直接使用以下分析结果，
**无需重新扫描整个代码库**（节省 token）：

---
{scan_content[:3000]}
---

请基于以上分析结果与 `new-ppt-workflow.md`，生成 `PLAN.md`（保存到项目根目录）：
- 先用 Read 检查 `PLAN.md` 是否已存在：已存在则用 Edit 更新，不存在则用 Write 创建
- 计划必须包含：per-shape策略矩阵、可读性预算、三层测试阈值、迭代策略
- 明确 COM 技术约束：PPT=pywin32，Excel=xlwings+COM，禁用 python-pptx
- 严格复用现有主干：`Main.py` 与 `src/` 中既有能力
"""
                except Exception:
                    scan_file = None  # 读取失败，走原流程

            if not scan_file or not scan_file.exists():
                base_prompt += """

⚠️ 重要提示：这是一个现有项目！

请按以下步骤工作：

1. **第一步：PPT代码库分析**
   - 核心读取：`Main.py`、`src/`、`new-ppt-workflow.md`、`repo-scan-result.md`
   - 识别可复用 COM 能力（shape控制、字体类、图表处理、GPT提炼）
   - **使用 Write 工具**生成 `CODEBASE_ANALYSIS.md`（保存到项目根目录）

2. **第二步：制定交付计划**
   - 先用 Read 检查 `PLAN.md` 是否已存在：已存在则用 Edit 更新，不存在则用 Write 创建
   - 计划必须包含：5步脚本链路、per-shape策略矩阵、可读性预算、三层测试阈值
   - 计划中必须明确：PPT=pywin32、Excel=xlwings+COM、禁用 python-pptx

记住：先复用既有脚本思路，再扩展新能力。
"""

        # 为不同agent注入差异化PPT上下文
        if agent_name == "developer":
            base_prompt += """

🔧 **Developer专项指令**
- 严格执行per-shape策略矩阵，不得全量GPT
- COM写入：文本仅写TextFrame.TextRange.Text，图表仅改ChartData
- 所有COM属性访问必须try-except（getattr对COM对象无效！）
- 可复用函数：GPT_5(), extract_info(), search(), color_key(), smart_color_text()
- 数据缺口必须写入 shape_data_gap_report.md，不得静默跳过
"""
        elif agent_name == "tester":
            base_prompt += """

🧪 **Tester专项指令**
- 对比对象：Template 2.1.pptx 第15页 vs codex X.Y.pptx 第1页
- 三层门禁全部达标才能通过（Visual>=98, Readability>=95, Semantic=100）
- 绝不允许"仅shape数量相同就通过"
- 必须输出 diff_result.json（结构化评分）和 fix-ppt.md（修复建议）
- fix-ppt.md中的修复路由：先改strategy -> 再改prompt -> 再改提取函数
"""
        elif agent_name == "optimizer":
            base_prompt += """

⚡ **Optimizer专项指令**
- 重点优化COM稳定性（重试/释放/超时处理）
- 加速迭代（缓存中间产物、减少重复COM读取）
- 硬约束：不改变视觉结果、不改变测试阈值、不改变策略矩阵
"""
        elif agent_name == "security":
            base_prompt += """

🔒 **Security专项指令**
- 检查所有.py文件中不得硬编码GPT/API key
- 输出路径不得覆盖 src/Template 2.1.pptx 和 2025 数据 v2.2.xlsx
- COM对象在异常路径上必须正确释放
- 产出 SECURITY_AUDIT.md
"""

        # 追加进度文件记录指令
        if progress_file:
            base_prompt += f"""

📝 **进度记录**
- 完成任务后，请将你的工作记录**追加**到进度文件: `{progress_file}`
- 先用 Read 读取现有内容，再用 Write 写入（保留已有内容 + 追加你的部分）
- 记录：你的角色名、任务描述、完成状态、关键输出摘要
"""

        return base_prompt


# ============================================================
# 2. AgentScheduler - 调度规划器
# ============================================================

class AgentScheduler:
    """规划执行阶段、管理agent配置"""

    # Agent配置映射（PPT流水线产物）
    AGENT_CONFIGS = {
        "architect": AgentConfig(
            name="architect",
            role_file=".claude/agents/01-arch.md",
            output_files=["PLAN.md"]
        ),
        "tech_lead": AgentConfig(
            name="tech_lead",
            role_file=".claude/agents/02-tech.md",
            output_files=["PLAN.md"]  # 审核并标注通过
        ),
        "developer": AgentConfig(
            name="developer",
            role_file=".claude/agents/03-dev.md",
            output_files=[
                "shape_detail_com.json", "shape_fingerprint_map.json",
                "shape_analysis_map.json", "prompt_specs.json", "readability_budget.json",
                "build_shape_content.json", "content_validation_report.md",
                "build-ppt-report.md", "post_write_readback.json"
            ]
        ),
        "tester": AgentConfig(
            name="tester",
            role_file=".claude/agents/04-test.md",
            output_files=["fix-ppt.md", "diff_result.json", "diff_semantic_report.md"]
        ),
        "optimizer": AgentConfig(
            name="optimizer",
            role_file=".claude/agents/05-opti.md",
            output_files=[]  # 直接优化代码
        ),
        "security": AgentConfig(
            name="security",
            role_file=".claude/agents/06-secu.md",
            output_files=["SECURITY_AUDIT.md"]
        ),
    }

    def plan_execution(self, complexity: TaskComplexity) -> List[List[str]]:
        """
        根据复杂度规划执行阶段
        返回：[[Phase1 agents], [Phase2 agents], ...]
        """
        if complexity == TaskComplexity.MINIMAL:
            return [
                ["developer"],
                ["tester"]
            ]
        elif complexity == TaskComplexity.SIMPLE:
            return [
                ["architect"],
                ["developer"],
                ["tester"]
            ]
        elif complexity == TaskComplexity.MODERATE:
            return [
                ["architect"],
                ["tech_lead"],
                ["developer"],
                ["tester", "optimizer"]
            ]
        else:  # COMPLEX
            return [
                ["architect"],
                ["tech_lead"],
                ["developer"],
                ["tester"],
                ["optimizer", "security"]
            ]

    def get_agent_config(self, agent_name: str) -> AgentConfig:
        """获取Agent配置"""
        return self.AGENT_CONFIGS[agent_name]

    def get_all_agent_names(self) -> List[str]:
        """获取所有可用的 agent 名称"""
        return list(self.AGENT_CONFIGS.keys())


# ============================================================
# 2.5 ManualTaskParser - 手动任务解析器
# ============================================================

class ManualTaskParser:
    """
    解析手动指定的 agent 任务

    支持语法：
      - @tech_lead 评审PPT方案
      - @developer 构建shape && @tester 差异测试
      - @architect 规划 -> @developer 实现 -> @tester 验证
      - @tech_lead 复盘 -> (@developer 修复 && @optimizer 提效) -> @tester 回归
    """

    # Agent 别名映射
    ALIASES = {
        "arch": "architect",
        "架构": "architect",
        "tech": "tech_lead",
        "技术": "tech_lead",
        "dev": "developer",
        "开发": "developer",
        "test": "tester",
        "测试": "tester",
        "opti": "optimizer",
        "优化": "optimizer",
        "sec": "security",
        "安全": "security",
    }

    def __init__(self, project_root: Path = None):
        self.scheduler = AgentScheduler()
        self.valid_agents = self.scheduler.get_all_agent_names()
        self.project_root = project_root or find_project_root()

    def is_manual_mode(self, user_input: str) -> bool:
        """检测是否是手动指定模式（包含 @agent，支持中文别名）"""
        return bool(re.search(r'@[\w\u4e00-\u9fff]+', user_input))

    def resolve_agent_name(self, name: str) -> Optional[str]:
        """解析 agent 名称（支持别名）"""
        name = name.lower().strip()
        if name in self.valid_agents:
            return name
        if name in self.ALIASES:
            return self.ALIASES[name]
        return None

    def parse(self, user_input: str) -> Tuple[List[List[Tuple[str, str]]], bool]:
        """
        解析手动指定的任务

        Args:
            user_input: 用户输入，如 "@tech_lead 审核代码 -> @developer 修复"

        Returns:
            (phases, success)
            phases: [[("agent_name", "task"), ...], ...]  # 每个 phase 包含并行的 agent-task 对
            success: 解析是否成功
        """
        user_input = user_input.strip()

        # 按 -> 分割成多个 phase（串行）
        phase_strs = re.split(r'\s*->\s*', user_input)

        phases = []
        for phase_str in phase_strs:
            phase_str = phase_str.strip()

            # 去除括号
            if phase_str.startswith('(') and phase_str.endswith(')'):
                phase_str = phase_str[1:-1].strip()

            # 按 && 分割成并行任务
            parallel_strs = re.split(r'\s*&&\s*', phase_str)

            phase_tasks = []
            for task_str in parallel_strs:
                task_str = task_str.strip()

                # 解析 @agent_name 任务描述（支持中文别名）
                match = re.match(r'@([\w\u4e00-\u9fff]+)\s+(.+)$', task_str)
                if match:
                    agent_raw, task = match.groups()
                    agent_name = self.resolve_agent_name(agent_raw)

                    if agent_name is None:
                        print(f"❌ 未知的 agent: @{agent_raw}")
                        print(f"   可用的 agents: {', '.join(self.valid_agents)}")
                        return [], False

                    task = task.strip()

                    # 检测是否为 .md 文件引用
                    if task.endswith('.md'):
                        md_file = self.project_root / task
                        if md_file.exists():
                            try:
                                task = md_file.read_text(encoding='utf-8')
                                print(f"📄 @{agent_name}: 从 {md_file.name} 读取任务描述")
                            except Exception as e:
                                print(f"⚠️ 无法读取 {task}: {e}")
                                return [], False
                        else:
                            print(f"❌ 文件不存在: {task}")
                            print(f"   完整路径: {md_file}")
                            return [], False

                    phase_tasks.append((agent_name, task))
                else:
                    print(f"❌ 无法解析任务: {task_str}")
                    print(f"   请使用格式: @agent_name 任务描述")
                    return [], False

            if phase_tasks:
                phases.append(phase_tasks)

        return phases, True

    def preview(self, phases: List[List[Tuple[str, str]]]) -> None:
        """预览执行计划"""
        print(f"\n📋 手动指定模式 - 执行计划：")
        print(f"   共 {len(phases)} 个阶段")

        for i, phase in enumerate(phases, 1):
            if len(phase) == 1:
                agent, task = phase[0]
                print(f"\n   Phase {i}: @{agent}")
                print(f"      任务: {task[:50]}{'...' if len(task) > 50 else ''}")
            else:
                agents = [f"@{a}" for a, _ in phase]
                print(f"\n   Phase {i}: {' && '.join(agents)}  (并行)")
                for agent, task in phase:
                    print(f"      @{agent}: {task[:40]}{'...' if len(task) > 40 else ''}")


# ============================================================
# 3. AgentExecutor - 执行引擎
# ============================================================

class AgentExecutor:
    """执行claude -p命令、管理子进程、解析输出"""

    def __init__(self, project_root: Path, max_budget: float = 10.0, max_concurrent: int = 2):
        self.project_root = project_root
        self.max_budget = max_budget
        self._semaphore = asyncio.Semaphore(max_concurrent)  # 限制并发数，避免API限流

    def _parse_agent_file(self, content: str) -> Tuple[Dict, str]:
        """
        解析 agent 文件，分离 YAML frontmatter 和正文

        Args:
            content: agent 文件的完整内容

        Returns:
            (metadata, body) - 元数据字典和正文内容
        """
        content = content.strip()

        # 检查是否以 --- 开头
        if not content.startswith('---'):
            # 没有 frontmatter，整个内容都是正文
            return {}, content

        # 更健壮的正则匹配 YAML frontmatter
        # 支持：---\n...\n--- 或 ---\r\n...\r\n--- (Windows换行)
        # 也支持 frontmatter 后面没有换行的情况
        patterns = [
            r'^---[ \t]*[\r\n]+(.*?)[\r\n]+---[ \t]*[\r\n]+(.*)$',  # 标准格式
            r'^---[ \t]*[\r\n]+(.*?)[\r\n]+---[ \t]*$',  # frontmatter 后无内容
            r'^---[ \t]*[\r\n]+---[ \t]*[\r\n]+(.*)$',  # 空 frontmatter
        ]

        metadata = {}
        body = content

        for i, pattern in enumerate(patterns):
            match = re.match(pattern, content, re.DOTALL)
            if match:
                if i == 2:  # 空 frontmatter 模式
                    body = match.group(1).strip() if match.lastindex >= 1 else ""
                elif i == 1:  # frontmatter 后无内容
                    frontmatter_str = match.group(1)
                    body = ""
                    # 解析 frontmatter
                    for line in frontmatter_str.split('\n'):
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            key, value = line.split(':', 1)
                            metadata[key.strip()] = value.strip()
                else:  # 标准格式
                    frontmatter_str = match.group(1)
                    body = match.group(2).strip()
                    # 解析 frontmatter
                    for line in frontmatter_str.split('\n'):
                        line = line.strip()
                        if ':' in line and not line.startswith('#'):
                            key, value = line.split(':', 1)
                            metadata[key.strip()] = value.strip()
                break

        return metadata, body

    def _check_architect_violation(self, line: str) -> Optional[str]:
        """
        检查 stream-json 单行是否显示 architect 尝试写入非 .md 文件

        检测逻辑：同一行 JSON 中同时出现 Write/Edit 工具名 + 非 .md 的 file_path
        """
        line = line.strip()
        if not line:
            return None

        # 快速预检：必须同时包含 file_path 和 Write/Edit
        if 'file_path' not in line:
            return None
        if 'Write' not in line and 'Edit' not in line:
            return None

        # 提取 file_path 值
        match = re.search(r'"file_path"\s*:\s*"([^"]+)"', line)
        if match:
            file_path = match.group(1)
            if not file_path.lower().endswith('.md'):
                return f"🚫 Architect 越权！尝试修改非 .md 文件: {file_path}"

        return None

    async def _monitor_architect_stream(
        self,
        process: asyncio.subprocess.Process
    ) -> Tuple[str, str, Optional[str]]:
        """
        实时监控 architect 的 stream-json 输出流

        逐行读取 stdout，检测到 Write/Edit 非 .md 文件时立刻 kill 进程。
        比事后校验更高效：节省时间和 token。

        Returns:
            (stdout_str, stderr_str, violation_msg)
            violation_msg 为 None 表示正常完成
        """
        stdout_lines = []
        violation = None
        MAX_LINES = 5000  # 防止 OOM

        while True:
            line = await process.stdout.readline()
            if not line:
                break
            line_str = line.decode('utf-8', errors='replace')
            stdout_lines.append(line_str)
            if len(stdout_lines) > MAX_LINES:
                stdout_lines = stdout_lines[-MAX_LINES:]

            # 实时检测越权行为
            violation = self._check_architect_violation(line_str)
            if violation:
                print(f"\n\n{violation}")
                print(f"⏹️  立即终止 Architect 进程，节省后续 token 消耗")
                process.kill()
                try:
                    await asyncio.wait_for(process.wait(), timeout=5.0)
                except asyncio.TimeoutError:
                    pass
                break

        # 读取剩余 stderr
        stderr_data = await process.stderr.read()
        stderr_str = stderr_data.decode('utf-8', errors='replace')

        stdout_str = ''.join(stdout_lines)
        return stdout_str, stderr_str, violation

    async def run_agent(
        self,
        config: AgentConfig,
        task_prompt: str,
        timeout: int = 600,
        session_id: Optional[str] = None
    ) -> ExecutionResult:
        """
        执行单个agent（异步）

        Args:
            config: Agent配置
            task_prompt: 任务提示词
            timeout: 超时时间（秒）
            session_id: 会话ID（可选，不提供则自动生成）
        """
        if session_id is None:
            session_id = str(uuid.uuid4())
        start_time = time.time()

        # 为 architect 追加权限限制（防止越权修改代码文件）
        if config.name == "architect":
            task_prompt += """

---
## ⚠️ 权限限制（必须严格遵守）

你是 Architect Agent，**只能**写入以下类型的文件：
- PLAN.md（实施计划）
- CODEBASE_ANALYSIS.md（代码库分析）
- 其他 .md 文档文件

❌ **绝对禁止**：
- 不得创建或修改任何源代码文件（.py, .js, .ts, .java, .go 等）
- 不得修改配置文件（package.json, requirements.txt 等）
- 不得运行测试或构建命令
- 不得执行任何代码实现

违反以上限制将导致你的输出被回滚。
"""

        # 读取并解析 agent 角色配置（分离 YAML frontmatter）
        role_file = self.project_root / config.role_file
        try:
            with open(role_file, 'r', encoding='utf-8') as f:
                content = f.read()
            metadata, role_prompt = self._parse_agent_file(content)
        except FileNotFoundError:
            return ExecutionResult(
                agent_name=config.name,
                status=AgentStatus.FAILED,
                session_id=session_id,
                exit_code=1,
                duration=0,
                cost=0,
                tokens=0,
                output_files=[],
                error_message=f"角色配置文件不存在: {config.role_file}"
            )

        # 从 metadata 中获取 model（如果有的话）
        agent_model = metadata.get('model', 'sonnet')

        # Windows 命令行长度限制修复（WinError 206）：
        # 当 task_prompt 过长时，写入临时文件让 agent 通过 Read 工具读取
        prompt_temp_file = None
        actual_prompt = task_prompt
        if len(task_prompt) > 4000:
            prompt_temp_file = self.project_root / ".claude" / f"prompt_{config.name}_{session_id[:8]}.txt"
            prompt_temp_file.parent.mkdir(parents=True, exist_ok=True)
            prompt_temp_file.write_text(task_prompt, encoding='utf-8')
            actual_prompt = f"请先使用 Read 工具读取文件 `{prompt_temp_file}` 获取完整任务指令，然后严格按照指令执行你的职责。"

        # 构建 claude 命令
        # 注意：architect 使用 plan 模式（只读），其他 agents 使用 skip-permissions（可写）
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
        ]

        # 所有 agent 使用 skip-permissions（architect 由 hook + stream monitor 防护）
        cmd.append("--dangerously-skip-permissions")

        # 进度指示器
        async def progress_indicator(agent_name: str, start: float):
            """周期性打印进度信息"""
            indicators = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
            idx = 0
            while True:
                elapsed = time.time() - start
                print(f"\r      {indicators[idx]} {agent_name} 工作中... ({elapsed:.0f}s)", end="", flush=True)
                idx = (idx + 1) % len(indicators)
                await asyncio.sleep(1)

        # 使用 semaphore 限制并发数（异步执行子进程）
        async with self._semaphore:
          try:
            # 设置环境变量，用于 hook 检测
            env = os.environ.copy()
            env['ORCHESTRATOR_RUNNING'] = 'true'
            env['ORCHESTRATOR_AGENT'] = config.name  # Hook 用此变量识别当前 agent

            process = await asyncio.create_subprocess_exec(
                *cmd,
                cwd=str(self.project_root),
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                env=env
            )

            # 启动进度指示器
            progress_task = asyncio.create_task(progress_indicator(config.name, start_time))

            # 等待完成（带超时）
            # architect 使用实时流监控，检测到写入非 .md 文件立刻终止
            violation_msg = None
            try:
                if config.name == "architect":
                    stdout_str, stderr_str, violation_msg = await asyncio.wait_for(
                        self._monitor_architect_stream(process),
                        timeout=timeout
                    )
                    stdout = stdout_str.encode('utf-8')
                    stderr = stderr_str.encode('utf-8')
                else:
                    stdout, stderr = await asyncio.wait_for(
                        process.communicate(),
                        timeout=timeout
                    )
            except asyncio.TimeoutError:
                progress_task.cancel()
                print()  # 换行
                process.kill()
                try:
                    await asyncio.wait_for(process.wait(), timeout=5.0)
                except asyncio.TimeoutError:
                    pass  # 强制终止后仍超时，忽略
                return ExecutionResult(
                    agent_name=config.name,
                    status=AgentStatus.FAILED,
                    session_id=session_id,
                    exit_code=-1,
                    duration=time.time() - start_time,
                    cost=0,
                    tokens=0,
                    output_files=[],
                    error_message=f"执行超时（{timeout}s）"
                )
            finally:
                # 停止进度指示器
                progress_task.cancel()
                try:
                    await progress_task
                except asyncio.CancelledError:
                    pass
                print()  # 换行，结束进度行

            # 如果 architect 被实时拦截，直接返回失败
            if violation_msg:
                cost, tokens = self._parse_stream_json(stdout.decode('utf-8', errors='replace'))
                return ExecutionResult(
                    agent_name=config.name,
                    status=AgentStatus.FAILED,
                    session_id=session_id,
                    exit_code=-2,
                    duration=time.time() - start_time,
                    cost=cost,
                    tokens=tokens,
                    output_files=[],
                    error_message=violation_msg
                )

            # 解析stream-json输出获取成本和tokens
            cost, tokens = self._parse_stream_json(stdout.decode('utf-8', errors='replace'))

            duration = time.time() - start_time

            # 检查输出文件是否生成
            output_files = self._check_output_files(config.output_files)

            status = AgentStatus.COMPLETED if process.returncode == 0 else AgentStatus.FAILED

            return ExecutionResult(
                agent_name=config.name,
                status=status,
                session_id=session_id,
                exit_code=process.returncode,
                duration=duration,
                cost=cost,
                tokens=tokens,
                output_files=output_files,
                error_message=stderr.decode('utf-8', errors='replace') if process.returncode != 0 else None
            )

          except Exception as e:
            return ExecutionResult(
                agent_name=config.name,
                status=AgentStatus.FAILED,
                session_id=session_id,
                exit_code=1,
                duration=time.time() - start_time,
                cost=0,
                tokens=0,
                output_files=[],
                error_message=str(e)
            )
          finally:
            # 清理临时 prompt 文件
            if prompt_temp_file and prompt_temp_file.exists():
                try:
                    prompt_temp_file.unlink()
                except (OSError, PermissionError):
                    pass

    def run_agent_interactive(
        self,
        config: AgentConfig,
        task_prompt: str,
        session_id: Optional[str] = None
    ) -> ExecutionResult:
        """
        以交互式模式执行agent（用于architect阶段）
        自动发送初始任务，用户可继续讨论直到满意

        Returns:
            ExecutionResult with basic info (详细成本等需手动检查)
        """
        if session_id is None:
            session_id = str(uuid.uuid4())
        start_time = time.time()

        # 读取并解析 agent 角色配置（分离 YAML frontmatter）
        role_file = self.project_root / config.role_file
        try:
            with open(role_file, 'r', encoding='utf-8') as f:
                content = f.read()
            metadata, role_prompt = self._parse_agent_file(content)
        except FileNotFoundError:
            return ExecutionResult(
                agent_name=config.name,
                status=AgentStatus.FAILED,
                session_id=session_id,
                exit_code=1,
                duration=0,
                cost=0,
                tokens=0,
                output_files=[],
                error_message=f"角色配置文件不存在: {config.role_file}"
            )

        # 从 metadata 中获取 model（如果有的话）
        agent_model = metadata.get('model', 'sonnet')

        # 构建初始提示词，明确指定输出文件位置
        output_instruction = """

---

## 输出要求

请将计划文件保存到项目根目录（使用相对路径）：
- `PLAN.md` - 实施计划（必须生成）：先检查是否已存在，已存在则用 Edit 更新，不存在则用 Write 创建
- `CODEBASE_ANALYSIS.md` - 代码库分析（如果是现有项目）

完成后请告知用户已生成上述文件，并输入 /exit 退出。
"""
        full_prompt = f"{role_prompt}\n\n---\n\n{task_prompt}{output_instruction}"

        print(f"\n{'='*60}", flush=True)
        print(f"🎯 启动交互式规划会话 - {config.name}", flush=True)
        print(f"{'='*60}", flush=True)
        print(f"📋 初始任务将自动发送，无需手动输入", flush=True)
        print(f"💡 你可以继续与 {config.name} 讨论，直到满意", flush=True)
        print(f"📄 完成后输入 /exit 退出会话", flush=True)
        print(f"{'='*60}\n", flush=True)

        # 构建交互式 claude 命令
        # 直接传入 prompt 参数，claude 会自动执行后保持交互模式
        # 注意：--max-budget-usd 只在 --print 模式下生效，交互式模式下忽略
        cmd = [
            "claude",
            "--model", agent_model,
            "--dangerously-skip-permissions",  # hooks 提供 architect 权限管控
            "--append-system-prompt", role_prompt,  # 角色定义作为系统提示
            task_prompt + output_instruction,  # 用户任务作为初始 prompt
        ]

        # 同步执行（阻塞等待用户交互）
        try:
            # 设置环境变量，用于 hook 检测
            env = os.environ.copy()
            env['ORCHESTRATOR_RUNNING'] = 'true'
            env['ORCHESTRATOR_AGENT'] = config.name  # Hook 用此变量识别当前 agent

            # 使用 subprocess.run，不重定向 stdin/stdout/stderr，让用户直接交互
            process = subprocess.run(
                cmd,
                cwd=str(self.project_root),
                env=env
            )

            duration = time.time() - start_time

            # 检查输出文件是否生成
            output_files = self._check_output_files(config.output_files)

            status = AgentStatus.COMPLETED if process.returncode == 0 else AgentStatus.FAILED

            # 交互式模式无法准确获取成本，返回估算值
            return ExecutionResult(
                agent_name=config.name,
                status=status,
                session_id=session_id,
                exit_code=process.returncode,
                duration=duration,
                cost=0.0,  # 交互式模式成本需手动查看
                tokens=0,
                output_files=output_files,
                error_message=None if process.returncode == 0 else "交互式会话异常退出"
            )

        except Exception as e:
            return ExecutionResult(
                agent_name=config.name,
                status=AgentStatus.FAILED,
                session_id=session_id,
                exit_code=1,
                duration=time.time() - start_time,
                cost=0,
                tokens=0,
                output_files=[],
                error_message=str(e)
            )

    async def run_parallel(
        self,
        configs: List[AgentConfig],
        prompts: Dict[str, str]
    ) -> Dict[str, ExecutionResult]:
        """并发执行多个agents"""
        tasks = [
            self.run_agent(config, prompts[config.name])
            for config in configs
        ]
        results = await asyncio.gather(*tasks)
        return {r.agent_name: r for r in results}

    def _parse_stream_json(self, stdout: str, verbose: bool = False) -> Tuple[float, int]:
        """
        解析stream-json输出获取成本和tokens（增强版）

        支持多种 JSON 结构：
        - {"cost": x, "tokens": y}
        - {"cost_usd": x, "total_tokens": y}
        - {"type": "result", "cost": x, ...}
        - {"usage": {"input_tokens": x, "output_tokens": y}}

        Args:
            stdout: claude 命令的标准输出
            verbose: 是否输出详细日志

        Returns:
            (cost, tokens) 元组
        """
        if not stdout or not stdout.strip():
            if verbose:
                print("  [调试] stream-json 输出为空")
            return 0.0, 0

        lines = stdout.strip().split('\n')
        best_cost = 0.0
        best_tokens = 0

        # 从后往前查找有效的 JSON 行
        for line in reversed(lines):
            line = line.strip()
            if not line:
                continue

            try:
                data = json.loads(line)

                # 优先查找 result 类型消息（通常是最终结果）
                if data.get('type') == 'result':
                    cost = data.get('cost_usd', data.get('cost', 0))
                    tokens = data.get('total_tokens', data.get('tokens', 0))
                    if cost > 0 or tokens > 0:
                        return float(cost), int(tokens)

                # 尝试多种字段名获取 cost（避免 or 短路：0 or x 返回 x）
                cost = data.get('cost_usd') if 'cost_usd' in data else data.get('cost', 0)

                # 尝试多种字段名获取 tokens
                tokens = data.get('tokens', 0)
                if tokens == 0:
                    tokens = data.get('total_tokens', 0)
                if tokens == 0 and 'usage' in data:
                    usage = data['usage']
                    tokens = usage.get('total_tokens', 0)
                    # 如果没有 total_tokens，尝试计算 input + output
                    if tokens == 0:
                        input_tokens = usage.get('input_tokens', 0)
                        output_tokens = usage.get('output_tokens', 0)
                        tokens = input_tokens + output_tokens

                # 保留找到的最大值（避免中间行覆盖最终结果）
                if cost > best_cost:
                    best_cost = float(cost)
                if tokens > best_tokens:
                    best_tokens = int(tokens)

                # 如果找到有效数据就返回
                if best_cost > 0 or best_tokens > 0:
                    return best_cost, best_tokens

            except json.JSONDecodeError as e:
                # 这行不是有效 JSON，继续尝试下一行
                if verbose:
                    print(f"  [调试] JSON 解析失败: {str(e)[:50]}")
                continue
            except (TypeError, ValueError, AttributeError) as e:
                if verbose:
                    print(f"  [调试] 数据类型转换失败: {e}")
                continue

        # 返回找到的最佳值（可能是 0）
        if verbose and best_cost == 0 and best_tokens == 0:
            print("  [调试] 未在输出中找到成本/tokens 信息")
        return best_cost, best_tokens

    def _check_output_files(self, expected_files: List[str]) -> List[str]:
        """检查输出文件是否存在"""
        existing = []
        for file in expected_files:
            file_path = self.project_root / file
            if file_path.exists():
                existing.append(file)
        return existing


# ============================================================
# 4. StateManager - 状态管理器
# ============================================================

class StateManager:
    """持久化状态到.claude/state.json"""

    def __init__(self, project_root: Path):
        self.project_root = project_root
        self.state_file = project_root / ".claude" / "state.json"
        self.state_file.parent.mkdir(parents=True, exist_ok=True)

    def save_state(self, state: Dict) -> None:
        """原子化保存状态"""
        # 确保目录存在
        self.state_file.parent.mkdir(parents=True, exist_ok=True)
        temp_file = self.state_file.with_suffix('.tmp')
        with open(temp_file, 'w', encoding='utf-8') as f:
            json.dump(state, f, indent=2, ensure_ascii=False)
        temp_file.replace(self.state_file)

    def load_state(self) -> Optional[Dict]:
        """加载状态"""
        if self.state_file.exists():
            with open(self.state_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None

    def clear_state(self) -> None:
        """清除状态"""
        if self.state_file.exists():
            self.state_file.unlink()


# ============================================================
# 5. ErrorHandler - 错误处理器
# ============================================================

class ErrorHandler:
    """重试机制、错误日志"""

    def __init__(self, project_root: Path, max_retries: int = 3):
        self.project_root = project_root
        self.max_retries = max_retries
        self.backoff_seconds = [5, 10, 20]
        self.error_log_file = project_root / ".claude" / "error_log.json"
        self.error_log_file.parent.mkdir(parents=True, exist_ok=True)

    async def retry_with_backoff(
        self,
        func,
        *args,
        **kwargs
    ) -> ExecutionResult:
        """
        重试最多3次，间隔5s/10s/20s
        3次失败后记录错误并返回
        """
        for attempt in range(self.max_retries):
            result = await func(*args, **kwargs)

            if result.status == AgentStatus.COMPLETED:
                return result

            # 如果不是最后一次尝试，等待后重试
            if attempt < self.max_retries - 1:
                wait_time = self.backoff_seconds[attempt]
                print(f"  [重试] {result.agent_name} 失败，{wait_time}秒后重试（{attempt + 1}/{self.max_retries}）")
                await asyncio.sleep(wait_time)

        # 3次重试后仍失败 → 记录错误
        self.log_error(result)
        return result

    def log_error(self, result: ExecutionResult) -> None:
        """记录错误到error_log.json"""
        error_entry = {
            "timestamp": datetime.now().isoformat(),
            "agent": result.agent_name,
            "exit_code": result.exit_code,
            "error_message": result.error_message,
            "retry_count": self.max_retries,
            "session_id": result.session_id
        }

        # 追加到错误日志（带容错）
        errors = []
        if self.error_log_file.exists():
            try:
                with open(self.error_log_file, 'r', encoding='utf-8') as f:
                    errors = json.load(f)
            except (json.JSONDecodeError, IOError):
                errors = []  # 文件损坏时重置为空列表

        errors.append(error_entry)

        with open(self.error_log_file, 'w', encoding='utf-8') as f:
            json.dump(errors, f, indent=2, ensure_ascii=False)


# ============================================================
# 6. ProgressMonitor - 进度监控器
# ============================================================

class ProgressMonitor:
    """实时进度显示、汇总报告"""

    def __init__(self, verbose: bool = False):
        self.verbose = verbose

    def display_phase_start(self, phase_num: int, agents: List[str]) -> None:
        """显示当前执行阶段"""
        print(f"\n{'='*60}")
        print(f"Phase {phase_num}: {', '.join(agents)}")
        print(f"{'='*60}")

    def display_agent_start(self, agent_name: str, session_id: str) -> None:
        """显示agent启动"""
        print(f"  [启动] {self._get_agent_display_name(agent_name)} (session: {session_id})")

    def display_agent_complete(
        self,
        result: ExecutionResult
    ) -> None:
        """显示agent完成"""
        status_icon = "✅" if result.status == AgentStatus.COMPLETED else "❌"

        # 如果有成本信息则显示，否则显示 tokens
        if result.cost > 0:
            cost_info = f"${result.cost:.4f}"
        elif result.tokens > 0:
            cost_info = f"{result.tokens:,} tokens"
        else:
            cost_info = "Pro 订阅"

        print(f"  {status_icon} {self._get_agent_display_name(result.agent_name)} - "
              f"{result.status.value} (耗时 {result.duration:.1f}s, {cost_info})")

        if result.status == AgentStatus.FAILED and result.error_message:
            print(f"      错误: {result.error_message[:100]}")

    def display_summary(
        self,
        all_results: Dict[str, ExecutionResult],
        total_duration: float
    ) -> None:
        """显示执行汇总"""
        total_cost = sum(r.cost for r in all_results.values())
        total_tokens = sum(r.tokens for r in all_results.values())

        print(f"\n{'='*60}")
        print(f"执行完成 - 总耗时 {total_duration:.1f}s")
        print(f"{'='*60}")

        # 智能显示成本或 tokens
        if total_cost > 0:
            print(f"总成本: ${total_cost:.4f}")
            print(f"总tokens: {total_tokens:,}")
        elif total_tokens > 0:
            print(f"总tokens: {total_tokens:,} (Pro 订阅模式)")
        else:
            print(f"计费模式: Pro 订阅（固定月费）")

        print(f"\nAgent 执行结果:")

        for name, result in all_results.items():
            status_icon = "✅" if result.status == AgentStatus.COMPLETED else "❌"

            # 显示成本或 tokens
            if result.cost > 0:
                cost_info = f"${result.cost:.4f}"
            elif result.tokens > 0:
                cost_info = f"{result.tokens:,} tokens"
            else:
                cost_info = "Pro 订阅"

            print(f"  {status_icon} {name:12s} - {result.status.value:10s} "
                  f"(耗时 {result.duration:.1f}s, {cost_info})")

            if result.output_files:
                for file in result.output_files:
                    print(f"      → 输出: {file}")

    def _get_agent_display_name(self, agent_name: str) -> str:
        """获取agent显示名称"""
        name_map = {
            "architect": "PPT流程架构师",
            "tech_lead": "PPT技术负责人",
            "developer": "PPT COM开发工程师",
            "tester": "PPT差异测试工程师",
            "optimizer": "PPT性能优化工程师",
            "security": "PPT交付审计工程师"
        }
        return name_map.get(agent_name, agent_name)


# ============================================================
# 7. Orchestrator - 主控类
# ============================================================

class Orchestrator:
    """协调所有模块，执行完整工作流"""

    def __init__(
        self,
        project_root: Path,
        max_budget: float = 10.0,
        max_retries: int = 3,
        verbose: bool = False,
        interactive_architect: bool = True,
        max_rounds: int = 1
    ):
        self.project_root = project_root
        self.task_parser = TaskParser(project_root)
        self.scheduler = AgentScheduler()
        self.executor = AgentExecutor(project_root, max_budget)
        self.state_manager = StateManager(project_root)
        self.error_handler = ErrorHandler(project_root, max_retries)
        self.monitor = ProgressMonitor(verbose)
        self.interactive_architect = interactive_architect
        self.max_rounds = max_rounds
        self.progress_file: Optional[Path] = None

    def _init_progress_file(self) -> Path:
        """
        初始化 claude-progress.md（固定文件名，每次运行重新创建）
        - 删除旧的 claude-progress.md 和历史编号文件
        - 创建全新的 claude-progress.md
        """
        base = self.project_root / "claude-progress.md"

        # 清理旧文件：固定名称 + 历史编号文件（claude-progress01.md 等）
        if base.exists():
            base.unlink()
        for old in self.project_root.glob("claude-progress[0-9][0-9].md"):
            old.unlink()

        base.write_text("# 任务进度记录\n\n", encoding='utf-8')
        return base

    def _cleanup_temp_agent_files(self) -> None:
        """清理 agent 生成的临时 md 文件（保留 claude-progress 和 PLAN.md）"""
        temp_files = [
            "CODEBASE_ANALYSIS.md",
            "SECURITY_AUDIT.md",
            "PROGRESS.md",
        ]
        # 清理PPT流水线归档文件
        for pattern in ["fix-ppt-round*.md", "diff_result-round*.json"]:
            for f in self.project_root.glob(pattern):
                try:
                    f.unlink()
                except (OSError, PermissionError):
                    pass

        for fname in temp_files:
            f = self.project_root / fname
            if f.exists():
                try:
                    f.unlink()
                except (OSError, PermissionError):
                    pass

        print("🧹 已清理临时文件")

    def _cleanup_old_state(self) -> None:
        """清理旧的状态文件和错误日志"""
        files_to_clean = [
            self.state_manager.state_file,
            self.state_manager.state_file.with_suffix('.tmp'),
            self.error_handler.error_log_file
        ]

        for file in files_to_clean:
            if file.exists():
                try:
                    file.unlink()
                except (OSError, PermissionError):
                    pass  # 忽略清理失败

        # 清理旧的临时提示文件
        claude_dir = self.project_root / ".claude"
        if claude_dir.exists():
            for temp_file in claude_dir.glob("prompt_*.txt"):
                try:
                    temp_file.unlink()
                except (OSError, PermissionError):
                    pass

    def _get_next_branch_number(self) -> int:
        """
        获取下一个分支流水号（带文件锁，防止并发竞态）

        Returns:
            3位流水号（从001开始）
        """
        counter_file = self.project_root / ".claude" / "branch_counter.txt"
        counter_file.parent.mkdir(parents=True, exist_ok=True)

        try:
            # 使用 a+ 模式：文件不存在时自动创建，避免竞态窗口
            with open(counter_file, 'a+', encoding='utf-8') as f:
                # 先移动到文件开头再加锁（a+ 模式打开后指针在 EOF）
                f.seek(0)

                # Windows 文件锁
                if sys.platform == 'win32':
                    import msvcrt
                    msvcrt.locking(f.fileno(), msvcrt.LK_LOCK, 1)

                try:
                    content = f.read().strip()
                    counter = int(content) if content else 0
                    counter += 1

                    f.seek(0)
                    f.truncate()
                    f.write(str(counter))
                    f.flush()

                    return counter
                finally:
                    # 释放锁
                    if sys.platform == 'win32':
                        f.seek(0)
                        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)

        except Exception:
            # 降级方案：毫秒时间戳 + 随机数，降低冲突概率
            import random
            return int(time.time() * 1000) % 100000 + random.randint(0, 99)

    def _create_feature_branch(self, task_description: str, first_agent: str = "arch") -> Optional[str]:
        """
        为任务创建 feature 分支

        Args:
            task_description: 任务描述（仅用于日志）
            first_agent: 首个执行的 agent 名称

        Returns:
            分支名称，如果失败则返回 None
        """
        # Agent 简写映射
        agent_abbrev = {
            "architect": "arch",
            "tech_lead": "tech",
            "developer": "dev",
            "tester": "test",
            "optimizer": "opti",
            "security": "sec",
        }

        # 获取 agent 简写
        abbrev = agent_abbrev.get(first_agent, first_agent[:4])

        try:
            # 检查是否在 git 仓库中
            result = subprocess.run(
                ["git", "rev-parse", "--git-dir"],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )
            if result.returncode != 0:
                return None  # 不是 git 仓库，跳过分支创建

            # 尝试创建分支，如果已存在则递增编号重试（最多尝试 10 次）
            for _ in range(10):
                branch_num = self._get_next_branch_number()
                branch_name = f"feature/{abbrev}-{branch_num:03d}"

                # 检查分支是否已存在
                check_result = subprocess.run(
                    ["git", "rev-parse", "--verify", branch_name],
                    cwd=str(self.project_root),
                    capture_output=True,
                    text=True,
                    encoding='utf-8'
                )

                if check_result.returncode != 0:
                    # 分支不存在，可以创建
                    result = subprocess.run(
                        ["git", "checkout", "-b", branch_name],
                        cwd=str(self.project_root),
                        capture_output=True,
                        text=True,
                        encoding='utf-8'
                    )

                    if result.returncode == 0:
                        print(f"🌿 已创建并切换到分支: {branch_name}")
                        return branch_name
                    else:
                        print(f"⚠️ 创建分支失败: {result.stderr}")
                        return None
                # 分支已存在，继续循环尝试下一个编号

            print(f"⚠️ 无法创建分支：尝试了多个编号都已存在")
            return None

        except Exception as e:
            print(f"⚠️ Git 操作失败: {e}")
            return None

    def _get_current_branch(self) -> Optional[str]:
        """获取当前 git 分支名"""
        try:
            result = subprocess.run(
                ["git", "rev-parse", "--abbrev-ref", "HEAD"],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )
            if result.returncode == 0:
                return result.stdout.strip()
            return None
        except Exception:
            return None

    def _create_agent_subbranch(self, parent_branch: str, agent_name: str) -> Optional[str]:
        """
        为特定 agent 创建隔离子分支

        Returns:
            子分支名（成功）或 None（失败）
        """
        try:
            subbranch_name = f"{parent_branch}-{agent_name}-{str(uuid.uuid4())[:8]}"

            result = subprocess.run(
                ["git", "checkout", "-b", subbranch_name],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )

            if result.returncode == 0:
                return subbranch_name
            else:
                print(f"⚠️ 创建子分支失败: {result.stderr}")
                return None

        except Exception as e:
            print(f"⚠️ 创建子分支异常: {e}")
            return None

    def _switch_to_branch(self, branch_name: str) -> bool:
        """切换到指定分支"""
        try:
            result = subprocess.run(
                ["git", "checkout", branch_name],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )
            return result.returncode == 0
        except Exception:
            return False

    def _commit_agent_changes(self, agent_name: str, task_desc: str) -> bool:
        """提交 agent 的所有更改"""
        try:
            subprocess.run(
                ["git", "add", "-A"],
                cwd=str(self.project_root),
                capture_output=True,
                timeout=30
            )

            commit_msg = f"[{agent_name}] {task_desc[:50]}"
            result = subprocess.run(
                ["git", "commit", "-m", commit_msg, "--allow-empty"],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )

            return result.returncode == 0

        except Exception:
            return False

    def _merge_subbranch(self, subbranch: str, target_branch: str) -> Tuple[bool, Optional[str]]:
        """
        将子分支合并到目标分支

        Returns:
            (成功, 冲突信息)
        """
        try:
            if not self._switch_to_branch(target_branch):
                return False, "无法切换到目标分支"

            result = subprocess.run(
                ["git", "merge", subbranch, "--no-edit"],
                cwd=str(self.project_root),
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30
            )

            if result.returncode == 0:
                return True, None
            else:
                if "CONFLICT" in result.stdout or "conflict" in result.stderr.lower():
                    subprocess.run(
                        ["git", "merge", "--abort"],
                        cwd=str(self.project_root),
                        capture_output=True,
                        timeout=30
                    )
                    return False, f"合并冲突: {result.stdout}"
                return False, result.stderr

        except Exception as e:
            return False, str(e)

    def _cleanup_subbranch(self, branch_name: str) -> None:
        """删除子分支（合并成功后）"""
        try:
            subprocess.run(
                ["git", "branch", "-d", branch_name],
                cwd=str(self.project_root),
                capture_output=True,
                timeout=30
            )
        except Exception:
            pass

    def _validate_architect_output(self) -> Tuple[bool, List[str]]:
        """
        校验 architect 执行后是否越权修改了非 .md 文件

        Returns:
            (is_clean, violated_files)
        """
        try:
            result = subprocess.run(
                ["git", "diff", "--name-only"],
                capture_output=True, text=True,
                cwd=str(self.project_root), encoding='utf-8'
            )
            changed_files = [f.strip() for f in result.stdout.strip().split('\n') if f.strip()]

            # 也检查未跟踪的新文件
            result2 = subprocess.run(
                ["git", "ls-files", "--others", "--exclude-standard"],
                capture_output=True, text=True,
                cwd=str(self.project_root), encoding='utf-8'
            )
            new_files = [f.strip() for f in result2.stdout.strip().split('\n') if f.strip()]

            all_files = changed_files + new_files
            violated = [f for f in all_files if not f.lower().endswith('.md')]

            return (len(violated) == 0, violated)
        except Exception:
            return (True, [])

    def _rollback_architect_violations(self, violated_files: List[str]) -> None:
        """回滚 architect 越权修改的文件"""
        for f in violated_files:
            file_path = self.project_root / f
            try:
                # 尝试还原已跟踪的文件
                subprocess.run(
                    ["git", "checkout", "--", f],
                    capture_output=True,
                    cwd=str(self.project_root)
                )
            except Exception:
                pass

            # 删除新创建的非 .md 文件（未跟踪的）
            if file_path.exists():
                try:
                    check = subprocess.run(
                        ["git", "ls-files", "--error-unmatch", f],
                        capture_output=True,
                        cwd=str(self.project_root)
                    )
                    if check.returncode != 0:
                        file_path.unlink()
                except Exception:
                    pass

        print(f"⚠️ Architect 越权修改了以下文件，已回滚:")
        for f in violated_files:
            print(f"   - {f}")

    async def execute(
        self,
        user_request: str,
        clean_start: bool = True,
        override_complexity: Optional[TaskComplexity] = None
    ) -> bool:
        """
        执行完整工作流

        Args:
            user_request: 用户需求描述
            clean_start: 是否清理旧状态（默认True，--resume时为False）
            override_complexity: 手动指定复杂度（可选，优先于自动解析）

        Returns:
            True if successful, False if failed
        """
        start_time = time.time()

        # Phase 0: 清理旧状态（新任务时）
        if clean_start:
            self._cleanup_old_state()
            print("🧹 已清理旧的状态文件和错误日志\n", flush=True)

        # 初始化进度文件
        self.progress_file = self._init_progress_file()
        print(f"📝 进度文件: {self.progress_file.name}", flush=True)

        # Phase 0.2: 解析任务
        print(f"📋 用户需求: {user_request}", flush=True)

        # 使用覆盖的复杂度，或自动解析
        if override_complexity:
            complexity = override_complexity
            task_prompt = user_request
            print(f"任务复杂度: {complexity.value}（用户指定）", flush=True)
        else:
            task_prompt, complexity = self.task_parser.parse(user_request)
            print(f"任务复杂度: {complexity.value}（自动解析）", flush=True)

        # Phase 0.5: 规划执行阶段
        phases = self.scheduler.plan_execution(complexity)
        print(f"执行计划: {len(phases)} 个阶段\n", flush=True)

        # Phase 0.1: 创建 feature 分支（新任务时，需要先知道首个 agent）
        feature_branch = None
        if clean_start and phases:
            first_agent = phases[0][0] if phases[0] else "arch"
            feature_branch = self._create_feature_branch(user_request, first_agent)

        # 初始化或恢复状态
        existing_state = None
        completed_agents = set()
        if not clean_start:
            existing_state = self.state_manager.load_state()
            if existing_state:
                # 获取已完成的 agents
                completed_agents = {
                    agent for agent, status in existing_state.get("agents_status", {}).items()
                    if status == "completed"
                }
                if completed_agents:
                    print(f"📂 跳过已完成的 agents: {', '.join(completed_agents)}", flush=True)

        if existing_state:
            # 恢复状态
            state = existing_state
            all_results = {}
            # 恢复已有结果用于统计
            for agent_name, result_dict in state.get("results", {}).items():
                if result_dict.get("status") == "completed":
                    all_results[agent_name] = ExecutionResult(
                        agent_name=result_dict.get("agent_name", agent_name),
                        status=AgentStatus.COMPLETED,
                        session_id=result_dict.get("session_id", ""),
                        exit_code=result_dict.get("exit_code", 0),
                        duration=result_dict.get("duration", 0),
                        cost=result_dict.get("cost", 0),
                        tokens=result_dict.get("tokens", 0),
                        output_files=result_dict.get("output_files", []),
                        error_message=result_dict.get("error_message")
                    )
        else:
            # 新任务，创建全新状态
            task_id = str(uuid.uuid4())
            state = {
                "task_id": task_id,
                "user_request": user_request,
                "complexity": complexity.value,
                "current_phase": 0,
                "agents_status": {},
                "results": {},
                "total_cost": 0.0,
                "total_tokens": 0
            }
            all_results = {}

        # 执行各阶段
        for phase_idx, agent_names in enumerate(phases, 1):
            # 过滤掉已完成的 agents
            remaining_agents = [name for name in agent_names if name not in completed_agents]
            if not remaining_agents:
                print(f"\n⏭️  Phase {phase_idx}: 所有 agents 已完成，跳过", flush=True)
                continue

            self.monitor.display_phase_start(phase_idx, remaining_agents)

            # 准备agent配置和提示词（只准备未完成的）
            configs = [self.scheduler.get_agent_config(name) for name in remaining_agents]
            progress_name = self.progress_file.name if self.progress_file else None
            prompts = {
                name: self.task_parser.generate_initial_prompt(user_request, agent_name=name, progress_file=progress_name)
                for name in remaining_agents
            }

            # 串行 or 并行执行
            if len(agent_names) == 1:
                # 单个agent：串行执行（带重试）
                config = configs[0]

                # 生成 session_id
                session_id = str(uuid.uuid4())

                # architect 可选择使用交互式模式
                if config.name == "architect" and self.interactive_architect:
                    print(f"\n💡 {self.monitor._get_agent_display_name(config.name)} 将在交互式模式下运行")
                    print(f"   你可以反复讨论计划，直到满意后退出会话")
                    print(f"   如需跳过交互，下次运行时添加 --auto-architect 参数\n")

                    # 交互式模式（阻塞，在异步上下文中运行同步函数）
                    result = await asyncio.to_thread(
                        self.executor.run_agent_interactive,
                        config,
                        prompts[config.name],
                        session_id
                    )
                else:
                    # 其他agents：无头模式（带重试）
                    self.monitor.display_agent_start(config.name, session_id)

                    result = await self.error_handler.retry_with_backoff(
                        self.executor.run_agent,
                        config,
                        prompts[config.name],
                        session_id=session_id
                    )

                self.monitor.display_agent_complete(result)
                all_results[config.name] = result

                # Architect 后置校验：检查是否越权修改了非 .md 文件
                if config.name == "architect" and result.status == AgentStatus.COMPLETED:
                    is_clean, violated = self._validate_architect_output()
                    if not is_clean:
                        self._rollback_architect_violations(violated)

                # 如果失败，终止执行
                if result.status == AgentStatus.FAILED:
                    print(f"\n❌ {config.name} 执行失败，终止流程")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            else:
                # 多个agents：并行执行 - 使用子分支隔离
                # 为每个agent生成session_id
                session_ids = {config.name: str(uuid.uuid4()) for config in configs}

                # 记录主分支
                main_branch = feature_branch if feature_branch else self._get_current_branch()
                agent_subbranches = {}

                for config in configs:
                    self.monitor.display_agent_start(config.name, session_ids[config.name])

                # 并行执行（每个 agent 在独立子分支）
                async def run_agent_isolated(config: AgentConfig, prompt: str, session_id: str) -> ExecutionResult:
                    # 创建子分支
                    subbranch = self._create_agent_subbranch(main_branch, config.name)
                    if subbranch:
                        agent_subbranches[config.name] = subbranch

                    # 执行 agent
                    result = await self.error_handler.retry_with_backoff(
                        self.executor.run_agent,
                        config,
                        prompt,
                        session_id=session_id
                    )

                    # 提交更改
                    if subbranch and result.status == AgentStatus.COMPLETED:
                        self._commit_agent_changes(config.name, user_request[:50])

                    return result

                tasks = [
                    run_agent_isolated(
                        config,
                        prompts[config.name],
                        session_ids[config.name]
                    )
                    for config in configs
                ]
                results = await asyncio.gather(*tasks)

                # 显示结果
                for result in results:
                    self.monitor.display_agent_complete(result)
                    all_results[result.agent_name] = result

                # 合并子分支
                if all(r.status == AgentStatus.COMPLETED for r in results) and agent_subbranches:
                    print(f"\n🔀 合并各 agent 的更改...")
                    merge_failures = []

                    for agent_name, subbranch in agent_subbranches.items():
                        success, conflict_info = self._merge_subbranch(subbranch, main_branch)
                        if success:
                            print(f"  ✅ {agent_name} 的更改已合并")
                            self._cleanup_subbranch(subbranch)
                        else:
                            merge_failures.append((agent_name, conflict_info))
                            print(f"  ❌ {agent_name} 合并失败: {conflict_info}")

                    if merge_failures:
                        print(f"\n⚠️ 检测到合并冲突！")
                        print(f"   以下分支保留供手动处理:")
                        for agent, _ in merge_failures:
                            print(f"     - {agent_subbranches.get(agent, 'unknown')}")
                elif agent_subbranches:
                    # 有 agent 失败，清理子分支
                    for subbranch in agent_subbranches.values():
                        self._cleanup_subbranch(subbranch)

                # 如果任何一个失败，终止执行
                if any(r.status == AgentStatus.FAILED for r in results):
                    failed_agents = [r.agent_name for r in results if r.status == AgentStatus.FAILED]
                    print(f"\n❌ 以下agents执行失败: {', '.join(failed_agents)}，终止流程")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            # 更新状态
            state["current_phase"] = phase_idx
            for name, result in all_results.items():
                state["agents_status"][name] = result.status.value
                # 转换 ExecutionResult 为可序列化的字典
                result_dict = asdict(result)
                result_dict["status"] = result.status.value  # 枚举 -> 字符串
                state["results"][name] = result_dict
            self.state_manager.save_state(state)

        # 显示汇总
        total_duration = time.time() - start_time
        self.monitor.display_summary(all_results, total_duration)

        # 保存最终状态
        self._save_final_state(state, all_results, total_duration)

        # 如果创建了 feature 分支，提示合并
        if feature_branch:
            print(f"\n{'='*60}")
            print(f"✅ 任务完成！当前在分支: {feature_branch}")
            print(f"{'='*60}")
            print(f"下一步操作：")
            print(f"  1. 检查生成的代码和文档")
            print(f"  2. 运行测试确保功能正常")
            print(f"  3. 提交更改：")
            print(f"     git add .")
            print(f"     git commit -m \"完成：{user_request[:50]}\"")
            print(f"  4. 合并到主分支：")
            print(f"     git checkout main")
            print(f"     git merge {feature_branch}")
            print(f"  5. 或创建 Pull Request 进行代码审查")
            print(f"{'='*60}\n")

        # 提示进度文件位置 & 清理临时文件
        if self.progress_file:
            print(f"📝 本次进度记录: {self.progress_file.name}")
        self._cleanup_temp_agent_files()

        return True

    def _save_final_state(
        self,
        state: Dict,
        all_results: Dict[str, ExecutionResult],
        total_duration: float
    ) -> None:
        """保存最终状态"""
        state["total_cost"] = sum(r.cost for r in all_results.values())
        state["total_tokens"] = sum(r.tokens for r in all_results.values())
        state["total_duration"] = total_duration
        self.state_manager.save_state(state)

    async def execute_from_plan(self, plan_content: str, existing_state: Optional[Dict] = None) -> bool:
        """
        从 PLAN.md 开始执行（跳过 architect 阶段）

        用于情景2：半自动模式，architect 已在 claude CLI 中完成
        也用于恢复中断的任务

        Args:
            plan_content: PLAN.md 的内容
            existing_state: 现有状态（用于恢复时跳过已完成的 agent）

        Returns:
            True if successful, False if failed
        """
        start_time = time.time()

        # 初始化进度文件
        self.progress_file = self._init_progress_file()
        print(f"📝 进度文件: {self.progress_file.name}", flush=True)

        # Bug Fix: 创建 feature 分支（从 plan 开始也需要分支隔离）
        feature_branch = None
        if not existing_state:  # 新任务才创建分支，恢复任务不创建
            feature_branch = self._create_feature_branch("from-plan", "tech")

        # 所有可能的 agents（跳过 architect）
        all_agents = ["tech_lead", "developer", "tester", "optimizer", "security"]

        # 如果有现有状态，过滤掉已完成的 agents
        if existing_state and existing_state.get("agents_status"):
            completed_agents = [
                agent for agent, status in existing_state["agents_status"].items()
                if status == "completed"
            ]
            remaining_agents = [a for a in all_agents if a not in completed_agents]
            print(f"📂 已完成的 agents: {', '.join(completed_agents) if completed_agents else '无'}")
            print(f"🔄 待执行的 agents: {', '.join(remaining_agents) if remaining_agents else '无'}")
        else:
            remaining_agents = all_agents

        if not remaining_agents:
            print("✅ 所有 agents 已完成，无需继续执行")
            return True

        # 构建提示词（引用 PLAN.md 而非嵌入全文，避免 Windows 命令行长度限制）
        progress_info = ""
        if self.progress_file:
            progress_info = f"\n📝 完成任务后，请将你的工作记录追加到进度文件: `{self.progress_file.name}`（先 Read 保留已有内容，再 Write 追加你的部分）\n"
        task_prompt = f"""请使用 Read 工具读取项目根目录的 `PLAN.md` 文件，然后根据实施计划执行你的职责。

请严格按照计划执行，确保与其他 agents 的工作保持一致。
{progress_info}"""

        # 初始化或恢复状态
        if existing_state:
            state = existing_state
            all_results = {}
            # 恢复已有结果
            for agent_name, result_dict in state.get("results", {}).items():
                if result_dict.get("status") == "completed":
                    # 重建 ExecutionResult 对象用于统计
                    all_results[agent_name] = ExecutionResult(
                        agent_name=result_dict.get("agent_name", agent_name),
                        status=AgentStatus.COMPLETED,
                        session_id=result_dict.get("session_id", ""),
                        exit_code=result_dict.get("exit_code", 0),
                        duration=result_dict.get("duration", 0),
                        cost=result_dict.get("cost", 0),
                        tokens=result_dict.get("tokens", 0),
                        output_files=result_dict.get("output_files", []),
                        error_message=result_dict.get("error_message")
                    )
        else:
            task_id = str(uuid.uuid4())
            state = {
                "task_id": task_id,
                "user_request": "从 PLAN.md 执行",
                "complexity": "from_plan",
                "current_phase": 1,  # 从 phase 1 开始（跳过 phase 0 architect）
                "agents_status": {"architect": "completed"},
                "results": {},
                "total_cost": 0.0,
                "total_tokens": 0
            }
            all_results = {}

        # 计算起始 phase 索引（architect 已跳过，从 phase 1 开始）
        start_phase_idx = len(all_agents) - len(remaining_agents) + 1

        # 执行剩余 agents
        for i, agent_name in enumerate(remaining_agents):
            phase_idx = start_phase_idx + i
            self.monitor.display_phase_start(phase_idx, [agent_name])

            config = self.scheduler.get_agent_config(agent_name)
            session_id = str(uuid.uuid4())

            self.monitor.display_agent_start(config.name, session_id)

            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent,
                config,
                task_prompt,
                session_id=session_id
            )

            self.monitor.display_agent_complete(result)
            all_results[config.name] = result

            # 更新状态
            state["current_phase"] = phase_idx
            state["agents_status"][config.name] = result.status.value
            # 转换 ExecutionResult 为可序列化的字典
            result_dict = asdict(result)
            result_dict["status"] = result.status.value
            state["results"][config.name] = result_dict
            self.state_manager.save_state(state)

            # 如果失败，终止执行
            if result.status == AgentStatus.FAILED:
                print(f"\n❌ {config.name} 执行失败，已保存状态")
                print(f"   修复问题后，运行 python mc-dir.py --resume 继续")
                self._save_final_state(state, all_results, time.time() - start_time)
                return False

        # 成功完成
        total_duration = time.time() - start_time
        self._save_final_state(state, all_results, total_duration)
        self.monitor.display_summary(all_results, total_duration)

        # 提示进度文件位置 & 清理临时文件
        if self.progress_file:
            print(f"📝 本次进度记录: {self.progress_file.name}")
        self._cleanup_temp_agent_files()

        return True

    async def execute_from_plan_with_loop(
        self,
        plan_content: str,
        existing_state: Optional[Dict] = None
    ) -> bool:
        """
        从 PLAN.md 开始执行，带多轮 developer-tester 循环

        跳过 architect（已完成），执行:
        tech_lead → developer ⇄ tester（循环）→ optimizer → security

        Args:
            plan_content: PLAN.md 的内容
            existing_state: 现有状态（用于恢复时跳过已完成的 agent）

        Returns:
            True if successful, False if failed
        """
        start_time = time.time()

        # 初始化进度文件
        self.progress_file = self._init_progress_file()
        print(f"📝 进度文件: {self.progress_file.name}", flush=True)

        # Bug Fix: 创建 feature 分支（从 plan 开始也需要分支隔离）
        feature_branch = None
        if not existing_state:  # 新任务才创建分支，恢复任务不创建
            feature_branch = self._create_feature_branch("from-plan-loop", "tech")

        # 构建提示词（引用 PLAN.md 而非嵌入全文，避免 Windows 命令行长度限制）
        progress_info = ""
        if self.progress_file:
            progress_info = f"\n📝 完成任务后，请将你的工作记录追加到进度文件: `{self.progress_file.name}`（先 Read 保留已有内容，再 Write 追加你的部分）\n"
        task_prompt = f"""请使用 Read 工具读取项目根目录的 `PLAN.md` 文件，然后根据实施计划执行你的职责。

请严格按照计划执行，确保与其他 agents 的工作保持一致。
{progress_info}"""

        # 初始化或恢复状态
        if existing_state:
            state = existing_state
            all_results = {}
            current_round = state.get("current_round", 1)
        else:
            task_id = str(uuid.uuid4())
            state = {
                "task_id": task_id,
                "user_request": "从 PLAN.md 执行（多轮模式）",
                "complexity": "from_plan_loop",
                "current_phase": 1,
                "current_round": 1,
                "agents_status": {"architect": "completed"},
                "results": {},
                "total_cost": 0.0,
                "total_tokens": 0
            }
            all_results = {}
            current_round = 1

        # Phase 1: 执行 tech_lead（只执行一次）
        if state.get("agents_status", {}).get("tech_lead") != "completed":
            print(f"\n{'='*60}")
            print(f"🔄 Phase 1: 技术审核")
            print(f"{'='*60}\n")

            config = self.scheduler.get_agent_config("tech_lead")
            session_id = str(uuid.uuid4())

            self.monitor.display_agent_start(config.name, session_id)

            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent,
                config,
                task_prompt,
                session_id=session_id
            )

            self.monitor.display_agent_complete(result)
            all_results["tech_lead"] = result

            state["agents_status"]["tech_lead"] = result.status.value
            result_dict = asdict(result)
            result_dict["status"] = result.status.value
            state["results"]["tech_lead"] = result_dict
            self.state_manager.save_state(state)

            if result.status == AgentStatus.FAILED:
                print(f"\n❌ tech_lead 执行失败")
                self._save_final_state(state, all_results, time.time() - start_time)
                return False

        # Phase 2: developer-tester 循环
        while current_round <= self.max_rounds:
            print(f"\n{'='*60}")
            print(f"🔄 Round {current_round}/{self.max_rounds}: 开发和测试")
            print(f"{'='*60}\n")

            # 准备本轮的任务提示
            round_prompt = task_prompt
            if current_round > 1:
                has_bugs, bug_summaries = self._check_bug_report()
                if bug_summaries:
                    bug_info = "\n".join(f"  - {b}" for b in bug_summaries[:10])
                    # 读取 fix-ppt.md 的修复建议
                    fix_advice = ""
                    fix_file = self.project_root / "fix-ppt.md"
                    if fix_file.exists():
                        try:
                            fix_advice = fix_file.read_text(encoding='utf-8')[:2000]
                        except (IOError, OSError):
                            pass
                    round_prompt = f"""{task_prompt}

---

⚠️ 上一轮PPT差异测试未通过（codex {current_round - 1}.0），请修复后生成 codex {current_round}.0：

{bug_info}

{('修复建议（来自fix-ppt.md）：' + chr(10) + fix_advice) if fix_advice else '请阅读 fix-ppt.md 获取详细修复建议。'}

修复优先级：1.检查shape策略 → 2.调整prompt → 3.改提取函数
"""

            # 执行 developer
            dev_key = f"developer_round{current_round}"
            if state.get("agents_status", {}).get(dev_key) != "completed":
                config = self.scheduler.get_agent_config("developer")
                session_id = str(uuid.uuid4())

                self.monitor.display_agent_start(f"developer (round {current_round})", session_id)

                result = await self.error_handler.retry_with_backoff(
                    self.executor.run_agent,
                    config,
                    round_prompt,
                    session_id=session_id
                )

                self.monitor.display_agent_complete(result)
                all_results[dev_key] = result

                state["agents_status"][dev_key] = result.status.value
                result_dict = asdict(result)
                result_dict["status"] = result.status.value
                state["results"][dev_key] = result_dict
                self.state_manager.save_state(state)

                if result.status == AgentStatus.FAILED:
                    print(f"\n❌ developer (round {current_round}) 执行失败")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            # 执行 tester
            tester_key = f"tester_round{current_round}"
            if state.get("agents_status", {}).get(tester_key) != "completed":
                config = self.scheduler.get_agent_config("tester")
                session_id = str(uuid.uuid4())

                self.monitor.display_agent_start(f"tester (round {current_round})", session_id)

                result = await self.error_handler.retry_with_backoff(
                    self.executor.run_agent,
                    config,
                    round_prompt,
                    session_id=session_id
                )

                self.monitor.display_agent_complete(result)
                all_results[tester_key] = result

                state["agents_status"][tester_key] = result.status.value
                result_dict = asdict(result)
                result_dict["status"] = result.status.value
                state["results"][tester_key] = result_dict
                self.state_manager.save_state(state)

                if result.status == AgentStatus.FAILED:
                    print(f"\n❌ tester (round {current_round}) 执行失败")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            # 检查PPT差异测试结果
            has_bugs, bug_summaries = self._check_bug_report()

            if not has_bugs:
                # 保底检测：如果是第1轮且没有 diff_result.json，tester 可能没正确生成
                diff_file = self.project_root / "diff_result.json"
                if current_round == 1 and not diff_file.exists() and self.max_rounds > 1:
                    print(f"\n⚠️ Round {current_round}: diff_result.json 不存在")
                    print(f"   Tester 可能没有生成测试报告，将继续下一轮确认...")
                    current_round += 1
                    state["current_round"] = current_round
                    self.state_manager.save_state(state)
                    continue  # 继续下一轮循环

                print(f"\n✅ Round {current_round}: PPT差异测试全部通过，继续执行后续阶段")
                break

            if current_round < self.max_rounds:
                print(f"\n⚠️ Round {current_round}: PPT差异测试有 {len(bug_summaries)} 项未通过")
                print(f"   将进入 Round {current_round + 1} 进行修复...")
                self._archive_bug_report(current_round)
            else:
                print(f"\n⚠️ 已达到最大循环次数 ({self.max_rounds})")
                print(f"   仍有 {len(bug_summaries)} 项未通过，请手动检查 fix-ppt.md 和 diff_result.json")

            current_round += 1
            state["current_round"] = current_round
            self.state_manager.save_state(state)

        # Phase 3: 执行 optimizer 和 security（只执行一次）
        phase3_agents = ["optimizer", "security"]
        print(f"\n{'='*60}")
        print(f"🔄 Phase 3: 优化和安全检查")
        print(f"{'='*60}\n")

        for agent_name in phase3_agents:
            if state.get("agents_status", {}).get(agent_name) == "completed":
                print(f"⏭️ 跳过已完成: {agent_name}")
                continue

            config = self.scheduler.get_agent_config(agent_name)
            session_id = str(uuid.uuid4())

            self.monitor.display_agent_start(config.name, session_id)

            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent,
                config,
                task_prompt,
                session_id=session_id
            )

            self.monitor.display_agent_complete(result)
            all_results[config.name] = result

            state["agents_status"][config.name] = result.status.value
            result_dict = asdict(result)
            result_dict["status"] = result.status.value
            state["results"][config.name] = result_dict
            self.state_manager.save_state(state)

            if result.status == AgentStatus.FAILED:
                print(f"\n❌ {config.name} 执行失败")
                self._save_final_state(state, all_results, time.time() - start_time)
                return False

        # 完成
        total_duration = time.time() - start_time
        self._save_final_state(state, all_results, total_duration)
        self.monitor.display_summary(all_results, total_duration)

        print(f"\n   执行了 {current_round} 轮 developer-tester 循环")

        # 提示进度文件位置 & 清理临时文件
        if self.progress_file:
            print(f"📝 本次进度记录: {self.progress_file.name}")
        self._cleanup_temp_agent_files()

        return True

    def _check_bug_report(self) -> Tuple[bool, List[str]]:
        """
        检查PPT差异测试结果（读取 diff_result.json）

        通过标准：Visual >= 98, Readability >= 95, Semantic == 100
        若 diff_result.json 不存在，视为测试未执行，返回 has_failures=True

        Returns:
            (has_failures, failure_summaries): 是否有未通过项，以及失败摘要列表
        """
        diff_file = self.project_root / "diff_result.json"

        if not diff_file.exists():
            print(f"   📋 diff_result.json 不存在，视为测试未执行")
            return True, ["diff_result.json 不存在，测试未执行"]

        try:
            content = diff_file.read_text(encoding='utf-8')
            diff_data = json.loads(content)
        except (IOError, OSError, json.JSONDecodeError) as e:
            print(f"   ⚠️ 无法读取/解析 diff_result.json: {e}")
            return True, [f"diff_result.json 读取失败: {e}"]

        # 提取三层评分
        visual = diff_data.get("visual_score", 0)
        readability = diff_data.get("readability_score", 0)
        semantic = diff_data.get("semantic_coverage", 0)
        overall_pass = diff_data.get("overall_pass", False)

        failure_summaries = []

        # 检查三层阈值
        if visual < 98:
            failure_summaries.append(f"Visual Score {visual:.1f} < 98（差距 {98-visual:.1f}）")
        if readability < 95:
            failure_summaries.append(f"Readability Score {readability:.1f} < 95（差距 {95-readability:.1f}）")
        if semantic < 100:
            failure_summaries.append(f"Semantic Coverage {semantic:.0f} < 100")

        # 检查per-shape失败项
        per_shape = diff_data.get("per_shape", [])
        failed_shapes = [s for s in per_shape if s.get("visual_score", 100) < 90 or not s.get("semantic_pass", True)]
        if failed_shapes:
            for s in failed_shapes[:5]:
                name = s.get("name", "unknown")
                issues = s.get("issues", [])
                issue_text = "; ".join(issues[:2]) if issues else "分数过低"
                failure_summaries.append(f"Shape [{name}]: {issue_text}")

        has_failures = len(failure_summaries) > 0

        # 调试输出
        if has_failures:
            print(f"   ❌ PPT差异测试未通过（Visual={visual:.1f}, Read={readability:.1f}, Sem={semantic:.0f}）:")
            for i, summary in enumerate(failure_summaries[:5], 1):
                print(f"      {i}. {summary[:80]}{'...' if len(summary) > 80 else ''}")
            if len(failure_summaries) > 5:
                print(f"      ... 还有 {len(failure_summaries) - 5} 项")
        else:
            print(f"   ✅ PPT差异测试通过（Visual={visual:.1f}, Read={readability:.1f}, Sem={semantic:.0f}）")

        return has_failures, failure_summaries

    def _archive_bug_report(self, round_num: int) -> None:
        """归档当前轮次的 fix-ppt.md 和 diff_result.json"""
        for filename in ["fix-ppt.md", "diff_result.json"]:
            src_file = self.project_root / filename
            if src_file.exists():
                stem = Path(filename).stem
                suffix = Path(filename).suffix
                archive_file = self.project_root / f"{stem}-round{round_num}{suffix}"
                try:
                    import shutil
                    shutil.copy2(src_file, archive_file)
                    print(f"📁 已归档 {filename} -> {archive_file.name}")
                except (IOError, OSError) as e:
                    print(f"⚠️ 归档 {filename} 失败: {e}")

    async def execute_with_loop(
        self,
        user_request: str,
        clean_start: bool = True,
        existing_state: Optional[Dict] = None,
        override_complexity: Optional[TaskComplexity] = None
    ) -> bool:
        """
        带多轮循环的执行模式

        developer-tester 会循环执行，直到：
        1. 没有未解决的 bug
        2. 达到最大循环次数 (max_rounds)

        Args:
            user_request: 用户请求
            clean_start: 是否清理旧状态
            existing_state: 现有状态（恢复时使用）
            override_complexity: 手动指定复杂度（可选，优先于自动解析）

        Returns:
            True if successful, False if failed
        """
        start_time = time.time()

        # 清理旧状态
        if clean_start:
            self._cleanup_old_state()
            print("🧹 已清理旧的状态文件\n")

        # 初始化进度文件
        self.progress_file = self._init_progress_file()
        print(f"📝 进度文件: {self.progress_file.name}", flush=True)

        # 解析任务复杂度
        if override_complexity:
            complexity = override_complexity
            print(f"📊 任务复杂度: {complexity.value}（用户指定）")
        else:
            _, complexity = self.task_parser.parse(user_request)
            print(f"📊 任务复杂度: {complexity.value}（自动解析）")

        # 获取执行计划
        phases = self.scheduler.plan_execution(complexity)

        # 创建 feature 分支（获取首个agent名称）
        first_agent = phases[0][0] if phases and phases[0] else "arch"
        feature_branch = self._create_feature_branch(user_request, first_agent)

        # 初始化状态
        task_id = str(uuid.uuid4())
        state = existing_state or {
            "task_id": task_id,
            "user_request": user_request,
            "complexity": complexity.value,
            "current_phase": 0,
            "current_round": 1,
            "agents_status": {},
            "results": {},
            "total_cost": 0.0,
            "total_tokens": 0
        }

        all_results = {}

        # 进度文件后缀（供各阶段 prompt 使用）
        progress_suffix = ""
        if self.progress_file:
            progress_suffix = f"\n\n📝 完成任务后，请将你的工作记录追加到进度文件: `{self.progress_file.name}`（先 Read 保留已有内容，再 Write 追加你的部分）"

        # 根据复杂度拆分执行阶段
        # phases 格式示例（COMPLEX）: [["architect"], ["tech_lead"], ["developer"], ["tester", "security", "optimizer"]]
        # phases 格式示例（MINIMAL）: [["developer"], ["tester"]]
        pre_loop_agents = []  # Phase 1: 规划阶段（architect, tech_lead）
        loop_agents = ["developer", "tester"]  # Phase 2: 开发-测试循环
        post_loop_agents = []  # Phase 3: 优化阶段（optimizer, security）

        # 从 phases 中提取各阶段的 agents
        for phase in phases:
            for agent in phase:
                if agent in ["developer", "tester"]:
                    # 这些 agent 在循环中执行，不放入 pre/post
                    continue
                elif agent in ["architect", "tech_lead"]:
                    if agent not in pre_loop_agents:
                        pre_loop_agents.append(agent)
                elif agent in ["optimizer", "security"]:
                    if agent not in post_loop_agents:
                        post_loop_agents.append(agent)

        # Phase 1: 执行 architect 和 tech_lead（只执行一次）
        phase1_agents = pre_loop_agents
        if phase1_agents:
            print(f"\n{'='*60}")
            print(f"🔄 Phase 1: 规划和设计")
            print(f"{'='*60}\n")
        else:
            print(f"\n⏭️ 跳过 Phase 1（当前复杂度无需规划阶段）\n")

        for agent_name in phase1_agents:
            if state.get("agents_status", {}).get(agent_name) == "completed":
                print(f"⏭️ 跳过已完成: {agent_name}")
                continue

            config = self.scheduler.get_agent_config(agent_name)
            session_id = str(uuid.uuid4())

            self.monitor.display_agent_start(config.name, session_id)

            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent,
                config,
                user_request + progress_suffix,
                session_id=session_id
            )

            self.monitor.display_agent_complete(result)
            all_results[config.name] = result

            # Architect 后置校验：检查是否越权修改了非 .md 文件
            if config.name == "architect" and result.status == AgentStatus.COMPLETED:
                is_clean, violated = self._validate_architect_output()
                if not is_clean:
                    self._rollback_architect_violations(violated)

            # 更新状态
            state["agents_status"][config.name] = result.status.value
            result_dict = asdict(result)
            result_dict["status"] = result.status.value
            state["results"][config.name] = result_dict
            self.state_manager.save_state(state)

            if result.status == AgentStatus.FAILED:
                print(f"\n❌ {config.name} 执行失败")
                self._save_final_state(state, all_results, time.time() - start_time)
                return False

            # architect 完成后读取 PLAN.md
            if agent_name == "architect":
                plan_file = self.project_root / "PLAN.md"
                if plan_file.exists():
                    try:
                        user_request = plan_file.read_text(encoding='utf-8', errors='replace')
                    except (IOError, OSError) as e:
                        print(f"⚠️ 无法读取 PLAN.md: {e}")
                        return False

        # Phase 2: developer-tester 循环
        current_round = state.get("current_round", 1)

        while current_round <= self.max_rounds:
            print(f"\n{'='*60}")
            print(f"🔄 Round {current_round}/{self.max_rounds}: 开发和测试")
            print(f"{'='*60}\n")

            # 准备本轮的任务提示
            round_prompt = user_request + progress_suffix
            if current_round > 1:
                # 如果是第2轮+，附加上一轮的差异测试结果
                has_bugs, bug_summaries = self._check_bug_report()
                if bug_summaries:
                    bug_info = "\n".join(f"  - {b}" for b in bug_summaries[:10])
                    # 读取 fix-ppt.md 的修复建议
                    fix_advice = ""
                    fix_file = self.project_root / "fix-ppt.md"
                    if fix_file.exists():
                        try:
                            fix_advice = fix_file.read_text(encoding='utf-8')[:2000]
                        except (IOError, OSError):
                            pass
                    round_prompt = f"""{user_request}

---

⚠️ 上一轮PPT差异测试未通过（codex {current_round - 1}.0），请修复后生成 codex {current_round}.0：

{bug_info}

{('修复建议（来自fix-ppt.md）：' + chr(10) + fix_advice) if fix_advice else '请阅读 fix-ppt.md 获取详细修复建议。'}

修复优先级：1.检查shape策略 → 2.调整prompt → 3.改提取函数
"""

            # 执行 developer
            dev_key = f"developer_round{current_round}"
            if state.get("agents_status", {}).get(dev_key) != "completed":
                config = self.scheduler.get_agent_config("developer")
                session_id = str(uuid.uuid4())

                self.monitor.display_agent_start(f"developer (round {current_round})", session_id)

                result = await self.error_handler.retry_with_backoff(
                    self.executor.run_agent,
                    config,
                    round_prompt,
                    session_id=session_id
                )

                self.monitor.display_agent_complete(result)
                all_results[dev_key] = result

                state["agents_status"][dev_key] = result.status.value
                result_dict = asdict(result)
                result_dict["status"] = result.status.value
                state["results"][dev_key] = result_dict
                self.state_manager.save_state(state)

                if result.status == AgentStatus.FAILED:
                    print(f"\n❌ developer (round {current_round}) 执行失败")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            # 执行 tester
            tester_key = f"tester_round{current_round}"
            if state.get("agents_status", {}).get(tester_key) != "completed":
                config = self.scheduler.get_agent_config("tester")
                session_id = str(uuid.uuid4())

                self.monitor.display_agent_start(f"tester (round {current_round})", session_id)

                result = await self.error_handler.retry_with_backoff(
                    self.executor.run_agent,
                    config,
                    round_prompt,
                    session_id=session_id
                )

                self.monitor.display_agent_complete(result)
                all_results[tester_key] = result

                state["agents_status"][tester_key] = result.status.value
                result_dict = asdict(result)
                result_dict["status"] = result.status.value
                state["results"][tester_key] = result_dict
                self.state_manager.save_state(state)

                if result.status == AgentStatus.FAILED:
                    print(f"\n❌ tester (round {current_round}) 执行失败")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            # 检查PPT差异测试结果
            has_bugs, bug_summaries = self._check_bug_report()

            if not has_bugs:
                # 保底检测：如果是第1轮且没有 diff_result.json，tester 可能没正确生成
                diff_file = self.project_root / "diff_result.json"
                if current_round == 1 and not diff_file.exists() and self.max_rounds > 1:
                    print(f"\n⚠️ Round {current_round}: diff_result.json 不存在")
                    print(f"   Tester 可能没有生成测试报告，将继续下一轮确认...")
                    current_round += 1
                    state["current_round"] = current_round
                    self.state_manager.save_state(state)
                    continue  # 继续下一轮循环

                print(f"\n✅ Round {current_round}: PPT差异测试全部通过，继续执行后续阶段")
                break

            if current_round < self.max_rounds:
                print(f"\n⚠️ Round {current_round}: PPT差异测试有 {len(bug_summaries)} 项未通过")
                print(f"   将进入 Round {current_round + 1} 进行修复...")
                # 归档本轮测试报告
                self._archive_bug_report(current_round)
            else:
                print(f"\n⚠️ 已达到最大循环次数 ({self.max_rounds})")
                print(f"   仍有 {len(bug_summaries)} 项未通过，请手动检查 fix-ppt.md 和 diff_result.json")

            current_round += 1
            state["current_round"] = current_round
            self.state_manager.save_state(state)

        # Phase 3: 执行 optimizer 和 security（只执行一次）
        phase3_agents = ["optimizer", "security"]
        print(f"\n{'='*60}")
        print(f"🔄 Phase 3: 优化和安全检查")
        print(f"{'='*60}\n")

        for agent_name in phase3_agents:
            if state.get("agents_status", {}).get(agent_name) == "completed":
                print(f"⏭️ 跳过已完成: {agent_name}")
                continue

            config = self.scheduler.get_agent_config(agent_name)
            session_id = str(uuid.uuid4())

            self.monitor.display_agent_start(config.name, session_id)

            result = await self.error_handler.retry_with_backoff(
                self.executor.run_agent,
                config,
                user_request + progress_suffix,
                session_id=session_id
            )

            self.monitor.display_agent_complete(result)
            all_results[config.name] = result

            state["agents_status"][config.name] = result.status.value
            result_dict = asdict(result)
            result_dict["status"] = result.status.value
            state["results"][config.name] = result_dict
            self.state_manager.save_state(state)

            if result.status == AgentStatus.FAILED:
                print(f"\n❌ {config.name} 执行失败")
                self._save_final_state(state, all_results, time.time() - start_time)
                return False

        # 完成
        total_duration = time.time() - start_time
        self._save_final_state(state, all_results, total_duration)
        self.monitor.display_summary(all_results, total_duration)

        # 打印分支信息
        if feature_branch:
            print(f"\n{'='*60}")
            print(f"✅ 任务完成！当前在分支: {feature_branch}")
            print(f"   执行了 {current_round} 轮 developer-tester 循环")
            print(f"{'='*60}\n")

        # 提示进度文件位置 & 清理临时文件
        if self.progress_file:
            print(f"📝 本次进度记录: {self.progress_file.name}")
        self._cleanup_temp_agent_files()

        return True

    async def execute_manual(
        self,
        phases: List[List[Tuple[str, str]]],
        clean_start: bool = True
    ) -> bool:
        """
        执行手动指定的 agent 任务

        Args:
            phases: [[("agent_name", "task"), ...], ...]
            clean_start: 是否清理旧状态

        Returns:
            True if successful, False if failed
        """
        start_time = time.time()

        # 清理旧状态
        if clean_start:
            self._cleanup_old_state()
            print("🧹 已清理旧的状态文件\n")

        # 初始化进度文件
        self.progress_file = self._init_progress_file()
        print(f"📝 进度文件: {self.progress_file.name}", flush=True)

        # 创建 feature 分支（使用首个 agent 名称）
        first_agent = phases[0][0][0] if phases and phases[0] else "arch"
        first_task = phases[0][0][1] if phases and phases[0] else "manual-task"
        feature_branch = self._create_feature_branch(first_task, first_agent)

        # 初始化状态
        task_id = str(uuid.uuid4())
        state = {
            "task_id": task_id,
            "mode": "manual",
            "current_phase": 0,
            "agents_status": {},
            "results": {},
            "total_cost": 0.0,
            "total_tokens": 0
        }

        all_results = {}

        # 执行各阶段
        for phase_idx, phase_tasks in enumerate(phases, 1):
            agent_names = [agent for agent, _ in phase_tasks]
            self.monitor.display_phase_start(phase_idx, agent_names)

            # 准备 agent 配置和提示词
            configs = []
            prompts = {}

            for agent_name, task in phase_tasks:
                config = self.scheduler.get_agent_config(agent_name)
                configs.append(config)
                progress_name = self.progress_file.name if self.progress_file else None
                prompts[agent_name] = self.task_parser.generate_initial_prompt(task, agent_name=agent_name, progress_file=progress_name)

            # 串行 or 并行执行
            if len(phase_tasks) == 1:
                # 单个 agent
                config = configs[0]
                agent_name = config.name
                session_id = str(uuid.uuid4())

                # architect 使用交互式模式
                if agent_name == "architect" and self.interactive_architect:
                    print(f"\n💡 {self.monitor._get_agent_display_name(agent_name)} 将在交互式模式下运行")

                    result = await asyncio.to_thread(
                        self.executor.run_agent_interactive,
                        config,
                        prompts[agent_name],
                        session_id
                    )
                else:
                    self.monitor.display_agent_start(agent_name, session_id)

                    result = await self.error_handler.retry_with_backoff(
                        self.executor.run_agent,
                        config,
                        prompts[agent_name],
                        session_id=session_id
                    )

                self.monitor.display_agent_complete(result)
                all_results[agent_name] = result

                # Architect 后置校验：检查是否越权修改了非 .md 文件
                if agent_name == "architect" and result.status == AgentStatus.COMPLETED:
                    is_clean, violated = self._validate_architect_output()
                    if not is_clean:
                        self._rollback_architect_violations(violated)

                if result.status == AgentStatus.FAILED:
                    print(f"\n❌ {agent_name} 执行失败，终止流程")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            else:
                # 多个 agent 并行执行 - 使用子分支隔离
                session_ids = {config.name: str(uuid.uuid4()) for config in configs}

                # 记录主分支
                main_branch = feature_branch if feature_branch else self._get_current_branch()
                agent_subbranches = {}

                for config in configs:
                    self.monitor.display_agent_start(config.name, session_ids[config.name])

                # 并行执行（每个 agent 在独立子分支）
                async def run_agent_isolated(config: AgentConfig, prompt: str, session_id: str, task_desc: str) -> ExecutionResult:
                    # 创建子分支
                    subbranch = self._create_agent_subbranch(main_branch, config.name)
                    if subbranch:
                        agent_subbranches[config.name] = subbranch

                    # 执行 agent
                    result = await self.error_handler.retry_with_backoff(
                        self.executor.run_agent,
                        config,
                        prompt,
                        session_id=session_id
                    )

                    # 提交更改
                    if subbranch and result.status == AgentStatus.COMPLETED:
                        self._commit_agent_changes(config.name, task_desc)

                    return result

                # 获取每个 agent 对应的任务描述
                agent_task_map = {agent: task for agent, task in phase_tasks}

                tasks = [
                    run_agent_isolated(
                        config,
                        prompts[config.name],
                        session_ids[config.name],
                        agent_task_map.get(config.name, "parallel-task")
                    )
                    for config in configs
                ]
                results = await asyncio.gather(*tasks)

                # 显示结果
                for result in results:
                    self.monitor.display_agent_complete(result)
                    all_results[result.agent_name] = result

                # 合并子分支
                if all(r.status == AgentStatus.COMPLETED for r in results) and agent_subbranches:
                    print(f"\n🔀 合并各 agent 的更改...")
                    merge_failures = []

                    for agent_name, subbranch in agent_subbranches.items():
                        success, conflict_info = self._merge_subbranch(subbranch, main_branch)
                        if success:
                            print(f"  ✅ {agent_name} 的更改已合并")
                            self._cleanup_subbranch(subbranch)
                        else:
                            merge_failures.append((agent_name, conflict_info))
                            print(f"  ❌ {agent_name} 合并失败: {conflict_info}")

                    if merge_failures:
                        print(f"\n⚠️ 检测到合并冲突！")
                        print(f"   以下分支保留供手动处理:")
                        for agent, _ in merge_failures:
                            print(f"     - {agent_subbranches.get(agent, 'unknown')}")
                elif agent_subbranches:
                    # 有 agent 失败，清理子分支
                    for subbranch in agent_subbranches.values():
                        self._cleanup_subbranch(subbranch)

                if any(r.status == AgentStatus.FAILED for r in results):
                    failed = [r.agent_name for r in results if r.status == AgentStatus.FAILED]
                    print(f"\n❌ 以下 agents 执行失败: {', '.join(failed)}")
                    self._save_final_state(state, all_results, time.time() - start_time)
                    return False

            # 更新状态
            state["current_phase"] = phase_idx
            for name, result in all_results.items():
                state["agents_status"][name] = result.status.value
                result_dict = asdict(result)
                result_dict["status"] = result.status.value
                state["results"][name] = result_dict
            self.state_manager.save_state(state)

        # 显示汇总
        total_duration = time.time() - start_time
        self.monitor.display_summary(all_results, total_duration)
        self._save_final_state(state, all_results, total_duration)

        # 提示合并
        if feature_branch:
            print(f"\n{'='*60}")
            print(f"✅ 手动任务完成！当前在分支: {feature_branch}")
            print(f"{'='*60}")
            print(f"下一步：git add . && git commit -m \"完成手动任务\"")
            print(f"{'='*60}\n")

        # 提示进度文件位置 & 清理临时文件
        if self.progress_file:
            print(f"📝 本次进度记录: {self.progress_file.name}")
        self._cleanup_temp_agent_files()

        return True


# ============================================================
# CLI接口
# ============================================================

def _open_file_in_editor(file_path: Path) -> None:
    """
    在用户默认编辑器中打开文件并等待关闭

    跨平台支持:
    - Windows: 使用 start /wait，回退到 notepad
    - Linux/Mac: 使用 $EDITOR，回退到 nano/vi
    """
    import shutil

    file_str = str(file_path)

    if sys.platform == 'win32':
        try:
            subprocess.run(['cmd', '/c', 'start', '/wait', '', file_str], check=True)
        except subprocess.CalledProcessError:
            subprocess.run(['notepad', file_str])
    else:
        editor = os.environ.get('EDITOR', '')
        if not editor:
            for ed in ['code', 'nano', 'vim', 'vi']:
                if shutil.which(ed):
                    editor = ed
                    break

        if editor:
            subprocess.run([editor, file_str])
        else:
            print(f"⚠️ 无法找到文本编辑器。请手动编辑: {file_path}")
            input("编辑完成后按回车继续...")


def _get_git_state(project_root: Path) -> tuple:
    """获取当前 git 状态快照（已修改文件 + 未跟踪文件）"""
    import subprocess
    try:
        r1 = subprocess.run(
            ["git", "diff", "--name-only"],
            cwd=str(project_root), capture_output=True, text=True, timeout=10
        )
        changed = set(f.strip() for f in r1.stdout.strip().split('\n') if f.strip())

        r2 = subprocess.run(
            ["git", "ls-files", "--others", "--exclude-standard"],
            cwd=str(project_root), capture_output=True, text=True, timeout=10
        )
        untracked = set(f.strip() for f in r2.stdout.strip().split('\n') if f.strip())

        return changed, untracked
    except Exception:
        return set(), set()


def _validate_architect_changes(project_root: Path, before_changed: set = None, before_untracked: set = None):
    """
    后置校验：检查 architect 是否越权修改了非 .md 文件
    只回滚 architect 会话期间新增的改动，不影响之前已有的未提交改动

    参数:
        before_changed: architect 启动前已修改的文件集合
        before_untracked: architect 启动前已存在的未跟踪文件集合
    """
    import subprocess

    if before_changed is None:
        before_changed = set()
    if before_untracked is None:
        before_untracked = set()

    try:
        after_changed, after_untracked = _get_git_state(project_root)

        # 只关注 architect 期间新增的改动
        new_changed = after_changed - before_changed
        new_untracked = after_untracked - before_untracked

        # 过滤出非 .md 文件
        violated_changed = [f for f in new_changed if not f.lower().endswith('.md')]
        violated_new = [f for f in new_untracked if not f.lower().endswith('.md')
                        and not f.startswith('.claude/')]

        if violated_changed or violated_new:
            print(f"\n{'='*60}")
            print(f"⚠️  ARCHITECT 后置校验：检测到越权修改！")
            print(f"{'='*60}")

            if violated_changed:
                print(f"\n   architect 期间被修改的非 .md 文件:")
                for f in sorted(violated_changed):
                    print(f"     ❌ {f}")
                subprocess.run(
                    ["git", "checkout", "--"] + list(violated_changed),
                    cwd=str(project_root), timeout=10
                )
                print(f"\n   ✅ 已回滚 {len(violated_changed)} 个被修改的文件")

            if violated_new:
                print(f"\n   architect 期间新建的非 .md 文件:")
                for f in sorted(violated_new):
                    print(f"     ❌ {f}")
                for f in violated_new:
                    file_path = project_root / f
                    if file_path.exists():
                        file_path.unlink()
                print(f"\n   ✅ 已删除 {len(violated_new)} 个越权创建的文件")

            print(f"\n{'='*60}\n")
        else:
            print(f"\n✅ Architect 后置校验通过：未检测到越权修改")

        # 日志：显示跳过的已有改动（帮助调试）
        skipped_changed = after_changed & before_changed
        if skipped_changed:
            non_md_skipped = [f for f in skipped_changed if not f.lower().endswith('.md')]
            if non_md_skipped:
                print(f"   ℹ️  跳过 {len(non_md_skipped)} 个 architect 之前已存在的改动（不回滚）")

    except Exception as e:
        print(f"\n⚠️ Architect 后置校验异常: {e}")


def semi_auto_mode(project_root: Path, config: dict):
    """
    情景2：半自动执行模式

    流程：
    1. 进入 claude CLI（plan 模式）讨论任务需求
    2. 生成 PLAN.md 后退出 claude
    3. 用户确认 PLAN.md
    4. 自动执行剩余 agents
    """
    import subprocess

    # 读取 architect 角色配置
    arch_file = project_root / ".claude" / "agents" / "01-arch.md"
    if arch_file.exists():
        with open(arch_file, 'r', encoding='utf-8') as f:
            content = f.read()
        # 分离 YAML frontmatter
        if content.startswith('---'):
            parts = content.split('---', 2)
            if len(parts) >= 3:
                arch_prompt = parts[2].strip()
            else:
                arch_prompt = content
        else:
            arch_prompt = content
    else:
        arch_prompt = "你是一个系统架构师，请分析需求并生成 PLAN.md"

    # 添加强制限制的系统提示
    project_root_str = str(project_root).replace('\\', '/')
    system_prompt = f"""{arch_prompt}

---

## ⚠️ 关键限制 - 必须严格遵守

**你是 Architect Agent，你的唯一任务是制定计划，而不是实现代码！**

### 🚨 PLAN.md 输出规则（非常重要！）

**所有输出文件必须保存在项目根目录，使用相对路径。**

**PLAN.md 生成规则：**
1. 先用 Read 工具检查 `PLAN.md` 是否已存在
2. **如果已存在** → 使用 **Edit 工具更新**（追加或修改相关内容，保留原有计划）
3. **如果不存在** → 使用 **Write 工具创建** `PLAN.md`
4. 路径必须是相对路径：`PLAN.md`（不要加任何目录前缀）

- ❌ 不要把计划写在对话中，必须写入文件
- ❌ 不要依赖 Claude CLI 的内置 plan 机制

**其他输出文件：**
| 文件名 | 位置 | 说明 |
|--------|------|------|
| `CODEBASE_ANALYSIS.md` | 项目根目录 | 代码库分析（现有项目） |

### 你必须做的事：
1. 分析用户需求
2. 如果是现有项目，先探索代码库并生成 `CODEBASE_ANALYSIS.md`
3. 检查 `PLAN.md` 是否已存在，按上述规则创建或更新
4. 完成后告知用户输入 `/exit` 退出会话

### 你绝对不能做的事：
- ❌ 不要编写任何实现代码
- ❌ 不要创建源代码文件（如 .py, .js, .ts 等）
- ❌ 不要修改现有的源代码
- ❌ 不要运行测试或构建命令
- ❌ 不要尝试"帮用户完成任务"

### 为什么？
你是多 Agent 流水线的第一个环节。你的输出（PLAN.md）将交给后续的 Developer、Tester、Security 等 agents 执行。如果你直接实现代码，就破坏了整个流程。

当用户描述完需求后，请开始分析并生成计划文件。
"""

    print(f"\n{'='*60}", flush=True)
    print(f"🎯 半自动模式 - 与 Architect 讨论任务", flush=True)
    print(f"{'='*60}", flush=True)
    print(f"💡 在 Claude CLI 中描述你的任务需求", flush=True)
    print(f"📄 讨论完成后，Architect 会生成 PLAN.md", flush=True)
    print(f"🚪 生成完毕后输入 /exit 退出，继续执行后续流程", flush=True)
    print(f"{'='*60}\n", flush=True)

    # 进入 claude CLI（hooks 提供 architect 权限管控，无需 plan mode）
    cmd = [
        "claude",
        "--dangerously-skip-permissions",
        "--append-system-prompt", system_prompt,
        "--max-budget-usd", str(config['max_budget']),
    ]

    env = os.environ.copy()
    env['ORCHESTRATOR_RUNNING'] = 'true'
    env['ORCHESTRATOR_AGENT'] = 'architect'  # Hook 用此变量检测 architect 阶段

    # 创建锁文件（Hook 备用检测方式）
    lock_file = project_root / ".claude" / "architect_active.lock"
    lock_file.parent.mkdir(parents=True, exist_ok=True)
    lock_file.write_text("architect session active", encoding='utf-8')

    # 记录 architect 启动前的 git 状态（用于后置校验的 baseline）
    before_changed, before_untracked = _get_git_state(project_root)
    print(f"  📸 [baseline] architect 启动前已有 {len(before_changed)} 个修改文件、{len(before_untracked)} 个未跟踪文件", flush=True)

    try:
        # 执行 claude（阻塞，用户交互）
        process = subprocess.run(cmd, cwd=str(project_root), env=env)
    finally:
        # 清理锁文件
        if lock_file.exists():
            lock_file.unlink()

    # 后置校验：暂时禁用（Hook 已能阻止越权，避免误回滚工作进度）
    # _validate_architect_changes(project_root, before_changed, before_untracked)

    # 检查 PLAN.md 是否生成
    plan_file = project_root / "PLAN.md"
    if not plan_file.exists():
        print(f"\n⚠️ 未检测到 PLAN.md，流程终止")
        print(f"   请重新运行并确保生成 PLAN.md")
        return False

    # 提示用户确认
    print(f"\n{'='*60}")
    print(f"📋 已检测到 PLAN.md")
    print(f"   位置: {plan_file}")
    print(f"{'='*60}")

    # 显示 PLAN.md 前几行（带容错）
    try:
        with open(plan_file, 'r', encoding='utf-8', errors='replace') as f:
            preview = f.read(500)
        print(f"\n--- PLAN.md 预览 ---")
        print(preview)
        if len(preview) >= 500:
            print("... (更多内容请查看文件)")
        print(f"--- 预览结束 ---\n")
    except (IOError, OSError, UnicodeDecodeError) as e:
        print(f"\n⚠️ 读取 PLAN.md 预览失败: {e}")
        print(f"   文件路径: {plan_file}")
        print(f"   请检查文件是否存在且可读\n")

    # 直接读取 PLAN.md 并执行后续 agents（跳过编辑/确认步骤）
    print(f"\n🚀 自动进入执行阶段...")

    # 读取 PLAN.md 作为任务描述（带容错）
    try:
        with open(plan_file, 'r', encoding='utf-8', errors='replace') as f:
            plan_content = f.read()
    except (IOError, OSError, UnicodeDecodeError) as e:
        print(f"\n❌ 无法读取 PLAN.md: {e}")
        print(f"   文件路径: {plan_file}")
        return False

    if not plan_content.strip():
        print(f"\n⚠️ PLAN.md 文件为空，无法继续执行")
        return False

    # 创建 orchestrator 执行剩余 agents
    max_rounds = config.get('max_rounds', 1)
    orchestrator = Orchestrator(
        project_root=project_root,
        max_budget=config['max_budget'],
        max_retries=config['max_retries'],
        verbose=config['verbose'],
        interactive_architect=False,  # architect 已完成
        max_rounds=max_rounds
    )

    # 执行剩余阶段（跳过 architect）
    print(f"\n🚀 开始执行后续 Agents...")
    if max_rounds > 1:
        print(f"   迭代模式: 最多 {max_rounds} 轮 developer-tester 循环")
        success = asyncio.run(orchestrator.execute_from_plan_with_loop(plan_content))
    else:
        success = asyncio.run(orchestrator.execute_from_plan(plan_content))

    return success


def from_plan_mode(project_root: Path, config: dict) -> bool:
    """
    从 PLAN.md 继续执行模式

    用于以下场景：
    - 用户已用其他 AI（如 GPT/Gemini/Grok）生成了 PLAN.md
    - 用户想跳过 architect 阶段节省 token
    - 直接执行 tech_lead 到 security 的后续 agents
    """
    plan_file = project_root / "PLAN.md"

    # 检查 PLAN.md 是否存在
    if not plan_file.exists():
        print(f"\n❌ 未找到 PLAN.md 文件")
        print(f"   请先生成计划文件：")
        print(f"   - 使用模式 1（半自动模式）生成")
        print(f"   - 或用其他 AI 生成后保存为 PLAN.md")
        return False

    # 读取 PLAN.md 内容
    try:
        with open(plan_file, 'r', encoding='utf-8', errors='replace') as f:
            plan_content = f.read()
    except (IOError, OSError) as e:
        print(f"\n❌ 无法读取 PLAN.md: {e}")
        return False

    if not plan_content.strip():
        print(f"\n⚠️ PLAN.md 文件为空，无法继续执行")
        return False

    # 显示 PLAN.md 预览
    print(f"\n{'='*60}")
    print(f"📋 检测到 PLAN.md")
    print(f"   位置: {plan_file}")
    print(f"{'='*60}")

    print(f"\n--- PLAN.md 预览 ---")
    preview = plan_content[:800]
    print(preview)
    if len(plan_content) > 800:
        print("... (更多内容请查看文件)")
    print(f"--- 预览结束 ---\n")

    # 确认执行
    confirm = input("确认跳过 Architect，执行后续 Agents？[Y/n] ").strip().lower()
    if confirm in ['n', 'no', '否']:
        print("已取消。")
        return False

    # 创建 orchestrator 执行剩余 agents
    max_rounds = config.get('max_rounds', 1)
    orchestrator = Orchestrator(
        project_root=project_root,
        max_budget=config['max_budget'],
        max_retries=config['max_retries'],
        verbose=config['verbose'],
        interactive_architect=False,
        max_rounds=max_rounds
    )

    print(f"\n🚀 开始执行后续 Agents（跳过 Architect）...")
    print(f"   将执行: tech_lead → developer → tester → optimizer → security")
    if max_rounds > 1:
        print(f"   迭代模式: 最多 {max_rounds} 轮 developer-tester 循环")
        success = asyncio.run(orchestrator.execute_from_plan_with_loop(plan_content))
    else:
        success = asyncio.run(orchestrator.execute_from_plan(plan_content))

    return success


def _ask_max_rounds() -> int:
    """询问用户选择迭代轮数"""
    print("""
开发-测试迭代轮数：
  1. 1轮（默认）- 线性执行，不循环
  2. 2轮 - 如有bug，developer-tester再迭代1次
  3. 3轮 - 最多迭代3次
""")
    rounds_choice = input("请选择 [1/2/3，直接回车=1]: ").strip()

    if rounds_choice == '2':
        return 2
    elif rounds_choice == '3':
        return 3
    else:
        return 1


def _ask_task_complexity() -> TaskComplexity:
    """询问用户选择任务复杂度"""
    print("""
任务复杂度：
  1. 简单任务 - 只用 developer + tester（2个agents，快速执行）
  2. 复杂任务 - 完整流程（6个agents，全面保障）
""")
    complexity_choice = input("请选择 [1/2，直接回车=2]: ").strip()

    if complexity_choice == '1':
        return TaskComplexity.MINIMAL
    else:
        return TaskComplexity.COMPLEX


def interactive_mode(project_root: Path):
    """交互式 CLI 模式 - 默认进入半自动模式"""
    print("""
╔════════════════════════════════════════════════════════════╗
║       🚀 mc-dir - 多Agent智能调度系统                       ║
╚════════════════════════════════════════════════════════════╝

选择执行模式：
  1. 半自动模式（推荐）- 进入 Claude CLI 讨论需求，生成 PLAN.md 后自动执行
  2. 从 PLAN.md 继续 - 跳过 Architect，直接从现有计划执行（节省 token）
  3. 全自动模式 - 输入任务后，Architect 自动规划并执行全流程
  4. （ADV）多agent模式* - 可同时指派多名 Agents🚀🚀🚀
  5. 退出
""")

    # 默认配置
    config = {
        'max_budget': 10.0,
        'max_retries': 3,
        'verbose': False,
        'auto_architect': False,
        'max_rounds': 1
    }

    choice = input("请选择 [1/2/3/4/5]: ").strip()

    if choice == '5':
        print("\n👋 再见！")
        return

    # 模式 1/2/3 都需要询问迭代轮数和任务复杂度
    if choice in ['1', '2', '3', '']:
        # 询问迭代轮数
        config['max_rounds'] = _ask_max_rounds()
        if config['max_rounds'] > 1:
            print(f"✓ 已设置: 最多 {config['max_rounds']} 轮 developer-tester 迭代\n")

        # 询问任务复杂度
        config['complexity'] = _ask_task_complexity()
        complexity_label = "简单任务（2个agents）" if config['complexity'] == TaskComplexity.MINIMAL else "复杂任务（6个agents）"
        print(f"✓ 已设置: {complexity_label}\n")

    if choice == '1' or choice == '':
        # 半自动模式
        # 注意：半自动模式会进入 Claude CLI 生成 PLAN.md，复杂度设置会被忽略
        if config.get('complexity') == TaskComplexity.MINIMAL:
            print("⚠️ 注意：半自动模式会由 Architect 自动规划，复杂度设置将被忽略\n")
        success = semi_auto_mode(project_root, config)
        if success:
            print("\n✅ 所有 Agents 执行完成！")
        return

    if choice == '2':
        # 从 PLAN.md 继续执行
        # 注意：PLAN.md 已存在，复杂度设置会被忽略
        if config.get('complexity') == TaskComplexity.MINIMAL:
            print("⚠️ 注意：从 PLAN.md 继续模式会按计划执行，复杂度设置将被忽略\n")
        success = from_plan_mode(project_root, config)
        if success:
            print("\n✅ 所有 Agents 执行完成！")
        return

    if choice == '3':
        # 全自动模式
        print("\n请输入任务描述（或 .md 文件路径）：")
        task_input = input("> ").strip()
        if not task_input:
            print("❌ 任务不能为空")
            return

        # 如果是 .md 文件，读取内容
        if task_input.endswith('.md'):
            task_file = project_root / task_input
            if task_file.exists():
                with open(task_file, 'r', encoding='utf-8') as f:
                    task_input = f.read()
            else:
                print(f"❌ 文件不存在: {task_file}")
                return

        orchestrator = Orchestrator(
            project_root=project_root,
            max_budget=config['max_budget'],
            max_retries=config['max_retries'],
            verbose=config['verbose'],
            interactive_architect=False,  # 全自动
            max_rounds=config['max_rounds']
        )

        print(f"\n🚀 全自动模式启动...")
        if config['max_rounds'] > 1:
            success = asyncio.run(orchestrator.execute_with_loop(
                task_input,
                override_complexity=config.get('complexity')
            ))
        else:
            success = asyncio.run(orchestrator.execute(
                task_input,
                override_complexity=config.get('complexity')
            ))

        if success:
            print("\n✅ 所有 Agents 执行完成！")
        return

    # 传统交互模式（选项 4）
    print("\n进入传统交互模式。输入 help 查看帮助，exit 退出。")

    while True:
        try:
            user_input = input("\n💬 有什么可以帮您？\n> ").strip()

            if not user_input:
                continue

            cmd_lower = user_input.lower()

            # 特殊命令
            if cmd_lower in ['exit', 'quit', 'q', '退出']:
                print("\n👋 再见！")
                break

            if cmd_lower in ['help', '?', '帮助']:
                print("""
📖 使用帮助
============================================================

【自动规划模式】直接描述需求：
  帮我写一个网页版的赛车游戏
  修复 src/main.py 中的登录 bug

【手动指定模式】使用 @agent 语法：
  @tech_lead 审核代码                    # 单个 agent
  @dev task1.md                          # 从 md 文件读取任务
  @dev task1.md && @opti task2.md        # 多 agent + md 文件
  @tech_lead 审核 && @security 安检      # 并行执行
  @tech_lead 审核 -> @developer 修复     # 串行执行
  @tech 审核 -> (@dev 修复 && @sec 安检) # 混合模式

特殊命令：
  help, ?       - 显示帮助
  agents        - 查看可用 agent 和别名
  config        - 查看/修改配置
  resume        - 恢复上次中断的任务
  status        - 查看当前状态
  exit, quit    - 退出程序

配置选项（在需求后添加）：
  --budget N    - 设置预算（USD）
  --auto        - 跳过交互式规划
  --verbose     - 详细日志
============================================================
""")
                continue

            if cmd_lower in ['agents', 'agent', '列表']:
                print("""
📋 可用的 Agents：
============================================================
  @architect  (别名: @arch, @架构)    - 系统架构师
  @tech_lead  (别名: @tech, @技术)    - 技术负责人
  @developer  (别名: @dev, @开发)     - 开发工程师
  @tester     (别名: @test, @测试)    - 测试工程师
  @optimizer  (别名: @opti, @优化)    - 优化专家
  @security   (别名: @sec, @安全)     - 安全专家

语法说明：
  ->   串行执行（前一个完成后执行下一个）
  &&   并行执行（同时执行）
  ()   分组（用于混合模式）

示例：
  @tech_lead 审核代码 -> @developer 根据建议修复
  @tester 测试 && @security 安全检查
============================================================
""")
                continue

            if cmd_lower == 'config':
                print(f"\n⚙️ 当前配置：")
                print(f"   预算上限: ${config['max_budget']}")
                print(f"   重试次数: {config['max_retries']}")
                print(f"   详细日志: {'是' if config['verbose'] else '否'}")
                print(f"   自动规划: {'是' if config['auto_architect'] else '否（交互式）'}")
                print(f"\n修改配置：config budget 20 / config verbose on")
                continue

            if cmd_lower.startswith('config '):
                parts = cmd_lower.split()
                if len(parts) >= 3:
                    key, value = parts[1], parts[2]
                    if key == 'budget':
                        config['max_budget'] = float(value)
                        print(f"✅ 预算设置为 ${config['max_budget']}")
                    elif key == 'verbose':
                        config['verbose'] = value in ['on', 'true', '1', '是']
                        print(f"✅ 详细日志: {'开启' if config['verbose'] else '关闭'}")
                    elif key == 'auto':
                        config['auto_architect'] = value in ['on', 'true', '1', '是']
                        print(f"✅ 自动规划: {'开启' if config['auto_architect'] else '关闭'}")
                continue

            # resume_mode 标志：用于后续执行时保留状态
            resume_mode = False

            if cmd_lower == 'resume':
                state_file = project_root / ".claude" / "state.json"
                if state_file.exists():
                    with open(state_file, 'r', encoding='utf-8') as f:
                        state = json.load(f)
                    print(f"📂 找到中断的任务: {state.get('user_request', '未知')}")
                    confirm = input("是否恢复？[Y/n] ").strip().lower()
                    if confirm not in ['n', 'no', '否']:
                        user_input = state['user_request']
                        resume_mode = True  # 标记为恢复模式，后续不清空状态
                        # 继续执行（落入后续逻辑）
                    else:
                        continue
                else:
                    print("❌ 没有找到可恢复的任务")
                    continue

            if cmd_lower == 'status':
                state_file = project_root / ".claude" / "state.json"
                if state_file.exists():
                    with open(state_file, 'r', encoding='utf-8') as f:
                        state = json.load(f)
                    print(f"\n📊 任务状态：")
                    print(f"   任务: {state.get('user_request', '未知')[:50]}")
                    print(f"   复杂度: {state.get('complexity', '未知')}")
                    print(f"   当前阶段: {state.get('current_phase', 0)}")
                    print(f"   总成本: ${state.get('total_cost', 0):.4f}")
                else:
                    print("📊 当前没有进行中的任务")
                continue

            if cmd_lower == 'clear':
                import os
                os.system('cls' if os.name == 'nt' else 'clear')
                continue

            # 解析命令行选项
            max_budget = config['max_budget']
            auto_architect = config['auto_architect']
            verbose = config['verbose']

            if '--budget' in user_input:
                import re
                match = re.search(r'--budget\s+(\d+(?:\.\d+)?)', user_input)
                if match:
                    max_budget = float(match.group(1))
                user_input = re.sub(r'--budget\s+\d+(?:\.\d+)?', '', user_input).strip()

            if '--auto' in user_input:
                auto_architect = True
                user_input = user_input.replace('--auto', '').strip()

            if '--verbose' in user_input:
                verbose = True
                user_input = user_input.replace('--verbose', '').strip()

            if not user_input:
                continue

            # 检测是否是手动指定模式
            manual_parser = ManualTaskParser(project_root)

            if manual_parser.is_manual_mode(user_input):
                # ========== 手动指定模式 ==========
                phases, success = manual_parser.parse(user_input)

                if not success:
                    continue

                # 预览执行计划
                manual_parser.preview(phases)
                print(f"   预算上限: ${max_budget}")

                confirm = input("\n确认执行？[Y/n] ").strip().lower()
                if confirm in ['n', 'no', '否']:
                    print("已取消")
                    continue

                # 创建 orchestrator 并执行手动任务
                orchestrator = Orchestrator(
                    project_root=project_root,
                    max_budget=max_budget,
                    max_retries=config['max_retries'],
                    verbose=verbose,
                    interactive_architect=not auto_architect
                )

                success = asyncio.run(orchestrator.execute_manual(phases, clean_start=True))

                if success:
                    print("\n✅ 手动任务完成！可以继续输入新需求。")
                else:
                    print("\n❌ 任务执行失败，请检查错误日志。")

            else:
                # ========== 自动规划模式 ==========
                task_parser = TaskParser(project_root)
                _, complexity = task_parser.parse(user_input)

                scheduler = AgentScheduler()
                phases = scheduler.plan_execution(complexity)
                total_agents = sum(len(p) for p in phases)

                print(f"\n📋 自动规划模式 - 任务预览：")
                print(f"   需求: {user_input[:60]}{'...' if len(user_input) > 60 else ''}")
                print(f"   复杂度: {complexity.value}")
                print(f"   执行阶段: {len(phases)} 个阶段，{total_agents} 个 Agent")
                print(f"   预算上限: ${max_budget}")
                print(f"   规划模式: {'自动' if auto_architect else '交互式'}")

                # 显示执行计划
                print(f"\n   执行计划：")
                for i, phase_agents in enumerate(phases, 1):
                    agent_names = ', '.join(phase_agents)
                    print(f"     Phase {i}: {agent_names}")

                confirm = input("\n确认执行？[Y/n] ").strip().lower()
                if confirm in ['n', 'no', '否']:
                    print("已取消")
                    continue

                # 创建 orchestrator 并执行
                orchestrator = Orchestrator(
                    project_root=project_root,
                    max_budget=max_budget,
                    max_retries=config['max_retries'],
                    verbose=verbose,
                    interactive_architect=not auto_architect
                )

                success = asyncio.run(orchestrator.execute(user_input, clean_start=not resume_mode))

                if success:
                    print("\n✅ 任务完成！可以继续输入新需求。")
                else:
                    print("\n❌ 任务执行失败，请检查错误日志。")

                # 重置 resume_mode 以便下次循环
                resume_mode = False

        except KeyboardInterrupt:
            print("\n\n⚠️ 中断当前任务")
            continue
        except EOFError:
            print("\n\n👋 再见！")
            break
        except Exception as e:
            print(f"\n❌ 错误: {e}")
            continue


def find_project_root() -> Path:
    """
    递归向上查找项目根目录（包含 .git 的目录）

    修复 Bug #7: 当在 src/ 子目录运行时，Path.cwd() 返回错误路径
    此函数通过查找 .git 目录确保返回真正的项目根目录

    Returns:
        Path: 项目根目录路径
    """
    current = Path.cwd()
    max_depth = 10  # 防止无限递归

    for _ in range(max_depth):
        if (current / '.git').exists():
            return current

        parent = current.parent
        if parent == current:  # 到达文件系统根目录
            break
        current = parent

    # 找不到 .git，使用当前目录
    return Path.cwd()


def _select_account() -> str:
    """
    选择 Claude 账户

    Returns:
        选中的账户标识 ('mc' 或 'xh')
    """
    print("""
╔════════════════════════════════════════════════════════════╗
║       🔐 Claude 账户选择                                    ║
╚════════════════════════════════════════════════════════════╝

可用账户：
  mc - Claude Pro 账户 (mc)
  xh - Claude Pro 账户 (xh)
""")

    while True:
        choice = input("请选择账户 [mc/xh，直接回车=mc]: ").strip().lower()

        if not choice:
            choice = 'mc'

        if choice in CLAUDE_CONFIG_DIRS:
            config_dir = CLAUDE_CONFIG_DIRS[choice]

            # 检查配置目录是否存在
            if not os.path.exists(config_dir):
                print(f"⚠️ 警告: 配置目录不存在: {config_dir}")
                print(f"   请先运行 'claude-{choice}' 初始化配置\n")
                continue

            # 设置环境变量
            os.environ['CLAUDE_CONFIG_DIR'] = config_dir
            print(f"✓ 已选择账户: {choice}")
            print(f"✓ 配置目录: {config_dir}\n")
            return choice
        else:
            print(f"❌ 无效选择: {choice}，请输入 'mc' 或 'xh'\n")


def main():
    """CLI入口"""
    # 步骤0: 选择 Claude 账户
    selected_account = _select_account()

    parser = argparse.ArgumentParser(
        description="mc-dir - 多Agent智能调度系统",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用方式：

  情景1 - 全自动执行（复杂任务从 md 文件读取）：
    python mc-dir.py task1.md --auto-architect

  情景2 - 半自动执行（进入 Claude CLI 讨论后自动执行）：
    python mc-dir.py

  恢复中断的任务：
    python mc-dir.py --resume
        """
    )

    parser.add_argument(
        "request",
        nargs="?",
        help="任务描述或 .md 文件路径（不指定则进入半自动模式）"
    )
    parser.add_argument(
        "--max-budget",
        type=float,
        default=10.0,
        help="最大预算（USD），默认10.0"
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=3,
        help="最大重试次数，默认3"
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="详细日志输出"
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="从上次中断处恢复"
    )
    parser.add_argument(
        "--auto-architect",
        action="store_true",
        help="全自动模式（跳过交互式规划）"
    )
    parser.add_argument(
        "--from-plan",
        action="store_true",
        help="从 PLAN.md 开始执行（跳过 architect，节省 token）"
    )
    parser.add_argument(
        "--max-rounds",
        type=int,
        default=1,
        help="developer-tester 循环最大轮数（默认1，即不循环）"
    )

    args = parser.parse_args()

    # 获取项目根目录（Bug #7 修复：使用 find_project_root() 而非 Path.cwd()）
    project_root = find_project_root()

    # --from-plan 模式：直接从 PLAN.md 开始
    if args.from_plan:
        plan_file = project_root / "PLAN.md"
        if not plan_file.exists():
            print(f"❌ PLAN.md 不存在，无法使用 --from-plan 模式")
            print(f"   请先生成 PLAN.md 文件")
            sys.exit(1)

        try:
            with open(plan_file, 'r', encoding='utf-8', errors='replace') as f:
                plan_content = f.read()
        except (IOError, OSError) as e:
            print(f"❌ 无法读取 PLAN.md: {e}")
            sys.exit(1)

        if not plan_content.strip():
            print(f"❌ PLAN.md 文件为空")
            sys.exit(1)

        print(f"📋 从 PLAN.md 开始执行（跳过 Architect）")
        orchestrator = Orchestrator(
            project_root=project_root,
            max_budget=args.max_budget,
            max_retries=args.max_retries,
            verbose=args.verbose,
            interactive_architect=False
        )

        try:
            success = asyncio.run(orchestrator.execute_from_plan(plan_content))
            if success:
                print("\n✅ 所有 Agents 执行完成！")
            sys.exit(0 if success else 1)
        except KeyboardInterrupt:
            print("\n\n⚠️ 用户中断执行")
            print("   状态已保存，可使用 --resume 恢复")
            sys.exit(130)

    # 无参数时进入半自动模式
    if not args.request and not args.resume:
        interactive_mode(project_root)
        return

    # 情景1：从 .md 文件读取任务描述
    user_request = args.request
    if user_request and user_request.endswith('.md'):
        task_file = project_root / user_request
        if task_file.exists():
            print(f"📄 从文件读取任务: {user_request}", flush=True)
            with open(task_file, 'r', encoding='utf-8') as f:
                user_request = f.read()
        else:
            print(f"❌ 任务文件不存在: {task_file}")
            sys.exit(1)

    # 创建orchestrator实例
    orchestrator = Orchestrator(
        project_root=project_root,
        max_budget=args.max_budget,
        max_retries=args.max_retries,
        verbose=args.verbose,
        interactive_architect=not args.auto_architect,
        max_rounds=args.max_rounds
    )

    # 恢复模式
    if args.resume:
        state = orchestrator.state_manager.load_state()
        if state:
            print(f"📂 恢复任务: {state['user_request'][:50]}...")
            # 检查是否是从 PLAN.md 执行的任务
            if state.get('complexity') == 'from_plan':
                # 读取 PLAN.md 继续执行
                plan_file = project_root / "PLAN.md"
                if plan_file.exists():
                    with open(plan_file, 'r', encoding='utf-8') as f:
                        plan_content = f.read()
                    try:
                        # 传入现有状态，跳过已完成的 agents
                        success = asyncio.run(orchestrator.execute_from_plan(plan_content, existing_state=state))
                        sys.exit(0 if success else 1)
                    except KeyboardInterrupt:
                        print("\n\n⚠️ 用户中断执行")
                        sys.exit(130)
                else:
                    print("❌ PLAN.md 不存在，无法恢复")
                    sys.exit(1)
            else:
                user_request = state['user_request']
        else:
            print("❌ 未找到可恢复的任务")
            sys.exit(1)

    # 执行
    try:
        # resume 模式不清理旧状态，新任务则清理
        clean_start = not args.resume

        # 如果 max_rounds > 1，使用带循环的执行模式
        if args.max_rounds > 1:
            print(f"🔄 多轮循环模式: 最多 {args.max_rounds} 轮 developer-tester 迭代")
            success = asyncio.run(orchestrator.execute_with_loop(user_request, clean_start=clean_start))
        else:
            success = asyncio.run(orchestrator.execute(user_request, clean_start=clean_start))

        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⚠️ 用户中断执行")
        sys.exit(130)
    except Exception as e:
        print(f"\n❌ 执行错误: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
