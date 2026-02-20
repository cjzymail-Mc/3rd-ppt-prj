#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Architect 越权防护 Hook

规则1: ExitPlanMode — 无条件拦截（不依赖 architect 检测）
规则2: Bash — 仅 architect 阶段拦截
规则3: Write/Edit 非.md — 仅 architect 阶段拦截

检测: 环境变量 ORCHESTRATOR_AGENT=architect 或锁文件（绝对路径）
格式: exit code 2 = 阻止, exit code 0 = 放行
"""
import sys
import json
import os
import datetime

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "guard_debug.log")


def log_debug(msg):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            f.write(f"[{ts}] {msg}\n")
    except Exception:
        pass


def find_lock_file():
    candidates = []
    # 方法A: 基于脚本位置（脚本在 .claude/hooks/，lock 在 .claude/）
    script_dir = os.path.dirname(os.path.abspath(__file__))
    claude_dir = os.path.dirname(script_dir)
    candidates.append(os.path.join(claude_dir, "architect_active.lock"))
    # 方法B: 基于 cwd
    candidates.append(os.path.join(".claude", "architect_active.lock"))
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def is_architect_active():
    env_val = os.environ.get("ORCHESTRATOR_AGENT", "")
    if env_val == "architect":
        log_debug(f"DETECTED architect via env var")
        return True
    lock = find_lock_file()
    if lock:
        log_debug(f"DETECTED architect via lock file: {lock}")
        return True
    log_debug(f"NOT architect. env='{env_val}', cwd={os.getcwd()}")
    return False


def main():
    try:
        input_data = json.load(sys.stdin)
    except Exception as e:
        log_debug(f"PARSE ERROR: {e}")
        sys.exit(0)

    tool_name = input_data.get("tool_name", "")
    tool_input = input_data.get("tool_input", {})
    log_debug(f"HOOK CALLED: tool={tool_name}")

    # =============================================
    # 规则1: ExitPlanMode — 无条件拦截
    # plan mode 批准→执行是越权根源，一律阻止
    # =============================================
    if tool_name == "ExitPlanMode":
        msg = (
            "ARCHITECT GUARD: ExitPlanMode blocked. "
            "Your PLAN.md has been saved. Do NOT execute the plan. "
            "Tell the user: 'PLAN.md is ready, please type /exit to proceed.'"
        )
        log_debug(f"BLOCKED (unconditional): ExitPlanMode")
        print(msg, file=sys.stderr)
        sys.exit(2)

    # =============================================
    # 规则2-3: 需要 architect 检测
    # 非 architect 阶段全部放行，不影响其他 agent
    # =============================================
    if not is_architect_active():
        log_debug(f"ALLOW (not architect): {tool_name}")
        sys.exit(0)

    # --- 以下仅 architect 阶段生效 ---

    if tool_name == "Bash":
        msg = "ARCHITECT GUARD: Bash blocked during architect phase."
        log_debug(f"BLOCKED: Bash")
        print(msg, file=sys.stderr)
        sys.exit(2)

    if tool_name in ["Write", "Edit"]:
        file_path = tool_input.get("file_path", "")
        if file_path.lower().endswith(".md"):
            log_debug(f"ALLOW (.md): {tool_name} -> {file_path}")
            sys.exit(0)
        msg = f"ARCHITECT GUARD: Blocked write to '{file_path}'. Architect can only write .md files."
        log_debug(f"BLOCKED: {tool_name} -> {file_path}")
        print(msg, file=sys.stderr)
        sys.exit(2)

    log_debug(f"ALLOW (other): {tool_name}")
    sys.exit(0)


if __name__ == "__main__":
    main()
