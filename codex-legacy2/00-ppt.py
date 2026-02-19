#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""统一入口：执行 new-ppt-workflow v3 全流程。"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
HISTORY = ROOT / "iteration_history.md"


def run(cmd: list[str]) -> int:
    print(f"[RUN] {' '.join(cmd)}")
    return subprocess.run(cmd, cwd=str(ROOT), check=False).returncode


def save_history(line: str):
    if not HISTORY.exists():
        HISTORY.write_text("# iteration_history\n\n", encoding="utf-8")
    with HISTORY.open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def main() -> int:
    ap = argparse.ArgumentParser(description="一键运行PPT新工作流")
    ap.add_argument("--start-version", default="1.0")
    ap.add_argument("--max-rounds", type=int, default=3)
    ap.add_argument("--from-step", type=int, default=1, choices=[1, 2, 3, 4])
    ap.add_argument("--to-step", type=int, default=4, choices=[1, 2, 3, 4])
    args = ap.parse_args()

    if args.from_step > args.to_step:
        print("[ERROR] from-step cannot be greater than to-step")
        return 2

    py = sys.executable

    base_major = int(float(args.start_version))
    base_minor = int(round((float(args.start_version) - base_major) * 10))

    for round_idx in range(args.max_rounds):
        version = f"{base_major}.{base_minor + round_idx}"
        target = f"codex {version}.pptx"
        print("\n" + "=" * 78)
        print(f"Round {round_idx + 1}/{args.max_rounds} - version {version}")
        print("=" * 78)

        plan = [
            (1, [py, "01-shape-detail.py"]),
            (2, [py, "02-shape-analysis.py"]),
            (3, [py, "03-build_shape.py"]),
            (3, [py, "03-build_ppt_com.py", "--version", version]),
            (4, [py, "04-shape_diff_test.py", "--target", target]),
        ]

        fail = None
        for step_no, cmd in plan:
            if not (args.from_step <= step_no <= args.to_step):
                continue
            rc = run(cmd)
            if rc != 0:
                fail = (step_no, rc)
                break

        ts = datetime.now().isoformat(timespec="seconds")
        if fail:
            save_history(f"- {ts} version={version} status=FAIL step={fail[0]} rc={fail[1]}")
            print(f"[FAIL] round stopped at step {fail[0]}, rc={fail[1]}")
            # diff fail允许进入下一轮；其他步骤失败立即终止
            if fail[0] != 4:
                return fail[1]
            continue

        # if step4 executed and passed, we can stop
        diff_result = ROOT / "diff_result.json"
        status = "unknown"
        if diff_result.exists():
            try:
                status = json.loads(diff_result.read_text(encoding="utf-8")).get("status", "unknown")
            except Exception:
                status = "unknown"

        save_history(f"- {ts} version={version} status={status.upper()}")
        if status == "ok" or args.to_step < 4:
            print(f"[DONE] workflow finished at version {version}, status={status}")
            return 0

    print("[DONE] reached max rounds; please inspect fix-ppt.md and iteration_history.md")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
