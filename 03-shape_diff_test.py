#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""兼容入口：转发到04-shape_diff_test.py。"""
import subprocess
import sys


def main() -> int:
    cmd = [sys.executable, "04-shape_diff_test.py"]
    return subprocess.run(cmd, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
