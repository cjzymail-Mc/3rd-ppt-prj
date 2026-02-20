#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""兼容入口：执行 Step2+Step3。"""
import subprocess
import sys


def main() -> int:
    py = sys.executable
    for cmd in ([py, "02-shape-analysis.py"], [py, "03-build_shape.py"], [py, "03-build_ppt_com.py"]):
        rc = subprocess.run(cmd, check=False).returncode
        if rc != 0:
            return rc
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
