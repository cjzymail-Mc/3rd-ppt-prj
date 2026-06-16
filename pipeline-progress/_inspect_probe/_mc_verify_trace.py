#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""主 Claude 独立验证 Step 2：用 fake shape 驱动 03b 的真实 except 埋点，
证明 com_api_failed_but_continued + shape_write_end 在运行时真的写进 jsonl。
纯 Python，无 COM/PowerPoint。"""
import importlib.util, os, sys, json
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.abspath(".")
spec = importlib.util.spec_from_file_location("b03", os.path.join(ROOT, "pipeline", "03b_build_ppt_com.py"))
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)

print("TraceLogger importable:", m._TraceLogger is not None)
if m._TraceLogger is None:
    print("[BLOCKER] office-com-helpers TraceLogger 没装/没 import 到 → trace 永远 no-op")
    sys.exit(1)

trace_path = os.path.join(ROOT, "pipeline-progress", "_inspect_probe", "_step2_trace_test.jsonl")
if os.path.exists(trace_path):
    os.remove(trace_path)
m._TRACE = m._TraceLogger(trace_path)

# fake shape: 有 TextFrame，但写 .Text 时抛错 → 驱动 _write_text 的 except 埋点
class FakeFont:
    Name = ""
class FakeTR:
    def __init__(self):
        self.Font = FakeFont()
    @property
    def Text(self):
        return "x"
    @Text.setter
    def Text(self, v):
        raise RuntimeError("forced COM write failure")
class FakeTF:
    def __init__(self):
        self.AutoSize = 0
        self.TextRange = FakeTR()
class FakeShape:
    Name = "FakeShape"
    HasTextFrame = -1
    def __init__(self):
        self.TextFrame = FakeTF()

r = m._write_text(FakeShape(), "hello world")
print("_write_text result:", r)

with m._trace_shape("FakeShape2", "text") as _sw:
    pass  # 正常退出 → shape_write_end ok

m._TRACE.close()

print("\n--- trace jsonl events ---")
with open(trace_path, encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if not line:
            continue
        ev = json.loads(line)
        print("event:", ev.get("event") or ev.get("name") or ev.get("type"), "|", json.dumps(ev, ensure_ascii=False))
