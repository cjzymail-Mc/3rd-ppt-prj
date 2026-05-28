#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""test_src_smoke.py — src/ 模块冒烟测试（yzr / zxh）

用途：在修改 src/yzr_ppt.py / src/zxh_ppt.py / src/_ppt_shared.py 后，
快速验证：
  1. Python 语法正确（import 不抛）
  2. SHAPES 定义结构完整
  3. _build_content / _build_rich_prompt 对 mock 数据不抛异常
  4. （可选）真实打开模板 PPT，校验 shape 名称存在 — 需要 PowerPoint COM

触发：
  - 每次修改 src/yzr_ppt.py 或 src/zxh_ppt.py 后手动跑
  - 共享模块 _ppt_shared.py 变更后必跑
  - 不做视觉 diff、不做完整 GPT 调用（交给 reviewer agent）

用法:
  python debug/test_src_smoke.py          # 仅纯 Python 测试
  python debug/test_src_smoke.py --com    # 加上 COM 打开模板校验 shape 名
"""

from __future__ import annotations

import sys
import traceback
from pathlib import Path

# 把项目根加入 sys.path，以便 `from src.xxx import ...` 可用
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


# ---------------------------------------------------------------------------
# Mock data — 最小可用的样本
# ---------------------------------------------------------------------------
_MOCK_ROWS = [
    ["姓名", "体重（KG）", "减震评分", "回弹评分", "稳定评分", "补充说明"],
    ["受访者A", 70, 8, 7, 9, "包裹性不错，但鞋带有点短，急停时抓地稍弱。"],
    ["受访者B", 75, 7, 8, 8, "整体舒适，后跟偶尔掉跟，缓震偏硬。"],
    ["受访者C", 68, 9, 8, 9, "上脚一体感好，止滑性在木地板表现稳定。"],
]


# ---------------------------------------------------------------------------
# Test helpers
# ---------------------------------------------------------------------------
_PASS = 0
_FAIL = 0
_FAIL_DETAIL: list[str] = []


def _t(name: str, fn):
    """Run one test; capture assertion/exception into _FAIL_DETAIL."""
    global _PASS, _FAIL
    try:
        fn()
        _PASS += 1
        print(f"  [ok] {name}")
    except Exception as exc:
        _FAIL += 1
        _FAIL_DETAIL.append(f"{name}: {exc}")
        print(f"  [FAIL] {name}: {exc}")
        traceback.print_exc()


# ---------------------------------------------------------------------------
# Per-module smoke tests
# ---------------------------------------------------------------------------
def _smoke_yzr():
    print("\n=== yzr_ppt ===")

    def _import():
        from src import yzr_ppt  # noqa: F401

    def _shapes_shape():
        from src.yzr_ppt import CODEX_SHAPES
        assert isinstance(CODEX_SHAPES, list) and CODEX_SHAPES
        for spec in CODEX_SHAPES:
            assert "name" in spec and "strategy" in spec, f"spec 缺字段: {spec}"

    def _build_content_ok():
        from src.yzr_ppt import CODEX_SHAPES, _build_content
        for spec in CODEX_SHAPES:
            if spec["strategy"] in ("skip", "extract_image"):
                continue
            # gpt_enabled=False → 走 fallback，不触发网络调用
            _build_content(spec, _MOCK_ROWS, gpt_enabled=False, model="")

    def _build_prompt_ok():
        from src.yzr_ppt import _build_rich_prompt
        p = _build_rich_prompt(
            budget={"max_chars": 200, "max_lines": 6},
            rows=_MOCK_ROWS,
            focus="优点",
        )
        assert isinstance(p, str) and len(p) > 50

    _t("import", _import)
    _t("CODEX_SHAPES 结构", _shapes_shape)
    _t("_build_content 不抛", _build_content_ok)
    _t("_build_rich_prompt 不抛", _build_prompt_ok)


def _smoke_zxh():
    print("\n=== zxh_ppt ===")

    def _import():
        from src import zxh_ppt  # noqa: F401

    def _shapes_shape():
        from src.zxh_ppt import ZXH_SHAPES
        assert isinstance(ZXH_SHAPES, list) and ZXH_SHAPES
        for spec in ZXH_SHAPES:
            assert "name" in spec and "strategy" in spec, f"spec 缺字段: {spec}"

    def _build_content_ok():
        from src.zxh_ppt import ZXH_SHAPES, _build_content
        for spec in ZXH_SHAPES:
            if spec["strategy"] in ("skip", "extract_image"):
                continue
            _build_content(
                spec, _MOCK_ROWS, gpt_enabled=False, model="",
                style_anchor=spec.get("template_text", ""),
            )

    def _build_prompt_p1p2_ok():
        from src.zxh_ppt import _build_rich_prompt
        p = _build_rich_prompt(
            budget={"max_chars": 97, "max_lines": 8},
            rows=_MOCK_ROWS,
            focus="修改建议",
            fmt="p1p2",
        )
        assert isinstance(p, str) and "P1" in p and "P2" in p

    _t("import", _import)
    _t("ZXH_SHAPES 结构", _shapes_shape)
    _t("_build_content 不抛", _build_content_ok)
    _t("_build_rich_prompt p1p2 不抛", _build_prompt_p1p2_ok)


# ---------------------------------------------------------------------------
# Optional: COM-based shape name validation
# ---------------------------------------------------------------------------
def _com_shape_check():
    """打开 Template 2.1.pptx，确认 SHAPES 中的 name 在对应 slide 真实存在。"""
    print("\n=== COM shape name check (--com) ===")
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        print(f"  [skip] win32com 不可用: {exc}")
        return

    tpl = ROOT / "src" / "Template 2.1.pptx"
    if not tpl.exists():
        print(f"  [skip] 模板不存在: {tpl}")
        return

    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        ppt = app.Presentations.Open(str(tpl), WithWindow=False)
    except Exception as exc:
        print(f"  [FAIL] 无法打开模板: {exc}")
        return

    try:
        from src.yzr_ppt import CODEX_SHAPES, _TEMPLATE_SLIDE as YZR_SLIDE
        from src.zxh_ppt import ZXH_SHAPES, _TEMPLATE_SLIDE as ZXH_SLIDE

        def _check(shapes_spec, slide_idx, label):
            def inner():
                slide = ppt.Slides(slide_idx)
                present = set()
                for i in range(1, slide.Shapes.Count + 1):
                    try:
                        present.add(slide.Shapes(i).Name)
                    except Exception:
                        pass
                missing = [s["name"] for s in shapes_spec if s["name"] not in present]
                assert not missing, f"Slide {slide_idx} 缺 shape: {missing}"
            _t(f"{label} shape 名在模板中存在", inner)

        _check(CODEX_SHAPES, YZR_SLIDE, "yzr")
        _check(ZXH_SHAPES, ZXH_SLIDE, "zxh")
    finally:
        try:
            ppt.Close()
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main(argv):
    want_com = "--com" in argv

    _smoke_yzr()
    _smoke_zxh()

    if want_com:
        _com_shape_check()
    else:
        print("\n(跳过 COM shape 校验；加 --com 启用)")

    print(f"\n==== summary: pass={_PASS}  fail={_FAIL} ====")
    if _FAIL:
        print("\n失败明细:")
        for line in _FAIL_DETAIL:
            print(f"  - {line}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
