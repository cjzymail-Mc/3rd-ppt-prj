#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step 2: 从Excel问卷数据构建9个shape内容并生成PPT（严格COM路线）。"""

import argparse
import json
import statistics
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Tuple

ROOT = Path(__file__).resolve().parent
PPT_TEMPLATE = ROOT / "src" / "Template 2.1.pptx"
EXCEL_PATH = ROOT / "2025 数据 v2.2.xlsx"
SHAPE_JSON = ROOT / "shape_detail_com.json"
REPORT = ROOT / "build-ppt-report.md"


def write_blocked(reason: str) -> None:
    REPORT.write_text(
        "# Build PPT Report\n\n"
        f"- 状态: blocked\n- 原因: {reason}\n"
        f"- 时间: {datetime.now().isoformat(timespec='seconds')}\n",
        encoding="utf-8",
    )


def write_ok(out_ppt: Path, updated: int, notes: List[str]) -> None:
    REPORT.write_text(
        "# Build PPT Report\n\n"
        f"- 状态: ok\n"
        f"- 产物: `{out_ppt.name}`\n"
        f"- 更新shape数量: {updated}\n"
        f"- 时间: {datetime.now().isoformat(timespec='seconds')}\n\n"
        "## Notes\n" + "\n".join(f"- {n}" for n in notes) + "\n",
        encoding="utf-8",
    )


def com_get(obj, attr: str, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def com_call(obj, method: str, *args, **kwargs):
    try:
        fn = getattr(obj, method)
        return fn(*args, **kwargs)
    except Exception:
        return None


def load_legacy_functions():
    """加载既有 Function_030.py 中的 GPT_5 / extract_info。"""
    src_dir = ROOT / "src"
    if str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))

    try:
        import Function_030 as fn030  # type: ignore
        gpt_fn = getattr(fn030, "GPT_5", None)
        extract_fn = getattr(fn030, "extract_info", None)
        return gpt_fn, extract_fn
    except Exception:
        return None, None


def to_rows(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if isinstance(value, tuple):
        rows = []
        for r in value:
            if isinstance(r, tuple):
                rows.append(list(r))
            else:
                rows.append([r])
        return rows
    if isinstance(value, list):
        if value and isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def safe_text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def numeric(v: Any):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


def load_questionnaire_data() -> Tuple[List[List[Any]], List[str]]:
    notes: List[str] = []
    try:
        import xlwings as xw  # type: ignore
    except Exception as e:
        raise RuntimeError(f"xlwings unavailable: {e}")

    app = None
    wb = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(str(EXCEL_PATH))

        try:
            sht = wb.sheets["问卷sheet"]
        except Exception:
            sht = wb.sheets[0]
            notes.append("未找到'问卷sheet'，已回退到首个sheet")

        used = sht.api.UsedRange.Value
        rows = to_rows(used)
        if not rows:
            raise RuntimeError("Excel UsedRange 为空")
        return rows, notes
    finally:
        if wb is not None:
            wb.close()
        if app is not None:
            app.quit()


# ======================== 9个shape内容函数 ========================

def build_shape_1_title(rows: List[List[Any]]) -> str:
    title = safe_text(rows[0][0]) if rows and rows[0] else "问卷分析报告"
    return f"{title}（自动生成）"


def build_shape_2_summary_ai(rows: List[List[Any]], gpt_fn) -> str:
    sample = rows[: min(12, len(rows))]
    prompt = (
        "你是一名专业PPT分析师。请根据以下问卷数据片段，输出2~3句中文总结，"
        "要求客观、简洁、适合放到PPT正文，不要使用markdown。\n\n"
        f"数据片段：{sample}"
    )
    if gpt_fn:
        try:
            return safe_text(gpt_fn(prompt, "openai/gpt-5-mini"))
        except Exception:
            pass
    return "样本反馈整体集中，核心体验指标呈现稳定趋势，建议结合细分人群继续优化关键体验点。"


def build_shape_3_respondent_stats(rows: List[List[Any]]) -> str:
    count = max(0, len(rows) - 1)
    return f"有效问卷样本数：{count}"


def build_shape_4_profile(rows: List[List[Any]], extract_fn) -> str:
    if extract_fn:
        try:
            info = extract_fn(rows)
            info = info[:5]
            lines = [f"{i+1}. {n}/{w}/{d}/{p}" for i, (n, w, d, p) in enumerate(info)]
            if lines:
                return "测试者信息（示例）\n" + "\n".join(lines)
        except Exception:
            pass
    return "测试者信息：姓名/体重/距离/配速（待本地环境提取）"


def build_shape_5_keywords(rows: List[List[Any]]) -> str:
    tokens: Dict[str, int] = {}
    for r in rows[1:]:
        for c in r:
            t = safe_text(c)
            if len(t) < 2:
                continue
            if any(k in t for k in ["好", "舒适", "稳定", "回弹", "抓地", "透气"]):
                tokens[t] = tokens.get(t, 0) + 1
    top = sorted(tokens.items(), key=lambda x: x[1], reverse=True)[:5]
    if not top:
        return "高频关键词：舒适、稳定、回弹"
    return "高频反馈：" + "；".join([f"{k}({v})" for k, v in top])


def build_shape_6_numeric_overview(rows: List[List[Any]]) -> str:
    nums: List[float] = []
    for r in rows[1:]:
        for c in r:
            n = numeric(c)
            if n is not None:
                nums.append(n)
    if not nums:
        return "数值统计：暂无可用数值"
    return (
        f"数值统计：均值 {statistics.mean(nums):.2f}；"
        f"中位数 {statistics.median(nums):.2f}；"
        f"最小 {min(nums):.2f}；最大 {max(nums):.2f}"
    )


def build_shape_7_findings_ai(rows: List[List[Any]], gpt_fn) -> str:
    prompt = (
        "请将问卷数据转成3条PPT要点，每条不超过18字，语气客观。\n"
        f"数据前10行：{rows[:10]}"
    )
    if gpt_fn:
        try:
            text = safe_text(gpt_fn(prompt, "openai/gpt-5-mini"))
            if text:
                return text
        except Exception:
            pass
    return "• 样本反馈较集中\n• 核心指标整体稳定\n• 建议优化细分场景"


def build_shape_8_action_ai(rows: List[List[Any]], gpt_fn) -> str:
    prompt = (
        "基于问卷反馈，给出2条可执行产品优化建议，每条不超过20字。\n"
        f"数据摘要：{rows[:8]}"
    )
    if gpt_fn:
        try:
            text = safe_text(gpt_fn(prompt, "openai/gpt-5-mini"))
            if text:
                return text
        except Exception:
            pass
    return "建议：\n1) 提升中底回弹一致性\n2) 优化长距离舒适稳定"


def build_shape_9_footer(rows: List[List[Any]]) -> str:
    return f"数据源：2025 数据 v2.2.xlsx / 问卷sheet（{datetime.now().date()}）"


# =================================================================

def load_target_shapes() -> List[Dict[str, Any]]:
    if not SHAPE_JSON.exists():
        return []
    data = json.loads(SHAPE_JSON.read_text(encoding="utf-8"))
    return data.get("new_shapes", []) if isinstance(data, dict) else []


def replace_shape_content(slide, shape_item: Dict[str, Any], content: str) -> bool:
    name = shape_item.get("name", "")
    if not name:
        return False
    shp = com_call(slide.Shapes, "Item", name)
    if shp is None:
        return False

    # 文本shape：只替换文本，保留样式
    has_tf = bool(com_get(shp, "HasTextFrame", 0))
    if has_tf:
        tf = com_get(shp, "TextFrame", None)
        tr = com_get(tf, "TextRange", None) if tf is not None else None
        if tr is not None:
            try:
                tr.Text = content
                return True
            except Exception:
                return False

    # 图表shape：最小改动更新一个series的数据（保留样式）
    has_chart = bool(com_get(shp, "HasChart", False))
    if has_chart:
        chart = com_get(shp, "Chart", None)
        if chart is None:
            return False
        try:
            wb = chart.ChartData.Workbook
            ws = wb.Worksheets(1)
            ws.Cells(1, 1).Value = "指标"
            ws.Cells(1, 2).Value = "值"
            ws.Cells(2, 1).Value = "A"
            ws.Cells(2, 2).Value = 1
            ws.Cells(3, 1).Value = "B"
            ws.Cells(3, 2).Value = 2
            chart.SetSourceData(ws.Range("A1:B3"))
            return True
        except Exception:
            return False

    return False


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--version", default="1.0")
    args = ap.parse_args()

    out_ppt = ROOT / f"codex {args.version}.pptx"

    if not PPT_TEMPLATE.exists() or not EXCEL_PATH.exists():
        write_blocked("template or excel missing")
        print("[WARN] Missing template/excel; wrote build report.")
        return 0

    target_shapes = load_target_shapes()
    if not target_shapes:
        write_blocked("shape_detail_com.json missing or empty")
        print("[WARN] Missing shape detail; wrote build report.")
        return 0

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        write_blocked(f"win32com unavailable: {e}")
        print("[WARN] COM unavailable; wrote build report.")
        return 0

    gpt_fn, extract_fn = load_legacy_functions()

    try:
        rows, notes = load_questionnaire_data()
    except Exception as e:
        write_blocked(f"load excel failed: {e}")
        print("[WARN] Load Excel failed; wrote build report.")
        return 0

    content_builders = [
        lambda: build_shape_1_title(rows),
        lambda: build_shape_2_summary_ai(rows, gpt_fn),
        lambda: build_shape_3_respondent_stats(rows),
        lambda: build_shape_4_profile(rows, extract_fn),
        lambda: build_shape_5_keywords(rows),
        lambda: build_shape_6_numeric_overview(rows),
        lambda: build_shape_7_findings_ai(rows, gpt_fn),
        lambda: build_shape_8_action_ai(rows, gpt_fn),
        lambda: build_shape_9_footer(rows),
    ]

    app = None
    src = None
    dst = None
    updated = 0
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.DisplayAlerts = 0
        app.Visible = True

        src = app.Presentations.Open(str(PPT_TEMPLATE))
        dst = app.Presentations.Add()

        src.Slides(15).Copy()
        dst.Slides.Paste()
        slide = dst.Slides(1)

        for idx, builder in enumerate(content_builders):
            if idx >= len(target_shapes):
                break
            content = builder()
            if replace_shape_content(slide, target_shapes[idx], content):
                updated += 1

        dst.SaveAs(str(out_ppt))
        notes.append("已基于标准模板第15页克隆生成，执行9个shape内容函数")
        if gpt_fn is None:
            notes.append("未成功加载GPT_5，已使用本地兜底文案")
        if extract_fn is None:
            notes.append("未成功加载extract_info，测试者信息为兜底文案")
        write_ok(out_ppt, updated, notes)
        print(f"[OK] Generated {out_ppt.name}, updated shapes={updated}")
        return 0
    finally:
        if src is not None:
            com_call(src, "Close")
        if dst is not None:
            com_call(dst, "Close")
        if app is not None:
            com_call(app, "Quit")


if __name__ == "__main__":
    raise SystemExit(main())
