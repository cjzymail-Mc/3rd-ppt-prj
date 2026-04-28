#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""diagnose_chart_write.py — PPT chart 写入诊断脚本.

用途：验证"原地 COM 写入 chart 数据"在跨机（Build 19929 / Build 4266）+ 加密文件场景
是否可行——**关键前提：chart 必须是未被污染的 fresh 状态**。

重要背景（路线重估）：
    XML surgery 路径已废弃 —— 办公室默认加密 pptx 是 CFB 容器，`zipfile` 无法读取。
    唯一剩余"100% 还原模板"的手段就是 COM 写入。
    而 COM 写入的罪魁祸首是 **BreakLink / Activate** —— 它们把 healthy chart 搞坏。
    纯 STRAT 1（裸 series.Values = tuple）已在同事机器健康 chart 上验证可行。

推荐使用（最小化污染）：
    python skills/diagnose_chart_write.py --strat1

    只跑纯裸写入，不调 BreakLink / Activate / Refresh / Workbook，
    用于验证 "fresh 模板 chart 能否直接被 COM 写入"。

完整诊断（会污染 chart 状态，慎用）：
    python skills/diagnose_chart_write.py --all

前置：
    1. PPT 打开一个 **fresh**（未被此前跑过 BreakLink 的）模板
    2. 选中 slide 上的 chart shape
    3. 运行（所有操作都在 PPT 内存中进行，不写磁盘）
"""
from __future__ import annotations

import os
import sys
import platform
import shutil
import time
import traceback
import zipfile
import re
from typing import Any, List, Optional, Tuple

# -------- 环境探测 -------------------------------------------------------
def env_report():
    print("=" * 60)
    print("[ENV] Python :", sys.version.replace("\n", " "))
    print("[ENV] OS     :", platform.platform())
    try:
        import win32com
        print("[ENV] pywin32:", win32com.__gen_path__[:60], "...")
    except Exception as e:
        print("[ENV] pywin32: ERROR", e)
    try:
        import pythoncom
        print("[ENV] pythoncom VT_ARRAY:", pythoncom.VT_ARRAY)
    except Exception as e:
        print("[ENV] pythoncom: ERROR", e)
    print()


# -------- 连接 PowerPoint -------------------------------------------------
def attach_ppt():
    import win32com.client
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
        print(f"[PPT] 连接已打开的 PowerPoint 成功")
    except Exception as e:
        print(f"[PPT] GetActiveObject 失败: {e}")
        print(f"[PPT] 请确保 PPT 已打开并选中一个 chart 后重试")
        sys.exit(1)

    try:
        ver = app.Version
        build = getattr(app, "Build", "?")
        print(f"[PPT] Version={ver}  Build={build}")
    except Exception as e:
        print(f"[PPT] 版本读取失败: {e}")
    return app


def get_selected_chart_shape(app):
    try:
        sel = app.ActiveWindow.Selection
        stype = int(sel.Type)
        print(f"[SEL] SelectionType={stype} (2=Shapes, 3=Text)")
        if stype not in (2, 3):
            print("[SEL] 请在 PPT 里选中 chart shape 后再跑")
            sys.exit(1)
        sr = sel.ShapeRange
        cnt = int(sr.Count)
        for i in range(1, cnt + 1):
            sh = sr.Item(i)
            try:
                has_chart = bool(sh.HasChart)
            except Exception:
                has_chart = False
            if has_chart:
                print(f"[SEL] 选中 chart shape: Name={sh.Name!r}")
                return sh
        print(f"[SEL] 选中的 {cnt} 个 shape 里没有 chart")
        sys.exit(1)
    except Exception as e:
        print(f"[SEL] 失败: {e}")
        traceback.print_exc()
        sys.exit(1)


# -------- Chart 基础信息 --------------------------------------------------
def report_chart(chart):
    def _try(attr_path, default="?"):
        try:
            obj = chart
            for a in attr_path.split("."):
                obj = getattr(obj, a)
            if callable(obj):
                return default
            return obj
        except Exception as e:
            return f"<err:{e}>"

    print(f"[CHART] Type={_try('ChartType')}")
    print(f"[CHART] HasTitle={_try('HasTitle')}")
    print(f"[CHART] ChartData.IsLinked={_try('ChartData.IsLinked')}")
    try:
        sc = chart.SeriesCollection()
        scount = int(sc.Count)
        print(f"[CHART] SeriesCount={scount}")
        for i in range(1, scount + 1):
            s = chart.SeriesCollection(i)
            try:
                vals = list(s.Values)
            except Exception as e:
                vals = f"<err:{e}>"
            try:
                xvals = list(s.XValues)
            except Exception as e:
                xvals = f"<err:{e}>"
            print(f"[CHART]   series{i} values={vals}")
            print(f"[CHART]   series{i} xvals ={xvals}")
    except Exception as e:
        print(f"[CHART] SeriesCollection 异常: {e}")


# -------- 写入策略 --------------------------------------------------------
def _readback(series):
    try:
        return list(series.Values)
    except Exception as e:
        return f"<readback err: {e}>"


def strat_1_plain_tuple(chart, labels, values):
    print("\n--- [STRAT 1] series.Values = tuple([...])  直接赋值 ---")
    try:
        s = chart.SeriesCollection(1)
        s.Values = tuple(values)
        s.XValues = tuple(labels)
        time.sleep(0.5)
        print(f"  写入完成，readback: {_readback(s)}")
    except Exception as e:
        print(f"  写入异常: {e}")


def strat_2_variant_wrap(chart, labels, values):
    print("\n--- [STRAT 2] VARIANT(VT_ARRAY|VT_R8, list)  显式 SAFEARRAY ---")
    try:
        import pythoncom
        from win32com.client import VARIANT
        s = chart.SeriesCollection(1)
        s.Values  = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,   list(values))
        s.XValues = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BSTR, list(labels))
        time.sleep(0.5)
        print(f"  写入完成，readback: {_readback(s)}")
    except Exception as e:
        print(f"  写入异常: {e}")


def strat_3_refresh_after(chart, labels, values):
    print("\n--- [STRAT 3] 写入 + chart.Refresh() ---")
    try:
        s = chart.SeriesCollection(1)
        s.Values = tuple(values)
        s.XValues = tuple(labels)
        try:
            chart.Refresh()
            print(f"  Refresh 成功")
        except Exception as e:
            print(f"  Refresh 异常: {e}")
        time.sleep(0.5)
        print(f"  写入完成，readback: {_readback(s)}")
    except Exception as e:
        print(f"  写入异常: {e}")


def strat_4_break_then_write(chart, labels, values):
    print("\n--- [STRAT 4] BreakLink → 不 Activate → 直接写 ---")
    try:
        try:
            chart.ChartData.BreakLink()
            time.sleep(0.3)
            print(f"  BreakLink OK")
        except Exception as e:
            print(f"  BreakLink 异常（继续）: {e}")
        s = chart.SeriesCollection(1)
        s.Values = tuple(values)
        s.XValues = tuple(labels)
        time.sleep(0.5)
        print(f"  写入完成，readback: {_readback(s)}")
        try:
            print(f"  IsLinked 现在 = {chart.ChartData.IsLinked}")
        except Exception:
            pass
    except Exception as e:
        print(f"  写入异常: {e}")


def _patch_chart_xml_in_pptx(pptx_path: str, match_labels: List[str],
                              new_labels: List[str], new_values: List[float]) -> Tuple[bool, str]:
    """Unzip pptx, find the chart{N}.xml whose strCache matches match_labels,
    rewrite its numCache + strCache, rezip.

    Returns (success, matched_chart_part_name).
    """
    tmp_path = pptx_path + ".patchtmp"
    matched_part = ""

    with zipfile.ZipFile(pptx_path, "r") as zin:
        names = zin.namelist()
        chart_parts = [n for n in names if re.match(r"ppt/charts/chart\d+\.xml$", n)]
        if not chart_parts:
            return False, "no chart parts found"

        # Find which chart matches the selected shape (by strCache labels)
        target = None
        for cp in chart_parts:
            xml = zin.read(cp).decode("utf-8", errors="replace")
            # Extract all <c:pt><c:v>...</c:v></c:pt> inside the first <c:strCache>
            m = re.search(r"<c:strCache\b[^>]*>(.*?)</c:strCache>", xml, re.DOTALL)
            if not m:
                continue
            strs_in_cache = re.findall(r"<c:v>([^<]*)</c:v>", m.group(1))
            # match heuristically: same count + at least first 3 labels identical
            if len(strs_in_cache) == len(match_labels) and all(
                strs_in_cache[i] == match_labels[i] for i in range(min(3, len(match_labels)))
            ):
                target = cp
                break
        if target is None:
            # fallback: if only one chart, use it
            if len(chart_parts) == 1:
                target = chart_parts[0]
            else:
                return False, f"no chart matched labels; candidates={chart_parts}"
        matched_part = target

        # Build replacement XML
        original_xml = zin.read(target).decode("utf-8", errors="replace")
        new_xml = _rewrite_caches(original_xml, new_labels, new_values)

        # Write new zip
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for n in names:
                if n == target:
                    zout.writestr(n, new_xml)
                else:
                    zout.writestr(n, zin.read(n))

    shutil.move(tmp_path, pptx_path)
    return True, matched_part


def _rewrite_caches(xml: str, new_labels: List[str], new_values: List[float]) -> str:
    """Replace the first <c:strCache> and the first <c:numCache> contents."""
    def _build_str_cache(labels):
        pts = "".join(
            f'<c:pt idx="{i}"><c:v>{_xml_escape(lab)}</c:v></c:pt>'
            for i, lab in enumerate(labels)
        )
        return f'<c:strCache><c:ptCount val="{len(labels)}"/>{pts}</c:strCache>'

    def _build_num_cache(values):
        pts = "".join(
            f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
            for i, v in enumerate(values)
        )
        return f'<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{len(values)}"/>{pts}</c:numCache>'

    # Replace first strCache
    xml = re.sub(
        r"<c:strCache\b[^>]*>.*?</c:strCache>",
        _build_str_cache(new_labels),
        xml, count=1, flags=re.DOTALL,
    )
    # Replace first numCache
    xml = re.sub(
        r"<c:numCache\b[^>]*>.*?</c:numCache>",
        _build_num_cache(new_values),
        xml, count=1, flags=re.DOTALL,
    )
    return xml


def _xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def strat_6_xml_surgery(app, shape, labels, values):
    """Save → Close → XML patch → Reopen → readback.

    This is the target production path: bypass COM entirely for chart data,
    works identically on both Build 19929 and Build 4266.
    """
    print("\n--- [STRAT 6] XML surgery: Save → zip-patch chart1.xml → Reopen ---")
    try:
        pres = shape.Parent.Parent  # shape → slide → presentation
        path = pres.FullName
        shape_name = shape.Name
        slide_index = shape.Parent.SlideIndex
        # capture current labels (used to locate the right chart{N}.xml)
        try:
            current_labels = [str(x) for x in shape.Chart.SeriesCollection(1).XValues]
        except Exception:
            current_labels = []
        print(f"  path={path}")
        print(f"  slide_index={slide_index}  shape_name={shape_name!r}")
        print(f"  current_labels={current_labels}")

        if not path or not os.path.isfile(path):
            print(f"  ✗ pptx 未保存到磁盘或路径无效，跳过")
            return

        # Save + close
        try:
            pres.Save()
            print(f"  Save OK")
        except Exception as e:
            print(f"  Save 异常: {e}")
            return
        try:
            pres.Close()
            print(f"  Close OK")
            time.sleep(1.5)
        except Exception as e:
            print(f"  Close 异常: {e}")
            return

        # Backup
        bak = path + ".strat6bak"
        try:
            shutil.copy2(path, bak)
            print(f"  备份 → {bak}")
        except Exception as e:
            print(f"  备份异常（继续）: {e}")

        # XML patch
        try:
            ok, matched = _patch_chart_xml_in_pptx(path, current_labels, labels, values)
            print(f"  XML patch: ok={ok}  matched_part={matched}")
            if not ok:
                print(f"  ✗ 补丁失败：{matched}")
                # restore + reopen
                if os.path.exists(bak):
                    shutil.copy2(bak, path)
                app.Presentations.Open(path)
                return
        except Exception as e:
            print(f"  XML patch 异常: {e}")
            traceback.print_exc()
            if os.path.exists(bak):
                shutil.copy2(bak, path)
            app.Presentations.Open(path)
            return

        # Reopen
        try:
            new_pres = app.Presentations.Open(path)
            time.sleep(1.0)
            print(f"  Reopen OK")
        except Exception as e:
            print(f"  Reopen 异常: {e}")
            return

        # Re-locate chart on the same slide by name, readback
        try:
            slide = new_pres.Slides(slide_index)
            found = None
            for i in range(1, int(slide.Shapes.Count) + 1):
                sh = slide.Shapes(i)
                if str(sh.Name) == shape_name and bool(sh.HasChart):
                    found = sh
                    break
            if found is None:
                print(f"  ✗ reopen 后找不到 shape {shape_name!r}")
                return
            s = found.Chart.SeriesCollection(1)
            rb_vals = _readback(s)
            try:
                rb_x = list(s.XValues)
            except Exception:
                rb_x = "<err>"
            print(f"  readback values = {rb_vals}")
            print(f"  readback xvals  = {rb_x}")
            print(f"  → 请肉眼确认 bars 是否为 {values}")
        except Exception as e:
            print(f"  readback 异常: {e}")
            traceback.print_exc()

    except Exception as e:
        print(f"  整体异常: {e}")
        traceback.print_exc()


def strat_5_activate_then_workbook(chart, labels, values):
    print("\n--- [STRAT 5] Activate → ChartData.Workbook.Sheets(1) 写 cell ---")
    try:
        try:
            chart.ChartData.Activate()
            time.sleep(1.0)
            print(f"  Activate OK")
        except Exception as e:
            print(f"  Activate 异常: {e}")
            print(f"  继续尝试直接访问 Workbook ...")
        try:
            wb = chart.ChartData.Workbook
            ws = wb.Worksheets(1)
            print(f"  Workbook.Worksheets(1).Name = {ws.Name}")
            # 按照标准 chart 模板：A1 空, B1 系列名, A2..AN 类别, B2..BN 值
            for i, (lab, val) in enumerate(zip(labels, values)):
                ws.Cells(i + 2, 1).Value = lab
                ws.Cells(i + 2, 2).Value = val
            time.sleep(0.5)
            s = chart.SeriesCollection(1)
            print(f"  写入完成，readback: {_readback(s)}")
        except Exception as e:
            print(f"  Workbook 写入异常: {e}")
            traceback.print_exc()
    except Exception as e:
        print(f"  整体异常: {e}")


# -------- 主流程 ---------------------------------------------------------
def main():
    env_report()
    app = attach_ppt()
    shp = get_selected_chart_shape(app)
    chart = shp.Chart

    # 命令行模式：
    #   --strat1（默认，推荐）只跑纯 series.Values = tuple，不污染 chart
    #   --all                   跑 STRAT 1-5（会污染 chart，诊断用）
    #   --strat6                XML surgery（已知加密环境不可用，保留做反向验证）
    mode = "strat1"
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        if arg in ("--all", "-a"):
            mode = "all"
        elif arg in ("--strat6", "--xml", "-6"):
            mode = "strat6"
        elif arg in ("--strat1", "-1"):
            mode = "strat1"

    print(f"\n>>> 模式：{mode}")
    print(">>> 写入前：chart 状态（务必确认 values / xvals 非空，否则 chart 已被历史写入污染）")
    report_chart(chart)

    if mode == "strat1":
        # === 纯 STRAT 1 验证 ===
        # 目标：证明 "fresh 模板上裸 series.Values = tuple 在双机均能工作"
        # 绝不调用 BreakLink / Activate / Refresh / Workbook —— 这些是已知凶手
        print("\n" + "=" * 60)
        print(">>> 纯 STRAT 1 验证（裸 COM 写入，不调用 BreakLink/Activate/Refresh）")
        print("=" * 60)
        labels = ["S1-抓地", "S1-缓震", "S1-包裹", "S1-抗扭", "S1-透气", "S1-防侧", "S1-耐久"]
        values = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0]
        strat_1_plain_tuple(chart, labels, values)

        print("\n>>> 写入后：chart 状态")
        report_chart(chart)

        print("\n" + "=" * 60)
        print("验收标准（纯 STRAT 1）：")
        print("  ✅ 通过：readback=[1..7]  且  PPT 肉眼可见 bars 为 1/2/3/4/5/6/7")
        print("  ❌ 失败：readback=[]      或  bars 消失")
        print()
        print("如果通过 → 生产代码 `_write_chart` 删掉 BreakLink/Activate 即可")
        print("如果失败 → COM 路线对 fresh 模板也走不通，需讨论 make_chart 兜底")
        print("请把以上日志 + chart 肉眼结果贴回对话")
        return

    # === 完整诊断模式 ===
    base_labels = ["A", "B", "C", "D", "E", "F", "G"]
    base_values_1 = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0]
    base_values_2 = [2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0]
    base_values_3 = [3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0]
    base_values_4 = [4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0]
    base_values_5 = [5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 9.0]
    strat6_labels = ["S6-a", "S6-b", "S6-c", "S6-d", "S6-e", "S6-f", "S6-g"]
    strat6_values = [6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0]

    if mode == "all":
        strat_1_plain_tuple(chart, base_labels, base_values_1)
        time.sleep(1.0)
        strat_2_variant_wrap(chart, base_labels, base_values_2)
        time.sleep(1.0)
        strat_3_refresh_after(chart, base_labels, base_values_3)
        time.sleep(1.0)
        strat_4_break_then_write(chart, base_labels, base_values_4)
        time.sleep(1.0)
        strat_5_activate_then_workbook(chart, base_labels, base_values_5)
        time.sleep(1.0)

        print("\n>>> STRAT 1-5 后：chart 状态")
        report_chart(chart)

    if mode == "strat6":
        # 已知加密文件不可用，保留仅做反向验证
        print("\n⚠️  加密 pptx 无法走 zipfile，STRAT 6 预期失败（反向验证用）")
        strat_6_xml_surgery(app, shp, strat6_labels, strat6_values)

    print("\n" + "=" * 60)
    print("诊断结束。请把日志 + chart 肉眼结果贴回对话。")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        print("\n[FATAL] 主流程异常：")
        traceback.print_exc()
