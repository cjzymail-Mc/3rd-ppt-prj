#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_copypicture.py
验证 CopyPicture(1, -4147) 是否产生矢量图（无 OLE 链接）。

测试流程：
1. 启动 Excel，写入测试数据并创建图表
2. CopyPicture(1, -4147) 复制到剪贴板
3. 启动 PPT，粘贴到 Slide
4. 检查粘贴后 shape 的类型（期望：图片 type=13，而非 OLE type=7）
5. 关闭 Excel、删除临时数据行，验证 PPT shape 不受影响
6. 清理
"""
import sys
import time
sys.stdout.reconfigure(encoding='utf-8')

import win32com.client

PASS = '[PASS]'
FAIL = '[FAIL]'

xl_app  = None
ppt_app = None

try:
    # ── 1. 启动 Excel，写入测试数据 ──────────────────────────────────────
    print('Step 1: 启动 Excel ...')
    xl_app = win32com.client.Dispatch('Excel.Application')
    xl_app.Visible = True
    wb = xl_app.Workbooks.Add()
    ws = wb.Sheets(1)

    # 写入简单数据
    ws.Range('A1').Value = '类别'
    ws.Range('B1').Value = '数值'
    ws.Range('A2').Value = 'A'
    ws.Range('B2').Value = 30
    ws.Range('A3').Value = 'B'
    ws.Range('B3').Value = 50
    ws.Range('A4').Value = 'C'
    ws.Range('B4').Value = 20

    # 创建图表
    chart_obj = ws.ChartObjects().Add(Left=10, Top=50, Width=300, Height=200)
    chart = chart_obj.Chart
    chart.ChartType = 51  # xlColumnClustered
    chart.SetSourceData(ws.Range('A1:B4'))
    time.sleep(1.0)
    print(f'  图表创建完成，ChartObjects数量={ws.ChartObjects().Count}')

    # ── 2. CopyPicture(1, -4147) ──────────────────────────────────────────
    print('Step 2: CopyPicture(Appearance=xlScreen, Format=xlPicture) ...')
    try:
        chart_obj.CopyPicture(1, -4147)
        copy_ok = True
        print(f'  {PASS} CopyPicture 成功')
    except Exception as e:
        copy_ok = False
        print(f'  {FAIL} CopyPicture 异常: {e}')
        # 回退 OLE
        chart_obj.Chart.Copy()
        print('  已回退 OLE Copy')

    time.sleep(0.8)

    # ── 3. 启动 PPT，粘贴 ────────────────────────────────────────────────
    print('Step 3: 启动 PPT，粘贴 shape ...')
    ppt_app = win32com.client.Dispatch('PowerPoint.Application')
    ppt_app.Visible = True
    prs = ppt_app.Presentations.Add()
    slide = prs.Slides.Add(1, 12)  # ppLayoutBlank=12
    time.sleep(0.5)

    ppt_app.Activate()
    time.sleep(0.3)

    shp = slide.Shapes.Paste()
    time.sleep(0.5)
    print(f'  粘贴成功，shape name={shp.Name}, type={shp.Type}')

    # ── 4. 验证 shape 类型 ───────────────────────────────────────────────
    # msoEmbeddedOLEObject=7, msoLinkedOLEObject=10, msoPicture=13, msoLinkedPicture=11
    shape_type = shp.Type
    if copy_ok:
        if shape_type == 13:
            print(f'  {PASS} shape.Type=13 (msoPicture) — 矢量图，无 OLE 链接 ✓')
        elif shape_type in (7, 10):
            print(f'  {FAIL} shape.Type={shape_type} (OLE 对象) — 仍有链接风险！')
        else:
            print(f'  shape.Type={shape_type}（其他类型，需人工核查）')
    else:
        print(f'  (OLE 回退) shape.Type={shape_type}')

    # ── 5. 删除 Excel 数据行，验证 PPT shape 不变 ───────────────────────
    print('Step 5: 删除 Excel 数据行（模拟 questionnaire 清理）...')
    ws.Range('A2:B4').EntireRow.Delete()
    time.sleep(0.5)

    # 如果 shape 是 OLE，删行可能导致 chart 崩溃；如果是图片，应无影响
    try:
        # 尝试访问 shape 属性（若 OLE 链接损坏，这里会异常）
        _ = shp.Width
        _ = shp.Height
        print(f'  {PASS} 删行后 PPT shape 仍可访问（Width={shp.Width:.0f}, Height={shp.Height:.0f}）')
    except Exception as e:
        print(f'  {FAIL} 删行后 shape 访问异常: {e}')

    # ── 6. 检查 HasChart（图片 shape 不含 Chart 属性）───────────────────
    if copy_ok:
        has_chart = getattr(shp, 'HasChart', None)
        if has_chart is False or has_chart is None:
            print(f'  {PASS} shape.HasChart=False — 确认为纯图片，无嵌入图表数据')
        else:
            try:
                hc = shp.HasChart
                print(f'  shape.HasChart={hc}')
            except:
                print(f'  shape 无 HasChart 属性 — 确认为图片类型')

    print('\n=== 测试结束 ===')

finally:
    # ── 清理 ────────────────────────────────────────────────────────────
    print('清理资源...')
    try:
        if ppt_app:
            for p in list(ppt_app.Presentations):
                p.Close()
            ppt_app.Quit()
    except Exception as e:
        print(f'  PPT 关闭异常: {e}')
    try:
        if xl_app:
            xl_app.DisplayAlerts = False
            for wb_ in list(xl_app.Workbooks):
                wb_.Close(SaveChanges=False)
            xl_app.Quit()
    except Exception as e:
        print(f'  Excel 关闭异常: {e}')
    print('清理完成')
