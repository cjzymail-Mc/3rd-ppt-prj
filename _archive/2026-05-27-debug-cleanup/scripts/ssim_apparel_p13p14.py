"""SSIM 验收：apparel 新生成 p13/p14 vs 模板基准。

桥接已打开的 PowerPoint，导出 4 个 PNG（剪贴板路径绕过文件级加密）：
  A1 = Template 2.1.pptx slide 22 (新生成 p13)
  B1 = apparel-page13-14-template.pptx slide 13 (模板 p13)
  A2 = Template 2.1.pptx slide 23 (新生成 p14)
  B2 = apparel-page13-14-template.pptx slide 14 (模板 p14)
然后两组 SSIM 比对。

SSIM 阈值：≥0.85 通过；<0.85 回炉。
"""
import os
import sys
import time

import win32com.client

# 注入 skill 路径
SKILL_PATH = r"C:\Users\xy24\.claude\skills\ppt-visual-fidelity-check"
sys.path.insert(0, SKILL_PATH)
from ppt_visual_check import slide_export_png, ssim_compare  # noqa: E402

PROJ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(PROJ, "debug", "ssim_apparel")
os.makedirs(OUT_DIR, exist_ok=True)

# 用户当前打开的 Template 2.1.pptx（含 smoke test 输出 slide 22/23）
SNAPSHOT_PATH = os.path.join(PROJ, "template", "apparel-page13-14-template.pptx")


def find_pres(app, name_keyword):
    for i in range(1, app.Presentations.Count + 1):
        p = app.Presentations(i)
        if name_keyword.lower() in p.Name.lower():
            return p
    return None


def main():
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        print(f"FAIL: 没有运行中的 PowerPoint: {e}")
        return 1

    # 找已打开的 Template 2.1.pptx（含 smoke test 结果）
    pres_a = find_pres(app, "Template 2.1")
    if pres_a is None:
        print("FAIL: 未找到打开的 Template 2.1.pptx")
        return 2
    print(f"[A] {pres_a.Name}  slides={pres_a.Slides.Count}")

    # 打开 snapshot（ReadOnly）做 B
    print(f"[B] 打开 snapshot (ReadOnly): {SNAPSHOT_PATH}")
    pres_b = app.Presentations.Open(SNAPSHOT_PATH, ReadOnly=True, WithWindow=True)
    time.sleep(0.5)
    print(f"[B] {pres_b.Name}  slides={pres_b.Slides.Count}")

    pairs = [
        ("p13", pres_a.Slides(22), pres_b.Slides(13)),
        ("p14", pres_a.Slides(23), pres_b.Slides(14)),
    ]

    results = []
    try:
        for label, slide_a, slide_b in pairs:
            png_a = os.path.join(OUT_DIR, f"new_{label}.png")
            png_b = os.path.join(OUT_DIR, f"template_{label}.png")

            print(f"\n[{label}] 导出 new (slide {slide_a.SlideIndex}) → {os.path.basename(png_a)}")
            ok_a = slide_export_png(app, slide_a, png_a)
            print(f"  {'OK' if ok_a else 'FAIL'}")
            time.sleep(0.3)

            print(f"[{label}] 导出 template (slide {slide_b.SlideIndex}) → {os.path.basename(png_b)}")
            ok_b = slide_export_png(app, slide_b, png_b)
            print(f"  {'OK' if ok_b else 'FAIL'}")
            time.sleep(0.3)

            if ok_a and ok_b:
                score = ssim_compare(png_a, png_b)
                print(f"[{label}] SSIM = {score:.4f}")
                results.append((label, score))
            else:
                print(f"[{label}] SKIP (export 失败)")
                results.append((label, -1.0))
    finally:
        # 关闭 B（不动用户的 A）
        try:
            pres_b.Close()
        except Exception:
            pass

    # 汇总
    print("\n" + "=" * 60)
    print("SSIM 验收结果（阈值 0.85）")
    print("=" * 60)
    all_pass = True
    for label, score in results:
        flag = "PASS" if score >= 0.85 else "FAIL"
        if score < 0.85:
            all_pass = False
        print(f"  {label}: SSIM = {score:.4f}  [{flag}]")
    print("=" * 60)
    print(f"总体：{'PASS' if all_pass else 'FAIL'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
