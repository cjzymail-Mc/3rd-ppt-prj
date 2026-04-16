"""
HTML 演示稿一键 Sanity Check
统一入口：依次运行编码检查、HTML结构检查、CSS语法检查

用法:
    python pipeline/sanity_check.py <html_file>          # 仅检查
    python pipeline/sanity_check.py <html_file> --fix     # 检查并自动修复编码问题
"""

import sys
from pathlib import Path

# 将 pipeline 目录加入路径
sys.path.insert(0, str(Path(__file__).parent))

from check_encoding import check_encoding
from check_html import check_html
from check_css import check_css


def run_all(filepath: str, fix: bool = False):
    print("=" * 60)
    print(f"  Sanity Check: {Path(filepath).name}")
    print("=" * 60)

    total_errors = 0
    total_warnings = 0

    # --- Step 1: 编码检查 ---
    print("\n[1/3] Encoding Check (curly quotes in HTML tags)")
    print("-" * 40)
    enc_issues = check_encoding(filepath, fix=fix)
    enc_errors = len(enc_issues)
    total_errors += enc_errors
    if not enc_issues:
        print("  PASS")
    elif fix:
        print(f"  FIXED {enc_errors} issue(s)")
        total_errors -= enc_errors  # 已修复不计入错误
    else:
        print(f"  FAIL: {enc_errors} curly quote(s) found")
        for i in enc_issues:
            print(f"    Line {i['line']}: {i['char']} ({i['char_name']})")

    # --- Step 2: HTML 结构检查 ---
    print(f"\n[2/3] HTML Structure Check")
    print("-" * 40)
    html_issues = check_html(filepath)
    for i in html_issues:
        if i['severity'] == 'ERROR':
            total_errors += 1
        else:
            total_warnings += 1
    if not html_issues:
        print("  PASS")

    # --- Step 3: CSS 语法检查 ---
    print(f"\n[3/3] CSS Syntax Check")
    print("-" * 40)
    css_issues = check_css(filepath)
    for i in css_issues:
        if i['severity'] == 'ERROR':
            total_errors += 1
        elif i['severity'] == 'WARNING':
            total_warnings += 1

    if not css_issues:
        print("  PASS")

    # --- 汇总 ---
    print("\n" + "=" * 60)
    if total_errors == 0 and total_warnings == 0:
        print("  ALL CHECKS PASSED")
    else:
        print(f"  SUMMARY: {total_errors} error(s), {total_warnings} warning(s)")
        if total_errors > 0:
            print("  ACTION REQUIRED: Fix errors before proceeding")
    print("=" * 60)

    return total_errors


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    filepath = sys.argv[1]
    fix = '--fix' in sys.argv

    if not Path(filepath).exists():
        print(f"ERROR: File not found: {filepath}")
        sys.exit(1)

    errors = run_all(filepath, fix=fix)
    sys.exit(1 if errors > 0 else 0)


if __name__ == '__main__':
    main()
