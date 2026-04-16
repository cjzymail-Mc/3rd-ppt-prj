"""
HTML 编码检查工具
检测 HTML 标签内的非 ASCII 引号（弯引号 U+201C/U+201D 等）
这是导致 CSS 样式失效的最常见 Bug
"""

import re
import sys
from pathlib import Path


# 需要检测的非 ASCII 引号字符
CURLY_QUOTES = {
    '\u201c': 'LEFT DOUBLE QUOTATION MARK',
    '\u201d': 'RIGHT DOUBLE QUOTATION MARK',
    '\u2018': 'LEFT SINGLE QUOTATION MARK',
    '\u2019': 'RIGHT SINGLE QUOTATION MARK',
    '\u00ab': 'LEFT-POINTING DOUBLE ANGLE QUOTATION MARK',
    '\u00bb': 'RIGHT-POINTING DOUBLE ANGLE QUOTATION MARK',
}


def check_encoding(filepath: str, fix: bool = False) -> list[dict]:
    """检查 HTML 文件中标签内的非 ASCII 引号

    Args:
        filepath: HTML 文件路径
        fix: 是否自动修复（替换为 ASCII 直引号）

    Returns:
        问题列表 [{line, col, char, char_name, tag_snippet}]
    """
    path = Path(filepath)
    content = path.read_text(encoding='utf-8')
    lines = content.split('\n')
    issues = []

    # 逐行扫描 HTML 标签
    for line_num, line in enumerate(lines, 1):
        for tag_match in re.finditer(r'<[^>]+>', line):
            tag = tag_match.group()
            tag_start = tag_match.start()
            for char, name in CURLY_QUOTES.items():
                pos = tag.find(char)
                while pos != -1:
                    issues.append({
                        'line': line_num,
                        'col': tag_start + pos + 1,
                        'char': repr(char),
                        'char_name': name,
                        'tag_snippet': tag[:80] + ('...' if len(tag) > 80 else ''),
                    })
                    pos = tag.find(char, pos + 1)

    # 自动修复
    if fix and issues:
        def fix_tag_quotes(m):
            tag = m.group(0)
            for char in CURLY_QUOTES:
                tag = tag.replace(char, '"' if char in '\u201c\u201d\u00ab\u00bb' else "'")
            return tag

        fixed = re.sub(r'<[^>]+>', fix_tag_quotes, content)
        path.write_text(fixed, encoding='utf-8')

    return issues


def main():
    if len(sys.argv) < 2:
        print("Usage: python check_encoding.py <html_file> [--fix]")
        sys.exit(1)

    filepath = sys.argv[1]
    fix = '--fix' in sys.argv

    if not Path(filepath).exists():
        print(f"ERROR: File not found: {filepath}")
        sys.exit(1)

    issues = check_encoding(filepath, fix=fix)

    if not issues:
        print(f"PASS: No curly quotes found in HTML tags")
    else:
        print(f"{'FIXED' if fix else 'FAIL'}: Found {len(issues)} curly quote(s) in HTML tags\n")
        for i in issues:
            print(f"  Line {i['line']}, Col {i['col']}: {i['char']} ({i['char_name']})")
            print(f"    Tag: {i['tag_snippet']}\n")

    sys.exit(0 if not issues or fix else 1)


if __name__ == '__main__':
    main()
