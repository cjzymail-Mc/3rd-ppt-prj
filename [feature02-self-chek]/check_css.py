"""
CSS 语法检查工具
检测 <style> 标签内的大括号平衡、常见语法错误
"""

import re
import sys
from pathlib import Path


def check_css(filepath: str) -> list[dict]:
    """检查 HTML 文件中 <style> 块的 CSS 语法

    Args:
        filepath: HTML 文件路径

    Returns:
        问题列表 [{severity, message, detail}]
    """
    content = Path(filepath).read_text(encoding='utf-8')
    issues = []

    # 提取所有 <style> 块
    style_blocks = re.findall(r'<style[^>]*>(.*?)</style>', content, re.DOTALL)
    if not style_blocks:
        issues.append({
            'severity': 'WARNING',
            'message': '未找到 <style> 标签',
            'detail': '文件中没有内嵌 CSS',
        })
        return issues

    css_all = '\n'.join(style_blocks)

    # --- 1. 大括号平衡 ---
    opens = css_all.count('{')
    closes = css_all.count('}')
    if opens != closes:
        issues.append({
            'severity': 'ERROR',
            'message': 'CSS 大括号不平衡',
            'detail': f'{{ 出现 {opens} 次, }} 出现 {closes} 次 (差异: {opens - closes})',
        })

    # --- 2. 常见语法错误 ---
    # 缺少分号（属性值后直接跟 } 且非注释）
    missing_semi = re.findall(r'[a-zA-Z0-9%)\s"\']\s*\n\s*\}', css_all)
    # 过滤掉注释内的情况
    for match in missing_semi:
        stripped = match.strip()
        if not stripped.startswith('/*') and not stripped.startswith('//'):
            issues.append({
                'severity': 'WARNING',
                'message': '可能缺少分号',
                'detail': f'在 "}} " 前发现: ...{match.strip()[:60]}',
            })

    # --- 3. 空规则检测 ---
    empty_rules = re.findall(r'([^\{\}]+)\{\s*\}', css_all)
    for rule in empty_rules:
        selector = rule.strip()
        if selector and not selector.startswith('/*'):
            issues.append({
                'severity': 'INFO',
                'message': f'空 CSS 规则: {selector[:50]}',
                'detail': '选择器存在但无任何属性声明',
            })

    # --- 4. 重复选择器检测 ---
    selectors = re.findall(r'([^\{\}/][^\{]*?)\s*\{', css_all)
    selector_counts = {}
    for s in selectors:
        s = s.strip()
        if s and not s.startswith('@') and not s.startswith('/*'):
            selector_counts[s] = selector_counts.get(s, 0) + 1
    for sel, count in selector_counts.items():
        if count > 1:
            issues.append({
                'severity': 'INFO',
                'message': f'重复选择器: {sel[:50]}',
                'detail': f'出现 {count} 次，可能导致样式覆盖',
            })

    # --- 5. class 引用一致性（CSS 中定义的 class vs HTML 中使用的 class）---
    css_classes = set(re.findall(r'\.([a-zA-Z_][\w-]*)', css_all))
    html_classes = set(re.findall(r'class="[^"]*\b([a-zA-Z_][\w-]*)\b', content))
    # 只报告 HTML 中使用但 CSS 中未定义的（可能拼写错误）
    undefined = html_classes - css_classes
    # 过滤掉常见的外部框架 class
    framework_classes = {'container', 'row', 'col', 'btn', 'active', 'hidden', 'show', 'fade'}
    undefined -= framework_classes
    if undefined and len(undefined) <= 10:  # 超过10个可能是误报
        issues.append({
            'severity': 'INFO',
            'message': f'HTML 中使用但 CSS 中未定义的 class',
            'detail': ', '.join(sorted(undefined)[:10]),
        })

    # --- 汇总 ---
    if not issues:
        print(f"PASS: CSS syntax OK ({opens} rules, {len(css_classes)} classes defined)")
    else:
        errors = sum(1 for i in issues if i['severity'] == 'ERROR')
        warnings = sum(1 for i in issues if i['severity'] == 'WARNING')
        infos = sum(1 for i in issues if i['severity'] == 'INFO')
        print(f"RESULT: {errors} error(s), {warnings} warning(s), {infos} info(s)\n")
        for i in issues:
            print(f"  [{i['severity']}] {i['message']}")
            print(f"    {i['detail']}\n")

    return issues


def main():
    if len(sys.argv) < 2:
        print("Usage: python check_css.py <html_file>")
        sys.exit(1)

    filepath = sys.argv[1]
    if not Path(filepath).exists():
        print(f"ERROR: File not found: {filepath}")
        sys.exit(1)

    issues = check_css(filepath)
    has_errors = any(i['severity'] == 'ERROR' for i in issues)
    sys.exit(1 if has_errors else 0)


if __name__ == '__main__':
    main()
