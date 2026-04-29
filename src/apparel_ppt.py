#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""apparel_ppt.py — apparel 服装测试模板 PPT 评测页生成（零 pipeline 依赖）.

模板：template/empty and standard-apparel.pptx 的 slide 2，已合并到
      src/Template 2.1.pptx 第 18 页（slide 19 是该模板的 standard 页，
      生成时 Clone slide 18 即可，与 Pipeline 产物对齐）。

模板特点（与 yzr/zxh 不同）：
  - 数据为 5 分制（5-scale），非 10 分制
  - 4 个分类（版型 / 面料 / 吸湿排汗 / 速干），每类一个独立 Chart
    + 一个圆形评分 (Oval) + 一个标题 TextBox
  - 优点 / 缺点 TextBox 由 GPT 生成
  - 一个受试者信息 TextBox（5 名样本）由 GPT 生成

公开 API: make_apparel_slide()
"""

from __future__ import annotations

import re
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Tuple

# 直接运行时（python src/apparel_ppt.py），将项目根目录加入 sys.path
if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

# src 内部依赖：GPT_5 + overlay helpers
GPT_5 = None
show_gpt_waiting_overlay = None
remove_gpt_waiting_overlay = None
try:
    from .Function_030 import GPT_5, show_gpt_waiting_overlay, remove_gpt_waiting_overlay  # type: ignore
except Exception:
    try:
        from src.Function_030 import GPT_5, show_gpt_waiting_overlay, remove_gpt_waiting_overlay  # type: ignore
    except Exception:
        try:
            from Function_030 import GPT_5, show_gpt_waiting_overlay, remove_gpt_waiting_overlay  # type: ignore
        except Exception:
            GPT_5 = None

# ---------------------------------------------------------------------------
# 共享纯数据工具（fix2）：常量 / 评分 / 文本裁剪 / Excel 列工具 / 通用 COM 写入
# 影响视觉输出的函数（_build_rich_prompt / _apply_keyword_color 配置）保留在
# 本文件内，以保证 per-template 微调能力。
# ---------------------------------------------------------------------------
_shared_import_ok = False
try:
    from ._ppt_shared import (  # type: ignore
        _RED, _BLUE, _BLACK,
        _ADVANTAGE_MARKERS, _DISADVANTAGE_MARKERS,
        _find_col, _classify_columns, _col_values,
        _xlwings_to_rows,
        clamp_text,
        _NAME_KEYWORDS, _WEIGHT_KEYWORDS,
        _com_get, _write_text,
        _apply_keyword_color,
    )
    _shared_import_ok = True
except Exception:
    pass
if not _shared_import_ok:
    try:
        from src._ppt_shared import (  # type: ignore
            _RED, _BLUE, _BLACK,
            _ADVANTAGE_MARKERS, _DISADVANTAGE_MARKERS,
            _find_col, _classify_columns, _col_values,
            _xlwings_to_rows,
            clamp_text,
            _NAME_KEYWORDS, _WEIGHT_KEYWORDS,
            _com_get, _write_text,
            _apply_keyword_color,
        )
    except Exception:
        from _ppt_shared import (  # type: ignore
            _RED, _BLUE, _BLACK,
            _ADVANTAGE_MARKERS, _DISADVANTAGE_MARKERS,
            _find_col, _classify_columns, _col_values,
            _xlwings_to_rows,
            clamp_text,
            _NAME_KEYWORDS, _WEIGHT_KEYWORDS,
            _com_get, _write_text,
            _apply_keyword_color,
        )

# Default GPT model
_MODEL = "openai/gpt-5.4"

# Clipboard copy-paste COM buffer (seconds)
_COPY_PASTE_DELAY = 1.5

# apparel 标准模板所在页（合并后的 Template 2.1.pptx 第 19 页）
# 备注：源模板 empty and standard-apparel.pptx 含 2 页：slide 1 = blank（7 shapes），
# slide 2 = standard（22 shapes）。
# 合并到 Template 2.1.pptx 末尾后总 19 页：slide 18 = blank、slide 19 = standard。
# 实测 Slide 18 是 blank（7 个零散 shape，含 Straight Connector 等占位），
# Slide 19 才是 18 个 shape 齐全的 standard 页。
_TEMPLATE_SLIDE = 19

# apparel 是 5 分制（与 yzr 10 分制不同）
_SCALE_MAX = 5

# ---------------------------------------------------------------------------
# 写作语调参考语料（13 条，"问题 // 优势"对照式）
# 仅供 GPT 学习行文风格和信息密度，不参与内容生成 —— prompt 里已强约束
# "参考文本仅作语调参考，不必复制其分类结构"，防止 GPT 抄分类。
# 注入位置：_build_rich_prompt 的 style_anchor 槽（≠ fallback，职责不同）
# ---------------------------------------------------------------------------
_STYLE_REFERENCE_CORPUS = (
    "版型：领口缺乏弹性，穿脱有阻碍 // 但整体版型宽松，上身无束缚，跑动中很自在\n"
    "面料手感：面料偏软塌，弹性不佳，体前屈时后背紧绷 // 但亲肤顺滑，不摩擦皮肤，长时间穿着依然舒适\n"
    "轻量化：出汗后局部会轻微变重 // 平时几乎感觉不到重量，起步轻盈，适合长距离\n"
    "舒适度：腋下接缝处长时间跑动略有磨感 // 其他部位无刺痒、无勒痕，整体舒适度较高\n"
    "运动限制：腰背部弹性余量不足，高抬腿时略有牵扯 // 日常慢跑和摆臂不受限，下蹲、跨步均可完成\n"
    "吸湿性能：胸口大汗时表面略有湿润感 // 吸汗非常快，汗液不会在身体停留，始终保持表层不滴流\n"
    "速干性能：湿态下拧干需约20分钟才全干 // 跑动中风吹即干，运动结束5分钟内基本恢复干爽\n"
    "排汗性能：高强度间歇训练后背部有短暂闷感 // 全程不贴身、不黏腻，贴身面保持干爽清凉\n"
    "接缝平整度：腋下与肩部接缝有轻微凸起 // 平缝工艺整平度高，长距离跑下来未出现明显摩擦红痕\n"
    "反光细节：侧身反光面积偏小，能见度一般 // 前后均有反光条，夜跑时正面和背面可视性良好\n"
    "口袋稳定性：缺少耳袋，钥匙等小物件无处安放 // 后腰拉链袋稳固不晃，手机放入后跑全马不跳动\n"
    "湿态脱衣便利性：袖口偏窄，湿臂取下略费劲 // 领口弹性适当，大汗后单手仍可较快脱下\n"
    "摩擦红痕：胸前logo胶印处有轻微痕迹 // 其余接缝与标签无擦伤，皮肤整体光洁无勒痕"
)

# ---------------------------------------------------------------------------
# 4 个分类的列名前缀（用于按分类切片 Chart 数据）
# 与模板 4 个 TextBox 标题对应：版型 / 面料 / 吸湿排汗 / 速干
# ---------------------------------------------------------------------------
_CATEGORY_KEYWORDS: Dict[str, List[str]] = {
    "版型":     ["版型"],
    "面料":     ["面料", "亲肤", "轻量"],
    "吸湿排汗": ["吸湿", "排汗"],
    "速干":     ["速干"],
}

# ---------------------------------------------------------------------------
# Hardcoded shape specs (apparel 标准模板在 Template 2.1.pptx 第 18 页)
# 来源：pipeline-progress/02-shape_analysis_map.json + 02-readability_budget.json
# 决策点（Developer 移植阶段做的判断，覆盖部分 Pipeline 默认）：
#   1. 4 个 Chart：Pipeline 给的 mean_extraction 把所有列堆一起，视觉上 4 chart
#      内容相同。改为按分类切片 → mean_extraction_filtered + category 参数。
#   2. 4 个 Oval：模板设计上是装饰性虚线圆圈，不放文字（分数另由分类标题
#      旁显示）。统一 skip，避免 score 文本残留在圈内。
#      （历史上曾尝试 score_category_mean 写整体均值，已废弃。）
#   3. 6 个 template_direct title：Clone 自然继承，全部 skip。
# ---------------------------------------------------------------------------
APPAREL_SHAPES = [
    # ---- 4 类分组：装饰圆圈（skip）+ 标题（skip）+ Chart ----
    {"name": "Oval 3",      "strategy": "skip"},  # 装饰虚线圆，不放文字
    {"name": "TextBox 6",   "strategy": "skip"},  # template_direct: 版型
    {"name": "Chart 12",    "strategy": "mean_extraction_filtered",
     "params": {"category": "版型"}},

    {"name": "Oval 13",     "strategy": "skip"},  # 装饰虚线圆，不放文字
    {"name": "TextBox 14",  "strategy": "skip"},  # template_direct: 面料
    {"name": "Chart 15",    "strategy": "mean_extraction_filtered",
     "params": {"category": "面料"}},

    {"name": "Oval 16",     "strategy": "skip"},  # 装饰虚线圆，不放文字
    {"name": "TextBox 17",  "strategy": "skip"},  # template_direct: 吸湿排汗
    {"name": "Chart 18",    "strategy": "mean_extraction_filtered",
     "params": {"category": "吸湿排汗"}},

    {"name": "Oval 19",     "strategy": "skip"},  # 装饰虚线圆，不放文字
    {"name": "TextBox 20",  "strategy": "skip"},  # template_direct: 速干
    {"name": "Chart 21",    "strategy": "mean_extraction_filtered",
     "params": {"category": "速干"}},

    # ---- 优点 / 缺点 / 受试者信息 ----
    {"name": "TextBox 23",  "strategy": "skip"},  # template_direct: "优点 strengths"
    {"name": "TextBox 24",  "strategy": "gpt_respondent_info",
     "budget": {"max_chars": 102, "max_lines": 5}},
    {"name": "TextBox 26",  "strategy": "skip"},  # template_direct: "缺点 drawbacks"
    {"name": "TextBox 8",   "strategy": "gpt_prompted",
     "params": {"source": "补充说明", "filter": "优点"},
     "budget": {"max_chars": 99, "max_lines": 3}},
    {"name": "TextBox 22",  "strategy": "gpt_prompted",
     "params": {"source": "补充说明", "filter": "缺点"},
     "budget": {"max_chars": 93, "max_lines": 2}},
    {"name": "Rectangle 25", "strategy": "skip"},  # template_direct: "I  <面料信息>"
]


# ---------------------------------------------------------------------------
# Tiny helper utilities (self-contained)
# ---------------------------------------------------------------------------

def _safe_text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _numeric(v: Any):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# apparel 专用：按分类切片提取均值
# ---------------------------------------------------------------------------

def _extract_means_for_category(rows: List[List[Any]],
                                category: str) -> List[Tuple[str, float]]:
    """提取某一分类（版型/面料/吸湿排汗/速干）下所有评分列的均值。

    关键过滤：
      - 列名包含 _CATEGORY_KEYWORDS[category] 中任一关键词
      - 列名不含其他类别的关键词（防误匹配，如"面料评价（文字描述）"被排除）
      - 该列 >70% 数据是 0~5 范围数值
    """
    if not rows or len(rows) < 2:
        return []
    headers = [_safe_text(h) for h in rows[0]]
    data = rows[1:]

    own_kws = _CATEGORY_KEYWORDS.get(category, [])
    if not own_kws:
        return []

    # 排除的"伪评分列"标记 ——
    # 注意："评价"二字在 apparel 列名里有歧义：
    #   - 评分列示例："1、【整体】版型评分（得分越高评价越合身，下同）" → 是评分列
    #   - 文本列示例："版型评价（文字描述）" → 是描述列
    # 所以不能简单用"评价"排除；用"文字描述/留言/建议"更精确。
    # 同时通过下面"列内 >50% 数据为 0~10 数值"的硬过滤兜底。
    text_markers = ["文字描述", "描述）", "留言", "提交人", "提交时间", "更新时间", "标题"]

    out: List[Tuple[str, float]] = []
    for ci, h in enumerate(headers):
        if not any(kw in h for kw in own_kws):
            continue
        # 排除"<分类>评价（文字描述）"等纯文本反馈列
        if any(m in h for m in text_markers):
            continue

        vals = []
        for row in data:
            if ci >= len(row):
                continue
            n = _numeric(row[ci])
            if n is not None and 0 <= n <= 10:
                vals.append(n)
        if not vals:
            continue
        if len(vals) / max(len(data), 1) < 0.5:
            continue

        mean_val = sum(vals) / len(vals)
        # 标签清理：apparel 模板列名格式形如
        #   "1、【腰围】版型评分（得分越高评价越合身，下同）"
        # chart 已按 category 分组（左下角有"版型/面料/..."标题），
        # value 标签只需保留差异点 → 取 【...】 内的字（如"腰围"）。
        m = re.search(r'【([^】]+)】', h)
        if m:
            clean_label = m.group(1).strip()
        else:
            # 兜底：无 【】 时退回旧清理逻辑
            clean_label = re.sub(r'^\d+[、.]\s*', '', h)
            clean_label = re.sub(r'（[^）]*）', '', clean_label).strip()
            clean_label = clean_label.replace("评分", "").strip()
        if not clean_label:
            clean_label = h[:8]
        out.append((clean_label, round(mean_val, 2)))

    return out


def _category_overall_mean(rows: List[List[Any]], category: str) -> float:
    """计算某一分类的整体均值（用于圆环评分）。

    取该分类下所有评分列的均值的均值。返回 0~5 的浮点数，无数据返回 0.0。
    """
    means = _extract_means_for_category(rows, category)
    if not means:
        return 0.0
    return round(sum(v for _, v in means) / len(means), 1)


# ---------------------------------------------------------------------------
# GPT prompt helpers (dynamic column version)
# ---------------------------------------------------------------------------

def _build_respondent_block(rows: List[List[Any]]) -> Tuple[str, int]:
    """Build a per-respondent data block for inclusion in GPT prompt.

    复刻自 yzr_ppt._build_respondent_block，但跑步服装问卷用"跑者姓名"
    作为名称列，已被 _NAME_KEYWORDS 的 "姓名" 关键词覆盖。
    """
    if not rows or len(rows) < 2:
        return "（无数据）", 0

    headers = [_safe_text(h) for h in rows[0]]
    n = len(rows) - 1

    name_col = _find_col(headers, _NAME_KEYWORDS)
    weight_col = _find_col(headers, _WEIGHT_KEYWORDS)
    height_col = _find_col(headers, ["身高", "height"])
    score_cols, text_cols = _classify_columns(headers, rows)

    blocks = []
    for i, row in enumerate(rows[1:], 1):
        fd: Dict[str, str] = {}
        for j, h in enumerate(headers):
            if j < len(row):
                fd[h] = _safe_text(row[j])

        name = fd.get(name_col, f"受访者{i}") if name_col else f"受访者{i}"
        weight = fd.get(weight_col, "") if weight_col else ""
        height = fd.get(height_col, "") if height_col else ""

        scores = ", ".join(
            f"{h.split('（')[0] if '（' in h else h}={fd[h]}"
            for h in score_cols if fd.get(h)
        )
        feedbacks = "\n  ".join(
            f"{h.split('（')[0] if '（' in h else h}: {fd[h]}"
            for h in text_cols if fd.get(h)
        )

        info_parts = []
        if height:
            info_parts.append(f"身高:{height}CM")
        if weight:
            info_parts.append(f"体重:{weight}KG")
        info_str = ("  " + " / ".join(info_parts)) if info_parts else ""
        parts = [f"【受访者{i}】{name}{info_str}"]
        if scores:
            parts.append(f"  各项评分: {scores}")
        if feedbacks:
            parts.append(f"  试穿反馈:\n  {feedbacks}")
        blocks.append("\n".join(parts))

    return "\n\n".join(blocks), n


# prompt_src:  pipeline-progress/02-prompt_specs.json (TextBox 8 / TextBox 22)
# synced_at:   2026-04-28
# synced_by:   Developer（apparel 模板移植时从 Pipeline 产物提取）
# 定制说明：   apparel 是服装试穿（不是篮球鞋），所以 task 描述统一改用"这件服装"；
#             5 分制（与 yzr 10 分制不同），prompt 里说明分值范围。
#             关键词染色仍用 【】 标记 → _apply_keyword_color 处理。
def _build_rich_prompt(
    budget: dict,
    rows: List[List[Any]],
    focus: str = "",
    content_source: str = "补充说明",
    style_anchor: str = "",
) -> str:
    """Build GPT prompt for gpt_prompted shapes (apparel 服装试穿专用).

    focus: '优点' or '缺点' → free-form summarization mode.
    """
    respondent_block, n = _build_respondent_block(rows)
    max_chars = budget.get("max_chars", 100)
    max_lines = budget.get("max_lines", 3)

    if focus:
        task_line = (
            f"请从{n}名测试者的实际反馈中，自由归纳这件服装的【{focus}】。\n"
            f"根据实际反馈内容自行决定分段维度（如版型/面料/吸湿排汗/速干），\n"
            f"在每条结论后注明（X/{n}）表示几分之几的测试者有此体验。\n"
            f"每段结论中，请将最核心的1-2个关键词用【】括起来（仅括词本身，不含标点），"
            f"例如：【版型】肩部偏松（2/{n}）。这些关键词后续会自动高亮显示。"
        )
        format_note = "- 参考文本仅作语调参考，不必复制其分类结构\n"
    else:
        task_line = (
            f"下面是{n}名测试者对这件服装的原始试穿反馈，"
            f"请帮我按分类汇总其中的【{content_source}】。\n"
            f"在每条结论后注明（X/{n}）表示几分之几的测试者有此反馈。\n"
            f"每段结论中，请将最核心的1-2个关键词用【】括起来（仅括词本身，不含标点）。"
        )
        format_note = "- 严格按照参考文本的格式、语调、陈述方式\n"

    return (
        f"【参考文本（参考语调和信息密度）】\n{style_anchor}\n\n"
        f"【你的任务】\n{task_line}\n\n"
        f"注意：\n"
        f"- 评分均为 5 分制（1=最差，5=最佳）\n"
        f"- 你只能分析已有数据，不能推测或编造\n"
        f"- 直接给出结论，不要展示分析过程\n"
        f"- {format_note}"
        f"- 总字数控制在{max_chars}字左右，不超过{max_lines}行\n"
        f"- 结论中请自然融入：'样本'（如'本次{n}名样本'）、'反馈'（如'样本反馈'）、'建议'（末尾给出改进建议）\n\n"
        f"【{n}名测试者原始反馈】\n{respondent_block}\n\n"
        f"直接输出结论，不需要任何前言。"
    )


# prompt_src:  pipeline-progress/02-prompt_specs.json (TextBox 24)
# synced_at:   2026-04-28
# synced_by:   Developer（受试者信息块独立 prompt，对应模板 A/B/C/D 列表）
# 标准成年人身高 / 体重 / BMI 的合理范围（apparel 测试人员均为标准身材，
# 用于识别填写时的常见单位混淆）。
_HEIGHT_CM_MAX = 210      # cm 上限
_WEIGHT_KG_MAX = 110      # kg 上限（>110 必为误填"斤"，先粗修一轮）
_BMI_OK = (16.0, 32.0)    # 合理 BMI 区间（含略瘦/略胖容差）


def _normalize_height_cm(raw: Any) -> str:
    """识别 m / cm 单位混淆并归一化为 cm 整数字符串。

    规则：
      - 数值 < 3   → 视为 m，×100 转 cm（1.65 → 165；1 → 100）
      - 数值 ≥ 3   → 视为 cm，原样保留（含异常值，留给上层定位）
      - 非数值     → 原样返回
    """
    n = _numeric(raw)
    if n is None:
        return _safe_text(raw)
    if n < 3:
        n = n * 100
    return str(int(round(n)))


def _normalize_weight_kg(raw: Any) -> str:
    """识别 kg / 斤 单位混淆并归一化为 kg 整数字符串（粗修阶段）。

    规则：
      - 数值 > 110 → 视为斤（1 kg = 2 斤），÷2 转 kg（130 → 65）
      - 数值 ≤ 110 → 暂保留（细修交给 _cross_validate_bmi）
      - 非数值     → 原样返回
    """
    n = _numeric(raw)
    if n is None:
        return _safe_text(raw)
    if n > _WEIGHT_KG_MAX:
        n = n / 2
    return str(int(round(n)))


def _cross_validate_bmi(h_str: str, w_str: str) -> Tuple[str, str]:
    """BMI 交叉验证：粗修后若 BMI 越界，再试体重 ÷2 是否落回合理区间。

    覆盖场景：标准身材测试者填了 100 kg —— 单看体重 ≤110 不触发粗修，
    但与 160 cm 组合 BMI=39 显然异常；÷2=50 kg 后 BMI=19.5 正常 → 采纳。

    返回 (height_cm_str, weight_kg_str)。任一非数值时原样返回。
    """
    try:
        h = float(h_str)
        w = float(w_str)
        if h <= 0 or w <= 0:
            return h_str, w_str
        bmi = w / (h / 100) ** 2
        if _BMI_OK[0] <= bmi <= _BMI_OK[1]:
            return h_str, w_str
        w2 = w / 2
        bmi2 = w2 / (h / 100) ** 2
        if _BMI_OK[0] <= bmi2 <= _BMI_OK[1]:
            return h_str, str(int(round(w2)))
    except Exception:
        pass
    return h_str, w_str


def _normalize_person(raw_h: Any, raw_w: Any) -> Tuple[str, str]:
    """联合归一化：身高/体重各自粗修 → BMI 交叉验证细修。"""
    return _cross_validate_bmi(
        _normalize_height_cm(raw_h),
        _normalize_weight_kg(raw_w),
    )


def _build_respondent_info_prompt(
    budget: dict,
    rows: List[List[Any]],
    style_anchor: str,
) -> str:
    """Build GPT prompt for TextBox 24 (受试者信息 Information).

    输出格式与模板原文对齐（A: 160CM/50KG 这种格式），用真实数据填充。
    """
    headers = [_safe_text(h) for h in rows[0]] if rows else []
    name_col = _find_col(headers, _NAME_KEYWORDS)
    weight_col = _find_col(headers, _WEIGHT_KEYWORDS)
    height_col = _find_col(headers, ["身高", "height"])

    info_lines = []
    for i, row in enumerate(rows[1:] if rows else [], 1):
        fd = {}
        for j, h in enumerate(headers):
            if j < len(row):
                fd[h] = _safe_text(row[j])
        name = fd.get(name_col, f"测试者{i}") if name_col else f"测试者{i}"
        raw_h = fd.get(height_col, "") if height_col else ""
        raw_w = fd.get(weight_col, "") if weight_col else ""
        height, weight = _normalize_person(raw_h, raw_w)
        info_lines.append(f"  {name}：身高 {height}CM，体重 {weight}KG")
    info_block = "\n".join(info_lines) if info_lines else "（无数据）"

    n = max(0, len(rows) - 1)
    max_chars = budget.get("max_chars", 102)
    max_lines = budget.get("max_lines", 5)

    return (
        f"【参考文本（参考排版格式和信息密度）】\n{style_anchor}\n\n"
        f"【你的任务】\n"
        f"以下是 {n} 名测试者的身高体重数据，请按参考文本的格式整理输出。\n"
        f"格式要求：第一行为标题『受试者信息 Information』，"
        f"接下来每行一名测试者，使用字母编号 A/B/C/D...，格式如 A: 160CM / 50 KG。\n\n"
        f"【测试者数据】\n{info_block}\n\n"
        f"注意：\n"
        f"- 严格按照参考文本的格式输出，不要添加额外分析\n"
        f"- 编号顺序按数据顺序，如 A=第1名测试者\n"
        f"- 总字数控制在{max_chars}字左右，不超过{max_lines}行\n"
        f"- 直接输出结果，不需要任何前言。"
    )


def _call_gpt(prompt: str, fallback: str, enabled: bool, model: str) -> str:
    """Call GPT_5 if enabled and available; return fallback otherwise."""
    if not enabled or GPT_5 is None:
        return fallback
    try:
        result = _safe_text(GPT_5(prompt, model))
        if result:
            return result
    except Exception:
        pass
    return fallback


def _build_respondent_info_fallback(rows: List[List[Any]]) -> str:
    """构造 TextBox 24 的 fallback —— 不调用 GPT 时直接生成受试者信息块。"""
    if not rows or len(rows) < 2:
        return "受试者信息 Information\r（无数据）"
    headers = [_safe_text(h) for h in rows[0]]
    weight_col = _find_col(headers, _WEIGHT_KEYWORDS)
    height_col = _find_col(headers, ["身高", "height"])

    lines = ["受试者信息 Information"]
    for i, row in enumerate(rows[1:], 1):
        fd = {}
        for j, h in enumerate(headers):
            if j < len(row):
                fd[h] = _safe_text(row[j])
        # 单位归一化（粗修 m→cm / 斤→kg + BMI 交叉验证细修）
        raw_h = fd.get(height_col, "?") if height_col else "?"
        raw_w = fd.get(weight_col, "?") if weight_col else "?"
        height, weight = _normalize_person(raw_h, raw_w)
        letter = chr(ord("A") + i - 1) if i <= 26 else str(i)
        lines.append(f"{letter}: {height}CM / {weight} KG")
    return "\r".join(lines)


# ---------------------------------------------------------------------------
# Content builder — routes by strategy
# ---------------------------------------------------------------------------

def _build_content(spec: dict, rows: List[List[Any]],
                   gpt_enabled: bool, model: str) -> str:
    """Build the text/data content for one shape spec."""
    strategy = spec["strategy"]
    params = spec.get("params", {})
    budget = spec.get("budget", {"max_chars": 100, "max_lines": 4})

    if strategy == "skip":
        return ""

    if strategy == "score_category_mean":
        cat = params.get("category", "")
        mean_val = _category_overall_mean(rows, cat)
        if mean_val <= 0:
            return ""
        # 5 分制评分，保留 1 位小数（如 4.5）
        return f"{mean_val:.1f}"

    if strategy == "mean_extraction_filtered":
        cat = params.get("category", "")
        means = _extract_means_for_category(rows, cat)
        if not means:
            return f"{cat}:0"
        return "\n".join(f"{k}:{v:.2f}" for k, v in means[:8])

    if strategy == "gpt_respondent_info":
        # TextBox 24：受试者信息（A/B/C/D 列表式）
        style_anchor = (
            "受试者信息 Information\r"
            "A: 160CM / 50 KG\r"
            "B: 165CM / 52 KG\r"
            "C: 162CM / 58 KG\r"
            "D: 168CM / 60 KG"
        )
        fallback = _build_respondent_info_fallback(rows)
        prompt = _build_respondent_info_prompt(budget, rows, style_anchor)
        result = _call_gpt(prompt, fallback, gpt_enabled, model)
        max_chars = budget.get("max_chars", 102)
        max_lines = budget.get("max_lines", 5)
        return clamp_text(result, max_chars, max_lines)

    if strategy == "gpt_prompted":
        focus = params.get("filter", "")
        src = params.get("source", "补充说明")
        # fallback 引用模板原文（pipeline-progress/01-shape_detail_com.json）
        fallback_map = {
            "优点": (
                "轻薄透气，舒适不闷热（4/4）\r"
                "吸湿排汗性能好，汗湿后衣服不会黏在身上（4/4）；\r"
                "速干性能极好（4/4）大部分时间都很干爽（1/4）"
            ),
            "缺点": (
                "肩部挖空偏小，会勒腋下，影响摆臂（2/4）；\r"
                "后背开叉偏高，跑动过程后背会漏出小三角（1/4）；"
            ),
        }
        fallback = fallback_map.get(focus, "样本反馈总体稳定，核心指标表现均衡。")
        prompt = _build_rich_prompt(
            budget, rows, focus=focus,
            content_source=src, style_anchor=_STYLE_REFERENCE_CORPUS,
        )
        result = _call_gpt(prompt, fallback, gpt_enabled, model)
        # 安全网：GPT 返回后强制 clamp 到 budget 范围内
        max_chars = budget.get("max_chars", 100)
        max_lines = budget.get("max_lines", 3)
        return clamp_text(result, max_chars, max_lines)

    return ""


# ---------------------------------------------------------------------------
# apparel 专用 chart helper —— xlwings 新建 + OLE 粘贴（fix4 路线）
# 与 yzr 不同点：
#   - 2D bar_clustered（非 3D，apparel 模板视觉是平面条形）
#   - 5 分制量程 0~6（5 分制 → max + 1，避免数据标签压住 bar 末端，硬规则）
#   - 数据为 N 个细分指标（如版型分类下的 整体/衣领/袖口/胸围/腰围）
# ---------------------------------------------------------------------------

def _prepare_apparel_chart_data(mc_sht, content: str, anchor_offset: int = 50):
    """Parse mean_extraction_filtered content and write to Excel as 2-column table.

    与 _prepare_yzr_chart_data 类似但用更大偏移避免冲突。
    每次调用使用不同的 anchor_offset，让 4 个 chart 各占一片区域。
    """
    import importlib

    try:
        fn030 = importlib.import_module("src.Function_030")
        origin = fn030.get_range(mc_sht)
    except Exception:
        try:
            fn030 = importlib.import_module("Function_030")
            origin = fn030.get_range(mc_sht)
        except Exception:
            origin = mc_sht.range("A1")

    try:
        rows_count = origin.api.CurrentRegion.Rows.Count
    except Exception:
        rows_count = 10

    anchor = origin.offset(row_offset=rows_count + anchor_offset, column_offset=0)

    parsed: List[Tuple[str, float]] = []
    for line in (content or "").splitlines():
        line = line.strip()
        if not line or ":" not in line:
            continue
        k, v = line.rsplit(":", 1)
        try:
            parsed.append((k.strip(), float(v.strip())))
        except Exception:
            parsed.append((k.strip(), 0.0))
    if not parsed:
        parsed = [("占位", 0.0)]

    table = [("指标", "均值")] + parsed
    anchor.value = table
    print(f"  [apparel-chart] 临时数据已写入：anchor=({anchor.row},{anchor.column})，{len(parsed)} 个指标")
    return anchor


def make_chart_for_apparel(mc_cell, mc_slide, Left, Top, Width, Height):
    """为 apparel 模板构建 2D 条形图，OLE 粘贴到 PPT。

    硬规则依赖：
      - `CutCopyMode = False`（断 OLE 热链接）
      - `MaximumScale = _SCALE_MAX + 1 = 6`（5 分制 → 6，规则 #bar-chart-max+1）
      - `Shapes.Paste()` 返回 ShapeRange，访问 .Chart 必须先 .Item(1)
    """
    import random
    import xlwings

    print("[apparel-chart] 开始 xlwings 建 2D 条形图 → OLE 粘贴")

    mc_sht = mc_cell.sheet
    try:
        mc_sht.select()
    except Exception:
        pass
    mc_cell.select()

    i0 = mc_cell.api.CurrentRegion.Row
    i = mc_cell.api.CurrentRegion.Rows.Count
    j0 = mc_cell.api.CurrentRegion.Column
    j = mc_cell.api.CurrentRegion.Columns.Count

    chart_left = mc_sht.cells(i0 + i - 2, j0 + 3).left
    chart_top = mc_sht.cells(i0 + i - 2, j0 + 3).top

    mc_chart1 = mc_sht.charts.add(chart_left, chart_top, width=Width, height=Height)
    # apparel 模板视觉是 2D 条形图
    try:
        mc_chart1.chart_type = "bar_clustered"
    except Exception:
        pass

    mc_chart1.set_source_data(
        mc_sht.range((i0, j0), (i0 + i - 1, j0 + j - 1))
    )

    # 隐藏图例 / 网格线
    mc_chart1.api[1].SetElement(100)
    mc_chart1.api[1].SetElement(328)

    # 固定数值轴量程：max = _SCALE_MAX + 1 = 6（硬规则 #bar-chart-max+1）
    _val_axis = mc_chart1.api[1].Axes(2)
    _val_axis.MinimumScaleIsAuto = False
    _val_axis.MaximumScaleIsAuto = False
    _val_axis.MinimumScale = 0
    _val_axis.MaximumScale = _SCALE_MAX + 1
    _val_axis.TickLabelPosition = -4142
    _val_axis.MajorTickMark = -4142
    _val_axis.MinorTickMark = -4142
    try:
        _val_axis.Format.Line.Visible = 0
    except Exception:
        pass
    print(f"  [apparel-chart] 坐标轴 0~{_SCALE_MAX + 1}，轴线/刻度/标签已隐藏")

    # 数据标签
    try:
        mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()
    except Exception:
        pass

    # 隐藏主标题（双调用，与 yzr-chart 一致）
    mc_chart1.api[1].SetElement(0)
    mc_chart1.api[1].SetElement(0)

    # OLE 复制
    mc_cell.select()
    mc_chart1.api[0].Copy()
    time.sleep(0.5 + random.random() * 0.3)

    mc_shape = mc_slide.Shapes.Paste()
    time.sleep(0.5)

    # 断 OLE 热链接（硬规则 #1）
    try:
        xlwings.apps.active.api.CutCopyMode = False
    except Exception:
        pass

    mc_shape.Left = Left
    mc_shape.Top = Top
    try:
        mc_shape.Width = Width
        mc_shape.Height = Height
    except Exception:
        pass

    # PPT 端再次隐藏主标题（ShapeRange.Chart 必须先 .Item(1)，硬规则）
    try:
        _shape_one = mc_shape.Item(1) if hasattr(mc_shape, "Item") else mc_shape
        _shape_one.Chart.HasTitle = False
        _shape_one.Chart.SetElement(0)
        print("  [apparel-chart] PPT 端主标题已隐藏")
    except Exception as _e:
        print(f"  [apparel-chart] PPT 端隐藏标题失败（{_e!r}）")

    print(f"  [apparel-chart] 已粘贴至 PPT（L={Left}, T={Top}, W={Width}, H={Height}）")
    return mc_chart1


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def make_apparel_slide(mc_sht, mc_ppt, mc_slide, sample_name: str,
                       mc_gpt: str = "n", mc_model: str = _MODEL):
    """Generate apparel evaluation slide.

    Clones Template 2.1.pptx 第 18 页 (apparel standard) 到末尾，
    然后按 APPAREL_SHAPES 写入。
    Returns the new slide object.
    """
    gpt_enabled = (mc_gpt == "y")
    print(f"\n[apparel] 开始生成评测页  sample={sample_name}  gpt={'开启' if gpt_enabled else '关闭'}")
    rows = _xlwings_to_rows(mc_sht)
    print(f"[apparel] 读取问卷数据：{len(rows)} 行（含标题行），{len(rows[0]) if rows else 0} 列")

    # === Clone pattern — identical to yzr/zxh ===
    X = mc_ppt.Slides.Count + 1
    print(f"[apparel] 克隆模板第 {_TEMPLATE_SLIDE} 页 → 新建第 {X} 页...")
    mc_ppt.Slides(_TEMPLATE_SLIDE).Copy()
    time.sleep(_COPY_PASTE_DELAY)
    new_slide = mc_ppt.Slides.Paste(X)
    time.sleep(1.0)

    # 显示等待遮罩（GPT 调用期间可见）
    _overlay = None
    if show_gpt_waiting_overlay is not None:
        try:
            _overlay = show_gpt_waiting_overlay(new_slide)
        except Exception:
            pass

    # === Per-shape content build and write ===
    print(f"[apparel] 开始逐 shape 写入，共 {len(APPAREL_SHAPES)} 个...")
    chart_anchor_offset = 50  # 每个 chart 数据块在 Excel 用不同偏移避免覆盖
    for spec in APPAREL_SHAPES:
        name     = spec["name"]
        strategy = spec["strategy"]

        if strategy == "skip":
            print(f"  [skip] {name} （template_direct，Clone 已继承）")
            continue

        # Find shape on the new slide
        shp = None
        try:
            shp = new_slide.Shapes(name)
        except Exception:
            pass
        if shp is None:
            print(f"  [未找到] {name}（模板中不存在此 shape，跳过）")
            continue

        print(f"  [处理] {name}  strategy={strategy}")

        # Build content
        content = _build_content(spec, rows, gpt_enabled, mc_model)

        # Route to correct writer
        if strategy == "mean_extraction_filtered":
            # fix4: 分发场景 chart 走从零制表路线
            try:
                L = float(_com_get(shp, "Left", 0))
                T = float(_com_get(shp, "Top", 0))
                W = float(_com_get(shp, "Width", 190))
                H = float(_com_get(shp, "Height", 100))
                shp.Delete()
            except Exception as _e:
                print(f"    [警告] 读取/删除模板 chart shape 失败: {_e}")
                continue

            # 写入临时数据 → 建图 → OLE 粘贴
            try:
                mc_cell = _prepare_apparel_chart_data(
                    mc_sht, content, anchor_offset=chart_anchor_offset,
                )
                chart_anchor_offset += len(content.splitlines()) + 5
            except Exception as _e:
                print(f"    [警告] 写入 Excel 临时数据失败: {_e}")
                continue

            try:
                _tmp_chart = make_chart_for_apparel(
                    mc_cell, new_slide, Left=L, Top=T, Width=W, Height=H,
                )
            except Exception as _e:
                print(f"    [警告] make_chart_for_apparel 失败: {_e}")
                continue

            # 清理：只删 chart，保留临时数据（与 yzr 一致，原因见 yzr_ppt 注释）
            _xl_app_api = None
            try:
                _xl_app_api = mc_sht.book.app.api
                _xl_app_api.DisplayAlerts = False
            except Exception:
                pass
            try:
                _tmp_chart.delete()
            except Exception:
                pass
            try:
                if _xl_app_api is not None:
                    _xl_app_api.DisplayAlerts = True
            except Exception:
                pass
        else:
            # 文本类 shape（score_category_mean / gpt_respondent_info / gpt_prompted）
            ok = _write_text(shp, content)
            if not ok:
                print(f"    [警告] _write_text 返回 False")
            if ok and strategy == "gpt_prompted":
                # 染色决策（developer.md 染色函数选用决策树）：
                # apparel 优点 / 缺点是 per-shape 单段语境（一个 shape 全是优点
                # 或全是缺点），符合"_apply_keyword_color"的 section context 用法。
                _apply_keyword_color(shp)

    # 工作完成，删除等待遮罩
    if _overlay is not None and remove_gpt_waiting_overlay is not None:
        try:
            remove_gpt_waiting_overlay(_overlay)
        except Exception:
            pass

    print(f"[apparel] 完成！新页在第 {new_slide.SlideIndex} 页")
    return new_slide


# ---------------------------------------------------------------------------
# 单独调试入口：连接已打开的 Excel + PPT，只跑 apparel 这一页
# 用法：python src/apparel_ppt.py
# 前置：1) Excel 已打开 template/source data-apparel.xlsx
#       2) PPT 已打开 src/Template 2.1.pptx（脚本会自动 Open，已打开则用现有）
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import os
    import win32com.client
    import xlwings

    _proj_root = str(Path(__file__).resolve().parent.parent)

    # 打开 PPT 模板
    mc_app = win32com.client.Dispatch("PowerPoint.Application")
    mc_app.DisplayAlerts = 0
    mc_app.Visible = True
    mc_ppt = mc_app.Presentations.Open(_proj_root + r"\src\Template 2.1.pptx")
    mc_slide = mc_ppt.Slides(mc_ppt.Slides.Count)

    # 连接已打开的 Excel（问卷 sheet）
    mc_book = xlwings.books.active
    mc_sht = None
    for s in mc_book.sheets:
        if "问卷" in s.name:
            mc_sht = s
            break
    if mc_sht is None:
        print("未找到包含'问卷'的 sheet，请先在 Excel 打开 template/source data-apparel.xlsx")
        exit(1)

    # apparel 数据源没有"基础信息"sheet，直接用 sheet 名作为 sample_name
    sample_name = mc_sht.name

    print(f"[debug] sample: {sample_name}")
    print(f"[debug] sheet:  {mc_sht.name}")
    print(f"[debug] slides: {mc_ppt.Slides.Count}")

    mc_gpt = "n"  # 调试默认不调 GPT，省时省钱

    new_slide = make_apparel_slide(
        mc_sht, mc_ppt, mc_slide, sample_name,
        mc_gpt=mc_gpt, mc_model=_MODEL,
    )
    print(f"[debug] 完成！新页在第 {new_slide.SlideIndex} 页")
    print(f"[debug] 注意：模板文件未保存，请手动检查后关闭（不要保存）")
