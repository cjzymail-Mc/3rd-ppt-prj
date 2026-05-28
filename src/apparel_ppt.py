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

# apparel 本地常量：纯白（COM BGR long）。用于 p13 跑量/训练标签的白字效果。
_WHITE = 16777215  # 0xFFFFFF — RGB(255,255,255)

# ---------------------------------------------------------------------------
# Pipeline trace (供 ppt-acceptance-check L4 行为层断言用)
# 来源：office-com-helpers skill 的 TraceLogger。无 skill 安装时降级 no-op。
# 由 make_apparel_p13_slide / make_apparel_p14_slide 的 trace_path kwarg 触发开/关。
# 配套契约：acceptance/apparel.json
# ---------------------------------------------------------------------------
import os as _os
_TraceLogger = None
try:
    _OCH_PATH = _os.path.expanduser(r"~/.claude/skills/office-com-helpers")
    if _os.path.isdir(_OCH_PATH) and _OCH_PATH not in sys.path:
        sys.path.insert(0, _OCH_PATH)
    from office_com_helpers import TraceLogger as _TraceLogger  # type: ignore
except Exception:
    _TraceLogger = None

_TRACE = None  # type: ignore  # set by make_apparel_p1{3,4}_slide when trace_path given


def _trace_event(name: str, **fields) -> None:
    """No-op when _TRACE is None。事件直接落 jsonl。"""
    if _TRACE is not None:
        try:
            _TRACE.event(name, **fields)
        except Exception:
            pass


class _NoopShapeWrite:
    """no-op context manager (when _TRACE is None or TraceLogger 不可用)."""
    def __enter__(self):
        return {}
    def __exit__(self, *a):
        return False


def _trace_shape(shape: str, strategy: str, slide=None):
    if _TRACE is not None:
        try:
            return _TRACE.shape_write(shape, strategy, slide=slide)
        except Exception:
            return _NoopShapeWrite()
    return _NoopShapeWrite()


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
_TEMPLATE_SLIDE = 19  # DEPRECATED 2026-05-26：旧 12 页布局，apparel 双页移植后弃用

# ---------------------------------------------------------------------------
# 双页移植 (2026-05-26) — page 13 / page 14 新模板常量
# Clone 源：src/Template 2.1.pptx（Main.py 实际打开的模板，apparel 双页 p13/p14
# 通过 _archive/2026-05-27-debug-cleanup/scripts/merge_apparel_template.py 一次性追加到末尾）
# 合并规则：原 19 页 → 追加 slide 13 → 追加 slide 14 → 总 21 页
#   slide 20 = p13（22 shapes，数据图表型）
#   slide 21 = p14（7 shapes，文字 bullet 型）
# ---------------------------------------------------------------------------
_TEMPLATE_PPTX_NAME = "Template 2.1.pptx"  # 在 src/ 下，由 Main.py 打开

_TEMPLATE_P13_SLIDE = 20
_TEMPLATE_P14_SLIDE = 21

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
# fix3 — 新问卷新增维度 + 元数据列，避免被 _classify_columns 当成文本反馈列
# 污染 GPT prompt（旧问卷不含这些列，传入空命中即 no-op，向后兼容）。
# 展示需求（三围/温度/支撑/跑量/场景）由后续 fix 处理。
# ---------------------------------------------------------------------------
_APPAREL_SKIP_KEYWORDS = [
    # fix3: 新问卷新增维度 + 元数据列，避免被 _classify_columns 当成文本反馈列。
    # 2026-05-26 双页移植注意：
    #   "累计 / 跑量"→ 列 G，现在由 _calc_total_km 专门处理，仍需 skip（非评分列）
    #   "温度"       → 列 AD/AE，由 _calc_temp_mode / _calc_chart63_data 处理，仍 skip
    #   "场景"       → 列 AC，由 _calc_train_ratio 处理，仍 skip
    #   "支撑"       → 列 AF，展示需求未定，维持 skip
    # 这些列不应出现在 GPT prompt 的 score_cols / text_cols 中（会污染 chart 和文本摘要）。
    "三围", "温度", "场景", "支撑", "跑量",
    "data_id", "提交", "更新", "标题", "品牌", "试穿产品",
]

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
# DEPRECATED 2026-05-26：旧 12 页布局，apparel 双页移植后弃用
# 由 make_apparel_slide() 使用；新代码请用 make_apparel_p13_slide / make_apparel_p14_slide
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
# Page 13 SHAPES（22 条，数据图表型）
# 坐标来源：_archive/2026-05-27-debug-cleanup/inspect/inspect-apparel-p1213/inspect_report.md Slide 13 段
# ---------------------------------------------------------------------------
APPAREL_P13_SHAPES = [
    # 装饰线 + 固定文字 → skip（Clone 自动继承）
    {"name": "Straight Connector 4", "strategy": "skip"},
    {"name": "Straight Connector 5", "strategy": "skip"},
    {"name": "TextBox 1",  "strategy": "skip"},   # "服装试穿反馈结果"
    {"name": "TextBox 32", "strategy": "skip"},   # "试穿反馈【 Athletes' Feedback】"

    # 装饰虚线圆 → skip
    {"name": "Oval 3",  "strategy": "skip"},
    {"name": "Oval 13", "strategy": "skip"},
    {"name": "Oval 16", "strategy": "skip"},
    {"name": "Oval 19", "strategy": "skip"},
    {"name": "Oval 49", "strategy": "skip"},

    # 分类评分标签（TextBox 含分类名 + 均分，如 "版型\n3.98 / 5"）
    # L/T/W/H: 30/122/125/73
    {"name": "TextBox 6",  "strategy": "category_score_label",
     "params": {"category": "版型",     "format": "版型\n{mean:.2f} / 5"}},
    # L/T/W/H: 327/125/125/69
    {"name": "TextBox 14", "strategy": "category_score_label",
     "params": {"category": "面料",     "format": "面料\n{mean:.2f} / 5",
                "value_size": 14}},  # 模板 TextBox 14 数值字号 = 14（其他 4 个标签是 16）
    # L/T/W/H: 30/282/125/73
    {"name": "TextBox 17", "strategy": "category_score_label",
     "params": {"category": "吸湿排汗", "format": "吸湿排汗\n{mean:.2f} / 5"}},
    # L/T/W/H: 327/280/125/73
    {"name": "TextBox 20", "strategy": "category_score_label",
     "params": {"category": "速干",     "format": "速干\n{mean:.2f} / 5"}},

    # 受试者信息（复用 gpt_respondent_info；9 人 → 动态扩行）
    # L/T/W/H: 831/54/125/148
    {"name": "TextBox 24", "strategy": "gpt_respondent_info",
     "budget": {"max_chars": 230, "max_lines": 10}},

    # 4 类分析 Chart（按 category 切片均值，OLE 粘贴路线）
    # Chart 7  L/T/W/H: 126/84/190/150
    {"name": "Chart 7",  "strategy": "mean_extraction_filtered",
     "params": {"category": "版型"}},
    # Chart 9  L/T/W/H: 406/106/224/98
    {"name": "Chart 9",  "strategy": "mean_extraction_filtered",
     "params": {"category": "面料"}},
    # Chart 10 L/T/W/H: 145/271/190/98
    {"name": "Chart 10", "strategy": "mean_extraction_filtered",
     "params": {"category": "吸湿排汗"}},
    # Chart 11 L/T/W/H: 436/268/190/98
    {"name": "Chart 11", "strategy": "mean_extraction_filtered",
     "params": {"category": "速干"}},

    # 温度适宜区间 stacked bar Chart 63
    # chart_type=58 (xlBarStacked)，3 系列（起点/区间/终点）× 2 行
    # L/T/W/H: 163/410/427/113
    {"name": "Chart 63", "strategy": "bar_stacked_temp_range"},

    # 新字段：适宜温度（列 AD 众数 bin）
    # L/T/W/H: 26/422/125/73
    {"name": "TextBox 50", "strategy": "temp_mode_label",
     "params": {"format": "适宜温度\n{mode_bin}"}},

    # 新字段：累计跑量（列 G 求和）
    # L/T/W/H: 848/227/92/55
    # 视觉契约（2026-05-28 对齐用户手工示范的 RR 2）：跨段，标题 11pt 白 / 数值 24pt 白
    {"name": "Rounded Rectangle 53", "strategy": "total_km_label",
     "params": {"format": "累计跑量km\n{sum_km}",
                "title_size": 11, "value_size": 24,
                "title_color": _WHITE, "value_color": _WHITE}},

    # 新字段：训练定位（列 AC 含"训练"的行数/总人数）
    # L/T/W/H: 849/309/92/55
    # 视觉契约（2026-05-28 对齐用户手工示范的 RR 7）：同段 2-run，标题 11pt 白 / 数值 24pt 白
    {"name": "Rounded Rectangle 55", "strategy": "train_ratio_label",
     "params": {"format": "定位日常训练\n{n}/{total}",  # \n 仅用于内部 title/value 拆分
                "title_size": 11, "value_size": 24,
                "title_color": _WHITE, "value_color": _WHITE,
                "same_line": True}},
]

# ---------------------------------------------------------------------------
# Page 14 SHAPES（7 条，文字 bullet 型）
# 坐标来源：_archive/2026-05-27-debug-cleanup/inspect/inspect-apparel-p1213/inspect_report.md Slide 14 段
# ---------------------------------------------------------------------------
APPAREL_P14_SHAPES = [
    # 装饰线 + 固定文字 → skip
    {"name": "Straight Connector 4", "strategy": "skip"},
    {"name": "Straight Connector 5", "strategy": "skip"},
    {"name": "TextBox 1",  "strategy": "skip"},   # "服装试穿反馈结果"
    {"name": "TextBox 32", "strategy": "skip"},   # "试穿反馈【 Athletes' Feedback】"

    # 受试者信息（同 p13；可由 caller 传 shared_info 复用，省一次 GPT）
    # L/T/W/H: 831/54/125/148
    {"name": "TextBox 24", "strategy": "gpt_respondent_info",
     "budget": {"max_chars": 230, "max_lines": 10}},

    # 优点 strengths bullet（保留"优点 strengths"标题首行，蓝字关键词染色）
    # L/T/W/H: 34/128/556/134
    {"name": "TextBox 23", "strategy": "gpt_strengths_bullet",
     "budget": {"max_chars": 230, "max_lines": 5}},

    # 缺点 drawbacks bullet（保留"缺点 drawbacks"标题首行，红字关键词染色）
    # L/T/W/H: 35/306/554/189
    {"name": "TextBox 26", "strategy": "gpt_drawbacks_bullet",
     "budget": {"max_chars": 330, "max_lines": 8}},
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

def _section_boundary(headers: List[str], category: str
                      ) -> Tuple[int, int] | None:
    """动态定位 category 在问卷里的评分列范围（按 section terminator 分组）.

    问卷设计规律（新/旧 apparel 都遵守）：每个 category 末尾必有一个
    "XX评价（文字描述）" 列作 section terminator。归属规则：
      - 找到含 _CATEGORY_KEYWORDS[category] 关键词的 terminator
      - 上一个 terminator + 1 ~ 当前 terminator - 1 之间的列归属本 category

    优势：不依赖每个评分列的 header 含 category 关键词，能正确归属
    "3、透气性评分"（不含"面料"）这类列。

    返回 (start_ci, end_ci_exclusive)；找不到对应 terminator 返回 None。
    """
    own_kws = _CATEGORY_KEYWORDS.get(category, [])
    if not own_kws:
        return None
    # 所有 section terminator 的 idx
    seps = [ci for ci, h in enumerate(headers)
            if "文字描述" in h or "描述）" in h]
    if not seps:
        return None
    # 当前 category 对应的 terminator pos
    own_pos = None
    for pos, ci in enumerate(seps):
        if any(kw in headers[ci] for kw in own_kws):
            own_pos = pos
            break
    if own_pos is None:
        return None
    own_ci = seps[own_pos]
    prev_ci = seps[own_pos - 1] if own_pos > 0 else -1
    return (prev_ci + 1, own_ci)


def _clean_chart_label(header: str, category: str) -> str:
    """清理 chart bar 标签 —— 优先取 【...】 内字（如"腰围"），
    否则去序号 / 括号说明 / "评分""性能"后缀 / category 前缀（如"面料"）。"""
    m = re.search(r'【([^】]+)】', header)
    if m:
        return m.group(1).strip()
    s = re.sub(r'^\d+[、.]\s*', '', header)
    s = re.sub(r'（[^）]*）', '', s).strip()
    s = s.replace("评分", "").replace("性能", "").strip()
    # 仅剥 category 名本身（如"面料"），不剥辅助识别 keyword（如"亲肤""轻量"）—
    # 否则"3、轻量化"会被剥成"化"。
    if s.startswith(category):
        s = s[len(category):].strip()
    return s or header[:8]


def _extract_means_for_category(rows: List[List[Any]],
                                category: str) -> List[Tuple[str, float]]:
    """提取某一分类（版型/面料/吸湿排汗/速干）下所有评分列的均值。

    动态识别策略（重构于 fix3 透气性遗漏 bug）：
      - 主路径：按 "XX评价（文字描述）" terminator 分组
        （`_section_boundary`）→ 不依赖每列名字含 category 关键词
      - 兜底：找不到 terminator 时按 _CATEGORY_KEYWORDS 关键词匹配
        （维持向后兼容，避免某些非标准问卷整体崩溃）

    硬过滤（无论走哪条路径）：
      - 复用 _classify_columns 的 score 列判定（含 _APPAREL_SKIP_KEYWORDS）
      - 该列 >50% 数据是 0~10 范围数值
    """
    if not rows or len(rows) < 2:
        return []
    headers = [_safe_text(h) for h in rows[0]]
    data = rows[1:]

    bounds = _section_boundary(headers, category)
    if bounds is not None:
        # 主路径：terminator 分组
        start, end = bounds
        # 复用 _classify_columns 的 score 列判定（已含 _APPAREL_SKIP_KEYWORDS skip）
        score_cols, _text_cols = _classify_columns(
            headers, rows, extra_skip_keywords=_APPAREL_SKIP_KEYWORDS,
        )
        score_set = set(score_cols)
        # 额外 skip：人物信息列（_classify_columns 未涵盖姓名/身高）
        # _classify_columns 的 _SKIP_KEYWORDS 含 "身高"，name_col/weight_col 也会 skip
        ci_list = [ci for ci in range(start, end)
                   if ci < len(headers) and headers[ci] in score_set]
    else:
        # 兜底路径：旧关键词匹配（无 terminator 的极端问卷）
        own_kws = _CATEGORY_KEYWORDS.get(category, [])
        if not own_kws:
            return []
        text_markers = ["文字描述", "描述）", "留言", "提交人", "提交时间", "更新时间", "标题"]
        ci_list = [ci for ci, h in enumerate(headers)
                   if any(kw in h for kw in own_kws)
                   and not any(m in h for m in text_markers)]

    out: List[Tuple[str, float]] = []
    for ci in ci_list:
        h = headers[ci]
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
        out.append((_clean_chart_label(h, category), round(mean_val, 2)))

    return out


def _category_overall_mean(rows: List[List[Any]], category: str) -> float:
    """计算某一分类的整体均值（用于圆环评分）。

    取该分类下所有评分列的均值的均值。返回 0~5 的浮点数，无数据返回 0.0。
    精度 round 到 2 位与下游 `{mean:.2f}` 显示对齐——若 round(1) 则末位永远是 0
    （3.98 会显示成 4.00、3.61 显示成 3.60），与模板原始 2 位显示不一致。
    """
    means = _extract_means_for_category(rows, category)
    if not means:
        return 0.0
    return round(sum(v for _, v in means) / len(means), 2)


# ---------------------------------------------------------------------------
# 双页新字段聚合函数（page 13 数据展示，2026-05-26）
# 源 sheet: 服装试穿问卷--紧身背心（10 行 × 36 列，表头 + 9 名受试者）
# ---------------------------------------------------------------------------

def _calc_total_km(rows: List[List[Any]]) -> int:
    """累计跑量求和（列 G `6、测试累计总跑量（km）`）。

    兼容格式：数字 63 / 字符串 "55km" / 字符串 "120"。
    re.findall(r"\\d+", str(val)) 取第一个整数，全样本求和。
    返回整数总和（如 671）。
    """
    col_h = _find_col(
        [_safe_text(h) for h in rows[0]] if rows else [],
        ["累计", "跑量"],
    )
    if not col_h or len(rows) < 2:
        return 0
    total = 0
    for row in rows[1:]:
        headers = [_safe_text(h) for h in rows[0]]
        ci = headers.index(col_h) if col_h in headers else -1
        if ci < 0 or ci >= len(row):
            continue
        val = row[ci]
        if val is None:
            continue
        nums = re.findall(r"\d+", str(val))
        if nums:
            total += int(nums[0])
    return total


def _calc_temp_mode(rows: List[List[Any]]) -> str:
    """适宜温度众数 bin（列 AD `适合的温度区间（体感温度）`）。

    取出现频次最高的 bin（如"15℃~25℃"），保留原始 ℃ 格式（不剥前缀 ℃）。

    修复（2026-05-27）：
      1. 使用更精确的关键词 ["适合的温度", "体感温度"] 避免与其他含"温度"的列
         （如 AE 实际穿着温度区间）产生误匹配。
      2. 规整化先于 Counter 统计：把全角波浪线 "℃～" 和 ASCII "℃~" 统一为
         "~" 后再计数，防止同一区间因编码差异被分裂为不同 key（如 "15℃~25℃"
         和 "15℃～25℃" 本应归为一类，但旧代码 Counter 视为两个不同 key）。
    """
    from collections import Counter
    headers = [_safe_text(h) for h in rows[0]] if rows else []
    # 精确关键词：先试"适合的温度"/"体感温度"，fallback 到旧的"适合"
    col_h = _find_col(headers, ["适合的温度", "体感温度"])
    if not col_h:
        col_h = _find_col(headers, ["适合", "体感"])
    if not col_h or len(rows) < 2:
        return ""
    ci = headers.index(col_h) if col_h in headers else -1
    if ci < 0:
        return ""
    raw_bins = [_safe_text(row[ci]) for row in rows[1:] if ci < len(row) and row[ci]]
    raw_bins = [b for b in raw_bins if b]
    if not raw_bins:
        return ""

    def _normalize_key(b: str) -> str:
        """仅统一波浪线，用于计数分组（不修改 ℃ 位置）。

        修复 Bug D：旧 _normalize 把 '5℃~15℃' 剥成 '5~15℃'（前 ℃ 丢失）。
        现在 key 仅做全角→半角波浪线统一，计数后返回原始 raw_bins 中
        出现次数最多的那个原始字符串，不改变 ℃ 位置。
        """
        return b.replace("～", "~").strip()

    # 按 normalized key 计数，但返回对应的原始 raw_bin（保留 ℃ 位置）
    key_to_raw: dict = {}
    for raw in raw_bins:
        k = _normalize_key(raw)
        if k not in key_to_raw:
            key_to_raw[k] = raw  # 记录第一次出现的原始值（代表该 bin）
    key_counts = Counter(_normalize_key(b) for b in raw_bins)
    most_key = key_counts.most_common(1)[0][0]
    return key_to_raw[most_key]


def _calc_train_ratio(rows: List[List[Any]]) -> Tuple[int, int]:
    """训练定位（列 AC `适宜的穿着场景`）。

    统计含"训练"的行数 / 总有效行数，返回 (n, total)。
    """
    headers = [_safe_text(h) for h in rows[0]] if rows else []
    col_h = _find_col(headers, ["穿着场景", "适宜的穿着"])
    if not col_h or len(rows) < 2:
        return (0, 0)
    ci = headers.index(col_h) if col_h in headers else -1
    if ci < 0:
        return (0, 0)
    n = 0
    total = 0
    for row in rows[1:]:
        val = row[ci] if ci < len(row) else None
        if val is None or _safe_text(val) == "":
            continue
        total += 1
        if "训练" in _safe_text(val):
            n += 1
    return (n, total)


def _calc_chart63_data(rows: List[List[Any]]) -> dict:
    """计算 Chart 63（xlBarStacked，3 系列 × 2 行）的数据。

    Chart 63 结构（来自 chart63_data.json）：
      series 1 "起点（占位）"  → 不可见的前置偏移
      series 2 "温度区间（℃）" → 可见的区间长度
      series 3 "终点（占位）"  → 35 - max 的后置填充

    两行（2 个 category）：
      row "体感适宜区间" → 源列 AD `适合的温度区间（体感温度）`
      row "实际穿着区间" → 源列 AE `实际穿着温度区间（℃）`

    对每一行：对该列所有 cell re.findall(r"\\d+", val) 收集全部数字，
      全集取 min/max，固定 35 为总长（不要改）。
      start = min(数字集)
      range_ = max(数字集) - min(数字集)
      end = 35 - max(数字集)

    返回 {
      "x_values": ["体感适宜区间", "实际穿着区间"],
      "s1_values": [start_感, start_实],   # 起点偏移
      "s2_values": [range_感, range_实],    # 区间长度
      "s3_values": [end_感,   end_实],      # 尾部填充
    }
    """
    headers = [_safe_text(h) for h in rows[0]] if rows else []

    # 列 AD：适合的温度区间（体感温度）—— 搜索词按优先级从精到粗
    col_ad = _find_col(headers, ["体感温度", "适合的温度", "适合", "体感"])
    # 列 AE：实际穿着温度区间（℃）—— "实际穿着" 在问卷中唯一，不会和 AD 冲突
    col_ae = _find_col(headers, ["实际穿着"])

    def _extract_temp_nums(col_h: str) -> List[int]:
        """从整列提取所有出现的整数（合并多行）。"""
        if not col_h:
            return []
        ci = headers.index(col_h) if col_h in headers else -1
        if ci < 0:
            return []
        all_nums: List[int] = []
        for row in rows[1:]:
            val = row[ci] if ci < len(row) else None
            if val is None:
                continue
            found = re.findall(r"\d+", str(val))
            all_nums.extend(int(x) for x in found)
        return all_nums

    def _to_stacked(nums: List[int]) -> Tuple[int, int, int]:
        if not nums:
            return (10, 10, 15)  # 兜底：若无数据用占位值
        lo = min(nums)
        hi = max(nums)
        return (lo, hi - lo, 35 - hi)

    nums_ad = _extract_temp_nums(col_ad)
    nums_ae = _extract_temp_nums(col_ae)

    start_ad, range_ad, end_ad = _to_stacked(nums_ad)
    start_ae, range_ae, end_ae = _to_stacked(nums_ae)

    return {
        "x_values": ["体感适宜区间", "实际穿着区间"],
        "s1_values": [start_ad, start_ae],
        "s2_values": [range_ad, range_ae],
        "s3_values": [end_ad,   end_ae],
    }


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
    score_cols, text_cols = _classify_columns(
        headers, rows, extra_skip_keywords=_APPAREL_SKIP_KEYWORDS,
    )

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


def _build_strengths_prompt(budget: dict, rows: List[List[Any]]) -> str:
    """Build GPT prompt for TextBox 23 (page 14 优点 strengths).

    输出格式：首行保留"优点 strengths"标题，然后逐条优点，
    每条后面注明 (n/N) 频次，关键词用【】标记以便染色。
    """
    respondent_block, n = _build_respondent_block(rows)
    max_chars = budget.get("max_chars", 230)
    max_lines = budget.get("max_lines", 5)

    sample_output = (
        "优点 strengths\r"
        "整体版型【合身】、修身显身材，上身贴合，动作舒展不受限（9/9）。\r"
        f"具备一定【支撑性】，覆盖中低到高强度训练（7/{n}）。\r"
        f"面料有【亲肤性】与一定耐用性，多名反馈不起球不勾丝"
    )

    return (
        f"【参考输出格式（参考语调、字数密度）】\n{sample_output}\n\n"
        f"【你的任务】\n"
        f"从以下 {n} 名测试者的原始反馈中，归纳这件服装的主要【优点】。\n"
        f"输出格式要求：\n"
        f"  - 第一行固定为：优点 strengths\n"
        f"  - 接下来每行一条优点，末尾注明（X/{n}）表示几分之几测试者有此体验\n"
        f"  - 每条结论中，将最核心的 1-2 个关键词用【】括起来（优势词，如【合身】【亲肤性】）\n"
        f"注意：\n"
        f"  - 评分为 5 分制（1=最差，5=最佳）\n"
        f"  - 只能分析已有数据，不能推测或编造\n"
        f"  - 总字数控制在 {max_chars} 字左右，不超过 {max_lines} 行（含标题行）\n"
        f"  - 直接输出结果，不需要任何前言。\n\n"
        f"【{n} 名测试者原始反馈】\n{respondent_block}\n\n"
        f"直接输出结论（第一行 = 优点 strengths）："
    )


def _build_drawbacks_prompt(budget: dict, rows: List[List[Any]]) -> str:
    """Build GPT prompt for TextBox 26 (page 14 缺点 drawbacks).

    输出格式：首行保留"缺点 drawbacks"标题，然后逐条缺点，
    每条后面注明 (n/N) 频次，关键词用【】标记以便染色（红字）。
    """
    respondent_block, n = _build_respondent_block(rows)
    max_chars = budget.get("max_chars", 330)
    max_lines = budget.get("max_lines", 8)

    sample_output = (
        "缺点 drawbacks\r"
        "面料【弹性不够】（2/9）\r"
        "前胸【闷热】、面料偏厚（4/9）\r"
        f"腋下【摩擦】、副乳【硌感】、胸下磨皮较集中，长距离或出汗后更明显（8/{n}）。\r"
        f"透气排汗不足较突出，局部有只吸不排、贴身感，速干较差（6/{n}）。\r"
        "建议改为【侧向开口】/【更低位置】设计。"
    )

    return (
        f"【参考输出格式（参考语调、字数密度）】\n{sample_output}\n\n"
        f"【你的任务】\n"
        f"从以下 {n} 名测试者的原始反馈中，归纳这件服装的主要【缺点】。\n"
        f"输出格式要求：\n"
        f"  - 第一行固定为：缺点 drawbacks\n"
        f"  - 中间每行一条缺点，末尾注明（X/{n}）表示几分之几测试者有此体验\n"
        f"  - 每条结论中，将最核心的 1-2 个关键词用【】括起来（问题词，如【摩擦】【闷热】）\n"
        f"  - **末行固定**：以「建议改为...」开头的一句具体改进建议（针对最主要的缺点），\n"
        f"    建议中也用【】标记 1-2 个改进方向关键词（如【侧向开口】【更低位置】）\n"
        f"注意：\n"
        f"  - 评分为 5 分制（1=最差，5=最佳）\n"
        f"  - 只能分析已有数据，不能推测或编造\n"
        f"  - 总字数控制在 {max_chars} 字左右，不超过 {max_lines} 行（含标题行 + 末尾建议行）\n"
        f"  - 直接输出结果，不需要任何前言。\n\n"
        f"【{n} 名测试者原始反馈】\n{respondent_block}\n\n"
        f"直接输出结论（第一行 = 缺点 drawbacks，末行 = 建议改为...）："
    )


def _call_gpt(prompt: str, fallback: str, enabled: bool, model: str,
              label: str = "gpt") -> str:
    """Call GPT_5 if enabled and available; return fallback otherwise.

    label: 用于 trace 事件名（如 "gpt_strengths" / "gpt_drawbacks"）。
           ppt-acceptance-check L4 行为层会 assert 这些事件出现。
    """
    if not enabled or GPT_5 is None:
        _trace_event(label, called=False,
                     reason=("mc_gpt=n" if not enabled else "GPT_5 unavailable"))
        return fallback
    try:
        result = _safe_text(GPT_5(prompt, model))
        if result:
            _trace_event(label, called=True, ok=True,
                         chars_out=len(result))
            return result
        _trace_event(label, called=True, ok=False, reason="empty result")
    except Exception as _e:
        _trace_event(label, called=True, ok=False, reason=f"exception:{_e}")
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
        # fix3: budget 按 n 动态化 —— 旧 max_lines=5 在 n>=5 时丢人（A1 bug）。
        # 标题 1 行 + 每名 1 行 = n+1 行；每行约 22 字（如 "A: 165CM / 50 KG"）。
        n = max(0, len(rows) - 1)
        budget = dict(budget)  # 不修改 SHAPES 里的原 dict
        budget["max_lines"] = max(budget.get("max_lines", 5), n + 1)
        budget["max_chars"] = max(budget.get("max_chars", 102), n * 22)

        style_anchor = (
            "受试者信息 Information\r"
            "A: 160CM / 50 KG\r"
            "B: 165CM / 52 KG\r"
            "C: 162CM / 58 KG\r"
            "D: 168CM / 60 KG"
        )
        fallback = _build_respondent_info_fallback(rows)
        prompt = _build_respondent_info_prompt(budget, rows, style_anchor)
        result = _call_gpt(prompt, fallback, gpt_enabled, model,
                           label="gpt_respondent_info")
        max_chars = budget["max_chars"]
        max_lines = budget["max_lines"]
        return clamp_text(result, max_chars, max_lines)

    if strategy == "gpt_prompted":
        focus = params.get("filter", "")
        src = params.get("source", "补充说明")
        # fix3 (B2): fallback 比例数字动态化 —— 旧版硬编码 (4/4) 是 4 人样本字面量，
        # 新问卷 n=9/10+ 时显然失真。fallback 极少触发，但保持字面一致性是基线。
        n = max(0, len(rows) - 1)
        fallback_map = {
            "优点": (
                f"轻薄透气，舒适不闷热（{n}/{n}）\r"
                f"吸湿排汗性能好（{n}/{n}）；\r"
                f"速干性能极好（{n}/{n}）大部分时间都很干爽"
            ),
            "缺点": (
                "肩部挖空偏小，会勒腋下（部分样本）；\r"
                "后背开叉偏高（部分样本）；"
            ),
        }
        fallback = fallback_map.get(focus, "样本反馈总体稳定，核心指标表现均衡。")
        prompt = _build_rich_prompt(
            budget, rows, focus=focus,
            content_source=src, style_anchor=_STYLE_REFERENCE_CORPUS,
        )
        label = "gpt_advantage" if focus == "优点" else (
            "gpt_disadvantage" if focus == "缺点" else "gpt_prompted")
        result = _call_gpt(prompt, fallback, gpt_enabled, model, label=label)
        # 安全网：GPT 返回后强制 clamp 到 budget 范围内
        max_chars = budget.get("max_chars", 100)
        max_lines = budget.get("max_lines", 3)
        return clamp_text(result, max_chars, max_lines)

    # ---- page 13 新 strategy（2026-05-26）----

    if strategy == "category_score_label":
        # 分类评分标签：例如"版型\n3.98 / 5"
        cat = params.get("category", "")
        fmt = params.get("format", "{mean:.2f} / 5")
        mean_val = _category_overall_mean(rows, cat)
        return fmt.replace("{mean:.2f}", f"{mean_val:.2f}")

    if strategy == "temp_mode_label":
        # 适宜温度众数 bin 标签：例如"适宜温度\n15~25℃"
        fmt = params.get("format", "适宜温度\n{mode_bin}")
        mode_bin = _calc_temp_mode(rows)
        return fmt.replace("{mode_bin}", mode_bin or "—")

    if strategy == "total_km_label":
        # 累计跑量标签：例如"累计跑量km\n671"
        fmt = params.get("format", "累计跑量km\n{sum_km}")
        sum_km = _calc_total_km(rows)
        return fmt.replace("{sum_km}", str(sum_km))

    if strategy == "train_ratio_label":
        # 训练定位标签：例如"定位日常训练7/9"
        fmt = params.get("format", "定位日常训练{n}/{total}")
        n_train, total = _calc_train_ratio(rows)
        return fmt.replace("{n}", str(n_train)).replace("{total}", str(total))

    # bar_stacked_temp_range：返回空字符串（数据计算在 make_apparel_p13_slide 内进行）
    if strategy == "bar_stacked_temp_range":
        return ""  # handled separately in make_apparel_p13_slide

    # ---- page 14 新 strategy（2026-05-26）----

    if strategy == "gpt_strengths_bullet":
        # 优点 bullet：保留"优点 strengths"标题首行，蓝字关键词染色
        n = max(0, len(rows) - 1)
        fallback = (
            f"优点 strengths\r"
            f"整体版型合身、修身显身材，上身贴合，动作舒展不受限（{n}/{n}）。\r"
            f"面料有亲肤性与一定耐用性，多名反馈不起球不勾丝"
        )
        prompt = _build_strengths_prompt(budget, rows)
        result = _call_gpt(prompt, fallback, gpt_enabled, model,
                           label="gpt_strengths")
        max_chars = budget.get("max_chars", 230)
        max_lines = budget.get("max_lines", 5)
        return clamp_text(result, max_chars, max_lines)

    if strategy == "gpt_drawbacks_bullet":
        # 缺点 bullet：保留"缺点 drawbacks"标题首行，红字关键词染色
        n = max(0, len(rows) - 1)
        fallback = (
            f"缺点 drawbacks\r"
            f"腋下摩擦、副乳硌感，长距离或出汗后更明显（{n}/{n}）。\r"
            f"前胸闷热、透气排汗不足较突出（{n}/{n}）。"
        )
        prompt = _build_drawbacks_prompt(budget, rows)
        result = _call_gpt(prompt, fallback, gpt_enabled, model,
                           label="gpt_drawbacks")
        max_chars = budget.get("max_chars", 330)
        max_lines = budget.get("max_lines", 8)
        return clamp_text(result, max_chars, max_lines)

    return ""


# ---------------------------------------------------------------------------
# apparel 专用：2-run 标签写入（Bug B 修复）
# 适用策略：category_score_label / temp_mode_label / total_km_label /
#           train_ratio_label
# 模板 run 结构（probe_5shape_runs.py 已验证）：
#   run0 = 类别名（如"版型\r"） → 黑色 RGB(0,0,0)，size=20，非粗
#   run1 = 数值（如"3.98 / 5"） → 红色 RGB(255,0,0)，size=16，非粗
# _write_text 会把整个 TextRange 覆盖为单 run，丢失 run1 的红色/小字号。
# 本函数写入前先 split("\n")，分别对 run0 和 run1 设定 rgb 和 size。
# ---------------------------------------------------------------------------

def _write_two_run_label(shp, content: str,
                         title_size: int = 20, value_size: int = 16,
                         title_color: int | None = None,
                         value_color: int | None = None,
                         same_line: bool = False) -> bool:
    """写入 2-run 标签 shape（类别名 + 数值，可独立控制字号 / 颜色 / 是否同段）.

    content 格式（\n 分隔 title/value，function 内部决定如何渲染到 PPT）：
        "版型\n3.98 / 5"
        "累计跑量km\n465"
        "定位日常训练\n7/9"

    渲染规则：
        same_line=False（默认）→ \n 转 \r 写成 2 段
            行 0（title） → title_color（默认 _BLACK）+ title_size（默认 20）+ bold
            行 1（value） → value_color（默认 _RED）+ value_size（默认 16）+ bold
        same_line=True → 把 title + value 拼成同段单行，用 Characters() 切片设样式
            （用户手工示范 RR 7 = "定位日常训练7/9" 同段不同字号的契约）

    如果 content 只有 1 行（无 \n），退化：整段按 title_size/title_color。
    """
    if title_color is None:
        title_color = _BLACK
    if value_color is None:
        value_color = _RED

    if not bool(_com_get(shp, "HasTextFrame", 0)):
        return False
    tf = _com_get(shp, "TextFrame", None)
    if tf is None:
        return False
    tr = _com_get(tf, "TextRange", None)
    if tr is None:
        return False

    try:
        tf.AutoSize = 0
    except Exception:
        pass

    parts = content.split("\n", 1)
    title_text = parts[0] if parts else ""
    value_text = parts[1].strip() if len(parts) > 1 else ""
    title_stripped = title_text.strip()

    # 决定 PPT 实际文本：same_line → 同段拼接；否则 \n → \r 分段
    if same_line and value_text:
        text_ppt = title_text + value_text  # 不 strip title，保字符长度精确
    else:
        text_ppt = content.replace("\n", "\r")

    try:
        tr.Text = text_ppt
        tr.Font.Name = "微软雅黑"
    except Exception:
        return False

    # 整段先 reset 成 title_color + 非粗（防模板残留样式继承）
    try:
        tr.Font.Color = title_color
        tr.Font.Bold = False
    except Exception:
        pass

    if same_line and value_text:
        # 同段 2-run 模式：Characters(start, length) 切片（PPT COM 是 1-based）
        t_len = len(title_text)
        v_len = len(value_text)
        try:
            tch = tr.Characters(1, t_len)
            tch.Font.Size = title_size
            tch.Font.Color = title_color
            tch.Font.Bold = True
        except Exception:
            pass
        try:
            vch = tr.Characters(t_len + 1, v_len)
            vch.Font.Size = value_size
            vch.Font.Color = value_color
            vch.Font.Bold = True
        except Exception:
            pass
    else:
        # 跨段模式：Paragraphs(n) 直接定位（fix6-b 选型，比 tr.Find 可靠）
        try:
            p1 = tr.Paragraphs(1)
            p1.Font.Size = title_size
            p1.Font.Color = title_color
            p1.Font.Bold = True
        except Exception:
            pass
        if value_text:
            try:
                p2 = tr.Paragraphs(2)
                p2.Font.Size = value_size
                p2.Font.Color = value_color
                p2.Font.Bold = True
            except Exception:
                pass

    return True


_TWO_RUN_STRATEGIES = frozenset({
    "category_score_label",
    "temp_mode_label",
    "total_km_label",
    "train_ratio_label",
})


# ---------------------------------------------------------------------------
# apparel 专用：优点/缺点 bullet 染色（Bug A 修复）
# 适用策略：gpt_strengths_bullet / gpt_drawbacks_bullet
#
# 与通用 _apply_keyword_color 的区别：
#   1. 首行（"优点 strengths" / "缺点 drawbacks"）跳过颜色重置，
#      保留模板原色（深红 0xc0 / 深青 0xc07000）。
#   2. 全局先 bold=False，避免模板 run 继承 bold 污染正文。
#   3. 关键词才 bold=True（其余正文保持非粗）。
# ---------------------------------------------------------------------------

def _apply_apparel_bullet_color(shp, bump_last_para_size: float | None = None) -> None:
    """apparel 优点/缺点 bullet 专用染色。

    规则：
      - 首行（如"优点 strengths\r"）：跳过颜色/bold 重置，保留模板原色
      - 其余行：关键词（【】标记）→ 对应颜色 + bold；正文 → 黑色 + 非粗

    bump_last_para_size (fix6-b, 2026-05-27)：
      传入字号（如 16.0）时，把最后一个段落（按 \\r 切分）的所有 run 字号设为该值。
      用于 TextBox 26 (gpt_drawbacks_bullet)：模板末尾"建议改为..."句子是 size 16
      （vs 正文 size 14），表达"改进建议"结构性强调。
      传 None 时跳过此步骤（默认行为，gpt_strengths_bullet 用此）。
    """
    try:
        tf = _com_get(shp, "TextFrame", None)
        if tf is None:
            return
        tr = tf.TextRange
        full_text = tr.Text

        keywords = list(dict.fromkeys(re.findall(r'【([^】]+)】', full_text)))
        if not keywords:
            return

        # 确定 section 颜色（优势→红，缺点→蓝）
        kw_color: dict = {}
        current_section = "neutral"
        for line in full_text.split('\r'):
            line_stripped = line.strip()
            if any(m in line_stripped for m in _ADVANTAGE_MARKERS):
                current_section = "advantage"
            elif any(m in line_stripped for m in _DISADVANTAGE_MARKERS):
                current_section = "disadvantage"
            for kw in re.findall(r'【([^】]+)】', line):
                if current_section == "advantage":
                    kw_color[kw] = _RED
                elif current_section == "disadvantage":
                    kw_color[kw] = _BLUE

        # 去除 【】 括号
        tr.Text = re.sub(r'[【】]', '', full_text)

        # 更新 full_text（去括号后）
        full_text_clean = tr.Text

        # 找首行结束位置（第一个 \r）
        first_para_end = full_text_clean.find('\r')

        # 全局 bold=False（清模板继承）
        try:
            tr.Font.Bold = False
        except Exception:
            pass

        # 非首行部分重置为黑色（首行跳过，保留模板原色）
        if first_para_end >= 0:
            # 首行之后的文本 range（从 first_para_end+1 到末尾）
            body_start = first_para_end + 2  # +1 for \r, +1 for 1-based index
            if body_start <= tr.Length:
                try:
                    body_range = tr.Characters(body_start, tr.Length - body_start + 1)
                    body_range.Font.Color = _BLACK
                except Exception:
                    # fallback：整段 reset 黑色（会覆盖首行，可接受）
                    try:
                        tr.Font.Color = _BLACK
                    except Exception:
                        pass

            # 首行标题恢复 bold（line 1333 全局清了 bold；模板首行是 bold=True，
            # acceptance L3 已抓出标题缺 bold）
            if first_para_end > 0:
                try:
                    title_range = tr.Characters(1, first_para_end)
                    title_range.Font.Bold = True
                except Exception:
                    pass
        else:
            # 无首行 \r，整段重置（退化为通用行为）
            try:
                tr.Font.Color = _BLACK
            except Exception:
                pass

        # 关键词：bold + color
        for kw, color in kw_color.items():
            start = 1
            while start <= tr.Length:
                found = tr.Find(kw, start)
                if found is None:
                    break
                try:
                    found.Font.Bold = True
                    found.Font.Color = color
                except Exception:
                    pass
                start = found.Start + found.Length

        # fix6-b: bump 末段字号（drawbacks 末尾"建议改为..."句子模板原为 size 16）
        if bump_last_para_size is not None:
            try:
                # 用 \r 切分（PowerPoint TextRange 段落分隔符），定位最后一段
                last_break = full_text_clean.rfind('\r')
                if last_break >= 0 and last_break + 1 < tr.Length:
                    last_start = last_break + 2  # +1 for \r, +1 for 1-based index
                    last_len = tr.Length - last_start + 1
                    if last_len > 0:
                        last_range = tr.Characters(last_start, last_len)
                        last_range.Font.Size = float(bump_last_para_size)
            except Exception:
                pass  # 字号 bump 失败不阻断（cosmetic）
    except Exception:
        pass  # coloring is cosmetic — never fail the build

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

    # fix3-bonus（chart 污染修复）：写入前清 anchor 周边残留，防 CurrentRegion 扩张
    # 根因：questionnaire / 上次 apparel 在 col 3~18 留临时数据，anchor 区域虽然
    # 自身 2 列干净，但右侧残留会让 anchor.api.CurrentRegion 一路扩到 col R（18 列），
    # 后续 make_chart_for_apparel set_source_data 抓 18 列 → series.name 变跑者名、
    # 17 个 values 全混进来。修法：写入前精确清 anchor footprint + 右侧 18 列。
    # 仅清当前 chart 的行 footprint（不影响其他 chart anchor 区域）。
    n_rows = len(table)
    try:
        mc_sht.range(
            (anchor.row, anchor.column),
            (anchor.row + n_rows - 1, anchor.column + 18),
        ).clear_contents()
    except Exception as _e:
        print(f"  [apparel-chart] clear 残留失败（{_e}），继续写入")

    anchor.value = table
    print(f"  [apparel-chart] 临时数据已写入：anchor=({anchor.row},{anchor.column})，{len(parsed)} 个指标")
    # fix3-bonus-4：返回 (anchor, n_rows)，让 make_chart_for_apparel 不依赖
    # expand("down") 探尺寸——expand 遇空 cell 才停，questionnaire runner 数据
    # 落在 anchor 下方时会被多吃（实测：版型 chart 多出"Alisa: 4" 一条 bar）。
    return anchor, n_rows


def make_chart_for_apparel(mc_cell, mc_slide, Left, Top, Width, Height,
                           n_rows=None):
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

    # fix3-bonus（chart 污染修复）：不走 mc_cell.api.CurrentRegion
    # —— CurrentRegion 会被旁边残留数据撑大到 18 列（实测 questionnaire 留的数据
    # 会让前两个 chart 的 anchor.CurrentRegion 扩到 A61:R68 / A72:R74）。
    # apparel 临时表固定 2 列（指标 | 均值），行数沿 anchor 同列 expand('down') 探。
    # 配合 _prepare_apparel_chart_data 写入前 .clear_contents()，双保险。
    i0 = mc_cell.row
    j0 = mc_cell.column
    j = 2  # apparel 固定 2 列
    # fix3-bonus-4：优先用 caller 显式传入的 n_rows（最准）；
    # 不传时回退到 expand("down")（向后兼容），但 expand 会被 questionnaire
    # runner 残留多吃（实测："Alisa: 4" 多出一条 bar）。
    if n_rows is not None:
        i = n_rows
    else:
        try:
            i = mc_cell.expand("down").last_cell.row - i0 + 1
        except Exception:
            i = 2  # 兜底：至少 1 header + 1 data
    print(f"  [apparel-chart] source range: i0={i0} j0={j0} rows={i} cols={j}")

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
                mc_cell, n_chart_rows = _prepare_apparel_chart_data(
                    mc_sht, content, anchor_offset=chart_anchor_offset,
                )
                chart_anchor_offset += len(content.splitlines()) + 5
            except Exception as _e:
                print(f"    [警告] 写入 Excel 临时数据失败: {_e}")
                continue

            try:
                _tmp_chart = make_chart_for_apparel(
                    mc_cell, new_slide, Left=L, Top=T, Width=W, Height=H,
                    n_rows=n_chart_rows,
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
            # fix3: TextBox 24 受试者信息按样本数动态拉长 Height
            # 模板原 Height 适配 6 行（1 标题 + 5 人）；n>5 时按 (n+1)/6 线性扩展。
            # 撞下方 shape 不做 boundary check（用户审核确认，人工处理）。
            if name == "TextBox 24" and ok:
                n_samples = max(0, len(rows) - 1)
                if n_samples > 5:
                    try:
                        original_h = float(_com_get(shp, "Height", 0))
                        new_h = original_h * (n_samples + 1) / 6
                        shp.Height = new_h
                        print(f"    [TextBox 24] Height 拉长：{original_h:.1f} → {new_h:.1f}（{n_samples} 名样本）")
                    except Exception as _e:
                        print(f"    [TextBox 24] 拉长 Height 失败: {_e}")

    # 工作完成，删除等待遮罩
    if _overlay is not None and remove_gpt_waiting_overlay is not None:
        try:
            remove_gpt_waiting_overlay(_overlay)
        except Exception:
            pass

    print(f"[apparel] 完成！新页在第 {new_slide.SlideIndex} 页")
    return new_slide


# ---------------------------------------------------------------------------
# Chart 63 原位注入：_write_chart 路线（ChartData.Workbook 直接改）
# Chart 63 是模板里已有的 xlBarStacked chart，有内嵌 workbook；
# 不走 OLE 粘贴路线，走 _write_chart 的 Activate → 写 SeriesCollection 路线。
# ---------------------------------------------------------------------------

def _write_chart63(shp, chart_data: dict) -> bool:
    """将温度区间数据写入 Chart 63（xlBarStacked，3 系列 × 2 行）。

    chart_data 来自 _calc_chart63_data()，格式：
      {
        "x_values": ["体感适宜区间", "实际穿着区间"],
        "s1_values": [start1, start2],  # 系列 1 起点偏移
        "s2_values": [range1, range2],  # 系列 2 区间长度（可见部分）
        "s3_values": [end1,   end2],    # 系列 3 尾部填充
      }

    写入策略（fix6-a 真修，2026-05-27）：
      - SeriesCollection.Values 直接赋值（pipeline 验证可行路径）
      - BreakLink 隔离 try 断外部链接（fix3 硬规则）
      - **不做 src 端"回读自证"**：SeriesCollection 回读返回 in-memory 值，
        不区分"写入持久化"vs"in-memory 同步"。持久化证明留给 L1 acceptance check
        (`chart_series_values` with Excel-derived expected) ——它从 .pptx 文件读 series
        值与 Excel 计算值对比，是真实磁盘验证（fix5 红旗 4 教训：src 端 hardcode 回读
        是伪验证，应由独立 layer 兜底）。

    历史背景：
      - fix5 §4 红旗：developer 加 `if abs(val - 5.0) < 0.5` 自证，5.0 恰好 = 模板默认
        → 永远通过 → 假持久化保证。
      - fix6-a 之前尝试 ChartData.Workbook 直写持久化路径，但 COM error -2147352567
        在非交互会话不稳定（Workbook 访问需 Activate，Activate 也不稳）。回到
        SeriesCollection-only 简单可靠路径，把持久化验证下沉到 acceptance L1。
    """
    chart = _com_get(shp, "Chart", None)
    if chart is None:
        print("  [Chart63] 未找到 Chart 对象，跳过")
        _trace_event("com_api_failed_but_continued",
                     shape="Chart 63", api="Shape.Chart",
                     reason="chart object is None")
        return False

    x_values = chart_data.get("x_values", ["体感适宜区间", "实际穿着区间"])
    s1 = chart_data.get("s1_values", [5, 15])
    s2 = chart_data.get("s2_values", [20, 17])
    s3 = chart_data.get("s3_values", [10, 3])

    print(f"  [Chart63] 准备写入: x={x_values} s1={s1} s2={s2} s3={s3}")

    # 抑制 PowerPoint 弹窗（fix6-a 2026-05-27）：
    # Chart 63 IsLinked=True 时 BreakLink/ChartData 操作触发"连接文件不可用"对话框，
    # 阻断脚本流程。设 Application.DisplayAlerts=1 (ppAlertsNone) 完全抑制。
    _ppt_app = None
    _prev_alerts = None
    try:
        _ppt_app = chart.Application
        _prev_alerts = _ppt_app.DisplayAlerts
        _ppt_app.DisplayAlerts = 1  # ppAlertsNone
    except Exception:
        pass

    # Pre-write BreakLink（隔离 try，失败不阻断后续写入）
    try:
        chart.ChartData.BreakLink()
        time.sleep(0.3)
        print("  [Chart63] BreakLink 完成（pre-write）")
    except Exception as _e:
        print(f"  [Chart63] BreakLink 异常（可忽略）: {_e}")

    # SeriesCollection.Values 写入（pipeline/03b_build_ppt_com.py::_write_chart 同款路径）
    try:
        series1 = chart.SeriesCollection(1)
        series1.XValues = tuple(x_values)
        series1.Values  = tuple(s1)
        series2 = chart.SeriesCollection(2)
        series2.XValues = tuple(x_values)
        series2.Values  = tuple(s2)
        series3 = chart.SeriesCollection(3)
        series3.XValues = tuple(x_values)
        series3.Values  = tuple(s3)
        time.sleep(0.3)
        print("  [Chart63] SeriesCollection 3 系列写入完成")
    except Exception as _e:
        print(f"  [Chart63] SeriesCollection 写入失败: {_e}")
        _trace_event("com_api_failed_but_continued",
                     shape="Chart 63", api="SeriesCollection.Values",
                     reason=str(_e))
        return False

    # Post-write BreakLink（fix3 硬规则）
    try:
        chart.ChartData.BreakLink()
        time.sleep(0.3)
        print("  [Chart63] BreakLink 完成（post-write）")
    except Exception as _e:
        print(f"  [Chart63] BreakLink 异常（可忽略）: {_e}")

    # 恢复 DisplayAlerts
    if _ppt_app is not None and _prev_alerts is not None:
        try:
            _ppt_app.DisplayAlerts = _prev_alerts
        except Exception:
            pass

    # 正面信号：写入 COM 调用未抛异常。真正持久化由 L1 acceptance 验证。
    _trace_event("chart63_write_ok", shape="Chart 63",
                 s1=list(s1), s2=list(s2), s3=list(s3))
    return True


def _write_shapes_to_slide(new_slide, shapes_list: list, rows: list,
                           mc_sht, gpt_enabled: bool, mc_model: str,
                           chart_anchor_offset: int = 50,
                           shared_info: str = "") -> int:
    """通用 shape 写入循环（供 p13 / p14 共用）.

    Returns updated chart_anchor_offset（供多次调用叠加）。
    shared_info: 如果非空，gpt_respondent_info strategy 直接用此值，不再调 GPT。
    """
    offset = chart_anchor_offset
    for spec in shapes_list:
        name     = spec["name"]
        strategy = spec["strategy"]

        if strategy == "skip":
            print(f"  [skip] {name}")
            continue

        # 找 shape
        shp = None
        try:
            shp = new_slide.Shapes(name)
        except Exception:
            pass
        if shp is None:
            print(f"  [未找到] {name}（模板中不存在，跳过）")
            continue

        print(f"  [处理] {name}  strategy={strategy}")

        # 特殊路由：Chart 63
        if strategy == "bar_stacked_temp_range":
            chart_data = _calc_chart63_data(rows)
            _write_chart63(shp, chart_data)
            continue

        # 特殊路由：mean_extraction_filtered（OLE 粘贴）
        if strategy == "mean_extraction_filtered":
            content = _build_content(spec, rows, gpt_enabled, mc_model)
            try:
                L = float(_com_get(shp, "Left", 0))
                T = float(_com_get(shp, "Top", 0))
                W = float(_com_get(shp, "Width", 190))
                H = float(_com_get(shp, "Height", 100))
                shp.Delete()
            except Exception as _e:
                print(f"    [警告] 读取/删除模板 chart shape 失败: {_e}")
                continue

            try:
                mc_cell, n_chart_rows = _prepare_apparel_chart_data(
                    mc_sht, content, anchor_offset=offset,
                )
                offset += len(content.splitlines()) + 5
            except Exception as _e:
                print(f"    [警告] 写入 Excel 临时数据失败: {_e}")
                continue

            try:
                _tmp_chart = make_chart_for_apparel(
                    mc_cell, new_slide, Left=L, Top=T, Width=W, Height=H,
                    n_rows=n_chart_rows,
                )
            except Exception as _e:
                print(f"    [警告] make_chart_for_apparel 失败: {_e}")
                continue

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
            continue

        # shared_info 复用逻辑（gpt_respondent_info 且 caller 已传入）
        if strategy == "gpt_respondent_info" and shared_info:
            ok = _write_text(shp, shared_info)
            if not ok:
                print(f"    [警告] _write_text（shared_info）返回 False")
            # TextBox 24 Height 动态拉长（同旧 make_apparel_slide 逻辑）
            n_samples = max(0, len(rows) - 1)
            if ok and n_samples > 5 and name == "TextBox 24":
                try:
                    original_h = float(_com_get(shp, "Height", 0))
                    new_h = original_h * (n_samples + 1) / 6
                    shp.Height = new_h
                    print(f"    [TextBox 24] Height 拉长（shared）: {original_h:.1f} → {new_h:.1f}")
                except Exception as _e:
                    print(f"    [TextBox 24] 拉长 Height 失败: {_e}")
            continue

        # 通用文本类
        content = _build_content(spec, rows, gpt_enabled, mc_model)
        # Bug B 修复：2-run 标签 strategy 走专用写入，保留模板红色数值格式
        if strategy in _TWO_RUN_STRATEGIES:
            # 允许 SHAPES 配置 override 默认 title_size/value_size/title_color/
            # value_color/same_line（per-shape 视觉契约，例如 RR 53/55 = 11pt/24pt 白色）
            _params = spec.get("params", {}) or {}
            _kw = {}
            for _k in ("title_size", "value_size",
                       "title_color", "value_color", "same_line"):
                if _k in _params:
                    _kw[_k] = _params[_k]
            ok = _write_two_run_label(shp, content, **_kw)
            if not ok:
                print(f"    [警告] _write_two_run_label 返回 False，回退 _write_text")
                ok = _write_text(shp, content)
        else:
            ok = _write_text(shp, content)
        if not ok:
            print(f"    [警告] _write_text 返回 False")

        # 染色
        # Bug A 修复：gpt_strengths_bullet / gpt_drawbacks_bullet 走专用函数，
        # 保留首行模板原色 + 只 bold 关键词；gpt_prompted 仍走通用函数。
        # fix6-b：gpt_drawbacks_bullet 额外把末段（"建议改为..."）字号 bump 到 16
        # （模板 TextBox 26 末尾结构性强调，size 14 → 16）。
        if ok and strategy == "gpt_drawbacks_bullet":
            _apply_apparel_bullet_color(shp, bump_last_para_size=16.0)
        elif ok and strategy == "gpt_strengths_bullet":
            _apply_apparel_bullet_color(shp)
        elif ok and strategy == "gpt_prompted":
            _apply_keyword_color(shp)

        # TextBox 24 Height 动态拉长
        n_samples = max(0, len(rows) - 1)
        if ok and n_samples > 5 and name == "TextBox 24":
            try:
                original_h = float(_com_get(shp, "Height", 0))
                new_h = original_h * (n_samples + 1) / 6
                shp.Height = new_h
                print(f"    [TextBox 24] Height 拉长: {original_h:.1f} → {new_h:.1f}")
            except Exception as _e:
                print(f"    [TextBox 24] 拉长 Height 失败: {_e}")

    return offset


# ---------------------------------------------------------------------------
# Public API — 双页入口（B 方案）
# ---------------------------------------------------------------------------

def make_apparel_p13_slide(mc_sht, mc_ppt, mc_slide, sample_name: str,
                           mc_gpt: str = "n", mc_model: str = _MODEL,
                           trace_path: str | None = "acceptance/apparel_trace.jsonl"):
    """生成 apparel page 13（数据图表型，22 shapes）.

    Clone src/Template 2.1.pptx 合并后 slide 20（_TEMPLATE_P13_SLIDE），
    然后按 APPAREL_P13_SHAPES 写入评分标签、4 类 chart、Chart 63 温度区间、
    累计跑量、训练定位、受试者信息。

    trace_path: jsonl 日志路径，供 ppt-acceptance-check L4 行为层断言用。
                传 None 关闭 trace（一般只在 unit test 时）。
                默认 "acceptance/apparel_trace.jsonl"，每次调用 append 进去。

    Returns 新 slide 对象。
    """
    global _TRACE
    _trace_owned = False
    if trace_path and _TraceLogger is not None and _TRACE is None:
        try:
            _TRACE = _TraceLogger(trace_path)
            _trace_owned = True
        except Exception:
            _TRACE = None

    try:
        gpt_enabled = (mc_gpt == "y")
        _trace_event("p13_start", sample=sample_name, gpt_enabled=gpt_enabled)
        print(f"\n[apparel-p13] 开始生成  sample={sample_name}  gpt={'开启' if gpt_enabled else '关闭'}")
        rows = _xlwings_to_rows(mc_sht)
        print(f"[apparel-p13] 读取问卷数据：{len(rows)} 行，{len(rows[0]) if rows else 0} 列")

        X = mc_ppt.Slides.Count + 1
        print(f"[apparel-p13] 克隆模板第 {_TEMPLATE_P13_SLIDE} 页 → 新建第 {X} 页...")
        mc_ppt.Slides(_TEMPLATE_P13_SLIDE).Copy()
        time.sleep(_COPY_PASTE_DELAY)
        new_slide = mc_ppt.Slides.Paste(X)
        time.sleep(1.0)

        _overlay = None
        if show_gpt_waiting_overlay is not None:
            try:
                _overlay = show_gpt_waiting_overlay(new_slide)
            except Exception:
                pass

        print(f"[apparel-p13] 逐 shape 写入，共 {len(APPAREL_P13_SHAPES)} 个...")
        _write_shapes_to_slide(
            new_slide, APPAREL_P13_SHAPES, rows, mc_sht,
            gpt_enabled, mc_model, chart_anchor_offset=50,
        )

        if _overlay is not None and remove_gpt_waiting_overlay is not None:
            try:
                remove_gpt_waiting_overlay(_overlay)
            except Exception:
                pass

        print(f"[apparel-p13] 完成！新页在第 {new_slide.SlideIndex} 页")
        _trace_event("p13_end", slide_index=new_slide.SlideIndex)
        return new_slide
    finally:
        if _trace_owned and _TRACE is not None:
            try:
                _TRACE.close()
            except Exception:
                pass
            _TRACE = None


def make_apparel_p14_slide(mc_sht, mc_ppt, mc_slide, sample_name: str,
                           mc_gpt: str = "n", mc_model: str = _MODEL,
                           shared_info: str = "",
                           trace_path: str | None = "acceptance/apparel_trace.jsonl"):
    """生成 apparel page 14（文字 bullet 型，7 shapes）.

    Clone src/Template 2.1.pptx 合并后 slide 21（_TEMPLATE_P14_SLIDE），
    然后按 APPAREL_P14_SHAPES 写入优点/缺点 bullet、受试者信息。

    shared_info: 可由 caller 传入 p13 阶段已生成的受试者信息字符串（复用，省一次 GPT）。
                 空字符串时 p14 自己调 GPT 生成。
    trace_path:  jsonl 日志路径，同 p13；append 模式与 p13 共用一份日志。

    Returns 新 slide 对象。
    """
    global _TRACE
    _trace_owned = False
    if trace_path and _TraceLogger is not None and _TRACE is None:
        try:
            _TRACE = _TraceLogger(trace_path)
            _trace_owned = True
        except Exception:
            _TRACE = None

    try:
        gpt_enabled = (mc_gpt == "y")
        _trace_event("p14_start", sample=sample_name, gpt_enabled=gpt_enabled,
                     shared_info_reused=bool(shared_info))
        print(f"\n[apparel-p14] 开始生成  sample={sample_name}  gpt={'开启' if gpt_enabled else '关闭'}")
        rows = _xlwings_to_rows(mc_sht)
        print(f"[apparel-p14] 读取问卷数据：{len(rows)} 行，{len(rows[0]) if rows else 0} 列")

        X = mc_ppt.Slides.Count + 1
        print(f"[apparel-p14] 克隆模板第 {_TEMPLATE_P14_SLIDE} 页 → 新建第 {X} 页...")
        mc_ppt.Slides(_TEMPLATE_P14_SLIDE).Copy()
        time.sleep(_COPY_PASTE_DELAY)
        new_slide = mc_ppt.Slides.Paste(X)
        time.sleep(1.0)

        _overlay = None
        if show_gpt_waiting_overlay is not None:
            try:
                _overlay = show_gpt_waiting_overlay(new_slide)
            except Exception:
                pass

        print(f"[apparel-p14] 逐 shape 写入，共 {len(APPAREL_P14_SHAPES)} 个...")
        _write_shapes_to_slide(
            new_slide, APPAREL_P14_SHAPES, rows, mc_sht,
            gpt_enabled, mc_model, chart_anchor_offset=150,
            shared_info=shared_info,
        )

        if _overlay is not None and remove_gpt_waiting_overlay is not None:
            try:
                remove_gpt_waiting_overlay(_overlay)
            except Exception:
                pass

        print(f"[apparel-p14] 完成！新页在第 {new_slide.SlideIndex} 页")
        _trace_event("p14_end", slide_index=new_slide.SlideIndex)
        return new_slide
    finally:
        if _trace_owned and _TRACE is not None:
            try:
                _TRACE.close()
            except Exception:
                pass
            _TRACE = None


# ---------------------------------------------------------------------------
# 就地覆写入口（供调试/验证，不 Clone 新 slide）
# ---------------------------------------------------------------------------

def rewrite_apparel_p13_slide(mc_sht, mc_ppt, slide_index: int,
                               mc_gpt: str = "n", mc_model: str = _MODEL,
                               trace_path: str | None = "acceptance/apparel_trace.jsonl"):
    """就地覆写已有 PPT 第 slide_index 页的 apparel p13 内容.

    与 make_apparel_p13_slide 的区别：不 Clone 新 slide，直接在现有页上重写所有
    APPAREL_P13_SHAPES（除 skip 类）。用于验证修复后的代码而不改变 PPT 总页数。

    trace_path: jsonl 日志路径，append 模式（与 make_apparel_p13_slide 共用格式）。
    Returns 被覆写的 slide 对象。
    """
    global _TRACE
    _trace_owned = False
    if trace_path and _TraceLogger is not None and _TRACE is None:
        try:
            _TRACE = _TraceLogger(trace_path)
            _trace_owned = True
        except Exception:
            _TRACE = None

    try:
        gpt_enabled = (mc_gpt == "y")
        sample_name = mc_sht.name
        _trace_event("p13_rewrite_start", slide_index=slide_index,
                     sample=sample_name, gpt_enabled=gpt_enabled)
        print(f"\n[apparel-p13-rewrite] 就地覆写第 {slide_index} 页  sample={sample_name}  gpt={'开启' if gpt_enabled else '关闭'}")

        rows = _xlwings_to_rows(mc_sht)
        print(f"[apparel-p13-rewrite] 读取问卷数据：{len(rows)} 行，{len(rows[0]) if rows else 0} 列")

        target_slide = mc_ppt.Slides(slide_index)
        print(f"[apparel-p13-rewrite] 目标 slide index={target_slide.SlideIndex}")

        print(f"[apparel-p13-rewrite] 逐 shape 写入，共 {len(APPAREL_P13_SHAPES)} 个...")
        _write_shapes_to_slide(
            target_slide, APPAREL_P13_SHAPES, rows, mc_sht,
            gpt_enabled, mc_model, chart_anchor_offset=50,
        )

        print(f"[apparel-p13-rewrite] 完成！第 {slide_index} 页已就地覆写")
        _trace_event("p13_end", slide_index=slide_index)
        return target_slide
    finally:
        if _trace_owned and _TRACE is not None:
            try:
                _TRACE.close()
            except Exception:
                pass
            _TRACE = None


# ---------------------------------------------------------------------------
# 单独调试入口：连接已打开的 Excel + PPT，只跑 apparel 这一页
# 用法：
#   python src/apparel_ppt.py p13   — 生成 page 13（数据图表型，Clone 新页）
#   python src/apparel_ppt.py p14   — 生成 page 14（文字 bullet 型，Clone 新页）
#   python src/apparel_ppt.py       — 生成 page 13 + page 14（双页，等同生产）
#
#   --overwrite-slide N [--gpt y]   — 就地覆写第 N 页（p13），不追加新页
#       用于验证修复后的代码而不改变 PPT 总页数。
#
# 前置：
#   1) Excel 已打开，含 sheet 名含"问卷"的工作表
#   2) PPT 已打开（脚本连接活动 Presentation，需已包含 p13/p14 模板页）
#   3) 模板合并：先运行 python _archive/2026-05-27-debug-cleanup/scripts/merge_apparel_template.py
#      把 template/apparel-page13-14-template.pptx 的 slide 13/14
#      合并到 src/Template 2.1.pptx（合并后 slide 20=p13、slide 21=p14）
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import win32com.client
    import xlwings

    _proj_root = str(Path(__file__).resolve().parent.parent)

    # 解析命令行参数
    # 支持：p13 / p14 / both / --overwrite-slide N [--gpt y]
    _args = sys.argv[1:]
    _overwrite_slide = None
    _mc_gpt = "n"

    # 提取 --overwrite-slide N
    if "--overwrite-slide" in _args:
        _idx = _args.index("--overwrite-slide")
        try:
            _overwrite_slide = int(_args[_idx + 1])
            _args = [a for i, a in enumerate(_args) if i not in (_idx, _idx + 1)]
        except (IndexError, ValueError):
            print("用法: --overwrite-slide N  （N 为 1-based slide index）")
            sys.exit(1)

    # 提取 --gpt y/n
    if "--gpt" in _args:
        _gidx = _args.index("--gpt")
        try:
            _mc_gpt = _args[_gidx + 1].lower()
            _args = [a for i, a in enumerate(_args) if i not in (_gidx, _gidx + 1)]
        except IndexError:
            pass

    _page_arg = _args[0].lower() if _args else "both"
    if _page_arg not in ("p13", "p14", "both"):
        print(f"用法: python src/apparel_ppt.py [p13|p14]  （默认生成双页）")
        _page_arg = "both"

    # 连接活动 PPT（需已打开合并后的 template）
    try:
        mc_app = win32com.client.GetActiveObject("PowerPoint.Application")
        mc_ppt = mc_app.ActivePresentation
        print(f"[debug] 连接活动 PPT: {mc_ppt.Name}  ({mc_ppt.Slides.Count} 页)")
    except Exception as _e:
        print(f"[debug] 无活动 PPT，尝试打开合并后模板...")
        mc_app = win32com.client.Dispatch("PowerPoint.Application")
        mc_app.DisplayAlerts = 0
        mc_app.Visible = True
        _template_path = _proj_root + r"\src\Template 2.1.pptx"
        mc_ppt = mc_app.Presentations.Open(_template_path)
    mc_slide = mc_ppt.Slides(mc_ppt.Slides.Count)

    # 连接已打开的 Excel（问卷 sheet）；若无活动 App，尝试打开 apparel 数据文件
    _xl_opened_by_us = False
    mc_book = None
    try:
        mc_book = xlwings.books.active
    except Exception:
        mc_book = None

    if mc_book is None:
        # 回退：用 win32com 直接打开 apparel 源数据
        print("[debug] 无活动 Excel，尝试打开 template/source data-apparel.xlsx ...")
        import os as _os
        _xl_app_fallback = win32com.client.Dispatch("Excel.Application")
        _xl_app_fallback.Visible = True
        _xl_app_fallback.DisplayAlerts = False
        _xl_path = _os.path.abspath(
            _os.path.join(_proj_root, "template", "source data-apparel.xlsx")
        )
        _xl_wb = _xl_app_fallback.Workbooks.Open(_xl_path)
        import time as _time
        _time.sleep(1.5)
        # 用 xlwings 重新接管（win32com 打开后 xlwings 可见）
        try:
            mc_book = xlwings.books.active
        except Exception:
            mc_book = None
        _xl_opened_by_us = True

    mc_sht = None
    if mc_book is not None:
        for s in mc_book.sheets:
            if "问卷" in s.name:
                mc_sht = s
                break
    if mc_sht is None:
        print("未找到包含'问卷'的 sheet，请先在 Excel 打开数据文件")
        sys.exit(1)

    sample_name = mc_sht.name

    print(f"[debug] overwrite_slide: {_overwrite_slide}")
    print(f"[debug] page_arg: {_page_arg}")
    print(f"[debug] gpt: {_mc_gpt}")
    print(f"[debug] sample:   {sample_name}")
    print(f"[debug] sheet:    {mc_sht.name}")
    print(f"[debug] slides:   {mc_ppt.Slides.Count}")
    print(f"[debug] _TEMPLATE_P13_SLIDE = {_TEMPLATE_P13_SLIDE}")
    print(f"[debug] _TEMPLATE_P14_SLIDE = {_TEMPLATE_P14_SLIDE}")

    # --- 就地覆写模式 ---
    if _overwrite_slide is not None:
        rewrite_slide = rewrite_apparel_p13_slide(
            mc_sht, mc_ppt, _overwrite_slide,
            mc_gpt=_mc_gpt, mc_model=_MODEL,
            trace_path="acceptance/apparel_trace.jsonl",
        )
        print(f"[debug] 就地覆写完成！第 {rewrite_slide.SlideIndex} 页已更新（总页数仍为 {mc_ppt.Slides.Count}）")
        print(f"[debug] trace 落盘：acceptance/apparel_trace.jsonl")
        sys.exit(0)

    # --- Clone 新页模式（原有行为）---
    if _page_arg in ("p13", "both"):
        new_slide = make_apparel_p13_slide(
            mc_sht, mc_ppt, mc_slide, sample_name,
            mc_gpt=_mc_gpt, mc_model=_MODEL,
        )
        print(f"[debug] p13 完成！新页在第 {new_slide.SlideIndex} 页")
        mc_slide = new_slide

    if _page_arg in ("p14", "both"):
        new_slide = make_apparel_p14_slide(
            mc_sht, mc_ppt, mc_slide, sample_name,
            mc_gpt=_mc_gpt, mc_model=_MODEL,
        )
        print(f"[debug] p14 完成！新页在第 {new_slide.SlideIndex} 页")

    print(f"[debug] 注意：模板文件未保存，请手动检查后关闭（不要保存）")
