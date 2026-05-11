# Function for import

# ★ ★ ★ ★ ★  公共信息（每个文件都复制一遍。。）  ★ ★ ★ ★ ★


# ----- 导入第三方Module ------
import ast
from datetime import datetime
import win32com.client
import xlwings

import os, time, random, re
from datetime import datetime



# ----- Main()独占：导入Class 和 Function 文件 ------


### 先解决路径问题
##import sys                           # 导入sys模块，来调取path 系统默认路径
##mc_path = os.getcwd()
##sys.path.append(mc_path)     # 将我的mod路径，添加至 系统默认路径
###print(sys.path)  # 检查下路径有没有问题



# 再导入模块：导入模块中所有的类的语句为：【 from modulname import * 】

# 2025升级，同文件夹目录下采用 【.方式】来import  ！！
from .Class_030 import *
#from Function_030 import *



#import time, os
#import os












# 全局变量 color + delay ---------------------



delay = 1     # 为了解决宇昂以及部分老旧电脑卡顿的问题，人工设置延时
#time.sleep(random.random()*delay)


black = 0                      #' 黑色'
white = 16777215              #' 白色'
gray = 14277081               #' 浅灰色'
light_gray = 15921906         #' 浅白灰色' （接近透明）

red = 255                     #' 红色'
green = 5287936               #' 绿色 '
dark_blue = 6299648           #' 墨蓝色 '
light_blue = 15773696         #' 天蓝色'

yellow = 65535                #' 亮黄 '
orange = 16727809             #' 橙色 '
purple = 10498160             #' 紫色 '



# # 全局变量 模板页面，例如，篮球矩阵模板是Slide(5)....  ---------------------
# dic_matrix ={
#
#     '减震回弹矩阵-篮球': 5 , \
#     '弯折扭转矩阵-篮球': 6 , \
#     '减震回弹矩阵-跑步': 7 , \
#     '弯折扭转矩阵-跑步': 8 , \
#     '步长重心振幅矩阵-跑步':12 , \
#      '重量厚度矩阵-跑步': 13     #注意，从12页开始，不要再往中间插入ppt了，直接在后面新增，不然main函数的代码（页码）都需要调整
#     }

# 使用全局变量模块，以后全局变量都放到 Global 文件中去吧
from .Global_var_030 import *
dic_matrix = get_value('dic_matrix')

# ★ ★ ★ ★ ★  公共信息（每个文件都复制一遍。。）  ★ ★ ★ ★ ★



# global 参数统一修改
main_model = "openai/gpt-5.4"    #2026-03-11 更新  https://openrouter.ai/openai
mini_model = "openai/gpt-5-mini"









## 【函数 function  函数 function  函数 function  函数 function  函数 function  函数 function  】
## ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★



# 2025，更新GPT的API和函数重构，没太大变化，只是更简洁了。
# 现在改用 OpenRouter 代理

from openai import OpenAI
import os
import traceback
import httpx
from openai import OpenAI
import os


def detect_system_proxy():
    """
    自动检测系统代理设置（Windows）

    Returns:
        str: 代理URL（如 "http://127.0.0.1:7897"），未检测到则返回 None
    """
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        )
        proxy_enabled = winreg.QueryValueEx(key, "ProxyEnable")[0]
        if proxy_enabled:
            proxy_server = winreg.QueryValueEx(key, "ProxyServer")[0]
            winreg.CloseKey(key)
            if proxy_server:
                # 如果代理地址不包含协议前缀，添加 http://
                if not proxy_server.startswith("http"):
                    proxy_server = f"http://{proxy_server}"
                print(f"[代理] 检测到系统代理: {proxy_server}")
                return proxy_server
        winreg.CloseKey(key)
    except Exception as e:
        print(f"[代理] 检测失败: {e}，将不使用代理")
    return None


# ===== 只改这里：换成 OpenRouter =====
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")

##print(OPENROUTER_API_KEY)
##a = input('.........................')

# 自动检测代理
_proxy_url = detect_system_proxy()
_http_client = httpx.Client(proxy=_proxy_url, timeout=30) if _proxy_url else httpx.Client(timeout=30)

client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key=OPENROUTER_API_KEY,
    http_client=_http_client,
    default_headers={
        "HTTP-Referer": "http://localhost",   # 或你的站点/项目地址
        "X-Title": "my-app",                  # 任意应用名
    })
# =====================================

messages = []     # 超级较真的 GPT-5.1 强迫我修正了 Claude 的小 bug


def GPT_5(mc_prompt, model):

    def initialize_dialogue():
        """
        首次对话或重置时调用：
        清空历史，并加入 system 提示
        """
        messages.clear()
        messages.extend([
            {
                "role": "system",
                "content": "你是一个非常专业、非常乐于助人的AI助手。"
            }
        ])

    def user_prompt(prompt):
        """添加用户消息（不带缓存标记）"""
        messages.append({"role": "user", "content": prompt})

    def continue_dialogue(completion):
        """添加助手回复（不带缓存标记）"""
        messages.append({'role': 'assistant', 'content': completion})

    def mark_history_as_cacheable():
        """将所有历史消息（除了最新的用户消息）标记为可缓存"""
        for i in range(len(messages) - 1):
            if 'cache_control' not in messages[i]:
                messages[i]['cache_control'] = {'type': 'ephemeral'}

    def dialogue(prompt):
        """发送对话请求"""
        print('      GPT 服务器请求中,需要一定的响应时间,请耐心等待答复.....\n')
        try:
            response = client.chat.completions.create(
                model=model,
                messages=prompt
            )

            print('                 成功获取GPT答复，Congratulations !!!\n\n\n')
            reply = response.choices[0].message.content

        except Exception as e:
            traceback.print_exc()
            reply = f"请求失败：{e}"
##            print('报错啦，请检查function.py 文件中的GPT_5 函数', e)
##            reply = '无法与服务器建立连接，请检查VPN设置，然后再尝试！'

        return reply

    # ===== 核心逻辑 =====

    if not messages:
        initialize_dialogue()
    else:
        mark_history_as_cacheable()

    user_prompt(mc_prompt)
    completion = dialogue(messages)
    continue_dialogue(completion)

    return completion


def get_messages():
    '''超级较真的 GPT-5.1 教我如何规范写代码。。。。。。'''
    return messages


# ========= 聊天室函数 =========
def chat_room(model="openai/gpt-5-mini"):  # 注意：OpenRouter的模型名需要带前缀
    """
    最简单的命令行聊天室：
    - 【SSRC（用户）：】提示输入
    - 输入内容为空或输入 '退出' / 'exit' / 'quit' 时结束
    """
    print("=== 简易 GPT 聊天室 已启动 ===")
    print("输入 '退出' / 'exit' / 'quit' 可结束对话。\n")

    while True:
        user_input = input("【SSRC（用户）：】")

        if user_input.strip() in ("退出", "exit", "quit", ""):
            print("\n聊天结束，再见！")
            break

        print("（GPT 正在思考中.....）\n")
        reply = GPT_5(user_input, model)
        print(f"【GPT（assistant）：】{reply}\n")








def _to_float(value):
    """安全转换为浮点数，失败返回 None。"""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip())
        except (ValueError, AttributeError):
            return None
    return None


def build_questionnaire_summary_prompt(sample_name, respondent_count, mean_rows, high_rows):
    """构建问卷汇总页（结论页）统一总结的提示词。"""
    mean_text = '；'.join([f"{k}={v:.1f}" for k, v in mean_rows])
    high_text = '；'.join([f"{k}={v:.1f}%" for k, v in high_rows])

    mc_prompt = (
        f"你是一名产品评测分析师。请基于以下问卷统计结果，对{sample_name}输出一段统一总结。"
        f"样本量：{respondent_count}。"
        f"均值统计：{mean_text}。"
        f"高分占比统计（>=8分）：{high_text}。"
        "要求：1）只输出一段自然语言结论；2）总字数不超过180字；3）不要输出分点、标题或分析过程；"
        "4）语气客观，突出主要优势和需要改进的点。"
    )

    return mc_prompt


def questionnaire_summary_slide(mc_sht, mc_ppt, mc_slide, sample_name, mc_gpt='n', mc_model='gpt-5-mini'):
    """在【5】与【6】之间生成问卷汇总页（轻量统计+图表+统一总结）。"""

    if mc_sht is None or '问卷' not in mc_sht.name:
        return mc_slide

    try:
        mc_cell0 = get_range(mc_sht)
        if mc_cell0 is None:
            return mc_slide

        raw_data = mc_cell0.api.CurrentRegion.Value
        clean_data = parse_survey_data(raw_data)

        if len(clean_data) < 2 or len(clean_data[0]) < 2:
            return mc_slide

        header = list(clean_data[0])
        metric_names = [str(x) for x in header[1:]]
        metric_values = []

        for col_idx in range(1, len(header)):
            col_data = []
            for row in clean_data[1:]:
                if col_idx < len(row):
                    num = _to_float(row[col_idx])
                    if num is not None:
                        col_data.append(num)
            metric_values.append(col_data)

        mean_rows = []
        high_rows = []
        for name, values in zip(metric_names, metric_values):
            if not values:
                continue
            avg = sum(values) / len(values)
            high_ratio = (sum(1 for x in values if x >= 8) / len(values)) * 100
            mean_rows.append((name, round(avg, 1)))
            high_rows.append((name, round(high_ratio, 1)))

        if not mean_rows:
            return mc_slide

        mean_rows.sort(key=lambda x: x[1], reverse=True)
        high_rows.sort(key=lambda x: x[1], reverse=True)

        # Template 2.1：第14页为空白问卷模板，第15页为标准问卷模板，优先使用标准模板
        summary_template_idx = 15 if len(mc_ppt.Slides) >= 15 else (14 if len(mc_ppt.Slides) >= 14 else 4)
        mc_ppt.Slides(summary_template_idx).Select()
        mc_ppt.Slides(summary_template_idx).Copy()
        X = len(mc_ppt.Slides) + 1
        time.sleep(random.random() * delay)
        mc_slide = mc_ppt.Slides.Paste(X)

        # 不清空模板内容，最大化复用现有版式；标题采用覆盖方式，避免新增 shape 干扰模板
        try:
            title_candidates = []
            for shp in mc_slide.Shapes:
                if shp.HasTextFrame and shp.TextFrame.HasText:
                    txt = str(shp.TextFrame.TextRange.Text or '').strip()
                    if txt:
                        title_candidates.append(shp)
            if title_candidates:
                title_shape = sorted(title_candidates, key=lambda s: (s.Top, s.Left))[0]
                title_shape.TextFrame.TextRange.Text = '问卷汇总分析'
            else:
                Title_1(mc_slide, Left=15, Top=15, Text='问卷汇总分析')
        except Exception:
            Title_1(mc_slide, Left=15, Top=15, Text='问卷汇总分析')

        i = mc_cell0.api.CurrentRegion.Rows.Count
        base_cell = mc_cell0.offset(row_offset=i + 8, column_offset=0)

        mean_table = [('指标', '均值')] + [(k, v) for k, v in mean_rows]
        base_cell.value = mean_table

        high_cell = base_cell.offset(row_offset=len(mean_table) + 2, column_offset=0)
        high_table = [('指标', '高分占比(%)')] + [(k, v) for k, v in high_rows]
        high_cell.value = high_table

        _chart1 = make_chart_for_questionnaire(base_cell, mc_slide, Left=40, Top=90, Width=360, Height=170)
        _chart2 = make_chart_for_questionnaire(high_cell, mc_slide, Left=500, Top=90, Width=360, Height=170)

        # 清理临时数据：先删 Excel chart 对象（切断 OLE 热链接），再删数据行。
        # DisplayAlerts=False 压制删行时可能出现的"错误公式引用"弹窗。
        _xl_app_api = None
        try:
            _xl_app_api = mc_sht.book.app.api
            _xl_app_api.DisplayAlerts = False
        except Exception:
            pass
        try:
            _chart2.delete()
        except Exception:
            pass
        try:
            high_cell.expand().api.EntireRow.Delete()
        except Exception:
            pass
        try:
            _chart1.delete()
        except Exception:
            pass
        try:
            base_cell.expand().api.EntireRow.Delete()
        except Exception:
            pass
        try:
            if _xl_app_api is not None:
                _xl_app_api.DisplayAlerts = True
        except Exception:
            pass

        respondent_count = max(len(clean_data) - 1, 0)
        top_mean = '、'.join([f"{k}{v:.1f}分" for k, v in mean_rows[:3]])
        top_high = '、'.join([f"{k}{v:.1f}%" for k, v in high_rows[:3]])
        stat_hint = f"样本量：{respondent_count}人；均值Top3：{top_mean}；高分占比Top3：{top_high}"
        Text_1(mc_slide, Left=40, Top=275, Text=stat_hint, scale=0.95)

        if mc_gpt == 'y':
            prompt = build_questionnaire_summary_prompt(sample_name, respondent_count, mean_rows, high_rows)
            summary_text = GPT_5(prompt, model=mc_model)
        else:
            summary_text = f"问卷共回收{respondent_count}份，{mean_rows[0][0]}等维度得分较高；建议持续优化{mean_rows[-1][0]}相关体验，以提升整体实战反馈。"

        result = Result_Bullet_small(mc_slide, Left=40, Top=330, Text=summary_text, scale=0.95)
        result.tr.ParagraphFormat.Bullet.Visible = 0

    except Exception as e:
        print(f"问卷汇总页生成失败：{e}")

    return mc_slide



# --------------------------- prompt 定制函数 ---------------------------

# 定义这个函数，用来定制 prompt，并请求一次GPT服务器获取一个 reply（completion）  ///  nonono  需要严格区分函数功能，聊天功能已经在GPT_5中严格实现了，这个只需要用来生成【问题】！@！@！@！


def gen_mc_prompt(**all_info):  # mc_cell 也装入 dict变量 all_info中算了~

    # 我需要针对每个图表 【图表1 / 2 / 3 ...】 都调用一次GPT函数，获取结论，因此要把 range对象输入到该函数中，使用*cell_list作为参数
    # 更新：知晓字典的好处之后，直接定义一个dict参数，然后在函数外部调整就行，这样就不用反复定义多个变量，多好 ~

    # # 去年写的人工GPT函数还能用。。。
    # mc_cell.select()
    # mc_sht = mc_cell.sheet  #不考虑那么多了， mc_sht 就是当前sheet，会一直切换遍历所有sheet
    # mc_book = mc_sht.book

    # 逐步开始构建 prompt //  注意，每次请求GPT，都需要将 prompt 清零，以免信息冗余
    #mc_prompt = ''


    # 【评判标准】
    # mc_prompt += '【测试标准】' + '\r' + str(all_info['测试标准'])

    mc_prompt = (

        f"【测试标准】\n{all_info['测试标准']}\n  {all_info['参考文本']}\n\n"

        f"【你的任务】\n 下面是我的实验测试数据，请帮我按照上述【测试标准】，对【{all_info['测试样品名称']}】的测试结果进行分析，并重点将它与{all_info['对比样品名称']}进行横向对比，并给出你的分析结论。\n"

        f"你只要分析已提供给你的数据即可，不要擅自进行推测、编造数据，记住，数据的准确性和严谨性是你的首要目标。\n "

        f"注意，你给出的分析结论需要严格按照【参考文本】的格式、语调、陈述方法，直接给出结论即可，你不要再重复一遍上面的判断依据 。\n "       #注意，你在给出分析结论答复时，需要严格按照【参考文本】的格式、语调、陈述方式，给出同样精炼和简洁的回答，并且你回答的总字数不能超过 160 个汉字。"

        f"切记，不要展示你的分析过程，直接给出结论！！\n "

        f"接下来是测试数据部分。\n【测试数据】\n{all_info['测试数据']}"

        f"针对上面的数据，给出你的结论，并按（1，2，3...）这种条目的形式展示出来，结论要简洁、严谨、准确无误。"

        f"⚠️ 重要：必须分析上述测试数据中的**所有指标**，一个都不能遗漏！"

        )


##    print('下面是目前传递给GPT服务器的文本信息/数据，请仔细检查\r')
##    print(mc_prompt+ '\r\r\r')
##    print(type(mc_prompt))
##    print(len(mc_prompt))
##    print(all_info['测试数据'])



    return mc_prompt #GPT_reply






# 问卷部分的 prompt
# 需要传入【runner】【temp_raw_data】2个参数。
def gen_questionnaire_prompt(runner,temp_raw_data):

    mc_prompt =(

                f'我回收了一份问卷，里面包含多名运动员的试穿反馈。这是第【{runner+1}】名运动员的试穿反馈。'
                    f'你帮我总结下这款产品的优点和缺点，总字数不要超过140字。'
                      f'在你的答复中，你认为需要重点标记的内容，用【】标记出来。这样更方便我阅读。'
                       f'注意，你要严格按照我发给你的模板，将你分析的内容分段展示，模板如下：'
                        '''
                        【优点】
                        1、
                        2、
                        ...
                         【缺点】
                         1、
                         2、
                         ...

                        '''

                        f'下面是第【{runner+1}】名运动员的试穿反馈：\n {str(temp_raw_data)}'

                          )


    return mc_prompt










# 总结分析所有信息 prompt
#   - sample_name: 必填
#   - sheet_summaries: 各测试方法 sheet 的 GPT 结论（list[str]，按顺序）
#   - questionnaire_summaries: 各运动员问卷反馈的 GPT 结论（list[str]）
# 显式注入先前结论，不再依赖 messages 上下文（plan4：更鲁棒、可被 clamp_text 控制结构）
def gen_result_prompt(sample_name, sheet_summaries=None, questionnaire_summaries=None):

    sheet_block = ""
    if sheet_summaries:
        sheet_block = "\n\n".join(
            f"【测试方法 {i+1}】\n{str(s).strip()}"
            for i, s in enumerate(sheet_summaries) if s and str(s).strip()
        )

    qn_block = ""
    if questionnaire_summaries:
        qn_block = "\n\n".join(
            f"【运动员 {i+1}】\n{str(s).strip()}"
            for i, s in enumerate(questionnaire_summaries) if s and str(s).strip()
        )

    info_section = ""
    if sheet_block or qn_block:
        info_section = (
            "【已知信息】\n"
            "以下是先前已经分析得出的结论，请直接基于此做总结，"
            "不要再展开重复分析：\n\n"
        )
        if sheet_block:
            info_section += sheet_block + "\n\n"
        if qn_block:
            info_section += qn_block + "\n\n"

    mc_prompt = (
        f"{info_section}"
        f"【你的任务】\n"
        f"请综合上述信息，对【{sample_name}】给出一份最终评测总结。\n\n"
        f"严格按以下三段式输出，每段用编号条目（1、2、3...）：\n"
        f"【优点】\n1、...\n2、...\n3、...\n\n"
        f"【缺点】\n1、...\n2、...\n3、...\n\n"
        f"【修改建议】\n1、...\n2、...\n3、...\n\n"
        f"硬性要求：\n"
        f"- 三段段头【优点】/【缺点】/【修改建议】必须保留，单独占一行。\n"
        f"- 每段至少 2-3 条；若确实无内容，保留段头并写\"暂无显著XX\"。\n"
        f"- **关键词标注规则**（必须严格遵守，外层程序按括号类型自动染色）：\n"
        f"    · 在【优点】段，把核心优势/性能词用半角尖括号 < > 包起来，例如：<回弹性能>。\n"
        f"    · 在【缺点】段，把核心问题/不足词用半角方括号 [ ] 包起来，例如：[支撑不足]。\n"
        f"    · 在【修改建议】段，把建议关键词用半角圆括号 ( ) 包起来，例如：(加厚中底)。\n"
        f"    · 仅包关键词本身、不含任何标点；段头【优点】/【缺点】/【修改建议】保持中文方括号 【】 不变。\n"
        f"    · 每条至少 1 个关键词标注；同一条结论可标多个关键词。\n"
        f"- 总字数 ≤ 270 汉字，总行数 ≤ 12 行（含 3 段头），尽量写满让页面饱满。\n"
        f"- 只能基于已有数据，不允许推测或编造。\n"
        f"- 直接输出结论，不重述题面、不展示分析过程、不输出空行、不使用 Markdown。\n"
    )

    return mc_prompt






# --------------------------- prompt 定制函数 ---------------------------
















# --------------------------- 弹窗函数 ---------------------------

##    GPT_弹窗选择按钮，替换input  ///

##   没有Claude辅助，根本搞定不了，该功能也属于画蛇添足，先放一放。。。最后再搞
import tkinter as tk
from tkinter import messagebox
import sys
import json
from pathlib import Path

# ===== 用户偏好（按钮记忆）持久化 =====
# 首次运行时弹窗按 first_default 倒计时；用户一旦选择，则下次倒计时改在用户选择上。
# 放在 src/ 下与 Function_030.py 同级，方便打包分发。
_PREFS_FILE = Path(__file__).resolve().parent / ".ppt_prefs.json"


def _load_prefs() -> dict:
    try:
        if _PREFS_FILE.exists():
            return json.loads(_PREFS_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_pref(key: str, value: str) -> None:
    try:
        prefs = _load_prefs()
        prefs[key] = value
        _PREFS_FILE.write_text(
            json.dumps(prefs, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as e:
        print(f"[prefs] 保存失败: {e}")


def _get_toplevel_hwnd(win):
    """取 Tk 顶层窗口的真实 Windows HWND（非子控件 HWND）。

    `win.winfo_id()` 返回的是 Tk 内部子控件 HWND，FlashWindowEx /
    SetWindowPos 对它静默失败 —— 这是老 bug 任务栏不闪烁的根因。
    正确做法：`win.wm_frame()` 直接返回顶层 HWND（含 0x 前缀的字符串）。
    """
    if sys.platform != 'win32':
        return None
    try:
        return int(win.wm_frame(), 16)
    except Exception:
        try:
            import ctypes
            hwnd = win.winfo_id()
            # 兜底：向上找两层 GetParent
            for _ in range(2):
                parent = ctypes.windll.user32.GetParent(hwnd)
                if parent:
                    hwnd = parent
            return hwnd
        except Exception:
            return None


def center_window(win, width, height):
    """将窗口居中显示（多显示器友好：居中到鼠标所在屏幕的工作区）。

    旧实现用 `winfo_screenwidth()` —— 多显示器/缩放下不稳定，弹窗经常
    跑到主屏中间但用户的 PPT/Excel 在副屏，导致用户看不见。
    新实现：Windows 下读光标当前所在的 Monitor，按工作区（避开任务栏）居中。
    """
    win.update_idletasks()

    if sys.platform == 'win32':
        try:
            import ctypes
            from ctypes import wintypes

            class POINT(ctypes.Structure):
                _fields_ = [('x', ctypes.c_long), ('y', ctypes.c_long)]

            class RECT(ctypes.Structure):
                _fields_ = [('left', ctypes.c_long), ('top', ctypes.c_long),
                            ('right', ctypes.c_long), ('bottom', ctypes.c_long)]

            class MONITORINFO(ctypes.Structure):
                _fields_ = [('cbSize', ctypes.c_ulong),
                            ('rcMonitor', RECT), ('rcWork', RECT),
                            ('dwFlags', ctypes.c_ulong)]

            user32 = ctypes.windll.user32
            pt = POINT()
            user32.GetCursorPos(ctypes.byref(pt))
            MONITOR_DEFAULTTONEAREST = 2
            mon = user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
            mi = MONITORINFO()
            mi.cbSize = ctypes.sizeof(MONITORINFO)
            user32.GetMonitorInfoW(mon, ctypes.byref(mi))
            wa = mi.rcWork  # 工作区，自动避开任务栏
            x = wa.left + (wa.right - wa.left - width) // 2
            y = wa.top + (wa.bottom - wa.top - height) // 2
            win.geometry(f"{width}x{height}+{x}+{y}")
            return
        except Exception as _e:
            print(f"[center_window] 多屏居中失败，回退主屏: {_e}")

    # Fallback：单屏 / 非 Windows
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - width) // 2
    y = (sh - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")


def flash_taskbar(win, count=8):
    """让任务栏图标闪烁直到窗口获得焦点。"""
    if sys.platform != 'win32':
        return
    hwnd = _get_toplevel_hwnd(win)
    if not hwnd:
        return
    try:
        import ctypes

        class FLASHWINFO(ctypes.Structure):
            _fields_ = [
                ('cbSize', ctypes.c_uint),
                ('hwnd', ctypes.c_void_p),
                ('dwFlags', ctypes.c_uint),
                ('uCount', ctypes.c_uint),
                ('dwTimeout', ctypes.c_uint),
            ]

        FLASHW_ALL = 0x00000003
        FLASHW_TIMERNOFG = 0x0000000C  # 闪到获得焦点为止

        flash_info = FLASHWINFO(
            ctypes.sizeof(FLASHWINFO),
            hwnd,
            FLASHW_ALL | FLASHW_TIMERNOFG,
            count,
            0,
        )
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(flash_info))
    except Exception as e:
        print(f"任务栏闪烁失败: {e}")


def force_window_front(win, beep=True, flash=True):
    """强制窗口置顶 + 提示用户（任务栏闪烁 + 系统蜂鸣）。

    修法要点（旧版老 bug 复盘）：
      1. 必须用 `_get_toplevel_hwnd` 取真正的顶层 HWND，
         否则 SetWindowPos / FlashWindowEx 静默失败。
      2. 不能永久 topmost=True —— 那样弹窗会一直挡 PPT。改为
         先 topmost True 推到前台，~400ms 后取消，并 SetWindowPos NOTOPMOST。
      3. 默认开启任务栏闪烁 + 系统蜂鸣，解决"用户无感"问题。
    """
    # tk 层先抢焦点
    win.attributes('-topmost', True)
    win.lift()
    win.focus_force()
    win.update_idletasks()

    if sys.platform == 'win32':
        hwnd = _get_toplevel_hwnd(win)
        if hwnd:
            try:
                import ctypes
                user32 = ctypes.windll.user32
                HWND_TOPMOST = -1
                HWND_NOTOPMOST = -2
                SWP_NOMOVE = 0x0002
                SWP_NOSIZE = 0x0001
                SWP_SHOWWINDOW = 0x0040
                FLAGS = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
                user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
                # 取消永久置顶（保留窗口在前，但允许后续切换）
                win.after(400, lambda: user32.SetWindowPos(
                    hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS))
                if beep:
                    user32.MessageBeep(0x00000040)  # MB_ICONINFORMATION
            except Exception as e:
                print(f"SetWindowPos失败: {e}")

    # tk 层在短延迟后取消 topmost，与 SetWindowPos 同步
    win.after(450, lambda: win.attributes('-topmost', False))
    win.after(500, lambda: win.focus_force())

    if flash:
        flash_taskbar(win)


# ===== 弹窗视觉常量（iOS systemGroupedBackground 风格 + 参考图饱和蓝） =====
# 关键设计：窗口底色用 iOS 系统分组背景灰 (#F2F2F7)，按钮反而是纯白卡片 ——
# "白色卡片浮于浅灰底"的经典 iOS 设置页分层；
# 默认按钮强调描边换成参考截图的饱和 Indigo (#4A6CF7)，去掉旧的红色。
_UI_FONT_FAMILY      = "Microsoft YaHei UI"  # 中文渲染干净（Win 自带）
_UI_BG               = "#F2F2F7"              # 窗口：iOS systemGroupedBackground
_UI_FG               = "#1C1C1E"              # 主文本：iOS label primary（更黑）
_UI_FG_MUTED         = "#8E8E93"              # 次要描述：iOS secondaryLabel
_UI_ACCENT           = "#4A6CF7"              # 默认按钮强调描边：饱和 Indigo（参考截图）
_UI_BTN_BG           = "#FFFFFF"              # 按钮：纯白卡片（与窗口浅灰拉开层次）
_UI_BTN_BG_HOVER     = "#E3F2FD"              # hover：浅蓝（与白色卡片对比清晰）
_UI_BTN_BORDER_HOVER = "#64B5F6"              # hover 描边：稍深蓝（强化层次）
_UI_BTN_BORDER       = "#E5E5EA"              # 静态描边：iOS separator gray


def _ask_with_countdown(
    title: str,
    message: str,
    options,            # list of (label, value[, description, fg])
    pref_key: str,
    first_default,
    countdown_seconds: int = 5,
    width: int = 360,
    on_close_value=None,
):
    """通用「倒计时 + 记忆」弹窗。

    options 每项可为 (label, value) 或 (label, value, description) 或
    (label, value, description, fg)。
    倒计时显示在「上次用户选过的按钮」（首次运行落到 first_default）。
    时间到自动按下默认按钮；用户点击任意按钮取消倒计时。
    返回值会写入 pref_key，下次调用时该值变为新的默认。
    """
    prefs = _load_prefs()
    default_value = prefs.get(pref_key, first_default)

    # 找到默认值在 options 中的索引；找不到则回退到首项
    default_idx = 0
    for i, opt in enumerate(options):
        if opt[1] == default_value:
            default_idx = i
            break
    default_value = options[default_idx][1]
    default_label = options[default_idx][0]

    state = {"choice": None, "after_id": None, "remaining": countdown_seconds}

    win = tk.Tk()
    win.title(title)
    win.resizable(False, False)
    win.configure(bg=_UI_BG)

    def cancel_countdown():
        if state["after_id"] is not None:
            try:
                win.after_cancel(state["after_id"])
            except Exception:
                pass
            state["after_id"] = None

    def select(v):
        state["choice"] = v
        cancel_countdown()
        try:
            win.quit()
            win.destroy()
        except Exception:
            pass

    fallback_close = on_close_value if on_close_value is not None else default_value
    win.protocol("WM_DELETE_WINDOW", lambda: select(fallback_close))

    # ===== 外层容器（统一内边距，告别 raw pack pady=10/8/2 的零散感） =====
    container = tk.Frame(win, bg=_UI_BG, padx=28, pady=24)
    container.pack(fill="both", expand=True)

    # ===== 标题文本 =====
    tk.Label(
        container,
        text=message,
        font=(_UI_FONT_FAMILY, 13, "bold"),
        fg=_UI_FG, bg=_UI_BG,
        anchor="w", justify="left",
        wraplength=max(width - 56, 240),
    ).pack(fill="x", pady=(0, 18))

    buttons = []
    last_idx = len(options) - 1
    for i, opt in enumerate(options):
        label, value = opt[0], opt[1]
        desc = opt[2] if len(opt) >= 3 else ""
        fg_desc = opt[3] if len(opt) >= 4 else _UI_FG_MUTED
        is_default = (i == default_idx)

        # 每个按钮 + 描述包成一个 row Frame，便于整体间距控制
        row = tk.Frame(container, bg=_UI_BG)
        row.pack(fill="x", pady=(0, 14 if i < last_idx else 0))

        rest_border = _UI_ACCENT if is_default else _UI_BTN_BORDER
        btn = tk.Button(
            row,
            text=label,
            font=(_UI_FONT_FAMILY, 11, "bold" if is_default else "normal"),
            fg=_UI_FG, bg=_UI_BTN_BG,
            activeforeground=_UI_FG,
            activebackground=_UI_BTN_BG_HOVER,
            relief="flat", bd=0,
            highlightthickness=2,
            highlightbackground=rest_border,
            highlightcolor=rest_border,
            cursor="hand2",
            padx=14, pady=9,
            command=lambda v=value: select(v),
        )
        btn.pack(fill="x")

        # Hover 反馈：浅蓝填充 + 略深蓝描边（一起变更出层次感）
        def _on_enter(_e, b=btn):
            try:
                b.config(
                    bg=_UI_BTN_BG_HOVER,
                    highlightbackground=_UI_BTN_BORDER_HOVER,
                    highlightcolor=_UI_BTN_BORDER_HOVER,
                )
            except Exception:
                pass

        def _on_leave(_e, b=btn, br=rest_border):
            try:
                b.config(
                    bg=_UI_BTN_BG,
                    highlightbackground=br,
                    highlightcolor=br,
                )
            except Exception:
                pass

        btn.bind("<Enter>", _on_enter)
        btn.bind("<Leave>", _on_leave)

        if desc:
            tk.Label(
                row,
                text=desc,
                font=(_UI_FONT_FAMILY, 9),
                fg=fg_desc, bg=_UI_BG,
                anchor="w", justify="left",
                wraplength=max(width - 56, 240),
            ).pack(fill="x", padx=2, pady=(5, 0))

        buttons.append(btn)

    default_btn = buttons[default_idx]

    def tick():
        if state["remaining"] <= 0:
            select(default_value)
            return
        try:
            default_btn.config(text=f"{default_label}（{state['remaining']}s）")
        except Exception:
            return  # 窗口已销毁
        state["remaining"] -= 1
        state["after_id"] = win.after(1000, tick)

    force_window_front(win)

    # 自然尺寸：让 tk 自己量好，再用入参 width 作下限保证横向不挤
    win.update_idletasks()
    natural_w = container.winfo_reqwidth()
    natural_h = container.winfo_reqheight()
    final_w = max(width, natural_w)
    final_h = natural_h
    center_window(win, final_w, final_h)

    state["after_id"] = win.after(1000, tick)
    win.mainloop()

    final = state["choice"] if state["choice"] is not None else default_value
    _save_pref(pref_key, final)
    return final


def ask_gpt_model():
    """
    交互选择 GPT 模型（带 5s 倒计时 + 记忆）：
    返回：
        'n'                  → 不启用 GPT
        'openai/gpt-5.4' 等  → 启用 GPT，返回模型 id
    """
    # 弹窗1：是否启用 GPT —— 默认按钮 "是"，5s 倒计时
    enable = _ask_with_countdown(
        title="需要启动 GPT 吗？",
        message="是否启用 GPT?（需建立 VPN 连接）",
        options=[
            ("是", "y"),
            ("否", "n"),
        ],
        pref_key="gpt_enable",
        first_default="y",
        countdown_seconds=5,
        width=320,
    )
    if enable != "y":
        return "n"

    # 弹窗2：选择 GPT 版本 —— 默认按钮 main_model（GPT-5.4），5s 倒计时
    version = _ask_with_countdown(
        title="选择 GPT 版本",
        message="请选择要使用的 GPT 模型:",
        options=[
            (main_model, main_model, "最新模型，速度稍慢，专业性强", "red"),
            (mini_model, mini_model, "速度最快，精确度可能欠缺", "gray"),
        ],
        pref_key="gpt_model",
        first_default=main_model,
        countdown_seconds=5,
        width=360,
        on_close_value="n",
    )
    return version or "n"


### 测试代码
##if __name__ == "__main__":
##    result = ask_gpt_model()
##    print(f"选择结果: {result}")
##


def ask_template_choice():
    """弹窗选择问卷模板（带 5s 倒计时 + 记忆），返回 'yzr' / 'zxh' / 'apparel'。

    首次运行倒计时落在 yzr；用户选过后，下次倒计时落在用户上次选的模板。
    """
    return _ask_with_countdown(
        title="选择问卷模板",
        message="请选择问卷模板:",
        options=[
            ("yzr模板", "yzr"),
            ("zxh模板", "zxh"),
            ("apparel 服装测试", "apparel"),
        ],
        pref_key="template_choice",
        first_default="yzr",
        countdown_seconds=5,
        width=300,
    )


# --------------------------- 弹窗函数 ---------------------------




























# ========================================  问卷处理相关函数 ========================================






def get_range(sheet):     #deepseek 差远了，GPT都比它强大；搞定了

##    '''
##    基于当前的sht，假设已经来到问卷sht
##
##    做一个问卷自动识别数据区域的函数，返回首个单元格即可，mc_cell即top_left_cell
##    '''
##    # 获取数据区域的范围
##    data_range = mc_sht.used_range
##
##    # 获取数据区域的左上角单元格
##    top_left_cell = data_range[0, 0]  # 获取左上角单元格（第一行第一列）
##
##    # 返回行和列的坐标 .row .column / 我更习惯返回单元格。。
##    return top_left_cell

    """
    used_range 太模糊了，升级下函数
    自动识别数据区域的左上角单元格
    :param sheet: xlwings Sheet 对象
    :return: (row, col) 或 None 如果没有数据
    """
    # 获取最大行列数（可以适当限制搜索范围提高速度）
    max_rows = 100
    max_cols = 100
    top_row = None
    left_col = None

    for r in range(1, max_rows + 1):
        row_values = sheet.range((r, 1), (r, max_cols)).value
        # 查找这一行第一个非空值单元格
        for c, val in enumerate(row_values, start=1):
            if val is not None and str(val).strip() != "":
                top_row = r
                left_col = c
                mc_cell = sheet.cells(top_row,left_col)
                return mc_cell   # 找到立即返回

    return None  # 没有数据









# 没有Claude协助的话，我已经不会写代码了。。。。。。。

# 目前想到的是，1、【性能相关指标】通常都很短，只有几个字；2、评分数值通常会大于1分，无论是5分制或者10分制 / //  其他问题、描述、这些可以用Claude目前提供的逻辑

# 关于问卷函数估计要持续微调，来适应各种不同问卷格式，无法一劳永逸。             ing............................................................................


import re

def parse_survey_data(data_tuple):
    """
    自动解析调查问卷数据，智能识别【性能评分】列和姓名列

    参数:
        data_tuple: Excel读取后的元组数据
                   格式为 (表头元组, 数据行1元组, 数据行2元组, ...)

    返回:
        元组，格式为 (表头元组, 数据行1元组, 数据行2元组, ...)
        包含姓名列和所有性能评分列（数字评分）
    """
    if not data_tuple or len(data_tuple) < 2:
        raise ValueError("数据至少需要包含表头和一行数据")

    # 第一行是表头
    header = data_tuple[0]

    # 找到姓名列的索引（匹配包含"姓名"或"昵称"的列）
    name_idx = None
    for idx, col in enumerate(header):
        if col and re.search(r'姓名|昵称', str(col)):
            name_idx = idx
            break

    if name_idx is None:
        raise ValueError("未找到姓名列")

    # 智能识别数字评分列：
    # 策略：检查该列的所有数据行，如果大部分都是1-10范围内的数字，则认为是评分列
    score_indices = []

    for col_idx in range(len(header)):
        # 跳过姓名列
        if col_idx == name_idx:
            continue

        # 收集该列所有数据
        col_values = []
        for row in data_tuple[1:]:
            if col_idx < len(row):
                col_values.append(row[col_idx])

        # 如果没有数据，跳过
        if not col_values:
            continue

        # 统计该列中有效数字的数量
        numeric_count = 0
        valid_range_count = 0  # 在合理评分范围内的数量

        for value in col_values:
            # 尝试转换为数字
            num_value = None
            if isinstance(value, (int, float)):
                num_value = float(value)
            elif isinstance(value, str):
                try:
                    num_value = float(value.strip())
                except (ValueError, AttributeError):
                    pass

            if num_value is not None:
                numeric_count += 1
                # 评分通常在1-10或0-10范围内
                if 0 <= num_value <= 10:
                    valid_range_count += 1

        # 判断条件：
        # 1. 至少80%的数据是数字
        # 2. 至少80%的数字在0-10范围内（典型评分范围）
        # 3. 列标题不是"轮次/第几轮"等非评分列（原来用 not has_one 会误排打了1分的性能列）
        total_count = len(col_values)
        col_header_str = str(header[col_idx]) if col_idx < len(header) else ""
        _round_keys = ["第几轮", "轮次", "轮反馈", "这是第几"]
        is_round_col = any(k in col_header_str for k in _round_keys)

        if (numeric_count >= total_count * 0.8 and
            valid_range_count >= numeric_count * 0.8 and
            not is_round_col):
            score_indices.append(col_idx)

    # 组合所有需要的列索引：姓名列 + 所有评分列
    target_indices = [name_idx] + score_indices

    # 构建结果
    result = []

    # 添加表头
    header_tuple = tuple(header[idx] for idx in target_indices)
    result.append(header_tuple)

    # 处理所有数据行
    for row in data_tuple[1:]:
        if not row:
            continue

        # 提取目标列的数据
        row_data = []
        for i, idx in enumerate(target_indices):
            if idx < len(row):
                value = row[idx]
                # 第一列是姓名，保持原样；其他列确保是数字
                if i == 0:
                    # 姓名列，转换为字符串
                    row_data.append(str(value) if value is not None else '')
                else:
                    # 评分列，确保是数字
                    if isinstance(value, (int, float)):
                        row_data.append(float(value))
                    elif isinstance(value, str):
                        try:
                            row_data.append(float(value.strip()))
                        except (ValueError, AttributeError):
                            row_data.append(None)
                    else:
                        row_data.append(None)
            else:
                row_data.append(None)

        result.append(tuple(row_data))

    return tuple(result)


#调试：将结果粘贴至新的数据区域，同时做好标记

##mc_sht = xlwings.books.active.sheets.active
##
### 假设已经来到mc_sht，模范main()函数
##mc_cell = get_range(mc_sht)
##
##    # debug 时简便方法：
##      mc_cell = xlwings.books.active.selection
##
##raw_data = mc_cell.api.CurrentRegion.Value
##
### 得到提取后的问卷数据
##results = parse_survey_data(raw_data)
##
### 将数据输出到指定区域 offset(10,10)
##






# 继续靠 Claude 。。。 提取【姓名/体重/距离/配速】信息
def extract_info(questionnaire_data):
    """
    从问卷数据中提取姓名、体重、距离和配速信息

    参数:
        questionnaire_data: 问卷数据列表，第一行为表头，后续为数据行
                           格式如: [['表头1', '表头2', ...], ['数据1', '数据2', ...], ...]

    返回:
        list: 包含每行数据的提取结果列表
              每个元素为 tuple: (name, body_weight, distance, pace) 所有返回值均为字符串类型
    """
    if not questionnaire_data or len(questionnaire_data) < 2:
        return []

    # 第一行是表头
    headers = questionnaire_data[0]

    # 结果列表
    results = []

    # 遍历每一行数据（从第二行开始）
    for data_row in questionnaire_data[1:]:
        # 初始化返回值
        name = ""
        body_weight = ""
        distance = ""
        pace = ""

        # 创建字段索引字典
        field_dict = {}
        for idx, header in enumerate(headers):
            if idx < len(data_row):
                field_dict[header] = str(data_row[idx]) if data_row[idx] is not None else ""

        # 1. 提取姓名
        # 查找包含"姓名"或"昵称"的字段
        for key, value in field_dict.items():
            if re.search(r'姓名|昵称', key, re.IGNORECASE):
                name = value
                break

        # 2. 提取体重
        # 查找包含"体重"或"重量"的字段
        for key, value in field_dict.items():
            if re.search(r'体重|重量', key, re.IGNORECASE):
                body_weight = value
                break

        # 3. 提取距离
        # 查找包含"距离"的字段
        for key, value in field_dict.items():
            if re.search(r'距离', key, re.IGNORECASE):
                distance = value
                break

        # 4. 提取配速
        # 查找包含"配速"或"速度"的字段
        for key, value in field_dict.items():
            if re.search(r'配速|速度', key, re.IGNORECASE):
                pace = value
                break

        # 确保所有返回值都是字符串类型
        name = str(name) if name else ""
        body_weight = str(body_weight) if body_weight else ""
        distance = str(distance) if distance else ""
        pace = str(pace) if pace else ""

        results.append((name, body_weight, distance, pace))

    return results







# 复制一页 【实战测试】 ppt

def questionnaire_ppt(mc_ppt,mc_slide):

    ''' 这个函数用来复制【实战测试】空白页  '''

    # 【5.1】先复制一遍空白模板页面
    mc_ppt.Slides(4).Select()
    mc_ppt.Slides(4).Copy()

    X = len(mc_ppt.Slides) + 1


    time.sleep(random.random()*delay)
    mc_slide = mc_ppt.Slides.Paste(X)  # 长期使用发现不行，还是要带参数比较稳定



    #要删除2次才能完全清空。。。  可能是延时的问题，碰到好多次了

    for shp in mc_slide.Shapes:    #删除所有的内容（shape）：插入一个完全空白的模板（白板）
        shp.Delete()
    for shp in mc_slide.Shapes:    #删除所有的内容（shape）：插入一个完全空白的模板（白板）
        shp.Delete()



    #  左上角主标题  -  实战测试（标题）
    Left = 15
    Top = 15
    Text =  '实战测试'
    title = Title_1(mc_slide,Left,Top,Text)


    #  中间加虚线 ============
    line_h = Line_Shape(mc_slide,480,30,480,490)     # 温习了很久才记起来使用方法
    line_h.shape.Line.ForeColor.RGB = red                # 当年写的 Class 确实是出类拔萃，代表了我编程的巅峰。。。


    return mc_slide          # 函数化之后，mc_slide 还停留在上一页，因此无法粘贴！！！ 再把新的返回，更新下 mc_slide 的最新状态






# 将main中的【问卷生成ppt】函数挪到这里试试看   # 挪动后结果发现需要不断新增参数，因为发生了嵌套。。。

def questionnaire_Excel(mc_sht, mc_ppt, mc_slide, mc_model, sample_name="", mc_gpt="y", summary_sink=None):

    #  这部分一空就是好多年。。  2025终于开动了，感谢AI时代。。

    #  ================================================= 问卷部分内容 ===================================================


    # 老规矩，先搞定1页中的【左半部分（即问卷只有1行数据】），再考虑【右边】、最后考虑【多页循环】


    # 0、--------------公共部分，放在循环外部 ！！！--------------

##    # 找到问卷 命名为 mc_sht
##    for c in mc_book.sheets:   # c = 针对所有的 sheet（一共n个sheet）
##
##        if '问卷' in c.name:
##
##            mc_sht = c   # 注意，这个设定只能识别1张问卷sheet，这也是合理的， 我不对问卷sht进行循环，我只在1张问卷中循环多行（多个人）的数据
##            mc_sht.select()




    # 如果存在问卷，再开始执行，否则就跳过。这部分写完后统一挪到 function.py中吧，作为一个庞大的函数、。。。。。



    if '问卷' in mc_sht.name:

        # 找到问卷数据区域 /  定位行列，就知道有几个人（多少行）数据了

        mc_cell0 = get_range(mc_sht)

        i0 = mc_cell0.api.CurrentRegion.Row

        i = mc_cell0.api.CurrentRegion.Rows.Count

        j0 = mc_cell0.api.CurrentRegion.Column

        j = mc_cell0.api.CurrentRegion.Columns.Count

        # 临时数据起始位置：远离原始数据区（含浮动图片），放到下方第 20 行处的安全缓冲区
        # 配合下方 try/finally 整行删除，确保不残留 + 不影响图片 mc_debug
        mc_cell = mc_cell0.offset(row_offset=i0+i+20,column_offset=0)                #.select()   加了这个就变成一个动作了。。


        # 问卷全部原始数据 raw_data
        raw_data = mc_cell0.api.CurrentRegion.Value

        # 防止问卷文本识别错误，导致ppt制作失败 ------------
        try:
            raw_data_title = raw_data[0]
            print(f"\n\n           问卷数据识别成功！开始处理问卷内容，注意一共收集到【{len(raw_data)-1}】名测试者的有效问卷")

        except Exception as e:   # 捕获所有异常
            print(f"\n\n           问卷数据识别异常，请检查【问卷】sheet中的内容！！: {e}")

            # 画蛇添足 - 在PPT中打字提示。。。
            Left = 50  # temp_p[0]
            Top = 400
            hint = (f"\n\n           问卷数据识别异常，请检查【问卷】sheet中的内容！！: {e}")
            hint_shape = Title_3(mc_slide, Left=Left, Top=Top, Text=hint,scale=4)  # 2024 注意搞清楚Class文件生成的是一个什么对象，不然很麻烦

            return mc_sht, mc_slide

       # 防止问卷文本识别错误，导致ppt制作失败 ------------



        raw_pure_date = tuple(raw_data[1:])   # 测试下 ，没问题，Tuple和list用法类似，只是无法直接修改，可以构建新的tuple，学到了

                                                                                # raw_data 就是全部问卷数据，包含大量杂乱信息

                                                                                 # raw_data信息量最全，发送给GPT生成优缺点需要使用raw_data

                                                                                 # 因此也需要逐条构建发送数据 【1行 raw_data_title + 1行 raw_pure_date】


        # 生成测试人员信息：【姓名/体重/距离/配速】，使用 Claude 的extract_info(questionnaire_data)函数
        # ing........
        runner_info = extract_info(raw_data)

                                    #  得到的数据为一个 list：  [('周永芳', '', '10-21km', '4min30-5min30'), ('郭德芳', '', '10-21km', '4min30-5min30'), ('杨清建', '', '5-10km', '5min30-6min30'), ('隋晗', '', '10-21km', '4min30-5min30'), ('罗勇', '', '10-21km', '5min30-6min30'), ('段崇文', '', '5-10km', '大于6min30')]
                                    #  里面有n名跑者信息，接下来要考虑如何遍历、在下面的循环中，如何输出到ppt






        # 得到提取后的问卷数据，应该是评分数据，试了1行、4行，都能生成，nice！
        clean_data = parse_survey_data(raw_data)
        clean_data_title = clean_data[0]
        clean_pure_date = tuple(clean_data[1:])
        # print(clean_pure_date)

                                                # clean_data 数据长这样 //  pure_data是去掉标题的数据，差别不大：

                                                # (('姓名', '这是第几轮反馈', '舒适度', '合脚性', '轻量感', '包裹感', '缓冲性能', '回弹性能', '稳定性', '过渡性', '地面反馈', '耐久性'),
                                                #   ('周永芳', 1.0, 10.0, 10.0, 10.0, 10.0, 8.0, 8.0, 9.0, 8.0, 9.0, 8.0),
                                                  #  ('郭德芳', 1.0, 10.0, 8.0, 9.0, 9.0, 10.0, 10.0, 10.0, 10.0, 10.0, 10.0),
                                                  #  ('杨清建', 1.0, 9.0, 9.0, 9.0, 10.0, 9.0, 9.0, 9.0, 9.0, 9.0, 8.0))

                                                   # clean_data是用来做问卷的，每次需要逐条构建问卷【1行 clean_data_title + 1行 clean_pure_date】


    # 0、--------------公共部分，放在循环外部 ！！！--------------






        # 1、-------------------------开始针对问卷【跑者1、2、3...】进行循环-------------------------------

        # 临时数据不再删除：OLE 嵌入图表的行级引用无法通过 CutCopyMode=False 完全断开，
        # 删行必然导致 PPT 图表数据丢失。改为保留临时数据，事后由用户手工清理 Excel。
        for runner in range(0, len(clean_pure_date)):                    # 每个runner 都是一行跑者评分数据

                print(f"\n\n           开始处理第【{runner+1}】份问卷..................")

            # 1、【左半部分】代码，先假设问卷只有1行数据，即1个 runner

                # 先做一个数据区域 new_data，然后适配我之前写的 mark_chart函数。。。 // 之前那个太复杂了，不想再折腾

                # 重写了一个函数 make_chart_for_questionnaire(mc_cell, mc_slide, Left=26, Top=168, Width=250, Height=150):

                # 也回顾了之前的 make_chart函数，发现太复杂了，当时我把【制表】+【位置排版】混合在一起了

                # 以后有空时，重写下这个make_chart函数，可以用Class来写，也可以拆分成不同函数。等以后有空再说吧 ==============================================================================================




                #+++++++++++++循环部分，没想到这个循环条件如此简单。。+++++++++++++++++++++++++++

                if runner % 2 == 0:    # 简单判断奇数偶数，[ 0, 1, 2, 3, 4, 5, 6.....]

                    name_Left = 26          # 跑者信息
                    name_Top = 81

                    chart_Left = 66         # 偶数 = 左边  chart微调后右移40像素=66
                    chart_Top = 168

                    text_Left = 26          # 优缺点
                    text_Top = 327

                    hint_Left = 15
                    hint_Top = 400

                else:

                    name_Left = 500
                    name_Top = 81

                    chart_Left = 543         #奇数 = 右边  chart微调后右移40像素=543
                    chart_Top = 168

                    text_Left = 500
                    text_Top = 327

                    hint_Left = 500
                    hint_Top = 400


                # 每页2名跑者，之后开始递增1页PPT
                if runner in range(2, 1000, 2):
                    mc_slide = questionnaire_ppt(mc_ppt,mc_slide)   # 先增加一个【实战测试】空白页，然后再根据实际情况（问卷条目）酌情增加

                #+++++++++++++++++++ 循环部分结束  +++++++++++++++++++++++++++


                # 1、首先展示跑者基本信息（排版时左、右） name_result

                name_result_text = (
                    f'测试者姓名：{runner_info[runner][0]}\n'
                    f'测试者体重：{runner_info[runner][1]}\n'
                    #f'平均距离：{runner_info[runner][2]}\n'
                    #f'配速区间：{runner_info[runner][3]}\n'

                    )


                name_result = Title_3(mc_slide, name_Left, name_Top, name_result_text, scale=3)

                name_result.tr.Font.Size = 16


                #Text.TextFrame.TextRange.ParagraphFormat.Bullet.Visible=0  # 试试看能否取消，不行只能 sleep了
                #time.sleep(1)
                #result.tr.ParagraphFormat.Bullet.Visible=0           # 调试起来好麻烦。。。。还是类的知识忘光了。。。改天恶补下 ---------------------------------------------------------------------------------------------
                #result.tr.ParagraphFormat.Bullet.Visible=0            # twice 就解决了。。。

                     #----------------------------------



                # 2、构建制表chart数据 & 发送给GPT的raw_data

                    #首先要做一个 发送给gpt的数据出来

                temp_raw_data = []
                temp_raw_data.append(raw_data_title)
                temp_raw_data.append(raw_pure_date[runner])


                # 基于 mc_cell0，循环生成 mc_cell，用来将数据在Excel中摆放好，便于生成条形图

                temp_data = []
                temp_data.append(clean_data_title)
                temp_data.append(clean_pure_date[runner])


                # mc_cell已经定义好了，接下来基于mc_cell来干活
                mc_cell.value = temp_data

                # 直接基于数据，生成 问卷条形图，优雅！！！   执行完之后，ppt中应该已经摆放好了
                # 后面重点设置下位置参数，配合循环即可！！ ------------------------------------------------------------------------------------------------  ing.........
                _tmp_chart = make_chart_for_questionnaire(mc_cell, mc_slide, Left=chart_Left, Top=chart_Top, Width=250, Height=150)

                # 删除 Excel chart 对象（每轮生成一个，Excel 不需要保留），
                # 但【不删除】临时数据行——删行会导致 PPT OLE 图表数据丢失。
                try:
                    _tmp_chart.delete()
                except Exception:
                    pass




                # 画蛇添足 - 在PPT中打字提示。。。

                hint = '[OPENAI]服务器请求发送中，请耐心等待.........'
                hint_shape = Title_3(mc_slide, Left=hint_Left, Top=hint_Top, Text=hint,scale=4)  # 2024 注意搞清楚Class文件生成的是一个什么对象，不然很麻烦




                #接下来重点考虑 gpt的工作： 需要再定制一个 prompt
                mc_prompt = gen_questionnaire_prompt(runner,temp_raw_data)


                mc_completion = GPT_5(mc_prompt, model=mc_model)    # 先统一用 5 来调试，后续太慢了再考虑用4.1之类
                #mc_completion = mc_reply

                # plan4：把每位 runner 的 GPT 结论累积给 Main.py 6.3 用
                if summary_sink is not None:
                    try:
                        summary_sink.append(mc_completion)
                    except Exception:
                        pass


                #假设我拿到了完美的 gpt 答复 mc_completion，接下来需要一个高级染色函数，优点用红色、缺点用蓝色  //  锦上添花，后面再弄吧  ------------------------------------------ ing.......



                # 画蛇添足 - 在PPT中打字提示。。。 然后别忘记删除
                hint_shape.tr.Text = '成功获取 [OPENAI] Reply ！'
                time.sleep(1)
                hint_shape.shape.Delete()  # TextRange 不能直接 Delete，必须使用 Parent 回到 Shape函数才能删除



                # 文字处理工作，先套用之前的参数，再微调 // 手工微调下，ok

                result = Result_Bullet_small(mc_slide, text_Left, text_Top, mc_completion, scale=0.5)


                clean_Text = ''.join(ch for ch in result.tr.Text if ch not in (' ', '\u3000') )
                result.tr.Text = clean_Text


                #Text.TextFrame.TextRange.ParagraphFormat.Bullet.Visible=0  # 试试看能否取消，不行只能 sleep了
                #time.sleep(1)
                result.tr.ParagraphFormat.Bullet.Visible=0           # 调试起来好麻烦。。。。还是类的知识忘光了。。。改天恶补下 ---------------------------------------------------------------------------------------------
                result.tr.ParagraphFormat.Bullet.Visible=0            # twice 就解决了。。。


                smart_color_text(result.tr, color_red=red, color_blue=light_blue)


                # 干完活之后，挪到下一个mc_cell
                mc_cell = mc_cell.offset(row_offset=5,column_offset=0)

        return mc_sht, mc_slide     # 为了程序后续继续能够顺利运行   # for 运行完之后再return，否则for runner 提前终止了
                       #[0]         #[1]














# ========================================== 问卷处理相关函数 ========================================














#import openai


##def GPT_4(mc_prompt):
##
##    global mc_key, mc_model,client
##
##    def initialize_dialogue():
##
##        global messages
##
##        messages = []
##
##    def user_prompt(prompt):
##
##        messages.append({"role": "user", "content": prompt})
##
##    def continue_dialogue(completion):
##
##        messages.append({'role': 'assistant', 'content': completion})
##
##    def dialogue(prompt):  # 目前先仅仅返回一个 reply  // 再加一个 usage
##        # 注意这个 prompt 其实是个 messages，里面的聊天记录可以一条，也可以很多条。。。。
##
##        # 这里添加一个长度模块，每次开始对话前，先判断下messages的长度
##        # 若超过3333个文字，就开始删除聊天记录，直到删到1000字以下（保留最近的聊天记录就行）  //  当然，如果最新的聊天记录超过666字，只好全部删除啦
##        # if len(str(messages)) > 3333:
##        #     for i in messages:
##        #         if len(str(messages)) > 666:
##        #             messages.remove(i)
##        #             print('（为保证连续对话，系统自动删除了1条聊天记录...）\n')
##
##        print('      GPT 服务器请求中，需要一定的响应时间，请耐心等待答复.....\n')
##
##        try:
##            response = client.chat.completions.create(
##
##                model=mc_model,  # 'gpt-4-0125-preview',   # gpt-3.5-turbo-0125
##                # gpt-3.5-turbo  /// nm，GPT-4的回答误导了我（gpt-4），需要不断拷问他。。。。text-davinci-004  ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★
##
##                messages=prompt)
##
##            # ------------- 经典两件套，读取reply ：str
##            print('                 成功获取GPT答复，Congratulations !!!\n\n\n')
##
##            reply = response.choices[0].message.content
##
##
##            # ------------- 读取usage ：dict {}
##            #usage = response['usage']
##
##        except openai.api_resources.connection_error.APIConnectionError as e:
##            print('报错啦，请检查function.py文件中的GPT_4 函数',e)
##
##            reply = '无法与服务器建立连接，请检查VPN设置，然后再尝试！'
##            #usage = None
##
##
##        return reply  #, usage  # 这个reply 是一个纯文本
##
##
##
##
##    initialize_dialogue()  # 第一次清零messages[]
##
##    # messages【0】 用于设定GPT的性格？ 这样竟然也能work...  // 这个本质上也属于 fine-tune 吧  ///  no，fine-tuned是另外收费的模型
##    messages = [
##        {"role": "system", "content": "你是一个非常专业、非常乐于助人的AI助手。"}
##    ]
##
##    # 1、 获取prompt，然后构建user prompt，封装进入 messages[] 中
##    user_prompt(mc_prompt)
##
##    # 2、 发送messages  ， 得到一个 response (type 为 tuple)
##    completion = dialogue(messages)
##
##
##    return completion









# 【Excel】
def search(mc_sht0,target,row_offset=0,column_offset=0):     # 设定一个offset参数，将两个函数合并（强化）
                                                            # 为了强化适应性，继续增加了查找不到时的return，避免报错 // 可以给绘图函数调用了
    '''搜索函数'''

    mc_sht0.select()                   # 每次执行查找之前，都需要select()相应的sheet ************

    find = mc_sht0.api.Cells.Find(What=target, After=mc_sht0.api.Cells(mc_sht0.api.Rows.Count, mc_sht0.api.Columns.Count),\
                      LookAt=xlwings.constants.LookAt.xlPart, LookIn=xlwings.constants.FindLookIn.xlFormulas,\
                      SearchOrder=xlwings.constants.SearchOrder.xlByColumns,\
                      SearchDirection=xlwings.constants.SearchDirection.xlNext, MatchCase=False)

    if find != None:

        find.Select()
        tempbook = xlwings.books.active
        tempbook.selection.offset(row_offset=row_offset,column_offset=column_offset).select()

        return tempbook.selection   #返回整个单元格，（可选：offset 参数），随意调用 row / column / value~

    else:

        return None







# 【Excel】 调用search函数 /// 实现测试方法的输出，同时输出打✔的单元格，以list形式？

def test_detail(mc_book):

    '''针对【基础信息】定制的测试详情函数，返回测试列表（测试详情描述、测试图片单元格、测试项目）'''

    mc_sht0 = mc_book.sheets['基础信息']

    # search 已带 offset(1,1)：返回的是【表头下一行 + 测试列表右一列】，即 (第一数据行, 测试项目列)
    start_cell = search(mc_sht0, '测试列表', row_offset=1, column_offset=1)
    if start_cell is None:
        return '', [], [], [], []

    first_data_row = start_cell.row       # 第一条数据行（不是表头行；老版用 end('up') 隐式回到这里）
    col_item       = start_cell.column    # 【测试项目】列（测试名所在列，连续非空，用来定位 test_end）
    col_marker     = col_item - 1         # 【测试列表】列（✔ 标记，非空=该行选中）

    # 用 range.end('down') 替代 selection 操作，找到最后数据行
    test_end = mc_sht0.range((first_data_row, col_item)).end('down').row

    if first_data_row > test_end:
        return '', [], [], [], []

    # 整块读取：(first_data_row, col_marker) ~ (test_end, col_marker+6)
    # 原来每行 ~9 次 COM 调用，现在合并成 1 次（老旧机器上 COM 开销显著）
    raw = mc_sht0.range(
        (first_data_row, col_marker),
        (test_end,       col_marker + 6)
    ).value

    # xlwings 单行时返回 list（非 list-of-list），统一格式
    if test_end == first_data_row:
        raw = [raw]

    test_list     = []   #【+1】测试项目
    temp_text     = ''   #【+2】测试方法描述（带 \r 的拼接字符串）
    pic_list      = []   #【+6】图片单元格 range
    test_standard = []   #【+4】测试标准
    test_result   = []   #【+5】参考文本（原 测试结论）

    for idx, row in enumerate(raw):
        if row is None or row[0] is None:   # 标记列（测试列表 ✔）为空 → 未选中，跳过
            continue

        actual_row = first_data_row + idx

        test_list.append(row[1])                               # 测试项目（测试名）
        detail = row[2] if row[2] is not None else ''          # 测试方法描述（健壮性：防 None 拼接报错）
        test_standard.append(row[4])                           # 测试评价标准
        test_result.append(row[5])                             # 示范文本
        pic_list.append(mc_sht0.range((actual_row, col_marker + 6)))   # 图片单元格（显式绑定 sheet）

        temp_text += detail + '\r'

    return temp_text, pic_list, test_list, test_result, test_standard

            # 【0】弱水三千 只取一瓢，我只要测试方法描述 test_detial —— 带【\r】的字符串 ； 搞定！

            # 【1】 结果还是要return pic_list，返回单元格的list，来自动筛选图片、挑出来

            # 【2】 再 return 一个【测试项目】列表 test_list，用来查找核对sheet名称，以便校验

            # 【3】 【GPT-参考文本】 继续返回测试结论话术 test_result。。。。

            # 【4】 【GPT-测试标准】 2024新增，返回test_standard，用于GPT 函数






























# 【PPT】

 # 复制 / 自动更新分页目录页码函数

content_count = 1

def content_slide(mc_ppt):

    global content_count # 这样才能修改函数外部的值

    '''定义一个函数，专门用来生成分页目录标题页，默认复制Slide(3)的模板，然后自动修改数字  a.isdigit() // 自动命名目录标题'''

    mc_ppt.Slides(3).Select()
    mc_ppt.Slides(3).Copy()
    X = len(mc_ppt.Slides) + 1

    time.sleep(random.random()*delay)
    temp_slide = mc_ppt.Slides.Paste(X)  # 默认粘贴到末尾 // 带参数更稳定

    # 然后自动更新数字目录

    for i in temp_slide.Shapes:  # 默认只有2个文本框，必须包含文字

        text = i.TextFrame.TextRange.Text

        #print('1')

        if text.isdigit() == True:        # 逐步调试才发现，【1】中包含一个空格，导致无法识别数字。。。 这种识别方式太不精确了，后续再升级。。。。。。。。。。。。。。。。

            #print('2')

            i.TextFrame.TextRange.Text = str(content_count)

        # 再自动命名目录标题
        else:

            if content_count == 1:

                i.TextFrame.TextRange.Text = '项目背景'

            elif content_count == 2:

                i.TextFrame.TextRange.Text = '基础性能测试'

            elif content_count == 3:

                i.TextFrame.TextRange.Text = '实战测试'

            elif content_count == 4:

                i.TextFrame.TextRange.Text = '总结'

            else:

                i.TextFrame.TextRange.Text = '其他'

    content_count += 1   # 如果没有global claim，将无法修改函数外部的变量







# 【PPT】
def color_key(text_range,key,color,bold=1):

    # 文字已生成，接下来用函数把关键词加红加粗  // 离线调试已ok  // 颜色已定义
    # start = text_range.Find(key).Start
    # length = len(key)
    # temp = text_range.Characters(start,length)
    # # --------- 格式调整 -----------
    # temp.Font.Color = color # 加颜色
    # temp.Font.Bold = bold      # 加粗 = 1   //



    # ------------------ 循环染色，所有字体都染色 / GPT还是厉害，不烧脑了
    start = 1
    while start < text_range.Length:
        found_range = text_range.Find(key, start)
        if found_range is None:
            break  # Exit the loop if the keyword is not found
        else:
            found_range.Font.Bold = bold
            found_range.Font.Color = color  # Change to your desired color

        start = found_range.Start + found_range.Length




#  耗尽了 Claude 免费token ，终于搞定了
def smart_color_text(text_range, color_red, color_blue, bold=1):
    """
    使用类似 color_key 的循环查找方式
    先删除【】，再逐个染色+加粗
    Claude 迭代到 v8 终于成功了。。。。。。。。

    """
    full_text = text_range.Text

    # === 收集所有需要染色的关键词 ===
    keywords_red = []
    keywords_blue = []

    idx_adv = full_text.find("【优点】")
    idx_disadv = full_text.find("【缺点】")

    # 提取优点部分的关键词
    if idx_adv != -1:
        adv_end = idx_disadv if (idx_disadv != -1 and idx_disadv > idx_adv) else len(full_text)
        part_adv = full_text[idx_adv:adv_end]
        keywords_red = re.findall(r"【(.*?)】", part_adv)

    # 提取缺点部分的关键词
    if idx_disadv != -1:
        part_disadv = full_text[idx_disadv:]
        keywords_blue = re.findall(r"【(.*?)】", part_disadv)

    # === 删除【】===
    cleaned_text = re.sub(r"[【】]", "", full_text)
    text_range.Text = cleaned_text

    # === ★★★ 先手动染色开头的 "优点" 和 "缺点" ★★★ ===
    # 因为它们在特殊位置，用 Characters 直接定位
    if idx_adv != -1:
        # "优点" 删除【】后在位置 1，长度 2
        優點_range = text_range.Characters(1, 2)
        優點_range.Font.Bold = bold
        優點_range.Font.Color = color_red

    if idx_disadv != -1:
        # 计算 "缺点" 删除【】后的位置
        # 原位置 idx_disadv，删除了之前所有【】
        brackets_before = full_text[:idx_disadv].count("【") + full_text[:idx_disadv].count("】")
        缺点_pos = idx_disadv - brackets_before + 1  # +1 因为 win32com 从1开始
        缺點_range = text_range.Characters(缺点_pos, 2)
        缺點_range.Font.Bold = bold
        缺點_range.Font.Color = color_blue

    # === 染色+加粗其他关键词（跳过"优点"和"缺点"）===
    # 优点部分：红色+加粗
    for keyword in keywords_red:
        if keyword in ["优点", "缺点"]:  # 跳过已处理的
            continue
        start = 1
        while start <= text_range.Length:
            found_range = text_range.Find(keyword, start)
            if found_range is None:
                break
            found_range.Font.Bold = bold
            found_range.Font.Color = color_red
            start = found_range.Start + found_range.Length

    # 缺点部分：蓝色+加粗
    for keyword in keywords_blue:
        if keyword in ["优点", "缺点"]:  # 跳过已处理的
            continue
        start = 1
        while start <= text_range.Length:
            found_range = text_range.Find(keyword, start)
            if found_range is None:
                break
            found_range.Font.Bold = bold
            found_range.Font.Color = color_blue
            start = found_range.Start + found_range.Length
















# 【PPT + Excel】


#  【Excel + PPT 图片排序函数】遍历返回值的难题未解决，暂时用双重嵌套改写下这个函数（2）

  # 通过05-18的学习，掌握了  ★★ 【*args & **kwargs】 ★★ 的用法，因此顺利把函数简化为一个

def mc_pic(*list_cells,temp_slide,Left,Top,scale,gap=2.5,ret=0):

    '''  这个函数，可以将兴趣区间内的图片（多个单元格 - list_cells），自动排成一行（整齐~~） 通过Left 和 Top来确定精确摆放位置  '''
    #def mc_pic(temp_cell,temp_slide,height,left,top,gap=2.5):  # 原参数

    pointer_left = Left

    for temp_cell in list_cells:

        # 【X.1】首先要进行基本定位，找到：【图组1】单元格的大小/位置信息

        # cell_h——单元格高度       mc_sht.cells(1,1).height  // mc_cell.height
        # cell_w——单元格宽度       mc_sht.cells(1,1).width  // mc_cell.width
        # cell_top——单元格纵坐标   mc_sht.cells(1,1).top    // mc_cell.top
        # cell_left——单元格横坐标  mc_sht.cells(1,1).left  // mc_cell.left

        # 这里假设我已经找到了 mc_cell, 将它赋值到 【temp_cell = mc_cell】



        # ★ 1格的兴趣区间 （图片左上顶点位于单元格内，均属于兴趣区间内）
        top_max = temp_cell.top + 1*temp_cell.height
        top_min = temp_cell.top - 0*temp_cell.height
        left_max = temp_cell.left + 1*temp_cell.width
        left_min = temp_cell.left - 0*temp_cell.width




        for i in range(0,len(temp_cell.sheet.shapes)):       # temp_cell.sheet  等价于  temp_cell.parent

            s_top = temp_cell.sheet.shapes[i].top
            s_left = temp_cell.sheet.shapes[i].left

            if s_top >= top_min and s_top <= top_max and s_left>=left_min and s_left<= left_max:

                #print(i)


                # 输出到本地 // 【 全新升级—— 粘贴到ppt~~~】

                temp_cell.sheet.shapes[i].api.Copy()

                time.sleep(random.random()*delay)
                temp_shape = temp_slide.Shapes.Paste()
                temp_shape.ScaleHeight(1,1,0)  # 04-20 【新增】先让图片恢复原来的大小，避免 Excel 中的略缩图失真   # 05-12 小bug，宇昂的图片失真的厉害(1,1,0)，只好把原始比例缩放取消了(3,0,0)

                                                                             # 结果发现即使改为(3,0,0)，失真的问题也无法解决，只能要求宇昂截图时注意比例控制，不能失真

                                                                                             #   ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★

                                                                                                # bug#01： 调试代码如下

                                                                                                ##>>> mc_book = xlwings.books.active
                                                                                                ##>>> mc_sht = mc_book.sheets.active
                                                                                                ##>>> mc_sht.shapes(1).api.Copy()
                                                                                                ##>>> temp_shape = mc_slide.Shapes.Paste()
                                                                                                ##>>> mc_sht.shapes(1).api.Copy()
                                                                                                ##>>> temp_shape = mc_slide.Shapes.Paste()
                                                                                                ##>>> temp_shape.ScaleHeight(1,1,0)
                                                                                                ##>>> mc_sht.shapes(1).api.Copy()
                                                                                                ##>>> temp_shape = mc_slide.Shapes.Paste()
                                                                                                ##>>> temp_shape.ScaleHeight(1,0,0)


                # ★接下来搞定位置、大小问题

                      #  - 【1】首先，Scale的问题。

                      # 屏幕高度540，【图片高度】=pic_h，【排版预期】=rate 在程序最开始部位设定，比例mc_scale=540/pic_h/rate

                      # 那么根据每张图，都要动态计算一次 mc_scale ， 来让所有的图都高度一致，这样才能确保图片能排整齐。

                pic_h = temp_shape.Height  # win32，注意大写，切记！！  2023 这里是图片的真实高度，假设是1080；同时假设我输入的scale是1（针对满屏而言），那么temp_scale 要 = 0.5，才能实现屏幕缩放

                temp_scale = 540 /pic_h * scale


                       # - 【2】 ★★★ ScaleHeight(【缩放比例=float】，【1=（推荐）依据原始比例，0 = 不依据原始比例】，【0-（推荐）左上角，1-中心，2-右下角】)


                temp_shape.ScaleHeight(temp_scale,0,0)





                       # - 【3】和往常一样，mc_shape 有 【Left / Top】 和 【 Height / Width】 4个参数来控制

                       #  我计划设置一个光标，用来存储下一个图片的【Top】 和 【Left】，就像word打字的光标、Excel的单元格一样

                       #  【原点】 —— 这个光标的初始位置，设定为 【Top】= 120    【Left】= 50

                       #  【间隙】 —— mc_gap=5， 通过调试确认

                temp_shape.Top = Top

                temp_shape.Left = pointer_left





                       # - 【4】 光标下移至下一个位置，目前只移动left就行   【下个位置】 = 当前zero_left + mc_gap + mc_shape.Width 当前图片宽度

                pointer_left += gap + temp_shape.Width



    #return temp_shape # 2023/05/09 把shape送出去，然后最后再微调下 ///////  05-17 这个必须有，因为在矩阵图函数中，图片粘贴完之后，还要微调居中。。。。。

                       # 05-17 碰到一个循环（双重）return的bug问题 /// 删掉之后，还能在粘贴完成后，继续微调吗？不行就只能改写这个函数了。bug★

                        # 只能尝试继续增加条件，避免循环Retrun的情况。

                               #也就是说，在make_chart函数中（这个函数本身会return四个值），因为无需继续调整图片位置，粘贴上去就行，那么就要避免return

                               #而在 make_matrix函数中（这个函数本身不会return任何东西），图片粘贴上去之后，还要微调居中，那么这个时候就要return

                                # bug★解决了！ 果然如我所想，循环Retrun会出大问题，因此要小心。。。。。。。



    if ret == 0:
        pass
    else:
        return temp_shape




     # - 【**】 通常写到这里就ok了，但问题是，我希望左右也能对齐，那么向右排序的问题就不好弄了，需要在循环外，另写一段代码。。。

        # 最后一段代码的思路如下： 将函数内部的循环，进行打包；  【然后在函数外面、主程序中调整整个图片合集的位置。】

        # bag = mc_ppt1.Slides(2).Shapes(1)  ///   bag.GroupItems(2).Height  这里其实等价：   bag.GroupItems(2) = Shapes(2)

        # 那么剩下的问题是，如何新增 GroupItems

        # mc_ppt1.Slides(2).Shapes(1).Type = 【 6 = 2张图片组合；  13 = 单张图片】

        # 组合的难题，暂时无法解决 ---------------------------------------------------------









#  【Excel + PPT chart 单独为问卷写一个全新的函数，避免出状况】


#  2025 调试===========
def make_chart_for_questionnaire(mc_cell, mc_slide, Left=26, Top=168, Width=250, Height=150):
    '''
    【制表函数2，专门服务问卷】一次生成1张图表，
    按照【Left =,Top =,Height =,Width =】 4个参数放到指定位置

    参数:
    mc_cell: 当前Excel中一个单元格对象（用于定位sheet）
    mc_slide: 当前PPT页面对象（slide）
    Left, Top, Height, Width: 图表在PPT中的位置和大小，单位是PowerPoint内部单位（points）
    求助GPT-5 / Claude都失败了，还是得靠自己，排版颗粒度太细了
    '''
    print('问卷制作chart开始——\n')

    # Excel 控制
    mc_sht = mc_cell.sheet
    mc_sht.select() # 先切换到对应的sheet，避免select失败
    mc_cell.select()
    mc_book = mc_sht.book

    i0 = mc_cell.api.CurrentRegion.Row
    i = mc_cell.api.CurrentRegion.Rows.Count
    j0 = mc_cell.api.CurrentRegion.Column
    j = mc_cell.api.CurrentRegion.Columns.Count

    chart_left = mc_sht.cells(i0 + i - 2, j0+3 ).left    # 往右上角移动（2，3）单元格
    chart_top = mc_sht.cells(i0 + i - 2, j0+3 ).top

    # 假设数据区域总是固定的，按下面这种格式排放

                        #  数据清洗不干净，先将就下
                        # (('姓名', '这是第几轮反馈', '舒适度', '合脚性', '轻量感', '包裹感', '缓冲性能', '回弹性能', '稳定性', '过渡性', '地面反馈', '耐久性'),
                            #   ('周永芳', 1.0, 10.0, 10.0, 10.0, 10.0, 8.0, 8.0, 9.0, 8.0, 9.0, 8.0),

    # 还是2023年的我更强。。。 Claude + GPT-5 都整不明白。。。。
    # 插入条形图，设置数据源
    mc_chart1 = mc_sht.charts.add(chart_left,chart_top,width=Width,height=Height)      # 目前250 × 150刚刚好，适合粘贴到ppt中
                                                                                                                                                # 发现一个小bug，标准格式是Width=,Height= //  我昨天写反了。。。。。 以前的函数没问题。。。
    mc_chart1.chart_type = 'bar_clustered'
    mc_chart1.set_source_data(mc_sht.range((i0,j0),(i0+i-1,j0+j-1)))

    # 自动识别量表范围：5分制 or 10分制（只有这两种情况）
    # 取评分列（跳过姓名列 j0 和"第几轮"列 j0+1），检查是否有分值 > 5
    try:
        score_vals = mc_sht.range((i0 + 1, j0 + 2), (i0 + i - 1, j0 + j - 1)).value
        # xlwings：单行返回 flat list，多行返回 list of list
        if score_vals and isinstance(score_vals[0], (list, tuple)):
            flat_vals = [v for row in score_vals for v in row if isinstance(v, (int, float))]
        else:
            flat_vals = [v for v in (score_vals or []) if isinstance(v, (int, float))]
        _scale_max = 10 if any(v > 5 for v in flat_vals) else 5
        print(f'  识别量表范围: {_scale_max} 分制（数据最大值={max(flat_vals, default=0):.1f}）')
    except Exception as _e:
        _scale_max = 10
        print(f'  量表范围识别异常，默认10分制: {_e}')

    # 感谢【Tools】 017 Docx 工具包。。。。 这些代码究竟是怎样写出来的。。。VBA的工程师们，写的时候完全没考虑过API使用感受。。。

    # ★★★★★★★★★ chart格式的调试接口，还好，这些代码是通用的，不受chart_type的影响（影响不大）★★★★★★★★★

    # =================================

    # 隐藏 【图例】
    mc_chart1.api[1].SetElement(100)    # 100 = 隐藏图例 // 101 = 显示图例（右侧）   102 = 显示图例（上侧）      103 = 显示图例（左侧）      104 = 显示图例（下侧）


    # 隐藏 【网络线】
    mc_chart1.api[1].SetElement(328)   #隐藏 y轴的网格线（横线）    (330) = 显示 y轴的网格线（横线）


    # 固定坐标轴量程（0 ~ scale_max），并隐藏坐标轴显示
    # 注意：不能先 Delete() 再设 scale——Delete 会触发 Excel 重新自适应，量程设置丢失。
    # 改为：先固定 MinimumScale/MaximumScale，再把轴线/刻度/标签全部隐藏，视觉效果与 Delete 相同。
    _val_axis = mc_chart1.api[1].Axes(2)   # Axes(1)=左侧分类轴, Axes(2)=下方数值轴
    _val_axis.MinimumScaleIsAuto = False
    _val_axis.MaximumScaleIsAuto = False
    # 坐标轴最大值 = 量表max + 1（5分制→6，10分制→11），给数据标签留出空间，
    # 避免 score = max 时数据标签压在 bar 末端
    _axis_max = _scale_max + 1
    _val_axis.MinimumScale = 0
    _val_axis.MaximumScale = _axis_max
    _val_axis.TickLabelPosition = -4142    # xlTickLabelPositionNone：隐藏刻度标签
    _val_axis.MajorTickMark = -4142        # xlTickMarkNone：隐藏主刻度线
    _val_axis.MinorTickMark = -4142        # xlTickMarkNone：隐藏副刻度线
    _val_axis.Format.Line.Visible = 0      # msoFalse：隐藏轴线本身
    print(f'  坐标轴已固定: 0 ~ {_axis_max}（{_scale_max}分制+1，预留数据标签空间），轴线/刻度/标签已隐藏')
    #    Axes(1) = 左侧坐标轴      Axes(2) = 下方坐标轴

    # 添加【数据标签】
    mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()

    # 隐藏 【主标题】
    mc_chart1.api[1].SetElement(0)   # 0 = 隐藏 // SetElement(2) = 显示主标题（默认，与源数据中的标题字符相同）
    mc_chart1.api[1].SetElement(0)
    #print('chart主标题已隐藏！twice')

    # chart 生成ok了，接下来完成复制即可 ---------------------------
    # 二选一，如果还是出现无法复制的错误，再说吧。。。  GPT已经把原因解释了，必须显示才能复制  // 继续控制屏幕缩放，来避免复制失败。。
    mc_app = xlwings.apps.active
    #mc_app.api.ActiveWindow.Zoom = 100
    mc_cell.select()

    # OLE 复制（粘贴原始图表，保留交互性和最高显示质量）
    mc_chart1.api[0].Copy()

    time.sleep(random.random()*delay)

    mc_shape = mc_slide.Shapes.Paste()
    time.sleep(0.5)  # 等待 PPT 完成 OLE embed 渲染

    # 清除 Excel 剪贴板，断开 OLE 热链接：
    # 粘贴后 Excel 仍保持 CutCopyMode=True（剪贴板激活），此时 PPT 与 Excel chart
    # 之间存在 COM 热链接，Excel 图表的任何变动（包括删行）都会刷新 PPT 显示。
    # 清除后 OLE embed 进入独立显示状态，后续删行/删 chart 不再影响 PPT 图表。
    try:
        xlwings.apps.active.api.CutCopyMode = False
    except Exception:
        pass

    # 位置暂时先用这个，手工排版的

##    Left = 26           #左
##    Top = 168
##
##    Left = 503         #右
##    Top = 168

    mc_shape.Left = Left
    mc_shape.Top = Top
##    mc_shape.Height = Height
##    mc_shape.Width = Width




    # 程序外 debug 使用
    #mc_chart1 = mc_sht.charts[0]
    print('问卷制作chart完成，已粘贴至ppt！')

    return mc_chart1







#  【Excel + PPT chart 图表生成函数】  目前为止最复杂的一个函数 ===================================================================================================


# 2025-12 修复宇昂发现的小bug，解决缩放问题
def Excel_zoom(sht, zoom=30):
    sht.api.Activate()  # 激活 sheet
    sht.api.Parent.Windows(1).Zoom = zoom

# 使用
#Excel_zoom(mc_sht, 60)




def make_chart(mc_sht,mc_slide):

    '''  【制表函数】 这个函数可以自动生成图表，然后复制到ppt中 '''

    #global sample_name   # chart 函数中竟然没用到这个。。

    #【part 1 根据不同图表数量，设定大小、位置参数】

    mc_book = mc_sht.book     # 不想搞太多参数，直接用sheet的parent

    temp_list = []

    for i in range(1,100):
        target = "图表"+str(i)
        find = search(mc_sht,target)

        if find != None:
            temp_list.append(find)           # temp_list =  图表1 / 图表2 / 图表3... 所在的range
        else:
            break

    i = len(temp_list)  # 一共能有i个图表 chart // 都放在  temp_list[0, 1, 2, 3 ...]  中，随时可以调用行、列、value~



     # 针对不同图表数量 i ，要设置不同chart  大小尺寸 / 位置的【初始值】，以便于PPT能排版容纳。图表超过4个，就默认等于3个的大小了（跑出屏幕也无所谓了），到时手工再调整

    if i == 1:

        Left =225
        Top =100
        Height =266
        Width =530

    elif i == 2:

        Left =72
        Top =120
        Height =240
        Width =360

    elif i == 3:

        Left =25
        Top =100
        Height =210
        Width =300

    elif i >= 4:   # 大小和3个相同

        Left =25
        Top =100
        Height =210
        Width =300


    #【part 2 针对当前sheet中的所有图表，进行遍历】

    for p_i in range(0,i):       # 一共能有i个图表 chart                        # 2023-05-10 这个for循环  调试出现bug ==========================================================================================

        temp_list[p_i].select()  # search是一个xlwings函数，需要用小写的select()   # 这里假设先选中 【图表1 - 单元格】

                                 # temp_list 包含的是 【图表1、2、3..】的单元格，假设这里选中了  【图表1】

                                 # 我准备在这里启用 【测试样品数量】，来独立识别，不限制样品数量，多少个都可以，拓展程序适用性

        # 这里仍然选中的是 【图表1 - 单元格】
        temp_i = temp_list[p_i].row
        temp_j = temp_list[p_i].column




        # 增加一个样品数量识别模块(control_count)，每个chart都单独识别一遍    ★★★★★★★★★★★★★★★★★★★
        temp_row = mc_book.selection.row
        temp_column = mc_book.selection.column
        mc_book.selection.end('down').select()
        control_count = mc_book.selection.row - temp_row    # 总样品数量 仍然放在 control_count 中，每个chart都单独识别（原来的代码不用改），并且自由了


        # 增加一个条件判断 // 2023-05 因为【柱状图】、【折线图】都用得到，并且需要为chat_GPT函数做准备
        temp_list[p_i].select()
        mc_book.selection.end('right').select()
        n_j = mc_book.selection.column - temp_j  # ★★★★★★★★★★★★★★★★★★★  n_j 表示测试条件数量 ///  control_count 表示测试样品数量


        # 2023-05-17 再增加一个chart 位置模块，避免看不见，就复制失败的bug
            #2024-02-23 再次出现bug，当我使用大量数据，chart跑到了数据最下方，看不见了

        chart_left = mc_sht.cells(temp_row + control_count + 1 ,temp_column).left
        chart_top = mc_sht.cells(temp_row + control_count + 1 ,temp_column).top

        chart_cell = temp_list[p_i].end('down').offset(row_offset=3,column_offset=0)     # 2024 试试看能否彻底解决这个历史遗留问题？







        # 并且将颜色储存在  color_value[]  中
        color_value = []                                                        # 每个chart的color都可以不同，因此这里需要清零
        color0 = [red,light_blue,green,orange,purple,gray,black,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray,light_gray]
                        # 设置一个默认配色方案~~~       第七个样品之后，全都默认浅灰色；如果超过40个样品，就超标了，这概率应该比较小。。。吧？


        for c in range(0,control_count):                    # 假设4个样品， 从 0 - 4              # 这个函数调试也是None。。。。 why ?? 结果是上面的column 偷懒复制，又出错了

            cell_color = mc_sht.cells(temp_row+1+c,temp_column-1).api.Interior.Color            # 颜色储存变量。。                                                                            # 不过调试发现一个更加严重的bug，那就是颜色失真了。。。。。。。。。。
            if cell_color != 16777215:
                #cell_color = int(mc_sht.cells(temp_row+1+c,temp_column-1).api.Interior.Color)            # 最后还是靠api解决了问题，没办法，源于VBA，自然绕不过去api，其他的xlwings返回的颜色为RGB格式，还是不好用。。
                color_value.append(cell_color)                                                              #   即使空着，也会默认为单元格为白色（16777215）
            else:
                color_value.append(color0[c])                                                                #  05-25 解决空白导致染色透明的问题，空白的话，那就用默认颜色吧【color0】



##          # 历史遗留（颜色识别失败，主要是十六进制转换2次导致失败），暂且保留吧
##            if mc_sht.cells(temp_row+1+c,temp_column-1).color != None:
##                rgb = mc_sht.cells(temp_row+1+c,temp_column-1).color  # Tuple
##                rgb = RGB_to_Hex_to_Dec(rgb)                          # Tuple to 整数
##                color_value.append(rgb)  # 将颜色（整数格式）储存起来
##
##            else:
##                mc_sht.cells(temp_row+1+c,temp_column).color = 0                  # 默认黑色
##                color_value.append(mc_sht.cells(temp_row+1+c,temp_column-1).color)  # 将颜色储存起来






        chart_type = mc_sht.cells(temp_row-2,temp_column).value   # 测试过，没问题

        temp_l1 = (0,1,2,3,4,5,6,7,8,9,10,11,12)   # 默认chart是按从左往右依次递增摆放，而条形图例外，因此单独定义条形图的递增参数

        #  （开始正式制表）=====================================================================================================

        if '柱状图' in chart_type:   # 替换掉 【chart_type == '柱状图':】  ，否则条件太严苛了，多个空格（弯弯误操作）也不行

            #【函数规范】 mc_chart1 = mc_sht.charts.add(left=0,top=0,width=355,height=211)    ★★★★★    # 通过API官方文档，发现了该函数的标准用法，包含4个参数！！！

             # 依次插入空白图表、设置图表类型
            mc_chart1 = mc_sht.charts.add(chart_left,chart_top,(355+30*(control_count-2))*1,211*1)       # ★★★ bug02（已解决）： 这里需要根据样品数量，来设定图表宽度；另外，为了匹配报告，设置为*0.66倍缩放 // 但11-04出现图片挂掉现象，只能取消缩放，改为*1
            mc_chart1.chart_type = 'column_clustered'   #设置一个柱状图


            # # 这里仍然选中的是 【图表1 - 单元格】
            # temp_i = temp_list[p_i].row
            # temp_j = temp_list[p_i].column

            # 增加一个条件判断，区分单条柱状图 / 2条 / 多条柱状图
            #mc_book.selection.end('right').select()
            #n_j = mc_book.selection.column - temp_j       # ★★★★★★★★★★★★★★★★★★★  n_j 表示测试条件数量 ///  control_count 表示测试样品数量

            mc_chart1.set_source_data(mc_sht.range((temp_i,temp_j)).expand())



            # 制表，格式重点  （慢慢来，不要搞太复杂，有需要再改）
            mc_chart1.api[1].SetElement(2)      #显示主标题
            mc_chart1.api[1].ChartTitle.Text = xlwings.Range((temp_i-4,temp_j)).value     # 修改主标题文本
            #mc_chart1.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Size = 13        # 修改图表主（正上方）标题大小  （最新版本，默认大小ok）
            mc_chart1.api[1].SetElement(309)  #显示 y轴 标题 （左侧，字体左侧旋转90°，别扭，但节约空间，英文适用）
            mc_chart1.api[1].Axes(2).AxisTitle.Text = xlwings.Range((temp_i-3,temp_j)).value          # 【左侧】 y轴标题（单位）的名字
            mc_chart1.api[1].SetElement(104)    #12/10 根据弯弯他们的要求，显示图例（下侧）



            # 目前只针对单个柱状图写了代码，需要再嵌套一层If



            if n_j==1:

                  ## 【3.1 - 单条柱状图生成】

                    # 以下这些是针对单条图写的
                mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()    # 添加数据标签          # ★★★ bug 03（11/04已解决）： 移植时发现的问题，竟然需要删除Full   【FullSeriesCollection(1)】- 【SeriesCollection(1)】


                #
                for c in range(0,control_count):

                    mc_chart1.api[1].SeriesCollection(1).Points(c+1).Format.Fill.ForeColor.RGB = color_value[c] #or -16776961     # 样品序列染色  这样改写是因为我可以设置了5个样品，结果出现未染色现象（因为没有5个样品的代码呀）


##                mc_chart1.api[1].SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = color_value[0] #or -16776961     # 样品1染色
##                mc_chart1.api[1].SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = color_value[1] #or 15773696      # 样品2染色
##
##
##                if control_count == 3:       # 2个对比竞品的话
##                    mc_chart1.api[1].SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = color_value[2] #or 5287936    # 样品3染色   # ★★★★★★★★★★★★ bug04： 【样品数量设定为3个，但实际只提供2个数据，会导致无法染色，报错】  目前最多做4个图表，有需要再增加。。。。 ★★★★★★★★★★★★★★★
##
##                if control_count == 4:        # 3个对比竞品的话
##                    mc_chart1.api[1].SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = color_value[2] #or 5287936     # 样品3染色
##                    mc_chart1.api[1].SeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = color_value[3] #or -16727809   # 样品4染色
##






                temp_min = min(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value)
                temp_max = max(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value)

                mc_chart1.api[1].Axes(2).MinimumScale= round(temp_min/2,1)     #比最小值小一半。。。  聪明啊~  11/05 第一次min生成一个list，需要再次min才能得到一个值
                mc_chart1.api[1].Axes(2).MaximumScale= round(temp_max*1.15,1) #比最大值大一点点。。
                # ------------ 原来 Range.value 会直接输出一个 list ， nice~

                #mc_chart1.api[1].Axes(2).MajorUnit = int((temp_max-temp_min)/4)     # 11/05 目前2、3、4都不太合适，暂时先自动吧，弯弯的需求（自动设定 scale ）暂时无法完美解决啦。。



            elif n_j==2:

                 ## 【3.2 - 2组柱状图生成】

                    # 以下这些是针对2条图写的 （ 默认起码有2条图，不然怎么做。。对吧，如果真只有1个，肯定会报错。。）

                mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()    # 添加数据标签1
                mc_chart1.api[1].SeriesCollection(2).ApplyDataLabels()    # 添加数据标签2
                mc_chart1.api[1].SeriesCollection(1).Format.Fill.ForeColor.RGB = color_value[0] #or -16776961      # 样品1染色
                mc_chart1.api[1].SeriesCollection(2).Format.Fill.ForeColor.RGB = color_value[1] #or 15773696       # 样品2染色



                temp_min = min(min(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value))
                temp_max = max(max(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value))

                mc_chart1.api[1].Axes(2).MinimumScale= round(temp_min/2,1)    #比最小值小一半。。。  聪明啊~  11/05 第一次min生成一个list，需要再次min才能得到一个值
                mc_chart1.api[1].Axes(2).MaximumScale= round(temp_max*1.15,1)   #比最大值大一点点。。




            elif n_j==3 and n_j > control_count:        # 11/17 ★★★★★★★★★★★★ bug 0x：杨祖锐发现bug，当样品为2、测试条件为3时（样品数量＜测试条件时），柱状图出现行列互换bug。新增一个if条件来解决该问题

                 ## 【3.3 - 3组柱状图生成】

                    # 以下这些是针对3条图写的（且样品数量较少、而测试条件较多时可用，实际上操作方式和2条图一模一样）

                mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()    # 添加数据标签1
                mc_chart1.api[1].SeriesCollection(2).ApplyDataLabels()    # 添加数据标签2

                mc_chart1.api[1].SeriesCollection(1).Format.Fill.ForeColor.RGB = color_value[0] #or -16776961      # 样品1染色
                mc_chart1.api[1].SeriesCollection(2).Format.Fill.ForeColor.RGB = color_value[1] #or 15773696       # 样品2染色


                temp_min = min(min(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value))
                temp_max = max(max(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value))

                mc_chart1.api[1].Axes(2).MinimumScale= round(temp_min/2,1)     #比最小值小一半。。。  聪明啊~  11/05 第一次min生成一个list，需要再次min才能得到一个值
                mc_chart1.api[1].Axes(2).MaximumScale= round(temp_max*1.15,1)  #比最大值大一点点。。



            elif n_j==3 and n_j <= control_count:

                 ## 【3.3 - 3组柱状图生成】

                    # 以下这些是针对3条图写的

                mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()    # 添加数据标签1
                mc_chart1.api[1].SeriesCollection(2).ApplyDataLabels()    # 添加数据标签2
                mc_chart1.api[1].SeriesCollection(3).ApplyDataLabels()    # 添加数据标签2
                mc_chart1.api[1].SeriesCollection(1).Format.Fill.ForeColor.RGB = color_value[0] #or -16776961     # 样品1染色
                mc_chart1.api[1].SeriesCollection(2).Format.Fill.ForeColor.RGB = color_value[1] #or 15773696      # 样品2染色
                mc_chart1.api[1].SeriesCollection(3).Format.Fill.ForeColor.RGB = color_value[2] #or 5287936       # 样品3染色


                temp_min = min(min(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value))
                temp_max = max(max(xlwings.Range((temp_i+1,temp_j+1),(temp_i+control_count,temp_j+n_j)).value))

                mc_chart1.api[1].Axes(2).MinimumScale= round(temp_min/2,1)    #比最小值小一半。。。  聪明啊~  11/05 第一次min生成一个list，需要再次min才能得到一个值
                mc_chart1.api[1].Axes(2).MaximumScale= round(temp_max*1.15,1)   #比最大值大一点点。。


            else:
                pass

                 ## 【3.4 - 多组柱状图暂时不生成，预留一个窗口后续写】


            #=================================
            #【柱状图】位置参数特殊处理（区别于邱岑的条形图）

            Left_adj = Left + (Width + 5)*(p_i)









        elif '折线图' in chart_type:   # == '折线图':       # 11/09   ★★★★★★★★★★★★★★★★★★★★★★  【  折线图入口Entrance  】  ★★★★★★★★★★★★★★★★★★★★★★★★★

                # 折线图不需要区分几条柱子，统一expand() 即可；
                # 但是折线图需要增加 if ，来区分2个样品（默认）、3个样品（control_count=3）、4个样品的条件（control_count=4），这样以便于染色、打标签等



                     # 依次插入空白图表、设置图表类型
            mc_chart1 = mc_sht.charts.add(chart_left,chart_top,(355+30*(control_count-2))*1,211*1)       # ★★★ bug02（已解决）： 这里需要根据样品数量，来设定图表宽度；另外，为了匹配报告，设置为*0.66倍缩放 // 但11-04出现图片挂掉现象，只能取消缩放，改为*1
            mc_chart1.chart_type = 'line'   #设置一个折线图


            # # 这里仍然选中的是 【图表1 - 单元格】
            # temp_i = temp_list[p_i].row
            # temp_j = temp_list[p_i].column


            # 增加一个条件判断，区分单条柱状图 / 2条 / 多条柱状图
            #mc_book.selection.end('right').select()
            #n_j = mc_book.selection.column - temp_j       # ★★★★★★★★★★★★★★★★★★★  n_j 表示测试条件数量 ///  control_count 表示测试样品数量

            mc_chart1.set_source_data(mc_sht.range((temp_i,temp_j)).expand())


                            # 制表，格式重点  （慢慢来，不要搞太复杂，有需要再改）

            mc_chart1.api[1].SetElement(2)      #显示主标题
            mc_chart1.api[1].ChartTitle.Text = xlwings.Range((temp_i-4,temp_j)).value     # 修改主标题文本
            #mc_chart1.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Size = 13        # 修改图表主（正上方）标题大小  （最新版本，默认大小ok）
            mc_chart1.api[1].SetElement(309)  #显示 y轴 标题 （左侧，字体左侧旋转90°，别扭，但节约空间，英文适用）
            mc_chart1.api[1].Axes(2).AxisTitle.Text = xlwings.Range((temp_i-3,temp_j)).value          # 【左侧】 y轴标题（单位）的名字


            mc_chart1.api[1].SeriesCollection(1).Format.Line.ForeColor.RGB = color_value[0] #or -16776961     # 样品1染色
            mc_chart1.api[1].SeriesCollection(2).Format.Line.ForeColor.RGB = color_value[1] #or 15773696      # 样品2染色



            if control_count == 3:       # 2个对比竞品的话，第3条折线图染色
                mc_chart1.api[1].SeriesCollection(3).Format.Line.ForeColor.RGB = color_value[2] #or 5287936     # 样品3染色



            if control_count == 4:        # 3个对比竞品的话，第3条 / 第4条 折线图染色
                mc_chart1.api[1].SeriesCollection(3).Format.Line.ForeColor.RGB = color_value[2] #or 5287936     # 样品3染色
                mc_chart1.api[1].SeriesCollection(4).Format.Line.ForeColor.RGB = color_value[3] #or -16727809   # 样品4染色



            #=================================
            #【折线图】位置参数特殊处理（区别于邱岑的条形图）

            Left_adj = Left + (Width + 5)*(p_i)





        elif '条形图' in chart_type: # == '条形图':      # 2023/05/16    ★★★★★★★★★★★★★★★★★★★★★★  【  条图入口 Entrance  】  ★★★★★★★★★★★★★★★★★★★★★★★★★

            mc_chart1 = mc_sht.charts.add(chart_left,chart_top,(355+30*(control_count-2))*1,211*1)
            mc_chart1.chart_type = 'bar_clustered'

            # # 这里仍然选中的是 【图表1 - 单元格】
            # temp_i = temp_list[p_i].row
            # temp_j = temp_list[p_i].column



            #mc_chart1.set_source_data(mc_sht.range((temp_i,temp_j)).expand())   # 自动生成了一个默认的图表 chart // 含【主标题】、【图例】、【横坐标轴】、 【网格线】，这些都要去掉
            # 2023-07-04 更新： 邱岑的图表出现异常，修改了数据源的读取范围
            mc_chart1.set_source_data(mc_sht.range((temp_i,temp_j),(temp_i+control_count,temp_j+1)))



             # 感谢【Tools】 017 Docx 工具包。。。。 这些代码究竟是怎样写出来的。。。VBA的工程师们，写的时候完全没考虑过API使用感受。。。

             # ★★★★★★★★★ chart格式的调试接口，还好，这些代码是通用的，不受chart_type的影响（影响不大）★★★★★★★★★

            # =================================

            # 隐藏 【主标题】
            mc_chart1.api[1].SetElement(0)   # 0 = 隐藏 // SetElement(2) = 显示主标题（默认，与源数据中的标题字符相同）


             # 隐藏 【图例】
            mc_chart1.api[1].SetElement(100)    # 100 = 隐藏图例 // 101 = 显示图例（右侧）   102 = 显示图例（上侧）      103 = 显示图例（左侧）      104 = 显示图例（下侧）


             # 隐藏 【网络线】
            mc_chart1.api[1].SetElement(328)   #隐藏 y轴的网格线（横线）    (330) = 显示 y轴的网格线（横线）


             # 删除坐标轴【网络线】  （目前只会删除，删除后不知道如何反向操作，不纠结了，找不到API）
            mc_chart1.api[1].Axes(2).Delete()     #    Axes(1) = 左侧坐标轴      Axes(2) = 下方坐标轴


             # 添加【数据标签】
            mc_chart1.api[1].SeriesCollection(1).ApplyDataLabels()


            # 试试看，坐标轴的最大、最小值设置为 0-5
            mc_chart1.api[1].Axes(2).MinimumScale = 0  # 比最小值小一半。。。  聪明啊~  11/05 第一次min生成一个list，需要再次min才能得到一个值
            mc_chart1.api[1].Axes(2).MaximumScale = 5  # 比最大值大一点点。。

            # 至此，暂时格式和邱岑做的表格完全相同了。



            #=================================

            #接下来搞定排版问题  // 大循环是，一共有 p_i 个chart 需要遍历  //  一共有 control_count 个样品数量（条）

            temp_l1 = (0,1,0,1,0,1,0,1,0,1,0,1)   # 设定一个递增Tuple，来实现循环位置公式， 直接调用 temp_l1[p_i] / temp_t1[p_i]   即可 // 预留冗余，即使12个图表也能排版，虽然超出边界了
            temp_t1 = (0,0,1,1,2,2,3,3,4,4,5,5)

            Left_adj =128 + temp_l1[p_i]*292   #针对条形图，单独定义大小和位置参数，来复制邱岑的排版格式
            Top =80 + temp_t1[p_i]*160
            Height =150                 # 粘贴到PPT中的大小相对比较固定  // 位置需要递增而已
            Width =190



            #=================================

            #染色问题  低于 <= 3.5 分 = 黑色   //  条形图的上下位置摆放是反的，但调用顺序是对的，搞定了  //  染色边框问题，没有再出现

            for l in range(0,control_count):

                if mc_sht.cells(temp_i+l+1,temp_j+1).value <= 3.5:

                    mc_chart1.api[1].SeriesCollection(1).Points(l+1).Format.Fill.ForeColor.RGB = black

                else:

                    mc_chart1.api[1].SeriesCollection(1).Points(l+1).Format.Fill.ForeColor.RGB = red




            #【 针对条形图，继续增加一个模块，因为 PPT中的（圆圈）和（标题文字）也属于chart的一部分，】

            #=================================

              # 虚线圆圈
            t_Left = Left_adj - 90   # chart位置 左移 90 像素     #它们的相对位置关系都是固定的，给个差值即可  //  因为chart位置已实现递增
            t_Top = Top + 18   # chart位置 下移 18 像素

            t_circle = Circle_Shape(mc_slide,Left = t_Left, Top = t_Top)


              # 文字

            t_Left = Left_adj - 95   # chart位置 左移 90 像素     #它们的相对位置关系都是固定的，给个差值即可  //  因为chart位置已实现递增
            t_Top = Top + 58   # chart位置 下移 18 像素
            Text = mc_sht.cells(temp_i-4,temp_j).value


            t_title = Title_3(mc_slide,Left = t_Left,Top = t_Top, Text = Text)






            #【 针对条形图，final增加一个模块。。。 把图片自动排序进去，这样可以省略部分工作量】

            #=================================

            if mc_sht.cells(temp_i,temp_j+2).value != None and '图片' in mc_sht.cells(temp_i,temp_j+2).value:   # 假如条形图里有附带图片，那么可以调用这个模块

                mc_pic(mc_sht.cells(temp_i+1,temp_j+2),  temp_slide = mc_slide,  Left = 660,  Top = 80,  scale = 0.2)    # ★★★★ 条形图对应的试穿图片大小比例控制  // 位置控制    bug （这里不用微调图片，要不别赋值了？也不行，最后取消了这个函数的双重Return才解决）★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

                                                                                                                                       # ret=0 不用写，默认=0
            else:    # 如果没有放置图片，那么就pass
                pass






        #elif:     # 这里计划增加杨祖锐的矩阵图，完全是另外一种函数，无需制表格，只要摆放位置和控制大小

                    # 可能不行，因为是完全不同的处理流程，需要全新的处理流程。。。 估计重新写个函数更靠谱// 到时再看看

                     # 如有必要，可以再增加其他的图表类型，★★★★★★★★★★★★★★★★★★★  先预留接口 【 雷达图 / 散点图 / 高级图表  】







        else:
            pass      # 万一输入错误，无法制图，直接跳过，到下一个图表







        #【part 3 图表都做好了，需要复制粘贴到指定位置了 mc_slide 】

        # 输出环节完全可以移植到这个循环中，避免误选中其他闲杂chart

        temp_list[p_i].select() # 先选中相关区域【图表1】，避免看不见图片、无法复制。。。 多2的bug


       # 二选一，如果还是出现无法复制的错误，再说吧。。。  GPT已经把原因解释了，必须显示才能复制  // 继续控制屏幕缩放，来避免复制失败。。
        #mc_app = xlwings.apps.active
        #mc_app.api.ActiveWindow.Zoom = 100

        Excel_zoom(mc_sht)
        chart_cell.select()
        #mc_book.selection.offset(row_offset=control_count*2,column_offset=0).select()

        # for _ in range(3):  # 尝试复制操作 3 次
        #     try:
        #         mc_chart1.api[0].Copy()
        #         break  # 如果复制成功，跳出循环
        #     except pywintypes.com_error:
        #         print("复制失败，1 秒后重试")
        #         time.sleep(1)  # 等待 1 秒后重试
        # else:
        #     print("复制操作失败")

        mc_chart1.api[0].Copy()






        time.sleep(random.random()*delay)

        mc_shape = mc_slide.Shapes.Paste()

        #mc_shape.Left = Left + Width*(p_i) + 5       # 只有这一个指标需要递增，实现排列
        mc_shape.Left = Left_adj                          # 由于每个chart的位置需要单独计算，分散到各个chart的if条件中去了
                                                            # 【2024】 Left 这个数据比较特殊，调试好几次，最终为了解决第三个chart跑出屏幕的问题，才发现递增条件错误，因此将Left 改为 Left_adj
        # print(f'    当前图表粘贴位置：Left = {Left_adj}\n')          # 调试位置错误的问题 （3个图表跑出屏幕的问题），已解决

        mc_shape.Top = Top

        mc_shape.Height = Height

        mc_shape.Width = Width










    return Left_adj,Top,Height,Width    # 最后还是需要输出，来为文字描述预留位置参数 // 这里输出最后一个chart的位置大小参数

    #print('一共成功输出' + str(p_i+1) + '个图表！')











#  【Excel + PPT matrix 生成函数 】

def make_matrix(mc_sht,mc_slide):

    ''' 终于来到这个函数，绘制矩阵图形 '''

    # 【步骤一】：传统定位操作，定位，输入坐标系关键参数

    mc_book = mc_sht.book     # 不想搞太多参数，直接用sheet的parent

    value_list = []   # 用来存储值（减震G，回弹R） / （弯折，扭转） ★★★★★★★★★★★★★★★★★★★★★★★★★★★★★

    pos_list = []    # 用来存储position 位置参数   ////  参照系、坐标系 frame of reference  //

    name_list = []   # 用来存储产品名称

    pic_list = []    # 用来储存照片的单元格

    target = "Matrix矩阵图"          # 只有1个图，直接名字复杂点，搜素它即可，不用for
    find = search(mc_sht,target)
    find.select()

    temp_i = find.row
    temp_j = find.column


    # 这里一共有 【n_i】 个样品数量，然后只有2个实验条件（数据栏）
    mc_book.selection.end('down').select()
    n_i  = mc_book.selection.row - temp_i







    #【步骤二】：设定关键坐标系参数，每个Matrix的原点都不一样，如果出现问题（例如摆放位置出现偏差），到时需要校准，借用【小工具】

    chart_type = mc_sht.cells(temp_i-2,temp_j).value

    if '减震回弹矩阵-篮球' in chart_type: # == '减震回弹矩阵-篮球':


        # 计算步骤挪到小工具里去了，这里直接给出【坐标原点】 和 【单位】
        delt_l = 3222.727272727273
        zero_l = (-1107.6363636363637,82.0)
        delt_t = -51.0           # 但是，正负早就体现在单位中了，每增加1单位，像素值是变大还是变小，其实【小工具】中早就计算出来了
        zero_t = (70.0,793.0)
        rev= -1                  #手工列一个rev值，Top轴在上面时，rev= -1，Top轴在下面，rev= 1

        adj = 0                 #手工设置一个微调值，来区分跑步和篮球


    elif '弯折扭转矩阵-篮球' in chart_type: # == '弯折扭转矩阵-篮球':

        # 计算结果
        delt_l = 41.1764705882353
        zero_l = (-293.764705882353,415.0)
        delt_t = -31.0
        zero_t = (61.0,508.0)
        rev= 1                  #手工列一个rev值，Top轴在上面时，rev= -1，Top轴在下面，rev= 1

        adj = 0                 #手工设置一个微调值，来区分跑步和篮球



    elif '减震回弹矩阵-跑步' in chart_type:  # == '减震回弹矩阵-跑步':

        # 计算结果
        delt_l = 5200.0
        zero_l = (-2275.0,81.0)
        delt_t = -51.0
        zero_t = (69.0,691.0)
        rev= -1                  #手工列一个rev值，Top轴在上面时，rev= -1，Top轴在下面，rev= 1

        adj = -6                 #手工设置一个微调值，来区分跑步和篮球


    elif '弯折扭转矩阵-跑步' in chart_type: # == '弯折扭转矩阵-跑步':

       # 计算结果
        delt_l = 40.875
        zero_l = (-331.625,414.0)
        delt_t = -32.42857142857143
        zero_t = (61.0,418.42857142857144)
        rev= 1                               #手工列一个rev值，Top轴在上面时，rev= -1，Top轴在下面，rev= 1

        adj = -6                      #手工设置一个微调值，来区分跑步和篮球



    elif '步长重心振幅矩阵-跑步' in chart_type:  # == '步长重心振幅矩阵-跑步':
        # 【小工具】计算结果
        delt_l = 52.357142857142854
        zero_l = (-7526.142857142857, 415)
        delt_t = -201.53846153846143
        zero_t = (61, 2118.230769230768)
        rev = 1

        adj = -6



    elif '重量厚度矩阵-跑步' in chart_type:  # == '重量厚度矩阵-跑步':
        # 【小工具】计算结果
        delt_l = -12.637931034482758
        zero_l = (3024.7241379310344, 415)
        delt_t = -17.466666666666665
        zero_t = (61, 909.0)
        rev = 1

        adj = -6



    else:
      print('矩阵关键参数尚未定义，请运行【小工具】重新计算！ 函数终止')
      return None





    # 【步骤三】 根据小工具计算的坐标系关键参数，计算对应的摆放位置（position）， 放在list中，用 tuple 的格式储存

    for n in range(0,n_i):

        value_L = mc_sht.cells(temp_i + 1 + n, temp_j + 1).value
        value_T = mc_sht.cells(temp_i + 1 + n, temp_j + 2).value


        # 然后才是计算图片摆放位置计算     # 【步骤三】： 随后，给定一个值得（value_L，value_T），计算出图片需要摆放的位置（Left, Top)

        Left = zero_l[0] +  value_L * delt_l   # 公式为什么变这样，我也不太理解。。。。
        Top  = zero_t[1] +  value_T * delt_t   # 调试发现  G用减法，R% 用加法   （未解之谜，ing.....） 其实小工具计算出来的【delt单位】已经包含正负了，没问题了



        # 储存到列表里，后面摆放的时候调用

        value_list.append((value_L,value_T))
        pos_list.append((Left, Top))              # ★ 最为重要的位置参数，所有的位置关系都基于它
        name_list.append(mc_sht.cells(temp_i + 1 + n, temp_j).value)
        pic_list.append(mc_sht.cells(temp_i + 1 + n, temp_j+3))




    # 【步骤四】 参数都计算完成了，接下来就是摆放图片了。这里直接调用 mc_pic()函数，不再重写 // 另外，加工下输出文字，美化UI


    for i in range(0, len(pic_list)):

        temp_group = []          # 05-31 根据杨祖锐的反馈，新增一个组合，便于调整位置，避免扎堆覆盖的问题
                                   # //  # 清零动作，这个list仅用来储存 Name，并且固定只有3个元素
                                    # 注意，mc_pic函数嵌套时出现过双重return的问题，这里可以return一个shape，而在make_chart中则不会return

        # 【1】摆放鞋款照片——  这是一个 PPT-Shape 对象，使用Shape 的 API  // scale = 46 / 540 ≈ 0.086
        temp_pic = mc_pic(pic_list[i],  temp_slide = mc_slide,  Left = pos_list[i][0],  Top = pos_list[i][1],  scale = 0.05,ret=1)  # 矩阵鞋款图的比例 ★★★★★★★★★★★★★★★★★★★★★★★★★★★

        # 【2】居中—— 将图片居中的公式（可挪用★★★★）
        temp_pic.Left = temp_pic.Left - temp_pic.Width * 0.5
        temp_pic.Top = temp_pic.Top - temp_pic.Height * 0.5

        temp_group.append(temp_pic.Name)   # 【0】 图片



        # 【3】标注名称、参数值 —— 又要重新定义一个文本 class 了吗？ no！ 直接微调就行


         # (1)鞋款名称                                                        # 矩阵图小文字的数据展示 位置参数  / 字体在Class 里 ★★★★★★★★★★★★★★★★★★★★★★★★★★★

        s_Left = temp_pic.Left + 0
        s_Top = temp_pic.Top - 4 + adj
        Text = name_list[i]
        temp = Text_small(mc_slide,s_Left,s_Top,Text)
        temp_group.append(temp.shape.Name)      # 【1】鞋款名称




         # (2)性能数值，需要慎重对待，加工下，美化数据
        s_Left = temp_pic.Left - 5                         # 右移？？个像素即可
        s_Top = temp_pic.Top +  temp_pic.Height - 8        # 下移？？个像素即可 ？？           ★★★★★★★★★★★★★★★★
        if '减震回弹' in chart_type: # == '减震回弹矩阵-篮球' or  chart_type =='减震回弹矩阵-跑步':

            t1 = str(round(mc_sht.cells(temp_i + 1 + i, temp_j + 1).value*100,1)) +'%' # L轴是回弹率
            t2 = str(round(mc_sht.cells(temp_i + 1 + i, temp_j + 2).value,1))      # T轴是减震G

        #elif chart_type == '弯折扭转矩阵-跑步':


        else:  # 其他暂时都是保留1位小数，先这样了，有需要再在上面添加  //  先L轴保留1位小数（t1）， 再T轴保留2位小数，这个一般是扭转力矩（t2）

            t1 = str(round(mc_sht.cells(temp_i + 1 + i, temp_j + 1).value,1))
            t2 = str(round(mc_sht.cells(temp_i + 1 + i, temp_j + 2).value,2))

        Text = '（' +  t1 +'，' +  t2 + '）'
        temp = Text_small(mc_slide,s_Left,s_Top,Text)
        temp_group.append(temp.shape.Name)      # 【2】性能参数

        # 然后，将这3个元素组合成一个group，试试看
        mc_slide.Shapes.Range((temp_group[0],temp_group[1],temp_group[2])).Group()

        # 大量矩阵图时，删除数据和鞋款名称，仅保留图片？
        #mc_slide.Shapes(temp_group[1]).Delete()   # 删除名称
        mc_slide.Shapes(temp_group[2]).Delete()   # 删除性能参数








        # 【4】描线（虚线、浅灰色、一条横、一条竖） ——— 这里又要拜托class出马了。。

        # 先定义【水平横线】的参数    # 有道云笔记中计算得到的 横线起止点坐标数值
        BeginX = zero_t [0]
        BeginY = pos_list[i][1]
        EndX  = pos_list[i][0] - 0.5*temp_pic.Width   # （Top轴始终在左边，没问题）
        EndY  = pos_list[i][1]


        # 再插入一条线  【水平横线】
        line_h = Line_Shape(mc_slide,BeginX,BeginY,EndX,EndY)    # 这里是一个Class，类似文本框，需要用line_h.shape来调用
        line_h.shape.ZOrder(1)   # 0 = 顶层？     # 1 = 底层？   是的，确定！！！  其他的先不管了





        # 再定义一条  【垂直纵线】的参数    # 有道云笔记中计算得到的 横线起止点坐标数值
        BeginX = pos_list[i][0]
        BeginY = zero_l [1]
        EndX  = pos_list[i][0]
        EndY  = pos_list[i][1] + rev * 0.5 * temp_pic.Height    #（Left轴偶尔在上，偶尔在下，始终从坐标轴出发，需要引入rev参数）

                                                            # =  图片摆放位置参数Top + rev* 0.5*图片高度  ，引入rev值，Top轴在上面时，rev= -1，Top轴在下面，rev= -1

        # 先插入一条线  【垂直纵线】
        line_v = Line_Shape(mc_slide,BeginX,BeginY,EndX,EndY)
        line_v.shape.ZOrder(1)  # 0 = 顶层？     # 1 = 底层？   是的，确定！！！  其他的先不管了


#print('看看chat之前的函数是否成功导入？')





#  【Excel 函数 】

# 完全不理解。。 为什么这样就能运行 ？ 代码放在function import区域就无法运行？ 不过，总算解决了
# mc_book = xlwings.books.active
# mc_sht0=mc_book.sheets['基础信息']
# target="测试样品"
# sample_name = search(mc_sht0,target,column_offset=1).value

#print('看看sample_name有没有被找到？AG')







# ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖
# def chat_GPT(*cell_list):   # 2023年为这个函数头疼不已，2024，我终于拥有了自己的GPT， real——GPT-4-turbo！
#
#     ''' 没想到这么快就开始挑战，文字机器人~  史上报错次数最多的函数，一直 debugger，连 Debugger 都卡死了'''
#     global sample_name             # 跨文件的 global ， 希望管用。。。  这样就可以不用定义，直接调用了。 吧？
#     chat:str = ''     # 这里需要挪到循环外部，避免被清空 // 备注下，来避免下划线提示。。。（PEB8代码规范。。）
#     chat2:str = ''
#     chat3:str = ''
#
#     # ----------------------*【第一部分：基础准备】*--------------------------------
#
#     # 由于（*args1,*args2）的方式违规了，只好将参数简化为  chat_GPT(mc_cell)  不纠结，赶紧继续~ // 最终还是升级为 *cell_list，因为默认 1页sheet = 1页ppt，需要生成一段chat
#     #for ct in range(0,len(cell_list)):
#     for mc_cell in cell_list:
#
#         #mc_cell = cell_list[ct]
#
#         mc_cell.select()
#         mc_sht = mc_cell.sheet
#         mc_book = mc_sht.book
#
#         # 构建几个关键的dict ，让实验结论看起来更像自然语言，例如
#             # chat_rev：用来保存正 / 负相关信息，例如G是越小越好，R%是越大越好
#             # chat_perform：用来保存性能描述，例如【减震：减震性能】【回弹：回弹性能】
#             #...........（后续）逐渐丰富keywords，增加各种dict来丰富语言结构
#             # attention！！！ 这些dict最好建议用同样的排序格式来生成，避免混淆。。 虽然dict和list不同，不存在顺序问题，但格式统一方便管理
#
#
#         # 注意：这两个dict 的 顺序和key是共享的，因此，最好严格同步
#         dic_chat_rev = {
#             '减震': -1, \
#             '回弹': 1, \
#             '弯折': 1, \
#             '扭转': 1,\
#             '（其他）':1
#         }       # 其他默认 = 1 （正相关）； 其实只需要定负相关的key：value 即可
#
#
#         dict_chat_perform = {
#             '减震': '减震性能', \
#             '回弹': '回弹性能', \
#             '弯折': '弯折力值', \
#             '扭转': '抗扭转性能',\
#             '（其他）': str(mc_sht.cells(temp_row, temp_column +1+ t).value) + '测试结果 '
#         }                              # 这个dict需要好好扩充下  ///  当【性能】关键字未在字典中定义时，默认会跑到最后一栏（其他），出现这个bug时再来扩充这个dict  ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★
#                                        # 当找不到 / 变成（其他）性能时，使用公式来获取该列的【测试项目】名称，代替性能描述，应该能糊弄过关~
#
#
#
#         list_deco1=[
#             ('的',''),\
#             ('在','测试中')\
#         ]
#
#
#
#         #增加一个样品数量识别模块(control_count)，每个chart都单独识别一遍    ★★
#         temp_row = mc_book.selection.row
#         temp_column = mc_book.selection.column
#         # 在其他函数中是使用temp_i,temp_j代替，而编写chat_GPT时忘记统一了 // 历史遗留bug , 希望不要出问题
#         temp_i = temp_row
#         temp_j = temp_column
#
#
#
#         mc_book.selection.end('down').select()
#         control_count = mc_book.selection.row - temp_row  # 总样品数量 仍然放在 control_count 中，每个chart都单独识别（原来的代码不用改），并且自由了
#
#         # 增加一个实验条件模块// 2023-05 因为【柱状图】、【折线图】都用得到，并且需要为chat_GPT函数做准备
#         mc_cell.select()
#         mc_book.selection.end('right').select()
#         n_j = mc_book.selection.column - temp_column  # ★★★  n_j 表示测试条件数量 ///  control_count 表示测试样品数量
#
#
#
#
#
#         # ----------------------*【第二部分：数据构建 - 装入list中 // 识别测试样和对比样】*--------------------------------
#
#         # 【2.1】 通过简单的循环，构建chat_sample 测试样品 / chat_test_item 测试项目 两个list
#         # 同时还获取了【关键试验样品（例如：飞影PB）】 的摆放位置，万一找不到，那就默认第一个样品
#         test_sample = list()              # = []
#         test_item = list()           # = []
#         compare_key = list()         # 对比样品可以有多个。。。。
#
#         sample_key = 0                 # sample 只有唯一的一个！ 【找到实验样品在list中的位置】  大家书写可能都不规范，（假设找不到 / 那就不找了），那么默认第一个sample是【实验样品】 //
#                                         # 另外，通过多个条件的判断，一定会有一个sample_key，默认值是0 （第一个样品）； 而有可能不存在 compare_key，这个时候就跟均值比？
#
#         for t in range(0,control_count):
#             test_sample.append(mc_sht.cells(temp_row+1+t,temp_column))                   #    test_sample = [样品1，样品2，样品3...】
#             if sample_name == mc_sht.cells(temp_row+1+t,temp_column).value:              # ★★★ 【实验样品】key_sample识别，【in】 or 【==】，需测试
#                 sample_key = t                                                         # 得到了 【实验样品】 在列表中的位置，注意 0 代表 【第一个数据】，和 range 的方式一样
#             elif '测试' in mc_sht.cells(temp_row+1+t,temp_column-1).value:
#                 sample_key = t
#             elif '对比' in mc_sht.cells(temp_row+1+t,temp_column-1).value:
#                 compare_key.append(t)
#
#         if compare_key == []:     # 为 compare_key 设定默认值？  //好像不需要，没有就不对比了呀。。 // 还是先试试吧
#             compare_key.append(control_count)
#
#
#         for t in range(0,n_j):
#             test_item.append(mc_sht.cells(temp_row,temp_column+1+t))                    #    test_item = [测试项目1，测试项目2，测试项目3...】
#
#
#
#
#         # 【2.2】  构建 sample_value 测试数据 list
#         # 构建均值/最大值/最小值list , 均需要用 n_j 进行遍历
#         chat_temp_list = list()       # 用来临时存储【减震 / 回弹】实验条件下，单列的值
#         sample_value = list()
#
#         average_value = list()
#         max_value = list()
#         min_value = list()
#
#         for c in range(0,n_j):
#             for r in range(0,control_count):                  # 这个循环、 以及这个list用来临时储存第1/2/3/4...列的值，例如   [减震1，减震2，减震3] / [ 回弹1，回弹2，回弹3]
#                 chat_temp_list.append(mc_sht.cells(temp_row+1+r,temp_column+1+c).value)
#
#             average_value.append(sum(chat_temp_list) / len(chat_temp_list))
#             max_value.append(max(chat_temp_list))                # 【max1,max2,max3...】  对应每个list 的最大值 ，使用n_j遍历
#             min_value.append(min(chat_temp_list))               # 【min1,min2,min3...】  对应每个list 的最小值 ，使用n_j遍历
#             sample_value.append(chat_temp_list)   # ★★★  所有的数据值，按照 sample_value=[  [减震1，减震2，减震3]，[ 回弹1，回弹2，回弹3]  ] 的 form , 需要用 n_j[i] 进行遍历
#
#             chat_temp_list = []        # 下一个（n_j） 实验条件循环赋值
#
#             # 至此，准备的数据工作完成了，测试样品、实验条件、数值都已经装入list中。可以下一步进行对比、判断了
#
#
#
#
#
#         # ----------------------*【第三部分：处理单个chart，生成评论文字。目前是调试阶段，假设有2个实验条件：减震 / 回弹】*--------------------------------
#
#         # chat 是最终的输出目标 // 最终的循环是按照 n_j的方式循环，因此构建的list需要严格按照同样的形式
#         #chat = ''   # 这里需要挪到循环外部，避免被清空
#
#         for t in range(0,n_j):     # 按照实验条件进行遍历
#
#             # 首先假设遍历来到了 【减震】这一列
#             # 在该循环内，我的最终目标是： 【飞影PB】 的 【减震性能/回弹性能】 【具有明显优势，在此次评测产品中表现最优 / 表现较好，高于同类均值 + xx%  / 表现较为一般，与同类竞品无明显优势】。
#
#
#             # --------------  首先是对该列的【实验条件】的线性关系进行初步判断： 正 / 负相关？ 这个非常重要 （值越大越好，还是越小越好？）
#             t_key = mc_sht.cells(temp_row,temp_column+1+t).value       # 假设【测试项目1】 = t_key = '减震G数据'  = 大家制表时填写的单元格的测试项目名称
#             chat_rev = 1                                                    # 由于大家制表时，不一定严格按照规范.. 先设置一个默认值，万一查找失败，默认 rev = 1 程序不会报错（虽然结果是反的，但可以手工改文字）
#             for key in dic_chat_rev:
#                 if key in t_key:                                     #  这里假设：如果 key ‘减震’  in   t_key '减震G数据'
#                     chat_rev = dic_chat_rev[key]                          #  key = ‘减震’ 时，查找字典发现， rev = -1
#
#                     #p_key = key
#                     break   # 把 key 拯救出来试试？ 不行再试上一句 一旦找到就break，不继续往下查找了/// 如果找不到，那么默认值就是1了
#
#
#
#
#
#
#             # ---------------------------------*【 final 核心文字部分：如何体验语言的生动、多样性，需要引入大量随机函数，并丰富dict语言库 】*----------------------------------------
#
#             campare_value = mc_sht.cells(temp_row + 1 + sample_key, temp_column + 1 + t).value        # campare_value 就是 【关键实验样品】的值，用来去跟其他竞品PK
#             key_sample_name = mc_sht.cells(temp_row + 1 + sample_key, temp_column).value            # key_sample_name =   【关键实验样品】的名称 // 正常是一样的，但万一出现刁钻的bug（例如我调试时），每张chart的关键样品名都不一样
#             #key_sample_name = sample_name                        # 90% 情况下都是成立的，但不成立时，也需要预留容错空间
#
#
#             # 前面已构建了 关键【实验样品】的位置，位于range(0,control_count) // test_sample 中的【sample_key】位置
#             if chat_rev == 1:        # 正相关条件，如【R%】   //
#
#                 # 如果【飞影PB】性能最优：
#                 if campare_value == max_value[t]:
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key]  + '具有明显优势，在此次评测产品中表现最优\r'
#
#
#                 # 如果【飞影PB】性能优于均值
#                 elif campare_value <= max_value[t] and campare_value >= average_value[t]:
#                     compare_per = abs(100 * round(((campare_value - average_value[t]) / average_value[t]), 1))
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key] + '表现较好，高于均值' + str(compare_per) + '%' + '\r'
#
#
#                 # 如果【飞影PB】性能低于均值
#                 elif  campare_value < average_value[t] or campare_value == min_value[t]:            # campare_value >= average_value[t]
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key] + '表现较为一般，与同类竞品相比无明显优势\r'
#
#
#
#                 # 如果【飞影PB】性能值超出范围 / 出错
#                 else:
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key] + '测试值超出范围，系统无法评价！\r'
#                     # print('值超出范围，无法评价！')
#
#
#
#             # 在该循环内，我的最终目标是： 【飞影PB】 的 【减震性能/回弹性能】 【具有明显优势，在此次评测产品中表现最优 / 表现较好，高于同类均值 + xx%  / 表现较为一般，与同类竞品无明显优势】。
#
#
#             else:    # 负相关条件，如【减震G】   chat_rev = -1 条件 /// 严格上讲，只可能存在两种情况，  1 / -1
#
#                 # 如果【飞影PB】性能最优：
#                 if campare_value == min_value[t]:
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key]  + '具有明显优势，在此次评测产品中表现最优\r'
#
#
#                 # 如果【飞影PB】性能优于均值
#                 elif campare_value >= min_value[t] and campare_value <= average_value[t]:
#                     compare_per = abs(100 * round(((campare_value - average_value[t]) / average_value[t]), 1))
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key] + '表现较好，高于均值' + str(compare_per) + '%' + '\r'
#
#
#                 # 如果【飞影PB】性能低于均值
#                 elif  campare_value > average_value[t] or campare_value == max_value[t]:            # campare_value >= average_value[t]
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key] + '表现较为一般，与同类竞品相比无明显优势\r'
#
#
#                 # 如果【飞影PB】性能值超出范围 / 出错
#                 else:
#                     chat += key + '：' + key_sample_name + '的' + dict_chat_perform[key] + '测试值超出范围，系统无法评价！\r'
#                     # print('值超出范围，无法评价！')
#
#
#     return chat           # 最后输出文字。。。呕心沥血呀
# ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖













# 【Python】 辅助函数，把上面的单元格依次pop出来，结合pic函数使用  ///  结果竟然实现不了，意味着改写成lambda函数，也无法实现这个功能。。。。 因为一旦函数执行遇到return，就结束了，意味着还没开始第二次循环，就结束了

   # 能否换个思路？ 不要return， 要pop。。。 结果的确pop了，但也丢失了。。。 都不行，因此，依次返回值的小问题，竟然也难倒了我。。。。。。。。。。

   # ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖ ✖         ing............

def pop_list(temp_list):

    #global temp_list

    while temp_list != []:
        temp_list.pop()




def rank_range(*l):
    '''简单的自制排序函数，将list的顺序输出为一个新的rank_list'''

    #rank_temp = l    # 这样生成的是Tuple  // 而我需要重构一个list
    rank_temp = []
    rank_list = []
    for t in l:
        rank_temp.append(t)    #rank_find.append(t)
    #rank_temp = rank_find   # copy 2个 新的list，一个rank_temp用来remove，一个rank_find用来index  ///  结果这样做，两个list竟然变成了双胞胎（list会联动）。。。。  奇葩呀，直接用tuple吧

    while rank_temp != []:
        rank_list.append(l.index(max(rank_temp)))
        rank_temp.remove(max(rank_temp))

    return rank_list   # 最后return的的确是个list。一个简单的排序函数，竟然写了这么久。。






















# 【PPT】 辅助函数，当字体可以大概率沿用，只需要微调时，对字体进行微调，不再重新设置Class，毕竟做那么多Class也没有用  ing...................................

def adj(text_range,size,c_color,b_color,bold=0,trs=0):

    # --------- 格式调整（微调） -----------
    text_range.Font.Bold = bold                            # 加粗   # 【0 = 不加粗  /  1 = 加粗】
    text_range.Font.Color = c_color                        # 字体颜色
    text_range.Parent.Parent.Fill.ForeColor.RGB = b_color  # 背景填充颜色
    mc_shape.Fill.Transparency = 0.39                      # 【 =1 完全透明   //   = 0 完全不透明    //    =0.39 常用黑色半透明 】
    text_range.Font.Size = size





def RGB_to_Hex_to_Dec(rgb):  # 这里需要用Tuple来存储rgb, 例如 rgb = (0, 176, 80)

    '''这里定义一个转换函数（CDSN），将RGB的（255,255,255）转换为 0xffffff，方便后续的操作'''

    color = '0x'

    for i in rgb:

        num = int(i)

        color += str(hex(num))[-2:].replace('x','0').upper()   # 这句代码让我深受震撼。。。 虽然我看不懂，很管用！！

    #return color 调试成功

    temp = str(color)

    #X = ast.literal_eval(temp)  # 是的 出问题的是这个代码。。。 VBA = 452960000  // 而我计算出来 = 45296  这。。。   16752384.0


    #调试代码
    #mc_cell = mc_book.selection
    #mc_cell.api.Interior.Color   可以返回一个整数


    X = int(temp[2:],16)

    return X   # 没问题，返回的是整数





# 【PPT】文件存储函数，借鉴 Excel

def ppt_save(mc_ppt,sample_name,mc_path):   # Excel_wb + 文件名   // 注意这个sample_name 是局部变量，作为参数传入，因此跟 global 的 sample_name 完全没关系
    #import os
    #from datetime import datetime
    today = str(datetime.today())
    today = today[0:10]

    j = 0
    while j < 20:
        if os.path.exists(mc_path + '\\' + today + ' ' +  '【' + sample_name + '】'  +  '测试报告' + 'v 1.' + str(j) + '.pptx') == True:
            j = j + 1
        else:
            mc_ppt.SaveAs(mc_path + '\\' + today + ' ' +  '【' + sample_name + '】'  +  '测试报告' + 'v 1.' + str(j) + '.pptx')
            break




# 针对路径的操作，真的是太考验基本功了，我一直没学会，只能试错


##>>> mc_ppt.SaveAs(r'D:\p\2023-04-24 py for PPT II adv\test')
##
##>>> os.path.exists(r'D:\p\2023-04-24 py for PPT II adv\test.pptx')
##True



##>>> mc_path
##'C:\\Users\\Administrator\\AppData\\Local\\Programs\\Python\\Python38\\lib'
##>>> print(mc_path)
##C:\Users\Administrator\AppData\Local\Programs\Python\Python38\lib
##






#print(mc_key)
print('Function文件导入完成！\n')   # 原来 import 会将整个程序 run 一遍 //  学习了














##=====================================================================================================================



##失败案例1：把GPT-5-mini当成函数使用。。。
##
##def gen_mc_prompt_question_data(mc_cell):  # 把GPT-5-mini当成函数使用，nice   ///  结果GPT失败了！！！  最终还是需要靠万能的 re 和 YYDS 的 Claude
##
##    '''
##    基于找到的单元格 也就是上面的函数
##
##    将问卷的内容打包成将混合数据，便于后续发送给 GPT-5-mini 处理成【问卷数据】
##    '''
##
##    raw_data = mc_cell.api.CurrentRegion.Value
##
##    # 逐步开始构建 prompt //  注意，每次请求GPT，都需要将 prompt 清零，以免信息冗余
##    mc_prompt = ''
##
##    # 将GPT当函数使。。
##
##    mc_prompt += (
##
##    f"我回收了一份Excel问卷，我得到的数据通常是下面这种格式："
##
##    f"【你的任务】\r 你需要帮我将我发给你的原始数据进行简单分析，提取其中关键的性能评分，问卷中其他内容（不是关于性能评分的内容）你必须忽略掉。"
##
##    f"最终你要生成一个优雅而整齐的list，返回给我。注意，你给我的回答必须是一个list，不要包含任何其他多余的东西。\r "
##
##    f"我示范给你看。假设我发送给你的原始数据是：\r "
##
##    f'''’序号	提交答卷时间	所用时间	来源	来源详情	来自IP	总分	1、姓名（昵称）	2、您穿着这款鞋子的总里程（km）	3、您穿着这款鞋子单次完成的最长里程（km）	4、您穿着这款鞋子最常见的使用配速	5、根据您的使用体验对鞋子的各项性能进行评分（1分最差，10分最好）—缓震性/舒适度	5、回弹性/推进性	5、稳定性/支撑性	5、抓地力/牵引力	5、长距离抗衰减性	5、灵活性/反馈度	5、 包裹性与贴合度	5、透气性	5、重量	5、耐久度	6、尺码选择：	7、穿着过程中，您是否遇到以下问题？(可多选)(磨脚 (请注明位置：脚后跟、脚踝、脚趾、足弓等))	7、(水泡 (请注明位置))	7、(黑趾甲/趾甲挤压)	7、(鞋内滑动/不跟脚)	7、(鞋带易松脱/难系紧)	7、(支撑不足导致脚踝不适)	7、(缓震不足导致膝盖/脚底疼痛)	7、(鞋面过早破损/开胶)	7、(鞋底齿纹过早磨损/脱落)	7、(以上均无)	8、(1)与您穿过的其他同类跑鞋相比，这双鞋的主要优势和劣势分别是什么？优势：___	8、(2)劣势：___	9、提供穿着后的鞋子照片（特别是磨损部位）	10、鞋子照片（多张备用1）	11、鞋子照片（多张备用2）	12、提供穿着鞋子跑步的数据截图	13、数据截图（多张备用1）	14、综合所有体验，您是否会考虑购买这款鞋（假设最终上市）	15、其他任何未尽的想法或建议？
##    1	2025/11/11 14:21:33	345秒	微信	N/A	39.144.251.51(福建-厦门)	64	隋晗1	150	3	2	9	3	8	7	4	6	6	7	6	8	1	0	0	0	1	0	0	0	0	0	0	初上脚有非常软的脚感，泡棉形变大，缓震极佳，总体舒适性好，在有氧跑的情景下，泄力感不强	跑步5-8km之后，前掌衰减明显，有变硬的趋势，且坡差过大，包裹和锁定偏差，用力调整鞋带后尚可接受	https://alifile.sojump.cn/338076422_1_q9_20251111142033679JJHG47.jpg?Expires=1770618115&OSSAccessKeyId=LTAI5t5yGPC18zF31HzHsQKG&Signature=UVkjjLVT4Ac2wvKDEzQf%2FRZnZrU%3D&response-content-disposition=attachment%3Bfilename%3D1_9_IMG_20251111_142021.jpg	(空)	(空)	https://alifile.sojump.cn/338076422_1_q12_20251111142043719F9TOUD.jpg?Expires=1770618115&OSSAccessKeyId=LTAI5t5yGPC18zF31HzHsQKG&Signature=kTfN%2Bfyw9OqkUwSaMXORBeNXyyE%3D&response-content-disposition=attachment%3Bfilename%3D1_12_IMG_20251110_205859.jpg	(空)	2	鞋子总体还是比较舒适的，可跑性也还不错，就是鞋垫的存在感太强，鞋垫和中底脚感比较割裂。
##    1	2025/11/11 14:21:33	345秒	微信	N/A	39.144.251.51(福建-厦门)	64	隋晗2	150	3	2	9	3	8	7	4	6	6	7	6	8	1	0	0	0	1	0	0	0	0	0	0	初上脚有非常软的脚感，泡棉形变大，缓震极佳，总体舒适性好，在有氧跑的情景下，泄力感不强	跑步5-8km之后，前掌衰减明显，有变硬的趋势，且坡差过大，包裹和锁定偏差，用力调整鞋带后尚可接受	https://alifile.sojump.cn/338076422_1_q9_20251111142033679JJHG47.jpg?Expires=1770618115&OSSAccessKeyId=LTAI5t5yGPC18zF31HzHsQKG&Signature=UVkjjLVT4Ac2wvKDEzQf%2FRZnZrU%3D&response-content-disposition=attachment%3Bfilename%3D1_9_IMG_20251111_142021.jpg	(空)	(空)	https://alifile.sojump.cn/338076422_1_q12_20251111142043719F9TOUD.jpg?Expires=1770618115&OSSAccessKeyId=LTAI5t5yGPC18zF31HzHsQKG&Signature=kTfN%2Bfyw9OqkUwSaMXORBeNXyyE%3D&response-content-disposition=attachment%3Bfilename%3D1_12_IMG_20251110_205859.jpg	(空)	2	鞋子总体还是比较舒适的，可跑性也还不错，就是鞋垫的存在感太强，鞋垫和中底脚感比较割裂。
##    '''
##
##    f"你给我的回答必须是下面这种格式："
##
##    f"(('1、姓名（昵称）', '5、回弹性/推进性', '5、稳定性/支撑性', '5、抓地力/牵引力', '5、长距离抗衰减性'), ('隋晗', 3.0, 8.0, 7.0, 4.0), ('陈晶', 6.0, 6.0, 7.0, 7.0))"
##
##    f"现在，我把原始数据传给你，开始你的任务：\r{raw_data}"
##
##    #f"注意，[步长/重心振幅比值] 是我非常关注的指标，如果基于我提供的数据你能计算这个指标，那么你千万不要忘记计算它；注意你只需给出计算结果和结论，你不需要展示计算过程。\r "     # 这里可以多备注一些指标，让GPT尽可能计算。
##
##    )
##
##    return mc_prompt

###Debug
##mc_cell = get_range(mc_sht)
##mc_prompt = gen_mc_prompt_question_data(mc_cell)
##reply = GPT-5(mc_prompt,model = 'gpt-5-mini')


# ===========================================================================
# GPT 等待遮罩（yzr_ppt / zxh_ppt 工作期间显示，完成后删除）
# ===========================================================================

def show_gpt_waiting_overlay(mc_slide):
    """在当前 Slide 顶层插入「GPT 服务器通讯中」遮罩矩形，返回 Shape 供后续删除。

    属性来自实测 PPT 手工调整结果：
      黑色填充 / 透明度 0.17 / 无边框 / 微软雅黑 40pt 白色粗体 / 垂直居中
    注意：先写文字再设尺寸，防止 TextFrame 自动调整触发意外缩放。
    """
    shp = mc_slide.Shapes.AddShape(1, 0, 213.3, 960, 141.7)

    # 填充
    shp.Fill.ForeColor.RGB = 0          # 黑色
    shp.Fill.Transparency = 0.17        # 实测透明度（约 83% 不透明）

    # 无边框
    shp.Line.Visible = 0

    # ① 先写文字 + 字体（文字可能触发 TextFrame 自动调整大小）
    tf = shp.TextFrame
    tf.VerticalAnchor = 3               # 垂直居中
    tr = tf.TextRange
    tr.Text = 'GPT 服务器通讯中，请耐心等待'
    tr.Font.NameFarEast = '微软雅黑'
    tr.Font.Name = '微软雅黑'
    tr.Font.Size = 40
    tr.Font.Bold = -1                   # 加粗
    tr.Font.Color.RGB = 16777215        # 白色
    tr.ParagraphFormat.Alignment = 1    # 左对齐

    # ② 再锁定尺寸（覆盖自动调整结果）
    shp.Left = 0
    shp.Top = 213.3
    shp.Width = 960
    shp.Height = 141.7

    # 置顶（msoBringToFront = 0）
    shp.ZOrder(0)

    return shp


def remove_gpt_waiting_overlay(shp):
    """删除「GPT 服务器通讯中」遮罩文本框。"""
    try:
        shp.Delete()
    except Exception:
        pass
