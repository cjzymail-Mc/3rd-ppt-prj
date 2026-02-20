


--------------------
ppt 标准工作流程总结
--------------------



我有一个现成的脚本，代码为【根目录main.py + /src】，我的脚本解析结果在repo-sanc-result.md中。

你现在是一个极度严谨的 PowerPoint COM 自动化专家 + Office自动化专家，
你精通xlwings 和 pywin32 对应的office Application COM API（PowerPoint / Excel 的 COM API）
在控制Excel时，你必须使用 xlwings + COM API 二者结合的方式，实现精准控制；
在控制PowePpoint时，你必须使用 pywin32 + win32com.client 通过真正的 PowerPoint.Application COM 接口来操作，
注意，**严禁使用 python-pptx** ！！！


当前repo中的根目录包含几个重要文件： 
【0-空白 ppt 模板.pptx】：仅保留基本的框架和装饰元素 
【1-标准 ppt 模板.pptx】：最终要实现的ppt效果（内容、排版都很完美） 
【2025 数据 v2.2.xlsx】文件中的【问卷sheet】：唯一的数据源，用来生成标准ppt模板中的文字、图表等所有内容
【Template 2.1.pptx】 已包含 【空白 ppt 模板】和【标准 ppt 模板】，分别位于【Template 2.1.pptx】文件 的第14页 和 15页


现在需要在原来的基础上开发新的功能。


------第1步----- 
首先，你需要写一个测试文件（01-shape-detail.py），分别读取【Template 2.1.pptx】中的（14页）【空白 ppt 模板】和（15页）【标准 ppt 模板】
 注意，你必须遍历读取所有元素shape，并识别出【标准 ppt 模板】中新增的元素（哪些元素是标准模板中新增的、空白模板中不存在的元素？）
  识别出新增的元素后，详细读取这些元素的所有信息，包括：

{
页面编号（slide_index），名称（name，如果名称为空，你根据它的特征，简短给它生成一个名字，并记住它）、
位置（top & left）、 大小（height & width）、 
字体（并检查该字体class文件是否存在，如果不存在就需要更新class_030.py文件，遵循class文件中的惯例新增）、 
属性（shape.Type，文本/图片/chart？）、 
内容（文字内容）
是否在组里（Parent.Group） }

 tips: tool 1 - shape top&left v1.1.py是我之前使用的一个小工具，你可以参考它，但我本地拥有 office 环境方便调用，你可能需要自己想办法
然后，将这些新增元素的信息，汇总成01-shape-detail.md（假设你识别出一共有9个新增元素），详细列出上面所有信息，并提供给我






------第2步----- 
分析 shapes 内容：
如果你成功识别出了这些【新增元素】，这些新增元素的内容都是源于【2025 数据 v2.2.xlsx】文件中的【问卷sheet】里的内容 
你需要写一个（02-shape-analysis.py）文件，通过它来逐一对比每个shape【新增元素】与【源数据】之间的关系
    （标准ppt模板里新增的9个shape的内容，完全源于 问卷sheet）
并输出一份标准文档：02-shape-analysis.md ，以便于其他agents根据这个标准文档，构建对应的函数工具，实现内容的精确提取和重构
注意事项：
    我之前的脚本代码【根目录main.py + /src】，里面的所有工具你都可以调用，不需要重新造轮子
    例如，你可以动态加载 src/Function_030.py 里的 GPT_5 与 extract_info，用于 AI 总结与复杂字段提取（并带兜底逻辑）






------第3步----- 
构建 shapes 内容：
    根据02-shape-analysis.md的要求，并结合我之前的脚本代码【根目录main.py + /src】里面的所有工具
    你帮我写一个py测试脚本（03-build_shape.py），主要包含9个函数（假设你识别出一共有9个新增shape元素）
    这9个函数的主要功能，是从【源数据-问卷sheet】中提取对应的信息，加工成对应9个shape的精确内容
    
    再次提醒，我原来脚本中包含【GPT-5通讯函数】，可以用来连接gpt服务器，实现问卷信息的总结和提取，你仅需构建对应的 prompt 即可
    你可以在新的脚本中（02-build_ppt_com.py）重复利用（import）这些现成的函数，不需要重写。





构建 ppt文件：
    另外，完成9个shape内容的构建后，
    你的这个（03-build_ppt_com.py）这个脚本，
    还需要完全按照【标准 ppt 模板】的样式，重新生成1个ppt文件，命名中包含【你的名字 + 版本号】
    （假设你是codex，那命名就叫做【codex 1.0.pptx】）
    你必须严格按照我的工作流程，来完成你的ppt制作（注意： 严禁使用 python-pptx ！！）：

    1. 启动 PowerPoint.Application（Visible=False）
    2. 创建一个新 Presentation 文件（假设你是codex，那命名就叫做【codex 1.0.pptx】）
    3. 复制【标准ppt模板（Template 2.1.pptx 第15页）】，source_slide.Duplicate() 或 Copy + Paste 的方式完整克隆 
        并保留源格式（这是关键：直接基于完美模板）（保留标准ppt模板 所有美工、动画、母版效果）
    4. shape 元素操作
    - 克隆后，根据shape_detail_com.json 精准找到9个目标形状（假设你识别出一共9个shape需要更新）
    - **只做内容替换**：
            - 文本形状 → 更新 .TextFrame.TextRange.Text（保留原有字体、颜色、段落格式）
            - 图表形状 → 读取【2-问卷模板.xlsx】对应数据，并通过 Chart.ChartData 或直接修改ppt-chart的 Series.Values / Categories 来更新数据（**保留图表所有样式、颜色、数据标签、轴格式**）
            - 数据仅需仅需简单的统计：例如均值、频率、汇总等，无需引入numpy等大型计算工具
            - 图片（如有）→ 如果需要替换则使用 .PictureFormat 或 Replace
    5. 所有数据必须来自 Excel（建议使用 xlwings 打开 Excel 读取）




------第4步----- 
你完成ppt文件生成后，需要再写一个测试脚本（04-shape_diff_test.py），
使用 COM 打开【标准模板（位于Template 2.1.pptx第15页）】和【codex.pptx（第1页，你生成的ppt应该只有1页）】，
逐 shape 对比【位置、大小、字体、内容（例如文字数量是否极为接近）颜色、图表类型等所有关键属性，】
给出详细差异报告。如果差异明显，必须自动调整代码并重新生成。

**最高优先级**：
- 视觉保真度必须达到 98%+（几乎和手动复制 slide 后改数据一样）
- 严禁使用 python-pptx 的任何函数
- 代码必须健壮、带错误处理和资源清理
- shape_detail_com.json 精确定位需要更新的shapes
- 生成fix-ppt.md文档，详细标注差异点，以及修复建议（修改哪个函数？如何修改？）



------第5步----- 
假设上一步的测试结果不满意，需要重复上面的【调整ppt - 测试ppt】步骤，
生成新的带编号ppt文件【codex 1.1/1.2/1.3.pptx 】，不要覆盖掉之前的ppt文件
直至你生成的最新ppt文件能够通过（04-shape_diff_test.py）的验证（视觉保真度必须达到 98%+（几乎和手动复制 slide 后改数据一样）
当然，我也会设置一个最大重复限制（例如最大重复3轮）



------第5步----- 
我作为用户，会人工检查最终生成的ppt文件，肉眼对比最新版【codex.pptx】和【标准模板】的差异
根据人工对比效果，再考虑是否继续下一轮的优化（直接进入【调整ppt - 测试ppt】步骤）













用户（我）：
我已经尽可能将制作ppt的过程拆解成详细步骤，为什么效果似乎还是不理想？  
你认为哪几个环节可以需要继续拆分成更细的任务包？   



---------------
codex对这个问题的答复是：
你这个问题非常关键：不是“流程有没有”，而是“流程有没有足够细的质量控制点”。

我已经按你的最新版目标，给出拆分建议并落到新文档 new-ppt-workflow.md 了。

我对“为什么效果不理想”的结论（简版）
最核心是 4 个缺口：

shape 匹配策略不够稳（仅 name/顺序容易漂移）。

内容构建缺少“可读性预算”（字数/行数/bullet 没被硬约束）。

diff 缺少语义与可读性门槛（只看样式不够）。

迭代没有“自动修复路由”，导致收敛慢。

我建议继续拆细的任务包（已写进文档）
重点新增拆分：

Step1：指纹构建 / 组内解析 / 模板漂移门禁。

Step2：字段映射 / Prompt规格 / 可读性预算。

Step3：9函数构建 / 内容校验 / COM写入 / 写后回读。

Step4：视觉diff / 可读性diff / 语义diff 三层门禁。

Step5：修复路由 / 收敛控制。

新流程产物
我还把新版标准流程整理成 Phase 0~4（20个任务），并补了建议新增产物与验收门槛（Visual≥98%、Readability≥95%、Semantic=100%）。