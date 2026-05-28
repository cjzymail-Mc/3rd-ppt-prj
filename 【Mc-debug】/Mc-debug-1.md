我有一个现成的脚本，现在需要开发新的功能：

功能需求描述：
将Excel【问卷模板-1】的内容读取、解析，然后生成PPT【篮球试穿问卷PPT】。
我可能有数十份问卷需要处理，因此需要一个智能处理函数，来完成数据解析、图表制作、排版、生成ppt。
你需要帮我制定一个开发计划，来完成该功能的开发。
我的repo解析结果在 repo-scan-result.md中。

注意事项：
注意，你的任务不是写代码，而是仔细阅读role.md的文档内容，按md文件的要求，制定完善的新功能计划





======================

gemini 和 gpt 都失败了， 但 codex 2分钟完成了任务！！！

专业业务能力出众！！！！


如果你以后会大量使用 Claude Agent，建议把 .claude/agents/、.claude/commands/、.claude/rules/（或 memory 里的规则文件）好好维护起来，这是目前最有效的「持久化 agent 知识」方式。
.claude/hooks/ 里如果有更多自定义脚本，也可以继续提交，但日志类的一定要 ignore。




-----------任务1：制定plan.md----------------

你仔细阅读 your-role.md 文档，然后根据role文档要求完成我的任务：

任务描述：
我有一个现成的脚本，代码为【根目录main.py + /src】，脚本解析结果在repo-sanc-result.md中。
现在需要在原来的基础上开发新的功能：

我的功能需求描述：
读取和解析Excel【问卷模板 - 1.xlsx】的内容，然后利用现有模板【ppt模板 - 1.pptx】，生成PPT【篮球试穿问卷ppt - 1.pptx】。
我可能有数十份问卷需要处理，需要生成数十份类似的ppt，因此需要一个智能处理函数，来完成数据解析、图表制作、排版、生成ppt。
你需要帮我制定一个开发计划，来完成该功能的开发。

注意事项：
再次强调，你的任务不是写代码，而是仔细阅读your-role.md的文档内容，按md文件的要求，制定完善的新功能开发计划


-------------------

我看了你的plan，有几点需要更新：
1. excel文件中的数据仅需简单处理（均值、频次统计等），不需要引入 numpy等大型计算插件。
2. 需要进行评论总结的地方，你可以直接调用我现有的GPT_5函数，你只需构建对应的提示词
3. 最大程度利用现有 ppt模板，避免复杂线条的绘制。你可能需要重点聚焦在图表绘制和排版功能
4. 我希望直接在 main 文件中新增这个功能，插入到（【5】实战测评部分 ）和（【6】结论部分）之间，插入位置大概位于main.py的 790行-800行之间。
5. 你不确定的地方，你可以向我提出问题（例如给我提供3个选项或者直接问我）




-------------------

好的，继续深度整合。将你的新增功能，整合到function中的 questionnaire_Excel函数里。

你可以保留你新增的这些build_questionnaire_summary_prompt等函数，在questionnaire_Excel中嵌套调用它即可。

这样我在main中仅需调用一次questionnaire_Excel函数，就能实现【问卷逐页处理（原有功能）  +  问卷汇总处理（你新增的功能）】

-------------------
接下来，继续深度整合

我将【ppt模板 - 1.pptx（只有1页ppt）】插入到了【Template 2.1.pptx】中，位于Template 2.1.pptx的第14页

我无法上传文件更新给你，但你可以按我更新后的文件执行（假设你已经可以从【Template 2.1.pptx】第14页获取到【ppt模板 - 1.pptx】完全一样的内容）

这样，你的新代码中可以继续沿用我之前的 【Template 2.1.pptx】，而不是调用2个ppt了。

你继续修改代码，实现仅从1个ppt文件读取，就能完成所有工作。




-------------------

























----------- 换个思路：手把手一点点教，才能做到最精确 ----------------


--------更新: ppt 工作流程--------
























-----任务2：试试 gemini jules-----        结果发现不如 codex












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














------最后的步骤：整合（待ppt效果ok之后再说，不急着整合）----- 
接下来你需要继续完成深度整合，将上面的py代码整合进入原来的【根目录main.py + /src】脚本程序中。注意：
接下来，我希望直接在 main 文件中你上面的功能，
插入到（【5】实战测评部分 ）和（【6】结论部分）之间，插入位置大概位于main.py的 785行-800行之间。




### 第2步：生成 ppt 主脚本（最高保真方案）
基于分析结果，编写主脚本 `build_codex_ppt_com.py`，严格遵循以下流程：

1. 启动 PowerPoint.Application（Visible=False）
2. 打开【1-标准 ppt 模板.pptx】作为源模板（这是关键：直接基于完美模板）
3. 创建一个新 Presentation（或直接 SaveAs 为 gemini-jules.pptx）
4. 对于需要的每一页：
   - 使用 source_slide.Duplicate() 或 Copy + Paste 的方式完整克隆 【1-标准 ppt 模板.pptx】对应的 slide（保留所有美工、动画、母版效果）
   - 克隆后，根据shape_detail_com.json 精准找到9个目标形状（假设你识别出一共9个shape需要更新）
   - **只做内容替换**：
        - 文本形状 → 更新 .TextFrame.TextRange.Text（保留原有字体、颜色、段落格式）
        - 图表形状 → 读取【2-问卷模板.xlsx】对应数据，并通过 Chart.ChartData 或直接修改ppt-chart的 Series.Values / Categories 来更新数据（**保留图表所有样式、颜色、数据标签、轴格式**）
        - 数据仅需仅需简单的统计：例如均值、频率、汇总等，无需引入numpy等大型计算工具
        - 图片（如有）→ 如果需要替换则使用 .PictureFormat 或 Replace
5. 所有数据必须来自 Excel（建议使用 xlwings 打开 Excel 读取）
6. 最后 SaveAs "gemini-jules.pptx" 并 Quit PowerPoint





























我的工作要求ppt的制作精度非常高，因此你必须将ppt当成一个架构精密的软件系统来对待
我已构建了软件工程通用的6个agents，位于 .claude/agents 目录中
我还构建了专业的agents控制器，位于根目录的 orchestrator.py

接下来我希望将它们改造成专业ppt工作团队，按照软件工程的逻辑，来生成ppt
即，用严谨的开发流程来生成ppt【规划、构建代码、测试、反馈、修改...（重复多轮后）→ 交付】

我原有的ppt生成脚本由【根目录main.py + /src】脚本程序构成，我的脚本解析结果在repo-sanc-result.md中
我原有的脚本虽然效率不高，但能够实现ppt的精准控制，你必须严格遵循我的原有思路，来进行接下来的工作
我接下来希望的工作流程为 ppt-workflow.md 文档记录
你根据这个 ppt-workflow.md 文档的工作流程 
帮我改写 【6个agents 的md文件】  和 【orchestrator.py】
将它们打造成专业的 ppt 软件工程师团队，来满足我对ppt的严苛要求





Traceback (most recent call last): File "D:\Technique Support\Claude Code Learning\3rd-ppt-prj\01-shape-detail.py", line 91, in <module> raise SystemExit(main()) File "D:\Technique Support\Claude Code Learning\3rd-ppt-prj\01-shape-detail.py", line 72, in main "in_group": bool(getattr(s, "ParentGroup", None)), File "C:\Users\xy24\AppData\Local\Programs\Python\Python312\Lib\site-packages\win32com\client\dynamic.py", line 627, in __getattr__ ret = self._oleobj_.Invoke(retEntry.dispid, 0, invoke_type, 1) pywintypes.com_error: (-2147352567, '发生意外。', (0, None, '指定的值超出了范围。', None, 0, -2147024891), None)           这个报错的直接原因是：代码里这一行在读 Shape.ParentGroup 的时候，PowerPoint COM 端抛异常了，而不是像 Python 属性那样“没有就返回 None”。     
你的脚本里是这么写的："in_group": bool(getattr(s, "ParentGroup", None)), 

01-shape-detail

为什么 getattr(..., None) 也会炸？

因为 win32com.client 的对象是动态派发：
getattr(s, "ParentGroup", None) 并不是“找不到属性就用 None”，而是会去调用 COM 的 Invoke 获取该属性值。
而在 PowerPoint 里，很多 shape（不在组里、某些特殊类型、占位符/图表/SmartArt/某些对象等）访问 ParentGroup 时会直接抛出 COM 异常（常见就是 DISP_E_EXCEPTION），于是你就看到：

pywintypes.com_error: (-2147352567, ... '指定的值超出了范围。' ...)

这是 PowerPoint 返回的异常，不是 Python 的 AttributeError。

常见触发场景

shape 根本不在 group 里（但 ParentGroup getter 不是返回空，而是抛错）

shape 类型比较特殊（占位符、图表、某些 OLE/ActiveX、SmartArt 等），ParentGroup 不适用

文档处于受保护/只读/受限状态时，一些属性访问也可能抛错（你这里更像“不适用属性”）    

 你要吸取教训，如果你提供的01-shape-detail/02-build_ppt_com/03-shape_diff_test 三个文件存在同样的错误，你要一次性都修正过来




很好，3个py程序不再报错了。不过，我按顺序运行了你提供的01-shape-detail/02-build_ppt_com/03-shape_diff_test 三个文件

01-shape-detail.py 成功识别了 新增的9个shape，并成功生成了对应的文件，ok


但02-build_ppt_com.py 问题很严重：
它仅仅完成了对标准模板的复制，没有实现从【源数据-Excel-问卷sheet】 生成【ppt】的功能。
你需要完全重写02-build_ppt_com.py ，来实现我所需要的功能。

你可以调用我之前的脚本中【function.py 】里的 GPT_5 函数，
来实现ai的调用，协助你从 Excel中提取信息并加工为 shape所需的内容（你可能需要在02-build_ppt_com.py中构建多个对应的prompt，来加工信息）

你还可以重复利用 【function.py 】中的 extract_info() 函数，提取复杂信息。如需更多功能，你可以参考extract_info() 函数，构建新函数


你再仔细阅读下我的需求，然后完成02-build_ppt_com.py的代码改写。

我期望的目标如下：

    （再重复粘贴一次上面的 md 文件）





02-build_ppt_com.py 问题很严重：
它生成的shape内容，和标准模板差距太大了，根本没办法阅读，毫无美观可言

而运行 03-shape_diff_test.py 后，生成的 Fix Report竟然显示通过？！
根本不是同一个东西，2个ppt差距太大了，你的这个检测标准也未免太低了点。

# PPT Diff & Fix Report

- 状态: ok
- template_shapes: 27
- target_shapes: 27
- 时间: 2026-02-12T14:26:18

> 详细字段级diff可继续扩展。



你再仔细阅读下我的需求，然后完成02-build_ppt_com.py / 03-shape_diff_test.py的代码改写。

我期望的目标如下：

    （再重复粘贴一次上面的 md 文件）





我按你最新要求把流程拆成了“分析 → 构建内容 → 生成PPT → 严格Diff测试”的分层实现，并保留旧入口兼容调用，避免你已有执行命令失效：02-build_ppt_com.py 现在会串行调用 02-shape-analysis.py、03-build_shape.py、03-build_ppt_com.py；03-shape_diff_test.py 转发到新的严格测试脚本 04-shape_diff_test.py。.

新增 02-shape-analysis.py：读取 shape_detail_com.json + Excel（xlwings + COM）并结合模板页文本，输出 shape 与数据源关系映射（shape_analysis_map.json），用于后续函数级构建标准化。.

新增 03-build_shape.py：实现 9 个 shape 内容函数，且复用你旧代码中的 GPT_5、extract_info，将内容统一输出为 build_shape_content.json，避免“拍脑袋文本”直接写入PPT。.

新增 03-build_ppt_com.py：严格保留“克隆标准模板第15页 + 只替换内容”的策略，文本走 TextFrame.TextRange.Text，图表走 ChartData 更新，尽量保持模板样式与视觉资产。.

新增 04-shape_diff_test.py 并提升检测标准：不再只比 shape 数量，而是逐 shape 计算位置/尺寸/类型/文本相似度/字体/颜色/图表类型加权分；报告只有在 overall >= 98% 且无低分shape时才 ok，否则 fail 并输出修复建议。.







所以我现在应该按哪个流程调用？你简单解释下，然后再给我写一个00-ppt.py，实现一键调用上面的所有流程


你现在用这个流程就行（简版）：

    python 00-ppt.py --version 1.0

自动跑：01 提取shape → 02 分析映射 → 03 构建内容 → 03 生成ppt → 04 严格diff。.

如果你只想重跑某一段（比如只重测diff）：

    python 00-ppt.py --version 1.0 --from-step 4 --to-step 4。.

如果目标文件不是默认 codex 1.0.pptx：

    python 00-ppt.py --target "codex 1.2.pptx" --from-step 4 --to-step 4。.






接下来讨论我最新的工作流程（最近一次发给你的【我期望的目标】那一大段文字，注意是最新版的）：
我已经尽可能将制作ppt的过程拆解成详细步骤，为什么效果似乎还是不理想？  
你认为哪几个环节可以需要继续拆分成更细的任务包？   
你给出你的调整建议，然后帮我更新成一份 new-ppt-workflow.md



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






非常好，你按你提供的最新 new-ppt-workflow.md 和 最新拆分计划，
重新调整所有代码（可以删除旧文件，提供全新的一批 py 脚本文件给我）。
我期望仍然拥有统一的控制入口，到时候我直接运行第一个文件（假设是00-ppt.py），
就能实现你上面的全新工作流。至于过程控制需要拆分成几个文件，你自己决定（我认为可以适当简化）。  

另外，
我之前的脚本代码【根目录main.py + /src】，里面的所有工具你都可以调用，不需要重新造轮子
例如，你可以动态加载 src/Function_030.py 里的 GPT_5 与 extract_info 等一些列函数，
用于 AI 总结与复杂字段提取（并带兜底逻辑）。
你也可以参考我的 function.py 文件中的 prompt，来精确控制 GPT_5 输出的文字内容，来匹配shape 内容的生成需要
你也可以重构新的extract_info函数，来提取更多关键信息，用于 shape内容的生成、问卷评分的提取、均值计算、频次统计等
当shape内容对比测试不通过时，
    你可能要优先调整 相关提示词（如果内容是由 GPT_5 函数生成），也可以考虑修改信息提取函数（如果内容是直接从问卷提取）









我测试了你的新代码，效果并没有太大改善。
几个主要问题：
1、需要生成的shape中，有一个chart（条形图），它的数据源已丢失，你可能需要重新从【源数据-问卷sheet】中获取对应的数据
    这个数据是每项评分的均值（例如有3个人填写了问卷，它是每项指标3个人填写评分的均值）
    你需要自己写一个数据提取函数（类似extract_info函数），来提取对应的数据

2、我已经强调过很多次了，你也可以参考我的 function.py 文件中的 prompt，
    来精确控制 GPT_5 输出的文字内容，来匹配shape 内容的生成需要
    但我在你的 所有py文件中，没有看见任何的prompt信息。你不利用GPT_5函数，怎么可能提取到准确的评论总结信息？
    你必须严格参考我在 Function_030.py 中的 GPT_5函数 和 gen_mc_prompt() gen_result_prompt()等函数的使用方式，你自己要构建对应的prompt，才能获取到准确的信息



build_prompt_with_legacy() 会优先调用 gen_result_prompt/gen_questionnaire_prompt/gen_mc_prompt；   
你是理解能力有问题。  我需要你根据【标准模板】 和 【源数据】之间的关联， 自己构建prompt。你沿用我的旧prompt有啥用？  
你必须思考， 你应该用什么样的prompt，才能从【源数据】提取出类似【标准模板】中那样的信息呢？






很好，你思路对了
但我再纠正你一点：不同shape需要用不同的方式来生成，而不是统一调用 GPT_5

假设你识别出9个shape需要更新；

针对每个shape，你都要思考：
    这个shape是什么东西？（标题？评论总结？图表评分？等级评价？图片？）
    这个shape的内容是【源数据】的哪个地方的信息加工出来的？
    假如这个shape是文本，那它的文本长度（文字数量）是多少？如何精确控制生成的长度与标准模板相同？
    你需要用什么方式才能提取出构建这个shape所需的内容？（调用GPT_5？还是直接写一个re正则处理函数提取？还是重新写一个类似extract_info函数？）
    如果这个shape的信息，在【源数据】中完全找不到，你需要记录下来，并输出给用户(输出一个 md文件让我确认，而不仅仅只是确认 prompt信息)
    

















--------------- grok ----------------

你现在是一个极度严谨的 PowerPoint COM 自动化专家，必须使用 pywin32 + win32com.client 通过真正的 PowerPoint.Application COM 接口来操作，**严禁使用 python-pptx**。
我有一个现成的脚本，代码为【根目录main.py + /src】，脚本解析结果在repo-sanc-result.md中，你仔细阅读该文件。
我的 branch codex3 中有 3 个核心文件：
- 【0-空白 ppt 模板.pptx】：基础框架
- 【1-标准 ppt 模板.pptx】：**视觉完美基准**（排版、美工、字体、颜色、图表样式、动画全部完美）
- 【2-问卷模板.xlsx】：数据源

**核心目标**：生成 `gemini-jules.pptx`，视觉效果必须极度接近【1-标准 ppt 模板】，肉眼几乎无法分辨差异。

**强制要求**：
- 全部使用 win32com.client.Dispatch("PowerPoint.Application")
- 必须后台运行（Visible = False）
- 操作完成后必须正确 Quit 并释放 COM 对象
- 所有关键 shape 必须通过 .Name 属性精准定位（请在代码中说明：用户需提前在【1-标准 ppt 模板】中给需要替换的形状手动命名，例如 "Title"、"MainChart"、"Q1_Bar"等）

### 第1步：深度分析（使用 COM）
请先写一个分析脚本 `analyze_templates_com.py`，使用 COM 同时打开【0-空白 ppt 模板.pptx】和【1-标准 ppt 模板.pptx】，完成以下任务：

1. 遍历每一页 slide
2. 遍历读取所有元素shape，并识别出【标准 ppt 模板】中新增的元素（哪些元素是标准模板中新增的、空白模板中不存在的元素？）
    识别出新增的元素后，详细读取这些元素的所有信息，包括下列主要内容：
    {
    位置（top & left）、 大小（height & width）、 
    字体（并检查该字体class文件是否存在，如果不存在就需要更新class_030.py文件，遵循class文件中的惯例新增）、 
    属性（文本/图片/chart？）、 
    内容（文字内容） }

2. 对于每个 Shape（包括分组内的形状），提取并输出到 `shape_detail_com.json`：
   - slide_index
   - shape.Name（如果没有名称，你需要自己进行编号、管理）
   - shape.Left, .Top, .Width, .Height
   - shape.Type（msoTextBox, msoChart, msoPicture, msoGroup 等）
   - 如果是文本：TextRange.Text、Font.Name、Font.Size、Font.Bold、Font.Color.RGB、ParagraphFormat.Alignment、LineSpacing 等
   - 如果是图表：Chart.ChartType、是否有 DataTable、Series 数量、样式关键信息
   - 是否在组里（Parent.Group）

输出shape_detail_com.json ， JSON 结构清晰，该json文件包含所有需要更新的shapes（假设你识别出一共9个shape需要更新）


### 第2步：生成主脚本（最高保真方案）
基于分析结果，编写主脚本 `build_codex_ppt_com.py`，严格遵循以下流程：

1. 启动 PowerPoint.Application（Visible=False）
2. 打开【1-标准 ppt 模板.pptx】作为源模板（这是关键：直接基于完美模板）
3. 创建一个新 Presentation（或直接 SaveAs 为 gemini-jules.pptx）
4. 对于需要的每一页：
   - 使用 source_slide.Duplicate() 或 Copy + Paste 的方式完整克隆 【1-标准 ppt 模板.pptx】对应的 slide（保留所有美工、动画、母版效果）
   - 克隆后，根据shape_detail_com.json 精准找到9个目标形状（假设你识别出一共9个shape需要更新）
   - **只做内容替换**：
        - 文本形状 → 更新 .TextFrame.TextRange.Text（保留原有字体、颜色、段落格式）
        - 图表形状 → 读取【2-问卷模板.xlsx】对应数据，并通过 Chart.ChartData 或直接修改ppt-chart的 Series.Values / Categories 来更新数据（**保留图表所有样式、颜色、数据标签、轴格式**）
        - 数据仅需仅需简单的统计：例如均值、频率、汇总等，无需引入numpy等大型计算工具
        - 图片（如有）→ 如果需要替换则使用 .PictureFormat 或 Replace
5. 所有数据必须来自 Excel（建议使用 xlwings 打开 Excel 读取）
6. 最后 SaveAs "gemini-jules.pptx" 并 Quit PowerPoint

请在脚本中包含以下关键函数：
- `open_powerpoint()`
- `duplicate_slide(source_ppt, target_ppt, slide_index)`
- `update_text_shape(slide, shape_name, new_text)`
- `update_chart_data(slide, shape_name, excel_data_dict)`  ← 重点实现图表数据更新
- `build_final_ppt()`

### 第3步：验证
生成完成后，写一个对比验证脚本，使用 COM 打开【1-标准 ppt 模板】和【gemini-jules.pptx】，逐 shape 对比位置、大小、字体、颜色、图表类型等关键属性，给出详细差异报告。如果差异明显，必须自动调整代码并重新生成。

**最高优先级**：
- 视觉保真度必须达到 98%+（几乎和手动复制 slide 后改数据一样）
- 严禁使用 python-pptx 的任何函数
- 代码必须健壮、带错误处理和资源清理
- shape_detail_com.json 精确定位需要更新的shapes



























































接下来安排第二轮任务：

在上一轮任务中，你让我在 1-标准 ppt 模板.pptx 里把后续要改的关键形状先命名好（如 Title、MainChart 等），
因为后面脚本是按 .Name 精准定位的。  
我希望你改为自动命名（或者自动重命名）。
因为我有大量ppt需要处理，我不可能一个个shape去手工命名。

文件更新：
1、Template 2.1.pptx  已包含 【空白 ppt 模板】和【标准 ppt 模板】，分别位于 Template ppt文件 的第14页 和 15页
2、【2-问卷模板.xlsx】中的数据已挪至【2025 数据 v2.2.xlsx】文件中

接下来，我最终目的是将你上一轮工作中完成的功能（问卷信息汇总成1页ppt的功能）融入原来的【根目录main.py + /src】脚本程序，要求如下：

我希望直接在 main 文件中新增这个功能，插入到（【5】实战测评部分 ）和（【6】结论部分）之间，插入位置大概位于main.py的 785行-800行之间。





第一轮任务回顾：（你已完成）


'''
你现在是一个极度严谨的 PowerPoint COM 自动化专家，必须使用 pywin32 + win32com.client 通过真正的 PowerPoint.Application COM 接口来操作，**严禁使用 python-pptx**。
我有一个现成的脚本，代码为【根目录main.py + /src】，脚本解析结果在repo-sanc-result.md中，你仔细阅读该文件。
我的 branch codex3 中有 3 个核心文件：
- 【0-空白 ppt 模板.pptx】：基础框架
- 【1-标准 ppt 模板.pptx】：**视觉完美基准**（排版、美工、字体、颜色、图表样式、动画全部完美）
- 【2-问卷模板.xlsx】：数据源

**核心目标**：生成 `gemini-jules.pptx`，视觉效果必须极度接近【1-标准 ppt 模板】，肉眼几乎无法分辨差异。

**强制要求**：
- 全部使用 win32com.client.Dispatch("PowerPoint.Application")
- 必须后台运行（Visible = False）
- 操作完成后必须正确 Quit 并释放 COM 对象
- 所有关键 shape 必须通过 .Name 属性精准定位（请在代码中说明：用户需提前在【1-标准 ppt 模板】中给需要替换的形状手动命名，例如 "Title"、"MainChart"、"Q1_Bar"等）

### 第1步：深度分析（使用 COM）
请先写一个分析脚本 `analyze_templates_com.py`，使用 COM 同时打开【0-空白 ppt 模板.pptx】和【1-标准 ppt 模板.pptx】，完成以下任务：

1. 遍历每一页 slide
2. 遍历读取所有元素shape，并识别出【标准 ppt 模板】中新增的元素（哪些元素是标准模板中新增的、空白模板中不存在的元素？）
    识别出新增的元素后，详细读取这些元素的所有信息，包括下列主要内容：
    {
    位置（top & left）、 大小（height & width）、 
    字体（并检查该字体class文件是否存在，如果不存在就需要更新class_030.py文件，遵循class文件中的惯例新增）、 
    属性（文本/图片/chart？）、 
    内容（文字内容） }

2. 对于每个 Shape（包括分组内的形状），提取并输出到 `shape_detail_com.json`：
   - slide_index
   - shape.Name（如果没有名称，你需要自己进行编号、管理）
   - shape.Left, .Top, .Width, .Height
   - shape.Type（msoTextBox, msoChart, msoPicture, msoGroup 等）
   - 如果是文本：TextRange.Text、Font.Name、Font.Size、Font.Bold、Font.Color.RGB、ParagraphFormat.Alignment、LineSpacing 等
   - 如果是图表：Chart.ChartType、是否有 DataTable、Series 数量、样式关键信息
   - 是否在组里（Parent.Group）

输出shape_detail_com.json ， JSON 结构清晰，该json文件包含所有需要更新的shapes（假设你识别出一共9个shape需要更新）


### 第2步：生成主脚本（最高保真方案）
基于分析结果，编写主脚本 `build_codex_ppt_com.py`，严格遵循以下流程：

1. 启动 PowerPoint.Application（Visible=False）
2. 打开【1-标准 ppt 模板.pptx】作为源模板（这是关键：直接基于完美模板）
3. 创建一个新 Presentation（或直接 SaveAs 为 gemini-jules.pptx）
4. 对于需要的每一页：
   - 使用 source_slide.Duplicate() 或 Copy + Paste 的方式完整克隆 【1-标准 ppt 模板.pptx】对应的 slide（保留所有美工、动画、母版效果）
   - 克隆后，根据shape_detail_com.json 精准找到9个目标形状（假设你识别出一共9个shape需要更新）
   - **只做内容替换**：
        - 文本形状 → 更新 .TextFrame.TextRange.Text（保留原有字体、颜色、段落格式）
        - 图表形状 → 读取【2-问卷模板.xlsx】对应数据，并通过 Chart.ChartData 或直接修改ppt-chart的 Series.Values / Categories 来更新数据（**保留图表所有样式、颜色、数据标签、轴格式**）
        - 数据仅需仅需简单的统计：例如均值、频率、汇总等，无需引入numpy等大型计算工具
        - 图片（如有）→ 如果需要替换则使用 .PictureFormat 或 Replace
5. 所有数据必须来自 Excel（建议使用 xlwings 打开 Excel 读取）
6. 最后 SaveAs "gemini-jules.pptx" 并 Quit PowerPoint

请在脚本中包含以下关键函数：
- `open_powerpoint()`
- `duplicate_slide(source_ppt, target_ppt, slide_index)`
- `update_text_shape(slide, shape_name, new_text)`
- `update_chart_data(slide, shape_name, excel_data_dict)`  ← 重点实现图表数据更新
- `build_final_ppt()`

### 第3步：验证
生成完成后，写一个对比验证脚本，使用 COM 打开【1-标准 ppt 模板】和【gemini-jules.pptx】，逐 shape 对比位置、大小、字体、颜色、图表类型等关键属性，给出详细差异报告。如果差异明显，必须自动调整代码并重新生成。

**最高优先级**：
- 视觉保真度必须达到 98%+（几乎和手动复制 slide 后改数据一样）
- 严禁使用 python-pptx 的任何函数
- 代码必须健壮、带错误处理和资源清理
- shape_detail_com.json 精确定位需要更新的shapes

'''