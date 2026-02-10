# 第 27个程序，基于Excel文件，自动生成PPT报告








# -------------------------------------------------------------------------
#2024更新： import模块（自动安装）  越来越复杂。。。不管了，能用就好
##import openai

import subprocess
import sys
import importlib.util

def install(package):
    
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        
packages = ["openai","xlwings"]   # 安装名称（例如：pip python-docx）★ ★ ★ ★ ★

modules = ["openai","xlwings"]           # 导入名称（例如：import docx）★ ★ ★ ★ ★ ★ 

for package, module in zip(packages, modules):
    is_installed = importlib.util.find_spec(module)
    if is_installed is None:
        print(f'开始安装{package}...\n')
        install(package)
        print(f'    {package}安装完成@！！\n')
        print(f'        尝试再次导入{module}\n')
        try:
            exec(f"import {module}")
            print(f'          {module}导入成功\n\n')
        except ImportError:
            print(f'{module}导入失败\n')
    else:
        print(f'{module}导入成功\n')
# -------------------------------------------------------------------------








import time
from openai import OpenAI

import os

##OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")

##openai_key = os.getenv('OPENAI_KEY')
##
##os.environ["OPENAI_API_KEY"] = openai_key    #"sk-4iuTJDvUdt3YJ6b6zv*****************"


##print(os.environ["OPENAI_API_KEY"])
##print(openai_key)
##a = input('Debugging...........')


# 2024 GPT-4 key
#mc_key = os.getenv('Mc_KEY')

#global mc_model, model_mini

#mc_model = 'gpt-5-mini'   # = 'gpt-5'    # 统一用mini调试吧      #gpt-5.1 也同步上线了
                                                                            # https://platform.openai.com/docs/pricing
#model_mini = 'gpt-realtime'   #'gpt-5-mini'

                # 'gpt-3.5-turbo-0125'
                # 'gpt-4-0125-preview'   2025-11✔
                # 'gpt-5'                2025-11✔
                # 'gpt-5-chat-latest'











# ★ ★ ★ ★ ★  公共信息（每个文件都复制一遍。。）  ★ ★ ★ ★ ★ 


# ----- 导入第三方Module ------
import ast
from datetime import datetime
import win32com.client
import xlwings

import os, time, random, re
from datetime import datetime


# ----- Main()独占：导入Class 和 Function 文件 ------

# 先解决路径问题
import sys                           # 导入sys模块，来调取path 系统默认路径
mc_path = os.getcwd()
sys.path.append(mc_path)     # 将我的mod路径，添加至 系统默认路径
#print(sys.path)  # 检查下路径有没有问题



# 在修改了系统路径前提下， 再导入模块：导入模块中所有的类的语句为：【 from modulname import * 】
# 2025 更新了 package 的导入方式，多亏张老师的教材！！ 感谢张老师
from src.Class_030 import *
from src.Function_030 import *







# 全局变量 color + delay ---------------------
    # 虽然我注释掉这部分赋值，也不会报错，但原理尚未完全明晰

#mc_gpt = input('是否启用GPT-5？ 请输入 y/n（默认值 = No！！）： y/n') or 'n' # 'y' # input('是否启用GPT-4？ 请输入 y/n（默认为No！！）： ') or 'n' #

#Claude 帮我写了全新的按钮窗体函数，nice
mc_model = ask_gpt_model()

if mc_model !='n':
    mc_gpt = 'y'   #如果选择了模型，那么就启用 GPT


delay = 2
#time.sleep(random.random()*delay)



black = 0                     #' 黑色'
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
    # 终于摆脱了一次修改3个文件的痛苦经历。。。。。。。。。。。

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
from src.Global_var_030 import *
dic_matrix = get_value('dic_matrix')




# ★ ★ ★ ★ ★  公共信息（每个文件都复制一遍。。）  ★ ★ ★ ★ ★ 










##import os
##print("HTTP_PROXY =", os.getenv("HTTP_PROXY"))
##print("HTTPS_PROXY=", os.getenv("HTTPS_PROXY"))
##print("ALL_PROXY  =", os.getenv("ALL_PROXY"))
##


##p = input('....................')











## 【程序主体 Main() 程序主体 Main() 程序主体 Main() 程序主体 Main() 程序主体 Main() 】
## ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★        



# ====================================================================================================================================================================

# 【0】调试准备动作



   # 【0-1】接管ppt文件与页面 ------------------- 前期调试使用接管模式，后面改为自动打开。✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔✔

##mc_app = win32com.client.Dispatch('PowerPoint.Application')
##mc_ppt = mc_app.ActivePresentation
###mc_slide = mc_app.ActiveWindow.Selection.SlideRange


mc_app = win32com.client.Dispatch('PowerPoint.Application')
mc_app.DisplayAlerts = 0        # 不警告 / 无弹窗
mc_app.Visible = True             #  or 不可见 = False
#mc_path = os.getcwd()
mc_ppt = mc_app.Presentations.Open(mc_path+r'\src\Template 2.1.pptx')


slide_all = len(mc_ppt.Slides)  # 模板一共有几页？ 到时在末尾全部删除掉即可



   # 【0-2】Excel 与 定位 ， 调用 offset_search() 函数

mc_book = xlwings.books.active         #11/05 最终决定延续过往的传统，使用 active 的 Excel，注意必须将其他 Excel文件关闭
mc_sht0=mc_book.sheets['基础信息']


target="测试样品"
sample_name=search(mc_sht0,target,column_offset=1).value
#print(sample_name)
#a = input('.................debug..............')
#sample_name = get_value('sample_name')

target="对比鞋款"
control=search(mc_sht0,target,column_offset=1).value

target="总样品数量"
sample_count=search(mc_sht0,target,column_offset=1).value

target="底模编号"
sample_num=search(mc_sht0,target,column_offset=1).value

target="功能描述"
sample_detail=search(mc_sht0,target,column_offset=1).value

target="样品图片"
sample_pic = search(mc_sht0,target,column_offset=1)
sample_pic = [sample_pic]


   # 【0-3】生成日期

today = str(datetime.today())
today = today[0:10]











# ========================================================================================


# 【1】首页，总目录，Title 内容制作
mc_ppt.Slides(1).Select()
mc_ppt.Slides(1).Copy()

time.sleep(random.random()*delay)
mc_slide = mc_ppt.Slides.Paste()  # 默认粘贴到末尾


Left = 77
Top = 162
Text = sample_name + '功能性测试报告'
title = Title_Max(mc_slide,Left,Top,Text)


for i in mc_slide.Shapes:

    if i.TextFrame.HasText == -1:   #  【0 = 不包含文本】 / 【-1 = 包含文本】

        temp = i.TextFrame.TextRange.Text

        if '日期' in temp:
            
            i.TextFrame.TextRange.Text = today






### 【2】然后，目录页面制作（总目录，分页目录） 暂时采用赋值粘贴的方式，后面再调整吧 ----------------------------------------------------------------------

# 复制总目录
mc_ppt.Slides(2).Select()
mc_ppt.Slides(2).Copy()

X = len(mc_ppt.Slides) + 1

time.sleep(random.random()*delay)
mc_slide = mc_ppt.Slides.Paste(X)  # 长期使用发现不行，还是要带参数比较稳定   

# 复制分页目录（） 函数
content_slide(mc_ppt)















### 【3】开始项目背景页面       ------------------------------------------------------------------------------------------------------------------------------

# 【3.1】先复制一遍空白模板页面
mc_ppt.Slides(4).Select()
mc_ppt.Slides(4).Copy()

X = len(mc_ppt.Slides) + 1

time.sleep(random.random()*delay)
mc_slide = mc_ppt.Slides.Paste(X)  # 长期使用发现不行，还是要带参数比较稳定


#要删除2次才能完全清空。。。
for shp in mc_slide.Shapes:    #删除所有的内容（shape）：插入一个完全空白的模板（白板）
    shp.Delete()
for shp in mc_slide.Shapes:    #删除所有的内容（shape）：插入一个完全空白的模板（白板）
    shp.Delete()




#  左上角主标题  -  背景  
Left = 15
Top = 15
Text =  '背景'
title = Title_1(mc_slide,Left,Top,Text)


#  副标题 - 测试背景

Left = 19
Top = 73
Text =  '测试背景'
title = Title_2(mc_slide,Left,Top,Text)






#  【3.2】项目背景 - 文字描述
#  局部文字格式修改，没想到这么麻烦  待解决 ===  已解决！！

Left = 38
Top = 217

#Text =  '科技中心收到1款 '+ sample_name + '（底模编号：' + sample_num  +   '），根据项目进度，基于该产品的科技功能定位，现对该款产品进行功能性测试。此次测试的对比样品分别为：' + control + '，本品及竞品的鞋款信息如图所示。'

try:
    Text = f'科技中心收到1款 {sample_name}(底模编号:{sample_num}),根据项目进度,基于该产品的科技功能定位,现对该款产品进行功能性测试。此次测试的对比样品分别为:{control},本品及竞品的鞋款信息如图所示。'
except:
    # 强制转换所有变量为字符串
    Text = f'科技中心收到1款 {str(sample_name)}(底模编号:{str(sample_num)}),根据项目进度,基于该产品的科技功能定位,现对该款产品进行功能性测试。此次测试的对比样品分别为:{str(control)},本品及竞品的鞋款信息如图所示。'


    
text = Text_1(mc_slide,Left,Top,Text)

# 加红加粗，分别凸显 【测试样】 、 【对比样】  /// 花了一下午时间~
key = sample_name
color_key(text.tr,key,red)

key = control
color_key(text.tr,key,light_blue)






#  【3.3】副标题 -  测试方法

Left = 19
Top = 327
Text='测试方法'
title = Title_2(mc_slide,Left,Top,Text)







#  测试方法 - 文字描述 （摘要）

Left = 394
Top = 361
Text='基础数据测试、实战测评、运动生物力学测试'
text = Text_1(mc_slide,Left,Top,Text)

# 加黑加粗
key = '基础数据测试、实战测评、运动生物力学测试'
color_key(text.tr,key,black)







#  测试方法 - 文字描述 （详细描述） from Excel  & 采用 Bullet格式

Left = 404
Top = 394
all_test_detail=test_detail(mc_book)    #调用函数，将测试方法文字扒出来：test_detail

text = Text_Bullet(mc_slide,Left,Top,all_test_detail[0],scale = 0.62)

# print(all_test_detail)
# print(type(all_test_detail))
# print(len(all_test_detail))
# p = input('调试中。。。。。')

# dict_test_detail = {
#     '测试列表':,
#     '测试方法':,
#     '测试标准':,
#     '参考文本':,
# }










#  【3.4】 是时候该处理图片函数了！ 结果还好，去年就写好了！！ 先复制过来，在修改试试

#通过上面的test_detail函数，已经将感兴趣（标记✔✔）的单元格装进 Text[1] 了。

mc_pic(*all_test_detail[1], temp_slide = mc_slide,Left = 43,Top = 376,scale=0.23)








# 再处理一次样品图片的排列 ====================================

                           #===================================这个区域的图片，需要小心，因为很有可能会太靠左，Left这个参数如果固定的话就比较坑爹。希望可以实现【组合 - 居中移动】的功能。

mc_pic(*sample_pic,temp_slide = mc_slide,Left = 120,Top = 89,scale=0.2)















### 【4】开始基础性测试部分--------------------------------------------------------------------------------------------------------------------------------

# 复制分页目录（） 函数，来到 【02、基础性能测试】页面
content_slide(mc_ppt)


# 然后开始Main()函数中最重要的循环，对【测试】sheet 进行遍历，依次按照顺序生成ppt  ★ 
 
  # 这边需要借用到 test_detail函数的结果 all_test_detail ——

  # all_test_detail[0] :  测试项目文字明细

  # all_test_detail[1] :  pic_list，返回单元格的list，来自动筛选图片、挑出来

  # all_test_detail[2] :  【测试项目】列表，用来查找核对sheet名称

  # all_test_detail[3] :  测试结论话术，对应的测试复制对应的话术


##  s>>> all_test_detail[2]
##  ['基本参数测试', '机械测试', '动作捕捉测试', '跑者实战测评']

##  c = chart  【基本参数测试1】、【基本参数测试2】 ...


# 为了保存2025年GPT的回答内容？ 但 GPT_5函数已经有历史聊天功能了呀，本质上是一种冗余。。。 先观察下看看 ==========================================================================================
#completion_list = []
messages = []


for c in mc_book.sheets:   # c = 针对所有的 sheet（一共4个sheet）

    for s in range(0,len(all_test_detail[2])):   # 分别检查名字，s 依次 = 0 / 1 / 2 / 3 / 4 .....
 
        if all_test_detail[2][s] in c.name:               # 例如，假设 all_test_detail[2][s] 依次 = '基础参数测试'  / 生物力学测试.... ， 检查  这个名字  是否包含在 当前【跑鞋机械力学测试1】sheet中
                                                          # 2024-02-21 调试过了，上述代码不能反过来写，因为sheet1【机械测试1】 sheet2【机械测试2】 这种情况必须存在

            Excel_zoom(c, 60)
            # 首先，复制一张空白内容页         
            mc_ppt.Slides(4).Select()
            mc_ppt.Slides(4).Copy()
            X = len(mc_ppt.Slides) + 1

            time.sleep(random.random()*delay)
            mc_slide = mc_ppt.Slides.Paste(X)  #  长期使用发现不行，还是要带参数比较稳定


            # 制作二级附标题
            Left = 43
            Top = 53
            Text = all_test_detail[2][s] +  '结果'          #例如：基本参数/机械/实战 测试  + 结果
            title = Title_2(mc_slide,Left,Top,Text)
            title.tr.ParagraphFormat.Bullet.Character = 252    #  ✔✔  | 详细见【022-Tools PPT】- （文本）小节的内容


            #进行制表
            temp_c = make_chart(c, mc_slide)  #  制表函数将会return    temp_c[0] Left,       temp_c[1] Top,    temp_c[2]  Height,   temp_c[3]  Width  // 这里输出的是最后一个chart的位置大小参数

                                              # bug ★ 循环双重return的问题，终于解决了（不能两个return函数相互嵌套，因为一旦return，函数就结束了）。 函数的接口设计有待提高呀。。。


            # # *------------------ 原代码 （已弃用） ----------------
            # # 补充 结论描述（参考文字）  继续再来一个 class 文本类
            # # 文本来源是 ——  all_test_detail[3]
            # # 真实文本框的位置大小参数为： Left = 15      Top = 352     Height = 65   Width = 941
            #
            # #Height = temp_p[2]  #temp_p[2]
            # #Width = 900   #  temp_p[3]    #文本Class中默认了Width参数，不用再操心
            #
            # Left  = 15    # temp_p[0]
            # Top   =  temp_c[1] + temp_c[2] + 5    # 在最有一张chart的基础上，继续下移5个像素
            #
            #
            #
            # text = Result_Bullet(mc_slide,Left,Top,all_test_detail[3][s],scale = 1)    # all_test_detail[3] 是一个list， 因此循环用 range 可以精确遍历
            #                                                                            # 现在有个问题是，\r不管用了。。。 =========
            # # *------------------ 原代码 （已弃用） ----------------


            
            # 是的，计划把结论描述（参考文字）部分使用chat_GPT()函数替换掉，让结论描述更精确。那么这个函数运行的前提，便是循环遍历（2层嵌套）
            # 参考 make_chart函数，对整个sheet进行【表1】【表2】【表3】...的遍历，函数本身可以处理当前sheet中的所有chart
            # 因此，这个chat_GPT（）函数也同样，需要对整个sheet的chart数据源进行处理，每个【实验条件】生成一段话。


            # 这段代码在 make_chart函数中重复了，但因为无法共享，只能再查找一遍。把当前sheet中的【图表1 / 2 / 3 ...】的 cells 标记出来
            
            mc_completion =''


            # 思考再三，还是在遍历图表的外层构建 all_info （因为我不想检查图表标题，而是简单基于sheet名称来判断测试方法）
                # 理想状况下，只需要构建len(测试方法) 数量的dict即可， 例如，3个测试方法 = 3个dict
                    # 含【cell】  【测试标准】 【参考文本】   【测试数据（放到图表遍历内部）】 【测试样品名称（例如飞影PB）】【对比样品名称（例如nike）】

            all_info = {}  # 用一次就清空掉，下一个测试方法 / sheet 再重新构建 all_info ，这个思路是对的，因为每页PPT请求1次GPT是合理的。 只要不是每个chart请求1次就行。
            
            all_info = {'测试标准': all_test_detail[4][s], '参考文本': all_test_detail[3][s], '测试样品名称': sample_name , '对比样品名称':control}
            
            test_data = []

            for i in range(1,100):
                
                target = "图表"+str(i)
                
                find = search(c,target)

                if find != None:    # and mc_gpt == 'y':
                    #cell_list.append(find)           # temp_list =  图表1 / 图表2 / 图表3... 所在的range对象 ； 针对这个对象 调用GPT函数


                    temp_data = str(find.api.CurrentRegion.Value) + '\n'
                    test_data.append(temp_data)
                    #print(test_data)
                    #a = input('.......debug.............................')


                else:
                    break



            if mc_gpt == 'y':


                # 针对每个 【图表】//【sheet】，调用一次GPT，获取评论
                all_info['测试数据'] = str(test_data)
                #print(f'针对【{c.name}】sheet，即将开始请求[OPENAI]服务器，使用模型为{mc_model} ——————\n\n')  # 优雅！！

                # 画蛇添足 - 在PPT中打字提示。。。
                Left = 15  # temp_p[0]
                Top = 400
                hint = '[OPENAI]服务器请求发送中，请耐心等待.........'
                hint_shape = Title_3(mc_slide,Left = Left,Top = Top, Text = hint,scale=4)   # 2024 注意搞清楚Class文件生成的是一个什么对象，不然很麻烦



                # 请求GPT服务器（如需调试，请进入 function.py 文件，直接对GPT函数进行调试！！！）
                mc_prompt = gen_mc_prompt(**all_info)
                #print(mc_prompt)
                #a = input('.......debug............................')

                mc_completion = GPT_5(mc_prompt,model=mc_model)
                
                #mc_completion = str(mc_reply)  # 先保留疑问，需要str吗？=============================================

                                            #messages.extend(mc_reply[1])     # 在 main.py 文件中，保存messages内容。试试看这样是否会重复

                                            #---------------------------------------------------- ing......  考虑下是否需要将GPT函数挪到main.py中，如果不挪，身为 global参数 的 messages 能否被实时修改？


                #mc_completion += f'{str(GPT_reply)}'

                #completion_list.append(mc_completion)   # 为最终结论生成，需要再请求一次GPT服务器


                # 画蛇添足 - 在PPT中打字提示。。。 然后别忘记删除
                hint_shape.tr.Text = '成功获取 [OPENAI] Reply ！'
                time.sleep(1)
                hint_shape.shape.Delete()  #TextRange 不能直接 Delete，必须使用 Parent 回到 Shape函数才能删除
                
                mc_completion = ('【GPT-5解析结论如下（该部分内容为AI自动生成，请认真审阅后，酌情使用）：】\n' + mc_completion)








                
            else: # 不启用GPT，这是所有人最常用的功能。。。。
                
                 # 源代码已经没啥用处了，先留着吧
                text = '' #chat_GPT(*cell_list)                    # 旧的GPT函数。。问题出在这里！！！！！！！！！！！！！！！！！！！！！！！！！！！！★★★★★ *args调用时注意格式 ★★★★★
                text += '\n' + all_test_detail[3][s] # 暂时先留着不动了
                mc_completion = f'当前未启用 GPT-5，历史参考文本如下：\n【历史参考文本】\n{all_test_detail[3][s]}'
            




            # 针对每个sheet，都使用一次 chat 函数，生成结论，最后用Result_Bullet函数，输出结论到ppt
            Left  = 15    # temp_p[0]
            Top   =  temp_c[1] + temp_c[2] + 5                # 在最后一张chart的基础上，继续下移5个像素
        
            Text = Result_Bullet(mc_slide,Left,Top,mc_completion,scale = 1)         # all_test_detail[3] 是一个list， 因此循环用 range 可以精确遍历



            # 对结论进行染色
            key = sample_name
            color_key(Text.tr, key, red)

            # 染色多个对比样品，re yyds
            key = [z.strip() for z in re.findall(r'[\u4e00-\u9fa5a-zA-Z0-9.\s]+', control)]    # re.findall(r'[\u4e00-\u9fa5a-zA-Z0-9\s]+',control)    # 切割control
            # print(key,'\r',control)
            # xxx = input('Debug.........')
            for value in key:
                color_key(Text.tr, value, light_blue)

            print('【Sheet '+c.name+'】 处理完成！\n\n\n')


    

    if '矩阵' in c.name:   # 对 【矩阵图】 进行操作  /// 这个不能放到第二个循环中，只能在第一个循环里。
            
        # 首先，复制一张空白内容页 /// 具体复制那一页，就按字典来了，不要轻易调整模板顺序
        temp = search(c,'Matrix矩阵图',-2).value
        temp = dic_matrix[temp]
        mc_ppt.Slides(temp).Select()
        mc_ppt.Slides(temp).Copy()
        X = len(mc_ppt.Slides) + 1

        time.sleep(random.random()*delay)
        mc_slide = mc_ppt.Slides.Paste(X)  #  长期使用发现不行，还是要带参数比较稳定


        #进行 matrix 制作
        make_matrix(c,mc_slide)
        print('【Sheet '+c.name+'】 处理完成！')
            





            











### 【5】实战测评部分 ======================================


# 复制分页目录（） 函数，来到 【03、实战测评】页面
content_slide(mc_ppt)



# 实战页面可能会有好几页，根据问卷情况，因此写一个嵌套函数，来避免重复代码，没别的意思。。。。

mc_slide = questionnaire_ppt(mc_ppt,mc_slide)   # 先增加一个【实战测试】空白页，然后再根据实际情况（问卷条目）酌情增加




mc_sht = None
for c in mc_book.sheets:   # c = 针对所有的 sheet（一共n个sheet）

    if '问卷' in c.name:

        mc_sht = c   # 注意，这个设定只能识别1张问卷sheet，这也是合理的， 我不对问卷sht进行循环，我只在1张问卷中循环多行（多个人）的数据
        mc_sht.select()
        Excel_zoom(mc_sht, 60)    # 2025-12 修订宇昂发现的bug，不是录频软件的问题

            



if mc_gpt == 'y' and mc_sht is not None:
    
    mc_work =questionnaire_Excel(mc_sht, mc_ppt, mc_slide, mc_model, sample_name=sample_name, mc_gpt=mc_gpt)     # 问卷逐页+汇总整合到同一函数
                                                                                                                            # 结果报错，需要传递 mc_model这个参数。。。 GPT也是这样建议我
                                                                                                                              # 再次报错，mc_ppt也需要传入

    mc_sht = mc_work[0]
    mc_slide = mc_work[1]


    



























### 【6】结论部分


# 复制分页目录（） 函数，来到 【04、结论】页面
content_slide(mc_ppt)



# 【6.1】先复制一遍结论模板页面
mc_ppt.Slides(10).Select()
mc_ppt.Slides(10).Copy()

X = len(mc_ppt.Slides) + 1

time.sleep(random.random()*delay)
mc_slide = mc_ppt.Slides.Paste(X)  # 长期使用发现不行，还是要带参数比较稳定 



# 【6.2】把鞋款图片摆放进去
mc_pic(*sample_pic,temp_slide = mc_slide,Left = 120,Top = 89,scale=0.2)




# 【6.3】 2024 干脆让GPT也生成结论看看
if mc_gpt == 'y':


    # 画蛇添足 - 在PPT中打字提示。。。
    Left = 15  # temp_p[0]
    Top = 400
    hint = '[OPENAI]服务器请求发送中，请耐心等待.........'
    hint_shape = Title_3(mc_slide, Left=Left, Top=Top, Text=hint,scale=4)  # 2024 注意搞清楚Class文件生成的是一个什么对象，不然很麻烦


    # 需要传入【sample_name】1个参数。
    mc_prompt=gen_result_prompt(sample_name)
    


    
    mc_completion = GPT_5(mc_prompt,model=mc_model)

    #mc_completion = mc_reply
    

    # 画蛇添足 - 在PPT中打字提示。。。 然后别忘记删除
    hint_shape.tr.Text = '成功获取 [OPENAI] Reply ！'
    time.sleep(1)
    hint_shape.shape.Delete()  # TextRange 不能直接 Delete，必须使用 Parent 回到 Shape函数才能删除



    Left = 51
    Top = 281
    
    Text = Result_Bullet(mc_slide, Left, Top, mc_completion, scale=1)

    # 对结论进行染色
    key = sample_name
    color_key(Text.tr, key, red)



    

    #这里需要Debug一下，看看messages如果只是在 funciton文件内部传输，能否实现？ 测试结果OK，没问题
    
##    all_in_one_msg = get_messages()
##    
##    for msg in all_in_one_msg:
##        
##        print(f'\n{msg}\n')
  

            # 这个简单的关于全局变量的争议，我跟GPT-5.1 请教了一下午。。。。。
            
            # 方法1：简单来说：import src.Function_030 as F
            
                            # 之后就能直接使用  all_messages = F.messages
                            
                                             # 这里就拿到了 function 文件里的全局 messages
            
             # 方法2：但我的用法是：from src.Function_030 import *

                             #技术上讲，那么在 main.py 里，确实可以直接写：all_messages = messages 

                                                 # 它会引用到 Function_030.py 里的那个全局 messages

                                                 #  但由于我并没有在 main文件中定义 messages，可能造成混淆和来源不清晰

                                                 #  因此GPT更推荐第一种方法 ！！ 或者直接定义一个函数，专门用来 return msg （如上）










### 【7】 Goodbye

mc_ppt.Slides(11).Select()
mc_ppt.Slides(11).Copy()

X = len(mc_ppt.Slides) + 1

time.sleep(random.random()*delay)
mc_slide = mc_ppt.Slides.Paste(X)  # 长期使用发现不行，还是要带参数比较稳定 














### 【8】 文件操作
for p in range(1,slide_all+1):
    mc_ppt.Slides(1).Delete()
    
ppt_save(mc_ppt,sample_name,mc_path)
print('请注意，文件已保存完成！')
print(mc_path)









# 小彩蛋：别急着退出，继续聊天，直到用户关闭程序
chat_room(mc_model)













##===========================================================================================

#mc_slide = mc_app.ActiveWindow.Selection.SlideRange

#mc_shape = mc_slide.Shapes.AddShape(5,Left=50, Top=50, Width=125, Height=17)

#mc_text_range = mc_slide.Shapes.AddTextbox(1,100,100,100,100).TextFrame.TextRange


##'''
##                        RGB = 16777215             #' 白色'
##                        RGB = 0                    #' 黑色'
##                        
##                        RGB = -16776961  // 255    #' 红色'
##                        RGB = 15773696             #' 天蓝色'
##
##                        
##                        RGB = 65535      #' 亮黄 '                                
##                        RGB = 5287936    #' 绿色 '
##                        RGB = -16727809  #' 橙色 '
##                        RGB = 10498160   #' 紫色 '
##                        RGB = 6299648    #' 墨蓝色 '
##
##
##
##            blck = 0                      #' 黑色'
##            white = 16777215              #' 白色'
##
##            red = 255                     #' 红色'
##            green = 5287936               #' 绿色 '
##            dark_blue = 6299648           #' 墨蓝色 '
##            light_blue = 15773696         #' 天蓝色'
##
##            yellow = 65535                #' 亮黄 '  
##            orange = 16727809             #' 橙色 '
##            purple = 10498160             #' 紫色 '
##                                                          
##'''




# 【目录】 ===================

##Slides(1) —— 【模板1】 总标题

##Slides(2) —— 【模板2】 总目录页

##Slides(3) —— 【模板3】 分页目录标题页   content_slide(mc_ppt)

##Slides(4) —— 【模板4】 正文页面

##Slides(5) —— 【模板5】 再见！

# 【目录】 ===================
















