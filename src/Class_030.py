# Class 用于 import

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
##


# 再导入模块：导入模块中所有的类的语句为：【 from modulname import * 】
#from Class_030 import *
#from Function_030 import *












# 全局变量 color + delay ---------------------

delay = 0
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










## 【真-全局变量  真-全局变量  真-全局变量  真-全局变量  真-全局变量  真-全局变量  】
## ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★

# ========== 不知是否 working.... //  还是先按我自己的思路吧。。。 结果思路相同。。。
                # 无法working也就算了，但是这样定义的函数，导致整个文件结构无法running。。。。

# def _init():
#     global _global_dict
#     _global_dict = {}
#
# def set_value(key,value):
#     _global_dict[key] = value
#
# def get_value(key,defValue=None):
#     try:
#         return _global_dict[key]
#     except KeyError:
#         return defValue
#
# mc_book = xlwings.books.active         #11/05 最终决定延续过往的传统，使用 active 的 Excel，注意必须将其他 Excel文件关闭
# mc_sht0=mc_book.sheets['基础信息']
# target="测试样品"
# #sample_name=search(mc_sht0,target,column_offset=1).value           # class不能再引用function了（自然无法使用search函数），否则死循环了
# mc_sht0.api.Cells.Find(What=target, After=mc_sht0.api.Cells(mc_sht0.api.Rows.Count, mc_sht0.api.Columns.Count),\
#                       LookAt=xlwings.constants.LookAt.xlPart, LookIn=xlwings.constants.FindLookIn.xlFormulas,\
#                       SearchOrder=xlwings.constants.SearchOrder.xlByColumns,\
#                       SearchDirection=xlwings.constants.SearchDirection.xlNext, MatchCase=False).select()
# mc_book.selection.offset(row_offset=0,column_offset=1).select()
# sample_name = mc_book.selection.value
#
# set_value('sample_name',sample_name)
















## 【文本-类 Class  文本-类 Class  文本-类 Class  文本-类 Class  文本-类 Class  文本-类 Class 】
## ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★



# =========================================================

class Text_Box:
    
    # 【Input 】: 尽量简化输入接口，只保留 Slide，意味着在当前ppt页面新建文本框
    
    # 【Output】: 1、 .textbox —— 用来从外部修改【test.textbox.】 Text 文本内容
                # 2、 .shape   —— 用来从外部修改【test.shape.】 Left / Top / Height / Width 位置大小信息 
    
    '定义一个不带任何格式的文本框（实质上，还是做了不少限制，例如，字体=微软雅黑 / 行间距=1 / 默认靠左对齐 / 取消填充边框 / 颜色 - 黑色 / ）'

    def __init__(self,Slide):   #我只需要将shape作为接口，自然就能编辑 Left,Top,Text 这些东西，试试看 // 果然没问题~

        temp = Slide.Shapes.AddTextbox(1,100,100,100,100).TextFrame.TextRange
        
        self.tr = temp              # 这样就能从外部，通过 【test.textbox.】 Text  调用和修改了！

        self.shape = temp.Parent.Parent  # 这样就能从外部，通过 【test.shape.】 Left / Top / Height / Width 调用了！


        # ------ Text_Box 默认格式 --------
        
        self.tr.Text = '默认文本框'

        self.tr.Font.NameFarEast =  "微软雅黑"          #==========中文内容的字体

        self.tr.Font.Name = "Arial"                    # ==========英文内容的字体  测试 OK

        self.tr.ParagraphFormat.SpaceWithin = 1.0  # 【1.5 = 1.5倍行间距   2.0 = 2倍行间距】

        self.tr.ParagraphFormat.Alignment = 1   #【文字对齐方式： Alignment   1=靠左 / 2=居中 / 3=靠右  】

        self.shape.Line.Visible = 0 # 取消边框     =1 增加边框（可见）

        self.shape.Fill.Transparency = 1   # 透明度 = 1  ~  无任何填充颜色

        self.tr.Font.Color.RGB = 0     #' 黑色'

# =========================================================






class Line_Shape:

    '定义一个虚线形状，用于矩阵图中描绘使用（浅灰色、无阴影、粗细 = 、Type = -2）'

    def __init__(self,Slide,BeginX,BeginY,EndX,EndY):
          
        mc_line = Slide.Shapes.AddLine(BeginX=BeginX, BeginY=BeginY,EndX=EndX, EndY=EndY)        #mc_line = Slide.Shapes.AddLine(BeginX=10, BeginY=10,EndX=25, EndY=25)

        self.shape = mc_line    # 这是一个 Shape 对象 
        

        # ------ Line 默认格式 --------

        self.shape.Shadow.Visible=0     # 取消阴影
        
        self.shape.Line.DashStyle = 4   # 经典虚线，大家都爱用~

        self.shape.Line.ForeColor.RGB = light_gray    #   = 14277081 浅灰色 //  调用全局变量没问题！

        self.shape.Line.Weight = 0.5   # 粗细 = 0.5 磅




        
class Circle_Shape:

    '定义一个圆形，暂时使用邱岑的格式和大小（浅粉色、无阴影、粗细磅数 = 6、Type = 9）'

    def __init__(self,Slide, Left, Top):
        
        mc_shape = Slide.Shapes.AddShape(9,Left=Left, Top=Top, Width=113, Height=113)
        
        self.shape = mc_shape   # 这是一个 Shape 对象// 其实这两句可以合并为一句，也不影响执行，已尝试过  ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★


        
        # ------ Line 默认格式 --------

        self.shape.Shadow.Visible=0     # 取消阴影

        self.shape.Fill.Transparency =1  # 取消填充颜色 = 填充颜色全透明    # 【 =1 完全透明   //   = 0 完全不透明    //    =0.39 常用黑色半透明 】

        self.shape.Line.DashStyle = 4        #【4 = 虚线   1 = 实线   2 = 圆点   3 = 细圆点    5 = 细原点 + 虚线    6 =2个细原点 + 虚线】

        self.shape.Line.ForeColor.RGB = 12435956   #【邱岑使用的颜色 / 浅粉】   # 调整圆形描边线的颜色

        self.shape.Line.Weight = 6.0              #【邱岑使用的线条磅数】




class Triangle_Shape:
    '定义一个三角形，用作坐标系的轴标尺。AutoShapeType = 7，形状搞定了）'

    def __init__(self, Slide, Left, Top):

        mc_shape = Slide.Shapes.AddShape(7, Left=Left, Top=Top, Width = 9, Height = 9)

        self.shape = mc_shape  # 这是一个 Shape 对象// 其实这两句可以合并为一句，也不影响执行，已尝试过  ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★

        # ------ Line 默认格式 --------

        self.shape.Shadow.Visible = 0      # 取消阴影

        self.shape.Line.Visible = 0        # 取消边框     =1 增加边框（可见）

        self.shape.Fill.ForeColor.RGB = 0



# =========================================================










# 【Class 试验区】=========================================================

##class Circle_Shape:
##
##    '定义一个圆形，暂时使用邱岑的格式和大小（浅粉色、无阴影、粗细磅数 = 6、Type = 9）'
##
##    def __init__(self,Slide, Left=39, Top=98, Width=113, Height=113):
##
##        self = Slide.Shapes.AddShape(9,Left=39, Top=98, Width=113, Height=113)
##
##
##        # ------ Line 默认格式 --------
##
##        self.Shadow.Visible=0     # 取消阴影
##
##        self.Line.DashStyle = 4        #【4 = 虚线   1 = 实线   2 = 圆点   3 = 细圆点    5 = 细原点 + 虚线    6 =2个细原点 + 虚线】
##
##        self.Line.ForeColor.RGB = 12435956   #【邱岑使用的颜色 / 浅粉】   # 调整圆形描边线的颜色
##
##        self.Line.Weight = 6.0              #【邱岑使用的线条磅数】



'''
1、上面的Class定义，self.shape 可以合并为一句，不影响程序运行

2、甚至更夸张，直接 self = Slide.Shapes.AddShape(9,Left=39, Top=98, Width=113, Height=113)  也没问题 /// 长见识了 ★ ★ ★ ★ ★ ★ ★ ★ ★ ★

     —— 只不过，这样定义的Class，接口消失了，没办法在外部再修改这个Class对象。很奇怪，接口就这样消失不见了

     —— 而上面第一种则不会消失，可正常调用接口，这也是我灵光一闪、无意中发现的方法，值得珍藏~~~ 



'''





















        
class Title_Max(Text_Box):
    
##    ''' 封面主标题：
##        文字内容：自定义（封面报告标题）
##        字体加粗：Bold = 1
##        字体  颜色： 白色
##        字体  大小： 54
##        高度Height： 69
##        宽度Width ： 400
##    '''

    def __init__(self,Slide,Left,Top,Text,scale=1):

        Text_Box.__init__(self,Slide)  # ★★★★★★★  继承的核心代码，竟然忘记写了

        self.tr.Text = Text

        # ------ Title_Max 格式 --------
        self.tr.Font.Bold = 1
        self.tr.Font.Color.RGB = 16777215     #' 白色'
        self.shape.Fill.ForeColor.RGB = 0 # 底纹（形状）填充黑色
        self.tr.Font.Size = 50
        self.shape.Fill.Transparency=1  # 透明度 = 1  ~  无任何填充颜色
        

        self.shape.Height = 69
        self.shape.Width = 560*scale

        self.shape.Left = Left
        self.shape.Top = Top








class Title_1(Title_Max):  #试试看能不能反复继承多次（第二次）？    竟然可以反复继承。。。。。。 然后微调标题文字

                           # 2023 反复继承，结果出现干扰，Text_Box 底纹自动变蓝色，必须使用更多限制条件，避免默认的格式出现意外的惊喜

    '内容页（左上角）主标题'

    
    def __init__(self,Slide,Left,Top,Text,scale=1):

        Title_Max.__init__(self,Slide,Left,Top,Text)  # ★★★★★★★  继承的核心代码，竟然忘记写了

        # ------ Title_1 格式 --------
        self.tr.Font.Size = 32
        self.shape.Fill.ForeColor.RGB = 0  # 底纹（形状）填充黑色
        self.shape.Fill.Transparency = 0     # 透明度 5 = 0  ~  显示填充颜色
        self.tr.ParagraphFormat.Alignment=2    #【文字对齐方式： Alignment   1=靠左 / 2=居中 / 3=靠右  】

        #我能直接调用self吗？ self = slide 好奇害死猫  果然不行，虽然 self = slide 在本质上是成立的，但没有定义，就不能使用！！    
        #return self.SlideID

        self.shape.Height = 50
        self.shape.Width = 160*scale




        


class Title_2(Title_Max):    #试试看能不能反复继承多次（第三次）？    竟然可以反复继承。。。   

                                
    '内容页（副）标题，带菱形项目编号'

    
    def __init__(self,Slide,Left,Top,Text,scale=1):

        Title_Max.__init__(self,Slide,Left,Top,Text)  # ★★★★★★★  继承的核心代码，竟然忘记写了

        # ------ Title_2 基本格式 --------
        self.tr.Font.Size = 20
        self.shape.Fill.Transparency=1   # 透明度 = 1  ~  无任何填充颜色
        #self.textrange.ParagraphFormat.Alignment=2   #【文字对齐方式： Alignment   1=靠左 / 2=居中 / 3=靠右  】
        self.tr.Font.Color.RGB = 0     #' 黑色'
        self.tr.Text = Text
        
        # ------ * Title_2 符号与编号格式 *  --------       
        self.tr.ParagraphFormat.Bullet.Visible = -1  # 【0 = 取消编号】 【-1 = 启用编号】        
        self.tr.ParagraphFormat.Bullet.Type = 1
        self.tr.ParagraphFormat.Bullet.Font.Name = 'Wingdings'   #    'Arial'  # 一个符号格式竟然破解了一上午。。。
        self.tr.ParagraphFormat.Bullet.Character = 117 


        self.shape.Height = 45
        self.shape.Width = 250*scale






class Title_3(Title_Max):       

                                
    '邱岑使用的条形图，左侧标题（普通文字、加粗、黑色字体、无边框底纹'

    
    def __init__(self,Slide,Left,Top,Text,scale=1):

        Title_Max.__init__(self,Slide,Left,Top,Text)  # ★★★★★★★  继承的核心代码，竟然忘记写了

        # ------ Title_2 基本格式 --------
        self.tr.Font.Size = 20
        self.shape.Fill.Transparency=1   # 透明度 = 1  ~  无任何填充颜色
        self.tr.ParagraphFormat.Alignment=1   #【文字对齐方式： Alignment   1=靠左 / 2=居中 / 3=靠右  】
        self.tr.Font.Color.RGB = 0     #' 黑色'
        self.tr.Text = Text

        
        self.shape.Height = 32
        self.shape.Width = 125*scale























class Text_1(Text_Box):       
                               
    '正文（普通），长度默认覆盖整个页面（可调节）'
    
    def __init__(self,Slide,Left,Top,Text,scale=1):  # 这里设置一个scale，1=满屏（默认）  0.5 = 1/2屏 这样就能方便调整★★

        Text_Box.__init__(self,Slide)  # ★★★★★★★  继承的核心代码，竟然又双叒叕........忘记写了

        self.tr.Text = Text

        self.tr.Font.Size = 18
        self.tr.ParagraphFormat.SpaceWithin = 1.5  # 【1.5 = 1.5倍行间距   2.0 = 2倍行间距】
        
        self.shape.Height = 69
        self.shape.Width = 902*scale

        self.shape.Left = Left
        self.shape.Top = Top






class Text_Bullet(Text_Box):       
                               
    '测试方法描述（带1、2、3、项目编号），长度默认覆盖整个页面（可调节）'
    
    def __init__(self,Slide,Left,Top,Text,scale=1):  # 这里设置一个scale，1=满屏（默认）  0.5 = 1/2屏 这样就能方便调整★★

        Text_Box.__init__(self,Slide)  # ★★★★★★★  继承的核心代码，竟然又双叒叕........忘记写了
        
        self.tr.Text = Text

        self.tr.Font.Size = 16
        self.tr.ParagraphFormat.SpaceWithin = 1.5  # 【1.5 = 1.5倍行间距   2.0 = 2倍行间距】
        
        self.shape.Height = 69
        self.shape.Width = 902*scale

        self.shape.Left = Left
        self.shape.Top = Top

        # ------ * Text_Bullet 符号与编号格式 *  --------       
        self.tr.ParagraphFormat.Bullet.Visible = -1  # 【0 = 取消编号】 【-1 = 启用编号】        
        self.tr.ParagraphFormat.Bullet.Type = 2
        self.tr.ParagraphFormat.Bullet.Font.Name = 'Arial'   #    'Arial'  # 一个符号格式竟然破解了一上午。。。
        self.tr.ParagraphFormat.Bullet.Style = 3





class Result_Bullet(Text_Box):       
                               
    '测试结论（带■项目编号），长度默认覆盖整个页面（可调节）'
    
    def __init__(self,Slide,Left,Top,Text,scale=1):  # 这里设置一个scale，1=满屏（默认）  0.5 = 1/2屏 这样就能方便调整★★

        Text_Box.__init__(self,Slide)  # ★★★★★★★  继承的核心代码，竟然又双叒叕........忘记写了
        
        self.tr.Text = Text

        self.tr.Font.Size = 16
        self.tr.ParagraphFormat.SpaceWithin = 1.5  # 【1.5 = 1.5倍行间距   2.0 = 2倍行间距】
        
        self.shape.Height = 100
        self.shape.Width = 902*scale

        self.shape.Left = Left
        self.shape.Top = Top

        # ------ * Text_Bullet 符号与编号格式 *  --------       
        self.tr.ParagraphFormat.Bullet.Visible = -1  # 【0 = 取消编号】 【-1 = 启用编号】        
        self.tr.ParagraphFormat.Bullet.Type = 1
        self.tr.ParagraphFormat.Bullet.Font.Name = 'Wingdings'   #    'Arial'  # 一个符号格式竟然破解了一上午。。。
        self.tr.ParagraphFormat.Bullet.Character = 110  




class Result_Bullet_small(Text_Box):       
                               
    '2025 问卷优缺点结论（不带■项目编号），长度默认覆盖整个页面（可调节）'
    
    def __init__(self,Slide,Left,Top,Text,scale=1):  # 这里设置一个scale，1=满屏（默认）  0.5 = 1/2屏 这样就能方便调整★★

        Text_Box.__init__(self,Slide)  # ★★★★★★★  继承的核心代码，竟然又双叒叕........忘记写了
        
        self.tr.Text = Text

        self.tr.Font.Size = 14
        self.tr.ParagraphFormat.SpaceWithin = 1.5  # 【1.5 = 1.5倍行间距   2.0 = 2倍行间距】
        
        self.shape.Height = 100
        self.shape.Width = 902*scale

        self.shape.Left = Left
        self.shape.Top = Top

        # ------ * Text_Bullet 符号与编号格式 *  --------       
        self.tr.ParagraphFormat.Bullet.Visible = 0  # 【0 = 取消编号】 【-1 = 启用编号】
                                                                           #  不知为何，修改不了，只能在外部用代码修改了
                                                                           # mc_shape.TextFrame.TextRange.ParagraphFormat.Bullet.Visible=0
        self.tr.ParagraphFormat.Bullet.Type = 1
        self.tr.ParagraphFormat.Bullet.Font.Name = 'Wingdings'   #    'Arial'  # 一个符号格式竟然破解了一上午。。。
        self.tr.ParagraphFormat.Bullet.Character = 110



        







class Text_small(Text_Box):       
                               
    '矩阵图中的备注小文字，没什么特别格式，就是字体非常小'
    
    def __init__(self,Slide,Left,Top,Text,scale=0.1):  # 这里设置一个scale，1=满屏（默认）  0.5 = 1/2屏 这样就能方便调整★★

        Text_Box.__init__(self,Slide)  # ★★★★★★★  继承的核心代码，竟然又双叒叕........忘记写了

        self.tr.Text = Text

        self.tr.Font.Size = 7   # 10 号字体显得有点大
        self.tr.ParagraphFormat.SpaceWithin = 1.5  # 【1.5 = 1.5倍行间距   2.0 = 2倍行间距】
        
        self.shape.Height = 20
        self.shape.Width = 902*scale

        self.shape.Left = Left
        self.shape.Top = Top


class Text_scale(Text_Box):
    '矩阵坐标系的单位数字，字体单独定制。。'

    def __init__(self, Slide, Left, Top, Text, scale=0.046):  # 这里设置一个scale，1=满屏（默认）  0.5 = 1/2屏 这样就能方便调整★★

        Text_Box.__init__(self, Slide)  # ★★★★★★★  继承的核心代码，竟然又双叒叕........忘记写了

        self.tr.Text = Text
        self.tr.Font.Bold = 1
        self.tr.ParagraphFormat.Alignment = 2     #【文字对齐方式： Alignment   1=靠左 / 2=居中 / 3=靠右  】  ★★★★★★★★★

        self.tr.Font.Size = 12  # 坐标轴的字体大小
        self.tr.ParagraphFormat.SpaceWithin = 1.5  # 【1.5 = 1.5倍行间距   2.0 = 2倍行间距】

        self.shape.Height = 22
        self.shape.Width = 902 * scale

        self.shape.Left = Left
        self.shape.Top = Top


print('Class文件导入完成！\n')




