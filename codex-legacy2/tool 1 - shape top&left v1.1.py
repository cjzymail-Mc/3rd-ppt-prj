
# 小工具1，用于识别ppt中，每个文本框 / shape 的位置信息

   # 经过一个上午的努力，顺利完成，也在这个过程中学习了新的技能（判断文本、复制粘贴Slide）




# 全局变量 color ---------------------

black = 0                      #' 黑色'
white = 16777215              #' 白色'

red = 255                     #' 红色'
green = 5287936               #' 绿色 '
dark_blue = 6299648           #' 墨蓝色 '
light_blue = 15773696         #' 天蓝色'

yellow = 65535                #' 亮黄 '
orange = 16727809             #' 橙色 '
purple = 10498160             #' 紫色 '




##color_value[0] #or -16776961     # 样品1染色
##color_value[1] #or 15773696      # 样品2染色
##color_value[2] #or 5287936     # 样品3染色
##color_value[3] #or -16727809   # 样品4染色
















import ast
from datetime import datetime
import xlwings

import os, time, random
from datetime import datetime
import win32com.client


# ----- Main()独占：导入Class 和 Function 文件 ------

# 先解决路径问题
import sys                           # 导入sys模块，来调取path 系统默认路径
mc_path = os.getcwd()
sys.path.append(mc_path)     # 将我的mod路径，添加至 系统默认路径
#print(sys.path)  # 检查下路径有没有问题



# 再导入模块：导入模块中所有的类的语句为：【 from modulname import * 】
from Class_030 import *
#from Function_030 import *
from Global_var_030 import *
dic_matrix = get_value('dic_matrix')





# 接管当前打开的ppt【程序】
try:
    mc_app = win32com.client.GetObject(Class="PowerPoint.Application")
    #print("已接管现有的 PowerPoint 应用程序")
except:
    # 如果没有已经打开的应用程序，则创建一个新的应用程序实例
    mc_app = win32com.client.Dispatch("PowerPoint.Application")
    #print("创建新的 PowerPoint 应用程序实例")


        #可以二选一，GPT教我的新方法 ！！
            #mc_app =win32com.client.GetObject(Class="PowerPoint.Application")
            #mc_app = win32com.client.Dispatch('PowerPoint.Application')










# 接管当前打开的ppt【文件】
##mc_ppt = mc_app.Presentations.Application.ActivePresentation
mc_ppt = mc_app.ActivePresentation  # 二选一，两者等效





# 接管当前【Slide】
mc_slide = mc_app.ActiveWindow.Selection.SlideRange








##print('本页一共有 '+str(len(mc_slide.Shapes))+' 个图形...')

# 2023 【文本Shape】 和 【图形Shape】 到底有什么区别？  ★★★ 如何判断是否包含文本

   # 目前为止，我这边有4种Shape需要处理：

      # 1、 【-1      文本框】        最简单，也最常见，偶尔有空的文本框，要小心

      # 2、 【0       外部图片】      外面复制粘贴过来的图片

      # 3、 【0 / -1  矢量绘制图形】  条形图、线条、框框、圆形等，ppt可以绘制的图形 /// 包含两种情况，有的图形有文本，有的没有文本

      # 4、 【0       chart图表】     一般是从Excel粘贴过来的图表，不建议修改

##for i in mc_slide.Shapes:
##
##    print(i.TextFrame.HasText)



print('''请选择：\n

0、接管当前页面的 Shapes(1)\n
1、【单页】位置信息标注\n
2、【单页】大小信息标注\n
3、【Excel】Chart1锁定\n

X、【Matrix】执行和调试 Matrix 程序\n
X2、【Matrix】自动生成坐标系\n''')


option=(input('------------------请输入: '))




if option=='0':  # 0、接管当前页面的 Shapes(1)

    mc_shape = mc_slide.Shapes(1)

    if mc_shape.TextFrame.HasText == -1:

        mc_tr = mc_shape.TextFrame.TextRange



















if option=='1':  # 1、【单页】位置信息标注\n

    for i in range(1,len(mc_slide.Shapes)+1):

        shape = mc_slide.Shapes(i)

        Left = str(round(shape.Left))
        Top  = str(round(shape.Top))

        if shape.TextFrame.HasText == -1:   #  【0 = 不包含文本】 / 【-1 = 包含文本】

            shape.TextFrame.TextRange.Text='Left = '+ Left +'\n'+'Top = '+Top

        else:  # 假设是不包含文本的shape，那么需要新增一个shape，标注下相关信息

            temp = shape.Parent.Shapes.AddShape(5,Left=Left, Top=Top, Width=125, Height=50)
            temp.TextFrame.TextRange.Text ='Left = '+ Left +'\n'+'Top = '+Top










if option=='2':  # 2、【单页】大小信息标注\n

    # 5月6日 发现一个bug，  i in mc_slide.Shapes， i 竟然无法等价于 mc_slide.Shapes(1)。。。 无法理喻 /// 先修正了

    for i in range(1,len(mc_slide.Shapes)+1):

        shape = mc_slide.Shapes(i)

        Height = str(round(shape.Height))
        Width  = str(round(shape.Width))

        if shape.TextFrame.HasText == -1:   #  【0 = 不包含文本】 / 【-1 = 包含文本】

            shape.TextFrame.TextRange.Text='Height = '+ Height +'\n'+'Width = '+ Width

        else:  # 假设是不包含文本的shape，那么需要新增一个shape，标注下相关信息

            left = shape.Left
            top = shape.Top

            temp = shape.Parent.Shapes.AddShape(5,Left=left, Top=top, Width=125, Height=50)
            temp.TextFrame.TextRange.Text ='Height = '+ Height +'\n'+'Width = '+ Width






if option=='3':  # 3、【单页】chart 锁定

    import xlwings

    mc_book = xlwings.books.active
    mc_sht = mc_book.sheets.active

    mc_chart1 = mc_sht.charts[0]

    mc_cell = mc_book.selection



























if option=='x':  # 5、执行和调试 Matrix 程序   /// 拓展了函数自动生成坐标系的功能，没想到这个函数会变得这么复杂。。。


    def scale_matrix():

        ''' 独立出来用于调试Matrix坐标系参数 '''

        # 【步骤一】：先手工输入 value_l1, value_l2, left1, left2, value_t1, value_t2, top1, top2，模拟调试过程 （原始数据来自杨祖锐的 减震-回弹 矩阵）

##        value_l1 = 0.38
##        value_l2 = 0.6
##        left_l1  = 117
##        left_l2  = 826
##        top_l    = 82
##
##        value_t1 = 13
##        value_t2 = 8
##        top_t1   = 130
##        top_t2   = 385
##        left_t   = 70

        print('第一步，先确定坐标轴的上下位置关系：Top轴在上面时，rev= -1，Top轴在下面，rev= 1 ————\n')
        rev = int(input('----请输入: rev = '))  # 手工列一个rev值，Top轴在上面时，rev= -1，Top轴在下面，rev= 1

        if rev == 1:   # 针对正常坐标轴

            print('第二步，输入 Left 轴上任意2个点的【2个值】和【2个座标】，以及它们公共【Top】值——\n')
            value_l1 = round(float(input('----请输入: value_l1（近端） = ')), 6)
            value_l2 = round(float(input('----请输入: value_l2（远端） = ')), 6)
            interval_l = round(float(input('----请输入: interval（L值间隔） = ')), 6)   # interval_l 表示水平坐标轴的值  递增/递减间隔
            # 因为ppt模板中坐标轴位置固定，因此起点/终点的位置其实也是固定的（下同）
            left_l1  = 118  #round(float(input('----请输入: left_l1 = ')), 6)
            left_l2  = 851  #round(float(input('----请输入: left_l2 = ')), 6)
            top_l    = 415  #round(float(input('----请输入: top_l = ')), 6)
            # 数值相对标尺的位移量 ---------------
            left_l_adj = -17
            left_t_adj = 9

            print('第三步，输入 Top 轴上任意2个点的【2个值】和【2个座标】，以及它们公共【Left】值——\n')
            value_t1 = round(float(input('----请输入: value_t1（近端） = ')), 6)
            value_t2 = round(float(input('----请输入: value_t2（远端） = ')), 6)
            interval_t = round(float(input('----请输入:  interval（T值间隔） = ')), 6)  # interval_t 表示垂直坐标轴的值  递增/递减间隔
            top_t1   = 385 #round(float(input('----请输入: top_t1 = ')), 6)
            top_t2   = 123 #round(float(input('----请输入: top_t2 = ')), 6)
            left_t   = 61  #round(float(input('----请输入: left_t = ')), 6)
            # 数值相对标尺的位移量 ---------------
            top_l_adj = -35
            top_t_adj = -10


        elif rev == -1:      # 针对Top轴向下的反向坐标轴


            print('第二步，输入 Left 轴上任意2个点的【2个值】和【2个座标】，以及它们公共【Top】值——\n')
            value_l1 = round(float(input('----请输入: value_l1（近端） = ')), 6)
            value_l2 = round(float(input('----请输入: value_l2（远端） = ')), 6)
            interval_l = round(float(input('----请输入:  interval（L值间隔） = ')), 6)  # interval_l 表示水平坐标轴的值  递增/递减间隔
            left_l1 = 117  # round(float(input('----请输入: left_l1 = ')), 6)
            left_l2 = 894  # round(float(input('----请输入: left_l2 = ')), 6)
            top_l = 80     # round(float(input('----请输入: top_l = ')), 6)
            # 数值相对标尺的位移量 ---------------
            left_l_adj = -17
            left_t_adj = -28


            print('第三步，输入 Top 轴上任意2个点的【2个值】和【2个座标】，以及它们公共【Left】值——\n')
            value_t1 = round(float(input('----请输入: value_t1（近端） = ')), 6)
            value_t2 = round(float(input('----请输入: value_t2（远端） = ')), 6)
            interval_t = round(float(input('----请输入:  interval（T值间隔） = ')), 6)  # interval_t 表示垂直坐标轴的值  递增/递减间隔
            top_t1 = 130  # round(float(input('----请输入: top_t1 = ')), 6)
            top_t2 = 385  # round(float(input('----请输入: top_t2 = ')), 6)
            left_t = 65   # round(float(input('----请输入: left_t = ')), 6)
            # 数值相对标尺的位移量 ---------------
            top_l_adj = -33
            top_t_adj = -12


        else:
            print('rev值无效输入，程序自动终止！')
            return None   # 自动终止函数





        # 【步骤二】： 然后，就能确定坐标原点，自动计算单位像素
        # 先计算 Left 方向轴上的原点和单位
        delt_l = (left_l1 - left_l2) / (value_l1 - value_l2)
        zero_l = ((left_l1 - value_l1*delt_l), top_l)   # 按 （Left , Top） 格式的Tuple


        # 再计算 Top 方向轴上的原点和单位
        delt_t = (top_t1 - top_t2) / (value_t1 - value_t2)
        zero_t = (left_t, (top_t1-value_t1*delt_t))
        #return delt_t,zero_t    #调试ok, 减震回弹的矩阵计算出来了


        # 最后，输出固定格式值：
        print('\n该矩阵坐标系【原点】和【单位】关键参数如下：\n（请将以下代码添加到make_matrix函数的dict中）\ndelt_l = '+ str(delt_l) +'\n'+'zero_l = ('+ str(zero_l[0]) +',' + str(zero_l[1]) +  ')\n' +         'delt_t = '+ str(delt_t) +'\n'+'zero_t = ('+ str(zero_t[0]) +',' + str(zero_t[1]) +  ')\n'  )




        # 【步骤三】  2023-05 新增模块： 自动生成坐标系标尺，从光秃秃的坐标轴开始 ======================================================================
        # 【3.1 先搞定横坐标轴】 —— 手工制作一个三角形 //
        mc_shape = Triangle_Shape(mc_slide,left_l1,top_l)   # 绘制一个默认的三角形，尖头朝上
        mc_shape = mc_shape.shape    # 按照我在class中的定义，mc_shap.shape = mc_slide.Shapes(1) ，因此这句代码不能少，注意！！！
        if rev == -1:
            mc_shape.Rotation = 180

        # 继续手工绘制一个 Text
        mc_text = Text_scale(Slide=mc_slide,Left=-100,Top=top_l+left_t_adj,Text=str(0))
        mc_text.tr.ParagraphFormat.Alignment = 2  # 因为和父级Class冲突，需要再运行一次，就不去修改父类了，避免节外生枝
        mc_text = mc_text.shape   # 这样赋值之后，还能调用text吗？ 我表示怀疑，保险起见，还是不折腾了，直接按照我在class中的定义//其实可以，我试过了，而且这种用法更稳定，后续可以不断复制


        value_temp = value_l1
        left_temp = left_l1
        count = abs((value_l2 - value_l1)/interval_l)  # 一共有7个interval需要填补，不过这个数据价值不大吧？ 还是有用的   // 注意，这里的count有可能是负数，因此使用abs！！
        interval_pix = abs((left_l2 - left_l1)/count)   # 每个标尺之间的间隔像素，类似坐标轴原点和【单位大小】
        sign = (value_l2-value_l1)<0 and -1 or 1   # sign 好像用不上了，先留着吧，留给循环中的value递增/递减使用


        for v in range(0,int(round(count+1,0))):   # 这个循环条件是对的，假设11-18，确实需要绘制8个图形，覆盖掉第一个shape
            # 三角形标尺
            mc_shape.Copy()
            mc_shape = mc_slide.Shapes.Paste()
            mc_shape.Top = top_l
            mc_shape.Left = left_temp



            # 坐标轴数值
            mc_text.Copy()
            mc_text = mc_slide.Shapes.Paste()  #  ！！！！  经过复制粘贴一次之后，这个Text 和 手工绘制那个Class就不一样了。不再是我定义的那个class了 // 结果，它变成了一个Shape！！！！
            if value_temp == int(value_temp):  # == round(value_temp,0):  #对应整数，如【1 / 2 / 3 】
                mc_text.TextFrame.TextRange.Text = str(int(value_temp))                      # 这里出问题了！！ bug ★ ★ ★   已修复

            elif type(value_temp) == float and value_temp >= 1: # round(value_temp,1) and value_temp > 1:  # 对应小数，如 【1.5 / 2.3】
                mc_text.TextFrame.TextRange.Text = str(round(value_temp, 1))

            elif type(value_temp) == float and value_temp < 1: #  value_temp == round(value_temp,2) and value_temp < 1:  # 对应百分比，如【55%】
                                                                # 终于找到问题根源，这句话根本没有被执行，因为 我以为 v=0.49，其实v=0.4900000000000000004 。。。  所以条件失败，
                mc_text.TextFrame.TextRange.Text = str(round(round(value_temp,2)*100, 0))[0:2] + "%"
                #print(mc_text.TextFrame.TextRange.Text)
            else:
                print('数据无法识别，出现float bug，请重新调试代码！！ 直接跳出循环吧')
                break



            mc_text.Top = top_l + left_t_adj
            mc_text.Left = left_temp + left_l_adj



            left_temp += interval_pix    #Left 恒定递增，始终是从左至右，无论rev值
            value_temp += (sign*interval_l)





       # 【3.2 再搞定纵坐标轴】 —— 手工制作一个三角形
        mc_shape = Triangle_Shape(mc_slide,left_t,top_t1)   # 绘制一个默认的三角形，尖头朝上
        mc_shape = mc_shape.shape    # 按照我在class中的定义，mc_shap.shape = mc_slide.Shapes(1) ，因此这句代码不能少，注意！！！
        mc_shape.Rotation = 90    # 无论 rev 正负， 三角形的尖头都要朝右边；

        # 继续手工绘制一个 Text
        mc_text = Text_scale(Slide=mc_slide,Left=-100,Top=top_t1+top_t_adj,Text=str(0),scale=0.04102)  # 纵轴的文本要大一些。。 奇怪，不应该呀
        mc_text = mc_text.shape   # 这样赋值之后，还能调用text吗？ 我表示怀疑，保险起见，还是不折腾了，直接按照我在class中的定义 // 还是这样赋值比较靠谱，代码稳定 而且后续可以不断复制

        value_temp = value_t1
        top_temp = top_t1

        count = abs((value_t2 - value_t1)/interval_t)  # 一共有7个interval需要填补，不过这个数据价值不大吧？ #// 注意，这里的count有可能是负数，因此使用abs！！
        interval_pix = abs((top_t2 - top_t1)/count)
        sign = (value_t2-value_t1)<0 and -1 or 1    # sign 好像用不上了，先留着吧 // 留给循环中的value递增/递减使用

        for v in range(0,int(round(count+1,0))):     # 这个循环条件是对的，假设11-18，确实需要绘制8个图形，覆盖掉第一个shape
            mc_shape.Copy()                                                                                    # 这个range 条件越来越复杂 ， 是为了应对G值坐标轴递减的情况
            mc_shape = mc_slide.Shapes.Paste()
            mc_shape.Top = top_temp
            mc_shape.Left = left_t

            # 坐标轴数值
            mc_text.Copy()
            mc_text = mc_slide.Shapes.Paste()  # 这个Text 和 手工绘制那个就不一样了。不再是我定义的那个class了吧？ 试试看？
            if value_temp == int(value_temp):  # == round(value_temp,0):  #对应整数，如【1 / 2 / 3 】
                mc_text.TextFrame.TextRange.Text = str(int(value_temp))                      # 这里出问题了！！ bug ★ ★ ★ ★    已修复

            elif type(value_temp) == float and value_temp >= 1: # round(value_temp,1) and value_temp > 1:  # 对应小数，如 【1.5 / 2.3】
                mc_text.TextFrame.TextRange.Text = str(round(value_temp, 1))

            elif type(value_temp) == float and value_temp < 1: #  value_temp == round(value_temp,2) and value_temp < 1:  # 对应百分比，如【55%】
                                                                # 终于找到问题根源，这句话根本没有被执行，因为 我以为 v=0.49，其实v=0.4900000000000000004 。。。  所以条件失败，
                mc_text.TextFrame.TextRange.Text = str(round(round(value_temp,2)*100, 0))[0:2] + "%"
                #print(mc_text.TextFrame.TextRange.Text)
            else:
                print('数据无法识别，出现float bug，请重新调试代码！！ 直接跳出循环吧')
                break

            mc_text.Top = top_temp + top_t_adj
            mc_text.Left = left_t + top_l_adj


            top_temp += (interval_pix*-1*rev)    #rev=1 top需要递减//rev=-1 top需要递增（无论value递增或者递减？是的）
            value_temp += (sign*interval_t)









    scale_matrix()









# ====================================












print('Done')








##
##if option=='3':  # 3、【全部】位置信息标注\n
##
##    # 从第一页开始，每页都复制一遍，然后用来标注位置信息
##
##    for p in range(1,len(mc_ppt.Slides)*2+1,2):
##        mc_ppt.Slides(p).Copy()
##        mc_slide = mc_ppt.Slides.Paste(p+1)
##
##
##        for i in mc_slide.Shapes:
##            Left = str(round(i.Left))
##            Top  = str(round(i.Top))
##
##            if i.TextFrame.HasText == -1:   #  【0 = 不包含文本】 / 【-1 = 包含文本】
##
##                i.TextFrame.TextRange.Text='Left = '+ Left +'\n'+'Top = '+Top
##
##            else:  # 假设是不包含文本的shape，那么需要新增一个shape，标注下相关信息
##
##                temp = i.Parent.Shapes.AddShape(5,Left=Left, Top=Top, Width=125, Height=50)
##                temp.TextFrame.TextRange.Text ='Left = '+ Left +'\n'+'Top = '+Top
##
##
##
##
##if option=='4':  # 4、【全部】大小信息标注\n
##
##    for p in range(1,len(mc_ppt.Slides)*2+1,2):
##        mc_ppt.Slides(p).Copy()
##        mc_slide = mc_ppt.Slides.Paste(p+1)
##
##        for i in mc_slide.Shapes:
##
##            Height = str(round(i.Height))
##            Width  = str(round(i.Width))
##
##            if i.TextFrame.HasText == -1:   #  【0 = 不包含文本】 / 【-1 = 包含文本】
##
##                i.TextFrame.TextRange.Text='Height = '+ Height +'\n'+'Width = '+ Width
##
##            else:  # 假设是不包含文本的shape，那么需要新增一个shape，标注下相关信息
##
##                temp = i.Parent.Shapes.AddShape(5,Left=i.Left, Top=i.Top, Width=125, Height=50)
##                temp.TextFrame.TextRange.Text ='Height = '+ Height +'\n'+'Width = '+ Width
##
