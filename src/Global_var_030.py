# global 全局变量模块




## 【真-全局变量  真-全局变量  真-全局变量  真-全局变量  真-全局变量  真-全局变量  】
## ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★ ★
##
## ========== 不知是否 working.... //  还是先按我自己的思路吧。。。 结果思路相同。。。
##                # 无法working也就算了，但是这样定义的函数，导致整个文件结构无法running。。。。

#global global_dict

# 这个全局变量字典中，目前只有一个变量：dic_matrix，以后只修改这一个文件就行  // 后续有需要，再按照字典的格式添加吧
global_dict = {
    'dic_matrix' : {
    '减震回弹矩阵-篮球': 5 , \
    '弯折扭转矩阵-篮球': 6 , \
    '减震回弹矩阵-跑步': 7 , \
    '弯折扭转矩阵-跑步': 8 , \
    '步长重心振幅矩阵-跑步':12 , \
     '重量厚度矩阵-跑步': 13     #注意，从12页开始，不要再往中间插入ppt了，直接在后面新增，不然main函数的代码（页码）都需要调整
    }




}

# def _init():     # 说实话，我没看懂。。。 // 改天再深究
#     global _global_dict
#
#     _global_dict = {}




def get_value(key):

    global global_dict
    try:
        return global_dict[key]
    except KeyError:
        print('全局变量读取失败，请检查 Global_var.py 文件！！')
        return None  # defValue







# def get_value(key,defValue=None):
#     global global_dict
#
#     try:
#         return global_dict[key]
#     except KeyError:
#         return defValue


print('Global_var_030 文件导入完成！\n')   # 原来 import 会将整个程序 run 一遍 //  学习了
