# import numpy as np
import pandas as pd
# import pyautogui


def blank_series():
    """为了生成Series，默认名字，省的重命名."""
    col_n = ['体检例数', '超声检查正常', '脏器灰阶成像*2/3', '残余尿测定', '床旁彩超加收*5',
             '腔内超声检查', '大排畸*6', '胎儿心脏超声*6', '脏器灰阶成像（NT+产科）', 
             '双胎加收*3', '胃肠超声*3', '脑黑质测定*2', '弹性成像', '门住体图文报告',
             '脏器声学造影*5', '临床操作超声引导*5', '介入操作例数*10', '消融例数*20', 
             '误时工作量例数', '来源住培学员', '疑难病例会诊例数', '夜班例数', '扣罚金额']
    return pd.Series(name='空白', dtype='int', data=None, index=col_n, )


a1 = blank_series()
a1['超声检查正常'] =1
a2 = blank_series()
a2['脏器灰阶成像*2/3'] =1
a3 = blank_series()
a3['体检例数'] =1
print(a1,a2)

lie_name_str='''
门住体图文报告	体检例数	肝纤维化和肝脂肪变测定*2	超声检查正常	脏器灰阶成像*2/3	残余尿测定	床旁彩超加收*5	腔内超声检查	"大排畸
*6"	"胎儿心脏超声
*6"	脏器灰阶成像（NT+产科）	双胎加收*3	胃肠超声*3	"脑黑质测定
*2"	脏器声学造影*5	"临床操作超声引导
*5"	"介入操作例数
*10"	消融例数*20	误时工作量例数	来源住培学员	疑难病例会诊例数	夜班例数	扣罚金额

'''
# print(lie_name_str)
# print(lie_name_str.translate(str.maketrans({'"': None, '\n': None})))
lie_name_str = lie_name_str.translate(str.maketrans({'"': None, '\n': None}))
lie_name_list = lie_name_str.split('\t')
lie_name_list.insert(0,'检查时间')
print(lie_name_list)
# lie_name_str.replace('"', '')
# print(lie_name_str.replace('\r', ''))


# get_name = '门住体报告'
def if_name(get_name):
    if get_name not in lie_name_list: print('缺少：' + get_name)

if_name('脏灰阶成像*2/3')

li = ['床旁彩超(二个部位),二维','床旁彩超（三个部位）,二维','床旁彩超(一个部位),二维','床旁彩超,二维']

for i in li:
    txt_str = i

    if  txt_str.count(r'三个部位'):
        print('三个')
    elif txt_str.count(r'一个部位'):
        print('一个')
    elif txt_str.count(r'床旁彩超') and not txt_str.count(r'个'):
       # if txt_str.find(r'个') == -1:
            print('床旁彩超')

