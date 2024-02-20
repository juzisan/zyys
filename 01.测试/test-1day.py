import pandas as pd
import numpy as np
import httpx
import re
import random
import pyautogui as pyau
import time

'''
'''
r = httpx.get('https://www.baidu.com/')
print(r)

col_n = ['日期', '项目', '体检例数', '腔内超声检查', '图文报告', '超声检查正常(包括双胎)',
         '双胎加收', '三维', "超声检查(胎儿系统)", "胎儿心脏超声",
         '残余尿测定', "床旁彩超加收", 'B超常规检', "脏器声学造影",
         '临床操作超声引导', '弹性成像' "介入操作", ]


def blank_series(s_name='空白'):
    """为了生成Series."""
    # 从参数导入 col_n
    return pd.Series(name=s_name, dtype='int', data=None, index=col_n)


aa = blank_series()
print(aa)

def one_day(day_num):
    """为了生成Series."""
    try:
        # del_count = group_day.get_group(day_num).apply(lambda row: one_do(row['检查部位'], row['患者类型']), axis=1)
        # groupby 对象需要用 get_group 才能调用,df用apply传递多个参数的时候要用lambda
        day_num * 2
    except KeyError:
        sum_series = blank_series()
        print(day_num, '日  没上班')
    else:
        # sum_series = del_count.sum()
        # sum_series.rename(day_num, inplace=True)  # 对 Series 重命名
        sum_series = day_num * 2
    sum_series = day_num * 2
    return sum_series


cc = map(one_day, range(1, 31))

for i in cc:
    print(i)