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



col_n = ['日期','项目' , '体检例数', '腔内超声检查', '图文报告', '超声检查正常(包括双胎)',
         '双胎加收', '三维', "超声检查(胎儿系统)", "胎儿心脏超声",
         '残余尿测定', "床旁彩超加收", 'B超常规检', "脏器声学造影",
         '临床操作超声引导', '弹性成像' "介入操作",]

def k_s():
    """为了生成Series."""
    # 从参数导入 col_n
    return pd.Series(name='空白', dtype='int', data=None, index=col_n)

aa = k_s()

aa['日期'] = '2023-5-22'
aa['项目'] = '[泌尿系彩超,三维]'
print(aa)

'''


aa = '我的三维三维你好'

if aa.count('三维'):
    print('找到')
else:
    print('无')

jishu = pd.Series(name='单个', dtype=float, data=None, index=['二维','三维'])
jishu['三维'] = 1
jishu['二维'] = 1
jishu.rename('rank',inplace=True)
print(jishu)

aa = "[膀胱残余尿测定彩超,二维]"
print(aa)


if re.match(r'^\[膀胱残余尿(.*?)$', aa):
    print('正确')
else:
    print('错误')

print('hell')
ab = ['[肝胆脾胰彩超,二维]','[泌尿系彩超,三维]','[胎儿中晚期妊娠彩超,二维]','[泌尿系及膀胱残余尿测定彩超,二维]',
'[甲状腺及颈部淋巴结彩超,三维]', '[乳腺及腋下淋巴结彩超,三维]']



for i in ab:
    if re.match(r'^\[(.*?),二维\]$', i):
        print(i,'二维彩超')
    elif re.match(r'^\[(.*?),三维\]$', i):
        print(i,'三维彩超')
    elif i.count('床旁'):
        print(i,'床旁加收')
    else:
        print('wrong')



def funnn():
    aaaa = [random.randint(11, 20) for i in range(5)]
    return aaaa

a1 = funnn()

a2 = funnn()

def duoduo(xx):
    gengduo = [funnn() for i in range(xx)]
    return gengduo

print(a1)
print(a2)
print(duoduo(6))

pyau.screenshot('shot1.png',region=(240,400,620,30))
time.sleep(0.75)

pyau.screenshot('shot2.png',region=(240,400,620,30))

for i in range(1):
    localtime = time.localtime(time.time())

    print(localtime)
    time.sleep(2)

'''