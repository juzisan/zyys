import pandas as pd
import numpy as np
import httpx
import re


'''
'''
r = httpx.get('https://www.baidu.com/')
print(r)



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

ab = ['三维', '三维经阴道','经阴道三维','床旁','经阴道二维']

for i in ab:
    if i.count('三维'):
        print(i,'三维彩超')
    elif i.count('经阴道'):
        print(i,'经阴道彩超')
    elif i.count('床旁'):
        print(i,'床旁加收')
    else:
        print('wrong')
