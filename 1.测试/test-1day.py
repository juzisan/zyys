import pandas as pd
import numpy as np
import httpx



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

aa = "[彩超（肝胆脾胰+子宫附件+乳腺）]"
print(aa)
print(aa.count(r'+') + 1)
