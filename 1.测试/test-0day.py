import pandas as pd
import numpy as np
from DrissionPage import ChromiumPage

from DrissionPage.easy_set import set_paths

# set_paths(browser_path=r'F:\Program Files\chrome-win')


# page = ChromiumPage()
# page.get('http://g1879.gitee.io/DrissionPageDocs')


p1 = -1
p2 = [-1, 0, 2]
for i in p2:
    if i:
        print(i, '没找到')
    else:
        print(i, '找到了')
print(f'{p2 = }')

d_2 = pd.DataFrame(np.random.randn(6, 3), ['班级', '人数', '姓名', '平均分', '总分', '最高分'],
                   columns=['ab', 'bc', 'cd'])
d_2 = (d_2 * 100).round(0)
print(d_2)
d_2.loc[d_2['ab'] < 60, 'ab'] = 0
d_2['行总和'] = d_2.sum(axis=1)
print(d_2)
s_1 = pd.Series(dtype=int, name='张三')
s_1['一'] = 13
s_1['二'] = 29
s_1['三'] = 37
# s_1.append([7])
print(f'{s_1 = }')
print('\033[1;32;40m start \033[0m')
'''
'''
data = [[1, 2, 3], [1, 5, 6], [7, 8, 9]]
df = pd.DataFrame(data, columns=["a", "b", "c"],
                  index=["owl", "toucan", "eagle"])
print(df)
try:
    aa = df.groupby(by=["a"]).get_group(5)
except KeyError:
    aa = 'kong'

print(aa)
