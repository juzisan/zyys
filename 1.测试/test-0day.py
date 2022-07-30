import pandas as pd
import numpy as np

p1 = -1
p2 = [-1, 0, 2]
for i in p2:
    if i:
        print(i, '没找到')
    else:
        print(i, '找到了')
print(f'{p2 = }')


d_2 = pd.DataFrame(np.random.randn(6, 3), ['班级', '人数', '姓名', '平均分', '总分', '最高分'], columns=['ab', 'bc', 'cd'])
d_2 = (d_2 * 100).round(0)
print(d_2)
d_2.loc[d_2['ab'] < 60, 'ab'] = '不及格'
print(d_2)
'''
'''