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

see_what2 = []
for i in range(6):
    see_what = round(pd.Series(np.random.random(6), ['班级', '人数', '姓名', '平均分', '总分', '最高分']) * 100, 0)
    see_what2.append(see_what)

d_1 = pd.concat(see_what2, axis=1)
print(d_1)

d_2 = pd.DataFrame(np.random.randn(6, 3), ['班级', '人数', '姓名', '平均分', '总分', '最高分'], columns=['ab', 'bc', 'cd'])
d_2 = (d_2 * 100).round(0)
see_what = round(pd.Series(np.random.random(3), ['ab', 'bc', 'cd']) * 100, 0)
print(see_what)
print(d_2)
d_3 = pd.DataFrame(see_what).T

d_2 = pd.concat([d_2, d_3])
print(d_2)

d_4 = pd.read_excel(io="小分表 - 哪个学校八年级期末考试-数学.xlsx", sheet_name=0, skiprows=1, dtype={'班级': 'str'})
print(d_4)
print(d_4['姓名'])
