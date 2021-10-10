import os
import re
import pandas as pd
import numpy as np
from requests_html import HTMLSession
import random
import scipy





values0 =[90,95,100,105,110,115,120,125,130,135,140,145,150]
values = ['red','green','magenta','chocolate','brown','deeppink',r'#000080']

df2 = pd.DataFrame(np.random.randint(0,100,size=(10, 3))) # 生成50行3列的dataframe

zitidaxiao = random.sample(values0 * df2[2].count(), df2[2].count())
yanse = random.sample(values * df2[2].count(), df2[2].count())

df2['字体'] = zitidaxiao
df2['颜色'] = yanse
print (df2)

aa = np.random.normal(loc=1.9, scale=1.2, size=375)*1.9
print (aa)