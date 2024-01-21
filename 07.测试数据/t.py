import os
import re
import pandas as pd
import numpy as np
from requests_html import HTMLSession
import random
import scipy
import time
import datetime

start_time = time.time()


def timer(func):  # 计时器
    def warpper():
        print('\033[1;32;40mstart\033[0m')
        time1 = time.time()
        func()
        seconds = time.time() - time1
        m, s = divmod(seconds, 60)
        print("\033[1;32;40mthe run time is %02d:%.6f\033[0m" % (m, s))
    return warpper


@timer
def main():
    print(''.join('zzzx'))

    values0 = [90, 95, 100, 105, 110, 115, 120, 125, 130, 135, 140, 145, 150]

    values = ['red', 'green', 'magenta', 'chocolate', 'brown', 'deeppink', r'#000080']

    df2 = pd.DataFrame(np.random.randint(0, 100, size=(10, 3)))  # 生成50行3列的dataframe

    zitidaxiao = random.sample(values0 * df2[2].count(), df2[2].count())
    yanse = random.sample(values * df2[2].count(), df2[2].count())

    df2['字体'] = zitidaxiao
    df2['颜色'] = yanse
    print(df2)

    aa = np.random.normal(loc=1.9, scale=1.2, size=10)*1.9
    print(aa)
    df2_size = 10  # 千
    #  df2_size = 1000  # 千
    #  df2_size = 1000000  # 百万
    df2 = pd.DataFrame(np.arange(df2_size))
    df2['真假0'] = np.random.choice([True, False], size=df2_size, replace=True, p=[0.2, 0.8])
    df2['真假1'] = np.random.choice([True, False], size=df2_size, replace=True, p=[0.1, 0.9])
    df2['真假2'] = np.random.choice([True, False], size=df2_size, replace=True, p=[0.27, 0.73])
    df2['真假3'] = np.random.choice([True, False], size=df2_size, replace=True, p=[0.65, 0.35])
    df2['真假4'] = np.random.choice([True, False], size=df2_size, replace=True, p=[0.28, 0.72])
    print(df2, df2['真假0'].describe(), df2['真假1'].describe(), df2['真假2'].describe(), df2['真假3'].describe(), df2['真假4'].describe())

    df2.to_csv("nihao.csv")


print('start')
if __name__ == "__main__":
    main()
