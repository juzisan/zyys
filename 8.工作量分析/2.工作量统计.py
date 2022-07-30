"""Created on Fri Jul  8 17:24:40 2022

@author: Mrlily
统计工作量
区分统计体检和门诊住院
体检都是二维，多个项目中间用 【+】 号分割
门诊住院区分二维三维，有些项目特殊，无法统计往诊
"""
import os
import re
import time

# import numpy as np
import pandas as pd


def timer(func):
    """ 计时器. """

    def warpper():
        """计时器内部."""
        print('\033[1;32;40mstart\033[0m')
        time1 = time.time()
        func()
        seconds = time.time() - time1
        m, s = divmod(seconds, 60)
        print("\033[1;32;40mthe run time is %02d:%.6f\033[0m" % (m, s))

    return warpper


def do_it(file_str):
    """程序开始."""
    # pd.set_option('display.max_rows', 100)
    yue = int(re.search(r'\d+', file_str)[0])
    ori_df = pd.read_excel(io=file_str, sheet_name=0, parse_dates=['检查时间'])
    # 读 excel '总分':'int'
    ori_df = ori_df[['检查时间', '患者类型', '检查部位,检查方法', '报告医生']]
    ori_df = ori_df.sort_values(by=['检查时间'], ascending=True)
    ori_df.rename(columns={'检查部位,检查方法': '检查部位'}, inplace=True)  # 把列重命名
    ori_df = ori_df[ori_df['检查时间'].dt.month.isin([yue])]  # 按月筛选

    tj_df = ori_df[ori_df['患者类型'] == '体检'].copy(deep=True)
    show_dict = {}
    tj_df.index = range(1, len(tj_df) + 1)  # 重建索引

    tj_exam_set = tj_df['检查部位'].unique()
    print(f'{tj_exam_set =}')
    tj_df['加号'] = tj_df['检查部位'].str.count(r'\+')
    print(tj_df)
    tj_sum = tj_df['加号'].sum() + tj_df['加号'].count()
    show_dict['体检'] = tj_sum
    print(f'{tj_sum  = }\n\n')

    mz_df = ori_df[ori_df['患者类型'] != '体检'].copy(deep=True)
    mz_df.index = range(1, len(mz_df) + 1)  # 重建索引
    mz_df['检查部位'].replace(r"\[|\]", '', regex=True, inplace=True)
    show_dict['报告数'] = mz_df.count()['报告医生']

    count_dict = {r'三维': r'三维', r'双胎': r'双胎', r'残余尿': r'残余尿',
                  r'55残余尿': r'^膀胱残余尿测定', r'经阴道': r'经阴道(?:.*)二维'}
    for key, value in count_dict.items():
        mz_df[key] = mz_df['检查部位'].str.contains(value, na=False, regex=True)
        mz_df[key] = mz_df[key].astype('int')
        show_dict[key] = mz_df[key].sum()
    print(mz_df)
    mz_df['二维'] = 1 - mz_df['三维']
    mz_df['二维'] = mz_df['二维'] - mz_df['经阴道']
    mz_df['二维'] = mz_df['二维'] - mz_df['55残余尿']
    print(mz_df[mz_df['二维'] < 0])
    # mz_df.loc[mz_df['二维'] < 0, '二维'] = 0
    show_dict['二维'] = mz_df['二维'].sum()
    mz_exam_set = mz_df['检查部位'].unique()

    print(f'{mz_exam_set  = }\n\n')
    print(f'{show_dict = }')


@timer
def main():
    """程序开始."""
    names = os.listdir(os.path.split(os.path.realpath(__file__))[0])
    names = [i for i in names if re.match(r'(.*?)\.(xls|xlsx)$', i)]
    print(names)
    if names:
        if len(names) > 1:
            print('请输入序号：')
            for i, value in enumerate(names):
                print(i, '代表：  ', value)
            select_num = int(input("输入转换的序号："))
            do_it(names[select_num])
        else:
            do_it(names[0])
    else:
        print('缺少文件')


if __name__ == "__main__":
    main()
    print('all ok')

'''



'''