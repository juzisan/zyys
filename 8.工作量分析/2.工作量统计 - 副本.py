"""Created on Fri Jul  8 17:24:40 2022.

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

global doct_df, exam_df, ori_df, col_n


def timer(func):
    """计时器."""

    def warpper():
        """计时器内部."""
        print('\033[1;32;40mstart\033[0m')
        time1 = time.time()
        func()
        seconds = time.time() - time1
        m, s = divmod(seconds, 60)
        print("\033[1;32;40mthe run time is %02d:%.6f\033[0m" % (m, s))

    return warpper


def one_do(p_name):
    """按医生统计."""
    global doct_df, exam_df, ori_df, col_n
    show_s = pd.Series(name='个数', dtype=float, data=None, index=col_n)

    if doct_df.loc[p_name, '在职']:
        p_df = ori_df[ori_df['报告医生'] == p_name].copy(deep=True)
        if p_df.empty:
            print(p_name, '没上班')
            pass
        else:
            tj_df = p_df[p_df['患者类型'] == '体检'].copy(deep=True)
            tj_df.index = range(1, len(tj_df) + 1)  # 重建索引
            # tj_exam_set = tj_df['检查部位'].unique()
            # print(f'{tj_exam_set =}')
            tj_df['加号'] = tj_df['检查部位'].str.count(r'\+')
            # print(tj_df)
            tj_sum = tj_df['加号'].sum() + tj_df['加号'].count()
            # print(f'{tj_sum  = }\n\n')
            mz_df = p_df[p_df['患者类型'] != '体检'].copy(deep=True)
            mz_df.index = range(1, len(mz_df) + 1)  # 重建索引

            for row in exam_df.itertuples():
                row_e = row.Index
                row_s = getattr(row, '搜索字段')
                mz_df[row_e] = mz_df['检查部位'].str.contains(row_s, na=False, regex=True)
                mz_df[row_e] = mz_df[row_e].astype('int')
                show_s[row_e] = mz_df[row_e].sum()
            mz_df['55残余尿'] = mz_df['检查部位'].str.contains(r'^膀胱残余尿测定', na=False, regex=True)
            show_s['图文报告'] = mz_df.count()['报告医生']
            mz_df['超声检查正常(包括双胎)'] = 1 - mz_df['三维']
            mz_df['超声检查正常(包括双胎)'] = mz_df['超声检查正常(包括双胎)'] - mz_df['腔内超声检查']
            mz_df['超声检查正常(包括双胎)'] = mz_df['超声检查正常(包括双胎)'] - mz_df['55残余尿']
            mz_df['超声检查正常(包括双胎)'] = mz_df['超声检查正常(包括双胎)'] - mz_df['胎儿心脏超声']
            mz_df['超声检查正常(包括双胎)'] = mz_df['超声检查正常(包括双胎)'] - mz_df['脏器声学造影']
            mz_df['超声检查正常(包括双胎)'] = mz_df['超声检查正常(包括双胎)'] - mz_df['临床操作超声引导']
            mz_df['超声检查正常(包括双胎)'] = mz_df['超声检查正常(包括双胎)'] - mz_df['弹性成像']

            show_s['体检例数'] = tj_sum
            if not mz_df[mz_df['超声检查正常(包括双胎)'] < 0].empty:
                print(p_name, '有错')
                mz_df.loc[mz_df['超声检查正常(包括双胎)'] < 0, '超声检查正常(包括双胎)'] = 0
            show_s['超声检查正常(包括双胎)'] = mz_df['超声检查正常(包括双胎)'].sum()
            income_num = show_s * exam_df['价格']
            show_s['工作量总计'] = income_num.sum(skipna=True)
            # print(f'{mz_exam_set  = }\n\n')
            # print(f'{show_s = }')
    else:
        pass
    show_s['报告医生'] = p_name
    # print(show_s)
    return show_s


def do_it(file_str):
    """程序开始."""
    global doct_df, exam_df, ori_df, col_n
    doct_df = pd.read_excel(io='科室信息.xlsx', sheet_name='报告医生', index_col=1, dtype={'在职': bool})
    exam_df = pd.read_excel(io='科室信息.xlsx', sheet_name='检查项目信息', index_col=1)
    print(exam_df)
    col_n = exam_df.index.tolist()
    print(f'{col_n = }')
    # pd.set_option('display.max_rows', 100)
    yue = int(re.search(r'\d+', file_str)[0])
    ori_df = pd.read_excel(io=file_str, sheet_name=0, parse_dates=['检查时间'])
    # 读 excel '总分':'int'
    ori_df = ori_df[['检查时间', '患者类型', '检查部位,检查方法', '报告医生']]
    ori_df = ori_df.sort_values(by=['检查时间'], ascending=True)
    ori_df.rename(columns={'检查部位,检查方法': '检查部位'}, inplace=True)  # 把列重命名
    ori_df['检查部位'].replace(r"\[|\]", '', regex=True, inplace=True)
    ori_df = ori_df[ori_df['检查时间'].dt.month.isin([yue])]  # 按月筛选
    xm_df = ori_df[ori_df['患者类型'] != '体检'].copy(deep=True)
    xm_set = list(xm_df['检查部位'].unique())
    xm_set.sort()
    print(xm_set)
    all_name = doct_df.index.tolist()
    print(all_name)
    a_list = [one_do(i) for i in all_name]
    d_df = pd.concat(a_list, axis=1).T
    d_df.set_index('报告医生', inplace=True)
    print(d_df)
    d_df.to_excel(str(yue)+'统计.xlsx')


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