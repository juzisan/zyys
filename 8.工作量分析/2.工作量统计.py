"""Created on Fri Jul  8 17:24:40 2022

@author: Mrlily
统计工作量
区分统计体检和门诊住院
体检都是二维，多个项目中间用 【+】 号分割
门诊住院区分二维三维，有些项目特殊，无法统计往诊
"""

import pandas as pd
import time


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


@timer
def main():
    """程序开始."""
    pd.set_option('display.max_rows', 100)
    original_dataframe = pd.read_excel('chengr.xls', parse_dates=[
        '检查时间'])  # 读 excel '总分':'int'
    original_dataframe = original_dataframe[[
        '检查时间', '患者类型', '检查部位,检查方法', '报告医生']]
    original_dataframe = original_dataframe.sort_values(
        by=['检查时间'], ascending=True)
    original_dataframe.rename(columns={'检查部位,检查方法': '检查部位'}, inplace=True)  # 把列重命名
    original_dataframe = original_dataframe[original_dataframe['检查时间'].dt.month.isin([6])]  # 按月筛选
    print(f'{original_dataframe  = }\n\n')
    analyze_lst = [['体检', original_dataframe[original_dataframe['患者类型'] == '体检']],
                   ['门诊住院', original_dataframe[original_dataframe['患者类型'] != '体检']]]
    print(f'{original_dataframe  = }\n\n')
    for typ, value in analyze_lst:
        value.index = range(1, len(value) + 1)  # 重建索引
        exam_set = value['检查部位'].unique()
        if typ == '体检':
            jiahao = value['检查部位'].apply(lambda x: x.count(r'+'))
            print(type(jiahao))
            value.insert(value.shape[1], '加号', jiahao)
            tijian_sum = value['加号'].sum() + value['加号'].count()
            print(f'{tijian_sum  = }\n\n')
        else:
            count_lst = [r'三维', r'双胎', r'残余尿', r'经阴道']
            for i in count_lst:
                value.insert(value.shape[1], i, value['检查部位'].apply(lambda x: bool(x.count(i))))
                if i == r'经阴道':
                    zhong_jian = value[i] & ~ value['三维']
                    del value[i]
                    value.insert(value.shape[1], i, zhong_jian)
                er_wei_sum = value[i].value_counts()
                print(f'{er_wei_sum}\n\n')
        print(f'{typ  = }\n\n')
        print(f'{value  = }\n\n')
        print(f'{exam_set  = }\n\n')


if __name__ == "__main__":
    main()
    print('all ok')

'''



'''
