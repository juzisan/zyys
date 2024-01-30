"""Created on Tues Jan  30 17:28:40 2024.

@author: Mrlily
统计工作量
区分统计体检和门诊住院
体检都是二维，多个项目中间用 【+】 号分割
门诊住院区分二维三维，有些项目特殊，无法统计往诊
"""
import os
import re
import time
import numpy as np
import pandas as pd
from argument import (col_n, NAME_ITEM_OLD, NAME_ITEM_OLD_REPLACE, )

global yue, NAME_RESIDUAL_URINE,  NAME_ITEM_OLD, NAME_ITEM_OLD_REPLACE

NAME_RESIDUAL_URINE = []  # 统计残余尿项目名称


def blank_series():
    """为了生成Series."""
    # 从参数导入 col_n
    return pd.Series(name='空白', dtype='int', data=None, index=col_n)


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


def one_do(txt_str, classify_person):
    """逻辑处理部分."""
    count_series = blank_series()  # 复制Series
    if classify_person == '门住':
        # 图文报告
        count_series['图文报告'] = 1
        if re.match(r'^\[(.*?),三维]', txt_str):
            count_series['三维'] = 1
            if txt_str.count(r'胎'):
                count_series['疑难病例会诊例数'] = 1
                # 胎儿三维再另加1个在疑难病例会诊例数里面
            # 残尿三维不以三维结尾
        elif re.match(r'^\[(.*?),二维]', txt_str):
            if txt_str.count(r'卵泡测定'):
                # print(txt_str)
                count_series['超声检查正常(包括双胎)'] = 1
            elif txt_str.count(r'经阴道'):
                count_series['腔内超声检查'] = 1
            elif txt_str.count(r'一个部位'):
                count_series['超声检查正常(包括双胎)'] = 1
                count_series['床旁彩超加收'] = 1
            elif txt_str.count(r'二个部位'):
                count_series['超声检查正常(包括双胎)'] = 2
                count_series['床旁彩超加收'] = 1
            elif txt_str.count(r'三个部位'):
                count_series['超声检查正常(包括双胎)'] = 3
                count_series['床旁彩超加收'] = 1
            elif txt_str.count(r'四个部位'):
                count_series['超声检查正常(包括双胎)'] = 4
                count_series['床旁彩超加收'] = 1
            elif txt_str.count(r'五个部位'):
                count_series['超声检查正常(包括双胎)'] = 5
                count_series['床旁彩超加收'] = 1
            else:
                count_series['超声检查正常(包括双胎)'] = 1
        else:
            print('缺少二维三维：', txt_str)
            count_series['超声检查正常(包括双胎)'] = 1

        # 残余尿
        if txt_str.count(r'残余'):
            count_series['残余尿测定'] = 1
            NAME_RESIDUAL_URINE.append(txt_str)
        if re.match(r'^\[膀胱残余尿(.*?)$', txt_str):
            count_series['超声检查正常(包括双胎)'] = 0
            # count_series['脏器灰阶立体成象'] = 0
    elif classify_person == '体检':
        count_series['体检例数'] = txt_str.count(r'+') + 1
    else:
        print('错误 ：患者类型')
    return count_series


def do_it(file_str):
    """程序开始."""
    global yue, NAME_RESIDUAL_URINE, NAME_ITEM_OLD_old
    ori_df = pd.read_excel(io=file_str, sheet_name=0, parse_dates=['检查时间'])
    ori_df = ori_df[['检查时间', '患者类型', '检查部位,检查方法', ]]
    ori_df = ori_df.sort_values(by=['检查时间'], ascending=True)
    ori_df.rename(columns={'检查部位,检查方法': '检查部位'}, inplace=True)  # 把列重命名
    ori_df['患者类型'] = ori_df['患者类型'].replace({'住院': '门住', '门诊': '门住'})
    # 把患者类型列的值重新替换一下
    ori_df['检查时间'] = ori_df['检查时间'].dt.day  # 检查时间只保留 day
    group_day = ori_df.groupby(['检查时间'])  # 按照检查时间分组
    name_item_new = ori_df[ori_df["患者类型"] == "门住"]

    name_item_new = name_item_new["检查部位"].str.rstrip("彩超,二三维])胰脾早中晚期妊娠（右）左")
    # NAME_ITEM_OLD = NAME_ITEM_OLD["检查部位"].str.rstrip(",二维")subset=["检查部位"],

    name_item_new = name_item_new.str.lstrip("[彩超(")
    name_item_new.replace(NAME_ITEM_OLD_REPLACE, inplace=True)
    NAME_ITEM_OLD.drop_duplicates(keep="first", inplace=True)
    NAME_ITEM_OLD = NAME_ITEM_OLD.sort_values(ascending=True)
    NAME_ITEM_OLD = NAME_ITEM_OLD.to_list()

    NAME_ITEM_OLD_old = list(filter(None, NAME_ITEM_OLD_old.split('\n')))
    # 删除list中的空元素,先用list转，否则不能正常比较
    # print(NAME_ITEM_OLD,'\n',NAME_ITEM_OLD_old)
    new_old = set(NAME_ITEM_OLD).difference(set(NAME_ITEM_OLD_old))
    print('新加：\n')
    print(*new_old, sep="\n")  # 显示没有的检查项目
    print('\n----------\n')
    print('旧多：\n')
    old_new = set(NAME_ITEM_OLD_old).difference(set(NAME_ITEM_OLD))
    print(*old_new, sep="\n")  # 多余的检查项目
    print('\n----------\n')

    zonghe = []  # 准备列表生成统计
    for i in range(1, 32):
        try:
            jisruan = group_day.get_group(i).apply(lambda row: one_do(row['检查部位'], row['患者类型']), axis=1)
            # groupby 对象需要用 get_group 才能调用,df用apply传递多个参数的时候要用lambda
        except KeyError:
            count_series = blank_series()
            print(i, '日  没上班')
        else:
            count_series = jisruan.sum()
        count_series.rename(i, inplace=True)  # 对 Series 重命名
        zonghe.append(count_series)

    zonghe = pd.concat(zonghe, axis=1).T  # 合并表，转置表
    zonghe.loc['总和'] = zonghe.apply(lambda x: x.sum())  # 各列求和，添加新的行
    print('\n----------\n')
    print(list(set(NAME_RESIDUAL_URINE)))
    zonghe.replace(0, np.nan, inplace=True)
    print(zonghe)
    zonghe.to_excel('统计好' + yue + '月.xlsx')


@timer
def main():
    """程序开始."""
    global yue
    names = os.listdir(os.path.split(os.path.realpath(__file__))[0])
    names = [i for i in names if re.match(r'(.*?)\.(xls|xlsx)$', i)]
    if names:
        if len(names) > 1:
            print('请输入序号：')
            for i, value in enumerate(names):
                print(i, '代表：  ', value)
            select_num = int(input("输入转换的序号："))
            file_n = names[select_num]
        else:
            file_n = names[0]
        yue = re.search(r'\d+', file_n)[0]
        do_it(names[0])
    else:
        print('缺少文件')


if __name__ == "__main__":
    main()
    print('all ok')

'''



'''
