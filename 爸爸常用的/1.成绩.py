"""Created on Sat Oct 19 19:52:49 2019.

@author: Mrlily
"""

import os
import re
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Border, Side, Font
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


def duo(xlsx_nam, data_xlsx):
    """重复部分."""
    xlsx_nam = xlsx_nam + '班'
    # 文件名

    wb = Workbook()  # 建立excel文件
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))  # 创建边框对象

    ws = wb.create_sheet(xlsx_nam)  # 创建工作簿名
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # 页面方向
    ws.page_setup.paperSize = ws.PAPERSIZE_A4  # 纸张大小
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0
    ws.page_margins.footer = 0  # 设置页边距
    ws.print_title_rows = '1:2'  # 设置打印标题为2行
    rows = dataframe_to_rows(data_xlsx)  # pandas 的dataframe 转openpyxl 的格式
    for row in rows:  # 不知道为啥有空行。openpyxl的序号从1开始。delete_rows()
        if len(row) > 2:
            ws.append(row)
    ws.column_dimensions['A'].width = 3  # 设置列宽
    ws.column_dimensions['B'].width = 9
    for i in range(3, 12):  # 最大列数目，是整数类型
        ws.column_dimensions[get_column_letter(i)].width = 4
    for i in range(12, ws.max_column):  # 最大列数目，是整数类型
        ws.column_dimensions[get_column_letter(i)].width = 5
    ws.column_dimensions[get_column_letter(ws.max_column)].width = 6
    for i in list(ws.rows)[0]:  # 生成器转list才能迭代，第一行
        i.alignment = Alignment(wrap_text=True)  # 设置文本自动换行
    for row in ws.rows:
        for cell in row:
            cell.border = thin_border  # 设置边框

    ws.insert_rows(1)  # 插入标题行
    ws.merge_cells(start_row=1, start_column=1, end_row=1,
                   end_column=ws.max_column)  # 合并标题行
    p_title = xlsx_nam
    ws['A1'] = p_title
    ws['A1'].alignment = Alignment(horizontal='center')  # 设置标题居中
    ws['A1'].font = Font(size=20)  # 设置字号大小

    del wb['Sheet']  # 删除默认sheet，保存文件
    wb.save('试卷分析' + xlsx_nam + '.xlsx')  # excel写入sheet


def chuli(xlsx_nam_chuli):
    """共有."""
    table_head_values = pd.read_excel(xlsx_nam_chuli, skiprows=1, dtype={
        '班级': 'str', })  # 读 excel '总分':'int'
    table_head_values = table_head_values.sort_values(by=['总分'], ascending=False)
    table_head_values.index = range(1, len(table_head_values) + 1)  # 重建索引
    bb = table_head_values.tail(1).applymap(lambda x: int(
        re.search(r'得分/共(\d+)分', x).group(1)), na_action='ignore')

    # table_head_values.loc[table_head_values.index[-1]].apply(lambda x : 6 if x.empty  else x )
    table_head_values.drop(table_head_values.tail(1).index, inplace=True)
    table_head_values = pd.concat([table_head_values, bb], ignore_index=True)
    table_head_values.rename(columns=column_dict, inplace=True)
    table_head_values_lst = table_head_values.columns.values.tolist()  # pandas列名转list
    # 共有

    ti_list = []
    # print (table_head_values_lst)
    big_problem = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', ]
    for k in big_problem:
        for j in table_head_values_lst:
            if j.find(k) > -1:
                ti_list.append(j)

    # print (ti_list)

    each_problem_value = table_head_values.tail(1)

    problem_value = each_problem_value[ti_list].T.to_dict(orient='dict')
    problem_value = problem_value.popitem()[1]
    sum_values = sum(problem_value.values())

    print('题值：', problem_value, "\n总分：", sum_values)
    # problem_value = problem_value[:1]

    table_head_values.drop(table_head_values.tail(1).index, inplace=True)

    del_column = []
    for i in table_head_values_lst:
        # 目的是删除多余的列。生成每题分值dict
        if i.find('.') > 0:
            pass
        elif i not in ['姓名', '班级', '总分']:
            del_column.append(i)
    for i in del_column:
        table_head_values_lst.remove(i)
    table_head_values = table_head_values[table_head_values_lst]  # 去掉无用列
    table_head_values.drop(table_head_values.tail(1).index, inplace=True)
    table_head_values["合计"] = table_head_values["总分"].astype('float') - sum_values  # 处理扣总分
    del table_head_values['总分']  # 删除总分列
    #  table_head_values_lst = table_head_values.columns.values.tolist()

    for j in problem_value:  # 变成负分.astype('float')
        table_head_values[j] = table_head_values[j].astype('float') - problem_value[j]
        table_head_values[j] = table_head_values[j].replace(0, np.nan)

    # print(table_head_values.dtypes)
    table_head_values = table_head_values.groupby(['班级'])  # 按照班级分类
    # print (table_head_values)
    combined_table = []
    for key, value in table_head_values:
        table_head_values_lst = value.columns.values.tolist()
        table_head_values_lst.remove('班级')
        value = value[table_head_values_lst]
        value.index = range(1, len(value) + 1)
        mean_mean = list(value.mean(numeric_only=True))  # pandas的mean出错太多了

        mean_mean.insert(0, '平均扣分')
        # mean_mean.insert(1,'班级')
        # print(dict(zip(table_head_values_lst, mean_mean)))

        mean_mean = pd.DataFrame(dict(zip(table_head_values_lst, mean_mean)), index=[0])
        # print(mean_mean)
        tj = pd.DataFrame([value.count()])

        tj['姓名'] = ['扣分人数']  # 添加姓名列，内容是后面的两项，姓名列改内容
        tj = tj[table_head_values_lst]  # 按照table_head_values.columns的顺序排列tj表，不排序就出错了，最后两行统计扣分人数和平均扣分

        # 把统计加入写入表

        value = value.append(mean_mean)
        value = value.append(tj)  # dataframe的append用加法

        combined_table.append([key, value])  # 列表append不用加法

    for e_class, e_class_data in combined_table:
        print(e_class)
        duo(e_class, e_class_data)
    print('OK')


# 数据区
column_dict = {'客 | 一.1': '一.1', '客 | 一.2': '一.2', '客 | 一.3': '一.3', '客 | 一.4': '一.4', '客 | 一.5': '一.5',
               '客 | 一.6': '一.6', '客 | 一.7': '一.7', '客 | 一.8': '一.8', '客 | 一.9': '一.9', '客 | 一.10': '一.10',
               '主 | 11-18': '二.11-18', '主 | 11-16': '二.11-16',
               '主 | 三.19': '三.19', '主 | 三.20': '三.20', '主 | 19': '三.19', '主 | 20': '三.20',
               '主 | 三.17': '三.17', '主 | 三.18': '三.18',
               '主 | 三.21': '三.21', '主 | 三.22': '三.22', '主 | 三.23': '三.23', '主 | 三.24': '三.24',
               '主 | 四.21': '四.21', '主 | 四.22': '四.22',
               '主 | 六.24': '六.24',
               '主 | 五.23': '五.23',
               '主 | 七.25': '七.25',
               '主 | 八.26': '八.26'}  # '':'',


# 数据区


@timer
def main():
    """程序开始."""
    names = os.listdir(os.path.split(os.path.realpath(__file__))[0])
    names = [i for i in names if re.match('小分表.*.xlsx', i)]

    chuli(names[0])


if __name__ == "__main__":
    main()
    print('all ok')

'''



'''
