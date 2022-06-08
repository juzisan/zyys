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


# from dataclasses import dataclass


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


def save_file(filename_str, data_dataframe):
    """重复部分."""
    filename_str = filename_str + '班'
    # 文件名

    wb = Workbook()  # 建立excel文件
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))  # 创建边框对象

    ws = wb.create_sheet(filename_str)  # 创建工作簿名
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # 页面方向
    ws.page_setup.paperSize = ws.PAPERSIZE_A4  # 纸张大小
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0
    ws.page_margins.footer = 0  # 设置页边距
    ws.print_title_rows = '1:2'  # 设置打印标题为2行
    rows = dataframe_to_rows(data_dataframe)  # pandas 的dataframe 转openpyxl 的格式
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
    p_title = filename_str
    ws['A1'] = p_title
    ws['A1'].alignment = Alignment(horizontal='center')  # 设置标题居中
    ws['A1'].font = Font(size=20)  # 设置字号大小

    del wb['Sheet']  # 删除默认sheet，保存文件
    wb.save('试卷分析' + filename_str + '.xlsx')  # excel写入sheet
    del wb, ws


def save_file2(filename_str, data_dataframe):
    wb = Workbook()  # 创建workbook实例
    ws = wb.active
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))  # 创建边框对象
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # 页面方向
    ws.page_setup.paperSize = ws.PAPERSIZE_A4  # 纸张大小
    ws.page_margins.left = 0.1
    ws.page_margins.right = 0
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0
    ws.page_margins.footer = 0  # 设置页边距
    ws.print_title_cols = 'A:A'

    for row in ws.rows:
        for cell in row:
            cell.border = thin_border  # 设置边框
    ws.insert_rows(0)  # 插入标题行

    for r in dataframe_to_rows(data_dataframe, index=True, header=True):
        ws.append(r)

    ws.column_dimensions['A'].width = 4  # 设置列宽
    for i in range(2, ws.max_column + 1):  # 最大列数目，是整数类型
        ws.column_dimensions[get_column_letter(i)].width = 5
    for i in range(3, ws.max_row + 1):  # 最大行数目，是整数类型
        ws.row_dimensions[i].height = 43

    for i in range(2, ws.max_column, 2):  # 合并表头
        end_i = i + 1
        ws.merge_cells(start_row=2, start_column=i,
                       end_row=2, end_column=end_i)

    for row in list(ws.rows)[1:]:
        for cell in row:
            cell.border = thin_border  # 设置边框

    ziti = ws['A2:' + get_column_letter(ws.max_column) + str(ws.max_row)]
    for row in list(ziti):  # 单元格选择区域是tuple，转换成list就能赋值了
        for cell in row:
            cell.font = Font(size=11)
            cell.alignment = Alignment(wrap_text=True)  # 设置文本自动换行

    ws['A3'] = ws['A4'].value  # 单元格值需要value
    ws.delete_rows(4)  # 删行
    ws.merge_cells('D1:U1')
    p_title = '数学' + filename_str  # '七年_________考试教学质量分析统计表（数学）'改成前面的了
    ws['D1'] = p_title
    ws['D1'].alignment = Alignment(horizontal='center')  # 设置标题居中
    ws['D1'].font = Font(size=20)  # 设置字号大小

    wb.save('汇总.xlsx')  # excel写入sheet
    del wb, ws


def chuli(filename_str_chuli):
    """共有."""
    p_title = re.findall('^小分表 - ([\\w\\W]*).xlsx$', filename_str_chuli)[0]
    print(p_title)
    original_dataframe = pd.read_excel(filename_str_chuli, skiprows=1, dtype={
        '班级': 'str', })  # 读 excel '总分':'int'
    original_dataframe = original_dataframe.sort_values(by=['总分'], ascending=False)
    original_dataframe.index = range(1, len(original_dataframe) + 1)  # 重建索引 从1开始
    score_dataframe = original_dataframe.tail(1).applymap(lambda x: int(
        re.search(r'得分/共(\d+)分', x).group(1)), na_action='ignore')  # 后面还得用它合并

    valid_score_dataframe = score_dataframe.dropna(axis=1, how='all')  # 最后一行，有数据的是每题分值
    original_dataframe.drop(original_dataframe.tail(1).index, inplace=True)  # 删除最后一行，之前的只是另存
    valid_score_dict = valid_score_dataframe.to_dict('index').popitem()[1]  # dataframe 转 dict 类型

    valid_score_dict_keys = valid_score_dict.keys()  # 获取键值 转 dict_keys 类型
    rename_dict = {}
    score_dict = {}
    keep_column_list = ['姓名', '班级', '总分']  # 列表中的键值将会是数据表中保留的列
    bignumber_index = 0
    bignumber_list = list(bignumber_dict.keys())
    for i in valid_score_dict_keys:
        j = i.replace("客 | ", "")
        j = j.replace("主 | ", "")
        index_str = list(j)[0]
        '''没有汉字大题的处理'''
        if index_str in bignumber_list:
            bignumber_index = bignumber_list.index(index_str) +1
        else:
            index_str = bignumber_list[bignumber_index]
            j = index_str + '.' + j
        rename_dict[i] = j
        score_dict[j] = valid_score_dict[i]
        keep_column_list.append(j)
        bignumber_dict[index_str].append(j)
    for i in list(bignumber_dict.keys()):
        if not bignumber_dict[i]:
            del bignumber_dict[i]
    print(bignumber_dict)
    # original_dataframe.loc[original_dataframe.index[-1]].apply(lambda x : 6 if x.empty  else x )
    sum_score_number = sum(score_dict.values())  # 求总分
    print('题值：', score_dict, "\n总分：    ", sum_score_number)  # 核对总分

    # original_dataframe = pd.concat([original_dataframe, score_dataframe], ignore_index=True) # 合并表格
    original_dataframe.rename(columns=rename_dict, inplace=True)  # 把列重命名
    original_dataframe = original_dataframe[keep_column_list]  # 去掉无用列

    """共有."""
    """1.生成负分表；2.生成汇总表"""

    original_dataframe["合计"] = original_dataframe["总分"].astype('float') - sum_score_number  # 处理扣总分
    analyze_dataframe = original_dataframe.copy()
    del original_dataframe['总分']  # 删除总分列

    for j in score_dict:  # 变成负分.astype('float')
        original_dataframe[j] = original_dataframe[j].astype('float') - score_dict[j]
        original_dataframe[j] = original_dataframe[j].replace(0, np.nan)

    # print(original_dataframe.dtypes)
    original_dataframe = original_dataframe.groupby(['班级'])  # 按照班级分类,分类并不重建索引，沿用没分类之前的索引
    print(original_dataframe)
    summary_list = []
    for key, value in original_dataframe:
        mean_series = value.mean(numeric_only=True)  # 平均扣分 行
        mean_series['姓名'] = "平均扣分"  # 重新赋值

        count_series = value.count()  # 扣分人数 行
        count_series['姓名'] = "扣分人数"  # 重新赋值
        value = value.append(mean_series, ignore_index=True)  # 插入平均扣分行
        value = value.append(count_series, ignore_index=True)  # 插入扣分人数行
        value['班级'] = str(key) + '班'
        summary_list.append(value[-2:].copy())

        del value['班级']  # 删除 班级 列
        value.index = range(1, len(value) + 1)  # 重建索引
        """1.负分表保存文件"""
        save_file(key, value)
        print(key, "班", type(value), len(value), '人')

    summary_dataframe = pd.concat(summary_list, ignore_index=True)  # 合并dataframe
    headline_list = summary_dataframe.columns.values.tolist()  # 把列名转换成 list 格式
    headline_list = headline_list[2:-1]  # 删除前面和后面的列名
    headline_new_list = []
    for i in headline_list:
        headline_new_list.append((i, '扣分人数'))
        headline_new_list.append((i, '平均扣分'))
    # 制作复合列名
    summary_dataframe = summary_dataframe.pivot_table(
        index='班级', columns='姓名', margins=False)  # 创建透视表，因为列名需要占用两行
    summary_dataframe = summary_dataframe[headline_new_list]  # 重现对列名进行格式化
    '''2.汇总表保存文件'''
    save_file2(p_title, summary_dataframe)
    print(summary_dataframe)
    '''3.数据分析'''
    '''3.1.全年级分析'''
    '''总人数count()后面不用[]，有筛选的就要用[]了'''
    valid_bignumber_dict = {}
    for bignumber_key, bignumber_value in bignumber_dict.items():
        analyze_dataframe[bignumber_key] = analyze_dataframe[bignumber_value].sum(axis=1)  # 根据列名求和
        valid_bignumber_dict[bignumber_key] = 
        
    keep_column_list = ['班级', '总分'] + list(bignumber_dict.keys())

    analyze_dataframe = analyze_dataframe[keep_column_list]  # 去掉无用列
    total_number = analyze_dataframe['总分'].count()
    pass_number = analyze_dataframe[analyze_dataframe["总分"] > sum_score_number*0.59].count()["总分"]  # 及格人数
    excellent_number = analyze_dataframe[analyze_dataframe["总分"] > 84.6].count()["总分"]  # 优秀人数
    f_number = analyze_dataframe[analyze_dataframe["总分"] < 39.7].count()["总分"]  # 过差人数
    '''新建一个空的dataframe，索引为all'''
    total_dataFrame = pd.DataFrame(index=["全年级"])
    total_dict = {'总人数': 'total_number',
                   '总分': 'sum_score_number',
                   '平均分': "round(analyze_dataframe['总分'].mean(),1)",
                   '及格人数': 'pass_number',
                   '及格率': 'round(pass_number/total_number*100,1)',
                   '优秀人数': 'excellent_number',
                   '优秀率': 'round(excellent_number/total_number*100,1)',
                   '过差人数': 'f_number',
                   '过差率': 'round(f_number/total_number*100,1)',
                   }

    for x in total_dict:
        exec("total_dataFrame['" + x + "'] = " + total_dict[x])

    
    print(total_dataFrame)
    '''3.1.全年级分析'''
    '''3.2.每班分析'''
    class_dict = {'本题满分': 'manfen_zhi',
                   '满分人数': 'manfen_renshu',
                   '满分率': 'round(manfen_renshu/zong_renshu*100,1)',
                   '及格人数': 'jige_renshu_zutifenxi',
                   '及格率': 'round(jige_renshu_zutifenxi/zong_renshu*100,1)',
                   '0分人数': 'lingfen_renshu_zutifenxi',
                   '0分率': 'round(lingfen_renshu_zutifenxi/zong_renshu*100,1)',
                   '本题平均分': 'round(yilie.mean(),1)',
                   '本题得分率': 'round(yilie.sum()/zong_renshu/manfen_zhi*100,1)',
                   '最高分': 'yilie.max()',
                   '最低分': 'yilie.min()',
                   }
    


    analyze_dataframe = analyze_dataframe.groupby(['班级'])  # 按照班级分类
    print(bignumber_dict)
    for key, value in analyze_dataframe:
        class_dataFrame = pd.DataFrame(index=[key])
        value.index = range(1, len(value) + 1)  # 重建索引
        perfect_number = value[value["总分"] > sum_score_number*0.99].count()["总分"]
        pass_number = value[value["总分"] > sum_score_number * 0.59].count()["总分"]
        zero_number = value[value["总分"] == 0].count()["总分"]



            

        print(key,value.tail(),)


        for x in class_dict:
            exec("class_dataFrame['" + x + "'] = " + class_dict[x])
        print(class_dataFrame)
    del total_number, pass_number, excellent_number, f_number, perfect_number, zero_number  # 为了不显示错误
    print('OK')


# 数据区

bignumber_dict = {'一':[],'二':[],'三':[],'四':[],'五':[],'六':[],'七':[],'八':[],'九':[],'十':[] }

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
