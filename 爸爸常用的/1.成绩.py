# -*- coding: utf-8 -*-
"""
Created on Sat Oct 19 19:52:49 2019

@author: Mrlily
"""


import os
import re
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment,Border,Side,Font




def duo(xlsx_nam,data_xlsx):
    print ()
    xlsx_nam=xlsx_nam+'班'
    #文件名

    wb = Workbook()#建立excel文件
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))#创建边框对象

    ws = wb.create_sheet(xlsx_nam)#创建工作簿名
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE#页面方向
    ws.page_setup.paperSize = ws.PAPERSIZE_A4#纸张大小
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0
    ws.page_margins.top = 0.2
    ws.page_margins.bottom = 0.2
    ws.page_margins.header = 0
    ws.page_margins.footer = 0#设置页边距
    ws.print_title_rows = '1:2'#设置打印标题为2行

    #print (dffff.tail())
    rows = dataframe_to_rows(data_xlsx)#pandas的dataframe转openpyxl的格式
    for row in rows:#不知道为啥有空行。openpyxl的序号从1开始。delete_rows()
        if len(row) >2:
            ws.append(row)
    ws.column_dimensions['A'].width = 3#设置列宽
    ws.column_dimensions['B'].width = 9
    for i in range(3,12):#最大列数目，是整数类型
        ws.column_dimensions[get_column_letter(i)].width = 4
    for i in range(12,ws.max_column):#最大列数目，是整数类型
        ws.column_dimensions[get_column_letter(i)].width = 5
    ws.column_dimensions[get_column_letter(ws.max_column)].width = 6
    for i in list(ws.rows)[0]:#生成器转list才能迭代，第一行
        i.alignment = Alignment(wrap_text=True)#设置文本自动换行
    for row in ws.rows:
        for cell in row:
            cell.border =thin_border#设置边框

    ws.insert_rows(1)#插入标题行
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)#合并标题行
    p_title = xlsx_nam
    ws['A1'] = p_title
    ws['A1'].alignment = Alignment(horizontal='center')#设置标题居中
    ws['A1'].font = Font(size=20)#设置字号大小

    del wb['Sheet']#删除默认sheet，保存文件
    wb.save('试卷分析' + xlsx_nam+'.xlsx')#excel写入sheat



def chuli(xlsx_nam_chuli):
#共有
    global biao_or
    global zong_renshu

    biao_or=pd.read_excel(xlsx_nam_chuli,skiprows=1,dtype={'班级':'str',} )#读excel'总分':'int'
    biao_or=biao_or.sort_values(by =['总分'], ascending=False)
    biao_or.index = range(1,len(biao_or)+1)#重建索引
    bb = biao_or.tail(1).applymap(lambda x: int(re.search(r'得分/共(\d+)分', x).group(1)), na_action='ignore')

    #biao_or.loc[biao_or.index[-1]].apply(lambda x : 6 if x.empty  else x )
    biao_or.drop(biao_or.tail(1).index,inplace=True)
    biao_or = pd.concat([biao_or, bb], ignore_index=True)
    biao_or.rename(columns=lieming_dict, inplace = True)
    biaotou = biao_or.columns.values.tolist()#pandas列名转list
#共有

    ti_list = []
    #print (biaotou)
    dati = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', ]
    for k in dati:
        for j in biaotou:
            if j.find(k) > -1 :
                ti_list.append(j)

    #print (ti_list)

    zutifenxi_fenzhi = biao_or.tail(1)


    tizhi = zutifenxi_fenzhi[ti_list].T.to_dict(orient='dict')
    tizhi = tizhi.popitem()[1]

    print ('题值',tizhi)
    #tizhi = tizhi[:1]


    biao_or.drop(biao_or.tail(1).index,inplace=True)




    shuchu=[]
    for i in biaotou:
        #目的是删除多余的列。生成每题分值dict
        if i.find('.') > 0 :
            pass
        elif i not in ['姓名','班级','总分']:
            shuchu.append(i)
    for i in shuchu:
        biaotou.remove(i)
    biao_or=biao_or[biaotou]#去掉无用列
    biao_or.drop(biao_or.tail(1).index,inplace=True)
    biao_or["合计"] = biao_or["总分"].astype('float') - 100#处理扣总分
    del biao_or['总分']#删除总分列
    biaotou = biao_or.columns.values.tolist()


    for j in tizhi: #变成负分.astype('float')
        biao_or[j]=biao_or[j].astype('float') - tizhi[j]
        biao_or[j]=biao_or[j].replace(0, np.nan)

    #print(biao_or.dtypes)
    biao_or=biao_or.groupby(['班级'])#按照班级分类
    #print (biao_or)
    huizongbiao=[]
    for key, value in biao_or:

        biaotou = value.columns.values.tolist()
        biaotou.remove('班级')
        value=value[biaotou]
        value.index = range(1,len(value)+1)
        mean_mean =value.mean().values.tolist()#pandas的mean出错太多了
        mean_mean.insert(0,'平均扣分')
        #mean_mean.insert(1,'班级')
        #print(dict(zip(biaotou, mean_mean)))

        mean_mean = pd.DataFrame(dict(zip(biaotou, mean_mean)),index=[0])
        #print(mean_mean)
        tj=pd.DataFrame([value.count()])


        tj['姓名']=['扣分人数']#添加姓名列，内容是后面的两项，姓名列改内容
        tj=tj[biaotou]#按照biao_or.columns的顺序排列tj表，不排序就出错了，最后两行统计扣分人数和平均扣分

        #把统计加入写入表

        value=value.append(mean_mean)
        value=value.append(tj)#dataframeappend用加法

        huizongbiao.append([key,value])#列表append不用加法



    for banji, banji_data in huizongbiao:
        print(banji)
        duo(banji, banji_data)
    print('OK')




#数据区
lieming_dict ={ '客 | 一.1':'一.1', '客 | 一.2':'一.2', '客 | 一.3':'一.3', '客 | 一.4':'一.4', '客 | 一.5':'一.5',
               '客 | 一.6':'一.6', '客 | 一.7':'一.7', '客 | 一.8':'一.8', '客 | 一.9':'一.9', '客 | 一.10':'一.10',
               '主 | 11-18':'二.11-18', '主 | 三.19':'三.19', '主 | 三.20':'三.20', '主 | 四.21':'四.21', '主 | 四.22':'四.22',
               '主 | 五.23':'五.23', '主 | 六.24':'六.24', '主 | 七.25':'七.25', '主 | 八.26':'八.26'}


#数据区




def main():
    #程序开始

    names = os.listdir(os.path.split(os.path.realpath(__file__))[0])
    names = [i for i in names if re.match('小分表.*.xlsx', i)]

    chuli(names[0])




if __name__ == "__main__":
    main()
    print ('all ok')

'''



'''