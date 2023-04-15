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
import numpy as np
import pandas as pd

global yue

col_n = ['体检例数', '腔内超声检查', '图文报告', '超声检查正常(包括双胎)', '双胎加收',
             '脏器灰阶立体成象', "超声检查(胎儿系统)", "胎儿心脏超声", '残余尿测定',
             "介入操作", "床旁彩超加收", 'B超常规检', "脏器声学造影", '临床操作超声引导', '弹性成像']
k_s = pd.Series(name='空白', dtype=float, data=None, index=col_n)
# 为了生成Series

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


def one_do(neir_str, leix_type):
    jishu = k_s.copy(deep=True) # 复制Series
    if leix_type == '门住':
        jishu['图文报告'] = 1
        if neir_str.count(r'三维'):
            jishu['脏器灰阶立体成象'] = 1
        else:
            if neir_str.count(r'经阴道'):
                jishu['腔内超声检查'] = 1
            elif neir_str.count(r'床旁彩超'):
                jishu['床旁彩超加收'] = 1
                if neir_str.count(r'一个部位'):
                    jishu['超声检查正常(包括双胎)'] = 1
                elif neir_str.count(r'二个部位'):
                    jishu['超声检查正常(包括双胎)'] = 2
                elif neir_str.count(r'三个部位'):
                    jishu['超声检查正常(包括双胎)'] = 3
                elif neir_str.count(r'四个部位'):
                    jishu['超声检查正常(包括双胎)'] = 4
                elif neir_str.count(r'五个部位'):
                    jishu['超声检查正常(包括双胎)'] = 5
                else:
                    jishu['超声检查正常(包括双胎)'] = 'wrong'
            else:
                jishu['超声检查正常(包括双胎)'] = 1
        if neir_str.count(r'残余尿'):
            jishu['残余尿测定'] = 1
            if neir_str.count(r'及残余尿'):
                jishu['超声检查正常(包括双胎)'] = 0
                jishu['脏器灰阶立体成象'] = 0
    elif leix_type == '体检':
        jishu['体检例数'] = neir_str.count(r'+') + 1
    else:
        print('wrong 患者类型')
    return jishu



def do_it(file_str):
    """程序开始."""
    global yue
    ori_df = pd.read_excel(io=file_str, sheet_name=0, parse_dates=['检查时间'])
    ori_df = ori_df[['检查时间', '患者类型', '检查部位,检查方法',]]
    ori_df = ori_df.sort_values(by=['检查时间'], ascending=True)
    ori_df.rename(columns={'检查部位,检查方法': '检查部位'}, inplace=True)  # 把列重命名
    ori_df['患者类型'] = ori_df['患者类型'].replace({'住院':'门住', '门诊':'门住'})
    # 把患者类型列的值重新替换一下
    ori_df['检查时间'] = ori_df['检查时间'].dt.day # 检查时间只保留 day
    fenzu = ori_df.groupby(['检查时间']) # 按照检查时间分组

    zonghe = [] # 准备列表生成统计
    for i in range(1,32):
        try:
            jisruan = fenzu.get_group(i).apply(lambda row :one_do(row['检查部位'],row['患者类型']),axis=1)
            # groupby 对象需要用 get_group 才能调用,df用apply传递多个参数的时候要用lambda
            jishu = jisruan.sum()
            jishu.rename(i,inplace=True) # 对 Series 重命名
        except:
            jishu = k_s.copy(deep=True)
            jishu.rename(i,inplace=True) # 对 Series 重命名
            print('没上班')
        zonghe.append(jishu)

    zonghe = pd.concat(zonghe, axis=1).T # 合并表，转置表
    zonghe.loc['总和'] = zonghe.apply(lambda x: x.sum()) # 各列求和，添加新的行
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
            yue = re.search(r'\d+', file_n)[0]
            do_it(file_n)
        else:
            do_it(names[0])
    else:
        print('缺少文件')


if __name__ == "__main__":
    main()
    print('all ok')

'''



'''