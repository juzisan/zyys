"""Created on Tues Jan  30 17:28:40 2024.

@author: Mrlily
统计工作量
区分统计体检和门诊住院
体检都是二维，多个项目中间用 【+】 号分割
门诊住院区分二维三维，有些项目特殊，无法统计往诊
"""
# import os
import glob
import re
import time

import numpy as np
import pandas as pd

name_residual_urine: list[str] = []  # 统计残余尿项目名称
rest_day_list: list[int] = []  # 统计休息日
group_day = pd.DataFrame()  # 先生成一个空表

c_f_lie_name_list: list = [c_pang_str := '床旁彩超加收*5',
                           m_z_t_t_w_bao_gao_str := '门住体图文报告',
                           c_s_j_c_zheng_chang_str := '超声检查正常',
                           san_wei_str := '脏器灰阶成像*2/3',
                           can_ke_str := '脏器灰阶成像（NT+产科）',
                           s_tai_str := '双胎加收*3',
                           qiang_nei_str := '腔内超声检查',
                           can_n_str := '残余尿测定',
                           gan_xian_wei_str := '肝纤维化和肝脂肪变测定*2',
                           ti_jian_str := '体检例数']

lie_name_str = '''
门住体图文报告	体检例数	肝纤维化和肝脂肪变测定*2	超声检查正常	脏器灰阶成像*2/3	残余尿测定	床旁彩超加收*5	腔内超声检查	"大排畸
*6"	"胎儿心脏超声
*6"	脏器灰阶成像（NT+产科）	双胎加收*3	胃肠超声*3	"脑黑质测定
*2"	脏器声学造影*5	"临床操作超声引导
*5"	"介入操作例数
*10"	消融例数*20	误时工作量例数	来源住培学员	疑难病例会诊例数	夜班例数	扣罚金额

'''

lie_name_str = lie_name_str.translate(str.maketrans({'"': None, '\n': None}))
lie_name_list = lie_name_str.split('\t')
lie_name_list.insert(0, '检查时间')

if que_set := set(c_f_lie_name_list) - set(lie_name_list): print('缺少的列名：    ', que_set)  # a元素不在b元素里，非空为True，空的话跳过


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


def blank_series():
    """为了生成Series，默认名字，省的重命名."""
    return pd.Series(name='空白', dtype='int', data=None, index=lie_name_list, )


def one_do(txt_str, classify_person, txt_date):
    count_series = blank_series()  # 复制Series
    count_series['检查时间'] = txt_date
    count_series[m_z_t_t_w_bao_gao_str] = 1

    match str(classify_person):
        case '门住':
            # 图文报告
            if re.search(r'三维', txt_str):
                count_series[san_wei_str] = 3
                if txt_str.count(r'胎'):
                    # 胎儿三维再另加1个在疑难病例会诊例数里面
                    # 残尿三维不以三维结尾
                    count_series[can_ke_str] = 1
                    if txt_str.count(r'双胎'):
                        count_series[s_tai_str] = 1
            elif re.search(r'二维', txt_str):
                if txt_str.count(r'卵泡测定'):
                    count_series[c_s_j_c_zheng_chang_str] = 1
                elif txt_str.count(r'经阴道'):
                    count_series[qiang_nei_str] = 1
                elif txt_str.count(r'一个部位'):
                    count_series[c_s_j_c_zheng_chang_str] = 1
                    count_series[c_pang_str] = 1
                    count_series[m_z_t_t_w_bao_gao_str] = 0
                elif txt_str.count(r'二个部位'):
                    count_series[c_s_j_c_zheng_chang_str] = 2
                    count_series[c_pang_str] = 1
                    count_series[m_z_t_t_w_bao_gao_str] = 0
                elif txt_str.count(r'三个部位'):
                    count_series[c_s_j_c_zheng_chang_str] = 3
                    count_series[c_pang_str] = 1
                    count_series[m_z_t_t_w_bao_gao_str] = 0
                elif txt_str.count(r'四个部位'):
                    count_series[c_s_j_c_zheng_chang_str] = 4
                    count_series[c_pang_str] = 1
                    count_series[m_z_t_t_w_bao_gao_str] = 0
                elif txt_str.count(r'五个部位'):
                    count_series[c_s_j_c_zheng_chang_str] = 5
                    count_series[c_pang_str] = 1
                    count_series[m_z_t_t_w_bao_gao_str] = 0
                elif txt_str.count(r'床旁彩超') and not txt_str.count(r'个部位'):
                    # 腹彩加收
                    count_series[c_s_j_c_zheng_chang_str] = 0
                    count_series[c_pang_str] = 1
                    count_series[m_z_t_t_w_bao_gao_str] = 0
                else:
                    count_series[c_s_j_c_zheng_chang_str] = 1
                    if txt_str.count(r'双胎'):
                        count_series[s_tai_str] = 1
            else:
                print('缺少二维三维：', txt_str)
                count_series[c_s_j_c_zheng_chang_str] = 1
            # 残余尿
            if txt_str.count(r'残余'):
                count_series[can_n_str] = 1
                name_residual_urine.append(txt_str)
            if re.match(r'^\[膀胱残余尿(.*?)$', txt_str):
                count_series[c_s_j_c_zheng_chang_str] = 0
        case '体检':
            if txt_str.count(r'肝纤维化和肝脂肪变测定'):
                count_series[gan_xian_wei_str] = 1
                count_series[m_z_t_t_w_bao_gao_str] = 0
            else:
                count_series[ti_jian_str] = txt_str.count(r'+') + 1
        case lost_type:
            print(f'错误：缺少患者类型 --- {lost_type}')

    return count_series


def one_day(day_num):
    """统计一天的工作量."""
    # global group_day
    try:
        df_count = group_day.get_group((day_num,))
        # group by 对象需要用 get_group 才能调用,df用apply传递多个参数的时候要用lambda
    except KeyError:  # 二选一，出错
        rest_day_list.append(day_num)
        return blank_series().rename(day_num)  # 重名名
    else:  # 二选一，正确
        return df_count.sum().rename(day_num)  # 重名名



def do_it(file_str):
    """程序开始."""
    global group_day
    ori_df = pd.read_excel(io=file_str, sheet_name=0, parse_dates=['检查时间'])
    ori_df = ori_df[['检查时间', '患者类型', '检查部位,检查方法', ]]
    ori_df = ori_df.sort_values(by=['检查时间'], ascending=True)
    ori_df.rename(columns={'检查部位,检查方法': '检查部位'}, inplace=True)  # 把列重命名
    ori_df['患者类型'] = ori_df['患者类型'].replace({'住院': '门住', '门诊': '门住'})
    # 把患者类型列的值重新替换一下
    ori_df['检查时间'] = ori_df['检查时间'].dt.day  # 检查时间只保留 day
    new_df = ori_df.apply(lambda row: one_do(row['检查部位'], row['患者类型'], row['检查时间']), axis=1)
    new_df = new_df[lie_name_list]

    group_day = new_df.groupby(['检查时间'])  # 按照检查时间分组

    combination_pd = pd.concat(map(one_day, range(1, 32)), axis=1).T  # 合并表，转置表
    del combination_pd['检查时间']
    combination_pd.loc['总和'] = combination_pd.apply(lambda x: x.sum())  # 各列求和，添加新的行    
    print('\n残尿：', list(set(name_residual_urine)))
    print('休息：', rest_day_list, '  共 ', len(rest_day_list), '天\n')
    combination_pd.replace(0, np.nan, inplace=True)
    print(combination_pd)
    return combination_pd

# @timer
def main():
    """程序开始."""
    names = glob.glob('*.xls')  # 只搜索xls扩展名
    match len(names):
        case 0:
            print('缺少文件')
            exit()
        case 1:
            file_n = names[0]
        case _:
            for i, value in enumerate(names):
                print(i, '代表：  ', value)
            try:
                select_num = int(input("输入转换的序号（ e 退出）："))
                file_n = names[select_num]
            except (ValueError, IndexError):
                print('请重新运行程序')
                exit()

    yue = re.search(r'\d+', file_n)[0]
    do_it(file_n).to_excel('统计好' + yue + '月.xlsx')


if __name__ == "__main__":
    main()
    print('all ok')

'''



'''
