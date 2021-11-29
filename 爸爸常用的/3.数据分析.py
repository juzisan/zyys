"""Created on Sat Oct 19 19:52:49 2019.

@author: Mrlily
"""

import os
import pandas as pd
import re
import time


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


# 数据区
lieming_dict = {'客 | 一.1': '一.1', '客 | 一.2': '一.2', '客 | 一.3': '一.3', '客 | 一.4': '一.4', '客 | 一.5': '一.5',
                '客 | 一.6': '一.6', '客 | 一.7': '一.7', '客 | 一.8': '一.8', '客 | 一.9': '一.9', '客 | 一.10': '一.10',
                '主 | 11-18': '二.11-18', '主 | 三.19': '三.19', '主 | 三.20': '三.20', '主 | 四.21': '四.21', '主 | 四.22': '四.22',
                '主 | 五.23': '五.23', '主 | 六.24': '六.24', '主 | 七.25': '七.25', '主 | 八.26': '八.26', '主 | www.26': '八.26'}


# 数据区


def chuli(xlsx_nam_chuli):
    """共有."""
    global biao_or
    global zong_renshu

    biao_or = pd.read_excel(xlsx_nam_chuli, skiprows=1, dtype={
        '班级': 'str', })  # 读excel'总分':'int'
    biao_or = biao_or.sort_values(by=['总分'], ascending=False)
    biao_or.index = range(1, len(biao_or) + 1)  # 重建索引
    bb = biao_or.tail(1).applymap(lambda x: int(
        re.search(r'得分/共(\d+)分', x).group(1)), na_action='ignore')

    # biao_or.loc[biao_or.index[-1]].apply(lambda x : 6 if x.empty  else x )
    biao_or.drop(biao_or.tail(1).index, inplace=True)
    biao_or = pd.concat([biao_or, bb], ignore_index=True)
    biao_or.rename(columns=lieming_dict, inplace=True)
    biaotou = list(biao_or.columns.values)  # pandas列名转list,先转成的迭代器
    # 共有

    print(type(biaotou))
    dati_list = []
    dati = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', ]
    for k in dati:
        qiuhe = [x for x in biaotou if x.find(k) > -1]
        if len(qiuhe) > 0:
            biao_or[k] = biao_or[qiuhe].sum(axis=1)  # 根据列名求和
            dati_list.append(k)
        else:
            break

    zutifenxi_fenzhi = biao_or.tail(1)
    # print (dati_list)

    '''   dati_list = set(dati_list)
    dati = list(filter(lambda n:n in dati_list, dati))#[ x for x in dati if x in dati_list]
    print(dati)'''

    tizhi = zutifenxi_fenzhi[dati_list].T.to_dict(orient='dict')
    print(tizhi)

    tizhi = tizhi.popitem()[1]
    tizhi = tizhi.items()  # 把dict转成list

    print(tizhi)
    # tizhi = tizhi[:2]

    biao_or.drop(biao_or.tail(1).index, inplace=True)

    # 在循环用exec引用

    zong_renshu = biao_or['总分'].count()  # 总人数，全局变量
    jige_renshu = biao_or[biao_or["总分"] > 59.6].count()["总分"]  # 及格人数
    youxiu_renshu = biao_or[biao_or["总分"] > 84.6].count()["总分"]  # 优秀人数
    guocha_renshu = biao_or[biao_or["总分"] < 39.7].count()["总分"]  # 过差人数

    # 在循环用exec引用

    zongti_pd = pd.DataFrame(index=["all"])

    zongti_dict = {'总人数': 'zong_renshu',
                   '平均分': "round(biao_or['总分'].mean(),1)",
                   '及格人数': 'jige_renshu',
                   '及格率': 'round(jige_renshu/zong_renshu*100,1)',
                   '优秀人数': 'youxiu_renshu',
                   '优秀率': 'round(youxiu_renshu/zong_renshu*100,1)',
                   '过差人数': 'guocha_renshu',
                   '过差率': 'round(guocha_renshu/zong_renshu*100,1)',
                   }

    for x in zongti_dict:
        exec("zongti_pd['" + x + "'] = " + zongti_dict[x])

    del jige_renshu, youxiu_renshu, guocha_renshu  # 为了不显示错误

    print(biao_or)

    hebin_df = map(zutifenxi, tizhi)  # 函数用一个变量更容易解决解包的问题
    hebin_df = pd.concat(hebin_df, ignore_index=False)

    print(zongti_pd)
    print(hebin_df)
    print('卷面满分：', hebin_df['本题满分'].sum())
    print('OK')


def zutifenxi(nei):
    """逐题分析."""
    global zuiti_pd
    yilie, manfen_zhi = nei

    zuiti_pd = pd.DataFrame(index=[yilie])
    yilie = biao_or[yilie]

    # 在循环用exec引用
    manfen_renshu = yilie[yilie == manfen_zhi].count()
    jige_renshu_zutifenxi = yilie[yilie > manfen_zhi * 0.6].count()
    lingfen_renshu_zutifenxi = yilie[yilie == 0].count()

    # 在循环用exec引用

    zixing_dict = {'本题满分': 'manfen_zhi',
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

    for x in zixing_dict:
        exec("zuiti_pd['" + x + "'] = " + zixing_dict[x])

    del manfen_renshu, jige_renshu_zutifenxi, lingfen_renshu_zutifenxi  # 为了不显示错误

    return zuiti_pd


@timer
def main():
    """程序开始."""
    names = os.listdir(os.path.split(os.path.realpath(__file__))[0])
    # [i for i in names if re.match('小分表.*.xlsx', i)]
    names = list(filter(lambda x: re.match('小分表.*.xlsx', x), names))

    chuli(names[0])


if __name__ == "__main__":
    main()
    print('all ok')

'''



'''
