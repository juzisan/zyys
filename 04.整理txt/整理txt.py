# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.

需要
输入序号
自动删除空行
自动删除段首空格
"""
import time
import codecs
import os


def timmer(func):
    def warpper(*args, **kwargs):
        print('\033[1;32;40mstart\033[0m')
        strat_time = time.time()
        func()
        if "time_pro" in globals().keys():
            time_pro1 = time_pro
        else:
            time_pro1 = 0
        seconds = time.time() - strat_time - time_pro1
        m, s = divmod(seconds, 60)
        print("\033[1;32;40mthe run time is %02d:%.6f\033[0m" % (m, s))

    return warpper


@timmer
def main():
    global time_pro

    file_name_lst = [i for i in os.listdir() if i.endswith('.txt')]
    for i, val in enumerate(file_name_lst):
        print(i, val)
    time_d_s = time.time()
    select_num = int(input("输入转换的序号："))

    print("请等待！")
    time_d_e = time.time()

    time_pro = time_d_e - time_d_s
    print(time_pro)
    with codecs.open(file_name_lst[select_num], 'r', 'utf-8') as f2:
        read_txt = f2.readlines()
        write_txt = []
        for i in read_txt:
            neirong_str = i.strip()
            neirong_str = neirong_str.strip("=")
            if neirong_str:
                write_txt.append(neirong_str + "\r\n")

        # f2.write(ww)

    with codecs.open(file_name_lst[select_num], 'w', 'utf-8') as f2:
        f2.writelines(write_txt)

    print('完成！')


if __name__ == "__main__":
    main()
