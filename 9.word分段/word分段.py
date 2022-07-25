import os
import re
import shutil
import time

from win32com.client import Dispatch


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


def do_it(file_n):
    dir_str = os.getcwd()  # 本地目录
    src_file = os.path.join(dir_str, file_n)
    file_name = '段落分析' + file_n
    dst_file = os.path.join(dir_str, file_name)
    if os.path.exists(dst_file):
        os.unlink(dst_file)  # 文件存在则删除文件，不然文件存在的时候，复制会出错
    shutil.copyfile(src_file, dst_file)  # 复制文件
    print(f'{dst_file = }')
    msword = Dispatch('Word.Application')  # 打开微软word程序，后面得关闭
    msword.Visible = 1  # 显示word窗口
    doc = msword.Documents.Open(dst_file)  # 打开word文件
    replace_t = ('xxx-xxx', False, False, False, False, False, True, 1, True, 'name_of_class + p_title', 2)
    msword.Selection.Find.Execute(*replace_t)
    # 教程不明确，替换是在应用里替换，所以前面用的是msword，而不是文档doc。
    doc.Save()
    doc.Close()  # 保存并退出
    msword.Quit()  # 手动关闭word程序


@timer
def main():
    """程序开始."""
    names = os.listdir(os.path.split(os.path.realpath(__file__))[0])
    names = [i for i in names if re.match(r'(.*?)\.(doc|docx)$', i)]
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