# -*- coding: utf-8 -*-

"""
Spyder Editor

This is a temporary script file.

需要
1. 数据.xlsx
2. add.png
3. 安装
pip install opencv-contrib-python
pip install dddd-ocr
pip install pyautogui
pip install pyperclip3
4. chrome 版本 119.0.6045.160（正式版本） （64 位）
5. OA 协同办公点完 绿色加号+
6. 数据.xlsx 中的列名是    序号 人员编码    姓名  家庭住址
"""

import time
import pandas as pd
import pyautogui
import pyperclip3

# import ddddocr

# ocr = ddddocr.DdddOcr(old=False)

# noinspection PyArgumentList
position_name_code = pyautogui.locateCenterOnScreen('qing_shu_ru_ren_yuan.png', grayscale=False)
# noinspection PyArgumentList
position_name_person = pyautogui.locateCenterOnScreen('qing_shu_ru_xing_ming.png', grayscale=False)


def apply_add(txt_index, name_code, name_person, txt_address):
    """添加通勤费条目."""
    print(txt_index, name_person)

    pyautogui.click(position_name_code)
    # time.sleep(1)
    pyautogui.write(name_code)
    time.sleep(0.1)

    pyautogui.click(position_name_person)
    time.sleep(0.3)
    pyautogui.press('tab')
    pyperclip3.copy(txt_address)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    pyautogui.write('40')

    return


def main():
    """循环执行."""
    print('屏幕大小：', pyautogui.size())
    print('鼠标位置：', pyautogui.position())
    time.sleep(5)
    # apply_conference('肝内囊性病变超声诊断','2024-02-08','17614231154')
    dataframe_1 = pd.read_excel('数据.xlsx', sheet_name='Sheet1',
                                dtype={'序号': str, '人员编码': str, '姓名': str, '家庭住址': str})
    # dataframe_1 = dataframe_1[:2]

    res = dataframe_1.values.tolist()
    for row in res:
        # print(*row)
        apply_add(*row)


if __name__ == "__main__":
    main()
    print('all ok')
