# -*- coding: utf-8 -*-

"""
Spyder Editor

This is a temporary script file.

需要
1. 数据.xlsx
2. add.png
3. 安装
pip install opencv-contrib-python
pip install ddddocr
pip install pyautogui
pip install pyperclip
4.  Chromium 版本 136.0.7070.0（开发者内部版本） （64 位）
5. OA 协同办公点完 绿色加号 +
6. 数据.xlsx 中 Sheet1 的列名是    序号 人员编码    姓名  家庭住址
"""

import time
import pandas as pd
import pyautogui
import pyperclip

# import ddddocr

# ocr = ddddocr.DdddOcr()
for i in range(5):
    time.sleep(1)
    print(i+1)

name_code_png: str = 'w1.png'  # "qing_shu_ru_ren_yuan.png"
name_person_png: str = 'w2.png'  # "qing_shu_ru_xing_ming.png"

# noinspection PyArgumentList
position_name_code = pyautogui.locateCenterOnScreen(image=name_code_png)  #  image=name_code_png
# noinspection PyArgumentList
position_name_person = pyautogui.locateCenterOnScreen(image=name_person_png)  #  image=name_person_png
print(position_name_code, position_name_person)


def apply_add(txt_index, name_code, name_person, txt_address):
    """添加通勤费条目."""
    print(txt_index, name_person, name_code)

    pyautogui.click(position_name_code)
    time.sleep(0.7)
    pyperclip.copy(name_code)
    pyautogui.hotkey('ctrl', 'v')
    # pyautogui.write(name_code)
    time.sleep(0.7)

    pyautogui.click(position_name_person)
    time.sleep(1)
    pyautogui.press('tab')
    pyperclip.copy(txt_address)
    # pyperclip.paste()
    # print(pyperclip.paste())
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    pyautogui.write('40')

    return


def main():
    """循环执行."""
    print('屏幕大小：', pyautogui.size())
    print('鼠标位置：', pyautogui.position())

    # apply_conference('肝内囊性病变超声诊断','2024-02-08','17614231154')
    dataframe_1 = pd.read_excel('数据.xlsx', sheet_name='Sheet1',
                                dtype={'序号': str, '人员编码': str, '姓名': str, '家庭住址': str})
    # dataframe_1 = dataframe_1[:2]
    dataframe_1 = dataframe_1.iloc[::-1]
    res = dataframe_1.values.tolist()
    for row in res:
        # print(*row)
        apply_add(*row)


if __name__ == "__main__":
    main()
    print('all ok')
