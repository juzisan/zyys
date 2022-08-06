# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.

需要
1. 数据.xls
2. add.png
3. 安装  pip install opencv-contrib-python
"""
import time

import pandas as pd
import pyautogui
import pyperclip3

print('屏幕大小：', pyautogui.size())
print('鼠标位置：', pyautogui.position())

'''载入人员数据'''
pd_tqf = pd.read_excel('数据.xls', index_col=0, sheet_name='通勤费', dtype={'人员编码': str, '金额': str})
# pd_tqf = pd_tqf[:3]

gao_xz = 0  # edge修正为0，chrome修正为45
'''加号的位置'''
x1, y1 = pyautogui.locateCenterOnScreen('add.png', confidence=0.5)
print('加号位置：', x1, y1)
y2 = y1 - 150

'''点击加号的位置'''
pyautogui.click(x1, y1)

'''等1s后'''
time.sleep(1)

'''开始循环'''
for row in pd_tqf.itertuples():
    time.sleep(0.3)
    pyautogui.click(950, 600 + gao_xz)
    pyautogui.press('tab')
    print(getattr(row, '姓名'))
    '''点击人员编码的位置'''

    pyautogui.write(getattr(row, '人员编码'))

    pyautogui.click(x1, y2 + gao_xz)
    '''点击姓名的位置'''
    po_xm = (1285, 600 + gao_xz)  # 姓名959,570 + gao_xz
    pyautogui.click(po_xm)
    time.sleep(0.5)

    pyautogui.press('tab')
    pyperclip3.copy(getattr(row, '家庭住址'))  # copy data to the clipboard
    pyautogui.hotkey('ctrl', 'v')  # retrieve clipboard contents
    time.sleep(0.3)
    pyautogui.press('tab')
    pyautogui.write(getattr(row, '金额'))

print('well done, OK')