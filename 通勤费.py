# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import time
import pyautogui
import pyperclip3
import pandas as pd

screenWidth, screenHeight = pyautogui.size()
currentMouseX, currentMouseY = pyautogui.position()
print(screenWidth, screenHeight, currentMouseX, currentMouseY)

pd_tqf = pd.read_excel('数据.xls', index_col=0, sheet_name='通勤费', dtype={'人员编码': str, '金额': str})
pd_tqf = pd_tqf[:4]

gao_xz = 0  # edge修正为0，chrome修正为45
po_jh = (1820, 550 + gao_xz)  # 加号1500,515 + gao_xz
pyautogui.click(po_jh)
time.sleep(2)

for row in pd_tqf.itertuples():
    print(getattr(row, '人员编码'))

    po_rybm = (1100, 600 + gao_xz)  # 人员编码775,570 + gao_xz
    pyautogui.click(po_rybm)

    pyautogui.write(getattr(row, '人员编码'))

    pyautogui.click(1415, 319 + gao_xz)

    po_xm = (1280, 600 + gao_xz)  # 姓名959,570 + gao_xz
    pyautogui.click(po_xm)
    time.sleep(0.5)

    pyautogui.press('tab')

    pyperclip3.copy(getattr(row, '家庭住址'))  # copy data to the clipboard
    pyautogui.hotkey('ctrl', 'v')    # retrieve clipboard contents

    pyautogui.press('tab')
    pyautogui.write(getattr(row, '金额'))
