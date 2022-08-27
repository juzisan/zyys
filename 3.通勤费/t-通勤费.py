# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.

需要
1. 数据.xls
2. add.png
3. 安装  pip install opencv-contrib-python
4. pip install ddddocr
"""
import time
import pandas as pd
import pyautogui
import pyperclip3
import ddddocr


print('屏幕大小：', pyautogui.size())
print('鼠标位置：', pyautogui.position())

ocr = ddddocr.DdddOcr(old=False)


def det_login():
    image = pyautogui.screenshot(region=(400, 550, 120, 50))
    res = ocr.classification(image)
    print(res)


def det(x):
    image = pyautogui.screenshot(region=x)
    res = ocr.classification(image)
    print(res)


image_ren = (1050, 580, 120, 40)
det(image_ren)
time.sleep(1)
