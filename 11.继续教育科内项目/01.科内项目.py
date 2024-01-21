# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.

需要
1. 数据.xlsx
2. add.png
3. 安装  pip install opencv-contrib-python
4. pip install dddd-ocr
"""
import time
import pandas as pd
import pyautogui
import pyperclip3
import ddddocr

ocr = ddddocr.DdddOcr(old=False)


def det(x):
    image = pyautogui.screenshot(region=x)
    res = ocr.classification(image)
    print(res)
    return res


def main():
    gao_xz = 0  # edge修正为0，chrome修正为45
    '''加号的位置'''
    print('屏幕大小：', pyautogui.size())
    print('鼠标位置：', pyautogui.position())
    pyautogui.click(350, 430 + gao_xz)  # 输入工资编号
    pyautogui.write('2046')
    time.sleep(0.5)
    pyautogui.click(350, 200 + gao_xz)  # 点击空白
    time.sleep(0.5)
    pyautogui.click(350, 500 + gao_xz)  # 输入密码
    pyautogui.write('chenG2046')
    time.sleep(0.5)
    pyautogui.click(350, 200 + gao_xz)  # 点击空白
    pyautogui.click(350, 570 + gao_xz)  # 输入验证码
    image_yzm = (450, 550, 140, 55)
    pyautogui.write(det(image_yzm))  # 识别验证码并输入
    time.sleep(0.5)
    pyautogui.click(350, 200 + gao_xz)  # 点击空白
    pyautogui.click(430, 670 + gao_xz)  # 点击登陆
    time.sleep(4)
    pyautogui.click(550, 160 + gao_xz)  # 协同办公
    time.sleep(3)
    pyautogui.click(120, 940 + gao_xz)  # 综合保障部
    time.sleep(1)
    pyautogui.click(120, 760 + gao_xz)  # 通勤费补助
    time.sleep(1)
    pyautogui.click(120, 810 + gao_xz)  # 通勤费补助明细
    time.sleep(3)

    '''载入人员数据'''
    pd_tqf = pd.read_excel('数据.xlsx', index_col=0, sheet_name='通勤费', dtype={'人员编码': str, '金额': str})
    # pd_tqf = pd_tqf[:3]
    # x1, y1 = pyautogui.locateCenterOnScreen('add.png', confidence=0.5)

    pyautogui.click(1850, 550)  # 点击加号

    time.sleep(1)

    '''开始循环'''
    for row in pd_tqf.itertuples():
        time.sleep(0.3)
        pyautogui.click(980, 600 + gao_xz)  # 点击序号
        pyautogui.press('tab')
        print(getattr(row, '姓名'), getattr(row, '人员编码'))
        pyperclip3.copy(getattr(row, '人员编码'))  # write在ZP时出错
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        pyautogui.click(1280, 300 + gao_xz)  # 点击空白
        pyautogui.click(1280, 600 + gao_xz)  # 点击请输入姓名
        time.sleep(0.5)
        pyautogui.click(1280, 300 + gao_xz)  # 点击空白
        pyautogui.click(1570, 650 + gao_xz)  # 点击请输入家庭住址
        pyperclip3.copy(getattr(row, '家庭住址'))  # copy data to the clipboard
        pyautogui.hotkey('ctrl', 'v')  # retrieve clipboard contents
        time.sleep(0.3)
        pyautogui.press('tab')
        pyautogui.write(getattr(row, '金额'))


if __name__ == "__main__":
    main()
    print('all ok')
