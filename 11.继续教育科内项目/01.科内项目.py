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
"""

import time
import pandas as pd
import pyautogui
import pyperclip3
# import ddddocr

# ocr = ddddocr.DdddOcr(old=False)


def click_center(picture):
    """点击图片中心."""
    # noinspection PyArgumentList
    pyautogui.click(pyautogui.locateCenterOnScreen(picture, grayscale=False))
    return


def apply_conference(txt_date, name_conference, name_person, txt_phone):
    """申请举办活动."""
    print(name_person)
    click_center('shen_qing_ju_ban_huo_dong.png')
    time.sleep(2)
    click_center('qing_shu_ru_huo_dong_ming_chen.png')
    time.sleep(0.5)
    pyperclip3.copy(name_conference)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    click_center('er_ji_xue_ke.png')
    time.sleep(0.5)
    pyautogui.move(0, 200)
    time.sleep(0.5)
    pyautogui.scroll(-170)
    time.sleep(0.5)
    click_center('ying_xing_yi_xue.png')

    time.sleep(1)
    click_center('san_ji_xue_ke.png')
    time.sleep(0.5)
    click_center('chao_sheng_zhen_duan_xue.png')

    time.sleep(1)
    click_center('huo_dong_xing_shi.png')
    time.sleep(0.5)
    click_center('lin_chuang_bing_li_tao_lun_hui.png')
    time.sleep(0.5)
    pyautogui.click()
    pyautogui.write('1.5')
    click_center('xuan_ze_ri_qi.png')
    pyautogui.write(txt_date)
    pyautogui.press('enter')
    time.sleep(1)
    pyperclip3.copy('15:00')
    click_center('qi_shi_shi_jian.png')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(1)
    pyperclip3.copy('16:30')
    click_center('jie_shu_shi_jian.png')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(1)
    click_center('ju_ban_di_dian.png')
    pyperclip3.copy('超声科教室')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('pagedown')
    time.sleep(1)
    click_center('lian_xi_ren_xing_ming.png')
    pyperclip3.copy('程日')
    pyautogui.hotkey('ctrl', 'v')
    click_center('lian_xi_ren_dian_hua.png')
    pyautogui.write('15698938681')
    pyautogui.hotkey('ctrl', 'v')
    click_center('ij_hua_pei_xun_ren_shu.png')
    pyautogui.write('30')
    pyautogui.hotkey('ctrl', 'v')
    click_center('huo_dong_ri_cheng_tian_jia.png')
    time.sleep(1)
    click_center('xuan_ze.png')
    time.sleep(1)
    click_center('qing_shu_ru_shou_ji_hao.png')
    pyautogui.write(txt_phone)
    click_center('sou_suo.png')
    time.sleep(1)
    pyautogui.press('tab')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click()
    pyautogui.press('tab', presses=3)
    pyperclip3.copy(name_conference)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')
    pyautogui.write('1.5')
    pyautogui.press('tab', presses=2)
    pyautogui.press('enter')
    time.sleep(1)
    click_center('huo_dong_shuo_ming.png')
    pyautogui.hotkey('ctrl', 'v')
    click_center('ti_jiao.png')
    time.sleep(3)

    return


def main():
    """循环执行."""
    print('屏幕大小：', pyautogui.size())
    print('鼠标位置：', pyautogui.position())
    # apply_conference('肝内囊性病变超声诊断','2024-02-08','17614231154')
    dataframe_1 = pd.read_excel('数据.xlsx', sheet_name='Sheet1',
                                dtype={'讲课时间': str, '授课内容': str, '主讲人': str, '手机号': str})
    # dataframe_1 = dataframe_1[:2]

    res = dataframe_1.values.tolist()
    for row in res:
        print(*row)
        apply_conference(*row)


if __name__ == "__main__":
    main()
    print('all ok')
