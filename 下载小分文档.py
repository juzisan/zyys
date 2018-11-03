# -*- coding: utf-8 -*-
#!/usr/bin/env python
"""
Created on Sat Nov  5 10:50:13 2016

@author: hello
"""

import os
import time#,datetime
import pandas as pd
from shuju import log_list
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By



def wait_time(xxx):
    for i in range(20):            # 循环60次，从0至59
        if i >= 19 :               # 当i大于等于59时，打印提示时间超时
            print("timeout")
            break
        try:                       # try代码块中出现找不到特定元素的异常会执行except中的代码
            eval(xxx) # 如果能查找到特定的元素id就提前退出循环
            break
        except:                    # 上面try代码块中出现异常，except中的代码会执行打印提示会继续尝试查找特定的元素id
            print("wait for find element")
        time.sleep(1)


real_path = os.path.split(os.path.realpath(__file__))[0]#py所在文件夹
print (real_path)
options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 1, 'download.default_directory': real_path}#下载文件夹设为py所在文件夹
options.add_experimental_option('prefs', prefs)

def xia_zai(xxx):
    browser = webdriver.Chrome(chrome_options=options)
    browser.get('http://fenxi.haofenshu.com/login/')#登录页面

    #time.sleep(3)

    loginName = browser.find_element_by_id('loginName')  # Find the search box
    loginName.send_keys(xxx['loginName'])

    password = browser.find_element_by_id('password')  # Find the search box
    password.send_keys(xxx['password'] + Keys.RETURN)#输入账号密码登录


    locator = (By.CLASS_NAME, 'close')
    try:
        WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator))
        #print (browser.find_element_by_class_name('close'))

    except:
        print ('wrong')

    guanbi_Button = browser.find_element_by_class_name('close')#想多了，根本不是弹窗
    guanbi_Button.click()#关闭弹出来的那个更新说明的窗口

    dakai_Button = browser.find_element_by_xpath('//*[@id="appComp"]/div/div[1]/div/div[2]/div/div/div[1]/div[2]/a')#想多了，根本不是弹窗
    for i in range(20):            # 循环60次，从0至59
        if i >= 19 :               # 当i大于等于59时，打印提示时间超时
            print("timeout")
            break
        try:                       # try代码块中出现找不到特定元素的异常会执行except中的代码
            dakai_Button.click()# 如果能查找到特定的元素id就提前退出循环
            break
        except:                    # 上面try代码块中出现异常，except中的代码会执行打印提示会继续尝试查找特定的元素id
            print("wait for find element")
        time.sleep(1)

    #点开最后一次考试
    locator = (By.XPATH, '//span[text()="重要报表"]')
    WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator)).click()#点开重要报表
    locator = (By.XPATH, '//li[text()="小分表"]')
    WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator)).click()#点开小分表

    locator = (By.XPATH, '//span[text()="数学"]')
    WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator)).click()#点开数学，有的老师不需要点数学

    locator = (By.XPATH, '//*[@id="appComp"]/div/div[2]/div/div[3]/span[2]/span[1]/input')
    WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator))
    checkbox = browser.find_elements_by_xpath("//input[@type='checkbox']")#获取checkbox内容一般为全体加班级个数


    if checkbox[0].is_selected():
        checkbox[0].click()

    names = os.listdir(real_path)
    names = [i for i in names if i.count('小分表')]
    print (names)
    if not names:
        locator = (By.XPATH, '//*[@id="appComp"]/div/div[2]/div/div[4]/div/div/button')
        WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator)).click()#点开下载

    for i in range(20):            # 循环60次，从0至59
        if i >= 19 :               # 当i大于等于59时，打印提示时间超时
            print("timeout")
            break
        names = os.listdir(real_path)
        names = [i for i in names if i.count('小分表')]
        if names:
            biao_or = pd.read_excel(names[0])
            break
        print("wait for find element")
        time.sleep(1)


    biao_tou = biao_or.columns.values.tolist()

    checkbox = checkbox[1:]
    print (len(checkbox))
    for i in checkbox:
        i.click()
        time.sleep(3)
        if i.is_selected():
            locator = (By.XPATH, '//button[@class="set-time"]')
            gong = browser.find_element_by_xpath('//*[@id="appComp"]/div/div[2]/div/div[6]/span')
            print (gong.get_attribute('innerText'))
            locator = (By.XPATH, '//*[@id="appComp"]/div/div[2]/div/div[6]/span/span/div/div')
            WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator)).click()

            check_list = browser.find_elements_by_xpath('//*[@id="appComp"]/div/div[2]/div/div[6]/span/span/div/div/ul/*/a')#显示多少条数据

            kuang_check_list = [disp.text for disp in check_list if disp.text]

            check_list[len(kuang_check_list) -1 ].click()

            dfs = pd.read_html(browser.page_source)#用pd处理快，但是这个页面出现的是并排的两个dataframe，合并就好了。
            dfs = pd.concat(dfs,axis=1)
            dfs.columns = biao_tou
            print (dfs)
            dfs.to_excel('9-' + i.get_attribute('value') + '.xlsx')

            #WebDriverWait(browser, 20, 0.5).until(EC.element_to_be_clickable(locator)).click()
            print ('下载')
        i.click()
    browser.quit()


def main():
    for log_one in log_list:
        xia_zai(log_one)


if __name__ == "__main__":
    main()
    print ('all ok')

'''



'''
