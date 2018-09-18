# -*- coding: utf-8 -*-
#!/usr/bin/env python
"""
Created on Sat Nov  5 10:50:13 2016

@author: hello
"""

import sys
import os
import re
import codecs
import copy
import math
import urllib
import urllib.request
import time,datetime
import operator
import zipfile
import shutil
import glob
import requests
import sqlite3
import pandas as pd
from multiprocessing import Pool as mpp
from multiprocessing.dummy import Pool as tpp
from math import log1p as ln
from matplotlib import pyplot
from pandas.tseries.offsets import *
from bs4 import BeautifulSoup
from shuju import all_list
from shuju import url_tou
from shuju import url_dict
from shuju import tou
from shuju import wei



def read_to_list(readonce=None):
    readonce =re.sub(r'\{\}','0',readonce)#eval函数需要对应的
    readonce =re.sub(r'null','0',readonce)#eval函数需要对应的null识别成变量
    findok =re.search(r'(?<=\[).*(?=\])',readonce)#【】需要转义
    pattern =re.compile(r'{.*?}')
    listlist=pattern.findall(findok.group())
    '''aa=[]
    for i in listlist:
        print (i)
        aa.append(eval(i))'''
    aa= [eval(x) for x in listlist]
    #aa= sorted(aa, key=lambda k: k['userName'][:1])#不能按照中文名字排序
    print ([x['userName'] for x in aa])
    return aa

def pr_url(url_str=None):
    return url_tou + url_str


def main():

    my_headers = {
        'User-Agent' : 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_9_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36',
        'Accept' : 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Encoding' : 'gzip',
        'Accept-Language' : 'zh-CN,zh;q=0.8,en;q=0.6,zh-TW;q=0.4'
    }
    sss     = requests.Session()
    r       = sss.get(pr_url(url_dict['yy']), headers = my_headers)


    r       = sss.post(pr_url(url_dict['xy']), headers = my_headers)

    a_str=r.text
    xa_str=a_str[a_str.find('var userGridData='):a_str.find('var userGridGridLoadUrl')]
    list_O = read_to_list(xa_str)
    #xingming = [x['userName'] for x in aa]
    xueyuan_list=[['学员', x['userName']+'_'+x['subjectName'], pr_url(url_dict['dl']) +x['loginName']+'&password='+x['password']] for x in list_O]
    print (len(xueyuan_list))


    r       = sss.post(pr_url(url_dict['ls']), headers = my_headers)

    a_str=r.text
    xa_str=a_str[a_str.find('var userGridData='):a_str.find('var userGridGridLoadUrl')]
    list_O = read_to_list(xa_str)
    #xingming = [x['userName'] for x in aa]
    daijiao_list=[['老师', x['userName']+'_'+x['departmentName'],  pr_url(url_dict['dl']) +x['loginName']+'&password='+x['password']] for x in list_O]
    print (len(daijiao_list))


    r       = sss.post(pr_url(url_dict['jm']), headers = my_headers)

    a_str=r.text
    xa_str=a_str[a_str.find('var userGridData='):a_str.find('var userGridGridLoadUrl')]
    list_O = read_to_list(xa_str)
    #xingming = [x['userName'] for x in aa]
    jiaomi_list=[['敎秘', x['userName']+'_'+x['departmentName'], pr_url(url_dict['dl']) +x['loginName']+'&password='+x['password']] for x in list_O]
    print (len(jiaomi_list))


    r       = sss.post(pr_url(url_dict['zr']), headers = my_headers)

    a_str=r.text
    xa_str=a_str[a_str.find('var userGridData='):a_str.find('var userGridGridLoadUrl')]
    list_O = read_to_list(xa_str)
    #xingming = [x['userName'] for x in aa]
    zhuren_list=[['主任',x['userName']+'_'+x['departmentName'], pr_url(url_dict['dl']) +x['loginName']+'&password='+x['password']] for x in list_O]
    print (len(zhuren_list))
    all_list.extend(xueyuan_list)
    all_list.extend(daijiao_list)
    all_list.extend(jiaomi_list)
    all_list.extend(zhuren_list)
    zhongjian = [str(index+1) + ",'"+x[0]+"','" +x[1] +"','" +x[2] +"'" for index,x in enumerate(all_list)]
    zhongzhi = '),('.join(zhongjian)



    with codecs.open('zyys.sql','w','utf-8') as f2:
        f2.write(tou + zhongzhi + wei)

    print ('完成')


main()
