# -*- coding: utf-8 -*-
#!/usr/bin/env python
"""
Created on Sat Nov  5 10:50:13 2016

@author: hello
"""


import codecs
import pandas as pd
from requests_html import HTMLSession
import random
import time
import os


def timmer(func):
    def warpper(*args,**kwargs):
        print ('\033[1;32;40mstart\033[0m')
        strat_time = time.time()
        func()
        seconds = time.time() - strat_time
        m, s = divmod(seconds, 60)
        print("\033[1;32;40mthe run time is %02d:%.6f\033[0m" %(m, s))
    return warpper


@timmer
def main():
    names = os.listdir()[:-1]


    print (names)
    for i in names:
        print (os.listdir(i))
        os.rename('/'.join([i,os.listdir(i)[0]]),'-'.join([i,os.listdir(i)[0]]))




if __name__ == "__main__":
    main()
