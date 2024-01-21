import pandas as pd
import numpy as np
import httpx
import re
import random
import pyautogui
import time

'''
'''
r = httpx.get('https://www.baidu.com/')
print(r)

sc = pyautogui.screenshot('screen.png')