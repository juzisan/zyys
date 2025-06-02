import pandas as pd
import numpy as np
import httpx
import re
import random
import pyautogui
import time
from bs4 import BeautifulSoup

'''
'''
r = httpx.get('https://tool.chinaz.com/dns/GitHub.com')
print(r)

print(r.encoding)

soup = BeautifulSoup(r.text, 'html.parser')
aa = soup.find_all('a')
## print(aa)
### with open('example.txt', 'w') as f:

    ##f.write(soup.text)