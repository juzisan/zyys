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
from shuju import url_dict
from shuju import html_tou

url_dl = url_dict['tou'] + url_dict['yy']#登陆链接
url_lj = url_dict['tou'] + url_dict['dl']#生成？号登陆链接
url_xy = url_dict['tou'] + url_dict['xy']#学员
url_ls = url_dict['tou'] + url_dict['ls']#老师
url_jm = url_dict['tou'] + url_dict['jm']#教秘
url_zr = url_dict['tou'] + url_dict['zr']#主任


def read_to_list(neir=None):

    values0 =['90','95','100','105','110','115','120','125','130','135','140','145','150']
    values = ['red','green','magenta','chocolate','brown','deeppink',r'#000080']


    kaitou = neir.find("[")
    jiewei = neir.rfind("]")
    neir = neir[kaitou:jiewei+1]

    df = pd.read_json(neir)

    df = df[['userName','loginName', 'password', 'departmentName', "subjectName"]]

    df['字体'] = random.sample(values0 * df['userName'].count(), df['userName'].count())
    df['颜色'] = random.sample(values * df['userName'].count(), df['userName'].count())

    df['link'] = '&rArr;<a href="' + url_lj +df['loginName']+'&password='+df['password']+'" style="font-size:'+df['字体']+'%;color:'+df['颜色']+'">'+df['userName']+'_'+df['subjectName']+'_'+df['departmentName']+'</a>&emsp;'


    return df['link'].tolist()



def main():
    session = HTMLSession()

    r = session.get(url_dl)#登陆医院
    print ('login finished')
#学员
    r = session.get(url_xy)

    textt = r.html.find('script', containing='var userGridData',first = True ).html
    #搜索script 内容为var userGridData，返回第一个，再转符串
    xueyuan_list = read_to_list(textt)
    print ("学员：", len(xueyuan_list))


#老师
    r = session.get(url_ls)
    textt = r.html.find('script', containing='var userGridData',first = True ).html
    daijiao_list = read_to_list(textt)
    print ("老师：", len(daijiao_list))


#教秘
    r = session.get(url_jm)
    textt = r.html.find('script', containing='var userGridData',first = True ).html
    jiaomi_list = read_to_list(textt)
    print ("教秘：", len(jiaomi_list))



#主任
    r = session.get(url_zr)
    textt = r.html.find('script', containing='var userGridData',first = True ).html
    zhuren_list = read_to_list(textt)
    print ("主任：", len(zhuren_list))


#合并html
    html_body = html_tou+r'学员</th><td>'
    for x in xueyuan_list:
        html_body+=x
    html_body+=r'''</td></tr>
	<tr style="border:2px  dashed;"><th>带教</th><td>'''
    for x in daijiao_list:
        html_body+=x
    html_body+=r'''</td></tr>
	<tr style="border:2px  dashed;"><th>教秘</th><td>'''
    for x in jiaomi_list:
        html_body+=x
    html_body+=r'''</td></tr>
	<tr style="border:2px  dashed;"><th>主任</th><td>'''
    for x in zhuren_list:
        html_body+=x
    html_body+='''</td></tr>
	</table></h5>
<pre>

</pre>
	</div>
</body>
</html>
'''




    with codecs.open('zyys.html','w','utf-8') as f2:
        f2.write(html_body)
    print ('OK')

main()
