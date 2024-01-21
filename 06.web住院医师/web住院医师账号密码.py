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
import time


def timmer(func):
    def warpper(*args,**kwargs):
        print ('\033[1;32;40mstart\033[0m')
        strat_time = time.time()
        func()
        seconds = time.time() - strat_time
        m, s = divmod(seconds, 60)
        print("\033[1;32;40mthe run time is %02d:%.6f\033[0m" %(m, s))
    return warpper






def read_to_list(neir=None):
    url_lj = url_dict['tou'] + url_dict['dl']#生成？号登陆链接

    values0 =['90','95','100','105','110','115','120','125','130','135','140','145','150']
    values = ['red','green','magenta','chocolate','brown','deeppink',r'#000080']


    json_neir = neir[neir.find("["):neir.rfind("]")+1]

    df = pd.read_json(json_neir)

    df = df[['userName','loginName', 'password', 'departmentName', "subjectName"]]
    hangshu = df['userName'].count()
    df['字体'] = random.sample(values0 * hangshu, hangshu)
    df['颜色'] = random.sample(values * hangshu, hangshu)

    df['link'] = '&rArr;<a href="' + url_lj +df['loginName']+'&password='+df['password']+'" style="font-size:'+df['字体']+'%;color:'+df['颜色']+'">'+df['userName']+'_'+df['subjectName']+'_'+df['departmentName']+'</a>&emsp;'


    return df['link'].tolist()


@timmer
def main():
    url_dl = url_dict['tou'] + url_dict['yy']#登陆链接
    url_xy = url_dict['tou'] + url_dict['xy']#学员
    url_ls = url_dict['tou'] + url_dict['ls']#老师
    url_jm = url_dict['tou'] + url_dict['jm']#教秘
    url_zr = url_dict['tou'] + url_dict['zr']#主任
    session = HTMLSession()
    s_get = session.get

    r = s_get(url_dl)#登陆医院
    print ('login')
#学员
    r = s_get(url_xy)

    textt = r.html.find('script', containing='var userGridData',first = True ).html
    #搜索script 内容为var userGridData，返回第一个，再转符串
    x_list = read_to_list(textt)
    print ("学员：", len(x_list))
    xueyuan_str = ''.join(x_list)


#老师
    r = s_get(url_ls)
    textt = r.html.find('script', containing='var userGridData',first = True ).html
    x_list = read_to_list(textt)
    print ("老师：", len(x_list))
    laoshi_str = ''.join(x_list)

#教秘
    r = s_get(url_jm)
    textt = r.html.find('script', containing='var userGridData',first = True ).html
    x_list = read_to_list(textt)
    print ("教秘：", len(x_list))
    jiaomi_str = ''.join(x_list)


#主任
    r = s_get(url_zr)
    textt = r.html.find('script', containing='var userGridData',first = True ).html
    x_list = read_to_list(textt)
    print ("主任：", len(x_list))
    zhuren_str = ''.join(x_list)

#合并html

    html_x = r'学员</th><td>'
    html_l = r'''</td></tr>
	<tr style="border:2px  dashed;"><th>老师</th><td>'''
    html_j =r'''</td></tr>
	<tr style="border:2px  dashed;"><th>教秘</th><td>'''
    html_z =r'''</td></tr>
	<tr style="border:2px  dashed;"><th>主任</th><td>'''
    html_wei ='''</td></tr>
	</table></h5>
<pre>

</pre>
	</div>
</body>
</html>
'''
    html_body = ''.join([html_tou,html_x,xueyuan_str,html_l,laoshi_str,html_j,jiaomi_str,html_z,zhuren_str,html_wei])



    with codecs.open('zyys.html','w','utf-8') as f2:
        f2.write(html_body)




if __name__ == "__main__":
    main()
