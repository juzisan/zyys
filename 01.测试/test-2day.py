import sympy
from sympy.abc import x, y
from sympy import solve
import pprint
import re

print(solve(x**2 - y, x, dict=True))




ques = sympy.solve([x + y - 1, x - y -3], [x, y])

print(ques)

print (sympy.sqrt(12))

print(sympy.sqrt(-1) )

print(sympy.sqrt(12).evalf(4))

tmp = sympy.series(sympy.exp(sympy.I*x), x, 0, 10)
pprint.pprint(tmp)


aaa ='二维'
# print(re.search(r'[二|三]维', aaa).group())
print('wu')

match str(aaa):
    case 'a':
        print(1)
    case '二维':
        print('二维')
    case '三维':
        print('三维')
    case lost_type:
        print(f'错误{lost_type}')
print('wu')

lie_name_str = '''
门住体图文报告	体检例数	肝纤维化和肝脂肪变测定*2	超声检查正常	脏器灰阶成像*2/3	残余尿测定	床旁彩超加收*5	腔内超声检查	"大排畸
*6"	"胎儿心脏超声
*6"	脏器灰阶成像（NT+产科）	双胎加收*3	胃肠超声*3	"脑黑质测定
*2"	脏器声学造影*5	"临床操作超声引导
*5"	"介入操作例数
*10"	消融例数*20	误时工作量例数	来源住培学员	疑难病例会诊例数	夜班例数	扣罚金额

'''

lie_name_str = lie_name_str.translate(str.maketrans({'"': None, '\n': None}))
lie_name_list = lie_name_str.split('\t')
lie_name_list.insert(0, '检查时间')
print('列名：', lie_name_list)
if '床旁彩超加收*5' not in lie_name_list:
    print('错误：  床旁彩超加收')
    exit()

x_d_list : list = [c_pang_str:= '床旁彩超加收*5',
                   m_z_t_t_w_bao_gao_str:= '门住体图文报告',
                   c_s_j_c_zheng_chang_str:= '超声检查正常',
                   san_wei_str:= '脏器灰阶成像*2/3',
                   can_ke_str:= '脏器灰阶成像（NT+产科）',
                   s_tai_str:= '双胎加收*3']

print(x_d_list)


a = [0,2,3]
b = [0,2,6,5,6,7,8,9,10]

if que_set :=set(a) - set(b) : print('缺少的列名：    ', que_set) # 非空，空的话跳过

