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