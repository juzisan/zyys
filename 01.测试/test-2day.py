import sympy
from sympy.abc import x, y
from sympy import solve
import pprint

print(solve(x**2 - y, x, dict=True))




ques = sympy.solve([x + y - 1, x - y -3], [x, y])

print(ques)

print (sympy.sqrt(12))

print(sympy.sqrt(-1) )

print(sympy.sqrt(12).evalf(4))

tmp = sympy.series(sympy.exp(sympy.I*x), x, 0, 10)
pprint.pprint(tmp)
