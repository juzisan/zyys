### 特殊注意

######

---
### 读取Excel文件
```python
import pandas as pd
df = pd.read_excel(io='file_str', sheet_name=0, dtype={'班级': 'str'}, skiprows=1)
```
#####
read_excel必须要有sheet_name参数，0为第一个，None为所有都读取，读到一个字典里

---
### dataframe以及series连接要用concat
####
1. concat连接的是一个列表
2. 如果要是两个series连接，axis=1
3. 连接两个series后的dataframe需要用大写的 .T 转置才能和想要的dataframe连接
4. 复制dataframe需要用 .copy(deep=True) 的深度复制才不是指向，能避免错误操作
5. concat，连接dataframe，可用 ignore_index=True 忽略索引
___
### dataframe的sum可以用 axis=1

---
### dataframe 的 count ，如果要是有筛选的话，得加上列名

---
### dataframe 的 astype ，用字典，改变数据类型为 int 或 str 类型

---
###

---