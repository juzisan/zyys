### 特殊注意

######

word替换使用的是程序Application，不是文档Documents

---

### 替换word文件中的内容：

```python
from win32com.client import Dispatch

msword = Dispatch('Word.Application')
msword.Visible = 1
doc = msword.Documents.Open('文件名')
r_t = ('OldStr', False, False, False, False, False, True, 1, True, 'NewStr', 2)
msword.Selection.Find.Execute(*r_t)
doc.Save()
doc.Close()
msword.Quit()
```

1. 导入包
2. 打开word程序，Dispatch和DispatchEx，区别是Ex在单独的进程里面打开文件
3. 设置word程序可见
4. 打开doc文件
5. 替换的参数太长了，所以分开两句，涉及的 11 个参数说明：

| 位置  | 参数     | 意义                                 |
|-----|--------|------------------------------------|
| 0   | OldStr | 搜索的关键字                             |
| 1   | True   | 区分大小写                              |
| 2   | True   | 完全匹配的单词，并非单词中的部分(全字匹配)             |
| 3   | True   | 使用通配符                              |
| 4   | True   | 同音                                 |
| 5   | True   | 查找单词的各种形式                          |
| 6   | True   | 向文档尾部搜索                            |
| 7   | 1      |                                    |
| 8   | True   | 带格式的文本                             |
| 9   | NewStr | 替换文本                               |
| 10  | 2      | 替换个数(0表示不替换，1表示只替换匹配到的第一个，2表示全部替换) |

6. 程序执行替换，而不是doc文档执行替换，用*号解包
7. 保存文档
8. 关闭文档
9. 关闭程序

---

### 替换word文件中的内容：

```python
import pandas as pd

df = pd.read_excel(io='file_str', sheet_name=0, dtype={'班级': 'str'}, skiprows=1)
```

#####

read_excel 必须要有 sheet_name 参数，0为第一个，None 为所有都读取，读到一个字典里

---

### 增加 word 替换文本函数 r_str

####

执行的时候用 * 号解包