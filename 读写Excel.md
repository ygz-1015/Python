# 读写Excel

我们经常需要 导入数据到Excel、从Excel导出数据、对Excel中的数据进行处理。

如果 要处理的数据量很大，人工操作非常费时间。

我们可以通过Python程序，自动化Excel的数据处理，帮我们节省大量的时间。



## xlrd - 读取Excel数据

如果我们只是要 `读取` Excel文件里面的数据进行处理，可以使用 `xlrd` 这个库。

首先我们安装xlrd库，执行下面的命令

```
pip install xlrd==1.2.0
```

注意：xlrd 新版本只支持 xls 格式，所以我们这里指定安装 1.2.0 老版本，可以支持xlsx格式。

这个文件里面有 3 个表单，分别记录了2018、2017、2016年 的月收入，如下所示

![img](D:\software\Typora\复制的图片文件\tut_20230417202725_69.png)

如果我们想用程序统计 2016、2017、2018 三年所有月收入的总和，但是不要包含 `打星号` 的那些月份。

怎么做？

一步步来，我们先学会如何用Python程序读取Excel单元格中的内容。



xlrd 库里面的 open_workbook 函数打开Excel文件，并且返回一个 `Book对象` ，这个对象代表打开的 Excel 文件。

可以通过这个Book对象得到Excel文件的很多信息，比如 获取 Excel 文件中表单(sheet) 的数量 和 所有 表单(sheet) 的名字。

我们可以用如下代码，读取 该文件中表单的数量和名称：

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

print(f"包含表单数量 {book.nsheets}")
print(f"表单的名分别为: {book.sheet_names()}")
```



要读取某个表单里单元格中的数据，必须要先获取 `表单（sheet）对象` 。

可以根据表单的索引 或者 表单名获取 表单对象，使用如下对应的方法

```python
# 表单索引从0开始，获取第一个表单对象
book.sheet_by_index(0)

# 获取名为2018的表单对象
book.sheet_by_name('2018')

# 获取所有的表单对象，放入一个列表返回
book.sheets()
```

表单对象所有属性和方法的描述， 可以[点击这里，查看官方文档](https://xlrd.readthedocs.io/en/latest/api.html#xlrd-sheet)



获取了表单对象后，可以根据其属性得到：

```
表单行数（nrows）
列数（ncols）
表单名（name）
表单索引（number）
```

如下

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

sheet = book.sheet_by_index(0)
print(f"表单名：{sheet.name} ")
print(f"表单索引：{sheet.number}")
print(f"表单行数：{sheet.nrows}")
print(f"表单列数：{sheet.ncols}")
```



获取了表单对象后，可以使用 `cell_value` 方法，参数为行号和列号，读取指定单元格中的文本内容。如下所示：

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

sheet = book.sheet_by_index(0)

# 行号、列号都是从0开始计算
print(f"单元格A1内容是: {sheet.cell_value(rowx=0, colx=0)}")
```

运行结果输出

```
单元格A1内容是: 月份
```

还可以使用 `row_values` 方法，参数为行号，读取指定行所有单元格的内容，存放在一个列表中返回。

如下所示：

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

sheet = book.sheet_by_index(0)

# 行号、列号都是从0开始计算
print(f"第一行内容是: {sheet.row_values(rowx=0)}")
```

运行结果输出

```
第一行内容是: ['月份', '收入']
```

还可以使用 `col_values` 方法，参数为列号，读取指定列所有单元格的内容，存放在一个列表中返回。

如下所示：

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

sheet = book.sheet_by_index(0)

# 行号、列号都是从0开始计算
print(f"第一列内容是: {sheet.col_values(colx=0)}")
```

运行结果输出

```
第一列内容是: ['月份', 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0]
```

可以看出，数字以小数的形式返回了。



有了这些方法，我们就可以完成一些数据处理任务了。比如我们要计算 2017年 全年的收入就可以这样

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

sheet = book.sheet_by_name('2017')

# 收入在第2列
incomes = sheet.col_values(colx=1,start_rowx=1)

print(f"2017年收入为: {sum(incomes)}")
```



那么我们怎么在汇总收入中，去掉包含星号的月份收入呢？

就需要我们查出哪些月份是带星号的，不要统计在内。

参考下面的代码

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

sheet = book.sheet_by_name('2017')

# 收入在第2列
incomes = sheet.col_values(colx=1,start_rowx=1)

print(f"2017年账面收入为: {int(sum(incomes))}")

# 去掉包含星号的月份收入
toSubstract = 0
# 月份在第1列
monthes = sheet.col_values(colx=0)

for row,month in enumerate(monthes):
    if type(month) is str and month.endswith('*'):
        income = sheet.cell_value(row,1)
        print(month,income)
        toSubstract += income

print(f"2017年真实收入为: {int(sum(incomes)- toSubstract)}")
```

最后，要得到3年的收入，就要获取所有的sheet对象，采用上面的计算方法，最后把收入相加。

如下所示：

```python
import xlrd

book = xlrd.open_workbook("income.xlsx")

# 得到所有sheet对象
sheets = book.sheets()

incomeOf3years = 0
for sheet in sheets:
    # 收入在第2列
    incomes = sheet.col_values(colx=1,start_rowx=1)
    # 去掉包含星号的月份收入
    toSubstract = 0
    # 月份在第1列
    monthes = sheet.col_values(colx=0)

    for row,month in enumerate(monthes):
        if type(month) is str and month.endswith('*'):
            income = sheet.cell_value(row,1)
            print(month,income)
            toSubstract += income

    actualIncome = int(sum(incomes)- toSubstract)
    print(f"{sheet.name}年真实收入为: {actualIncome}")
    incomeOf3years += actualIncome

print(f'全部收入为{incomeOf3years}')
```

## openpyxl

新建Excel，写入数据

[点击这里，边看视频讲解，边学习以下内容](https://www.bilibili.com/video/av87258657?p=3)

`xlrd` 只能读取Excel内容，如果你要 `创建` 一个新的Excel并 `写入` 数据，可以使用 `openpyxl` 库。

`openpyxl` 库既可以读文件、也可以写文件、也可以修改文件。

但是，openpyxl 库不支持老版本 Office2003 的 `xls` 格式的Excel文档，如果要读写xls格式的文档，可以使用 Excel 进行相应的格式转化。

执行 `pip install openpyxl` 安装该库

[点击这里，查看openpyxl参考文档](https://openpyxl.readthedocs.io/en/stable/)

下面的代码，演示了 openpyxl 的 一些基本用法。

```python
import openpyxl

# 创建一个Excel workbook 对象
book = openpyxl.Workbook()

# 创建时，会自动产生一个sheet，通过active获取
sh = book.active

# 修改当前 sheet 标题为 工资表
sh.title = '工资表'

# 保存文件
book.save('信息.xlsx')

# 增加一个名为 '年龄表' 的sheet，放在最后
sh1 = book.create_sheet('年龄表-最后')

# 增加一个 sheet，放在最前
sh2 = book.create_sheet('年龄表-最前',0)

# 增加一个 sheet，指定为第2个表单
sh3 = book.create_sheet('年龄表2',1)

# 根据名称获取某个sheet对象
sh = book['工资表']

# 给第一个单元格写入内容
sh['A1'] = '你好'

# 获取某个单元格内容
print(sh['A1'].value)

# 根据行号列号， 给第一个单元格写入内容，
# 注意和 xlrd 不同，是从 1 开始
sh.cell(2,2).value = '白月黑羽'

# 根据行号列号， 获取某个单元格内容
print(sh.cell(1, 1).value)

book.save('信息.xlsx')
```



下面的示例代码 将 保存在字典中的年龄表的内容 写入到excel文件中

```python
import openpyxl

name2Age = {
    '张飞' :  38,
    '赵云' :  27,
    '许褚' :  36,
    '典韦' :  38,
    '关羽' :  39,
    '黄忠' :  49,
    '徐晃' :  43,
    '马超' :  23,
}

# 创建一个Excel workbook 对象
book = openpyxl.Workbook()

# 创建时，会自动产生一个sheet，通过active获取
sh = book.active

sh.title = '年龄表'

# 写标题栏
sh['A1'] =  '姓名'
sh['B1'] =  '年龄'

# 写入内容
row = 2

for name,age in name2Age.items():
    sh.cell(row, 1).value = name
    sh.cell(row, 2).value = age
    row += 1

# 保存文件
book.save('信息.xlsx')
```



如果你的数据在一个列表或者元组中，可以使用append方法在sheet的末尾添加新行，写入数据，比如

```python
import openpyxl

name2Age = [
    ['张飞' ,  38 ] ,
    ['赵云' ,  27 ] ,
    ['许褚' ,  36 ] ,
    ['典韦' ,  38 ] ,
    ['关羽' ,  39 ] ,
    ['黄忠' ,  49 ] ,
    ['徐晃' ,  43 ] ,
    ['马超' ,  23 ]
]

# 创建一个Excel workbook 对象
book = openpyxl.Workbook()
sh = book.active
sh.title = '年龄表'

# 写标题栏
sh['A1'] =  '姓名'
sh['B1'] =  '年龄'

for row in name2Age:
    # 添加到下一行的数据
    sh.append(row)

# 保存文件
book.save('信息.xlsx')
```

### 打开现有的Excel文件

可以使用 `openpyxl.load_workbook()` 方法打开现有的Excel文件，可以打开 `xlsx xlsm xltx xltm` 这些格式，不能打开老的 `xls` 格式

```python
from openpyxl import load_workbook

book = load_workbook(filename='income.xlsx')
sheet = book.active
```

### 读取所有单元格内容

如果我们只想获取文件中单元格的内容，可以通过 `Worksheet.values` 属性，这是产生每行数据的 generator

比如：

```python
from openpyxl import load_workbook

book = load_workbook(filename='income.xlsx')
values = book.active.values 

for row in values: 
    # row 是tuple, 包含该行每个cell数据
    for value in row:
        print(value)
```

### 直接获取某个单元格

可以这样 `sh['C3']` 根据 列名，行号 直接访问某个单元格,

也可以这样 `sh.cell(row=3, column=3)` 根据 行号，列号， 直接访问某个单元格。 注意：行号列号都是从1开始，不是从0开始

这2种方法，返回的是都是单元格 `Cell` 对象，然后可以根据Cell对象属性获取各种信息。

比如：

```python
from openpyxl import load_workbook

wb = load_workbook(filename='d:/1.xlsx')
sh = wb.active

# 可以这样访问 C3 单元格
cell = sh['C3']
# 也可以这样访问
cell = sh.cell(row=3, column=3)

print(cell.value) # 单元格内容
```

### 获取多个单元格

要获取某个区域的所有单元格，可以直接通过 左上，右下 两个单元格坐标获取

比如 `sh['A1':'C3']` 返回一个tuple， 里面依次存放 `A1, B1, C1, A2, B2, C2, A3, B3, C3` Cell 对象

```python
from openpyxl import load_workbook

wb = load_workbook(filename='d:/1.xlsx')
sh = wb.active

print(sh['A1':'C3'])
```



要获取某列所有单元格对象，直接根据列号字母，比如 `sh['C']` 返回一个tuple，里面存放第3列所有 Cell 对象

要获取多列，用冒号隔开，比如 `sh['C:D']` 返回一个tuple，里面存放第3，4 列所有 Cell 对象



要获取某行所有单元格对象，直接根据行号数字，比如 `sh['2']` 返回一个tuple，里面存放第2行所有 Cell 对象

要获取多列，用冒号隔开，比如 `sh['2:3']` 返回一个tuple，里面存放第2，3 行所有 Cell 对象

### 修改单元格内容

[点击这里，边看视频讲解，边学习以下内容](https://www.bilibili.com/video/av87258657?p=4)

可以这样 `修改` Excel 单元格内容

```python
import openpyxl

# 加载 excel 文件
wb = openpyxl.load_workbook('income.xlsx')

# 得到sheet对象
sheet = wb['2017']

sheet['A1'] = '修改一下'

## 指定不同的文件名，可以另存为别的文件
wb.save('income-1.xlsx')
```

### 插入行、插入列

sheet 对象的 `insert_rows` 和 `insert_cols` 方法，分别用来插入 `行` 和 `列`， 比如

```python
import openpyxl

wb = openpyxl.load_workbook('income.xlsx')
sheet = wb['2018']

# 在第2行的位置插入1行
sheet.insert_rows(2)

# 在第3行的位置插入3行
sheet.insert_rows(3,3)

# 在第2列的位置插入1列
sheet.insert_cols(2)

# 在第2列的位置插入3列
sheet.insert_cols(2,3)

## 指定不同的文件名，可以另存为别的文件
wb.save('income-1.xlsx')
```

### 删除行、删除列

sheet 对象的 `delete_rows` 和 `delete_cols` 方法，分别用来删除 `行` 和 `列`， 比如

```python
import openpyxl

wb = openpyxl.load_workbook('income.xlsx')
sheet = wb['2018']

# 在第2行的位置删除1行
sheet.delete_rows(2)

# 在第3行的位置删除3行
sheet.delete_rows(3,3)

# 在第2列的位置删除1列
sheet.delete_cols(2)

# 在第3列的位置删除3列
sheet.delete_cols(3,3)

## 指定不同的文件名，可以另存为别的文件
wb.save('income-1.xlsx')
```

### 文字 颜色、字体、大小

单元格里面的 `样式风格` （包括 颜色、字体、大小、下划线 等） 都是通过 `Font` 对象设定的

```python
import openpyxl
# 导入Font对象 和 colors 颜色常量
from openpyxl.styles import Font,colors

wb = openpyxl.load_workbook('income.xlsx')
sheet = wb['2018']

# 指定单元格字体颜色，
sheet['A1'].font = Font(color=colors.RED, #使用预置的颜色常量
                        size=15,    # 设定文字大小
                        bold=True,  # 设定为粗体
                        italic=True # 设定为斜体
                        )

# 也可以使用RGB数字表示的颜色
sheet['B1'].font = Font(color="981818")

# 指定整行 字体风格， 这里指定的是第3行
font = Font(color="981818")
for y in range(1, 100): # 第 1 到 100 列
    sheet.cell(row=3, column=y).font = font

# 指定整列 字体风格， 这里指定的是第2列
font = Font(bold=True)
for x in range(1, 100): # 第 1 到 100 行
    sheet.cell(row=x, column=2).font = font

wb.save('income-1.xlsx')
```

### 背景色

```python
import openpyxl
# 导入Font对象 和 colors 颜色常量
from openpyxl.styles import PatternFill

wb = openpyxl.load_workbook('income.xlsx')
sheet = wb['2018']

# 指定 某个单元格背景色
sheet['A1'].fill = PatternFill("solid", "E39191")

# 指定 整行 背景色， 这里指定的是第2行
fill = PatternFill("solid", "E39191")
for y in range(1, 100): # 第 1 到 100 列
    sheet.cell(row=2, column=y).fill = fill

wb.save('income-1.xlsx')
```

### 插入图片

下面是插入图片的代码

```python
import openpyxl
from openpyxl.drawing.image import Image

wb = openpyxl.load_workbook('income.xlsx')
sheet = wb['2018']

# 在第1行，第4列 的位置插入图片
sheet.add_image(Image('1.png'), 'D1')

## 指定不同的文件名，可以另存为别的文件
wb.save('income-1.xlsx')
```

## Excel COM接口

在Windows平台上，还可以通过 Excel 应用的 COM 接口 来对Excel进行操作。

这个方法相当于使用Python程序 通过 Excel应用 自己去修改，当然没有任何的副作用

而且能够实现一些特殊的功能，比如 自动打印Excel、合并单元格 等。

COM接口的特点是：打开文件快，读写速度慢。

使用 Excel COM 接口 打开 `超大` Excel文件 比上面的两个库 要快很多。因为Excel程序本身的优化，可以部分加载，而上面的两个库是全部先读入内存。

如果你只是从 大Excel文件中 读取或修改少量数据，Excel COM 接口会快很多。

但是，如果你要读取大Excel中的大量数据，不要使用 COM接口，会非常的慢。



使用 Excel COM 接口，首先需要安装pywin32库，在命令行窗口输入如下命令：

```
pip install pywin32
```

比如可以这样修改

```python
import win32com.client
excel = win32com.client.Dispatch("Excel.Application")

# excel.Visible = True     # 可以让excel 可见

# 这里填写要修改的Excel文件的绝对路径
workbook = excel.Workbooks.Open(r"d:\tmp\income1.xlsx")

# 得到 2017 表单
sheet = workbook.Sheets('2017')

# 修改表单第一行第一列单元格内容
# com接口，单元格行号、列号从1开始
sheet.Cells(1,1).Value="你好"

# 保存内容
workbook.Save()

# 关闭该Excel文件
workbook.Close()

# excel进程退出
excel.Quit()

# 释放相关资源
sheet = None
book = None
excel.Quit()
excel = None
```

运行一下可以发现Excel内容也能修改。

关于使用com接口操作Excel具体细节， 可以[点击这里，参考微软官方文档](https://docs.microsoft.com/en-us/visualstudio/vsto/excel-object-model-overview?view=vs-2019)

COM 接口、xlrd库 打开数据文件 和 读取数据的性能对比，大家可以参考下面这段代码。

```python
import time

def byCom():
    t1 = time.time()
    import win32com.client
    excel = win32com.client.Dispatch("Excel.Application")

    # excel.Visible = True     # 可以让excel 可见
    workbook = excel.Workbooks.Open(r"h:\tmp\ruijia\数据.xlsx")

    sheet = workbook.Sheets(2)

    print(sheet.Cells(2,15).Value)
    print(sheet.UsedRange.Rows.Count)  #多少行

    t2 = time.time()
    print(f'打开: 耗时{t2 - t1}秒')

    total = 0
    for row in range(2,sheet.UsedRange.Rows.Count+1):
        value = sheet.Cells(row,15).Value
        if type(value) not in [int,float]:
            continue
        total += value

    print(total)

    t3 = time.time()
    print(f'读取数据: 耗时{t3 - t2}秒')


def byXlrd():
    t1 = time.time()
    import xlrd

    # 加载 excel 文件
    srcBook = xlrd.open_workbook("数据.xlsx")
    sheet = srcBook.sheet_by_index(1)

    print(sheet.cell_value(rowx=1,colx=14))
    print(sheet.nrows) #多少行

    t2 = time.time()
    print(f'打开: 耗时{t2 - t1}秒')

    total = 0
    for row in range(1,sheet.nrows):
        value = sheet.cell_value(row, 14)
        if type(value) == str:
            continue
        total += value

    print(total)

    t3 = time.time()
    print(f'读取数据: 耗时{t3 - t2}秒')

byCom()
byXlrd()
```