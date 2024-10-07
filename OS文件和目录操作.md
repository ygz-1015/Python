# 文件和目录操作

本章总结了一些常见的 文件 和目录的操作。

注意：初次学习，本节的内容不需要背。 只需要记得有哪些操作方法，有个印象。在实际工作中需要用到时，再回来查看即可。

## 创建目录

os.makedirs 可以递归的创建目录结构，比如

```python
import os
os.makedirs('tmp/python/fileop',exist_ok=True)
```

会在当前工作目录下面创建 tmp目录，在tmp目录下面再创建 python目录，在Python目录下面再创建fileop目录

`exist_ok=True` 指定了，如果某个要创建的目录已经存在，也不报错

## 删除文件或目录

os.remove 可以删除一个文件，比如

```python
os.remove('sdf.py')
```



shutil.rmtree() 可以递归的删除 某个目录 和 该目录里面所有的子目录和子文件

比如

```python
import shutil
shutil.rmtree('tmp', ignore_errors=True)
```

注意：参数 `ignore_errors` 值设置为 `True` ，表示忽略删除过程中的错误（比如要删除的目录不存在），不会抛出异常。

## 拷贝文件

shutil 模块里面有很多 目录文件操作的函数

拷贝文件，可以使用shutil模块的copyfile函数。

比如

```python
from shutil import copyfile

# 拷贝 d:/tools/first.py 到 e:/first.py
copyfile('d:/tools/first.py', 'e:/first.py')
```

注意，如果拷贝前，e:/first.py 已经存在，则会被拷贝覆盖，所以使用该函数一定要小心。

## 拷贝目录

如果我们要拷贝一个目录里面所有的内容（包括子目录和文件、子目录里面的子目录和文件，等等）到另外一个目录中，可以使用 shutil的copytree函数。

如下所示

```python
from shutil import copytree

# 拷贝 d:/tools/aaa 目录中所有的内容 到 e:/bbb 中
copytree('d:/tools/aaa', 'e:/new/bbb')
```

注意拷贝前， 目标目录必须 `不存在` ，否则会报错。

上面的代码执行前面，如果 e:/new/bbb 已经存在，执行到copytree时，就会报错

上面的代码执行前面，如果 e:/new 这个目录都不存在，执行到copytree时，就会 创建 e:/new 目录，再创建 e:/new/bbb 目录，再拷贝 d:/tools/aaa 目录中所有的内容 到 e:/new/bbb 中。

上面的代码执行前面，如果 e:/new 这个目录存在，但是 e:/new/bbb 不存在，执行到copytree时，就只会 创建 e:/new/bbb ，再拷贝 d:/tools/aaa 目录中所有的内容 到 e:/new/bbb 中。

## 修改文件名、目录名

要修改文件名、目录名，可以使用os模块的rename函数。

比如

```python
import os

# 修改目录名 d:/tools/aaa 为 d:/tools/bbb
os.rename('d:/tools/aaa','d:/tools/bbb')

# 修改文件名 d:/tools/first.py 为 d:/tools/second.py
os.rename('d:/tools/first.py','d:/tools/second.py')
```

注意，Linux 系统上，如果重命名之前 d:/tools/second.py 已经存在，则会被覆盖，所以使用该函数一定要小心。

## 判断文件、目录是否存在

如果我们需要判断一个指定路径的文件或者目录是否存在，可以使用下面的方法

```python
import os
os.path.exists('d:/systems/cmd.exe')
os.path.exists('d:/systems')
```

exists方法返回值为True表示 存在，否则表示不存在。



如果你要判断指定路径是否是文件，可以这样

```python
import os

# 返回值为True 表示是文件
os.path.isfile('d:/systems/cmd.exe')
```



如果你要判断指定路径是否是目录，可以这样

```python
import os

# 返回值为True 表示是目录
os.path.isdir('d:/systems')
```

## 当前工作目录

我们的程序运行时，是有 `当前工作目录` 的。

程序代码中所有的对文件的访问，如果使用相对路径，都是基于这个当前工作目录。

有时候我们需要得到程序的当前工作目录的路径，可以使用

```python
cwd = os.getcwd()
```

如果需要改变当前工作目录到另外的路径，可以使用

```python
os.chdir(path)
```

参数就是 新的当前工作目录 路径地址。

## 递归的遍历目录下面所有的文件

假如我们要获取某个目录中所有的 文件， 包括子目录里面的文件。 可以使用 os库中的walk方法

比如我们要得到某个目录下面所有的子目录 和所有的文件，存放在两个列表中

可以这样写代码

```python
import os

# 目标目录
targetDir = r'd:\tmp\util\dist\check'
files = []
dirs  = []

# 下面的三个变量 dirpath, dirnames, filenames
# dirpath 代表当前遍历到的目录名
# dirnames 是列表对象，存放当前dirpath中的所有子目录名
# filenames 是列表对象，存放当前dirpath中的所有文件名

for (dirpath, dirnames, filenames) in os.walk(targetDir):
    files += filenames
    dirs += dirnames

print(files)
print(dirs)
```

如果要得到某个目录下所有文件的全路径可以这样

```python
import os

# 目标目录
targetDir = r'd:\tmp\util\dist\check'

for (dirpath, dirnames, filenames) in os.walk(targetDir):
    for fn in filenames:
        # 把 dirpath 和 每个文件名拼接起来 就是全路径
        fpath = os.path.join(dirpath, fn)
```

## 得到目录中所有的文件和子目录名

```python
import os

# 目标目录
targetDir = r'd:\tmp\util\dist\check'


files =  os.listdir(targetDir)
print(files)
```

listdir返回的是该目录下面所有的文件和子目录。

如果我们只需要获取目录中所有的文件，或者只需要子目录，可以这样

```python
import os
from os.path import isfile, join,isdir

# 目标目录
targetDir = r'd:\tmp\util\dist\check'

# 所有的文件
print([f for f in os.listdir(targetDir) if isfile(join(targetDir, f))])

# 所有的目录
print([f for f in os.listdir(targetDir) if isdir(join(targetDir, f))])
```

## 得到目录中指定扩展名的文件和子目录

可以使用glob库

```python
import glob
exes = glob.glob(r'd:\tmp\*.txt')

print(exes)
```