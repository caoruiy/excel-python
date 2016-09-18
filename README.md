# excel-python
A simple API of read and write excel  file based on xlrd, xlwt and xlutils
> 该模块封装了xlrd，xlwt，xlutils的操作，提供了更加简洁的API来读写Excel文件（中文说明在本文最后，[点此跳转][china]）
You can use this package to read and write Excel file like this:

## APIs

First of all, you should instantiate the Excel object
you can do like this:
```
from expy.excel import Excel
excel = Excel()

excel = Excel( file="../../excel/data.xlsx", sheet="sheet1", rebuild=True)

excel = Excel( "../../excel/data.xlsx", "sheet1", True)
```

parameters:
* _file_ : which file you want to open or create. By default, the _file_ is generated in the following format: **%Y%m%d-%H%M%S** eg: 20160918-125622.xlsx
* _sheet_ : which sheet in the file you opened or created you want to write. Default using the first sheet
* _rebuild_ : if you want to re-create the file when this file exist. The default value is _False_

### how to read
you can read the excel file like this:
```
sheet = excel.read()

sheet = excel.read(sheet='sheet1')

sheet = excel.read("sheet1")
```

Open the first sheet by default, also, You can also specify the name of the sheet by _sheet_
You can use all the properties and methods of the object which you use _xlrd_ open a file

Most commonly used method and properties like:
```
sheet.cell(1, 0).value

# get all rows
sheet.nrows

# get all cols
sheet.ncols

#get sheet name
sheet.name
```
Of course, you can use:
```
sheet(1,0)
```
get the date at (1, 0), the same to _sheet.cell(1, 0).value_

### how to write
you can write like this:
```
excel.write(row=1, col=0, content='ben', sheet='sheet1')
```
parameters:
* _row_, _col_ : row and column you write
* _content_ : the content you want to write
* _sheet_ : which sheet you want to write

Some simple use like:
```
excel.write(1, 0, 'ben')
excel.write(1, 0, 'ben', 'sheet1')
```
You can add a line to the end of the file by the following method:
```
excel.write(['ben'], sheet='sheet1')
```

If you want to write a title in the first line of the file,you can use _title_ function like this:
```
excel.title(["name", "age"])
```
Sometimes you need to insert a title in the middle of the file, You can insert a row of title Compulsive like this:
```
excel.title(["name", "age"], absTitle=True)
```

### save
Through the above method, all of the data is written in a list(a list Maintained by The _Excel_ object), and there is no real write to the file,

you can use _save_ method write all the data to the file, like this:
```
excel.save()
```

### Error
If an error occurs, You can view the error message like this:
```
excel.error()
```

### Some available properties

Get file name:
```
excel.file
```

Get Error code:
```
excel.code
```

get Error message:
```
excel.msg
```

# china-apis
# 中文帮助文档

## APIs

首先你需要实例化 **Excel** 对象，你可以这样做：
```
from expy.excel import Excel
excel = Excel()

excel = Excel( file="../../excel/data.xlsx", sheet="sheet1", rebuild=True)

excel = Excel( "../../excel/data.xlsx", "sheet1", True)
```

参数说明:
* _file_ : 你想打开或者创建的文件名. 默认情况下, _file_ 会按照下面这样的格式自动生成一个文件名: **%Y%m%d-%H%M%S** 例如: 20160918-125622.xlsx
* _sheet_ : 你想操作的sheet名. 默认情况下使用第一张sheet
* _rebuild_ : 当写文件时，如果该文件存在，是否重写该文件（意思就是说是否删除原文件并新建一个空的文件）. 默认值为： _False_，不重建

### 读文件
你可以这样读取文件内容:
```
# 默认情况下使用实例化时提供的_sheet_, 如果实例化时没有提供_sheet_, 默认使用第一张sheet
sheet = excel.read()

# 读取指定的sheet内容
sheet = excel.read(sheet='sheet1')
# 简写
sheet = excel.read("sheet1")
```

默认情况下打开第一张sheet, 当然, 你也可以通过 _sheet_ 参数指定sheet名称，

sheet对象包含了使用xlrd对象打开Excel文件后提供的所有属性和方法，

下面列举一些常用的方法和属性：

```
# 获取指定行列的值
sheet.cell(1, 0).value

# 获取打开表的总行号
sheet.nrows

# 获取打开表的总列号
sheet.ncols

# 获取打开表的名称
sheet.name
```
额外的，该模块还提供了一个简便方法，如下:
```
sheet(1,0)
```
用来获取指定行列的值，该方法等同与： _sheet.cell(1, 0).value_

### 写入文件
你可以像下面那样写入内容:
```
excel.write(row=1, col=0, content='ben', sheet='sheet1')
```
参数:
_row_, _col_ : 写入的行号，列号
_content_ : 写入的内容
_sheet_ : 写入的sheet名称

简写:
```
excel.write(1, 0, 'ben')
excel.write(1, 0, 'ben', 'sheet1')
```
你还可以写入一行数据:
```
excel.write(['ben'], sheet='sheet1')
```
上面的用法相当于在文件末尾追加了一行数据

如果你新建了一个文件，并想在文件第一行写入一行标题,你可以使用 _title_ 方法:
```
excel.title(["name", "age"])
```
如果文件中已经存在一些数据了，你仍然想强制的追加一行标题来写入点别的内容,你可以像下面那样强制的追加一行标题:
```
excel.title(["name", "age"], absTitle=True)
```

### 保存
上面的方法，所以的数据只是写入到了一个由_Excel_对象维护的list中, 并没有真正的写入文件,

调用_save_方法把内容正在的写入文件:
```
excel.save()
```

### 错误
如果在使用时发生了错误，你可以这样查看错误信息:
```
excel.error()
```

### 一些可以的属性

获取文件名:
```
excel.file
```

获取错误码:
```
excel.code
```

获取错误信息:
```
excel.msg
```

[china]: #china-apis '中文帮助文档'