# -*- coding=utf-8 -*-

from expy.excel import Excel

# 读写Excel示例

# 实例化一个excel对象
# 该对象支持两个参数:
# file:指定打开或写入的文件, 可以包含任意文件路径，路径中不存在的文件夹会自动创建
# sheet:指定打开或写入文件的哪一个sheet表, 不指定默认写入sheet1表，默认读取第一个sheet表
# 以下格式都可以
excel = Excel()
excel = Excel( file="../../excel/data.xlsx", sheet="sheet1")
excel = Excel( file="../../excel/data.xlsx")
excel = Excel( "../../excel/data.xlsx", "sheet1")

# 写入标题
# 以一个list列表的方式新增标题, 在写入文件内容之前，你需要先调用该方法，才能保证标题写在第一行，因为该方法并不强制把方法放在第一行
# 该方法默认只在新建文件时才执行
excel.title([u"姓名", u"年龄",u"性别"])
# 如果文件已经存在内容，你还想新增标题，可以如下强制写u标题信息
excel.title([u"姓名", u"年龄",u"性别"], absTitle=True)

# 写入内容
# write方法用于写入数据，write参数如下：
# write(row, col, content, sheet)
# row 行号
# col 列号
# content 写入的内容
# sheet   写入到哪一张表
# 以下方法都可以
excel.write(row=1, col=0, content=u'小明', sheet='sheet1')
excel.write(row=1, col=0, content=u'小明')
excel.write(1, 1, 'text')
# 还有一种简便的写入方式，写入一行数据：
excel.write([u"小明", 19, u"男"])
# 也可以多张表一起写入
# 下面的代码表示：
# 在 A 表 0行0列 写入 你好
# 在 B 表 0行0列 写入 你好
# 在 B 表 0行1列 写入 你好
# 第三个write没有指定sheet参数，默认使用最后指定的sheet
excel.write(0,0,"你好",'A').write(0,0,"你好",'B').write(0,1,"你好")

# 保存时只要调用save方法，就会写入到文件
excel.write(0,0,"你好",'A').write(0,0,"你好",'B').write(0,1,"你好").save()
# 指定特定的sheet
excel.write([u"小明", 19, u"男"], sheet="A")
# -------------------------------------------------------------------------------

# 读取文件
# 上述实例化的excel对象还有一个read方法，用于读取文件
# read方法有一个参数sheet,用于读取指定表，不指定时使用实例化时指定的表名，如果都没有指定，默认打开第一个表
# 以下调用方法都可以
sheet = excel.read()
sheet = excel.read(sheet='sheet1')
sheet = excel.read("sheet1")

# 获取的sheet对象支持xlrd中获取的sheet对象所以的参数
sheet.nrows # 行总数
sheet.ncols # 列总数
sheet.name # 表名

# 除此之外，提供一个简单的读取数据的方法
# 读取0行1列的数据
sheet(0,1)

# Excel对象有以下几个属性可以使用：

# 获取完成的文件名，包括文件路径
excel.file

# 发生错误时的，错误码和错误信息
excel.code # 0 为没有发生错误
excel.msg
# error方法返回包含code和msg的错误字典，一般用不到
excel.error()

