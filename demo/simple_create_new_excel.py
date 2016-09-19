# -*- coding=utf-8 -*-

from expy.excel import Excel

# 新建一个excel，并写入一些用户信息。
# 支持链式操作
# create a new excel file and write some users infomations
# Support chain operation

Excel().title(["name", "age"]).write(["ben", 20]).write(2, 0, "nolly").write(2, 1, 18).save()


# the same to:
# 上面的实例等同于：

# excel = Excel()
# excel = title(["name", "age"])
# excel = write(["ben", 20])
# excel = write(2, 0, "nolly")
# excel = write(2, 1, 18)
# excel = save()
