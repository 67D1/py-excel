#coding:utf-8


import xlrd     #导入XLRD
import xlwt     #导入XLWT

##读取原文件部分：
data = xlrd.open_workbook('1.xls')  #打开目标文件（同一目录下 文件名1.XLS）
table = data.sheets()[0]    #读取目标EXCEL表单第一张
nrows = table.nrows     #行数
ncols = table.ncols     #列数


a=[[0 for x in range(nrows)] for y in range(ncols)]     #定义一组二维数组 下标分别为EXCEL的行和列

## A1 = table.row(0)[0].value 调用EXCEL的行格式 注意()和[]

## B3 = table.col(1)[2].value 调用EXCEL的列格式 注意()和[]


for i in range(nrows):
    for u in range(ncols):
        a[i][u]=  table.row(i)[u].value     #给数组a赋值 讲表里的内容赋值给a

##读取比对部分：
data2 = xlrd.open_workbook('2.xls')
table2 = data2.sheets()[0]
nrows2 = table2.nrows
ncols2 = table.ncols

b=[[0 for x2 in range(nrows)] for y2 in range(ncols)]

for i in range(nrows2):
    for u in range(ncols2):
        b[i][u]=  table2.row(i)[u].value

print b[0][1]


##写入部分
wb = xlwt.Workbook()        #新建一个文件
sh = wb.add_sheet('test',cell_overwrite_ok=True)        #新建一个表单，注意cell_overwrite_ok=True，要求可以重复写入

for x in range(nrows):      #复制原文件内容
    for y in range(ncols):
        sh.write(x,y,a[x][y])

for x in range(nrows2):     #对比文件第一列中是否有重复的，如果有则复制该行中第三列内容到新文件
    for y in range(nrows):
        if b[x][0] == a[y][0] :
            sh.write(y, 3, b[x][3])

wb.save('tttt.xls')     #保存
