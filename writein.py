# coding:utf-8
import sys

reload(sys)
sys.setdefaultencoding('utf8')
#    __author__ = '郭 璞'
#    __date__ = '2016/8/20'
#    __Desc__ = 将excel上的数据读入到数据库

import xlrd
import MySQLdb

# 读取EXCEL中的数据到数据库中
workbook = xlrd.open_workbook(r'./example.xlsx')
table = workbook.sheet_by_index(0)

table_header = []
table_body = []
table_rows = table.nrows
table_cols = table.ncols
# 获取excel的表头信息，将作为存储与数据库中的字段
table_header = table.row_values(0)
print '字段信息：\n',
str = ''
for item in table_header:
    str += item
    str += "    "
print '\n',str
# 获取excel中所有数据
def getDataByRow(table,table_rows,table_body):
    for row_number in range(1,table_rows):
        table_body.append(table.row_values(row_number))
    return table_body

table_body = getDataByRow(table=table,table_rows=table_rows,table_body=table_body)
for item in table_body:
    row = ''
    for i in item:
        row =row + u'%s'%i
        row += "    "
    print row

table_header_type = ['varchar(100)','int(100)','varchar(255)','varchar(100)']

# ------------------------------开始处理数据库相关
# 判断当前项是否为数字，整形或者是浮点型的数字均可以判断
def isnum(protype):
    import re
    # 调用正则
    reg = re.compile(r'^[-+]?[0-9]+?\.[0-9]+?$')
    result = reg.match(protype)
    if result:
        return True
    else:
        return False

# 将一个包含任意类型的数组，转化成可以作为插入数据库的值的串
def arr2str(arr):
    fields = ''
    for item in range(len(arr)):
        arr_item =u'%s'%arr[item]
        if isnum(arr_item) or arr_item.isdigit():
            pass
        else:
            arr_item ="'"+u'%s'%arr_item+"'"
        fields += arr_item + ","
    return fields.rstrip(',')

conn = MySQLdb.connect('localhost','root','mysql','test')
cursor = conn.cursor()
print cursor
# 创建数据库表
def createTable(cursor,table_name,table_header,table_header_type,primary_key=None):
    sql = "create table "+u'%s'%table_name+"("
    for item in range(len(table_header)):
        sql = sql + u'%s'%table_header[item] + " "+ u'%s'%table_header_type[item]+" ,"
    sql += "primary key("+u'%s'%primary_key+")"
    sql +=");"
    print sql
    # cursor.execute(sql)


# 将数据存储进数据库
def storeData(cursor,table_name,table_header,table_body):
    for row in range(len(table_body)):
        values = arr2str(table_body[row])

        fields = ""
        for item in range(len(table_header)):
            if (item+1) == len(table_header):
                fields +=(u'%s'%table_header[item])
            else:
                fields += (u'%s' % table_header[item] + ",")
        fields.rstrip(",")
        sql = 'insert into '+ table_name+'('+u'%s'%fields+')' + ' values('+u'%s'%values+');'
        print sql
        # cursor.execute(sql)


createTable(cursor,'example',table_header,table_header_type,table_header[0])
storeData(cursor,'example',table_header,table_body)




