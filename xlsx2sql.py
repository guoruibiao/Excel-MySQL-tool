# coding:utf-8
import sys
import os
reload(sys)
sys.setdefaultencoding('utf8')
#    __author__ = '郭 璞'
#    __date__ = '2016/8/22'
#    __Desc__ = 根据给定的excel表格，一键导出为可执行的sql语句，用来给用户一个查看以及修改的机会，减少出错的可能性。
import xlrd

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

# 根据excel的路径生成数据库表的表名称
def get_table_name_from_excelpath(excel_path):
    if excel_path == None:
        print u'%s'%("未指定Excel的完整路径\n")
        exit(0)
    else:
        excel_name = excel_path.split('/')[-1].strip('.xlsx')
        return excel_name

class XlsxTool():

    # 初始化相关参数
    def __init__(self):
        self.table_header = []
        self.xlsx_name = None
        self.table_header_type = []
        self.table_body = []
        self.table = None
        self.table_rows = None
        self.table_cols = None
        self.table_name = None
        self.sql_create_table = ""
        self.sql_data_inflate = []

    # 从给定的路径中读取excel数据表，并将数据读取到init之后的变量中，方便下文的使用
    def readxlsx(self,xlsx_name=None):
        if type(xlsx_name)==None:
            print u'%s'%("excel表名称为空，请检查后重试！\n")
            exit(0)
        else:
            self.xlsx_name = xlsx_name
            self.table_name = get_table_name_from_excelpath(xlsx_name)
        workbook = xlrd.open_workbook(self.xlsx_name)
        self.table = workbook.sheet_by_index(0)
        self.table_rows = self.table.nrows
        self.table_cols = self.table_cols
        self.table_header = self.table.row_values(0)

        for row in range(1,self.table_rows):
            self.table_body.append(self.table.row_values(row))

    # 根据传进来的表的头的类型，生成数据库中的建表字段。注意传入的类型一定要符合数据库的语段要求
    def create_table(self,table_header_type,primary_key_name=None):
        if table_header_type == None:
            print u'%s'%("数据库表头类型为必须项，且为按照数据库语句规则的列表！\n")
            exit(0)
        else:
            self.table_header_type = table_header_type
        # 创建数据库表
        sql = "create table " + u'%s' % self.table_name + "("
        for item in range(len(self.table_header)):
            sql = sql + u'%s' % self.table_header[item] + " " + u'%s' % self.table_header_type[item] + " ,"
        sql += "primary key(" + u'%s' % primary_key_name + ")"
        sql += ");"
        self.sql_create_table = sql

    # 将excel表格中的所有的数据，生成insert 语句，为接下来的向数据库中导入数据做准备
    def create_insert_sqls(self):

        for row in range(len(self.table_body)):
            values = arr2str(self.table_body[row])

            fields = ""
            for item in range(len(self.table_header)):
                if (item + 1) == len(self.table_header):
                    fields += (u'%s' % self.table_header[item])
                else:
                    fields += (u'%s' % self.table_header[item] + ",")
            fields.rstrip(",")
            sql = 'insert into ' + self.table_name + '(' + u'%s' % fields + ')' + ' values(' + u'%s' % values + ');'
            self.sql_data_inflate .append(sql)


class Xlsx2sql():

    def __init__(self,xlsxtool):
        self.xlsxtool = xlsxtool

    def generate(self,excel_file_path,table_header_type,primary_key_name,output_sql_file):
        self.xlsxtool.readxlsx(excel_file_path)
        self.xlsxtool.create_table(table_header_type=table_header_type,primary_key_name=primary_key_name)
        self.xlsxtool.create_insert_sqls()
        generate_data = self.xlsxtool.sql_create_table
        for item in range(len(self.xlsxtool.sql_data_inflate)):
            generate_data+=self.xlsxtool.sql_data_inflate[item]
            print self.xlsxtool.sql_data_inflate[item]
        file = open(output_sql_file,'wb')
        file.write(generate_data)
        file.close()
        print "尊敬的用户，%s Excel表中的数据已成功转换成,生成文件的路径为：%s，待导入数据库执行的sql文件，" \
              "请检查语法合格后导入数据库！\n"%(excel_file_path,output_sql_file)
