# coding:utf-8
import sys

reload(sys)
sys.setdefaultencoding('utf8')
#    __author__ = '郭 璞'
#    __date__ = '2016/8/22'
#    __Desc__ = 测试文件


print '-----s-----------------------------\n'
from xlsx2sql import XlsxTool,Xlsx2sql
tool = XlsxTool()
# table_header_type = ['varchar(100) not null','int(100) ','varchar(255)','varchar(100)']
table_header_type = ['int(100) not null','varchar(255) ','varchar(255)','varchar(255)','varchar(30)','varchar(30)','varchar(100) not null']
release = Xlsx2sql(tool)
release.generate(r'./readout.xlsx',table_header_type,"id",r'./readout.sql')