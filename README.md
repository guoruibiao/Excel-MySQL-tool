# Excel-MySQL-tool
:sparkles:Excel一键导入MySQL数据库，MySQL数据库一键读取到Excel表格。

## 依赖
- Python2.7.11 环境
- xlrd
- xlwt
- MySQLdb

## mysql-->>excel

- 首先将`readout.py`导入到你的工作目录下
 - `host`                数据库访问链接，远程的使用IP也是可以的
 - `user`                数据库用户名
 - `password`         数据库密码
 - `db`                    要操作的数据库名称
 - `output_file`    excel文件输出完整路径

- 然后使用如下的调用语句即可。
```
# 结果测试
if __name__ == "__main__":
    export('localhost','root','mysql','test','datetest',r'datetest.xlsx')
    
```

## excel -->> MySQL

- 首先也是导入`xlsx2sql.py`。
- 然后就可以使用如下的语句来实现咯。
```
from xlsx2sql import XlsxTool,Xlsx2sql
tool = XlsxTool()
table_header_type = ['int(100) not null','varchar(255) ','varchar(255)','varchar(255)','varchar(30)','varchar(30)','varchar(100) not null']
release = Xlsx2sql(tool)
release.generate(r'./readout.xlsx',table_header_type,"id",r'./readout.sql')

```


## 案例展示
## 图片展示

由于数据库输出到excel比较简单，这里就简单的展示一下excel转汉城数据库的案例吧。

- excel原始内容
![Excel数据源](http://img.blog.csdn.net/20160822214106332)


- 运行完刚才的示例代码后会在当前目录下生成一个readout.sql文件，复制里面的内容到数据库中。
![输出结果](http://img.blog.csdn.net/20160822214406226)
