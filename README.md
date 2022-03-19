# DealTemperatureExcel
This is a windows exe software base on python. It is used to deal the excel which record indoor temperature


(1). 将 python 文件转换成 EXE 文件的命令：

pyinstaller -F wendu.py --hidden-import pyexcel_xls --hidden-import pyexcel_io.readers.csvr --hidden-import pyexcel_io.readers.csvz --hidden-import pyexcel_io.readers.tsv --hidden-import pyexcel_io.readers.tsvz --hidden-import pyexcel_io.writers.csvw --hidden-import pyexcel_io.writers.csvz --hidden-import pyexcel_io.writers.tsv --hidden-import pyexcel_io.writers.tsvz --hidden-import pyexcel_io.database.importers.django --hidden-import pyexcel_io.database.importers.sqlalchemy --hidden-import pyexcel_io.database.exporters.django --hidden-import pyexcel_io.database.exporters.sqlalchemy


(2). 使用方法

    1）. Excel 文件以及文件存放的格式如下, 必须这样放 Excel 表格， 如 excel_test_file 文件夹所示.

        <Excel的总文件夹>
                |
                |_ <EXCEL 分文件夹1> 
                |        |
                |        |_ excel表格1
                |        |
                |        |_ excel表格2
                |        |
                |        |_ <excel其他表格> 
                |        |
                |        |_ ...
                |
                |
                |_ <其他EXCEL份文件>
                         |
                         |_ <excel表格1>
                         |
                         |_ ...
            
