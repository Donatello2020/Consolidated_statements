import os
import sys
import xlwings as xw

# import configparser
curpath = os.path.dirname(os.path.realpath(__file__))
xlspath = os.path.join(curpath, 'xlsbook')
xls_name = []
for _xls_name in os.listdir(xlspath):
    if _xls_name.endswith('xlsx'):
        xls_name.append(_xls_name)
# print(xls_name)
if not xls_name:
    print('没有找到Excel表格,请检查当前目录')
    sys.exit(0)
for _xls_name in xls_name:
    wb = xw.Book(os.path.join(xlspath, _xls_name))
    wb.api.Application.ErrorCheckingOptions.BackgroundChecking = False  # 关闭Excel错误检查
    break
