# coding=utf-8
import xlwings as xw
from xlwings import constants

wb = xw.Book(r'D:\python\Consolidated statements\AJE.xlsx')
sht = wb.sheets('Sheet1')
rng = sht.range('e2:e15')


def set_validation(range_1):
    range_1.api.Validation.Delete()
    range_1.api.Validation.Add(3, 3, 1, 'AAA,2,3')
    return


set_validation(rng)
