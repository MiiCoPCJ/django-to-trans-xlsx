from openpyxl import *


def line_title(sheet,r):
    list_title = ['字段名','字段英文名','数据类型','默认','描述']

    c = 1
    for list in list_title:
        sheet.cell(row=r, column=c, value=list)
        c = c + 1

def foreign_title(sheet,r,length):
    list_title = ['外键','外键表','描述']

    c = 1
    for list in list_title:
        sheet.cell(row=r+length, column=c, value=list)
        c = c + 1