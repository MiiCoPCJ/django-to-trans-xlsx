import os
import re
from openpyxl import *
from common.format import *

# 执行文件的路径
path = os.path.split(os.path.realpath(__file__))[0]
# 读取model数据 ,去除#注明部分
pre = 'users'
str = ''

# 开启
wb = Workbook()
sheet = wb.active
sheet.title = pre

# sheet的行r和列c
r = 1
c = 1

f = open(path + '/../resource/users_models.py','r',encoding='utf-8')
for line in f.readlines():
    line = line.strip()
    if not len(line) or line.startswith('#'):
        continue
    str += line + '\n'
f.close()

# 正则 查找
pattern = re.compile(r'class\s+\w+\(models.Model\):.*?verbose_name\s?=\s?verbose_name_plural\s?=\s?(\'|")\w+(\'|")', re.S)

find = re.finditer(pattern,str)
# 历遍
for m in find:
    model = m.group()

    # 表名 英文
    data_name = re.search('class\s+(\w+)\(models.Model\):', model).group(1).lower()
    # 表名 中文
    data_name_cn = re.search('verbose_name_plural\s?=\s?(\'|")(\w+)(\'|")', model).group(2)

    sheet.cell(row=r,column=c,value='表名称:')
    sheet.cell(row=r, column=c+1, value=data_name_cn+'('+data_name+')')
    r = r + 1
    line_title(sheet,r)

    # 获取字段部分
    field = re.search('class\s+\w+\(models.Model\):(.*)class\sMeta:', model, re.S).group(1)
    stp = field.strip().split('\n')
    length = len(stp)
    for sp in stp:
        # 外键处理
        Tip = True
        key = re.search('(\w+)\s?=\s?models\.ForeignKey\((.*)\)',sp)
        if key is not None and len(key.group()):
            key_name = key.group(1)
            foreign = key.group(2).split(',')[0].strip()
            # 查找外键所在表
            link_key = re.search('from\s(\w+).models\simport\s'+foreign, str)
            if link_key is not None and len(link_key.group()):
                dl = link_key.group(1) + foreign.lower()
            else:
                dl = pre + foreign.lower()

            if Tip:
                foreign_title(sheet,r,length)
                Tip = False
            sheet.cell(row=r + length + 1, column=1, value=key_name)
            sheet.cell(row=r + length + 1, column=2, value=dl)

            des = re.search('\s?'+foreign+'\s?,(.*)',key.group(2)).group(1).strip()
            sheet.cell(row=r + length + 1, column=3, value=des)



        d = re.search('(\w+)\s?=\s?models\.(\w+)Field\((.*)\)',sp)

        if d is not None and len(d.group()):
            field_name = d.group(1)
            field_type = d.group(2)
            field_name_cn = re.search('verbose_name\s?=\s?(\'|")(\w+)(\'|")', d.group(3)).group(2)

            sheet.cell(row=r, column=1, value=field_name_cn)
            sheet.cell(row=r, column=2, value=field_name)

            if field_type == 'Decimal':
                max_digits = re.search('max_digits\s?=\s?(\d+)\s?(,)?', d.group(3)).group(1)
                decimal_places = re.search('decimal_places\s?=\s?(\d+)\s?(,)?', d.group(3)).group(1)
                sheet.cell(row=r, column=3, value=field_type+'('+max_digits+','+decimal_places+')')
            if field_type == 'Char':
                max_length = re.search('max_length\s?=\s?(\d+)\s?(,)?', d.group(3)).group(1)
                sheet.cell(row=r, column=3, value=field_type + '(' + max_length + ')')

        r = r + 1

    wb.save(path + '/../doc/model.xlsx')





# xlsx文件只存一个
# if os.path.isfile(path + "/../doc/model.xlsx"):
#     os.remove(path + '/../doc/model.xlsx')
#     wb = Workbook()
#     wb.save(path + '/../doc/model.xlsx')
# else:
#     wb = Workbook()
#     wb.save(path + '/../doc/model.xlsx')

# 查找resource目录所有model文件
# dir = path + "/../resource/"
# for parent, dirnames, filenames in os.walk(dir):
#     for filename in filenames:
#         file_path = os.path.join(parent, filename)
