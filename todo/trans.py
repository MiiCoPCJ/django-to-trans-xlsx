import os
import re
from openpyxl import *

# 执行文件的路径
path = os.path.split(os.path.realpath(__file__))[0]
# 读取model数据
f = open(path + '/../resource/users_models.py','r',encoding='utf-8')
str = f.read()
f.close()
# 正则 查找
# pattern = re.compile(r'class\s+\w+\(models.Model\):.*?verbose_name\s?=\s?verbose_name_plural\s?=\s?(\'|")\w+(\'|")', re.S)
#
# find = re.finditer(pattern,str)
# 历遍
# for m in find:
#     print(m.group())
#     print('?????')

# xlsx文件只存一个
# if os.path.isfile(path + "/../doc/model.xlsx"):
#     os.remove(path + '/../doc/model.xlsx')
#     wb = Workbook()
#     wb.save(path + '/../doc/model.xlsx')
# else:
#     wb = Workbook()
#     wb.save(path + '/../doc/model.xlsx')

# 查找resource目录所有model文件
dir = path + "/../resource/"
for parent, dirnames, filenames in os.walk(dir):
    for filename in filenames:
        file_path = os.path.join(parent, filename)
