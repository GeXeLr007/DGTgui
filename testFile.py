# import os
#
# path = os.getcwd() + '\\data'  # 文件夹目录
# print(path)
#
# files = os.listdir(path)  # 得到文件夹下的所有文件名称
# s = []
# for file in files:  # 遍历文件夹
#     if not os.path.isdir(file) and os.path.splitext(file)[-1][1:] == "xls":  # 判断是否是文件夹，不是文件夹且是xls文件才打开
#         print(file)

import matplotlib.pyplot as plt
import pandas as pd
import os
import xlsxwriter
from xlutils.copy import copy
import xlrd
from PIL import Image
import sys

fig = plt.figure(1)
figName = 'foo.jpg'
fig.savefig(figName)

img = Image.open(figName)
img.save('foo.bmp')

# writer = pd.ExcelWriter('savepicture.xlsx', engine='xlsxwriter')
workbook = xlrd.open_workbook('savepicture.xls')
newWb = copy(workbook)
sheet = newWb.add_sheet('test2')
sheet.insert_bitmap('foo.bmp', 0, 0, scale_x=0.5, scale_y=0.5)
newWb.save('savepicture.xls')
os.remove(figName)
os.remove('foo.bmp')
