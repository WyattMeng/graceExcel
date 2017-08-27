# -*- coding: utf-8 -*-
"""
Created on Fri Aug 25 13:52:47 2017

@author: Maddox.Meng
"""

#源、目标excel
path_src = 'C:\Users\maddox.meng\Desktop\E10000.xlsx'
path_dest= r'C:\Users\maddox.meng\Desktop\GET.xlsx'
path_tst= 'C:\Users\maddox.meng\Desktop\src_dest_sheetList.xlsx'
p = 'C:\Users\maddox.meng\Desktop\zzz.xlsx'

from openpyxl import load_workbook
import openpyxl; 
print openpyxl.__version__

#修改GET
wb = load_workbook(filename = path_dest, data_only = False, guess_types = False)
##print wb.get_active_sheet
ws = wb.get_sheet_by_name(u'应收利息')
ws.cell(row=4,column=4).value = 9999
wb.save(path_dest)


#修改path_tst
#wb = load_workbook(path_tst)
#ws = wb.get_sheet_by_name('Sheet1')
#ws.cell(row=1,column=3).value = 232
#wb.save(path_tst)

import xlrd
import xlwt
workbook = xlrd.open_workbook(path_dest)
sheet2 = workbook.sheet_by_name(u'应收利息')
print sheet2.name,sheet2.nrows,sheet2.ncols

'''
1.spyder内置openpyxl当前版本2.4.1，怎么更新到最新
CMD里指定版本号更新，自动卸载之前的2.4.1：#######pip install openpyxl==2.5.0a3
###########重启spyder，解决了问题啦！！！'NoneType' object has no attribute 'group' load_workbook
'''