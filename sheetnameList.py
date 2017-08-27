# -*- coding: utf-8 -*-
"""
Created on Thu Aug 24 10:37:00 2017
读取某个excel的所有sheetname
@author: Maddox.Meng
"""

import xlrd
import pandas as pd

xlsfile = r'C:\Users\maddox.meng\Desktop\GET.xlsx'
book = xlrd.open_workbook(xlsfile) 
#xlrd用于获取每个sheet的sheetname  
#count = len(book.sheets())
with pd.ExcelWriter('newxls.xls') as writer:
    for sheet in book.sheets():
        print sheet.name
        df = pd.read_excel(xlsfile,sheet.name,index_col = None,na_values= ['9999'])
        df.to_excel(writer,sheet_name = sheet.name)