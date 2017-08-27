# -*- coding: utf-8 -*-
"""
Created on Tue Aug 22 11:28:32 2017

填数到Grace-Excel Template从底稿

@author: Maddox.Meng
"""

#import sys
#reload(sys)  
#sys.setdefaultencoding('utf8')
import pandas as pd
import numpy as np
import xlrd
import xlsxwriter
from xlutils.copy import copy
import os
from formatString_M import formatString
from getCYPY import getCY
import datetime

#源、目标excel
path_src = 'C:\Users\maddox.meng\Desktop\E10000.xlsx'
path_dest= 'C:\Users\maddox.meng\Desktop\GET.xlsx'

#得到GET表的CY PY
CY = getCY(path_dest)[0]
PY = getCY(path_dest)[1]

#目标excel,Grace ET所有sheet列表
wb = xlrd.open_workbook(path_dest)
destSheetList=[]
for sheet in wb.sheets():
    destSheetList.append(sheet.name)


#源excel E10000的 第1列 的以数字开头的cell值————对应目标excel的sheetList
#（此处先假设已经确定E10000第1列为‘科目名称’）
df = pd.read_excel(path_src, 'E10000',header=None)
#df = pd.read_excel(path_src, 'risktable',header=None)
#for item in df.icol(0): #icol(i) is deprecated. Please use .iloc[:,i]
srcSheetList=[]
for item in df.iloc[:,0]:    #[row, column], 所有row第0列
    #if item is not np.nan and item[0].isdigit() is True and item != CY is True and item != PY is True:
    if item is not np.nan and isinstance(item, datetime.datetime) is False and item[0].isdigit() is True:    
        #print item
        #print "".join(item.split()) #为什么replace(' ','')只去掉字符中3连续空格的一个空格？？        
        srcSheetList.append("".join(item.split()))


#匹配 srcSheetList & destSheetList 元素
list = []
for item_s in srcSheetList:
    for item_d in destSheetList:
        #if item_d.find(item_s.split('.')[-1]) != -1: #is True
        if item_d == item_s.split('.')[-1]:
            #print item_s,'--->>>',item_d
            list.append(item_s+'--->>>'+item_d)
        if item_s.find('.') != -1 and item_d.find(item_s[0:3]) != -1:
            #print item_s,'--->>>',item_d
            list.append(item_s+'--->>>'+item_d)
print repr(list).decode("unicode-escape")  
print list[0].split('--->>>')[0]          
       
        
        
'''
1.如何模糊匹配srcSheetList和destSheetList

2.如何匹配
  '4.3逾期贷款'
  '发放贷款及垫款-4.3逾期贷款CY'、
  '发放贷款及垫款-4.3逾期贷款PY'
  
3.如何向现有xlsx写入数据

4.怎么知道Grace ET里哪些数据需要被填入 

5.有没有可能把E10000数据做成dict

6.读risktable的列1，excel显示2015年12月31日, 实际python输出为2015-12-31 00:00:007.

7. if xxxx ==1 / !=-1 / is True

8.打印中文dict, print repr(dict).decode("unicode-escape")
'''