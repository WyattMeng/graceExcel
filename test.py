# -*- coding: utf-8 -*-
"""
Created on Fri Aug 25 09:48:14 2017

@author: Maddox.Meng
"""

import pandas as pd
import numpy as np
#import xlrd
#import xlsxwriter
#from xlutils.copy import copy
#import os
#from formatString_M import formatString
from getCYPY import getCY

import datetime

#源、目标excel
path_src = 'C:\Users\maddox.meng\Desktop\E10000.xlsx'
path_dest= 'C:\Users\maddox.meng\Desktop\GET.xlsx'

#得到GET表的CY PY
getCY(path_dest)

def datetime_toString(dt):
    return dt.strftime("%Y年%m月%d日") 

df = pd.read_excel(path_src, 'risktable',header=None)


'''如果是datetime.datetime类型，xxxx，所以if下面三个and也是有顺序的
isinstance先把datetime去掉，再用item[0]判断。否则会提示datetime没有getitem方法'''
srcSheetList=[]
for item in df.iloc[:,0]:    #[row, column], 所有row第0列
    if item is not np.nan and isinstance(item, datetime.datetime) is False and item[0].isdigit() is True:
        print item
        #print datetime_toString(item)
        

        
    


'''
2015-12-31 00:00:00 === <type 'datetime.datetime'>
发放贷款及垫款 === <type 'unicode'>

python, 判断一个东西的type是否为某个type

未import datetime————datetime is not defined
from datetime import datetime————AttributeError: type object 'datetime.datetime' has no attribute 'datetime'
'datetime.datetime'————把第一个datetime当成datetime module了。那么第一个dt本来应该是什么？？
'''   