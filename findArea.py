# -*- coding: utf-8 -*-
"""
Created on Sat Aug 26 17:20:44 2017

@author: Maddox.Meng
"""

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

df = pd.read_excel(path_src, 'E10000',header=None)


key = u'3.应收利息'

x_min=0
for item in df.iloc[:,0]:
    if item is not np.nan and formatString(item) == key:
        #print df.iloc[x,0] #x=8
        break  #必须要break，否则x会到最后一行108
    x_min+=1
    
#找 '3.应收利息' 往下的第一个带数字开头的item，以确定area的height
x_diff = 0
for item in df.iloc[x_min+1:,0]:
    if item is not np.nan and item[0].isdigit() is True:
        print item
        print x_diff
        break
    x_diff+=1

x_max = x_min + x_diff

#E10000的'3.应收利息'区域，1st row = x_min, last row = x_max

df_area = df[x_min:x_max]
#print df_area
for item in df_area[:,0]:
    if item is not np.nan:
        print item    
    