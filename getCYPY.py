# -*- coding: utf-8 -*-
"""
Created on Thu Aug 24 16:18:41 2017
get CY PY值 from GET's Instraction sheet
@author: Maddox.Meng
"""

import pandas as pd
#import numpy as np
#from formatString_M import formatString

def getCY(xlsxPath):

    #源、目标excel
    path_src = 'C:\Users\maddox.meng\Desktop\E10000.xlsx'
    path_dest= 'C:\Users\maddox.meng\Desktop\GET.xlsx'
    
    sheet = 'Instraction'
    df = pd.read_excel(xlsxPath, sheet, header=None)
    #print df.shape #out:(23.3), 23 rows, 3 columns. type:tuple
    #print type(df.shape)
    Count_Row=df.shape[0] #gives number of row count
    Count_Col=df.shape[1] #gives number of col count
    #print df.iloc[0]
    #print df.iloc[:,0]
    #print df.iloc[0:1]
    x = 0
    y = 0
    while x < Count_Row:
        while y < Count_Col:
            if df.iloc[x,y] == 'CY':
                #print x,y
                CY = df.iloc[x,y+1]         ####假设日期就在'CY'的右边一列 
    #        elif df.iloc[x,y] == 'PY': 
    #            PY = df.iloc[x,y+1]
            y+=1
        x+=1    
    
    #CY年份减1得到PY
    PY = unicode(int(CY[0:4])-1)+CY[4:]
    #print 'CY=',CY
    #print 'PY=',PY
    return [CY,PY]


######################
#print getCY(path_dest)[1]

'''
1.一个dataframe有多少row、column
2.iloc[x,y]就是定位到第x行、第y列的数据
3.iloc[0]打印出结果是：
    0             CY
    1    2017年12月31日
    2            NaN
    Name: 0, dtype: object
4.要打印出第1行，必须切片df[0:1]
5.判断4个连续的数字   
6.Instraction里CY PY各对应3中形式的日期格式
    2017年12月31日
    2017年度
    2017
  应当把他们存为dateList，以此去例如sheet 【发放贷款及垫款-4.3逾期贷款CY】匹配年份。
 
 
'''