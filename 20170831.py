# -*- coding: utf-8 -*-
"""
Created on Thu Aug 31 11:37:19 2017

@author: Maddox.Meng
"""

import os
import pandas as pd
import numpy as np
from xlrd import open_workbook
from numpy import float64
#import os.path
s = os.sep #根据unix或win，s为\或/
path = r'C:\Users\maddox.meng\Desktop\哈尔滨GET表格及贷款、存款WP和贷款PBC\底稿'
root = 'c:'+s+'Users\maddox.meng\Desktop\哈尔滨GET表格及贷款、存款WP和贷款PBC\底稿'+s
path2= 'C:\\Users\\maddox.meng\\Desktop\\automation'


#for rt,dirs,files in os.walk(path2):
#    for f in files:
#        print f


'''1'''
#list文件夹里所有底稿文件 
awps = []
for i in os.listdir(path2):
    if os.path.isfile(os.path.join(path2,i)):
        awps.append(os.path.join(path2,i))

'''2''' 
#遍历每个底稿    
#for awp in awps:
#    df = pd.read_excel(awp)

#'''3'''
##遍历一个底稿的每个sheet
#wb = open_workbook(awps[0])
#for sheet in wb.sheet_names():
#    df = pd.read_excel(awps[0], sheet, header=None)
        
df = pd.read_excel(awps[0],header=None)

'''4'''
#确定某excel的某sheet的第几列存放对应GET表的sheetname。逻辑是出现文字的第一列
#这个excel的第0列是sheetname
def findSheetnameCol(dataframe):
    for y in range(0,df.shape[-1]):
        x = 0
        for cell in df[y]:
            if cell is not np.nan:
                #print x,y,cell
                #break
                return y
            x+=1

'''5'''
#列出底稿里某sheet的某column里存放的sheetname和row number
def sheetNames(column):
    sheetnames=[]
    rownumbers=[]
    x = 0       
    for cell in df[findSheetnameCol(df)]:
        if cell is not np.nan:
            #print x,cell
            sheetnames.append(cell)
            rownumbers.append(x)
        x+=1    
    print sheetnames,rownumbers    

sheetnames = [u'BS_\u8d44\u4ea7', u'PL_\u51c0\u5229\u6da6', u'EQ_CY', u'EQ_PY']
rownumbers = [4, 41, 71, 87]
dict={}
'''6'''
#根据上面的row number，加上df.shape可以确定每个sheetname的area
#[4, 41, 71, 87], df.shape = (101, 21), [4:41,0:21],[41:71,0:21],[71:87,0:21],[87:101,0:21]
for i in range(0,len(sheetnames)):
    if i < len(sheetnames) - 1 :
#        x_bot = rownumbers[i+1]
#        area = '[%s:%s, 0:%s]' % (rownumbers[i],x_bot,df.shape[1])
#        dict[i] = ({'sheet':sheetnames[i],
#                   'area': area})
        dict[i] = {'sheet':sheetnames[i], 'x_min':rownumbers[i], 'x_max':rownumbers[i+1], 'y_min':0,'y_max':df.shape[1]}
    else:
#        x_bot = df.shape[0]
#        area = '[%s:%s, 0:%s]' % (rownumbers[i],x_bot,df.shape[1])
#        dict[i] = ({'sheet':sheetnames[i],
#                   'area': area})   
        dict[i] = {'sheet':sheetnames[i], 'x_min':rownumbers[i], 'x_max':df.shape[0], 'y_min':0,'y_max':df.shape[1]}    

print dict

#BS_资产的area是
print df.iloc[dict[0]['x_min']:dict[0]['x_max'], dict[0]['y_min']:dict[0]['y_max']]





#找底稿里关于年份的cell，关键字是‘201’
x=0    
for x in range(0,df.shape[0]):
    y=0
    for y in range(0,df.shape[1]):
        cell = df.iloc[x,y]
        #if cell is not np.nan and cell is not np.float64 and cell.find('年') != -1:
        if isinstance(cell,unicode) is True and cell.find('201') !=-1:
        #print type(cell)
            print cell,x,y
        y+=1        
    x+=1
    
  
    
    
'''
1.SyntaxError: (unicode error) 'unicodeescape' codec can't decode bytes in position 2-3: truncated \UXXXXXXXX escape
  r'C:/....'
'''    