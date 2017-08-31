# -*- coding: utf-8 -*-
"""
Created on Thu Aug 31 11:37:19 2017

@author: Maddox.Meng
"""

import os
import pandas as pd
import numpy as np
from numpy import float64
#import os.path
s = os.sep #根据unix或win，s为\或/
path = r'C:\Users\maddox.meng\Desktop\哈尔滨GET表格及贷款、存款WP和贷款PBC\底稿'
root = 'c:'+s+'Users\maddox.meng\Desktop\哈尔滨GET表格及贷款、存款WP和贷款PBC\底稿'+s
path2= 'C:\\Users\\maddox.meng\\Desktop\\automation'


#for rt,dirs,files in os.walk(path2):
#    for f in files:
#        print f
 
awps = []
for i in os.listdir(path2):
    if os.path.isfile(os.path.join(path2,i)):
        awps.append(os.path.join(path2,i))
 
#for awp in awps:
#    df = pd.read_excel(awp)
        
df = pd.read_excel(awps[0],header=None)

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