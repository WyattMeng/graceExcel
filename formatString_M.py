# -*- coding: utf-8 -*-
"""
Created on Thu Aug 24 14:49:43 2017
去掉字符串左、右、中的空格
@author: Maddox.Meng
"""

def formatString(foo):
    return "".join(foo.split())    
    #return foo.strip().replace(' ','')

#print formatString(' a a ')