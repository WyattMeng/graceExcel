# -*- coding: utf-8 -*-
"""
Created on Sun Aug 27 16:04:03 2017
处理带感叹号的cell value，即到另一个sheet中寻找值
@author: Maddox.Meng
"""

def jumpToSheet(cellValue,wb):
    if cellValue.find('!') == -1:
        return cellValue
    else:
        #print cv.split('!')[0].lstrip('=')
        sheetname = cellValue.split('!')[0].lstrip('=')
        #sheetname = "'"+sheetname+"'"
        #print sheetname
        #print '%s'%sheetname
        ws_temp = wb.get_sheet_by_name(sheetname) 
        #print cv.split('!')[-1] 
        return ws_temp[cellValue.split('!')[-1]].value 