# -*- coding: utf-8 -*-
"""
Created on Fri Jun 12 20:57:17 2015

@author: Ashwin
"""

import process
filename = 'TimeTable.xlsx'
wb, xl = process.ExcelUtilities.Excel(process.os.getcwd(), filename, visible = False)
dict_sheet = {'Class':wb.Sheets('Class'),
              'Room':wb.Sheets('Room'),
              'Faculty':wb.Sheets('Faculty')}
    

def maketable(res, listhead):
    '''
    res should be one of ['Class', 'Room', 'Faculty']
    '''
    sheet = wb.Sheets(res)
    irow0 = 3
    icol0 = 1 #NOT 0!
    
    dc = process.df[res][listhead].to_dict()
    for key in dc:
        text = dc[key]
        nlegend = process.makematrix(sheet, text, irow0, icol0, bool_legend = True, bool_read = False)
        irow0 += nlegend + 8 + len(process.listdays)
    
maketable('Class','Name')
maketable('Room','Index')
maketable('Faculty','Index')  
wb.Save()  
process.wb.Close()
wb.Close()