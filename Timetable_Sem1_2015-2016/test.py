# -*- coding: utf-8 -*-
"""
Created on Sun Jul 19 19:12:41 2015

@author: Ashwin
"""

import win32com.client
import os
import scipy

thisdir = os.getcwd()
filename = 'Timetable.xlsx'
xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
xl.Visible = False
wb = xl.Workbooks.Open(thisdir+'/'+filename)

sh = wb.Sheets('Class')
rng = sh.Range("L4:L8")

#To give this range a name 'TTU'
rng.__setattr__('Name','TTU')
#To access the name of the range
print rng.Name.Name

ie = win32com.client.gencache.EnsureDispatch("Chrome.Application")
ie.Visible = True