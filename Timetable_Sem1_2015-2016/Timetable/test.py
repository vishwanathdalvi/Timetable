# -*- coding: utf-8 -*-
"""
Created on Sat May 30 17:06:23 2015

@author: Ashwin
"""

import os
import ExcelUtilities
thisdir = os.getcwd()
filename = 'Test.xlsx'
wb, xl = ExcelUtilities.Excel(thisdir, filename)
