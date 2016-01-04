# -*- coding: utf-8 -*-
"""
Created on Sun Aug 24 16:03:55 2014

@author: Ashwin
"""

import win32com.client


def Excel(pathname, filename, visible = True):
    '''
    wb, xl = Excel(pathname, filename, visible=True)
    '''
    try:
        xl = win32com.client.GetActiveObject("Excel.Application")
        xl.Visible = visible; bool_found = False
        for i in range(xl.Workbooks.Count):
            name = str(xl.Workbooks(i+1).Name)
            if name == filename:
                bool_found = True
                print 'File Found Open.  Acquiring.'            
                wb = xl.Workbooks(i+1)
                break
        if not bool_found:
            xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
            print 'File Not Found.  Opening file ...'
            wb = xl.Workbooks.Open(pathname+'/'+filename)
    except:
        xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
        print 'File Not Found.  Opening file ...'
        wb = xl.Workbooks.Open(pathname+'/'+filename)
    return wb, xl

def getActiveCell(xl):
    aw = xl.ActiveWindow
    cell = aw.ActiveCell
    return cell

def Word(pathname, filename, visible = True):
    '''
    doc, wrd = Word(pathname, filename, visible=True)
    '''
    try:
        wrd = win32com.client.GetActiveObject("Word.Application")
        wrd.Visible = visible; bool_found = False
        for i in range(wrd.Documents.Count):
            name = str(wrd.Documents(i+1).Name)
            if name == filename:
                bool_found = True
                print 'File Found Open.  Acquiring.'            
                doc = wrd.Documents(i+1)
                break
        if not bool_found:
            wrd = win32com.client.gencache.EnsureDispatch("Word.Application")
            print 'File Not Found.  Opening file ...'
            doc = wrd.Documents.Open(pathname+'/'+filename)
    except:
        wrd = win32com.client.gencache.EnsureDispatch("Word.Application")
        print 'File Not Found.  Opening file ...'
        doc = wrd.Documents.Open(pathname+'/'+filename)
    return doc, wrd

class Zakian:
    '''
    http://code.activestate.com/recipes/127469-numerical-inversion-of-laplace-transforms-through-/
    '''
    def __init__(self, F):
        self.F = F
    def f(self, t):
        a = 12.83767675+1.666063445j, 12.22613209+5.012718792j,\
        10.93430308+8.409673116j, 8.776434715+11.92185389j,\
        5.225453361+15.72952905j
        
        K = -36902.08210+196990.4257j, 61277.02524-95408.62551j,\
        -28916.56288+18169.18531j, +4655.361138-1.901528642j,\
        -118.7414011-141.3036911j
        
        summ = 0.0
        
        if t == 0:
            print "\n"
            print "ERROR:   Inverse transform can not be calculated for t=0"
            print "WARNING: Routine zakian() exiting. \n"
            return ("Error")
            
        for j in range(0,5):
            summ += (K[j]*self.F(a[j]/t)).real
        return 2.0*summ/t