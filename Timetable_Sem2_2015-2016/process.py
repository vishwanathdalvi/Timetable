# -*- coding: utf-8 -*-
"""
Created on Thu May 21 12:58:02 2015

@author: Ashwin
"""

import ExcelUtilities
import os
import pandas
import scipy
import matplotlib.pyplot as plt

filename = 'TimeTable_Database.xlsm'
wb, xl = ExcelUtilities.Excel(os.getcwd(), filename, visible = False)
sheetdb = wb.Sheets("Database")
sheetfaculty = wb.Sheets("Faculty")
sheetclass = wb.Sheets("Class")
sheetroom = wb.Sheets("Room")
sheetdept = wb.Sheets("Departments")
worksheet = wb.Sheets("Worksheet")

xl_file = pandas.ExcelFile(filename)
df = {}


listslots = [1,2,3,4,5,6,7,8,9]
listslotnames = ['0830-0925','0930-1025','1040-1135','1140-1235',
                 '1330-1425','1430-1525','1540-1635','1640-1735',
                 '1740-1835']

listsymbols = ['', '*', '!','#','$','^','&','?','+','~','>']
                 
listdays = ['Mon','Tue','Wed','Thu','Fri','Sat']

dict_pos_slot = {}
dict_slot_pos = {}
for iday in xrange(len(listdays)):
    j = iday+1
    dict_pos_slot[j] = {}
    for i in listslots:    
        slot = listdays[iday]+str(i)
        dict_pos_slot[j][i] = slot
        dict_slot_pos[slot] = [j,i]
        

def reparse(df):
    xl_file = pandas.ExcelFile(filename)
    df['Database'] = xl_file.parse("Database")
    df['Faculty'] = xl_file.parse("Faculty")
    df['Class'] = xl_file.parse("Class")
    df['Room'] = xl_file.parse("Room")
    df['Departments'] = xl_file.parse("Departments")
    
    df['Database'].Room.fillna('')
    df['Database'].Slot.fillna('')
    df['Database'].Index = df['Database'].Code + '_' + df['Database'].Class + '_' + df['Database'].Department
    df['Database']['Index2'] = df['Database'].Faculty+'_'+df['Database'].Room.fillna('')

def makematrix(sheet, text, irow0, icol0, bool_legend = False, bool_read = False):
    matrix = scipy.zeros((len(listdays)+1,len(listslots)+1),dtype = 'S100')
    matrix[0] = [text]+listslotnames
    matrix[:,0] = [text]+listdays
    
    db = df['Database']
    listcol = list(db.columns)
    indroom = listcol.index('Room')
    
    if bool_read:
        dslt = db.Slot.fillna('').to_dict()
        slot_array = [dslt[key] for key in dslt.keys()]
        
        drm = db.Room.fillna('').to_dict()
        room_array = [drm[key] for key in drm.keys()]
        
    legend = scipy.zeros((100, len(db.columns)), dtype = 'S500')
    legend[0] = list(db.columns)
    
    dontshow = ['Batch','Elective', 'Index2']
    
    for column in dontshow:
        ncol = list(db.columns).index(column)
        legend[0][ncol] = ''
    
    #Getting Truncated Database
    if text in list(df['Faculty'].Index):
        trdb = db[db.Faculty == text]
        ddd = df['Faculty'][df['Faculty'].Index == text].iloc[0].fillna('')  #Note: .iloc[[0]] will return a dataframe
        dept1 = df['Departments'][df['Departments'].Index == ddd.Department.strip()].iloc[0].Name
        dept2 = ''        
        if ddd.Department2 in list(df['Departments'].Index):
            dept2 =', '+df['Departments'][df['Departments'].Index == ddd.Department2.strip()].iloc[0].Name 
        desc = ddd.Index+': '+ddd.Name+' ('+dept1+dept2+')'
    elif text in list(df['Room'].Index):
        trdb = db[db.Room == text]
        ddd = df['Room'][df['Room'].Index == text].iloc[0]
        desc = ddd.Index+': '+ddd.Description
    elif text in list(df['Class'].Index):
        trdb = db[db.Class == text]
        ddd = df['Class'][df['Class'].Index == text].iloc[0]
        desc = ddd.Index+': '+ddd.Description
    elif '-' in text:
        splttext = text.split('-')
        clss = splttext[0]
        dept = splttext[1]
        if clss in list(df['Class'].Index) and dept in list(df['Departments'].Index):
            if dept in ['PO','CO']:
                trdb = db[(db.Class == clss) & ((db.Department == 'TECH') | (db.Department == dept) | (db.Department == 'POCO'))]
            else:
                trdb = db[(db.Class == clss) & ((db.Department == 'TECH') | (db.Department == dept))]
            ddd = df['Class'][(df['Class'].Index == clss) & (df['Class'].Department == dept)].iloc[0]
            desc = ddd.Index + '-' + ddd.Department + ': '+ddd.Description
        else:
            sheet.Range("A1").Value = 'No Faculty, Room or Class by name %s'%text
            return
    else:
        sheet.Range("A1").Value = 'No Faculty, Room or Class by name %s'%text
        return 
        
    #Getting positions of matrix
    topleft = (irow0, icol0)
    topright = (irow0, icol0+len(listslots))
    botleft = (irow0+len(listdays), icol0)
    botright = (irow0+len(listdays), icol0+len(listslots))    
    
    cellTL = sheet.Cells(topleft[0], topleft[1])
    cellBR = sheet.Cells(botright[0], botright[1])
    cellBL = sheet.Cells(botleft[0], botleft[1])
    cellTR = sheet.Cells(topright[0], topright[1])
    
    #Reading from worksheet
    cellLT = sheet.Cells(topleft[0]+1, topleft[1]+1)
    rng = sheet.Range(cellLT.Address+":"+cellBR.Address)
    if bool_read:
        matrix[1:,1:] = rng.Value
        matrix[matrix == 'None'] = ''    

    #Grouping
    g = trdb.groupby('Index')
    ngrp = 0
    dict_groups = {}
    dict_index = {}
    dict_legend = {}
    nlegend = 0 #Row of legend matrix
    for group in g.groups:
        ngrp += 1
        dict_groups[ngrp] = {}
        nsubgrp = -1
        gg = g.get_group(group).groupby('Index2')
        for subgroup in gg.groups:
            nlegend += 1
            nsubgrp += 1
            dict_groups[ngrp][nsubgrp] = gg.get_group(subgroup).fillna('')
            data = dict_groups[ngrp][nsubgrp].iloc[0].to_dict()
            index = str(ngrp)+listsymbols[nsubgrp]
            data['Slot'] = len(gg.get_group(subgroup))
            if index in dict_index:
                dict_index[index] += [[ngrp, nsubgrp]]
            else:
                dict_index[index] = [[ngrp, nsubgrp]]
            legend[nlegend][0] = index
            dict_legend[index] = nlegend
            donshw = dontshow + ['Index']
            if bool_read:
                #Getting room from Legend
                indrm = listcol.index('Room')
                cellroom = sheet.Cells(irow0+len(listdays)+2+nlegend, 
                             icol0+indrm)
                donshw = dontshow + ['Index'] + ['Room']
                legend[nlegend][indrm] = cellroom.Value
            for column in listcol:
                if column not in donshw:
                    ind = listcol.index(column)
                    legend[nlegend][ind] = data[column]
            
            
            
            if not bool_read:
                #Putting Slots in Matrix
                ddict = dict_groups[ngrp][nsubgrp].to_dict()
                dictslot = ddict['Slot']
                dictbatch = ddict['Batch']
                for ind in dictslot:
                    slot = dictslot[ind]
                    btch = dictbatch[ind].strip()
                    if btch == 'All':
                        btch = ''
                    if slot in dict_slot_pos:
                        [i, j] = dict_slot_pos[slot]
                        slt = matrix[i][j].strip()
                        if slt == '':                    
                            matrix[i][j] = index+btch
                        else:
                            matrix[i][j] = slt+','+index+btch
    if bool_read:
        #Getting slots from Matrix
        dict_vals = {}
        for i in xrange(len(matrix)):
            for j in xrange(len(matrix[i])):
                if i > 0 and j > 0:
                    vl = matrix[i][j].strip()
                    if len(vl) > 0:
                        listval = vl.split(',')
                        for val in listval:
                            val = val.strip()
                            if val[-1] in ['A', 'B', 'C', 'D']:
                                btch = val[-1]
                                index = val[:-1]
                            else:
                                btch = 'All'
                                index = val
                            index = index.split('.')[0]
                            if val not in dict_vals:
                                dict_vals[val] = []
                            [[ngrp, nsubgrp]] = dict_index[index]
                            grps = dict_groups[ngrp][nsubgrp]
                            listkeys = grps[grps['Batch'] == btch].to_dict()['Batch'].keys()
                            ilegend = dict_legend[index]
                            room = legend[ilegend][indroom]
                            slt = dict_pos_slot[i][j]                                    
                            for key in listkeys:
                                if key not in dict_vals[val]:
                                    dict_vals[val] += [key]
                                    slot_array[key] = slt
                                    room_array[key] = room
                                    break
                                    
        slot_array = scipy.array(slot_array, dtype = 'S10').reshape((len(slot_array),1))
        room_array = scipy.array(room_array, dtype = 'S10').reshape((len(room_array),1))
        rng_slot = sheetdb.Range("I2:I%d"%(2+len(slot_array)-1))
        rng_slot.Value = slot_array
        rng_room = sheetdb.Range("H2:H%d"%(2+len(room_array)-1))
        rng_room.Value = room_array

    #Writing to worksheet
        
    rng = sheet.Range(cellTL.Address+":"+cellBR.Address)
    rng.Value = None
    rng.WrapText = True
    rng.Font.Bold = False
    rng.Font.Italic = False
    rng.Font.Size = 8
    rng.VerticalAlignment = 2
    rng.HorizontalAlignment = 3
    
    rng = sheet.Range(cellTL.Address+":"+cellTR.Address)
    rng.Font.Italic = True
    rng.Font.Size = 9
    rng = sheet.Range(cellTL.Address+":"+cellBL.Address)
    rng.Font.Italic = True
    rng.Font.Size = 9
    
    rng = sheet.Range(cellTL.Address+":"+cellBR.Address)
    rng.Value = matrix
    
    rng = sheet.Range(cellTL.Address)
    rng.Font.Bold = True
    rng.Font.Italic = False
    rng.Font.Size = 12
    
    rng = sheet.Cells(irow0-1, icol0)
    rng.Value = desc
    rng.Font.Bold = False
    rng.Font.Italic = True
    rng.Font.Size = 9
    rng.WrapText = False
    
    if bool_legend:        
        celltl = sheet.Cells(irow0+len(listdays)+2, 
                             icol0)
        cellbr = sheet.Cells(irow0+len(listdays)+2+100, 
                             icol0+len(df['Database'].columns)-1)
        celltr = sheet.Cells(irow0+len(listdays)+2,
                             icol0+len(df['Database'].columns)-1)
        cellbl = sheet.Cells(irow0+len(listdays)+2,
                             icol0)
        rng = sheet.Range(celltl.Address+":"+cellbr.Address)
        rng.WrapText = False
        rng.Value = None
        rng.Font.Bold = False
        rng.Font.Italic = False
        rng.Font.Size = 9
        rng.HorizontalAlignment = 2
        rng.VerticalAlignment = 2
        rng.Value = legend
    
        rng = sheet.Range(celltl.Address+":"+celltr.Address)
        rng.Font.Italic = True
    
    return nlegend

def gettext():
    sheet = worksheet
    rng = sheet.Range("C2")
    text = rng.Value
    return text

def getlist():
    sheet = worksheet
    rng = sheet.Range("O2")
    try:
        number = int(rng.Value)
        rng = sheet.Range("O3:O%d"%(2+number))
        listtext = scipy.array(rng.Value).reshape((1,number))[0]
        return list(listtext)
    except:
        return []
    
def getmatrix(sheet = worksheet, irow0 = 2, icol0 = 3, text = False):
    reparse(df)
    bool_ = False
    if text:
        bool_ = True
    if not bool_:
        text = gettext()
    nlegend = makematrix(sheet, text, irow0, icol0, bool_legend = True, bool_read = False)
    listtext = getlist()
    icol = 20
    irow = 2
    if not bool_:
        for text in listtext:
            nlegend = makematrix(worksheet, text, irow, icol, bool_legend = False, bool_read = False)
            irow += len(listdays) + 4
    
def update():
    reparse(df)
    text = gettext()
    nlegend = makematrix(worksheet, text, 2, 3, bool_legend = True, bool_read = True)
    wb.Save()
    getmatrix()
    

reparse(df)  
if __name__ == '__main__':
    xl.Visible = True
    #Infinite loop 
    import pywintypes
    boolgo = True
    while boolgo:
        plt.pause(0.1)
        try:
            cell = worksheet.Range("A1")
            text = cell.Value
            if text == 'Pull':
                cell.Value = ''
                worksheet.Range("A2").Value = '' 
                getmatrix()
            elif text == 'Push':
                cell.Value = ''
                concell = worksheet.Range("A2")
                confirm = concell.Value
                if confirm == 'Update':
                    concell.Value = ''
                    update()
                else:
                    concell.Value = 'Not Updated!'
        except pywintypes.com_error:
            continue
                
    