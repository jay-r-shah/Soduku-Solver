# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 10:57:43 2015

@author: Jay
"""

import os
import win32com.client as win
import numpy as np
import math
#from sys import exit
#import time

sud=np.zeros((9,9))
thisdir=os.getcwd()+'/'
filename='Sudoku.xlsx'
xl=win.gencache.EnsureDispatch('Excel.Application')
xl.Visible=True
wb=xl.Workbooks.Open(thisdir+filename)
sheet=wb.Sheets('Solution')

#start=time.time()
for row in range(9):
    for col in range(9):
        if math.isnan(np.float64(sheet.Cells(row+1,col+1).Value)):
            sud[row,col]=0
        else:
            sud[row,col]=sheet.Cells(row+1,col+1).Value
            sheet.Cells(row+12,col+1).Value=sheet.Cells(row+1,col+1).Value
            sheet.Cells(row+12,col+1).Interior.ColorIndex=4

def grid2str(grid):
    p=''
    for row in range(9):
        for col in range(9):
            p=p+str(int(grid[row,col]))
    return p

def check_row(i,j):
    if i//9==j//9:
        return True #True if cell with zero is in the same row as cell j

def check_col(i,j):
    if i%9==j%9:
        return True #True if cell with zero is in the same col as cell j

def check_block(i,j):
    if (i//27 == j//27 and i%9//3 == j%9//3):
        return True #You get the drift

def sol(sudo):
    i=sudo.find('0')
    if i==-1:
        for row in range(9):
            for col in range(9):
                sheet.Cells(row+12,col+1).Value=sudo[row*9+col]
    else:
        excluded_nums=set()
        for j in range(81):
            if check_row(i,j) or check_col(i,j) or check_block(i,j):
                excluded_nums.add(sudo[j]) #adds to a set of impossible nums
        for m in '123456789':
            if m not in excluded_nums:
                sol(sudo[:i]+m+sudo[i+1:])

print(sol(grid2str(sud)))