
import Brain
import NewMenu
import xlwt
from xlrd import open_workbook
import sys
from tkinter import *
import tkinter.filedialog as tf


excel_names=[]
for param in sys.argv:
    excel_names.append(param)
    
wb = open_workbook(excel_names[1])
wb2 = open_workbook(excel_names[2])
wb3 = open_workbook(excel_names[3])

s = wb.sheet_by_index(0)
s2 = wb2.sheet_by_index(0)
sB = wb3.sheet_by_index(0)

M = [[" "]*s.ncols]*s.nrows
M2 = [[" "]*s2.ncols]*s2.nrows
MB = [[" "]*sB.ncols]*sB.nrows

for row in range(s.nrows):
    M[row] = s.row_values(row)

for row in range(s2.nrows):
    M2[row] = s2.row_values(row)

for row in range(sB.nrows):
    MB[row] = sB.row_values(row)

Brain.equalizer(M,M2)
Brain.equalizer(MB,M2)

print(M)
print(M2)
print(MB)


root = Tk()
root.config(bg="#E6E6FA")
o = NewMenu.Main_Menu(M,M2,MB)

root.mainloop()
