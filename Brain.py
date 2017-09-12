import xlwt
from xlrd import open_workbook
from tkinter import *
import sys
from xlutils.filter import process,XLRDReader,XLWTWriter

#Create table

class Table:
    def __init__(self,parent,M,color):
        F = Frame(parent)
        
##        rows = 100
##        if len(M) > 100:
##            rows = len(M)
##        colns = 100
##        if len(M[0]) > 100:
##            rows = len(M[0])
            
        for i in range(len(M)):
            for j in range(len(M[i])):
                try:
                    if M[i][j] != "-":
                        l = Label(F,text='%s' % (M[i][j]), relief=RIDGE,bg = 'white')
                        l.grid(row=i, column=j, sticky=NSEW)
                    else:
                        l = Label(F,text='%s' % (M[i][j]), relief=RIDGE,bg = color)
                        l.grid(row=i, column=j, sticky=NSEW)
                except IndexError:
                    l = Label(F,text=" " , relief=RIDGE,bg = 'white')
                    l.grid(row=i, column=j, sticky=NSEW)
                
        

        F.pack(side = TOP)


#Delete str 
style_red = xlwt.XFStyle()
pattern_red = xlwt.Pattern()
pattern_red.pattern = xlwt.Pattern.SOLID_PATTERN
pattern_red.pattern_fore_colour = xlwt.Style.colour_map['red']
style_red.pattern = pattern_red

#added str
style_green = xlwt.XFStyle()
pattern_green = xlwt.Pattern()
pattern_green.pattern = xlwt.Pattern.SOLID_PATTERN
pattern_green.pattern_fore_colour = xlwt.Style.colour_map['green']
style_green.pattern = pattern_green

#changed str
style_delete = xlwt.XFStyle()
pattern_delete = xlwt.Pattern()
pattern_delete.pattern = xlwt.Pattern.SOLID_PATTERN
pattern_delete.pattern_fore_colour = xlwt.Style.colour_map['blue']
style_delete.pattern = pattern_delete

def del_add(i,k,sB,sN):
    b = 0
##    print(sN.ncols)     
    for t in range(i,sN.nrows):
##        print (sB.cell(i,0).value, sN.cell(t,0).value)
        if sB.cell(i-k,0).value == sN.cell(t,0).value:
            b = 1
            return b
    return b

def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb,'unknown.xls'),
        w
        )
    return w.output[0][1], w.style_list

#main diff
def Diff(sB,sN,s3):
    k = 0
    for i in range(min(sB.nrows, sN.nrows)):
        for j in range(min(sB.ncols, sN.ncols)):
            try:
                if sB.cell(i-k,j).value != sN.cell(i,j).value:
##                    print(k)
##                    print("Отличие")
##                    print(i,j)
                    if j == 0:
                        c = del_add(i,k,sB,sN)
##                        print("c=",c)
                        if c == 0:
##                            print("dell")
                            k -= 1
                            for t in range(sN.nrows):
                                s3.write(i,t,sN.cell(i,t).value, style_delete)
                            break
                            
                            
                        else:
##                            print("add")
                            k+=1
                            for t in range(sN.nrows):
                                s3.write(i,t,sN.cell(i,t).value, style_green)
                            break

                            
                    else:
##                        print("chenge")
                        s3.write(i,j,sN.cell(i,j).value, style_red)
            except IndexError:
                print (i,j,k)
                print("error")
                break
    return s3


wb = open_workbook("base_.xls",formatting_info=True)
wb2 = open_workbook("base_1.xls",formatting_info=True)
wb3,style_list = copy2(wb2)

sB = wb.sheet_by_index(0)
sN = wb2.sheet_by_index(0)
s3 = wb3.get_sheet(0)


s3 = Diff(sB,sN,s3)

wb3.save("Diff.xls")
                    
                   

#выравнивает столбцы и строки
def equalizer(A,B):

    if len(A[0]) > len(B[0]):
        A,B = B,A

    

## добавляет в средину столбец
    for i in range(len(B[0])):
        if A[0][i] != B[0][i]:
            A[0].insert(i,"-")
            for n in range(len(A)):
                if n != 0:
                    A[n].insert(i,"-")
            if (len(A[0]) == len(B[0])):
                break
        
##  добавляет в конец столбец   
    if A[0][-1] != B[0][-1]:
        A[0].append(B[0][-1])
        for n in range(len(A)):
            if n != 0:
                A[n].append("-") 

    if len(A) > len(B):
        A,B = B,A
## добавляем в середину строку
    for i in range(len(B)):
        try:
            if A[i][0] != B[i][0]:
                A.insert(i,["-"]*len(B[i]))
        except IndexError:
             A.append(["-"]*len(B[-1]))
        if (len(A) == len(B)):
            break

       
        




def diff_list(M,M2):

    R = []
    for i in range(max(len(M),len(M2))):
        R.append([])
        for g in range(max(len(M[i]),len(M2[i]))):
            R[i].append(" ")
            
    for i in range(max(len(M),len(M2))):
        for j in range(max(len(M[i]),len(M2[i]))):
            if M[i][j] == M2[i][j]:
                pass
            else:
                R[i][j] = M2[i][j]       
    return R

def full_list(M,MB):

    R = []
    for i in range(max(len(M),len(MB))):
        R.append([])
        for g in range(max(len(M[i]),len(MB[i]))):
            R[i].append(" ")
            
    for i in range(max(len(M),len(MB))):
        for j in range(max(len(M[i]),len(MB[i]))):
            if M[i][j] != " ":
                R[i][j] = M[i][j]
            else:
                R[i][j] = MB[i][j]       
    return R
    

    

    





