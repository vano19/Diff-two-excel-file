import Brain
import xlwt
from xlrd import open_workbook
from tkinter import *


class Main_Menu:
    def __init__(self,M,M2,MB):
        self.M = Brain.diff_list(MB,M)
        self.M2 = Brain.diff_list(MB,M2)
        self.MB = MB
        self.F1 = Frame(height = 400, width = 700)
        self.F1.pack()
        self.F2 = Frame(bg="#E6E6FA")
        self.F2.pack()
        self.F3 = Frame(height = 100, width = 700)
        self.F3.pack()
        
## Первый уровень--------------------       
        self.L_res = Label(self.F1, text = "Результат сравнения",font = "Arial 16",bg="#E6E6FA")
        self.L_res.pack(side = TOP, fill = X)
        
## Здесь будет таблица---------------------------------------------
        self.f1_1 = Frame(self.F1,height = 400, width = 600,bg="#FFDEAD")
        self.f1_1.pack(side = TOP)
        self.L_table = Label(self.f1_1,height = 25, width = 80,
                             text = "Здесь будет таблица",bg = 'grey')
        self.L_table.pack(fill = X)
        
##Второй уровень-------------------------------------------------------
        self.f2_1 = Frame(self.F2,height = 10, width = 700,bg="#E6E6FA")
        self.f2_1.pack(side = TOP)
        
        self.L_place = Label(self.F2,text = "Путь к файлу для сохранения", font = "Arial 12",bg="#E6E6FA")
        self.L_place.pack(side = TOP, fill = X)
        
        self.E_place = Entry(self.F2, width = 70, bd=3)
        self.E_place.pack()

##Третий уровень----------------------------------------------------
        def Merge(event):
            book = xlwt.Workbook(encoding="utf-8")
            sheet1 = book.add_sheet("merge")

            path = self.E_place.get()
            if len(path) == 0:
                print("Введите путь к файлу!")
                return 0
            
            #Заполнения окончательного файла   
            for i in range(len(self.M)):
                for j in range(len(self.M[i])):
                    #Проверка на конфликт
                    if (self.M[i][j] != " " and self.M2[i][j] != " "): 
                        print("Обнаружен конфликт! " + "%s -> %s <- %s" %(self.M[i][j],self.MB[i][j],self.M2[i][j]))
                        while(True):
                            s = input("С какого файла взять значение (1 или 2): ")
                            
                            if s == "1":
                                sheet1.write(i,j,M[i][j])
                                del s
                                break
                            elif s == "2":
                                sheet1.write(i,j,M2[i][j])
                                del s
                                break
                            
                    elif (self.M[i][j] != " "):
                        sheet1.write(i,j,M[i][j])

                    elif (self.M2[i][j] != " "):
                        sheet1.write(i,j,M2[i][j])

                    else:
                        sheet1.write(i,j,MB[i][j])

            book.save(path)
                            
     
        self.f3_1 = Frame(self.F3,height = 10, width = 700,bg="#E6E6FA")
        self.f3_1.pack(side = TOP)

        self.B_merge = Button(self.f3_1,text = "Merge", height = 2, width = 15,bg="#ADD8E6")
        self.B_merge.bind('<Button-1>',Merge)
        self.B_merge.pack(side = RIGHT)



