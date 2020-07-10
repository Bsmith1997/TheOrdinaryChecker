#!/usr/bin/env python
from xlrd import open_workbook
import sys
import re
from tkinter import messagebox
from tkinter import *


def main():
    book = open_workbook('TheOrdinaryConflicts.xlsx')
    sheet = book.sheet_by_index(0)
    product_col = 0
    conflict_col = 1 
    group_col = 2
    toCheck = txt.get()
    x = []
    y = []
    z = []
    for row in range(sheet.nrows):
        if sheet.cell(row,product_col).value == toCheck:
            group = sheet.cell(row,group_col).value
            lbl2 = Label(f2, font='Helvetica 14 bold',text = toCheck + " conficts with: ", background="coral1")
            lbl2.place(anchor="c",relx=.5, rely=.4)
            for rows in range(sheet.nrows):
                if group in sheet.row_values(rows)[1]:
                   if sheet.row_values(rows)[0] != toCheck and sheet.row_values(rows)[0] != "Product":
                        x.append(sheet.row_values(rows)[0])
                if toCheck in sheet.row_values(rows)[1]:
                    if sheet.row_values(rows)[0] != toCheck and sheet.row_values(rows)[0] != "Product":
                        y.append(sheet.row_values(rows)[0])
                if sheet.row_values(rows)[0] in sheet.cell(row,1).value:
                    if sheet.row_values(rows)[0] != toCheck and sheet.row_values(rows)[0] != "Product":
                        z.append(sheet.row_values(rows)[0])
                if sheet.row_values(rows)[2] in sheet.cell(row,1).value:
                    if sheet.row_values(rows)[0] != toCheck and sheet.row_values(rows)[0] != "Product":
                        z.append(sheet.row_values(rows)[0])

    x_set = set(x)
    y_set = set(y)
    z_set = set(z)
    print(x_set)
    print(y_set)
    print(z_set)

    total = x + y + z
    total_set = set(total)
   
    if not total_set:
        lbl4 = Label(f2, text="This product does not have any conficts. \n If this does not make sense then please check your spelling, \n this is case and spelling sensitive!", background="coral1", font='Helvetica 14 bold')
        lbl4.place(anchor="c",relx=.5, rely=.7)
    else: 
        
        final = "\n ".join(str(e) for e in total_set)
        print(final)
        lbl4 = Text(f2, background="coral1")
        lbl4.insert(INSERT, final)
        lbl4.config(state='disabled')
        lbl4.place(anchor="c",relx=.5, rely=.7)
        v = Scrollbar(window, orient= VERTICAL) 
        v.pack(side = RIGHT, fill = Y)
        lbl4.config(yscrollcommand = v.set)
        v.config(command)
        

window = Tk()
window.geometry("750x600")
f1 = Frame(width=800, height=600, background="white")
f2 = Frame(width=775, height=575, background="SkyBlue3")

f1.pack(fill="both", expand=True, padx=20, pady=20)
f2.place(in_=f1, anchor="c", relx=.5, rely=.5)


window.title("Does it confict?")

lbl = Label(f2, text="Enter the name of the product to check conficts:", background="SkyBlue3", font='Helvetica 14 bold')
lbl.place(anchor="c",relx=.5, rely=.2)
#lbl.grid(column=5, row=0)

txt = Entry(f2,width=10)

txt.place(anchor="c",relx=.5, rely=.25)
txt.focus()

btn = Button(f2, text="Check it!", command = main, bg="SkyBlue4")

btn.place(anchor="c",relx=.5, rely=.3)


window.mainloop()
	    	

if __name__ == "__main__":
    main()



