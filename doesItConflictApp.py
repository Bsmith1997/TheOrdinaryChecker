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
            lbl2 = Label(f2, font='Helvetica 14 bold',text = toCheck + " conficts with: ")
            lbl2.place(anchor="c",relx=.5, rely=.4)
            for rows in range(sheet.nrows):
                if group in sheet.row_values(rows)[1]:
                    x.append(sheet.row_values(rows)[0])
                if toCheck in sheet.row_values(rows)[1]:
                    y.append(sheet.row_values(rows)[0])
                if sheet.row_values(rows)[0] in sheet.cell(row,1).value:
                    z.append(sheet.row_values(rows)[0])
    x_set = set(x)
    y_set = set(y)
    z_set = set(z)

    total = x + y + z
    total_set = set(total)
   
    final = "\n ".join(str(e) for e in total_set)
    lbl4 = Label(f2, text=final)
    lbl4.place(anchor="c",relx=.5, rely=.5)

window = Tk()
window.geometry("500x500")
f1 = Frame(width=800, height=600, background="white")
f2 = Frame(width=775, height=575, background="white")
f1.pack(fill="both", expand=True, padx=20, pady=2)
f2.place(in_=f1, anchor="c", relx=.5, rely=.5)

f1.pack(fill="both", expand=True, padx=20, pady=20)
f2.place(in_=f1, anchor="c", relx=.5, rely=.5)

window.title("Does it confict?")

lbl = Label(f2, text="Enter the name of the product to check conficts:")
lbl.place(anchor="c",relx=.5, rely=.2)
#lbl.grid(column=5, row=0)

txt = Entry(f2,width=10)

txt.place(anchor="c",relx=.5, rely=.25)
txt.focus()

btn = Button(f2, text="Check it!", command = main)

btn.place(anchor="c",relx=.5, rely=.3)

window.mainloop()
	    	

if __name__ == "__main__":
    main()



