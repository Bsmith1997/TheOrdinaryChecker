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
    conflict_col = 1 #Just an example
    group_col = 2
    toCheck = txt.get()
    x = []
    y = []
    z = []
    #toCheck = input("Enter the name of the product to check conficts: ")
    for row in range(sheet.nrows):
        if sheet.cell(row,product_col).value == toCheck:
            group = sheet.cell(row,group_col).value
            lbl2 = Label(window, font='Helvetica 14 bold',text = toCheck + " conficts with: ")
            lbl2.grid(column=0, row=5)
            #messagebox.showinfo('Message title',toCheck + " conficts with: ")
            for rows in range(sheet.nrows):
                if group in sheet.row_values(rows)[1]:
                    x.append(sheet.row_values(rows)[0])
                    #lbl3 = Label(window, text=x)
                    #lbl3.grid(column=0)
                if toCheck in sheet.row_values(rows)[1]:
                    y.append(sheet.row_values(rows)[0])
                    #lbl4 = Label(window, text=y)
                    #lbl4.grid(column=0)
                if sheet.row_values(rows)[0] in sheet.cell(row,1).value:
                    z.append(sheet.row_values(rows)[0])
    x_set = set(x)
    y_set = set(y)
    z_set = set(z)

    total = x + y + z
    total_set = set(total)
   
    final = "\n ".join(str(e) for e in total_set)
    lbl4 = Label(window, text=final)
    lbl4.grid(column=0)

window = Tk()
window.geometry("500x500")

window.title("Does it confict?")

lbl = Label(window, text="Enter the name of the product to check conficts:")

lbl.grid(column=0, row=0)

txt = Entry(window,width=10)

txt.grid(column=0, row=2)
txt.focus()


btn = Button(window, text="Check it!", command = main)

btn.grid(column=0, row=4)

window.mainloop()
	    	

if __name__ == "__main__":
    main()



