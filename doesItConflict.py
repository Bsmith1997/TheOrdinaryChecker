from xlrd import open_workbook
import sys
import re

def main():
    book = open_workbook('TheOrdinaryConflicts.xlsx')
    sheet = book.sheet_by_index(0)
    product_col = 0
    conflict_col = 1 #Just an example
    group_col = 2
    x = []
    y = []
    z = []
    toCheck = input("Enter the name of the product to check conficts: ")
    for row in range(sheet.nrows):
        if sheet.cell(row,product_col).value == toCheck:
            group = sheet.cell(row,group_col).value
            print(toCheck, "conficts with: ")
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

    print(set(total))    
    #print(y)
    #print(z)

if __name__ == "__main__":
    main()

#toCheck = input("Enter the name of the product to check conficts: ")

#print(data)
#for index, row in data.iterrows():
 #   genval = row[['Product ','Conflicts','Group']]
  #  if(row['Product']=='Toning solution'):
   # 	genval['side_chosen']='left'
    #	df = df.append(genval,ignore_index=True)
#df.to_excel('yourProductConflictsWith.xlsx')


