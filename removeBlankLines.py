import openpyxl, sys, tkinter
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
root = tkinter.Tk()

messagebox.showinfo("Warning", "Make sure to save a copy of the file in case the program does something unexpected!")
fileName = filedialog.askopenfilename(initialdir = "C:\\", filetypes = [("Excel files", ("*.xlsx", "*.xls"))])
try:
    workbook = openpyxl.load_workbook(fileName)
except openpyxl.utils.exceptions.InvalidFileException:
    print('Error: File is not an Excel file')

row = int(input("Enter a row to begin at>"))
maxRow = int(input("Enter a row to end at>"))
columns = input("Enter the columns you want to check separated by commas(no spaces)>")
dontDelete = input("Enter any rows you want to keep as column,value separated by spaces: e.g: A,Inspection F,Total>")

columnList = columns.split(',') 
pairsList = None

if not dontDelete == '':
    pairs = dontDelete.split(',')
    pairsList = []
    for pair in pairs:
        split = pair.split(',')
        pairsList.append([split[0], split[1]])

sheet = workbook[workbook.sheetnames[0]]

while row < maxRow:
    strRow = str(row)
    shouldDelete = True 
    for column in columnList:
        if not sheet[column + strRow].value == None:
           shouldDelete = False 
    if shouldDelete and not pairsList == None:
        for pair in pairsList:
            if sheet[pair[0] + strRow].strip() == pair[1]:
                shouldDelete = False
    if shouldDelete:
        sheet.delete_rows(row, 1)
        row -= 1
        maxRow -= 1
    row += 1
    progressBar["value"] = row
workbook.save(fileName)
