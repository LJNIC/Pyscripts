import openpyxl, sys, tkinter, shlex
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

def progress(count, total, status=''):
    bar_len = 50
    filled_len = int(round(bar_len * count / float(total)))

    percents = round(100.0 * count / float(total), 1)
    bar = '#' * filled_len + ' ' * (bar_len - filled_len)

    sys.stdout.write('[%s] %s%s Working...%s\r' % (bar, percents, '%', status))
    sys.stdout.flush()

root = tkinter.Tk()
root.withdraw()

messagebox.showinfo("Warning", "Make sure to save a copy of the file in case the program does something unexpected!\n\nNote: The program might take a while to finish depending on the file's length and options chosen.")
try:
    fileName = filedialog.askopenfilename(initialdir = "C:\\", filetypes = [("Excel files", ("*.xlsx", "*.xls"))])
except FileNotFoundError:
    print('File not found or file not chosen. Closing program...')
    quit()

try:
    workbook = openpyxl.load_workbook(fileName)
except openpyxl.utils.exceptions.InvalidFileException:
    messagebox.showinfo('Error: File is not an Excel file')

row = int(input("Enter a row to begin at>"))
maxRow = int(input("Enter a row to end at>"))
barRow = row
barTotal = maxRow
columns = input("Enter the columns you want to check separated by commas(no spaces)>")
dontDelete = input("Enter any rows you want to keep as column,value separated by spaces: e.g: A,'Inspection Results' F,'Total'>")

columnList = columns.split(',') 
pairsList = None
if not dontDelete == '':
    pairs = shlex.split(dontDelete)
    pairsList = []
    for pair in pairs:
        split = pair.split(',')
        pairsList.append([split[0], split[1].strip("'")])

print(pairsList[0][0], pairsList[0][1])
sheet = workbook[workbook.sheetnames[0]]
progress(barRow, barTotal)
while row < maxRow:
    strRow = str(row)
    shouldDelete = True 
    for column in columnList:
        if not sheet[column + strRow].value == None:
           shouldDelete = False 
           break
    if shouldDelete and not pairsList == None:
        for pair in pairsList:
            if not sheet[pair[0] + strRow].value == None and sheet[pair[0] + strRow].value.strip() == pair[1]:
                shouldDelete = False
                break
    if shouldDelete == True:
        sheet.delete_rows(row, 1)
        row -= 1
        maxRow -= 1
    barRow += 1
    row += 1
    progress(barRow, barTotal)

workbook.save(fileName)
