import xlrd
import xlwt
import os
from os import listdir

sheet_name ="Sheet1"
output_file_name= "Merge.xlsx"

def program():
	arr = readFiles()
	writeFile(arr)


def readFiles():
    files= []
    path = os.path.dirname(os.path.abspath(__file__))
    for f in listdir(path):
        if f[-5:] == ".xlsx":
            files.append(f)
    rows= [] # each element is a row that contains 5 values
    for file in files:
        rows = rows + readFile(file)
    return rows


#returns an array, where each element is an array itslef
def readFile(file):
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)
    row_num  = 0
    has_next = True
    rows = []
    while (has_next):
        row_num = row_num+1
        row = readRow(sheet,row_num)
        if (len(row) >0):
            rows.append(row)
        else:
            has_next = False
    return rows


#This method reads a row and returns an array of  value
def readRow(row,row_num):
    arr_rows = []
    
    try:
        for x in range(0,4):
            val =  row.cell(row_num,x).value
            arr_rows.append(val)
    except Exception, e:
        ""

    return arr_rows


def_cols = ["a","b","c","d"]
def writeFile(arr):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)  

    for i in range(0, len(def_cols)):
        sheet.write(0, i, def_cols[i])

    i = 1
    for a in arr:
        j = 0
        for b in a:
            sheet.write(i, j,b)
            j = j+1
        i = i+1
    workbook.save(output_file_name)

program()



















