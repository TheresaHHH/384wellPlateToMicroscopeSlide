import pandas as pd
import numpy as np
import openpyxl
import sys

fileName= sys.argv[1]
op0 = input('The fileName is '+ fileName +'. Enter 0 to input new fileName:')
if op0 == '0':
    fileName = input("Enter fileName.xlsx:");
else:
    print(fileName + ' will be read. Please close it if you have opened it.');

#####it only runs when when you run this file it more than once 
newSheet= 'RotatedTable';
wb = openpyxl.load_workbook(fileName);
if newSheet in wb.sheetnames:
    del wb[newSheet]
    wb.save(fileName);
  
##Please make sure the extension is .xlsx
#reads the excel file
excel_File = pd.ExcelFile(fileName);    
excelSheets = excel_File.sheet_names;

#How many columns and rows in each small table?
firstTableColumn = 12;
firstTableRow = 4;

firstTables = [];
