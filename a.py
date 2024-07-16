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

#Read the excel file and create tables
for sheets in range(len(excelSheets)):
    #Dataframe
    file = excel_File.parse(excelSheets[sheets], header=None, names=None, index_col=None, usecols=None) ;

    ##How many tables are on the left side?
    for j in range(4):
        firstTable=[];
        for i in range(firstTableRow):
            tempRow = file.loc[i+j*firstTableRow, 0 : firstTableColumn-1].values.tolist();
            firstTable.extend(tempRow);
        firstTables.insert(len(firstTables), firstTable);

    ##How many tables are on the right side?
    for j in range(7):
        firstTable=[];
        for i in range(firstTableRow):
            tempRow = file.loc[i+j*firstTableRow, firstTableColumn: firstTableColumn*2-1].values.tolist();
            firstTable.extend(tempRow);
        firstTables.insert(len(firstTables), firstTable);        
    print('sheet: '+ str(sheets) + ' is read.');

'''
#testing    
print('firstTables')
for i in range(len(firstTables)):
    print(str(i))
    print(firstTables[i])
'''
