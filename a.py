import pandas as pd
import numpy as np

#reads the excel file
excel_File=pd.ExcelFile('Book4.xlsx');
excelSheets = excel_File.sheet_names;

#How many columns and rows in each small table?
firstTableColumn = 12;
firstTableRow = 4;

firstTables = [];

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
    for j in range(4):
        firstTable=[];
        for i in range(firstTableRow):
            tempRow = file.loc[i+j*firstTableRow, firstTableColumn: firstTableColumn*2-1].values.tolist();
            #firstTable.insert(len(firstTable), tempRow);
            firstTable.extend(tempRow);
        firstTables.insert(len(firstTables), firstTable);        
    print('sheetName: '+ str(sheets) + ' is read.');
 
 
 # Organizes the cells into tables based on their coordinators Adds 'blank to each table
zippedTables0 = [list(x) for x in zip(*firstTables)];
zippedTables0 = [sublist + ['blank'] for sublist in zippedTables0];
#print(zippedTables0);

#48 pre-rotated tables in 4 rows. 12 tables in a row. 
##Please change the values if different. 
secondTableColumn = 12;
secondTableRow = 4;
zippedTables1 = [];
for i in range(secondTableRow):
    zippedTables1.insert(len(zippedTables1), zippedTables0[i*secondTableColumn:secondTableColumn*(i+1)]);
 
#Rotates to the right. The order of cells is not correct. 
rotatedTables0 = np.rot90(zippedTables1, -1);
rotatedTables1 = rotatedTables0.tolist();
#print(rotatedTables1)

# Fixes the order of cells.
##Please replace 7 with the desired value. How many cells in a row? 
for i in range(len(rotatedTables1)):
    for j in range(len(rotatedTables1[i])):
            temp = [rotatedTables1[i][j][z : z+7] for z in range(0, len(rotatedTables1[i][j]),7)];
            ##reverse the data in each row.
            temp = [sublist[::-1] for sublist in temp];
            rotatedTables1[i][j] = temp;
