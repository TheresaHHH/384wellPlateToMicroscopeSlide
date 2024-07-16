Language: Python 3.12.4

Libraries: openpyxl, pandas, numpy


Description:

a.py read the raw data 11 tables of 48 cells. 12 cells in a row.
Organized them into new 48 tables according to their cordinatates. 3 occurrences of the original data set in each table.
Positioned these 48 tables into 12 tables in a row. Now a big table is created.
Rotated this big table 90 degrees clockwise.
Output the result as BeautifulTable.txt and BeautifulTables.xlsx if needed.

c.py convert rotatedTable to a txt file only.

instruction:

Install the libraries openpyxl, pandas, and numpy.

Please use command "pip install pandas numpy openpyxl".

Have a.py, c.py and your excel file in the same folder.

Please close excel file while you running the py file, or you will get an error message.

python a.py fileName.xlsx

python c.py Result.xlsx
