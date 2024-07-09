Language: Python 3.12.4
Libraries: openpyxl, pandas, numpy
Install the libraries openpyxl, pandas, and numpy. 
Please use command "pip install pandas numpy openpyxl".

Have a.py, b.py, c.py and your excel file, in the same folder.
Please close excel file while you running the py file, or you will get an error message.

Description:
a.py *in progress* / reads the raw data 11 4x12tables, then assigns them into 48 6x6tables and creates a new excel file called Result.xlsx
b.py *in progress* / reads the table in Result.xlsx, then overwrite the helping table in the Result.xlsx with the RotatedResult table
c.py convert rotated table to a txt file.

instruction:
please run the following commands in this order 
python a.py fileName.xlsx
python b.py Result.xlsx
python c.py Result.xlsx
