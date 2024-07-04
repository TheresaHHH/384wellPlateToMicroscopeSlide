import openpyxl

# Load the workbook
wb = openpyxl.load_workbook('Book4.xlsx')

# Check if 'Result' sheet exists and if so, delete it
if 'Result' in wb.sheetnames:
    del wb['Result']

# Select the sheet you want to copy from
source = wb['Sheet1']

# Create a new sheet and select it
destination = wb.create_sheet('Result')

# Copy the values from the source sheet to the destination sheet
for row in source:
    for cell in row:
        destination[cell.coordinate].value = cell.value

# Save the workbook
