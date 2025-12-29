from openpyxl import load_workbook

workbook = load_workbook("my_workbook.xlsx")

# Accessing the active worksheet:
active_worksheet = workbook.active
print(active_worksheet.cell(row=4,column=4).value)

# Accessing a worksheet by name:
worksheet_by_name = workbook["Sheet3"]
print(worksheet_by_name.cell(row=3,column=3).value)

# Accessing a worksheet by index (index starts from 0):
worksheet_by_index = workbook.worksheets[1]
print(worksheet_by_index.cell(row=2,column=2).value)

# Check if a worksheet exists before accessing it (avoid errors):
if "Sheet2" in workbook.sheetnames:
    worksheet2 = workbook["Sheet2"]
    print(worksheet2.cell(2,2).value)

if "Sheet10" in workbook.sheetnames:
    worksheet10 = workbook["Sheet10"]
else:
    print("Sheet10 is not existed!")