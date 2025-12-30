from openpyxl import load_workbook

workbook = load_workbook("my_workbook.xlsx")

worksheet = workbook["Sheet2"]
worksheet.title = "New Sheet Name"

workbook.save("my_workbook_renamed.xlsx")