from openpyxl import load_workbook

workbook = load_workbook("my_workbook.xlsx")

# Deleting a worksheet by name
sheet_to_delete = workbook["Sheet3"]
workbook.remove(sheet_to_delete)

# Deleting a worksheet by index
workbook.remove(workbook.worksheets[1])

workbook.save("my_workbook_modified.xlsx")