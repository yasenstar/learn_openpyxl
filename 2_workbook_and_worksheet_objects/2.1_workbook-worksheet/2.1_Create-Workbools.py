from openpyxl import Workbook

workbook = Workbook()

worksheet = workbook.active

worksheet.cell(2,3).value = "Hello Excel!"

workbook.save("my_new_workbook.xlsx")