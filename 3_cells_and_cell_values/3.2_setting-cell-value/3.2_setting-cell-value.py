from openpyxl import Workbook

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Hello Excel"
sheet['B2'] = 42
sheet["C1"] = 3.1415
sheet["F2"] = "=SUM(B2,C1)"

sheet.cell(row = 4, column = 5).value = True

workbook.save("output.xlsx")