from openpyxl import Workbook
from openpyxl.styles import numbers

workbook = Workbook()
sheet = workbook.active

sheet["A1"], sheet["A2"], sheet["A3"], sheet["A4"] = 1234.56, 1234.56, 1234.56, 1234.56

sheet["B1"], sheet["B2"], sheet["B3"], sheet["B4"] = 1234.56, 1234.56, 1234.56, 1234.56

sheet["B1"].number_format = numbers.FORMAT_PERCENTAGE
sheet["B2"].number_format = numbers.FORMAT_PERCENTAGE_00
sheet["B3"].number_format = '0.000'
sheet["B4"].number_format = '##,##0'

workbook.save("number_formats.xlsx")