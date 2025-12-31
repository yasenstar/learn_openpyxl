from openpyxl import Workbook
from datetime import datetime
from openpyxl.styles import numbers

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = datetime(2025, 12, 31, 13, 19, 00)
sheet["A2"] = "2025/11/05"
sheet["A2"].number_format = numbers.FORMAT_DATE_YYYYMMDD2
sheet["A3"] = datetime(2025, 12, 31, 13, 19, 00)
sheet["A3"].number_format = numbers.FORMAT_DATE_TIME6

workbook.save("date_time_sample.xlsx")