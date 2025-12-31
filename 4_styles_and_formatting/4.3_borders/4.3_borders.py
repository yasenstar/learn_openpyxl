from openpyxl import Workbook
from openpyxl.styles import Border, Side

workbook = Workbook()
sheet = workbook.active

sheet["B2"] = "Thin Border"

sheet["D2"] = "Thick Border"

thin_border = Border(
    left = Side(style = 'thin'),
    right = Side(style = 'thin'),
    top = Side(style = 'thin'),
    bottom = Side(style = 'dashDotDot', color = "0000FF")
)

thick_border = Border(
    left = Side(style = 'thick', color = "FF0000"),
    right = Side(style = 'thick'),
    top = Side(style = 'thick'),
    bottom = Side(style = 'thick', color = "0000FF")
)

sheet["B2"].border = thin_border

sheet["D2"].border = thick_border

workbook.save("border_styles.xlsx")