from openpyxl import Workbook
from openpyxl.styles import Font

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Styled Text"

my_font = Font(
    name = "Arial",
    size = 20,
    bold = True,
    italic = True,
    underline = "double",
    strikethrough = True,
    shadow = True,
    color = "0000FF"
)

sheet["A1"].font = my_font

workbook.save("font_styles.xlsx")