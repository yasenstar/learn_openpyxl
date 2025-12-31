from openpyxl import Workbook
from openpyxl.styles import Alignment

workbook = Workbook()
sheet = workbook.active

sheet["C5"] = "Center Aligned Text"

my_alignment = Alignment(
    horizontal = "center",
    vertical = "top",
    wrapText = True
)

sheet["C5"].alignment = my_alignment

workbook.save("alignment_styles.xlsx")