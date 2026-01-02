from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

workbook = Workbook()
sheet = workbook.active

# Prepare Data
data = [
    ["Name", "Age", "City"],
    ["Alice", 30, "New York"],
    ["Bob", 25, "London"],
    ["Charlie", 35, "Paris"]
]

for row in data:
    sheet.append(row)

# Create Table
table = Table(displayName="MyTable", ref="A1:C4")

# Define a table style
custom_style = TableStyleInfo(
    name = "TableStyleMedium9",
    showFirstColumn=True,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=True,
    # pivotButton=True  # no more this option
)

custom_style.font = "Arial,20"
custom_style.fill = {
    "type": "solid",
    "fgColor": "FFFFCC"
}

table.tableStyleInfo = custom_style

# Add the table to the worksheet
sheet.add_table(table)

workbook.save("custom_table_style.xlsx")

# 2025/01/12: the custom style is not effective, need further investigation