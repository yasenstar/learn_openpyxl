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
style = TableStyleInfo(
    name = "TableStyleMedium9",
    showFirstColumn=True,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=True
)
table.tableStyleInfo = style

# Add the table to the worksheet
sheet.add_table(table)

workbook.save("create_table_sample.xlsx")