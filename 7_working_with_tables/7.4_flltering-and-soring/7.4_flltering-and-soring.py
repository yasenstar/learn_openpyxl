from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.filters import AutoFilter

# 1. Create workbook and data

workbook = Workbook()
sheet = workbook.active
sheet.title = "Inventory"

data = [
    ["Product", "Category", "Quantity", "Price"],
    ["Laptop", "Electronics", 5, 1200],
    ["Mouse", "Electronics", 25, 25],
    ["Desk Chair", "Furniture", 10, 200],
    ["Monitor", "Electronics", 8, 400],
    ["Bookshelf", "Furniture", 3, 150],
    ["Keyboard", "Electronics", 15, 45]
]

for row in data:
    sheet.append(row)

# 2. Define the range of the table (A1 to D7)

table_range = "A1:D7"

# 3. Create a Table object
myTable = Table(displayName="InventoryTable", ref=table_range)

# 4. Add a Table Style

myStyle = TableStyleInfo(
    name = "TableStyleMedium9",
    showFirstColumn = False,
    showLastColumn = False,
    showRowStripes = True,
    showColumnStripes = False
)
myTable.tableStyleInfo = myStyle

# 5. Define Filtering Logic (Metadata)

myTable.autoFilter = AutoFilter(ref=table_range)

myTable.autoFilter.add_filter_column(1, ["Electronics"])

# 6. Define Soring Logic (Metadata)

myTable.autoFilter.add_sort_condition("A2:A7", descending = False)

# 7. Add the table to the worksheet

sheet.add_table(myTable)

# 8. Save Workbook

workbook.save("Filtering-and-Sorting.xlsx")