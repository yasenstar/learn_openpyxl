# openpyxl - 7. Working with Tables

- [openpyxl - 7. Working with Tables](#openpyxl---7-working-with-tables)
  - [7.1 Creating Tables](#71-creating-tables)
  - [7.2 Formatting Tables](#72-formatting-tables)
  - [7.3 Working with Table Data](#73-working-with-table-data)
  - [7.4 Filtering and Sorting](#74-filtering-and-sorting)

## 7.1 Creating Tables

openpyxl provides functionality to work with Excel tables (also known as `ListObjects`).

Creating a table involves specifying a range of cells and optionally providing a table name.

```python
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
```

![create-table](img/create-table.png)

Source Code Reference:

- `worksheet/table.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/worksheet/table.py

## 7.2 Formatting Tables

Table formatting involves applying styles to the table’s appearance, such as header row styles, banded rows, total row, etc.

This is done by modifying the tableStyleInfo attribute of the Table object. You can choose from predefined styles or create custom styles.

```python
```

Source Code Reference:

- `TableStyleInfo`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/worksheet/table.py#L41

## 7.3 Working with Table Data

```python
```

## 7.4 Filtering and Sorting

```python
```

---

Last Updated at: 12/28/2025, 8:18:39 AM 