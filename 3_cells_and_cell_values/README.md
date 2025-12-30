# openpyxl - 3. Cells and Cell Values

- [openpyxl - 3. Cells and Cell Values](#openpyxl---3-cells-and-cell-values)
  - [3.1 Accessing Cell Values](#31-accessing-cell-values)
  - [3.2 Setting Cell Values](#32-setting-cell-values)
  - [3.3 Data Types](#33-data-types)
  - [3.4 Formulas and Calculations](#34-formulas-and-calculations)
  - [3.5 Number Formatting](#35-number-formatting)
  - [3.6 Dates and Times](#36-dates-and-times)
  - [3.7 Working with Cell Ranges](#37-working-with-cell-ranges)

## 3.1 Accessing Cell Values

Cell values are accessed using several methods.

The most common method is indexing using the cell's coordinates as a string (e.g., "A1") or using the `cell()`1 method, which takes row and column numbers, see below sample:

```python
from openpyxl import load_workbook

workbook = load_workbook("my_workbook.xlsx")
sheet = workbook.active

# Accessing using string index
cell_value = sheet["A2"].value
print(f"Value of A1 is: {cell_value}")

# Accessing using cell() method (row, column)
cell_value1 = sheet.cell(row=2, column=4).value
print(f"Value of D2 is: {cell_value1}")

# Checking for a None value (empty cell)
if sheet["F1"].value is None:
    print("Cell F1 is Empty")

if sheet["A200"].value is None:
    print("Cell A200 is Empty")
else:
    print(f"Value of A200 is: {sheet["A200"].value}")
```

Source code reference:

- `Worksheet.cell(self, row, column, value=None)`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/worksheet/worksheet.py?ref_type=heads#L220

Note:

1. the `cell` coordinates are 1-based (the top-left call is A1).
2. an empty `cell` will have a `value` of `None`.

Remeber - same as Excel - that cell coordinates are 1-based (the top-left cell is "A1").

An empty cell will have a `value` of `None`.

## 3.2 Setting Cell Values

Setting a `cell`'s value is equally simple, as below sample:

```python
```

You can assign various data types to cells.

## 3.3 Data Types

openpyxl hands several data types as below:

| Data Type | Description |
| --- | --- |
| Numbers | Integers, floats, etc. are stored as numbers. |
| Strings | Text values are stored as strings |
| Booleans | `True` and `False` are supported. |
| Dates and Times | These are stored as Python `datetime` objects. |
| Formulas | Formulas are stored as strings, but openpyxl can evalute some simple formulas |
| None | Represents an empty cell |

Source code reference:

- `VALID_TYPES` in `cell.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/cell/cell.py?ref_type=heads#L58

## 3.4 Formulas and Calculations

openpyxl can handle formulas in cells, as below sample:

```python
```

## 3.5 Number Formatting

To format numbers, use the `number_format` property of the cell:

```python
```

Source Code Reference:

- `cell.py`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/cell/cell.py?ref_type=heads

## 3.6 Dates and Times

Dates and times are represented using Python's `datetime` objects:

```python
```

openpyxl handls the conversion between Python's `datetime` objects and Excel's date system automatically.

## 3.7 Working with Cell Ranges

You can efficiently work with ranges of cells using `sheet.iter_rows()` and `sheet.iter_cols()`. These methods provide iterators to traverse ranges efficiently:

```python
```

Source code reference:

- `Worksheet.iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False)`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/worksheet/worksheet.py?ref_type=heads#L405
  - `_cells_by_row(self, min_col, min_row, max_col, max_row, values_only=False)`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/worksheet/worksheet.py?ref_type=heads#L444
- `Worksheet.iter_cols(self, min_col=None, max_col=None, min_row=None, max_row=None, values_only=False)`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/worksheet/worksheet.py?ref_type=heads#L472
  - `_cells_by_col(self, min_col, min_row, max_col, max_row, values_only=False)`: https://foss.heptapod.net/openpyxl/openpyxl/-/blob/branch/default/openpyxl/worksheet/worksheet.py?ref_type=heads#L510

---

Last Updated at: 12/28/2025, 8:14:43 AM 