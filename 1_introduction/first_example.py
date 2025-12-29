from openpyxl import load_workbook, Workbook

# Reding from an existing file
workbook = load_workbook("example.xlsx")
worksheet = workbook.active

# Accessing cell values
cell_value = worksheet['A1'].value
print(f"Value of A1: {cell_value}")

# Iterating through rows
for row in worksheet.iter_rows(
    min_row=1,
    max_row=5,
    min_col=1,
    max_col=2
):
    for cell in row:
        print(f"Cell value [{row}][{cell}] is: {cell.value}")

# Writing to a new file
new_workbook = Workbook()
new_worksheet = new_workbook.active
new_worksheet['A1'] = "Hello, OpenPyXl"
new_worksheet['B2'] = 123

# Save the new workbook
new_workbook.save("new_example1.xlsx")