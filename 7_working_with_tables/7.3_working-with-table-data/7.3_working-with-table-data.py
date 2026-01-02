from openpyxl import load_workbook

workbook = load_workbook("table_sample.xlsx")
sheet = workbook.active
table = sheet["MyTable"]

# Accessing a specific cell in a table
cell_value = table["A2"].value
print(f"Value of A2: {cell_value}")

workbook.save("updated_table.xlsx")