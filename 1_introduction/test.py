from openpyxl import Workbook

dataset = [
    ["Month", "Days", "memo"],
    ["January", 31, "test"],
    ["February", 28, "test"],
    ["March", 31, "test"],
    ["April", 30, "test"],
    ["May", 31, "test"]
]

my_workbook = Workbook()
my_worksheet = my_workbook.active

for row in range(1,len(dataset)+1):
    for col in range(1,len(dataset[0])+1):
        my_worksheet.cell(row,col).value = dataset[row-1][col-1]
        print(f"Cell value [{row}][{col}] is: {my_worksheet.cell(row,col).value}")

my_workbook.save("my_file.xlsx")