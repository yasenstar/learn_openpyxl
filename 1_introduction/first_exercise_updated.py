# Name: first_exercise
# Purpose: practive writing into spreadsheet via iterating
# Author: Xiaoqi Zhao
# Date: 2025/12/29

from openpyxl import Workbook

# Provide the dataset for writing into the spreadsheet
dataset = [
    ["Month", "Days"],
    ["January", 31],
    ["February", 28],
    ["March", 31],
    ["April", 30]
]

# Writing to a new file
new_workbook = Workbook()
new_worksheet = new_workbook.active

# Iterating through rows
for row in range(1,len(dataset)+1):
    for col in range(1,len(dataset[0])+1):
        new_worksheet.cell(row,col).value = dataset[row-1][col-1]
        print(f"Cell value [{row}][{col}] is: {new_worksheet.cell(row,col).value}")

# Save the new workbook
new_workbook.save("first_exercise.xlsx")