from openpyxl import Workbook

workbook = Workbook()

# Create a new worksheet name "Data" at the end
new_worksheet = workbook.create_sheet(title="Data")

# Create a new worksheet named "Summary" before the sheet named "Data"
summary_worksheet = workbook.create_sheet(
    title = "Summary",
    index = 1
    # before = new_worksheet # TypeError: Workbook.create_sheet() got an unexpected keyword argument 'before'
)

sales_worksheet = workbook.create_sheet(
    title = "Sales",
    index = 2
    # before = new_worksheet # TypeError: Workbook.create_sheet() got an unexpected keyword argument 'before'
)

workbook.save("workbook_with_sheets.xlsx")