from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference

# 0. Initialize WorkBook and WorkSheet

workbook = Workbook()
sheet = workbook.active
sheet.title = "Sales Report"

# 1. Prepare the Data

rows = [
    ["Month", "Online Sales", "InStore Sales"],
    ["Jan", 150, 100],
    ["Feb", 180, 120],
    ["Mar", 210, 180],
    ["Apr", 190, 210],
    ["May", 250, 200],
    ["Jun", 300, 230],
    ["Jul", 270, 190],
]

for row in rows:
    sheet.append(row)

# 2. Initialize the Chart - Line Chart

myChart = LineChart()
myChart.title = "Monthly Sales Trends"
myChart.x_axis.title = "Month"
myChart.y_axis.title = "Sales Revenue ($)"

# 3. Define the Data (Column B and C)

myData = Reference(sheet, min_col=2, max_col=3, min_row=2, max_row = 8)
myChart.add_data(myData)

# 4. Define Categories (Column A)

myCats = Reference(sheet, min_col=1, min_row=2, max_row=8)
myChart.set_categories(myCats)

# 5. Add chart into the sheet

sheet.add_chart(myChart, "E2")

# 6. Save workbook

workbook.save("Line_Chart_Sample.xlsx")