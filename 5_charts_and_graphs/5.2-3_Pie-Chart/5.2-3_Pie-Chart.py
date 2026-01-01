from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList

# 0. Initialize WorkBook and WorkSheet

workbook = Workbook()
sheet = workbook.active
sheet.title = "Traffic Data"

# 1. Prepare the Data

rows = [
    ["Source", "Visitors"],
    ["Organic Search", 4500],
    ["Direct", 2500],
    ["Social Media", 1000],
    ["Referral", 800],
    ["Other", 1300]
]

for row in rows:
    sheet.append(row)

# 2. Initialize the Chart - Pie Chart

myChart = PieChart()
myChart.title = "Website Traffic Sources Statistics"

# 3. Define the Data (Column B)

myData = Reference(sheet, min_col=2, max_col=2, min_row=2, max_row = 6)
myChart.add_data(myData, titles_from_data = False)

# 4. Define Categories (Column A)

myCats = Reference(sheet, min_col=1, max_col=1, min_row=2, max_row=6)
myChart.set_categories(myCats)

# Optional: add Percent Labels

myChart.dataLabels = DataLabelList()
myChart.dataLabels.showPercent = True
myChart.dataLabels.showCategoryname = False

# 5. Add chart into the sheet

sheet.add_chart(myChart, "D2")

# 6. Save workbook

workbook.save("Pie_Chart_Sample.xlsx")
print("Pie Chart generated successfully!")