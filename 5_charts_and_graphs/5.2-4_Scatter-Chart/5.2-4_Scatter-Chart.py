from openpyxl import Workbook
from openpyxl.chart import Reference, ScatterChart, Series

# 0. Initialize Workbook & Worksheet

workbook = Workbook()
sheet = workbook.active

# 1. Preparation of the Data

sheet.title = "Scatter Data"

data = [
    ["Ad Spend", "New Customer"],
    [100, 12],
    [200, 25],
    [300, 30],
    [400, 45],
    [500, 70],
    [600, 86]
]

# 2. Add data into Worksheet

for row in data:
    sheet.append(row)

# 3. Initialize Chart: Scatter Chart

myChart = ScatterChart()
myChart.title = "Ad Spend vs. Customer Aquisition"
myChart.x_axis.title = "Budget ($)"
myChart.y_axis.title = "New Customers"
# myChart.style = 15
myChart.scatterStyle = "marker"

# 4. Define X & Y Reference

xValues = Reference(sheet, min_col=1, max_col=1, min_row=2, max_row=7)
yValues = Reference(sheet, min_col=2, max_col=2, min_row=2, max_row=7)

# 5. Create Series object to Link X and Y

mySeries = Series(yValues, xValues, title = "Acquisition Rate")
myChart.series.append(mySeries)

# Optional: Markers

mySeries.marker.symbol = "circle"
mySeries.marker.graphicalProperties.solidFill = "0000FF"

# Add the Chart to the Worksheet

sheet.add_chart(myChart, "D2")

# Save the file

workbook.save("Scatter_Chart_Sample.xlsx")
print("Scatter Chart generated successfully!")