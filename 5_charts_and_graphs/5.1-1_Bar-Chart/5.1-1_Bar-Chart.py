from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference

workbook = Workbook()
sheet = workbook.active

# Prepare Data
data = [
    ['Category', 'Value'],
    ['A', 10],
    ['B', 15],
    ['C', 20],
    ['D', 17]
]

# Adding Data into worksheet
for row in data:
    sheet.append(row)

# Create Chart Data Reference
chart_data = Reference(
    sheet,
    min_col = 2,
    max_col = 2,
    min_row = 2,
    max_row = 5
)

# Create Bar Chart
my_chart = BarChart()
my_chart.add_data(chart_data)

# Add Chart into Worksheet
sheet.add_chart(my_chart, "D1")

workbook.save("bar_chart_sample.xlsx")