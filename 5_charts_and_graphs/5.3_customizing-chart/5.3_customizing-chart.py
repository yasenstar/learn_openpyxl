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
my_chart.title = "Sales Data"
my_chart.add_data(chart_data)

my_chart.x_axis.title = "Categories"
my_chart.y_axis.title = "Sales Figures"

my_chart.series[0].name = "Sales Figures"

my_chart.legend = None

my_chart.width = 10
my_chart.height = 10

# Add Chart into Worksheet
sheet.add_chart(my_chart, "D2")

workbook.save("bar_chart_sample.xlsx")