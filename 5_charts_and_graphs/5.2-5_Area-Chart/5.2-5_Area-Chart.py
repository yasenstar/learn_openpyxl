from openpyxl import Workbook
from openpyxl.chart import Reference, AreaChart

workbook = Workbook()
sheet = workbook.active
sheet.title = "Web Analytics"

data = [
    ["Day", "Organic Traffic", "Paid Traffic"],
    ["Mon", 40, 20],
    ["Tue", 45, 25],
    ["Wed", 50, 30],
    ["Thu", 35, 27],
    ["Fri", 70, 46],
    ["Sat", 85, 54],
    ["Sun", 65, 40]
]

for row in data:
    sheet.append(row)

chart = AreaChart()
chart.title = "Weekly Traffic Volume"
chart.style = 42
chart.x_axis.title = "Day of Week"
chart.y_axis.title = "Visits (thousands)"

data_ref = Reference(sheet, min_col=2, max_col=3, min_row=1, max_row=8)
chart.add_data(data_ref, titles_from_data=True)

cats_ref = Reference(sheet, min_col=1, max_col=1, min_row=2, max_row=8)
chart.set_categories(cats_ref)

chart.grouping = "stacked"

sheet.add_chart(chart, "E2")

workbook.save("Area_Chart_Example.xlsx")