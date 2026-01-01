from openpyxl import Workbook
from openpyxl.chart import Reference, RadarChart

workbook = Workbook()
sheet = workbook.active
sheet.title = "Skill Assessmnet"

data = [
    ["Metric", "Developer A", "Developer B", "Developer C"],
    ["Coding", 90, 94, 85],
    ["Debugging", 86, 89, 93],
    ["Documentation", 70, 95, 84],
    ["Speed", 85, 70, 80],
    ["Teamwork", 93, 85, 65]
]

for row in data:
    sheet.append(row)

chart = RadarChart()
chart.title = "Developer Skills Assessment"
chart.style = 26

data = Reference(sheet, min_col=2, max_col=4, min_row=1, max_row=6)
labels = Reference(sheet, min_col=1, max_col=1, min_row=2, max_row=6)

chart.add_data(data, titles_from_data=True)
chart.set_categories(labels)

sheet.add_chart(chart, "F2")

workbook.save("Radar_Chart_Sample.xlsx")