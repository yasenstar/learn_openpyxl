from openpyxl import Workbook
from openpyxl.chart import Reference, DoughnutChart

workbook = Workbook()
sheet = workbook.active
sheet.title = "Market Share for Cloud"

# Warning for too long Sheet title: seems max is 30
# UserWarning: Title is more than 31 characters. Some applications may not be able to read the file
#   warnings.warn("Title is more than 31 characters. Some applications may not be able to read the file")

data = [
    ["Category", "Percentage"],
    ["SaaS", 35],
    ["PaaS", 20],
    ["IaaS", 40],
    ["Others", 5]
]

for row in data:
    sheet.append(row)

chart = DoughnutChart()
data = Reference(sheet, min_col=2, max_col=2, min_row=1, max_row=5)
labels = Reference(sheet, min_col=1, max_col=1, min_row=2, max_row=5)

chart.add_data(data, titles_from_data=True)
chart.set_categories(labels)
chart.title = "Revenue Distribution"
chart.style = 30

chart.holeSize = 50

sheet.add_chart(chart, "D2")

workbook.save("Doughnut_Chart_Sample.xlsx")