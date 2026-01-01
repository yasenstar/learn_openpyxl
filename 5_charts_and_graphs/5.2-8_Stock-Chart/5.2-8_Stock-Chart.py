from openpyxl import Workbook
from openpyxl.chart import Reference, StockChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.updown_bars import UpDownBars

workbook = Workbook()
sheet = workbook.active
sheet.title = "StockData"

data = [
    ["Date", "Open", "High", "Low", "Close"],
    ["2024-12-01", 100, 110, 95, 105],
    ["2024-12-02", 105, 115, 99, 112],
    ["2024-12-03", 112, 113, 102, 108],
    ["2024-12-04", 108, 120, 94, 118],
    ["2024-12-05", 118, 118, 104, 106],
    ["2024-12-06", 106, 124, 105, 113],
    ["2024-12-07", 113, 118, 111, 115],
    ["2024-12-08", 115, 115, 98, 123],
    ["2024-12-09", 123, 119, 109, 111]
]

for row in data:
    sheet.append(row)

chart = StockChart()
chart.title = "Stock Price Analysis"

data = Reference(sheet, min_col=2, max_col=5, min_row=1, max_row=10)
labels = Reference(sheet, min_col=1, max_col=1, min_row=2, max_row=10)

chart.add_data(data, titles_from_data=True)
chart.set_categories(labels)

volume_series = chart.series[0]

chart.hiLowLines = ChartLines()

chart.upDownBars = UpDownBars()

chart.x_axis.title = "Price (US$)"
chart.y_axis.title = "Data"

sheet.add_chart(chart, "G2")

workbook.save("Stock_Chart_Sample.xlsx")
print("Stock Chart Generated Successfully!")