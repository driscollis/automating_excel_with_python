# scatter_chart.py

from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series


def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    rows = [
        ["Book", "Kindle", "Paperback"],
        [1, 9.99, 25.99],
        [2, 9.99, 25.99],
        [3, 9.99, 25.99],
        [4, 4.99, 29.99],
        [5, 4.99, 29.99],
        [6, 24.99, 29.99],
        [7, 24.99, 65.00],
        [8, 24.99, 69.00],
        [9, 24.99, 69.00],
    ]

    for row in rows:
        sheet.append(row)

    chart = ScatterChart()
    chart.title = "Scatter Chart"
    chart.style = 13
    chart.x_axis.title = 'Size'
    chart.y_axis.title = 'Percentage'

    xvalues = Reference(sheet, min_col=1, min_row=2, max_row=7)
    for i in range(2, 4):
        values = Reference(sheet, min_col=i, min_row=1, max_row=7)
        series = Series(values, xvalues, title_from_data=True)
        chart.series.append(series)

    sheet.add_chart(chart, "E2")

    workbook.save(filename)


if __name__ == "__main__":
    main("scatter_chart.xlsx")