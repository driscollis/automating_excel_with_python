# line_chart.py

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference


def create_excel_data(sheet):
    data_rows = [
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

    for row in data_rows:
        sheet.append(row)


def create_bar_chart(sheet):
    chart = LineChart()
    chart.title = "Line Chart"
    chart.style = 15
    chart.y_axis.title = 'Sales'
    chart.x_axis.title = 'Books'

    data = Reference(sheet, min_col=2, min_row=2, max_col=3, max_row=9)
    chart.add_data(data)

    # Style the lines
    series_1 = chart.series[0]
    series_1.marker.symbol = "triangle"
    # Marker filling color
    series_1.marker.graphicalProperties.solidFill = "FF0000"
    # Marker outline color
    series_1.marker.graphicalProperties.line.solidFill = "FF0000"

    series_1.graphicalProperties.line.noFill = True

    series_2 = chart.series[1]
    series_2.graphicalProperties.line.solidFill = "00AAAA"
    series_2.graphicalProperties.line.dashStyle = "sysDot"
    series_2.graphicalProperties.line.width = 100050 # width in EMUs

    sheet.add_chart(chart, "E2")


def main():
    workbook = Workbook()
    sheet = workbook.active
    create_excel_data(sheet)
    create_bar_chart(sheet)
    workbook.save("line_chart.xlsx")


if __name__ == "__main__":
    main()