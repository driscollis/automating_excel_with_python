# area_chart.py

from openpyxl import Workbook
from openpyxl.chart import AreaChart, Reference


def create_excel_data(sheet):
    data_rows = [
        ["Book", "Kindle", "Paperback"],
        ["Python 101", 9.99, 15.99],
        ["Python 201", 9.99, 25.99],
        ["ReportLab", 9.99, 25.99],
        ["wxPython", 4.99, 29.99],
        ["Jupyter", 14.99, 39.99],
    ]

    for row in data_rows:
        sheet.append(row)


def create_chart(sheet):
    chart = AreaChart()
    chart.style = 23
    chart.title = "Book Sales"
    chart.x_axis.title = "Book Types"
    chart.y_axis.title = "Prices"

    data = Reference(worksheet=sheet,
                     min_row=1,
                     max_row=10,
                     min_col=2,
                     max_col=3)
    chart.add_data(data, titles_from_data=True)
    sheet.add_chart(chart, "E2")


def main():
    workbook = Workbook()
    sheet = workbook.active
    create_excel_data(sheet)
    create_chart(sheet)
    workbook.save("area_chart.xlsx")


if __name__ == "__main__":
    main()