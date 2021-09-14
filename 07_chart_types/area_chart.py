# area_chart.py

from openpyxl import Workbook
from openpyxl.chart import AreaChart, Reference


def create_excel_data(sheet):
    data_rows = [
        ["Book", "Kindle", "Paperback"],
        ["Python 101", 15, 5],
        ["Python 201", 5, 1],
        ["ReportLab", 10, 0],
        ["wxPython", 2, 2],
        ["Jupyter", 25, 15],
    ]

    for row in data_rows:
        sheet.append(row)


def create_chart(sheet):
    chart = AreaChart()
    chart.style = 23
    chart.title = "Book Sales"
    chart.x_axis.title = "Book"
    chart.y_axis.title = "Copies Sold"

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
    workbook.save("area_chart_3d.xlsx")


if __name__ == "__main__":
    main()