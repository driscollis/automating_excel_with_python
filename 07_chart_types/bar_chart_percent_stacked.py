# bar_chart_percent_stacked.py

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference


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
    ]

    for row in data_rows:
        sheet.append(row)


def create_bar_chart(sheet):
    bar_chart = BarChart()
    bar_chart.type = "bar"
    bar_chart.style = 13
    bar_chart.grouping = "percentStacked"
    bar_chart.overlap = 100
    bar_chart.title = 'Percent Stacked Chart'

    data = Reference(worksheet=sheet,
                     min_row=1,
                     max_row=10,
                     min_col=2,
                     max_col=3)
    bar_chart.add_data(data, titles_from_data=True)
    sheet.add_chart(bar_chart, "E2")


def main():
    workbook = Workbook()
    sheet = workbook.active
    create_excel_data(sheet)
    create_bar_chart(sheet)
    workbook.save("bar_chart_percent_stacked.xlsx")


if __name__ == "__main__":
    main()