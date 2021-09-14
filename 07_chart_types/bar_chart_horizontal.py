# bar_chart_horizontal.py

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference


def create_excel_data(sheet):
    data_rows = [
        ('Number', 'Batch 1', 'Batch 2'),
        (2, 10, 30),
        (3, 40, 60),
        (4, 50, 70),
        (5, 20, 10),
        (6, 10, 40),
        (7, 50, 30),
    ]
    for row in data_rows:
        sheet.append(row)


def create_bar_chart(sheet):
    bar_chart = BarChart()
    bar_chart.type = "bar"

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
    workbook.save("bar_chart_horizontal.xlsx")


if __name__ == "__main__":
    main()