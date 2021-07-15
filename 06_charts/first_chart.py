# first_chart.py

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference


def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    # Add data to spreadsheet
    data_rows = [
        ["Book", "Kindle", "Paperback"],
        [1, 9.99, 15.99],
        [2, 9.99, 25.99],
        [3, 9.99, 25.99],
        [4, 4.99, 29.99],
        [5, 14.99, 39.99],
    ]

    for row in data_rows:
        sheet.append(row)

    # Create the bar chart
    bar_chart = BarChart()

    data = Reference(worksheet=sheet,
                     min_row=1,
                     max_row=10,
                     min_col=2,
                     max_col=3)
    bar_chart.add_data(data, titles_from_data=True)
    sheet.add_chart(bar_chart, "E2")

    workbook.save(filename)


if __name__ == "__main__":
    main("bar_chart.xlsx")