# first_chart_categories.py

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference


def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    # Add data to spreadsheet
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

    # Create the bar chart
    bar_chart = BarChart()
    bar_chart.title = "Book Sales"
    bar_chart.x_axis.title = "Book Types"
    bar_chart.y_axis.title = "Prices"
    bar_chart.set_x_axis({'num_font':  {'rotation': 45}})


    data = Reference(worksheet=sheet,
                     min_row=1,
                     max_row=10,
                     min_col=2,
                     max_col=3)
    bar_chart.add_data(data, titles_from_data=True)

    # Add categories
    categories = Reference(worksheet=sheet,
                           min_row=2,
                           max_row=10,
                           min_col=1,
                           max_col=1)
    bar_chart.set_categories(categories)
    sheet.add_chart(bar_chart, "E2")

    workbook.save(filename)


if __name__ == "__main__":
    main("bar_chart_categories.xlsx")