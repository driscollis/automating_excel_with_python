# pie_chart.py

from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.series import DataPoint


def main(filename):
    workbook = Workbook()
    sheet = workbook.active

    data = [
        ['Pizza', 'Sold'],
        ['Pepperoni', 50],
        ['Sausage', 30],
        ['Cheese', 10],
        ['Supreme', 40],
    ]

    for row in data:
        sheet.append(row)

    chart = PieChart()
    chart.title = "Pizza Pie Chart"
    labels = Reference(sheet, min_col=1, min_row=2, max_row=5)
    data = Reference(sheet, min_col=2, min_row=1, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(labels)

    # Cut the first slice of pizza from the pie
    slice = DataPoint(idx=0, explosion=20)
    chart.series[0].data_points = [slice]

    sheet.add_chart(chart, "E2")

    workbook.save(filename)


if __name__ == "__main__":
    main("pie_chart.xlsx")