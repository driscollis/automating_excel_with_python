# pie_chart_3d.py

from copy import deepcopy

from openpyxl import Workbook
from openpyxl.chart import PieChart3D, Reference


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

    pie = PieChart3D()
    labels = Reference(sheet, min_col=1, min_row=2, max_row=5)
    data = Reference(sheet, min_col=2, min_row=1, max_row=5)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Pizza Pies sold by type"
    sheet.add_chart(pie, "E2")

    workbook.save(filename)


if __name__ == "__main__":
    main("pie_chart_3d.xlsx")
